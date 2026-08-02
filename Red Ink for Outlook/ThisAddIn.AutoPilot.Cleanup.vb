' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.Cleanup.vb
' Purpose:
'   Inky AutoPilot retention cleanup for Outlook.
'   Automatically deletes AutoPilot-processed mail threads after a configured
'   number of hours, including items already moved to Deleted Items.
'
' Architecture:
'  - Uses hidden MAPI properties to group an incoming mail with all AutoPilot
'    replies and follow-up notices.
'  - Starts the retention clock only after a substantive reply is sent.
'  - Runs a periodic cleanup timer alongside the AutoPilot session.
'  - Scans mailbox folders and Deleted Items to remove expired tagged items.
'  - Keeps Outlook COM access on the UI thread via `SwitchToUi`.
'
' Retention Model:
'  - `AutoDeleteAfterHours = 0` disables cleanup.
'  - When enabled, the original mail and all AutoPilot replies in the same group
'    are eligible for deletion after the configured retention window.
'  - Cleanup includes normal folders and Deleted Items.
'
' Security & Reliability:
'  - Deletion is based on hidden MAPI metadata, not EntryID or subject text.
'  - The cleanup process is best-effort and fail-safe.
'  - COM objects are released after use to avoid Outlook resource leaks.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Threading.Tasks
Imports Microsoft.Office.Interop.Outlook

Partial Public Class ThisAddIn

    Private Const AP_CleanupGroupIdProperty As String =
        "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/X-RedInk-AutoDeleteGroupId"
    Private Const AP_CleanupEligibleProperty As String =
        "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/X-RedInk-AutoDeleteEligible"
    Private Const AP_CleanupAnsweredUtcProperty As String =
        "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/X-RedInk-AutoDeleteAnsweredUtc"
    Private Const AP_CleanupDeleteAfterUtcProperty As String =
        "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/X-RedInk-AutoDeleteAfterUtc"

    ' When True, an auto-delete cleanup pass runs immediately at AutoPilot startup.
    ' Default False: cleanup runs only on the periodic timer.
    Private Const AP_RunAutoDeleteCleanupOnStartup As Boolean = False

    Private Const AP_AutoDeleteTimerIntervalSeconds As Integer = 6 * 60 * 60
    Private Const AP_AutoDeleteUiYieldItemInterval As Integer = 100
    Private Const AP_AutoDeleteUiYieldFolderInterval As Integer = 10
    Private Const AP_AutoDeleteProgressItemInterval As Integer = 500
    Private Const AP_AutoDeleteProgressFolderInterval As Integer = 25
    Private Const AP_AutoDeleteDiagnosticRejectLogLimit As Integer = 20
    Private Const AP_AutoDeleteDiagnosticErrorLogLimit As Integer = 20

    Private _apAutoDeleteTimer As System.Threading.Timer = Nothing
    Private _apAutoDeleteCheckRunning As Integer = 0
    Private _apAutoDeleteDiagnosticRejectLogCount As Integer = 0
    Private _apAutoDeleteDiagnosticErrorLogCount As Integer = 0

    Private Structure AutoDeleteCleanupStats
        Public ScannedCount As Integer
        Public DeletedCount As Integer
        Public ErrorCount As Integer
        Public FolderCount As Integer
        Public LastProgressLogScannedCount As Integer
        Public LastProgressLogFolderCount As Integer
    End Structure

    Friend Sub StartAutoDeleteTimer()
        StopAutoDeleteTimer()

        If _apConfig Is Nothing OrElse _apConfig.AutoDeleteAfterHours <= 0 Then Return

        _apAutoDeleteTimer = New System.Threading.Timer(
            AddressOf AutoDeleteTimerCallback,
            Nothing,
            dueTime:=TimeSpan.FromSeconds(AP_AutoDeleteTimerIntervalSeconds),
            period:=TimeSpan.FromSeconds(AP_AutoDeleteTimerIntervalSeconds))

        ApDashboardLog($"🗑 Auto-delete timer started ({_apConfig.AutoDeleteAfterHours}h retention, including Deleted Items).", "info")
    End Sub

    Friend Sub StopAutoDeleteTimer()
        Try : _apAutoDeleteTimer?.Dispose() : Catch : End Try
        _apAutoDeleteTimer = Nothing
    End Sub

    Private Async Sub AutoDeleteTimerCallback(state As Object)
        If Not _apActive Then Return

        Try
            Dim ct = _apCts?.Token
            If ct Is Nothing OrElse ct.Value.IsCancellationRequested Then Return
            Await RunAutoDeleteCleanupAsync(ct.Value, "timer")
        Catch ex As OperationCanceledException
            ' Expected during shutdown
        Catch ex As System.Exception
            ApDashboardLog($"🗑 Auto-delete timer error: {ex.Message}", "warn")
        End Try
    End Sub

    Friend Async Function RunAutoDeleteCleanupAsync(ct As CancellationToken,
                                                    Optional trigger As String = "manual") As Task
        If _apConfig Is Nothing OrElse _apConfig.AutoDeleteAfterHours <= 0 Then Return

        If Interlocked.CompareExchange(_apAutoDeleteCheckRunning, 1, 0) <> 0 Then
            ApDashboardLog($"🗑 Auto-delete cleanup already running — skipped duplicate {trigger} request.", "step")
            Return
        End If

        Dim nowUtc = DateTime.UtcNow
        Dim pass1 As AutoDeleteCleanupStats
        Dim pass2 As AutoDeleteCleanupStats
        Dim sw = Stopwatch.StartNew()

        '_apAutoDeleteDiagnosticRejectLogCount = 0
        '_apAutoDeleteDiagnosticErrorLogCount = 0

        Try
            ApDashboardLog($"🗑 Auto-delete cleanup started ({trigger}).", "info")

            ThrowIfAutoDeleteCancelled(ct)
            Await SwitchToUi(
                Sub()
                    pass1 = DeleteExpiredAutoPilotItemsOutsideDeletedItems(ct, nowUtc)
                End Sub)

            ApDashboardLog(
                $"🗑 Auto-delete mailbox folders pass: {pass1.DeletedCount} item(s) removed, {pass1.ErrorCount} error(s), {pass1.ScannedCount} candidate item(s) matched across {pass1.FolderCount} folder(s).",
                If(pass1.ErrorCount > 0, "warn", "step"))

            ThrowIfAutoDeleteCancelled(ct)
            Await SwitchToUi(
                Sub()
                    pass2 = DeleteExpiredAutoPilotItemsInsideDeletedItems(ct, nowUtc)
                End Sub)

            ApDashboardLog(
                $"🗑 Auto-delete Deleted Items pass: {pass2.DeletedCount} item(s) removed, {pass2.ErrorCount} error(s), {pass2.ScannedCount} candidate item(s) matched across {pass2.FolderCount} folder(s).",
                If(pass2.ErrorCount > 0, "warn", "step"))

            Dim totalDeleted = pass1.DeletedCount + pass2.DeletedCount
            Dim totalErrors = pass1.ErrorCount + pass2.ErrorCount
            Dim totalScanned = pass1.ScannedCount + pass2.ScannedCount
            Dim totalFolders = pass1.FolderCount + pass2.FolderCount

            If totalScanned = 0 Then
                ApDashboardLog(
                    "🗑 Auto-delete found no eligible mails with cleanup metadata. Existing mails in Inbox, Sent Items, or Deleted Items are ignored unless they were previously stamped for auto-delete.",
                    "step")
            End If

            ApDashboardLog(
                $"🗑 Auto-delete cleanup finished in {sw.Elapsed.TotalSeconds:F1}s: {totalDeleted} item(s) removed, {totalErrors} error(s), {totalScanned} candidate item(s) matched across {totalFolders} folder(s).",
                If(totalErrors > 0, "warn", "info"))
        Catch ex As OperationCanceledException
            ApDashboardLog("🗑 Auto-delete cleanup cancelled.", "step")
        Catch ex As System.Exception
            ApDashboardLog($"🗑 Auto-delete cleanup error: {ex.Message}", "warn")
        Finally
            sw.Stop()
            Interlocked.Exchange(_apAutoDeleteCheckRunning, 0)
        End Try
    End Function

    Private Sub ThrowIfAutoDeleteCancelled(ct As CancellationToken)
        If ct.IsCancellationRequested Then
            Throw New OperationCanceledException(ct)
        End If
    End Sub

    Private Sub LogAutoDeleteRejectDiagnostic(message As String)
        'If _apAutoDeleteDiagnosticRejectLogCount >= AP_AutoDeleteDiagnosticRejectLogLimit Then Return

        '_apAutoDeleteDiagnosticRejectLogCount += 1
        'ApDashboardLog($"🗑 Auto-delete diagnostic: {message}", "step")
    End Sub

    Private Sub LogAutoDeleteErrorDiagnostic(message As String)
        'If _apAutoDeleteDiagnosticErrorLogCount >= AP_AutoDeleteDiagnosticErrorLogLimit Then Return

        '_apAutoDeleteDiagnosticErrorLogCount += 1
        'ApDashboardLog($"🗑 Auto-delete error diagnostic: {message}", "warn")
    End Sub

    Private Function GetAutoDeleteMailDiagnosticLabel(mi As MailItem, folderPath As String) As String
        Dim subjectText As String = ""
        Dim messageClassText As String = ""

        If mi IsNot Nothing Then
            Try
                subjectText = If(mi.Subject, "")
            Catch
            End Try

            Try
                messageClassText = If(mi.MessageClass, "")
            Catch
            End Try
        End If

        subjectText = subjectText.Replace(vbCr, " ").Replace(vbLf, " ").Trim()

        Return $"Folder='{folderPath}', Subject='{subjectText}', MessageClass='{messageClassText}'"
    End Function

    Private Sub PumpAutoDeleteUi(ByRef stats As AutoDeleteCleanupStats,
                                 ct As CancellationToken,
                                 currentScope As String,
                                 Optional force As Boolean = False)

        ThrowIfAutoDeleteCancelled(ct)

        Dim shouldLog As Boolean = force

        If Not shouldLog Then
            If stats.ScannedCount > 0 AndAlso
               stats.ScannedCount - stats.LastProgressLogScannedCount >= AP_AutoDeleteProgressItemInterval Then
                shouldLog = True
            ElseIf stats.FolderCount > 0 AndAlso
                   stats.FolderCount Mod AP_AutoDeleteProgressFolderInterval = 0 AndAlso
                   stats.FolderCount <> stats.LastProgressLogFolderCount Then
                shouldLog = True
            End If
        End If

        If shouldLog Then
            stats.LastProgressLogScannedCount = stats.ScannedCount
            stats.LastProgressLogFolderCount = stats.FolderCount

            Dim scopeLabel = If(String.IsNullOrWhiteSpace(currentScope), "(unknown folder)", currentScope)

            ApDashboardLog(
                $"🗑 Auto-delete progress: {stats.DeletedCount} item(s) removed, {stats.ErrorCount} error(s), {stats.ScannedCount} item(s) scanned across {stats.FolderCount} folder(s). Current: {scopeLabel}",
                "step")
        End If

        If force OrElse
           (stats.ScannedCount > 0 AndAlso stats.ScannedCount Mod AP_AutoDeleteUiYieldItemInterval = 0) OrElse
           (stats.FolderCount > 0 AndAlso stats.FolderCount Mod AP_AutoDeleteUiYieldFolderInterval = 0) Then
            Try
                System.Windows.Forms.Application.DoEvents()
            Catch
            End Try

            ThrowIfAutoDeleteCancelled(ct)
        End If
    End Sub

    Private Function GetAutoDeleteCutoffUtc() As DateTime?
        If _apConfig Is Nothing OrElse _apConfig.AutoDeleteAfterHours <= 0 Then Return Nothing
        Return DateTime.UtcNow.AddHours(_apConfig.AutoDeleteAfterHours)
    End Function

    Private Function TryGetAutoDeleteTargetStoreId(session As Microsoft.Office.Interop.Outlook.NameSpace,
                                                   ByRef storeId As String) As Boolean
        storeId = ""

        If session Is Nothing Then Return False
        If _apConfig Is Nothing OrElse String.IsNullOrWhiteSpace(_apConfig.MonitoredMailbox) Then Return True

        For i As Integer = 1 To session.Accounts.Count
            Dim acct As Account = Nothing
            Dim deliveryStore As Store = Nothing

            Try
                acct = session.Accounts(i)
                If acct Is Nothing Then Continue For
                If String.IsNullOrWhiteSpace(acct.SmtpAddress) Then Continue For
                If Not acct.SmtpAddress.Equals(_apConfig.MonitoredMailbox, StringComparison.OrdinalIgnoreCase) Then Continue For

                deliveryStore = acct.DeliveryStore
                If deliveryStore Is Nothing Then Return False

                storeId = If(deliveryStore.StoreID, "")
                Return Not String.IsNullOrWhiteSpace(storeId)
            Catch
            Finally
                If deliveryStore IsNot Nothing Then Try : Marshal.ReleaseComObject(deliveryStore) : Catch : End Try
                If acct IsNot Nothing Then Try : Marshal.ReleaseComObject(acct) : Catch : End Try
            End Try
        Next

        Return False
    End Function

    Private Function ShouldProcessAutoDeleteStore(store As Store,
                                                  targetStoreId As String) As Boolean
        If store Is Nothing Then Return False
        If String.IsNullOrWhiteSpace(targetStoreId) Then Return True

        Try
            Return store.StoreID.Equals(targetStoreId, StringComparison.OrdinalIgnoreCase)
        Catch
            Return False
        End Try
    End Function

    Private Function GetCleanupGroupId(mi As MailItem) As String
        If mi Is Nothing Then Return Nothing

        Try
            Dim value = CStr(mi.PropertyAccessor.GetProperty(AP_CleanupGroupIdProperty))
            If Not String.IsNullOrWhiteSpace(value) Then Return value.Trim()
        Catch
        End Try

        Return Nothing
    End Function

    Friend Function GetOrCreateCleanupGroupId(mi As MailItem) As String
        If mi Is Nothing Then Return Nothing
        If _apConfig Is Nothing OrElse _apConfig.AutoDeleteAfterHours <= 0 Then Return Nothing

        Dim existing = GetCleanupGroupId(mi)
        If Not String.IsNullOrWhiteSpace(existing) Then Return existing

        Dim groupId = Guid.NewGuid().ToString("N")

        Try
            mi.PropertyAccessor.SetProperty(AP_CleanupGroupIdProperty, groupId)
            mi.Save()
        Catch ex As System.Exception
            Debug.WriteLine($"[AutoPilot] Failed to create cleanup group ID: {ex.Message}")
        End Try

        Return groupId
    End Function

    Friend Sub StampCleanupMetadata(mi As MailItem,
                                    groupId As String,
                                    isEligible As Boolean,
                                    answeredUtc As DateTime?,
                                    deleteAfterUtc As DateTime?,
                                    Optional saveItem As Boolean = True)

        If mi Is Nothing OrElse String.IsNullOrWhiteSpace(groupId) Then Return

        Try
            Dim pa = mi.PropertyAccessor
            pa.SetProperty(AP_CleanupGroupIdProperty, groupId)
            pa.SetProperty(AP_CleanupEligibleProperty, If(isEligible, "true", "false"))
            pa.SetProperty(
                AP_CleanupAnsweredUtcProperty,
                If(answeredUtc.HasValue,
                   answeredUtc.Value.ToString("o", CultureInfo.InvariantCulture),
                   ""))
            pa.SetProperty(
                AP_CleanupDeleteAfterUtcProperty,
                If(deleteAfterUtc.HasValue,
                   deleteAfterUtc.Value.ToString("o", CultureInfo.InvariantCulture),
                   ""))

            If saveItem Then mi.Save()
        Catch ex As System.Exception
            Debug.WriteLine($"[AutoPilot] StampCleanupMetadata error: {ex.Message}")
        End Try
    End Sub


    ''' <summary>
    ''' Schedules an auto-reply/out-of-office mail for auto-deletion by stamping it with
    ''' the cleanup group ID of its originating AutoPilot conversation. Such mails are not
    ''' processed, but should still be removed by the retention timer alongside the group.
    ''' </summary>
    Friend Sub TagAutoReplyOrOofForCleanup(oofMail As MailItem)
        If oofMail Is Nothing Then Return
        If _apConfig Is Nothing OrElse _apConfig.AutoDeleteAfterHours <= 0 Then Return

        Dim groupId = TryGetCleanupGroupIdFromConversation(oofMail)
        If String.IsNullOrWhiteSpace(groupId) Then Return

        Dim deleteAfterUtc = GetAutoDeleteCutoffUtc()
        If Not deleteAfterUtc.HasValue Then Return

        StampCleanupMetadata(
            oofMail,
            groupId,
            isEligible:=True,
            answeredUtc:=DateTime.UtcNow,
            deleteAfterUtc:=deleteAfterUtc,
            saveItem:=True)

        ApDashboardLog(
            $"🗑 Auto-delete scheduled for auto-reply/OOF in group {groupId} in {_apConfig.AutoDeleteAfterHours}h.",
            "step")
    End Sub

    ''' <summary>
    ''' Finds an existing cleanup group ID for a mail by inspecting the item itself and then
    ''' walking its conversation thread for a sibling that already carries a group ID.
    ''' </summary>
    Private Function TryGetCleanupGroupIdFromConversation(mi As MailItem) As String
        If mi Is Nothing Then Return Nothing

        ' 1. The item itself may already carry a group id.
        Dim ownGroupId = GetCleanupGroupId(mi)
        If Not String.IsNullOrWhiteSpace(ownGroupId) Then Return ownGroupId

        ' 2. Walk the conversation thread to find a sibling with a cleanup group id.
        Dim conversation As Conversation = Nothing
        Try
            conversation = mi.GetConversation()
        Catch
            ' GetConversation can fail on some store types
        End Try

        If conversation IsNot Nothing Then
            Try
                Dim rootItems As SimpleItems = conversation.GetRootItems()
                If rootItems IsNot Nothing Then
                    For Each rootItem As Object In rootItems
                        Dim found = FindCleanupGroupIdInConversationNode(conversation, rootItem, 0)
                        If Not String.IsNullOrWhiteSpace(found) Then Return found
                    Next
                End If
            Catch
                ' Conversation API can throw on certain store types
            End Try
        End If

        Return Nothing
    End Function

    ''' <summary>Recursively searches conversation tree nodes for an existing cleanup group ID.</summary>
    Private Function FindCleanupGroupIdInConversationNode(conv As Conversation,
                                                          item As Object,
                                                          depth As Integer) As String
        If depth > AP_MaxThreadDepth Then Return Nothing

        Try
            Dim nodeMail = TryCast(item, MailItem)
            If nodeMail IsNot Nothing Then
                Dim gid = GetCleanupGroupId(nodeMail)
                If Not String.IsNullOrWhiteSpace(gid) Then Return gid
            End If

            Dim children As SimpleItems = conv.GetChildren(item)
            If children IsNot Nothing Then
                For Each child As Object In children
                    Dim found = FindCleanupGroupIdInConversationNode(conv, child, depth + 1)
                    If Not String.IsNullOrWhiteSpace(found) Then Return found
                Next
            End If
        Catch
            ' Ignore individual node errors
        End Try

        Return Nothing
    End Function

    Friend Sub MarkMailGroupAsAnsweredAndEligible(originalMail As MailItem)
        If originalMail Is Nothing Then Return

        Dim deleteAfterUtc = GetAutoDeleteCutoffUtc()
        If Not deleteAfterUtc.HasValue Then Return

        Dim groupId = GetOrCreateCleanupGroupId(originalMail)
        If String.IsNullOrWhiteSpace(groupId) Then Return

        Dim answeredUtc = DateTime.UtcNow

        StampCleanupMetadata(
            originalMail,
            groupId,
            isEligible:=True,
            answeredUtc:=answeredUtc,
            deleteAfterUtc:=deleteAfterUtc,
            saveItem:=True)

        Dim stampedCount As Integer = 0
        ApplyEligibilityToGroupInAllStores(groupId, answeredUtc, deleteAfterUtc.Value, stampedCount)

        ApDashboardLog(
            $"🗑 Auto-delete scheduled for group {groupId} in {_apConfig.AutoDeleteAfterHours}h ({stampedCount} item(s) marked).",
            "step")
    End Sub

    Friend Sub MarkMailGroupRepliesAsAnsweredAndEligible(originalMail As MailItem)
        If originalMail Is Nothing Then Return

        Dim deleteAfterUtc = GetAutoDeleteCutoffUtc()
        If Not deleteAfterUtc.HasValue Then Return

        Dim groupId = GetOrCreateCleanupGroupId(originalMail)
        If String.IsNullOrWhiteSpace(groupId) Then Return

        Dim answeredUtc = DateTime.UtcNow
        Dim stampedCount As Integer = 0

        ApplyEligibilityToGroupInAllStores(groupId, answeredUtc, deleteAfterUtc.Value, stampedCount)

        ApDashboardLog(
            $"🗑 Auto-delete scheduled for sent replies in group {groupId} in {_apConfig.AutoDeleteAfterHours}h ({stampedCount} item(s) marked).",
            "step")
    End Sub

    Private Sub ApplyEligibilityToGroupInAllStores(groupId As String,
                                                   answeredUtc As DateTime,
                                                   deleteAfterUtc As DateTime,
                                                   ByRef stampedCount As Integer)

        Dim session As Microsoft.Office.Interop.Outlook.NameSpace = Nothing

        Try
            session = Application.GetNamespace("MAPI")

            Dim targetStoreId As String = ""
            If Not TryGetAutoDeleteTargetStoreId(session, targetStoreId) Then Return

            For i As Integer = 1 To session.Stores.Count
                Dim store As Store = Nothing
                Dim sentItems As MAPIFolder = Nothing
                Dim deletedItems As MAPIFolder = Nothing

                Try
                    store = session.Stores(i)
                    If Not ShouldProcessAutoDeleteStore(store, targetStoreId) Then Continue For

                    Try
                        sentItems = store.GetDefaultFolder(OlDefaultFolders.olFolderSentMail)
                        ApplyEligibilityToGroupInFolderTree(sentItems, groupId, answeredUtc, deleteAfterUtc, stampedCount)
                    Catch
                    Finally
                        If sentItems IsNot Nothing Then Try : Marshal.ReleaseComObject(sentItems) : Catch : End Try
                        sentItems = Nothing
                    End Try

                    Try
                        deletedItems = store.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems)
                        ApplyEligibilityToGroupInFolderTree(deletedItems, groupId, answeredUtc, deleteAfterUtc, stampedCount)
                    Catch
                    Finally
                        If deletedItems IsNot Nothing Then Try : Marshal.ReleaseComObject(deletedItems) : Catch : End Try
                        deletedItems = Nothing
                    End Try

                    Exit For
                Catch
                Finally
                    If store IsNot Nothing Then Try : Marshal.ReleaseComObject(store) : Catch : End Try
                End Try
            Next
        Catch ex As System.Exception
            Debug.WriteLine($"[AutoPilot] ApplyEligibilityToGroupInAllStores error: {ex.Message}")
        Finally
            If session IsNot Nothing Then Try : Marshal.ReleaseComObject(session) : Catch : End Try
        End Try
    End Sub

    Private Function IsAutoDeleteMailFolder(folder As MAPIFolder) As Boolean
        If folder Is Nothing Then Return False

        Try
            If folder.DefaultItemType <> OlItemType.olMailItem Then Return False
        Catch
            Return False
        End Try

        ' A folder that defaults to mail items can still report a blank or non-"IPM.Note"
        ' DefaultMessageClass. This is common for Deleted Items and for some Sent Items /
        ' IMAP folders. Rejecting those folders here previously caused the entire Deleted
        ' Items folder (the "Bin") to be skipped, so expired AutoPilot items were never
        ' propagated as eligible nor deleted. Each candidate item is still validated via
        ' TryCast to MailItem and by its cleanup metadata before deletion, so scanning
        ' these folders is safe. Only reject folders whose DefaultMessageClass clearly
        ' identifies a non-mail folder (appointments, contacts, tasks, notes, journals).
        Try
            Dim defaultMessageClass = If(folder.DefaultMessageClass, "")
            If String.IsNullOrWhiteSpace(defaultMessageClass) Then Return True
            If defaultMessageClass.StartsWith("IPM.Note", StringComparison.OrdinalIgnoreCase) Then Return True
            If defaultMessageClass.StartsWith("IPM.Post", StringComparison.OrdinalIgnoreCase) Then Return True

            Return False
        Catch
            Return True
        End Try
    End Function

    Private Sub ApplyEligibilityToGroupInFolderTree(folder As MAPIFolder,
                                                    groupId As String,
                                                    answeredUtc As DateTime,
                                                    deleteAfterUtc As DateTime,
                                                    ByRef stampedCount As Integer)

        If folder Is Nothing Then Return

        Dim items As Items = Nothing
        Dim subFolders As Folders = Nothing

        Try
            If IsAutoDeleteMailFolder(folder) Then
                items = folder.Items

                For i As Integer = items.Count To 1 Step -1
                    Dim obj As Object = Nothing
                    Dim mi As MailItem = Nothing

                    Try
                        obj = items(i)
                        mi = TryCast(obj, MailItem)
                        If mi Is Nothing Then Continue For

                        If Not String.Equals(GetCleanupGroupId(mi), groupId, StringComparison.OrdinalIgnoreCase) Then Continue For

                        StampCleanupMetadata(
                            mi,
                            groupId,
                            isEligible:=True,
                            answeredUtc:=answeredUtc,
                            deleteAfterUtc:=deleteAfterUtc,
                            saveItem:=True)

                        stampedCount += 1
                    Catch
                    Finally
                        If mi IsNot Nothing Then Try : Marshal.ReleaseComObject(mi) : Catch : End Try
                        If obj IsNot Nothing AndAlso Not ReferenceEquals(obj, mi) Then
                            Try : Marshal.ReleaseComObject(obj) : Catch : End Try
                        End If
                    End Try
                Next
            End If

            subFolders = folder.Folders
            For i As Integer = 1 To subFolders.Count
                Dim child As MAPIFolder = Nothing
                Try
                    child = subFolders(i)
                    ApplyEligibilityToGroupInFolderTree(child, groupId, answeredUtc, deleteAfterUtc, stampedCount)
                Catch
                Finally
                    If child IsNot Nothing Then Try : Marshal.ReleaseComObject(child) : Catch : End Try
                End Try
            Next
        Catch
        Finally
            If subFolders IsNot Nothing Then Try : Marshal.ReleaseComObject(subFolders) : Catch : End Try
            If items IsNot Nothing Then Try : Marshal.ReleaseComObject(items) : Catch : End Try
        End Try
    End Sub

    Private Function DeleteExpiredAutoPilotItemsOutsideDeletedItems(ct As CancellationToken,
                                                                    nowUtc As DateTime) As AutoDeleteCleanupStats
        Dim stats As New AutoDeleteCleanupStats()
        Dim session As Microsoft.Office.Interop.Outlook.NameSpace = Nothing

        Try
            session = Application.GetNamespace("MAPI")

            Dim targetStoreId As String = ""
            If Not TryGetAutoDeleteTargetStoreId(session, targetStoreId) Then Return stats

            For i As Integer = 1 To session.Stores.Count
                ThrowIfAutoDeleteCancelled(ct)

                Dim store As Store = Nothing
                Dim root As MAPIFolder = Nothing
                Dim deletedItems As MAPIFolder = Nothing
                Dim storeLabel As String = $"Store {i}"

                Try
                    store = session.Stores(i)
                    If Not ShouldProcessAutoDeleteStore(store, targetStoreId) Then Continue For

                    Try
                        storeLabel = If(store.DisplayName, storeLabel)
                    Catch
                    End Try

                    root = store.GetRootFolder()
                    Try
                        deletedItems = store.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems)
                    Catch
                        deletedItems = Nothing
                    End Try

                    ApDashboardLog($"🗑 Auto-delete scanning mailbox folders in {storeLabel}...", "step")

                    DeleteExpiredItemsOutsideDeletedItemsTree(
                        root,
                        If(deletedItems IsNot Nothing, deletedItems.EntryID, ""),
                        nowUtc,
                        stats,
                        ct)
                Catch ex As OperationCanceledException
                    Throw
                Catch
                Finally
                    If deletedItems IsNot Nothing Then Try : Marshal.ReleaseComObject(deletedItems) : Catch : End Try
                    If root IsNot Nothing Then Try : Marshal.ReleaseComObject(root) : Catch : End Try
                    If store IsNot Nothing Then Try : Marshal.ReleaseComObject(store) : Catch : End Try
                End Try
            Next
        Catch ex As OperationCanceledException
            Throw
        Catch ex As System.Exception
            Debug.WriteLine($"[AutoPilot] DeleteExpiredAutoPilotItemsOutsideDeletedItems error: {ex.Message}")
        Finally
            If session IsNot Nothing Then Try : Marshal.ReleaseComObject(session) : Catch : End Try
        End Try

        Return stats
    End Function

    Private Function DeleteExpiredAutoPilotItemsInsideDeletedItems(ct As CancellationToken,
                                                                   nowUtc As DateTime) As AutoDeleteCleanupStats
        Dim stats As New AutoDeleteCleanupStats()
        Dim session As Microsoft.Office.Interop.Outlook.NameSpace = Nothing

        Try
            session = Application.GetNamespace("MAPI")

            Dim targetStoreId As String = ""
            If Not TryGetAutoDeleteTargetStoreId(session, targetStoreId) Then Return stats

            For i As Integer = 1 To session.Stores.Count
                ThrowIfAutoDeleteCancelled(ct)

                Dim store As Store = Nothing
                Dim deletedItems As MAPIFolder = Nothing
                Dim storeLabel As String = $"Store {i}"

                Try
                    store = session.Stores(i)
                    If Not ShouldProcessAutoDeleteStore(store, targetStoreId) Then Continue For

                    Try
                        storeLabel = If(store.DisplayName, storeLabel)
                    Catch
                    End Try

                    deletedItems = store.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems)
                    ApDashboardLog($"🗑 Auto-delete scanning Deleted Items in {storeLabel}...", "step")
                    DeleteExpiredItemsInFolderTree(deletedItems, nowUtc, stats, ct)
                Catch ex As OperationCanceledException
                    Throw
                Catch
                Finally
                    If deletedItems IsNot Nothing Then Try : Marshal.ReleaseComObject(deletedItems) : Catch : End Try
                    If store IsNot Nothing Then Try : Marshal.ReleaseComObject(store) : Catch : End Try
                End Try
            Next
        Catch ex As OperationCanceledException
            Throw
        Catch ex As System.Exception
            Debug.WriteLine($"[AutoPilot] DeleteExpiredAutoPilotItemsInsideDeletedItems error: {ex.Message}")
        Finally
            If session IsNot Nothing Then Try : Marshal.ReleaseComObject(session) : Catch : End Try
        End Try

        Return stats
    End Function

    Private Sub DeleteExpiredItemsOutsideDeletedItemsTree(folder As MAPIFolder,
                                                          deletedItemsEntryId As String,
                                                          nowUtc As DateTime,
                                                          ByRef stats As AutoDeleteCleanupStats,
                                                          ct As CancellationToken)

        If folder Is Nothing Then Return

        ThrowIfAutoDeleteCancelled(ct)

        Dim folderPath As String = "(unknown folder)"
        Try
            folderPath = folder.FolderPath
        Catch
        End Try

        stats.FolderCount += 1
        PumpAutoDeleteUi(stats, ct, folderPath)

        Try
            If Not String.IsNullOrWhiteSpace(deletedItemsEntryId) AndAlso
               folder.EntryID.Equals(deletedItemsEntryId, StringComparison.OrdinalIgnoreCase) Then
                Return
            End If
        Catch
        End Try

        DeleteExpiredItemsInCurrentFolder(folder, nowUtc, stats, ct)

        Dim subFolders As Folders = Nothing
        Try
            subFolders = folder.Folders
            For i As Integer = 1 To subFolders.Count
                Dim child As MAPIFolder = Nothing
                Try
                    child = subFolders(i)
                    DeleteExpiredItemsOutsideDeletedItemsTree(child, deletedItemsEntryId, nowUtc, stats, ct)
                Catch ex As OperationCanceledException
                    Throw
                Catch
                Finally
                    If child IsNot Nothing Then Try : Marshal.ReleaseComObject(child) : Catch : End Try
                End Try
            Next
        Catch ex As OperationCanceledException
            Throw
        Catch
        Finally
            If subFolders IsNot Nothing Then Try : Marshal.ReleaseComObject(subFolders) : Catch : End Try
        End Try
    End Sub

    Private Sub DeleteExpiredItemsInFolderTree(folder As MAPIFolder,
                                               nowUtc As DateTime,
                                               ByRef stats As AutoDeleteCleanupStats,
                                               ct As CancellationToken)

        If folder Is Nothing Then Return

        ThrowIfAutoDeleteCancelled(ct)

        Dim folderPath As String = "(unknown folder)"
        Try
            folderPath = folder.FolderPath
        Catch
        End Try

        stats.FolderCount += 1
        PumpAutoDeleteUi(stats, ct, folderPath)

        DeleteExpiredItemsInCurrentFolder(folder, nowUtc, stats, ct)

        Dim subFolders As Folders = Nothing
        Try
            subFolders = folder.Folders
            For i As Integer = 1 To subFolders.Count
                Dim child As MAPIFolder = Nothing
                Try
                    child = subFolders(i)
                    DeleteExpiredItemsInFolderTree(child, nowUtc, stats, ct)
                Catch ex As OperationCanceledException
                    Throw
                Catch
                Finally
                    If child IsNot Nothing Then Try : Marshal.ReleaseComObject(child) : Catch : End Try
                End Try
            Next
        Catch ex As OperationCanceledException
            Throw
        Catch
        Finally
            If subFolders IsNot Nothing Then Try : Marshal.ReleaseComObject(subFolders) : Catch : End Try
        End Try
    End Sub

    Private Sub DeleteExpiredItemsInCurrentFolder(folder As MAPIFolder,
                                                  nowUtc As DateTime,
                                                  ByRef stats As AutoDeleteCleanupStats,
                                                  ct As CancellationToken)

        Dim items As Items = Nothing
        Dim folderPath As String = "(unknown folder)"
        Dim processedInFolder As Integer = 0

        Try
            Try
                folderPath = folder.FolderPath
            Catch
            End Try

            If Not IsAutoDeleteMailFolder(folder) Then Return

            items = folder.Items

            For i As Integer = items.Count To 1 Step -1
                Dim obj As Object = Nothing
                Dim mi As MailItem = Nothing

                Try
                    processedInFolder += 1
                    If processedInFolder Mod AP_AutoDeleteUiYieldItemInterval = 0 Then
                        PumpAutoDeleteUi(stats, ct, folderPath)
                    End If

                    obj = items(i)
                    mi = TryCast(obj, MailItem)
                    If mi Is Nothing Then Continue For

                    If Not ShouldAutoDeleteMail(mi, nowUtc, folderPath) Then Continue For

                    stats.ScannedCount += 1
                    PumpAutoDeleteUi(stats, ct, folderPath)

                    mi.Delete()
                    stats.DeletedCount += 1
                Catch ex As OperationCanceledException
                    Throw
                Catch ex As System.Exception
                    stats.ErrorCount += 1
                    LogAutoDeleteErrorDiagnostic($"Folder='{folderPath}', ItemIndex={i}: {ex.Message}")
                Finally
                    If mi IsNot Nothing Then Try : Marshal.ReleaseComObject(mi) : Catch : End Try
                    If obj IsNot Nothing AndAlso Not ReferenceEquals(obj, mi) Then
                        Try : Marshal.ReleaseComObject(obj) : Catch : End Try
                    End If
                End Try
            Next
        Catch ex As OperationCanceledException
            Throw
        Catch ex As System.Exception
            stats.ErrorCount += 1
            ApDashboardLog($"🗑 Auto-delete folder scan error: {folderPath} — {ex.Message}", "warn")
        Finally
            If items IsNot Nothing Then Try : Marshal.ReleaseComObject(items) : Catch : End Try
        End Try
    End Sub

    Private Function BuildAutoDeleteRestriction(nowUtc As DateTime) As String
        Dim eligibleProperty = ChrW(34) & AP_CleanupEligibleProperty & ChrW(34)
        Dim deleteAfterProperty = ChrW(34) & AP_CleanupDeleteAfterUtcProperty & ChrW(34)
        Dim nowUtcText = nowUtc.ToString("o", CultureInfo.InvariantCulture).Replace("'", "''")

        Return "@SQL=" &
            eligibleProperty & " = 'true' AND " &
            deleteAfterProperty & " <= '" & nowUtcText & "'"
    End Function

    Private Function ShouldAutoDeleteMail(mi As MailItem,
                                          nowUtc As DateTime,
                                          folderPath As String) As Boolean

        If mi Is Nothing Then Return False

        Dim label = GetAutoDeleteMailDiagnosticLabel(mi, folderPath)

        Dim groupId = GetCleanupGroupId(mi)
        If String.IsNullOrWhiteSpace(groupId) Then Return False

        Dim eligibleRaw As String = Nothing
        If Not TryGetNamedStringProperty(mi, AP_CleanupEligibleProperty, eligibleRaw) Then
            LogAutoDeleteRejectDiagnostic($"{label}, GroupId='{groupId}': missing eligible flag.")
            Return False
        End If

        If Not "true".Equals(eligibleRaw, StringComparison.OrdinalIgnoreCase) Then
            LogAutoDeleteRejectDiagnostic($"{label}, GroupId='{groupId}': eligible flag is '{eligibleRaw}'.")
            Return False
        End If

        Dim deleteAfterRaw As String = Nothing
        If Not TryGetNamedStringProperty(mi, AP_CleanupDeleteAfterUtcProperty, deleteAfterRaw) Then
            LogAutoDeleteRejectDiagnostic($"{label}, GroupId='{groupId}': missing delete-after timestamp.")
            Return False
        End If

        If String.IsNullOrWhiteSpace(deleteAfterRaw) Then
            LogAutoDeleteRejectDiagnostic($"{label}, GroupId='{groupId}': blank delete-after timestamp.")
            Return False
        End If

        Dim deleteAfterUtc As DateTime
        If Not TryParseCleanupUtcValue(deleteAfterRaw, deleteAfterUtc) Then
            LogAutoDeleteRejectDiagnostic($"{label}, GroupId='{groupId}': invalid delete-after timestamp '{deleteAfterRaw}'.")
            Return False
        End If

        If deleteAfterUtc > nowUtc Then
            LogAutoDeleteRejectDiagnostic(
                $"{label}, GroupId='{groupId}': delete-after {deleteAfterUtc.ToString("o", CultureInfo.InvariantCulture)} is later than now {nowUtc.ToString("o", CultureInfo.InvariantCulture)}.")
            Return False
        End If

        Return True
    End Function

    Private Function TryGetNamedStringProperty(mi As MailItem, propertyName As String, ByRef value As String) As Boolean
        value = Nothing
        If mi Is Nothing Then Return False

        Try
            Dim raw = mi.PropertyAccessor.GetProperty(propertyName)
            If raw Is Nothing Then Return False
            value = CStr(raw)
            Return True
        Catch
            Return False
        End Try
    End Function

    Private Function TryParseCleanupUtcValue(raw As String, ByRef value As DateTime) As Boolean
        value = DateTime.MinValue
        If String.IsNullOrWhiteSpace(raw) Then Return False

        Dim parsed As DateTime

        If DateTime.TryParseExact(
            raw,
            "o",
            CultureInfo.InvariantCulture,
            DateTimeStyles.RoundtripKind,
            parsed) Then

            value = parsed.ToUniversalTime()
            Return True
        End If

        If DateTime.TryParse(
            raw,
            CultureInfo.InvariantCulture,
            DateTimeStyles.AssumeUniversal Or DateTimeStyles.AdjustToUniversal,
            parsed) Then

            value = parsed.ToUniversalTime()
            Return True
        End If

        Return False
    End Function

    Private Function TryGetNamedDateUtcProperty(mi As MailItem, propertyName As String, ByRef value As DateTime) As Boolean
        value = DateTime.MinValue

        Dim raw As String = Nothing
        If Not TryGetNamedStringProperty(mi, propertyName, raw) Then Return False

        Return TryParseCleanupUtcValue(raw, value)
    End Function

End Class
