' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.
' For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Commands.InboxBoard.vb
' Purpose:
'   Interactive Inbox board (Kanban-style) for Outlook mails. Displays categorized
'   and optionally flagged messages as draggable cards, supports category/flag
'   updates via drag-and-drop, board-folder organization, and asynchronous AI
'   summaries.
'
' Architecture:
'  - Data collection:
'      * Scans default Inbox and builds board entries from MailItem metadata.
'      * Includes categorized mails; optionally includes flagged mails.
'      * Prompts for load cap when qualifying mail volume exceeds configurable threshold.
'  - Conversation mode:
'      * Optional grouping by ConversationTopic into a representative card with
'        aggregated message count and merged metadata.
'  - Column model:
'      * Category columns from Outlook category master list + pinned columns from settings.
'      * Optional flagged columns:
'          - single synthetic flagged column, or
'          - date-bucketed flag columns (Overdue/Today/Tomorrow/This Week/.../Done).
'  - Folder model:
'      * User-defined board folders (separate from Outlook folders) persisted in settings.
'      * Card assignment persisted per mail via UserProperty: `RedInkBoardFolder`.
'  - UI host:
'      * WinForms dialog with WebView2 rendering full HTML/CSS/JS board.
'      * Supports search, column filter, column reorder, field toggles, theme toggle,
'        sort per column, and card drag/drop gestures.
'  - JS ↔ VB bridge:
'      * WebView2 postMessage actions for move, moveAdd, unmark, open, reload,
'        settings persistence, folder operations, and card refresh.
'  - AI summaries:
'      * Background batch summarization using existing `LLM()` helper and
'        `SP_InboxBoard` prompt.
'      * EntryID-based summary cache persisted in `My.Settings` (bounded size).
'
' COM / Threading:
'  - Outlook COM operations run on UI thread (with `ComRetry` access pattern).
'  - LLM summary generation is asynchronous with cancellation on reload/close.
'
' Persistence (`My.Settings`):
'  - Window geometry/state, theme, card field toggles, column filter/order,
'    pinned category/flag columns, load threshold + last load count,
'    include-flagged/hide-done/grouping toggles, summary language,
'    summary cache, board folder list, expanded folder state.
' =============================================================================


Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.Threading
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Outlook
Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.WinForms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

#Region "InboxBoard Constants"

    ''' <summary>Number of mails per AI summary batch.</summary>
    Private Const InboxBoard_SummaryBatchSize As Integer = 5

    ''' <summary>Maximum body excerpt length sent for summarization.</summary>
    Private Const InboxBoard_SummaryBodyCap As Integer = 2000

    ''' <summary>Maximum cached summaries to persist in My.Settings.</summary>
    Private Const InboxBoard_SummaryCacheMax As Integer = 500

    ''' <summary>Column ID used for the synthetic "Flagged" column.</summary>
    Private Const InboxBoard_FlaggedColumnId As String = "⚑ Flagged"

    ''' <summary>Color used for the synthetic "Flagged" column.</summary>
    Private Const InboxBoard_FlaggedColumnColor As String = "#ef4444"

    ''' <summary>Default load threshold: prompt user when qualifying mails exceed this number.</summary>
    Private Const InboxBoard_DefaultLoadThreshold As Integer = 500

    ''' <summary>Remembers the user's last load-count choice so Refresh reuses it.</summary>
    Private _inboxBoardLastMaxToLoad As Integer = 0

    ''' <summary>In-memory summary cache: EntryID → summary text.</summary>
    Private _inboxBoardSummaryCache As Dictionary(Of String, String) = Nothing

    ''' <summary>MAPI UserProperty name for board folder assignment.</summary>
    Private Const InboxBoard_FolderPropertyName As String = "RedInkBoardFolder"

    ''' <summary>Maximum number of board folders a user can create.</summary>
    Private Const InboxBoard_MaxFolders As Integer = 20

    ''' <summary>Display label for the currently tracked folder (updated on each scan).</summary>
    Private _inboxBoardTrackedFolderLabel As String = "Inbox"


    ''' <summary>Column IDs for date-bucketed flagged columns.</summary>
    Private Const InboxBoard_FlagBucket_Overdue As String = "⚑ Overdue"
    Private Const InboxBoard_FlagBucket_Today As String = "⚑ Today"
    Private Const InboxBoard_FlagBucket_Tomorrow As String = "⚑ Tomorrow"
    Private Const InboxBoard_FlagBucket_ThisWeek As String = "⚑ This Week"
    Private Const InboxBoard_FlagBucket_NextWeek As String = "⚑ Next Week"
    Private Const InboxBoard_FlagBucket_Next4Weeks As String = "⚑ Next 4 Weeks"
    Private Const InboxBoard_FlagBucket_Later As String = "⚑ Later"
    Private Const InboxBoard_FlagBucket_NoDueDate As String = "⚑ Flagged"
    Private Const InboxBoard_FlagBucket_Done As String = "✓ Done"

    ''' <summary>Colors for each flag date bucket.</summary>
    Private Shared ReadOnly InboxBoard_FlagBucketColors As Dictionary(Of String, String) = New Dictionary(Of String, String)(StringComparer.Ordinal) From {
        {InboxBoard_FlagBucket_Overdue, "#dc2626"},
        {InboxBoard_FlagBucket_Today, "#ef4444"},
        {InboxBoard_FlagBucket_Tomorrow, "#f97316"},
        {InboxBoard_FlagBucket_ThisWeek, "#eab308"},
        {InboxBoard_FlagBucket_NextWeek, "#22c55e"},
        {InboxBoard_FlagBucket_Next4Weeks, "#3b82f6"},
        {InboxBoard_FlagBucket_Later, "#8b5cf6"},
        {InboxBoard_FlagBucket_NoDueDate, "#6b7280"},
        {InboxBoard_FlagBucket_Done, "#9ca3af"}
    }

    ''' <summary>Ordered list of all flag bucket column IDs (display order).</summary>
    Private Shared ReadOnly InboxBoard_FlagBucketOrder As String() = {
        InboxBoard_FlagBucket_Overdue,
        InboxBoard_FlagBucket_Today,
        InboxBoard_FlagBucket_Tomorrow,
        InboxBoard_FlagBucket_ThisWeek,
        InboxBoard_FlagBucket_NextWeek,
        InboxBoard_FlagBucket_Next4Weeks,
        InboxBoard_FlagBucket_Later,
        InboxBoard_FlagBucket_NoDueDate,
        InboxBoard_FlagBucket_Done
    }

#End Region

#Region "InboxBoard Data Classes"

    ''' <summary>Represents a categorized mail displayed on the board.</summary>
    Private Class InboxBoardEntry
        Public Property EntryID As String
        Public Property Subject As String
        Public Property SenderName As String
        Public Property SenderEmail As String
        Public Property Recipients As String
        Public Property ReceivedTime As DateTime
        Public Property IsRead As Boolean
        Public Property MessageCount As Integer
        Public Property Categories As String          ' First category name
        Public Property AllCategories As String       ' Semicolon-separated
        Public Property BodyExcerpt As String
        Public Property Summary As String             ' AI-generated, initially empty
        Public Property CategoryColor As String       ' Hex color for the category
        Public Property ConversationTopic As String   ' For conversation grouping
        Public Property ConversationEntryIDs As List(Of String) ' All EntryIDs in this conversation group
        Public Property IsGrouped As Boolean          ' True when this card represents a grouped conversation
        Public Property IsFlagged As Boolean          ' True when mail has FlagStatus = olFlagMarked
        Public Property IsCompleted As Boolean        ' True when mail has FlagStatus = olFlagComplete
        Public Property FlagRequest As String         ' Flag label text (e.g. "Follow up", "Reply")
        Public Property FlagDueDate As DateTime?      ' Optional due date from TaskDueDate
        Public Property FlagBucket As String          ' Which date-bucket column this card belongs to (if flagged)
        Public Property BoardFolder As String             ' Board folder name (from UserProperty), empty = not in a folder

    End Class

    ''' <summary>Represents a column (category) on the board.</summary>
    Private Class InboxBoardColumn
        Public Property CategoryName As String
        Public Property Color As String               ' Hex color
        Public Property Tag As String                 ' Display tag
        Public Property IsFlaggedColumn As Boolean    ' True for the synthetic "Flagged" column
    End Class

#End Region

#Region "InboxBoard Summary Cache"

    ''' <summary>
    ''' Loads the summary cache from My.Settings into the in-memory dictionary.
    ''' </summary>
    Private Sub LoadSummaryCache()
        If _inboxBoardSummaryCache IsNot Nothing Then Return
        _inboxBoardSummaryCache = New Dictionary(Of String, String)(StringComparer.Ordinal)
        Try
            Dim json As String = My.Settings.InboxBoardSummaryCache
            If Not String.IsNullOrEmpty(json) Then
                Dim obj As JObject = JObject.Parse(json)
                For Each prop In obj.Properties()
                    _inboxBoardSummaryCache(prop.Name) = CStr(prop.Value)
                Next
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Saves the in-memory summary cache to My.Settings, capping at InboxBoard_SummaryCacheMax entries.
    ''' </summary>
    Private Sub SaveSummaryCache()
        Try
            If _inboxBoardSummaryCache Is Nothing Then Return

            ' If over limit, remove oldest entries (by insertion order — Dictionary preserves order in .NET 4.x)
            While _inboxBoardSummaryCache.Count > InboxBoard_SummaryCacheMax
                Dim firstKey As String = _inboxBoardSummaryCache.Keys.First()
                _inboxBoardSummaryCache.Remove(firstKey)
            End While

            Dim obj As New JObject()
            For Each kvp In _inboxBoardSummaryCache
                obj(kvp.Key) = kvp.Value
            Next
            My.Settings.InboxBoardSummaryCache = obj.ToString(Formatting.None)
            My.Settings.Save()
        Catch
        End Try


    End Sub

    ''' <summary>
    ''' Looks up a cached summary for a given EntryID.
    ''' </summary>
    Private Function GetCachedSummary(entryId As String) As String
        LoadSummaryCache()
        Dim result As String = Nothing
        If _inboxBoardSummaryCache.TryGetValue(entryId, result) Then Return result
        Return Nothing
    End Function

    ''' <summary>
    ''' Stores a summary in the cache (in-memory + persistent).
    ''' </summary>
    Private Sub CacheSummary(entryId As String, summary As String)
        LoadSummaryCache()
        _inboxBoardSummaryCache(entryId) = summary
    End Sub

    ''' <summary>
    ''' Clears the entire summary cache (in-memory + persistent).
    ''' </summary>
    Private Sub ClearSummaryCache()
        _inboxBoardSummaryCache = New Dictionary(Of String, String)(StringComparer.Ordinal)
        Try
            My.Settings.InboxBoardSummaryCache = ""
            My.Settings.Save()
        Catch
        End Try
    End Sub

#End Region


#Region "InboxBoard Folder Operations"

    ''' <summary>
    ''' Reads the board-folder assignment from a mail item's user property.
    ''' </summary>
    ''' <returns>
    ''' Folder name if assigned; otherwise empty string.
    ''' </returns>
    Private Function GetMailBoardFolder(mi As MailItem) As String
        Try
            Dim prop As Outlook.UserProperty = Nothing
            Try
                prop = mi.UserProperties.Find(InboxBoard_FolderPropertyName)
            Catch ex As System.Exception
                Debug.WriteLine($"[InboxBoard] GetMailBoardFolder Find error: {ex.Message}")
            End Try
            If prop IsNot Nothing Then
                Dim val As String = CStr(prop.Value)
                Debug.WriteLine($"[InboxBoard] GetMailBoardFolder: Found '{val}' on '{mi.Subject}'")
                If Not String.IsNullOrWhiteSpace(val) Then Return val.Trim()
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"[InboxBoard] GetMailBoardFolder error: {ex.Message}")
        End Try
        Return ""
    End Function

    ''' <summary>
    ''' Sets or clears the board-folder assignment user property on a mail item.
    ''' </summary>
    ''' <param name="entryId">Target MailItem EntryID.</param>
    ''' <param name="folderName">
    ''' Folder name to assign; empty value clears assignment.
    ''' </param>
    Private Sub SetMailBoardFolder(entryId As String, folderName As String)
        Try
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
            Dim mi As MailItem = TryCast(ComRetry(Of Object)(Function() ns.GetItemFromID(entryId)), MailItem)
            If mi Is Nothing Then
                Debug.WriteLine($"[InboxBoard] SetMailBoardFolder: Could not resolve entryId {entryId}")
                Return
            End If

            ' Always remove any existing property first to avoid stale COM references
            Try
                Dim existing As Outlook.UserProperty = Nothing
                Try
                    existing = mi.UserProperties.Find(InboxBoard_FolderPropertyName)
                Catch ex As System.Exception
                    Debug.WriteLine($"[InboxBoard] SetMailBoardFolder Find error: {ex.Message}")
                End Try
                If existing IsNot Nothing Then
                    Debug.WriteLine($"[InboxBoard] SetMailBoardFolder: Deleting existing prop, old value = '{existing.Value}'")
                    existing.Delete()
                End If
            Catch ex As System.Exception
                Debug.WriteLine($"[InboxBoard] SetMailBoardFolder Delete error: {ex.Message}")
            End Try

            If Not String.IsNullOrWhiteSpace(folderName) Then
                ' Add a fresh property — use True for AddToFolderFields so Exchange
                ' cached mode stores persist the property reliably
                Dim prop As Outlook.UserProperty = mi.UserProperties.Add(
                    InboxBoard_FolderPropertyName, OlUserPropertyType.olText, True)
                prop.Value = folderName.Trim()
                Debug.WriteLine($"[InboxBoard] SetMailBoardFolder: Added prop '{InboxBoard_FolderPropertyName}' = '{folderName.Trim()}' on '{mi.Subject}'")
            Else
                Debug.WriteLine($"[InboxBoard] SetMailBoardFolder: Cleared prop on '{mi.Subject}'")
            End If

            mi.Save()
            Debug.WriteLine($"[InboxBoard] SetMailBoardFolder: Save() succeeded for '{mi.Subject}'")

            ' Verify the property was persisted by re-reading it
            Dim verify As Outlook.UserProperty = Nothing
            Try
                verify = mi.UserProperties.Find(InboxBoard_FolderPropertyName)
            Catch
            End Try
            If Not String.IsNullOrWhiteSpace(folderName) Then
                If verify IsNot Nothing Then
                    Debug.WriteLine($"[InboxBoard] SetMailBoardFolder: VERIFIED value = '{verify.Value}'")
                Else
                    Debug.WriteLine($"[InboxBoard] SetMailBoardFolder: *** VERIFY FAILED — property not found after Save! ***")
                End If
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"[InboxBoard] SetMailBoardFolder error: {ex.Message}")
            Debug.WriteLine($"[InboxBoard] SetMailBoardFolder stack: {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' Returns the list of user-defined board folder names from My.Settings.
    ''' </summary>
    Private Function GetBoardFolderNames() As List(Of String)
        Dim result As New List(Of String)()
        Try
            Dim saved As String = If(My.Settings.InboxBoardFolders, "")
            If Not String.IsNullOrEmpty(saved) Then
                For Each part In saved.Split({";"c}, StringSplitOptions.RemoveEmptyEntries)
                    Dim t = part.Trim()
                    If Not String.IsNullOrEmpty(t) AndAlso Not result.Any(Function(f) f.Equals(t, StringComparison.OrdinalIgnoreCase)) Then
                        result.Add(t)
                    End If
                Next
            End If
        Catch
        End Try
        Return result
    End Function

    ''' <summary>
    ''' Saves the board folder names list to My.Settings.
    ''' </summary>
    Private Sub SaveBoardFolderNames(names As List(Of String))
        Try
            My.Settings.InboxBoardFolders = String.Join(";", names)
            My.Settings.Save()
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Returns the set of currently expanded folder names from My.Settings.
    ''' </summary>
    Private Function GetExpandedFolders() As HashSet(Of String)
        Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Try
            Dim saved As String = If(My.Settings.InboxBoardExpandedFolders, "")
            If Not String.IsNullOrEmpty(saved) Then
                For Each part In saved.Split({";"c}, StringSplitOptions.RemoveEmptyEntries)
                    Dim t = part.Trim()
                    If Not String.IsNullOrEmpty(t) Then result.Add(t)
                Next
            End If
        Catch
        End Try
        Return result
    End Function

    ''' <summary>
    ''' Saves the expanded folder set to My.Settings.
    ''' </summary>
    Private Sub SaveExpandedFolders(expanded As HashSet(Of String))
        Try
            My.Settings.InboxBoardExpandedFolders = String.Join(";", expanded)
            My.Settings.Save()
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Renames a board folder: updates UserProperty on all affected mails
    ''' and updates My.Settings.
    ''' </summary>
    Private Sub RenameBoardFolder(oldName As String, newName As String, mails As List(Of InboxBoardEntry))
        If String.IsNullOrWhiteSpace(oldName) OrElse String.IsNullOrWhiteSpace(newName) Then Return
        If oldName.Equals(newName, StringComparison.OrdinalIgnoreCase) Then Return

        ' Update all mails in the folder
        For Each m In mails
            If Not String.IsNullOrEmpty(m.BoardFolder) AndAlso
               m.BoardFolder.Equals(oldName, StringComparison.OrdinalIgnoreCase) Then
                ' Update all entry IDs in the group
                Dim idsToUpdate As List(Of String) = If(m.IsGrouped AndAlso m.ConversationEntryIDs IsNot Nothing AndAlso m.ConversationEntryIDs.Count > 1,
                                                        m.ConversationEntryIDs,
                                                        New List(Of String) From {m.EntryID})
                For Each id In idsToUpdate
                    SetMailBoardFolder(id, newName)
                Next
                m.BoardFolder = newName
            End If
        Next

        ' Update settings
        Dim folders = GetBoardFolderNames()
        Dim idx = folders.FindIndex(Function(f) f.Equals(oldName, StringComparison.OrdinalIgnoreCase))
        If idx >= 0 Then
            folders(idx) = newName
        Else
            folders.Add(newName)
        End If
        SaveBoardFolderNames(folders)

        ' Update expanded state
        Dim expanded = GetExpandedFolders()
        If expanded.Remove(oldName) Then expanded.Add(newName)
        SaveExpandedFolders(expanded)
    End Sub

    ''' <summary>
    ''' Deletes a board folder: clears UserProperty on all affected mails
    ''' (releasing them back to the main board) and removes from My.Settings.
    ''' </summary>
    Private Sub DeleteBoardFolder(folderName As String, mails As List(Of InboxBoardEntry))
        If String.IsNullOrWhiteSpace(folderName) Then Return

        ' Clear folder assignment on all mails
        For Each m In mails
            If Not String.IsNullOrEmpty(m.BoardFolder) AndAlso
               m.BoardFolder.Equals(folderName, StringComparison.OrdinalIgnoreCase) Then
                Dim idsToUpdate As List(Of String) = If(m.IsGrouped AndAlso m.ConversationEntryIDs IsNot Nothing AndAlso m.ConversationEntryIDs.Count > 1,
                                                        m.ConversationEntryIDs,
                                                        New List(Of String) From {m.EntryID})
                For Each id In idsToUpdate
                    SetMailBoardFolder(id, "")
                Next
                m.BoardFolder = ""
            End If
        Next

        ' Remove from settings
        Dim folders = GetBoardFolderNames()
        folders.RemoveAll(Function(f) f.Equals(folderName, StringComparison.OrdinalIgnoreCase))
        SaveBoardFolderNames(folders)

        ' Remove from expanded
        Dim expanded = GetExpandedFolders()
        expanded.Remove(folderName)
        SaveExpandedFolders(expanded)
    End Sub

#End Region

#Region "InboxBoard Entry Point"

    ''' <summary>
    ''' Main entry point for the Inbox Board feature.
    ''' </summary>
    Public Sub InboxBoard()
        Try
            Using progress As New InboxBoardProgressScope(
                "Inbox Board",
                "Preparing inbox board…",
                1)

                ReportInboxBoardProgress(progress, 0, 1, "Preparing inbox board…")
                If progress.CancelRequested Then Return

                ' _inboxBoardLastMaxToLoad starts at 0 each session so the user is
                ' always prompted on first launch when mails exceed the threshold.
                ' It is only set after the user makes a selection, allowing in-session
                ' reloads (Refresh button) to reuse the choice without re-prompting.

                ReportInboxBoardProgress(progress, 0, 1, "Loading summary cache…")
                LoadSummaryCache()
                If progress.CancelRequested Then Return

                Dim mails As List(Of InboxBoardEntry) = CollectCategorizedInboxMails(0, progress)
                If mails Is Nothing Then Return
                If progress.CancelRequested Then Return

                ReportInboxBoardProgress(progress, 1, 1, "Building board columns…")
                Dim columns As List(Of InboxBoardColumn) = BuildBoardColumns(mails)
                If columns Is Nothing Then columns = New List(Of InboxBoardColumn)()

                If progress.CancelRequested Then Return

                ShowInboxBoardForm(mails, columns)
            End Using

        Catch ex As System.Exception
            ShowCustomMessageBox($"Inbox Board error: {ex.Message}", $"{AN} - Inbox Board")
        End Try
    End Sub

#End Region

#Region "InboxBoard Mail Collection"

    Private NotInheritable Class InboxBoardProgressScope
        Implements IDisposable

        Private ReadOnly _cts As New CancellationTokenSource()
        Private ReadOnly _ready As New ManualResetEventSlim(False)
        Private _uiThread As Thread
        Private _form As Form
        Private _closed As Integer = 0

        Public Sub New(headerText As String,
                       initialLabel As String,
                       Optional max As Integer = 100)

            ProgressBarModule.CancelOperation = False
            ProgressBarModule.GlobalProgressMax = Math.Max(1, max)
            ProgressBarModule.GlobalProgressValue = 0
            ProgressBarModule.GlobalProgressLabel = If(initialLabel, "")

            _uiThread = New Thread(
                Sub()
                    Try
                        System.Windows.Forms.Application.EnableVisualStyles()

                        Dim frm As System.Windows.Forms.Form =
                            CType(New Global.SharedLibrary.SharedLibrary.DPIProgressForm(headerText, initialLabel), System.Windows.Forms.Form)

                        _form = frm

                        AddHandler frm.Shown,
                            Sub()
                                Try
                                    _ready.Set()
                                Catch
                                End Try
                            End Sub

                        System.Windows.Forms.Application.Run(frm)
                    Catch
                        Try
                            _ready.Set()
                        Catch
                        End Try
                    End Try
                End Sub)

            _uiThread.IsBackground = True
            _uiThread.SetApartmentState(ApartmentState.STA)
            _uiThread.Start()

            Try
                _ready.Wait(750)
            Catch
            End Try
        End Sub

        Public Shared Sub Report(current As Integer,
                                 Optional max As Integer = -1,
                                 Optional label As String = Nothing)
            If max >= 1 Then ProgressBarModule.GlobalProgressMax = max
            If label IsNot Nothing Then ProgressBarModule.GlobalProgressLabel = label
            ProgressBarModule.GlobalProgressValue =
                Math.Max(0, Math.Min(current, ProgressBarModule.GlobalProgressMax))
        End Sub

        Public ReadOnly Property CancelRequested As Boolean
            Get
                Return ProgressBarModule.CancelOperation OrElse _cts.IsCancellationRequested
            End Get
        End Property

        Public Sub Dispose() Implements IDisposable.Dispose
            If Interlocked.Exchange(_closed, 1) <> 0 Then Return

            Try
                ProgressBarModule.CancelOperation = True

                Dim f = _form
                If f IsNot Nothing Then
                    Try
                        If f.IsHandleCreated AndAlso Not f.IsDisposed Then
                            f.BeginInvoke(New System.Action(
                                Sub()
                                    Try
                                        If Not f.IsDisposed Then f.Close()
                                    Catch
                                    End Try
                                End Sub))
                        End If
                    Catch
                    End Try
                End If
            Finally
                Try
                    If _uiThread IsNot Nothing AndAlso _uiThread.IsAlive Then
                        If Not _uiThread.Join(1000) Then
                            Try : _uiThread.Interrupt() : Catch : End Try
                        End If
                    End If
                Catch
                End Try

                Try
                    _ready.Dispose()
                Catch
                End Try
            End Try
        End Sub
    End Class


    Private Class InboxBoardScanProgressState
        Public Property ExpectedItems As Integer
        Public Property ProcessedItems As Integer
    End Class

    Private Shared Function ShouldReportInboxBoardProgress(current As Integer, total As Integer) As Boolean
        If total <= 0 Then Return False
        If current <= 1 OrElse current >= total Then Return True
        If total <= 100 Then Return True
        Return (current Mod 25) = 0
    End Function

    Private Shared Sub ReportInboxBoardProgress(progress As InboxBoardProgressScope,
                                                current As Integer,
                                                max As Integer,
                                                label As String)
        If progress Is Nothing Then Return
        InboxBoardProgressScope.Report(current, Math.Max(1, max), label)
    End Sub

    ''' <summary>
    ''' Scans the configured target folder (defaulting to default Inbox) for mails that
    ''' have at least one category assigned, and optionally mails that are flagged for follow-up.
    ''' If qualifying mails exceed the user's load threshold, asks how many to load.
    ''' When "include subfolders" is enabled, recursively scans child folders as well.
    ''' </summary>
    Private Function CollectCategorizedInboxMails(Optional maxOverride As Integer = 0,
                                                  Optional progress As InboxBoardProgressScope = Nothing) As List(Of InboxBoardEntry)

        Dim ownsProgress As Boolean = (progress Is Nothing)
        If ownsProgress Then
            progress = New InboxBoardProgressScope(
                "Inbox Board",
                "Preparing inbox board…",
                1)
        End If

        Try
            ReportInboxBoardProgress(progress, 0, 1, "Resolving target folder…")

            ' Resolve target folder from My.Settings (StoreId + FolderPath)
            Dim targetFolder As MAPIFolder = GetInboxBoardTargetFolder()

            If progress IsNot Nothing AndAlso progress.CancelRequested Then Return Nothing

            If targetFolder Is Nothing Then
                ShowCustomMessageBox("Could not access the target folder.", $"{AN} - Inbox Board")
                Return Nothing
            End If

            ' Read subfolder setting
            Dim includeSubfolders As Boolean = True
            Try : includeSubfolders = My.Settings.InboxBoardIncludeSubfolders : Catch : End Try

            ' Read flag-related settings
            Dim includeFlagged As Boolean = False
            Dim hideDoneFlags As Boolean = True
            Try : includeFlagged = My.Settings.InboxBoardIncludeFlagged : Catch : End Try
            Try : hideDoneFlags = My.Settings.InboxBoardHideDoneFlags : Catch : End Try

            ' Read user-configurable load threshold (0 = use default)
            Dim loadThreshold As Integer = 0
            Try : loadThreshold = My.Settings.InboxBoardLoadThreshold : Catch : End Try
            If loadThreshold <= 0 Then loadThreshold = InboxBoard_DefaultLoadThreshold

            Dim allMailItems As List(Of MailItem) = Nothing
            Dim totalQualifying As Integer = 0

            ' Stage 1: scan mailbox + count qualifying mails
            ReportInboxBoardProgress(
                progress,
                0,
                1,
                $"Scanning ""{_inboxBoardTrackedFolderLabel}""…")

            Dim scanState As New InboxBoardScanProgressState()
            allMailItems = CollectMailItemsFromFolder(targetFolder, includeSubfolders, progress, scanState)
            If progress.CancelRequested Then Return Nothing

            Dim scanStageMax As Integer = Math.Max(1, allMailItems.Count * 2)

            ReportInboxBoardProgress(
                progress,
                allMailItems.Count,
                scanStageMax,
                $"Checking {allMailItems.Count:N0} mails…")

            Dim countIndex As Integer = 0
            For Each mi In allMailItems
                If progress.CancelRequested Then Return Nothing

                countIndex += 1

                Try
                    Dim cats As String = ""
                    Try : cats = ComRetry(Of String)(Function() If(mi.Categories, "")) : Catch : End Try
                    Dim hasCats As Boolean = Not String.IsNullOrWhiteSpace(cats)

                    Dim qualifies As Boolean = hasCats
                    If Not qualifies Then
                        Dim flagStatus As Integer = 0
                        Try : flagStatus = ComRetry(Of Integer)(Function() CInt(mi.FlagStatus)) : Catch : End Try
                        If flagStatus = 2 Then
                            qualifies = True
                        ElseIf flagStatus = 1 AndAlso Not hideDoneFlags Then
                            qualifies = True
                        End If
                    End If

                    If qualifies Then totalQualifying += 1
                Catch
                End Try

                If ShouldReportInboxBoardProgress(countIndex, allMailItems.Count) Then
                    ReportInboxBoardProgress(
                        progress,
                        allMailItems.Count + countIndex,
                        scanStageMax,
                        $"Checking qualifying mails… {countIndex:N0}/{allMailItems.Count:N0}")
                End If
            Next

            If totalQualifying = 0 Then
                Dim hasPinnedColumns As Boolean = False
                Try
                    Dim pinned As String = If(My.Settings.InboxBoardColumns, "")
                    If Not String.IsNullOrEmpty(pinned) Then hasPinnedColumns = True
                Catch
                End Try

                Dim hasPinnedFlagColumns As Boolean = False
                Try
                    Dim pinnedFlags As String = If(My.Settings.InboxBoardPinnedFlagColumns, "")
                    If Not String.IsNullOrEmpty(pinnedFlags) Then hasPinnedFlagColumns = True
                Catch
                End Try

                If Not hasPinnedColumns AndAlso Not hasPinnedFlagColumns Then
                    ShowCustomMessageBox($"No categorized or flagged mails found in ""{_inboxBoardTrackedFolderLabel}"".", $"{AN} - Inbox Board")
                    Return Nothing
                End If

                Return New List(Of InboxBoardEntry)()
            End If

            Dim maxToLoad As Integer = totalQualifying
            If maxOverride > 0 Then maxToLoad = Math.Min(maxOverride, totalQualifying)

            ' Prompt user if qualifying mails exceed the threshold.
            ' Keep the existing behavior: reuse the remembered choice if available.
            If maxOverride <= 0 AndAlso totalQualifying > loadThreshold Then
                If _inboxBoardLastMaxToLoad > 0 Then
                    maxToLoad = _inboxBoardLastMaxToLoad
                Else
                    Dim items As New List(Of SelectionItem)()
                    items.Add(New SelectionItem($"{loadThreshold} mails", loadThreshold))
                    If loadThreshold * 2 < totalQualifying Then
                        items.Add(New SelectionItem($"{loadThreshold * 2} mails", loadThreshold * 2))
                    End If
                    If loadThreshold * 5 < totalQualifying Then
                        items.Add(New SelectionItem($"{loadThreshold * 5} mails", loadThreshold * 5))
                    End If
                    If loadThreshold * 10 < totalQualifying Then
                        items.Add(New SelectionItem($"{loadThreshold * 10} mails", loadThreshold * 10))
                    End If
                    items.Add(New SelectionItem($"All ({totalQualifying} mails)", totalQualifying))

                    Dim chosen As Integer = SelectValue(items, loadThreshold,
                        $"Found {totalQualifying} qualifying mails in ""{_inboxBoardTrackedFolderLabel}"" (threshold: {loadThreshold})." & vbCrLf &
                        "How many should be loaded?",
                        $"{AN} - Inbox Board")
                    If chosen = 0 Then Return Nothing
                    maxToLoad = chosen
                End If
            End If

            _inboxBoardLastMaxToLoad = maxToLoad
            Try
                My.Settings.InboxBoardLastLoadCount = maxToLoad
                My.Settings.Save()
            Catch
            End Try

            ' Stage 2: build board entries
            Dim buildStageMax As Integer = Math.Max(1, allMailItems.Count + 4)

            ReportInboxBoardProgress(progress, 0, buildStageMax, "Loading Outlook categories…")

            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
            Dim categoryColors As Dictionary(Of String, String) = GetOutlookCategoryColors(ns)

            Dim entries As New List(Of InboxBoardEntry)()
            Dim extractIndex As Integer = 0

            For Each mi In allMailItems
                If progress.CancelRequested Then Return Nothing

                extractIndex += 1

                Try
                    Dim entry As InboxBoardEntry = BuildInboxBoardEntry(mi, categoryColors, includeFlagged, hideDoneFlags)
                    If entry IsNot Nothing Then entries.Add(entry)
                Catch
                End Try

                If ShouldReportInboxBoardProgress(extractIndex, allMailItems.Count) Then
                    ReportInboxBoardProgress(
                        progress,
                        extractIndex,
                        buildStageMax,
                        $"Building inbox board… {extractIndex:N0}/{allMailItems.Count:N0}")
                End If
            Next

            If progress.CancelRequested Then Return Nothing

            ReportInboxBoardProgress(progress, allMailItems.Count + 1, buildStageMax, "Sorting mails…")
            entries.Sort(Function(a, b) b.ReceivedTime.CompareTo(a.ReceivedTime))

            If entries.Count > maxToLoad Then
                entries.RemoveRange(maxToLoad, entries.Count - maxToLoad)
            End If

            If progress.CancelRequested Then Return Nothing

            ReportInboxBoardProgress(progress, allMailItems.Count + 2, buildStageMax, "Grouping conversations…")

            Dim groupConversations As Boolean = False
            Try : groupConversations = My.Settings.InboxBoardGroupConversations : Catch : End Try
            If groupConversations Then
                entries = GroupByConversation(entries)
            End If

            If progress.CancelRequested Then Return Nothing

            ReportInboxBoardProgress(progress, allMailItems.Count + 3, buildStageMax, "Applying cached summaries…")

            For Each entry In entries
                If progress.CancelRequested Then Return Nothing

                Dim cached As String = GetCachedSummary(entry.EntryID)
                If Not String.IsNullOrEmpty(cached) Then
                    entry.Summary = cached
                End If
            Next

            If progress.CancelRequested Then Return Nothing

            ReportInboxBoardProgress(progress, allMailItems.Count + 4, buildStageMax, "Finalizing inbox board…")

            Dim knownFolders = GetBoardFolderNames()
            Dim foldersChanged As Boolean = False
            For Each entry In entries
                If progress.CancelRequested Then Return Nothing

                If Not String.IsNullOrEmpty(entry.BoardFolder) Then
                    If Not knownFolders.Any(Function(f) f.Equals(entry.BoardFolder, StringComparison.OrdinalIgnoreCase)) Then
                        knownFolders.Add(entry.BoardFolder)
                        foldersChanged = True
                    End If
                End If
            Next
            If foldersChanged Then SaveBoardFolderNames(knownFolders)

            Return entries

        Finally
            If ownsProgress AndAlso progress IsNot Nothing Then
                progress.Dispose()
            End If
        End Try
    End Function

    ''' <summary>
    ''' Groups entries by ConversationTopic. For each conversation with 2+ mails,
    ''' keeps the latest mail as the representative card. Single-topic and no-topic
    ''' entries are preserved as individual ungrouped cards.
    ''' </summary>
    Private Function GroupByConversation(entries As List(Of InboxBoardEntry)) As List(Of InboxBoardEntry)
        ' Group by ConversationTopic (case-insensitive). Mails without a topic stay ungrouped.
        Dim grouped As New Dictionary(Of String, List(Of InboxBoardEntry))(StringComparer.OrdinalIgnoreCase)
        Dim ungrouped As New List(Of InboxBoardEntry)()

        For Each entry In entries
            If String.IsNullOrWhiteSpace(entry.ConversationTopic) Then
                entry.IsGrouped = False
                ungrouped.Add(entry)
            Else
                Dim key As String = entry.ConversationTopic.Trim()
                If Not grouped.ContainsKey(key) Then
                    grouped(key) = New List(Of InboxBoardEntry)()
                End If
                grouped(key).Add(entry)
            End If
        Next

        Dim result As New List(Of InboxBoardEntry)()

        For Each kvp In grouped
            Dim convMails As List(Of InboxBoardEntry) = kvp.Value

            ' If only one mail shares this topic, treat it as ungrouped
            If convMails.Count = 1 Then
                convMails(0).IsGrouped = False
                result.Add(convMails(0))
                Continue For
            End If

            ' Sort by ReceivedTime descending — first item is the latest
            convMails.Sort(Function(a, b) b.ReceivedTime.CompareTo(a.ReceivedTime))

            Dim representative As InboxBoardEntry = convMails(0)
            representative.MessageCount = convMails.Count
            representative.IsGrouped = True
            representative.ConversationEntryIDs = convMails.Select(Function(m) m.EntryID).ToList()

            ' If any mail in the conversation is unread, mark the card as unread
            If convMails.Any(Function(m) Not m.IsRead) Then
                representative.IsRead = False
            End If

            ' If any mail in the conversation is flagged, mark the card as flagged
            If convMails.Any(Function(m) m.IsFlagged) Then
                representative.IsFlagged = True
                ' Use the flag request from the first flagged mail
                Dim firstFlagged = convMails.FirstOrDefault(Function(m) m.IsFlagged)
                If firstFlagged IsNot Nothing Then
                    If String.IsNullOrEmpty(representative.FlagRequest) Then
                        representative.FlagRequest = firstFlagged.FlagRequest
                    End If
                    If Not representative.FlagDueDate.HasValue AndAlso firstFlagged.FlagDueDate.HasValue Then
                        representative.FlagDueDate = firstFlagged.FlagDueDate
                    End If
                End If
            End If

            ' Recompute flag bucket for the representative card
            representative.FlagBucket = GetFlagBucket(representative)

            ' Merge all categories from all mails in the conversation
            Dim allCats As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each m In convMails
                If Not String.IsNullOrEmpty(m.AllCategories) Then
                    For Each catPart In m.AllCategories.Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
                        Dim t = catPart.Trim()
                        If Not String.IsNullOrEmpty(t) Then allCats.Add(t)
                    Next
                End If
            Next
            representative.AllCategories = String.Join(", ", allCats)
            If allCats.Count > 0 Then representative.Categories = allCats.First()

            ' Use board folder from the latest mail that has one set
            If String.IsNullOrEmpty(representative.BoardFolder) Then
                For Each m In convMails
                    If Not String.IsNullOrEmpty(m.BoardFolder) Then
                        representative.BoardFolder = m.BoardFolder
                        Exit For
                    End If
                Next
            End If

            ' Check if any mail in the conversation has a cached summary
            If String.IsNullOrEmpty(representative.Summary) Then
                For Each m In convMails
                    Dim cached As String = GetCachedSummary(m.EntryID)
                    If Not String.IsNullOrEmpty(cached) Then
                        representative.Summary = cached
                        Exit For
                    End If
                Next
            End If

            result.Add(representative)
        Next

        ' Add all ungrouped mails — these were previously lost when toggling grouping off/on
        result.AddRange(ungrouped)

        ' Re-sort by received time descending
        result.Sort(Function(a, b) b.ReceivedTime.CompareTo(a.ReceivedTime))

        Return result
    End Function

    ''' <summary>
    ''' Builds an InboxBoardEntry from a MailItem.
    ''' Returns Nothing if the mail does not qualify for the board.
    ''' </summary>
    Private Function BuildInboxBoardEntry(mi As MailItem, categoryColors As Dictionary(Of String, String),
                                          includeFlagged As Boolean, hideDoneFlags As Boolean) As InboxBoardEntry
        Dim entry As New InboxBoardEntry()

        entry.EntryID = ComRetry(Of String)(Function() mi.EntryID)
        entry.Subject = ComRetry(Of String)(Function() If(mi.Subject, "(no subject)"))
        entry.SenderName = ComRetry(Of String)(Function() If(mi.SenderName, ""))
        entry.SenderEmail = ""
        Try : entry.SenderEmail = ComRetry(Of String)(Function() If(mi.SenderEmailAddress, "")) : Catch : End Try
        entry.Recipients = ""
        Try : entry.Recipients = ComRetry(Of String)(Function() CStr(mi.To)) : Catch : End Try
        entry.ReceivedTime = ComRetry(Of Date)(Function() mi.ReceivedTime)
        entry.IsRead = ComRetry(Of Boolean)(Function() Not mi.UnRead)

        ' MessageCount = 1 for ungrouped cards; GroupByConversation overwrites it for grouped cards.
        entry.MessageCount = 1
        entry.IsGrouped = False

        ' ConversationTopic for grouping
        entry.ConversationTopic = ""
        Try : entry.ConversationTopic = ComRetry(Of String)(Function() If(mi.ConversationTopic, "")) : Catch : End Try
        entry.ConversationEntryIDs = New List(Of String)() From {entry.EntryID}

        ' Flag properties
        entry.IsFlagged = False
        entry.IsCompleted = False
        entry.FlagRequest = ""
        entry.FlagDueDate = Nothing
        Try
            Dim flagStatus As Integer = ComRetry(Of Integer)(Function() CInt(mi.FlagStatus))
            ' olFlagMarked = 2, olFlagComplete = 1, olNoFlag = 0
            If flagStatus = 2 Then
                entry.IsFlagged = True
            ElseIf flagStatus = 1 Then
                entry.IsCompleted = True
            End If
        Catch
        End Try
        If entry.IsFlagged OrElse entry.IsCompleted Then
            Try : entry.FlagRequest = ComRetry(Of String)(Function() If(mi.FlagRequest, "")) : Catch : End Try
            Try
                Dim dueDate As DateTime = ComRetry(Of Date)(Function() mi.TaskDueDate)
                ' Outlook uses 1/1/4501 as "no date"
                If dueDate.Year < 4000 Then entry.FlagDueDate = dueDate
            Catch
            End Try
        End If

        ' Categories
        Dim cats As String = ""
        Try : cats = ComRetry(Of String)(Function() If(mi.Categories, "")) : Catch : End Try
        entry.AllCategories = cats

        Dim hasCats As Boolean = Not String.IsNullOrEmpty(cats)

        If hasCats Then
            Dim firstCat As String = cats.Split({","c, ";"c})(0).Trim()
            entry.Categories = firstCat
            entry.CategoryColor = "#6b7280" ' default gray
            If categoryColors.ContainsKey(firstCat) Then
                entry.CategoryColor = categoryColors(firstCat)
            End If
        Else
            ' No category — only include if flagged and includeFlagged is on
            If includeFlagged AndAlso entry.IsFlagged Then
                entry.Categories = ""
                entry.CategoryColor = InboxBoard_FlaggedColumnColor
            ElseIf includeFlagged AndAlso entry.IsCompleted AndAlso Not hideDoneFlags Then
                entry.Categories = ""
                entry.CategoryColor = "#9ca3af" ' muted gray for completed
            Else
                Return Nothing ' Does not qualify
            End If
        End If

        ' Body excerpt for AI summarization
        entry.BodyExcerpt = ""
        Try
            Dim body As String = GetMailBody(mi)
            If Not String.IsNullOrEmpty(body) Then
                body = GetLatestMailBody(body)
                If body.Length > InboxBoard_SummaryBodyCap Then
                    body = body.Substring(0, InboxBoard_SummaryBodyCap)
                End If
                entry.BodyExcerpt = body
            End If
        Catch
        End Try

        ' Compute flag bucket for date-grouped flag columns
        entry.FlagBucket = GetFlagBucket(entry)

        ' Board folder assignment (UserProperty)
        entry.BoardFolder = GetMailBoardFolder(mi)

        ' Compute flag bucket for date-grouped flag columns
        entry.FlagBucket = GetFlagBucket(entry)

        entry.Summary = ""
        Return entry
    End Function

    ''' <summary>
    ''' Determines which flag date-bucket a mail belongs to based on its flag status
    ''' and TaskDueDate. Returns the bucket column ID string.
    ''' </summary>
    Private Shared Function GetFlagBucket(entry As InboxBoardEntry) As String
        If entry.IsCompleted Then Return InboxBoard_FlagBucket_Done
        If Not entry.IsFlagged Then Return Nothing

        If Not entry.FlagDueDate.HasValue Then Return InboxBoard_FlagBucket_NoDueDate

        Dim dueDate As DateTime = entry.FlagDueDate.Value.Date
        Dim today As DateTime = DateTime.Today

        ' Overdue: before today
        If dueDate < today Then Return InboxBoard_FlagBucket_Overdue

        ' Today
        If dueDate = today Then Return InboxBoard_FlagBucket_Today

        ' Tomorrow
        If dueDate = today.AddDays(1) Then Return InboxBoard_FlagBucket_Tomorrow

        ' This week: up to end of current week (Sunday)
        Dim endOfWeek As DateTime = today.AddDays(7 - CInt(today.DayOfWeek))
        If dueDate <= endOfWeek Then Return InboxBoard_FlagBucket_ThisWeek

        ' Next week: Monday–Sunday of next week
        Dim endOfNextWeek As DateTime = endOfWeek.AddDays(7)
        If dueDate <= endOfNextWeek Then Return InboxBoard_FlagBucket_NextWeek

        ' Next 4 weeks: within 4 calendar weeks from today
        If dueDate <= today.AddDays(28) Then Return InboxBoard_FlagBucket_Next4Weeks

        ' Later: everything else
        Return InboxBoard_FlagBucket_Later
    End Function

    ''' <summary>
    ''' Gets a mapping of category name → hex color from Outlook's category list.
    ''' </summary>
    Private Function GetOutlookCategoryColors(ns As Outlook.NameSpace) As Dictionary(Of String, String)
        Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        ' Map OlCategoryColor enum values to hex colors
        Dim colorMap As New Dictionary(Of Integer, String)() From {
            {0, "#dc2626"},   ' olCategoryColorNone -> red
            {1, "#dc2626"},   ' Red
            {2, "#f97316"},   ' Orange
            {3, "#854d0e"},   ' Peach/Brown
            {4, "#eab308"},   ' Yellow
            {5, "#22c55e"},   ' Green
            {6, "#14b8a6"},   ' Teal
            {7, "#06b6d4"},   ' Olive -> Cyan
            {8, "#3b82f6"},   ' Blue
            {9, "#8b5cf6"},   ' Purple
            {10, "#d946ef"},  ' Maroon -> Fuchsia
            {11, "#64748b"},  ' Steel
            {12, "#6b7280"},  ' DarkSteel
            {13, "#9ca3af"},  ' Gray
            {14, "#475569"},  ' DarkGray
            {15, "#1e293b"},  ' Black
            {16, "#991b1b"},  ' DarkRed
            {17, "#c2410c"},  ' DarkOrange
            {18, "#a16207"},  ' DarkPeach
            {19, "#a3e635"},  ' DarkYellow -> Lime
            {20, "#15803d"},  ' DarkGreen
            {21, "#0d9488"},  ' DarkTeal
            {22, "#0284c7"},  ' DarkOlive -> Sky
            {23, "#1d4ed8"},  ' DarkBlue
            {24, "#7c3aed"},  ' DarkPurple
            {25, "#a21caf"}   ' DarkMaroon -> DarkFuchsia
        }

        Try
            Dim categories As Outlook.Categories = ns.Categories
            For Each cat As Outlook.Category In categories
                Try
                    Dim name As String = cat.Name
                    Dim colorVal As Integer = CInt(cat.Color)
                    Dim hexColor As String = "#6b7280"
                    If colorMap.ContainsKey(colorVal) Then hexColor = colorMap(colorVal)
                    result(name) = hexColor
                Catch
                End Try
            Next
        Catch
        End Try

        Return result
    End Function


#End Region

#Region "InboxBoard Column Building"

    ''' <summary>
    ''' Builds the list of board columns from found categories + always-present columns from settings.
    ''' If flagged mails are included, adds a synthetic "⚑ Flagged" column.
    ''' </summary>
    Private Function BuildBoardColumns(mails As List(Of InboxBoardEntry)) As List(Of InboxBoardColumn)
        Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
        Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
        Dim categoryColors As Dictionary(Of String, String) = GetOutlookCategoryColors(ns)

        ' Collect unique categories from ALL categories on every mail (not just the first)
        Dim usedCats As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each m In mails
            If Not String.IsNullOrEmpty(m.AllCategories) Then
                For Each catPart In m.AllCategories.Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
                    Dim catName As String = catPart.Trim()
                    If Not String.IsNullOrEmpty(catName) Then usedCats.Add(catName)
                Next
            End If
        Next

        ' Add always-present columns from settings
        Dim alwaysPresent As String = ""
        Try : alwaysPresent = If(My.Settings.InboxBoardColumns, "") : Catch : End Try
        If Not String.IsNullOrEmpty(alwaysPresent) Then
            For Each col In alwaysPresent.Split({";"c}, StringSplitOptions.RemoveEmptyEntries)
                Dim trimmed As String = col.Trim()
                If Not String.IsNullOrEmpty(trimmed) Then usedCats.Add(trimmed)
            Next
        End If

        ' Build column list
        Dim columns As New List(Of InboxBoardColumn)()
        For Each catName In usedCats
            Dim hexColor As String = "#6b7280"
            If categoryColors.ContainsKey(catName) Then hexColor = categoryColors(catName)

            columns.Add(New InboxBoardColumn With {
                .CategoryName = catName,
                .Color = hexColor,
                .Tag = catName,
                .IsFlaggedColumn = False
            })
        Next

        ' Add flagged column(s) when the setting is on
        Dim includeFlagged As Boolean = False
        Try : includeFlagged = My.Settings.InboxBoardIncludeFlagged : Catch : End Try
        If includeFlagged Then
            Dim groupByDate As Boolean = False
            Try : groupByDate = My.Settings.InboxBoardGroupFlagsByDate : Catch : End Try
            Dim hideDoneFlags As Boolean = True
            Try : hideDoneFlags = My.Settings.InboxBoardHideDoneFlags : Catch : End Try

            If groupByDate Then
                ' Add one column per date bucket that has at least one mail, plus pinned buckets
                Dim usedBuckets As New HashSet(Of String)(StringComparer.Ordinal)
                For Each m In mails
                    If m.IsFlagged OrElse m.IsCompleted Then
                        Dim bucket = GetFlagBucket(m)
                        If bucket IsNot Nothing Then usedBuckets.Add(bucket)
                    End If
                Next

                ' Always show Today and Tomorrow even if empty, so users can drag into them
                'usedBuckets.Add(InboxBoard_FlagBucket_Today)
                'usedBuckets.Add(InboxBoard_FlagBucket_Tomorrow)

                ' Add pinned flag columns from settings (user-chosen always-visible buckets)
                Dim pinnedFlagCols As String = ""
                Try : pinnedFlagCols = If(My.Settings.InboxBoardPinnedFlagColumns, "") : Catch : End Try
                If Not String.IsNullOrEmpty(pinnedFlagCols) Then
                    For Each part In pinnedFlagCols.Split({";"c}, StringSplitOptions.RemoveEmptyEntries)
                        Dim t = part.Trim()
                        If Not String.IsNullOrEmpty(t) Then usedBuckets.Add(t)
                    Next
                End If

                ' Remove Done bucket if hidden
                If hideDoneFlags Then usedBuckets.Remove(InboxBoard_FlagBucket_Done)

                ' Insert in the defined display order, at position 0 (before category columns)
                Dim insertIdx As Integer = 0
                For Each bucketId In InboxBoard_FlagBucketOrder
                    If usedBuckets.Contains(bucketId) Then
                        Dim bucketColor As String = "#6b7280"
                        If InboxBoard_FlagBucketColors.ContainsKey(bucketId) Then bucketColor = InboxBoard_FlagBucketColors(bucketId)
                        columns.Insert(insertIdx, New InboxBoardColumn With {
                            .CategoryName = bucketId,
                            .Color = bucketColor,
                            .Tag = bucketId,
                            .IsFlaggedColumn = True
                        })
                        insertIdx += 1
                    End If
                Next
            Else
                ' Single "⚑ Flagged" column (original behavior)
                columns.Insert(0, New InboxBoardColumn With {
                    .CategoryName = InboxBoard_FlaggedColumnId,
                    .Color = InboxBoard_FlaggedColumnColor,
                    .Tag = InboxBoard_FlaggedColumnId,
                    .IsFlaggedColumn = True
                })
            End If
        End If

        ' Apply saved column order from My.Settings
        Dim savedOrder As String = ""
        Try : savedOrder = If(My.Settings.InboxBoardColumnOrder, "") : Catch : End Try
        If Not String.IsNullOrEmpty(savedOrder) Then
            Dim orderList As New List(Of String)()
            For Each part In savedOrder.Split({";"c}, StringSplitOptions.RemoveEmptyEntries)
                Dim t = part.Trim()
                If Not String.IsNullOrEmpty(t) Then orderList.Add(t)
            Next
            If orderList.Count > 0 Then
                columns.Sort(Function(a, b)
                                 Dim idxA As Integer = orderList.FindIndex(Function(o) o.Equals(a.CategoryName, StringComparison.OrdinalIgnoreCase))
                                 Dim idxB As Integer = orderList.FindIndex(Function(o) o.Equals(b.CategoryName, StringComparison.OrdinalIgnoreCase))
                                 ' Columns not in the saved order go to the end, preserving their relative position
                                 If idxA < 0 Then idxA = Integer.MaxValue
                                 If idxB < 0 Then idxB = Integer.MaxValue
                                 Return idxA.CompareTo(idxB)
                             End Function)
            End If
        End If

        Return columns
    End Function

#End Region

#Region "InboxBoard Target Folder"

    ''' <summary>
    ''' Resolves the target MAPI folder for the board based on persisted
    ''' My.Settings (StoreId + FolderPath). Falls back to the default Inbox.
    ''' Also sets <see cref="_inboxBoardTrackedFolderLabel"/> for UI display.
    ''' </summary>
    ''' <returns>The resolved MAPIFolder, or Nothing if unreachable.</returns>
    Private Function GetInboxBoardTargetFolder() As MAPIFolder
        Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
        Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")

        Dim savedStoreId As String = ""
        Dim savedFolderPath As String = ""
        Try : savedStoreId = If(My.Settings.InboxBoardStoreId, "") : Catch : End Try
        Try : savedFolderPath = If(My.Settings.InboxBoardFolderPath, "") : Catch : End Try

        ' If both are empty, use the currently open folder in the Explorer
        If String.IsNullOrWhiteSpace(savedStoreId) AndAlso String.IsNullOrWhiteSpace(savedFolderPath) Then
            Dim explorerFolder As MAPIFolder = Nothing
            Try
                Dim explorer As Outlook.Explorer = outlookApp.ActiveExplorer()
                If explorer IsNot Nothing Then
                    explorerFolder = ComRetry(Of MAPIFolder)(Function() explorer.CurrentFolder)
                End If
            Catch
            End Try

            If explorerFolder Is Nothing Then
                explorerFolder = ComRetry(Of MAPIFolder)(Function() ns.GetDefaultFolder(OlDefaultFolders.olFolderInbox))
            End If

            _inboxBoardTrackedFolderLabel = GetFolderDisplayLabel(explorerFolder)
            Return explorerFolder
        End If

        ' Try to resolve the store
        Dim targetStore As Outlook.Store = Nothing
        If Not String.IsNullOrWhiteSpace(savedStoreId) Then
            Try
                For Each st As Outlook.Store In ns.Stores
                    Try
                        If st.StoreID = savedStoreId Then
                            targetStore = st
                            Exit For
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try
        End If

        ' If no specific folder path, use the store's default Inbox
        If String.IsNullOrWhiteSpace(savedFolderPath) Then
            If targetStore IsNot Nothing Then
                Try
                    Dim storeInbox As MAPIFolder = ComRetry(Of MAPIFolder)(Function() targetStore.GetDefaultFolder(OlDefaultFolders.olFolderInbox))
                    If storeInbox IsNot Nothing Then
                        _inboxBoardTrackedFolderLabel = GetFolderDisplayLabel(storeInbox)
                        Return storeInbox
                    End If
                Catch
                End Try
            End If

            ' Fallback
            Dim fallback As MAPIFolder = ComRetry(Of MAPIFolder)(Function() ns.GetDefaultFolder(OlDefaultFolders.olFolderInbox))
            _inboxBoardTrackedFolderLabel = GetFolderDisplayLabel(fallback)
            Return fallback
        End If

        ' Resolve a specific folder by path within the store
        Dim targetFolder As MAPIFolder = Nothing
        If targetStore IsNot Nothing Then
            targetFolder = FindFolderByPathInStore(targetStore, savedFolderPath)
        End If

        ' If store-based lookup failed, try all stores (in case the store was re-added)
        If targetFolder Is Nothing Then
            Try
                For Each st As Outlook.Store In ns.Stores
                    targetFolder = FindFolderByPathInStore(st, savedFolderPath)
                    If targetFolder IsNot Nothing Then Exit For
                Next
            Catch
            End Try
        End If

        If targetFolder IsNot Nothing Then
            _inboxBoardTrackedFolderLabel = GetFolderDisplayLabel(targetFolder)
            Return targetFolder
        End If

        ' Final fallback
        Debug.WriteLine($"[InboxBoard] Could not resolve saved folder StoreId={savedStoreId}, Path={savedFolderPath}. Falling back to default Inbox.")
        Dim defaultInbox As MAPIFolder = ComRetry(Of MAPIFolder)(Function() ns.GetDefaultFolder(OlDefaultFolders.olFolderInbox))
        _inboxBoardTrackedFolderLabel = GetFolderDisplayLabel(defaultInbox)
        Return defaultInbox
    End Function

    ''' <summary>
    ''' Navigates a store's folder tree to find a folder by its relative path
    ''' (e.g. "Inbox\Projects\Active"). Path parts are separated by backslash.
    ''' </summary>
    Private Function FindFolderByPathInStore(store As Outlook.Store, relativePath As String) As MAPIFolder
        If store Is Nothing OrElse String.IsNullOrWhiteSpace(relativePath) Then Return Nothing
        Try
            Dim parts As String() = relativePath.Split(New Char() {"\"c}, StringSplitOptions.RemoveEmptyEntries)
            If parts.Length = 0 Then Return Nothing

            Dim root As MAPIFolder = store.GetRootFolder()
            Dim current As MAPIFolder = root

            For Each partName In parts
                Dim found As Boolean = False
                Dim subFolders As Outlook.Folders = current.Folders
                For Each subF As MAPIFolder In subFolders
                    If String.Equals(subF.Name, partName, StringComparison.OrdinalIgnoreCase) Then
                        current = subF
                        found = True
                        Exit For
                    End If
                Next
                If Not found Then Return Nothing
            Next

            Return current
        Catch
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Builds a human-readable display label for a folder, e.g. "user@example.com / Inbox".
    ''' </summary>
    Private Function GetFolderDisplayLabel(folder As MAPIFolder) As String
        If folder Is Nothing Then Return "Inbox"
        Try
            Dim folderPath As String = folder.FolderPath
            ' FolderPath looks like "\\account@mail.com\Inbox\Sub"
            ' Strip the leading \\ and split into store-name + relative path
            Dim trimmed As String = folderPath.TrimStart("\"c)
            Dim sepIdx As Integer = trimmed.IndexOf("\"c)
            If sepIdx > 0 Then
                Dim storeName As String = trimmed.Substring(0, sepIdx)
                Dim relPath As String = trimmed.Substring(sepIdx + 1)
                Return storeName & " / " & relPath.Replace("\", " / ")
            End If
            Return trimmed.Replace("\", " / ")
        Catch
            Return "Inbox"
        End Try
    End Function

    ''' <summary>
    ''' Builds the relative folder path (within its store) from the full FolderPath.
    ''' E.g. "\\account\Inbox\Sub" → "Inbox\Sub".
    ''' </summary>
    Private Shared Function GetRelativeFolderPath(folder As MAPIFolder) As String
        If folder Is Nothing Then Return ""
        Try
            Dim folderPath As String = folder.FolderPath
            Dim trimmed As String = folderPath.TrimStart("\"c)
            Dim sepIdx As Integer = trimmed.IndexOf("\"c)
            If sepIdx > 0 Then Return trimmed.Substring(sepIdx + 1)
            Return trimmed
        Catch
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Prompts the user to select a mailbox (Store) and then a folder within it.
    ''' Persists the choice to My.Settings. Returns True if the user confirmed,
    ''' False if cancelled.
    ''' </summary>
    Private Function PromptSelectTargetFolder() As Boolean
        Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
        Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")

        ' ── Step 1: Select mailbox (Store) ──
        Dim stores As New List(Of Outlook.Store)()
        Try
            For Each st As Outlook.Store In ns.Stores
                Try
                    ' Only include stores that have a valid display name and root folder
                    Dim rootF As MAPIFolder = st.GetRootFolder()
                    If rootF IsNot Nothing Then stores.Add(st)
                Catch
                End Try
            Next
        Catch
        End Try

        If stores.Count = 0 Then
            ShowCustomMessageBox("No mailbox stores found.", $"{AN} - Inbox Board")
            Return False
        End If

        ' Read currently saved StoreId for pre-selection
        Dim currentStoreId As String = ""
        Try : currentStoreId = If(My.Settings.InboxBoardStoreId, "") : Catch : End Try

        Dim storeItems As New List(Of SelectionItem)()
        Dim defaultStoreIdx As Integer = 1
        For i As Integer = 0 To stores.Count - 1
            Dim displayName As String = ""
            Try : displayName = stores(i).DisplayName : Catch : End Try
            If String.IsNullOrWhiteSpace(displayName) Then displayName = $"Store {i + 1}"
            storeItems.Add(New SelectionItem(displayName, i + 1))
            If Not String.IsNullOrWhiteSpace(currentStoreId) AndAlso stores(i).StoreID = currentStoreId Then
                defaultStoreIdx = i + 1
            End If
        Next

        Dim selectedStoreVal As Integer = SelectValue(
            storeItems, defaultStoreIdx,
            "Select the mailbox (account) to track:",
            $"{AN} - Inbox Board — Mailbox")
        If selectedStoreVal <= 0 Then Return False

        Dim chosenStore As Outlook.Store = stores(selectedStoreVal - 1)

        ' ── Step 2: Select folder within the store ──
        Dim folderPaths As New List(Of String)()
        Dim folderObjects As New List(Of MAPIFolder)()
        Try
            Dim rootFolder As MAPIFolder = chosenStore.GetRootFolder()
            CollectFoldersRecursive(rootFolder, folderPaths, folderObjects, 0)
        Catch
        End Try

        If folderPaths.Count = 0 Then
            ShowCustomMessageBox("No folders found in the selected mailbox.", $"{AN} - Inbox Board")
            Return False
        End If

        ' Read currently saved folder path for pre-selection
        Dim currentFolderPath As String = ""
        Try : currentFolderPath = If(My.Settings.InboxBoardFolderPath, "") : Catch : End Try

        Dim folderItems As New List(Of SelectionItem)()
        Dim defaultFolderIdx As Integer = 1
        For i As Integer = 0 To folderPaths.Count - 1
            folderItems.Add(New SelectionItem(folderPaths(i), i + 1))
            ' Pre-select either the saved folder or the Inbox
            Dim relPath As String = GetRelativeFolderPath(folderObjects(i))
            If Not String.IsNullOrWhiteSpace(currentFolderPath) AndAlso
               chosenStore.StoreID = currentStoreId AndAlso
               relPath.Equals(currentFolderPath, StringComparison.OrdinalIgnoreCase) Then
                defaultFolderIdx = i + 1
            ElseIf String.IsNullOrWhiteSpace(currentFolderPath) Then
                ' Default to Inbox folder
                Try
                    If folderObjects(i).DefaultItemType = OlItemType.olMailItem AndAlso
                       folderObjects(i).Name.Equals("Inbox", StringComparison.OrdinalIgnoreCase) Then
                        defaultFolderIdx = i + 1
                    End If
                Catch
                End Try
            End If
        Next

        Dim selectedFolderVal As Integer = SelectValue(
            folderItems, defaultFolderIdx,
            $"Select the folder to track in ""{chosenStore.DisplayName}"":",
            $"{AN} - Inbox Board — Folder")
        If selectedFolderVal <= 0 Then Return False

        Dim chosenFolder As MAPIFolder = folderObjects(selectedFolderVal - 1)
        Dim chosenRelPath As String = GetRelativeFolderPath(chosenFolder)

        ' ── Persist ──
        Try
            My.Settings.InboxBoardStoreId = chosenStore.StoreID
            My.Settings.InboxBoardFolderPath = chosenRelPath
            My.Settings.Save()
        Catch
        End Try

        ' Reset load count so user is re-prompted on next scan
        _inboxBoardLastMaxToLoad = 0

        Debug.WriteLine($"[InboxBoard] Target folder set: Store={chosenStore.DisplayName}, Path={chosenRelPath}")
        Return True
    End Function

    ''' <summary>
    ''' Recursively collects all folders in a store for display in the folder picker.
    ''' Caps at a maximum depth and total folder count to prevent crashes on large mailboxes.
    ''' </summary>
    Private Sub CollectFoldersRecursive(folder As MAPIFolder, paths As List(Of String),
                                         folders As List(Of MAPIFolder), depth As Integer)
        ' Guard: cap recursion depth and total folder count to protect against
        ' huge Exchange/public-folder trees that can exhaust COM resources.
        Const MaxDepth As Integer = 15
        Const MaxFolders As Integer = 500

        If depth > MaxDepth OrElse paths.Count >= MaxFolders Then Return

        Try
            ' Build an indented display name
            Dim indent As String = New String(" "c, depth * 3)
            Dim displayName As String = indent & folder.Name
            paths.Add(displayName)
            folders.Add(folder)

            Dim subFolders As Outlook.Folders = Nothing
            Try
                subFolders = folder.Folders
            Catch
                Return
            End Try

            If subFolders Is Nothing Then Return

            Dim subCount As Integer = 0
            Try : subCount = subFolders.Count : Catch : Return : End Try

            For i As Integer = 1 To subCount
                If paths.Count >= MaxFolders Then Exit For
                Try
                    Dim idx As Integer = i
                    Dim subF As MAPIFolder = TryCast(subFolders.Item(idx), MAPIFolder)
                    If subF IsNot Nothing Then
                        CollectFoldersRecursive(subF, paths, folders, depth + 1)
                    End If
                Catch
                End Try
            Next
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Collects all MailItems from a folder, optionally including subfolders.
    ''' Reports progress and supports cancellation while scanning.
    ''' </summary>
    Private Function CollectMailItemsFromFolder(folder As MAPIFolder,
                                            includeSubfolders As Boolean,
                                            progress As InboxBoardProgressScope,
                                            state As InboxBoardScanProgressState) As List(Of MailItem)

        Dim result As New List(Of MailItem)()
        If folder Is Nothing Then Return result
        If progress IsNot Nothing AndAlso progress.CancelRequested Then Return result

        Dim folderLabel As String = GetFolderDisplayLabel(folder)

        Try
            Dim folderItems As Outlook.Items = ComRetry(Of Outlook.Items)(Function() folder.Items)
            Dim totalItems As Integer = 0
            Try : totalItems = ComRetry(Of Integer)(Function() folderItems.Count) : Catch : End Try

            If totalItems > 0 Then
                state.ExpectedItems += totalItems
                ReportInboxBoardProgress(
                    progress,
                    state.ProcessedItems,
                    state.ExpectedItems,
                    $"Scanning {folderLabel}…")
            End If

            For i As Integer = 1 To totalItems
                If progress IsNot Nothing AndAlso progress.CancelRequested Then Exit For

                Try
                    Dim idx As Integer = i
                    Dim item As Object = ComRetry(Function() folderItems.Item(idx))
                    Dim mi As MailItem = TryCast(item, MailItem)
                    If mi IsNot Nothing Then result.Add(mi)
                Catch
                End Try

                state.ProcessedItems += 1

                If ShouldReportInboxBoardProgress(state.ProcessedItems, state.ExpectedItems) Then
                    ReportInboxBoardProgress(
                        progress,
                        state.ProcessedItems,
                        state.ExpectedItems,
                        $"Scanning {folderLabel}… {state.ProcessedItems:N0}/{state.ExpectedItems:N0}")
                End If
            Next
        Catch
        End Try

        If progress IsNot Nothing AndAlso progress.CancelRequested Then Return result

        If includeSubfolders Then
            Try
                Dim subFolders As Outlook.Folders = folder.Folders
                For Each subF As MAPIFolder In subFolders
                    If progress IsNot Nothing AndAlso progress.CancelRequested Then Exit For
                    result.AddRange(CollectMailItemsFromFolder(subF, True, progress, state))
                Next
            Catch
            End Try
        End If

        Return result
    End Function

    ''' <summary>
    ''' Resets the tracked folder to the default Inbox and clears persisted settings.
    ''' </summary>
    Private Sub ResetInboxBoardTargetFolder()
        Try
            My.Settings.InboxBoardStoreId = ""
            My.Settings.InboxBoardFolderPath = ""
            My.Settings.Save()
        Catch
        End Try
        _inboxBoardLastMaxToLoad = 0
        _inboxBoardTrackedFolderLabel = "Inbox"
    End Sub

#End Region

#Region "InboxBoard WebView2 Form"

    ''' <summary>
    ''' Shows the board in a WebView2-hosted WinForms dialog.
    ''' </summary>
    Private Sub ShowInboxBoardForm(mails As List(Of InboxBoardEntry), columns As List(Of InboxBoardColumn))
        Dim frm As New Form()
        frm.Text = $"{AN} - Inbox Board"
        frm.MinimumSize = New Drawing.Size(900, 500)
        frm.FormBorderStyle = FormBorderStyle.Sizable
        frm.MaximizeBox = True
        frm.MinimizeBox = True
        frm.ShowInTaskbar = True

        ' Restore window position, size and maximized state from My.Settings
        Dim savedX As Integer = 0, savedY As Integer = 0
        Dim savedW As Integer = 0, savedH As Integer = 0
        Dim savedMax As Integer = 0
        Try
            savedX = My.Settings.InboxBoardWindowX
            savedY = My.Settings.InboxBoardWindowY
            savedW = My.Settings.InboxBoardWindowW
            savedH = My.Settings.InboxBoardWindowH
            savedMax = My.Settings.InboxBoardWindowMax
        Catch
        End Try

        If savedW >= 900 AndAlso savedH >= 500 Then
            frm.Size = New Drawing.Size(savedW, savedH)
            ' Verify the saved position is still on a visible screen
            Dim savedRect As New Drawing.Rectangle(savedX, savedY, savedW, savedH)
            Dim onScreen As Boolean = False
            For Each scr As Screen In Screen.AllScreens
                If scr.WorkingArea.IntersectsWith(savedRect) Then
                    onScreen = True
                    Exit For
                End If
            Next
            If onScreen Then
                frm.StartPosition = FormStartPosition.Manual
                frm.Location = New Drawing.Point(savedX, savedY)
            Else
                frm.StartPosition = FormStartPosition.CenterScreen
            End If
        Else
            frm.Size = New Drawing.Size(1500, 850)
            frm.StartPosition = FormStartPosition.CenterScreen
        End If

        If savedMax <> 0 Then frm.WindowState = FormWindowState.Maximized

        Try
            Dim bmp As New Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
            frm.Icon = Drawing.Icon.FromHandle(bmp.GetHicon())
            bmp.Dispose()
        Catch
        End Try

        Dim webView As New WebView2() With {
            .Dock = DockStyle.Fill
        }
        frm.Controls.Add(webView)

        ' Track mails + columns for callbacks
        Dim boardMails As List(Of InboxBoardEntry) = mails
        Dim boardColumns As List(Of InboxBoardColumn) = columns
        Dim summaryCts As New CancellationTokenSource()

        ' Temp file path for the board HTML (gives a real origin so localStorage works)
        Dim tempHtmlPath As String = ""

        AddHandler frm.FormClosing, Sub(s, e)
                                        summaryCts.Cancel()

                                        ' Persist summary cache on close
                                        SaveSummaryCache()

                                        ' Save window position, size and maximized state
                                        Try
                                            If frm.WindowState = FormWindowState.Maximized Then
                                                My.Settings.InboxBoardWindowMax = 1
                                                My.Settings.InboxBoardWindowX = frm.RestoreBounds.X
                                                My.Settings.InboxBoardWindowY = frm.RestoreBounds.Y
                                                My.Settings.InboxBoardWindowW = frm.RestoreBounds.Width
                                                My.Settings.InboxBoardWindowH = frm.RestoreBounds.Height
                                            Else
                                                My.Settings.InboxBoardWindowMax = 0
                                                My.Settings.InboxBoardWindowX = frm.Location.X
                                                My.Settings.InboxBoardWindowY = frm.Location.Y
                                                My.Settings.InboxBoardWindowW = frm.Size.Width
                                                My.Settings.InboxBoardWindowH = frm.Size.Height
                                            End If
                                            My.Settings.Save()
                                        Catch
                                        End Try

                                        ' Clean up temp file
                                        Try
                                            If Not String.IsNullOrEmpty(tempHtmlPath) AndAlso System.IO.File.Exists(tempHtmlPath) Then
                                                System.IO.File.Delete(tempHtmlPath)
                                            End If
                                        Catch
                                        End Try
                                    End Sub

        AddHandler frm.Shown, Async Sub(s, e)
                                  Try
                                      Dim userDataFolder As String = SharedLibrary.SharedLibrary.SharedMethods.GetWebView2UserDataFolder()
                                      Dim env As CoreWebView2Environment = Await CoreWebView2Environment.CreateAsync(Nothing, userDataFolder, Nothing)
                                      Await webView.EnsureCoreWebView2Async(env)

                                      AddHandler webView.CoreWebView2.ProcessFailed,
                                          Sub(pfSender, pfArgs)
                                              SharedLibrary.SharedLibrary.SharedMethods.LogWebView2ProcessFailed("InboxBoard", pfArgs.ProcessFailedKind.ToString(), pfArgs.ExitCode.ToString())
                                              Try
                                                  If pfArgs.ProcessFailedKind = CoreWebView2ProcessFailedKind.RenderProcessExited Then
                                                      webView.Reload()
                                                  End If
                                              Catch
                                              End Try
                                          End Sub

                                      ' Handle messages from JS — use TryGetWebMessageAsString()
                                      ' because JS sends JSON.stringify(...) which is a string,
                                      ' not a raw JSON object. WebMessageAsJson would double-encode it.
                                      AddHandler webView.CoreWebView2.WebMessageReceived,
                                          Sub(sender, args)
                                              Dim raw As String = args.TryGetWebMessageAsString()
                                              If Not String.IsNullOrEmpty(raw) Then
                                                  HandleBoardMessage(raw, boardMails, boardColumns, webView, frm, summaryCts)
                                              End If
                                          End Sub

                                      ' Build HTML and write to a temp file so the page has a real
                                      ' file:// origin (about:blank blocks localStorage/sessionStorage).
                                      Dim html As String = BuildInboxBoardHtml(boardMails, boardColumns)
                                      tempHtmlPath = System.IO.Path.Combine(
                                          System.IO.Path.GetTempPath(),
                                          "RedInk_InboxBoard_" & Guid.NewGuid().ToString("N") & ".html")
                                      System.IO.File.WriteAllText(tempHtmlPath, html, Encoding.UTF8)
                                      webView.CoreWebView2.Navigate("file:///" & tempHtmlPath.Replace("\"c, "/"c))

                                      ' Start async summary generation after a short delay
                                      Await System.Threading.Tasks.Task.Delay(2000)
                                      If Not summaryCts.IsCancellationRequested Then
                                          Await GenerateSummariesAsync(boardMails, webView, summaryCts.Token)
                                      End If
                                  Catch ex As System.Exception
                                      Debug.WriteLine($"[InboxBoard] WebView2 init error: {ex.Message}")
                                  End Try
                              End Sub

        frm.Show()
    End Sub

    ''' <summary>
    ''' Handles messages sent from the board's JavaScript via postMessage.
    ''' </summary>
    Private Sub HandleBoardMessage(
    jsonMessage As String,
    mails As List(Of InboxBoardEntry),
    columns As List(Of InboxBoardColumn),
    webView As WebView2,
    frm As Form,
    ByRef summaryCts As CancellationTokenSource)

        Try
            Dim msg As JObject = JObject.Parse(jsonMessage)
            Dim action As String = If(CStr(msg("action")), "")

            Select Case action
                Case "move"
                    ' Replace: remove all existing categories, set only the new one
                    ' For grouped conversations, update ALL mails in the group
                    Dim entryId As String = CStr(msg("entryId"))
                    Dim newCategory As String = CStr(msg("newCategory"))

                    ' If moving to any Flagged bucket column, set the flag + update due date
                    If InboxBoard_FlagBucketColors.ContainsKey(newCategory) OrElse newCategory = InboxBoard_FlaggedColumnId Then
                        Dim entry = mails.FirstOrDefault(Function(m) m.EntryID = entryId)
                        If entry IsNot Nothing Then
                            ' When moving between flag buckets, only update the representative mail's flag.
                            ' The other mails in the conversation keep their own independent flag state.
                            ' Only apply to all grouped mails when moving FROM a category column TO flagged.
                            Dim movingBetweenFlagBuckets As Boolean = (entry.IsFlagged OrElse entry.IsCompleted) AndAlso String.IsNullOrEmpty(entry.AllCategories)
                            Dim idsToUpdate As List(Of String)
                            If movingBetweenFlagBuckets Then
                                idsToUpdate = New List(Of String) From {entryId}
                            Else
                                idsToUpdate = If(entry.IsGrouped AndAlso entry.ConversationEntryIDs IsNot Nothing AndAlso entry.ConversationEntryIDs.Count > 1,
                                                 entry.ConversationEntryIDs,
                                                 New List(Of String) From {entryId})
                            End If
                            Dim newDueDate As DateTime? = GetDueDateForBucket(newCategory)
                            For Each id In idsToUpdate
                                Try
                                    Dim outlookAppFlag As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
                                    Dim nsFlag As Outlook.NameSpace = outlookAppFlag.GetNamespace("MAPI")
                                    Dim miFlag As MailItem = TryCast(ComRetry(Of Object)(Function() nsFlag.GetItemFromID(id)), MailItem)
                                    If miFlag IsNot Nothing Then
                                        ' Only clear categories when entry actually had categories before
                                        ' (don't wipe categories when just moving between flag buckets)
                                        If Not String.IsNullOrEmpty(entry.AllCategories) Then
                                            miFlag.Categories = ""
                                        End If
                                        If newCategory = InboxBoard_FlagBucket_Done Then
                                            miFlag.FlagStatus = OlFlagStatus.olFlagComplete
                                        Else
                                            ' Always (re-)establish the flag and dates. Outlook needs
                                            ' FlagStatus set before dates, and TaskStartDate before
                                            ' TaskDueDate when the new start is after the current due.
                                            miFlag.FlagStatus = OlFlagStatus.olFlagMarked
                                            If newDueDate.HasValue Then
                                                miFlag.TaskStartDate = newDueDate.Value
                                                miFlag.TaskDueDate = newDueDate.Value
                                            Else
                                                miFlag.TaskStartDate = #1/1/4501#
                                                miFlag.TaskDueDate = #1/1/4501#
                                            End If
                                        End If
                                        miFlag.Save()
                                    End If
                                Catch ex As System.Exception
                                    Debug.WriteLine($"[InboxBoard] move-to-flag error: {ex.Message}")
                                End Try
                            Next
                            ' Only blank categories in local state if they were actually cleared in Outlook
                            If Not String.IsNullOrEmpty(entry.AllCategories) Then
                                entry.Categories = ""
                                entry.AllCategories = ""
                            End If
                            If newCategory = InboxBoard_FlagBucket_Done Then
                                entry.IsFlagged = False
                                entry.IsCompleted = True
                            Else
                                entry.IsFlagged = True
                                entry.IsCompleted = False
                            End If
                            entry.FlagDueDate = GetDueDateForBucket(newCategory)
                            entry.FlagBucket = newCategory
                            entry.CategoryColor = If(InboxBoard_FlagBucketColors.ContainsKey(newCategory), InboxBoard_FlagBucketColors(newCategory), InboxBoard_FlaggedColumnColor)
                        End If
                        PushUpdatedCardToJs(entryId, mails, webView)
                        Return
                    End If

                    Dim entryMove = mails.FirstOrDefault(Function(m) m.EntryID = entryId)
                    If entryMove IsNot Nothing Then
                        ' Update all mails in the conversation group (or just the single mail)
                        Dim idsToUpdate As List(Of String) = If(entryMove.IsGrouped AndAlso entryMove.ConversationEntryIDs IsNot Nothing AndAlso entryMove.ConversationEntryIDs.Count > 1,
                                                                 entryMove.ConversationEntryIDs,
                                                                 New List(Of String) From {entryId})
                        For Each id In idsToUpdate
                            UpdateMailCategory(id, newCategory)
                            ' If moving out of Flagged column, clear the flag
                            If entryMove.IsFlagged Then SetMailFlag(id, False)
                        Next

                        ' Update local state
                        entryMove.Categories = newCategory
                        entryMove.AllCategories = newCategory
                        If entryMove.IsFlagged Then
                            entryMove.IsFlagged = False
                            entryMove.IsCompleted = False
                        End If
                        Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
                        Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
                        Dim cc = GetOutlookCategoryColors(ns)
                        entryMove.CategoryColor = If(cc.ContainsKey(newCategory), cc(newCategory), "#6b7280")
                    End If

                    ' Push updated card data back to JS so category badges refresh
                    PushUpdatedCardToJs(entryId, mails, webView)

                Case "moveAdd"
                    ' Add: keep existing categories, append the new one
                    ' For grouped conversations, update ALL mails in the group
                    Dim entryId As String = CStr(msg("entryId"))
                    Dim addCategory As String = CStr(msg("newCategory"))

                    ' If adding to any Flagged bucket column, set the flag + update due date
                    If InboxBoard_FlagBucketColors.ContainsKey(addCategory) OrElse addCategory = InboxBoard_FlaggedColumnId Then
                        Dim entry = mails.FirstOrDefault(Function(m) m.EntryID = entryId)
                        If entry IsNot Nothing Then
                            Dim idsToUpdate As List(Of String) = If(entry.IsGrouped AndAlso entry.ConversationEntryIDs IsNot Nothing AndAlso entry.ConversationEntryIDs.Count > 1,
                                                                     entry.ConversationEntryIDs,
                                                                     New List(Of String) From {entryId})
                            Dim newDueDate As DateTime? = GetDueDateForBucket(addCategory)
                            For Each id In idsToUpdate
                                Try
                                    Dim outlookAppFlag As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
                                    Dim nsFlag As Outlook.NameSpace = outlookAppFlag.GetNamespace("MAPI")
                                    Dim miFlag As MailItem = TryCast(ComRetry(Of Object)(Function() nsFlag.GetItemFromID(id)), MailItem)
                                    If miFlag IsNot Nothing Then
                                        If addCategory = InboxBoard_FlagBucket_Done Then
                                            miFlag.FlagStatus = OlFlagStatus.olFlagComplete
                                        Else
                                            miFlag.FlagStatus = OlFlagStatus.olFlagMarked
                                            If newDueDate.HasValue Then
                                                miFlag.TaskStartDate = newDueDate.Value
                                                miFlag.TaskDueDate = newDueDate.Value
                                            Else
                                                miFlag.TaskStartDate = #1/1/4501#
                                                miFlag.TaskDueDate = #1/1/4501#
                                            End If
                                        End If
                                        miFlag.Save()
                                    End If
                                Catch ex As System.Exception
                                    Debug.WriteLine($"[InboxBoard] moveAdd-to-flag error: {ex.Message}")
                                End Try
                            Next
                            If addCategory = InboxBoard_FlagBucket_Done Then
                                entry.IsFlagged = False
                                entry.IsCompleted = True
                            Else
                                entry.IsFlagged = True
                                entry.IsCompleted = False
                            End If
                            entry.FlagDueDate = GetDueDateForBucket(addCategory)
                            entry.FlagBucket = addCategory
                        End If
                        PushUpdatedCardToJs(entryId, mails, webView)
                        Return
                    End If

                    Dim entryAdd = mails.FirstOrDefault(Function(m) m.EntryID = entryId)
                    If entryAdd IsNot Nothing Then
                        ' Update all mails in the conversation group (or just the single mail)
                        Dim idsToUpdate As List(Of String) = If(entryAdd.IsGrouped AndAlso entryAdd.ConversationEntryIDs IsNot Nothing AndAlso entryAdd.ConversationEntryIDs.Count > 1,
                                                                 entryAdd.ConversationEntryIDs,
                                                                 New List(Of String) From {entryId})
                        For Each id In idsToUpdate
                            AddMailCategory(id, addCategory)
                        Next

                        ' Update local state
                        Dim existing As New List(Of String)()
                        If Not String.IsNullOrEmpty(entryAdd.AllCategories) Then
                            For Each part In entryAdd.AllCategories.Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
                                Dim t = part.Trim()
                                If Not String.IsNullOrEmpty(t) Then existing.Add(t)
                            Next
                        End If
                        If Not existing.Any(Function(c) c.Equals(addCategory, StringComparison.OrdinalIgnoreCase)) Then
                            existing.Add(addCategory)
                        End If
                        entryAdd.AllCategories = String.Join(", ", existing)
                        entryAdd.Categories = existing(0)
                        Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
                        Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
                        Dim cc = GetOutlookCategoryColors(ns)
                        entryAdd.CategoryColor = If(cc.ContainsKey(entryAdd.Categories), cc(entryAdd.Categories), "#6b7280")
                    End If

                    ' Push updated card data back to JS so category badges refresh
                    PushUpdatedCardToJs(entryId, mails, webView)

                Case "unmark"
                    ' Remove ONE category from the mail(s). If the card has multiple categories,
                    ' only the category of the column where unmark was clicked is removed.
                    ' If it was the last category, the card is removed entirely.
                    ' For grouped conversations, all mails in the group are updated.
                    Dim entryId As String = CStr(msg("entryId"))
                    Dim categoryToRemove As String = If(CStr(msg("category")), "")

                    Dim entry = mails.FirstOrDefault(Function(m) m.EntryID = entryId)
                    If entry Is Nothing Then Return

                    ' If unmarking from any Flagged bucket column, clear the flag
                    If InboxBoard_FlagBucketColors.ContainsKey(categoryToRemove) OrElse categoryToRemove = InboxBoard_FlaggedColumnId Then
                        Dim idsToUpdate As List(Of String) = If(entry.IsGrouped AndAlso entry.ConversationEntryIDs IsNot Nothing AndAlso entry.ConversationEntryIDs.Count > 1,
                                                                 entry.ConversationEntryIDs,
                                                                 New List(Of String) From {entryId})
                        For Each id In idsToUpdate
                            SetMailFlag(id, False)
                        Next
                        entry.IsFlagged = False
                        entry.IsCompleted = False

                        ' If the card has no categories either, remove it from the board
                        If String.IsNullOrEmpty(entry.AllCategories) Then
                            mails.RemoveAll(Function(m) m.EntryID = entryId)
                            Try
                                webView.CoreWebView2.PostWebMessageAsJson(
                                    JsonConvert.SerializeObject(New With {.action = "removeCard", .entryId = entryId}))
                            Catch
                            End Try
                        Else
                            PushUpdatedCardToJs(entryId, mails, webView)
                        End If
                        Return
                    End If

                    ' Determine all entry IDs to process (group or single)
                    Dim idsToProcess As List(Of String) = If(entry.IsGrouped AndAlso entry.ConversationEntryIDs IsNot Nothing AndAlso entry.ConversationEntryIDs.Count > 1,
                                                             entry.ConversationEntryIDs,
                                                             New List(Of String) From {entryId})

                    ' Check how many categories the card currently has
                    Dim currentCats As New List(Of String)()
                    If Not String.IsNullOrEmpty(entry.AllCategories) Then
                        For Each part In entry.AllCategories.Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
                            Dim t = part.Trim()
                            If Not String.IsNullOrEmpty(t) Then currentCats.Add(t)
                        Next
                    End If

                    Dim isLastCategory As Boolean = (currentCats.Count <= 1) OrElse
                                                     String.IsNullOrEmpty(categoryToRemove)

                    If isLastCategory Then
                        ' If flagged, keep on board (in Flagged column) but remove categories
                        If entry.IsFlagged Then
                            For Each id In idsToProcess
                                ClearMailCategories(id)
                            Next
                            entry.Categories = ""
                            entry.AllCategories = ""
                            entry.CategoryColor = InboxBoard_FlaggedColumnColor
                            PushUpdatedCardToJs(entryId, mails, webView)
                        Else
                            ' Last (or only) category: clear all categories from all mails, remove card
                            For Each id In idsToProcess
                                ClearMailCategories(id)
                                ' Also clear any flag so it doesn't linger in Outlook
                                If entry.IsFlagged OrElse entry.IsCompleted Then
                                    SetMailFlag(id, False)
                                End If
                            Next
                            entry.IsFlagged = False
                            entry.IsCompleted = False
                            mails.RemoveAll(Function(m) m.EntryID = entryId)

                            ' Tell JS to remove the card
                            Try
                                webView.CoreWebView2.PostWebMessageAsJson(
                                    JsonConvert.SerializeObject(New With {.action = "removeCard", .entryId = entryId}))
                            Catch
                            End Try
                        End If
                    Else
                        ' Multiple categories: remove only the clicked category from all mails
                        For Each id In idsToProcess
                            RemoveMailCategory(id, categoryToRemove)
                        Next

                        ' Update local state
                        Dim remaining = currentCats.Where(Function(c) Not c.Equals(categoryToRemove, StringComparison.OrdinalIgnoreCase)).ToList()
                        entry.AllCategories = String.Join(", ", remaining)
                        entry.Categories = If(remaining.Count > 0, remaining(0), "")

                        Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
                        Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
                        Dim cc = GetOutlookCategoryColors(ns)
                        entry.CategoryColor = If(remaining.Count > 0 AndAlso cc.ContainsKey(remaining(0)), cc(remaining(0)), "#6b7280")

                        ' Push updated card to JS — renderBoard will remove it from the column
                        PushUpdatedCardToJs(entryId, mails, webView)
                    End If

                Case "open"
                    ' Open mail in Outlook Inspector
                    Dim entryId As String = CStr(msg("entryId"))
                    OpenMailInInspector(entryId)

                Case "reload"
                    ' Re-scan inbox and push new data, reusing the user's last load-count choice
                    summaryCts.Cancel()
                    Dim newCts As New CancellationTokenSource()
                    summaryCts = newCts

                    Dim reloadMax As Integer = 0
                    Dim newMails = CollectCategorizedInboxMails(reloadMax)
                    If newMails IsNot Nothing Then
                        mails.Clear()
                        mails.AddRange(newMails)
                        Dim newColumns = BuildBoardColumns(mails)
                        columns.Clear()
                        columns.AddRange(newColumns)

                        Dim dataJson As String = BuildBoardDataJson(mails, columns)
                        Try
                            webView.CoreWebView2.PostWebMessageAsJson(
                                JsonConvert.SerializeObject(New With {.action = "reloadData", .data = dataJson}))
                        Catch
                        End Try

                        ' Re-generate summaries
                        Dim capturedMails As List(Of InboxBoardEntry) = mails
                        Dim capturedWebView As WebView2 = webView
                        Dim capturedToken As CancellationToken = newCts.Token
                        System.Threading.Tasks.Task.Run(Async Function() As System.Threading.Tasks.Task
                                                            Await System.Threading.Tasks.Task.Delay(500)
                                                            frm.BeginInvoke(CType(Async Sub()
                                                                                      Await GenerateSummariesAsync(capturedMails, capturedWebView, capturedToken)
                                                                                  End Sub, System.Windows.Forms.MethodInvoker))
                                                        End Function)
                    End If

                Case "setPinnedColumns"
                    ' Save pinned column names to My.Settings
                    Try
                        Dim colsToken As JToken = msg("columns")
                        If colsToken IsNot Nothing AndAlso colsToken.Type = JTokenType.Array Then
                            Dim names As New List(Of String)()
                            For Each t As JToken In colsToken
                                Dim n As String = CStr(t)
                                If Not String.IsNullOrWhiteSpace(n) Then names.Add(n.Trim())
                            Next
                            My.Settings.InboxBoardColumns = String.Join(";", names)
                            My.Settings.Save()
                        End If
                    Catch ex2 As System.Exception
                        Debug.WriteLine($"[InboxBoard] setPinnedColumns error: {ex2.Message}")
                    End Try

                Case "setPinnedFlagColumns"
                    ' Save pinned flag bucket column names to My.Settings and reload
                    Try
                        Dim colsToken As JToken = msg("columns")
                        If colsToken IsNot Nothing AndAlso colsToken.Type = JTokenType.Array Then
                            Dim names As New List(Of String)()
                            For Each t As JToken In colsToken
                                Dim n As String = CStr(t)
                                If Not String.IsNullOrWhiteSpace(n) Then names.Add(n.Trim())
                            Next
                            My.Settings.InboxBoardPinnedFlagColumns = String.Join(";", names)
                            My.Settings.Save()
                        End If
                    Catch ex2 As System.Exception
                        Debug.WriteLine($"[InboxBoard] setPinnedFlagColumns error: {ex2.Message}")
                    End Try

                    ' Trigger a reload to apply the change
                    HandleBoardReload(mails, columns, webView, frm, summaryCts)

                Case "setColumnOrder"
                    ' Save user's column order to My.Settings
                    Try
                        Dim colsToken As JToken = msg("columns")
                        If colsToken IsNot Nothing AndAlso colsToken.Type = JTokenType.Array Then
                            Dim names As New List(Of String)()
                            For Each t As JToken In colsToken
                                Dim n As String = CStr(t)
                                If Not String.IsNullOrWhiteSpace(n) Then names.Add(n.Trim())
                            Next
                            My.Settings.InboxBoardColumnOrder = String.Join(";", names)
                            My.Settings.Save()

                            ' Update in-memory column order to match
                            Dim newOrder As New List(Of InboxBoardColumn)()
                            For Each name In names
                                Dim col = columns.FirstOrDefault(Function(c) c.CategoryName.Equals(name, StringComparison.OrdinalIgnoreCase))
                                If col IsNot Nothing Then newOrder.Add(col)
                            Next
                            ' Append any columns not in the saved order
                            For Each col In columns
                                If Not newOrder.Any(Function(c) c.CategoryName.Equals(col.CategoryName, StringComparison.OrdinalIgnoreCase)) Then
                                    newOrder.Add(col)
                                End If
                            Next
                            columns.Clear()
                            columns.AddRange(newOrder)
                        End If
                    Catch ex2 As System.Exception
                        Debug.WriteLine($"[InboxBoard] setColumnOrder error: {ex2.Message}")
                    End Try

                Case "setTheme"
                    ' Persist theme choice to My.Settings
                    Try
                        Dim theme As String = CStr(msg("theme"))
                        If Not String.IsNullOrEmpty(theme) Then
                            My.Settings.InboxBoardTheme = theme
                            My.Settings.Save()
                        End If
                    Catch
                    End Try

                Case "setFieldSettings"
                    ' Persist card field toggles as JSON string
                    Try
                        Dim fields As String = CStr(msg("fields"))
                        If Not String.IsNullOrEmpty(fields) Then
                            My.Settings.InboxBoardFields = fields
                            My.Settings.Save()
                        End If
                    Catch
                    End Try

                Case "setColumnFilter"
                    ' Persist column filter selection
                    Try
                        Dim filterVal As String = CStr(msg("filter"))
                        My.Settings.InboxBoardColumnFilter = If(filterVal, "")
                        My.Settings.Save()
                    Catch
                    End Try

                Case "setGroupConversations"
                    ' Persist conversation grouping toggle and reload
                    Try
                        Dim enabled As Boolean = CBool(msg("enabled"))
                        My.Settings.InboxBoardGroupConversations = enabled
                        My.Settings.Save()
                    Catch
                    End Try

                    ' Trigger a reload to apply the grouping change
                    HandleBoardReload(mails, columns, webView, frm, summaryCts)

                Case "setIncludeFlagged"
                    ' Persist include-flagged toggle and reload
                    Try
                        Dim enabled As Boolean = CBool(msg("enabled"))
                        My.Settings.InboxBoardIncludeFlagged = enabled
                        My.Settings.Save()
                    Catch
                    End Try

                    ' Trigger a reload to apply the change
                    HandleBoardReload(mails, columns, webView, frm, summaryCts)

                Case "setHideDoneFlags"
                    ' Persist hide-done-flags toggle and reload
                    Try
                        Dim enabled As Boolean = CBool(msg("enabled"))
                        My.Settings.InboxBoardHideDoneFlags = enabled
                        My.Settings.Save()
                    Catch
                    End Try

                    ' Trigger a reload to apply the change
                    HandleBoardReload(mails, columns, webView, frm, summaryCts)

                Case "setGroupFlagsByDate"
                    ' Persist group-flags-by-date toggle and reload
                    Try
                        Dim enabled As Boolean = CBool(msg("enabled"))
                        My.Settings.InboxBoardGroupFlagsByDate = enabled
                        My.Settings.Save()
                    Catch
                    End Try

                    ' Trigger a reload to apply the change
                    HandleBoardReload(mails, columns, webView, frm, summaryCts)

                Case "setSummaryLanguage"
                    ' Persist summary language selection; clear cache if language changed
                    Try
                        Dim lang As String = If(CStr(msg("language")), "").Trim()
                        Dim oldLang As String = ""
                        Try : oldLang = If(My.Settings.InboxBoardSummaryLanguage, "") : Catch : End Try

                        My.Settings.InboxBoardSummaryLanguage = lang
                        My.Settings.Save()

                        ' If the language actually changed, clear the cache and re-generate
                        If Not lang.Equals(oldLang, StringComparison.OrdinalIgnoreCase) Then
                            ClearSummaryCache()

                            ' Clear in-memory summaries so they will be re-generated
                            For Each m In mails
                                m.Summary = ""
                            Next

                            ' Push cleared summaries to JS so cards show "Generating…"
                            For Each m In mails
                                Try
                                    webView.CoreWebView2.PostWebMessageAsJson(
                                        JsonConvert.SerializeObject(New With {
                                            .action = "updateSummary",
                                            .entryId = m.EntryID,
                                            .summary = ""
                                        }))
                                Catch
                                End Try
                            Next

                            ' Restart summary generation with the new language
                            summaryCts.Cancel()
                            Dim newCtsLang As New CancellationTokenSource()
                            summaryCts = newCtsLang

                            Dim capturedMailsLang As List(Of InboxBoardEntry) = mails
                            Dim capturedWebViewLang As WebView2 = webView
                            Dim capturedTokenLang As CancellationToken = newCtsLang.Token
                            System.Threading.Tasks.Task.Run(Async Function() As System.Threading.Tasks.Task
                                                                Await System.Threading.Tasks.Task.Delay(500)
                                                                frm.BeginInvoke(CType(Async Sub()
                                                                                          Await GenerateSummariesAsync(capturedMailsLang, capturedWebViewLang, capturedTokenLang)
                                                                                      End Sub, System.Windows.Forms.MethodInvoker))
                                                            End Function)
                        End If
                    Catch
                    End Try
                Case "moveToFolder"
                    ' Drag card into a board folder
                    Dim entryId As String = CStr(msg("entryId"))
                    Dim folderName As String = CStr(msg("folderName"))
                    If String.IsNullOrWhiteSpace(folderName) Then Return

                    Dim entry = mails.FirstOrDefault(Function(m) m.EntryID = entryId)
                    If entry IsNot Nothing Then
                        Dim idsToUpdate As List(Of String) = If(entry.IsGrouped AndAlso entry.ConversationEntryIDs IsNot Nothing AndAlso entry.ConversationEntryIDs.Count > 1,
                                                                 entry.ConversationEntryIDs,
                                                                 New List(Of String) From {entryId})
                        For Each id In idsToUpdate
                            SetMailBoardFolder(id, folderName)
                        Next
                        entry.BoardFolder = folderName
                    End If

                    ' Push full reload so folder tile counts update and card moves
                    Dim dataJsonFolder As String = BuildBoardDataJson(mails, columns)
                    Try
                        webView.CoreWebView2.PostWebMessageAsJson(
                            JsonConvert.SerializeObject(New With {.action = "reloadData", .data = dataJsonFolder}))
                    Catch
                    End Try

                Case "removeFromFolder"
                    ' Drag card out of a folder (or click remove)
                    Dim entryId As String = CStr(msg("entryId"))

                    Dim entry = mails.FirstOrDefault(Function(m) m.EntryID = entryId)
                    If entry IsNot Nothing Then
                        Dim idsToUpdate As List(Of String) = If(entry.IsGrouped AndAlso entry.ConversationEntryIDs IsNot Nothing AndAlso entry.ConversationEntryIDs.Count > 1,
                                                                 entry.ConversationEntryIDs,
                                                                 New List(Of String) From {entryId})
                        For Each id In idsToUpdate
                            SetMailBoardFolder(id, "")
                        Next
                        entry.BoardFolder = ""
                    End If

                    Dim dataJsonUnfolder As String = BuildBoardDataJson(mails, columns)
                    Try
                        webView.CoreWebView2.PostWebMessageAsJson(
                            JsonConvert.SerializeObject(New With {.action = "reloadData", .data = dataJsonUnfolder}))
                    Catch
                    End Try

                Case "createFolder"
                    Dim folderName As String = If(CStr(msg("name")), "").Trim()
                    If String.IsNullOrWhiteSpace(folderName) Then Return

                    Dim folders = GetBoardFolderNames()
                    If folders.Count >= InboxBoard_MaxFolders Then Return
                    If folders.Any(Function(f) f.Equals(folderName, StringComparison.OrdinalIgnoreCase)) Then Return

                    folders.Add(folderName)
                    SaveBoardFolderNames(folders)

                    Dim dataJsonCreate As String = BuildBoardDataJson(mails, columns)
                    Try
                        webView.CoreWebView2.PostWebMessageAsJson(
                            JsonConvert.SerializeObject(New With {.action = "reloadData", .data = dataJsonCreate}))
                    Catch
                    End Try

                Case "renameFolder"
                    Dim oldName As String = If(CStr(msg("oldName")), "").Trim()
                    Dim newName As String = If(CStr(msg("newName")), "").Trim()
                    If String.IsNullOrWhiteSpace(oldName) OrElse String.IsNullOrWhiteSpace(newName) Then Return

                    RenameBoardFolder(oldName, newName, mails)

                    Dim dataJsonRename As String = BuildBoardDataJson(mails, columns)
                    Try
                        webView.CoreWebView2.PostWebMessageAsJson(
                            JsonConvert.SerializeObject(New With {.action = "reloadData", .data = dataJsonRename}))
                    Catch
                    End Try

                Case "deleteFolder"
                    Dim folderName As String = If(CStr(msg("name")), "").Trim()
                    If String.IsNullOrWhiteSpace(folderName) Then Return

                    DeleteBoardFolder(folderName, mails)

                    Dim dataJsonDelete As String = BuildBoardDataJson(mails, columns)
                    Try
                        webView.CoreWebView2.PostWebMessageAsJson(
                            JsonConvert.SerializeObject(New With {.action = "reloadData", .data = dataJsonDelete}))
                    Catch
                    End Try

                Case "setLoadThreshold"
                    ' Persist user-configurable load threshold
                    Try
                        Dim threshold As Integer = CInt(msg("threshold"))
                        If threshold < 0 Then threshold = 0
                        My.Settings.InboxBoardLoadThreshold = threshold
                        My.Settings.Save()
                        ' Reset remembered load count so next reload uses the new threshold
                        _inboxBoardLastMaxToLoad = 0
                    Catch
                    End Try

                Case "toggleFolderExpand"
                    Dim folderName As String = If(CStr(msg("name")), "").Trim()
                    Dim expanded As Boolean = CBool(msg("expanded"))
                    Dim expandedSet = GetExpandedFolders()
                    If expanded Then
                        expandedSet.Add(folderName)
                    Else
                        expandedSet.Remove(folderName)
                    End If
                    SaveExpandedFolders(expandedSet)

                Case "changeTargetFolder"
                    ' Prompt user to select a new mailbox and folder
                    Dim changed As Boolean = PromptSelectTargetFolder()
                    If changed Then
                        ' Trigger a full reload with the new folder
                        HandleBoardReload(mails, columns, webView, frm, summaryCts)
                        ' Push the updated folder label to JS
                        Try
                            webView.CoreWebView2.PostWebMessageAsJson(
                                JsonConvert.SerializeObject(New With {
                                    .action = "updateFolderLabel",
                                    .label = _inboxBoardTrackedFolderLabel
                                }))
                        Catch
                        End Try
                    End If

                Case "resetTargetFolder"
                    ' Reset to default Inbox
                    ResetInboxBoardTargetFolder()
                    HandleBoardReload(mails, columns, webView, frm, summaryCts)
                    Try
                        webView.CoreWebView2.PostWebMessageAsJson(
                            JsonConvert.SerializeObject(New With {
                                .action = "updateFolderLabel",
                                .label = _inboxBoardTrackedFolderLabel
                            }))
                    Catch
                    End Try

                Case "setIncludeSubfolders"
                    ' Persist include-subfolders toggle and reload
                    Try
                        Dim enabled As Boolean = CBool(msg("enabled"))
                        My.Settings.InboxBoardIncludeSubfolders = enabled
                        My.Settings.Save()
                    Catch
                    End Try
                    HandleBoardReload(mails, columns, webView, frm, summaryCts)

            End Select

        Catch ex As System.Exception
            Debug.WriteLine($"[InboxBoard] HandleBoardMessage error: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Helper to trigger a full board reload (used by multiple settings toggles).
    ''' </summary>
    Private Sub HandleBoardReload(
        mails As List(Of InboxBoardEntry),
        columns As List(Of InboxBoardColumn),
        webView As WebView2,
        frm As Form,
        ByRef summaryCts As CancellationTokenSource)

        summaryCts.Cancel()
        Dim newCts As New CancellationTokenSource()
        summaryCts = newCts

        Dim reloadMax As Integer = 0
        Dim newMails = CollectCategorizedInboxMails(reloadMax)
        If newMails IsNot Nothing Then
            mails.Clear()
            mails.AddRange(newMails)
            Dim newColumns = BuildBoardColumns(mails)
            columns.Clear()
            columns.AddRange(newColumns)
            Dim dataJson As String = BuildBoardDataJson(mails, columns)
            Try
                webView.CoreWebView2.PostWebMessageAsJson(
                    JsonConvert.SerializeObject(New With {.action = "reloadData", .data = dataJson}))
            Catch
            End Try
            Dim capturedMails As List(Of InboxBoardEntry) = mails
            Dim capturedWebView As WebView2 = webView
            Dim capturedToken As CancellationToken = newCts.Token
            System.Threading.Tasks.Task.Run(Async Function() As System.Threading.Tasks.Task
                                                Await System.Threading.Tasks.Task.Delay(500)
                                                frm.BeginInvoke(CType(Async Sub()
                                                                          Await GenerateSummariesAsync(capturedMails, capturedWebView, capturedToken)
                                                                      End Sub, System.Windows.Forms.MethodInvoker))
                                            End Function)
        End If
    End Sub

    ''' <summary>
    ''' Pushes updated card data (allCategories, category, color, flag info) to JS after a move/moveAdd.
    ''' </summary>
    Private Sub PushUpdatedCardToJs(entryId As String, mails As List(Of InboxBoardEntry), webView As WebView2)
        Try
            Dim entry = mails.FirstOrDefault(Function(m) m.EntryID = entryId)
            If entry Is Nothing Then Return

            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
            Dim categoryColors = GetOutlookCategoryColors(ns)

            Dim allCatsArr As New JArray()
            If Not String.IsNullOrEmpty(entry.AllCategories) Then
                For Each catPart In entry.AllCategories.Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
                    Dim catName As String = catPart.Trim()
                    If Not String.IsNullOrEmpty(catName) Then
                        Dim catObj As New JObject()
                        catObj("name") = catName
                        catObj("color") = If(categoryColors.ContainsKey(catName), categoryColors(catName), "#6b7280")
                        allCatsArr.Add(catObj)
                    End If
                Next
            End If

            webView.CoreWebView2.PostWebMessageAsJson(
                JsonConvert.SerializeObject(New With {
                    .action = "updateCard",
                    .entryId = entryId,
                    .category = entry.Categories,
                    .categoryColor = entry.CategoryColor,
                    .allCategories = allCatsArr,
                    .isFlagged = entry.IsFlagged,
                    .isCompleted = entry.IsCompleted,
                    .flagRequest = If(entry.FlagRequest, ""),
                    .flagDueDate = If(entry.FlagDueDate.HasValue, entry.FlagDueDate.Value.ToString("yyyy-MM-dd"), ""),
                    .flagBucket = If(entry.FlagBucket, ""),
                    .boardFolder = If(entry.BoardFolder, "")
                }))
        Catch
        End Try
    End Sub

#End Region

#Region "InboxBoard Outlook COM Operations"

    ''' <summary>
    ''' Sets the TaskDueDate on a flagged mail item.
    ''' Pass Nothing to clear the due date.
    ''' </summary>
    Private Sub SetMailDueDate(entryId As String, dueDate As DateTime?)
        Try
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
            Dim mi As MailItem = TryCast(ComRetry(Of Object)(Function() ns.GetItemFromID(entryId)), MailItem)
            If mi IsNot Nothing Then
                If dueDate.HasValue Then
                    mi.TaskDueDate = dueDate.Value
                    mi.TaskStartDate = dueDate.Value
                Else
                    ' Set to Outlook's "no date" sentinel (1/1/4501)
                    mi.TaskDueDate = #1/1/4501#
                    mi.TaskStartDate = #1/1/4501#
                End If
                mi.Save()
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"[InboxBoard] SetMailDueDate error: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Returns a representative due date for a given flag bucket column ID.
    ''' Used when dragging a card into a date-bucketed column.
    ''' Returns Nothing for the "no due date" / generic flagged bucket.
    ''' </summary>
    Private Shared Function GetDueDateForBucket(bucketId As String) As DateTime?
        Dim today As DateTime = DateTime.Today

        Select Case bucketId
            Case InboxBoard_FlagBucket_Overdue
                ' Set to yesterday (it's already overdue)
                Return today.AddDays(-1)
            Case InboxBoard_FlagBucket_Today
                Return today
            Case InboxBoard_FlagBucket_Tomorrow
                Return today.AddDays(1)
            Case InboxBoard_FlagBucket_ThisWeek
                ' Set to end of current week (Friday)
                Dim daysUntilFriday As Integer = (5 - CInt(today.DayOfWeek) + 7) Mod 7
                If daysUntilFriday = 0 Then daysUntilFriday = 5 ' If today is Friday, keep Friday
                Return today.AddDays(daysUntilFriday)
            Case InboxBoard_FlagBucket_NextWeek
                ' Set to next Monday + 4 (next Friday)
                Dim daysUntilNextMonday As Integer = ((1 - CInt(today.DayOfWeek)) + 7) Mod 7
                If daysUntilNextMonday = 0 Then daysUntilNextMonday = 7
                Return today.AddDays(daysUntilNextMonday + 4) ' Next Friday
            Case InboxBoard_FlagBucket_Next4Weeks
                Return today.AddDays(21) ' 3 weeks out
            Case InboxBoard_FlagBucket_Later
                Return today.AddDays(42) ' 6 weeks out
            Case InboxBoard_FlagBucket_Done
                Return Nothing ' Don't change due date for completed items
            Case Else
                Return Nothing ' Generic flagged / no due date
        End Select
    End Function

    ''' <summary>
    ''' Changes the category on a mail item (replaces all existing categories).
    ''' </summary>
    Private Sub UpdateMailCategory(entryId As String, newCategory As String)
        Try
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
            Dim mi As MailItem = TryCast(ComRetry(Of Object)(Function() ns.GetItemFromID(entryId)), MailItem)
            If mi IsNot Nothing Then
                mi.Categories = newCategory
                mi.Save()
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"[InboxBoard] UpdateMailCategory error: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Adds a category to a mail item, keeping existing categories.
    ''' </summary>
    Private Sub AddMailCategory(entryId As String, categoryToAdd As String)
        Try
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
            Dim mi As MailItem = TryCast(ComRetry(Of Object)(Function() ns.GetItemFromID(entryId)), MailItem)
            If mi IsNot Nothing Then
                Dim existing As String = ComRetry(Of String)(Function() If(mi.Categories, ""))
                Dim cats As New List(Of String)()
                If Not String.IsNullOrEmpty(existing) Then
                    For Each part In existing.Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
                        Dim t = part.Trim()
                        If Not String.IsNullOrEmpty(t) Then cats.Add(t)
                    Next
                End If
                If Not cats.Any(Function(c) c.Equals(categoryToAdd, StringComparison.OrdinalIgnoreCase)) Then
                    cats.Add(categoryToAdd)
                End If
                mi.Categories = String.Join(", ", cats)
                mi.Save()
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"[InboxBoard] AddMailCategory error: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Removes a single category from a mail item, keeping all other categories.
    ''' </summary>
    Private Sub RemoveMailCategory(entryId As String, categoryToRemove As String)
        Try
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
            Dim mi As MailItem = TryCast(ComRetry(Of Object)(Function() ns.GetItemFromID(entryId)), MailItem)
            If mi IsNot Nothing Then
                Dim existing As String = ComRetry(Of String)(Function() If(mi.Categories, ""))
                Dim cats As New List(Of String)()
                If Not String.IsNullOrEmpty(existing) Then
                    For Each part In existing.Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
                        Dim t = part.Trim()
                        If Not String.IsNullOrEmpty(t) AndAlso
                           Not t.Equals(categoryToRemove, StringComparison.OrdinalIgnoreCase) Then
                            cats.Add(t)
                        End If
                    Next
                End If
                mi.Categories = String.Join(", ", cats)
                mi.Save()
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"[InboxBoard] RemoveMailCategory error: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Clears all categories from a mail item.
    ''' </summary>
    Private Sub ClearMailCategories(entryId As String)
        Try
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
            Dim mi As MailItem = TryCast(ComRetry(Of Object)(Function() ns.GetItemFromID(entryId)), MailItem)
            If mi IsNot Nothing Then
                mi.Categories = ""
                mi.Save()
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"[InboxBoard] ClearMailCategories error: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Sets or clears the follow-up flag on a mail item.
    ''' When setting, uses olFlagMarked with "Follow up" request.
    ''' When clearing, calls ClearTaskFlag to remove all flag/task properties.
    ''' </summary>
    Private Sub SetMailFlag(entryId As String, flagOn As Boolean)
        Try
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
            Dim mi As MailItem = TryCast(ComRetry(Of Object)(Function() ns.GetItemFromID(entryId)), MailItem)
            If mi IsNot Nothing Then
                If flagOn Then
                    mi.FlagStatus = OlFlagStatus.olFlagMarked
                    mi.FlagRequest = "Follow up"
                Else
                    mi.ClearTaskFlag()
                End If
                mi.Save()
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"[InboxBoard] SetMailFlag error: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Marks a mail item's flag as complete (olFlagComplete).
    ''' </summary>
    Private Sub SetMailFlagComplete(entryId As String)
        Try
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
            Dim mi As MailItem = TryCast(ComRetry(Of Object)(Function() ns.GetItemFromID(entryId)), MailItem)
            If mi IsNot Nothing Then
                mi.FlagStatus = OlFlagStatus.olFlagComplete
                mi.Save()
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"[InboxBoard] SetMailFlagComplete error: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Opens a mail in an Outlook Inspector window.
    ''' </summary>
    Private Sub OpenMailInInspector(entryId As String)
        Try
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
            Dim mi As MailItem = TryCast(ComRetry(Of Object)(Function() ns.GetItemFromID(entryId)), MailItem)
            If mi IsNot Nothing Then
                mi.Display(False)
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"[InboxBoard] OpenMailInInspector error: {ex.Message}")
        End Try
    End Sub

#End Region

#Region "InboxBoard AI Summaries"

    ''' <summary>
    ''' Returns the language instruction to append to the summary system prompt.
    ''' Uses the free-text language field; empty means auto-detect.
    ''' </summary>
    Private Function GetSummaryLanguageInstruction() As String
        Dim lang As String = ""
        Try : lang = If(My.Settings.InboxBoardSummaryLanguage, "").Trim() : Catch : End Try

        ' If empty, treat as auto-detect — no explicit instruction
        If String.IsNullOrEmpty(lang) Then Return ""

        Return $" Write all summaries in {lang}."
    End Function

    ''' <summary>
    ''' Generates AI summaries for mails in batches and pushes them to the board.
    ''' Skips mails that already have a cached summary.
    ''' </summary>
    Private Async Function GenerateSummariesAsync(
        mails As List(Of InboxBoardEntry),
        webView As WebView2,
        ct As CancellationToken) As System.Threading.Tasks.Task

        Dim unsummarized = mails.Where(Function(m) String.IsNullOrEmpty(m.Summary) AndAlso
                                                    Not String.IsNullOrEmpty(m.BodyExcerpt)).ToList()
        If unsummarized.Count = 0 Then Return

        Dim langInstruction As String = GetSummaryLanguageInstruction()
        Dim batchIndex As Integer = 0

        While batchIndex < unsummarized.Count
            If ct.IsCancellationRequested Then Return

            Dim batchStart As Integer = batchIndex
            Dim batchEnd As Integer = Math.Min(batchIndex + InboxBoard_SummaryBatchSize - 1, unsummarized.Count - 1)

            Dim userPrompt As New StringBuilder()
            For i As Integer = batchStart To batchEnd
                Dim mail As InboxBoardEntry = unsummarized(i)
                Dim mailNum As Integer = i - batchStart + 1
                userPrompt.AppendLine($"<EMAIL id=""{mailNum}"">")
                userPrompt.AppendLine($"Subject: {mail.Subject}")
                userPrompt.AppendLine($"From: {mail.SenderName}")
                userPrompt.AppendLine($"Date: {mail.ReceivedTime:yyyy-MM-dd}")
                If Not String.IsNullOrEmpty(mail.BodyExcerpt) Then
                    userPrompt.AppendLine($"Body: {mail.BodyExcerpt}")
                End If
                userPrompt.AppendLine("</EMAIL>")
            Next

            Try
                Dim systemPrompt As String = InterpolateAtRuntime(SP_InboxBoard) & langInstruction
                Dim response As String = Await LLM(
                    systemPrompt, userPrompt.ToString(),
                    "", "", 0, False, True)

                If Not String.IsNullOrWhiteSpace(response) Then
                    ParseAndApplySummaries(response, unsummarized, batchStart, batchEnd, webView)
                End If
            Catch ex As System.Exception
                Debug.WriteLine($"[InboxBoard] Summary batch error: {ex.Message}")
            End Try

            batchIndex = batchEnd + 1
        End While

        ' Persist cache after all batches are done
        SaveSummaryCache()
    End Function

    ''' <summary>
    ''' Parses summary JSON and pushes updates to the board via postMessage.
    ''' Also caches summaries for future use.
    ''' </summary>
    Private Sub ParseAndApplySummaries(
        response As String,
        mails As List(Of InboxBoardEntry),
        batchStart As Integer, batchEnd As Integer,
        webView As WebView2)

        Try
            Dim jsonText As String = response.Trim()
            If jsonText.StartsWith("```") Then
                Dim startIdx As Integer = jsonText.IndexOf("[")
                Dim endIdx As Integer = jsonText.LastIndexOf("]")
                If startIdx >= 0 AndAlso endIdx > startIdx Then
                    jsonText = jsonText.Substring(startIdx, endIdx - startIdx + 1)
                End If
            End If

            Dim arr As JArray = JArray.Parse(jsonText)

            For Each item As JObject In arr
                Dim id As Integer = If(item("id") IsNot Nothing, CInt(item("id")), 0)
                Dim summary As String = If(item("summary") IsNot Nothing, CStr(item("summary")), "")
                If String.IsNullOrEmpty(summary) Then Continue For

                Dim absoluteIndex As Integer = batchStart + id - 1
                If absoluteIndex < batchStart OrElse absoluteIndex > batchEnd OrElse absoluteIndex >= mails.Count Then Continue For

                mails(absoluteIndex).Summary = summary

                ' Cache the summary
                CacheSummary(mails(absoluteIndex).EntryID, summary)

                ' Push to JS
                Try
                    Dim entryId As String = mails(absoluteIndex).EntryID
                    webView.CoreWebView2.PostWebMessageAsJson(
                        JsonConvert.SerializeObject(New With {
                            .action = "updateSummary",
                            .entryId = entryId,
                            .summary = summary
                        }))
                Catch
                End Try
            Next
        Catch ex As System.Exception
            Debug.WriteLine($"[InboxBoard] ParseAndApplySummaries error: {ex.Message}")
        End Try
    End Sub

#End Region

#Region "InboxBoard JSON Helpers"

    ''' <summary>
    ''' Builds a JSON string with card and column data for JS initialization.
    ''' </summary>
    Private Function BuildBoardDataJson(mails As List(Of InboxBoardEntry), columns As List(Of InboxBoardColumn)) As String
        Dim categoryColors As Dictionary(Of String, String) = Nothing

        Dim cardsArr As New JArray()
        For i As Integer = 0 To mails.Count - 1
            Dim m = mails(i)
            Dim card As New JObject()
            card("entryId") = m.EntryID
            card("subject") = If(m.Subject, "")
            card("senderName") = If(m.SenderName, "")
            card("senderEmail") = If(m.SenderEmail, "")
            card("recipients") = If(m.Recipients, "")
            card("date") = m.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
            card("messages") = m.MessageCount
            card("isGrouped") = m.IsGrouped
            card("category") = If(m.Categories, "")
            card("categoryColor") = If(m.CategoryColor, "#6b7280")
            card("summary") = If(m.Summary, "")
            card("isRead") = m.IsRead
            card("isConversation") = m.IsGrouped
            card("isFlagged") = m.IsFlagged
            card("isCompleted") = m.IsCompleted
            card("flagRequest") = If(m.FlagRequest, "")
            card("flagDueDate") = If(m.FlagDueDate.HasValue, m.FlagDueDate.Value.ToString("yyyy-MM-dd"), "")
            card("flagBucket") = If(m.FlagBucket, "")
            card("boardFolder") = If(m.BoardFolder, "")

            ' Build array of all categories with colors for multi-category badges
            Dim allCatsArr As New JArray()
            If Not String.IsNullOrEmpty(m.AllCategories) Then
                If categoryColors Is Nothing Then
                    Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
                    Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
                    categoryColors = GetOutlookCategoryColors(ns)
                End If
                For Each catPart In m.AllCategories.Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
                    Dim catName As String = catPart.Trim()
                    If Not String.IsNullOrEmpty(catName) Then
                        Dim catObj As New JObject()
                        catObj("name") = catName
                        catObj("color") = If(categoryColors.ContainsKey(catName), categoryColors(catName), "#6b7280")
                        allCatsArr.Add(catObj)
                    End If
                Next
            End If
            card("allCategories") = allCatsArr

            cardsArr.Add(card)
        Next

        Dim colsArr As New JArray()
        For Each col In columns
            Dim c As New JObject()
            c("id") = col.CategoryName
            c("title") = col.CategoryName
            c("tag") = If(col.Tag, col.CategoryName)
            c("color") = If(col.Color, "#6b7280")
            c("isFlaggedColumn") = col.IsFlaggedColumn
            colsArr.Add(c)
        Next

        Dim root As New JObject()
        root("cards") = cardsArr
        root("columns") = colsArr

        ' Board folders
        Dim foldersArr As New JArray()
        Dim folderNames = GetBoardFolderNames()
        Dim expandedFolders = GetExpandedFolders()
        For Each fn In folderNames
            Dim fo As New JObject()
            fo("name") = fn
            fo("count") = mails.Where(Function(m) Not String.IsNullOrEmpty(m.BoardFolder) AndAlso m.BoardFolder.Equals(fn, StringComparison.OrdinalIgnoreCase)).Count()
            fo("expanded") = expandedFolders.Contains(fn)
            foldersArr.Add(fo)
        Next
        root("folders") = foldersArr

        root("trackedFolder") = If(_inboxBoardTrackedFolderLabel, "Inbox")

        Return root.ToString(Formatting.None)
    End Function

#End Region

#Region "InboxBoard HTML Builder"

    ''' <summary>
    ''' Builds the complete HTML page for the Inbox Board.
    ''' </summary>
    Private Function BuildInboxBoardHtml(mails As List(Of InboxBoardEntry), columns As List(Of InboxBoardColumn)) As String
        Dim dataJson As String = BuildBoardDataJson(mails, columns)

        ' Read persisted settings to inject into JS as constants
        Dim savedTheme As String = ""
        Dim savedFields As String = ""
        Dim savedColumnFilter As String = ""
        Dim savedGroupConversations As Boolean = False
        Dim savedSummaryLanguage As String = ""
        Dim savedIncludeFlagged As Boolean = False
        Dim savedHideDoneFlags As Boolean = True
        Dim savedGroupFlagsByDate As Boolean = False
        Dim savedIncludeSubfolders As Boolean = True
        Try : savedIncludeSubfolders = My.Settings.InboxBoardIncludeSubfolders : Catch : End Try
        Try : savedTheme = If(My.Settings.InboxBoardTheme, "") : Catch : End Try
        Try : savedFields = If(My.Settings.InboxBoardFields, "") : Catch : End Try
        Try : savedColumnFilter = If(My.Settings.InboxBoardColumnFilter, "") : Catch : End Try
        Try : savedGroupConversations = My.Settings.InboxBoardGroupConversations : Catch : End Try
        Try : savedSummaryLanguage = If(My.Settings.InboxBoardSummaryLanguage, "") : Catch : End Try
        Try : savedIncludeFlagged = My.Settings.InboxBoardIncludeFlagged : Catch : End Try
        Try : savedHideDoneFlags = My.Settings.InboxBoardHideDoneFlags : Catch : End Try
        Try : savedGroupFlagsByDate = My.Settings.InboxBoardGroupFlagsByDate : Catch : End Try

        ' If no saved language yet, default to INI_Language1 (full language name, e.g. "German")
        ' but do NOT persist it — let the user confirm by editing the field first.
        ' This way the field shows a sensible default without locking it in.
        If String.IsNullOrEmpty(savedSummaryLanguage) Then
            Try
                Dim lang1 As String = If(_context.INI_Language1, "")
                If Not String.IsNullOrEmpty(lang1) Then
                    savedSummaryLanguage = lang1
                End If
            Catch
            End Try
        End If

        ' Build folder definitions for JS
        Dim savedFoldersJson As String = "[]"
        Try
            Dim folderNames = GetBoardFolderNames()
            Dim expandedFolders = GetExpandedFolders()
            Dim fArr As New JArray()
            For Each fn In folderNames
                Dim fo As New JObject()
                fo("name") = fn
                fo("expanded") = expandedFolders.Contains(fn)
                fArr.Add(fo)
            Next
            savedFoldersJson = fArr.ToString(Formatting.None)
        Catch
        End Try

        Dim sb As New StringBuilder(32000)
        sb.AppendLine("<!DOCTYPE html>")
        sb.AppendLine("<html lang=""en"">")
        sb.AppendLine("<head>")
        sb.AppendLine("<meta charset=""UTF-8"">")
        sb.AppendLine("<meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">")
        sb.AppendLine($"<title>{AN} - Inbox Board</title>")
        sb.AppendLine("<style>")
        sb.AppendLine(GetBoardCss())
        sb.AppendLine("</style>")
        sb.AppendLine("</head>")
        sb.AppendLine("<body>")
        sb.AppendLine(GetBoardHeaderHtml())
        sb.AppendLine("<div class=""board"" id=""board""></div>")
        sb.AppendLine("<div class=""toast-container"" id=""toastContainer""></div>")
        sb.AppendLine("<div class=""drag-mode-label"" id=""dragModeLabel""><span class=""plus"">+</span> Add category (Ctrl held)</div>")
        sb.AppendLine("<script>")
        sb.AppendLine($"const INIT_DATA = {dataJson};")
        sb.AppendLine($"const SAVED_THEME = {JsonConvert.SerializeObject(savedTheme)};")
        sb.AppendLine($"const SAVED_FIELDS = {JsonConvert.SerializeObject(savedFields)};")
        sb.AppendLine($"const SAVED_COLUMN_FILTER = {JsonConvert.SerializeObject(savedColumnFilter)};")
        sb.AppendLine($"const SAVED_GROUP_CONVERSATIONS = {If(savedGroupConversations, "true", "false")};")
        sb.AppendLine($"const SAVED_SUMMARY_LANGUAGE = {JsonConvert.SerializeObject(savedSummaryLanguage)};")
        sb.AppendLine($"const SAVED_INCLUDE_FLAGGED = {If(savedIncludeFlagged, "true", "false")};")
        sb.AppendLine($"const SAVED_HIDE_DONE_FLAGS = {If(savedHideDoneFlags, "true", "false")};")
        sb.AppendLine($"const SAVED_GROUP_FLAGS_BY_DATE = {If(savedGroupFlagsByDate, "true", "false")};")
        sb.AppendLine($"const FLAGGED_COLUMN_ID = {JsonConvert.SerializeObject(InboxBoard_FlaggedColumnId)};")
        sb.AppendLine($"const SAVED_FOLDERS = {savedFoldersJson};")
        sb.AppendLine($"const SAVED_INCLUDE_SUBFOLDERS = {If(savedIncludeSubfolders, "true", "false")};")
        sb.AppendLine($"const SAVED_TRACKED_FOLDER = {JsonConvert.SerializeObject(_inboxBoardTrackedFolderLabel)};")
        sb.AppendLine(GetBoardJavaScript())
        sb.AppendLine("</script>")
        sb.AppendLine("</body>")
        sb.AppendLine("</html>")

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Returns the full CSS stylesheet used by the Inbox Board HTML host.
    ''' </summary>
    Private Function GetBoardCss() As String
        Return "
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
:root {
  --bg: #f0f2f5; --header-bg: #ffffff; --header-border: #e0e0e0;
  --col-bg: #f7f8fa; --col-border: #e2e4e8; --card-bg: #ffffff;
  --card-border: #e0e0e0; --card-shadow: rgba(0,0,0,0.06);
  --text: #1a1a1a; --text-secondary: #555; --text-muted: #888;
  --input-bg: #f0f2f5; --input-border: #d0d0d0;
  --drop-highlight: rgba(59,130,246,0.12); --drop-line: #3b82f6;
  --drop-highlight-add: rgba(34,197,94,0.12); --drop-line-add: #22c55e;
  --toast-bg: #323232; --toast-text: #fff;
  --overlay-bg: rgba(0,0,0,0.05); --settings-bg: #ffffff; --settings-border: #e0e0e0;
}
[data-theme=""dark""] {
  --bg: #1a1b1e; --header-bg: #25262b; --header-border: #373a40;
  --col-bg: #25262b; --col-border: #373a40; --card-bg: #2c2e33;
  --card-border: #373a40; --card-shadow: rgba(0,0,0,0.2);
  --text: #e0e0e0; --text-secondary: #b0b0b0; --text-muted: #777;
  --input-bg: #2c2e33; --input-border: #444;
  --drop-highlight: rgba(59,130,246,0.2);
  --drop-highlight-add: rgba(34,197,94,0.2); --drop-line-add: #22c55e;
  --toast-bg: #444; --toast-text: #fff;
  --overlay-bg: rgba(0,0,0,0.2); --settings-bg: #2c2e33; --settings-border: #444;
}
body { font-family: system-ui, 'Segoe UI', Roboto, Arial, sans-serif;
  background: var(--bg); color: var(--text); min-height: 100vh; transition: background 0.3s, color 0.3s; }
.header { background: var(--header-bg); border-bottom: 1px solid var(--header-border);
  padding: 10px 20px; display: flex; align-items: center; gap: 12px; flex-wrap: wrap;
  position: sticky; top: 0; z-index: 100; transition: background 0.3s; }
.header .topline { display: flex; align-items: center; gap: 8px; min-width: 0; }
.header .topline img.logo { width: 24px; height: 24px; border-radius: 6px; display: block; }
.header .topline .brandbig { font-size: 18px; font-weight: 700; white-space: nowrap; }
.header .topline .sub { font-size: 15px; color: var(--text-muted); white-space: nowrap; }
.header-controls { display: flex; align-items: center; gap: 8px; flex: 1; justify-content: flex-end; flex-wrap: wrap; }
.search-input { padding: 6px 10px; border: 1px solid var(--input-border); border-radius: 6px;
  background: var(--input-bg); color: var(--text); font-size: 13px; width: 200px; outline: none; }
.search-input:focus { border-color: #3b82f6; }
.search-input::placeholder { color: var(--text-muted); }
.filter-select { padding: 6px 8px; border: 1px solid var(--input-border); border-radius: 6px;
  background: var(--input-bg); color: var(--text); font-size: 13px; outline: none; cursor: pointer; }
.icon-btn { width: 34px; height: 34px; border: 1px solid var(--input-border); border-radius: 6px;
  background: var(--input-bg); color: var(--text); cursor: pointer; display: flex;
  align-items: center; justify-content: center; font-size: 16px; position: relative; }
.icon-btn:hover { border-color: #3b82f6; }
.icon-btn:active { transform: scale(0.93); }
.settings-wrapper { position: relative; }
.settings-panel { position: absolute; top: 40px; right: 0; background: var(--settings-bg);
  border: 1px solid var(--settings-border); border-radius: 8px; padding: 10px 14px;
  min-width: 220px; max-height: 420px; overflow-y: auto;
  box-shadow: 0 4px 16px var(--card-shadow); z-index: 200; display: none; }
.settings-panel.open { display: block; }
.settings-panel h3 { font-size: 12px; font-weight: 600; margin-bottom: 6px;
  color: var(--text-secondary); text-transform: uppercase; letter-spacing: 0.5px; }
.settings-panel h3.section-divider { margin-top: 10px; padding-top: 8px;
  border-top: 1px solid var(--settings-border); }
.settings-panel label { display: flex; align-items: center; gap: 6px; padding: 3px 0;
  font-size: 13px; cursor: pointer; color: var(--text); }
.settings-panel input[type=""checkbox""] { width: 15px; height: 15px; accent-color: #3b82f6; cursor: pointer; }
.settings-panel select { padding: 4px 6px; border: 1px solid var(--input-border); border-radius: 4px;
  background: var(--input-bg); color: var(--text); font-size: 12px; outline: none; cursor: pointer; width: 100%; }
.settings-panel .cat-dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }
.lang-input-wrapper { position: relative; width: 100%; }
.lang-input-wrapper input[type=""text""] { padding: 4px 6px; border: 1px solid var(--input-border); border-radius: 4px;
  background: var(--input-bg); color: var(--text); font-size: 12px; outline: none; width: 100%; }
.lang-input-wrapper input[type=""text""]:focus { border-color: #3b82f6; }
.lang-input-wrapper .lang-overlay { position: absolute; top: 0; left: 0; right: 0; bottom: 0;
  padding: 4px 6px; font-size: 12px; color: var(--text-muted); pointer-events: none;
  display: flex; align-items: center; opacity: 0.7; }
.lang-input-wrapper input[type=""text""]:not([data-empty=""true""]) + .lang-overlay { display: none; }
.lang-input-wrapper input[type=""text""]:focus + .lang-overlay { display: none; }
.board { display: flex; gap: 14px; padding: 16px 20px; overflow-x: auto; overflow-y: hidden;
  min-height: calc(100vh - 56px); align-items: flex-start; cursor: default; }
.board.panning { cursor: grabbing; user-select: none; }
.column { min-width: 220px; max-width: 340px; flex: 0 0 280px; background: var(--col-bg);
  border: 1px solid var(--col-border); border-radius: 10px; display: flex; flex-direction: column;
  max-height: calc(100vh - 90px); }
.column.drag-over { background: var(--drop-highlight); }
.column.drag-over-add { background: var(--drop-highlight-add); }
.column-header { padding: 10px 12px; font-weight: 600; font-size: 13px; display: flex;
  align-items: center; gap: 6px; border-bottom: 1px solid var(--col-border); user-select: none;
  flex-shrink: 0; }
.column-header .dot { width: 9px; height: 9px; border-radius: 50%; flex-shrink: 0; }
.column-header .count { margin-left: auto; background: var(--input-bg); color: var(--text-muted);
  font-size: 11px; font-weight: 500; padding: 1px 7px; border-radius: 10px; }
.column-header .sort-btn { background: none; border: none; color: var(--text-muted); cursor: pointer;
  font-size: 14px; padding: 2px 4px; border-radius: 4px; display: flex; align-items: center;
  justify-content: center; position: relative; margin-left: 4px; flex-shrink: 0; }
.column-header .sort-btn:hover { color: #3b82f6; background: var(--drop-highlight); }
.sort-dropdown { position: absolute; top: 100%; right: 0; background: var(--settings-bg);
  border: 1px solid var(--settings-border); border-radius: 6px; padding: 4px 0;
  min-width: 160px; box-shadow: 0 4px 12px var(--card-shadow); z-index: 300; display: none; }
.sort-dropdown.open { display: block; }
.sort-dropdown .sort-option { padding: 5px 12px; font-size: 12px; color: var(--text);
  cursor: pointer; white-space: nowrap; display: flex; align-items: center; gap: 6px; }
.sort-dropdown .sort-option:hover { background: var(--drop-highlight); }
.sort-dropdown .sort-option .sort-check { width: 14px; font-size: 11px; color: #3b82f6; }
.card-list { padding: 6px; flex: 1; min-height: 50px; display: flex; flex-direction: column; gap: 0;
  overflow-y: auto; overflow-x: hidden; }
.card-list::-webkit-scrollbar { width: 5px; }
.card-list::-webkit-scrollbar-track { background: transparent; }
.card-list::-webkit-scrollbar-thumb { background: var(--text-muted); border-radius: 3px; opacity: 0.4; }
.card-list::-webkit-scrollbar-thumb:hover { opacity: 0.7; }
.drop-indicator { height: 3px; background: var(--drop-line); border-radius: 2px;
  margin: 2px 4px; opacity: 0; flex-shrink: 0; }
.drop-indicator.visible { opacity: 1; }
.drop-indicator.add-mode { background: var(--drop-line-add); }
.card { background: var(--card-bg); border: 1px solid var(--card-border); border-radius: 8px;
  padding: 9px 11px; cursor: grab; box-shadow: 0 1px 3px var(--card-shadow);
  user-select: none; margin-bottom: 6px; transition: box-shadow 0.2s, transform 0.15s; flex-shrink: 0; }
.card:hover { box-shadow: 0 3px 8px var(--card-shadow); transform: translateY(-1px); }
.card.dragging { opacity: 0.5; cursor: grabbing; transform: rotate(2deg); }
.card.hidden { display: none; }
.card.folder-hidden { display: none; }
.card.conversation { border-left: 3px solid #3b82f6; }
.card.flagged { border-left: 3px solid #ef4444; }
.card.conversation.flagged { border-left: 3px solid #ef4444; }
.card-subject { font-weight: 600; font-size: 13px; margin-bottom: 4px; color: var(--text); line-height: 1.3;
  overflow: hidden; text-overflow: ellipsis; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; }
.card-sender { font-size: 11px; color: var(--text-secondary); margin-bottom: 3px;
  overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.card-meta { display: flex; align-items: center; gap: 8px; font-size: 10px;
  color: var(--text-muted); margin-bottom: 3px; flex-wrap: wrap; }
.card-summary { font-size: 11px; color: var(--text-secondary); line-height: 1.3;
  margin-bottom: 4px; font-style: italic; min-height: 14px; }
.card-summary.loading { color: var(--text-muted); }
.card-footer { display: flex; justify-content: flex-end; gap: 4px; }
.card-btn { background: none; border: none; color: var(--text-muted); cursor: pointer;
  font-size: 13px; padding: 2px 4px; border-radius: 4px; display: flex;
  align-items: center; gap: 3px; position: relative; }
.card-btn:hover { color: #3b82f6; background: var(--drop-highlight); }
.card-btn .tip { position: absolute; bottom: 100%; right: 0; background: var(--toast-bg);
  color: var(--toast-text); font-size: 10px; padding: 3px 6px; border-radius: 4px;
  white-space: nowrap; opacity: 0; pointer-events: none; margin-bottom: 3px; }
.card-btn:hover .tip { opacity: 1; }
.card-btn.folder-remove-btn:hover { color: #ef4444; }
.unread-dot { width: 7px; height: 7px; border-radius: 50%; background: #3b82f6;
  display: inline-block; margin-right: 4px; flex-shrink: 0; }
.conv-icon { font-size: 10px; color: var(--text-muted); margin-right: 3px; }
.flag-badge { font-size: 10px; color: #ef4444; margin-right: 3px; }
.flag-meta { font-size: 10px; color: #ef4444; white-space: nowrap; }
.flag-meta.completed { color: #22c55e; text-decoration: line-through; }
.card-categories { display: flex; flex-wrap: wrap; gap: 3px; margin-bottom: 4px; }
.card-cat-badge { font-size: 9px; padding: 1px 6px; border-radius: 8px; color: #fff;
  white-space: nowrap; line-height: 1.4; font-weight: 500; }
.card-folder-badge { font-size: 10px; color: #b91c1c; margin-bottom: 3px; white-space: nowrap;
  overflow: hidden; text-overflow: ellipsis; }
.toast-container { position: fixed; bottom: 16px; right: 16px; display: flex;
  flex-direction: column-reverse; gap: 6px; z-index: 1000; pointer-events: none; }
.toast { background: var(--toast-bg); color: var(--toast-text); padding: 8px 14px;
  border-radius: 8px; font-size: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.2);
  max-width: 380px; line-height: 1.3; animation: toastIn 0.3s ease; pointer-events: auto; }
.toast.fade-out { animation: toastOut 0.3s ease forwards; }
@keyframes toastIn { from { opacity:0; transform:translateY(10px); } to { opacity:1; transform:translateY(0); } }
@keyframes toastOut { from { opacity:1; transform:translateY(0); } to { opacity:0; transform:translateY(10px); } }
.icon-svg { width: 15px; height: 15px; display: inline-block; vertical-align: middle; }
.drag-mode-label { position: fixed; bottom: 16px; left: 16px; background: var(--toast-bg);
  color: var(--toast-text); padding: 6px 12px; border-radius: 8px; font-size: 12px;
  box-shadow: 0 2px 8px rgba(0,0,0,0.2); z-index: 1000; pointer-events: none;
  display: none; }
.drag-mode-label.visible { display: block; }
.drag-mode-label .plus { color: #22c55e; font-weight: 700; margin-right: 4px; }
.column-header { cursor: grab; }
.column-header:active { cursor: grabbing; }
.column.col-dragging { opacity: 0.5; }
.column.col-drag-over { border-left: 3px solid #3b82f6; }
.column.folder-column { flex: 0 0 220px; min-width: 180px; max-width: 240px; }
.column.folder-column .column-header { cursor: default; }
.folder-tile { background: var(--col-bg); border: 2px dashed var(--col-border); border-radius: 8px;
  padding: 9px 11px; cursor: pointer; user-select: none; margin-bottom: 6px;
  transition: background 0.2s, border-color 0.2s; display: flex; align-items: center; gap: 8px; }
.folder-tile:hover { border-color: #dc2626; background: rgba(185,28,28,0.08); }
.folder-tile.drag-over { border-color: #dc2626; background: rgba(185,28,28,0.08); border-style: solid; }
.folder-tile.selected { border-color: #991b1b; background: rgba(153,27,27,0.15); border-style: solid; }
.folder-tile.show-all { border-color: var(--text-muted); }
.folder-tile.show-all:hover { border-color: #ef4444; }
.folder-tile.show-foldered-toggle { border-color: var(--text-muted); border-style: dashed; }
.folder-tile.show-foldered-toggle:hover { border-color: #b91c1c; }
.folder-tile .folder-icon { font-size: 16px; flex-shrink: 0; }
.folder-tile .folder-name { font-weight: 600; font-size: 13px; flex: 1; overflow: hidden;
  text-overflow: ellipsis; white-space: nowrap; }
.folder-tile .folder-count { font-size: 11px; color: var(--text-muted); background: var(--input-bg);
  padding: 1px 7px; border-radius: 10px; }
.folder-manage-item { display: flex; align-items: center; gap: 4px; padding: 2px 0; font-size: 12px; }
.folder-manage-item .folder-manage-name { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.folder-manage-item button { background: none; border: none; color: var(--text-muted); cursor: pointer;
  font-size: 12px; padding: 2px 4px; border-radius: 3px; }
.folder-manage-item button:hover { color: #ef4444; background: var(--drop-highlight); }
"
    End Function

    ''' <summary>Returns the header HTML for the board.</summary>
    Private Function GetBoardHeaderHtml() As String
        Dim logoDataUrl As String = GetLogoDataUrl()
        Dim brandHtml As String = System.Net.WebUtility.HtmlEncode(If(Not String.IsNullOrWhiteSpace(AN), AN, "Red Ink"))

        Dim logoImg As String = ""
        If Not String.IsNullOrWhiteSpace(logoDataUrl) Then
            logoImg = $"<img class=""logo"" src=""{System.Net.WebUtility.HtmlEncode(logoDataUrl)}"" alt=""logo"">"
        End If

        ' Build pinned columns checkboxes from all Outlook categories
        Dim pinnedHtml As New StringBuilder()
        Dim pinnedSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Try
            Dim alwaysPresent As String = If(My.Settings.InboxBoardColumns, "")
            If Not String.IsNullOrEmpty(alwaysPresent) Then
                For Each col In alwaysPresent.Split({";"c}, StringSplitOptions.RemoveEmptyEntries)
                    Dim trimmed As String = col.Trim()
                    If Not String.IsNullOrEmpty(trimmed) Then pinnedSet.Add(trimmed)
                Next
            End If
        Catch
        End Try

        Try
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
            Dim categoryColors As Dictionary(Of String, String) = GetOutlookCategoryColors(ns)
            Dim categories As Outlook.Categories = ns.Categories
            For Each cat As Outlook.Category In categories
                Try
                    Dim catName As String = cat.Name
                    Dim escapedName As String = System.Net.WebUtility.HtmlEncode(catName)
                    Dim hexColor As String = "#6b7280"
                    If categoryColors.ContainsKey(catName) Then hexColor = categoryColors(catName)
                    Dim isChecked As String = If(pinnedSet.Contains(catName), " checked", "")
                    pinnedHtml.AppendLine($"        <label><input type=""checkbox"" data-pinned-cat=""{escapedName}""{isChecked}><span class=""cat-dot"" style=""background:{hexColor}""></span> {escapedName}</label>")
                Catch
                End Try
            Next
        Catch
        End Try

        ' Conversation grouping checkbox
        Dim groupChecked As String = ""
        Try : If My.Settings.InboxBoardGroupConversations Then groupChecked = " checked"
        Catch : End Try

        ' Include flagged checkbox
        Dim flaggedChecked As String = ""
        Try : If My.Settings.InboxBoardIncludeFlagged Then flaggedChecked = " checked"
        Catch : End Try

        ' Hide done flags checkbox
        Dim hideDoneChecked As String = ""
        Try : If My.Settings.InboxBoardHideDoneFlags Then hideDoneChecked = " checked"
        Catch : End Try

        ' Group flagged by date checkbox
        Dim groupFlagsByDateChecked As String = ""
        Try : If My.Settings.InboxBoardGroupFlagsByDate Then groupFlagsByDateChecked = " checked"
        Catch : End Try

        ' Summary language: free text input with "Auto-detect" overlay
        Dim savedLang As String = ""
        Try : savedLang = If(My.Settings.InboxBoardSummaryLanguage, "") : Catch : End Try
        Dim escapedLang As String = System.Net.WebUtility.HtmlEncode(savedLang)
        Dim dataEmpty As String = If(String.IsNullOrEmpty(savedLang), " data-empty=""true""", "")

        ' Load threshold setting
        Dim savedThreshold As Integer = 0
        Try : savedThreshold = My.Settings.InboxBoardLoadThreshold : Catch : End Try
        If savedThreshold <= 0 Then savedThreshold = InboxBoard_DefaultLoadThreshold

        ' Include subfolders checkbox
        Dim subfoldersChecked As String = ""
        Try : If My.Settings.InboxBoardIncludeSubfolders Then subfoldersChecked = " checked"
        Catch : subfoldersChecked = " checked" : End Try

        Return $"
<div class=""header"">
  <div class=""topline"">
    {logoImg}
    <div class=""brandbig"">{brandHtml}</div>
    <div class=""sub"">Inbox Board</div>
  </div>
  <div class=""header-controls"">
    <input type=""text"" class=""search-input"" id=""searchInput"" placeholder=""Search subjects, senders…"">
    <select class=""filter-select"" id=""columnFilter""><option value=""all"">All Columns</option></select>
    <div class=""settings-wrapper"">
      <button class=""icon-btn"" id=""settingsBtn"" title=""Card display settings"">
        <svg class=""icon-svg"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round""><circle cx=""12"" cy=""12"" r=""3""/><path d=""M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.32 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z""/></svg>
      </button>
      <div class=""settings-panel"" id=""settingsPanel"">
        <h3>Card Fields</h3>
        <label><input type=""checkbox"" data-field=""sender"" checked> Sender</label>
        <label><input type=""checkbox"" data-field=""date"" checked> Date</label>
        <label><input type=""checkbox"" data-field=""messageCount"" checked> Message count</label>
        <label><input type=""checkbox"" data-field=""summary"" checked> AI Summary</label>
        <h3 class=""section-divider"">Grouping</h3>
        <label><input type=""checkbox"" id=""groupConversationsChk""{groupChecked}> Group conversations</label>
        <h3 class=""section-divider"">Flags</h3>
        <label><input type=""checkbox"" id=""includeFlaggedChk""{flaggedChecked}> Include flagged mails</label>
        <label><input type=""checkbox"" id=""hideDoneFlagsChk""{hideDoneChecked}> Hide completed flags</label>
        <label><input type=""checkbox"" id=""groupFlagsByDateChk""{groupFlagsByDateChecked}> Group flagged by due date</label>
        <h3 class=""section-divider"">Summary Language</h3>
        <div class=""lang-input-wrapper"">
          <input type=""text"" id=""summaryLanguageInput"" value=""{escapedLang}""{dataEmpty}>
          <div class=""lang-overlay"">Auto-detect</div>
        </div>
        <h3 class=""section-divider"">Load Limit</h3>
        <div style=""font-size:11px;color:var(--text-muted);line-height:1.4;padding:0 0 4px 0"">
          Prompt when qualifying mails exceed this number. Takes effect on next Reload.
        </div>
        <input type=""number"" id=""loadThresholdInput"" value=""{savedThreshold}"" min=""50"" max=""50000"" step=""50""
          style=""padding:4px 6px;border:1px solid var(--input-border);border-radius:4px;background:var(--input-bg);color:var(--text);font-size:12px;outline:none;width:100%"">
        <h3 class=""section-divider"">Tracked Folder</h3>
        <div style=""font-size:11px;color:var(--text-muted);line-height:1.4;padding:0 0 4px 0"">
          Currently tracking:
        </div>
        <div id=""trackedFolderLabel"" style=""font-size:12px;font-weight:600;color:var(--text);padding:2px 0 6px 0;word-break:break-all""></div>
        <div style=""display:flex;gap:4px;margin-bottom:6px"">
          <button id=""changeFolderBtn"" style=""flex:1;padding:5px 8px;border:1px solid var(--input-border);border-radius:4px;background:var(--input-bg);color:var(--text);font-size:12px;cursor:pointer"">Change…</button>
          <button id=""resetFolderBtn"" style=""padding:5px 8px;border:1px solid var(--input-border);border-radius:4px;background:var(--input-bg);color:var(--text);font-size:12px;cursor:pointer"" title=""Reset to default Inbox"">↺</button>
        </div>
        <label><input type=""checkbox"" id=""includeSubfoldersChk""{subfoldersChecked}> Include subfolders</label>
        <h3 class=""section-divider"">Pinned Columns</h3>
{pinnedHtml.ToString().TrimEnd()}
        <h3 class=""section-divider"">Pinned Flag Columns</h3>
        <div id=""pinnedFlagColumnsSection"" style=""display:{If(groupFlagsByDateChecked <> "", "block", "none")}"">
{GetPinnedFlagColumnsHtml()}
        </div>
        <div id=""pinnedFlagColumnsHint"" style=""display:{If(groupFlagsByDateChecked = "", "block", "none")};font-size:11px;color:var(--text-muted);padding:2px 0"">
          Enable &quot;Group flagged by due date&quot; first
        </div>
        <h3 class=""section-divider"">Board Folders</h3>
        <div id=""folderManageSection"">
          <div style=""display:flex;gap:4px;margin-bottom:6px"">
            <input type=""text"" id=""newFolderInput"" placeholder=""New folder name…"" style=""flex:1;padding:4px 6px;border:1px solid var(--input-border);border-radius:4px;background:var(--input-bg);color:var(--text);font-size:12px;outline:none"">
            <button id=""addFolderBtn"" style=""padding:4px 8px;border:1px solid var(--input-border);border-radius:4px;background:var(--input-bg);color:var(--text);font-size:12px;cursor:pointer"">+</button>
          </div>
          <div id=""folderList""></div>
        </div>
        <h3 class=""section-divider"">Drag &amp; Drop</h3>
        <div style=""font-size:11px;color:var(--text-muted);line-height:1.4;padding:2px 0"">
          <b>Drag</b> → replace all categories<br>
          <b>Ctrl + Drag</b> → add category<br>
          <b>Drag onto 📁</b> → move to folder
        </div>
      </div>
    </div>
    <button class=""icon-btn"" id=""reloadBtn"" title=""Reload mails from Inbox"">
      <svg class=""icon-svg"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round""><polyline points=""23 4 23 10 17 10""/><polyline points=""1 20 1 14 7 14""/><path d=""M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15""/></svg>
    </button>
    <button class=""icon-btn"" id=""themeToggle"" title=""Toggle theme"">
      <svg class=""icon-svg"" id=""themeIcon"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round""><circle cx=""12"" cy=""12"" r=""5""/><line x1=""12"" y1=""1"" x2=""12"" y2=""3""/><line x1=""12"" y1=""21"" x2=""12"" y2=""23""/><line x1=""4.22"" y1=""4.22"" x2=""5.64"" y2=""5.64""/><line x1=""18.36"" y1=""18.36"" x2=""19.78"" y2=""19.78""/><line x1=""1"" y1=""12"" x2=""3"" y2=""12""/><line x1=""21"" y1=""12"" x2=""23"" y2=""12""/><line x1=""4.22"" y1=""19.78"" x2=""5.64"" y2=""18.36""/><line x1=""18.36"" y1=""5.64"" x2=""19.78"" y2=""4.22""/></svg>
    </button>
  </div>
</div>"
    End Function

    ''' <summary>Builds checkboxes for pinning individual flag date-bucket columns.</summary>
    Private Function GetPinnedFlagColumnsHtml() As String
        Dim sb As New StringBuilder()
        Dim pinnedSet As New HashSet(Of String)(StringComparer.Ordinal)
        Try
            Dim saved As String = If(My.Settings.InboxBoardPinnedFlagColumns, "")
            If Not String.IsNullOrEmpty(saved) Then
                For Each part In saved.Split({";"c}, StringSplitOptions.RemoveEmptyEntries)
                    Dim t = part.Trim()
                    If Not String.IsNullOrEmpty(t) Then pinnedSet.Add(t)
                Next
            End If
        Catch
        End Try

        For Each bucketId In InboxBoard_FlagBucketOrder
            Dim escapedId As String = System.Net.WebUtility.HtmlEncode(bucketId)
            Dim bucketColor As String = "#6b7280"
            If InboxBoard_FlagBucketColors.ContainsKey(bucketId) Then bucketColor = InboxBoard_FlagBucketColors(bucketId)
            Dim isChecked As String = If(pinnedSet.Contains(bucketId), " checked", "")
            sb.AppendLine($"          <label><input type=""checkbox"" data-pinned-flag=""{escapedId}""{isChecked}><span class=""cat-dot"" style=""background:{bucketColor}""></span> {escapedId}</label>")
        Next

        Return sb.ToString().TrimEnd()
    End Function

    ''' <summary>Returns the JavaScript for the board.</summary>
    Private Function GetBoardJavaScript() As String
        Dim js As New StringBuilder(30000)
        js.AppendLine("// --- State ---")
        js.AppendLine("var cards = [];")
        js.AppendLine("var COLUMNS = [];")
        js.AppendLine("var fieldSettings = { sender: true, date: true, messageCount: true, summary: true };")
        js.AppendLine("var isDragging = false;")
        js.AppendLine("var ctrlHeld = false;")
        js.AppendLine("var langCommitTimer = null;")
        js.AppendLine("var columnSortState = {};")
        js.AppendLine("var boardFolders = (typeof SAVED_FOLDERS !== 'undefined') ? SAVED_FOLDERS : [];")
        js.AppendLine("var selectedFolder = null;")
        js.AppendLine("var hideFoldered = true;")
        js.AppendLine()
        js.AppendLine("function init(data) {")
        js.AppendLine("  try {")
        js.AppendLine("    var d = (typeof data === 'string') ? JSON.parse(data) : data;")
        js.AppendLine("    COLUMNS = d.columns || [];")
        js.AppendLine("    cards = (d.cards || []).map(function(c, i) { return Object.assign({}, c, { id: i + 1 }); });")
        js.AppendLine("    if (d.folders) { boardFolders = d.folders; }")
        js.AppendLine("    if (d.trackedFolder) { var lbl = document.getElementById('trackedFolderLabel'); if (lbl) lbl.textContent = d.trackedFolder; }")
        js.AppendLine("    console.log('InboxBoard init: ' + cards.length + ' cards, ' + COLUMNS.length + ' columns');")
        js.AppendLine("    populateColumnFilter();")
        js.AppendLine("    renderBoard();")
        js.AppendLine("    loadFieldSettings();")
        js.AppendLine("    applyFieldVisibility();")
        js.AppendLine("    restoreColumnFilter();")
        js.AppendLine("  } catch(err) { document.body.innerHTML = '<pre style=""color:red;padding:20px"">' + err.message + '\n' + err.stack + '</pre>'; }")
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("function populateColumnFilter() {")
        js.AppendLine("  var sel = document.getElementById('columnFilter');")
        js.AppendLine("  sel.innerHTML = '<option value=""all"">All Columns</option>';")
        js.AppendLine("  COLUMNS.forEach(function(col) {")
        js.AppendLine("    var opt = document.createElement('option');")
        js.AppendLine("    opt.value = col.id;")
        js.AppendLine("    opt.textContent = col.title;")
        js.AppendLine("    sel.appendChild(opt);")
        js.AppendLine("  });")
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("function restoreColumnFilter() {")
        js.AppendLine("  if (SAVED_COLUMN_FILTER) {")
        js.AppendLine("    var sel = document.getElementById('columnFilter');")
        js.AppendLine("    for (var i = 0; i < sel.options.length; i++) {")
        js.AppendLine("      if (sel.options[i].value === SAVED_COLUMN_FILTER) { sel.selectedIndex = i; break; }")
        js.AppendLine("    }")
        js.AppendLine("    applySearchFilter();")
        js.AppendLine("  }")
        js.AppendLine("}")
        js.AppendLine()
        ' --- Board panning ---
        js.AppendLine("(function() {")
        js.AppendLine("  var board = null, isPanning = false, startX = 0, startY = 0, scrollLeftStart = 0, scrollTopStart = 0;")
        js.AppendLine("  function getBoard() { if (!board) board = document.getElementById('board'); return board; }")
        js.AppendLine("  document.addEventListener('mousedown', function(e) {")
        js.AppendLine("    var b = getBoard(); if (!b) return;")
        js.AppendLine("    if (e.button !== 1 && e.button !== 2) return;")
        js.AppendLine("    var target = e.target;")
        js.AppendLine("    if (target.closest('.card') || target.closest('.card-btn') || target.closest('.sort-dropdown') || target.closest('.header')) return;")
        js.AppendLine("    if (!target.closest('.board') && !target.closest('.column') && target !== b) return;")
        js.AppendLine("    e.preventDefault(); isPanning = true; startX = e.clientX; startY = e.clientY;")
        js.AppendLine("    scrollLeftStart = b.scrollLeft; scrollTopStart = window.scrollY; b.classList.add('panning');")
        js.AppendLine("  });")
        js.AppendLine("  document.addEventListener('mousemove', function(e) {")
        js.AppendLine("    if (!isPanning) return; var b = getBoard(); if (!b) return; e.preventDefault();")
        js.AppendLine("    b.scrollLeft = scrollLeftStart - (e.clientX - startX);")
        js.AppendLine("    window.scrollTo(window.scrollX, scrollTopStart - (e.clientY - startY));")
        js.AppendLine("  });")
        js.AppendLine("  document.addEventListener('mouseup', function() { if (!isPanning) return; isPanning = false; var b = getBoard(); if (b) b.classList.remove('panning'); });")
        js.AppendLine("  document.addEventListener('contextmenu', function(e) { if (e.target.closest('.board') || e.target.closest('.column')) e.preventDefault(); });")
        js.AppendLine("})();")
        js.AppendLine()
        ' --- Persistence ---
        js.AppendLine("function saveState() {")
        js.AppendLine("  try {")
        js.AppendLine("    var state = COLUMNS.map(function(col) {")
        js.AppendLine("      var colEl = document.querySelector('.column[data-column=""' + CSS.escape(col.id) + '""]');")
        js.AppendLine("      if (!colEl) return { id: col.id, cardIds: [] };")
        js.AppendLine("      var cids = [];")
        js.AppendLine("      colEl.querySelector('.card-list').querySelectorAll('.card').forEach(function(c) { cids.push(c.dataset.entryId); });")
        js.AppendLine("      return { id: col.id, cardIds: cids };")
        js.AppendLine("    });")
        js.AppendLine("    localStorage.setItem('rib-columns', JSON.stringify(state));")
        js.AppendLine("  } catch(err) { console.error('saveState error:', err); }")
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("function saveFieldSettings() { window.chrome.webview.postMessage(JSON.stringify({ action: 'setFieldSettings', fields: JSON.stringify(fieldSettings) })); }")
        js.AppendLine("function loadFieldSettings() {")
        js.AppendLine("  if (SAVED_FIELDS) { try { fieldSettings = JSON.parse(SAVED_FIELDS); } catch(e) {} }")
        js.AppendLine("  document.querySelectorAll('#settingsPanel input[type=""checkbox""][data-field]').forEach(function(cb) { cb.checked = fieldSettings[cb.dataset.field] !== false; });")
        js.AppendLine("}")
        js.AppendLine()
        ' --- Theme ---
        js.AppendLine("function applyTheme(theme) {")
        js.AppendLine("  document.documentElement.setAttribute('data-theme', theme);")
        js.AppendLine("  var icon = document.getElementById('themeIcon');")
        js.AppendLine("  if (theme === 'dark') { icon.innerHTML = '<path d=""M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z""/>'; }")
        js.AppendLine("  else { icon.innerHTML = '<circle cx=""12"" cy=""12"" r=""5""/><line x1=""12"" y1=""1"" x2=""12"" y2=""3""/><line x1=""12"" y1=""21"" x2=""12"" y2=""23""/><line x1=""4.22"" y1=""4.22"" x2=""5.64"" y2=""5.64""/><line x1=""18.36"" y1=""18.36"" x2=""19.78"" y2=""19.78""/><line x1=""1"" y1=""12"" x2=""3"" y2=""12""/><line x1=""21"" y1=""12"" x2=""23"" y2=""12""/><line x1=""4.22"" y1=""19.78"" x2=""5.64"" y2=""18.36""/><line x1=""18.36"" y1=""5.64"" x2=""19.78"" y2=""4.22""/>'; }")
        js.AppendLine("}")
        js.AppendLine("function initTheme() { var theme = SAVED_THEME || (window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'); applyTheme(theme); }")
        js.AppendLine("document.getElementById('themeToggle').addEventListener('click', function() { var cur = document.documentElement.getAttribute('data-theme'); var next = cur === 'dark' ? 'light' : 'dark'; applyTheme(next); window.chrome.webview.postMessage(JSON.stringify({ action: 'setTheme', theme: next })); });")
        js.AppendLine()
        ' --- Settings panel ---
        js.AppendLine("document.getElementById('settingsBtn').addEventListener('click', function(e) { e.stopPropagation(); document.getElementById('settingsPanel').classList.toggle('open'); });")
        js.AppendLine("document.addEventListener('click', function(e) {")
        js.AppendLine("  var p = document.getElementById('settingsPanel');")
        js.AppendLine("  if (!p.contains(e.target) && e.target !== document.getElementById('settingsBtn')) p.classList.remove('open');")
        js.AppendLine("  if (!e.target.closest('.sort-btn') && !e.target.closest('.sort-dropdown')) document.querySelectorAll('.sort-dropdown.open').forEach(function(dd) { dd.classList.remove('open'); });")
        js.AppendLine("});")
        js.AppendLine("document.querySelectorAll('#settingsPanel input[type=""checkbox""][data-field]').forEach(function(cb) { cb.addEventListener('change', function() { fieldSettings[cb.dataset.field] = cb.checked; saveFieldSettings(); applyFieldVisibility(); }); });")
        js.AppendLine()
        ' --- Toggles ---
        js.AppendLine("document.getElementById('groupConversationsChk').addEventListener('change', function() { window.chrome.webview.postMessage(JSON.stringify({ action: 'setGroupConversations', enabled: this.checked })); });")
        js.AppendLine("document.getElementById('includeFlaggedChk').addEventListener('change', function() { window.chrome.webview.postMessage(JSON.stringify({ action: 'setIncludeFlagged', enabled: this.checked })); });")
        js.AppendLine("document.getElementById('hideDoneFlagsChk').addEventListener('change', function() { window.chrome.webview.postMessage(JSON.stringify({ action: 'setHideDoneFlags', enabled: this.checked })); });")
        js.AppendLine("document.getElementById('groupFlagsByDateChk').addEventListener('change', function() {")
        js.AppendLine("  var sec = document.getElementById('pinnedFlagColumnsSection'); var hint = document.getElementById('pinnedFlagColumnsHint');")
        js.AppendLine("  if (sec) sec.style.display = this.checked ? 'block' : 'none';")
        js.AppendLine("  if (hint) hint.style.display = this.checked ? 'none' : 'block';")
        js.AppendLine("  window.chrome.webview.postMessage(JSON.stringify({ action: 'setGroupFlagsByDate', enabled: this.checked }));")
        js.AppendLine("});")
        js.AppendLine()
        ' --- Summary language ---
        js.AppendLine("(function() {")
        js.AppendLine("  var langInput = document.getElementById('summaryLanguageInput');")
        js.AppendLine("  var lastCommitted = SAVED_SUMMARY_LANGUAGE || '';")
        js.AppendLine("  if (SAVED_SUMMARY_LANGUAGE) langInput.value = SAVED_SUMMARY_LANGUAGE;")
        js.AppendLine("  function updateEmpty() { langInput.setAttribute('data-empty', langInput.value.trim() === '' ? 'true' : 'false'); }")
        js.AppendLine("  updateEmpty();")
        js.AppendLine("  function commitLang() { if (langCommitTimer) { clearTimeout(langCommitTimer); langCommitTimer = null; } var val = langInput.value.trim(); if (val !== lastCommitted) { lastCommitted = val; window.chrome.webview.postMessage(JSON.stringify({ action: 'setSummaryLanguage', language: val })); showToast(val ? 'Summary language: ' + val : 'Summary language: Auto-detect'); } }")
        js.AppendLine("  langInput.addEventListener('input', function() { updateEmpty(); if (langCommitTimer) clearTimeout(langCommitTimer); langCommitTimer = setTimeout(commitLang, 1500); });")
        js.AppendLine("  langInput.addEventListener('blur', function() { if (langCommitTimer) { clearTimeout(langCommitTimer); langCommitTimer = null; } commitLang(); });")
        js.AppendLine("  langInput.addEventListener('keydown', function(e) { if (e.key === 'Enter') { e.preventDefault(); if (langCommitTimer) { clearTimeout(langCommitTimer); langCommitTimer = null; } commitLang(); langInput.blur(); } });")
        js.AppendLine("})();")
        js.AppendLine()
        ' --- Pinned columns ---
        js.AppendLine("document.querySelectorAll('#settingsPanel input[type=""checkbox""][data-pinned-cat]').forEach(function(cb) { cb.addEventListener('change', function() { var pinned = []; document.querySelectorAll('#settingsPanel input[type=""checkbox""][data-pinned-cat]').forEach(function(c) { if (c.checked) pinned.push(c.dataset.pinnedCat); }); window.chrome.webview.postMessage(JSON.stringify({ action: 'setPinnedColumns', columns: pinned })); showToast('Pinned columns updated'); }); });")
        js.AppendLine("document.querySelectorAll('#settingsPanel input[type=""checkbox""][data-pinned-flag]').forEach(function(cb) { cb.addEventListener('change', function() { var pinned = []; document.querySelectorAll('#settingsPanel input[type=""checkbox""][data-pinned-flag]').forEach(function(c) { if (c.checked) pinned.push(c.dataset.pinnedFlag); }); window.chrome.webview.postMessage(JSON.stringify({ action: 'setPinnedFlagColumns', columns: pinned })); showToast('Pinned flag columns updated'); }); });")
        js.AppendLine()
        ' --- Board folder management in settings ---
        js.AppendLine("function renderFolderManageList() {")
        js.AppendLine("  var listEl = document.getElementById('folderList'); if (!listEl) return;")
        js.AppendLine("  listEl.innerHTML = '';")
        js.AppendLine("  var allNames = getAllFolderNames();")
        js.AppendLine("  allNames.forEach(function(fn){")
        js.AppendLine("    var item = document.createElement('div'); item.className = 'folder-manage-item';")
        js.AppendLine("    var nameSpan = document.createElement('span'); nameSpan.className = 'folder-manage-name'; nameSpan.textContent = '\uD83D\uDCC1 ' + fn;")
        js.AppendLine("    var renameBtn = document.createElement('button'); renameBtn.title = 'Rename'; renameBtn.textContent = '\u270E';")
        js.AppendLine("    renameBtn.addEventListener('click', function(){ showPrompt('Rename folder \u0022' + fn + '\u0022 to:', fn, function(nn){ if(nn && nn.trim() && nn.trim()!==fn) window.chrome.webview.postMessage(JSON.stringify({action:'renameFolder',oldName:fn,newName:nn.trim()})); }); });")
        js.AppendLine("    var deleteBtn = document.createElement('button'); deleteBtn.title = 'Delete'; deleteBtn.textContent = '\u2716';")
        js.AppendLine("    deleteBtn.addEventListener('click', function(){ showConfirm('Delete folder \u0022' + fn + '\u0022? Cards will return to the main board.', function(){ window.chrome.webview.postMessage(JSON.stringify({action:'deleteFolder',name:fn})); }); });")
        js.AppendLine("    item.appendChild(nameSpan); item.appendChild(renameBtn); item.appendChild(deleteBtn);")
        js.AppendLine("    listEl.appendChild(item);")
        js.AppendLine("  });")
        js.AppendLine("  if(allNames.length===0){ var hint=document.createElement('div'); hint.style.cssText='font-size:11px;color:var(--text-muted);padding:2px 0'; hint.textContent='No folders yet. Create one above.'; listEl.appendChild(hint); }")
        js.AppendLine("}")
        js.AppendLine("document.getElementById('addFolderBtn').addEventListener('click', function(){ var input = document.getElementById('newFolderInput'); var name = input.value.trim(); if(!name) return; window.chrome.webview.postMessage(JSON.stringify({action:'createFolder',name:name})); input.value = ''; showToast('\uD83D\uDCC1 Folder ""' + name + '"" created'); });")
        js.AppendLine("document.getElementById('newFolderInput').addEventListener('keydown', function(e){ if(e.key==='Enter'){e.preventDefault(); document.getElementById('addFolderBtn').click();} });")
        js.AppendLine()
        ' --- Load threshold setting ---
        js.AppendLine("(function() {")
        js.AppendLine("  var threshInput = document.getElementById('loadThresholdInput'); if (!threshInput) return;")
        js.AppendLine("  var threshTimer = null;")
        js.AppendLine("  function commitThreshold() { if (threshTimer) { clearTimeout(threshTimer); threshTimer = null; } var val = parseInt(threshInput.value, 10); if (isNaN(val) || val < 50) val = 50; if (val > 50000) val = 50000; threshInput.value = val; window.chrome.webview.postMessage(JSON.stringify({ action: 'setLoadThreshold', threshold: val })); showToast('Load threshold: ' + val + ' mails (active on next Reload)'); }")
        js.AppendLine("  threshInput.addEventListener('input', function() { if (threshTimer) clearTimeout(threshTimer); threshTimer = setTimeout(commitThreshold, 1500); });")
        js.AppendLine("  threshInput.addEventListener('blur', function() { if (threshTimer) { clearTimeout(threshTimer); threshTimer = null; } commitThreshold(); });")
        js.AppendLine("  threshInput.addEventListener('keydown', function(e) { if (e.key === 'Enter') { e.preventDefault(); if (threshTimer) { clearTimeout(threshTimer); threshTimer = null; } commitThreshold(); threshInput.blur(); } });")
        js.AppendLine("})();")
        js.AppendLine()
        ' --- Tracked folder controls ---
        js.AppendLine("document.getElementById('changeFolderBtn').addEventListener('click', function() {")
        js.AppendLine("  window.chrome.webview.postMessage(JSON.stringify({ action: 'changeTargetFolder' }));")
        js.AppendLine("});")
        js.AppendLine("document.getElementById('resetFolderBtn').addEventListener('click', function() {")
        js.AppendLine("  window.chrome.webview.postMessage(JSON.stringify({ action: 'resetTargetFolder' }));")
        js.AppendLine("  showToast('Reset to default Inbox');")
        js.AppendLine("});")
        js.AppendLine("document.getElementById('includeSubfoldersChk').addEventListener('change', function() {")
        js.AppendLine("  window.chrome.webview.postMessage(JSON.stringify({ action: 'setIncludeSubfolders', enabled: this.checked }));")
        js.AppendLine("});")
        js.AppendLine("(function() {")
        js.AppendLine("  var lbl = document.getElementById('trackedFolderLabel');")
        js.AppendLine("  if (lbl) lbl.textContent = (typeof SAVED_TRACKED_FOLDER !== 'undefined') ? SAVED_TRACKED_FOLDER : 'Inbox';")
        js.AppendLine("})();")
        js.AppendLine()
        ' --- Folder helpers (new: dedicated folder column + filter) ---
        js.AppendLine("function cardMatchesColumn(c, col) {")
        js.AppendLine("  if (col.isFlaggedColumn) {")
        js.AppendLine("    if (!(c.isFlagged || c.isCompleted)) return false;")
        js.AppendLine("    if (col.id === FLAGGED_COLUMN_ID) return true;")
        js.AppendLine("    if (c.flagBucket) return c.flagBucket === col.id;")
        js.AppendLine("    return false;")
        js.AppendLine("  }")
        js.AppendLine("  if (c.allCategories && c.allCategories.length > 0) return c.allCategories.some(function(ac){ return ac.name === col.id; });")
        js.AppendLine("  return c.category === col.id;")
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("function getAllFolderNames() {")
        js.AppendLine("  var allNames = boardFolders.map(function(f){ return f.name; });")
        js.AppendLine("  cards.forEach(function(c){ if(c.boardFolder && allNames.indexOf(c.boardFolder)<0) allNames.push(c.boardFolder); });")
        js.AppendLine("  return allNames;")
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("function applyFolderFilter() {")
        js.AppendLine("  document.querySelectorAll('.card').forEach(function(cardEl) {")
        js.AppendLine("    var card = cards.find(function(c) { return c.entryId === cardEl.dataset.entryId; });")
        js.AppendLine("    if (!card) return;")
        js.AppendLine("    if (selectedFolder) {")
        js.AppendLine("      if (card.boardFolder === selectedFolder) cardEl.classList.remove('folder-hidden');")
        js.AppendLine("      else cardEl.classList.add('folder-hidden');")
        js.AppendLine("    } else if (hideFoldered && card.boardFolder) {")
        js.AppendLine("      cardEl.classList.add('folder-hidden');")
        js.AppendLine("    } else {")
        js.AppendLine("      cardEl.classList.remove('folder-hidden');")
        js.AppendLine("    }")
        js.AppendLine("  });")
        js.AppendLine("  document.querySelectorAll('.folder-tile:not(.show-all):not(.show-foldered-toggle)').forEach(function(tile) {")
        js.AppendLine("    if (tile.dataset.folderName === selectedFolder) tile.classList.add('selected');")
        js.AppendLine("    else tile.classList.remove('selected');")
        js.AppendLine("  });")
        js.AppendLine("  var showAllBtn = document.getElementById('folderShowAll');")
        js.AppendLine("  if (showAllBtn) showAllBtn.style.display = selectedFolder ? '' : 'none';")
        js.AppendLine("  var toggleBtn = document.getElementById('folderToggleHide');")
        js.AppendLine("  if (toggleBtn) {")
        js.AppendLine("    if (selectedFolder) { toggleBtn.style.display = 'none'; }")
        js.AppendLine("    else {")
        js.AppendLine("      toggleBtn.style.display = '';")
        js.AppendLine("      var toggleName = toggleBtn.querySelector('.folder-name');")
        js.AppendLine("      if (toggleName) toggleName.textContent = hideFoldered ? 'Show foldered cards' : 'Hide foldered cards';")
        js.AppendLine("      var toggleIcon = toggleBtn.querySelector('.folder-icon');")
        js.AppendLine("      if (toggleIcon) toggleIcon.textContent = hideFoldered ? '\uD83D\uDC41' : '\uD83D\uDEAB';")
        js.AppendLine("    }")
        js.AppendLine("  }")
        js.AppendLine("  updateCounts();")
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("function createFolderTile(folderName) {")
        js.AppendLine("  var tile = document.createElement('div');")
        js.AppendLine("  tile.className = 'folder-tile' + (selectedFolder === folderName ? ' selected' : '');")
        js.AppendLine("  tile.dataset.folderName = folderName;")
        js.AppendLine("  var icon = document.createElement('span'); icon.className = 'folder-icon'; icon.textContent = '\uD83D\uDCC1';")
        js.AppendLine("  var name = document.createElement('span'); name.className = 'folder-name'; name.textContent = folderName;")
        js.AppendLine("  var count = document.createElement('span'); count.className = 'folder-count';")
        js.AppendLine("  count.textContent = cards.filter(function(c){ return c.boardFolder === folderName; }).length;")
        js.AppendLine("  tile.appendChild(icon); tile.appendChild(name); tile.appendChild(count);")
        js.AppendLine("  tile.addEventListener('click', function(){")
        js.AppendLine("    if (selectedFolder === folderName) { selectedFolder = null; showToast('Showing all cards'); }")
        js.AppendLine("    else { selectedFolder = folderName; showToast('\uD83D\uDCC1 Filtering: ' + folderName); }")
        js.AppendLine("    applyFolderFilter();")
        js.AppendLine("  });")
        js.AppendLine("  tile.addEventListener('dragover', function(e){ e.preventDefault(); e.stopPropagation(); e.dataTransfer.dropEffect='move'; tile.classList.add('drag-over'); });")
        js.AppendLine("  tile.addEventListener('dragleave', function(e){ if(!tile.contains(e.relatedTarget)) tile.classList.remove('drag-over'); });")
        js.AppendLine("  tile.addEventListener('drop', function(e){")
        js.AppendLine("    e.preventDefault(); e.stopPropagation(); tile.classList.remove('drag-over');")
        js.AppendLine("    var entryId = e.dataTransfer.getData('text/plain'); if(!entryId) return;")
        js.AppendLine("    var card = cards.find(function(c){ return c.entryId===entryId; });")
        js.AppendLine("    if(card && card.boardFolder===folderName){ showToast('Already in '+folderName); return; }")
        js.AppendLine("    window.chrome.webview.postMessage(JSON.stringify({action:'moveToFolder',entryId:entryId,folderName:folderName}));")
        js.AppendLine("    showToast('\uD83D\uDCC1 Moved to '+folderName);")
        js.AppendLine("  });")
        js.AppendLine("  return tile;")
        js.AppendLine("}")
        js.AppendLine()
        ' --- Reload button ---
        js.AppendLine("document.getElementById('reloadBtn').addEventListener('click', function() { window.chrome.webview.postMessage(JSON.stringify({ action: 'reload' })); });")
        js.AppendLine()
        ' --- Ctrl key tracking ---
        js.AppendLine("document.addEventListener('keydown', function(e) { if (e.key === 'Control' && !ctrlHeld) { ctrlHeld = true; updateDragModeVisuals(); } });")
        js.AppendLine("document.addEventListener('keyup', function(e) { if (e.key === 'Control') { ctrlHeld = false; updateDragModeVisuals(); } });")
        js.AppendLine("window.addEventListener('blur', function() { ctrlHeld = false; updateDragModeVisuals(); });")
        js.AppendLine()
        js.AppendLine("function updateDragModeVisuals() {")
        js.AppendLine("  var label = document.getElementById('dragModeLabel');")
        js.AppendLine("  if (isDragging && ctrlHeld) {")
        js.AppendLine("    label.classList.add('visible');")
        js.AppendLine("    document.querySelectorAll('.column.drag-over').forEach(function(el) { el.classList.add('drag-over-add'); });")
        js.AppendLine("    document.querySelectorAll('.drop-indicator.visible').forEach(function(el) { el.classList.add('add-mode'); });")
        js.AppendLine("  } else {")
        js.AppendLine("    label.classList.remove('visible');")
        js.AppendLine("    document.querySelectorAll('.column.drag-over-add').forEach(function(el) { el.classList.remove('drag-over-add'); });")
        js.AppendLine("    document.querySelectorAll('.drop-indicator.add-mode').forEach(function(el) { el.classList.remove('add-mode'); });")
        js.AppendLine("  }")
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("function applyFieldVisibility() {")
        js.AppendLine("  document.querySelectorAll('.card-sender').forEach(function(el) { el.style.display = fieldSettings.sender ? '' : 'none'; });")
        js.AppendLine("  document.querySelectorAll('.card-meta .card-date').forEach(function(el) { el.style.display = fieldSettings.date ? '' : 'none'; });")
        js.AppendLine("  document.querySelectorAll('.card-meta .card-msg-count').forEach(function(el) { el.style.display = fieldSettings.messageCount ? '' : 'none'; });")
        js.AppendLine("  document.querySelectorAll('.card-summary').forEach(function(el) { el.style.display = fieldSettings.summary ? '' : 'none'; });")
        js.AppendLine("}")
        js.AppendLine()
        ' --- Column sorting ---
        js.AppendLine("var SORT_OPTIONS = [")
        js.AppendLine("  { key: 'date-desc', label: 'Date (newest first)' }, { key: 'date-asc', label: 'Date (oldest first)' },")
        js.AppendLine("  { key: 'subject-asc', label: 'Subject (A\u2013Z)' }, { key: 'subject-desc', label: 'Subject (Z\u2013A)' },")
        js.AppendLine("  { key: 'sender-asc', label: 'Sender (A\u2013Z)' }, { key: 'sender-desc', label: 'Sender (Z\u2013A)' },")
        js.AppendLine("  { key: 'unread', label: 'Unread first' }")
        js.AppendLine("];")
        js.AppendLine("function sortColumnCards(colId, sortKey) {")
        js.AppendLine("  columnSortState[colId] = sortKey;")
        js.AppendLine("  var colEl = document.querySelector('.column[data-column=""' + CSS.escape(colId) + '""]'); if (!colEl) return;")
        js.AppendLine("  var list = colEl.querySelector('.card-list'); if (!list) return;")
        js.AppendLine("  var cardEls = Array.from(list.querySelectorAll('.card'));")
        js.AppendLine("  cardEls.sort(function(a, b) {")
        js.AppendLine("    var ca = cards.find(function(c) { return c.entryId === a.dataset.entryId; });")
        js.AppendLine("    var cb = cards.find(function(c) { return c.entryId === b.dataset.entryId; });")
        js.AppendLine("    if (!ca || !cb) return 0;")
        js.AppendLine("    switch (sortKey) {")
        js.AppendLine("      case 'date-desc': return (cb.date || '').localeCompare(ca.date || '');")
        js.AppendLine("      case 'date-asc': return (ca.date || '').localeCompare(cb.date || '');")
        js.AppendLine("      case 'subject-asc': return (ca.subject || '').localeCompare(cb.subject || '', undefined, { sensitivity: 'base' });")
        js.AppendLine("      case 'subject-desc': return (cb.subject || '').localeCompare(ca.subject || '', undefined, { sensitivity: 'base' });")
        js.AppendLine("      case 'sender-asc': return (ca.senderName || '').localeCompare(cb.senderName || '', undefined, { sensitivity: 'base' });")
        js.AppendLine("      case 'sender-desc': return (cb.senderName || '').localeCompare(ca.senderName || '', undefined, { sensitivity: 'base' });")
        js.AppendLine("      case 'unread': return (ca.isRead === cb.isRead) ? 0 : (ca.isRead ? 1 : -1);")
        js.AppendLine("      default: return 0;")
        js.AppendLine("    }")
        js.AppendLine("  });")
        js.AppendLine("  cardEls.forEach(function(el) { list.appendChild(el); });")
        js.AppendLine("  var dd = colEl.querySelector('.sort-dropdown');")
        js.AppendLine("  if (dd) dd.querySelectorAll('.sort-check').forEach(function(ck) { ck.textContent = ck.parentElement.dataset.sortKey === sortKey ? '\u2713' : ''; });")
        js.AppendLine("  showToast('Sorted by ' + SORT_OPTIONS.find(function(o) { return o.key === sortKey; }).label);")
        js.AppendLine("}")
        js.AppendLine()
        ' --- Render Board (folder column + category columns) ---
        js.AppendLine("function renderBoard() {")
        js.AppendLine("  var board = document.getElementById('board');")
        js.AppendLine("  board.innerHTML = '';")
        js.AppendLine()
        ' Folder column (only if folders exist)
        js.AppendLine("  var allFolderNames = getAllFolderNames();")
        js.AppendLine("  if (allFolderNames.length > 0) {")
        js.AppendLine("    var folderCol = document.createElement('div');")
        js.AppendLine("    folderCol.className = 'column folder-column';")
        js.AppendLine("    folderCol.dataset.column = '__folders__';")
        js.AppendLine("    var fHeader = document.createElement('div'); fHeader.className = 'column-header'; fHeader.draggable = false; fHeader.style.cursor = 'default';")
        js.AppendLine("    var fTitle = document.createElement('span'); fTitle.textContent = '\uD83D\uDCC1 Folders';")
        js.AppendLine("    var fCount = document.createElement('span'); fCount.className = 'count'; fCount.textContent = allFolderNames.length;")
        js.AppendLine("    fHeader.appendChild(fTitle); fHeader.appendChild(fCount);")
        js.AppendLine("    folderCol.appendChild(fHeader);")
        js.AppendLine("    var fList = document.createElement('div'); fList.className = 'card-list'; fList.dataset.column = '__folders__';")
        ' Show All button (only visible when a folder is selected)
        js.AppendLine("    var showAllBtn = document.createElement('div');")
        js.AppendLine("    showAllBtn.id = 'folderShowAll';")
        js.AppendLine("    showAllBtn.className = 'folder-tile show-all';")
        js.AppendLine("    showAllBtn.style.display = selectedFolder ? '' : 'none';")
        js.AppendLine("    showAllBtn.innerHTML = '<span class=""folder-icon"">\u2716</span><span class=""folder-name"">Show All</span>';")
        js.AppendLine("    showAllBtn.addEventListener('click', function(){ selectedFolder = null; applyFolderFilter(); showToast('Showing all cards'); });")
        js.AppendLine("    fList.appendChild(showAllBtn);")
        ' Toggle hide/show foldered cards button
        js.AppendLine("    var toggleHideBtn = document.createElement('div');")
        js.AppendLine("    toggleHideBtn.id = 'folderToggleHide';")
        js.AppendLine("    toggleHideBtn.className = 'folder-tile show-foldered-toggle';")
        js.AppendLine("    toggleHideBtn.style.display = selectedFolder ? 'none' : '';")
        js.AppendLine("    toggleHideBtn.innerHTML = '<span class=""folder-icon"">' + (hideFoldered ? '\uD83D\uDC41' : '\uD83D\uDEAB') + '</span><span class=""folder-name"">' + (hideFoldered ? 'Show foldered cards' : 'Hide foldered cards') + '</span>';")
        js.AppendLine("    toggleHideBtn.addEventListener('click', function(){ hideFoldered = !hideFoldered; applyFolderFilter(); showToast(hideFoldered ? 'Hiding foldered cards' : 'Showing all cards'); });")
        js.AppendLine("    fList.appendChild(toggleHideBtn);")
        js.AppendLine("    allFolderNames.forEach(function(fn){ fList.appendChild(createFolderTile(fn)); });")
        js.AppendLine("    folderCol.appendChild(fList);")
        js.AppendLine("    board.appendChild(folderCol);")
        js.AppendLine("  }")
        js.AppendLine()
        ' Category columns — ALL cards shown (folder filter hides via CSS class)
        js.AppendLine("  COLUMNS.forEach(function(col) {")
        js.AppendLine("    var colEl = document.createElement('div');")
        js.AppendLine("    colEl.className = 'column';")
        js.AppendLine("    colEl.dataset.column = col.id;")
        js.AppendLine("    var colCards = cards.filter(function(c) { return cardMatchesColumn(c, col); });")
        js.AppendLine("    var headerDiv = document.createElement('div'); headerDiv.className = 'column-header';")
        js.AppendLine("    var dot = document.createElement('span'); dot.className = 'dot'; dot.style.background = col.color;")
        js.AppendLine("    var titleSpan = document.createElement('span'); titleSpan.textContent = col.title;")
        js.AppendLine("    var countSpan = document.createElement('span'); countSpan.className = 'count'; countSpan.textContent = colCards.length;")
        js.AppendLine("    headerDiv.appendChild(dot); headerDiv.appendChild(titleSpan); headerDiv.appendChild(countSpan);")
        js.AppendLine()
        ' Sort button
        js.AppendLine("    var sortBtnWrapper = document.createElement('div'); sortBtnWrapper.style.position = 'relative'; sortBtnWrapper.style.display = 'inline-flex';")
        js.AppendLine("    var sortBtn = document.createElement('button'); sortBtn.className = 'sort-btn'; sortBtn.title = 'Sort cards';")
        js.AppendLine("    sortBtn.innerHTML = '<svg width=""14"" height=""14"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round""><path d=""M7 15l5 5 5-5""/><path d=""M7 9l5-5 5 5""/></svg>';")
        js.AppendLine("    var sortDD = document.createElement('div'); sortDD.className = 'sort-dropdown';")
        js.AppendLine("    var currentSort = columnSortState[col.id] || '';")
        js.AppendLine("    SORT_OPTIONS.forEach(function(opt) {")
        js.AppendLine("      var item = document.createElement('div'); item.className = 'sort-option'; item.dataset.sortKey = opt.key;")
        js.AppendLine("      var check = document.createElement('span'); check.className = 'sort-check'; check.textContent = (opt.key === currentSort) ? '\u2713' : '';")
        js.AppendLine("      var labelSpan = document.createElement('span'); labelSpan.textContent = opt.label;")
        js.AppendLine("      item.appendChild(check); item.appendChild(labelSpan);")
        js.AppendLine("      item.addEventListener('click', function(e) { e.stopPropagation(); sortColumnCards(col.id, opt.key); sortDD.classList.remove('open'); });")
        js.AppendLine("      sortDD.appendChild(item);")
        js.AppendLine("    });")
        js.AppendLine("    sortBtn.addEventListener('click', function(e) { e.stopPropagation(); document.querySelectorAll('.sort-dropdown.open').forEach(function(dd) { if (dd !== sortDD) dd.classList.remove('open'); }); sortDD.classList.toggle('open'); });")
        js.AppendLine("    sortBtnWrapper.appendChild(sortBtn); sortBtnWrapper.appendChild(sortDD); headerDiv.appendChild(sortBtnWrapper);")
        js.AppendLine()
        js.AppendLine("    colEl.appendChild(headerDiv);")
        js.AppendLine("    var list = document.createElement('div'); list.className = 'card-list'; list.dataset.column = col.id;")
        js.AppendLine("    colEl.appendChild(list);")
        js.AppendLine("    colCards.forEach(function(card) { list.appendChild(createCardEl(card, col)); });")
        js.AppendLine("    setupDropZone(list);")
        js.AppendLine("    board.appendChild(colEl);")
        js.AppendLine("    if (columnSortState[col.id]) sortColumnCards(col.id, columnSortState[col.id]);")
        js.AppendLine("  });")
        js.AppendLine("  applyFieldVisibility();")
        js.AppendLine("  applySearchFilter();")
        js.AppendLine("  applyFolderFilter();")
        js.AppendLine("  setupColumnDrag();")
        js.AppendLine("  renderFolderManageList();")
        js.AppendLine("}")
        js.AppendLine()
        ' --- Column drag (skip folder column) ---
        js.AppendLine("var colDragSrc=null;")
        js.AppendLine("function setupColumnDrag(){")
        js.AppendLine("  document.querySelectorAll('.column:not(.folder-column)').forEach(function(colEl){")
        js.AppendLine("    var header=colEl.querySelector('.column-header');")
        js.AppendLine("    header.draggable=true;")
        js.AppendLine("    header.addEventListener('dragstart',function(e){ if(e.target.closest('.card') || e.target.closest('.sort-btn') || e.target.closest('.sort-dropdown')){e.preventDefault();return;} colDragSrc=colEl; colEl.classList.add('col-dragging'); e.dataTransfer.effectAllowed='move'; e.dataTransfer.setData('text/x-column',colEl.dataset.column); });")
        js.AppendLine("    header.addEventListener('dragend',function(){ colDragSrc=null; document.querySelectorAll('.column').forEach(function(c){c.classList.remove('col-dragging','col-drag-over');}); });")
        js.AppendLine("    colEl.addEventListener('dragover',function(e){ if(!colDragSrc||colDragSrc===colEl)return; e.preventDefault(); e.dataTransfer.dropEffect='move'; document.querySelectorAll('.column.col-drag-over').forEach(function(c){c.classList.remove('col-drag-over');}); colEl.classList.add('col-drag-over'); });")
        js.AppendLine("    colEl.addEventListener('dragleave',function(e){ if(!colEl.contains(e.relatedTarget))colEl.classList.remove('col-drag-over'); });")
        js.AppendLine("    colEl.addEventListener('drop',function(e){")
        js.AppendLine("      if(!colDragSrc||colDragSrc===colEl)return; e.preventDefault();")
        js.AppendLine("      var board=document.getElementById('board'); var allCols=[...board.querySelectorAll('.column:not(.folder-column)')];")
        js.AppendLine("      var fromIdx=allCols.indexOf(colDragSrc); var toIdx=allCols.indexOf(colEl);")
        js.AppendLine("      if(fromIdx<0||toIdx<0)return;")
        js.AppendLine("      if(fromIdx<toIdx){board.insertBefore(colDragSrc,colEl.nextSibling);}else{board.insertBefore(colDragSrc,colEl);}")
        js.AppendLine("      document.querySelectorAll('.column').forEach(function(c){c.classList.remove('col-dragging','col-drag-over');});")
        js.AppendLine("      colDragSrc=null;")
        js.AppendLine("      var newOrder=[]; board.querySelectorAll('.column:not(.folder-column)').forEach(function(c){newOrder.push(c.dataset.column);});")
        js.AppendLine("      COLUMNS.sort(function(a,b){return newOrder.indexOf(a.id)-newOrder.indexOf(b.id);});")
        js.AppendLine("      window.chrome.webview.postMessage(JSON.stringify({action:'setColumnOrder',columns:newOrder}));")
        js.AppendLine("      showToast('Column order updated');")
        js.AppendLine("    });")
        js.AppendLine("  });")
        js.AppendLine("}")
        js.AppendLine()
        ' --- Create card element ---
        js.AppendLine("function createCardEl(card, col) {")
        js.AppendLine("  var el = document.createElement('div');")
        js.AppendLine("  var cls = 'card';")
        js.AppendLine("  if (card.isConversation) cls += ' conversation';")
        js.AppendLine("  if (card.isFlagged) cls += ' flagged';")
        js.AppendLine("  el.className = cls;")
        js.AppendLine("  el.dataset.entryId = card.entryId; el.dataset.id = card.id; el.draggable = true;")
        js.AppendLine()
        js.AppendLine("  var subjectDiv = document.createElement('div'); subjectDiv.className = 'card-subject';")
        js.AppendLine("  if (!card.isRead) { var dot = document.createElement('span'); dot.className = 'unread-dot'; subjectDiv.appendChild(dot); }")
        js.AppendLine("  if (card.isFlagged) { var fb = document.createElement('span'); fb.className = 'flag-badge'; fb.textContent = '\u2691 '; subjectDiv.appendChild(fb); }")
        js.AppendLine("  if (card.isConversation) { var ci = document.createElement('span'); ci.className = 'conv-icon'; ci.textContent = '\uD83D\uDCEC '; subjectDiv.appendChild(ci); }")
        js.AppendLine("  subjectDiv.appendChild(document.createTextNode(card.subject));")
        js.AppendLine("  el.appendChild(subjectDiv);")
        js.AppendLine()
        ' Multi-category badges
        js.AppendLine("  if (card.allCategories && card.allCategories.length > 1) {")
        js.AppendLine("    var catsDiv = document.createElement('div'); catsDiv.className = 'card-categories';")
        js.AppendLine("    card.allCategories.forEach(function(ac) { var badge = document.createElement('span'); badge.className = 'card-cat-badge'; badge.textContent = ac.name; badge.style.background = ac.color; if (ac.name !== col.id) badge.style.opacity = '0.55'; catsDiv.appendChild(badge); });")
        js.AppendLine("    el.appendChild(catsDiv);")
        js.AppendLine("  }")
        js.AppendLine()
        ' Folder badge
        js.AppendLine("  if (card.boardFolder) {")
        js.AppendLine("    var folderBadge = document.createElement('div'); folderBadge.className = 'card-folder-badge';")
        js.AppendLine("    folderBadge.textContent = '\uD83D\uDCC1 ' + card.boardFolder;")
        js.AppendLine("    el.appendChild(folderBadge);")
        js.AppendLine("  }")
        js.AppendLine()
        js.AppendLine("  var senderDiv = document.createElement('div'); senderDiv.className = 'card-sender'; senderDiv.textContent = card.senderName || card.senderEmail; el.appendChild(senderDiv);")
        js.AppendLine()
        js.AppendLine("  var metaDiv = document.createElement('div'); metaDiv.className = 'card-meta';")
        js.AppendLine("  var dateSpan = document.createElement('span'); dateSpan.className = 'card-date'; dateSpan.textContent = card.date; metaDiv.appendChild(dateSpan);")
        js.AppendLine("  var msgSpan = document.createElement('span'); msgSpan.className = 'card-msg-count';")
        js.AppendLine("  if (card.isGrouped && card.messages > 1) msgSpan.textContent = card.messages + ' mail' + (card.messages !== 1 ? 's' : '');")
        js.AppendLine("  metaDiv.appendChild(msgSpan);")
        js.AppendLine("  if (card.isFlagged || card.isCompleted) {")
        js.AppendLine("    var flagSpan = document.createElement('span'); flagSpan.className = 'flag-meta' + (card.isCompleted ? ' completed' : '');")
        js.AppendLine("    var flagText = card.flagRequest || (card.isFlagged ? 'Flagged' : 'Done'); if (card.flagDueDate) flagText += ' \u2022 ' + card.flagDueDate;")
        js.AppendLine("    flagSpan.textContent = flagText; metaDiv.appendChild(flagSpan);")
        js.AppendLine("  }")
        js.AppendLine("  el.appendChild(metaDiv);")
        js.AppendLine()
        js.AppendLine("  var summaryDiv = document.createElement('div'); summaryDiv.className = card.summary ? 'card-summary' : 'card-summary loading';")
        js.AppendLine("  summaryDiv.dataset.entryId = card.entryId; summaryDiv.textContent = card.summary || 'Generating summary\u2026'; el.appendChild(summaryDiv);")
        js.AppendLine()
        ' Footer buttons
        js.AppendLine("  var footerDiv = document.createElement('div'); footerDiv.className = 'card-footer';")
        js.AppendLine("  var openBtn = document.createElement('button'); openBtn.className = 'card-btn'; openBtn.dataset.action = 'open'; openBtn.dataset.entryId = card.entryId; openBtn.title = 'Open in Outlook';")
        js.AppendLine("  openBtn.innerHTML = '<svg class=""icon-svg"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round""><path d=""M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z""/><polyline points=""22,6 12,13 2,6""/></svg><span class=""tip"">Open in Outlook</span>';")
        js.AppendLine("  footerDiv.appendChild(openBtn);")
        js.AppendLine("  var unmarkBtn = document.createElement('button'); unmarkBtn.className = 'card-btn'; unmarkBtn.dataset.action = 'unmark'; unmarkBtn.dataset.entryId = card.entryId; unmarkBtn.dataset.column = col.id;")
        js.AppendLine("  unmarkBtn.title = 'Remove from ' + col.title;")
        js.AppendLine("  unmarkBtn.innerHTML = '<svg class=""icon-svg"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round""><line x1=""18"" y1=""6"" x2=""6"" y2=""18""/><line x1=""6"" y1=""6"" x2=""18"" y2=""18""/></svg><span class=""tip"">Remove from ' + escHtml(col.title) + '</span>';")
        js.AppendLine("  footerDiv.appendChild(unmarkBtn);")
        ' Remove-from-folder button (only when card has a folder)
        js.AppendLine("  if (card.boardFolder) {")
        js.AppendLine("    var removeFolderBtn = document.createElement('button'); removeFolderBtn.className = 'card-btn folder-remove-btn';")
        js.AppendLine("    removeFolderBtn.dataset.action = 'removeFromFolder'; removeFolderBtn.dataset.entryId = card.entryId;")
        js.AppendLine("    removeFolderBtn.title = 'Remove from ' + card.boardFolder;")
        js.AppendLine("    removeFolderBtn.innerHTML = '<svg class=""icon-svg"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round""><path d=""M3 7v13a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V7""/><path d=""M8 7V5a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2""/><line x1=""1"" y1=""7"" x2=""23"" y2=""7""/></svg><span class=""tip"">Remove from folder</span>';")
        js.AppendLine("    removeFolderBtn.addEventListener('click', function(e){ e.stopPropagation(); window.chrome.webview.postMessage(JSON.stringify({action:'removeFromFolder',entryId:card.entryId})); showToast('Removed from '+card.boardFolder); });")
        js.AppendLine("    footerDiv.insertBefore(removeFolderBtn, footerDiv.firstChild);")
        js.AppendLine("  }")
        js.AppendLine("  el.appendChild(footerDiv);")
        js.AppendLine()
        ' Drag events
        js.AppendLine("  el.addEventListener('dragstart', function(e) { isDragging = true; el.classList.add('dragging'); e.dataTransfer.effectAllowed = 'copyMove'; e.dataTransfer.setData('text/plain', card.entryId); updateDragModeVisuals(); });")
        js.AppendLine("  el.addEventListener('dragend', function() { isDragging = false; el.classList.remove('dragging'); clearAllIndicators(); document.getElementById('dragModeLabel').classList.remove('visible'); });")
        js.AppendLine()
        ' Button click handler
        js.AppendLine("  el.querySelectorAll('.card-btn').forEach(function(btn) {")
        js.AppendLine("    btn.addEventListener('click', function(e) {")
        js.AppendLine("      e.stopPropagation(); var act = btn.dataset.action; var eid = btn.dataset.entryId;")
        js.AppendLine("      if (act === 'open') window.chrome.webview.postMessage(JSON.stringify({ action: 'open', entryId: eid }));")
        js.AppendLine("      else if (act === 'unmark') { var colId = btn.dataset.column || ''; window.chrome.webview.postMessage(JSON.stringify({ action: 'unmark', entryId: eid, category: colId })); }")
        js.AppendLine("    });")
        js.AppendLine("  });")
        js.AppendLine("  return el;")
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("function escHtml(str) { var d = document.createElement('div'); d.textContent = str; return d.innerHTML; }")
        js.AppendLine()
        ' --- Drag & Drop ---
        js.AppendLine("function setupDropZone(list) {")
        js.AppendLine("  list.addEventListener('dragover', function(e) { e.preventDefault(); ctrlHeld = e.ctrlKey; e.dataTransfer.dropEffect = ctrlHeld ? 'copy' : 'move'; var col = list.closest('.column'); col.classList.add('drag-over'); if (ctrlHeld) col.classList.add('drag-over-add'); else col.classList.remove('drag-over-add'); showDropIndicator(list, getDragAfterElement(list, e.clientY)); });")
        js.AppendLine("  list.addEventListener('dragleave', function(e) { if (!list.contains(e.relatedTarget)) { var col = list.closest('.column'); col.classList.remove('drag-over'); col.classList.remove('drag-over-add'); clearIndicators(list); } });")
        js.AppendLine("  list.addEventListener('drop', function(e) {")
        js.AppendLine("    e.preventDefault(); var isAddMode = e.ctrlKey; ctrlHeld = e.ctrlKey;")
        js.AppendLine("    var column = list.closest('.column'); column.classList.remove('drag-over'); column.classList.remove('drag-over-add');")
        js.AppendLine("    var entryId = e.dataTransfer.getData('text/plain');")
        js.AppendLine("    var cardEl = document.querySelector('.card[data-entry-id=""' + CSS.escape(entryId) + '""]');")
        js.AppendLine("    if (!cardEl) return;")
        js.AppendLine("    var oldList = cardEl.closest('.card-list');")
        js.AppendLine("    var oldColumnId = oldList ? oldList.dataset.column : '';")
        js.AppendLine("    var newColumnId = list.dataset.column;")
        js.AppendLine("    clearAllIndicators();")
        js.AppendLine("    if (oldColumnId === newColumnId && !isAddMode) { var after = getDragAfterElement(list, e.clientY); if (after) list.insertBefore(cardEl, after); else list.appendChild(cardEl); saveState(); return; }")
        js.AppendLine("    if (isAddMode) {")
        js.AppendLine("      var card = cards.find(function(c) { return c.entryId === entryId; });")
        js.AppendLine("      if (card && (card.isFlagged || card.isCompleted)) { var targetCol = COLUMNS.find(function(cc) { return cc.id === newColumnId; }); if (targetCol && targetCol.isFlaggedColumn && card.flagBucket === newColumnId) { showToast('Already in ' + newColumnId); return; } }")
        js.AppendLine("      if (card && card.allCategories && card.allCategories.some(function(ac) { return ac.name === newColumnId; })) { showToast('Already in ' + newColumnId); return; }")
        js.AppendLine("      window.chrome.webview.postMessage(JSON.stringify({ action: 'moveAdd', entryId: entryId, newCategory: newColumnId }));")
        js.AppendLine("      var addCol = COLUMNS.find(function(c) { return c.id === newColumnId; });")
        js.AppendLine("      showToast('+ Added to ' + (addCol ? addCol.title : newColumnId));")
        js.AppendLine("    } else {")
        js.AppendLine("      var after = getDragAfterElement(list, e.clientY); if (after) list.insertBefore(cardEl, after); else list.appendChild(cardEl);")
        js.AppendLine("      var card = cards.find(function(c) { return c.entryId === entryId; }); if (card) card.category = newColumnId;")
        js.AppendLine("      updateCounts(); saveState();")
        js.AppendLine("      window.chrome.webview.postMessage(JSON.stringify({ action: 'move', entryId: entryId, newCategory: newColumnId }));")
        js.AppendLine("      var newCol = COLUMNS.find(function(c) { return c.id === newColumnId; });")
        js.AppendLine("      showToast('Moved to ' + (newCol ? newCol.title : newColumnId));")
        js.AppendLine("    }")
        js.AppendLine("  });")
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("function getDragAfterElement(list, y) { var els = Array.from(list.querySelectorAll('.card:not(.dragging)')); var closest = null, closestOffset = Number.POSITIVE_INFINITY; els.forEach(function(child) { var box = child.getBoundingClientRect(); var offset = y - box.top - box.height / 2; if (offset < 0 && -offset < closestOffset) { closestOffset = -offset; closest = child; } }); return closest; }")
        js.AppendLine("function showDropIndicator(list, after) { clearIndicators(list); var ind = document.createElement('div'); ind.className = 'drop-indicator visible' + (ctrlHeld ? ' add-mode' : ''); if (after) list.insertBefore(ind, after); else list.appendChild(ind); }")
        js.AppendLine("function clearIndicators(list) { list.querySelectorAll('.drop-indicator').forEach(function(el) { el.remove(); }); }")
        js.AppendLine("function clearAllIndicators() { document.querySelectorAll('.drop-indicator').forEach(function(el) { el.remove(); }); document.querySelectorAll('.column.drag-over').forEach(function(el) { el.classList.remove('drag-over'); }); document.querySelectorAll('.column.drag-over-add').forEach(function(el) { el.classList.remove('drag-over-add'); }); }")
        js.AppendLine()
        js.AppendLine("function updateCounts() { document.querySelectorAll('.column:not(.folder-column)').forEach(function(colEl) { var list = colEl.querySelector('.card-list'); if (!list) return; var count = list.querySelectorAll('.card:not(.hidden):not(.folder-hidden)').length; var countEl = colEl.querySelector('.count'); if (countEl) countEl.textContent = count; }); }")
        js.AppendLine()
        js.AppendLine("function showToast(message) { var container = document.getElementById('toastContainer'); var toast = document.createElement('div'); toast.className = 'toast'; toast.textContent = message; container.appendChild(toast); setTimeout(function() { toast.classList.add('fade-out'); setTimeout(function() { toast.remove(); }, 300); }, 2500); }")
        js.AppendLine()
        ' --- Custom confirm/prompt dialogs (no browser chrome text) ---
        js.AppendLine("function showConfirm(message, onYes) {")
        js.AppendLine("  var overlay = document.createElement('div');")
        js.AppendLine("  overlay.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:var(--overlay-bg);z-index:2000;display:flex;align-items:center;justify-content:center';")
        js.AppendLine("  var box = document.createElement('div');")
        js.AppendLine("  box.style.cssText = 'background:var(--settings-bg);border:1px solid var(--settings-border);border-radius:10px;padding:20px 24px;min-width:320px;max-width:420px;box-shadow:0 8px 32px rgba(0,0,0,0.2);font-size:13px;color:var(--text)';")
        js.AppendLine("  var msg = document.createElement('div'); msg.style.cssText = 'margin-bottom:16px;line-height:1.5'; msg.textContent = message;")
        js.AppendLine("  var btnRow = document.createElement('div'); btnRow.style.cssText = 'display:flex;justify-content:flex-end;gap:8px';")
        js.AppendLine("  var cancelBtn = document.createElement('button'); cancelBtn.textContent = 'Cancel';")
        js.AppendLine("  cancelBtn.style.cssText = 'padding:6px 16px;border:1px solid var(--input-border);border-radius:6px;background:var(--input-bg);color:var(--text);font-size:13px;cursor:pointer';")
        js.AppendLine("  var yesBtn = document.createElement('button'); yesBtn.textContent = 'Delete';")
        js.AppendLine("  yesBtn.style.cssText = 'padding:6px 16px;border:none;border-radius:6px;background:#ef4444;color:#fff;font-size:13px;cursor:pointer;font-weight:600';")
        js.AppendLine("  cancelBtn.addEventListener('click', function(){ overlay.remove(); });")
        js.AppendLine("  yesBtn.addEventListener('click', function(){ overlay.remove(); onYes(); });")
        js.AppendLine("  overlay.addEventListener('click', function(e){ if(e.target===overlay) overlay.remove(); });")
        js.AppendLine("  btnRow.appendChild(cancelBtn); btnRow.appendChild(yesBtn);")
        js.AppendLine("  box.appendChild(msg); box.appendChild(btnRow);")
        js.AppendLine("  overlay.appendChild(box); document.body.appendChild(overlay); yesBtn.focus();")
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("function showPrompt(message, defaultVal, onOk) {")
        js.AppendLine("  var overlay = document.createElement('div');")
        js.AppendLine("  overlay.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:var(--overlay-bg);z-index:2000;display:flex;align-items:center;justify-content:center';")
        js.AppendLine("  var box = document.createElement('div');")
        js.AppendLine("  box.style.cssText = 'background:var(--settings-bg);border:1px solid var(--settings-border);border-radius:10px;padding:20px 24px;min-width:320px;max-width:420px;box-shadow:0 8px 32px rgba(0,0,0,0.2);font-size:13px;color:var(--text)';")
        js.AppendLine("  var msg = document.createElement('div'); msg.style.cssText = 'margin-bottom:10px;line-height:1.5'; msg.textContent = message;")
        js.AppendLine("  var input = document.createElement('input'); input.type = 'text'; input.value = defaultVal || '';")
        js.AppendLine("  input.style.cssText = 'width:100%;padding:6px 10px;border:1px solid var(--input-border);border-radius:6px;background:var(--input-bg);color:var(--text);font-size:13px;outline:none;margin-bottom:16px';")
        js.AppendLine("  var btnRow = document.createElement('div'); btnRow.style.cssText = 'display:flex;justify-content:flex-end;gap:8px';")
        js.AppendLine("  var cancelBtn = document.createElement('button'); cancelBtn.textContent = 'Cancel';")
        js.AppendLine("  cancelBtn.style.cssText = 'padding:6px 16px;border:1px solid var(--input-border);border-radius:6px;background:var(--input-bg);color:var(--text);font-size:13px;cursor:pointer';")
        js.AppendLine("  var okBtn = document.createElement('button'); okBtn.textContent = 'OK';")
        js.AppendLine("  okBtn.style.cssText = 'padding:6px 16px;border:none;border-radius:6px;background:#3b82f6;color:#fff;font-size:13px;cursor:pointer;font-weight:600';")
        js.AppendLine("  cancelBtn.addEventListener('click', function(){ overlay.remove(); });")
        js.AppendLine("  okBtn.addEventListener('click', function(){ overlay.remove(); onOk(input.value); });")
        js.AppendLine("  input.addEventListener('keydown', function(e){ if(e.key==='Enter'){e.preventDefault(); overlay.remove(); onOk(input.value);} if(e.key==='Escape') overlay.remove(); });")
        js.AppendLine("  overlay.addEventListener('click', function(e){ if(e.target===overlay) overlay.remove(); });")
        js.AppendLine("  btnRow.appendChild(cancelBtn); btnRow.appendChild(okBtn);")
        js.AppendLine("  box.appendChild(msg); box.appendChild(input); box.appendChild(btnRow);")
        js.AppendLine("  overlay.appendChild(box); document.body.appendChild(overlay); input.focus(); input.select();")
        js.AppendLine("}")
        ' --- Search & Filter ---
        js.AppendLine("document.getElementById('searchInput').addEventListener('input', applySearchFilter);")
        js.AppendLine("document.getElementById('columnFilter').addEventListener('change', function() { applySearchFilter(); window.chrome.webview.postMessage(JSON.stringify({ action: 'setColumnFilter', filter: document.getElementById('columnFilter').value })); });")
        js.AppendLine()
        js.AppendLine("function applySearchFilter() {")
        js.AppendLine("  var query = document.getElementById('searchInput').value.toLowerCase().trim();")
        js.AppendLine("  var colFilter = document.getElementById('columnFilter').value;")
        js.AppendLine("  document.querySelectorAll('.column:not(.folder-column)').forEach(function(colEl) { var colId = colEl.dataset.column; colEl.style.display = (colFilter === 'all' || colFilter === colId) ? '' : 'none'; });")
        js.AppendLine("  document.querySelectorAll('.card').forEach(function(cardEl) {")
        js.AppendLine("    var card = cards.find(function(c) { return c.entryId === cardEl.dataset.entryId; }); if (!card) return;")
        js.AppendLine("    if (!query) { cardEl.classList.remove('hidden'); return; }")
        js.AppendLine("    var match = (card.subject || '').toLowerCase().indexOf(query) >= 0 || (card.senderName || '').toLowerCase().indexOf(query) >= 0 || (card.senderEmail || '').toLowerCase().indexOf(query) >= 0;")
        js.AppendLine("    if (match) cardEl.classList.remove('hidden'); else cardEl.classList.add('hidden');")
        js.AppendLine("  });")
        js.AppendLine("  updateCounts();")
        js.AppendLine("}")
        js.AppendLine()
        ' --- Messages from VB.NET ---
        js.AppendLine("if (window.chrome && window.chrome.webview) {")
        js.AppendLine("  window.chrome.webview.addEventListener('message', function(event) {")
        js.AppendLine("    var msg = (typeof event.data === 'string') ? JSON.parse(event.data) : event.data;")
        js.AppendLine("    if (msg.action === 'updateSummary') {")
        js.AppendLine("      document.querySelectorAll('.card-summary').forEach(function(el) { if (el.dataset.entryId === msg.entryId) { if (msg.summary) { el.textContent = msg.summary; el.classList.remove('loading'); } else { el.textContent = 'Generating summary\u2026'; el.classList.add('loading'); } } });")
        js.AppendLine("      var card = cards.find(function(c) { return c.entryId === msg.entryId; }); if (card) card.summary = msg.summary;")
        js.AppendLine("    } else if (msg.action === 'removeCard') {")
        js.AppendLine("      document.querySelectorAll('.card').forEach(function(el) { if (el.dataset.entryId === msg.entryId) el.remove(); });")
        js.AppendLine("      updateCounts(); cards = cards.filter(function(c) { return c.entryId !== msg.entryId; }); showToast('Category removed');")
        js.AppendLine("    } else if (msg.action === 'reloadData') {")
        js.AppendLine("      var d = (typeof msg.data === 'string') ? JSON.parse(msg.data) : msg.data; init(d); showToast('Board reloaded');")
        js.AppendLine("    } else if (msg.action === 'updateCard') {")
        js.AppendLine("      var card = cards.find(function(c) { return c.entryId === msg.entryId; });")
        js.AppendLine("      if (card) { card.category = msg.category; card.categoryColor = msg.categoryColor; card.allCategories = msg.allCategories || [];")
        js.AppendLine("        card.isFlagged = msg.isFlagged || false; card.isCompleted = msg.isCompleted || false;")
        js.AppendLine("        card.flagRequest = msg.flagRequest || ''; card.flagDueDate = msg.flagDueDate || '';")
        js.AppendLine("        card.flagBucket = msg.flagBucket || ''; card.boardFolder = msg.boardFolder || ''; }")
        js.AppendLine("      renderBoard();")
        js.AppendLine("    } else if (msg.action === 'updateFolderLabel') {")
        js.AppendLine("      var lbl = document.getElementById('trackedFolderLabel');")
        js.AppendLine("      if (lbl) lbl.textContent = msg.label || 'Inbox';")
        js.AppendLine("      showToast('\uD83D\uDCC2 Now tracking: ' + (msg.label || 'Inbox'));")
        js.AppendLine("    }")
        js.AppendLine("  });")
        js.AppendLine("}")
        js.AppendLine()
        ' --- Init ---
        js.AppendLine("function startBoard() {")
        js.AppendLine("  try { initTheme(); } catch(e) { console.error('initTheme error:', e); }")
        js.AppendLine("  requestAnimationFrame(function() { requestAnimationFrame(function() { init(INIT_DATA); }); });")
        js.AppendLine("}")
        js.AppendLine("if (document.readyState === 'loading') { document.addEventListener('DOMContentLoaded', startBoard); } else { startBoard(); }")

        Return js.ToString()
    End Function

#End Region



End Class
