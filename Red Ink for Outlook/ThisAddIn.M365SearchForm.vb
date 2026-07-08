' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: M365SearchForm.vb
' Purpose: Lightweight harness for SharedLibrary.M365Service. Lets the user
'          enter a query, lists matching emails, and opens the selected
'          message inside Outlook itself (or, as a last resort, in the
'          browser via webLink).
'
' Wire-up: From any ribbon button (or Immediate window), call:
'              M365SearchTest.Show(_context)
' =============================================================================

Option Explicit On
Option Strict On

Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Drawing
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Markdig
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports Outlook = Microsoft.Office.Interop.Outlook

''' <summary>
''' Modeless Red-Ink-styled M365 mail search window with open-in-Outlook.
''' </summary>
Public Class M365SearchTestForm
    Inherits Form
    Implements IProgress(Of M365SearchProgress)

    Private Const AISearch_MaxIterations As Integer = 30
    Private Const AISearch_MinCandidates As Integer = 10
    Private Const AISearch_MaxCandidateHits As Integer = 300
    Private Const AISearch_ReviewBatchSize As Integer = 12
    Private Const AISearch_BodyCap_PerCandidate As Integer = 6000
    Private Const AISearch_BodyCap_FinalSummary As Integer = 8000

    Private Const M365Search_SummaryBatchSize As Integer = 5
    Private Const M365Search_SummaryBodyCap As Integer = 2500

    Private NotInheritable Class AiResolvedMail
        Public Property GlobalRef As Integer
        Public Property Hit As M365SearchHit
        Public Property Message As M365Message
    End Class

    Private NotInheritable Class AiFinalSelection
        Public Property Summary As String = ""
        Public Property OrderedGlobalRefs As New List(Of Integer)()
    End Class

    Private NotInheritable Class AiSearchInvocation
        Public Property Query As String = ""
        Public Property Sources As M365SearchSources = M365SearchSources.Mail
        Public Property MaxPerSource As Integer = 25
        Public Property FromDate As Date?
        Public Property ToDate As Date?
        Public Property KqlExtra As String = ""
        Public Property NextFromIndex As Integer = 0
        Public Property Exhausted As Boolean

        Public Function StableKey() As String
            Return Query.Trim().ToLowerInvariant() & "|" &
                   Sources.ToString() & "|" &
                   MaxPerSource.ToString() & "|" &
                   If(FromDate.HasValue, FromDate.Value.ToString("yyyy-MM-dd"), "") & "|" &
                   If(ToDate.HasValue, ToDate.Value.ToString("yyyy-MM-dd"), "") & "|" &
                   If(KqlExtra, "").Trim().ToLowerInvariant()
        End Function
    End Class

    Private _aiRecordedSearches As New List(Of AiSearchInvocation)()

    Private ReadOnly _context As ISharedContext
    Private _hits As New List(Of M365SearchHit)()
    Private _cts As CancellationTokenSource
    Private _aiHeartbeat As System.Windows.Forms.Timer
    Private _aiHeartbeatStart As DateTime
    Private _aiHeartbeatPrefix As String = ""
    Private _hasPinnedSummary As Boolean
    Private _aiLastUserPrompt As String = ""
    Private _aiCandidateHits As New List(Of M365SearchHit)()
    Private _aiSearchAvailable As Boolean
    Private _aiSearchUnavailableReason As String = ""
    Private ReadOnly _aiMessageCacheLock As New Object()
    Private ReadOnly _aiMessageCache As New Dictionary(Of String, M365Message)(StringComparer.OrdinalIgnoreCase)

    ' ── UI fields ───────────────────────────────────────────────────────
    Private WithEvents txtQuery As TextBox
    Private WithEvents btnSearch As Button
    Private WithEvents btnOpen As Button
    Private WithEvents btnSignIn As Button
    Private WithEvents btnSignOut As Button
    Private WithEvents btnGetMore As Button
    Private WithEvents numMax As NumericUpDown
    Private WithEvents lvResults As ListView
    Private WithEvents btnClose As Button
    Private headerPanel As Panel
    Private footerPanel As Panel
    Private splitMain As SplitContainer
    Private lblQuery As Label
    Private lblStatus As Label
    Private lblUser As Label
    Private lblMax As Label
    Private WithEvents btnAISearch As Button
    Private txtSummary As WebBrowser
    Private lblSummary As Label
    Private lblAiStats As Label
    Private pb As ProgressBar

    ' Standard Red Ink button padding/margin (matches ShowSelectionForm etc.).
    Private Shared ReadOnly StdButtonPadding As New Padding(8, 4, 5, 4)
    Private Shared ReadOnly StdButtonMargin As New Padding(0, 0, 10, 0)

    Public Sub New(context As ISharedContext)
        If context Is Nothing Then Throw New ArgumentNullException(NameOf(context))
        _context = context
        BuildUi()
    End Sub

    Private ReadOnly _toolTip As New ToolTip() With {
        .ShowAlways = True,
        .AutoPopDelay = 20000,
        .InitialDelay = 400,
        .ReshowDelay = 150
    }
    Private _dialogOwnerScope As System.IDisposable = Nothing

    Protected Overrides Sub OnHandleCreated(e As System.EventArgs)
        MyBase.OnHandleCreated(e)

        If _dialogOwnerScope Is Nothing Then
            _dialogOwnerScope = SharedMethods.PushDialogOwner(Me)
        End If
    End Sub

    Protected Overrides Sub OnHandleDestroyed(e As System.EventArgs)
        Dim scope As System.IDisposable = _dialogOwnerScope
        _dialogOwnerScope = Nothing

        If scope IsNot Nothing Then
            Try
                scope.Dispose()
            Catch
            End Try
        End If

        MyBase.OnHandleDestroyed(e)
    End Sub

    Private Function GetAiHarvestMaxPerCall() As Integer
        Return CInt(Math.Max(1D, Math.Min(CDec(AISearch_MaxCandidateHits), numMax.Value)))
    End Function

    Private Sub ConfigureToolTips()
        _toolTip.SetToolTip(lblMax, "Search: maximum rows requested. AI Search / Get more: preferred maximum hits requested per m365_search call. Overall AI candidate pool is still capped at 100.")
        _toolTip.SetToolTip(numMax, "Search: maximum rows requested. AI Search / Get more: preferred maximum hits requested per m365_search call. Overall AI candidate pool is still capped at 100.")
        _toolTip.SetToolTip(btnSearch, "Run a direct Microsoft 365 mail search and show up to Max hits. Fast, metadata-first search.")
        _toolTip.SetToolTip(btnAISearch, "Run multi-step AI mail search: broad candidate gathering first, then full-text review of the candidate mails. Slower, but more accurate for semantic requests.")
        _toolTip.SetToolTip(btnGetMore, "Run another AI candidate-gathering pass for the current AI query. Already known candidates are passed back to the model, and duplicates are merged out host-side.")
        _toolTip.SetToolTip(btnSignIn, "Sign in to Microsoft 365 so Red Ink can search and read your mailbox through Microsoft Graph.")
        _toolTip.SetToolTip(btnSignOut, "Sign out of Microsoft 365 and clear the current Microsoft Graph session for this add-in.")
    End Sub

    Private Sub ClearAiMessageCache()
        SyncLock _aiMessageCacheLock
            _aiMessageCache.Clear()
        End SyncLock
    End Sub

    Private Shared Function BuildMessageCacheKeys(messageId As String,
                                                  internetMessageId As String) As List(Of String)
        Dim keys As New List(Of String)()
        Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        Dim graphKey As String = NormalizeGraphIdForCompare(messageId)
        If Not String.IsNullOrWhiteSpace(graphKey) Then
            Dim key As String = "id:" & graphKey
            If seen.Add(key) Then keys.Add(key)
        End If

        Dim imidKey As String = NormalizeInternetMessageId(internetMessageId)
        If Not String.IsNullOrWhiteSpace(imidKey) Then
            Dim key As String = "imid:" & imidKey
            If seen.Add(key) Then keys.Add(key)
        End If

        Return keys
    End Function

    Private Sub CacheMessage(message As M365Message)
        If message Is Nothing Then Return

        Dim keys As List(Of String) = BuildMessageCacheKeys(message.Id, message.InternetMessageId)
        If keys.Count = 0 Then Return

        SyncLock _aiMessageCacheLock
            For Each key In keys
                _aiMessageCache(key) = message
            Next
        End SyncLock
    End Sub

    Private Function TryGetCachedMessage(hit As M365SearchHit) As M365Message
        If hit Is Nothing Then Return Nothing

        Dim keys As List(Of String) = BuildMessageCacheKeys(GetHitMessageId(hit), TryGetInternetMessageId(hit))
        If keys.Count = 0 Then Return Nothing

        SyncLock _aiMessageCacheLock
            For Each key In keys
                Dim cached As M365Message = Nothing
                If _aiMessageCache.TryGetValue(key, cached) Then
                    Return cached
                End If
            Next
        End SyncLock

        Return Nothing
    End Function

    Private Function TryGetHitFromListItem(item As ListViewItem) As M365SearchHit
        If item Is Nothing Then Return Nothing

        Dim directHit As M365SearchHit = TryCast(item.Tag, M365SearchHit)
        If directHit IsNot Nothing Then
            Return directHit
        End If

        If item.Tag Is Nothing Then Return Nothing

        Try
            Dim hitIdx As Integer = CInt(item.Tag)
            If hitIdx >= 0 AndAlso hitIdx < _hits.Count Then
                Return _hits(hitIdx)
            End If
        Catch
        End Try

        Return Nothing
    End Function

    ' ════════════════════════════════════════════════════════════════════
    '  UI construction
    ' ════════════════════════════════════════════════════════════════════
    Private Sub BuildUi()
        Me.Text = Globals.ThisAddIn.AN & "- Search Microsoft 365 Mails"
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.MinimumSize = New Size(1100, 540)
        Me.ClientSize = New Size(1400, 700)
        Me.AutoScaleMode = AutoScaleMode.Dpi
        Me.AutoScaleDimensions = New SizeF(96.0F, 96.0F)
        Me.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)

        ' Red Ink logo as window icon.
        Try
            Dim bmp As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
            Me.Icon = Icon.FromHandle(bmp.GetHicon())
        Catch
        End Try

        ' ── Header (search bar + user label) ─────────────────────────────
        headerPanel = New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 100
        }

        lblQuery = New Label() With {
            .Text = "Search query:", .AutoSize = True,
            .Location = New Point(12, 16)}
        headerPanel.Controls.Add(lblQuery)

        txtQuery = New TextBox() With {
            .Location = New Point(95, 12),
            .Size = New Size(380, 23),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left}
        headerPanel.Controls.Add(txtQuery)

        lblMax = New Label() With {
            .Text = "Max:", .AutoSize = True,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .Location = New Point(485, 16)}
        headerPanel.Controls.Add(lblMax)

        numMax = New NumericUpDown() With {
            .Minimum = 1, .Maximum = 500, .Value = 25,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .Location = New Point(520, 12), .Size = New Size(60, 23)}
        headerPanel.Controls.Add(numMax)

        btnSearch = New Button() With {
            .Text = "Search",
            .AutoSize = True,
            .Padding = StdButtonPadding,
            .Margin = StdButtonMargin,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .Location = New Point(590, 9)}
        headerPanel.Controls.Add(btnSearch)
        Me.AcceptButton = btnSearch

        btnAISearch = New Button() With {
            .Text = "AI Search",
            .AutoSize = True,
            .Padding = StdButtonPadding,
            .Margin = StdButtonMargin,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .Location = New Point(660, 9)}
        headerPanel.Controls.Add(btnAISearch)

        btnGetMore = New Button() With {
            .Text = "Get more",
            .AutoSize = True,
            .Padding = StdButtonPadding,
            .Margin = StdButtonMargin,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .Location = New Point(750, 9),
            .Enabled = False}
        headerPanel.Controls.Add(btnGetMore)

        btnSignIn = New Button() With {
            .Text = "Sign in",
            .AutoSize = True,
            .Padding = StdButtonPadding,
            .Margin = StdButtonMargin,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .Location = New Point(730, 9)}
        headerPanel.Controls.Add(btnSignIn)

        btnSignOut = New Button() With {
            .Text = "Sign out",
            .AutoSize = True,
            .Padding = StdButtonPadding,
            .Margin = New Padding(0),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .Location = New Point(800, 9)}
        headerPanel.Controls.Add(btnSignOut)

        ' "(not signed in)" — 20 px padding above and below.
        lblUser = New Label() With {
            .Text = "(not signed in)", .AutoSize = True, .ForeColor = Color.DimGray,
            .Location = New Point(12, 60)}
        headerPanel.Controls.Add(lblUser)

        lblAiStats = New Label() With {
            .Text = "",
            .AutoEllipsis = True,
            .ForeColor = Color.DimGray,
            .TextAlign = ContentAlignment.MiddleRight,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right,
            .Location = New Point(320, 60),
            .Size = New Size(headerPanel.ClientSize.Width - 332, 20)}
        headerPanel.Controls.Add(lblAiStats)

        ' ── Footer (progress, status, Open + Close) ──────────────────────
        footerPanel = New Panel() With {
            .Dock = DockStyle.Bottom,
            .Height = 80
        }

        pb = New ProgressBar() With {
            .Location = New Point(12, 10),
            .Size = New Size(footerPanel.ClientSize.Width - 260, 18),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right,
            .Style = ProgressBarStyle.Continuous}
        footerPanel.Controls.Add(pb)

        lblStatus = New Label() With {
            .Text = "Ready.",
            .AutoEllipsis = True,
            .Location = New Point(12, 36),
            .Size = New Size(footerPanel.ClientSize.Width - 260, 22),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right}
        footerPanel.Controls.Add(lblStatus)

        btnOpen = New Button() With {
            .Text = "Open (Preview)",
            .AutoSize = True,
            .Padding = StdButtonPadding,
            .Margin = New Padding(0),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .Location = New Point(footerPanel.ClientSize.Width - 160, 8),
            .Enabled = False}
        footerPanel.Controls.Add(btnOpen)

        btnClose = New Button() With {
            .Text = "Close",
            .AutoSize = True,
            .Padding = StdButtonPadding,
            .Margin = New Padding(0),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .Location = New Point(footerPanel.ClientSize.Width - 70, 8)}
        footerPanel.Controls.Add(btnClose)
        Me.CancelButton = btnClose

        ' ── Center (results + resizable AI summary) ──────────────────────
        splitMain = New SplitContainer() With {
            .Dock = DockStyle.Fill,
            .Orientation = Orientation.Horizontal,
            .SplitterWidth = 6,
            .Panel1MinSize = 140,
            .Panel2MinSize = 90,
            .FixedPanel = FixedPanel.None
        }
        splitMain.Panel1.Padding = New Padding(12, 6, 12, 2)
        splitMain.Panel2.Padding = New Padding(12, 4, 12, 6)

        lvResults = New ListView() With {
            .Dock = DockStyle.Fill,
            .View = View.Details,
            .FullRowSelect = True,
            .GridLines = True,
            .HideSelection = False,
            .MultiSelect = False}
        lvResults.Columns.Add("Ref", 55)
        lvResults.Columns.Add("Att", 40)
        lvResults.Columns.Add("Date", 120)
        lvResults.Columns.Add("Time", 75)
        lvResults.Columns.Add("From", 200)
        lvResults.Columns.Add("To", 220)
        lvResults.Columns.Add("Subject", 360)
        lvResults.Columns.Add("AI Summary", 440)
        splitMain.Panel1.Controls.Add(lvResults)

        ' Order: textbox first (Fill), label second (Top), so the label
        ' docks above the textbox and never overlaps it.
        txtSummary = New WebBrowser() With {
            .Dock = DockStyle.Fill,
            .AllowWebBrowserDrop = False,
            .IsWebBrowserContextMenuEnabled = True,
            .WebBrowserShortcutsEnabled = True,
            .ScriptErrorsSuppressed = True}
        splitMain.Panel2.Controls.Add(txtSummary)

        lblSummary = New Label() With {
            .Text = "AI Summary:",
            .Dock = DockStyle.Top,
            .Height = 28,
            .Padding = New Padding(0, 6, 0, 6),
            .TextAlign = ContentAlignment.MiddleLeft}
        splitMain.Panel2.Controls.Add(lblSummary)

        ' Add docked containers in the order Top → Bottom → Fill so each
        ' takes its slice and Fill consumes what's left.
        Me.Controls.Add(splitMain)
        Me.Controls.Add(footerPanel)
        Me.Controls.Add(headerPanel)

        ' Initial split: list ~60%, summary ~40% so the summary is tall
        ' enough to read several paragraphs without scrolling.
        Try
            Dim h As Integer = Math.Max(splitMain.Panel1MinSize + splitMain.Panel2MinSize + splitMain.SplitterWidth + 20,
                                        Me.ClientSize.Height - headerPanel.Height - footerPanel.Height)
            splitMain.SplitterDistance = CInt(h * 0.58)
        Catch
        End Try

        ' Pack the right-aligned button cluster + footer buttons after
        ' the controls have their preferred sizes.
        LayoutTopRowRight()
        LayoutFooterRight()
        ConfigureToolTips()
        RefreshAiSearchAvailability()
    End Sub

    ''' <summary>
    ''' Re-positions the top-row right-aligned controls (Search / AI Search
    ''' / Sign in / Sign out plus Max label + NumericUpDown) using each
    ''' button's actual PreferredSize, so AutoSize + Padding never causes
    ''' overlap.
    ''' </summary>
    Private Sub LayoutTopRowRight()
        If headerPanel Is Nothing OrElse btnSearch Is Nothing OrElse
           btnAISearch Is Nothing OrElse btnSignIn Is Nothing OrElse
           btnSignOut Is Nothing Then Return

        Const rightMargin As Integer = 12
        Const buttonGap As Integer = 6
        Const groupGap As Integer = 12
        Dim topY As Integer = 9
        Dim labelY As Integer = 16
        Dim numY As Integer = 12

        Dim x As Integer = headerPanel.ClientSize.Width - rightMargin

        For Each b As Button In New Button() {btnSignOut, btnSignIn, btnGetMore, btnAISearch, btnSearch}
            Dim w As Integer = b.PreferredSize.Width
            x -= w
            b.SetBounds(x, topY, w, b.Height)
            x -= buttonGap
        Next

        x -= (groupGap - buttonGap)
        x -= numMax.Width
        numMax.SetBounds(x, numY, numMax.Width, numMax.Height)

        Dim lblW As Integer = lblMax.PreferredSize.Width
        x -= 4
        x -= lblW
        lblMax.SetBounds(x, labelY, lblW, lblMax.Height)

        ' Place txtQuery just to the right of lblQuery and stretch it so
        ' it stops a fixed gap before lblMax, regardless of font/DPI.
        Const labelTextGap As Integer = 8
        Dim queryLeft As Integer = lblQuery.Left + lblQuery.PreferredSize.Width + labelTextGap
        Dim queryRight As Integer = x - groupGap
        Dim newWidth As Integer = Math.Max(120, queryRight - queryLeft)
        txtQuery.SetBounds(queryLeft, txtQuery.Top, newWidth, txtQuery.Height)

        If lblAiStats IsNot Nothing Then
            Dim statsLeft As Integer = Math.Max(320, txtQuery.Left)
            Dim statsWidth As Integer = Math.Max(120, headerPanel.ClientSize.Width - statsLeft - 12)
            lblAiStats.SetBounds(statsLeft, 60, statsWidth, lblAiStats.Height)
        End If

    End Sub

    Private Shared Sub SetAiDisplayRef(hit As M365SearchHit, aiRef As Integer)
        If hit Is Nothing Then Return

        Try
            If hit.RawJson Is Nothing Then
                hit.RawJson = New JObject()
            End If

            If aiRef > 0 Then
                hit.RawJson("aiRef") = aiRef
            Else
                Dim prop As JProperty = hit.RawJson.Property("aiRef")
                If prop IsNot Nothing Then prop.Remove()
            End If
        Catch
        End Try
    End Sub

    Private Shared Function GetAiDisplayRef(hit As M365SearchHit) As String
        If hit Is Nothing OrElse hit.RawJson Is Nothing Then Return ""

        Try
            Dim tok As JToken = hit.RawJson("aiRef")
            If tok Is Nothing Then Return ""

            Dim n As Integer = 0
            If Integer.TryParse(tok.ToString(), n) AndAlso n > 0 Then
                Return n.ToString()
            End If
        Catch
        End Try

        Return ""
    End Function

    Private Shared Function MeasureSingleLineLabelHeight(label As Label) As Integer
        If label Is Nothing Then Return 22

        Dim measured As Size = TextRenderer.MeasureText(
        "Ag",
        label.Font,
        New Size(Integer.MaxValue, Integer.MaxValue),
        TextFormatFlags.SingleLine Or TextFormatFlags.NoPadding)

        Return Math.Max(22, measured.Height + 6)
    End Function

    ''' <summary>Right-aligns Open + Close in the footer using PreferredSize.</summary>
    Private Sub LayoutFooterRight()
        If footerPanel Is Nothing OrElse btnOpen Is Nothing OrElse btnClose Is Nothing OrElse
       pb Is Nothing OrElse lblStatus Is Nothing Then Return

        Const leftMargin As Integer = 12
        Const rightMargin As Integer = 12
        Const topMargin As Integer = 10
        Const rowGap As Integer = 8
        Const buttonGap As Integer = 6
        Const bottomMargin As Integer = 10

        Dim buttonHeight As Integer = Math.Max(btnOpen.PreferredSize.Height, btnClose.PreferredSize.Height)
        Dim progressHeight As Integer = Math.Max(pb.Height, 18)
        Dim statusHeight As Integer = MeasureSingleLineLabelHeight(lblStatus)

        Dim x As Integer = footerPanel.ClientSize.Width - rightMargin

        Dim wClose As Integer = btnClose.PreferredSize.Width
        x -= wClose
        btnClose.SetBounds(x, 8, wClose, buttonHeight)
        x -= buttonGap

        Dim wOpen As Integer = btnOpen.PreferredSize.Width
        x -= wOpen
        btnOpen.SetBounds(x, 8, wOpen, buttonHeight)

        Dim leftAreaWidth As Integer = Math.Max(80, x - leftMargin - 8)
        pb.SetBounds(leftMargin, topMargin, leftAreaWidth, progressHeight)
        lblStatus.SetBounds(leftMargin, pb.Bottom + rowGap, leftAreaWidth, statusHeight)

        Dim neededHeight As Integer = Math.Max(
        lblStatus.Bottom + bottomMargin,
        Math.Max(btnOpen.Bottom, btnClose.Bottom) + bottomMargin)

        If footerPanel.Height <> neededHeight Then
            footerPanel.Height = neededHeight
        End If
    End Sub

    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        LayoutTopRowRight()
        LayoutFooterRight()
        ResizeResultsColumns()
    End Sub


    Private Sub ResetAiStats()
        UiPost(Sub()
                   If lblAiStats IsNot Nothing Then
                       lblAiStats.Text = ""
                   End If
               End Sub)
    End Sub

    Private Sub UpdateAiStats(retrievedCount As Integer,
                          Optional reviewedCount As Integer? = Nothing,
                          Optional keptCount As Integer? = Nothing,
                          Optional shownCount As Integer? = Nothing)
        UiPost(Sub()
                   If lblAiStats Is Nothing Then Return

                   Dim parts As New List(Of String) From {
                   $"Retrieved: {Math.Max(0, retrievedCount)}/{AISearch_MaxCandidateHits}"
               }

                   If reviewedCount.HasValue Then
                       parts.Add($"Reviewed: {Math.Max(0, reviewedCount.Value)}")
                   End If

                   If keptCount.HasValue Then
                       parts.Add($"Kept: {Math.Max(0, keptCount.Value)}")
                   End If

                   If shownCount.HasValue Then
                       parts.Add($"Shown: {Math.Max(0, shownCount.Value)}")
                   End If

                   lblAiStats.Text = String.Join("   •   ", parts)
               End Sub)
    End Sub

    ' ════════════════════════════════════════════════════════════════════
    '  UI-thread marshalling helpers
    ' ════════════════════════════════════════════════════════════════════
    ''' <summary>
    ''' Runs <paramref name="action"/> on the form's UI thread. Safe to call
    ''' from any thread; no-ops if the form is gone. Use this for every UI
    ''' touch that follows an Await — Outlook's main thread does not have a
    ''' WindowsFormsSynchronizationContext, so ConfigureAwait(True) may
    ''' resume on a thread-pool thread.
    ''' </summary>
    Private Sub UiPost(action As Action)
        If action Is Nothing Then Return
        If Not IsFormUsable() Then Return
        Try
            If Me.InvokeRequired Then
                Me.BeginInvoke(action)
            Else
                action()
            End If
        Catch
            ' Form vanished between the check and the invoke — ignore.
        End Try
    End Sub

    ''' <summary>Synchronous variant of <see cref="UiPost"/>.</summary>
    Private Sub UiInvoke(action As Action)
        If action Is Nothing Then Return
        If Not IsFormUsable() Then Return
        Try
            If Me.InvokeRequired Then
                Me.Invoke(action)
            Else
                action()
            End If
        Catch
        End Try
    End Sub


    ' ════════════════════════════════════════════════════════════════════
    '  Lifecycle
    ' ════════════════════════════════════════════════════════════════════
    Protected Overrides Async Sub OnShown(e As EventArgs)
        MyBase.OnShown(e)
        Try
            Dim user = Await M365Service.GetSignedInUserAsync(_context).ConfigureAwait(False)
            UiPost(Sub()
                       lblUser.Text = If(String.IsNullOrEmpty(user),
                                         "(not signed in)",
                                         "Signed in as: " & user)
                   End Sub)
        Catch
        End Try
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        Try : _cts?.Cancel() : Catch : End Try
        Try : _summaryCts?.Cancel() : Catch : End Try
        Try : StopAiHeartbeat() : Catch : End Try
        MyBase.OnFormClosing(e)
    End Sub

    ' ════════════════════════════════════════════════════════════════════
    '  Sign in / out
    ' ════════════════════════════════════════════════════════════════════


    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Try : _cts?.Cancel() : Catch : End Try
        Me.Close()
    End Sub


    Private Async Sub btnSignIn_Click(sender As Object, e As EventArgs) Handles btnSignIn.Click
        SetBusy(True, "Signing in…")
        Try
            Await M365Service.SignInAsync(_context).ConfigureAwait(False)
            Dim user = Await M365Service.GetSignedInUserAsync(_context).ConfigureAwait(False)
            UiPost(Sub()
                       lblUser.Text = "Signed in as: " & user
                       lblStatus.Text = "Signed in."
                   End Sub)
        Catch ex As Exception When IsCancellation(ex)
            UiPost(Sub() lblStatus.Text = "Sign-in cancelled.")
        Catch ex As Exception
            Dim msg As String = BuildExceptionDetails(ex)
            Debug.WriteLine("[M365] Sign-in failed:")
            Debug.WriteLine(msg)
            ShowErrorWithClipboard(
                "Microsoft 365",
                "Sign-in failed.",
                msg,
                "Sign-in failed.")
        Finally
            UiPost(Sub() SetBusy(False))
        End Try
    End Sub

    Private Shared Function BuildExceptionDetails(ex As Exception) As String
        If ex Is Nothing Then Return ""

        Dim sb As New System.Text.StringBuilder()
        Dim current As Exception = ex
        Dim level As Integer = 0

        While current IsNot Nothing
            If level = 0 Then
                sb.AppendLine(current.GetType().FullName)
            Else
                sb.AppendLine()
                sb.AppendLine("Inner exception " & level.ToString() & ":")
                sb.AppendLine(current.GetType().FullName)
            End If

            If Not String.IsNullOrWhiteSpace(current.Message) Then
                sb.AppendLine(current.Message)
            End If

            If Not String.IsNullOrWhiteSpace(current.StackTrace) Then
                sb.AppendLine()
                sb.AppendLine(current.StackTrace)
            End If

            current = current.InnerException
            level += 1
        End While

        Return sb.ToString().Trim()
    End Function

    Private Sub ShowErrorWithClipboard(caption As String,
                                       summary As String,
                                       details As String,
                                       statusText As String)
        UiPost(Sub()
                   Dim copied As Boolean = TryCopyTextToClipboard(details)
                   Dim msg As String = summary
                   If copied Then
                       msg &= vbCrLf & vbCrLf & "Technical details were copied to the clipboard."
                   ElseIf Not String.IsNullOrWhiteSpace(details) Then
                       msg &= vbCrLf & vbCrLf & "Technical details could not be copied to the clipboard."
                   End If

                   SharedMethods.ShowCustomMessageBox(msg, caption)
                   lblStatus.Text = statusText
               End Sub)
    End Sub

    Private Shared Function TryCopyTextToClipboard(text As String) As Boolean
        If String.IsNullOrWhiteSpace(text) Then Return False
        Try
            Clipboard.SetText(text, TextDataFormat.UnicodeText)
            Return True
        Catch
            Return False
        End Try
    End Function

    Private Async Function EnrichMailHitsAsync(hits As IList(Of M365SearchHit),
                                               ct As CancellationToken) As Task
        If hits Is Nothing OrElse hits.Count = 0 Then Return

        Dim total As Integer = hits.Count
        Dim ok As Integer = 0
        Dim failed As Integer = 0
        Dim skipped As Integer = 0
        Dim firstError As String = ""

        For i As Integer = 0 To hits.Count - 1
            ct.ThrowIfCancellationRequested()

            Dim hit = hits(i)
            If hit Is Nothing Then
                skipped += 1
                Debug.WriteLine($"[M365Search.Enrich] #{i + 1}: SKIPPED (hit is Nothing)")
                Continue For
            End If

            Dim id As String = GetHitMessageId(hit)
            Dim msg As M365Message = Nothing

            If String.IsNullOrWhiteSpace(id) Then
                Dim imid As String = TryGetInternetMessageId(hit)
                If Not String.IsNullOrWhiteSpace(imid) Then
                    Try
                        msg = Await M365Service.GetMessageByInternetMessageIdAsync(
                            _context,
                            imid,
                            M365MessageFields.Recipients,
                            ct).ConfigureAwait(False)

                        If msg IsNot Nothing Then
                            ApplyPreviewMessageToHit(hit, msg)
                            id = GetHitMessageId(hit)
                        End If
                    Catch ex As OperationCanceledException
                        Throw
                    Catch ex As Exception
                        Debug.WriteLine($"[M365Search.Enrich] #{i + 1} GetMessageByInternetMessageIdAsync EX: {ex.Message}")
                    End Try
                End If
            End If

            If String.IsNullOrWhiteSpace(id) Then
                If String.IsNullOrWhiteSpace(hit.AdditionalText) AndAlso
                   Not hit.LastModifiedUtc.HasValue AndAlso
                   String.IsNullOrWhiteSpace(hit.Author) AndAlso
                   String.IsNullOrWhiteSpace(hit.Title) Then

                    skipped += 1
                    Debug.WriteLine($"[M365Search.Enrich] #{i + 1}: SKIPPED (no id and no usable metadata)")
                Else
                    ok += 1
                    Debug.WriteLine($"[M365Search.Enrich] #{i + 1}: kept metadata-only hit")
                End If

                Continue For
            End If

            Dim needsFetch As Boolean =
                msg Is Nothing OrElse
                String.IsNullOrWhiteSpace(hit.AdditionalText) OrElse
                String.IsNullOrWhiteSpace(hit.Author) OrElse
                String.IsNullOrWhiteSpace(hit.Title) OrElse
                String.IsNullOrWhiteSpace(hit.WebUrl) OrElse
                Not hit.LastModifiedUtc.HasValue

            If needsFetch Then
                UiPost(Sub() lblStatus.Text = $"Loading recipients {i + 1}/{total}…")

                Try
                    msg = Await M365Service.GetMessageAsync(
                        _context,
                        id,
                        M365MessageFields.Recipients,
                        ct).ConfigureAwait(False)
                Catch ex As OperationCanceledException
                    Throw
                Catch ex As Exception
                    failed += 1
                    If String.IsNullOrEmpty(firstError) Then firstError = ex.Message
                    Debug.WriteLine($"[M365Search.Enrich] #{i + 1} EX: {ex.GetType().Name}: {ex.Message}")
                End Try
            End If

            If msg IsNot Nothing Then
                ApplyPreviewMessageToHit(hit, msg)
                If String.IsNullOrWhiteSpace(hit.AdditionalText) AndAlso msg.RawJson IsNot Nothing Then
                    hit.AdditionalText = ExtractToFromJson(msg.RawJson, "enrich.raw")
                End If
            End If

            If String.IsNullOrWhiteSpace(hit.AdditionalText) AndAlso
               Not hit.LastModifiedUtc.HasValue AndAlso
               String.IsNullOrWhiteSpace(hit.Author) AndAlso
               String.IsNullOrWhiteSpace(hit.Title) Then

                skipped += 1
                Debug.WriteLine($"[M365Search.Enrich] #{i + 1}: SKIPPED (message unresolved and no usable metadata)")
                Continue For
            End If

            ok += 1
        Next

        Debug.WriteLine($"[M365Search.Enrich] done: total={total} ok={ok} failed={failed} skipped={skipped} firstError='{firstError}'")
        UiPost(Sub() lblStatus.Text = $"Enriched {ok}/{total} mail(s)." &
                                      If(skipped > 0, $" {skipped} had no id.", "") &
                                      If(failed > 0, $" {failed} failed. First error: {firstError}", ""))
    End Function

    Private Shared Function BuildRecipientList(recipients As IEnumerable(Of String)) As String
        If recipients Is Nothing Then Return ""

        Dim items = recipients.
            Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
            Select(Function(s) s.Trim()).
            ToList()

        If items.Count = 0 Then Return ""
        Return String.Join("; ", items)
    End Function

    Private Async Sub btnSignOut_Click(sender As Object, e As EventArgs) Handles btnSignOut.Click
        SetBusy(True, "Signing out…")
        Try
            Await M365Service.SignOutAsync(_context).ConfigureAwait(False)
            UiPost(Sub()
                       lblUser.Text = "(not signed in)"
                       lblStatus.Text = "Signed out."
                   End Sub)
        Catch ex As Exception When IsCancellation(ex)
            UiPost(Sub() lblStatus.Text = "Sign-out cancelled.")
        Catch ex As Exception
            Dim msg As String = ex.Message
            UiPost(Sub() lblStatus.Text = "Sign-out failed: " & msg)
        Finally
            UiPost(Sub() SetBusy(False))
        End Try
    End Sub


    ''' <summary>True if the form is alive and safe to touch from a continuation.</summary>
    Private Function IsFormUsable() As Boolean
        Return Not Me.IsDisposed AndAlso Not Me.Disposing AndAlso Me.IsHandleCreated
    End Function

    ''' <summary>
    ''' Returns True for any flavour of "user cancelled" — covers
    ''' OperationCanceledException, MSAL's authentication_canceled, and
    ''' WebView2/WAM cancel paths that surface as plain Exceptions.
    ''' </summary>
    Private Shared Function IsCancellation(ex As Exception) As Boolean
        If ex Is Nothing Then Return False
        If TypeOf ex Is OperationCanceledException Then Return True

        Dim t As Type = ex.GetType()
        Dim typeName As String = If(t?.FullName, "")
        If typeName.IndexOf("Msal", StringComparison.OrdinalIgnoreCase) >= 0 Then
            Try
                Dim codeProp = t.GetProperty("ErrorCode")
                If codeProp IsNot Nothing Then
                    Dim code As String = TryCast(codeProp.GetValue(ex), String)
                    If Not String.IsNullOrEmpty(code) AndAlso
                       (code.IndexOf("cancel", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                        code.Equals("authentication_canceled", StringComparison.OrdinalIgnoreCase)) Then
                        Return True
                    End If
                End If
            Catch
            End Try
        End If

        Dim msg As String = If(ex.Message, "")
        Return msg.IndexOf("cancel", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               msg.IndexOf("user_cancel", StringComparison.OrdinalIgnoreCase) >= 0
    End Function

    ' ════════════════════════════════════════════════════════════════════
    '  Search
    ' ════════════════════════════════════════════════════════════════════
    Private Async Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Dim q = txtQuery.Text.Trim()
        If String.IsNullOrEmpty(q) Then
            SharedMethods.ShowCustomMessageBox("Enter a search term first.", "Microsoft 365")
            Return
        End If

        _cts?.Cancel()
        _cts = New CancellationTokenSource()

        SetBusy(True, "Searching…")
        lvResults.Items.Clear()
        _hits.Clear()
        btnOpen.Enabled = False
        _hasPinnedSummary = False
        SetSummaryMarkdown("")
        ResetAiStats()
        ClearAiMessageCache()

        Try
            Dim opts As New M365SearchOptions() With {
                .MaxPerSource = CInt(numMax.Value),
                .Parallel = True}

            Dim res = Await M365Service.SearchAsync(
                _context,
                q,
                M365SearchSources.Mail,
                opts,
                Me,
                _cts.Token).ConfigureAwait(False)

            Await EnrichMailHitsAsync(res.Hits, _cts.Token).ConfigureAwait(False)

            Dim user = Await M365Service.GetSignedInUserAsync(_context).ConfigureAwait(False)

            UiPost(Sub()
                       _hits = res.Hits
                       PopulateResults(_hits)
                       StartRowSummaries()
                       If Not String.IsNullOrEmpty(user) Then
                           lblUser.Text = "Signed in as: " & user
                       End If
                       lblStatus.Text = $"{_hits.Count} hit(s)." &
                           If(res.HasErrors, " Errors: " & String.Join("; ", res.ErrorsBySource.Values), "")
                   End Sub)
        Catch ex As OperationCanceledException
            UiPost(Sub() lblStatus.Text = "Cancelled.")
        Catch ex As Exception When IsCancellation(ex)
            UiPost(Sub() lblStatus.Text = "Cancelled.")
        Catch ex As Exception
            Dim msg As String = BuildExceptionDetails(ex)
            Debug.WriteLine("[M365] Search failed:")
            Debug.WriteLine(msg)
            ShowErrorWithClipboard(
                "Microsoft 365",
                "Microsoft 365 search failed.",
                msg,
                "Search failed.")
        Finally
            UiPost(Sub() SetBusy(False))
        End Try
    End Sub


    Private Async Sub btnAISearch_Click(sender As Object, e As EventArgs) Handles btnAISearch.Click

        RefreshAiSearchAvailability()

        If Not _aiSearchAvailable Then
            SharedMethods.ShowCustomMessageBox(_aiSearchUnavailableReason, "AI Mail Search")
            Return
        End If

        Dim userPrompt As String = ""
        Dim lastPrompt As String = ""
        Dim lastPromptInstruct As String = ""

        Try
            lastPrompt = If(My.Settings.LastSearchPrompt, "")
        Catch
            lastPrompt = ""
        End Try

        If Not String.IsNullOrWhiteSpace(lastPrompt) Then
            lastPromptInstruct = " Press Ctrl+P to reinsert your last prompt."
        End If

        Try
            userPrompt = SharedMethods.ShowCustomInputBox(
            "Describe the e-mails you are looking for. The AI will search " &
            "your Microsoft 365 mailbox, gather candidate messages, review their full text, and " &
            "show the relevant ones." & lastPromptInstruct,
            "AI Mail Search",
            SimpleInput:=False,
            Context:=_context,
            CtrlP:=lastPrompt)
        Catch
        End Try
        If String.IsNullOrWhiteSpace(userPrompt) OrElse userPrompt = "ESC" Then Return

        userPrompt = userPrompt.Trim()

        Try
            My.Settings.LastSearchPrompt = userPrompt
            My.Settings.Save()
        Catch
        End Try

        Await RunAiSearchPipelineAsync(userPrompt, appendMore:=False).ConfigureAwait(False)
    End Sub

    Private Async Sub btnGetMore_Click(sender As Object, e As EventArgs) Handles btnGetMore.Click
        If String.IsNullOrWhiteSpace(_aiLastUserPrompt) Then
            SharedMethods.ShowCustomMessageBox("Run an AI search first.", "AI Mail Search")
            Return
        End If

        If _aiRecordedSearches Is Nothing OrElse _aiRecordedSearches.Count = 0 Then
            SharedMethods.ShowCustomMessageBox("No recorded phase-1 searches are available for paging. Run AI Search again.", "AI Mail Search")
            Return
        End If

        _cts?.Cancel()
        _cts = New CancellationTokenSource()

        SetBusy(True, "Getting more candidates…")
        StartAiHeartbeat("Getting more candidates")

        Try
            Dim oldCount As Integer = _aiCandidateHits.Count
            Dim moreHits As List(Of M365SearchHit) =
                Await GetMoreCandidatesFromRecordedQueriesAsync(_cts.Token).ConfigureAwait(False)

            _aiCandidateHits = MergeCandidateHits(_aiCandidateHits, moreHits, AISearch_MaxCandidateHits)

            If _aiCandidateHits.Count = oldCount Then
                UiPost(Sub()
                           UpdateAiStats(_aiCandidateHits.Count)
                           UpdateGetMoreEnabled()
                           lblStatus.Text = "No additional candidates were found from the recorded phase-1 searches."
                       End Sub)
                Return
            End If

            Await ReviewAndDisplayCandidatePoolAsync(_aiLastUserPrompt, _cts.Token).ConfigureAwait(False)

        Catch ex As OperationCanceledException
            UiPost(Sub() lblStatus.Text = "Get more cancelled.")
        Catch ex As Exception When IsCancellation(ex)
            UiPost(Sub() lblStatus.Text = "Get more cancelled.")
        Catch ex As Exception
            Dim msg As String = BuildExceptionDetails(ex)
            Debug.WriteLine("[M365] Get more failed:" & vbCrLf & msg)
            ShowErrorWithClipboard(
                "AI Mail Search",
                "Getting more candidates failed.",
                msg,
                "Get more failed.")
        Finally
            UiPost(Sub() StopAiHeartbeat())
            UiPost(Sub() SetBusy(False))
        End Try
    End Sub

    Private Async Function RunAiSearchPipelineAsync(userPrompt As String,
                                                    appendMore As Boolean) As Task
        If String.IsNullOrWhiteSpace(userPrompt) Then Return

        _cts?.Cancel()
        _cts = New CancellationTokenSource()

        Dim aiMaxPerCall As Integer = GetAiHarvestMaxPerCall()

        If Not appendMore Then
            lvResults.Items.Clear()
            _hits.Clear()
            _aiCandidateHits.Clear()
            _aiRecordedSearches.Clear()
            btnOpen.Enabled = False
            _hasPinnedSummary = False
            SetSummaryMarkdown("")
            ResetAiStats()
            ClearAiMessageCache()
        End If

        SetBusy(True, "AI is searching your mailbox…")
        StartAiHeartbeat("AI is searching your mailbox")

        Try
            Dim harvest = Await ExecuteAiCandidateHarvestAsync(userPrompt, appendMore, aiMaxPerCall, _cts.Token).ConfigureAwait(False)
            If Not IsFormUsable() Then Return

            Debug.WriteLine("[M365Search] AI candidate raw response:" & vbCrLf & If(harvest.Item1, "(null)"))

            _aiRecordedSearches = MergeSearchInvocations(_aiRecordedSearches, CaptureAiSearchInvocations(harvest.Item2))

            Dim candidateRefs As List(Of Integer) = ExtractEmailRefs(harvest.Item1)
            Dim toolHits As List(Of M365SearchHit) = BuildHitsFromSearchToolResponses(harvest.Item2)

            Dim harvestedHits As New List(Of M365SearchHit)()
            If candidateRefs.Count > 0 Then
                harvestedHits = MatchAiReturnedRefsToToolHits(candidateRefs, toolHits)
                Debug.WriteLine("[AISearch] Matched candidate refs to tool hits: " & harvestedHits.Count.ToString())
            Else
                harvestedHits = toolHits
                Debug.WriteLine("[AISearch] No candidate refs returned; using raw search hits.")
            End If

            _aiCandidateHits = MergeCandidateHits(_aiCandidateHits, harvestedHits, AISearch_MaxCandidateHits)
            _aiLastUserPrompt = userPrompt

            UpdateAiStats(_aiCandidateHits.Count)
            Await ReviewAndDisplayCandidatePoolAsync(userPrompt, _cts.Token).ConfigureAwait(False)

        Catch ex As OperationCanceledException
            UiPost(Sub() lblStatus.Text = "AI search cancelled.")
        Catch ex As Exception When IsCancellation(ex)
            UiPost(Sub() lblStatus.Text = "AI search cancelled.")
        Catch ex As Exception
            Dim msg As String = BuildExceptionDetails(ex)
            Debug.WriteLine("[M365] AI search failed:" & vbCrLf & msg)
            ShowErrorWithClipboard(
                "AI Mail Search",
                "AI mail search failed.",
                msg,
                "AI search failed.")
        Finally
            UiPost(Sub() StopAiHeartbeat())
            UiPost(Sub() SetBusy(False))
        End Try
    End Function

    Private Async Function ExecuteAiCandidateHarvestAsync(userPrompt As String,
                                                          appendMore As Boolean,
                                                          aiMaxPerCall As Integer,
                                                          ct As CancellationToken) As Task(Of Tuple(Of String, List(Of ThisAddIn.ToolResponse)))

        Dim addIn = Globals.ThisAddIn
        Dim chatScope As IDisposable = Nothing
        Dim previousOtherPrompt As String = addIn.OtherPrompt

        Dim altPath As String = _context.INI_AlternateModelPath
        Dim haveAltPath As Boolean = Not String.IsNullOrWhiteSpace(altPath)

        Dim backupConfig As ModelConfig = Nothing
        Dim backupOriginalLoaded As Boolean = False
        Dim toolModelApplied As Boolean = False
        Dim aiSearchModel As ModelConfig = Nothing

        Dim previousMaxIters As Integer = addIn.INI_ToolingMaximumIterations
        Dim itersOverridden As Boolean = False

        Try
            addIn.OtherPrompt = userPrompt
            Globals.ThisAddIn.OtherPrompt = userPrompt

            Dim m365Tools = SharedLibrary.SharedLibrary.M365ToolService.GetTools(_context)
            Dim searchTool = m365Tools.FirstOrDefault(
        Function(t) String.Equals(t?.ToolName,
                                  SharedLibrary.SharedLibrary.M365ToolService.SearchToolName,
                                  StringComparison.OrdinalIgnoreCase))
            If searchTool Is Nothing Then
                Throw New InvalidOperationException("The Microsoft 365 search tool is not available.")
            End If

            Dim selectedTools As New List(Of ModelConfig) From {searchTool}

            chatScope = addIn.EnterChatAgentScope()

            If Not TryGetAiSearchDefaultModel(aiSearchModel) Then
                Throw New InvalidOperationException(_aiSearchUnavailableReason)
            End If

            backupConfig = SharedMethods.GetCurrentConfig(_context)
            backupOriginalLoaded = SharedMethods.originalConfigLoaded

            SharedMethods.originalConfig = backupConfig
            SharedMethods.originalConfigLoaded = True
            SharedMethods.ApplyModelConfig(_context, aiSearchModel)
            toolModelApplied = True

            addIn.INI_ToolingMaximumIterations = AISearch_MaxIterations
            itersOverridden = True

            Dim effectiveMaxPerCall As Integer = Math.Max(1, Math.Min(aiMaxPerCall, AISearch_MaxCandidateHits))

            Dim retryPolicy As String =
                vbCrLf & vbCrLf &
                "RETRY POLICY (candidate gathering):" & vbCrLf &
                $"- Use m365_search only. Gather a broad pool of distinct candidate mail hits for later full-text review, up to {AISearch_MaxCandidateHits} total candidates." & vbCrLf &
                $"- In EVERY m365_search call, set max_per_source to {effectiveMaxPerCall}." & vbCrLf &
                "- Start broad with objective anchors only. Do NOT use semantic review words such as problem, issue, risk, complaint or their German variants as initial KQL terms." & vbCrLf &
                "- If the request is about problems/issues with a person or matter, first gather mails involving that person or matter broadly, then let later phases decide relevance from full text." & vbCrLf &
                "- Use meaningfully different queries across calls. Do not just repeat the same query." & vbCrLf &
                "- Return candidate hit numbers in <candidate_refs>, not ids." & vbCrLf &
                "- If there is an ALREADY KNOWN CANDIDATES block, search for additional candidates beyond those already known. If old mails reappear in search results, do not return them again unless no other candidate exists."

            Dim knownCandidatesBlock As String = BuildKnownCandidatesBlock(_aiCandidateHits)

            Dim resolvedSysPrompt As String = addIn.InterpolateAtRuntime(If(_context.SP_AIMailSearch1, ""))
            resolvedSysPrompt &= retryPolicy

            If appendMore AndAlso Not String.IsNullOrWhiteSpace(knownCandidatesBlock) Then
                resolvedSysPrompt &= vbCrLf & vbCrLf & knownCandidatesBlock
            End If

            Dim finalText As String = Await addIn.ExecuteToolingLoop(
                sysCommand:=resolvedSysPrompt,
                userText:=userPrompt,
                selectedTools:=selectedTools,
                useSecondAPI:=True,
                otherPrompt:=userPrompt,
                hideSplash:=True,
                cancellationToken:=ct,
                finalResponseContract:=SharedLibrary.Agents.ToolingFinalResponseContract.RawCallerText).ConfigureAwait(False)

            Dim toolResponses As List(Of ThisAddIn.ToolResponse) = addIn.GetLastCompletedToolResponsesSnapshot()
            Return Tuple.Create(finalText, toolResponses)

        Finally
            Try
                If itersOverridden Then addIn.INI_ToolingMaximumIterations = previousMaxIters
            Catch
            End Try
            Try
                If toolModelApplied Then
                    If SharedMethods.originalConfigLoaded Then
                        SharedMethods.RestoreDefaults(_context, SharedMethods.originalConfig)
                    End If
                    SharedMethods.originalConfigLoaded = backupOriginalLoaded
                    If backupConfig IsNot Nothing Then
                        SharedMethods.ApplyModelConfig(_context, backupConfig)
                    End If
                End If
            Catch
            End Try
            Try : chatScope?.Dispose() : Catch : End Try
            Try : addIn.OtherPrompt = previousOtherPrompt : Catch : End Try
        End Try
    End Function

    Private Shared Function GetAttachmentIndicator(hit As M365SearchHit) As String
        If hit Is Nothing Then Return ""

        Try
            Dim resource As JObject = TryCast(hit.RawJson?("resource"), JObject)
            If resource Is Nothing Then Return ""

            Dim tok As JToken = resource("hasAttachments")
            If tok Is Nothing OrElse tok.Type = JTokenType.Null Then Return ""

            If tok.Type = JTokenType.Boolean AndAlso tok.Value(Of Boolean)() Then
                Return "📎"
            End If

            Dim parsed As Boolean = False
            If Boolean.TryParse(tok.ToString(), parsed) AndAlso parsed Then
                Return "📎"
            End If
        Catch
        End Try

        Return ""
    End Function

    Private Function BuildKnownCandidatesBlock(knownHits As IEnumerable(Of M365SearchHit)) As String
        If knownHits Is Nothing Then Return ""

        Dim list As List(Of M365SearchHit) = knownHits.Where(Function(h) h IsNot Nothing).Take(40).ToList()
        If list.Count = 0 Then Return ""

        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine("ALREADY KNOWN CANDIDATES:")
        sb.AppendLine("Search for ADDITIONAL candidates beyond these already-known mails.")
        sb.AppendLine("Do not return the same mail again if new candidates can be found.")
        For Each hit In list
            Dim dtText As String = ""
            Dim dt = GetMailDate(hit)
            If dt.HasValue Then dtText = dt.Value.ToLocalTime().ToString("yyyy-MM-dd HH:mm")
            sb.AppendLine("- " &
                          If(String.IsNullOrWhiteSpace(dtText), "(no date)", dtText) &
                          " | From: " & If(hit.Author, "") &
                          " | Subject: " & If(hit.Title, ""))
        Next

        Return sb.ToString().Trim()
    End Function

    Private Function MergeCandidateHits(existingHits As IEnumerable(Of M365SearchHit),
                                       newHits As IEnumerable(Of M365SearchHit),
                                       maxCount As Integer) As List(Of M365SearchHit)
        Dim merged As New List(Of M365SearchHit)()
        Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        Dim appendHit =
            Sub(hit As M365SearchHit)
                If hit Is Nothing Then Exit Sub
                Dim key As String = GetCandidateStableKey(hit)
                If String.IsNullOrWhiteSpace(key) Then Exit Sub
                If seen.Add(key) Then merged.Add(hit)
            End Sub

        If existingHits IsNot Nothing Then
            For Each hit In existingHits
                appendHit(hit)
                If merged.Count >= maxCount Then Return merged
            Next
        End If

        If newHits IsNot Nothing Then
            For Each hit In newHits
                appendHit(hit)
                If merged.Count >= maxCount Then Return merged
            Next
        End If

        Return merged
    End Function

    Private Function GetCandidateStableKey(hit As M365SearchHit) As String
        If hit Is Nothing Then Return ""

        Dim graphId As String = NormalizeGraphIdForCompare(GetHitMessageId(hit))
        If Not String.IsNullOrWhiteSpace(graphId) Then Return "id:" & graphId

        Dim imid As String = NormalizeInternetMessageId(TryGetInternetMessageId(hit))
        If Not String.IsNullOrWhiteSpace(imid) Then Return "imid:" & imid

        Dim webUrl As String = If(hit.WebUrl, "").Trim()
        If Not String.IsNullOrWhiteSpace(webUrl) Then Return "url:" & webUrl

        Return "fallback:" &
               If(hit.Author, "") & "|" &
               If(hit.Title, "") & "|" &
               If(GetMailDate(hit).HasValue, GetMailDate(hit).Value.ToString("o"), "")
    End Function

    Private Async Function FetchResolvedCandidateMailsAsync(candidateHits As IList(Of M365SearchHit),
                                                            ct As CancellationToken) As Task(Of List(Of AiResolvedMail))
        Dim result As New List(Of AiResolvedMail)()
        If candidateHits Is Nothing OrElse candidateHits.Count = 0 Then Return result

        Dim missingIds As New List(Of String)()

        For Each hit In candidateHits
            If hit Is Nothing Then Continue For
            If TryGetCachedMessage(hit) Is Nothing Then
                Dim id As String = GetHitMessageId(hit)
                If Not String.IsNullOrWhiteSpace(id) Then
                    missingIds.Add(id)
                End If
            End If
        Next

        If missingIds.Count > 0 Then
            Dim fetched As List(Of M365Message) =
                Await M365Service.GetMessagesBatchAsync(
                    _context,
                    missingIds.Distinct(StringComparer.OrdinalIgnoreCase),
                    M365MessageFields.Body Or M365MessageFields.Recipients,
                    ct).ConfigureAwait(False)

            For Each msg In fetched
                CacheMessage(msg)
            Next
        End If

        For i As Integer = 0 To candidateHits.Count - 1
            ct.ThrowIfCancellationRequested()

            Dim hit As M365SearchHit = candidateHits(i)
            Dim msg As M365Message = Nothing

            Dim currentIndex As Integer = i + 1
            UiPost(Sub() lblStatus.Text = $"Downloading full text {currentIndex}/{candidateHits.Count}…")

            If hit IsNot Nothing Then
                msg = TryGetCachedMessage(hit)
            End If

            If msg Is Nothing AndAlso hit IsNot Nothing Then
                Dim imid As String = NormalizeInternetMessageId(TryGetInternetMessageId(hit))
                If Not String.IsNullOrWhiteSpace(imid) Then
                    Try
                        msg = Await M365Service.GetMessageByInternetMessageIdAsync(
                            _context,
                            imid,
                            M365MessageFields.Body Or M365MessageFields.Recipients,
                            ct).ConfigureAwait(False)

                        CacheMessage(msg)
                    Catch ex As Exception
                        Debug.WriteLine("[AISearch] Fallback GetMessageByInternetMessageIdAsync failed: " & ex.Message)
                    End Try
                End If
            End If

            If msg IsNot Nothing Then
                ApplyPreviewMessageToHit(hit, msg)
            End If

            result.Add(New AiResolvedMail() With {
                .GlobalRef = i + 1,
                .Hit = hit,
                .Message = msg
            })
        Next

        Return result
    End Function

    Private Async Function ReviewCandidatesInBatchesAsync(userPrompt As String,
                                                          resolvedCandidates As IList(Of AiResolvedMail),
                                                          ct As CancellationToken) As Task(Of List(Of Integer))
        Dim result As New List(Of Integer)()
        If resolvedCandidates Is Nothing OrElse resolvedCandidates.Count = 0 Then Return result

        Dim addIn = Globals.ThisAddIn
        Dim seen As New HashSet(Of Integer)()

        For startIdx As Integer = 0 To resolvedCandidates.Count - 1 Step AISearch_ReviewBatchSize
            ct.ThrowIfCancellationRequested()

            Dim batch As List(Of AiResolvedMail) =
                resolvedCandidates.Skip(startIdx).Take(AISearch_ReviewBatchSize).ToList()

            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine("<user_request>" & userPrompt & "</user_request>")
            For Each item In batch
                sb.AppendLine(BuildAiAnalysisMailBlock(item, AISearch_BodyCap_PerCandidate))
            Next

            Dim currentEnd As Integer = Math.Min(startIdx + AISearch_ReviewBatchSize, resolvedCandidates.Count)
            UiPost(Sub() lblStatus.Text = $"Reviewing full text batch {startIdx + 1}-{currentEnd} of {resolvedCandidates.Count}…")

            Dim raw As String =
                Await addIn.LLM(_context.SP_AIMailSearch2, sb.ToString(), "", "", 0, False, True).ConfigureAwait(False)

            Dim parsedRefs As List(Of Integer) = ParseJsonIntArray(raw)

            Dim reviewedSoFar As Integer = Math.Min(currentEnd, resolvedCandidates.Count)
            UiPost(Sub() lblStatus.Text = $"Reviewed {reviewedSoFar}/{resolvedCandidates.Count} candidate mails…")

            Dim batchAllowed As New HashSet(Of Integer)(batch.Select(Function(x) x.GlobalRef))

            For Each r In parsedRefs
                If batchAllowed.Contains(r) AndAlso seen.Add(r) Then
                    result.Add(r)
                End If
            Next
        Next

        Return result
    End Function

    Private Async Function BuildFinalGridSelectionAsync(userPrompt As String,
                                                        shortlistedCandidates As IList(Of AiResolvedMail),
                                                        ct As CancellationToken) As Task(Of AiFinalSelection)

        Dim finalSelection As New AiFinalSelection()

        If shortlistedCandidates Is Nothing OrElse shortlistedCandidates.Count = 0 Then
            finalSelection.Summary = "No matching e-mails were found."
            Return finalSelection
        End If

        UiPost(Sub() lblStatus.Text = $"Building final summary from {shortlistedCandidates.Count} reviewed mail(s)…")

        Dim addIn = Globals.ThisAddIn
        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine("<user_request>" & userPrompt & "</user_request>")
        For Each item In shortlistedCandidates
            sb.AppendLine(BuildAiAnalysisMailBlock(item, AISearch_BodyCap_FinalSummary))
        Next

        Dim raw As String =
            Await addIn.LLM(_context.SP_AIMailSearch3, sb.ToString(), "", "", 0, False, True).ConfigureAwait(False)

        Dim summary As String = ExtractTaggedBlock(raw, "summary").Trim()
        Dim orderedRefs As List(Of Integer) = ExtractEmailRefs(raw)
        Dim allowed As New HashSet(Of Integer)(shortlistedCandidates.Select(Function(x) x.GlobalRef))
        Dim seen As New HashSet(Of Integer)()

        For Each r In orderedRefs
            If allowed.Contains(r) AndAlso seen.Add(r) Then
                finalSelection.OrderedGlobalRefs.Add(r)
            End If
        Next

        If finalSelection.OrderedGlobalRefs.Count = 0 Then
            For Each item In shortlistedCandidates
                If seen.Add(item.GlobalRef) Then finalSelection.OrderedGlobalRefs.Add(item.GlobalRef)
            Next
        End If

        finalSelection.Summary = If(String.IsNullOrWhiteSpace(summary),
                                    "No matching e-mails were found.",
                                    summary)

        Return finalSelection
    End Function

    Private Function BuildAiAnalysisMailBlock(item As AiResolvedMail,
                                              bodyCap As Integer) As String
        Dim hit As M365SearchHit = item?.Hit
        Dim msg As M365Message = item?.Message

        Dim subjectText As String = If(If(msg?.Subject, ""), If(hit?.Title, ""))
        Dim fromText As String = ""
        If msg IsNot Nothing Then
            fromText = If(String.IsNullOrWhiteSpace(msg.From), msg.FromAddress, msg.From)
        End If
        If String.IsNullOrWhiteSpace(fromText) Then fromText = If(If(hit?.Author, ""), "")

        Dim toText As String = ""
        If msg IsNot Nothing Then
            toText = BuildRecipientList(msg.To_)
        End If
        If String.IsNullOrWhiteSpace(toText) Then
            toText = If(If(hit?.AdditionalText, ""), "")
        End If

        Dim dtIsoText As String = ""
        Dim dtAnchorText As String = ""
        Dim dt = If(msg?.SentUtc, If(msg?.ReceivedUtc, GetMailDate(hit)))
        If dt.HasValue Then
            Dim dtLocal = dt.Value.ToLocalTime()
            dtIsoText = dtLocal.ToString("yyyy-MM-dd HH:mm", Globalization.CultureInfo.InvariantCulture)
            dtAnchorText = dtLocal.ToString("dd MMM yyyy HH:mm", Globalization.CultureInfo.InvariantCulture)
        End If

        Dim bodyText As String = ""
        If msg IsNot Nothing Then
            bodyText = GetPreviewBodyText(msg)
        End If
        If String.IsNullOrWhiteSpace(bodyText) Then
            bodyText = If(If(hit?.Summary, ""), "")
        End If
        If bodyText.Length > bodyCap Then
            bodyText = bodyText.Substring(0, bodyCap)
        End If

        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine($"<EMAIL ref=""{item.GlobalRef}"">")
        sb.AppendLine("DateISO: " & dtIsoText)
        sb.AppendLine("DateAnchor: " & dtAnchorText)
        sb.AppendLine("From: " & fromText)
        sb.AppendLine("To: " & toText)
        sb.AppendLine("Subject: " & subjectText)
        sb.AppendLine("Body:")
        sb.AppendLine(bodyText)
        sb.AppendLine("</EMAIL>")
        Return sb.ToString()
    End Function

    Private Shared Function ParseJsonIntArray(rawJson As String) As List(Of Integer)
        Dim result As New List(Of Integer)()
        If String.IsNullOrWhiteSpace(rawJson) Then Return result

        Dim text As String = rawJson.Trim()
        If text.StartsWith("```", StringComparison.Ordinal) Then
            Dim startArr As Integer = text.IndexOf("["c)
            Dim endArr As Integer = text.LastIndexOf("]"c)
            If startArr >= 0 AndAlso endArr > startArr Then
                text = text.Substring(startArr, endArr - startArr + 1)
            End If
        End If

        Dim arr As JArray = Nothing
        Try
            arr = JArray.Parse(text)
        Catch
            Return result
        End Try

        Dim seen As New HashSet(Of Integer)()
        For Each tok As JToken In arr
            Dim n As Integer = 0
            If tok.Type = JTokenType.Integer Then
                n = tok.Value(Of Integer)()
            ElseIf tok.Type = JTokenType.Object Then
                Dim obj As JObject = CType(tok, JObject)
                If obj("ref") IsNot Nothing Then
                    Integer.TryParse(obj("ref").ToString(), n)
                ElseIf obj("id") IsNot Nothing Then
                    Integer.TryParse(obj("id").ToString(), n)
                End If
            Else
                Integer.TryParse(tok.ToString(), n)
            End If

            If n > 0 AndAlso seen.Add(n) Then
                result.Add(n)
            End If
        Next

        Return result
    End Function


    Private Shared Function ExtractTaggedBlock(text As String, tag As String) As String
        If String.IsNullOrEmpty(text) Then Return ""
        Dim pattern As String = "<" & tag & ">(?<v>[\s\S]*?)</" & tag & ">"
        Dim m = System.Text.RegularExpressions.Regex.Match(
            text, pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        If m.Success Then Return m.Groups("v").Value
        Return ""
    End Function

    Private Shared Function ExtractEmailIds(text As String) As List(Of String)
        Dim block = ExtractTaggedBlock(text, "email_ids")
        Dim list As New List(Of String)()
        If String.IsNullOrWhiteSpace(block) Then Return list

        Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim seps = New Char() {ControlChars.Lf, ControlChars.Cr, ","c, ";"c}

        For Each raw In block.Split(seps, StringSplitOptions.RemoveEmptyEntries)
            Dim id = raw.Trim().
                         Trim(""""c).
                         Trim("'"c).
                         Trim("`"c).
                         Trim().
                         TrimStart("-"c, "*"c, " "c, "•"c).
                         Trim().
                         Trim("<"c, ">"c).
                         Trim()

            If id.Length > 0 AndAlso seen.Add(id) Then
                list.Add(id)
            End If
        Next

        Return list
    End Function

    Private Shared Function ExtractEmailRefs(text As String) As List(Of Integer)
        Dim block As String = ExtractTaggedBlock(text, "candidate_refs")
        If String.IsNullOrWhiteSpace(block) Then
            block = ExtractTaggedBlock(text, "email_refs")
        End If
        If String.IsNullOrWhiteSpace(block) Then
            block = ExtractTaggedBlock(text, "selected_refs")
        End If
        If String.IsNullOrWhiteSpace(block) Then
            block = ExtractTaggedBlock(text, "email_hit_numbers")
        End If
        If String.IsNullOrWhiteSpace(block) Then
            block = ExtractTaggedBlock(text, "email_hits")
        End If

        Dim result As New List(Of Integer)()
        If String.IsNullOrWhiteSpace(block) Then Return result

        Dim seen As New HashSet(Of Integer)()
        Dim matches = System.Text.RegularExpressions.Regex.Matches(block, "\d+")

        For Each m As System.Text.RegularExpressions.Match In matches
            Dim n As Integer
            If Integer.TryParse(m.Value, n) AndAlso n > 0 AndAlso seen.Add(n) Then
                result.Add(n)
            End If
        Next

        Return result
    End Function


    Private Shared Function NormalizeInternetMessageId(value As String) As String
        Dim s As String = If(value, "").Trim()
        If String.IsNullOrWhiteSpace(s) Then Return ""
        s = s.Trim("<"c, ">"c).Trim()
        Return s
    End Function

    ''' <summary>Adapts a fetched M365 message to the existing M365SearchHit shape used by the listview.</summary>
    Private Shared Function MailToHit(m As M365Message) As M365SearchHit
        Dim h As New M365SearchHit()
        If m Is Nothing Then Return h

        h.Id = m.Id
        h.Title = m.Subject
        h.Author = If(String.IsNullOrWhiteSpace(m.From), m.FromAddress, m.From)
        h.WebUrl = m.WebLink
        h.AdditionalText = BuildRecipientList(m.To_)

        ' Always wrap raw JSON so GetMailTo / TryGetInternetMessageId find a "resource".
        If m.RawJson IsNot Nothing Then
            h.RawJson = New JObject(New JProperty("resource", m.RawJson))
        End If

        ' If parsed To_ was empty but raw JSON has toRecipients, fill from there.
        If String.IsNullOrWhiteSpace(h.AdditionalText) AndAlso m.RawJson IsNot Nothing Then
            h.AdditionalText = ExtractToFromJson(m.RawJson, "MailToHit.raw")
        End If

        If m.ReceivedUtc.HasValue Then
            h.LastModifiedUtc = m.ReceivedUtc.Value
        ElseIf m.SentUtc.HasValue Then
            h.LastModifiedUtc = m.SentUtc.Value
        End If

        Return h
    End Function

    Private Sub PopulateResults(hits As List(Of M365SearchHit))
        lvResults.BeginUpdate()
        Try
            lvResults.Items.Clear()

            Dim ordered = hits.OrderByDescending(Function(h) GetMailDate(h)).ToList()
            _hits = ordered

            For i = 0 To ordered.Count - 1
                Dim h = ordered(i)
                Dim dtUtc As DateTime? = GetMailDate(h)
                Dim datePart As String = ""
                Dim timePart As String = ""

                If dtUtc.HasValue Then
                    Dim dtLocal = dtUtc.Value.ToLocalTime()
                    datePart = dtLocal.ToString("yyyy-MM-dd")
                    timePart = dtLocal.ToString("HH:mm")
                End If

                Dim it As New ListViewItem(GetAiDisplayRef(h))
                it.SubItems.Add(GetAttachmentIndicator(h))
                it.SubItems.Add(datePart)
                it.SubItems.Add(timePart)
                it.SubItems.Add(If(h.Author, ""))
                it.SubItems.Add(GetMailTo(h))
                it.SubItems.Add(If(h.Title, "(no subject)"))
                it.SubItems.Add("")
                it.Tag = h
                lvResults.Items.Add(it)
            Next
        Finally
            lvResults.EndUpdate()
        End Try

        ResizeResultsColumns()
    End Sub


    ''' <summary>
    ''' Builds a compact "To" string from the raw Graph payload's
    ''' toRecipients[].emailAddress.{name|address}. Returns "" if absent.
    ''' </summary>
    Private Function GetMailTo(h As M365SearchHit) As String
        If h Is Nothing Then
            Debug.WriteLine("[M365Search.To] hit is Nothing")
            Return ""
        End If

        Dim hitId As String = If(h.Id, "(no id)")

        If Not String.IsNullOrWhiteSpace(h.AdditionalText) Then
            Debug.WriteLine($"[M365Search.To] {hitId} -> AdditionalText='{h.AdditionalText}'")
            Return h.AdditionalText.Trim()
        End If

        Debug.WriteLine($"[M365Search.To] {hitId} -> AdditionalText is empty/whitespace")

        Dim resourceObj As JObject = TryCast(h.RawJson?("resource"), JObject)
        Debug.WriteLine($"[M365Search.To] {hitId} -> resource exists: {resourceObj IsNot Nothing}")

        Dim flatResourceTo As String = If(resourceObj?("to")?.ToString(), "")
        If Not String.IsNullOrWhiteSpace(flatResourceTo) Then
            Debug.WriteLine($"[M365Search.To] {hitId} -> resource.to='{flatResourceTo}'")
            Return flatResourceTo.Trim()
        End If

        Dim fromResource As String = ExtractToFromJson(resourceObj, $"{hitId}/resource")
        If Not String.IsNullOrWhiteSpace(fromResource) Then Return fromResource

        Dim flatTopTo As String = If(h.RawJson?("to")?.ToString(), "")
        If Not String.IsNullOrWhiteSpace(flatTopTo) Then
            Debug.WriteLine($"[M365Search.To] {hitId} -> top.to='{flatTopTo}'")
            Return flatTopTo.Trim()
        End If

        Dim fromTop As String = ExtractToFromJson(h.RawJson, $"{hitId}/top")
        If Not String.IsNullOrWhiteSpace(fromTop) Then Return fromTop

        Try
            If h.RawJson IsNot Nothing Then
                Dim keys = String.Join(",", h.RawJson.Properties().Select(Function(p) p.Name))
                Dim json As String = h.RawJson.ToString(Formatting.None)
                If json.Length > 400 Then json = json.Substring(0, 400) & "…"
                Debug.WriteLine($"[M365Search.To] {hitId} -> RawJson keys=[{keys}] body={json}")
            Else
                Debug.WriteLine($"[M365Search.To] {hitId} -> RawJson is Nothing")
            End If
        Catch
        End Try

        Return ""
    End Function

    Private Shared Function ExtractToFromJson(o As JObject, where As String) As String
        If o Is Nothing Then
            Debug.WriteLine($"[M365Search.To] ExtractToFromJson@{where}: object is Nothing")
            Return ""
        End If
        Try
            Dim arr = TryCast(o("toRecipients"), JArray)
            If arr Is Nothing Then
                Debug.WriteLine($"[M365Search.To] ExtractToFromJson@{where}: no toRecipients")
                Return ""
            End If
            Debug.WriteLine($"[M365Search.To] ExtractToFromJson@{where}: toRecipients count={arr.Count}")

            Dim names As New List(Of String)()
            For Each tok In arr
                Dim ea = TryCast(tok("emailAddress"), JObject)
                If ea Is Nothing Then
                    Debug.WriteLine($"[M365Search.To] ExtractToFromJson@{where}: tok has no emailAddress: {tok.ToString(Formatting.None)}")
                    Continue For
                End If
                Dim nm As String = If(ea("name")?.ToString(), "")
                Dim ad As String = If(ea("address")?.ToString(), "")
                Dim disp As String = If(Not String.IsNullOrWhiteSpace(nm), nm, ad)
                If Not String.IsNullOrWhiteSpace(disp) Then names.Add(disp.Trim())
            Next
            Dim joined = String.Join("; ", names)
            Debug.WriteLine($"[M365Search.To] ExtractToFromJson@{where}: produced='{joined}'")
            Return joined
        Catch ex As Exception
            Debug.WriteLine($"[M365Search.To] ExtractToFromJson@{where}: EX {ex.Message}")
            Return ""
        End Try
    End Function

    ''' <summary>Prefers SentDateTime from the raw Graph payload, falls back to LastModifiedUtc.</summary>
    Private Function GetMailDate(h As M365SearchHit) As DateTime?
        Try
            Dim r = TryCast(h?.RawJson?("resource"), JObject)
            If r IsNot Nothing Then
                Dim sentTok = r("sentDateTime")
                If sentTok IsNot Nothing AndAlso sentTok.Type <> JTokenType.Null Then
                    Dim dt As DateTime
                    If DateTime.TryParse(sentTok.ToString(), Nothing,
                                         Globalization.DateTimeStyles.AdjustToUniversal Or Globalization.DateTimeStyles.AssumeUniversal,
                                         dt) Then
                        Return dt
                    End If
                End If
                Dim recvTok = r("receivedDateTime")
                If recvTok IsNot Nothing AndAlso recvTok.Type <> JTokenType.Null Then
                    Dim dt As DateTime
                    If DateTime.TryParse(recvTok.ToString(), Nothing,
                                         Globalization.DateTimeStyles.AdjustToUniversal Or Globalization.DateTimeStyles.AssumeUniversal,
                                         dt) Then
                        Return dt
                    End If
                End If
            End If
        Catch
        End Try
        Return h?.LastModifiedUtc
    End Function

    Private Sub lvResults_SelectedIndexChanged(sender As Object, e As EventArgs) _
        Handles lvResults.SelectedIndexChanged
        btnOpen.Enabled = lvResults.SelectedItems.Count > 0
        UpdateSummaryPaneFromSelection()
    End Sub

    Private Sub lvResults_DoubleClick(sender As Object, e As EventArgs) _
        Handles lvResults.DoubleClick
        OpenSelectedHit()
    End Sub

    Private Sub btnOpen_Click(sender As Object, e As EventArgs) Handles btnOpen.Click
        OpenSelectedHit()
    End Sub

    ' ════════════════════════════════════════════════════════════════════
    '  Open hit — Outlook classic first, OWA web link as fallback
    ' ════════════════════════════════════════════════════════════════════
    Private Async Sub OpenSelectedHit()
        If lvResults.SelectedItems.Count = 0 Then Return

        Dim selectedItem As ListViewItem = lvResults.SelectedItems(0)
        Dim hit As M365SearchHit = TryGetHitFromListItem(selectedItem)
        If hit Is Nothing Then Return

        Dim previewMessage As M365Message = Nothing

        SetBusy(True, "Loading e-mail text…")
        Try
            previewMessage = Await GetPreviewMessageAsync(hit).ConfigureAwait(False)
        Catch ex As Exception When IsCancellation(ex)
            UiPost(Sub() lblStatus.Text = "Open cancelled.")
            Return
        Catch ex As Exception
            Dim msg As String = BuildExceptionDetails(ex)
            Debug.WriteLine("[M365Search] Preview download failed:")
            Debug.WriteLine(msg)
            ShowErrorWithClipboard(
                "Microsoft 365",
                "The selected e-mail could not be downloaded for preview.",
                msg,
                "Preview failed.")
            Return
        Finally
            UiInvoke(Sub() SetBusy(False))
        End Try

        Dim previewChoice As PreviewDialogChoice =
            Await ShowHtmlPreviewDialogAsync(hit, previewMessage).ConfigureAwait(False)

        If previewChoice <> PreviewDialogChoice.OpenInOutlook Then
            UiPost(Sub() lblStatus.Text = "Preview closed.")
            Return
        End If

        UiInvoke(Sub() SetBusy(True, "Opening message…"))
        Try
            Await OpenHitUsingExistingPipelineAsync(hit, previewMessage).ConfigureAwait(False)
        Catch ex As Exception When IsCancellation(ex)
            UiPost(Sub() lblStatus.Text = "Open cancelled.")
        Catch ex As Exception
            Dim msg As String = BuildExceptionDetails(ex)
            Debug.WriteLine("[M365Search] Open failed:")
            Debug.WriteLine(msg)
            ShowErrorWithClipboard(
                "Microsoft 365",
                "The selected message could not be opened.",
                msg,
                "Open failed.")
        Finally
            UiInvoke(Sub() SetBusy(False))
        End Try
    End Sub


    Private Enum PreviewDialogChoice
        ClosePreview = 0
        OpenInOutlook = 1
    End Enum

    Private Function ShowHtmlPreviewDialogAsync(hit As M365SearchHit,
                                            message As M365Message) As Task(Of PreviewDialogChoice)
        Dim html As String = BuildPreviewDialogHtml(hit, message)
        Dim tcs As New TaskCompletionSource(Of PreviewDialogChoice)()

        SharedMethods.ShowHTMLCustomMessageBox(
        html,
        "Microsoft 365",
        extraButtonText:="Open in Outlook",
        extraButtonAction:=Sub()
                               tcs.TrySetResult(PreviewDialogChoice.OpenInOutlook)
                           End Sub,
        CloseAfterExtra:=True,
        onClose:=Sub()
                     tcs.TrySetResult(PreviewDialogChoice.ClosePreview)
                 End Sub)

        Return tcs.Task
    End Function

    Private Function BuildPreviewDialogHtml(hit As M365SearchHit,
                                        message As M365Message) As String
        Dim subjectText As String = If(If(message?.Subject, ""), "").Trim()
        If String.IsNullOrWhiteSpace(subjectText) Then subjectText = If(If(hit?.Title, ""), "").Trim()
        If String.IsNullOrWhiteSpace(subjectText) Then subjectText = "(no subject)"

        Dim fromText As String = ""
        If message IsNot Nothing Then
            fromText = If(String.IsNullOrWhiteSpace(message.From), message.FromAddress, message.From)
        End If
        If String.IsNullOrWhiteSpace(fromText) Then fromText = If(If(hit?.Author, ""), "").Trim()
        If String.IsNullOrWhiteSpace(fromText) Then fromText = "(unknown sender)"

        Dim toText As String = ""
        If message IsNot Nothing Then
            toText = BuildRecipientList(message.To_)
        End If
        If String.IsNullOrWhiteSpace(toText) Then toText = GetMailTo(hit)

        Dim ccText As String = ""
        If message IsNot Nothing Then
            ccText = BuildRecipientList(message.Cc)
        End If

        Dim mailDate As DateTime? = Nothing
        If message IsNot Nothing Then
            If message.SentUtc.HasValue Then
                mailDate = message.SentUtc.Value
            ElseIf message.ReceivedUtc.HasValue Then
                mailDate = message.ReceivedUtc.Value
            End If
        End If
        If Not mailDate.HasValue Then mailDate = GetMailDate(hit)

        Dim dateText As String = ""
        If mailDate.HasValue Then
            dateText = mailDate.Value.ToLocalTime().ToString("yyyy-MM-dd HH:mm")
        End If

        Dim bodyHtml As String = BuildPreviewBodyHtml(message)

        Return "<html><head><meta http-equiv=""X-UA-Compatible"" content=""IE=edge"" />" &
        "<style>" &
        "body{font-family:'Segoe UI',sans-serif;font-size:9pt;color:#222;margin:0;padding:0;background:#fff;}" &
        ".wrap{padding:18px 20px 20px 20px;}" &
        ".banner{background:#fff8db;border:1px solid #ead27a;padding:10px 12px;margin-bottom:16px;}" &
        ".banner strong{display:block;margin-bottom:4px;}" &
        ".meta{border-collapse:collapse;width:100%;margin-bottom:18px;}" &
        ".meta th,.meta td{padding:6px 8px;border-bottom:1px solid #e5e5e5;vertical-align:top;text-align:left;}" &
        ".meta th{width:90px;color:#555;font-weight:600;white-space:nowrap;}" &
        ".mail-body{border:1px solid #ddd;padding:14px;background:#fff;}" &
        ".mail-plain{white-space:pre-wrap;word-wrap:break-word;font-family:Consolas,'Courier New',monospace;}" &
        "blockquote{margin:8px 0 8px 12px;padding-left:8px;border-left:3px solid #ddd;color:#555;}" &
        "table{border-collapse:collapse;}" &
        "th,td{border:1px solid #ddd;padding:4px 6px;}" &
        "</style></head><body>" &
        "<div class=""wrap"">" &
        "<div class=""banner"">" &
        "<strong>The full e-mail content below was downloaded from Microsoft 365.</strong>" &
        "You can select and copy from this preview. Opening it in Outlook may be slow and done online." &
        "</div>" &
        "<table class=""meta"">" &
        "<tr><th>Subject</th><td>" & HtmlEncode(subjectText) & "</td></tr>" &
        "<tr><th>From</th><td>" & HtmlEncode(fromText) & "</td></tr>" &
        If(String.IsNullOrWhiteSpace(toText), "", "<tr><th>To</th><td>" & HtmlEncode(toText) & "</td></tr>") &
        If(String.IsNullOrWhiteSpace(ccText), "", "<tr><th>Cc</th><td>" & HtmlEncode(ccText) & "</td></tr>") &
        If(String.IsNullOrWhiteSpace(dateText), "", "<tr><th>Date</th><td>" & HtmlEncode(dateText) & "</td></tr>") &
        "</table>" &
        "<div class=""mail-body"">" & bodyHtml & "</div>" &
        "</div></body></html>"
    End Function

    Private Shared Function BuildPreviewBodyHtml(message As M365Message) As String
        If message Is Nothing Then
            Return "<div class=""mail-plain"">(No message body text was returned.)</div>"
        End If

        Dim body As String = If(message.Body, "")
        If String.IsNullOrWhiteSpace(body) Then
            body = If(message.BodyPreview, "")
        End If

        If String.IsNullOrWhiteSpace(body) Then
            Return "<div class=""mail-plain"">(No message body text was returned.)</div>"
        End If

        If String.Equals(message.BodyContentType, "html", StringComparison.OrdinalIgnoreCase) Then
            Return body
        End If

        Return "<div class=""mail-plain"">" &
           HtmlEncode(NormalizePreviewText(body)).Replace(vbCrLf, "<br>").Replace(vbLf, "<br>") &
           "</div>"
    End Function

    Private Shared Function HtmlEncode(value As String) As String
        Return System.Net.WebUtility.HtmlEncode(If(value, ""))
    End Function


    Private Async Function GetPreviewMessageAsync(hit As M365SearchHit) As Task(Of M365Message)
        If hit Is Nothing Then Return Nothing

        Dim cached As M365Message = TryGetCachedMessage(hit)
        If cached IsNot Nothing Then
            ApplyPreviewMessageToHit(hit, cached)
            Return cached
        End If

        Dim resolvedId As String = GetHitMessageId(hit)
        Dim imidRaw As String = TryGetInternetMessageId(hit)
        Dim imid As String = NormalizeInternetMessageId(imidRaw)
        Dim webUrl As String = If(hit.WebUrl, "").Trim()
        Dim conversationId As String = TryGetConversationId(hit)

        If String.IsNullOrWhiteSpace(resolvedId) AndAlso Not String.IsNullOrWhiteSpace(webUrl) Then
            Dim fromLink As String = ExtractMessageIdFromWebLink(webUrl)
            If Not String.IsNullOrWhiteSpace(fromLink) Then
                resolvedId = fromLink
                hit.Id = fromLink
            End If
        End If

        Dim msg As M365Message = Nothing
        Dim lastEx As Exception = Nothing
        Dim attempts As New List(Of String)()

        Debug.WriteLine("[M365Search] Preview resolve start:")
        Debug.WriteLine("  hit.Id        = '" & If(hit.Id, "(null)") & "'")
        Debug.WriteLine("  resolvedId    = '" & If(resolvedId, "(null)") & "'")
        Debug.WriteLine("  imid          = '" & If(imid, "(null)") & "'")
        Debug.WriteLine("  hit.WebUrl    = '" & If(webUrl, "(null)") & "'")
        Debug.WriteLine("  conversationId= '" & If(conversationId, "(null)") & "'")

        If Not String.IsNullOrWhiteSpace(resolvedId) Then
            attempts.Add("GetMessageAsync(id='" & resolvedId & "')")
            Try
                msg = Await M365Service.GetMessageAsync(
                _context,
                resolvedId,
                M365MessageFields.Body Or M365MessageFields.Recipients).ConfigureAwait(False)
            Catch ex As OperationCanceledException
                Throw
            Catch ex As Exception
                lastEx = ex
                Debug.WriteLine("[M365Search] GetMessageAsync(id) failed: " & ex.Message)
            End Try
        End If

        If msg Is Nothing AndAlso Not String.IsNullOrWhiteSpace(imid) Then
            Dim variants = New String() {imid, "<" & imid & ">", imidRaw}
            For Each v In variants.Where(Function(s) Not String.IsNullOrWhiteSpace(s)).Distinct()
                attempts.Add("GetMessageByInternetMessageIdAsync('" & v & "')")
                Try
                    msg = Await M365Service.GetMessageByInternetMessageIdAsync(
                    _context,
                    v,
                    M365MessageFields.Body Or M365MessageFields.Recipients).ConfigureAwait(False)
                Catch ex As OperationCanceledException
                    Throw
                Catch ex As Exception
                    lastEx = ex
                    Debug.WriteLine("[M365Search] GetMessageByInternetMessageIdAsync('" & v & "') failed: " & ex.Message)
                End Try
                If msg IsNot Nothing Then Exit For
            Next
        End If

        If msg Is Nothing AndAlso Not String.IsNullOrWhiteSpace(conversationId) Then
            attempts.Add("GetMailThreadAsync(conversationId='" & conversationId & "')")
            Try
                msg = Await TryGetPreviewMessageFromConversationAsync(hit, conversationId).ConfigureAwait(False)
            Catch ex As OperationCanceledException
                Throw
            Catch ex As Exception
                lastEx = ex
                Debug.WriteLine("[M365Search] GetMailThreadAsync(conversationId) failed: " & ex.Message)
            End Try
        End If

        If msg Is Nothing Then
            attempts.Add("TryBuildPreviewMessageFromAiToolResponses")
            msg = TryBuildPreviewMessageFromAiToolResponses(hit, resolvedId)
        End If

        If msg Is Nothing Then
            Debug.WriteLine("[M365Search] All preview attempts failed:" & vbCrLf &
                        "  - " & String.Join(vbCrLf & "  - ", attempts))

            If lastEx IsNot Nothing Then
                Throw New InvalidOperationException(
                "The selected e-mail could not be resolved on Microsoft 365 (" &
                lastEx.Message & ").", lastEx)
            End If

            Throw New InvalidOperationException(
            "The selected e-mail could not be resolved to a Microsoft 365 message. " &
            "id='" & If(resolvedId, "") & "', internetMessageId='" & If(imid, "") & "'.")
        End If

        CacheMessage(msg)
        ApplyPreviewMessageToHit(hit, msg)
        Return msg
    End Function



    Private Shared Function TryGetConversationId(hit As M365SearchHit) As String
        Try
            Dim r = TryCast(hit?.RawJson?("resource"), JObject)
            If r Is Nothing Then Return ""
            Return If(r("conversationId")?.ToString(), "").Trim()
        Catch
            Return ""
        End Try
    End Function

    Private Async Function TryGetPreviewMessageFromConversationAsync(hit As M365SearchHit,
                                                                 conversationId As String) As Task(Of M365Message)
        If hit Is Nothing OrElse String.IsNullOrWhiteSpace(conversationId) Then Return Nothing

        Dim messages = Await M365Service.GetMailThreadAsync(
        _context,
        conversationId,
        New M365ThreadOptions() With {
            .MaxMessages = 100,
            .Ascending = False,
            .IncludeMailBody = True}).ConfigureAwait(False)

        If messages Is Nothing OrElse messages.Count = 0 Then Return Nothing

        Dim wantedImid As String = NormalizeInternetMessageId(TryGetInternetMessageId(hit))
        Dim wantedWebUrl As String = If(hit.WebUrl, "").Trim()

        If Not String.IsNullOrWhiteSpace(wantedImid) Then
            Dim exactByImid = messages.FirstOrDefault(
            Function(m) String.Equals(
                NormalizeInternetMessageId(If(m?.InternetMessageId, "")),
                wantedImid,
                StringComparison.OrdinalIgnoreCase))
            If exactByImid IsNot Nothing Then Return exactByImid
        End If

        If Not String.IsNullOrWhiteSpace(wantedWebUrl) Then
            Dim exactByWebUrl = messages.FirstOrDefault(
            Function(m) String.Equals(
                If(m?.WebLink, "").Trim(),
                wantedWebUrl,
                StringComparison.OrdinalIgnoreCase))
            If exactByWebUrl IsNot Nothing Then Return exactByWebUrl
        End If

        Return Nothing
    End Function


    ''' <summary>
    ''' Pulls the Graph message id out of an OWA webLink's ItemID query parameter.
    ''' Returns the Graph URL-safe id form used by /me/messages/{id}.
    ''' </summary>
    Private Shared Function ExtractMessageIdFromWebLink(webLink As String) As String
        If String.IsNullOrWhiteSpace(webLink) Then Return ""

        Try
            Dim m = System.Text.RegularExpressions.Regex.Match(
                webLink,
                "[?&]ItemID=(?<v>[^&]+)",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase)

            If Not m.Success Then Return ""

            Dim rawId As String = Uri.UnescapeDataString(m.Groups("v").Value).Trim()
            If String.IsNullOrWhiteSpace(rawId) Then Return ""

            Return ToGraphUrlSafeId(rawId)
        Catch
            Return ""
        End Try
    End Function


    Private Function TryBuildPreviewMessageFromAiToolResponses(hit As M365SearchHit,
                                                               preferredMessageId As String) As M365Message
        Dim responses As List(Of ThisAddIn.ToolResponse) = Nothing
        Try
            responses = Globals.ThisAddIn.GetLastCompletedToolResponsesSnapshot()
        Catch ex As Exception
            Debug.WriteLine("[M365Search] AI tool snapshot unavailable: " & ex.Message)
        End Try

        If responses Is Nothing OrElse responses.Count = 0 Then Return Nothing

        Dim preferredIdNorm As String = NormalizeGraphIdForCompare(preferredMessageId)
        Dim hitWebUrl As String = If(hit?.WebUrl, "").Trim()
        Dim hitImid As String = NormalizeInternetMessageId(TryGetInternetMessageId(hit))

        For Each r In responses
            If r Is Nothing OrElse Not r.Success Then Continue For
            If Not String.Equals(r.ToolName,
                                 SharedLibrary.SharedLibrary.M365ToolService.GetMailToolName,
                                 StringComparison.OrdinalIgnoreCase) Then
                Continue For
            End If

            Dim candidate As M365Message = ParsePreviewMessageFromGetMailResponse(r.Response)
            If candidate Is Nothing Then Continue For

            Dim candidateIdNorm As String = NormalizeGraphIdForCompare(candidate.Id)
            If Not String.IsNullOrWhiteSpace(preferredIdNorm) AndAlso
               String.Equals(candidateIdNorm, preferredIdNorm, StringComparison.OrdinalIgnoreCase) Then
                Return candidate
            End If

            If Not String.IsNullOrWhiteSpace(hitImid) AndAlso
               Not String.IsNullOrWhiteSpace(candidate.InternetMessageId) AndAlso
               String.Equals(NormalizeInternetMessageId(candidate.InternetMessageId),
                             hitImid,
                             StringComparison.OrdinalIgnoreCase) Then
                Return candidate
            End If

            If Not String.IsNullOrWhiteSpace(hitWebUrl) AndAlso
               Not String.IsNullOrWhiteSpace(candidate.WebLink) AndAlso
               String.Equals(candidate.WebLink.Trim(),
                             hitWebUrl,
                             StringComparison.OrdinalIgnoreCase) Then
                Return candidate
            End If
        Next

        Debug.WriteLine("[M365Search] AI tool preview fallback found no exact identifier match.")
        Return Nothing
    End Function

    Private Function ParsePreviewMessageFromGetMailResponse(responseText As String) As M365Message
        If String.IsNullOrWhiteSpace(responseText) Then Return Nothing

        Dim m = System.Text.RegularExpressions.Regex.Match(
        responseText,
        "<MAIL\s+id=""(?<id>[^""]*)""\s+title=""(?<title>[^""]*)""\s*>(?<body>[\s\S]*?)</MAIL>",
        System.Text.RegularExpressions.RegexOptions.IgnoreCase)

        If Not m.Success Then Return Nothing

        Dim messageId As String = ToGraphUrlSafeId(m.Groups("id").Value.Trim())
        Dim title As String = m.Groups("title").Value.Trim()
        Dim body As String = m.Groups("body").Value.Trim()

        Return New M365Message() With {
        .Id = messageId,
        .Subject = title,
        .Body = body,
        .BodyContentType = "text"
    }
    End Function


    Private Sub ApplyPreviewMessageToHit(hit As M365SearchHit, message As M365Message)
        If hit Is Nothing OrElse message Is Nothing Then Return

        Dim graphId As String = ToGraphUrlSafeId(If(message.Id, ""))
        If Not String.IsNullOrWhiteSpace(graphId) Then
            hit.Id = graphId
        End If

        If Not String.IsNullOrWhiteSpace(message.Subject) Then
            hit.Title = message.Subject
        End If

        Dim fromDisplay As String = If(String.IsNullOrWhiteSpace(message.From), message.FromAddress, message.From)
        If Not String.IsNullOrWhiteSpace(fromDisplay) Then
            hit.Author = fromDisplay
        End If

        If Not String.IsNullOrWhiteSpace(If(message.WebLink, "").Trim()) Then
            hit.WebUrl = message.WebLink.Trim()
        End If

        Dim toLine As String = BuildRecipientList(message.To_)
        If Not String.IsNullOrWhiteSpace(toLine) Then
            hit.AdditionalText = toLine
        End If

        Dim preferredDate As DateTime? = Nothing
        If message.SentUtc.HasValue Then
            preferredDate = message.SentUtc.Value
        ElseIf message.ReceivedUtc.HasValue Then
            preferredDate = message.ReceivedUtc.Value
        End If
        If preferredDate.HasValue Then
            hit.LastModifiedUtc = preferredDate.Value
        End If

        Try
            If message.RawJson IsNot Nothing Then
                Dim wrapped As JObject = If(hit.RawJson, New JObject())
                wrapped("hitId") = If(graphId, "")
                wrapped("resource") = message.RawJson.DeepClone()
                hit.RawJson = wrapped
            End If
        Catch
        End Try
    End Sub
    Private Function BuildPreviewDialogText(hit As M365SearchHit, message As M365Message) As String
        Dim sb As New System.Text.StringBuilder()
        Dim subjectText As String = If(If(message?.Subject, ""), "").Trim()
        If String.IsNullOrWhiteSpace(subjectText) Then subjectText = If(If(hit?.Title, ""), "").Trim()
        If String.IsNullOrWhiteSpace(subjectText) Then subjectText = "(no subject)"

        Dim fromText As String = ""
        If message IsNot Nothing Then
            fromText = If(String.IsNullOrWhiteSpace(message.From), message.FromAddress, message.From)
        End If
        If String.IsNullOrWhiteSpace(fromText) Then fromText = If(If(hit?.Author, ""), "").Trim()
        If String.IsNullOrWhiteSpace(fromText) Then fromText = "(unknown sender)"

        Dim toText As String = ""
        If message IsNot Nothing Then
            toText = BuildRecipientList(message.To_)
        End If
        If String.IsNullOrWhiteSpace(toText) Then toText = GetMailTo(hit)

        Dim ccText As String = ""
        If message IsNot Nothing Then
            ccText = BuildRecipientList(message.Cc)
        End If

        Dim mailDate As DateTime? = Nothing
        If message IsNot Nothing Then
            If message.SentUtc.HasValue Then
                mailDate = message.SentUtc.Value
            ElseIf message.ReceivedUtc.HasValue Then
                mailDate = message.ReceivedUtc.Value
            End If
        End If
        If Not mailDate.HasValue Then mailDate = GetMailDate(hit)

        Dim bodyText As String = GetPreviewBodyText(message)
        If String.IsNullOrWhiteSpace(bodyText) Then
            bodyText = "(No message body text was returned.)"
        End If

        sb.AppendLine("The full e-mail text below was downloaded from Microsoft 365.")
        sb.AppendLine("Opening it in Outlook may still be slow.")
        sb.AppendLine()
        sb.AppendLine("Choose ""Open in Outlook"" to continue, or ""Close"" to stay in the search window.")
        sb.AppendLine()
        sb.AppendLine(New String("-"c, 72))
        sb.AppendLine("Subject: " & subjectText)
        sb.AppendLine("From: " & fromText)
        If Not String.IsNullOrWhiteSpace(toText) Then sb.AppendLine("To: " & toText)
        If Not String.IsNullOrWhiteSpace(ccText) Then sb.AppendLine("Cc: " & ccText)
        If mailDate.HasValue Then sb.AppendLine("Date: " & mailDate.Value.ToLocalTime().ToString("yyyy-MM-dd HH:mm"))
        sb.AppendLine(New String("-"c, 72))
        sb.AppendLine()
        sb.AppendLine(bodyText)

        Return sb.ToString().Trim()
    End Function
    Private Shared Function GetPreviewBodyText(message As M365Message) As String
        If message Is Nothing Then Return ""
        Dim bodyText As String = If(message.Body, "")
        If String.IsNullOrWhiteSpace(bodyText) Then
            bodyText = If(message.BodyPreview, "")
        End If
        If String.IsNullOrWhiteSpace(bodyText) Then Return ""

        If String.Equals(message.BodyContentType, "html", StringComparison.OrdinalIgnoreCase) Then
            bodyText = ConvertHtmlToPlainText(bodyText)
        End If

        Return NormalizePreviewText(bodyText)
    End Function
    Private Shared Function ConvertHtmlToPlainText(html As String) As String
        If String.IsNullOrWhiteSpace(html) Then Return ""
        Dim text As String = html

        text = System.Text.RegularExpressions.Regex.Replace(text, "(?is)<\s*br\s*/?\s*>", vbCrLf)
        text = System.Text.RegularExpressions.Regex.Replace(text, "(?is)</\s*(p|div|tr|table|h[1-6])\s*>", vbCrLf & vbCrLf)
        text = System.Text.RegularExpressions.Regex.Replace(text, "(?is)<\s*li[^>]*>", "- ")
        text = System.Text.RegularExpressions.Regex.Replace(text, "(?is)</\s*li\s*>", vbCrLf)
        text = System.Text.RegularExpressions.Regex.Replace(text, "(?is)<[^>]+>", "")
        text = System.Net.WebUtility.HtmlDecode(text)
        text = text.Replace(ChrW(160), " "c)

        Return NormalizePreviewText(text)
    End Function
    Private Shared Function NormalizePreviewText(text As String) As String
        If String.IsNullOrWhiteSpace(text) Then Return ""
        Dim value As String = text.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
        value = value.Replace(vbTab, "    ")
        value = System.Text.RegularExpressions.Regex.Replace(value, "[ \t]+\n", vbLf)
        value = System.Text.RegularExpressions.Regex.Replace(value, "\n{3,}", vbLf & vbLf)

        Return value.Replace(vbLf, vbCrLf).Trim()
    End Function


    Private Async Function OpenHitUsingExistingPipelineAsync(hit As M365SearchHit, previewMessage As M365Message) As Task
        If hit Is Nothing Then Return

        Dim resolvedId As String = GetHitMessageId(hit)
        If String.IsNullOrWhiteSpace(resolvedId) AndAlso previewMessage IsNot Nothing Then
            resolvedId = ToGraphUrlSafeId(If(previewMessage.Id, ""))
            If Not String.IsNullOrWhiteSpace(resolvedId) Then hit.Id = resolvedId
        End If

        Dim imid As String = NormalizeInternetMessageId(TryGetInternetMessageId(hit))
        If String.IsNullOrWhiteSpace(imid) AndAlso previewMessage IsNot Nothing Then
            imid = NormalizeInternetMessageId(If(previewMessage.InternetMessageId, ""))
        End If

        Dim webUrl As String = If(hit.WebUrl, "").Trim()
        If String.IsNullOrWhiteSpace(webUrl) AndAlso previewMessage IsNot Nothing Then
            webUrl = If(previewMessage.WebLink, "").Trim()
            hit.WebUrl = webUrl
        End If

        Dim openedLocally As Boolean = False

        If Not String.IsNullOrWhiteSpace(imid) Then
            UiInvoke(Sub()
                         Try
                             Dim app = Globals.ThisAddIn.Application
                             If app Is Nothing Then Return

                             Dim mail = FindMailInDefaultStore(app, imid)
                             If mail IsNot Nothing Then
                                 mail.Display(False)
                                 openedLocally = True
                                 lblStatus.Text = "Opened in Outlook."
                             End If
                         Catch ex As Exception
                             Debug.WriteLine("[M365Search] Exact Outlook open failed: " & ex.Message)
                         End Try
                     End Sub)
        End If

        If openedLocally Then Return

        If (String.IsNullOrWhiteSpace(imid) OrElse String.IsNullOrWhiteSpace(webUrl)) AndAlso
           Not String.IsNullOrWhiteSpace(resolvedId) Then
            Try
                Dim full = Await M365Service.GetMessageAsync(
                    _context, resolvedId, M365MessageFields.Headers).ConfigureAwait(False)

                If full IsNot Nothing Then
                    If String.IsNullOrWhiteSpace(imid) Then
                        imid = NormalizeInternetMessageId(If(full.InternetMessageId, ""))
                    End If
                    If String.IsNullOrWhiteSpace(webUrl) Then
                        webUrl = If(full.WebLink, "").Trim()
                        hit.WebUrl = webUrl
                    End If
                End If
            Catch ex As Exception
                Debug.WriteLine("[M365Search] GetMessageAsync(headers) failed: " & ex.Message)
            End Try
        End If

        If Not openedLocally AndAlso Not String.IsNullOrWhiteSpace(imid) Then
            UiInvoke(Sub()
                         Try
                             Dim app = Globals.ThisAddIn.Application
                             If app Is Nothing Then Return

                             Dim mail = FindMailInDefaultStore(app, imid)
                             If mail IsNot Nothing Then
                                 mail.Display(False)
                                 openedLocally = True
                                 lblStatus.Text = "Opened in Outlook."
                             End If
                         Catch ex As Exception
                             Debug.WriteLine("[M365Search] Exact Outlook open after Graph failed: " & ex.Message)
                         End Try
                     End Sub)
        End If

        If openedLocally Then Return

        If Not String.IsNullOrWhiteSpace(webUrl) Then
            UiInvoke(Sub()
                         Try
                             Process.Start(New ProcessStartInfo(webUrl) With {.UseShellExecute = True})
                             lblStatus.Text = "Opened in Outlook on the web."
                         Catch ex As Exception
                             Dim details As String =
                                 "WebUrl=" & webUrl & vbCrLf & vbCrLf &
                                 BuildExceptionDetails(ex)
                             ShowErrorWithClipboard(
                                 "Microsoft 365",
                                 "The selected message could not be opened.",
                                 details,
                                 "Open failed.")
                         End Try
                     End Sub)
            Return
        End If

        Dim diag As String =
            "Hit.Id=" & If(hit.Id, "(null)") & vbCrLf &
            "ResolvedId=" & If(resolvedId, "(null)") & vbCrLf &
            "Hit.WebUrl=" & If(hit.WebUrl, "(null)") & vbCrLf &
            "InternetMessageId=" & If(imid, "(null)") & vbCrLf & vbCrLf &
            "RawJson:" & vbCrLf &
            If(hit.RawJson Is Nothing, "(null)", hit.RawJson.ToString())

        ShowErrorWithClipboard(
            "Microsoft 365",
            "The message could not be located in Outlook and no web link is available.",
            diag,
            "Open failed (no link).")
    End Function


    ''' <summary>Reads internetMessageId from the search hit's raw payload, if present.</summary>
    Private Function TryGetInternetMessageId(hit As M365SearchHit) As String
        Try
            Dim r = TryCast(hit?.RawJson?("resource"), JObject)
            If r Is Nothing Then Return ""
            Return If(r("internetMessageId")?.ToString(), "")
        Catch
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Bounded search restricted to the user's DEFAULT store (the active Outlook
    ''' classic mailbox). Tries the standard folders first (fast, indexed), then
    ''' walks the rest with a small depth cap. Never touches other stores
    ''' (archives, PSTs, shared mailboxes) so it cannot hang on those.
    ''' </summary>
    Private Function FindMailInDefaultStore(app As Outlook.Application, imid As String) As Outlook.MailItem
        If app Is Nothing OrElse String.IsNullOrWhiteSpace(imid) Then Return Nothing

        Dim core As String = imid.Trim()
        If core.StartsWith("<") Then core = core.TrimStart("<"c)
        If core.EndsWith(">") Then core = core.TrimEnd(">"c)
        Dim coreEsc As String = core.Replace("'", "''")

        ' Try several filters (with-brackets DASL, without-brackets DASL, Jet).
        Dim filters As String() = {
            "@SQL=""urn:schemas:mailheader:message-id"" = '<" & coreEsc & ">'",
            "@SQL=""urn:schemas:mailheader:message-id"" = '" & coreEsc & "'",
            "[InternetMessageID] = '<" & coreEsc & ">'",
            "[InternetMessageID] = '" & coreEsc & "'"
        }

        Dim session = app.Session
        If session Is Nothing Then Return Nothing

        Dim quickTargets = New Outlook.OlDefaultFolders() {
            Outlook.OlDefaultFolders.olFolderInbox,
            Outlook.OlDefaultFolders.olFolderSentMail,
            Outlook.OlDefaultFolders.olFolderDrafts,
            Outlook.OlDefaultFolders.olFolderDeletedItems,
            Outlook.OlDefaultFolders.olFolderJunk,
            Outlook.OlDefaultFolders.olFolderOutbox
        }

        For Each ft In quickTargets
            Try
                Dim f = TryCast(session.GetDefaultFolder(ft), Outlook.Folder)
                For Each filterText In filters
                    Dim m = TryFindMail(f, filterText)
                    If m IsNot Nothing Then Return m
                Next
            Catch ex As Exception
                Debug.WriteLine("[M365Search] Quick folder skipped: " & ex.Message)
            End Try
        Next

        Try
            Dim defaultStore = session.DefaultStore
            If defaultStore Is Nothing Then Return Nothing
            Dim root = TryCast(defaultStore.GetRootFolder(), Outlook.Folder)
            For Each filterText In filters
                Dim m = WalkFoldersBounded(root, filterText, currentDepth:=0, maxDepth:=4)
                If m IsNot Nothing Then Return m
            Next
        Catch ex As Exception
            Debug.WriteLine("[M365Search] Default store walk failed: " & ex.Message)
        End Try

        Return Nothing
    End Function

    Private Function TryFindMail(folder As Outlook.Folder, filter As String) As Outlook.MailItem
        If folder Is Nothing Then Return Nothing
        Try
            If folder.DefaultItemType <> Outlook.OlItemType.olMailItem Then Return Nothing
            Return TryCast(folder.Items.Find(filter), Outlook.MailItem)
        Catch
            Return Nothing
        End Try
    End Function

    Private Function WalkFoldersBounded(folder As Outlook.Folder,
                                        filter As String,
                                        currentDepth As Integer,
                                        maxDepth As Integer) As Outlook.MailItem
        If folder Is Nothing OrElse currentDepth > maxDepth Then Return Nothing

        Dim m = TryFindMail(folder, filter)
        If m IsNot Nothing Then Return m

        Try
            For Each sub_ As Outlook.Folder In folder.Folders
                Dim r = WalkFoldersBounded(sub_, filter, currentDepth + 1, maxDepth)
                If r IsNot Nothing Then Return r
            Next
        Catch
        End Try

        Return Nothing
    End Function


    ''' <summary>
    ''' Walks every Outlook store and folder looking for a mail item whose
    ''' InternetMessageID matches. Stops at the first match.
    ''' </summary>
    Private Function FindMailItemByInternetMessageId(app As Outlook.Application, imid As String) As Outlook.MailItem
        If app Is Nothing OrElse String.IsNullOrEmpty(imid) Then Return Nothing
        Dim filter = "[InternetMessageID] = '" & imid.Replace("'", "''") & "'"
        For Each store As Outlook.Store In app.Session.Stores
            Try
                Dim root As Outlook.Folder = CType(store.GetRootFolder(), Outlook.Folder)
                Dim found = SearchFolderRecursive(root, filter)
                If found IsNot Nothing Then Return found
            Catch ex As Exception
                Debug.WriteLine("[M365Search] Store skipped: " & ex.Message)
            End Try
        Next
        Return Nothing
    End Function

    Private Function SearchFolderRecursive(folder As Outlook.Folder, filter As String) As Outlook.MailItem
        If folder Is Nothing Then Return Nothing
        Try
            If folder.DefaultItemType = Outlook.OlItemType.olMailItem Then
                Dim found = TryCast(folder.Items.Find(filter), Outlook.MailItem)
                If found IsNot Nothing Then Return found
            End If
        Catch ex As Exception
            Debug.WriteLine("[M365Search] Folder skipped '" & folder.Name & "': " & ex.Message)
        End Try
        Try
            For Each sub_ As Outlook.Folder In folder.Folders
                Dim r = SearchFolderRecursive(sub_, filter)
                If r IsNot Nothing Then Return r
            Next
        Catch
        End Try
        Return Nothing
    End Function

    ' ════════════════════════════════════════════════════════════════════
    '  IProgress<M365SearchProgress>
    ' ════════════════════════════════════════════════════════════════════
    Public Sub Report(value As M365SearchProgress) Implements IProgress(Of M365SearchProgress).Report
        If InvokeRequired Then
            Try
                BeginInvoke(New Action(Sub() ApplyProgress(value)))
            Catch
            End Try
        Else
            ApplyProgress(value)
        End If
    End Sub

    Private Sub ApplyProgress(p As M365SearchProgress)
        If Not IsFormUsable() Then Return
        If p.Percent < 0 Then
            pb.Style = ProgressBarStyle.Marquee
        Else
            pb.Style = ProgressBarStyle.Continuous
            pb.Value = Math.Max(0, Math.Min(100, p.Percent))
        End If
        lblStatus.Text = $"{p.Stage}{If(p.Source = M365SearchSources.None, "", " — " & p.Source.ToString())}: {p.Message}  ({p.HitsSoFar} hit(s))"
    End Sub

    Private Sub SetBusy(busy As Boolean, Optional status As String = Nothing)
        btnSearch.Enabled = Not busy
        btnAISearch.Enabled = Not busy AndAlso _aiSearchAvailable
        btnSignIn.Enabled = Not busy
        btnSignOut.Enabled = Not busy
        If busy Then
            btnGetMore.Enabled = False
        End If
        txtQuery.Enabled = Not busy
        numMax.Enabled = Not busy
        Cursor = If(busy, Cursors.AppStarting, Cursors.Default)
        If status IsNot Nothing Then lblStatus.Text = status
        If Not busy Then
            pb.Style = ProgressBarStyle.Continuous
            pb.Value = 0
            UpdateGetMoreEnabled()
        End If
    End Sub

    Private Sub StartAiHeartbeat(prefix As String)
        StopAiHeartbeat()
        _aiHeartbeatPrefix = prefix
        _aiHeartbeatStart = DateTime.Now
        _aiHeartbeat = New System.Windows.Forms.Timer() With {.Interval = 1000}
        AddHandler _aiHeartbeat.Tick, AddressOf AiHeartbeat_Tick
        _aiHeartbeat.Start()
        pb.Style = ProgressBarStyle.Marquee
        pb.MarqueeAnimationSpeed = 30
    End Sub

    Private Sub StopAiHeartbeat()
        Try
            If _aiHeartbeat IsNot Nothing Then
                _aiHeartbeat.Stop()
                RemoveHandler _aiHeartbeat.Tick, AddressOf AiHeartbeat_Tick
                _aiHeartbeat.Dispose()
                _aiHeartbeat = Nothing
            End If
        Catch
        End Try
        Try
            pb.MarqueeAnimationSpeed = 0
            pb.Style = ProgressBarStyle.Continuous
            pb.Value = 0
        Catch
        End Try
    End Sub

    Private Sub AiHeartbeat_Tick(sender As Object, e As EventArgs)
        If Not IsFormUsable() Then Return
        Try
            Dim elapsed As TimeSpan = DateTime.Now - _aiHeartbeatStart
            lblStatus.Text = $"{_aiHeartbeatPrefix} ({CInt(elapsed.TotalSeconds)}s elapsed)…"
        Catch
        End Try
    End Sub

    Private _summaryCts As CancellationTokenSource

    Private Sub StartRowSummaries()

        Try : _summaryCts?.Cancel() : Catch : End Try
        _summaryCts = New CancellationTokenSource()
        Dim ct = _summaryCts.Token

        ' Snapshot the rows to summarize (row index -> hit).
        Dim snapshot As New List(Of Tuple(Of Integer, M365SearchHit))()
        For i As Integer = 0 To lvResults.Items.Count - 1
            Dim it = lvResults.Items(i)
            Dim hit As M365SearchHit = TryGetHitFromListItem(it)
            If hit Is Nothing Then Continue For
            snapshot.Add(Tuple.Create(i, hit))
        Next
        If snapshot.Count = 0 Then Return

        Task.Run(Async Function()
                     Try
                         Await GenerateRowSummariesAsync(snapshot, ct).ConfigureAwait(False)
                     Catch ex As OperationCanceledException
                     Catch ex As Exception
                         Debug.WriteLine("[M365Search] Summary worker failed: " & ex.Message)
                     End Try
                 End Function)
    End Sub

    Private Async Function GenerateRowSummariesAsync(rows As List(Of Tuple(Of Integer, M365SearchHit)),
                                                     ct As CancellationToken) As Task
        Dim addIn = Globals.ThisAddIn
        Dim batchSize As Integer = M365Search_SummaryBatchSize
        Dim total As Integer = rows.Count
        Dim done As Integer = 0
        Dim batchNo As Integer = 0
        Dim batchStart As Integer = 0

        While batchStart < total
            ct.ThrowIfCancellationRequested()
            batchNo += 1
            Dim batchEnd As Integer = Math.Min(batchStart + batchSize - 1, total - 1)
            Debug.WriteLine($"[M365Search.Sum] batch #{batchNo}: rows {batchStart + 1}..{batchEnd + 1}")

            Dim items As New List(Of Tuple(Of Integer, M365SearchHit, String))()
            For i As Integer = batchStart To batchEnd
                ct.ThrowIfCancellationRequested()
                Dim row = rows(i)
                Dim hit = row.Item2
                Dim body As String = ""
                Try
                    Dim cached As M365Message = TryGetCachedMessage(hit)
                    If cached IsNot Nothing Then
                        body = GetPreviewBodyText(cached)
                    Else
                        Dim mailId As String = GetHitMessageId(hit)
                        If hit IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(mailId) Then
                            body = Await M365Service.GetMessageBodyAsync(
                                _context, mailId, asPlainText:=True, ct:=ct).ConfigureAwait(False)
                        End If
                    End If
                    Debug.WriteLine($"[M365Search.Sum] body fetched id={If(hit?.Id, "?")} len={If(body, "").Length}")
                Catch ex As OperationCanceledException
                    Throw
                Catch ex As Exception
                    Debug.WriteLine($"[M365Search.Sum] body EX id={If(hit?.Id, "?")}: {ex.Message}")
                End Try
                If body IsNot Nothing AndAlso body.Length > M365Search_SummaryBodyCap Then
                    body = body.Substring(0, M365Search_SummaryBodyCap)
                End If
                items.Add(Tuple.Create(row.Item1, hit, If(body, "")))
            Next

            Dim sb As New System.Text.StringBuilder()
            For k As Integer = 0 To items.Count - 1
                Dim hit = items(k).Item2
                If hit Is Nothing Then Continue For
                Dim mailNum As Integer = k + 1
                sb.AppendLine($"<EMAIL id=""{mailNum}"">")
                sb.AppendLine($"Subject: {If(hit.Title, "")}")
                sb.AppendLine($"From: {If(hit.Author, "")}")
                If hit.LastModifiedUtc.HasValue Then
                    sb.AppendLine($"Date: {hit.LastModifiedUtc.Value.ToLocalTime():yyyy-MM-dd}")
                End If
                If Not String.IsNullOrWhiteSpace(items(k).Item3) Then
                    sb.AppendLine($"Body: {items(k).Item3}")
                End If
                sb.AppendLine("</EMAIL>")
            Next
            Dim userPromptText As String = sb.ToString()
            Debug.WriteLine($"[M365Search.Sum] LLM userPrompt length={userPromptText.Length}")

            Dim systemPrompt As String = ""
            Try
                systemPrompt = addIn.InterpolateAtRuntime(_context.SP_InboxBoard)
            Catch ex As Exception
                Debug.WriteLine($"[M365Search.Sum] InterpolateAtRuntime EX: {ex.Message}")
            End Try
            Debug.WriteLine($"[M365Search.Sum] system prompt len={If(systemPrompt, "").Length}")

            Dim raw As String = ""
            Try
                raw = Await addIn.LLM(systemPrompt, userPromptText, "", "", 0, False, True).ConfigureAwait(False)
                Debug.WriteLine($"[M365Search.Sum] LLM returned len={If(raw, "").Length} preview={If(raw, "").Substring(0, Math.Min(200, If(raw, "").Length))}")
            Catch ex As OperationCanceledException
                Throw
            Catch ex As Exception
                Debug.WriteLine($"[M365Search.Sum] LLM EX {ex.GetType().Name}: {ex.Message}")
            End Try

            ApplyRowSummaries(raw, items)

            done += items.Count
            UiPost(Sub() lblStatus.Text = $"AI summaries: {Math.Min(done, total)}/{total} done.")
            batchStart = batchEnd + 1
        End While
    End Function

    Private Sub ApplyRowSummaries(rawJson As String,
                                  items As List(Of Tuple(Of Integer, M365SearchHit, String)))
        If items Is Nothing OrElse items.Count = 0 Then
            Debug.WriteLine("[M365Search.Sum] ApplyRowSummaries: no items")
            Return
        End If
        If String.IsNullOrWhiteSpace(rawJson) Then
            Debug.WriteLine("[M365Search.Sum] ApplyRowSummaries: rawJson empty")
            Return
        End If
        Try
            Dim text As String = rawJson.Trim()
            If text.StartsWith("```") Then
                Dim s As Integer = text.IndexOf("[")
                Dim ee As Integer = text.LastIndexOf("]")
                If s >= 0 AndAlso ee > s Then text = text.Substring(s, ee - s + 1)
            End If

            Dim arr As JArray = Nothing
            Try
                arr = JArray.Parse(text)
            Catch ex As Exception
                Debug.WriteLine($"[M365Search.Sum] JArray.Parse EX: {ex.Message}; text preview={text.Substring(0, Math.Min(200, text.Length))}")
                Return
            End Try
            Debug.WriteLine($"[M365Search.Sum] parsed {arr.Count} entries")

            For Each item As JObject In arr
                Dim id As Integer = If(item("id") IsNot Nothing, CInt(item("id")), 0)
                Dim summary As String = If(item("summary") IsNot Nothing, CStr(item("summary")), "")
                Debug.WriteLine($"[M365Search.Sum] entry id={id} summary.len={If(summary, "").Length}")
                If id < 1 OrElse id > items.Count OrElse String.IsNullOrWhiteSpace(summary) Then Continue For

                Dim rowIdx As Integer = items(id - 1).Item1
                Dim cleaned As String = summary.Trim()
                UiPost(Sub()
                           If rowIdx < 0 OrElse rowIdx >= lvResults.Items.Count Then
                               Debug.WriteLine($"[M365Search.Sum] rowIdx {rowIdx} out of range")
                               Return
                           End If
                           Dim it = lvResults.Items(rowIdx)
                           Dim summaryColIdx As Integer = lvResults.Columns.Count - 1
                           While it.SubItems.Count <= summaryColIdx
                               it.SubItems.Add("")
                           End While
                           it.SubItems(summaryColIdx).Text = cleaned
                           If Not _hasPinnedSummary AndAlso lvResults.SelectedItems.Count > 0 AndAlso
                              lvResults.SelectedItems(0) Is it Then
                               SetSummaryMarkdown(cleaned)
                           End If
                           Debug.WriteLine($"[M365Search.Sum] row #{rowIdx} updated col {summaryColIdx}")
                       End Sub)
            Next
        Catch ex As Exception
            Debug.WriteLine($"[M365Search.Sum] ApplyRowSummaries EX: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Returns a usable message id for a search hit. Graph's /search/query
    ''' for mail puts the message id in the top-level "hitId", not always in
    ''' resource.id, so we may need to back-fill.
    ''' </summary>
    Private Shared Function GetHitMessageId(h As M365SearchHit) As String
        If h Is Nothing Then Return ""

        Dim id As String = If(h.Id, "").Trim()

        Try
            If String.IsNullOrWhiteSpace(id) AndAlso h.RawJson IsNot Nothing Then
                Dim hitIdTok = h.RawJson("hitId")
                If hitIdTok IsNot Nothing Then id = hitIdTok.ToString().Trim()

                If String.IsNullOrWhiteSpace(id) Then
                    Dim r = TryCast(h.RawJson("resource"), JObject)
                    If r IsNot Nothing Then
                        Dim rid = r("id")
                        If rid IsNot Nothing Then id = rid.ToString().Trim()
                    End If
                End If
            End If
        Catch
        End Try

        id = ToGraphUrlSafeId(id)

        If Not String.IsNullOrWhiteSpace(id) AndAlso String.IsNullOrWhiteSpace(h.Id) Then
            h.Id = id
        End If

        Return If(id, "")
    End Function

    Private Sub ResizeResultsColumns()
        If lvResults Is Nothing OrElse lvResults.Columns.Count = 0 Then Return

        Const verticalScrollbarAllowance As Integer = 24
        Const minSummaryWidth As Integer = 260

        Dim fixedWidth As Integer = 0
        For i As Integer = 0 To lvResults.Columns.Count - 2
            fixedWidth += lvResults.Columns(i).Width
        Next

        Dim available As Integer = lvResults.ClientSize.Width - fixedWidth - verticalScrollbarAllowance
        If available < minSummaryWidth Then available = minSummaryWidth

        lvResults.Columns(lvResults.Columns.Count - 1).Width = available
    End Sub

    Private Sub UpdateSummaryPaneFromSelection()
        If _hasPinnedSummary Then Return
        If txtSummary Is Nothing OrElse lvResults Is Nothing Then Return

        If lvResults.SelectedItems.Count = 0 Then
            SetSummaryMarkdown("")
            Return
        End If

        Dim summaryColIdx As Integer = lvResults.Columns.Count - 1
        Dim it = lvResults.SelectedItems(0)
        If it Is Nothing Then
            SetSummaryMarkdown("")
            Return
        End If

        If it.SubItems.Count <= summaryColIdx Then
            SetSummaryMarkdown("")
            Return
        End If

        SetSummaryMarkdown(If(it.SubItems(summaryColIdx).Text, ""))
    End Sub

    Private Function MarkdownToHtml(md As String) As String
        Try
            Dim pipeline As Markdig.MarkdownPipeline =
                New Markdig.MarkdownPipelineBuilder().
                    UseAdvancedExtensions().
                    UseSoftlineBreakAsHardlineBreak().
                    UsePipeTables().
                    UseGridTables().
                    UseListExtras().
                    UseFootnotes().
                    UseDefinitionLists().
                    UseAbbreviations().
                    UseAutoLinks().
                    UseTaskLists().
                    UseMathematics().
                    UseFigures().
                    UseGenericAttributes().
                    Build()

            Return Markdig.Markdown.ToHtml(If(md, ""), pipeline)
        Catch
            Return System.Net.WebUtility.HtmlEncode(If(md, "")).Replace(vbCrLf, "<br>").Replace(vbLf, "<br>")
        End Try
    End Function

    Private Function WrapSummaryHtml(bodyHtml As String) As String
        Return "<html><head><meta http-equiv=""X-UA-Compatible"" content=""IE=edge"" />" &
            "<style>" &
            "body{font-family:'Segoe UI',sans-serif;font-size:9pt;color:#222;margin:10px;line-height:1.35;}" &
            "p{margin:0 0 8px 0;}" &
            "ul,ol{margin-top:4px;margin-bottom:8px;}" &
            "code,pre{font-family:Consolas,'Courier New',monospace;}" &
            "pre{white-space:pre-wrap;word-wrap:break-word;background:#f6f6f6;padding:8px;border:1px solid #ddd;}" &
            "table{border-collapse:collapse;}" &
            "th,td{border:1px solid #ddd;padding:4px 6px;}" &
            "blockquote{margin:8px 0 8px 12px;padding-left:8px;border-left:3px solid #ddd;color:#555;}" &
            "</style></head><body>" &
            bodyHtml &
            "</body></html>"
    End Function

    Private Sub SetSummaryMarkdown(markdown As String)
        If txtSummary Is Nothing Then Return
        Dim html As String = MarkdownToHtml(If(markdown, ""))
        txtSummary.DocumentText = WrapSummaryHtml(html)
    End Sub

    Private Shared Function TryParseUtc(value As String) As DateTime?
        If String.IsNullOrWhiteSpace(value) Then Return Nothing
        Dim dt As DateTime
        If DateTime.TryParse(value, Globalization.CultureInfo.InvariantCulture,
                             Globalization.DateTimeStyles.AdjustToUniversal Or
                             Globalization.DateTimeStyles.AssumeUniversal, dt) Then
            Return dt
        End If
        Return Nothing
    End Function

    Private Function BuildHitsFromSearchToolResponses(responses As IEnumerable(Of ThisAddIn.ToolResponse)) As List(Of M365SearchHit)
        Dim hits As New List(Of M365SearchHit)()
        If responses Is Nothing Then Return hits

        Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim expectedName As String = SharedLibrary.SharedLibrary.M365ToolService.SearchToolName

        Dim idx As Integer = 0
        For Each r In responses
            idx += 1
            If r Is Nothing OrElse Not r.Success Then Continue For
            If Not String.Equals(r.ToolName, expectedName, StringComparison.OrdinalIgnoreCase) Then Continue For
            If String.IsNullOrWhiteSpace(r.Response) Then Continue For

            Dim parsed As JToken = Nothing
            Try
                parsed = JToken.Parse(r.Response)
            Catch ex As Exception
                Debug.WriteLine($"[AISearch] tool resp #{idx}: JSON parse failed: {ex.Message}")
            End Try
            If parsed Is Nothing Then Continue For

            Dim arr As JArray = Nothing
            If TypeOf parsed Is JObject Then
                Dim root As JObject = CType(parsed, JObject)
                arr = TryCast(root("hits"), JArray)
                If arr Is Nothing Then arr = TryCast(root("results"), JArray)
                If arr Is Nothing Then arr = TryCast(root("items"), JArray)
            ElseIf TypeOf parsed Is JArray Then
                arr = CType(parsed, JArray)
            End If

            If arr Is Nothing Then Continue For

            For Each tok As JObject In arr.OfType(Of JObject)()
                Dim src As String = If(tok("source")?.ToString(), "")
                If Not String.IsNullOrWhiteSpace(src) AndAlso
                   Not String.Equals(src, "mail", StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If

                Dim graphId As String = If(tok("id")?.ToString(),
                                           If(tok("hitId")?.ToString(), "")).Trim()

                Dim internetMessageId As String =
                    If(tok("internet_message_id")?.ToString(),
                       If(tok("internetMessageId")?.ToString(), "")).Trim()

                If String.IsNullOrWhiteSpace(graphId) Then
                    Dim webUrlCandidate As String = If(tok("web_url")?.ToString(),
                                                       If(tok("webLink")?.ToString(), "")).Trim()
                    If Not String.IsNullOrWhiteSpace(webUrlCandidate) Then
                        graphId = ExtractMessageIdFromWebLink(webUrlCandidate)
                    End If
                End If

                Dim dedupKey As String = If(Not String.IsNullOrWhiteSpace(graphId), graphId, internetMessageId)
                If String.IsNullOrWhiteSpace(dedupKey) OrElse Not seen.Add(dedupKey) Then Continue For

                Dim title As String = If(tok("title")?.ToString(),
                                         If(tok("subject")?.ToString(), ""))
                Dim summary As String = If(tok("summary")?.ToString(), "")
                Dim author As String = If(tok("author")?.ToString(),
                                          If(tok("from")?.ToString(), ""))
                Dim webUrl As String = If(tok("web_url")?.ToString(),
                                          If(tok("webLink")?.ToString(), ""))

                Dim sentDateStr As String = If(tok("sentDateTime")?.ToString(), "")
                Dim receivedDateStr As String = If(tok("receivedDateTime")?.ToString(),
                                                   If(tok("date")?.ToString(), ""))

                Dim toLine As String = ""
                Dim toRecipients = TryCast(tok("toRecipients"), JArray)
                If toRecipients IsNot Nothing Then
                    Dim names As New List(Of String)()
                    For Each rcp In toRecipients
                        Dim ea = TryCast(rcp("emailAddress"), JObject)
                        If ea Is Nothing Then Continue For
                        Dim nm As String = If(ea("name")?.ToString(), "")
                        Dim ad As String = If(ea("address")?.ToString(), "")
                        Dim disp As String = If(Not String.IsNullOrWhiteSpace(nm), nm, ad)
                        If Not String.IsNullOrWhiteSpace(disp) Then names.Add(disp.Trim())
                    Next
                    toLine = String.Join("; ", names.Distinct(StringComparer.OrdinalIgnoreCase))
                End If

                Dim hasAttachments As Boolean = False
                Dim hasAttachmentsTok As JToken = tok("has_attachments")
                If hasAttachmentsTok IsNot Nothing AndAlso hasAttachmentsTok.Type <> JTokenType.Null Then
                    If hasAttachmentsTok.Type = JTokenType.Boolean Then
                        hasAttachments = hasAttachmentsTok.Value(Of Boolean)()
                    Else
                        Boolean.TryParse(hasAttachmentsTok.ToString(), hasAttachments)
                    End If
                End If

                Dim h As New M365SearchHit() With {
                    .Source = M365SearchSources.Mail,
                    .Id = graphId,
                    .Title = title,
                    .Summary = summary,
                    .Author = author,
                    .AdditionalText = toLine,
                    .WebUrl = webUrl,
                    .LastModifiedUtc = TryParseUtc(If(Not String.IsNullOrWhiteSpace(sentDateStr), sentDateStr, receivedDateStr)),
                    .RawJson = New JObject(
                        New JProperty("hitId", graphId),
                        New JProperty("resource",
                            New JObject(
                                New JProperty("id", graphId),
                                New JProperty("internetMessageId", internetMessageId),
                                New JProperty("conversationId", If(tok("conversation_id")?.ToString(), "")),
                                New JProperty("webLink", webUrl),
                                New JProperty("subject", title),
                                New JProperty("sentDateTime", sentDateStr),
                                New JProperty("receivedDateTime", receivedDateStr),
                                New JProperty("hasAttachments", hasAttachments),
                                New JProperty("toRecipients", If(toRecipients, New JArray()))
                            )))
                }

                hits.Add(h)
            Next
        Next

        Debug.WriteLine($"[AISearch] BuildHitsFromSearchToolResponses produced {hits.Count} hit(s).")
        Return hits
    End Function

    Private Function MatchAiReturnedIdsToToolHits(
        returnedIds As IEnumerable(Of String),
        toolHits As IEnumerable(Of M365SearchHit)) As List(Of M365SearchHit)

        Dim result As New List(Of M365SearchHit)()
        If returnedIds Is Nothing OrElse toolHits Is Nothing Then Return result

        Dim byGraphId As New Dictionary(Of String, M365SearchHit)(StringComparer.OrdinalIgnoreCase)
        Dim byInternetMessageId As New Dictionary(Of String, M365SearchHit)(StringComparer.OrdinalIgnoreCase)
        Dim added As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        Dim toolHitCount As Integer = 0

        For Each h In toolHits
            If h Is Nothing Then Continue For
            toolHitCount += 1

            Dim graphId As String = GetHitMessageId(h)
            Dim graphIdKey As String = NormalizeGraphIdForCompare(graphId)
            If Not String.IsNullOrWhiteSpace(graphIdKey) AndAlso Not byGraphId.ContainsKey(graphIdKey) Then
                byGraphId.Add(graphIdKey, h)
            End If

            Dim internetMessageId As String = NormalizeInternetMessageId(TryGetInternetMessageId(h))
            If Not String.IsNullOrWhiteSpace(internetMessageId) AndAlso
               Not byInternetMessageId.ContainsKey(internetMessageId) Then
                byInternetMessageId.Add(internetMessageId, h)
            End If
        Next

        Debug.WriteLine($"[AISearch] toolHits={toolHitCount}, graphIds={byGraphId.Count}, internetMessageIds={byInternetMessageId.Count}")

        For Each rawId In returnedIds
            Dim id As String = If(rawId, "").Trim()
            If String.IsNullOrWhiteSpace(id) Then Continue For

            Dim matched As M365SearchHit = Nothing
            Dim graphKey As String = NormalizeGraphIdForCompare(id)
            Dim internetKey As String = NormalizeInternetMessageId(id)

            If byGraphId.TryGetValue(graphKey, matched) OrElse
               byInternetMessageId.TryGetValue(internetKey, matched) Then

                Dim graphId As String = ToGraphUrlSafeId(GetHitMessageId(matched))
                If Not String.IsNullOrWhiteSpace(graphId) Then
                    matched.Id = graphId
                    If added.Add(graphId) Then
                        result.Add(matched)
                    End If
                End If
            Else
                Debug.WriteLine("[AISearch] No tool-hit match for returned id: " & id)
                Debug.WriteLine("[AISearch]   graphKey=" & graphKey)
                Debug.WriteLine("[AISearch]   internetKey=" & internetKey)
            End If
        Next

        Debug.WriteLine("[AISearch] matched result count=" & result.Count)
        Return result
    End Function


    Private Function MatchAiReturnedRefsToToolHits(
        returnedRefs As IEnumerable(Of Integer),
        toolHits As IEnumerable(Of M365SearchHit)) As List(Of M365SearchHit)

        Dim result As New List(Of M365SearchHit)()
        If returnedRefs Is Nothing OrElse toolHits Is Nothing Then Return result

        Dim orderedHits As List(Of M365SearchHit) = toolHits.Where(Function(h) h IsNot Nothing).ToList()
        Dim added As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each n In returnedRefs
            If n < 1 OrElse n > orderedHits.Count Then
                Debug.WriteLine("[AISearch] Ignoring out-of-range email_ref n=" & n.ToString())
                Continue For
            End If

            Dim hit As M365SearchHit = orderedHits(n - 1)
            If hit Is Nothing Then Continue For

            Dim graphId As String = GetHitMessageId(hit)
            If String.IsNullOrWhiteSpace(graphId) Then
                graphId = If(hit.Id, "").Trim()
            End If

            If String.IsNullOrWhiteSpace(graphId) Then
                graphId = If(TryGetInternetMessageId(hit), "").Trim()
            End If

            Dim dedupKey As String = If(graphId, "").Trim()
            If String.IsNullOrWhiteSpace(dedupKey) Then
                dedupKey = "ref:" & n.ToString()
            End If

            If added.Add(dedupKey) Then
                result.Add(hit)
            End If
        Next

        Debug.WriteLine("[AISearch] matched refs result count=" & result.Count.ToString())
        Return result
    End Function


    Private Function FindMailInDefaultStoreByHeuristics(app As Outlook.Application,
                                                    subject As String,
                                                    author As String,
                                                    mailDateUtc As DateTime?) As Outlook.MailItem
        If app Is Nothing Then Return Nothing
        If String.IsNullOrWhiteSpace(subject) AndAlso String.IsNullOrWhiteSpace(author) AndAlso Not mailDateUtc.HasValue Then
            Return Nothing
        End If

        Dim session = app.Session
        If session Is Nothing OrElse session.DefaultStore Is Nothing Then Return Nothing

        Dim quickTargets = New Outlook.OlDefaultFolders() {
            Outlook.OlDefaultFolders.olFolderInbox,
            Outlook.OlDefaultFolders.olFolderSentMail,
            Outlook.OlDefaultFolders.olFolderDrafts,
            Outlook.OlDefaultFolders.olFolderDeletedItems,
            Outlook.OlDefaultFolders.olFolderJunk,
            Outlook.OlDefaultFolders.olFolderOutbox
        }

        For Each ft In quickTargets
            Try
                Dim folder = TryCast(session.GetDefaultFolder(ft), Outlook.Folder)
                Dim found = FindMailInFolderByHeuristics(folder, subject, author, mailDateUtc, 750)
                If found IsNot Nothing Then Return found
            Catch ex As Exception
                Debug.WriteLine("[M365Search] Heuristic folder skipped: " & ex.Message)
            End Try
        Next

        Return Nothing
    End Function

    Private Function FindMailInFolderByHeuristics(folder As Outlook.Folder,
                                                  subject As String,
                                                  author As String,
                                                  mailDateUtc As DateTime?,
                                                  scanCap As Integer) As Outlook.MailItem
        If folder Is Nothing Then Return Nothing

        Try
            If folder.DefaultItemType <> Outlook.OlItemType.olMailItem Then Return Nothing
        Catch
            Return Nothing
        End Try

        Dim subjectNorm As String = NormalizeMailCompare(subject)
        Dim authorNorm As String = NormalizeMailCompare(author)

        Try
            Dim items = folder.Items
            If items Is Nothing Then Return Nothing

            Dim count As Integer = 0
            Dim maxCount As Integer = Math.Min(items.Count, scanCap)

            For i As Integer = 1 To maxCount
                Dim mi = TryCast(items(i), Outlook.MailItem)
                If mi Is Nothing Then Continue For

                count += 1

                If IsHeuristicMailMatch(mi, subjectNorm, authorNorm, mailDateUtc) Then
                    Debug.WriteLine($"[M365Search] Heuristic local match in folder '{folder.Name}' after scanning {count} item(s).")
                    Return mi
                End If
            Next
        Catch ex As Exception
            Debug.WriteLine("[M365Search] Heuristic folder scan failed '" & folder.Name & "': " & ex.Message)
        End Try

        Return Nothing
    End Function

    Private Function IsHeuristicMailMatch(mail As Outlook.MailItem,
                                          subjectNorm As String,
                                          authorNorm As String,
                                          mailDateUtc As DateTime?) As Boolean
        If mail Is Nothing Then Return False

        Dim subj As String = ""
        Dim senderName As String = ""
        Dim senderAddress As String = ""
        Dim localDate As DateTime? = Nothing

        Try : subj = NormalizeMailCompare(mail.Subject) : Catch : End Try
        Try : senderName = NormalizeMailCompare(mail.SenderName) : Catch : End Try
        Try : senderAddress = NormalizeMailCompare(mail.SenderEmailAddress) : Catch : End Try

        Try
            ' Sent items usually have SentOn; received items ReceivedTime.
            If mail.Sent Then
                localDate = mail.SentOn.ToUniversalTime()
            Else
                localDate = mail.ReceivedTime.ToUniversalTime()
            End If
        Catch
        End Try

        Dim subjectMatch As Boolean =
            String.IsNullOrWhiteSpace(subjectNorm) OrElse
            String.Equals(subj, subjectNorm, StringComparison.OrdinalIgnoreCase)

        Dim authorMatch As Boolean =
            String.IsNullOrWhiteSpace(authorNorm) OrElse
            senderName.Contains(authorNorm) OrElse
            authorNorm.Contains(senderName) OrElse
            senderAddress.Contains(authorNorm) OrElse
            authorNorm.Contains(senderAddress)

        Dim dateMatch As Boolean = True
        If mailDateUtc.HasValue AndAlso localDate.HasValue Then
            Dim delta As TimeSpan = (localDate.Value - mailDateUtc.Value).Duration()
            dateMatch = delta.TotalMinutes <= 10
        End If

        ' Require subject, and either sender or date alignment when available.
        If Not subjectMatch Then Return False
        If Not String.IsNullOrWhiteSpace(authorNorm) AndAlso Not authorMatch Then Return False
        If mailDateUtc.HasValue AndAlso Not dateMatch Then Return False

        Return True
    End Function

    Private Shared Function NormalizeMailCompare(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return ""
        Return value.Trim().ToLowerInvariant()
    End Function

    Private Shared Function NormalizeGraphIdForCompare(value As String) As String
        Dim s As String = If(value, "").Trim()
        If String.IsNullOrWhiteSpace(s) Then Return ""
        s = s.Replace(" ", "")
        ' Compare in standard Base64 form
        s = s.Replace("-"c, "+"c).Replace("_"c, "/"c)
        Return s
    End Function

    Private Shared Function ToGraphUrlSafeId(value As String) As String
        Dim s As String = NormalizeGraphIdForCompare(value)
        If String.IsNullOrWhiteSpace(s) Then Return ""
        Return s.Replace("+"c, "-"c).Replace("/"c, "_"c)
    End Function


    Private Shared Function ParseAiSourceName(name As String) As M365SearchSources
        Select Case If(name, "").Trim().ToLowerInvariant()
            Case "mail" : Return M365SearchSources.Mail
            Case "onedrive" : Return M365SearchSources.OneDrive
            Case "sharepoint" : Return M365SearchSources.SharePoint
            Case "sharepoint_sites" : Return M365SearchSources.SharePointSites
            Case "sharepoint_listitems" : Return M365SearchSources.SharePointListItems
            Case "teams" : Return M365SearchSources.Teams
            Case "calendar" : Return M365SearchSources.Calendar
            Case "onenote" : Return M365SearchSources.OneNote
            Case "people" : Return M365SearchSources.People
            Case "all_files" : Return M365SearchSources.OneDrive Or M365SearchSources.SharePoint
            Case "all_sharepoint" : Return M365SearchSources.SharePoint Or M365SearchSources.SharePointSites Or M365SearchSources.SharePointListItems
            Case "all" : Return M365SearchSources.All
            Case Else : Return M365SearchSources.None
        End Select
    End Function

    Private Shared Function ParseAiSourcesToken(tok As JToken) As M365SearchSources
        If tok Is Nothing Then Return M365SearchSources.None

        Dim combined As M365SearchSources = M365SearchSources.None

        If tok.Type = JTokenType.Array Then
            For Each item In CType(tok, JArray)
                combined = combined Or ParseAiSourceName(item.ToString())
            Next
        ElseIf tok.Type = JTokenType.String Then
            combined = ParseAiSourceName(tok.ToString())
        End If

        Return combined
    End Function

    Private Shared Function ParseNullableDate(value As String) As Date?
        If String.IsNullOrWhiteSpace(value) Then Return Nothing
        Dim dt As DateTime
        Dim inv = Globalization.CultureInfo.InvariantCulture
        Dim styles = Globalization.DateTimeStyles.AssumeUniversal Or
                     Globalization.DateTimeStyles.AdjustToUniversal

        ' Strict ISO first (what the LLM / tool args always send).
        Dim isoFormats As String() = {
            "yyyy-MM-dd",
            "yyyy-MM-dd'T'HH:mm:ss",
            "yyyy-MM-dd'T'HH:mm:ss'Z'",
            "yyyy-MM-dd'T'HH:mm:ssK",
            "yyyy-MM-dd'T'HH:mm:ss.FFFFFFF'Z'",
            "yyyy-MM-dd'T'HH:mm:ss.FFFFFFFK"
        }
        If DateTime.TryParseExact(value.Trim(), isoFormats, inv, styles, dt) Then
            Return dt.Date
        End If

        ' Fallback: invariant general parse (never current culture).
        If DateTime.TryParse(value, inv, styles, dt) Then Return dt.Date
        Return Nothing
    End Function

    Private Function ParseAiSearchInvocation(tr As ThisAddIn.ToolResponse) As AiSearchInvocation
        If tr Is Nothing OrElse Not tr.Success Then Return Nothing
        If Not String.Equals(tr.ToolName, SharedLibrary.SharedLibrary.M365ToolService.SearchToolName, StringComparison.OrdinalIgnoreCase) Then Return Nothing

        Dim query As String = ""
        Dim sources As M365SearchSources = M365SearchSources.None
        Dim maxPerSource As Integer = 25
        Dim fromIndex As Integer = 0
        Dim fromDate As Date? = Nothing
        Dim toDate As Date? = Nothing
        Dim kqlExtra As String = ""

        Try
            If Not String.IsNullOrWhiteSpace(tr.OriginalCallJson) Then
                Dim root As JToken = JToken.Parse(tr.OriginalCallJson)
                Dim fc As JObject = TryCast(root("functionCall"), JObject)
                If fc Is Nothing Then fc = TryCast(root, JObject)

                Dim args As JObject = TryCast(fc?("args"), JObject)
                If args IsNot Nothing Then
                    query = If(args("query")?.ToString(), "").Trim()
                    sources = ParseAiSourcesToken(args("sources"))
                    maxPerSource = Math.Max(1, Math.Min(CInt(If(args("max_per_source"), 25)), 500))
                    fromIndex = Math.Max(0, CInt(If(args("from_index"), 0)))
                    fromDate = ParseNullableDate(If(args("from_date")?.ToString(), ""))
                    toDate = ParseNullableDate(If(args("to_date")?.ToString(), ""))
                    kqlExtra = If(args("kql_extra")?.ToString(), "")
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine("[AISearch] Failed to parse OriginalCallJson for search invocation: " & ex.Message)
        End Try

        If String.IsNullOrWhiteSpace(query) Then
            Try
                Dim envelope As JObject = JObject.Parse(If(tr.Response, ""))
                query = If(envelope("query")?.ToString(), "").Trim()
                If sources = M365SearchSources.None Then
                    sources = ParseAiSourceName(If(envelope("requested_sources")?.ToString(), ""))
                End If
                maxPerSource = Math.Max(1, Math.Min(CInt(If(envelope("requested_max_per_source"), maxPerSource)), 500))
                fromIndex = Math.Max(0, CInt(If(envelope("requested_from_index"), fromIndex)))
                If String.IsNullOrWhiteSpace(kqlExtra) Then
                    kqlExtra = If(envelope("requested_kql_extra")?.ToString(), "")
                End If
            Catch ex As Exception
                Debug.WriteLine("[AISearch] Failed to parse search response envelope: " & ex.Message)
            End Try
        End If

        If String.IsNullOrWhiteSpace(query) Then Return Nothing
        If sources = M365SearchSources.None Then sources = M365SearchSources.Mail

        Return New AiSearchInvocation() With {
            .Query = query,
            .Sources = sources,
            .MaxPerSource = maxPerSource,
            .FromDate = fromDate,
            .ToDate = toDate,
            .KqlExtra = kqlExtra,
            .NextFromIndex = fromIndex + maxPerSource,
            .Exhausted = False
        }
    End Function

    Private Function CaptureAiSearchInvocations(toolResponses As IEnumerable(Of ThisAddIn.ToolResponse)) As List(Of AiSearchInvocation)
        Dim result As New List(Of AiSearchInvocation)()
        If toolResponses Is Nothing Then Return result

        For Each tr In toolResponses
            Dim parsed As AiSearchInvocation = ParseAiSearchInvocation(tr)
            If parsed IsNot Nothing Then result.Add(parsed)
        Next

        Return result
    End Function

    Private Function MergeSearchInvocations(existingItems As IEnumerable(Of AiSearchInvocation),
                                            newItems As IEnumerable(Of AiSearchInvocation)) As List(Of AiSearchInvocation)
        Dim merged As New Dictionary(Of String, AiSearchInvocation)(StringComparer.OrdinalIgnoreCase)

        If existingItems IsNot Nothing Then
            For Each item In existingItems
                If item Is Nothing Then Continue For
                Dim key As String = item.StableKey()
                If Not merged.ContainsKey(key) Then
                    merged.Add(key, item)
                End If
            Next
        End If

        If newItems IsNot Nothing Then
            For Each item In newItems
                If item Is Nothing Then Continue For
                Dim key As String = item.StableKey()
                If merged.ContainsKey(key) Then
                    merged(key).NextFromIndex = Math.Max(merged(key).NextFromIndex, item.NextFromIndex)
                Else
                    merged.Add(key, item)
                End If
            Next
        End If

        Return merged.Values.ToList()
    End Function

    Private Sub UpdateGetMoreEnabled()
        UiPost(Sub()
                   If btnGetMore Is Nothing Then Return
                   Dim haveRoom As Boolean = _aiCandidateHits.Count < AISearch_MaxCandidateHits
                   Dim haveSearches As Boolean = _aiRecordedSearches.Any(Function(x) x IsNot Nothing AndAlso Not x.Exhausted)
                   btnGetMore.Enabled = haveRoom AndAlso haveSearches
               End Sub)
    End Sub

    Private Async Function GetMoreCandidatesFromRecordedQueriesAsync(ct As CancellationToken) As Task(Of List(Of M365SearchHit))
        Dim additional As New List(Of M365SearchHit)()
        If _aiRecordedSearches Is Nothing OrElse _aiRecordedSearches.Count = 0 Then Return additional

        Dim remainingRoom As Integer = AISearch_MaxCandidateHits - _aiCandidateHits.Count
        If remainingRoom <= 0 Then Return additional

        Dim perPageTarget As Integer = Math.Max(1, Math.Min(GetAiHarvestMaxPerCall(), remainingRoom))

        For Each inv In _aiRecordedSearches.Where(Function(x) x IsNot Nothing AndAlso Not x.Exhausted).ToList()
            ct.ThrowIfCancellationRequested()

            Dim pageSize As Integer = Math.Max(1, Math.Min(inv.MaxPerSource, perPageTarget))
            Dim opts As New M365SearchOptions() With {
                .MaxPerSource = pageSize,
                .FromIndex = inv.NextFromIndex,
                .From = inv.FromDate,
                .To = inv.ToDate,
                .KqlExtra = inv.KqlExtra,
                .Parallel = True
            }

            UiPost(Sub() lblStatus.Text = $"Get more: searching next page at offset {inv.NextFromIndex} ({pageSize} requested)…")

            Dim res As M365SearchResult = Await M365Service.SearchAsync(
                _context,
                inv.Query,
                inv.Sources,
                opts,
                Nothing,
                ct).ConfigureAwait(False)

            If res Is Nothing OrElse res.Hits.Count = 0 Then
                inv.Exhausted = True
                Continue For
            End If

            additional.AddRange(res.Hits)
            inv.NextFromIndex += pageSize

            If res.Hits.Count < pageSize Then
                inv.Exhausted = True
            End If

            If additional.Count >= perPageTarget Then Exit For
        Next

        Return additional
    End Function

    Private Async Function ReviewAndDisplayCandidatePoolAsync(userPrompt As String,
                                                          ct As CancellationToken) As Task
        If _aiCandidateHits.Count = 0 Then
            UiPost(Sub()
                       _hits.Clear()
                       lvResults.Items.Clear()
                       _hasPinnedSummary = True
                       SetSummaryMarkdown("No matching e-mails were found.")
                       UpdateAiStats(0, 0, 0, 0)
                       UpdateGetMoreEnabled()
                       lblStatus.Text = "AI search found no candidate mails."
                   End Sub)
            Return
        End If

        UiPost(Sub() lblStatus.Text = $"Reviewing {_aiCandidateHits.Count} candidate mail(s) in full text…")

        Dim resolvedCandidates As List(Of AiResolvedMail) =
        Await FetchResolvedCandidateMailsAsync(_aiCandidateHits, ct).ConfigureAwait(False)

        Dim reviewedCount As Integer = resolvedCandidates.Count

        Dim shortlistRefs As List(Of Integer) =
        Await ReviewCandidatesInBatchesAsync(userPrompt, resolvedCandidates, ct).ConfigureAwait(False)

        Dim shortlistSet As New HashSet(Of Integer)(shortlistRefs)
        Dim shortlistResolved As List(Of AiResolvedMail) =
        resolvedCandidates.Where(Function(x) shortlistSet.Contains(x.GlobalRef)).ToList()

        Dim keptCount As Integer = shortlistResolved.Count

        UpdateAiStats(_aiCandidateHits.Count, reviewedCount, keptCount)
        UiPost(Sub() lblStatus.Text = $"Full-text review kept {keptCount} of {reviewedCount} reviewed candidate mails. Building final result…")

        Dim finalSelection As AiFinalSelection =
        Await BuildFinalGridSelectionAsync(userPrompt, shortlistResolved, ct).ConfigureAwait(False)

        If shortlistResolved.Count = 0 Then
            UiPost(Sub()
                       _hits = New List(Of M365SearchHit)()
                       lvResults.Items.Clear()
                       _hasPinnedSummary = True
                       SetSummaryMarkdown("No matching e-mails were found.")
                       UpdateAiStats(_aiCandidateHits.Count, reviewedCount, keptCount, 0)
                       UpdateGetMoreEnabled()
                       lblStatus.Text = $"0 relevant mails shown from {keptCount} kept / {_aiCandidateHits.Count} retrieved."
                   End Sub)
            Return
        End If

        Dim byGlobalRef As Dictionary(Of Integer, AiResolvedMail) =
        shortlistResolved.ToDictionary(Function(x) x.GlobalRef)

        Dim finalHits As New List(Of M365SearchHit)()
        Dim addedRefs As New HashSet(Of Integer)()

        ' First: honor the AI-provided order.
        For Each globalRef In finalSelection.OrderedGlobalRefs
            Dim resolved As AiResolvedMail = Nothing
            If byGlobalRef.TryGetValue(globalRef, resolved) AndAlso
           resolved IsNot Nothing AndAlso
           resolved.Hit IsNot Nothing AndAlso
           addedRefs.Add(globalRef) Then

                SetAiDisplayRef(resolved.Hit, globalRef)
                finalHits.Add(resolved.Hit)
            End If
        Next

        ' Then: append every remaining kept mail so nothing mentioned in the
        ' summary context can disappear from the grid.
        For Each resolved In shortlistResolved
            If resolved Is Nothing OrElse resolved.Hit Is Nothing Then Continue For
            If addedRefs.Add(resolved.GlobalRef) Then
                SetAiDisplayRef(resolved.Hit, resolved.GlobalRef)
                finalHits.Add(resolved.Hit)
            End If
        Next

        Await EnrichMailHitsAsync(finalHits, ct).ConfigureAwait(False)

        UiPost(Sub()
                   _hits = finalHits
                   PopulateResults(_hits)
                   _hasPinnedSummary = True
                   SetSummaryMarkdown(If(String.IsNullOrWhiteSpace(finalSelection.Summary),
                                     "No matching e-mails were found.",
                                     finalSelection.Summary.Trim()))
                   UpdateAiStats(_aiCandidateHits.Count, reviewedCount, keptCount, _hits.Count)
                   UpdateGetMoreEnabled()
                   lblStatus.Text = $"Showing {_hits.Count} relevant mail(s) from {keptCount} kept / {_aiCandidateHits.Count} retrieved."
               End Sub)

        UiPost(Sub() StartRowSummaries())
    End Function


    Private Function TryGetAiSearchDefaultModel(ByRef modelConfig As ModelConfig) As Boolean
        modelConfig = Nothing
        _aiSearchUnavailableReason = ""

        Dim altPath As String = If(_context?.INI_AlternateModelPath, "")
        If String.IsNullOrWhiteSpace(altPath) Then
            _aiSearchUnavailableReason = "AI Search is unavailable because INI_AlternateModelPath is not configured."
            Return False
        End If

        Dim candidate As ModelConfig = Nothing
        Dim unsupportedDefaults As New List(Of String)()

        If SharedMethods.TryGetSpecialTaskModelConfig(_context, altPath, "ToolDefaultModel", candidate) Then
            If SharedMethods.ModelSupportsTooling(candidate) Then
                modelConfig = candidate
                Return True
            End If

            unsupportedDefaults.Add("ToolDefaultModel")
        End If

        candidate = Nothing

        If SharedMethods.TryGetSpecialTaskModelConfig(_context, altPath, "AgentDefaultModel", candidate) Then
            If SharedMethods.ModelSupportsTooling(candidate) Then
                modelConfig = candidate
                Return True
            End If

            unsupportedDefaults.Add("AgentDefaultModel")
        End If

        If unsupportedDefaults.Count > 0 Then
            _aiSearchUnavailableReason =
                "AI Search is unavailable because the configured " &
                String.Join(" or ", unsupportedDefaults) &
                " does not support tool calling."
        Else
            _aiSearchUnavailableReason =
                "AI Search is unavailable because no tooling-capable ToolDefaultModel or AgentDefaultModel is configured in the alternate models INI."
        End If

        Return False
    End Function

    Private Sub RefreshAiSearchAvailability()
        Dim aiModel As ModelConfig = Nothing
        _aiSearchAvailable = TryGetAiSearchDefaultModel(aiModel)

        If btnAISearch IsNot Nothing Then
            btnAISearch.Enabled = _aiSearchAvailable
            _toolTip.SetToolTip(
                btnAISearch,
                If(_aiSearchAvailable,
                   "Run multi-step AI mail search: broad candidate gathering first, then full-text review of the candidate mails. Slower, but more accurate for semantic requests.",
                   _aiSearchUnavailableReason))
        End If

        If Not _aiSearchAvailable AndAlso btnGetMore IsNot Nothing Then
            btnGetMore.Enabled = False
        End If
    End Sub

End Class

''' <summary>Convenience launcher — call from any ribbon button or the Immediate window.</summary>
Public Module M365SearchTest
    Public Sub Show(context As ISharedContext)
        Dim f As New M365SearchTestForm(context)
        f.Show()
    End Sub
End Module

