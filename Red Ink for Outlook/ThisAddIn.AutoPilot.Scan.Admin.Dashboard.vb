' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.Scan.Admin.Dashboard.vb
' Purpose: Modeless AutoPilot Scan Handler admin dashboard.
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

Namespace ScanInbox
    Friend Class AdminHost
        Public Property GetMailboxes As System.Func(Of System.Collections.Generic.List(Of System.String))
        Public Property RuntimeSummary As System.Func(Of System.String)
        Public Property ReadActivity As System.Func(Of System.Threading.Tasks.Task(Of System.Collections.Generic.List(Of AdminActivity)))
        Public Property ReadSearchProgress As System.Func(Of System.Threading.Tasks.Task(Of System.Collections.Generic.List(Of AdminSearchProgress)))
    End Class

    Friend Class AdminDashboard
        Inherits System.Windows.Forms.Form
        Private ReadOnly _session As AdminSession
        Private ReadOnly _host As AdminHost
        Private ReadOnly _tabs As New System.Windows.Forms.TabControl With {.Dock = System.Windows.Forms.DockStyle.Fill}
        Private ReadOnly _recipients As System.Windows.Forms.DataGridView = AdminUi.Grid("Name", "Email", "Organization", "Status", "Used by recipient sets")
        Private ReadOnly _sets As System.Windows.Forms.DataGridView = AdminUi.Grid("Recipient set", "Scanner senders", "Mailbox", "Default recipient", "Status", "Allowed recipients")
        Private ReadOnly _activity As System.Windows.Forms.DataGridView = AdminUi.Grid("Received (local)", "Recipient set", "Status", "Delivery", "Issue")
        Private ReadOnly _search As System.Windows.Forms.TextBox = AdminUi.Text("")
        Private ReadOnly _statusFilter As New System.Windows.Forms.ComboBox With {.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList}
        Private ReadOnly _summary As New System.Windows.Forms.Label With {.AutoSize = True, .Dock = System.Windows.Forms.DockStyle.Fill, .UseMnemonic = False}
        Private ReadOnly _status As New System.Windows.Forms.Label With {.AutoSize = True, .Dock = System.Windows.Forms.DockStyle.Fill, .UseMnemonic = False}
        Private ReadOnly _configurationPages As New System.Collections.Generic.List(Of System.Windows.Forms.TabPage)()
        Private ReadOnly _timer As New System.Windows.Forms.Timer With {.Interval = 15000}
        Private _items As New System.Collections.Generic.List(Of AdminActivity)()
        Private _busy As System.Boolean
        Private _refreshing As System.Boolean
        Private _closing As System.Boolean
        Private _bypassClose As System.Boolean

        Public Sub New(path As System.String, host As AdminHost)
            _host = host
            _session = New AdminSession(path)
            AdminUi.Style(Me, "AutoPilot Scan Handler", New System.Drawing.Size(1120, 740))
            MinimumSize = New System.Drawing.Size(780, 520)
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            ShowInTaskbar = True
            Dim root = AdminUi.Layout()
            root.AutoSize = False
            root.RowCount = 4
            root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100))
            root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            root.Controls.Add(_summary, 0, 0)
            root.Controls.Add(_tabs, 0, 1)
            root.Controls.Add(_status, 0, 2)
            Dim reload = AdminUi.Button("Reload", Sub() ReloadAsync())
            Dim backup = AdminUi.Button("Restore previous version…", AddressOf RestoreBackup)
            Dim export = AdminUi.Button("Export configuration…", AddressOf ExportConfiguration)
            Dim import = AdminUi.Button("Import configuration…", AddressOf ImportConfiguration)
            Dim close = AdminUi.Button("Close", Sub() Me.Close())
            root.Controls.Add(AdminUi.Footer(close, reload, export, import, backup), 0, 3)
            Controls.Add(root)
            AddHandler root.SizeChanged, Sub()
                                             Dim width = System.Math.Max(120, root.ClientSize.Width - root.Padding.Horizontal - 12)
                                             _status.MaximumSize = New System.Drawing.Size(width, 0)
                                             _summary.MaximumSize = New System.Drawing.Size(width, 0)
                                         End Sub
            Dim page = DataPage("Recipient register", _recipients, AdminUi.Buttons(AdminUi.Button("Add recipient", Sub() EditRecipient(True)), AdminUi.Button("Import CSV…", AddressOf ImportNames), AdminUi.Button("Export CSV…", AddressOf ExportRecipients), AdminUi.Button("Edit", Sub() EditRecipient(False)), AdminUi.Button("Remove", AddressOf DeleteRecipient)))
            _configurationPages.Add(page)
            page = DataPage("Recipient sets", _sets, AdminUi.Buttons(AdminUi.Button("Add recipient set", Sub() EditSet(True)), AdminUi.Button("Edit", Sub() EditSet(False)), AdminUi.Button("Duplicate", AddressOf DuplicateSet), AdminUi.Button("Enable / disable", AddressOf ToggleSet), AdminUi.Button("Remove", AddressOf DeleteSet)))
            _configurationPages.Add(page)
            Dim activityPage = DataPage("Activity", _activity, AdminUi.Buttons(AdminUi.Button("Refresh", Sub() RefreshActivityAsync()), AdminUi.Button("Search progress", AddressOf ShowSearchProgress), AdminUi.Button("Details", AddressOf ShowDetails)))
            Dim activityRoot = DirectCast(activityPage.Controls(0), System.Windows.Forms.TableLayoutPanel)
            activityRoot.RowCount = 3
            activityRoot.RowStyles.Insert(0, New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            activityRoot.SetRow(_activity, 1)
            activityRoot.SetRow(activityRoot.Controls.Cast(Of System.Windows.Forms.Control)().Single(Function(c) TypeOf c Is System.Windows.Forms.FlowLayoutPanel), 2)
            _search.MinimumSize = New System.Drawing.Size(220, 0)
            _search.AccessibleName = "Search activity"
            AdminUi.Help(_activity, "Activity records are retained to prevent duplicate delivery. Keep scan content controls attachment and analysis copies, not this history. Failed or uncertain packages retain their files for recovery.")
            _statusFilter.Items.AddRange({"All packages", "Needs attention", "In progress", "Complete"})
            _statusFilter.SelectedIndex = 0
            activityRoot.Controls.Add(AdminUi.Buttons(New System.Windows.Forms.Label With {.Text = "Search", .AutoSize = True, .Anchor = System.Windows.Forms.AnchorStyles.Left}, _search, _statusFilter), 0, 0)
            AddHandler _search.TextChanged, Sub() DisplayActivity()
            AddHandler _statusFilter.SelectedIndexChanged, Sub() DisplayActivity()
            AddHandler _recipients.CellDoubleClick, Sub(sender, args)
                                                        If args.RowIndex >= 0 Then EditRecipient(False)
                                                    End Sub
            AddHandler _sets.CellDoubleClick, Sub(sender, args)
                                                  If args.RowIndex >= 0 Then EditSet(False)
                                              End Sub
            AddHandler _activity.CellDoubleClick, Sub(sender, args)
                                                      If args.RowIndex >= 0 Then ShowDetails()
                                                  End Sub
            AddHandler Shown, Sub()
                                  SetBusy(False)
                                  ReloadAsync()
                              End Sub
            AddHandler _timer.Tick, Sub() RefreshActivityAsync()
            AddHandler FormClosing, AddressOf HandleClosing
            AddHandler FormClosed, Sub()
                                       _timer.Stop()
                                       _timer.Dispose()
                                   End Sub
            SetBusy(True)
        End Sub

        Private Function DataPage(title As System.String, grid As System.Windows.Forms.DataGridView, buttons As System.Windows.Forms.Control) As System.Windows.Forms.TabPage
            Dim page As New System.Windows.Forms.TabPage(title) With {.UseVisualStyleBackColor = True}
            Dim root = AdminUi.Layout()
            root.AutoSize = False
            root.RowCount = 2
            root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100))
            root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            root.Controls.Add(grid, 0, 0)
            Dim actions = TryCast(buttons, System.Windows.Forms.FlowLayoutPanel)
            If actions IsNot Nothing Then
                Dim ordered = actions.Controls.Cast(Of System.Windows.Forms.Control)().Reverse().ToArray()
                actions.Controls.Clear()
                actions.Controls.AddRange(ordered)
                actions.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft
                actions.Padding = New System.Windows.Forms.Padding(0, 4, 0, 0)
            End If
            root.Controls.Add(buttons, 0, 1)
            page.Controls.Add(root)
            _tabs.TabPages.Add(page)
            Return page
        End Function

        Private Sub SetBusy(busy As System.Boolean)
            _busy = busy
            _tabs.Enabled = Not busy
            UseWaitCursor = busy
        End Sub

        Private Function HasChanges() As System.Boolean
            Return _session.HasChanges
        End Function

        Private Function DiscardChanges() As System.Boolean
            Return Not HasChanges() OrElse AdminUi.Confirm("Discard unsaved changes?", "Discard")
        End Function

        Private Async Sub ReloadAsync()
            If _busy Then Return
            If Not DiscardChanges() Then Return
            SetBusy(True)
            Try
                Await System.Threading.Tasks.Task.Run(Sub() _session.Reload())
                If IsDisposed Then Return
                BindConfiguration()
                SetStatus(If(_session.LoadError = "", "Changes are saved when you click Save in an editor. Topic routing is used only when no named recipient matches.", "Configuration unavailable: " & _session.LoadError & " Import a configuration or restore the previous version."), _session.LoadError <> "")
                _timer.Start()
            Catch ex As System.Exception
                If Not IsDisposed Then
                    BindConfiguration()
                    SetStatus("Could not load configuration: " & ex.Message, True)
                End If
            Finally
                If Not IsDisposed Then SetBusy(False)
            End Try
            Await RefreshActivityCoreAsync()
        End Sub

        Private Sub BindConfiguration()
            For Each page In _configurationPages
                page.Enabled = _session.Configuration IsNot Nothing
            Next
            If _session.Configuration Is Nothing Then
                _recipients.Rows.Clear()
                _sets.Rows.Clear()
                Return
            End If
            RefreshConfigurationLists()
        End Sub

        Private Sub RefreshConfigurationLists()
            Dim recipientId = AdminUi.SelectedId(_recipients)
            Dim setId = AdminUi.SelectedId(_sets)
            _recipients.Rows.Clear()
            For Each person In _session.Configuration.Recipients.OrderBy(Function(p) p.Name)
                Dim index = _recipients.Rows.Add(person.Name, person.SmtpAddress, person.Organization, If(person.Active, "Enabled", "Disabled"), _session.RecipientReferences(person.Id))
                _recipients.Rows(index).Tag = person.Id

            Next
            _sets.Rows.Clear()
            For Each rule In _session.Configuration.RuleSets.OrderBy(Function(r) r.Name)
                Dim person = _session.Configuration.Recipients.SingleOrDefault(Function(p) p.Id = rule.DefaultRecipientId)
                Dim index = _sets.Rows.Add(rule.Name, System.String.Join(", ", rule.ScannerSenders), If(rule.Mailbox = "", "All monitored", rule.Mailbox), If(person Is Nothing, "Not selected", person.Name), If(rule.Active, "Enabled", "Disabled"), System.String.Join(", ", _session.Configuration.Recipients.Where(Function(p) rule.AllowedRecipientIds.Contains(p.Id)).Select(Function(p) p.Name)))
                _sets.Rows(index).Tag = rule.Id
                _sets.Rows(index).Cells(4).Style.ForeColor = If(rule.Active, System.Drawing.Color.DarkGreen, System.Drawing.Color.DarkOrange)

            Next
            SelectRow(_recipients, recipientId)
            SelectRow(_sets, setId)
            _summary.Text = _host.RuntimeSummary.Invoke() & "  ·  " & _recipients.Rows.Count.ToString() & " recipients  ·  " & _sets.Rows.Count.ToString() & " recipient sets  ·  Revision " & _session.Configuration.Revision.ToString()
        End Sub

        Private Shared Sub SelectRow(grid As System.Windows.Forms.DataGridView, id As System.String)
            grid.ClearSelection()
            Dim selected = grid.Rows.Cast(Of System.Windows.Forms.DataGridViewRow)().FirstOrDefault(Function(row) CStr(row.Tag) = id)
            If selected Is Nothing AndAlso grid.Rows.Count > 0 Then selected = grid.Rows(0)
            If selected IsNot Nothing Then
                grid.CurrentCell = selected.Cells(0)
                selected.Selected = True
            End If
        End Sub

        Private Sub EditRecipient(create As System.Boolean)
            Dim id = AdminUi.SelectedId(_recipients)
            If Not create AndAlso id = "" Then Return
            Dim value = If(create, Nothing, _session.Configuration.Recipients.Single(Function(p) p.Id = id))
            Using editor As New AdminRecipientEditor(_session, value)
                If editor.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
                    RefreshConfigurationLists()
                    SetStatus("Recipient saved.")
                End If
            End Using
        End Sub

        Private Sub ExportRecipients()
            If _busy Then Return
            Using dialog As New System.Windows.Forms.SaveFileDialog With {.Title = "Export recipients", .Filter = "CSV recipients (*.csv)|*.csv", .DefaultExt = "csv", .AddExtension = True, .FileName = "Scan-Recipients.csv", .OverwritePrompt = True}
                If dialog.ShowDialog(Me) <> System.Windows.Forms.DialogResult.OK Then Return
                _session.ExportRecipientsFile(dialog.FileName)
                SetStatus("Recipients exported. To apply edits, import the CSV with Update existing recipients enabled.")
            End Using
        End Sub

        Private Sub ImportNames()
            Using editor As New AdminRecipientImport(_session)
                If editor.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
                    RefreshConfigurationLists()
                    SetStatus("Recipient import saved.")
                End If
            End Using
        End Sub

        Private Sub DeleteRecipient()
            Dim id = AdminUi.SelectedId(_recipients)
            If id = "" Then Return
            Try
                If Not AdminUi.Confirm("Remove the selected recipient?") Then Return
                _session.SaveChange(Sub() _session.DeleteRecipient(id))
                RefreshConfigurationLists()
                SetStatus("Recipient removed.")
            Catch ex As System.Exception
                AdminUi.ErrorMessage(Me, ex.Message)
            End Try
        End Sub

        Private Sub EditSet(create As System.Boolean)
            Dim id = AdminUi.SelectedId(_sets)
            If Not create AndAlso id = "" Then Return
            If Not _session.Configuration.Recipients.Any(Function(p) p.Active) Then
                AdminUi.ErrorMessage(Me, "Add at least one active recipient first.")
                _tabs.SelectedIndex = 0
                Return
            End If
            Using editor As New AdminRuleEditor(_session, If(create, Nothing, _session.Configuration.RuleSets.Single(Function(r) r.Id = id)), _host.GetMailboxes.Invoke())
                If editor.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
                    RefreshConfigurationLists()
                    SetStatus("Recipient set saved.")
                End If
            End Using
        End Sub

        Private Sub DuplicateSet()
            Dim id = AdminUi.SelectedId(_sets)
            If id = "" Then Return
            Dim value = AdminSession.Copy(_session.Configuration.RuleSets.Single(Function(r) r.Id = id))
            value.Id = System.Guid.NewGuid().ToString("N")
            value.Name &= " (copy)"
            value.Active = False
            Try
                _session.SaveChange(Sub() _session.PutRule(value))
                RefreshConfigurationLists()
                SetStatus("Disabled copy saved. Edit its scanner or mailbox before enabling.")
            Catch ex As System.Exception
                AdminUi.ErrorMessage(Me, ex.Message)
            End Try
        End Sub

        Private Sub ToggleSet()
            Dim id = AdminUi.SelectedId(_sets)
            If id = "" Then Return
            Try
                Dim value = AdminSession.Copy(_session.Configuration.RuleSets.Single(Function(r) r.Id = id))
                value.Active = Not value.Active
                _session.SaveChange(Sub() _session.PutRule(value))
                RefreshConfigurationLists()
                SetStatus("Recipient set status saved.")
            Catch ex As System.Exception
                AdminUi.ErrorMessage(Me, ex.Message)
            End Try
        End Sub

        Private Sub DeleteSet()
            Dim id = AdminUi.SelectedId(_sets)
            If id = "" Then Return
            If Not AdminUi.Confirm("Remove this recipient set? Existing packages are retained; pending deliveries may require review.") Then Return
            _session.SaveChange(Sub() _session.Configuration.RuleSets.RemoveAll(Function(r) r.Id = id))
            RefreshConfigurationLists()
            SetStatus("Recipient set removed.")
        End Sub

        Private Sub ExportConfiguration()
            If _busy Then Return
            Using dialog As New System.Windows.Forms.SaveFileDialog With {.Title = "Export Scan Handler configuration", .Filter = "JSON configuration (*.json)|*.json", .DefaultExt = "json", .AddExtension = True, .FileName = "RedInk-Scan-Configuration-" & System.DateTime.Now.ToString("yyyyMMdd-HHmm") & ".json", .OverwritePrompt = True}
                If dialog.ShowDialog(Me) <> System.Windows.Forms.DialogResult.OK Then Return
                _session.ExportFile(dialog.FileName)
                SetStatus("Configuration exported. The JSON file can be edited externally and imported, or kept as a backup. Scan documents and activity are not included.")
            End Using
        End Sub

        Private Sub ImportConfiguration()
            If _busy Then Return
            Using dialog As New System.Windows.Forms.OpenFileDialog With {.Title = "Import Scan Handler configuration or backup", .Filter = "JSON configuration (*.json;*.bak)|*.json;*.bak|All files (*.*)|*.*", .CheckFileExists = True}
                If dialog.ShowDialog(Me) <> System.Windows.Forms.DialogResult.OK Then Return
                If Not AdminUi.Confirm("Replace the recipient register and all recipient sets with this file? Existing scan history is retained. Invalid files will not be applied.", "Import") Then Return
                _session.ImportFile(dialog.FileName)
                BindConfiguration()
                SetStatus("Configuration imported and saved.")
            End Using
        End Sub

        Private Sub RestoreBackup()
            If _busy Then Return
            If Not AdminUi.Confirm("Restore the automatic backup from the last configuration save? This replaces the current recipients and recipient sets. Scan history is retained.", "Restore") Then Return
            _session.SaveChange(Sub() _session.LoadBackup())
            BindConfiguration()
            SetStatus("Previous configuration restored and saved.")
        End Sub

        Private Async Sub RefreshActivityAsync()
            Await RefreshActivityCoreAsync()
        End Sub

        Private Async Function RefreshActivityCoreAsync() As System.Threading.Tasks.Task
            If _refreshing OrElse IsDisposed OrElse _closing Then Return
            _refreshing = True
            Try
                Dim items = Await _host.ReadActivity.Invoke()
                If IsDisposed Then Return
                _items = items
                DisplayActivity()
                If _session.Configuration IsNot Nothing Then _summary.Text = _host.RuntimeSummary.Invoke() & "  ·  " & _session.Configuration.Recipients.Count.ToString() & " recipients  ·  " & _session.Configuration.RuleSets.Count.ToString() & " recipient sets"
            Catch ex As System.Exception
                If Not IsDisposed Then SetStatus("Activity unavailable: " & ex.Message, True)
            Finally
                _refreshing = False
            End Try
        End Function

        Private Sub DisplayActivity()
            Dim selected = AdminUi.SelectedId(_activity)
            Dim query = _search.Text.Trim()
            _activity.Rows.Clear()
            For Each item In _items.OrderByDescending(Function(a) a.Received)
                If query <> "" AndAlso (item.Id & " " & item.RuleSet & " " & item.Mailbox & " " & item.Issue).IndexOf(query, System.StringComparison.OrdinalIgnoreCase) < 0 Then Continue For
                Dim attention = item.Status = "failed" OrElse item.Status = "blocked" OrElse item.Status = "uncertain" OrElse item.Issue <> ""
                If _statusFilter.SelectedIndex = 1 AndAlso Not attention Then Continue For
                If _statusFilter.SelectedIndex = 2 AndAlso (item.Status = "complete" OrElse attention) Then Continue For
                If _statusFilter.SelectedIndex = 3 AndAlso item.Status <> "complete" Then Continue For
                Dim index = _activity.Rows.Add(item.Received.ToLocalTime().ToString("g"), item.RuleSet, item.Status, item.Deliveries, item.Issue)
                Dim row = _activity.Rows(index)
                row.Tag = item.Id
                row.Cells(2).Style.ForeColor = If(item.Status = "complete", System.Drawing.Color.DarkGreen, If(attention, System.Drawing.Color.DarkRed, System.Drawing.Color.DarkOrange))
                If item.Id = selected Then row.Selected = True
            Next
        End Sub

        Private Sub ShowSearchProgress()
            Using dialog As New AdminSearchProgressDialog(_host.ReadSearchProgress)
                dialog.ShowDialog(Me)
            End Using
        End Sub

        Private Sub ShowDetails()
            Dim id = AdminUi.SelectedId(_activity)
            Dim item = _items.SingleOrDefault(Function(a) a.Id = id)
            If item Is Nothing Then Return
            Using dialog As New AdminOperationDetails(item)
                dialog.ShowDialog(Me)
            End Using
        End Sub

        Private Sub SetStatus(message As System.String, Optional errorState As System.Boolean = False)
            _status.Text = message
            _status.ForeColor = If(errorState, System.Drawing.Color.DarkRed, System.Drawing.SystemColors.ControlText)
        End Sub

        Private Sub HandleClosing(sender As System.Object, args As System.Windows.Forms.FormClosingEventArgs)
            ' User-initiated close may be deferred while editing; Office shutdown must
            ' not leave this modeless Red Ink window (and its refresh timer) running.
            If _busy AndAlso Not Globals.ThisAddIn.IsHostModelessClosing Then
                args.Cancel = True
                SetStatus("Wait for the current operation to finish.")
                Return
            End If
            _closing = True
        End Sub
    End Class

    Friend Class AdminSearchProgressDialog
        Inherits System.Windows.Forms.Form
        Private ReadOnly _read As System.Func(Of System.Threading.Tasks.Task(Of System.Collections.Generic.List(Of AdminSearchProgress)))
        Private ReadOnly _grid As System.Windows.Forms.DataGridView = AdminUi.Grid("Mailbox", "Last completed search (local)", "Status")
        Private ReadOnly _refresh As System.Windows.Forms.Button
        Private _busy As System.Boolean
        Public Sub New(read As System.Func(Of System.Threading.Tasks.Task(Of System.Collections.Generic.List(Of AdminSearchProgress))))
            AdminUi.Style(Me, "Search progress", New System.Drawing.Size(800, 450))
            ShowInTaskbar = False
            _read = read
            Dim root = AdminUi.Layout()
            root.AutoSize = False
            root.RowCount = 2
            root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100))
            root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            _refresh = AdminUi.Button("Refresh", Sub() RefreshProgressAsync())
            Dim close = AdminUi.Button("Close", Sub() Me.Close())
            CancelButton = close
            root.Controls.Add(_grid, 0, 0)
            root.Controls.Add(AdminUi.Footer(close, _refresh), 0, 1)
            Controls.Add(root)
            AddHandler Shown, Sub() RefreshProgressAsync()
        End Sub

        Private Async Sub RefreshProgressAsync()
            If _busy Then Return
            _busy = True
            _refresh.Enabled = False
            Try
                Dim progress = Await _read.Invoke()
                If IsDisposed Then Return
                _grid.Rows.Clear()
                For Each value In progress
                    _grid.Rows.Add(value.Mailbox, If(value.ThroughUtc.HasValue, value.ThroughUtc.GetValueOrDefault().ToLocalTime().ToString("g"), ""), If(value.Issue = "", "Ready for next search", value.Issue))
                Next
            Catch ex As System.Exception
                If Not IsDisposed Then AdminUi.ErrorMessage(Me, ex.Message)
            Finally
                _busy = False
                If Not IsDisposed Then _refresh.Enabled = True
            End Try
        End Sub
    End Class

    Friend Class AdminOperationDetails
        Inherits System.Windows.Forms.Form
        Public Sub New(item As AdminActivity)
            AdminUi.Style(Me, "Package details", New System.Drawing.Size(900, 650))
            ShowInTaskbar = False
            Dim root = AdminUi.Layout()
            root.AutoSize = False
            root.RowCount = 2
            root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100))
            root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            Dim tabs As New System.Windows.Forms.TabControl With {.Dock = System.Windows.Forms.DockStyle.Fill}
            Dim fields = AdminUi.Fields(tabs, "Overview")
            Dim operation = item.Operation
            Detail(fields, "Status", item.Status)
            Detail(fields, "Received", item.Received.ToLocalTime().ToString("f"))
            Detail(fields, "Recipient set", item.RuleSet)
            Detail(fields, "Mailbox", item.Mailbox)
            Detail(fields, "Issue", item.Issue)
            If operation IsNot Nothing Then
                Detail(fields, "Subject", operation.Subject)
                If operation.Decision IsNot Nothing Then
                    Detail(fields, "Routing basis", System.String.Join(System.Environment.NewLine, operation.Decision.Reasons))
                    Detail(fields, "Review", System.String.Join(System.Environment.NewLine, operation.Decision.Review))
                End If
                Dim names = AdminUi.Grid("Original attachment", "Delivery name", "Note")
                For Each doc In operation.Documents
                    Dim name = operation.AttachmentNames.SingleOrDefault(Function(n) n.DocumentId = doc.Id)
                    names.Rows.Add(doc.FileName, If(name Is Nothing, doc.FileName, name.FileName), If(name Is Nothing, "", name.Note))
                Next
                AddGrid(tabs, "Attachments", names)
                Dim deliveries = AdminUi.Grid("Part", "Status", "Submitted (local)")
                For Each delivery In operation.Deliveries
                    deliveries.Rows.Add(delivery.Id, delivery.Status, If(delivery.SubmittedUtc.HasValue, delivery.SubmittedUtc.GetValueOrDefault().ToLocalTime().ToString("g"), ""))
                Next
                AddGrid(tabs, "Deliveries", deliveries)
                If operation.Decision IsNot Nothing Then
                    Dim dates = AdminUi.Grid("Date", "Event", "Details")
                    For Each value In operation.Decision.Events
                        dates.Rows.Add(value.DateValue.ToString("yyyy-MM-dd"), value.Title, value.Description)
                    Next
                    AddGrid(tabs, "Dates", dates)
                End If
                Dim technical As New System.Windows.Forms.TabPage("Technical details")
                technical.Controls.Add(New System.Windows.Forms.TextBox With {.Multiline = True, .ReadOnly = True, .ScrollBars = System.Windows.Forms.ScrollBars.Both, .Dock = System.Windows.Forms.DockStyle.Fill, .WordWrap = False,
                    .Text = Newtonsoft.Json.JsonConvert.SerializeObject(operation, Newtonsoft.Json.Formatting.Indented), .AccessibleName = "Read-only technical package record"})
                tabs.TabPages.Add(technical)
            End If
            root.Controls.Add(tabs, 0, 0)
            Dim close = AdminUi.Button("Close", Sub() Me.Close())
            CancelButton = close
            root.Controls.Add(AdminUi.Footer(close), 0, 1)
            Controls.Add(root)
        End Sub

        Private Shared Sub Detail(fields As System.Windows.Forms.TableLayoutPanel, name As System.String, value As System.String)
            AdminUi.Field(fields, name, New System.Windows.Forms.Label With {.Text = If(value, ""), .AutoSize = True, .UseMnemonic = False})
        End Sub

        Private Shared Sub AddGrid(tabs As System.Windows.Forms.TabControl, title As System.String, grid As System.Windows.Forms.DataGridView)
            Dim page As New System.Windows.Forms.TabPage(title)
            page.Controls.Add(grid)
            tabs.TabPages.Add(page)
        End Sub
    End Class
End Namespace
