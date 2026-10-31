' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SelectorGroupsEditorForm.vb
' PURPOSE
'   Edits and saves personal selector group layout metadata in a Windows Forms dialog.
'
' KEY ARCHITECTURE
'   • UI: Edits one-level group assignments in a Windows Forms grid.
'   • Persistence: Saves only personal presentation metadata, not tool permissions.
' =============================================================================

' Explicit layout editor. Saving never writes tool selections or central resources.
Option Strict On
Option Explicit On

Namespace SharedLibrary
    Public NotInheritable Class SelectorGroupsEditorForm
        Inherits System.Windows.Forms.Form

        Private ReadOnly _store As SelectorGroupingStore
        Private ReadOnly _items As System.Collections.Generic.List(Of SelectorGroupItem)
        Private _document As SelectorGroupDocument
        Private ReadOnly _grid As New System.Windows.Forms.DataGridView()
        Private ReadOnly _search As New System.Windows.Forms.TextBox()
        Private ReadOnly _groupName As New System.Windows.Forms.TextBox()
        Private ReadOnly _source As New System.Windows.Forms.Label()
        Private ReadOnly _save As New System.Windows.Forms.Button()
        Private _ownedIcon As System.Drawing.Icon

        Public Sub New(store As SelectorGroupingStore, items As System.Collections.Generic.IEnumerable(Of SelectorGroupItem))
            _store = store
            _items = New System.Collections.Generic.List(Of SelectorGroupItem)(items)
            _document = store.Document.Copy()
            Me.Text = SharedMethods.AN & " - Selector groups"
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
            Me.MinimizeBox = False
            Me.MaximizeBox = True
            Me.ShowInTaskbar = False
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
            Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0F, 96.0F)
            Me.ClientSize = New System.Drawing.Size(950, 590)
            Me.MinimumSize = New System.Drawing.Size(760, 500)
            Me.ShowIcon = True
            InitializeBranding()

            Dim outer As New System.Windows.Forms.TableLayoutPanel() With {
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .Padding = New System.Windows.Forms.Padding(12),
                .ColumnCount = 1,
                .RowCount = 6
            }
            outer.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
            For i As System.Int32 = 0 To 5
                outer.RowStyles.Add(New System.Windows.Forms.RowStyle(If(i = 3, System.Windows.Forms.SizeType.Percent, System.Windows.Forms.SizeType.AutoSize), If(i = 3, 100.0F, 0.0F)))
            Next
            Dim explanation As New System.Windows.Forms.Label() With {
                .Text = "Type a group name beside an entry, or select several rows and assign a name below. Leave the name empty for unassigned entries. This changes presentation only, never which tools are enabled." & System.Environment.NewLine &
                        "Save personal replaces your local layout; Export creates a file you can share or place centrally. Explicit layout saves are independent of OK/Cancel in the tool selector.",
                .AutoSize = True,
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .Margin = New System.Windows.Forms.Padding(0, 0, 0, 8)
            }
            _source.AutoSize = True
            _source.Dock = System.Windows.Forms.DockStyle.Fill
            _source.Margin = New System.Windows.Forms.Padding(0, 0, 0, 8)
            UpdateSourceLabel()
            _search.Dock = System.Windows.Forms.DockStyle.Fill
            _search.AccessibleName = "Filter entries"
            AddHandler _search.TextChanged, AddressOf FilterRows
            Dim searchRow As New System.Windows.Forms.TableLayoutPanel() With {.AutoSize = True, .Dock = System.Windows.Forms.DockStyle.Fill, .ColumnCount = 2}
            searchRow.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize))
            searchRow.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
            searchRow.Controls.Add(New System.Windows.Forms.Label() With {.Text = "Filter:", .AutoSize = True, .Anchor = System.Windows.Forms.AnchorStyles.Left}, 0, 0)
            searchRow.Controls.Add(_search, 1, 0)

            _grid.Dock = System.Windows.Forms.DockStyle.Fill
            _grid.AllowUserToAddRows = False
            _grid.AllowUserToDeleteRows = False
            _grid.AllowUserToResizeRows = False
            _grid.RowHeadersVisible = False
            _grid.MultiSelect = True
            _grid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
            _grid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
            _grid.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText
            _grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            _grid.ColumnHeadersDefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
            _grid.ColumnHeadersDefaultCellStyle.WrapMode = System.Windows.Forms.DataGridViewTriState.True
            UpdateHeaderPadding()
            _grid.Columns.Add(New System.Windows.Forms.DataGridViewTextBoxColumn() With {.Name = "entryLabel", .HeaderText = "Entry", .ReadOnly = True, .FillWeight = 48.0F, .SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable})
            _grid.Columns.Add(New System.Windows.Forms.DataGridViewTextBoxColumn() With {.Name = "entryId", .HeaderText = "Stable ID", .ReadOnly = True, .FillWeight = 26.0F, .SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable})
            _grid.Columns.Add(New System.Windows.Forms.DataGridViewTextBoxColumn() With {.Name = "groupName", .HeaderText = "Group (editable)", .MaxInputLength = 100, .FillWeight = 26.0F, .SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable})
            FillRows()

            ' All controls occupy the same auto-sized row. Horizontal-only anchors
            ' centre their different natural heights, including after DPI/font changes.
            Dim assignRow As New System.Windows.Forms.TableLayoutPanel() With {
                .AutoSize = True,
                .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .ColumnCount = 4,
                .RowCount = 1
            }
            assignRow.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize))
            assignRow.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
            assignRow.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize))
            assignRow.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize))
            assignRow.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            assignRow.Controls.Add(New System.Windows.Forms.Label() With {
                .Text = "Group for selected rows:",
                .AutoSize = True,
                .Anchor = System.Windows.Forms.AnchorStyles.Left,
                .Margin = New System.Windows.Forms.Padding(0, 3, 6, 3)
            }, 0, 0)
            _groupName.Width = 220
            _groupName.MaxLength = 100
            _groupName.AccessibleName = "Group for selected entries; empty means unassigned"
            _groupName.Anchor = System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right
            _groupName.Margin = New System.Windows.Forms.Padding(0, 3, 6, 3)
            assignRow.Controls.Add(_groupName, 1, 0)
            Dim assignButton As System.Windows.Forms.Button = CreateButton("Assign selected", AddressOf AssignSelected)
            assignButton.Anchor = System.Windows.Forms.AnchorStyles.Left
            assignButton.Margin = New System.Windows.Forms.Padding(0, 3, 6, 3)
            assignRow.Controls.Add(assignButton, 2, 0)
            Dim clearButton As System.Windows.Forms.Button = CreateButton("Clear selected", AddressOf ClearSelected)
            clearButton.Anchor = System.Windows.Forms.AnchorStyles.Left
            clearButton.Margin = New System.Windows.Forms.Padding(0, 3, 0, 3)
            assignRow.Controls.Add(clearButton, 3, 0)

            Dim buttons As New System.Windows.Forms.FlowLayoutPanel() With {
                .AutoSize = True,
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft,
                .WrapContents = True
            }
            Dim cancel As System.Windows.Forms.Button = CreateButton("Cancel", Nothing)
            cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
            _save.Text = "Save personal"
            _save.AutoSize = True
            _save.Padding = New System.Windows.Forms.Padding(8, 4, 8, 4)
            _save.Enabled = _store.CanSaveLocal
            AddHandler _save.Click, AddressOf SavePersonal
            buttons.Controls.Add(cancel)
            buttons.Controls.Add(_save)
            buttons.Controls.Add(CreateButton("Export…", AddressOf ExportLayout))
            buttons.Controls.Add(CreateButton("Import…", AddressOf ImportLayout))
            Dim reset As System.Windows.Forms.Button = CreateButton("Use shared / flat", AddressOf ResetLayout)
            reset.Enabled = _store.CanSaveLocal
            buttons.Controls.Add(reset)
            outer.Controls.Add(explanation, 0, 0)
            outer.Controls.Add(_source, 0, 1)
            outer.Controls.Add(searchRow, 0, 2)
            outer.Controls.Add(_grid, 0, 3)
            outer.Controls.Add(assignRow, 0, 4)
            outer.Controls.Add(buttons, 0, 5)
            Me.Controls.Add(outer)
            Me.CancelButton = cancel
            ' Enter in a grid cell must commit the edit, not accidentally save the file.
        End Sub

        Private Sub InitializeBranding()
            Try
                ' This existing helper returns a caller-owned clone of the standard
                ' Red Ink logo and releases the temporary native HICON itself.
                _ownedIcon = SharedMethods.CreateMarkdownEditorIcon()
                Me.Icon = _ownedIcon
            Catch ex As System.Exception
                SelectorGroupingStore.Log("editor-branding-fallback", _store.SourceFile, ex)
            End Try
        End Sub

        Private Sub UpdateHeaderPadding()
            ' Header height is content-driven, not a hard-coded pixel height. Use
            ' logical spacing for high-DPI displays; auto-sizing also fits wrapping.
            Dim verticalPadding As System.Int32 = System.Math.Max(1, CInt(System.Math.Ceiling(6.0R * Me.DeviceDpi / 96.0R)))
            _grid.ColumnHeadersDefaultCellStyle.Padding = New System.Windows.Forms.Padding(0, verticalPadding, 0, verticalPadding)
        End Sub

        Protected Overrides Sub OnLoad(args As System.EventArgs)
            MyBase.OnLoad(args)
            UpdateHeaderPadding()
        End Sub

        Protected Overrides Sub OnDpiChanged(args As System.Windows.Forms.DpiChangedEventArgs)
            MyBase.OnDpiChanged(args)
            UpdateHeaderPadding()
        End Sub

        Protected Overrides Sub Dispose(disposing As System.Boolean)
            Try
                If disposing AndAlso _ownedIcon IsNot Nothing Then
                    Me.Icon = Nothing
                    _ownedIcon.Dispose()
                    _ownedIcon = Nothing
                End If
            Finally
                MyBase.Dispose(disposing)
            End Try
        End Sub

        Protected Overrides Sub OnShown(args As System.EventArgs)
            MyBase.OnShown(args)
            SharedMethods.ForceDialogToForeground(Me)
            SharedMethods.AttachForeignForegroundWatchdog(Me)
        End Sub

        Private Shared Function CreateButton(text As System.String, handler As System.EventHandler) As System.Windows.Forms.Button
            Dim button As New System.Windows.Forms.Button() With {.Text = text, .AutoSize = True, .Padding = New System.Windows.Forms.Padding(8, 4, 8, 4)}
            If handler IsNot Nothing Then AddHandler button.Click, handler
            Return button
        End Function

        Private Sub UpdateSourceLabel()
            _source.Text = "Current: " & If(_store.SourceFile.Length = 0, "no group file (flat list)", _store.SourceFile) & System.Environment.NewLine &
                           "Personal: " & If(_store.LocalFile.Length = 0, "not configured", _store.LocalFile) & System.Environment.NewLine &
                           "Shared: " & If(_store.CentralFile.Length = 0, "not configured", _store.CentralFile)
            If _store.Warning.Length > 0 Then _source.Text &= System.Environment.NewLine & _store.Warning
            If Not _store.CanSaveLocal Then _source.Text &= System.Environment.NewLine & "Configure a separate local resource folder, or export the layout. Shared files are never changed here."
        End Sub

        Private Sub FillRows()
            _grid.Rows.Clear()
            For Each item As SelectorGroupItem In _items
                Dim index As System.Int32 = _grid.Rows.Add(item.Label, item.Id, _document.GroupFor(item.Id))
                _grid.Rows(index).Tag = item.Id
            Next
            FilterRows(Me, System.EventArgs.Empty)
        End Sub

        Private Sub FilterRows(sender As System.Object, args As System.EventArgs)
            If Not _grid.EndEdit() Then Return
            _grid.CurrentCell = Nothing
            Dim search As System.String = _search.Text.Trim()
            For Each row As System.Windows.Forms.DataGridViewRow In _grid.Rows
                Dim matches As System.Boolean = search.Length = 0
                For Each cell As System.Windows.Forms.DataGridViewCell In row.Cells
                    If System.Convert.ToString(cell.Value, System.Globalization.CultureInfo.InvariantCulture).IndexOf(search, System.StringComparison.OrdinalIgnoreCase) >= 0 Then matches = True
                Next
                row.Visible = matches
                ' Bulk assignment cannot modify hidden selected rows after a filter change.
                If Not matches Then row.Selected = False
            Next
        End Sub

        Private Sub AssignSelected(sender As System.Object, args As System.EventArgs)
            SetSelectedGroup(_groupName.Text.Trim())
        End Sub

        Private Sub ClearSelected(sender As System.Object, args As System.EventArgs)
            SetSelectedGroup("")
        End Sub

        Private Sub SetSelectedGroup(name As System.String)
            If Not _grid.EndEdit() Then Return
            For Each row As System.Windows.Forms.DataGridViewRow In _grid.SelectedRows
                If row.Visible Then row.Cells("groupName").Value = name
            Next
        End Sub

        Private Function EditedDocument() As SelectorGroupDocument
            If Not _grid.EndEdit() Then Throw New System.FormatException("Finish editing the current cell first.")
            Dim assignments As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)
            For Each row As System.Windows.Forms.DataGridViewRow In _grid.Rows
                assignments(DirectCast(row.Tag, System.String)) = System.Convert.ToString(row.Cells("groupName").Value, System.Globalization.CultureInfo.InvariantCulture).Trim()
            Next
            Return _document.WithAssignments(assignments)
        End Function

        Private Sub SavePersonal(sender As System.Object, args As System.EventArgs)
            Try
                Dim result As SelectorGroupDocument = EditedDocument()
                If _store.Warning.Length > 0 Then
                    Using SharedMethods.PushDialogOwner(Me)
                        If SharedMethods.ShowCustomYesNoBox("A group file could not be loaded. Save this layout as your personal version? The previous local file, if readable, is retained as a backup.", "Save personal", "Cancel", header:=Me.Text) <> 1 Then Return
                    End Using
                End If
                _store.SaveLocal(result)
                Me.DialogResult = System.Windows.Forms.DialogResult.OK
                Me.Close()
            Catch ex As System.Exception
                ReportFailure("The layout could not be saved. No tool selections were changed.", ex)
            End Try
        End Sub

        Private Sub ImportLayout(sender As System.Object, args As System.EventArgs)
            Try
                Using dialog As New System.Windows.Forms.OpenFileDialog()
                    dialog.Filter = "Selector group files (*.json)|*.json"
                    dialog.CheckFileExists = True
                    dialog.Multiselect = False
                    dialog.FileName = SelectorGroupingStore.FileName
                    Dim owner As System.Windows.Forms.IWin32Window = SharedMethods.IfOwnerOnCurrentThread(Me)
                    If If(owner IsNot Nothing, dialog.ShowDialog(owner), dialog.ShowDialog()) <> System.Windows.Forms.DialogResult.OK Then Return
                    Dim imported As SelectorGroupDocument = SelectorGroupingStore.ReadDocument(dialog.FileName)
                    Using SharedMethods.PushDialogOwner(Me)
                        If SharedMethods.ShowCustomYesNoBox("Replace the layout being edited with this file? Nothing is applied until you choose Save personal.", "Import", "Cancel", header:=Me.Text) <> 1 Then Return
                    End Using
                    _document = imported
                    _search.Clear()
                    FillRows()
                End Using
            Catch ex As System.Exception
                ReportFailure("The group file could not be imported. The current layout has not been saved or replaced on disk.", ex)
            End Try
        End Sub

        Private Sub ExportLayout(sender As System.Object, args As System.EventArgs)
            Try
                Dim result As SelectorGroupDocument = EditedDocument()
                Using dialog As New System.Windows.Forms.SaveFileDialog()
                    dialog.Filter = "Selector group files (*.json)|*.json"
                    dialog.DefaultExt = "json"
                    dialog.AddExtension = True
                    dialog.OverwritePrompt = True
                    dialog.FileName = SelectorGroupingStore.FileName
                    Dim owner As System.Windows.Forms.IWin32Window = SharedMethods.IfOwnerOnCurrentThread(Me)
                    If If(owner IsNot Nothing, dialog.ShowDialog(owner), dialog.ShowDialog()) <> System.Windows.Forms.DialogResult.OK Then Return
                    If SelectorGroupingStore.SamePath(dialog.FileName, _store.LocalFile) OrElse SelectorGroupingStore.SamePath(dialog.FileName, _store.CentralFile) Then
                        Inform("Choose a different export location. Use Save personal for your own active layout; copy an exported file to the shared folder to publish it centrally.")
                        Return
                    End If
                    SelectorGroupingStore.ExportDocument(dialog.FileName, result)
                    Inform("The layout was exported. To share it, use Import / Save personal on another computer, or place selector-groups.json in the shared resource folder. Your current tool selections were not changed.")
                End Using
            Catch ex As System.Exception
                ReportFailure("The group file could not be exported.", ex)
            End Try
        End Sub

        Private Sub ResetLayout(sender As System.Object, args As System.EventArgs)
            Using SharedMethods.PushDialogOwner(Me)
                If SharedMethods.ShowCustomYesNoBox("Stop using your personal layout and return to the shared file (or the flat list when none exists)? Your personal file is kept as a backup. Unsaved layout edits will be discarded; tool selections are unchanged.", "Use shared / flat", "Cancel", header:=Me.Text) <> 1 Then Return
            End Using
            Try
                _store.UseCentralOrFlat()
                Me.DialogResult = System.Windows.Forms.DialogResult.OK
                Me.Close()
            Catch ex As System.Exception
                ReportFailure("The personal layout could not be reset.", ex)
            End Try
        End Sub

        Private Sub Inform(message As System.String)
            Using SharedMethods.PushDialogOwner(Me)
                SharedMethods.ShowCustomMessageBox(message, header:=Me.Text)
            End Using
        End Sub

        Private Sub ReportFailure(message As System.String, failure As System.Exception)
            SelectorGroupingStore.Log("editor-failure", _store.SourceFile, failure)
            ' Only our validation/conflict messages are shown, not exception types/stacks.
            If TypeOf failure Is System.FormatException OrElse TypeOf failure Is System.InvalidOperationException Then
                message &= System.Environment.NewLine & failure.Message
            ElseIf TypeOf failure Is Newtonsoft.Json.JsonException Then
                message &= System.Environment.NewLine & "Check the JSON format and schema_version in the selected file."
            Else
                message &= System.Environment.NewLine & "Check folder access and whether the file changed in another window, then reopen Groups and try again."
            End If
            Inform(message)
        End Sub
    End Class
End Namespace
