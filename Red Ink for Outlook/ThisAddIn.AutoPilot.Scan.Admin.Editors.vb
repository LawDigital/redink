' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.Scan.Admin.Editors.vb
' Purpose: Accessible, layout-based administration editors.
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
    Friend NotInheritable Class AdminUi

        Public Shared Sub Style(form As System.Windows.Forms.Form, title As System.String, size As System.Drawing.Size)
            form.Text = "Red Ink" & If(title = "AutoPilot Scan Handler", " " & title, " AutoPilot Scan Handler — " & title)
            form.Font = New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
            form.AutoScaleDimensions = New System.Drawing.SizeF(7.0F, 15.0F)
            form.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            form.ClientSize = size
            form.MinimumSize = New System.Drawing.Size(640, 480)
            form.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
            form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
            form.TopMost = False
            form.MaximizeBox = True
            form.BackColor = System.Drawing.SystemColors.Control
            ' Same icon source as CreateSchedulerDashboardForm; no in-form branding.
            Try
                form.Icon = System.Drawing.Icon.FromHandle((New System.Drawing.Bitmap(GetLogoBitmap(LogoType.Standard))).GetHicon())
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine(ex.Message)
            End Try
            AddHandler form.Shown,
                Sub()
                    Dim area = System.Windows.Forms.Screen.FromControl(form).WorkingArea
                    form.MinimumSize = New System.Drawing.Size(System.Math.Min(form.MinimumSize.Width, area.Width), System.Math.Min(form.MinimumSize.Height, area.Height))
                    form.Size = New System.Drawing.Size(System.Math.Min(form.Width, area.Width), System.Math.Min(form.Height, area.Height))
                    form.Location = New System.Drawing.Point(System.Math.Max(area.Left, System.Math.Min(form.Left, area.Right - form.Width)), System.Math.Max(area.Top, System.Math.Min(form.Top, area.Bottom - form.Height)))
                End Sub
        End Sub

        Public Shared Function Layout(Optional columns As System.Int32 = 1) As System.Windows.Forms.TableLayoutPanel
            Dim table As New System.Windows.Forms.TableLayoutPanel With {.Dock = System.Windows.Forms.DockStyle.Fill, .ColumnCount = columns, .RowCount = 0, .Padding = New System.Windows.Forms.Padding(8), .AutoSize = True, .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink}
            For index As System.Int32 = 0 To columns - 1
                table.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, CSng(100.0 / columns)))
            Next
            Return table
        End Function

        Public Shared Function Fields(tabs As System.Windows.Forms.TabControl, title As System.String) As System.Windows.Forms.TableLayoutPanel
            Dim page As New System.Windows.Forms.TabPage(title) With {.AutoScroll = True, .UseVisualStyleBackColor = True}
            Dim result = Layout(2)
            result.Dock = System.Windows.Forms.DockStyle.Top
            result.GrowStyle = System.Windows.Forms.TableLayoutPanelGrowStyle.AddRows
            result.ColumnStyles.Clear()
            result.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 30))
            result.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 70))
            page.Controls.Add(result)
            tabs.TabPages.Add(page)
            Return result
        End Function

        Public Shared Sub Field(table As System.Windows.Forms.TableLayoutPanel, text As System.String, control As System.Windows.Forms.Control)
            Dim row = table.RowCount
            table.RowCount += 1
            table.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            Dim label As New System.Windows.Forms.Label With {.Text = text, .AutoSize = True, .Anchor = System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Top, .Margin = New System.Windows.Forms.Padding(3, 9, 12, 8), .UseMnemonic = False}
            control.Anchor = System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right Or System.Windows.Forms.AnchorStyles.Top
            control.Margin = New System.Windows.Forms.Padding(3, 5, 3, 5)
            control.AccessibleName = text
            table.Controls.Add(label, 0, row)
            table.Controls.Add(control, 1, row)
            AddHandler table.SizeChanged,
                Sub()
                    Dim available = System.Math.Max(180, table.ClientSize.Width - table.Padding.Horizontal)
                    label.MaximumSize = New System.Drawing.Size(System.Math.Max(70, CInt(available * 0.3) - label.Margin.Horizontal), 0)
                    If TypeOf control Is System.Windows.Forms.Label OrElse TypeOf control Is System.Windows.Forms.CheckBox Then
                        control.MaximumSize = New System.Drawing.Size(System.Math.Max(100, CInt(available * 0.7) - control.Margin.Horizontal), 0)
                    End If
                End Sub
        End Sub

        Public Shared Sub Help(control As System.Windows.Forms.Control, explanation As System.String)
            Dim tip As New System.Windows.Forms.ToolTip With {.AutoPopDelay = 20000, .ShowAlways = True}
            tip.SetToolTip(control, explanation)
            AddHandler control.Disposed, Sub() tip.Dispose()
        End Sub

        Public Shared Function Text(value As System.String, Optional multiline As System.Boolean = False) As System.Windows.Forms.TextBox
            Dim input As New System.Windows.Forms.TextBox With {.Text = If(value, ""), .Dock = System.Windows.Forms.DockStyle.Fill}
            If multiline Then
                input.Multiline = True
                input.AcceptsReturn = True
                input.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
                input.MinimumSize = New System.Drawing.Size(180, 90)
            End If
            Return input
        End Function

        Public Shared Function Check(text As System.String, value As System.Boolean) As System.Windows.Forms.CheckBox
            Return New System.Windows.Forms.CheckBox With {.Text = text, .Checked = value, .AutoSize = True, .UseMnemonic = False}
        End Function

        Public Shared Function Number(value As System.Int32, minimum As System.Int32, maximum As System.Int32) As System.Windows.Forms.NumericUpDown
            Return New System.Windows.Forms.NumericUpDown With {.Minimum = minimum, .Maximum = maximum, .Value = System.Math.Max(minimum, System.Math.Min(maximum, value)), .ThousandsSeparator = True}
        End Function

        Public Shared Function Button(text As System.String, action As System.Action) As System.Windows.Forms.Button
            Dim result As New System.Windows.Forms.Button With {.Text = text, .AutoSize = True, .Padding = New System.Windows.Forms.Padding(10, 5, 10, 5), .UseVisualStyleBackColor = True}
            AddHandler result.Click,
                Sub()
                    Try
                        action()
                    Catch ex As System.Exception
                        ErrorMessage(result.FindForm(), ex.Message)
                    End Try
                End Sub
            Return result
        End Function

        Public Shared Function Buttons(ParamArray controls As System.Windows.Forms.Control()) As System.Windows.Forms.FlowLayoutPanel
            Dim panel As New System.Windows.Forms.FlowLayoutPanel With {.AutoSize = True, .Dock = System.Windows.Forms.DockStyle.Fill, .WrapContents = True, .Margin = New System.Windows.Forms.Padding(0)}
            panel.Controls.AddRange(controls)
            Return panel
        End Function

        Public Shared Function Grid(ParamArray headings As System.String()) As System.Windows.Forms.DataGridView
            Dim result As New System.Windows.Forms.DataGridView With {.Dock = System.Windows.Forms.DockStyle.Fill, .ReadOnly = True, .AllowUserToAddRows = False, .AllowUserToDeleteRows = False,
                .AllowUserToResizeRows = False, .MultiSelect = False, .RowHeadersVisible = False, .SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect,
                .AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill, .BackgroundColor = System.Drawing.SystemColors.Window, .BorderStyle = System.Windows.Forms.BorderStyle.None,
                .ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize, .AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells,
                .EnableHeadersVisualStyles = True}
            result.DefaultCellStyle.Font = New System.Drawing.Font("Segoe UI", 8.5F)
            result.DefaultCellStyle.WrapMode = System.Windows.Forms.DataGridViewTriState.False
            result.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(245, 245, 250)
            For Each heading In headings
                Dim index = result.Columns.Add("c" & result.Columns.Count.ToString(), heading)
                result.Columns(index).MinimumWidth = 70
                result.Columns(index).SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
            Next
            Return result
        End Function

        Public Shared Function SelectedId(grid As System.Windows.Forms.DataGridView) As System.String
            If grid.SelectedRows.Count = 0 Then Return ""
            Return TryCast(grid.SelectedRows(0).Tag, System.String)
        End Function

        Public Shared Function Confirm(message As System.String, Optional accept As System.String = "Remove") As System.Boolean
            Return ShowCustomYesNoBox(message, accept, "Cancel", header:="Red Ink" & " AutoPilot Scan Handler") = 1
        End Function

        Public Shared Function Footer(ParamArray controls As System.Windows.Forms.Control()) As System.Windows.Forms.FlowLayoutPanel
            Dim panel = Buttons(controls)
            panel.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft
            panel.Padding = New System.Windows.Forms.Padding(0, 4, 0, 0)
            Return panel
        End Function

        Public Shared Sub ErrorMessage(owner As System.Windows.Forms.Form, message As System.String)
            ShowCustomMessageBox(message, header:="Red Ink" & " AutoPilot Scan Handler")
        End Sub
    End Class

    Friend Class AdminEditor
        Inherits System.Windows.Forms.Form
        Protected ReadOnly Tabs As New System.Windows.Forms.TabControl With {.Dock = System.Windows.Forms.DockStyle.Fill, .Multiline = True}
        Private ReadOnly _error As New System.Windows.Forms.Label With {.AutoSize = True, .ForeColor = System.Drawing.Color.DarkRed, .Dock = System.Windows.Forms.DockStyle.Fill, .UseMnemonic = False}
        Protected ReadOnly ApplyButton As System.Windows.Forms.Button

        Public Sub New(title As System.String)
            AdminUi.Style(Me, title, New System.Drawing.Size(760, 660))
            ShowInTaskbar = False
            MinimizeBox = False
            Dim root = AdminUi.Layout()
            root.AutoSize = False
            root.RowCount = 3
            root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100))
            root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            ApplyButton = AdminUi.Button("OK", AddressOf ApplyChanges)
            Dim cancel = AdminUi.Button("Cancel", Sub() Close())
            cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
            CancelButton = cancel
            ' Enter in a list entry must not accidentally accept the entire editor.
            Dim buttons = AdminUi.Footer(cancel, ApplyButton)
            buttons.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft
            root.Controls.Add(Tabs, 0, 0)
            root.Controls.Add(_error, 0, 1)
            root.Controls.Add(buttons, 0, 2)
            Controls.Add(root)
            AddHandler root.SizeChanged, Sub() _error.MaximumSize = New System.Drawing.Size(System.Math.Max(180, root.ClientSize.Width - root.Padding.Horizontal - 12), 0)
        End Sub

        Protected Overridable Sub Commit()
        End Sub

        Private Sub ApplyChanges()
            Try
                Commit()
                DialogResult = System.Windows.Forms.DialogResult.OK
                Close()
            Catch ex As System.Exception
                _error.Text = ex.Message
                AdminUi.ErrorMessage(Me, ex.Message)
            End Try
        End Sub
    End Class

    Friend Class AdminTextList
        Inherits System.Windows.Forms.UserControl
        Private ReadOnly _list As New System.Windows.Forms.ListBox With {.Dock = System.Windows.Forms.DockStyle.Fill, .IntegralHeight = False, .MinimumSize = New System.Drawing.Size(140, 90)}
        Private ReadOnly _input As System.Windows.Forms.TextBox = AdminUi.Text("")
        Private ReadOnly _positiveNumbers As System.Boolean
        Public ReadOnly Property Values As System.Collections.Generic.List(Of System.String)
            Get
                Return _list.Items.Cast(Of System.String)().ToList()
            End Get
        End Property

        Public Sub New(values As System.Collections.Generic.IEnumerable(Of System.String), Optional positiveNumbers As System.Boolean = False)
            _positiveNumbers = positiveNumbers
            AutoSize = True
            AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
            MinimumSize = New System.Drawing.Size(200, 0)
            Dim root = AdminUi.Layout()
            root.Dock = System.Windows.Forms.DockStyle.Top
            root.Padding = New System.Windows.Forms.Padding(0)
            root.RowCount = 3
            For index As System.Int32 = 0 To 2
                root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            Next
            _list.Items.AddRange(values.Cast(Of System.Object)().ToArray())
            root.Controls.Add(_list, 0, 0)
            root.Controls.Add(_input, 0, 1)
            Dim add = AdminUi.Button("Add", AddressOf AddValue)
            Dim remove = AdminUi.Button("Remove selected", AddressOf RemoveSelected)
            add.Enabled = False
            remove.Enabled = False
            AddHandler _input.TextChanged, Sub() add.Enabled = Not System.String.IsNullOrWhiteSpace(_input.Text)
            AddHandler _list.SelectedIndexChanged, Sub() remove.Enabled = _list.SelectedIndex >= 0
            root.Controls.Add(AdminUi.Buttons(add, remove), 0, 2)
            Controls.Add(root)
            AddHandler _input.KeyDown, Sub(sender, args)
                                           If args.KeyCode = System.Windows.Forms.Keys.Enter Then
                                               AddValue()
                                               args.SuppressKeyPress = True
                                           End If
                                       End Sub
        End Sub

        Private Sub RemoveSelected()
            Dim index = _list.SelectedIndex
            If index < 0 Then Return
            _list.Items.RemoveAt(index)
            If _list.Items.Count > 0 Then _list.SelectedIndex = System.Math.Min(index, _list.Items.Count - 1)
        End Sub

        Private Sub AddValue()
            Dim value = _input.Text.Trim()
            If value = "" Then Return
            If _positiveNumbers Then
                Dim number As System.Int32
                If Not System.Int32.TryParse(value, number) OrElse number < 1 OrElse number > 3650 Then
                    AdminUi.ErrorMessage(FindForm(), "Enter a whole number of days from 1 to 3,650.")
                    Return
                End If
                value = number.ToString(System.Globalization.CultureInfo.InvariantCulture)
            End If
            Dim existing = Values.FindIndex(Function(item) item.Equals(value, System.StringComparison.OrdinalIgnoreCase))
            _list.SelectedIndex = If(existing < 0, _list.Items.Add(value), existing)
            _input.Clear()
            _input.Focus()
        End Sub

        Public Sub FlushPending()
            If _input.Text.Trim() <> "" Then
                Dim text = _input.Text.Trim()
                AddValue()
                If _input.Text.Trim() <> "" Then Throw New System.IO.InvalidDataException("Finish or clear the pending list entry: " & text)
            End If
        End Sub
    End Class

    Friend Class AdminRecipientPicker
        Inherits System.Windows.Forms.CheckedListBox
        Private ReadOnly _people As System.Collections.Generic.List(Of RegisteredRecipient)
        Public Sub New(people As System.Collections.Generic.IEnumerable(Of RegisteredRecipient), selected As System.Collections.Generic.IEnumerable(Of System.String))
            _people = people.Where(Function(p) p.Active).OrderBy(Function(p) p.Name).ToList()
            CheckOnClick = True
            IntegralHeight = False
            MinimumSize = New System.Drawing.Size(220, 160)
            AccessibleName = "Recipients"
            For Each person In _people
                Items.Add(person.Name & " <" & person.SmtpAddress & ">", selected.Contains(person.Id))
            Next
        End Sub
        Public ReadOnly Property SelectedIds As System.Collections.Generic.List(Of System.String)
            Get
                Return CheckedIndices.Cast(Of System.Int32)().Select(Function(i) _people(i).Id).ToList()
            End Get
        End Property
    End Class

    Friend Class AdminRecipientEditor
        Inherits AdminEditor
        Private ReadOnly _session As AdminSession
        Private ReadOnly _value As RegisteredRecipient
        Private ReadOnly _name As System.Windows.Forms.TextBox
        Private ReadOnly _email As System.Windows.Forms.TextBox
        Private ReadOnly _active As System.Windows.Forms.CheckBox
        Private ReadOnly _aliases As AdminTextList
        Private ReadOnly _organization As System.Windows.Forms.TextBox
        Private ReadOnly _location As System.Windows.Forms.TextBox
        Private ReadOnly _attributes As AdminTextList
        Private ReadOnly _instruction As System.Windows.Forms.TextBox

        Public Sub New(session As AdminSession, person As RegisteredRecipient)
            MyBase.New(If(person Is Nothing, "Add recipient", "Edit recipient"))
            _session = session
            ApplyButton.Text = "Save"
            _value = If(person Is Nothing, New RegisteredRecipient(), AdminSession.Copy(person))
            Dim fields = AdminUi.Fields(Tabs, "Contact")
            _name = AdminUi.Text(_value.Name)
            _email = AdminUi.Text(_value.SmtpAddress)
            _active = AdminUi.Check("Available for recipient sets", _value.Active)
            _organization = AdminUi.Text(_value.Organization)
            _location = AdminUi.Text(_value.Location)
            _aliases = New AdminTextList(_value.Aliases)
            AdminUi.Field(fields, "Full name *", _name)
            AdminUi.Field(fields, "Email address *", _email)
            AdminUi.Field(fields, "Status", _active)
            AdminUi.Field(fields, "Organization", _organization)
            AdminUi.Field(fields, "Location", _location)
            AdminUi.Field(fields, "Alternative names", _aliases)
            fields = AdminUi.Fields(Tabs, "Matching & instructions")
            _attributes = New AdminTextList(_value.Attributes)
            _instruction = AdminUi.Text(_value.Instruction, True)
            AdminUi.Field(fields, "Matching attributes", _attributes)
            AdminUi.Field(fields, "Instructions for this recipient", _instruction)
        End Sub

        Protected Overrides Sub Commit()
            _aliases.FlushPending()
            _attributes.FlushPending()
            _value.Name = _name.Text
            _value.SmtpAddress = _email.Text
            _value.Active = _active.Checked
            _value.Organization = _organization.Text.Trim()
            _value.Location = _location.Text.Trim()
            _value.Aliases = _aliases.Values
            _value.Attributes = _attributes.Values
            _value.Instruction = _instruction.Text
            _session.SaveChange(Sub() _session.PutRecipient(_value))
        End Sub
    End Class

    Friend Class AdminExtraRuleEditor
        Inherits AdminEditor
        Private ReadOnly _value As AdditionalRule
        Private ReadOnly _active As System.Windows.Forms.CheckBox
        Private ReadOnly _kind As New System.Windows.Forms.ComboBox With {.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList}
        Private ReadOnly _criteria As System.Windows.Forms.TextBox
        Private ReadOnly _recipients As AdminRecipientPicker
        Public Property Result As AdditionalRule

        Public Sub New(value As AdditionalRule, people As System.Collections.Generic.IEnumerable(Of RegisteredRecipient))
            MyBase.New(If(value Is Nothing, "Add routing rule", "Edit routing rule"))
            _value = If(value Is Nothing, New AdditionalRule(), AdminSession.Copy(value))
            Dim fields = AdminUi.Fields(Tabs, "Rule")
            _active = AdminUi.Check("Apply this rule", _value.Active)
            _kind.Items.AddRange({"Exact text or reference", "Meaning / semantic match"})
            _kind.SelectedIndex = If(_value.Kind = "semantic", 1, 0)
            _criteria = AdminUi.Text(_value.Criteria, True)
            _recipients = New AdminRecipientPicker(people, _value.RecipientIds)
            AdminUi.Field(fields, "Status", _active)
            AdminUi.Field(fields, "Match type", _kind)
            AdminUi.Field(fields, "Text, reference or matching condition *", _criteria)
            AdminUi.Field(fields, "Recipients when no name matches *", _recipients)
        End Sub

        Protected Overrides Sub Commit()
            If System.String.IsNullOrWhiteSpace(_criteria.Text) Then Throw New System.IO.InvalidDataException("Enter the text or condition to match.")
            If _recipients.SelectedIds.Count = 0 Then Throw New System.IO.InvalidDataException("Select at least one topic recipient.")
            _value.Active = _active.Checked
            _value.Kind = If(_kind.SelectedIndex = 1, "semantic", "exact")
            _value.Criteria = _criteria.Text.Trim()
            _value.RecipientIds = _recipients.SelectedIds
            Result = _value
        End Sub
    End Class
    Friend Class AdminRecipientImport
        Inherits AdminEditor
        Private ReadOnly _session As AdminSession
        Private ReadOnly _text As System.Windows.Forms.TextBox = AdminUi.Text("", True)
        Private ReadOnly _separator As New System.Windows.Forms.ComboBox With {.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList}
        Private ReadOnly _header As System.Windows.Forms.CheckBox = AdminUi.Check("First row contains column headings", False)
        Private ReadOnly _set As New System.Windows.Forms.ComboBox With {.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList}
        Private ReadOnly _grid As System.Windows.Forms.DataGridView = AdminUi.Grid("Row", "Name", "Email", "Organization", "Result")
        Private _previewKey As System.String = ""
        Public Sub New(session As AdminSession)
            MyBase.New("Import recipients")
            _session = session
            ApplyButton.Text = "Import"
            ApplyButton.Enabled = False
            Dim fields = AdminUi.Fields(Tabs, "Names and addresses")
            AdminUi.Field(fields, "Columns", New System.Windows.Forms.Label With {.Text = "Name | Email | Organization (optional). Paste cells copied from Excel, or open a CSV file.", .AutoSize = True})
            _text.MinimumSize = New System.Drawing.Size(180, 180)
            _text.MaxLength = 2000000
            _text.AcceptsTab = True
            AdminUi.Field(fields, "Data", _text)
            _separator.Items.AddRange({"Tab (Excel paste)", "Semicolon", "Comma"})
            _separator.SelectedIndex = 0
            AdminUi.Field(fields, "Column separator", _separator)
            AdminUi.Field(fields, "Headings", _header)
            _set.Items.Add(New AdminChoice With {.Id = "", .Label = "Register only — assign to a set later"})
            For Each rule In session.Configuration.RuleSets.OrderBy(Function(r) r.Name)
                _set.Items.Add(New AdminChoice With {.Id = rule.Id, .Label = rule.Name})
            Next
            _set.SelectedIndex = 0
            AdminUi.Field(fields, "Add new recipients to", _set)
            AdminUi.Help(_set, "Optionally make the new recipients available for matching in one recipient set. Existing recipients, topic rules and the default recipient remain unchanged.")
            AdminUi.Field(fields, "", AdminUi.Buttons(AdminUi.Button("Open CSV…", AddressOf OpenFile), AdminUi.Button("Preview", AddressOf Preview)))
            Dim page As New System.Windows.Forms.TabPage("Preview") With {.UseVisualStyleBackColor = True}
            page.Controls.Add(_grid)
            Tabs.TabPages.Add(page)
            AddHandler _text.TextChanged, Sub() InvalidatePreview()
            AddHandler _separator.SelectedIndexChanged, Sub() InvalidatePreview()
            AddHandler _header.CheckedChanged, Sub() InvalidatePreview()
            AddHandler _set.SelectedIndexChanged, Sub() InvalidatePreview()
        End Sub

        Private Function Separator() As System.String
            Return New System.String() {vbTab, ";", ","}(_separator.SelectedIndex)
        End Function

        Private Sub InvalidatePreview()
            _previewKey = ""
            _grid.Rows.Clear()
            ApplyButton.Enabled = False
        End Sub

        Private Function CurrentKey() As System.String
            Return Registry.Hash(_text.Text & ChrW(0) & Separator() & _header.Checked.ToString() & DirectCast(_set.SelectedItem, AdminChoice).Id)
        End Function

        Private Sub OpenFile()
            Using dialog As New System.Windows.Forms.OpenFileDialog With {.Title = "Import names and email addresses", .Filter = "CSV or tab-separated text (*.csv;*.tsv;*.txt)|*.csv;*.tsv;*.txt", .CheckFileExists = True}
                If dialog.ShowDialog(Me) <> System.Windows.Forms.DialogResult.OK Then Return
                If New System.IO.FileInfo(dialog.FileName).Length > 4000000 Then Throw New System.IO.InvalidDataException("File is too large. Import at most 1,000 rows at a time.")
                ' UTF-8 (with or without BOM), or BOM-marked UTF-16. Never silently replace invalid characters.
                Using reader As New System.IO.StreamReader(dialog.FileName, New System.Text.UTF8Encoding(False, True), True)
                    _text.Text = reader.ReadToEnd()
                End Using
                _separator.SelectedIndex = If(System.IO.Path.GetExtension(dialog.FileName).Equals(".tsv", System.StringComparison.OrdinalIgnoreCase), 0, 1)
            End Using
        End Sub

        Private Sub Preview()
            Dim rows = _session.PreviewRecipients(_text.Text, Separator(), _header.Checked)
            _grid.Rows.Clear()
            For Each row In rows
                Dim index = _grid.Rows.Add(row.Line, row.Name, row.Email, row.Organization, If(row.Issue = "", row.Status, row.Issue))
                If row.Issue <> "" Then _grid.Rows(index).DefaultCellStyle.ForeColor = System.Drawing.Color.DarkRed
            Next
            _previewKey = CurrentKey()
            ApplyButton.Enabled = rows.Any(Function(r) r.Add) AndAlso Not rows.Any(Function(r) r.Issue <> "")
            Tabs.SelectedIndex = 1
        End Sub

        Protected Overrides Sub Commit()
            If _previewKey = "" OrElse _previewKey <> CurrentKey() Then Throw New System.IO.InvalidDataException("Preview the current data before importing.")
            _session.ImportRecipients(_text.Text, Separator(), _header.Checked, DirectCast(_set.SelectedItem, AdminChoice).Id)
        End Sub
    End Class
End Namespace
