' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.Scan.Admin.RuleEditor.vb
' Purpose: Recipient set and filename controls.
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

Namespace ScanInbox
    Friend Class AdminChoice
        Public Property Id As System.String
        Public Property Label As System.String
        Public Overrides Function ToString() As System.String
            Return Label
        End Function
    End Class

    Friend Class AdminNameEditor
        Inherits System.Windows.Forms.UserControl
        Private ReadOnly _mode As New System.Windows.Forms.ComboBox With {.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList}
        Private ReadOnly _parts As New System.Windows.Forms.ListBox With {.IntegralHeight = False, .MinimumSize = New System.Drawing.Size(160, 100), .Dock = System.Windows.Forms.DockStyle.Fill}
        Private ReadOnly _field As New System.Windows.Forms.ComboBox With {.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList}
        Private ReadOnly _format As New System.Windows.Forms.ComboBox With {.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList}
        Private ReadOnly _separator As New System.Windows.Forms.ComboBox With {.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList}
        Private ReadOnly _custom As System.Windows.Forms.TextBox = AdminUi.Text("")
        Private ReadOnly _builder As System.Windows.Forms.TableLayoutPanel
        Private ReadOnly _preview As System.Windows.Forms.TextBox = AdminUi.Text("", True)
        Private ReadOnly _timeZone As System.Func(Of System.String)
        Private _loading As System.Boolean
        Public ReadOnly Property Pattern As System.String
            Get
                Select Case _mode.SelectedIndex
                    Case 0 : Return AttachmentNaming.DefaultTemplate
                    Case 1 : Return AdminSession.ComposeName(_parts.Items.Cast(Of AdminChoice)().Select(Function(p) p.Id), SeparatorValue())
                    Case Else : Return _custom.Text
                End Select
            End Get
        End Property

        Public Sub New(pattern As System.String, timeZone As System.Func(Of System.String))
            _timeZone = timeZone
            AutoSize = True
            AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
            Dock = System.Windows.Forms.DockStyle.Top
            _loading = True
            Dim root = AdminUi.Layout()
            root.Dock = System.Windows.Forms.DockStyle.Top
            root.Padding = New System.Windows.Forms.Padding(0)
            _mode.Items.AddRange({"Standard (recommended)", "Build from fields", "Custom pattern"})
            _custom.Text = pattern
            _mode.SelectedIndex = If(AttachmentNaming.IsStandardTemplate(pattern), 0, 2)
            _builder = AdminUi.Layout()
            _builder.Dock = System.Windows.Forms.DockStyle.Top
            _builder.Padding = New System.Windows.Forms.Padding(0)
            _builder.RowCount = 4
            For index As System.Int32 = 0 To 3
                _builder.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            Next
            _field.Items.AddRange({"Received date", "Received time", "Document date", "Package identifier", "Attachment number", "Content description", "Original name"})
            _separator.Items.AddRange({"Underscore", "Hyphen", "Space"})
            _separator.SelectedIndex = 0
            _field.SelectedIndex = 0
            UpdateFormats()
            Dim tools = AdminUi.Buttons(_field, _format, AdminUi.Button("Add field", AddressOf AddPart))
            _builder.Controls.Add(tools, 0, 0)
            _builder.Controls.Add(_parts, 0, 1)
            _builder.Controls.Add(AdminUi.Buttons(AdminUi.Button("Move up", Sub() MovePart(-1)), AdminUi.Button("Move down", Sub() MovePart(1)), AdminUi.Button("Remove", AddressOf RemovePart)), 0, 2)
            _builder.Controls.Add(AdminUi.Buttons(New System.Windows.Forms.Label With {.Text = "Separator", .AutoSize = True, .Anchor = System.Windows.Forms.AnchorStyles.Left}, _separator), 0, 3)
            For Each choice In New AdminChoice() {
                New AdminChoice With {.Id = "{received:yyyyMMdd_HHmmss}", .Label = "Received date and time"},
                New AdminChoice With {.Id = "{package}", .Label = "Package identifier"},
                New AdminChoice With {.Id = "{index:000}", .Label = "Attachment number (001)"},
                New AdminChoice With {.Id = "{content}", .Label = "Content description"}}
                _parts.Items.Add(choice)
            Next
            _preview.ReadOnly = True
            _preview.MinimumSize = New System.Drawing.Size(180, 60)
            _preview.AccessibleName = "Example filenames for two attachments"
            root.RowCount = 6
            For index As System.Int32 = 0 To 5
                root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            Next
            root.Controls.Add(_mode, 0, 0)
            root.Controls.Add(_builder, 0, 1)
            root.Controls.Add(_custom, 0, 2)
            root.Controls.Add(New System.Windows.Forms.Label With {.Text = "Preview · two example attachments", .AutoSize = True, .Margin = New System.Windows.Forms.Padding(3, 8, 3, 3)}, 0, 3)
            root.Controls.Add(_preview, 0, 4)
            root.Controls.Add(New System.Windows.Forms.Label With {.Text = "Original file extension is preserved. Duplicate names get a suffix.", .AutoSize = True, .UseMnemonic = False}, 0, 5)
            For Each control As System.Windows.Forms.Control In New System.Windows.Forms.Control() {_mode, _custom, _preview}
                control.Anchor = System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right
            Next
            Controls.Add(root)
            AddHandler root.SizeChanged, Sub()
                                             For Each label In root.Controls.OfType(Of System.Windows.Forms.Label)()
                                                 label.MaximumSize = New System.Drawing.Size(System.Math.Max(120, root.ClientSize.Width - label.Margin.Horizontal), 0)
                                             Next
                                         End Sub
            AddHandler _mode.SelectedIndexChanged, Sub() UpdatePreview()
            AddHandler _field.SelectedIndexChanged, Sub() UpdateFormats()
            AddHandler _custom.TextChanged, Sub() UpdatePreview()
            AddHandler _separator.SelectedIndexChanged, Sub() UpdatePreview()
            _loading = False
            UpdatePreview()
        End Sub

        Private Function SeparatorValue() As System.String
            Return If(_separator.SelectedIndex = 1, "-", If(_separator.SelectedIndex = 2, " ", "_"))
        End Function

        Private Sub UpdateFormats()
            _format.Items.Clear()
            Select Case _field.SelectedIndex
                Case 0, 2 : _format.Items.AddRange({"yyyyMMdd", "yyyy-MM-dd", "dd.MM.yyyy"})
                Case 1 : _format.Items.AddRange({"HHmmss", "HH-mm-ss", "HHmm"})
                Case 4 : _format.Items.AddRange({"000", "00", "0"})
                Case Else : _format.Items.Add("Default")
            End Select
            _format.SelectedIndex = 0
            _format.Enabled = _field.SelectedIndex <= 2 OrElse _field.SelectedIndex = 4
        End Sub

        Private Function NewPart() As AdminChoice
            Dim names = New System.String() {"received", "received", "documentDate", "package", "index", "content", "original"}
            Dim token = "{" & names(_field.SelectedIndex) & If(_format.Enabled, ":" & CStr(_format.SelectedItem), "") & "}"
            Return New AdminChoice With {.Id = token, .Label = CStr(_field.SelectedItem) & If(_format.Enabled, " (" & CStr(_format.SelectedItem) & ")", "")}
        End Function

        Private Sub AddPart()
            Dim part = NewPart()
            If _mode.SelectedIndex = 2 Then
                _custom.SelectedText = part.Id
                _custom.Focus()
            Else
                _parts.SelectedIndex = _parts.Items.Add(part)
                _mode.SelectedIndex = 1
            End If
            UpdatePreview()
        End Sub

        Private Sub MovePart(direction As System.Int32)
            Dim index = _parts.SelectedIndex
            Dim target = index + direction
            If index < 0 OrElse target < 0 OrElse target >= _parts.Items.Count Then Return
            Dim value = _parts.Items(index)
            _parts.Items.RemoveAt(index)
            _parts.Items.Insert(target, value)
            _parts.SelectedIndex = target
            UpdatePreview()
        End Sub

        Private Sub RemovePart()
            If _parts.SelectedIndex >= 0 Then _parts.Items.RemoveAt(_parts.SelectedIndex)
            UpdatePreview()
        End Sub

        Public Sub UpdatePreview()
            If _loading Then Return
            _builder.Visible = _mode.SelectedIndex <> 0
            _parts.Visible = _mode.SelectedIndex = 1
            _builder.Controls(2).Visible = _mode.SelectedIndex = 1
            _builder.Controls(3).Visible = _mode.SelectedIndex = 1
            _custom.Visible = _mode.SelectedIndex = 2
            Try
                Dim rule As New RuleSet With {.AttachmentNameTemplate = Pattern, .TimeZoneId = _timeZone()}
                _preview.Text = System.String.Join(System.Environment.NewLine, AdminSession.PreviewNames(rule).Select(Function(n) n.FileName))
                _preview.ForeColor = System.Drawing.SystemColors.WindowText
            Catch ex As System.Exception
                _preview.Text = ex.Message
                _preview.ForeColor = System.Drawing.Color.DarkRed
            End Try
        End Sub
    End Class

    Friend Class AdminRuleEditor
        Inherits AdminEditor
        Private ReadOnly _session As AdminSession
        Private ReadOnly _value As RuleSet
        Private ReadOnly _name As System.Windows.Forms.TextBox
        Private ReadOnly _active As System.Windows.Forms.CheckBox
        Private ReadOnly _senders As AdminTextList
        Private ReadOnly _mailbox As New System.Windows.Forms.ComboBox With {.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown}
        Private ReadOnly _people As AdminRecipientPicker
        Private ReadOnly _default As New System.Windows.Forms.ComboBox With {.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList}
        Private ReadOnly _zone As New System.Windows.Forms.ComboBox With {.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList}
        Private ReadOnly _deadline As New System.Windows.Forms.ComboBox With {.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList}
        Private ReadOnly _reminders As AdminTextList
        Private ReadOnly _retention As System.Windows.Forms.NumericUpDown
        Private ReadOnly _separate As System.Windows.Forms.CheckBox
        Private ReadOnly _rename As System.Windows.Forms.CheckBox
        Private ReadOnly _naming As AdminNameEditor
        Private ReadOnly _namingInstruction As System.Windows.Forms.TextBox
        Private ReadOnly _subject As System.Windows.Forms.TextBox
        Private ReadOnly _guidance As System.Windows.Forms.TextBox
        Private ReadOnly _rules As System.Windows.Forms.DataGridView = AdminUi.Grid("Match", "Condition", "Recipient assignment", "Status")
        Private ReadOnly _extras As System.Collections.Generic.List(Of AdditionalRule)
        Private ReadOnly _allowedPage As System.Windows.Forms.TabPage
        Private ReadOnly _defaultId As System.String

        Public Sub New(session As AdminSession, value As RuleSet, mailboxes As System.Collections.Generic.IEnumerable(Of System.String))
            MyBase.New(If(value Is Nothing, "Add recipient set", "Edit recipient set"))
            _session = session
            ApplyButton.Text = "Save"
            _value = If(value Is Nothing, New RuleSet(), AdminSession.Copy(value))
            _extras = _value.AdditionalRules
            _defaultId = _value.DefaultRecipientId
            Dim fields = AdminUi.Fields(Tabs, "Scanner source")
            _name = AdminUi.Text(_value.Name)
            _active = AdminUi.Check("Use this set when the Scan Handler is running", _value.Active)
            _senders = New AdminTextList(_value.ScannerSenders)
            _mailbox.Items.Add("All monitored mailboxes")
            _mailbox.Items.AddRange(mailboxes.Distinct(System.StringComparer.OrdinalIgnoreCase).Cast(Of System.Object)().ToArray())
            _mailbox.Text = If(_value.Mailbox = "", "All monitored mailboxes", _value.Mailbox)
            AdminUi.Field(fields, "Set name *", _name)
            AdminUi.Field(fields, "Status", _active)
            AdminUi.Field(fields, "Scanner email addresses *", _senders)
            AdminUi.Field(fields, "Receiving mailbox", _mailbox)
            fields = AdminUi.Fields(Tabs, "Recipient assignment")
            _allowedPage = Tabs.TabPages(Tabs.TabPages.Count - 1)
            _people = New AdminRecipientPicker(session.Configuration.Recipients, _value.AllowedRecipientIds)
            _separate = AdminUi.Check("Send separately to each recipient", _value.SeparateMessages)
            AdminUi.Field(fields, "Recipients available for matching *", _people)
            AdminUi.Field(fields, "Default recipient (unmatched mail) *", _default)
            AdminUi.Field(fields, "Separate emails", _separate)
            AdminUi.Help(_separate, "On: each recipient receives a separate email with their own instructions. Off: recipients share one email; recipient-specific additions are visible to everyone on it.")
            AdminUi.Field(fields, "Matching order", New System.Windows.Forms.Label With {.Text = "Named addressee → topic rule → default. Ambiguous or incomplete analysis goes to the default.", .AutoSize = True})
            fields = AdminUi.Fields(Tabs, "Dates & retention")
            For Each zone In System.TimeZoneInfo.GetSystemTimeZones()
                _zone.Items.Add(New AdminChoice With {.Id = zone.Id, .Label = zone.DisplayName})
            Next
            For Each item As AdminChoice In _zone.Items
                If item.Id = _value.TimeZoneId Then _zone.SelectedItem = item
            Next
            If _zone.SelectedIndex < 0 Then _zone.Items.Add(New AdminChoice With {.Id = _value.TimeZoneId, .Label = _value.TimeZoneId}) : _zone.SelectedIndex = _zone.Items.Count - 1
            _deadline.Items.AddRange({"Calculate calendar-day deadlines", "Do not calculate relative deadlines"})
            _deadline.SelectedIndex = If(_value.DeadlineProfile = "review_only", 1, 0)
            _reminders = New AdminTextList(_value.ReminderDays.Select(Function(d) d.ToString(System.Globalization.CultureInfo.InvariantCulture)), True)
            _retention = AdminUi.Number(_value.RetentionDays, 1, 3650)
            AdminUi.Field(fields, "Receipt stamps", New System.Windows.Forms.Label With {.Text = "Detected automatically, even without instructions. A reliable current stamp takes precedence over the email receipt date.", .AutoSize = True})
            AdminUi.Field(fields, "Date / time zone", _zone)
            AdminUi.Field(fields, "Relative deadlines", _deadline)
            AdminUi.Help(_deadline, "Calculate: periods stated in calendar days run from the day after receipt. Off: relative deadlines appear in the recipient email without a due-date calendar attachment. Explicit dates and follow-up reminders remain enabled. AutoPilot never pauses for a reviewer; ambiguous or unsupported periods are reported to the recipient without inventing a date.")
            AdminUi.Field(fields, "Follow-up days after receipt *", _reminders)
            AdminUi.Field(fields, "Keep scan content (days)", _retention)
            AdminUi.Help(_retention, "Stores copies of attachments and analysis for recipient follow-up questions. The period starts when the recipient copy is created; expired copies are removed during AutoPilot maintenance. Completed temporary files are removed after delivery. Failed or uncertain packages are kept for recovery. Activity records remain; this setting does not delete mailbox emails.")
            fields = AdminUi.Fields(Tabs, "Attachment names")
            _rename = AdminUi.Check("Rename attachments using AI", _value.RenameAttachments)
            _naming = New AdminNameEditor(_value.AttachmentNameTemplate, AddressOf ZoneId)
            _namingInstruction = AdminUi.Text(_value.AttachmentNamingInstruction, True)
            _namingInstruction.MaxLength = 20000
            AdminUi.Field(fields, "Status", _rename)
            AdminUi.Field(fields, "Filename format", _naming)
            AdminUi.Field(fields, "Content description instructions", _namingInstruction)
            _naming.Enabled = _rename.Checked
            _namingInstruction.Enabled = _rename.Checked
            AddHandler _rename.CheckedChanged, Sub()
                                                  _naming.Enabled = _rename.Checked
                                                  _namingInstruction.Enabled = _rename.Checked
                                              End Sub
            AddHandler _zone.SelectedIndexChanged, Sub() _naming.UpdatePreview()
            Dim rulePage As New System.Windows.Forms.TabPage("Topic routing") With {.UseVisualStyleBackColor = True}
            Dim ruleRoot = AdminUi.Layout()
            ruleRoot.AutoSize = False
            ruleRoot.RowCount = 2
            ruleRoot.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100))
            ruleRoot.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            ruleRoot.Controls.Add(_rules, 0, 0)
            ruleRoot.Controls.Add(AdminUi.Footer(AdminUi.Button("Remove", AddressOf DeleteExtra), AdminUi.Button("Edit", Sub() EditExtra(False)), AdminUi.Button("Add rule", Sub() EditExtra(True))), 0, 1)
            rulePage.Controls.Add(ruleRoot)
            Tabs.TabPages.Add(rulePage)
            fields = AdminUi.Fields(Tabs, "Instructions")
            _subject = AdminUi.Text(_value.SubjectTemplate)
            _subject.MaxLength = 240
            AdminUi.Field(fields, "Email subject (optional)", _subject)
            AdminUi.Help(_subject, "Leave blank for Scan: followed by the document subject, or a matter, company or topic. Enter fixed text, optionally including {subject} for that automatic description. Example: Incoming mail: {subject}")
            _guidance = AdminUi.Text(_value.Guidance, True)
            AdminUi.Field(fields, "Document analysis instructions", _guidance)
            AdminUi.Help(_guidance, "Optional instructions for interpreting all documents in this scanner email, for example how to recognize a receipt stamp or a matter reference. These do not change allowed recipients or deadline calculation. For recipient-specific output use the recipient instructions.")
            AddHandler _people.ItemCheck, Sub()
                                             If IsHandleCreated AndAlso Not IsDisposed Then
                                                 BeginInvoke(New System.Action(Sub()
                                                                                   If Not IsDisposed Then RefreshDefault()
                                                                               End Sub))
                                             End If
                                         End Sub
            AddHandler _rules.CellDoubleClick, Sub(sender, args)
                                                   If args.RowIndex >= 0 Then EditExtra(False)
                                               End Sub
            RefreshDefault()
            RefreshRules()
        End Sub

        Private Function ZoneId() As System.String
            Dim choice = TryCast(_zone.SelectedItem, AdminChoice)
            Return If(choice Is Nothing, _value.TimeZoneId, choice.Id)
        End Function

        Private Sub RefreshDefault()
            Dim selected = TryCast(_default.SelectedItem, AdminChoice)
            Dim id = If(selected Is Nothing, _defaultId, selected.Id)
            _default.Items.Clear()
            For Each person In _session.Configuration.Recipients.Where(Function(p) p.Active AndAlso _people.SelectedIds.Contains(p.Id)).OrderBy(Function(p) p.Name)
                Dim choice As New AdminChoice With {.Id = person.Id, .Label = person.Name & " <" & person.SmtpAddress & ">"}
                _default.Items.Add(choice)
                If person.Id = id Then _default.SelectedItem = choice
            Next
            If _default.SelectedIndex < 0 AndAlso _default.Items.Count = 1 Then _default.SelectedIndex = 0
        End Sub

        Private Sub RefreshRules()
            _rules.Rows.Clear()
            For Each rule In _extras
                Dim names = _session.Configuration.Recipients.Where(Function(p) rule.RecipientIds.Contains(p.Id)).Select(Function(p) p.Name)
                Dim index = _rules.Rows.Add(If(rule.Kind = "exact", "Exact text", "Semantic"), rule.Criteria, System.String.Join(", ", names), If(rule.Active, "Enabled", "Disabled"))
                _rules.Rows(index).Tag = rule.Id
            Next
        End Sub

        Private Sub EditExtra(create As System.Boolean)
            Dim id = AdminUi.SelectedId(_rules)
            If Not create AndAlso id = "" Then Return
            Dim people = _session.Configuration.Recipients.Where(Function(p) p.Active AndAlso _people.SelectedIds.Contains(p.Id)).ToList()
            If people.Count = 0 Then
                Tabs.SelectedTab = _allowedPage
                AdminUi.ErrorMessage(Me, "Select allowed recipients first.")
                Return
            End If
            Dim current = If(create, Nothing, _extras.Single(Function(r) r.Id = id))
            Using editor As New AdminExtraRuleEditor(current, people)
                If editor.ShowDialog(Me) <> System.Windows.Forms.DialogResult.OK Then Return
                Dim index = _extras.FindIndex(Function(r) r.Id = editor.Result.Id)
                If index < 0 Then
                    _extras.Add(editor.Result)
                Else
                    _extras(index) = editor.Result
                End If
            End Using
            RefreshRules()
        End Sub

        Private Sub DeleteExtra()
            Dim id = AdminUi.SelectedId(_rules)
            If id <> "" Then _extras.RemoveAll(Function(r) r.Id = id)
            RefreshRules()
        End Sub

        Protected Overrides Sub Commit()
            _senders.FlushPending()
            _reminders.FlushPending()
            If System.String.IsNullOrWhiteSpace(_name.Text) Then
                Tabs.SelectedIndex = 0
                _name.Focus()
                Throw New System.IO.InvalidDataException("Enter a recipient set name.")
            End If
            If _senders.Values.Count = 0 Then
                Tabs.SelectedIndex = 0
                Throw New System.IO.InvalidDataException("Add at least one scanner email address.")
            End If
            For Each sender In _senders.Values
                Try
                    Registry.Address(sender)
                Catch ex As System.Exception
                    Tabs.SelectedIndex = 0
                    Throw New System.IO.InvalidDataException("Invalid scanner email address: " & sender, ex)
                End Try
            Next
            If _reminders.Values.Count = 0 Then
                Tabs.SelectedIndex = 2
                Throw New System.IO.InvalidDataException("Add at least one follow-up period in days.")
            End If
            RefreshDefault()
            Dim fallback = TryCast(_default.SelectedItem, AdminChoice)
            If fallback Is Nothing Then
                Tabs.SelectedTab = _allowedPage
                Throw New System.IO.InvalidDataException("Select allowed recipients and a fallback recipient.")
            End If
            _value.Name = _name.Text.Trim()
            _value.Active = _active.Checked
            _value.ScannerSenders = _senders.Values
            _value.Mailbox = If(_mailbox.Text = "All monitored mailboxes", "", _mailbox.Text.Trim())
            _value.AllowedRecipientIds = _people.SelectedIds
            _value.DefaultRecipientId = fallback.Id
            If _extras.Any(Function(r) r.RecipientIds.Any(Function(id) Not _value.AllowedRecipientIds.Contains(id))) Then
                Throw New System.IO.InvalidDataException("An additional rule still uses a removed recipient. Update that rule or reselect the recipient.")
            End If
            _value.TimeZoneId = ZoneId()
            _value.DeadlineProfile = If(_deadline.SelectedIndex = 1, "review_only", "calendar_days")
            _value.ReminderDays = _reminders.Values.Select(Function(d) System.Int32.Parse(d, System.Globalization.CultureInfo.InvariantCulture)).ToList()
            _value.RetentionDays = CInt(_retention.Value)
            _value.SeparateMessages = _separate.Checked
            _value.RenameAttachments = _rename.Checked
            _value.AttachmentNameTemplate = _naming.Pattern
            _value.AttachmentNamingInstruction = _namingInstruction.Text
            _value.Guidance = _guidance.Text
            _value.SubjectTemplate = _subject.Text.Trim()
            _session.SaveChange(Sub() _session.PutRule(_value))
        End Sub
    End Class
End Namespace
