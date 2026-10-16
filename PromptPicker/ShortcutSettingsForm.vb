' Short, coherent native settings dialog with no second branding/header strip.
' Modifiers and candidate keys retain the original validation/registration flow.
Friend NotInheritable Class ShortcutSettingsForm
    Inherits System.Windows.Forms.Form

    Private NotInheritable Class KeyOption
        Public ReadOnly Property VirtualKey As Integer
        Public ReadOnly Property Display As String

        Public Sub New(code As Integer)
            VirtualKey = code
            Display = CType(code, System.Windows.Forms.Keys).ToString()
        End Sub

        Public Overrides Function ToString() As String
            Return Display
        End Function
    End Class

    Private ReadOnly ctrlCheck As New System.Windows.Forms.CheckBox With {.Text = "Ctrl", .Checked = True, .AutoSize = True}
    Private ReadOnly altBox As New System.Windows.Forms.CheckBox With {.Text = "Alt", .Checked = True, .AutoSize = True}
    Private ReadOnly shiftBox As New System.Windows.Forms.CheckBox With {.Text = "Shift", .AutoSize = True}
    Private ReadOnly keyInput As New System.Windows.Forms.ComboBox With {.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList}

    Public ReadOnly Property Shortcut As PickerSettings
        Get
            Dim modifiers As UInteger = If(ctrlCheck.Checked, Native.MOD_CONTROL, 0UI) Or If(altBox.Checked, Native.MOD_ALT, 0UI) Or If(shiftBox.Checked, Native.MOD_SHIFT, 0UI)
            Dim selected As KeyOption = TryCast(keyInput.SelectedItem, KeyOption)
            Return New PickerSettings With {.HotkeyModifiers = modifiers, .HotkeyKey = If(selected Is Nothing, CInt(System.Windows.Forms.Keys.P), selected.VirtualKey)}
        End Get
    End Property

    Public Sub New(current As PickerSettings)
        Text = "Tastenkombination ändern"
        ClientSize = New System.Drawing.Size(515, 218)
        AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        MinimumSize = New System.Drawing.Size(515, 230)
        FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        MaximizeBox = False
        MinimizeBox = False
        ShowInTaskbar = False
        TopMost = True
        BackColor = PickerUi.Surface
        Font = New System.Drawing.Font("Segoe UI", 9.0F)
        Icon = System.Drawing.Icon.ExtractAssociatedIcon(System.Windows.Forms.Application.ExecutablePath)

        Dim layout As New System.Windows.Forms.TableLayoutPanel With {
            .Dock = System.Windows.Forms.DockStyle.Fill,
            .Padding = New System.Windows.Forms.Padding(17, 15, 17, 12),
            .ColumnCount = 1, .RowCount = 5
        }
        layout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
        layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
        layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 47.0F))
        layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
        layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
        layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 45.0F))
        Controls.Add(layout)
        layout.Controls.Add(PickerUi.Caption("Tastenkombination"), 0, 0)

        Dim keys As New System.Windows.Forms.TableLayoutPanel With {
            .Dock = System.Windows.Forms.DockStyle.Fill,
            .ColumnCount = 4, .RowCount = 1,
            .Margin = New System.Windows.Forms.Padding(0, 3, 0, 0)
        }
        keys.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 72.0F))
        keys.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 67.0F))
        keys.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 81.0F))
        keys.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
        keys.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
        For Each toggle As System.Windows.Forms.CheckBox In {ctrlCheck, altBox, shiftBox}
            toggle.Anchor = System.Windows.Forms.AnchorStyles.Left
            toggle.Margin = New System.Windows.Forms.Padding(0)
            toggle.ForeColor = PickerUi.Ink
        Next
        keys.Controls.Add(ctrlCheck, 0, 0)
        keys.Controls.Add(altBox, 1, 0)
        keys.Controls.Add(shiftBox, 2, 0)
        keyInput.Dock = System.Windows.Forms.DockStyle.Fill
        keyInput.Margin = New System.Windows.Forms.Padding(1, 5, 0, 5)
        keyInput.Font = New System.Drawing.Font("Segoe UI", 9.5F)
        keyInput.AccessibleName = "Taste"
        keys.Controls.Add(keyInput, 3, 0)
        layout.Controls.Add(keys, 0, 1)

        Dim hint As New System.Windows.Forms.Label With {
            .Text = "Mindestens Ctrl oder Alt auswählen.",
            .AutoSize = True, .ForeColor = PickerUi.Muted,
            .Margin = New System.Windows.Forms.Padding(0, 8, 0, 0),
            .MaximumSize = New System.Drawing.Size(465, 0)
        }
        layout.Controls.Add(hint, 0, 2)

        Dim buttonRow As New System.Windows.Forms.FlowLayoutPanel With {
            .Dock = System.Windows.Forms.DockStyle.Fill,
            .WrapContents = False,
            .FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft,
            .Margin = New System.Windows.Forms.Padding(0, 6, 0, 0)
        }
        Dim apply As System.Windows.Forms.Button = PickerUi.ActionButton("Übernehmen", "check", True, 133)
        Dim cancel As System.Windows.Forms.Button = PickerUi.ActionButton("Abbrechen", "close", False, 123)
        cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        buttonRow.Controls.Add(apply)
        buttonRow.Controls.Add(cancel)
        layout.Controls.Add(buttonRow, 0, 4)

        For code As Integer = 65 To 90
            keyInput.Items.Add(New KeyOption(code))
        Next
        For code As Integer = 112 To 123
            keyInput.Items.Add(New KeyOption(code))
        Next
        ctrlCheck.Checked = (current.HotkeyModifiers And Native.MOD_CONTROL) <> 0UI
        altBox.Checked = (current.HotkeyModifiers And Native.MOD_ALT) <> 0UI
        shiftBox.Checked = (current.HotkeyModifiers And Native.MOD_SHIFT) <> 0UI
        For i As Integer = 0 To keyInput.Items.Count - 1
            Dim entry As KeyOption = DirectCast(keyInput.Items(i), KeyOption)
            If entry.VirtualKey = current.HotkeyKey Then
                keyInput.SelectedIndex = i
                Exit For
            End If
        Next
        If keyInput.SelectedIndex < 0 Then keyInput.SelectedIndex = 0
        AddHandler apply.Click, Sub(sender As Object, e As System.EventArgs)
                                    Dim requested As PickerSettings = Shortcut
                                    If Not ShortcutService.IsValid(requested.HotkeyModifiers, requested.HotkeyKey) Then
                                        System.Windows.Forms.MessageBox.Show(Me, "Mindestens Ctrl oder Alt auswählen.", "Tastenkombination", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
                                        Return
                                    End If
                                    DialogResult = System.Windows.Forms.DialogResult.OK
                                    Close()
                                End Sub
        AcceptButton = apply
        CancelButton = cancel
    End Sub
End Class
