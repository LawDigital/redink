' Compact, DPI-aware native dialog. Source writes are enforced in PromptLibraryStore;
' this form never touches configuration and never chooses a destination path.
Friend NotInheritable Class SavePromptForm
    Inherits System.Windows.Forms.Form

    Private ReadOnly titleInput As New System.Windows.Forms.TextBox()
    Private ReadOnly categoryInput As New System.Windows.Forms.TextBox()
    Private ReadOnly promptInput As New System.Windows.Forms.TextBox()

    Public ReadOnly Property PromptTitle As String
        Get
            Return titleInput.Text
        End Get
    End Property

    Public ReadOnly Property PromptCategory As String
        Get
            Return categoryInput.Text
        End Get
    End Property

    Public ReadOnly Property PromptBody As String
        Get
            Return promptInput.Text
        End Get
    End Property

    Public Sub New(selection As String, suggestedCategory As String)
        Text = "Prompt speichern"
        ClientSize = New System.Drawing.Size(560, 340)
        MinimumSize = New System.Drawing.Size(480, 310)
        StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
        MaximizeBox = False
        MinimizeBox = False
        ShowInTaskbar = False
        TopMost = True
        BackColor = PickerUi.Surface
        Font = New System.Drawing.Font("Segoe UI", 9.0F)
        Icon = System.Drawing.Icon.ExtractAssociatedIcon(System.Windows.Forms.Application.ExecutablePath)

        ' Auto-size labels and an expanding prompt area avoid the cropped fields
        ' and one-line footer that failed on Windows with non-default DPI scaling.
        Dim outer As New System.Windows.Forms.TableLayoutPanel With {
            .Dock = System.Windows.Forms.DockStyle.Fill,
            .Padding = New System.Windows.Forms.Padding(16, 15, 16, 13),
            .ColumnCount = 1, .RowCount = 3,
            .BackColor = BackColor
        }
        outer.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
        outer.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
        outer.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
        outer.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 47.0F))
        Controls.Add(outer)

        Dim fields As New System.Windows.Forms.TableLayoutPanel With {
            .Dock = System.Windows.Forms.DockStyle.Top,
            .AutoSize = True, .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
            .ColumnCount = 2, .RowCount = 2,
            .Margin = New System.Windows.Forms.Padding(0, 0, 0, 12)
        }
        fields.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 58.0F))
        fields.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 42.0F))
        fields.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
        fields.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 34.0F))
        Dim titleLabel As System.Windows.Forms.Label = PickerUi.Caption("Titel")
        Dim categoryLabel As System.Windows.Forms.Label = PickerUi.Caption("Kategorie · optional")
        categoryLabel.Margin = New System.Windows.Forms.Padding(10, 0, 0, 5)
        fields.Controls.Add(titleLabel, 0, 0)
        fields.Controls.Add(categoryLabel, 1, 0)
        PickerUi.StyleTextBox(titleInput)
        PickerUi.StyleTextBox(categoryInput)
        titleInput.PlaceholderText = "Bezeichnung des Prompts"
        categoryInput.PlaceholderText = "Ohne Kategorie"
        categoryInput.Text = If(suggestedCategory = "(Uncategorized)" OrElse suggestedCategory = "Alle Kategorien", "", suggestedCategory)
        titleInput.Dock = System.Windows.Forms.DockStyle.Top
        categoryInput.Dock = System.Windows.Forms.DockStyle.Top
        titleInput.Margin = New System.Windows.Forms.Padding(0, 1, 0, 0)
        categoryInput.Margin = New System.Windows.Forms.Padding(10, 1, 0, 0)
        fields.Controls.Add(titleInput, 0, 1)
        fields.Controls.Add(categoryInput, 1, 1)
        outer.Controls.Add(fields, 0, 0)

        Dim promptArea As New System.Windows.Forms.TableLayoutPanel With {
            .Dock = System.Windows.Forms.DockStyle.Fill,
            .ColumnCount = 1, .RowCount = 2,
            .Margin = New System.Windows.Forms.Padding(0)
        }
        promptArea.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
        promptArea.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
        promptArea.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
        Dim promptLabel As System.Windows.Forms.Label = PickerUi.Caption("Prompt")
        promptLabel.Margin = New System.Windows.Forms.Padding(0, 0, 0, 7)
        promptArea.Controls.Add(promptLabel, 0, 0)
        PickerUi.StyleTextBox(promptInput)
        promptInput.Multiline = True
        promptInput.AcceptsReturn = True
        promptInput.AcceptsTab = True
        promptInput.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        promptInput.Dock = System.Windows.Forms.DockStyle.Fill
        promptInput.Text = If(selection, "")
        promptInput.Margin = New System.Windows.Forms.Padding(0)
        promptArea.Controls.Add(promptInput, 0, 1)
        outer.Controls.Add(promptArea, 0, 1)

        Dim footer As New System.Windows.Forms.TableLayoutPanel With {
            .Dock = System.Windows.Forms.DockStyle.Fill,
            .ColumnCount = 2, .RowCount = 1,
            .Margin = New System.Windows.Forms.Padding(0, 11, 0, 0)
        }
        footer.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
        footer.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize))
        footer.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
        Dim buttons As New System.Windows.Forms.FlowLayoutPanel With {
            .AutoSize = True, .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
            .WrapContents = False,
            .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
            .Anchor = System.Windows.Forms.AnchorStyles.Right Or System.Windows.Forms.AnchorStyles.Top,
            .Margin = New System.Windows.Forms.Padding(0)
        }
        Dim cancel As System.Windows.Forms.Button = PickerUi.ActionButton("Abbrechen", "close", False, 120)
        cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Dim save As System.Windows.Forms.Button = PickerUi.ActionButton("Speichern", "check", True, 120)
        AddHandler save.Click, AddressOf SaveClicked
        buttons.Controls.Add(cancel)
        buttons.Controls.Add(save)
        footer.Controls.Add(buttons, 1, 0)
        outer.Controls.Add(footer, 0, 2)
        AcceptButton = save
        CancelButton = cancel
        AddHandler Shown, Sub(sender As Object, e As System.EventArgs) titleInput.Focus()
    End Sub

    Private Sub SaveClicked(sender As Object, e As System.EventArgs)
        If System.String.IsNullOrWhiteSpace(titleInput.Text) Then
            System.Windows.Forms.MessageBox.Show(Me, "Bitte einen Titel eingeben.", "Prompt speichern", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
            titleInput.Focus()
            Return
        End If
        If System.String.IsNullOrWhiteSpace(promptInput.Text) Then
            System.Windows.Forms.MessageBox.Show(Me, "Bitte einen Prompt eingeben.", "Prompt speichern", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
            promptInput.Focus()
            Return
        End If
        DialogResult = System.Windows.Forms.DialogResult.OK
        Close()
    End Sub
End Class
