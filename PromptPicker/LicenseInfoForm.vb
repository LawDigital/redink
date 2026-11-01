' Part of Red Ink (Prompt Picker). Calm, read-only license viewer launched from the tray.
Option Strict On
Option Explicit On

Friend NotInheritable Class LicenseInfoForm
    Inherits System.Windows.Forms.Form

    Public Sub New()
        Text = "Red Ink Prompt Picker – Lizenzhinweise"
        ClientSize = New System.Drawing.Size(740, 550)
        MinimumSize = New System.Drawing.Size(490, 375)
        StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        ShowInTaskbar = False
        TopMost = True
        BackColor = PickerUi.Surface
        Font = New System.Drawing.Font("Segoe UI", 9.0F)
        Icon = System.Drawing.Icon.ExtractAssociatedIcon(System.Windows.Forms.Application.ExecutablePath)

        Dim layout As New System.Windows.Forms.TableLayoutPanel With {
            .Dock = System.Windows.Forms.DockStyle.Fill,
            .Padding = New System.Windows.Forms.Padding(17, 14, 17, 12),
            .ColumnCount = 1,
            .RowCount = 4
        }
        layout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
        layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
        layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
        layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
        layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 45.0F))
        Controls.Add(layout)

        Dim heading As New System.Windows.Forms.Label With {
            .Text = "Red Ink · Lizenz und Drittanbieterhinweise",
            .AutoSize = True,
            .Font = New System.Drawing.Font("Segoe UI", 10.0F, System.Drawing.FontStyle.Regular),
            .ForeColor = PickerUi.Ink,
            .Margin = New System.Windows.Forms.Padding(0, 0, 0, 4)
        }
        layout.Controls.Add(heading, 0, 0)

        Dim copyrightInfo As New System.Windows.Forms.Label With {
            .Text = "© 2025– LawDigital Ltd., Schweiz · redink.ai",
            .AutoSize = True,
            .ForeColor = PickerUi.Muted,
            .Margin = New System.Windows.Forms.Padding(0, 0, 0, 11)
        }
        layout.Controls.Add(copyrightInfo, 0, 1)

        Dim viewer As New System.Windows.Forms.TextBox With {
            .Text = LegalResources.LicenseText,
            .Dock = System.Windows.Forms.DockStyle.Fill,
            .Margin = New System.Windows.Forms.Padding(0),
            .Multiline = True,
            .ReadOnly = True,
            .ScrollBars = System.Windows.Forms.ScrollBars.Vertical,
            .WordWrap = True,
            .BackColor = System.Drawing.Color.White,
            .ForeColor = PickerUi.Ink,
            .BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle,
            .Font = New System.Drawing.Font("Segoe UI", 9.0F),
            .AccessibleName = "Lizenztext"
        }
        layout.Controls.Add(viewer, 0, 2)

        Dim buttons As New System.Windows.Forms.FlowLayoutPanel With {
            .Dock = System.Windows.Forms.DockStyle.Fill,
            .FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft,
            .WrapContents = False,
            .Margin = New System.Windows.Forms.Padding(0, 10, 0, 0)
        }
        Dim closeButton As System.Windows.Forms.Button = PickerUi.ActionButton("Schliessen", "close", False, 115)
        closeButton.DialogResult = System.Windows.Forms.DialogResult.OK
        buttons.Controls.Add(closeButton)
        layout.Controls.Add(buttons, 0, 3)
        AcceptButton = closeButton
        CancelButton = closeButton
    End Sub
End Class
