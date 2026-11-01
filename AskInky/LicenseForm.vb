' Part of "Red Ink" (Ask Inky)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: LicenseForm.vb
' Purpose: Displays the bundled Red Ink and third-party license notices offline.
' Architecture: Read-only WinForms viewer, independent of Helper connectivity.
' Security: Content is embedded, not downloaded or interpreted as executable HTML.
' =============================================================================
Option Strict On
Option Explicit On

Friend NotInheritable Class LicenseForm
    Inherits System.Windows.Forms.Form

    Public Sub New(language As String)
        Text = "Ask Inky – " & UiStrings.GetText("desktop.thirdPartyLicenses", language)
        ClientSize = New System.Drawing.Size(760, 580)
        MinimumSize = New System.Drawing.Size(420, 350)
        StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Font = New System.Drawing.Font("Segoe UI", 9.0F)
        Icon = System.Drawing.Icon.ExtractAssociatedIcon(System.Windows.Forms.Application.ExecutablePath)
        Dim viewer As New System.Windows.Forms.TextBox() With {
            .Multiline = True, .ReadOnly = True, .ScrollBars = System.Windows.Forms.ScrollBars.Both,
            .WordWrap = True, .Dock = System.Windows.Forms.DockStyle.Fill,
            .Font = New System.Drawing.Font("Consolas", 9.0F),
            .Text = LegalResources.LicenseText
        }
        Controls.Add(viewer)
    End Sub
End Class
