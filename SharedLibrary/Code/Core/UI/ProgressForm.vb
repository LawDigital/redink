' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ProgressForm.vb
' Purpose:
'   Provides a small modal WinForms dialog that displays a header, a progress bar,
'   a status line, and a Cancel button.
'
' Architecture:
'   - The form polls `ProgressBarModule` on a UI timer (250 ms) to update:
'       - `ProgressBarModule.GlobalProgressMax`   -> progress bar maximum
'       - `ProgressBarModule.GlobalProgressValue` -> progress bar value (clamped)
'       - `ProgressBarModule.GlobalProgressLabel` -> status label text
'   - If `ProgressBarModule.CancelOperation` becomes True, the form closes itself
'     and returns `DialogResult.Cancel`.
'   - Clicking Cancel sets `ProgressBarModule.CancelOperation` to True.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary

    ''' <summary>
    ''' Modal progress/cancellation dialog that reflects progress state stored in <c>ProgressBarModule</c>.
    ''' </summary>
    Public Class ProgressForm
        Inherits System.Windows.Forms.Form

        ''' <summary>Progress bar reflecting <c>ProgressBarModule.GlobalProgressValue</c> and <c>.GlobalProgressMax</c>.</summary>
        Private WithEvents progressBar As System.Windows.Forms.ProgressBar

        ''' <summary>Header label shown at the top of the dialog.</summary>
        Private WithEvents lblHeader As System.Windows.Forms.Label

        ''' <summary>Status label shown under the progress bar.</summary>
        Private WithEvents lblStatus As System.Windows.Forms.Label

        ''' <summary>Cancel button that sets <c>ProgressBarModule.CancelOperation</c>.</summary>
        Private WithEvents btnCancel As System.Windows.Forms.Button

        ''' <summary>UI timer used to periodically poll <c>ProgressBarModule</c> and refresh the UI.</summary>
        Private WithEvents uiTimer As System.Windows.Forms.Timer

        ''' <summary>
        ''' Initializes a new instance of the progress dialog.
        ''' </summary>
        ''' <param name="headerText">Text displayed in the header label.</param>
        ''' <param name="initialLabel">Initial text displayed in the status label.</param>
        Public Sub New(headerText As String, initialLabel As String)
            ' --- Use Font scaling ---
            Dim standardFont As New System.Drawing.Font(
    "Segoe UI",
    9.0F,
    System.Drawing.FontStyle.Regular,
    System.Drawing.GraphicsUnit.Point)

            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.Font = standardFont
            Me.AutoSize = True
            Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ShowInTaskbar = False
            Me.Text = SharedMethods.AN ' headerText

            ' --- Set icon ---
            Dim bmp As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
            Me.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

            ' --- Header Label ---
            lblHeader = New System.Windows.Forms.Label() With {
    .Text = headerText,
    .AutoSize = True
}

            ' --- ProgressBar ---
            progressBar = New System.Windows.Forms.ProgressBar() With {
    .Minimum = 0,
    .Maximum = ProgressBarModule.GlobalProgressMax,
    .Dock = System.Windows.Forms.DockStyle.Fill
}

            ' --- Status Label ---
            lblStatus = New System.Windows.Forms.Label() With {
    .Text = initialLabel,
    .AutoSize = True,
    .Dock = System.Windows.Forms.DockStyle.Fill
}

            ' --- Cancel Button ---
            btnCancel = New System.Windows.Forms.Button() With {
    .Text = "Cancel",
    .AutoSize = True
}
            AddHandler btnCancel.Click, AddressOf btnCancel_Click

            ' --- Layout in TableLayoutPanel ---
            Dim layout As New System.Windows.Forms.TableLayoutPanel() With {
    .AutoSize = True,
    .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
    .Dock = System.Windows.Forms.DockStyle.Fill,
    .Padding = New System.Windows.Forms.Padding(10),
    .ColumnCount = 1,
    .RowCount = 4
}
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))

            layout.Controls.Add(lblHeader, 0, 0)
            layout.Controls.Add(progressBar, 0, 1)
            layout.Controls.Add(lblStatus, 0, 2)
            layout.Controls.Add(btnCancel, 0, 3)

            Me.Controls.Add(layout)

            ' --- UI timer for periodic updates ---
            uiTimer = New System.Windows.Forms.Timer() With {
    .Interval = 250 ' Update every 250 ms
}
            AddHandler uiTimer.Tick, AddressOf Timer_Tick
            uiTimer.Start()
        End Sub

        ''' <summary>
        ''' Polls <c>ProgressBarModule</c> and updates the progress bar and status label; closes on cancellation.
        ''' </summary>
        ''' <param name="sender">Timer instance.</param>
        ''' <param name="e">Event arguments.</param>
        Private Sub Timer_Tick(sender As Object, e As EventArgs)
            Try
                ' Update the progress bar maximum and value.
                progressBar.Maximum = ProgressBarModule.GlobalProgressMax
                progressBar.Value = Math.Min(ProgressBarModule.GlobalProgressValue, progressBar.Maximum)

                ' Update the status text.
                lblStatus.Text = ProgressBarModule.GlobalProgressLabel

                ' If the cancel flag is set, close the form with a Cancel result.
                If ProgressBarModule.CancelOperation Then
                    Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
                    Me.Close()
                End If
            Catch ex As System.Exception
                ' Possible exception if the form is closing.
                System.Diagnostics.Debug.WriteLine("Timer error: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Sets the global cancel flag.
        ''' </summary>
        ''' <param name="sender">The Cancel button.</param>
        ''' <param name="e">Event arguments.</param>
        Private Sub btnCancel_Click(sender As Object, e As EventArgs)
            ProgressBarModule.CancelOperation = True
        End Sub

        ''' <summary>
        ''' Stops the UI timer and sets the cancel flag when the form is closed.
        ''' </summary>
        ''' <param name="e">Form closed event arguments.</param>
        Protected Overrides Sub OnFormClosed(e As System.Windows.Forms.FormClosedEventArgs)
            uiTimer.Stop()
            ProgressBarModule.CancelOperation = True
            MyBase.OnFormClosed(e)
        End Sub
    End Class


End Namespace