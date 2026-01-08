' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: DPIProgressForm.vb
' Purpose: Implements a small modal Windows Forms progress dialog with a progress bar,
'          status text, and a Cancel button. The UI is periodically refreshed from
'          shared state exposed by `ProgressBarModule`.
'
' Architecture:
'  - UI Elements: Header label, progress bar, status label, and Cancel button.
'  - State Source: Reads progress maximum/value/label and cancel flag from `ProgressBarModule`.
'  - Update Loop: A WinForms `Timer` triggers periodic UI refreshes (default: 250 ms).
'  - Cancellation: Clicking Cancel sets `ProgressBarModule.CancelOperation`; the timer also
'    closes the form if cancellation is detected.
'  - DPI Awareness: Uses WinForms autoscaling (`AutoScaleMode.Font`) for DPI/font scaling.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary

    ''' <summary>
    ''' Modal progress dialog that displays progress and status text and allows the user to cancel.
    ''' </summary>
    Public Class DPIProgressForm
        Inherits System.Windows.Forms.Form

        ''' <summary>
        ''' Progress bar showing aggregated progress as provided by <c>ProgressBarModule</c>.
        ''' </summary>
        Private WithEvents progressBar As System.Windows.Forms.ProgressBar

        ''' <summary>
        ''' Header label shown at the top of the dialog.
        ''' </summary>
        Private WithEvents lblHeader As System.Windows.Forms.Label

        ''' <summary>
        ''' Status label showing the current progress message.
        ''' </summary>
        Private WithEvents lblStatus As System.Windows.Forms.Label

        ''' <summary>
        ''' Button that triggers cancellation by setting <c>ProgressBarModule.CancelOperation</c>.
        ''' </summary>
        Private WithEvents btnCancel As System.Windows.Forms.Button

        ''' <summary>
        ''' Timer used to periodically refresh the UI from <c>ProgressBarModule</c>.
        ''' </summary>
        Private WithEvents uiTimer As System.Windows.Forms.Timer

        ''' <summary>
        ''' Initializes a new instance of the <see cref="DPIProgressForm"/> class.
        ''' </summary>
        ''' <param name="headerText">The caption text shown in the form title bar.</param>
        ''' <param name="initialLabel">The initial status label text.</param>
        Public Sub New(headerText As String, initialLabel As String)
            ' --- Auto-scale for DPI and font ---
            Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0F, 96.0F)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font

            ' --- Form properties ---
            Me.ClientSize = New System.Drawing.Size(400, 220)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ShowInTaskbar = False
            Me.Text = headerText

            ' Set icon
            Dim bmp As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
            Me.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

            ' Standard font
            Dim standardFont As New System.Drawing.Font(
    "Segoe UI",
    9.0F,
    System.Drawing.FontStyle.Regular,
    System.Drawing.GraphicsUnit.Point)

            ' --- Header label ---
            lblHeader = New System.Windows.Forms.Label()
            lblHeader.Text = "Progress ..."
            lblHeader.AutoSize = True
            lblHeader.Font = standardFont
            lblHeader.Location = New System.Drawing.Point(10, 10)
            lblHeader.Anchor = System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left
            Me.Controls.Add(lblHeader)

            ' --- ProgressBar ---
            progressBar = New System.Windows.Forms.ProgressBar()
            progressBar.Minimum = 0
            progressBar.Maximum = ProgressBarModule.GlobalProgressMax
            progressBar.Size = New System.Drawing.Size(Me.ClientSize.Width - 20, 25)
            progressBar.Location = New System.Drawing.Point(10, 40)
            progressBar.Anchor = System.Windows.Forms.AnchorStyles.Top Or
                     System.Windows.Forms.AnchorStyles.Left Or
                     System.Windows.Forms.AnchorStyles.Right
            Me.Controls.Add(progressBar)

            ' --- Status label ---
            lblStatus = New System.Windows.Forms.Label()
            lblStatus.Text = initialLabel
            lblStatus.AutoSize = False
            lblStatus.Font = standardFont
            lblStatus.Location = New System.Drawing.Point(10, 75)
            lblStatus.Size = New System.Drawing.Size(Me.ClientSize.Width - 20, 20)
            lblStatus.Anchor = System.Windows.Forms.AnchorStyles.Top Or
                 System.Windows.Forms.AnchorStyles.Left Or
                 System.Windows.Forms.AnchorStyles.Right
            Me.Controls.Add(lblStatus)

            ' --- Cancel button ---
            btnCancel = New System.Windows.Forms.Button()
            btnCancel.Text = "Cancel"
            btnCancel.Font = standardFont
            btnCancel.AutoSize = True
            btnCancel.Location = New System.Drawing.Point(10, 120)
            btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left
            AddHandler btnCancel.Click, AddressOf btnCancel_Click
            Me.Controls.Add(btnCancel)

            ' --- Resize event for dynamic adjustments ---
            AddHandler Me.ClientSizeChanged, AddressOf Form_Resize

            ' --- UI timer for periodic updates ---
            uiTimer = New System.Windows.Forms.Timer()
            uiTimer.Interval = 250 ' Update every 250 ms
            AddHandler uiTimer.Tick, AddressOf Timer_Tick
            uiTimer.Start()
        End Sub

        ''' <summary>
        ''' Updates control widths when the client size changes.
        ''' </summary>
        ''' <param name="sender">The event sender.</param>
        ''' <param name="e">Event arguments.</param>
        Private Sub Form_Resize(sender As Object, e As EventArgs)
            progressBar.Size = New System.Drawing.Size(Me.ClientSize.Width - 20, progressBar.Height)
            lblStatus.Size = New System.Drawing.Size(Me.ClientSize.Width - 20, lblStatus.Height)
        End Sub

        ' Timer tick event updates the progress bar and status label.

        ''' <summary>
        ''' Periodically refreshes the progress bar and status label from <c>ProgressBarModule</c>,
        ''' and closes the form if cancellation is requested.
        ''' </summary>
        ''' <param name="sender">The event sender.</param>
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
                ' It is possible to get an exception if the form is closing.
                System.Diagnostics.Debug.WriteLine("Timer error: " & ex.Message)
            End Try
        End Sub

        ' When the Cancel button is clicked, set the global cancel flag.

        ''' <summary>
        ''' Handles the Cancel button click by setting <c>ProgressBarModule.CancelOperation</c> to <c>True</c>.
        ''' </summary>
        ''' <param name="sender">The event sender.</param>
        ''' <param name="e">Event arguments.</param>
        Private Sub btnCancel_Click(sender As Object, e As EventArgs)
            ProgressBarModule.CancelOperation = True
        End Sub

        ' Stop the timer when the form is closed.

        ''' <summary>
        ''' Stops the UI timer and sets the global cancel flag when the form is closed.
        ''' </summary>
        ''' <param name="e">Provides data for the <see cref="System.Windows.Forms.Form.FormClosed"/> event.</param>
        Protected Overrides Sub OnFormClosed(e As System.Windows.Forms.FormClosedEventArgs)
            uiTimer.Stop()
            ProgressBarModule.CancelOperation = True
            MyBase.OnFormClosed(e)
        End Sub
    End Class

End Namespace