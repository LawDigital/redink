' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: DPIProgressForm.vb
' Purpose: Implements a small Windows Forms progress dialog with a progress bar,
'          status text, and a Cancel button. The UI is periodically refreshed from
'          shared state exposed by `ProgressBarModule`.
'
' Architecture:
'  - UI Elements: Header label, progress bar, status label, and Cancel button.
'  - State Source: Reads progress maximum/value/label and cancel flag from `ProgressBarModule`.
'  - Update Loop: A WinForms `Timer` triggers periodic UI refreshes (default: 250 ms).
'  - Cancellation: Clicking Cancel sets `ProgressBarModule.CancelOperation`; the timer also
'    closes the form if cancellation is detected.
'  - DPI Awareness: Uses WinForms autoscaling (`AutoScaleMode.Font`) and layout panels
'    instead of absolute positioning, so wrapped status text and button spacing scale correctly.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary

    ''' <summary>
    ''' Progress dialog that displays progress and status text and allows the user to cancel.
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
        ''' Root layout panel.
        ''' </summary>
        Private layoutRoot As System.Windows.Forms.TableLayoutPanel

        ''' <summary>
        ''' Button row container.
        ''' </summary>
        Private buttonPanel As System.Windows.Forms.FlowLayoutPanel

        ''' <summary>
        ''' Initializes a new instance of the <see cref="DPIProgressForm"/> class.
        ''' </summary>
        ''' <param name="headerText">The caption text shown in the form title bar.</param>
        ''' <param name="initialLabel">The initial status label text.</param>
        Public Sub New(headerText As String, initialLabel As String)
            ' --- Auto-scale for DPI and font ---
            Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0F, 96.0F)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font

            Dim standardFont As New System.Drawing.Font(
                "Segoe UI",
                9.0F,
                System.Drawing.FontStyle.Regular,
                System.Drawing.GraphicsUnit.Point)

            Me.Font = standardFont
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ShowInTaskbar = False
            Me.TopMost = True
            Me.Text = headerText
            Me.AutoSize = True
            Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
            Me.Padding = New System.Windows.Forms.Padding(0)
            Me.MinimumSize = New System.Drawing.Size(440, 0)

            ' Set icon
            Dim bmp As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
            Me.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

            ' --- Root layout ---
            layoutRoot = New System.Windows.Forms.TableLayoutPanel() With {
                .AutoSize = True,
                .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .Padding = New System.Windows.Forms.Padding(12),
                .ColumnCount = 1,
                .RowCount = 4
            }
            layoutRoot.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
            layoutRoot.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            layoutRoot.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            layoutRoot.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            layoutRoot.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            Me.Controls.Add(layoutRoot)

            ' --- Header label ---
            lblHeader = New System.Windows.Forms.Label() With {
                .Text = headerText,
                .AutoSize = True,
                .Font = standardFont,
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .Margin = New System.Windows.Forms.Padding(0, 0, 0, 8)
            }
            layoutRoot.Controls.Add(lblHeader, 0, 0)

            ' --- Progress bar ---
            progressBar = New System.Windows.Forms.ProgressBar() With {
                .Minimum = 0,
                .Maximum = ProgressBarModule.GlobalProgressMax,
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .Height = 22,
                .Margin = New System.Windows.Forms.Padding(0, 0, 0, 8)
            }
            layoutRoot.Controls.Add(progressBar, 0, 1)

            ' --- Status label ---
            lblStatus = New System.Windows.Forms.Label() With {
                .Text = initialLabel,
                .AutoSize = True,
                .Font = standardFont,
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .Margin = New System.Windows.Forms.Padding(0, 0, 0, 8)
            }
            layoutRoot.Controls.Add(lblStatus, 0, 2)

            ' --- Cancel button row ---
            buttonPanel = New System.Windows.Forms.FlowLayoutPanel() With {
                .AutoSize = True,
                .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft,
                .WrapContents = False,
                .Margin = New System.Windows.Forms.Padding(0)
            }

            btnCancel = New System.Windows.Forms.Button() With {
                .Text = "Cancel",
                .AutoSize = True,
                .Margin = New System.Windows.Forms.Padding(0)
            }
            AddHandler btnCancel.Click, AddressOf btnCancel_Click
            buttonPanel.Controls.Add(btnCancel)

            layoutRoot.Controls.Add(buttonPanel, 0, 3)

            ' --- Layout updates ---
            AddHandler Me.Shown, AddressOf Form_Shown
            AddHandler Me.ClientSizeChanged, AddressOf Form_Resize

            ' --- UI timer for periodic updates ---
            uiTimer = New System.Windows.Forms.Timer() With {
                .Interval = 250
            }
            AddHandler uiTimer.Tick, AddressOf Timer_Tick
            uiTimer.Start()
        End Sub

        ''' <summary>
        ''' Applies width-dependent wrapping for the status label.
        ''' </summary>
        Private Sub UpdateStatusLayout()
            Try
                Dim horizontalPadding As Integer = layoutRoot.Padding.Left + layoutRoot.Padding.Right
                Dim usableWidth As Integer = Math.Max(200, Me.ClientSize.Width - horizontalPadding)
                lblStatus.MaximumSize = New System.Drawing.Size(usableWidth, 0)
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' Updates wrapping after the form is first shown.
        ''' </summary>
        Private Sub Form_Shown(sender As Object, e As System.EventArgs)
            UpdateStatusLayout()
            SharedMethods.ForceDialogToForeground(Me)
        End Sub

        ''' <summary>
        ''' Updates control wrapping when the client size changes.
        ''' </summary>
        Private Sub Form_Resize(sender As Object, e As System.EventArgs)
            UpdateStatusLayout()
        End Sub

        ''' <summary>
        ''' Periodically refreshes the progress bar and status label from <c>ProgressBarModule</c>,
        ''' and closes the form if cancellation is requested.
        ''' </summary>
        ''' <param name="sender">The event sender.</param>
        ''' <param name="e">Event arguments.</param>
        Private Sub Timer_Tick(sender As Object, e As System.EventArgs)
            Try
                progressBar.Maximum = Math.Max(1, ProgressBarModule.GlobalProgressMax)
                progressBar.Value = Math.Min(ProgressBarModule.GlobalProgressValue, progressBar.Maximum)

                Dim newStatus As String = ProgressBarModule.GlobalProgressLabel
                If lblStatus.Text <> newStatus Then
                    lblStatus.Text = newStatus
                    UpdateStatusLayout()
                End If

                If ProgressBarModule.CancelOperation Then
                    Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
                    Me.Close()
                End If
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("Timer error: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Handles the Cancel button click by setting <c>ProgressBarModule.CancelOperation</c> to <c>True</c>.
        ''' </summary>
        ''' <param name="sender">The event sender.</param>
        ''' <param name="e">Event arguments.</param>
        Private Sub btnCancel_Click(sender As Object, e As System.EventArgs)
            ProgressBarModule.CancelOperation = True
        End Sub

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
