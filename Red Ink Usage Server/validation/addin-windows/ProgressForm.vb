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
        Private _ownedFont As System.Drawing.Font
        Private _ownedBitmap As System.Drawing.Bitmap
        Private _ownedIcon As System.Drawing.Icon

        ''' <summary>
        ''' Initializes a new instance of the progress dialog.
        ''' </summary>
        ''' <param name="headerText">Text displayed in the header label.</param>
        ''' <param name="initialLabel">Initial text displayed in the status label.</param>
        Public Sub New(headerText As String, initialLabel As String)
            MyBase.New()
            Try
                InitializeView(headerText, initialLabel)
            Catch
                Me.Dispose()
                Throw
            End Try
        End Sub

        Private Sub InitializeView(headerText As String, initialLabel As String)
            ' --- Use Font scaling ---
            _ownedFont = New System.Drawing.Font(
    "Segoe UI",
    9.0F,
    System.Drawing.FontStyle.Regular,
    System.Drawing.GraphicsUnit.Point)
            Dim standardFont As System.Drawing.Font = _ownedFont

            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.Font = standardFont
            Me.AutoSize = True
            Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ShowInTaskbar = False
            Me.TopMost = True
            Me.Text = SharedMethods.AN ' headerText

            ' --- Set icon ---
            _ownedBitmap = New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
            Dim iconHandle As System.IntPtr = _ownedBitmap.GetHicon()
            Try
                Using borrowedIcon As System.Drawing.Icon = System.Drawing.Icon.FromHandle(iconHandle)
                    _ownedIcon = CType(borrowedIcon.Clone(), System.Drawing.Icon)
                End Using
            Finally
                DestroyProgressIcon(iconHandle)
            End Try
            Me.Icon = _ownedIcon

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
        ''' Lifts the ownerless progress dialog above the host window. It runs on a
        ''' dedicated background STA thread, so TopMost alone is defeated by Windows'
        ''' foreground lock; the AllowSetForegroundWindow + SetForegroundWindow
        ''' handshake raises it deterministically without a cross-thread owner.
        ''' </summary>
        Protected Overrides Sub OnShown(e As System.EventArgs)
            MyBase.OnShown(e)
            SharedMethods.ForceDialogToForeground(Me)

            ' Watchdog: if the host process (Word) raises its own native prompt
            ' (e.g. "Save changes?" while the user closes Word) it can gain focus yet
            ' stay hidden behind this TopMost dialog, deadlocking the UI. A short poll
            ' reliably detects that same-process foreground window and drops this
            ' dialog out of the topmost band so the prompt becomes visible. Does not
            ' change TopMost or owner behavior; self-disposes on FormClosed.
            SharedMethods.AttachForeignForegroundWatchdog(Me)
        End Sub

        Protected Overrides Sub Dispose(disposing As System.Boolean)
            Try
                If disposing AndAlso uiTimer IsNot Nothing Then
                    uiTimer.Stop()
                    RemoveHandler uiTimer.Tick, AddressOf Timer_Tick
                    uiTimer.Dispose()
                    uiTimer = Nothing
                End If
            Finally
                Try
                    MyBase.Dispose(disposing)
                Finally
                    If disposing Then
                        If _ownedIcon IsNot Nothing Then
                            _ownedIcon.Dispose()
                            _ownedIcon = Nothing
                        End If
                        If _ownedBitmap IsNot Nothing Then
                            _ownedBitmap.Dispose()
                            _ownedBitmap = Nothing
                        End If
                        If _ownedFont IsNot Nothing Then
                            _ownedFont.Dispose()
                            _ownedFont = Nothing
                        End If
                    End If
                End Try
            End Try
        End Sub

        <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
        Private Shared Function DestroyProgressIcon(iconHandle As System.IntPtr) As System.Boolean
        End Function

        Protected Overrides Sub OnFormClosed(e As System.Windows.Forms.FormClosedEventArgs)
            If uiTimer IsNot Nothing Then uiTimer.Stop()
            ProgressBarModule.CancelOperation = True
            MyBase.OnFormClosed(e)
        End Sub
    End Class


End Namespace
