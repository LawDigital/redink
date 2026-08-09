' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: LogWindow.vb
' PURPOSE
'   Modeless tooling log UI used by the tooling pipeline (e.g., ExecuteToolingLoop)
'   to display real-time progress and allow user cancellation.
'
' KEY ARCHITECTURE
'   • UI: RichTextBox log + button strip (Cancel/Close, Keep Open, Copy, Clear, Abort Job)
'   • Threading: Safe cross-thread updates via InvokeRequired/BeginInvoke
'   • Lifecycle: CancelRequested event; MarkComplete switches Cancel → Close
'   • AbortJobRequested: Aborts only the current job without stopping the session
'   • Auto-close: Optional countdown timer with "Keep Open" override
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Drawing
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary

''' <summary>
''' Modeless tooling log window used to display real-time tool execution output.
''' </summary>
''' <remarks>
''' This window is designed to be called from background tooling tasks:
''' <list type="bullet">
'''   <item><description><see cref="AppendLog"/> is thread-safe (marshals to the UI thread).</description></item>
'''   <item><description><see cref="CancelRequested"/> is raised when the user clicks Cancel (stops the entire session).</description></item>
'''   <item><description><see cref="AbortJobRequested"/> is raised when the user clicks Abort Job (aborts only the current job).</description></item>
'''   <item><description><see cref="MarkComplete"/> switches the Cancel button to Close and optionally starts an auto-close countdown.</description></item>
''' </list>
''' </remarks>
Public Class LogWindow
    Inherits Form

    ''' <summary>Rich text area used as the log output surface.</summary>
    Private ReadOnly rtbLog As RichTextBox

    ''' <summary>
    ''' Cancel/Close button. Initially raises <see cref="CancelRequested"/>; becomes Close after <see cref="MarkComplete"/>.
    ''' </summary>
    Private ReadOnly btnCancel As Button

    ''' <summary>Copies the full log text to the clipboard.</summary>
    Private ReadOnly btnCopy As Button

    ''' <summary>Clears the current log output.</summary>
    Private ReadOnly btnClear As Button

    ''' <summary>
    ''' Cancels the auto-close countdown and keeps the window open.
    ''' Visible only while the countdown is running.
    ''' </summary>
    Private ReadOnly btnKeepOpen As Button

    ''' <summary>
    ''' Aborts only the currently processing job without stopping the entire session.
    ''' Hidden by default; shown via <see cref="ShowAbortJobButton"/>.
    ''' </summary>
    Private ReadOnly btnAbortJob As Button

    ''' <summary>Tracks whether the initial bottom-right placement was already applied.</summary>
    Private _initialPositionSet As Boolean = False

    ''' <summary>Timer used to implement the auto-close countdown.</summary>
    Private _autoCloseTimer As Timer = Nothing

    ''' <summary>Seconds remaining in the current auto-close countdown.</summary>
    Private _autoCloseSecondsRemaining As Integer = 0

    ''' <summary>
    ''' Base text used for the primary button. Starts as "Close" but the actual button may show a countdown suffix.
    ''' </summary>
    Private _closeButtonBaseText As String = "Close"
    Private _restoringPlacement As Boolean = False

    ''' <summary>
    ''' Raised when the user requests cancellation of the current tooling run.
    ''' </summary>
    ''' <remarks>
    ''' The tooling loop should subscribe to this event and abort as soon as possible.
    ''' </remarks>
    Public Event CancelRequested As EventHandler

    ''' <summary>
    ''' Raised when the user requests abortion of only the current job.
    ''' </summary>
    ''' <remarks>
    ''' AutoPilot subscribes to this event to cancel the per-job CancellationTokenSource
    ''' without stopping the entire session. The processing pump continues with the next mail.
    ''' </remarks>
    Public Event AbortJobRequested As EventHandler

    ''' <summary>
    ''' Starts (or restarts) the auto-close countdown and updates the primary button text.
    ''' </summary>
    ''' <param name="seconds">
    ''' Countdown duration in seconds. Values &lt;= 0 are coerced to 1.
    ''' </param>
    ''' <remarks>
    ''' Thread-safe: marshals to the UI thread if called from a background thread.
    ''' Shows <see cref="btnKeepOpen"/> while the countdown is active.
    ''' </remarks>
    Public Sub StartAutoCloseCountdown(Optional seconds As Integer = 30)
        If seconds <= 0 Then seconds = 1

        If Me.InvokeRequired Then
            If Me.IsHandleCreated Then Me.BeginInvoke(Sub() StartAutoCloseCountdown(seconds))
            Return
        End If

        _autoCloseSecondsRemaining = seconds

        If _autoCloseTimer Is Nothing Then
            _autoCloseTimer = New Timer() With {.Interval = 1000}
            AddHandler _autoCloseTimer.Tick, AddressOf AutoCloseTimerTick
        End If

        btnCancel.Text = $"{_closeButtonBaseText} ({_autoCloseSecondsRemaining}s)"
        btnKeepOpen.Visible = True
        _autoCloseTimer.Stop()
        _autoCloseTimer.Start()
    End Sub

    ''' <summary>
    ''' Decrements the countdown timer and closes the form when it reaches zero.
    ''' </summary>
    ''' <param name="sender">Timer instance.</param>
    ''' <param name="e">Event arguments.</param>
    Private Sub AutoCloseTimerTick(sender As Object, e As EventArgs)
        _autoCloseSecondsRemaining -= 1

        If _autoCloseSecondsRemaining <= 0 Then
            _autoCloseTimer.Stop()
            Try
                Me.Close()
            Catch
            End Try
            Return
        End If

        btnCancel.Text = $"{_closeButtonBaseText} ({_autoCloseSecondsRemaining}s)"
    End Sub

    ''' <summary>
    ''' Stops the auto-close countdown and resets related UI state.
    ''' </summary>
    ''' <remarks>
    ''' Resets remaining seconds, restores the primary button text, and hides <see cref="btnKeepOpen"/>.
    ''' </remarks>
    Private Sub StopAutoCloseCountdown()
        If _autoCloseTimer IsNot Nothing Then
            _autoCloseTimer.Stop()
        End If
        _autoCloseSecondsRemaining = 0
        If btnCancel IsNot Nothing Then
            btnCancel.Text = _closeButtonBaseText
        End If
        If btnKeepOpen IsNot Nothing Then
            btnKeepOpen.Visible = False
        End If
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the log window and creates all controls.
    ''' </summary>
    ''' <remarks>
    ''' The window is modeless and not shown in the taskbar. The form handle is created immediately
    ''' to ensure <see cref="Control.BeginInvoke(System.Delegate)"/> can be used even before the form is shown.
    ''' </remarks>
    Public Sub New()
        Me.Text = $"{ThisAddIn.AN} Tooling Log"
        Me.Width = 600
        Me.Height = 400
        Me.StartPosition = FormStartPosition.Manual
        Me.FormBorderStyle = FormBorderStyle.Sizable
        Me.MinimumSize = New Size(450, 300)
        Me.TopMost = False
        Me.MaximizeBox = False
        Me.MinimizeBox = True
        Me.ShowInTaskbar = True
        Me.AutoScaleMode = AutoScaleMode.Font

        ' Double buffering prevents rendering artifacts on buttons/text
        Me.DoubleBuffered = True

        Try
            Me.Icon = Icon.FromHandle((New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))).GetHicon())
        Catch
        End Try

        Dim mainPanel As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill, .ColumnCount = 1, .RowCount = 2, .Padding = New Padding(10)
        }
        mainPanel.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        mainPanel.RowStyles.Add(New RowStyle(SizeType.AutoSize))

        rtbLog = New RichTextBox() With {
            .Dock = DockStyle.Fill, .ReadOnly = True, .BorderStyle = BorderStyle.None,
            .ScrollBars = RichTextBoxScrollBars.Vertical, .WordWrap = False, .DetectUrls = False,
            .BackColor = Me.BackColor, .HideSelection = False, .ShortcutsEnabled = True,
            .Font = New Font("Consolas", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
        }
        mainPanel.Controls.Add(rtbLog, 0, 0)

        Dim buttonPanel As New FlowLayoutPanel() With {
            .Dock = DockStyle.Fill, .FlowDirection = FlowDirection.RightToLeft, .AutoSize = True
        }

        btnCancel = New Button() With {.Text = "Cancel", .AutoSize = True, .Padding = New Padding(10, 5, 10, 5)}
        AddHandler btnCancel.Click, AddressOf OnCancelClick

        btnKeepOpen = New Button() With {.Text = "Keep Open", .AutoSize = True, .Padding = New Padding(10, 5, 10, 5), .Visible = False}
        AddHandler btnKeepOpen.Click, AddressOf OnKeepOpenClick

        btnAbortJob = New Button() With {.Text = "Abort Job", .AutoSize = True, .Padding = New Padding(10, 5, 10, 5), .Visible = False}
        AddHandler btnAbortJob.Click, AddressOf OnAbortJobClick

        btnCopy = New Button() With {.Text = "Copy Log", .AutoSize = True, .Padding = New Padding(10, 5, 10, 5)}
        AddHandler btnCopy.Click, AddressOf OnCopyClick

        btnClear = New Button() With {.Text = "Clear", .AutoSize = True, .Padding = New Padding(10, 5, 10, 5)}
        AddHandler btnClear.Click, AddressOf OnClearClick

        buttonPanel.Controls.Add(btnCancel)
        buttonPanel.Controls.Add(btnKeepOpen)
        buttonPanel.Controls.Add(btnAbortJob)
        buttonPanel.Controls.Add(btnCopy)
        buttonPanel.Controls.Add(btnClear)
        mainPanel.Controls.Add(buttonPanel, 0, 1)
        Me.Controls.Add(mainPanel)

        ' Force handle creation immediately. This ensures BeginInvoke works
        ' even if called before the window is fully Shown.
        Me.CreateControl()
    End Sub

    ''' <summary>
    ''' Applies the initial window placement and performs an initial refresh.
    ''' </summary>
    ''' <param name="e">Event arguments.</param>
    ''' <remarks>
    ''' Positions the window near the bottom-right of the primary screen working area once per instance.
    ''' </remarks>
    Protected Overrides Sub OnShown(e As EventArgs)
        MyBase.OnShown(e)
        If Not _initialPositionSet Then
            RestoreWindowPlacement()
        End If
        Me.BringToFront()
        Me.Refresh() ' Ensure initial paint
    End Sub

    Protected Overrides Sub OnMove(e As EventArgs)
        MyBase.OnMove(e)

        If Me.WindowState = FormWindowState.Normal Then
            SaveWindowPlacement()
        End If
    End Sub

    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        SaveWindowPlacement()
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        SaveWindowPlacement()
        MyBase.OnFormClosing(e)
    End Sub

    Private Sub RestoreWindowPlacement()
        _restoringPlacement = True

        Try
            Dim savedWidth As Integer = My.Settings.ToolingLogWidth
            Dim savedHeight As Integer = My.Settings.ToolingLogHeight
            Dim hasSavedBounds As Boolean = savedWidth > 0 AndAlso savedHeight > 0

            If hasSavedBounds Then
                Dim restoredBounds As Rectangle =
                    GetSafePlacementBounds(
                        My.Settings.ToolingLogX,
                        My.Settings.ToolingLogY,
                        savedWidth,
                        savedHeight)

                Me.SetBounds(
                    restoredBounds.X,
                    restoredBounds.Y,
                    restoredBounds.Width,
                    restoredBounds.Height)
            Else
                PositionWindowDefault()
            End If

            If My.Settings.ToolingLogMaximized Then
                Me.WindowState = FormWindowState.Maximized
            ElseIf My.Settings.ToolingLogMinimized Then
                Me.WindowState = FormWindowState.Minimized
            Else
                Me.WindowState = FormWindowState.Normal
            End If
        Catch
            PositionWindowDefault()
            Me.WindowState = FormWindowState.Normal
        Finally
            _initialPositionSet = True
            _restoringPlacement = False
        End Try
    End Sub

    Private Sub PositionWindowDefault()
        Dim wa As Rectangle = Screen.PrimaryScreen.WorkingArea
        Dim width As Integer = Math.Min(Me.Width, wa.Width)
        Dim height As Integer = Math.Min(Me.Height, wa.Height)

        Me.SetBounds(
            Math.Max(wa.Left, wa.Right - width - 40),
            Math.Max(wa.Top, wa.Bottom - height - 40),
            width,
            height)
    End Sub

    Private Function GetSafePlacementBounds(savedX As Integer,
                                            savedY As Integer,
                                            savedWidth As Integer,
                                            savedHeight As Integer) As Rectangle

        Dim width As Integer = If(savedWidth > 0, savedWidth, Me.Width)
        Dim height As Integer = If(savedHeight > 0, savedHeight, Me.Height)

        Dim probe As New Rectangle(savedX, savedY, Math.Max(1, width), Math.Max(1, height))
        Dim area As Rectangle = Screen.FromRectangle(probe).WorkingArea

        Dim minWidth As Integer = Math.Min(Me.MinimumSize.Width, area.Width)
        Dim minHeight As Integer = Math.Min(Me.MinimumSize.Height, area.Height)

        width = Math.Max(minWidth, Math.Min(width, area.Width))
        height = Math.Max(minHeight, Math.Min(height, area.Height))

        Dim x As Integer = Math.Max(area.Left, Math.Min(savedX, area.Right - width))
        Dim y As Integer = Math.Max(area.Top, Math.Min(savedY, area.Bottom - height))

        Return New Rectangle(x, y, width, height)
    End Function

    Private Sub SaveWindowPlacement()
        If _restoringPlacement OrElse Me.IsDisposed OrElse Not Me.IsHandleCreated Then
            Return
        End If

        Try
            Dim boundsToSave As Rectangle =
                If(Me.WindowState = FormWindowState.Normal,
                   Me.Bounds,
                   Me.RestoreBounds)

            If boundsToSave.Width > 0 AndAlso boundsToSave.Height > 0 Then
                My.Settings.ToolingLogX = boundsToSave.X
                My.Settings.ToolingLogY = boundsToSave.Y
                My.Settings.ToolingLogWidth = boundsToSave.Width
                My.Settings.ToolingLogHeight = boundsToSave.Height
            End If

            My.Settings.ToolingLogMinimized = (Me.WindowState = FormWindowState.Minimized)
            My.Settings.ToolingLogMaximized = (Me.WindowState = FormWindowState.Maximized)
            My.Settings.Save()
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Appends a line to the log output (timestamped and color-coded by level).
    ''' </summary>
    ''' <param name="message">Message text to append.</param>
    ''' <param name="level">
    ''' Optional severity/category for color selection (e.g., "info", "warn", "error", "success", "step", "llm").
    ''' </param>
    ''' <remarks>
    ''' Thread-safe: uses <see cref="Control.BeginInvoke(System.Delegate)"/> when called off the UI thread.
    ''' Messages are ignored if the form or log control has been disposed.
    ''' </remarks>
    Public Sub AppendLog(message As String, Optional level As String = "info")
        If Me.IsDisposed OrElse rtbLog.IsDisposed Then Return

        If Me.InvokeRequired Then
            ' Check handle again before invoking to avoid race conditions during close
            If Me.IsHandleCreated Then
                Me.BeginInvoke(Sub() AppendLogInternal(message, level))
            End If
        Else
            AppendLogInternal(message, level)
        End If
    End Sub

    ''' <summary>
    ''' Performs the actual append work on the UI thread.
    ''' </summary>
    ''' <param name="message">Message text.</param>
    ''' <param name="level">Severity/category used for choosing the text color.</param>
    ''' <remarks>
    ''' Uses <see cref="Control.SuspendLayout"/> to reduce flicker during frequent updates.
    ''' </remarks>
    Private Sub AppendLogInternal(message As String, level As String)
        If String.IsNullOrEmpty(message) Then Return
        If rtbLog.IsDisposed Then Return

        Dim displayMessage As String = NormalizeLogMessage(message)
        If displayMessage.Length = 0 Then Return

        Dim timestamp = DateTime.Now.ToString("HH:mm:ss")
        Dim textColor As Color = Color.Black
        Select Case (If(level, "info")).ToLowerInvariant()
            Case "error", "err", "fail" : textColor = Color.DarkRed
            Case "warn", "warning" : textColor = Color.DarkOrange
            Case "success", "ok" : textColor = Color.DarkGreen
            Case "step" : textColor = Color.DarkBlue
            Case "llm" : textColor = Color.DarkMagenta
        End Select

        rtbLog.SuspendLayout()
        Try
            rtbLog.SelectionStart = rtbLog.TextLength
            rtbLog.SelectionColor = Color.Gray
            rtbLog.AppendText($"[{timestamp}] ")
            rtbLog.SelectionStart = rtbLog.TextLength
            rtbLog.SelectionColor = textColor
            rtbLog.AppendText(displayMessage & Environment.NewLine)
            rtbLog.SelectionStart = rtbLog.TextLength
            rtbLog.ScrollToCaret()
        Finally
            rtbLog.ResumeLayout()
        End Try
    End Sub

    Private Shared Function NormalizeLogMessage(message As String) As String
        Dim text As String = If(message, "")
        text = text.Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ").Replace(vbTab, " ").Trim()

        Do While text.Contains("  ")
            text = text.Replace("  ", " ")
        Loop

        Return text
    End Function

    ''' <summary>
    ''' Handles the Cancel button click and raises <see cref="CancelRequested"/>.
    ''' </summary>
    ''' <param name="sender">Event sender.</param>
    ''' <param name="e">Event arguments.</param>
    ''' <remarks>
    ''' Disables the button immediately to prevent duplicate cancel requests.
    ''' </remarks>
    Private Sub OnCancelClick(sender As Object, e As EventArgs)
        btnCancel.Enabled = False
        btnCancel.Text = "Cancelling..."
        RaiseEvent CancelRequested(Me, EventArgs.Empty)
    End Sub

    ''' <summary>
    ''' Cancels the auto-close countdown and keeps the window open.
    ''' </summary>
    ''' <param name="sender">Event sender.</param>
    ''' <param name="e">Event arguments.</param>
    Private Sub OnKeepOpenClick(sender As Object, e As EventArgs)
        StopAutoCloseCountdown()
        AppendLog("Auto-close cancelled - window will remain open", "step")
    End Sub

    ''' <summary>
    ''' Handles the Abort Job button click and raises <see cref="AbortJobRequested"/>.
    ''' </summary>
    ''' <param name="sender">Event sender.</param>
    ''' <param name="e">Event arguments.</param>
    ''' <remarks>
    ''' Temporarily disables the button for 3 seconds to prevent rapid repeated clicks,
    ''' then re-enables it for the next job. Unlike Cancel, this does not permanently
    ''' disable the button because new jobs will follow.
    ''' </remarks>
    Private Sub OnAbortJobClick(sender As Object, e As EventArgs)
        btnAbortJob.Enabled = False
        btnAbortJob.Text = "Aborting..."
        RaiseEvent AbortJobRequested(Me, EventArgs.Empty)

        ' Re-enable after a short delay so the button is ready for the next job
        Dim reEnableTimer As New Timer() With {.Interval = 3000}
        AddHandler reEnableTimer.Tick, Sub(s As Object, ev As EventArgs)
                                           reEnableTimer.Stop()
                                           reEnableTimer.Dispose()
                                           If Not btnAbortJob.IsDisposed Then
                                               btnAbortJob.Enabled = True
                                               btnAbortJob.Text = "Abort Job"
                                           End If
                                       End Sub
        reEnableTimer.Start()
    End Sub

    ''' <summary>
    ''' Shows or hides the Abort Job button.
    ''' </summary>
    ''' <param name="visible">True to show, False to hide.</param>
    ''' <remarks>
    ''' Thread-safe: marshals to the UI thread if called from a background thread.
    ''' Used by AutoPilot to expose per-job abort functionality on the dashboard.
    ''' </remarks>
    Public Sub ShowAbortJobButton(visible As Boolean)
        If Me.IsDisposed Then Return

        If Me.InvokeRequired Then
            If Me.IsHandleCreated Then Me.BeginInvoke(Sub() ShowAbortJobButton(visible))
            Return
        End If

        btnAbortJob.Visible = visible
    End Sub

    ''' <summary>
    ''' Copies the current log contents to the clipboard.
    ''' </summary>
    ''' <param name="sender">Event sender.</param>
    ''' <param name="e">Event arguments.</param>
    Private Sub OnCopyClick(sender As Object, e As EventArgs)
        If Not String.IsNullOrEmpty(rtbLog.Text) Then Clipboard.SetText(rtbLog.Text)
    End Sub

    ''' <summary>
    ''' Clears the log output.
    ''' </summary>
    ''' <param name="sender">Event sender.</param>
    ''' <param name="e">Event arguments.</param>
    Private Sub OnClearClick(sender As Object, e As EventArgs)
        rtbLog.Clear()
    End Sub

    ''' <summary>
    ''' Marks the tooling operation as completed and switches the primary button from Cancel to Close.
    ''' </summary>
    ''' <remarks>
    ''' Thread-safe: may be called from background threads. After completion, the window title is updated and an
    ''' auto-close countdown is started using <c>ThisAddIn.ToolingLog_AutoCloseDefaultSeconds</c>.
    ''' </remarks>
    Public Sub MarkComplete()
        If Me.IsDisposed Then Return

        If Me.InvokeRequired Then
            If Me.IsHandleCreated Then Me.BeginInvoke(Sub() MarkCompleteInternal())
        Else
            MarkCompleteInternal()
        End If
    End Sub

    ''' <summary>
    ''' Applies completed-state UI changes on the UI thread.
    ''' </summary>
    ''' <remarks>
    ''' Stops any current countdown, rewires the Cancel button to Close the form, updates the window title,
    ''' hides the Abort Job button, and starts the default auto-close countdown.
    ''' </remarks>
    Private Sub MarkCompleteInternal()
        If btnCancel.IsDisposed Then Return

        StopAutoCloseCountdown()

        btnCancel.Text = "Close"
        _closeButtonBaseText = "Close"
        btnCancel.Enabled = True
        RemoveHandler btnCancel.Click, AddressOf OnCancelClick
        AddHandler btnCancel.Click, Sub()
                                        StopAutoCloseCountdown()
                                        Me.Close()
                                    End Sub

        ' Hide abort job button — session is complete
        If Not btnAbortJob.IsDisposed Then
            btnAbortJob.Visible = False
        End If

        Me.Text = $"{ThisAddIn.AN} Tooling Log (Complete)"

        StartAutoCloseCountdown(ThisAddIn.ToolingLog_AutoCloseDefaultSeconds)
    End Sub

End Class
