' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: LogWindow.vb
' Purpose: Provides a persistent log window that uses WinForms autoscaling
'          (AutoScaleMode.Font) and displays scrolling log entries without flashing.
'
' Architecture:
'  - Public API (Module LogWindow): ShowLogWindow / HideLogWindow / CloseLogWindow /
'    AppendLog / SetTitle / ClearLog.
'  - Thread-safety: All access to the cached form instance is guarded by SyncLock.
'    UI updates are marshaled to the form thread via InvokeOnForm.
'  - Lifetime: The underlying form is created on demand (EnsureFormCreated) and is
'    typically hidden (not disposed) when the user closes it (OnFormClosing).
'  - Close policy: Optional OnCloseRequested callback can veto the user-initiated close.
'  - Non-activation: The form does not take focus when shown/clicked
'    (ShowWithoutActivation + WM_MOUSEACTIVATE).
'  - UI: RichTextBox used for log display; context menu and Ctrl+A/C supported.
'    Entries are timestamped and color-coded by level.
'  - Initial placement: On first show, the window is positioned bottom-right within
'    the primary screen working area with fixed margins.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Namespace SharedLibrary

    ''' <summary>
    ''' Manages a persistent log window for execution feedback.
    ''' Thread-safe. The underlying form uses WinForms autoscaling via <see cref="Form.AutoScaleMode"/>.
    ''' </summary>
    Public Module LogWindow

        ''' <summary>
        ''' Cached instance of the log window form. Created on demand and typically hidden (not disposed).
        ''' </summary>
        Private _logForm As LogWindowForm = Nothing

        ''' <summary>
        ''' Synchronizes access to <see cref="_logForm"/> and ensures thread-safe updates.
        ''' </summary>
        Private ReadOnly _syncLock As New Object()

        ''' <summary>
        ''' Callback invoked when the user attempts to close the log window (e.g., via the X button or the Close button).
        ''' Return True to allow the close action to proceed (the window is hidden); return False to cancel the close action.
        ''' </summary>
        Public Property OnCloseRequested As Func(Of Boolean)

        ''' <summary>
        ''' Shows the log window (creates it if necessary) and optionally clears previous content.
        ''' </summary>
        ''' <param name="clearOnShow">If True, clears previous log entries when showing.</param>
        Public Sub ShowLogWindow(Optional clearOnShow As Boolean = True)
            SyncLock _syncLock
                EnsureFormCreated()
                If _logForm IsNot Nothing Then
                    InvokeOnForm(Sub()
                                     If clearOnShow Then
                                         _logForm.ClearLog()
                                     End If
                                     If Not _logForm.Visible Then
                                         _logForm.Show()
                                     End If
                                     _logForm.BringToFront()
                                 End Sub)
                End If
            End SyncLock
        End Sub

        ''' <summary>
        ''' Hides the log window without destroying it.
        ''' </summary>
        Public Sub HideLogWindow()
            SyncLock _syncLock
                If _logForm IsNot Nothing AndAlso _logForm.Visible Then
                    InvokeOnForm(Sub() _logForm.Hide())
                End If
            End SyncLock
        End Sub

        ''' <summary>
        ''' Closes and disposes the log window.
        ''' </summary>
        Public Sub CloseLogWindow()
            SyncLock _syncLock
                If _logForm IsNot Nothing Then
                    InvokeOnForm(Sub()
                                     _logForm.Close()
                                     _logForm.Dispose()
                                 End Sub)
                    _logForm = Nothing
                End If
            End SyncLock
        End Sub

        ''' <summary>
        ''' Appends a log entry to the window. The window scrolls to show the latest entry.
        ''' </summary>
        ''' <param name="message">The message to append.</param>
        ''' <param name="level">Optional log level for color coding (info, warn, error, success).</param>
        Public Sub AppendLog(message As String, Optional level As String = "info")
            SyncLock _syncLock
                EnsureFormCreated()
                If _logForm IsNot Nothing Then
                    InvokeOnForm(Sub() _logForm.AppendLogEntry(message, level))
                End If
            End SyncLock
        End Sub

        ''' <summary>
        ''' Updates the window title.
        ''' </summary>
        ''' <param name="title">New title text. If empty/whitespace, defaults to the application name.</param>
        Public Sub SetTitle(title As String)
            SyncLock _syncLock
                If _logForm IsNot Nothing Then
                    InvokeOnForm(Sub() _logForm.Text = If(String.IsNullOrWhiteSpace(title), SharedMethods.AN, title))
                End If
            End SyncLock
        End Sub

        ''' <summary>
        ''' Clears all log entries from the window.
        ''' </summary>
        Public Sub ClearLog()
            SyncLock _syncLock
                If _logForm IsNot Nothing Then
                    InvokeOnForm(Sub() _logForm.ClearLog())
                End If
            End SyncLock
        End Sub

        ''' <summary>
        ''' Ensures the log window form exists. If a main UI form exists, creation is marshaled to its UI thread.
        ''' </summary>
        Private Sub EnsureFormCreated()
            If _logForm Is Nothing OrElse _logForm.IsDisposed Then
                ' Must create on UI thread
                If Application.OpenForms.Count > 0 Then
                    Dim mainForm = Application.OpenForms(0)
                    If mainForm.InvokeRequired Then
                        mainForm.Invoke(Sub() _logForm = New LogWindowForm())
                    Else
                        _logForm = New LogWindowForm()
                    End If
                Else
                    _logForm = New LogWindowForm()
                End If
            End If
        End Sub

        ''' <summary>
        ''' Executes <paramref name="action"/> on the log window UI thread.
        ''' No-op if the form instance is missing or disposed.
        ''' </summary>
        ''' <param name="action">Action to execute on the form thread.</param>
        Private Sub InvokeOnForm(action As Action)
            If _logForm Is Nothing OrElse _logForm.IsDisposed Then Return
            Try
                If _logForm.InvokeRequired Then
                    _logForm.Invoke(action)
                Else
                    action()
                End If
            Catch ex As ObjectDisposedException
                ' Form was disposed between check and invoke
            Catch ex As InvalidOperationException
                ' Handle edge cases during form lifecycle
            End Try
        End Sub

        ''' <summary>
        ''' Internal form class for the log window.
        ''' </summary>
        Private Class LogWindowForm
            Inherits Form

            ''' <summary>
            ''' Log output control.
            ''' </summary>
            Private _logBox As RichTextBox

            ''' <summary>
            ''' Container for action buttons at the bottom of the window.
            ''' </summary>
            Private _bottomPanel As Panel

            ''' <summary>
            ''' Button that triggers a close request (typically hides the window).
            ''' </summary>
            Private _closeButton As Button

            ''' <summary>
            ''' Button that clears the log contents.
            ''' </summary>
            Private _clearButton As Button

            ''' <summary>
            ''' Button that copies the entire log contents to the clipboard.
            ''' </summary>
            Private _copyButton As Button

            ''' <summary>
            ''' Tracks whether the initial on-screen position has already been set.
            ''' </summary>
            Private _initialPositionSet As Boolean = False

            ''' <summary>
            ''' Standard UI font used by this form.
            ''' </summary>
            Private ReadOnly _standardFont As New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)

            ''' <summary>
            ''' Monospace font used for log output.
            ''' </summary>
            Private ReadOnly _monoFont As New Font("Consolas", 9.0F, FontStyle.Regular, GraphicsUnit.Point)

            ''' <summary>
            ''' Creates a new instance of the log window form and initializes its UI.
            ''' </summary>
            Public Sub New()
                InitializeComponent()
            End Sub

            ''' <summary>
            ''' Builds the UI and wires event handlers for the log window.
            ''' </summary>
            Private Sub InitializeComponent()
                Me.SuspendLayout()

                ' Form settings - non-modal, user can continue working
                Me.Text = SharedMethods.AN
                Me.StartPosition = FormStartPosition.Manual
                Me.FormBorderStyle = FormBorderStyle.Sizable
                Me.MinimumSize = New Size(450, 300)
                Me.Size = New Size(600, 400)
                Me.ShowInTaskbar = False
                Me.TopMost = False
                Me.MaximizeBox = False
                Me.MinimizeBox = False
                Me.Font = _standardFont
                Me.AutoScaleMode = AutoScaleMode.Font
                Me.ShowIcon = True

                ' (1) Use logo as icon (if available).
                Try
                    Using bmpIcon As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                        Dim hIcon As IntPtr = bmpIcon.GetHicon()
                        Try
                            Using ico As Icon = Icon.FromHandle(hIcon)
                                Me.Icon = CType(ico.Clone(), Icon)
                            End Using
                        Finally
                            DestroyIcon(hIcon)
                        End Try
                    End Using
                Catch
                    ' Ignore if resource not available
                End Try

                ' Do not grab focus automatically when shown
                Me.TabStop = False

                ' Keep key input handling within this form (without activating it on show/click).
                Me.KeyPreview = True

                ' (2) Bottom panel with buttons - increased height for more bottom padding
                _bottomPanel = New Panel() With {
                    .Dock = DockStyle.Bottom,
                    .Height = 70,
                    .Padding = New Padding(12, 8, 12, 24)
                }

                _closeButton = New Button() With {
                    .Text = "Close",
                    .AutoSize = True,
                    .Font = _standardFont,
                    .Padding = New Padding(8, 4, 8, 4)
                }
                AddHandler _closeButton.Click, Sub(s, e) HandleCloseRequest()

                _clearButton = New Button() With {
                    .Text = "Clear",
                    .AutoSize = True,
                    .Font = _standardFont,
                    .Padding = New Padding(8, 4, 8, 4)
                }
                AddHandler _clearButton.Click, Sub(s, e) ClearLog()

                _copyButton = New Button() With {
                    .Text = "Copy",
                    .AutoSize = True,
                    .Font = _standardFont,
                    .Padding = New Padding(8, 4, 8, 4)
                }
                AddHandler _copyButton.Click, Sub(s, e) CopyLogToClipboard()

                ' Position buttons using a layout container to avoid AutoSize/DPI timing issues.
                Dim buttonFlow As New FlowLayoutPanel() With {
                    .Dock = DockStyle.Top,
                    .AutoSize = True,
                    .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    .WrapContents = False,
                    .Margin = Padding.Empty,
                    .Padding = Padding.Empty,
                    .FlowDirection = FlowDirection.LeftToRight
                }
                buttonFlow.Controls.Add(_closeButton)
                buttonFlow.Controls.Add(_clearButton)
                buttonFlow.Controls.Add(_copyButton)
                _bottomPanel.Controls.Add(buttonFlow)

                ' Log text box - allow mouse selection + right-click copy/paste
                _logBox = New RichTextBox() With {
                    .Dock = DockStyle.Fill,
                    .ReadOnly = True,
                    .Font = _monoFont,
                    .BorderStyle = BorderStyle.None,
                    .ScrollBars = RichTextBoxScrollBars.Vertical,
                    .WordWrap = True,
                    .DetectUrls = False,
                    .BackColor = Me.BackColor,
                    .HideSelection = False,
                    .ShortcutsEnabled = True
                }

                ' Context menu to allow mouse-driven copy
                Dim cms As New ContextMenuStrip()
                Dim miCopy As New ToolStripMenuItem("Copy")
                AddHandler miCopy.Click, Sub() CopySelectionOrAllToClipboard()
                cms.Items.Add(miCopy)

                Dim miSelectAll As New ToolStripMenuItem("Select All")
                AddHandler miSelectAll.Click, Sub() _logBox.SelectAll()
                cms.Items.Add(miSelectAll)

                _logBox.ContextMenuStrip = cms

                ' Keyboard shortcuts (Ctrl+C / Ctrl+A)
                AddHandler _logBox.KeyDown,
                    Sub(sender As Object, e As KeyEventArgs)
                        If e.Control AndAlso e.KeyCode = Keys.A Then
                            _logBox.SelectAll()
                            e.Handled = True
                            e.SuppressKeyPress = True
                            Return
                        End If
                        If e.Control AndAlso e.KeyCode = Keys.C Then
                            CopySelectionOrAllToClipboard()
                            e.Handled = True
                            e.SuppressKeyPress = True
                            Return
                        End If
                    End Sub

                ' Wrapper panel for log box with padding
                Dim logPanel As New Panel() With {
                    .Dock = DockStyle.Fill,
                    .Padding = New Padding(12, 12, 12, 0)
                }
                _logBox.Dock = DockStyle.Fill
                logPanel.Controls.Add(_logBox)

                Me.Controls.Add(logPanel)
                Me.Controls.Add(_bottomPanel)

                Me.ResumeLayout(False)
            End Sub

            ''' <summary>
            ''' Prevents the form from activating when shown.
            ''' </summary>
            Protected Overrides ReadOnly Property ShowWithoutActivation As Boolean
                Get
                    Return True
                End Get
            End Property

            ''' <summary>
            ''' Prevents activation when the user clicks the form by handling <c>WM_MOUSEACTIVATE</c>.
            ''' </summary>
            ''' <param name="m">The Windows message to process.</param>
            Protected Overrides Sub WndProc(ByRef m As Message)
                Const WM_MOUSEACTIVATE As Integer = &H21
                Const MA_NOACTIVATE As Integer = 3

                If m.Msg = WM_MOUSEACTIVATE Then
                    m.Result = CType(MA_NOACTIVATE, IntPtr)
                    Return
                End If

                MyBase.WndProc(m)
            End Sub

            ''' <summary>
            ''' Positions window at bottom-right on initial show only.
            ''' </summary>
            Private Sub PositionWindowInitial()
                ' (3) Bottom-right of the working area with 40px margin from right and bottom
                Const marginRight As Integer = 40
                Const marginBottom As Integer = 40

                Dim wa = Screen.PrimaryScreen.WorkingArea
                Me.Location = New Point(
                    wa.Right - Me.Width - marginRight,
                    wa.Bottom - Me.Height - marginBottom
                )
            End Sub

            ''' <summary>
            ''' Handles a close request by invoking <see cref="LogWindow.OnCloseRequested"/> (if set) and hiding the form when allowed.
            ''' </summary>
            Private Sub HandleCloseRequest()
                Dim callback = LogWindow.OnCloseRequested
                If callback IsNot Nothing Then
                    Dim allowClose = callback.Invoke()
                    If allowClose Then
                        Me.Hide()
                    End If
                Else
                    Me.Hide()
                End If
            End Sub

            ''' <summary>
            ''' Appends a timestamped, color-coded log entry and scrolls to the latest entry.
            ''' </summary>
            ''' <param name="message">The message to append.</param>
            ''' <param name="level">Log level used for color coding.</param>
            Public Sub AppendLogEntry(message As String, level As String)
                If String.IsNullOrEmpty(message) Then Return

                Dim timestamp = DateTime.Now.ToString("HH:mm:ss.fff")
                Dim prefix = $"[{timestamp}] "

                ' Determine color based on level
                Dim textColor As Color
                Select Case level.ToLowerInvariant()
                    Case "error", "err", "fail"
                        textColor = Color.DarkRed
                    Case "warn", "warning"
                        textColor = Color.DarkOrange
                    Case "success", "ok"
                        textColor = Color.DarkGreen
                    Case "step"
                        textColor = Color.DarkBlue
                    Case "llm"
                        textColor = Color.DarkMagenta
                    Case Else
                        textColor = Color.Black
                End Select

                _logBox.SuspendLayout()
                Try
                    _logBox.SelectionStart = _logBox.TextLength
                    _logBox.SelectionLength = 0
                    _logBox.SelectionColor = Color.Gray
                    _logBox.AppendText(prefix)

                    _logBox.SelectionStart = _logBox.TextLength
                    _logBox.SelectionLength = 0
                    _logBox.SelectionColor = textColor
                    _logBox.AppendText(message & Environment.NewLine)

                    _logBox.SelectionStart = _logBox.TextLength
                    _logBox.ScrollToCaret()
                Finally
                    _logBox.ResumeLayout()
                End Try
            End Sub

            ''' <summary>
            ''' Clears all text from the log display.
            ''' </summary>
            Public Sub ClearLog()
                _logBox.Clear()
            End Sub

            ''' <summary>
            ''' Copies the entire log text to the clipboard.
            ''' </summary>
            Private Sub CopyLogToClipboard()
                CopyTextToClipboard(_logBox.Text)
            End Sub

            ''' <summary>
            ''' Copies the current selection to the clipboard; if nothing is selected, copies the entire log.
            ''' </summary>
            Private Sub CopySelectionOrAllToClipboard()
                Dim txt = If(_logBox.SelectionLength > 0, _logBox.SelectedText, _logBox.Text)
                CopyTextToClipboard(txt)
            End Sub

            ''' <summary>
            ''' Copies the provided text to the clipboard using <see cref="SharedMethods.PutInClipboard"/> with a fallback to <see cref="Clipboard.SetText"/>.
            ''' </summary>
            ''' <param name="text">Text to place on the clipboard.</param>
            Private Sub CopyTextToClipboard(text As String)
                Try
                    ' Use existing library clipboard helper (STA-safe); avoids COM/OLE issues.
                    SharedMethods.PutInClipboard(If(text, String.Empty))
                Catch
                    Try
                        Clipboard.SetText(If(text, String.Empty))
                    Catch
                    End Try
                End Try
            End Sub

            ''' <summary>
            ''' Cancels user-initiated closes and hides the form instead.
            ''' </summary>
            ''' <param name="e">Form closing event args.</param>
            Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
                ' Hide instead of close when user clicks X
                If e.CloseReason = CloseReason.UserClosing Then
                    e.Cancel = True
                    HandleCloseRequest()
                Else
                    MyBase.OnFormClosing(e)
                End If
            End Sub

            ''' <summary>
            ''' Positions the window on first show and keeps it on top of other windows via <see cref="Control.BringToFront"/>.
            ''' </summary>
            ''' <param name="e">Event args.</param>
            Protected Overrides Sub OnShown(e As EventArgs)
                MyBase.OnShown(e)
                ' Position only on first show
                If Not _initialPositionSet Then
                    PositionWindowInitial()
                    _initialPositionSet = True
                End If
                Me.BringToFront()
            End Sub

            <DllImport("user32.dll", SetLastError:=True)>
            Private Shared Function DestroyIcon(hIcon As IntPtr) As Boolean
            End Function
        End Class

    End Module

End Namespace