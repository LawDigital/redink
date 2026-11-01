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
'  - Thread-safety: Producers enqueue actions; an explicitly captured Office UI
'    dispatcher drains them in order. No lock is held across UI work or posting.
'  - Lifetime: Create/show/dispose only on that Office UI thread, registered in
'    its HostModelessFormRegistry. User close still hides; host shutdown disposes.
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

        ' Only the host UI thread touches _logForm and the registry. The lock
        ' protects queue/dispatcher metadata, never a Form or a UI invocation.
        Private _logForm As LogWindowForm = Nothing
        Private ReadOnly _syncLock As New System.Object()
        Private ReadOnly _pending As New System.Collections.Generic.Queue(Of System.Action)()
        Private _uiContext As System.Threading.SynchronizationContext
        Private _uiThreadId As System.Int32
        Private _hostRegistry As HostModelessFormRegistry
        Private _drainPosted As System.Boolean
        Private _shuttingDown As System.Boolean
        Private _closeCallback As System.Func(Of System.Boolean)

        ''' <summary>Called once by the owning Office host on its UI thread.</summary>
        Public Sub InitializeUiDispatcher(context As System.Threading.SynchronizationContext,
                                          registry As HostModelessFormRegistry)
            If context Is Nothing Then Throw New System.ArgumentNullException(NameOf(context))
            If registry Is Nothing Then Throw New System.ArgumentNullException(NameOf(registry))
            If System.Threading.Thread.CurrentThread.GetApartmentState() <> System.Threading.ApartmentState.STA Then
                Throw New System.InvalidOperationException("The log window dispatcher must be initialized on the Office UI thread.")
            End If
            SharedMethods.InitializeHostUiDispatcher(context, registry)
            SyncLock _syncLock
                If _shuttingDown Then Throw New System.InvalidOperationException("The Office host is already closing.")
                Dim currentThreadId As System.Int32 = System.Threading.Thread.CurrentThread.ManagedThreadId
                If _uiContext IsNot Nothing AndAlso (_uiThreadId <> currentThreadId OrElse _uiContext IsNot context OrElse _hostRegistry IsNot registry) Then
                    Throw New System.InvalidOperationException("The log window is already assigned to another UI dispatcher.")
                End If
                _uiContext = context
                _uiThreadId = currentThreadId
                _hostRegistry = registry
            End SyncLock
        End Sub

        Public Property OnCloseRequested As System.Func(Of System.Boolean)
            Get
                SyncLock _syncLock
                    Return _closeCallback
                End SyncLock
            End Get
            Set(value As System.Func(Of System.Boolean))
                SyncLock _syncLock
                    If Not _shuttingDown Then _closeCallback = value
                End SyncLock
            End Set
        End Property

        Public Sub ShowLogWindow(Optional clearOnShow As System.Boolean = True)
            EnqueueOnUi(
                Sub()
                    EnsureFormCreatedOnUi()
                    If _logForm Is Nothing Then Return
                    If clearOnShow Then _logForm.ClearLog()
                    _hostRegistry.Show(_logForm)
                    _logForm.BringToFront()
                End Sub)
        End Sub

        Public Sub HideLogWindow()
            EnqueueOnUi(Sub()
                            If _logForm IsNot Nothing AndAlso Not _logForm.IsDisposed Then _logForm.Hide()
                        End Sub)
        End Sub

        Public Sub CloseLogWindow()
            EnqueueOnUi(AddressOf DisposeFormOnUi)
        End Sub

        Public Sub AppendLog(message As System.String, Optional level As System.String = "info")
            EnqueueOnUi(Sub()
                            EnsureFormCreatedOnUi()
                            If _logForm IsNot Nothing Then _logForm.AppendLogEntry(message, level)
                        End Sub)
        End Sub

        Public Sub SetTitle(title As System.String)
            EnqueueOnUi(Sub()
                            If _logForm IsNot Nothing AndAlso Not _logForm.IsDisposed Then
                                _logForm.Text = If(System.String.IsNullOrWhiteSpace(title), SharedMethods.AN, title)
                            End If
                        End Sub)
        End Sub

        Public Sub ClearLog()
            EnqueueOnUi(Sub()
                            If _logForm IsNot Nothing AndAlso Not _logForm.IsDisposed Then _logForm.ClearLog()
                        End Sub)
        End Sub

        Private Sub EnqueueOnUi(action As System.Action)
            Dim dispatcher As System.Threading.SynchronizationContext = Nothing
            Dim missingDispatcher As System.Boolean
            SyncLock _syncLock
                If _shuttingDown Then Return
                missingDispatcher = _uiContext Is Nothing
                If Not missingDispatcher Then
                    _pending.Enqueue(action)
                    If Not _drainPosted Then
                        _drainPosted = True
                        dispatcher = _uiContext
                    End If
                End If
            End SyncLock
            If missingDispatcher Then
                LogUiFailure("DISPATCHER_UNAVAILABLE", Nothing)
            ElseIf dispatcher IsNot Nothing Then
                PostDrain(dispatcher)
            End If
        End Sub

        Private Sub PostDrain(dispatcher As System.Threading.SynchronizationContext)
            Try
                ' Never use Send/Invoke or hold _syncLock while calling a dispatcher.
                dispatcher.Post(Sub(state As System.Object) DrainOnUi(), Nothing)
            Catch ex As System.Exception
                SyncLock _syncLock
                    _pending.Clear()
                    _drainPosted = False
                End SyncLock
                LogUiFailure("POST_FAILED", ex)
            End Try
        End Sub

        Private Sub DrainOnUi()
            If System.Threading.Thread.CurrentThread.ManagedThreadId <> _uiThreadId Then
                SyncLock _syncLock
                    _pending.Clear()
                    _drainPosted = False
                End SyncLock
                LogUiFailure("WRONG_DISPATCHER_THREAD", Nothing)
                Return
            End If
            ' Bound each batch so a busy producer cannot monopolize the Office UI.
            For count As System.Int32 = 1 To 64
                Dim action As System.Action
                SyncLock _syncLock
                    If _shuttingDown OrElse _pending.Count = 0 Then
                        _drainPosted = False
                        Return
                    End If
                    action = _pending.Dequeue()
                End SyncLock
                Try
                    action.Invoke()
                Catch ex As System.Exception
                    LogUiFailure("UPDATE_FAILED", ex)
                End Try
            Next
            Dim dispatcher As System.Threading.SynchronizationContext
            SyncLock _syncLock
                If _shuttingDown OrElse _pending.Count = 0 Then
                    _drainPosted = False
                    Return
                End If
                dispatcher = _uiContext
            End SyncLock
            PostDrain(dispatcher)
        End Sub

        Private Sub EnsureFormCreatedOnUi()
            If _hostRegistry Is Nothing OrElse _hostRegistry.IsClosing Then Return
            If _logForm IsNot Nothing AndAlso Not _logForm.IsDisposed Then Return
            Dim created As New LogWindowForm()
            Try
                created.InitializeView()
                _hostRegistry.Register(created)
                _logForm = created
            Catch
                created.Dispose()
                Throw
            End Try
        End Sub

        Private Sub DisposeFormOnUi()
            Dim closing As LogWindowForm = _logForm
            _logForm = Nothing
            If closing IsNot Nothing AndAlso Not closing.IsDisposed Then closing.ClosePermanently()
        End Sub

        ''' <summary>Called on the owning host UI thread before registry teardown; idempotent.</summary>
        Public Sub Shutdown()
            SyncLock _syncLock
                If _shuttingDown Then Return
                If _uiThreadId <> 0 AndAlso System.Threading.Thread.CurrentThread.ManagedThreadId <> _uiThreadId Then
                    Throw New System.InvalidOperationException("The log window must shut down on its owning Office UI thread.")
                End If
                _shuttingDown = True
                _pending.Clear()
                _drainPosted = False
                _closeCallback = Nothing
            End SyncLock
            Try
                DisposeFormOnUi()
            Finally
                SyncLock _syncLock
                    _uiContext = Nothing
                    _hostRegistry = Nothing
                End SyncLock
            End Try
        End Sub

        Private Sub LogUiFailure(stage As System.String, failure As System.Exception)
            Try
                If failure Is Nothing Then
                    RiCrashLogger.LogMarker("LOG_WINDOW_" & stage, "No UI action was performed on an unverified thread.")
                Else
                    RiCrashLogger.LogException("LOG_WINDOW_" & stage, failure)
                End If
            Catch
                System.Diagnostics.Debug.WriteLine("Log window " & stage)
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
            Private _standardFont As Font

            ''' <summary>
            ''' Monospace font used for log output.
            ''' </summary>
            Private _monoFont As Font
            Private _ownedIcon As Icon
            Private _contextMenu As ContextMenuStrip
            Private _allowPermanentClose As System.Boolean

            ''' <summary>
            ''' Creates a new instance of the log window form and initializes its UI.
            ''' </summary>
            Public Sub New()
                MyBase.New()
            End Sub

            Public Sub InitializeView()
                _standardFont = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
                _monoFont = New Font("Consolas", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
                InitializeComponent()
            End Sub

            Public Sub ClosePermanently()
                _allowPermanentClose = True
                Try
                    Me.Close()
                Finally
                    If Not Me.IsDisposed Then Me.Dispose()
                End Try
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
                                _ownedIcon = CType(ico.Clone(), Icon)
                                Me.Icon = _ownedIcon
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
                _contextMenu = New ContextMenuStrip()
                Dim cms As ContextMenuStrip = _contextMenu
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

                Dim timestamp = DateTime.Now.ToString("HH:mm:ss")
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
                ' Preserve ordinary user-close confirmation/hide. Host exit and
                ' explicit disposal must never display an abort question.
                Dim hostClosing As System.Boolean = _hostRegistry IsNot Nothing AndAlso _hostRegistry.IsClosing
                If e.CloseReason = CloseReason.UserClosing AndAlso Not _allowPermanentClose AndAlso Not hostClosing Then
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

            Protected Overrides Sub Dispose(disposing As System.Boolean)
                Try
                    MyBase.Dispose(disposing)
                Finally
                    If disposing Then
                        If _contextMenu IsNot Nothing Then
                            _contextMenu.Dispose()
                            _contextMenu = Nothing
                        End If
                        If _ownedIcon IsNot Nothing Then
                            _ownedIcon.Dispose()
                            _ownedIcon = Nothing
                        End If
                        If _monoFont IsNot Nothing Then
                            _monoFont.Dispose()
                            _monoFont = Nothing
                        End If
                        If _standardFont IsNot Nothing Then
                            _standardFont.Dispose()
                            _standardFont = Nothing
                        End If
                    End If
                End Try
            End Sub

            <DllImport("user32.dll", SetLastError:=True)>
            Private Shared Function DestroyIcon(hIcon As IntPtr) As Boolean
            End Function
        End Class

    End Module

End Namespace