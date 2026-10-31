' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SplashScreenCountDown.vb
' Purpose: UI-free controller for a countdown splash on one dedicated STA.
'
' Contract:
'  - Construction never creates a Form, Control, native handle or UI context.
'  - ShowAsync awaits readiness without blocking the caller's message pump.
'    Startup is bounded; close/cancellation before readiness is remembered.
'  - The private view constructs, displays, updates and disposes ALL of its
'    WinForms/GDI resources on its own STA, inside the same Using/Run lifetime.
'  - Producers enqueue data only. A UI timer drains it on the owning thread;
'    there are no cross-thread Invoke/Send calls or background countdown tasks.
'  - Close/Dispose request shutdown; Completion signals actual owner-thread
'    cleanup. Neither waits synchronously or cancels a successful request.
'  - ESC cancels RequestCancellationToken and raises the compatible event.
'    A display failure cancels the same token, but is distinguished by Failure.
'  - The splash is deliberately ownerless/modeless: never attach an Office
'    window from another thread. SharedMethods.DialogOwner remains unchanged.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary
    Public NotInheritable Class SplashScreenCountDown
        Implements System.IDisposable

        Private Const StartupTimeoutMilliseconds As System.Int32 = 10000
        Private ReadOnly stateLock As New System.Object()
        Private ReadOnly instanceId As System.String = System.Guid.NewGuid().ToString("N")
        Private ReadOnly initialText As System.String
        Private ReadOnly requestedWidth As System.Int32
        Private ReadOnly requestedHeight As System.Int32
        Private ReadOnly initialSeconds As System.Int32
        Private ReadOnly pendingUpdates As New System.Collections.Generic.Queue(Of SplashUpdate)()
        Private ReadOnly ready As New System.Threading.Tasks.TaskCompletionSource(Of System.Boolean)(System.Threading.Tasks.TaskCreationOptions.RunContinuationsAsynchronously)
        Private ReadOnly stopped As New System.Threading.Tasks.TaskCompletionSource(Of System.Boolean)(System.Threading.Tasks.TaskCreationOptions.RunContinuationsAsynchronously)
        Private ReadOnly requestAbort As New System.Threading.CancellationTokenSource()
        Private ReadOnly requestAbortToken As System.Threading.CancellationToken
        Private splashThread As System.Threading.Thread
        Private startRequested As System.Boolean
        Private closeRequested As System.Boolean
        Private cleanupCompleted As System.Boolean
        Private displayFailure As SplashDisplayException

        ''' <summary>Raised for ESC only, never for normal Close/Dispose or display failure.</summary>
        Public Event CancelRequested As System.EventHandler

        Public Sub New(Optional ByVal customText As System.String = "Please wait …",
                       Optional ByVal formWidth As System.Int32 = 0,
                       Optional ByVal formHeight As System.Int32 = 0,
                       Optional ByVal countdownSeconds As System.Int32 = 0)
            initialText = If(customText, System.String.Empty)
            requestedWidth = formWidth
            requestedHeight = formHeight
            initialSeconds = System.Math.Max(0, countdownSeconds)
            requestAbortToken = requestAbort.Token
            LogLifecycle("CREATED")
        End Sub

        ''' <summary>
        ''' Link this token to the request CTS. It is canceled for ESC or a running
        ''' view failure, not for normal Close. The cached token is safe to read
        ''' even after owner-thread cleanup has disposed its source.
        ''' </summary>
        Public ReadOnly Property RequestCancellationToken As System.Threading.CancellationToken
            Get
                Return requestAbortToken
            End Get
        End Property

        Public ReadOnly Property Completion As System.Threading.Tasks.Task
            Get
                Return stopped.Task
            End Get
        End Property

        Public ReadOnly Property Failure As SplashDisplayException
            Get
                SyncLock stateLock
                    Return displayFailure
                End SyncLock
            End Get
        End Property

        Public Sub ThrowIfFailed()
            Dim failureSnapshot As SplashDisplayException = Failure
            If failureSnapshot IsNot Nothing Then Throw failureSnapshot
        End Sub

        ''' <summary>
        ''' Starts at most one dedicated STA and asynchronously waits for its first
        ''' displayed/message-loop tick. No Wait/Result/Join on an Office thread.
        ''' </summary>
        Public Async Function ShowAsync(Optional cancellationToken As System.Threading.CancellationToken = Nothing) As System.Threading.Tasks.Task
            cancellationToken.ThrowIfCancellationRequested()
            Using cancellationRegistration As System.Threading.CancellationTokenRegistration =
                    cancellationToken.Register(AddressOf Close)
                StartThreadIfNeeded()
                Using startupWait As New System.Threading.CancellationTokenSource()
                    Dim deadline As System.Threading.Tasks.Task =
                        System.Threading.Tasks.Task.Delay(StartupTimeoutMilliseconds, startupWait.Token)
                    Try
                        Dim completed As System.Threading.Tasks.Task =
                            Await System.Threading.Tasks.Task.WhenAny(ready.Task, deadline).ConfigureAwait(False)
                        ThrowIfFailed()
                        cancellationToken.ThrowIfCancellationRequested()
                        If completed IsNot ready.Task Then
                            RecordFailure(New System.TimeoutException("Splash readiness deadline exceeded."))
                            Close()
                        End If
                        ThrowIfFailed()
                        Dim wasReady As System.Boolean = Await ready.Task.ConfigureAwait(False)
                        ThrowIfFailed()
                        If Not wasReady OrElse IsClosing Then
                            Throw New System.OperationCanceledException("The progress display was closed before the request started.", requestAbortToken)
                        End If
                    Finally
                        startupWait.Cancel()
                    End Try
                End Using
            End Using
        End Function

        Private Sub StartThreadIfNeeded()
            Try
                SyncLock stateLock
                    If startRequested OrElse closeRequested Then Return
                    startRequested = True
                    splashThread = New System.Threading.Thread(AddressOf RunView)
                    splashThread.IsBackground = True
                    splashThread.Name = "Red Ink countdown " & instanceId.Substring(0, 8)
                    splashThread.SetApartmentState(System.Threading.ApartmentState.STA)
                    splashThread.Start()
                End SyncLock
            Catch ex As System.Exception
                RecordFailure(ex)
                Close()
                CompleteStopped()
            End Try
        End Sub

        Private Sub RunView()
            Dim exceptionHandler As System.Threading.ThreadExceptionEventHandler =
                Sub(sender As System.Object, e As System.Threading.ThreadExceptionEventArgs)
                    HandleViewFailure(e.Exception)
                    Try
                        System.Windows.Forms.Application.ExitThread()
                    Catch exitException As System.Exception
                        LogFailure("MESSAGE_LOOP_EXIT_FAILED", exitException)
                    End Try
                End Sub
            Dim handlerInstalled As System.Boolean = False
            Try
                If IsClosing Then Return
                ' This dedicated UI thread never borrows a caller/Office context.
                System.Threading.SynchronizationContext.SetSynchronizationContext(Nothing)
                ' WinForms stores ThreadException on the CURRENT thread context.
                ' Install/remove only on our own STA; never change the host's policy.
                System.Windows.Forms.Application.SetUnhandledExceptionMode(
                    System.Windows.Forms.UnhandledExceptionMode.CatchException, True)
                AddHandler System.Windows.Forms.Application.ThreadException, exceptionHandler
                handlerInstalled = True
                Using view As New SplashView(Me)
                    ' Separate initialization ensures Using disposes partially
                    ' initialized resources if logo/font/layout creation fails.
                    view.InitializeView(initialText, requestedWidth, requestedHeight, initialSeconds)
                    If IsClosing Then Return
                    LogLifecycle("UI_CREATED")
                    System.Windows.Forms.Application.Run(view)
                End Using
                If Not IsClosing Then
                    HandleViewFailure(New System.InvalidOperationException("Splash message loop ended unexpectedly."))
                End If
            Catch ex As System.Exception
                HandleViewFailure(ex)
            Finally
                Try
                    If handlerInstalled Then RemoveHandler System.Windows.Forms.Application.ThreadException, exceptionHandler
                Catch cleanupException As System.Exception
                    LogFailure("EXCEPTION_HANDLER_CLEANUP_FAILED", cleanupException)
                Finally
                    CompleteStopped()
                End Try
            End Try
        End Sub

        ''' <summary>Idempotent, non-blocking shutdown request; never invokes a Control.</summary>
        Public Sub Close()
            Dim firstClose As System.Boolean
            Dim neverStarted As System.Boolean
            SyncLock stateLock
                firstClose = Not closeRequested
                closeRequested = True
                pendingUpdates.Clear()
                neverStarted = Not startRequested
            End SyncLock
            ready.TrySetResult(False)
            If firstClose Then LogLifecycle("CLOSE_REQUESTED")
            If neverStarted Then CompleteStopped()
        End Sub

        Public Sub Dispose() Implements System.IDisposable.Dispose
            Close()
        End Sub

        Public Sub UpdateMessage(ByVal newMessage As System.String)
            EnqueueUpdate(New SplashUpdate(False, 0, If(newMessage, System.String.Empty)))
        End Sub

        Public Sub RestartCountdown(ByVal seconds As System.Int32,
                                    Optional ByVal newBaseText As System.String = Nothing)
            EnqueueUpdate(New SplashUpdate(True, System.Math.Max(0, seconds), newBaseText))
        End Sub

        Private Sub EnqueueUpdate(update As SplashUpdate)
            SyncLock stateLock
                If closeRequested Then Return
                pendingUpdates.Enqueue(update)
            End SyncLock
        End Sub

        Private Function TakeUpdates() As SplashUpdate()
            SyncLock stateLock
                Dim updates As SplashUpdate() = pendingUpdates.ToArray()
                pendingUpdates.Clear()
                Return updates
            End SyncLock
        End Function

        Private ReadOnly Property IsClosing As System.Boolean
            Get
                SyncLock stateLock
                    Return closeRequested
                End SyncLock
            End Get
        End Property

        Private Sub SignalReady()
            Dim firstReady As System.Boolean
            SyncLock stateLock
                If closeRequested Then Return
                firstReady = ready.TrySetResult(True)
            End SyncLock
            If firstReady Then LogLifecycle("READY")
        End Sub

        ' Only the owning STA calls this method, before disposing requestAbort.
        Private Sub RequestUserCancellation()
            If IsClosing Then Return
            LogLifecycle("USER_CANCEL")
            Try
                RaiseEvent CancelRequested(Me, System.EventArgs.Empty)
            Catch ex As System.Exception
                LogFailure("CANCEL_HANDLER_FAILED", ex)
            Finally
                CancelRequestOnOwnerThread()
                Close()
            End Try
        End Sub

        Private Sub HandleViewFailure(ex As System.Exception)
            RecordFailure(ex)
            ' RunView and UI timer are on the source's lifetime-owning STA.
            CancelRequestOnOwnerThread()
            Close()
        End Sub

        Private Sub CancelRequestOnOwnerThread()
            Try
                requestAbort.Cancel()
            Catch ex As System.Exception
                ' A consumer's cancellation callback must not escape the UI loop.
                LogFailure("CANCEL_CALLBACK_FAILED", ex)
            End Try
        End Sub

        Private Sub RecordFailure(ex As System.Exception)
            Dim isFirstFailure As System.Boolean
            SyncLock stateLock
                isFirstFailure = displayFailure Is Nothing
                If isFirstFailure Then displayFailure = New SplashDisplayException(ex)
            End SyncLock
            If isFirstFailure Then LogFailure("FAILED", ex)
        End Sub

        Private Sub CompleteStopped()
            SyncLock stateLock
                If cleanupCompleted Then Return
                cleanupCompleted = True
                closeRequested = True
                pendingUpdates.Clear()
            End SyncLock
            ready.TrySetResult(False)
            Try
                requestAbort.Dispose()
            Catch ex As System.Exception
                LogFailure("TOKEN_DISPOSE_FAILED", ex)
            Finally
                LogLifecycle("STOPPED")
                stopped.TrySetResult(True)
            End Try
        End Sub

        Private Sub LogLifecycle(stage As System.String)
            Try
                RiCrashLogger.LogMarker("SPLASH_" & stage,
                    "Instance=" & instanceId & "; ManagedThread=" &
                    System.Threading.Thread.CurrentThread.ManagedThreadId.ToString(System.Globalization.CultureInfo.InvariantCulture))
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("Splash lifecycle logging failed: " & ex.ToString())
            End Try
        End Sub

        Private Sub LogFailure(stage As System.String, ex As System.Exception)
            Try
                RiCrashLogger.LogException("SPLASH_" & stage & "; Instance=" & instanceId, ex)
            Catch logException As System.Exception
                System.Diagnostics.Debug.WriteLine(ex.ToString())
                System.Diagnostics.Debug.WriteLine(logException.ToString())
            End Try
        End Sub

        ''' <summary>Public message contains no technical details; InnerException is for diagnostics.</summary>
        Public NotInheritable Class SplashDisplayException
            Inherits System.Exception

            Public Sub New(innerException As System.Exception)
                MyBase.New("The progress window could not be displayed. The request was stopped. Please try again.", innerException)
            End Sub
        End Class

        Private NotInheritable Class SplashUpdate
            Public ReadOnly IsRestart As System.Boolean
            Public ReadOnly Seconds As System.Int32
            Public ReadOnly Text As System.String

            Public Sub New(restart As System.Boolean, duration As System.Int32, message As System.String)
                IsRestart = restart
                Seconds = duration
                Text = message
            End Sub
        End Class

        ''' <summary>Private owner-thread-only form. No caller ever receives a Control reference.</summary>
        Private NotInheritable Class SplashView
            Inherits System.Windows.Forms.Form

            Private Const LayoutPadding As System.Int32 = 10
            Private ReadOnly controller As SplashScreenCountDown
            Private ReadOnly countdownClock As New System.Diagnostics.Stopwatch()
            Private lblMessage As System.Windows.Forms.Label
            Private picLogo As System.Windows.Forms.PictureBox
            Private ownedLogo As System.Drawing.Bitmap
            Private ownedFont As System.Drawing.Font
            Private uiTimer As System.Windows.Forms.Timer
            Private baseText As System.String
            Private countdownSeconds As System.Int32
            Private lastRemainingSeconds As System.Int32
            Private requestedFormWidth As System.Int32
            Private requestedFormHeight As System.Int32
            Private exitRequested As System.Boolean
            Private disposalLogged As System.Boolean

            Public Sub New(ownerController As SplashScreenCountDown)
                MyBase.New()
                controller = ownerController
                controller.LogLifecycle("UI_CONSTRUCTED")
            End Sub

            Public Sub InitializeView(message As System.String, width As System.Int32,
                                      height As System.Int32, seconds As System.Int32)
                requestedFormWidth = width
                requestedFormHeight = height
                Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
                Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
                Me.BackColor = System.Drawing.ColorTranslator.FromWin32(&H8000000F)
                Me.KeyPreview = True
                Me.TopMost = True

                ' Retain the original logo/font/layout; only dispose resources we own.
                ownedLogo = New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                picLogo = New System.Windows.Forms.PictureBox() With {
                    .Image = ownedLogo,
                    .SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
                }
                Me.Controls.Add(picLogo)
                ownedFont = New System.Drawing.Font("Segoe UI", 10.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
                lblMessage = New System.Windows.Forms.Label() With {
                    .Font = ownedFont,
                    .AutoSize = True,
                    .TextAlign = System.Drawing.ContentAlignment.MiddleLeft
                }
                Me.Controls.Add(lblMessage)
                baseText = message
                countdownSeconds = seconds
                lastRemainingSeconds = seconds
                ApplyMessageLayout(If(seconds > 0, $"{message} {seconds}s", message), True)
                If seconds > 0 Then countdownClock.Start()

                uiTimer = New System.Windows.Forms.Timer() With {.Interval = 100}
                AddHandler uiTimer.Tick, AddressOf OnUiTick
                ' The timer is also the shutdown mailbox consumer. Starting it here
                ' permits ExitThread even when a late close suppresses initial Show.
                uiTimer.Start()
            End Sub

            Protected Overrides Sub SetVisibleCore(value As System.Boolean)
                If value AndAlso controller.IsClosing Then Return
                MyBase.SetVisibleCore(value)
            End Sub

            Private Sub OnUiTick(sender As System.Object, e As System.EventArgs)
                Try
                    If controller.IsClosing Then
                        ExitView()
                        Return
                    End If
                    For Each update As SplashUpdate In controller.TakeUpdates()
                        If controller.IsClosing Then Exit For
                        If update.IsRestart Then
                            If update.Text IsNot Nothing Then baseText = update.Text
                            countdownSeconds = update.Seconds
                            lastRemainingSeconds = countdownSeconds
                            countdownClock.Restart()
                            ApplyMessageLayout($"{baseText} {countdownSeconds}s", False)
                        Else
                            ApplyMessageLayout(update.Text, False)
                        End If
                    Next
                    If countdownClock.IsRunning AndAlso countdownSeconds > 0 Then
                        Dim remaining As System.Int32 = System.Math.Max(0,
                            CInt(System.Math.Ceiling(countdownSeconds - countdownClock.Elapsed.TotalSeconds)))
                        If remaining <> lastRemainingSeconds Then
                            lastRemainingSeconds = remaining
                            ApplyMessageLayout($"{baseText} {remaining}s", False)
                        End If
                        If remaining = 0 Then countdownClock.Stop()
                    End If
                    If controller.IsClosing Then
                        ExitView()
                    ElseIf Me.Visible AndAlso Me.IsHandleCreated Then
                        controller.SignalReady()
                    End If
                Catch ex As System.Exception
                    controller.HandleViewFailure(ex)
                    ExitView()
                End Try
            End Sub

            Private Sub ExitView()
                If exitRequested Then Return
                exitRequested = True
                Try
                    If uiTimer IsNot Nothing Then uiTimer.Stop()
                    countdownClock.Stop()
                    Me.Close()
                Finally
                    ' Close on a never-shown form need not raise FormClosed. Always
                    ' exit THIS thread's loop, never Application.Exit (all hosts).
                    System.Windows.Forms.Application.ExitThread()
                End Try
            End Sub

            Protected Overrides Sub OnKeyDown(e As System.Windows.Forms.KeyEventArgs)
                MyBase.OnKeyDown(e)
                If e.KeyCode = System.Windows.Forms.Keys.Escape Then
                    e.Handled = True
                    e.SuppressKeyPress = True
                    controller.RequestUserCancellation()
                    ExitView()
                End If
            End Sub

            Protected Overrides Sub OnFormClosing(e As System.Windows.Forms.FormClosingEventArgs)
                ' Native user close (for example Alt+F4) is cancellation, not a
                ' technical display failure. Programmatic completion has closed
                ' the controller already and must not cancel the request.
                If Not controller.IsClosing Then controller.RequestUserCancellation()
                MyBase.OnFormClosing(e)
            End Sub

            Private Sub ApplyMessageLayout(messageText As System.String, allowShrink As System.Boolean)
                lblMessage.Text = messageText
                Dim textSize As System.Drawing.Size =
                    System.Windows.Forms.TextRenderer.MeasureText(messageText, lblMessage.Font)
                lblMessage.Size = textSize
                Dim logoSize As System.Int32 = textSize.Height
                picLogo.SetBounds(LayoutPadding, LayoutPadding, logoSize, logoSize)
                Dim labelX As System.Int32 = picLogo.Right + LayoutPadding
                Dim labelY As System.Int32 = LayoutPadding + (logoSize - textSize.Height) \ 2
                lblMessage.SetBounds(labelX, labelY, textSize.Width, textSize.Height)
                Dim clientW As System.Int32 = lblMessage.Right + LayoutPadding
                Dim clientH As System.Int32 = System.Math.Max(picLogo.Bottom, lblMessage.Bottom) + LayoutPadding
                If requestedFormWidth > 0 Then
                    clientW = requestedFormWidth
                ElseIf Not allowShrink Then
                    clientW = System.Math.Max(Me.ClientSize.Width, clientW)
                End If
                If requestedFormHeight > 0 Then
                    clientH = requestedFormHeight
                ElseIf Not allowShrink Then
                    clientH = System.Math.Max(Me.ClientSize.Height, clientH)
                End If
                Me.ClientSize = New System.Drawing.Size(clientW, clientH)
                lblMessage.Refresh()
            End Sub

            Protected Overrides Sub Dispose(disposing As System.Boolean)
                Try
                    If disposing Then
                        If uiTimer IsNot Nothing Then
                            uiTimer.Stop()
                            RemoveHandler uiTimer.Tick, AddressOf OnUiTick
                            uiTimer.Dispose()
                            uiTimer = Nothing
                        End If
                        countdownClock.Stop()
                    End If
                Finally
                    Try
                        MyBase.Dispose(disposing)
                    Finally
                        If disposing Then
                            If ownedLogo IsNot Nothing Then
                                ownedLogo.Dispose()
                                ownedLogo = Nothing
                            End If
                            If ownedFont IsNot Nothing Then
                                ownedFont.Dispose()
                                ownedFont = Nothing
                            End If
                            If Not disposalLogged Then
                                disposalLogged = True
                                controller.LogLifecycle("UI_DISPOSED")
                            End If
                        End If
                    End Try
                End Try
            End Sub

            <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
            Private Shared Function ReleaseCapture() As System.Boolean
            End Function

            <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
            Private Shared Function SendMessage(hWnd As System.IntPtr, wMsg As System.Int32,
                                               wParam As System.IntPtr, lParam As System.IntPtr) As System.IntPtr
            End Function

            Protected Overrides Sub OnMouseDown(e As System.Windows.Forms.MouseEventArgs)
                MyBase.OnMouseDown(e)
                If e.Button = System.Windows.Forms.MouseButtons.Left Then
                    ReleaseCapture()
                    SendMessage(Me.Handle, &HA1, New System.IntPtr(2), System.IntPtr.Zero)
                End If
            End Sub
        End Class
    End Class
End Namespace
