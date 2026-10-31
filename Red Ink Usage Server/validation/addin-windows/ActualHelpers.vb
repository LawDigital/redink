Option Strict On
Option Explicit On
Imports System.IO
Namespace SharedLibrary
Public Partial Class SharedMethods
        ' Explicit host-owned dispatcher. Initialization happens on the actual Office STA.
        ' Never discover a dispatcher from OpenForms, foreground windows, or an arbitrary caller.
        Private Shared ReadOnly _hostUiDispatchLock As New System.Object()
        Private Shared _hostUiDispatchContext As System.Threading.SynchronizationContext
        Private Shared _hostUiDispatchThreadId As System.Int32
        Private Shared _hostUiDispatchRegistry As HostModelessFormRegistry

        Public Shared Sub InitializeHostUiDispatcher(context As System.Threading.SynchronizationContext,
                                                      registry As HostModelessFormRegistry)
            If context Is Nothing Then Throw New System.ArgumentNullException(NameOf(context))
            If registry Is Nothing Then Throw New System.ArgumentNullException(NameOf(registry))
            If System.Threading.Thread.CurrentThread.GetApartmentState() <> System.Threading.ApartmentState.STA Then
                Throw New System.InvalidOperationException("The host UI dispatcher must be initialized on the Office STA.")
            End If
            SyncLock _hostUiDispatchLock
                If _hostUiDispatchRegistry IsNot Nothing AndAlso
                   Not _hostUiDispatchRegistry.ShutdownToken.IsCancellationRequested AndAlso
                   (_hostUiDispatchThreadId <> System.Threading.Thread.CurrentThread.ManagedThreadId OrElse
                    _hostUiDispatchRegistry IsNot registry) Then
                    Throw New System.InvalidOperationException("A different live Office UI dispatcher is already registered.")
                End If
                _hostUiDispatchContext = context
                _hostUiDispatchThreadId = System.Threading.Thread.CurrentThread.ManagedThreadId
                _hostUiDispatchRegistry = registry
            End SyncLock
        End Sub

        Public Shared ReadOnly Property HasHostUiDispatcher As System.Boolean
            Get
                SyncLock _hostUiDispatchLock
                    Return _hostUiDispatchContext IsNot Nothing AndAlso _hostUiDispatchRegistry IsNot Nothing AndAlso
                           Not _hostUiDispatchRegistry.ShutdownToken.IsCancellationRequested
                End SyncLock
            End Get
        End Property

        Public Shared ReadOnly Property IsOnHostUiThread As System.Boolean
            Get
                SyncLock _hostUiDispatchLock
                    Return _hostUiDispatchContext IsNot Nothing AndAlso _hostUiDispatchRegistry IsNot Nothing AndAlso
                           Not _hostUiDispatchRegistry.ShutdownToken.IsCancellationRequested AndAlso
                           System.Threading.Thread.CurrentThread.ManagedThreadId = _hostUiDispatchThreadId AndAlso
                           System.Threading.Thread.CurrentThread.GetApartmentState() = System.Threading.ApartmentState.STA
                End SyncLock
            End Get
        End Property

        Private Shared Function AwaitHostUiThread() As OfficeUiThreadAwaitable
            SyncLock _hostUiDispatchLock
                Dim token As System.Threading.CancellationToken = System.Threading.CancellationToken.None
                If _hostUiDispatchRegistry IsNot Nothing Then token = _hostUiDispatchRegistry.ShutdownToken
                Return New OfficeUiThreadAwaitable(_hostUiDispatchContext, _hostUiDispatchThreadId, token)
            End SyncLock
        End Function

        ''' <summary>
        ''' Executes the whole action on the captured host UI thread. The returned Task
        ''' reports completion, cancellation or failure; its caller must dispatch its own
        ''' later UI work explicitly if it does not preserve a UI synchronization context.
        ''' No shared lock is held while posting or executing the action.
        ''' </summary>
        Public Shared Async Function PostToHostUiAsync(action As System.Action) As System.Threading.Tasks.Task
            If action Is Nothing Then Throw New System.ArgumentNullException(NameOf(action))
            Await AwaitHostUiThread()
            action.Invoke()
        End Function

        Public Shared Async Function PostToHostUiAsync(Of TResult)(action As System.Func(Of TResult)) As System.Threading.Tasks.Task(Of TResult)
            If action Is Nothing Then Throw New System.ArgumentNullException(NameOf(action))
            Await AwaitHostUiThread()
            Return action.Invoke()
        End Function

        Private Shared Function RequireHostUiRegistry() As HostModelessFormRegistry
            SyncLock _hostUiDispatchLock
                If Not IsOnHostUiThread Then
                    Throw New System.InvalidOperationException("Host UI forms must be managed on the initialized Office UI thread.")
                End If
                Return _hostUiDispatchRegistry
            End SyncLock
        End Function

        Public Shared Function RegisterHostUiForm(form As System.Windows.Forms.Form) As System.Windows.Forms.IWin32Window
            Return RequireHostUiRegistry().Register(form)
        End Function

        Public Shared Sub ShowHostUiForm(form As System.Windows.Forms.Form)
            RequireHostUiRegistry().Show(form)
        End Sub


        <System.ThreadStatic>
        Private Shared _hostNeutralStaDepth As System.Int32

        Public Shared ReadOnly Property IsOnHostNeutralSta As System.Boolean
            Get
                Return _hostNeutralStaDepth > 0 AndAlso
                       System.Threading.Thread.CurrentThread.GetApartmentState() = System.Threading.ApartmentState.STA
            End Get
        End Property

        ''' <summary>
        ''' Runs a complete host-neutral dialog or OLE operation on an STA. Existing STA
        ''' callers stay on their thread, preserving safe modal ownership. MTA callers
        ''' use a background STA and never transfer an Office owner to that thread.
        ''' The action must dispose its forms on its own thread and must not access Office COM.
        ''' Worker failures are rethrown on the caller so existing error handlers still run.
        ''' </summary>
        Public Shared Sub RunOnStaThread(action As System.Action,
                                        Optional threadName As System.String = "Red Ink STA",
                                        Optional waitTimeoutMilliseconds As System.Int32 = System.Threading.Timeout.Infinite,
                                        Optional alwaysUseWorker As System.Boolean = False)
            If action Is Nothing Then Throw New System.ArgumentNullException(NameOf(action))
            RunOnStaThread(Of System.Boolean)(
                Function()
                    action.Invoke()
                    Return True
                End Function, threadName, waitTimeoutMilliseconds, alwaysUseWorker)
        End Sub

        Public Shared Function RunOnStaThread(Of TResult)(action As System.Func(Of TResult),
                                                          Optional threadName As System.String = "Red Ink STA",
                                                          Optional waitTimeoutMilliseconds As System.Int32 = System.Threading.Timeout.Infinite,
                                                          Optional alwaysUseWorker As System.Boolean = False) As TResult
            If action Is Nothing Then Throw New System.ArgumentNullException(NameOf(action))
            If waitTimeoutMilliseconds < System.Threading.Timeout.Infinite Then
                Throw New System.ArgumentOutOfRangeException(NameOf(waitTimeoutMilliseconds))
            End If
            If Not alwaysUseWorker AndAlso System.Threading.Thread.CurrentThread.GetApartmentState() = System.Threading.ApartmentState.STA Then
                Dim previousScopeDepth As System.Int32 = _hostNeutralStaDepth
                _hostNeutralStaDepth += 1
                Try
                    Return action.Invoke()
                Finally
                    _hostNeutralStaDepth = previousScopeDepth
                End Try
            End If

            Dim result As TResult = Nothing
            Dim failure As System.Runtime.ExceptionServices.ExceptionDispatchInfo = Nothing
            Dim worker As New System.Threading.Thread(
                Sub()
                    ' This is a new, dedicated worker, never an Office/UI thread. Do not
                    ' adopt a caller's synchronization context or marshal back to its owner.
                    System.Threading.SynchronizationContext.SetSynchronizationContext(Nothing)
                    _hostNeutralStaDepth += 1
                    Try
                        result = action.Invoke()
                    Catch ex As System.Exception
                        failure = System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(ex)
                    Finally
                        _hostNeutralStaDepth -= 1
                        ' A modal dialog may install WinForms' marshaling control. Dispose
                        ' that thread's context before it exits; do not leave a SystemEvents
                        ' subscriber pointing at a dead worker/message pump.
                        Dim workerUiContext As System.Windows.Forms.WindowsFormsSynchronizationContext =
                            TryCast(System.Threading.SynchronizationContext.Current,
                                    System.Windows.Forms.WindowsFormsSynchronizationContext)
                        Try
                            System.Windows.Forms.Application.ExitThread()
                            If workerUiContext IsNot Nothing Then workerUiContext.Dispose()
                        Catch cleanupError As System.Exception
                            If failure Is Nothing Then
                                failure = System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(cleanupError)
                            Else
                                System.Diagnostics.Debug.WriteLine("STA context cleanup: " & cleanupError.Message)
                            End If
                        Finally
                            System.Threading.SynchronizationContext.SetSynchronizationContext(Nothing)
                        End Try
                    End Try
                End Sub)
            worker.IsBackground = True
            worker.Name = threadName
            worker.SetApartmentState(System.Threading.ApartmentState.STA)
            worker.Start()
            ' Modal callers wait only off the Office STA. Short OLE operations can explicitly
            ' request a worker with a bounded wait. A timed-out native call is never aborted/retried.
            If Not worker.Join(waitTimeoutMilliseconds) Then
                Throw New System.TimeoutException("The STA operation did not complete within " &
                    waitTimeoutMilliseconds.ToString(System.Globalization.CultureInfo.InvariantCulture) & " milliseconds.")
            End If
            If failure IsNot Nothing Then failure.Throw()
            Return result
        End Function
        Friend Shared Function ReportLegacyTextReaderError(value As System.String, ByRef readError As System.String) As System.String
            readError = If(value, "reader_failed")
            Return value
        End Function
        Private Shared Function SuppressLegacyTextReaderError(value As System.String, ByRef readError As System.String) As System.String
            readError = If(value, "reader_failed")
            Return System.String.Empty
        End Function
        Public Shared Function ReadRtfAsText(ByVal rtfPath As String, Optional ReturnErrorInsteadOfEmpty As Boolean = True, Optional ByRef readError As System.String = Nothing) As String
            readError = System.String.Empty
            Try
                Dim rtfContent As String = File.ReadAllText(rtfPath)
                Return RtfToPlainText(rtfContent)
            Catch ex As System.Exception
                Return If(ReturnErrorInsteadOfEmpty, ReportLegacyTextReaderError($"Error reading RTF: {ex.Message}", readError), SuppressLegacyTextReaderError($"Error reading RTF: {ex.Message}", readError))
            End Try
        End Function
        Public Shared Function RtfToPlainText(rtfContent As System.String) As System.String
            Return ReadRtfOnSta(Sub(rtb As System.Windows.Forms.RichTextBox) rtb.Rtf = rtfContent)
        End Function
        Public Shared Function LoadRtfFileAsPlainText(rtfPath As System.String) As System.String
            Return ReadRtfOnSta(
                Sub(rtb As System.Windows.Forms.RichTextBox)
                    rtb.LoadFile(rtfPath, System.Windows.Forms.RichTextBoxStreamType.RichText)
                End Sub)
        End Function
        Private Shared Function ReadRtfOnSta(loadRtf As System.Action(Of System.Windows.Forms.RichTextBox)) As System.String
            Return RunOnStaThread(Of System.String)(
                Function() As System.String
                    Dim previousContext As System.Threading.SynchronizationContext = System.Threading.SynchronizationContext.Current
                    Dim previousAutoInstall As System.Boolean = System.Windows.Forms.WindowsFormsSynchronizationContext.AutoInstall
                    Try
                        ' AutoInstall is thread-local. A hidden parsing control must not
                        ' install a dispatch context on a temporary STA with no message loop.
                        System.Windows.Forms.WindowsFormsSynchronizationContext.AutoInstall = False
                        Using rtb As New System.Windows.Forms.RichTextBox()
                            rtb.Visible = False
                            loadRtf.Invoke(rtb)
                            Return rtb.Text
                        End Using
                    Finally
                        System.Threading.SynchronizationContext.SetSynchronizationContext(previousContext)
                        System.Windows.Forms.WindowsFormsSynchronizationContext.AutoInstall = previousAutoInstall
                    End Try
                End Function, "Red Ink RTF")
        End Function
        Private NotInheritable Class OcrChunkStatusDialog
            Implements System.IDisposable

            Private ReadOnly _syncRoot As New Object()
            Private ReadOnly _ready As New System.Threading.Tasks.TaskCompletionSource(Of System.Boolean)(System.Threading.Tasks.TaskCreationOptions.RunContinuationsAsynchronously)
            Private ReadOnly _stopped As New System.Threading.Tasks.TaskCompletionSource(Of System.Boolean)(System.Threading.Tasks.TaskCreationOptions.RunContinuationsAsynchronously)
            Private ReadOnly _cancellation As New System.Threading.CancellationTokenSource()
            Private ReadOnly _cancellationToken As System.Threading.CancellationToken
            Private _cleanupCompleted As System.Boolean
            Private _disposeRequested As System.Boolean
            Private _cancellationDisposed As System.Boolean
            Private _uiThread As System.Threading.Thread = Nothing
            Private ReadOnly _operationName As System.String
            Private ReadOnly _sourceName As System.String
            Private _progressMaximum As System.Int32 = 0
            Private _progressValue As System.Int32 = 0
            Private _statusText As String = "Starting OCR..."
            Private ReadOnly _totalClock As New System.Diagnostics.Stopwatch()
            Private ReadOnly _requestClock As New System.Diagnostics.Stopwatch()
            Private _phase As System.String = "Starting OCR..."
            Private _cancelled As Boolean = False
            Private _closeRequested As Boolean = False
            Private _form As System.Windows.Forms.Form = Nothing

            Public Sub New(Optional operationName As System.String = "OCR", Optional sourceName As System.String = "")
                _cancellationToken = _cancellation.Token
                _operationName = operationName
                _sourceName = sourceName
                _statusText = "Starting " & operationName & "..."
                _phase = _statusText
            End Sub

            Public Async Function ShowAsync(initialText As System.String,
                                            Optional cancellationToken As System.Threading.CancellationToken = Nothing) As System.Threading.Tasks.Task
                cancellationToken.ThrowIfCancellationRequested()
                SyncLock _syncRoot
                    If _closeRequested Then Throw New System.OperationCanceledException("The progress window was already closed.")
                    If _uiThread IsNot Nothing Then Throw New System.InvalidOperationException("The progress window has already been started.")
                    _statusText = initialText
                    _totalClock.Restart()
                    _requestClock.Reset()
                    _phase = "Starting " & _operationName & "..."
                    _uiThread = New System.Threading.Thread(AddressOf UiThreadMain) With {
                        .IsBackground = True,
                        .Name = "Red Ink extraction progress"
                    }
                End SyncLock
                Using cancellationRegistration As System.Threading.CancellationTokenRegistration = cancellationToken.Register(AddressOf Close)
                    Try
                        _uiThread.SetApartmentState(System.Threading.ApartmentState.STA)
                        _uiThread.Start()
                    Catch ex As System.Exception
                        _ready.TrySetException(ex)
                        CompleteLifetime()
                    End Try
                    Using startupDeadline As New System.Threading.CancellationTokenSource()
                        Try
                            Dim deadline As System.Threading.Tasks.Task = System.Threading.Tasks.Task.Delay(10000, startupDeadline.Token)
                            Dim first As System.Threading.Tasks.Task = Await System.Threading.Tasks.Task.WhenAny(_ready.Task, deadline).ConfigureAwait(False)
                            cancellationToken.ThrowIfCancellationRequested()
                            If first IsNot _ready.Task Then
                                Close()
                                Throw New System.TimeoutException("The extraction progress window did not become ready in time.")
                            End If
                            If Not Await _ready.Task.ConfigureAwait(False) Then
                                Throw New System.OperationCanceledException("The extraction progress window closed before startup completed.")
                            End If
                        Finally
                            startupDeadline.Cancel()
                        End Try
                    End Using
                End Using
            End Function

            Public ReadOnly Property Completion As System.Threading.Tasks.Task
                Get
                    Return _stopped.Task
                End Get
            End Property

            Public Sub UpdateStatus(text As String)
                SyncLock _syncRoot
                    _statusText = text
                End SyncLock
            End Sub

            Public Sub SetProgress(value As System.Int32, maximum As System.Int32)
                SyncLock _syncRoot
                    _progressMaximum = System.Math.Max(0, maximum)
                    _progressValue = System.Math.Max(0, System.Math.Min(value, _progressMaximum))
                End SyncLock
            End Sub

            Public Sub SetPhase(phase As System.String)
                SyncLock _syncRoot
                    If Not _cancelled Then _phase = phase
                End SyncLock
            End Sub

            Public Sub BeginModelWait()
                SyncLock _syncRoot
                    If _cancelled Then Return
                    _requestClock.Restart()
                    _phase = "Waiting for the model response..."
                End SyncLock
            End Sub

            Public Sub EndModelWait()
                SyncLock _syncRoot
                    _requestClock.Stop()
                    If Not _cancelled Then _phase = "Model request finished."
                End SyncLock
            End Sub

            Private Shared Function FormatElapsed(elapsed As System.TimeSpan) As System.String
                Return System.Math.Floor(elapsed.TotalHours).ToString("00", System.Globalization.CultureInfo.InvariantCulture) & ":" &
                    elapsed.Minutes.ToString("00", System.Globalization.CultureInfo.InvariantCulture) & ":" &
                    elapsed.Seconds.ToString("00", System.Globalization.CultureInfo.InvariantCulture)
            End Function

            Private Function GetDisplayText() As System.String
                SyncLock _syncRoot
                    Dim sourceHeading As System.String = If(System.String.IsNullOrWhiteSpace(_sourceName), System.String.Empty, _sourceName & System.Environment.NewLine & System.Environment.NewLine)
                    Dim text As System.String = sourceHeading & _statusText & System.Environment.NewLine & System.Environment.NewLine & _phase &
                        System.Environment.NewLine & "Total elapsed: " & FormatElapsed(_totalClock.Elapsed)
                    If _requestClock.IsRunning Then
                        text &= System.Environment.NewLine & "Waiting this request: " & FormatElapsed(_requestClock.Elapsed) &
                            System.Environment.NewLine & "Progress updates after the response has been validated."
                    End If
                    Return text
                End SyncLock
            End Function

            Public ReadOnly Property IsCancelled As Boolean
                Get
                    SyncLock _syncRoot
                        Return _cancelled
                    End SyncLock
                End Get
            End Property

            Public ReadOnly Property CancellationToken As System.Threading.CancellationToken
                Get
                    Return _cancellationToken
                End Get
            End Property

            Private Sub RequestCancel()
                SyncLock _syncRoot
                    _cancelled = True
                    _statusText = "Cancelling " & _operationName & "..."
                    _phase = "Cancellation requested; waiting for the request to stop."
                    _requestClock.Stop()
                End SyncLock
                ' Invoke transport cancellation outside the status lock: callbacks may
                ' complete on another thread and must not wait for UI state ownership.
                Try
                    _cancellation.Cancel()
                Catch ex As System.ObjectDisposedException
                    ' A late window-close event after operation disposal is harmless.
                End Try
            End Sub

            Private Shared Function GetLowerMiddleLocation(formSize As System.Drawing.Size) As System.Drawing.Point
                Dim wa As System.Drawing.Rectangle = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea
                Dim x As Integer = wa.Left + ((wa.Width - formSize.Width) \ 2)
                Dim y As Integer = wa.Top + CInt((wa.Height * 0.75R) - (formSize.Height / 2.0R))

                If x < wa.Left Then x = wa.Left
                If y < wa.Top Then y = wa.Top
                If x + formSize.Width > wa.Right Then x = wa.Right - formSize.Width
                If y + formSize.Height > wa.Bottom Then y = wa.Bottom - formSize.Height

                Return New System.Drawing.Point(x, y)
            End Function

            Private Sub UiThreadMain()
                Dim localForm As System.Windows.Forms.Form = Nothing
                Dim bmpIcon As System.Drawing.Bitmap = Nothing
                Dim ownedIcon As System.Drawing.Icon = Nothing
                Dim standardFont As System.Drawing.Font = Nothing
                Dim refreshTimer As System.Windows.Forms.Timer = Nothing
                Dim exceptionHandler As System.Threading.ThreadExceptionEventHandler =
                    Sub(sender As System.Object, e As System.Threading.ThreadExceptionEventArgs)
                        HandleUiFailure(e.Exception)
                        System.Windows.Forms.Application.ExitThread()
                    End Sub
                Dim handlerInstalled As System.Boolean
                Try
                    System.Threading.SynchronizationContext.SetSynchronizationContext(Nothing)
                    System.Windows.Forms.Application.SetUnhandledExceptionMode(System.Windows.Forms.UnhandledExceptionMode.CatchException, True)
                    AddHandler System.Windows.Forms.Application.ThreadException, exceptionHandler
                    handlerInstalled = True
                    SyncLock _syncRoot
                        If _closeRequested Then Return
                    End SyncLock
                    localForm = New System.Windows.Forms.Form() With {
                        .Opacity = 0,
                        .Text = SharedMethods.AN & " " & _operationName,
                        .FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog,
                        .StartPosition = System.Windows.Forms.FormStartPosition.Manual,
                        .MaximizeBox = False,
                        .MinimizeBox = False,
                        .ShowInTaskbar = False,
                        .TopMost = True,
                        .KeyPreview = True,
                        .AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font,
                        .AutoSize = False
                    }

                    bmpIcon = New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                    Dim iconHandle As System.IntPtr = bmpIcon.GetHicon()
                    Try
                        Using borrowedIcon As System.Drawing.Icon = System.Drawing.Icon.FromHandle(iconHandle)
                            ownedIcon = CType(borrowedIcon.Clone(), System.Drawing.Icon)
                        End Using
                    Finally
                        DestroyProgressIcon(iconHandle)
                    End Try
                    localForm.Icon = ownedIcon
                    standardFont = New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
                    localForm.Font = standardFont

                    Dim wa As System.Drawing.Rectangle = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea
                    Dim paddingAll As Integer = 20
                    Dim gapAboveButtons As Integer = 10
                    Dim spacerExtra As Integer = 20
                    Dim minContentWidth As Integer = 320
                    Dim maxWindowWidth As Integer = System.Math.Min(wa.Width, System.Math.Max(420, CInt(System.Math.Floor(wa.Width * 0.4))))
                    Dim maxWindowHeight As Integer = CInt(System.Math.Floor(wa.Height * 0.6))

                    Dim cancelButton As New System.Windows.Forms.Button() With {
                        .Text = "Cancel",
                        .AutoSize = True,
                        .Font = standardFont,
                        .Margin = New System.Windows.Forms.Padding(0)
                    }

                    Dim bottomFlow As New System.Windows.Forms.FlowLayoutPanel() With {
                        .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
                        .AutoSize = True,
                        .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                        .Margin = New System.Windows.Forms.Padding(0)
                    }
                    bottomFlow.Controls.Add(cancelButton)
                    bottomFlow.PerformLayout()

                    Dim reservedBottomHeight As Integer = bottomFlow.PreferredSize.Height + gapAboveButtons

                    Dim statusLabel As New System.Windows.Forms.Label() With {
                        .Text = GetDisplayText(),
                        .Font = standardFont,
                        .AutoSize = True,
                        .Margin = New System.Windows.Forms.Padding(0)
                    }

                    Dim getLabelPreferred As System.Func(Of Integer, System.Drawing.Size) =
                        Function(w As Integer) As System.Drawing.Size
                            statusLabel.MaximumSize = New System.Drawing.Size(System.Math.Max(1, w), 0)
                            Return statusLabel.GetPreferredSize(New System.Drawing.Size(System.Math.Max(1, w), 0))
                        End Function

                    Dim maxContentWidth As Integer = System.Math.Max(minContentWidth, maxWindowWidth - 2 * paddingAll)
                    Dim pref As System.Drawing.Size = getLabelPreferred(maxContentWidth)
                    Dim contentWidth As Integer = System.Math.Max(minContentWidth, System.Math.Min(maxContentWidth, pref.Width))
                    pref = getLabelPreferred(contentWidth)

                    Dim maxBodyHeightNoScroll As Integer = System.Math.Max(100, maxWindowHeight - reservedBottomHeight - spacerExtra - 2 * paddingAll)

                    While (pref.Height > maxBodyHeightNoScroll) AndAlso ((contentWidth + 2 * paddingAll) < maxWindowWidth)
                        Dim stepW As Integer = System.Math.Max(20, (maxWindowWidth - 2 * paddingAll - contentWidth) \ 3)
                        contentWidth = System.Math.Min(maxWindowWidth - 2 * paddingAll, contentWidth + stepW)
                        pref = getLabelPreferred(contentWidth)
                    End While

                    Dim bodyPanelHeight As Integer = System.Math.Max(90, System.Math.Min(pref.Height, maxBodyHeightNoScroll))

                    Dim bodyPanel As New System.Windows.Forms.Panel() With {
                        .AutoSize = False,
                        .AutoScroll = True,
                        .Dock = System.Windows.Forms.DockStyle.Fill,
                        .Size = New System.Drawing.Size(contentWidth, bodyPanelHeight),
                        .Margin = New System.Windows.Forms.Padding(0),
                        .Padding = New System.Windows.Forms.Padding(0)
                    }

                    statusLabel.MaximumSize = New System.Drawing.Size(contentWidth, 0)
                    bodyPanel.Controls.Add(statusLabel)
                    statusLabel.Location = New System.Drawing.Point(0, 0)

                    Dim table As New System.Windows.Forms.TableLayoutPanel() With {
                        .Dock = System.Windows.Forms.DockStyle.Fill,
                        .ColumnCount = 1,
                        .RowCount = 3,
                        .Padding = New System.Windows.Forms.Padding(paddingAll),
                        .AutoSize = False,
                        .Margin = New System.Windows.Forms.Padding(0)
                    }
                    table.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
                    table.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, bodyPanelHeight))
                    table.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, spacerExtra))
                    table.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))

                    table.Controls.Add(bodyPanel, 0, 0)

                    Dim spacer As New System.Windows.Forms.Panel() With {
                        .Height = spacerExtra,
                        .Dock = System.Windows.Forms.DockStyle.Fill,
                        .Width = 1,
                        .Margin = New System.Windows.Forms.Padding(0)
                    }
                    Dim preparationProgress As New System.Windows.Forms.ProgressBar() With {
                        .Dock = System.Windows.Forms.DockStyle.Fill,
                        .Visible = System.String.Equals(_operationName, "Markdown", System.StringComparison.Ordinal),
                        .Style = System.Windows.Forms.ProgressBarStyle.Marquee,
                        .MarqueeAnimationSpeed = 30
                    }
                    spacer.Controls.Add(preparationProgress)
                    table.Controls.Add(spacer, 0, 1)

                    Dim bottomHost As New System.Windows.Forms.Panel() With {
                        .AutoSize = True,
                        .Margin = New System.Windows.Forms.Padding(0)
                    }
                    bottomHost.Padding = New System.Windows.Forms.Padding(0, gapAboveButtons, 0, 0)
                    bottomHost.Controls.Add(bottomFlow)
                    table.Controls.Add(bottomHost, 0, 2)

                    localForm.Controls.Clear()
                    localForm.Controls.Add(table)

                    Dim clientW As Integer = contentWidth + 2 * paddingAll
                    Dim clientH As Integer = bodyPanelHeight + spacerExtra + reservedBottomHeight + 2 * paddingAll
                    clientW = System.Math.Min(clientW, maxWindowWidth)
                    clientH = System.Math.Min(clientH, maxWindowHeight)
                    localForm.ClientSize = New System.Drawing.Size(clientW, clientH)
                    localForm.Location = GetLowerMiddleLocation(localForm.Size)

                    AddHandler cancelButton.Click,
                        Sub(sender As Object, e As System.EventArgs)
                            RequestCancel()
                        End Sub

                    AddHandler localForm.KeyDown,
                        Sub(sender As Object, e As System.Windows.Forms.KeyEventArgs)
                            If e.KeyCode = System.Windows.Forms.Keys.Escape Then
                                RequestCancel()
                                e.SuppressKeyPress = True
                            End If
                        End Sub

                    AddHandler localForm.FormClosing,
                        Sub(sender As Object, e As System.Windows.Forms.FormClosingEventArgs)
                            Dim closeRequested As Boolean

                            SyncLock _syncRoot
                                closeRequested = _closeRequested
                            End SyncLock

                            If Not closeRequested Then
                                RequestCancel()
                            End If
                        End Sub

                    AddHandler localForm.Shown,
                        Sub(sender As Object, e As System.EventArgs)
                            SharedMethods.ForceDialogToForeground(localForm)
                            _ready.TrySetResult(True)
                        End Sub

                    refreshTimer = New System.Windows.Forms.Timer() With {
                        .Interval = 250
                    }

                    AddHandler refreshTimer.Tick,
                        Sub(sender As Object, e As System.EventArgs)
                            Dim latestText As String = ""
                            Dim closeRequested As Boolean = False
                            Dim progressMaximum As System.Int32
                            Dim progressValue As System.Int32

                            SyncLock _syncRoot
                                latestText = GetDisplayText()
                                closeRequested = _closeRequested
                                progressMaximum = _progressMaximum
                                progressValue = _progressValue
                            End SyncLock
                            If preparationProgress.Visible Then
                                If progressMaximum > 0 Then
                                    preparationProgress.Style = System.Windows.Forms.ProgressBarStyle.Blocks
                                    If preparationProgress.Maximum <> progressMaximum Then
                                        ' Reset before lowering Maximum between files/phases.
                                        preparationProgress.Value = 0
                                        preparationProgress.Maximum = progressMaximum
                                    End If
                                    If preparationProgress.Value <> progressValue Then preparationProgress.Value = progressValue
                                Else
                                    preparationProgress.Style = System.Windows.Forms.ProgressBarStyle.Marquee
                                End If
                            End If

                            If statusLabel.Text <> latestText Then
                                statusLabel.Text = latestText
                                Dim availableWidth As System.Int32 = System.Math.Max(1, bodyPanel.ClientSize.Width - System.Windows.Forms.SystemInformation.VerticalScrollBarWidth)
                                Dim updatedHeight As System.Int32 = System.Math.Max(90, System.Math.Min(getLabelPreferred(availableWidth).Height, maxBodyHeightNoScroll))
                                table.RowStyles(0).Height = updatedHeight
                                localForm.ClientSize = New System.Drawing.Size(clientW, updatedHeight + spacerExtra + reservedBottomHeight + 2 * paddingAll)
                                table.PerformLayout()
                            End If

                            If closeRequested Then
                                refreshTimer.Stop()
                                localForm.Close()
                            End If
                        End Sub

                    SyncLock _syncRoot
                        _form = localForm
                    End SyncLock

                    SyncLock _syncRoot
                        If _closeRequested Then Return
                    End SyncLock
                    refreshTimer.Start()
                    localForm.Opacity = 1

                    Dim owner As System.Windows.Forms.IWin32Window = SharedMethods.ResolveSameThreadDialogOwner()
                    Dim ownerScope As System.IDisposable = Nothing

                    Try
                        ownerScope = SharedMethods.PushDialogOwner(localForm)

                        If owner IsNot Nothing Then
                            localForm.ShowDialog(owner)
                        Else
                            localForm.ShowDialog()
                        End If
                    Finally
                        If ownerScope IsNot Nothing Then
                            Try
                                ownerScope.Dispose()
                            Catch
                            End Try
                        End If
                    End Try

                Catch ex As System.Exception
                    HandleUiFailure(ex)
                Finally
                    ' Every UI/GDI resource is disposed on its creation STA, including
                    ' initialization failures and a close before ShowDialog starts.
                    Try
                        If refreshTimer IsNot Nothing Then
                            refreshTimer.Stop()
                            refreshTimer.Dispose()
                        End If
                        If localForm IsNot Nothing Then localForm.Dispose()
                        If ownedIcon IsNot Nothing Then ownedIcon.Dispose()
                        If bmpIcon IsNot Nothing Then bmpIcon.Dispose()
                        If standardFont IsNot Nothing Then standardFont.Dispose()
                    Finally
                        If handlerInstalled Then RemoveHandler System.Windows.Forms.Application.ThreadException, exceptionHandler
                        Dim workerContext As System.Windows.Forms.WindowsFormsSynchronizationContext =
                            TryCast(System.Threading.SynchronizationContext.Current, System.Windows.Forms.WindowsFormsSynchronizationContext)
                        Try
                            System.Windows.Forms.Application.ExitThread()
                        Finally
                            If workerContext IsNot Nothing Then workerContext.Dispose()
                            System.Threading.SynchronizationContext.SetSynchronizationContext(Nothing)
                            SyncLock _syncRoot
                                _form = Nothing
                            End SyncLock
                            CompleteLifetime()
                        End Try
                    End Try
                End Try
            End Sub

            Private Sub HandleUiFailure(failure As System.Exception)
                Try
                    RiCrashLogger.LogException("EXTRACTION_PROGRESS_FAILED", failure)
                Catch
                    System.Diagnostics.Debug.WriteLine(failure.ToString())
                End Try
                _ready.TrySetException(failure)
                Try
                    RequestCancel()
                Catch
                    ' Cancellation callbacks must not escape the worker message loop.
                End Try
                Close()
            End Sub

            <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
            Private Shared Function DestroyProgressIcon(iconHandle As System.IntPtr) As System.Boolean
            End Function

            Public Sub Close()
                SyncLock _syncRoot
                    _closeRequested = True
                    _totalClock.Stop()
                    _requestClock.Stop()
                End SyncLock

                Dim localForm As System.Windows.Forms.Form = Nothing

                SyncLock _syncRoot
                    localForm = _form
                End SyncLock

                If localForm IsNot Nothing AndAlso localForm.IsHandleCreated Then
                    Try
                        localForm.BeginInvoke(New System.Action(Sub() localForm.Close()))
                    Catch
                    End Try
                End If

                _ready.TrySetResult(False)
                If _uiThread Is Nothing Then CompleteLifetime()
            End Sub

            Private Sub CompleteLifetime()
                Dim shouldDisposeCancellation As System.Boolean
                SyncLock _syncRoot
                    If _cleanupCompleted Then Return
                    _cleanupCompleted = True
                    _closeRequested = True
                    shouldDisposeCancellation = _disposeRequested
                End SyncLock
                _ready.TrySetResult(False)
                If shouldDisposeCancellation Then DisposeCancellation()
                _stopped.TrySetResult(True)
            End Sub

            Private Sub DisposeCancellation()
                SyncLock _syncRoot
                    If _cancellationDisposed Then Return
                    _cancellationDisposed = True
                End SyncLock
                _cancellation.Dispose()
            End Sub

            Public Sub Dispose() Implements System.IDisposable.Dispose
                Dim cleanupCompleted As System.Boolean
                SyncLock _syncRoot
                    _disposeRequested = True
                    cleanupCompleted = _cleanupCompleted
                End SyncLock
                Close()
                If cleanupCompleted Then DisposeCancellation()
            End Sub
        End Class
End Class
    Public NotInheritable Class OfficeUiThreadAwaitable
        Implements System.Runtime.CompilerServices.INotifyCompletion

        Private ReadOnly _context As System.Threading.SynchronizationContext
        Private ReadOnly _threadId As System.Int32
        Private ReadOnly _shutdownToken As System.Threading.CancellationToken
        Private _dispatchFailure As System.Runtime.ExceptionServices.ExceptionDispatchInfo

        Public Sub New(context As System.Threading.SynchronizationContext,
                       threadId As System.Int32,
                       shutdownToken As System.Threading.CancellationToken)
            _context = context
            _threadId = threadId
            _shutdownToken = shutdownToken
        End Sub

        Public Function GetAwaiter() As OfficeUiThreadAwaitable
            Return Me
        End Function

        Public ReadOnly Property IsCompleted As System.Boolean
            Get
                ' Missing initialization and cancellation fail synchronously in GetResult.
                Return _context Is Nothing OrElse _threadId <= 0 OrElse
                       _shutdownToken.IsCancellationRequested OrElse
                       System.Threading.Thread.CurrentThread.ManagedThreadId = _threadId
            End Get
        End Property

        Public Sub GetResult()
            _shutdownToken.ThrowIfCancellationRequested()
            If _dispatchFailure IsNot Nothing Then _dispatchFailure.Throw()
            If _context Is Nothing OrElse _threadId <= 0 Then
                Throw New System.InvalidOperationException("The Office UI thread has not been initialized.")
            End If
            If System.Threading.Thread.CurrentThread.ManagedThreadId <> _threadId Then
                Throw New System.InvalidOperationException("The continuation did not reach the captured Office UI thread.")
            End If
        End Sub

        Public Sub OnCompleted(continuation As System.Action) Implements System.Runtime.CompilerServices.INotifyCompletion.OnCompleted
            If continuation Is Nothing Then Throw New System.ArgumentNullException(NameOf(continuation))
            Dim completed As System.Int32 = 0
            Dim registration As System.Threading.CancellationTokenRegistration
            Dim resumeContinuation As System.Action =
                Sub()
                    ' A queued UI callback and shutdown may race. Resume exactly once.
                    If System.Threading.Interlocked.Exchange(completed, 1) <> 0 Then Return
                    Try
                        continuation.Invoke()
                    Finally
                        registration.Dispose()
                    End Try
                End Sub

            Try
                If _shutdownToken.CanBeCanceled Then
                    registration = _shutdownToken.Register(
                        Sub()
                            ' Never synchronously wait for or execute Office work during shutdown.
                            ' GetResult throws cancellation before any continuation UI/COM work.
                            System.Threading.ThreadPool.QueueUserWorkItem(Sub(state As System.Object) resumeContinuation.Invoke())
                        End Sub)
                End If
                If System.Threading.Volatile.Read(completed) = 0 Then
                    _context.Post(Sub(state As System.Object) resumeContinuation.Invoke(), Nothing)
                End If
            Catch ex As System.Exception
                _dispatchFailure = System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(ex)
                ' Rejected dispatch must fault the awaiting task, never strand it.
                resumeContinuation.Invoke()
            Finally
                ' Covers cancellation/dispatch completing before Register returned its handle.
                If System.Threading.Volatile.Read(completed) <> 0 Then registration.Dispose()
            End Try
        End Sub
    End Class
    Public NotInheritable Class HostModelessFormRegistry
        Private ReadOnly _uiThreadId As System.Int32 = System.Threading.Thread.CurrentThread.ManagedThreadId
        Private ReadOnly _registered As New System.Collections.Generic.List(Of System.Windows.Forms.Form)()
        Private _closing As System.Boolean
        Private ReadOnly _shutdownCts As New System.Threading.CancellationTokenSource()

        ' An explicit host-lifetime signal. Background STA tools may observe this
        ' token but must clean up their own windows on their own STA thread.
        Public ReadOnly Property ShutdownToken As System.Threading.CancellationToken
            Get
                Return _shutdownCts.Token
            End Get
        End Property

        Public ReadOnly Property IsClosing As System.Boolean
            Get
                Return _closing
            End Get
        End Property

        Private Shared Sub LogModelessShutdown(eventName As System.String, details As System.String)
            Try
                RiCrashLogger.LogMarker(eventName, details)
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine(eventName & ": " & details & "; log failed: " & ex.Message)
            End Try
        End Sub

        Private Sub RequireUiThread()
            If System.Threading.Thread.CurrentThread.ManagedThreadId <> _uiThreadId Then
                Throw New System.InvalidOperationException("Modeless Office windows must be managed on the owning Office UI thread.")
            End If
        End Sub

        Private Shared Function VerifiedOfficeOwner() As System.Windows.Forms.IWin32Window
            Dim hwnd As System.IntPtr = SharedMethods.GetHostWindowHandle()
            If hwnd = System.IntPtr.Zero OrElse Not SharedMethods.IsWindowOnCurrentThread(hwnd) Then
                Return Nothing
            End If
            Return SharedMethods.IfOwnerOnCurrentThread(New WindowWrapper(hwnd))
        End Function

        ''' <summary>Registers a Red Ink-owned form and returns a validated native Office owner.</summary>
        Public Function Register(form As System.Windows.Forms.Form) As System.Windows.Forms.IWin32Window
            RequireUiThread()
            If _closing Then Throw New System.InvalidOperationException("The Office host is closing; new modeless windows are forbidden.")
            If form Is Nothing OrElse form.IsDisposed Then Throw New System.ArgumentException("A live form is required.", NameOf(form))
            If Not _registered.Contains(form) Then
                AddHandler form.FormClosed, AddressOf RegisteredFormClosed
                AddHandler form.Disposed, AddressOf RegisteredFormDisposed
                _registered.Add(form)
            End If
            Dim owner As System.Windows.Forms.IWin32Window = VerifiedOfficeOwner()
            If owner Is Nothing Then
                LogModelessShutdown("MODELESS_HOST_OWNER_UNAVAILABLE", form.GetType().FullName & ": no verified same-thread Office HWND; explicit registry cleanup remains active.")
            End If
            Return owner
        End Function

        ''' <summary>Shows or raises a registered form, normally owned by the Office main window.</summary>
        Public Sub Show(form As System.Windows.Forms.Form)
            Dim owner As System.Windows.Forms.IWin32Window = Register(form)
            Try
                If Not form.Visible Then
                    If owner IsNot Nothing Then
                        form.Show(owner)
                    Else
                        form.Show()
                    End If
                Else
                    form.BringToFront()
                End If
            Catch
                If Not form.Visible Then Unregister(form)
                Throw
            End Try
        End Sub

        Private Sub RegisteredFormClosed(sender As System.Object, e As System.Windows.Forms.FormClosedEventArgs)
            Dim form As System.Windows.Forms.Form = TryCast(sender, System.Windows.Forms.Form)
            If form IsNot Nothing Then Unregister(form)
        End Sub

        Private Sub RegisteredFormDisposed(sender As System.Object, e As System.EventArgs)
            Dim form As System.Windows.Forms.Form = TryCast(sender, System.Windows.Forms.Form)
            If form IsNot Nothing Then Unregister(form)
        End Sub

        Private Sub Unregister(form As System.Windows.Forms.Form)
            If Not _registered.Contains(form) Then Return
            RemoveHandler form.Disposed, AddressOf RegisteredFormDisposed
            RemoveHandler form.FormClosed, AddressOf RegisteredFormClosed
            _registered.Remove(form)
        End Sub

        ''' <summary>
        ''' Closes only this add-in's registered modeless windows. A form whose FormClosing
        ''' handler cancels the close is force-disposed during host exit; failure is logged.
        ''' </summary>
        Public Sub CloseAll()
            RequireUiThread()
            If _closing Then Return
            _closing = True
            ' Cancel before closing forms. Polling workers can finish their own STA
            ' cleanup without a synchronous wait from the Office UI thread.
            Try
                _shutdownCts.Cancel()
            Catch ex As System.Exception
                LogModelessShutdown("MODELESS_HOST_CANCEL_FAILED", ex.ToString())
            End Try
            Dim forms As System.Windows.Forms.Form() = _registered.ToArray()
            For index As System.Int32 = forms.Length - 1 To 0 Step -1
                Dim form As System.Windows.Forms.Form = forms(index)
                Try
                    If Not form.IsDisposed Then form.Close()
                Catch ex As System.Exception
                    LogModelessShutdown("MODELESS_HOST_CLOSE_FAILED", form.GetType().FullName & ": " & ex.Message)
                Finally
                    Try
                        If Not form.IsDisposed Then
                            LogModelessShutdown("MODELESS_HOST_FORCE_DISPOSE", form.GetType().FullName)
                            form.Dispose()
                        End If
                    Catch ex As System.Exception
                        LogModelessShutdown("MODELESS_HOST_DISPOSE_FAILED", form.GetType().FullName & ": " & ex.Message)
                    End Try
                    Unregister(form)
                End Try
            Next
            _registered.Clear()
        End Sub
    End Class
End Namespace
