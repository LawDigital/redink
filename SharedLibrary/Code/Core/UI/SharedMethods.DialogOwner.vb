' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.DialogOwner.vb
' Purpose:
'   Provides an ambient, thread-local stack of IWin32Window owners for modal
'   dialogs created by the shared Show* helpers (SelectValue, ShowCustomMessageBox,
'   ShowCustomYesNoBox, ShowCustomVariableInputForm, ShowCustomInputBox,
'   ShowHTMLCustomMessageBox, ShowRTFCustomMessageBox, ShowTextFileEditor, ...).
'
'   Why:
'     Several shared dialogs historically call ShowDialog() with no owner. When
'     they are spawned from a WinForms form that is TopMost (e.g. DiscussInky,
'     HelpMeInky, Form1, Outlook chat forms), the unowned modal child ends up
'     behind the parent. With an ambient owner pushed by the parent form, every
'     Show* helper transparently parents the dialog to that form so the Z-order
'     is correct, without changing call sites.
'
'   How:
'     - Callers wrap dialog-spawning code in:
'           Using SharedMethods.PushDialogOwner(Me)
'               ShowCustomMessageBox(...)
'               SelectValue(...)
'           End Using
'     - Each Show* helper resolves the owner via SharedMethods.ResolveDialogOwner()
'       and passes the result to ShowDialog(owner).
'     - Resolution order:
'         1. The top of the thread-local stack (the form that pushed itself).
'         2. The Office host window hwnd (Word/Excel/Outlook) via GetOfficeApplicationHwnd.
'         3. Nothing (caller falls back to ShowDialog() with no owner).
'
'   IMPORTANT (cross-host modal deadlock prevention):
'     Never pass a foreign or unverified owner directly to ShowDialog(owner)
'     (e.g. New WindowWrapper(ownerHandle), Form.ActiveForm, or a caller-supplied
'     owner/ownerForm/ownerHandle). Such an owner may belong to another UI thread
'     or Office host process; ShowDialog disables it and may never re-enable it,
'     freezing that host. Before using any such owner, call
'     OfficeWindowWatchdog.InspectDialogOwner(...) to log the attempt and
'     SharedMethods.IfOwnerOnCurrentThread(...) to reject a cross-thread owner
'     (falling back to an ownerless dialog). For ambient owners, prefer
'     SharedMethods.ResolveSameThreadDialogOwner(); do NOT call the internal
'     ResolveDialogOwner() directly for ShowDialog. See the "Modal Dialog
'     Ownership (Cross-Host Deadlock Prevention)" section in CONTRIBUTING.md.
' =============================================================================
'
' KEY ARCHITECTURE
'   • Surface: SharedMethods, DialogOwnerEntry (VB implementation).
'   • Entry points/helpers: InitializeHostWindow, IsWindowInCurrentProcess,
'     IsWindowOnCurrentThread.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Threading
Imports System.Windows.Forms

Namespace SharedLibrary
    Partial Public Class SharedMethods

        ' Host identity is configured by each host on its UI thread. No Office COM
        ' is used here, so recovery is also safe from the diagnostic timer.
        Private Shared ReadOnly _hostWindowLock As New System.Object()
        Private Shared _hostWindowClass As System.String
        Private Shared _hostWindowThread As System.Int32
        Private Shared _hostWindowHandle As System.IntPtr

        Public Shared Sub InitializeHostWindow(windowClass As System.String,
                                               Optional candidate As System.IntPtr = Nothing)
            SyncLock _hostWindowLock
                _hostWindowClass = windowClass
                _hostWindowThread = NativeMethods.GetCurrentThreadId()
                _hostWindowHandle = If(IsHostWindowCandidate(candidate), candidate, System.IntPtr.Zero)
            End SyncLock
        End Sub

        Public Shared Function IsWindowInCurrentProcess(hwnd As System.IntPtr) As System.Boolean
            If hwnd = System.IntPtr.Zero OrElse Not NativeMethods.IsWindow(hwnd) Then Return False
            Dim pid As System.Int32 = 0
            If NativeMethods.GetWindowThreadProcessId(hwnd, pid) = 0 Then Return False
            Using process As System.Diagnostics.Process = System.Diagnostics.Process.GetCurrentProcess()
                Return pid = process.Id
            End Using
        End Function

        Public Shared Function IsWindowOnCurrentThread(hwnd As System.IntPtr) As System.Boolean
            If Not IsWindowInCurrentProcess(hwnd) Then Return False
            Dim pid As System.Int32 = 0
            Return NativeMethods.GetWindowThreadProcessId(hwnd, pid) = NativeMethods.GetCurrentThreadId()
        End Function

        Private Shared Function IsHostWindowCandidate(hwnd As System.IntPtr) As System.Boolean
            If System.String.IsNullOrEmpty(_hostWindowClass) OrElse Not IsWindowInCurrentProcess(hwnd) Then Return False
            Dim pid As System.Int32 = 0
            If NativeMethods.GetWindowThreadProcessId(hwnd, pid) <> _hostWindowThread Then Return False
            Dim name As New System.Text.StringBuilder(256)
            NativeMethods.GetClassName(hwnd, name, name.Capacity)
            Return System.String.Equals(name.ToString(), _hostWindowClass, System.StringComparison.Ordinal)
        End Function

        Public Shared Function GetHostWindowHandle() As System.IntPtr
            SyncLock _hostWindowLock
                If IsHostWindowCandidate(_hostWindowHandle) AndAlso NativeMethods.IsWindowVisible(_hostWindowHandle) Then
                    Return _hostWindowHandle
                End If
                Dim previous As System.IntPtr = _hostWindowHandle
                _hostWindowHandle = System.IntPtr.Zero
                If Not System.String.IsNullOrEmpty(_hostWindowClass) Then
                    ' Never search another process or guess another Office application.
                    NativeMethods.EnumWindows(
                        Function(hwnd As System.IntPtr, state As System.IntPtr) As System.Boolean
                            If IsHostWindowCandidate(hwnd) AndAlso NativeMethods.IsWindowVisible(hwnd) Then
                                _hostWindowHandle = hwnd
                                Return False
                            End If
                            Return True
                        End Function, System.IntPtr.Zero)
                End If
                If previous <> System.IntPtr.Zero AndAlso previous <> _hostWindowHandle Then
                    RiCrashLogger.LogMarker("HOST_WINDOW_RESELECTED",
                        "PreviousHWND=" & previous.ToString() & "; CurrentHWND=" & _hostWindowHandle.ToString() &
                        "; Reason=previous window no longer eligible; same-process/UI-thread/class recovery only.")
                End If
                Return _hostWindowHandle
            End SyncLock
        End Function

        Public Shared Sub SetHostWindowHandle(candidate As System.IntPtr)
            SyncLock _hostWindowLock
                ' Compatibility setter; startup must configure the class/thread first.
                _hostWindowHandle = If(IsHostWindowCandidate(candidate), candidate, System.IntPtr.Zero)
            End SyncLock
        End Sub

        ''' <summary>
        ''' Thread-local stack of dialog owners. Each managed UI thread keeps its
        ''' own stack so background threads cannot accidentally inherit an owner.
        ''' </summary>
        ' Entries are unique per PushDialogOwner scope. Modeless forms may close out of
        ' creation order, so matching only the top owner would leave a stale form alive
        ' in the ambient stack. Never infer ownership from window titles or processes.
        Private NotInheritable Class DialogOwnerEntry
            Public ReadOnly Property Owner As System.Windows.Forms.IWin32Window

            Public Sub New(ownerWindow As System.Windows.Forms.IWin32Window)
                Owner = ownerWindow
            End Sub
        End Class

        Private Shared ReadOnly _ownerStack As New System.Threading.ThreadLocal(Of System.Collections.Generic.Stack(Of DialogOwnerEntry))(
            Function() New System.Collections.Generic.Stack(Of DialogOwnerEntry)())

        ''' <summary>
        ''' Pushes <paramref name="owner"/> onto the ambient dialog-owner stack for
        ''' the current thread. The returned token must be disposed (typically via
        ''' a <c>Using</c> block) to pop the owner again.
        ''' </summary>
        ''' <param name="owner">
        ''' The window to use as owner for shared modal dialogs while the token is
        ''' alive. If <c>Nothing</c>, the call is a no-op and the returned token
        ''' does nothing on dispose.
        ''' </param>
        Public Shared Function PushDialogOwner(owner As IWin32Window) As IDisposable
            If owner Is Nothing Then Return New NullOwnerScope()
            Dim entry As New DialogOwnerEntry(owner)
            _ownerStack.Value.Push(entry)
            Return New OwnerScope(entry)
        End Function

        ''' <summary>
        ''' Resolves the ambient dialog owner candidate for the current thread.
        ''' Returns the top of the thread-local stack if any, else a
        ''' <see cref="WindowWrapper"/> around the Office host window if known,
        ''' otherwise <c>Nothing</c>.
        ''' </summary>
        ''' <remarks>
        ''' Internal helper only. Callers should normally use
        ''' <see cref="ResolveSameThreadDialogOwner"/> before passing an owner to
        ''' <c>ShowDialog(owner)</c>. This method may return a window that belongs
        ''' to another UI thread or Office host process; using such an owner
        ''' directly can disable that foreign window and cause modal deadlocks.
        ''' This method exists so <see cref="ResolveSameThreadDialogOwner"/> can
        ''' first resolve the best candidate and then reject it when it is not on
        ''' the current thread.
        ''' </remarks>
        Private Shared Function ResolveDialogOwner() As IWin32Window
            Dim resolved As IWin32Window = ResolveDialogOwnerCore()
            OfficeWindowWatchdog.InspectDialogOwner(resolved, "ResolveDialogOwner", Nothing)
            Return resolved
        End Function

        Private Shared Function ResolveDialogOwnerCore() As IWin32Window
            Dim stack = _ownerStack.Value
            While stack IsNot Nothing AndAlso stack.Count > 0
                Dim top As DialogOwnerEntry = stack.Peek()
                Dim asForm As System.Windows.Forms.Form = TryCast(top.Owner, System.Windows.Forms.Form)
                ' A handle can be destroyed before the scope is disposed. Skip all
                ' stale top entries rather than hiding the next still-live owner.
                If top.Owner Is Nothing OrElse
                   (asForm IsNot Nothing AndAlso (asForm.IsDisposed OrElse asForm.Disposing)) Then
                    stack.Pop()
                    Continue While
                End If
                Return top.Owner
            End While

            Dim hwnd As IntPtr = GetOfficeApplicationHwnd()
            If hwnd <> IntPtr.Zero Then
                Return New WindowWrapper(hwnd)
            End If

            Return Nothing
        End Function

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

        ''' <summary>
        ''' Resolves the ambient dialog owner but returns it ONLY when its window
        ''' belongs to the current (calling) thread. A modal dialog shown with a
        ''' cross-thread owner attaches the two threads' input queues; on close,
        ''' Windows must synchronously re-enable and re-activate the owner window on
        ''' its thread, which can deadlock the host if that thread is simultaneously
        ''' blocked marshalling work back to the caller. Rejecting a cross-thread
        ''' owner (falling back to an ownerless, TopMost dialog) removes that hazard
        ''' deterministically without changing the same-thread behavior callers rely
        ''' on for correct Z-order.
        ''' </summary>
        Public Shared Function ResolveSameThreadDialogOwner() As IWin32Window
            Dim owner As IWin32Window = ResolveDialogOwner()
            Return IfOwnerOnCurrentThread(owner)
        End Function


        ''' <summary>
        ''' Forces an ownerless, TopMost dialog to the foreground reliably, including
        ''' when it was created on a background/STA thread where Windows' foreground
        ''' lock would otherwise keep it behind the active Office window. Safe no-op
        ''' on failure. Does NOT assign a cross-thread owner, so it cannot reintroduce
        ''' the modal-close deadlock. Use this from a dialog's Shown handler instead
        ''' of the TopMost=False/TopMost=True toggle, which drops the topmost band and
        ''' can let an ownerless window fall behind the host.
        ''' </summary>
        Public Shared Sub ForceDialogToForeground(dialog As Form)
            If dialog Is Nothing Then Return
            Try
                dialog.TopMost = True
                NativeMethods.AllowSetForegroundWindow(-1)
                NativeMethods.SetForegroundWindow(dialog.Handle)
                dialog.Activate()
                dialog.BringToFront()
            Catch
            End Try
        End Sub

        ' One state per form, shared by the timer and any Deactivate handler. All
        ' methods run on the dialog's own UI thread. Never marshal to the Office STA.
        Private Shared ReadOnly _foregroundWatches As New System.Runtime.CompilerServices.ConditionalWeakTable(Of System.Windows.Forms.Form, ForegroundWatch)()
        Private Shared ReadOnly _watchProcessId As System.Int32 = ReadWatchProcessId()

        Private Shared Function ReadWatchProcessId() As System.Int32
            Using process As System.Diagnostics.Process = System.Diagnostics.Process.GetCurrentProcess()
                Return process.Id
            End Using
        End Function

        Public Shared Sub AttachForeignForegroundWatchdog(dialog As System.Windows.Forms.Form)
            If dialog Is Nothing OrElse dialog.Disposing OrElse dialog.IsDisposed OrElse dialog.InvokeRequired Then Return
            Try
                _foregroundWatches.GetValue(dialog, Function(owner) New ForegroundWatch(owner)).Start()
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("[DialogWatchdog] Cannot start foreground monitoring: " & ex.ToString())
            End Try
        End Sub

        Private NotInheritable Class ForegroundWatch
            Private ReadOnly _form As System.Windows.Forms.Form
            Private _timer As System.Windows.Forms.Timer
            Private _lastTarget As System.IntPtr
            Private _retry As System.Diagnostics.Stopwatch
            Private _processing As System.Boolean
            Private _restoreTopMost As System.Boolean
            Private _reportedError As System.Boolean
            Private _stopped As System.Boolean

            Friend Sub New(form As System.Windows.Forms.Form)
                _form = form
                AddHandler _form.FormClosed, AddressOf Closed
                AddHandler _form.Disposed, AddressOf Disposed
                AddHandler _form.Activated, AddressOf Activated
            End Sub

            Friend Sub Start()
                If _stopped OrElse _timer IsNot Nothing Then Return
                _timer = New System.Windows.Forms.Timer() With {.Interval = 100}
                Try
                    AddHandler _timer.Tick, AddressOf Tick
                    _timer.Start()
                Catch
                    StopMonitoring()
                    Throw
                End Try
            End Sub

            Private Sub Tick(sender As System.Object, e As System.EventArgs)
                CheckForeground()
            End Sub

            Private Sub Activated(sender As System.Object, e As System.EventArgs)
                _lastTarget = System.IntPtr.Zero
                If Not _restoreTopMost OrElse _stopped Then Return
                _restoreTopMost = False
                Try
                    _form.TopMost = True
                Catch ex As System.Exception
                    ReportOnce(ex.ToString())
                End Try
            End Sub

            Private Sub Closed(sender As System.Object, e As System.Windows.Forms.FormClosedEventArgs)
                StopMonitoring()
            End Sub

            Private Sub Disposed(sender As System.Object, e As System.EventArgs)
                StopMonitoring()
            End Sub

            Private Sub StopMonitoring()
                If _stopped Then Return
                _stopped = True
                RemoveHandler _form.FormClosed, AddressOf Closed
                RemoveHandler _form.Disposed, AddressOf Disposed
                RemoveHandler _form.Activated, AddressOf Activated
                If _timer IsNot Nothing Then
                    RemoveHandler _timer.Tick, AddressOf Tick
                    _timer.Stop()
                    _timer.Dispose()
                    _timer = Nothing
                End If
                _foregroundWatches.Remove(_form)
            End Sub

            Private Sub ReportOnce(message As System.String)
                If _reportedError Then Return
                _reportedError = True
                System.Diagnostics.Debug.WriteLine("[DialogWatchdog] " & message)
            End Sub

            Friend Sub CheckForeground()
                If _processing OrElse _stopped OrElse _form.IsDisposed OrElse _form.Disposing OrElse
                    Not _form.IsHandleCreated OrElse _form.InvokeRequired OrElse Not _form.Visible Then Return
                _processing = True
                Try
                    Dim target As System.IntPtr = NativeMethods.GetForegroundWindow()
                    ' No stable foreground during activation hand-off. No work or log.
                    If target = System.IntPtr.Zero OrElse target = _form.Handle Then
                        _lastTarget = System.IntPtr.Zero
                        Return
                    End If
                    Dim pid As System.Int32 = 0
                    If NativeMethods.GetWindowThreadProcessId(target, pid) = 0 OrElse pid <> _watchProcessId Then
                        _lastTarget = System.IntPtr.Zero
                        Return
                    End If
                    If target = _lastTarget AndAlso Not _form.TopMost Then
                        If _retry Is Nothing OrElse _retry.ElapsedMilliseconds < 1000 Then Return
                    End If
                    If ContainsExpandedComboBox(_form) Then
                        _lastTarget = System.IntPtr.Zero
                        Return
                    End If
                    If NativeMethods.GetForegroundWindow() <> target Then Return
                    _lastTarget = target
                    _retry = Nothing
                    If _form.TopMost Then
                        _restoreTopMost = True
                        _form.TopMost = False
                        System.Diagnostics.Debug.WriteLine("[PromoteForeign] Yielded TopMost to a foreground window of this Office process.")
                    End If
                    ' The target already has focus. Raise its Z-order, not its input
                    ' focus, and never wait on a different input queue.
                    If _stopped OrElse NativeMethods.GetForegroundWindow() <> target Then
                        _lastTarget = System.IntPtr.Zero
                        Return
                    End If
                    Const flags As System.UInt32 = &H1UI Or &H2UI Or &H10UI Or &H4000UI ' NOSIZE/NOMOVE/NOACTIVATE/ASYNCWINDOWPOS
                    If Not NativeMethods.SetWindowPos(target, System.IntPtr.Zero, 0, 0, 0, 0, flags) Then
                        _retry = System.Diagnostics.Stopwatch.StartNew()
                        ReportOnce("Cannot raise the same-process foreground window; retrying while it remains in front.")
                    End If
                Catch ex As System.Exception
                    _retry = System.Diagnostics.Stopwatch.StartNew()
                    ReportOnce(ex.ToString())
                Finally
                    _processing = False
                End Try
            End Sub
        End Class


        ''' <summary>
        ''' A WinForms ComboBox uses a separate native popup HWND while its list is open.
        ''' That HWND can become the foreground window even though it belongs to the
        ''' dialog's own UI, and must not be treated as a foreign Office dialog.
        ''' </summary>
        Private Shared Function ContainsExpandedComboBox(root As System.Windows.Forms.Control) As System.Boolean
            If root Is Nothing OrElse root.IsDisposed Then Return False
            Dim combo As System.Windows.Forms.ComboBox = TryCast(root, System.Windows.Forms.ComboBox)
            If combo IsNot Nothing AndAlso combo.DroppedDown Then Return True
            For Each child As System.Windows.Forms.Control In root.Controls
                If ContainsExpandedComboBox(child) Then Return True
            Next
            Return False
        End Function

        ''' <summary>
        ''' When a TopMost shared dialog is showing and a window belonging to the
        ''' SAME (host) process — e.g. Word's native "Save changes?" prompt raised
        ''' while the user is closing the application — steals the input focus, that
        ''' native prompt can end up hidden behind our TopMost dialog. It then holds
        ''' focus but is invisible, so the UI appears deadlocked.
        '''
        ''' This helper detects that situation from the dialog's Deactivate event and
        ''' promotes the foreign foreground window above our dialog (making it visible
        ''' without stealing focus) so the user can act on it. It temporarily yields
        ''' TopMost, restores it on reactivation and never changes the dialog owner.
        ''' </summary>
        Public Shared Sub PromoteForeignForegroundDialog(dialog As System.Windows.Forms.Form)
            If dialog Is Nothing OrElse dialog.Disposing OrElse dialog.IsDisposed OrElse
                Not dialog.IsHandleCreated OrElse dialog.InvokeRequired Then Return
            Try
                _foregroundWatches.GetValue(dialog, Function(owner) New ForegroundWatch(owner)).CheckForeground()
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("[DialogWatchdog] Cannot check foreground: " & ex.ToString())
            End Try
        End Sub

        Public Shared Function IfOwnerOnCurrentThread(owner As IWin32Window) As IWin32Window
            If owner Is Nothing Then Return Nothing

            Dim control As System.Windows.Forms.Control = TryCast(owner, System.Windows.Forms.Control)
            If control IsNot Nothing AndAlso (control.IsDisposed OrElse control.Disposing OrElse
                Not control.IsHandleCreated OrElse control.InvokeRequired) Then
                RiCrashLogger.LogMarker("DIALOG_OWNER_REJECTED",
                    "Reason=disposed, uncreated or cross-thread control; using ownerless fallback.")
                Return Nothing
            End If
            Dim h As System.IntPtr
            Try
                h = owner.Handle
            Catch
                Return Nothing
            End Try
            If h = IntPtr.Zero Then Return Nothing

            If IsWindowOnCurrentThread(h) Then Return owner
            OfficeWindowWatchdog.InspectDialogOwner(owner, "RejectedOwner", "IfOwnerOnCurrentThread")
            RiCrashLogger.LogMarker("DIALOG_OWNER_REJECTED",
                "OwnerHWND=" & h.ToString() & "; Reason=invalid, foreign-process or foreign-thread owner; using ownerless fallback.")
            Return Nothing
        End Function

        ''' <summary>
        ''' RAII token returned by <see cref="PushDialogOwner"/>; pops the owner
        ''' from the thread-local stack on dispose. Tolerates being disposed more
        ''' than once and tolerates stack drift if a child also pushed/popped.
        ''' </summary>
        Private NotInheritable Class OwnerScope
            Implements IDisposable

            Private ReadOnly _expected As DialogOwnerEntry
            Private ReadOnly _threadId As System.Int32 = System.Threading.Thread.CurrentThread.ManagedThreadId
            Private _disposed As Boolean

            Public Sub New(expected As DialogOwnerEntry)
                _expected = expected
            End Sub

            Public Sub Dispose() Implements IDisposable.Dispose
                If _disposed Then Return
                ' The stack is thread-local. Disposing it on another thread must
                ' never touch that thread's unrelated dialog owners.
                If System.Threading.Thread.CurrentThread.ManagedThreadId <> _threadId Then
                    Try
                        RiCrashLogger.LogMarker("DIALOG_OWNER_SCOPE_WRONG_THREAD",
                            "Dialog owner scope disposal requires its creating UI thread.")
                    Catch ex As System.Exception
                        System.Diagnostics.Debug.WriteLine("Dialog owner scope disposal on wrong thread: " & ex.Message)
                    End Try
                    Return
                End If
                _disposed = True
                Try
                    Dim stack As System.Collections.Generic.Stack(Of DialogOwnerEntry) = _ownerStack.Value
                    If stack Is Nothing OrElse stack.Count = 0 Then Return
                    If Object.ReferenceEquals(stack.Peek(), _expected) Then
                        stack.Pop()
                        Return
                    End If

                    ' Remove precisely this scope without disturbing newer scopes.
                    ' In particular, two scopes can refer to the SAME form instance.
                    Dim newer As New System.Collections.Generic.Stack(Of DialogOwnerEntry)()
                    While stack.Count > 0
                        Dim current As DialogOwnerEntry = stack.Pop()
                        If Object.ReferenceEquals(current, _expected) Then Exit While
                        newer.Push(current)
                    End While
                    While newer.Count > 0
                        stack.Push(newer.Pop())
                    End While
                Catch ex As System.Exception
                    Try
                        RiCrashLogger.LogMarker("DIALOG_OWNER_SCOPE_RELEASE_FAILED", ex.ToString())
                    Catch
                        System.Diagnostics.Debug.WriteLine("Dialog owner scope release failed: " & ex.Message)
                    End Try
                End Try
            End Sub
        End Class

        ''' <summary>No-op token used when a null owner is pushed.</summary>
        Private NotInheritable Class NullOwnerScope
            Implements IDisposable
            Public Sub Dispose() Implements IDisposable.Dispose
            End Sub
        End Class

    End Class

    ''' <summary>
    ''' Awaitable that posts the continuation itself to a captured Office UI thread.
    ''' Completing a Task on that thread does not switch an awaiting MTA caller to it.
    ''' Capture context and thread ID on the host STA; this type never invents a UI
    ''' context or accepts a different thread. It is intended for one Await per instance.
    ''' Host shutdown completes pending awaits as cancelled even if the UI stops pumping.
    ''' </summary>
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

    ''' <summary>
    ''' Per-Office-add-in lifetime for explicitly registered modeless windows. Never enumerates
    ''' Application.OpenForms (which may include windows belonging to another add-in).
    ''' Registration and shutdown run on the Office UI thread; the validated Office HWND is
    ''' used as the native owner when available, so a host window's destruction also closes
    ''' modeless windows if the VSTO Shutdown event is not delivered.
    ''' </summary>
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
