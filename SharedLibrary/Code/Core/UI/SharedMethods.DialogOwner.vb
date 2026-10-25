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

        ''' <summary>
        ''' Attaches a lightweight, self-disposing watchdog to a TopMost shared dialog
        ''' so it can surface a same-process foreground window (e.g. Word's native
        ''' "Save changes?" prompt) that would otherwise stay hidden behind it.
        '''
        ''' A single Deactivate check is unreliable: Windows may report
        ''' GetForegroundWindow() as zero or still our own handle while activation is
        ''' in transit, so the foreign prompt can be missed permanently. This helper
        ''' polls on a short WinForms timer that runs on the dialog's own UI thread
        ''' through the message pump ShowDialog is already spinning, then calls the
        ''' existing PromoteForeignForegroundDialog logic. The timer is created,
        ''' started, and — via the dialog's FormClosed event — stopped and disposed
        ''' automatically, so callers do not have to change their ShowDialog control
        ''' flow. Ownership and TopMost policy are untouched. Safe no-op on failure.
        ''' </summary>
        Public Shared Sub AttachForeignForegroundWatchdog(dialog As Form)
            If dialog Is Nothing Then Return
            Try
                Dim watchdog As New System.Windows.Forms.Timer()
                watchdog.Interval = 100

                Dim tick As EventHandler = Nothing
                tick =
                    Sub(s As Object, ev As System.EventArgs)
                        PromoteForeignForegroundDialog(dialog)
                    End Sub

                Dim closed As System.Windows.Forms.FormClosedEventHandler = Nothing
                closed =
                    Sub(s As Object, ev As System.Windows.Forms.FormClosedEventArgs)
                        Try
                            RemoveHandler dialog.FormClosed, closed
                            RemoveHandler watchdog.Tick, tick
                            watchdog.Stop()
                            watchdog.Dispose()
                        Catch
                        End Try
                    End Sub

                AddHandler watchdog.Tick, tick
                AddHandler dialog.FormClosed, closed
                watchdog.Start()
            Catch
                ' Never throw from dialog setup.
            End Try
        End Sub


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
        ''' and foreground) so the user can act on it. It only touches windows of our
        ''' own process and does NOT change the dialog's TopMost state or its owner,
        ''' so it introduces no Z-order/owner regressions. Safe no-op on failure.
        ''' </summary>
        Public Shared Sub PromoteForeignForegroundDialog(dialog As Form)
            If dialog Is Nothing Then Return
            Try
                System.Diagnostics.Debug.WriteLine("[PromoteForeign] Deactivate fired.")

                Dim fg As IntPtr = NativeMethods.GetForegroundWindow()
                System.Diagnostics.Debug.WriteLine("[PromoteForeign] fg=" & fg.ToString() & " dialog=" & dialog.Handle.ToString())
                If fg = IntPtr.Zero OrElse fg = dialog.Handle Then
                    System.Diagnostics.Debug.WriteLine("[PromoteForeign] EXIT: fg is zero or is our own dialog.")
                    Return
                End If

                ' Restrict to windows of our own (host) process, e.g. Word's prompts.
                Dim fgPid As Integer = 0
                NativeMethods.GetWindowThreadProcessId(fg, fgPid)
                If fgPid = 0 Then
                    System.Diagnostics.Debug.WriteLine("[PromoteForeign] EXIT: fgPid is 0.")
                    Return
                End If

                Dim ourPid As Integer = System.Diagnostics.Process.GetCurrentProcess().Id
                System.Diagnostics.Debug.WriteLine("[PromoteForeign] fgPid=" & fgPid & " ourPid=" & ourPid)
                If fgPid <> ourPid Then
                    System.Diagnostics.Debug.WriteLine("[PromoteForeign] EXIT: foreground window belongs to a different process.")
                    Return
                End If

                ' An expanded ComboBox's native list is an intentional, same-process
                ' popup, not a competing Office dialog. Preserve the parent TopMost
                ' state while it is open; otherwise the list can disappear immediately.
                If ContainsExpandedComboBox(dialog) Then
                    System.Diagnostics.Debug.WriteLine("[PromoteForeign] EXIT: dialog ComboBox dropdown is open.")
                    Return
                End If

                System.Diagnostics.Debug.WriteLine("[PromoteForeign] Proceeding to move/demote our dialog.")

                ' A same-process window (e.g. Word's native "Save changes?" or
                ' "You cannot close Microsoft Word because a dialogue box is open"
                ' prompt) has taken the foreground. Because our dialog is TopMost, it
                ' shares the topmost band with — and keeps covering — that prompt, so
                ' merely re-promoting the prompt does not surface it. Instead, drop our
                ' own dialog out of the topmost band so the foreign prompt can come to
                ' the front, and remember to restore TopMost once we are re-activated
                ' (i.e. after the user dismisses that prompt).
                If dialog.TopMost Then
                    dialog.TopMost = False

                    ' One-shot restore: when our dialog regains activation, put it back
                    ' into the topmost band so later focus changes still keep it above
                    ' the host. Detaches itself so it only fires once per demotion.
                    Dim restore As EventHandler = Nothing
                    restore =
                        Sub(s As Object, ev As System.EventArgs)
                            Try
                                RemoveHandler dialog.Activated, restore
                                dialog.TopMost = True
                            Catch
                            End Try
                        End Sub
                    AddHandler dialog.Activated, restore
                End If

                Const HWND_TOP As Integer = 0
                Const SWP_NOMOVE As UInteger = &H2UI
                Const SWP_NOSIZE As UInteger = &H1UI
                Const SWP_SHOWWINDOW As UInteger = &H40UI

                NativeMethods.SetWindowPos(fg, New IntPtr(HWND_TOP), 0, 0, 0, 0,
                                           SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)
                NativeMethods.SetForegroundWindow(fg)
            Catch
                ' Never throw from an event-driven watchdog.
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

        Private Sub Unregister(form As System.Windows.Forms.Form)
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
