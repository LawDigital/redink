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
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Threading
Imports System.Windows.Forms

Namespace SharedLibrary
    Partial Public Class SharedMethods

        ''' <summary>
        ''' Thread-local stack of dialog owners. Each managed UI thread keeps its
        ''' own stack so background threads cannot accidentally inherit an owner.
        ''' </summary>
        Private Shared ReadOnly _ownerStack As New ThreadLocal(Of Stack(Of IWin32Window))(
            Function() New Stack(Of IWin32Window)())

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
            _ownerStack.Value.Push(owner)
            Return New OwnerScope(owner)
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
            If stack IsNot Nothing AndAlso stack.Count > 0 Then
                Dim top = stack.Peek()
                ' Defensive: skip disposed forms (e.g. if a caller forgot to pop).
                Dim asForm = TryCast(top, Form)
                If asForm Is Nothing OrElse Not asForm.IsDisposed Then
                    Return top
                End If
            End If

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

            Dim h As IntPtr
            Try
                h = owner.Handle
            Catch
                Return Nothing
            End Try
            If h = IntPtr.Zero Then Return Nothing

            Dim pid As Integer = 0
            Dim ownerThreadId As Integer = NativeMethods.GetWindowThreadProcessId(h, pid)
            If ownerThreadId = 0 Then Return Nothing

            If ownerThreadId = NativeMethods.GetCurrentThreadId() Then
                Return owner
            End If

            Return Nothing
        End Function

        ''' <summary>
        ''' RAII token returned by <see cref="PushDialogOwner"/>; pops the owner
        ''' from the thread-local stack on dispose. Tolerates being disposed more
        ''' than once and tolerates stack drift if a child also pushed/popped.
        ''' </summary>
        Private NotInheritable Class OwnerScope
            Implements IDisposable

            Private ReadOnly _expected As IWin32Window
            Private _disposed As Boolean

            Public Sub New(expected As IWin32Window)
                _expected = expected
            End Sub

            Public Sub Dispose() Implements IDisposable.Dispose
                If _disposed Then Return
                _disposed = True
                Try
                    Dim stack = _ownerStack.Value
                    If stack IsNot Nothing AndAlso stack.Count > 0 AndAlso
                       Object.ReferenceEquals(stack.Peek(), _expected) Then
                        stack.Pop()
                    End If
                Catch
                    ' Never throw from dispose.
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
End Namespace
