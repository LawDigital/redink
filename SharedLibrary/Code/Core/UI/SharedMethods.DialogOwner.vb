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
        ''' Resolves the best available owner for a shared modal dialog on the
        ''' current thread. Returns the top of the thread-local stack if any, else
        ''' a <see cref="WindowWrapper"/> around the Office host window if known,
        ''' otherwise <c>Nothing</c>.
        ''' </summary>
        Public Shared Function ResolveDialogOwner() As IWin32Window
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
