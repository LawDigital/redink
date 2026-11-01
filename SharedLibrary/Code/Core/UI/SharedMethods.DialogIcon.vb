' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.DialogIcon.vb
' PURPOSE
'   Creates and applies owned Windows Forms dialog icons with safe native HICON lifetime.
'
' KEY ARCHITECTURE
'   • Ownership: Clone Icon.FromHandle before releasing the native HICON.
'   • Lifecycle: ApplyStandardDialogIcon leases icons safely to Windows Forms.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary
    Partial Public Class SharedMethods
        ' GetHicon allocates an HICON. Icon.FromHandle only borrows it. Never use the
        ' borrowed wrapper as a Form.Icon or destroy a shared/caller-supplied HICON.
        Friend Shared Function CreateOwnedDialogIcon() As System.Drawing.Icon
            Dim ownsSource As System.Boolean = False
            Dim source As System.Drawing.Bitmap = ResolveLogoBitmap(LogoType.Standard, False, ownsSource)
            Try
                Using working As New System.Drawing.Bitmap(source)
                    Dim handle As System.IntPtr = working.GetHicon()
                    Try
                        Using borrowed As System.Drawing.Icon = System.Drawing.Icon.FromHandle(handle)
                            Return DirectCast(borrowed.Clone(), System.Drawing.Icon)
                        End Using
                    Finally
                        If Not NativeMethods.DestroyIcon(handle) Then
                            Dim code As System.Int32 = System.Runtime.InteropServices.Marshal.GetLastWin32Error()
                            System.Diagnostics.Trace.TraceWarning("[DialogIcon] DestroyIcon failed (Win32 {0}).", code)
                        End If
                    End Try
                End Using
            Finally
                If ownsSource AndAlso source IsNot Nothing Then source.Dispose()
            End Try
        End Function

        ' Called on the form's creation thread, before or after Show. No handle is
        ' requested for thread dispatch and no Office/COM object is accessed.
        Private Shared ReadOnly _dialogIcons As New System.Runtime.CompilerServices.ConditionalWeakTable(Of System.Windows.Forms.Form, FormIconLease)()

        Public Shared Sub ApplyStandardDialogIcon(form As System.Windows.Forms.Form)
            If form Is Nothing Then Throw New System.ArgumentNullException(NameOf(form))
            If form.Disposing OrElse form.IsDisposed Then Return
            If form.IsHandleCreated AndAlso form.InvokeRequired Then Throw New System.InvalidOperationException("Set the dialog icon on the form's owning thread.")
            Dim replacement As System.Drawing.Icon = Nothing
            Try
                replacement = CreateOwnedDialogIcon()
                Dim lease As FormIconLease = _dialogIcons.GetValue(form, Function(owner) New FormIconLease(owner))
                form.Icon = replacement
                lease.Adopt(replacement)
                replacement = Nothing
            Catch ex As System.Exception When TypeOf ex Is System.ArgumentException OrElse
                                              TypeOf ex Is System.IO.IOException OrElse
                                              TypeOf ex Is System.Runtime.InteropServices.ExternalException
                ' Branding is optional: do not cancel an import/translation just because
                ' an image could not be decoded. Do not mask threading or binding errors.
                System.Diagnostics.Trace.TraceWarning("[DialogIcon] {0}: keeping the existing icon ({1}: {2}).", form.GetType().Name, ex.GetType().Name, ex.Message)
            Finally
                If replacement IsNot Nothing Then replacement.Dispose()
            End Try
        End Sub

        Private NotInheritable Class FormIconLease
            Private _owned As System.Drawing.Icon
            Private _disposed As System.Boolean
            Friend Sub New(owner As System.Windows.Forms.Form)
                AddHandler owner.Disposed, AddressOf Released
            End Sub
            Friend Sub Adopt(icon As System.Drawing.Icon)
                If _disposed Then
                    icon.Dispose()
                    Return
                End If
                Dim previous As System.Drawing.Icon = _owned
                _owned = icon
                If previous IsNot Nothing Then previous.Dispose()
            End Sub
            Private Sub Released(sender As System.Object, e As System.EventArgs)
                _disposed = True
                Dim owner As System.Windows.Forms.Form = DirectCast(sender, System.Windows.Forms.Form)
                RemoveHandler owner.Disposed, AddressOf Released
                _dialogIcons.Remove(owner)
                If _owned IsNot Nothing Then _owned.Dispose()
                _owned = Nothing
            End Sub
        End Class
    End Class
End Namespace
