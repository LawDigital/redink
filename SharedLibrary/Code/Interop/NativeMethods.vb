' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: NativeMethods.vb
' Purpose: Provides minimal P/Invoke declarations for calling Windows (Win32) APIs
'          required by this library.
'
' Architecture / How it works:
'  - Exposes `Shared` (static) interop methods in a dedicated type to keep unmanaged
'    imports centralized and easy to audit.
'  - Uses `DllImport` to declare the native `user32.dll` function `SetForegroundWindow`.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary

    ''' <summary>
    ''' Win32 API declarations used by this library.
    ''' </summary>
    Public Class NativeMethods

        <Runtime.InteropServices.DllImport("user32.dll")>
        Public Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
        End Function

        <Runtime.InteropServices.DllImport("user32.dll")>
        Public Shared Function GetForegroundWindow() As IntPtr
        End Function

        <Runtime.InteropServices.DllImport("user32.dll")>
        Public Shared Function GetWindowThreadProcessId(hWnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
        End Function

        <Runtime.InteropServices.DllImport("kernel32.dll")>
        Public Shared Function GetCurrentThreadId() As Integer
        End Function

        <Runtime.InteropServices.DllImport("user32.dll")>
        Public Shared Function AllowSetForegroundWindow(dwProcessId As Integer) As Boolean
        End Function

        <Runtime.InteropServices.DllImport("user32.dll")>
        Public Shared Function SetWindowPos(hWnd As IntPtr, hWndInsertAfter As IntPtr, X As Integer, Y As Integer, cx As Integer, cy As Integer, uFlags As UInteger) As Boolean
        End Function

    End Class

End Namespace
