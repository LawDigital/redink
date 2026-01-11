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

        ''' <summary>
        ''' Brings the specified window to the foreground (activates it).
        ''' </summary>
        ''' <param name="hWnd">Handle to the window to activate.</param>
        ''' <returns><see langword="True" /> if the window was brought to the foreground; otherwise <see langword="False" />.</returns>
        <Runtime.InteropServices.DllImport("user32.dll")>
        Public Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
        End Function

    End Class

End Namespace