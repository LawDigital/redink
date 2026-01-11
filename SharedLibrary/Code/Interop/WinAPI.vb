' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: WinAPI.vb
' Purpose: Provides minimal P/Invoke declarations for calling Windows (Win32) APIs
'          required by this library.
'
' Architecture / How it works:
'  - Exposes `Public` Win32 API declarations in a dedicated module to keep unmanaged
'    imports centralized and easy to audit.
'  - Uses `DllImport` to declare the native `user32.dll` function `FindWindow`.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Runtime.InteropServices

Namespace SharedLibrary

    ''' <summary>
    ''' Win32 API declarations used by this library.
    ''' </summary>
    Module WinAPI

        ''' <summary>
        ''' Retrieves a handle to the top-level window whose class name and/or window name matches the specified strings.
        ''' </summary>
        ''' <param name="lpClassName">The window class name. Use <see langword="Nothing" /> to ignore the class name.</param>
        ''' <param name="lpWindowName">The window name (title). Use <see langword="Nothing" /> to ignore the window name.</param>
        ''' <returns>
        ''' A window handle if found; otherwise <see cref="IntPtr.Zero" />.
        ''' </returns>
        <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
        Public Function FindWindow(lpClassName As String, lpWindowName As String) As IntPtr
        End Function

    End Module

End Namespace