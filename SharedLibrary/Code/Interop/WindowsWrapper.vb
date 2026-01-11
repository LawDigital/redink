' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: WindowsWrapper.vb
' Purpose: Provides a managed wrapper around a native window handle (HWND) by
'          implementing `IWin32Window` for consumption by Windows Forms APIs.
'
' Architecture / How it works:
'  - Stores an `IntPtr` window handle provided by the caller.
'  - Exposes the handle via `IWin32Window.Handle` so WinForms dialogs/components
'    can be parented to an existing native window.
' =============================================================================



Option Strict On
Option Explicit On

Imports System.Windows.Forms

Namespace SharedLibrary

    ''' <summary>
    ''' Wraps a native window handle (HWND) to provide <see cref="IWin32Window" /> implementation
    ''' for use with managed Windows Forms APIs that require parent window references.
    ''' </summary>
    ''' <remarks>
    ''' This class is particularly useful when hosting dialogs from Office add-ins,
    ''' where the Office application window handle must be wrapped to ensure proper
    ''' modal dialog parenting and Z-order behavior.
    ''' </remarks>
    Public Class WindowWrapper
        Implements System.Windows.Forms.IWin32Window

        ''' <summary>
        ''' Stores the native window handle (HWND) represented by this instance.
        ''' </summary>
        Private _hwnd As IntPtr

        ''' <summary>
        ''' Initializes a new instance of the <see cref="WindowWrapper" /> class with the specified window handle.
        ''' </summary>
        ''' <param name="handle">The native window handle (HWND) to wrap.</param>
        ''' <remarks>
        ''' The caller is responsible for ensuring the handle remains valid for the lifetime
        ''' of this wrapper instance.
        ''' </remarks>
        Public Sub New(handle As IntPtr)
            _hwnd = handle
        End Sub

        ''' <summary>
        ''' Gets the native window handle (HWND) represented by this wrapper.
        ''' </summary>
        ''' <value>The native window handle (HWND).</value>
        Public ReadOnly Property Handle As IntPtr Implements IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property

    End Class

End Namespace