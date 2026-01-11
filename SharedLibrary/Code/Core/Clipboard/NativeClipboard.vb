' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: NativeClipboard.vb
' Purpose: Provides P/Invoke declarations for Win32 clipboard and GDI functions used
'          by this library to read Enhanced Metafile (EMF) data from the clipboard.
'
' Architecture:
'  - Clipboard access (user32.dll):
'     - OpenClipboard / CloseClipboard: Opens/closes the clipboard for the calling thread.
'     - IsClipboardFormatAvailable: Checks if a specific clipboard format is present.
'     - GetClipboardData: Retrieves a handle to the clipboard data in the requested format.
'  - EMF handle management (gdi32.dll):
'     - CopyEnhMetaFile: Duplicates an EMF handle (optionally writing it to a file when a path is supplied).
'     - DeleteEnhMetaFile: Frees an EMF handle created/duplicated by GDI functions.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Runtime.InteropServices

Namespace SharedLibrary

    Friend Module NativeClipboardX

        ''' <summary>
        ''' Clipboard format identifier for Enhanced Metafile (CF_ENHMETAFILE).
        ''' </summary>
        Friend Const CF_ENHMETAFILE As UInteger = 14

        ''' <summary>
        ''' Opens the clipboard for examination and prevents other applications from modifying it.
        ''' </summary>
        ''' <param name="hWnd">Handle to the window to be associated with the open clipboard (may be <see cref="IntPtr.Zero"/>).</param>
        ''' <returns><see langword="True"/> on success; otherwise <see langword="False"/>.</returns>
        <DllImport("user32.dll", SetLastError:=True)>
        Friend Function OpenClipboard(hWnd As IntPtr) As Boolean
        End Function

        ''' <summary>
        ''' Closes the clipboard.
        ''' </summary>
        ''' <returns><see langword="True"/> on success; otherwise <see langword="False"/>.</returns>
        <DllImport("user32.dll")>
        Friend Function CloseClipboard() As Boolean
        End Function

        ''' <summary>
        ''' Checks whether the specified clipboard format is available.
        ''' </summary>
        ''' <param name="fmt">Clipboard format identifier.</param>
        ''' <returns><see langword="True"/> if the format is available; otherwise <see langword="False"/>.</returns>
        <DllImport("user32.dll")>
        Friend Function IsClipboardFormatAvailable(fmt As UInteger) As Boolean
        End Function

        ''' <summary>
        ''' Retrieves a handle to data in the specified clipboard format.
        ''' </summary>
        ''' <param name="fmt">Clipboard format identifier.</param>
        ''' <returns>
        ''' A handle to the clipboard data if successful; otherwise <see cref="IntPtr.Zero"/>.
        ''' </returns>
        <DllImport("user32.dll")>
        Friend Function GetClipboardData(fmt As UInteger) As IntPtr
        End Function

        ''' <summary>
        ''' Duplicates an enhanced metafile handle. If <paramref name="lpszFile"/> is provided,
        ''' the function may write the metafile to the specified file.
        ''' </summary>
        ''' <param name="hEmfSrc">Handle to the source enhanced metafile.</param>
        ''' <param name="lpszFile">Optional file path, or <see langword="Nothing"/>.</param>
        ''' <returns>A handle to the copied enhanced metafile, or <see cref="IntPtr.Zero"/> on failure.</returns>
        <DllImport("gdi32.dll")>
        Friend Function CopyEnhMetaFile(hEmfSrc As IntPtr,
                                       lpszFile As String) As IntPtr
        End Function

        ''' <summary>
        ''' Deletes an enhanced metafile handle.
        ''' </summary>
        ''' <param name="hemf">Handle to the enhanced metafile to delete.</param>
        ''' <returns><see langword="True"/> on success; otherwise <see langword="False"/>.</returns>
        <DllImport("gdi32.dll")>
        Friend Function DeleteEnhMetaFile(hemf As IntPtr) As Boolean
        End Function

    End Module

End Namespace