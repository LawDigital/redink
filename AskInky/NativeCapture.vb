' Part of "Red Ink" (Ask Inky)
' Copyright (c) LawDigital AG, Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: NativeCapture.vb
' Purpose: Win32 declarations for robust selection capture, clipboard snapshots and paste-back.
' Architecture: Windows native host with embedded Office Web-based interface.
' Notes: Red Ink License and Ask Inky third-party attributions are embedded in LICENSE.txt.
' =============================================================================
' Part of Red Ink (AskInky). P18 W367: SendInput replaces SendKeys; clipboard snapshots copy raw bytes.
Friend Module NativeCapture
    Public Const INPUT_KEYBOARD As UInteger = 1UI
    Public Const KEYEVENTF_EXTENDEDKEY As UInteger = 1UI
    Public Const KEYEVENTF_KEYUP As UInteger = 2UI
    Public Const VK_C As UShort = &H43US
    Public Const VK_V As UShort = &H56US
    Public Const VK_INSERT As UShort = &H2DUS
    Public Const VK_CONTROL_U As UShort = &H11US
    Public Const GMEM_MOVEABLE As UInteger = 2UI
    Public Const EM_GETSEL As Integer = &HB0
    Public Const WM_GETTEXT As Integer = &HD
    Public Const WM_GETTEXTLENGTH As Integer = &HE
    Public Const SMTO_ABORTIFHUNG As UInteger = 2UI
    Public Const PROCESS_QUERY_LIMITED_INFORMATION As UInteger = &H1000UI
    Public Const TOKEN_QUERY As UInteger = 8UI
    Public Const TokenElevationClass As Integer = 20

    <System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)>
    Public Structure KEYBDINPUT
        Public wVk As UShort
        Public wScan As UShort
        Public dwFlags As UInteger
        Public time As UInteger
        Public dwExtraInfo As System.IntPtr
    End Structure

    ' INPUT on x64: DWORD type + 4 bytes padding + union (largest member MOUSEINPUT = 32 bytes).
    <System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Explicit, Size:=40)>
    Public Structure INPUT
        <System.Runtime.InteropServices.FieldOffset(0)> Public type As UInteger
        <System.Runtime.InteropServices.FieldOffset(8)> Public ki As KEYBDINPUT
    End Structure

    <System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)>
    Public Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure

    <System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)>
    Public Structure GUITHREADINFO
        Public cbSize As Integer
        Public flags As Integer
        Public hwndActive As System.IntPtr
        Public hwndFocus As System.IntPtr
        Public hwndCapture As System.IntPtr
        Public hwndMenuOwner As System.IntPtr
        Public hwndMoveSize As System.IntPtr
        Public hwndCaret As System.IntPtr
        Public rcCaret As RECT
    End Structure

    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
    Public Function SendInput(count As UInteger, <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPArray)> inputs As INPUT(), size As Integer) As UInteger
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
    Public Function OpenClipboard(owner As System.IntPtr) As Boolean
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
    Public Function CloseClipboard() As Boolean
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
    Public Function EmptyClipboard() As Boolean
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
    Public Function EnumClipboardFormats(format As UInteger) As UInteger
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
    Public Function GetClipboardData(format As UInteger) As System.IntPtr
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
    Public Function SetClipboardData(format As UInteger, data As System.IntPtr) As System.IntPtr
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True, CharSet:=System.Runtime.InteropServices.CharSet.Unicode)>
    Public Function RegisterClipboardFormat(name As String) As UInteger
    End Function
    <System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True)>
    Public Function GlobalLock(handle As System.IntPtr) As System.IntPtr
    End Function
    <System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True)>
    Public Function GlobalUnlock(handle As System.IntPtr) As Boolean
    End Function
    <System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True)>
    Public Function GlobalSize(handle As System.IntPtr) As System.UIntPtr
    End Function
    <System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True)>
    Public Function GlobalAlloc(flags As UInteger, bytes As System.UIntPtr) As System.IntPtr
    End Function
    <System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True)>
    Public Function GlobalFree(handle As System.IntPtr) As System.IntPtr
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
    Public Function GetGUIThreadInfo(threadId As UInteger, ByRef info As GUITHREADINFO) As Boolean
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True, CharSet:=System.Runtime.InteropServices.CharSet.Unicode)>
    Public Function GetClassName(hWnd As System.IntPtr, buffer As System.Text.StringBuilder, size As Integer) As Integer
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True, CharSet:=System.Runtime.InteropServices.CharSet.Unicode, EntryPoint:="SendMessageTimeoutW")>
    Public Function SendMessageTimeout(hWnd As System.IntPtr, msg As Integer, wParam As System.IntPtr, lParam As System.IntPtr, flags As UInteger, timeout As UInteger, ByRef result As System.IntPtr) As System.IntPtr
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True, CharSet:=System.Runtime.InteropServices.CharSet.Unicode, EntryPoint:="SendMessageTimeoutW")>
    Public Function SendMessageTimeoutText(hWnd As System.IntPtr, msg As Integer, wParam As System.IntPtr, buffer As System.Text.StringBuilder, flags As UInteger, timeout As UInteger, ByRef result As System.IntPtr) As System.IntPtr
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
    Public Function SendMessageTimeoutRef(hWnd As System.IntPtr, msg As Integer, ByRef wParam As Integer, ByRef lParam As Integer, flags As UInteger, timeout As UInteger, ByRef result As System.IntPtr) As System.IntPtr
    End Function
    <System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True)>
    Public Function OpenProcess(access As UInteger, inherit As Boolean, processId As UInteger) As System.IntPtr
    End Function
    <System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True)>
    Public Function CloseHandle(handle As System.IntPtr) As Boolean
    End Function
    <System.Runtime.InteropServices.DllImport("advapi32.dll", SetLastError:=True)>
    Public Function OpenProcessToken(process As System.IntPtr, access As UInteger, ByRef token As System.IntPtr) As Boolean
    End Function
    <System.Runtime.InteropServices.DllImport("advapi32.dll", SetLastError:=True)>
    Public Function GetTokenInformation(token As System.IntPtr, infoClass As Integer, ByRef elevation As Integer, length As Integer, ByRef returned As Integer) As Boolean
    End Function
    <System.Runtime.InteropServices.DllImport("kernel32.dll")>
    Public Function GetCurrentProcess() As System.IntPtr
    End Function
End Module
