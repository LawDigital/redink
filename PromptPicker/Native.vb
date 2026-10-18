Friend Module Native
    Public Const WM_HOTKEY As Integer = &H312
    Public Const WM_REBIND As Integer = &H8000 + &H141
    Public Const VK_LBUTTON As Integer = &H1
    Public Const VK_RBUTTON As Integer = &H2
    Public Const VK_MBUTTON As Integer = &H4
    Public Const VK_CONTROL As Integer = &H11
    Public Const VK_MENU As Integer = &H12
    Public Const VK_SHIFT As Integer = &H10
    Public Const MOD_ALT As UInteger = 1UI
    Public Const MOD_CONTROL As UInteger = 2UI
    Public Const MOD_SHIFT As UInteger = 4UI
    Public Const MOD_NOREPEAT As UInteger = &H4000UI
    Public Const INPUT_KEYBOARD As UInteger = 1UI
    Public Const KEYEVENTF_KEYUP As UInteger = &H2UI
    Public Const VK_V As UShort = &H56US

    <System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)>
    Public Structure NativeRectangle
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure

    <System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)>
    Public Structure GuiThreadInfo
        Public Size As UInteger
        Public Flags As UInteger
        Public Active As System.IntPtr
        Public Focus As System.IntPtr
        Public Capture As System.IntPtr
        Public MenuOwner As System.IntPtr
        Public MoveSize As System.IntPtr
        Public Caret As System.IntPtr
        Public CaretRectangle As NativeRectangle
    End Structure

    ' SendInput is process-independent, unlike SendKeys (which can dispatch to
    ' a changed active control after the WebView2 dialog closes).
    <System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)>
    Public Structure InputRecord
        Public InputType As UInteger
        Public Data As InputData
    End Structure

    <System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Explicit)>
    Public Structure InputData
        <System.Runtime.InteropServices.FieldOffset(0)>
        Public Keyboard As KeyboardRecord
        ' The native INPUT union is sized by MOUSEINPUT (32 bytes on x64).
        ' Without it, SendInput rejects the 32-byte keyboard-only record.
        <System.Runtime.InteropServices.FieldOffset(0)>
        Public Mouse As MouseRecord
    End Structure

    <System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)>
    Public Structure MouseRecord
        Public Dx As Integer
        Public Dy As Integer
        Public MouseData As UInteger
        Public Flags As UInteger
        Public Time As UInteger
        Public ExtraInfo As System.IntPtr
    End Structure

    <System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)>
    Public Structure KeyboardRecord
        Public VirtualKey As UShort
        Public ScanCode As UShort
        Public Flags As UInteger
        Public Time As UInteger
        Public ExtraInfo As System.IntPtr
    End Structure

    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
    Public Function SendInput(count As UInteger, input As InputRecord(), recordSize As Integer) As UInteger
    End Function

    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
    Public Function RegisterHotKey(hwnd As System.IntPtr, id As Integer, modifiers As UInteger, key As UInteger) As Boolean
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
    Public Function UnregisterHotKey(hwnd As System.IntPtr, id As Integer) As Boolean
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Public Function GetForegroundWindow() As System.IntPtr
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Public Function GetWindowThreadProcessId(hwnd As System.IntPtr, ByRef processId As UInteger) As UInteger
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
    Public Function GetGUIThreadInfo(threadId As UInteger, ByRef info As GuiThreadInfo) As Boolean
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Public Function IsChild(parent As System.IntPtr, child As System.IntPtr) As Boolean
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
    Public Function AttachThreadInput(first As UInteger, second As UInteger, attach As Boolean) As Boolean
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
    Public Function SetFocus(hwnd As System.IntPtr) As System.IntPtr
    End Function
    <System.Runtime.InteropServices.DllImport("kernel32.dll")>
    Public Function GetCurrentThreadId() As UInteger
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Public Function GetClipboardSequenceNumber() As UInteger
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Public Function GetAsyncKeyState(key As Integer) As Short
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Public Function SetForegroundWindow(hwnd As System.IntPtr) As Boolean
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
    Public Function PostMessage(hwnd As System.IntPtr, message As Integer, wParam As System.IntPtr, lParam As System.IntPtr) As Boolean
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Public Function IsWindow(hwnd As System.IntPtr) As Boolean
    End Function
End Module
