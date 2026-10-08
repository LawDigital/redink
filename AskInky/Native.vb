' Part of Red Ink (AskInky). No keyboard/mouse hook and no text-insertion API.
Friend Module Native
    Public Const WM_HOTKEY As Integer = &H312
    Public Const VK_CONTROL As Integer = &H11
    Public Const VK_MENU As Integer = &H12
    Public Const VK_SHIFT As Integer = &H10
    Public Const MOD_CONTROL As UInteger = 2UI
    Public Const MOD_SHIFT As UInteger = 4UI
    Public Const MOD_ALT As UInteger = 1UI
    Public Const MOD_NOREPEAT As UInteger = &H4000UI

    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
    Public Function RegisterHotKey(hWnd As System.IntPtr, id As Integer, modifiers As UInteger, key As UInteger) As Boolean
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
    Public Function UnregisterHotKey(hWnd As System.IntPtr, id As Integer) As Boolean
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Public Function GetClipboardSequenceNumber() As UInteger
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Public Function GetAsyncKeyState(key As Integer) As Short
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Public Function GetForegroundWindow() As System.IntPtr
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Public Function GetWindowThreadProcessId(window As System.IntPtr, ByRef processId As UInteger) As UInteger
    End Function
End Module
