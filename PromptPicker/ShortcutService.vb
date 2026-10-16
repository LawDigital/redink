' Hotkey implementation aligned with AskInky.HotkeyService: two-slot apply so a
' conflicting new hotkey cannot unregister an already-working old shortcut.
Friend NotInheritable Class ShortcutService
    Implements System.IDisposable
    Private ReadOnly handle As System.IntPtr
    Private activeId As Integer
    Private activeModifiers As UInteger
    Private activeKey As Integer
    Private nextId As Integer = 45723

    Public Sub New(windowHandle As System.IntPtr)
        handle = windowHandle
    End Sub

    Public ReadOnly Property IsRegistered As Boolean
        Get
            Return activeId <> 0
        End Get
    End Property

    Public Shared Function IsValid(modifiers As UInteger, key As Integer) As Boolean
        Dim isKey As Boolean = (key >= 65 AndAlso key <= 90) OrElse (key >= 48 AndAlso key <= 57) OrElse (key >= 112 AndAlso key <= 135)
        Return isKey AndAlso (modifiers And (Native.MOD_CONTROL Or Native.MOD_ALT)) <> 0UI AndAlso (modifiers And Not 7UI) = 0UI
    End Function

    Public Shared Function Describe(modifiers As UInteger, key As Integer) As String
        Return If((modifiers And Native.MOD_CONTROL) <> 0UI, "Ctrl + ", "") &
               If((modifiers And Native.MOD_SHIFT) <> 0UI, "Shift + ", "") &
               If((modifiers And Native.MOD_ALT) <> 0UI, "Alt + ", "") &
               CType(key, System.Windows.Forms.Keys).ToString()
    End Function

    Private Function TryRegister(id As Integer, settings As PickerSettings) As Boolean
        If Native.RegisterHotKey(handle, id, settings.HotkeyModifiers Or Native.MOD_NOREPEAT, CUInt(settings.HotkeyKey)) Then Return True
        Dim err As Integer = System.Runtime.InteropServices.Marshal.GetLastWin32Error()
        If err = 87 Then
            If Native.RegisterHotKey(handle, id, settings.HotkeyModifiers, CUInt(settings.HotkeyKey)) Then Return True
            err = System.Runtime.InteropServices.Marshal.GetLastWin32Error()
        End If
        AppLog.Write("hotkey.register.failed.win32." & err.ToString(System.Globalization.CultureInfo.InvariantCulture))
        Return False
    End Function

    Public Function Apply(value As PickerSettings) As Boolean
        If Not IsValid(value.HotkeyModifiers, value.HotkeyKey) Then Return False
        If activeId <> 0 AndAlso activeModifiers = value.HotkeyModifiers AndAlso activeKey = value.HotkeyKey Then Return True
        Dim candidate As Integer = nextId
        nextId = If(candidate = 45723, 45724, 45723)
        If Not TryRegister(candidate, value) Then Return False
        If activeId <> 0 Then Native.UnregisterHotKey(handle, activeId)
        activeId = candidate
        activeModifiers = value.HotkeyModifiers
        activeKey = value.HotkeyKey
        Return True
    End Function

    Public Function Rebind(value As PickerSettings) As Boolean
        If activeId <> 0 Then
            Native.UnregisterHotKey(handle, activeId)
            activeId = 0
        End If
        Return Apply(value)
    End Function

    Public Function IsMessage(id As Integer) As Boolean
        Return id = activeId AndAlso activeId <> 0
    End Function

    Public Sub Dispose() Implements System.IDisposable.Dispose
        If activeId <> 0 Then Native.UnregisterHotKey(handle, activeId)
        activeId = 0
    End Sub
End Class
