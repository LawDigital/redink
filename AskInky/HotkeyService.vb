' Part of Red Ink (AskInky). Register candidates before releasing the working shortcut.
Friend NotInheritable Class HotkeyService
    Implements System.IDisposable
    Private ReadOnly windowHandle As System.IntPtr
    Private activeId As Integer
    Private activeModifiers As UInteger
    Private activeKey As Integer

    Public Sub New(handle As System.IntPtr)
        windowHandle = handle
    End Sub

    Public Shared Function IsValid(modifiers As UInteger, key As Integer) As Boolean
        Dim allowedKey As Boolean = (key >= 65 AndAlso key <= 90) OrElse (key >= 48 AndAlso key <= 57) OrElse (key >= 112 AndAlso key <= 122) OrElse (key >= 124 AndAlso key <= 135)
        ' No bare key, Shift-only letter, Windows-key combination, or debugger-reserved F12.
        Return allowedKey AndAlso (modifiers And 3UI) <> 0UI AndAlso (modifiers And Not 7UI) = 0UI
    End Function

    Public Shared Function Format(modifiers As UInteger, key As Integer) As String
        Return If((modifiers And Native.MOD_CONTROL) <> 0UI, "Ctrl + ", "") &
               If((modifiers And Native.MOD_SHIFT) <> 0UI, "Shift + ", "") &
               If((modifiers And Native.MOD_ALT) <> 0UI, "Alt + ", "") &
               If(key >= 48 AndAlso key <= 57, System.Convert.ToChar(key).ToString(), CType(key, System.Windows.Forms.Keys).ToString())
    End Function

    Public Function IsActive(id As Integer) As Boolean
        Return activeId <> 0 AndAlso id = activeId
    End Function

    Public Function TryApply(config As ClientConfig, persist As Boolean) As Boolean
        If Not IsValid(config.HotkeyModifiers, config.HotkeyKey) Then Return False
        If activeId <> 0 AndAlso config.HotkeyModifiers = activeModifiers AndAlso config.HotkeyKey = activeKey Then
            If persist Then ConfigStore.Save(config)
            Return True
        End If
        Dim candidateId As Integer = If(activeId = 1, 2, 1)
        If Not Native.RegisterHotKey(windowHandle, candidateId, config.HotkeyModifiers Or Native.MOD_NOREPEAT, CUInt(config.HotkeyKey)) Then
            ClientLog.Write("hotkey.registration.failed." & System.Runtime.InteropServices.Marshal.GetLastWin32Error().ToString(System.Globalization.CultureInfo.InvariantCulture))
            Return False
        End If
        Try
            If persist Then ConfigStore.Save(config)
        Catch ex As System.Exception
            Native.UnregisterHotKey(windowHandle, candidateId)
            Throw
        End Try
        If activeId <> 0 AndAlso Not Native.UnregisterHotKey(windowHandle, activeId) Then ClientLog.Write("hotkey.unregister.failed")
        activeId = candidateId
        activeModifiers = config.HotkeyModifiers
        activeKey = config.HotkeyKey
        Return True
    End Function

    Public Sub Dispose() Implements System.IDisposable.Dispose
        If activeId <> 0 AndAlso Not Native.UnregisterHotKey(windowHandle, activeId) Then ClientLog.Write("hotkey.unregister.failed")
        activeId = 0
    End Sub
End Class
