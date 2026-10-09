' Part of "Red Ink" (Ask Inky)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: HotkeyService.vb
' Purpose: Registers the configurable Windows global shortcut, with VDI-friendly
'          diagnostics, bounded compatibility fallback and reconnect recovery.
' Architecture: Win32 RegisterHotKey on the dedicated message-only UI window.
' Notes: Never intercepts keyboard events or overrides other applications' keys.
' =============================================================================
Friend NotInheritable Class HotkeyService
    Implements System.IDisposable

    Private ReadOnly windowHandle As System.IntPtr
    Private activeId As Integer
    Private activeModifiers As UInteger
    Private activeKey As Integer
    Private nextId As Integer = 1
    Private ReadOnly slotOffset As Integer

    Public Sub New(handle As System.IntPtr, Optional idOffset As Integer = 0)
        windowHandle = handle
        nextId = 1 + idOffset
        slotOffset = idOffset
    End Sub

    Public ReadOnly Property IsRegistered As Boolean
        Get
            Return activeId <> 0
        End Get
    End Property

    Public Shared Function IsValid(modifiers As UInteger, key As Integer) As Boolean
        Dim allowedKey As Boolean = (key >= 65 AndAlso key <= 90) OrElse (key >= 48 AndAlso key <= 57) OrElse (key >= 112 AndAlso key <= 122) OrElse (key >= 124 AndAlso key <= 135)
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

    Private Function TryRegister(id As Integer, modifiers As UInteger, key As Integer) As Boolean
        If Native.RegisterHotKey(windowHandle, id, modifiers Or Native.MOD_NOREPEAT, CUInt(key)) Then
            ClientLog.Write("hotkey.registered.norepeat")
            Return True
        End If
        Dim errorCode As Integer = System.Runtime.InteropServices.Marshal.GetLastWin32Error()
        ' Some VDI/older Windows hosts reject MOD_NOREPEAT (ERROR_INVALID_PARAMETER).
        ' Never retry a conflicting shortcut (ERROR_HOTKEY_ALREADY_REGISTERED).
        If errorCode = 87 Then
            ClientLog.Write("hotkey.norepeat.unsupported.retry-without-flag")
            If Native.RegisterHotKey(windowHandle, id, modifiers, CUInt(key)) Then
                ClientLog.Write("hotkey.registered.compatibility")
                Return True
            End If
            errorCode = System.Runtime.InteropServices.Marshal.GetLastWin32Error()
        End If
        ClientLog.Write("hotkey.registration.failed.win32." & errorCode.ToString(System.Globalization.CultureInfo.InvariantCulture))
        Return False
    End Function

    Public Function TryApply(config As ClientConfig, persist As Boolean) As Boolean
        If Not IsValid(config.HotkeyModifiers, config.HotkeyKey) Then
            ClientLog.Write("hotkey.registration.invalid-config")
            Return False
        End If
        If activeId <> 0 AndAlso config.HotkeyModifiers = activeModifiers AndAlso config.HotkeyKey = activeKey Then
            If persist Then ConfigStore.Save(config)
            Return True
        End If
        Dim candidateId As Integer = If(nextId = 1 + slotOffset, 1 + slotOffset, 2 + slotOffset)
        nextId = If(candidateId = 1 + slotOffset, 2 + slotOffset, 1 + slotOffset)
        If Not TryRegister(candidateId, config.HotkeyModifiers, config.HotkeyKey) Then Return False
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

    ' Called on the UI thread after Windows session unlock or remote reconnection.
    ' An existing registration may have been lost across a virtual-session transition.
    Public Function Rebind(config As ClientConfig) As Boolean
        If activeId <> 0 Then
            If Not Native.UnregisterHotKey(windowHandle, activeId) Then ClientLog.Write("hotkey.rebind.unregister.failed")
            activeId = 0
        End If
        ClientLog.Write("hotkey.rebind.attempt")
        Return TryApply(config, False)
    End Function

    Public Sub Dispose() Implements System.IDisposable.Dispose
        If activeId <> 0 AndAlso Not Native.UnregisterHotKey(windowHandle, activeId) Then ClientLog.Write("hotkey.unregister.failed")
        activeId = 0
    End Sub
End Class
