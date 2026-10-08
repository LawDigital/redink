' Part of "Red Ink" (Ask Inky)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: AssistantContext.vb
' Purpose: Ask Inky desktop functionality in the modular Windows client.
' Architecture: Windows native host with embedded Office Web-based interface.
' Notes: Red Ink License and Ask Inky third-party attributions are embedded in LICENSE.txt.
' =============================================================================
' Part of Red Ink (AskInky). Tray lifecycle and one configurable global keyboard shortcut only.
Public NotInheritable Class AssistantContext
    Inherits System.Windows.Forms.ApplicationContext
    Private ReadOnly host As HotkeyWindow
    Private ReadOnly hotkeys As HotkeyService
    Private ReadOnly tray As System.Windows.Forms.NotifyIcon
    Private ReadOnly menu As System.Windows.Forms.ContextMenuStrip
    Private form As AssistantForm
    Private settingsForm As SettingsForm
    Private config As ClientConfig
    Private capturing As Boolean
    Private busy As Boolean
    Private exiting As Boolean

    Public Sub New()
        config = ConfigStore.Load()
        host = New HotkeyWindow(Me)
        hotkeys = New HotkeyService(host.Handle)
        menu = New System.Windows.Forms.ContextMenuStrip With {.ShowImageMargin = False}
        menu.Items.Add("", Nothing, Sub(sender As Object, e As System.EventArgs) ShowInky(False))
        menu.Items.Add("", Nothing, Sub(sender As Object, e As System.EventArgs) ShowSettings())
        menu.Items.Add("", Nothing, Sub(sender As Object, e As System.EventArgs) ExitThread())
        tray = New System.Windows.Forms.NotifyIcon With {.Text = "Red Ink", .ContextMenuStrip = menu, .Icon = System.Drawing.Icon.ExtractAssociatedIcon(System.Windows.Forms.Application.ExecutablePath), .Visible = True}
        AddHandler tray.MouseClick, Sub(sender As Object, e As System.Windows.Forms.MouseEventArgs)
                                        If e.Button = System.Windows.Forms.MouseButtons.Left Then menu.Show(System.Windows.Forms.Cursor.Position)
                                    End Sub
        ApplyMenuLanguage()
        If Not hotkeys.TryApply(config, False) Then tray.ShowBalloonTip(6000, "Red Ink", UiStrings.GetText("desktop.hotkeyConflict", config.UiLanguage), System.Windows.Forms.ToolTipIcon.Warning)
        If ConfigStore.LoadFailed Then tray.ShowBalloonTip(6000, "Red Ink", UiStrings.GetText("desktop.configLoadFailed", config.UiLanguage), System.Windows.Forms.ToolTipIcon.Warning)
    End Sub

    Private Sub ApplyMenuLanguage()
        menu.Items(0).Text = UiStrings.GetText("desktop.open", config.UiLanguage)
        DirectCast(menu.Items(0), System.Windows.Forms.ToolStripMenuItem).ShortcutKeyDisplayString = HotkeyService.Format(config.HotkeyModifiers, config.HotkeyKey)
        menu.Items(1).Text = UiStrings.GetText("settings", config.UiLanguage)
        menu.Items(2).Text = UiStrings.GetText("desktop.exit", config.UiLanguage)
    End Sub

    Friend Sub SetBusy(value As Boolean)
        busy = value
        If Not exiting Then menu.Items(1).Enabled = Not value
    End Sub

    Friend Async Sub ShowInky(captureSelection As Boolean)
        If exiting OrElse capturing Then Return
        If settingsForm IsNot Nothing AndAlso Not settingsForm.IsDisposed Then
            settingsForm.Activate()
            Return
        End If
        capturing = True
        Try
            Dim source As System.IntPtr = Native.GetForegroundWindow()
            Dim process As UInteger
            Native.GetWindowThreadProcessId(source, process)
            Dim capture As Boolean = captureSelection AndAlso Not busy AndAlso source <> System.IntPtr.Zero AndAlso process <> CUInt(System.Environment.ProcessId)
            Dim selected As String = Nothing
            If capture Then selected = Await SelectionService.CaptureAsync(source)
            If exiting Then Return
            If form Is Nothing OrElse form.IsDisposed Then form = New AssistantForm(Me)
            If capture Then form.LoadSelection(selected)
            If form.WindowState = System.Windows.Forms.FormWindowState.Minimized Then form.WindowState = System.Windows.Forms.FormWindowState.Normal
            form.Show()
            form.Activate()
            form.FocusQuestion()
        Catch ex As System.Exception
            ClientLog.Write("assistant.open.failed", ex)
            tray.ShowBalloonTip(5000, "Red Ink", HelperClient.Describe(ex, config.UiLanguage), System.Windows.Forms.ToolTipIcon.Warning)
        Finally
            capturing = False
        End Try
    End Sub

    Friend Sub ShowSettings()
        If exiting OrElse busy Then Return
        If settingsForm IsNot Nothing AndAlso Not settingsForm.IsDisposed Then
            settingsForm.Activate()
            Return
        End If
        settingsForm = New SettingsForm(Me, config)
        AddHandler settingsForm.FormClosed, Sub(sender As Object, e As System.Windows.Forms.FormClosedEventArgs)
                                                settingsForm = Nothing
                                                If form IsNot Nothing AndAlso Not form.IsDisposed Then form.SetSettingsOpen(False)
                                            End Sub
        If form IsNot Nothing AndAlso Not form.IsDisposed AndAlso form.Visible Then
            form.SetSettingsOpen(True)
            settingsForm.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
            ' Explicit same-thread local Form only. Never adopt the source application's HWND.
            settingsForm.Show(form)
        Else
            settingsForm.Show()
        End If
        settingsForm.Activate()
    End Sub

    Friend Function TryCommit(candidate As ClientConfig) As Boolean
        If Not hotkeys.TryApply(candidate, True) Then Return False
        config = candidate
        If Not config.ShowLlmDebug Then
            HelperClient.LastRequest = ""
            HelperClient.LastResponse = ""
        End If
        ApplyMenuLanguage()
        If form IsNot Nothing AndAlso Not form.IsDisposed Then form.ApplySettings()
        Return True
    End Function

    Friend Sub CommitCredentials(registered As ClientConfig)
        ' One-time registration must survive closing settings, independent of unsaved hotkey edits.
        Dim candidate As ClientConfig = ConfigStore.Clone(config)
        candidate.HelperUrl = registered.HelperUrl
        candidate.CredentialOrigin = registered.CredentialOrigin
        candidate.UserId = registered.UserId
        candidate.LicenseUser = registered.LicenseUser
        candidate.ApiKeyProtected = registered.ApiKeyProtected
        candidate.DeviceTokenProtected = registered.DeviceTokenProtected
        ConfigStore.Save(candidate)
        config = candidate
    End Sub

    Friend Function Settings() As ClientConfig
        Return config
    End Function

    Private Sub HotkeyPressed(id As Integer)
        If Not hotkeys.IsActive(id) Then Return
        If settingsForm IsNot Nothing AndAlso Not settingsForm.IsDisposed AndAlso settingsForm.IsRecordingHotkey Then
            settingsForm.RecordCurrentHotkey(config.HotkeyModifiers, config.HotkeyKey)
        Else
            ShowInky(True)
        End If
    End Sub

    Protected Overrides Sub ExitThreadCore()
        If exiting Then Return
        exiting = True
        If settingsForm IsNot Nothing AndAlso Not settingsForm.IsDisposed Then settingsForm.Dispose()
        If form IsNot Nothing AndAlso Not form.IsDisposed Then form.Shutdown()
        hotkeys.Dispose()
        host.Dispose()
        tray.Visible = False
        tray.Dispose()
        menu.Dispose()
        MyBase.ExitThreadCore()
    End Sub

    Private NotInheritable Class HotkeyWindow
        Inherits System.Windows.Forms.NativeWindow
        Implements System.IDisposable
        Private ReadOnly owner As AssistantContext
        Public Sub New(parent As AssistantContext)
            owner = parent
            CreateHandle(New System.Windows.Forms.CreateParams With {.Parent = New System.IntPtr(-3)})
        End Sub
        Protected Overrides Sub WndProc(ByRef message As System.Windows.Forms.Message)
            If message.Msg = Native.WM_HOTKEY Then owner.HotkeyPressed(message.WParam.ToInt32())
            MyBase.WndProc(message)
        End Sub
        Public Sub Dispose() Implements System.IDisposable.Dispose
            DestroyHandle()
        End Sub
    End Class
End Class
