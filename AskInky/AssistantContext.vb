' Part of "Red Ink" (Ask Inky)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: AssistantContext.vb
' Purpose: Manages tray lifecycle, VDI hotkey recovery and the assistant windows.
' Architecture: Windows native host with embedded Office Web-based interface.
' Notes: Red Ink License and Ask Inky third-party attributions are embedded in LICENSE.txt.
' =============================================================================
' Part of Red Ink (AskInky). Tray lifecycle and one configurable global keyboard shortcut only.
Public NotInheritable Class AssistantContext
    Inherits System.Windows.Forms.ApplicationContext
    Private ReadOnly host As HotkeyWindow
    Private ReadOnly hotkeys As HotkeyService
    Private ReadOnly ocrHotkeys As HotkeyService
    Private ReadOnly tray As System.Windows.Forms.NotifyIcon
    Private ReadOnly menu As System.Windows.Forms.ContextMenuStrip
    Private ReadOnly hotkeyRecoveryTimer As System.Windows.Forms.Timer
    Private reconnectPending As Integer
    Private sessionEventsSubscribed As Boolean
    Private lastHotkeyWarning As System.DateTime = System.DateTime.MinValue
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
        ocrHotkeys = New HotkeyService(host.Handle, 2)
        menu = New System.Windows.Forms.ContextMenuStrip With {.ShowImageMargin = False}
        menu.Items.Add("", Nothing, Sub(sender As Object, e As System.EventArgs) ShowInky(False))
        menu.Items.Add("", Nothing, Sub(sender As Object, e As System.EventArgs) StartOcrCapture())
        menu.Items.Add("", Nothing, Sub(sender As Object, e As System.EventArgs) ShowSettings())
        menu.Items.Add("", Nothing, Sub(sender As Object, e As System.EventArgs) ExitThread())
        tray = New System.Windows.Forms.NotifyIcon With {.Text = "Red Ink", .ContextMenuStrip = menu, .Icon = System.Drawing.Icon.ExtractAssociatedIcon(System.Windows.Forms.Application.ExecutablePath), .Visible = True}
        AddHandler tray.MouseClick, Sub(sender As Object, e As System.Windows.Forms.MouseEventArgs)
                                        If e.Button = System.Windows.Forms.MouseButtons.Left Then menu.Show(System.Windows.Forms.Cursor.Position)
                                    End Sub
        ApplyMenuLanguage()
        hotkeyRecoveryTimer = New System.Windows.Forms.Timer With {.Interval = 15000}
        AddHandler hotkeyRecoveryTimer.Tick, AddressOf RecoverHotkeyOnUiThread
        Try
            AddHandler Microsoft.Win32.SystemEvents.SessionSwitch, AddressOf SessionSwitched
            sessionEventsSubscribed = True
        Catch ex As System.Exception
            ClientLog.Write("hotkey.session-events.unavailable", ex)
        End Try
        hotkeyRecoveryTimer.Start()
        If Not hotkeys.TryApply(config, False) Then WarnHotkeyUnavailable()
        If Not ocrHotkeys.TryApply(OcrShortcutConfig(config), False) Then WarnHotkeyUnavailable()
        If ConfigStore.LoadFailed Then tray.ShowBalloonTip(6000, "Red Ink", UiStrings.GetText("desktop.configLoadFailed", config.UiLanguage), System.Windows.Forms.ToolTipIcon.Warning)
    End Sub

    Private Sub WarnHotkeyUnavailable()
        ' Do not repeatedly interrupt VDI users with balloons on every retry.
        If System.DateTime.UtcNow - lastHotkeyWarning < System.TimeSpan.FromMinutes(10) Then Return
        lastHotkeyWarning = System.DateTime.UtcNow
        ClientLog.Write("hotkey.unavailable.tray-menu-remains-available")
        tray.ShowBalloonTip(6000, "Red Ink", UiStrings.GetText("desktop.hotkeyConflict", config.UiLanguage), System.Windows.Forms.ToolTipIcon.Warning)
    End Sub

    Private Sub SessionSwitched(sender As Object, e As Microsoft.Win32.SessionSwitchEventArgs)
        If exiting Then Return
        Select Case e.Reason
            Case Microsoft.Win32.SessionSwitchReason.SessionUnlock,
                 Microsoft.Win32.SessionSwitchReason.RemoteConnect,
                 Microsoft.Win32.SessionSwitchReason.ConsoleConnect
                System.Threading.Interlocked.Exchange(reconnectPending, 1)
                ClientLog.Write("hotkey.session.reconnect-observed")
                If Not Native.PostMessage(host.Handle, Native.WM_ASKINKY_REBIND, System.IntPtr.Zero, System.IntPtr.Zero) Then
                    ClientLog.Write("hotkey.session.rebind-post-failed")
                End If
        End Select
    End Sub

    Private Sub RecoverHotkeyOnUiThread(sender As Object, e As System.EventArgs)
        If exiting OrElse capturing Then Return
        Try
            Dim reconnect As Boolean = System.Threading.Interlocked.Exchange(reconnectPending, 0) <> 0
            If reconnect Then
                If Not hotkeys.Rebind(config) Then WarnHotkeyUnavailable()
                If Not ocrHotkeys.Rebind(OcrShortcutConfig(config)) Then WarnHotkeyUnavailable()
            Else
                If Not hotkeys.IsRegistered AndAlso Not hotkeys.TryApply(config, False) Then WarnHotkeyUnavailable()
                If Not ocrHotkeys.IsRegistered AndAlso Not ocrHotkeys.TryApply(OcrShortcutConfig(config), False) Then WarnHotkeyUnavailable()
            End If
        Catch ex As System.Exception
            ClientLog.Write("hotkey.recovery.failed", ex)
        End Try
    End Sub

    Private Shared Function OcrShortcutConfig(source As ClientConfig) As ClientConfig
        Dim value As ClientConfig = ConfigStore.Clone(source)
        value.HotkeyModifiers = source.OcrHotkeyModifiers
        value.HotkeyKey = source.OcrHotkeyKey
        Return value
    End Function

    Private Sub ApplyMenuLanguage()
        menu.Items(0).Text = UiStrings.GetText("desktop.open", config.UiLanguage)
        DirectCast(menu.Items(0), System.Windows.Forms.ToolStripMenuItem).ShortcutKeyDisplayString = HotkeyService.Format(config.HotkeyModifiers, config.HotkeyKey)
        menu.Items(1).Text = UiStrings.GetText("desktop.ocrMenu", config.UiLanguage)
        DirectCast(menu.Items(1), System.Windows.Forms.ToolStripMenuItem).ShortcutKeyDisplayString = HotkeyService.Format(config.OcrHotkeyModifiers, config.OcrHotkeyKey)
        menu.Items(2).Text = UiStrings.GetText("settings", config.UiLanguage)
        menu.Items(3).Text = UiStrings.GetText("desktop.exit", config.UiLanguage)
    End Sub

    Friend Sub SetBusy(value As Boolean)
        busy = value
        If Not exiting Then menu.Items(2).Enabled = Not value
    End Sub

    Friend Async Sub ShowInky(captureSelection As Boolean)
        If exiting OrElse capturing Then Return
        If settingsForm IsNot Nothing AndAlso Not settingsForm.IsDisposed Then
            settingsForm.BringToFront()
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
            WindowPlacement.EnsureVisible(form)
            form.Show()
            ' VDI remote desktops may reject foreground activation after asynchronous
            ' clipboard capture. SetForegroundWindow first; use a short topmost pulse
            ' only if the form is still behind another application's window.
            form.BringToFront()
            Native.SetForegroundWindow(form.Handle)
            form.Activate()
            If Native.GetForegroundWindow() <> form.Handle Then
                Dim wasTopMost As Boolean = form.TopMost
                form.TopMost = True
                form.BringToFront()
                form.Activate()
                form.TopMost = wasTopMost
                ClientLog.Write("assistant.foreground.topmost-fallback")
            End If
            form.FocusQuestion()
        Catch ex As System.Exception
            ClientLog.Write("assistant.open.failed", ex)
            tray.ShowBalloonTip(5000, "Red Ink", HelperClient.Describe(ex, config.UiLanguage), System.Windows.Forms.ToolTipIcon.Warning)
        Finally
            capturing = False
        End Try
    End Sub

    Friend Async Sub StartOcrCapture()
        If exiting OrElse capturing OrElse busy Then Return
        If settingsForm IsNot Nothing AndAlso Not settingsForm.IsDisposed Then Return
        capturing = True
        Try
            Dim png As Byte() = Nothing
            Dim restoreForm As Boolean = form IsNot Nothing AndAlso Not form.IsDisposed AndAlso form.Visible
            If restoreForm Then form.Hide()
            ' Yield to the window compositor so Ask Inky is not captured in the image.
            Await System.Threading.Tasks.Task.Delay(180)
            Using picker As New ScreenRegionPicker()
                If picker.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
                    If restoreForm Then form.Show()
                    Return
                End If
                png = picker.SelectedPng()
            End Using
            If restoreForm Then form.Show()
            If png Is Nothing Then Return
            ClientLog.Write("ocr.capture.selected.bytes=" & png.Length.ToString(System.Globalization.CultureInfo.InvariantCulture))
            Dim text As String
            Using progress As New OcrProgressForm()
                progress.Show()
                progress.Refresh()
                Try
                    text = Await HelperClient.OcrImageAsync(config, png, System.Threading.CancellationToken.None)
                Finally
                    progress.Close()
                End Try
            End Using
            If exiting Then Return
            If form Is Nothing OrElse form.IsDisposed Then form = New AssistantForm(Me)
            form.LoadSelection(text)
            form.Show()
            form.BringToFront()
            Native.SetForegroundWindow(form.Handle)
            form.Activate()
            form.FocusQuestion()
            ClientLog.Write("ocr.capture.completed")
        Catch ex As System.Exception
            ClientLog.Write("ocr.capture.failed", ex)
            tray.ShowBalloonTip(6000, "Red Ink", HelperClient.Describe(ex, config.UiLanguage), System.Windows.Forms.ToolTipIcon.Warning)
            If form IsNot Nothing AndAlso Not form.IsDisposed Then form.Show()
        Finally
            capturing = False
        End Try
    End Sub

    Friend Sub RememberAssistantPosition(position As System.Drawing.Point)
        If exiting Then Return
        If form Is Nothing OrElse form.IsDisposed OrElse form.WindowState <> System.Windows.Forms.FormWindowState.Normal Then Return
        ' Only persist a location if the form is actually visible on an attached monitor.
        If Not WindowPlacement.IsMeaningfullyVisible(form.Bounds) Then Return
        If config.AssistantLeft.HasValue AndAlso config.AssistantTop.HasValue AndAlso
           config.AssistantLeft.Value = position.X AndAlso config.AssistantTop.Value = position.Y Then Return
        config.AssistantLeft = position.X
        config.AssistantTop = position.Y
        Try
            ConfigStore.Save(config)
        Catch ex As System.Exception
            ClientLog.Write("assistant.position.save.failed", ex)
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
        ' Modeless Show(owner) does not guarantee CenterParent positioning.
        ' Position explicitly against the current assistant rectangle, then constrain
        ' the entire Settings window to its monitor working area.
        Dim visibleOwner As Boolean = form IsNot Nothing AndAlso Not form.IsDisposed AndAlso form.Visible
        If visibleOwner Then
            form.SetSettingsOpen(True)
            WindowPlacement.CenterOverOwner(settingsForm, form)
            settingsForm.Show(form)
        Else
            WindowPlacement.CenterOnActiveScreen(settingsForm)
            settingsForm.Show()
        End If

        ' Windows/VDI may deny foreground activation on the first request.
        ' A transient topmost transition helps bring the settings forward;
        ' the dialog must never stay permanently topmost.
        settingsForm.BringToFront()
        settingsForm.Activate()
        If Not settingsForm.ContainsFocus Then
            settingsForm.TopMost = True
            settingsForm.BringToFront()
            settingsForm.Activate()
            settingsForm.TopMost = False
        End If
    End Sub

    Friend Function TryCommit(candidate As ClientConfig) As Boolean
        If candidate.HotkeyModifiers = candidate.OcrHotkeyModifiers AndAlso candidate.HotkeyKey = candidate.OcrHotkeyKey Then Return False
        Dim previous As ClientConfig = ConfigStore.Clone(config)
        If Not ocrHotkeys.TryApply(OcrShortcutConfig(candidate), False) Then Return False
        If Not hotkeys.TryApply(candidate, False) Then
            ocrHotkeys.TryApply(OcrShortcutConfig(previous), False)
            Return False
        End If
        Try
            ConfigStore.Save(candidate)
        Catch
            hotkeys.TryApply(previous, False)
            ocrHotkeys.TryApply(OcrShortcutConfig(previous), False)
            Throw
        End Try
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
        If Not hotkeys.IsActive(id) AndAlso Not ocrHotkeys.IsActive(id) Then Return
        ClientLog.Write(If(ocrHotkeys.IsActive(id), "ocr.hotkey.message.received", "hotkey.message.received"))
        If settingsForm IsNot Nothing AndAlso Not settingsForm.IsDisposed AndAlso settingsForm.IsRecordingHotkey Then
            If settingsForm.RecordingOcrHotkey Then
                settingsForm.RecordCurrentHotkey(config.OcrHotkeyModifiers, config.OcrHotkeyKey, True)
            Else
                settingsForm.RecordCurrentHotkey(config.HotkeyModifiers, config.HotkeyKey, False)
            End If
        ElseIf ocrHotkeys.IsActive(id) Then
            StartOcrCapture()
        Else
            ShowInky(True)
        End If
    End Sub

    Protected Overrides Sub ExitThreadCore()
        If exiting Then Return
        exiting = True
        If sessionEventsSubscribed Then
            Try
                RemoveHandler Microsoft.Win32.SystemEvents.SessionSwitch, AddressOf SessionSwitched
            Catch ex As System.Exception
                ClientLog.Write("hotkey.session-events.cleanup.failed", ex)
            End Try
        End If
        hotkeyRecoveryTimer.Stop()
        hotkeyRecoveryTimer.Dispose()
        If settingsForm IsNot Nothing AndAlso Not settingsForm.IsDisposed Then settingsForm.Dispose()
        If form IsNot Nothing AndAlso Not form.IsDisposed Then form.Shutdown()
        hotkeys.Dispose()
        ocrHotkeys.Dispose()
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
            If message.Msg = Native.WM_HOTKEY Then
                owner.HotkeyPressed(message.WParam.ToInt32())
            ElseIf message.Msg = Native.WM_ASKINKY_REBIND Then
                owner.RecoverHotkeyOnUiThread(Me, System.EventArgs.Empty)
            End If
            MyBase.WndProc(message)
        End Sub
        Public Sub Dispose() Implements System.IDisposable.Dispose
            DestroyHandle()
        End Sub
    End Class
End Class
