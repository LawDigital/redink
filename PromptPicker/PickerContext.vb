' Standalone tray lifetime and hotkey dispatch; no VSTO/Office COM dependencies.
Friend NotInheritable Class PickerContext
    Inherits System.Windows.Forms.ApplicationContext

    Private ReadOnly host As PickerMessageWindow
    Private ReadOnly keys As ShortcutService
    Private ReadOnly menu As System.Windows.Forms.ContextMenuStrip
    Private ReadOnly tray As System.Windows.Forms.NotifyIcon
    Private ReadOnly foregroundPoll As System.Windows.Forms.Timer
    Private ReadOnly recovery As System.Windows.Forms.Timer
    Private currentSettings As PickerSettings
    Private activePicker As PickerForm
    Private lastExternalWindow As System.IntPtr
    Private running As Boolean
    Private exiting As Boolean
    Private eventSubscribed As Boolean

    Public Sub New()
        currentSettings = SettingsStore.Load()
        host = New PickerMessageWindow(Me)
        keys = New ShortcutService(host.Handle)
        menu = New System.Windows.Forms.ContextMenuStrip With {.ShowImageMargin = False}
        menu.Items.Add("Prompt Picker öffnen", Nothing, Sub(sender As Object, e As System.EventArgs) BeginPicker(False))
        menu.Items.Add("Tastenkombination ändern …", Nothing, Sub(sender As Object, e As System.EventArgs) ShowShortcutSettings())
        menu.Items.Add(New System.Windows.Forms.ToolStripSeparator())
        menu.Items.Add("Lizenzhinweise …", Nothing, Sub(sender As Object, e As System.EventArgs) ShowLicenseInformation())
        menu.Items.Add("Beenden", Nothing, Sub(sender As Object, e As System.EventArgs) ExitThread())
        DirectCast(menu.Items(0), System.Windows.Forms.ToolStripMenuItem).ShortcutKeyDisplayString = ShortcutService.Describe(currentSettings.HotkeyModifiers, currentSettings.HotkeyKey)
        tray = New System.Windows.Forms.NotifyIcon With {
            .Text = "Red Ink Prompt Picker", .ContextMenuStrip = menu,
            .Icon = System.Drawing.Icon.ExtractAssociatedIcon(System.Windows.Forms.Application.ExecutablePath), .Visible = True
        }
        AddHandler tray.DoubleClick, Sub(sender As Object, e As System.EventArgs) BeginPicker(False)
        foregroundPoll = New System.Windows.Forms.Timer With {.Interval = 350}
        AddHandler foregroundPoll.Tick, AddressOf RememberExternalForeground
        foregroundPoll.Start()
        RememberExternalForeground(Me, System.EventArgs.Empty)
        recovery = New System.Windows.Forms.Timer With {.Interval = 15000}
        AddHandler recovery.Tick, AddressOf CheckHotkey
        recovery.Start()
        Try
            AddHandler Microsoft.Win32.SystemEvents.SessionSwitch, AddressOf OnSessionSwitch
            eventSubscribed = True
        Catch ex As System.Exception
            AppLog.Write("session.monitor.unavailable", ex)
        End Try
        If Not keys.Apply(currentSettings) Then WarnHotkeyUnavailable()
        AppLog.Write("tray.started " & ShortcutService.Describe(currentSettings.HotkeyModifiers, currentSettings.HotkeyKey))
    End Sub

    Private Sub WarnHotkeyUnavailable()
        tray.ShowBalloonTip(5500, "Red Ink Prompt Picker", "Der Hotkey ist bereits belegt. Du kannst den Picker über das Tray öffnen oder eine andere Tastenkombination wählen.", System.Windows.Forms.ToolTipIcon.Warning)
    End Sub

    Private Sub RememberExternalForeground(sender As Object, e As System.EventArgs)
        If running Then Return
        Dim window As System.IntPtr = Native.GetForegroundWindow()
        If IsExternal(window) Then lastExternalWindow = window
    End Sub

    Private Shared Function IsExternal(window As System.IntPtr) As Boolean
        If window = System.IntPtr.Zero OrElse Not Native.IsWindow(window) Then Return False
        Dim processId As UInteger
        Native.GetWindowThreadProcessId(window, processId)
        Return processId <> 0UI AndAlso processId <> CUInt(System.Environment.ProcessId)
    End Function

    Friend Sub OnHotkey(id As Integer)
        If keys.IsMessage(id) Then BeginPicker(True)
    End Sub

    Friend Sub OnRebind()
        If Not exiting AndAlso Not keys.Rebind(currentSettings) Then WarnHotkeyUnavailable()
    End Sub

    Private Sub OnSessionSwitch(sender As Object, args As Microsoft.Win32.SessionSwitchEventArgs)
        If exiting Then Return
        If args.Reason = Microsoft.Win32.SessionSwitchReason.SessionUnlock OrElse
           args.Reason = Microsoft.Win32.SessionSwitchReason.RemoteConnect OrElse
           args.Reason = Microsoft.Win32.SessionSwitchReason.ConsoleConnect Then
            Native.PostMessage(host.Handle, Native.WM_REBIND, System.IntPtr.Zero, System.IntPtr.Zero)
        End If
    End Sub

    Private Sub CheckHotkey(sender As Object, e As System.EventArgs)
        If Not exiting AndAlso Not running AndAlso Not keys.IsRegistered AndAlso Not keys.Apply(currentSettings) Then
            AppLog.Write("hotkey.recovery.retry.failed")
        End If
    End Sub

    Private Async Sub BeginPicker(calledByHotkey As Boolean)
        If exiting Then Return
        If running Then
            If activePicker IsNot Nothing AndAlso Not activePicker.IsDisposed Then
                activePicker.Activate()
                activePicker.BringToFront()
            End If
            Return
        End If
        Dim target As System.IntPtr = If(calledByHotkey, Native.GetForegroundWindow(), lastExternalWindow)
        If Not IsExternal(target) Then target = lastExternalWindow
        If Not IsExternal(target) Then
            System.Windows.Forms.MessageBox.Show("Keine zuvor aktive Zielanwendung erkannt. Bitte zuerst die gewünschte Anwendung aktivieren.", "Red Ink Prompt Picker", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
            Return
        End If
        ' Preserve the actual editing control before our picker steals focus.
        ' This is target-agnostic: Win32 may host the editor in a child HWND.
        Dim originalFocus As System.IntPtr = If(Native.GetForegroundWindow() = target, SelectionClipboardService.CaptureFocusedControl(target), System.IntPtr.Zero)
        running = True
        Try
            Dim selection As String = ""
            ' Selection capture is best effort and only when the target still has foreground.
            ' For a tray-menu invocation Windows can focus Explorer; browsing still works.
            If Native.GetForegroundWindow() = target Then
                selection = Await SelectionClipboardService.ReadSelectionAsync(target)
            End If
            If exiting Then Return
            Dim config As LibraryConfiguration = Await WordConfigReader.LoadAsync()
            If exiting Then Return
            Dim prompts As System.Collections.Generic.List(Of PromptEntry) = PromptLibraryStore.Load(config)
            AppLog.Write("picker.open; local-configured=" & config.HasLocal.ToString() & "; entries=" & prompts.Count.ToString(System.Globalization.CultureInfo.InvariantCulture))
            Using picker As New PickerForm(config, selection, prompts, currentSettings)
                activePicker = picker
                Dim dialogOutcome As System.Windows.Forms.DialogResult = picker.ShowDialog()
                picker.CaptureWindowSettings()
                Try
                    SettingsStore.Save(currentSettings)
                Catch ex As System.Exception
                    AppLog.Write("picker.geometry.save.failed", ex)
                End Try
                If dialogOutcome = System.Windows.Forms.DialogResult.OK Then
                    ' Do not keep the modal WebView2 alive while restoring input
                    ' to the external target. The selected text is immutable here.
                    picker.Hide()
                    Dim inserted As Boolean = Await SelectionClipboardService.InsertAsync(target, originalFocus, picker.ChosenPrompt)
                    If Not inserted Then
                        System.Windows.Forms.MessageBox.Show("Die Zielanwendung konnte nicht sicher aktiviert werden oder das Einfügen war nicht möglich. Der Prompt wurde nicht absichtlich in ein anderes Fenster eingefügt.", "Red Ink Prompt Picker", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning)
                    End If
                End If
            End Using
        Catch ex As System.Exception
            AppLog.Write("picker.open.failed", ex)
            System.Windows.Forms.MessageBox.Show("Prompt Picker konnte nicht geöffnet werden: " & ex.Message & System.Environment.NewLine & System.Environment.NewLine & "Konfigurationsquelle: " & WordConfigReader.ResolveSource(), "Red Ink Prompt Picker", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
        Finally
            activePicker = Nothing
            running = False
        End Try
    End Sub

    Private Sub ShowLicenseInformation()
        If exiting Then Return
        Try
            Using license As New LicenseInfoForm()
                license.ShowDialog()
            End Using
        Catch ex As System.Exception
            AppLog.Write("license.view.failed", ex)
            System.Windows.Forms.MessageBox.Show("Die Lizenzhinweise konnten nicht angezeigt werden.", "Red Ink Prompt Picker", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub ShowShortcutSettings()
        If running OrElse exiting Then Return
        Using settings As New ShortcutSettingsForm(currentSettings)
            If settings.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then Return
            Dim desired As PickerSettings = settings.Shortcut
            ' Hotkey edits must not erase the persisted picker geometry.
            desired.PickerX = currentSettings.PickerX
            desired.PickerY = currentSettings.PickerY
            desired.PickerClientWidth = currentSettings.PickerClientWidth
            desired.PickerClientHeight = currentSettings.PickerClientHeight
            desired.PreviewPercent = currentSettings.PreviewPercent
            If Not keys.Apply(desired) Then
                System.Windows.Forms.MessageBox.Show("Diese Tastenkombination ist belegt oder ungültig. Der bisherige Hotkey bleibt erhalten.", "Red Ink Prompt Picker", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning)
                Return
            End If
            Try
                SettingsStore.Save(desired)
                currentSettings = desired
                DirectCast(menu.Items(0), System.Windows.Forms.ToolStripMenuItem).ShortcutKeyDisplayString = ShortcutService.Describe(currentSettings.HotkeyModifiers, currentSettings.HotkeyKey)
            Catch ex As System.Exception
                AppLog.Write("hotkey.settings.save.failed", ex)
                keys.Apply(currentSettings)
                System.Windows.Forms.MessageBox.Show("Hotkey konnte nicht gespeichert werden: " & ex.Message, "Red Ink Prompt Picker", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            End Try
        End Using
    End Sub

    Protected Overrides Sub ExitThreadCore()
        exiting = True
        If eventSubscribed Then
            RemoveHandler Microsoft.Win32.SystemEvents.SessionSwitch, AddressOf OnSessionSwitch
            eventSubscribed = False
        End If
        recovery.Stop()
        recovery.Dispose()
        foregroundPoll.Stop()
        foregroundPoll.Dispose()
        If activePicker IsNot Nothing AndAlso Not activePicker.IsDisposed Then activePicker.Close()
        keys.Dispose()
        host.Dispose()
        tray.Visible = False
        tray.Dispose()
        menu.Dispose()
        AppLog.Write("tray.exit")
        MyBase.ExitThreadCore()
    End Sub
End Class

Friend NotInheritable Class PickerMessageWindow
    Inherits System.Windows.Forms.NativeWindow
    Implements System.IDisposable

    Private ReadOnly app As PickerContext

    Public Sub New(owner As PickerContext)
        app = owner
        CreateHandle(New System.Windows.Forms.CreateParams With {.Caption = "Prompt Picker Hotkey"})
    End Sub

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        If m.Msg = Native.WM_HOTKEY Then
            app.OnHotkey(m.WParam.ToInt32())
            Return
        End If
        If m.Msg = Native.WM_REBIND Then
            app.OnRebind()
            Return
        End If
        MyBase.WndProc(m)
    End Sub

    Public Sub Dispose() Implements System.IDisposable.Dispose
        DestroyHandle()
    End Sub
End Class
