Friend NotInheritable Class PickerSettings
    Public Property HotkeyModifiers As UInteger = Native.MOD_CONTROL Or Native.MOD_ALT
    Public Property HotkeyKey As Integer = CInt(System.Windows.Forms.Keys.P)
    ' UI-only settings; never change redink.ini. Nullable coordinates keep older
    ' settings.json files backward compatible and retain CenterScreen on first use.
    Public Property PickerX As Integer? = Nothing
    Public Property PickerY As Integer? = Nothing
    ' First-run width: 760 * 1.5 = 1140 client pixels. Previously customised
    ' window sizes continue to win over the default on subsequent launches.
    Public Const CurrentPickerLayoutVersion As Integer = 2
    Public Property PickerLayoutVersion As Integer = CurrentPickerLayoutVersion
    Public Property PickerClientWidth As Integer = 1140
    Public Property PickerClientHeight As Integer = 540
    Public Property PreviewPercent As Integer = 55
End Class

Friend NotInheritable Class SettingsStore
    Private Sub New()
    End Sub
    Private Shared ReadOnly Property ConfigPath As String
        Get
            Return System.IO.Path.Combine(AppLog.Root, "settings.json")
        End Get
    End Property

    Public Shared Function Load() As PickerSettings
        Try
            If Not System.IO.File.Exists(ConfigPath) Then Return New PickerSettings()
            Dim json As String = System.IO.File.ReadAllText(ConfigPath)
            Dim data As PickerSettings = System.Text.Json.JsonSerializer.Deserialize(Of PickerSettings)(json)
            If data IsNot Nothing AndAlso ShortcutService.IsValid(data.HotkeyModifiers, data.HotkeyKey) Then
                ' v3 persisted its *default* width as 760. An ordinary default
                ' change would never reach upgraded installations. Apply the
                ' one-time migration only to that legacy default, not to
                ' deliberately resized windows. An explicit version field
                ' prevents migrating a user-sized 760px window on later runs.
                Using document As System.Text.Json.JsonDocument = System.Text.Json.JsonDocument.Parse(json)
                    Dim legacyVersion As System.Text.Json.JsonElement
                    Dim isLegacy As Boolean = Not document.RootElement.TryGetProperty(NameOf(PickerSettings.PickerLayoutVersion), legacyVersion)
                    If Not isLegacy AndAlso legacyVersion.ValueKind = System.Text.Json.JsonValueKind.Number Then
                        Dim version As Integer
                        isLegacy = legacyVersion.TryGetInt32(version) AndAlso version < PickerSettings.CurrentPickerLayoutVersion
                    End If
                    If isLegacy AndAlso data.PickerClientWidth = 760 Then data.PickerClientWidth = 1140
                End Using
                data.PickerLayoutVersion = PickerSettings.CurrentPickerLayoutVersion
                Return data
            End If
        Catch ex As System.Exception
            AppLog.Write("settings.read.failed.default-used", ex)
        End Try
        Return New PickerSettings()
    End Function

    Public Shared Sub Save(value As PickerSettings)
        System.IO.Directory.CreateDirectory(AppLog.Root)
        Dim temp As String = ConfigPath & "." & System.Guid.NewGuid().ToString("N") & ".tmp"
        Try
            System.IO.File.WriteAllText(temp, System.Text.Json.JsonSerializer.Serialize(value, New System.Text.Json.JsonSerializerOptions With {.WriteIndented = True}), New System.Text.UTF8Encoding(False))
            System.IO.File.Move(temp, ConfigPath, True)
        Finally
            If System.IO.File.Exists(temp) Then System.IO.File.Delete(temp)
        End Try
    End Sub
End Class
