' Part of "Red Ink" (Ask Inky)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: ConfigStore.vb
' Purpose: Ask Inky desktop functionality in the modular Windows client.
' Architecture: Windows native host with embedded Office Web-based interface.
' Notes: Red Ink License and Ask Inky third-party attributions are embedded in LICENSE.txt.
' =============================================================================
' Part of Red Ink (AskInky). User settings and DPAPI-protected Helper credentials.
Public NotInheritable Class ClientConfig
    Public Property HelperUrl As String = "http://localhost"
    Public Property UserId As String = ""
    Public Property LicenseUser As String = ""
    Public Property HotkeyModifiers As UInteger = 3UI
    Public Property HotkeyKey As Integer = CInt(System.Windows.Forms.Keys.K)
    Public Property OcrHotkeyModifiers As UInteger = 3UI
    Public Property OcrHotkeyKey As Integer = CInt(System.Windows.Forms.Keys.O)
    ' Optional persisted assistant position; Nothing means centered on the active screen.
    Public Property AssistantLeft As Integer? = Nothing
    Public Property AssistantTop As Integer? = Nothing
    Public Property DarkMode As Boolean = False
    Public Property UiLanguage As String = "de"
    Public Property FontSize As Integer = 12
    Public Property Temperature As Double = 0.7
    Public Property Model As String = ""
    Public Property ShowLlmDebug As Boolean = False
    Public Property SystemPrompt As String = "Antworte direkt, sachlich und knapp auf die Frage. Keine Begrüssung, Selbstvorstellung, Höflichkeitsfloskeln oder einleitende Ankündigung. Nur auf ausdrücklichen Wunsch ausführlich antworten. Verwende Markdown, wenn es der Lesbarkeit dient."
    Public Property ApiKeyProtected As String = ""
    Public Property DeviceTokenProtected As String = ""
    Public Property CredentialOrigin As String = ""
    ' Registry policy is opt-out; remember only a SHA-256 of the successful one-time link.
    Public Property IgnoreRegistryHelperLink As Boolean = False
    Public Property AppliedRegistryLinkFingerprint As String = ""
    ' Obsolete UseRightClick/ConfirmSend keys in older JSON are deliberately ignored.
End Class

Public NotInheritable Class ConfigStore
    Private Shared ReadOnly ConfigPath As String = System.IO.Path.Combine(StoragePaths.Root, "config.json")
    ' Keep the original entropy: previously registered devices must continue to decrypt.
    Private Shared ReadOnly Entropy As Byte() = System.Text.Encoding.UTF8.GetBytes("RedInk.AskInky.DeviceAuth.v1")
    Public Shared Property LoadFailed As Boolean
    Public Shared Property HasSavedHelperHost As Boolean

    Public Shared Function Load() As ClientConfig
        LoadFailed = False
        HasSavedHelperHost = False
        Try
            If System.IO.File.Exists(ConfigPath) Then
                Dim json As String = System.IO.File.ReadAllText(ConfigPath, System.Text.Encoding.UTF8)
                Dim result As ClientConfig = System.Text.Json.JsonSerializer.Deserialize(Of ClientConfig)(json)
                If result Is Nothing Then Throw New System.IO.InvalidDataException("Empty configuration")
                ' Property initializers supply localhost even when the JSON has no HelperUrl.
                ' Only an explicitly persisted nonempty host counts as user configuration.
                Using document As System.Text.Json.JsonDocument = System.Text.Json.JsonDocument.Parse(json)
                    Dim savedUrl As System.Text.Json.JsonElement
                    HasSavedHelperHost = document.RootElement.ValueKind = System.Text.Json.JsonValueKind.Object AndAlso
                        document.RootElement.TryGetProperty("HelperUrl", savedUrl) AndAlso
                        savedUrl.ValueKind = System.Text.Json.JsonValueKind.String AndAlso
                        Not String.IsNullOrWhiteSpace(savedUrl.GetString())
                End Using
                result.UiLanguage = UiStrings.NormalizeLanguage(result.UiLanguage)
                If result.FontSize < 6 OrElse result.FontSize > 16 Then result.FontSize = 12
                If System.Double.IsNaN(result.Temperature) OrElse result.Temperature < 0 OrElse result.Temperature > 1 Then result.Temperature = 0.7
                ' Upgrade only the exact historical pair of unmodified defaults. A custom
                ' primary or OCR shortcut must never be rewritten on upgrade.
                ' Legacy: Ctrl+Shift+K (6/75) and Ctrl+Alt+Shift+K (7/75).
                If result.HotkeyModifiers = 6UI AndAlso result.HotkeyKey = 75 AndAlso
                   result.OcrHotkeyModifiers = 7UI AndAlso result.OcrHotkeyKey = 75 Then
                    result.HotkeyModifiers = 3UI
                    result.HotkeyKey = 75
                    result.OcrHotkeyModifiers = 3UI
                    result.OcrHotkeyKey = 79
                    Try
                        Save(result)
                        ClientLog.Write("hotkey.defaults.legacy-pair-upgraded")
                    Catch migrationError As System.Exception
                        ClientLog.Write("hotkey.defaults.migration-save-failed", migrationError)
                    End Try
                End If
                ' Bind legacy credentials to the server already saved with them, never to a new target.
                If String.IsNullOrEmpty(result.CredentialOrigin) Then result.CredentialOrigin = HelperClient.GetOrigin(result.HelperUrl)
                Return result
            End If
        Catch ex As System.Exception
            LoadFailed = True
            HasSavedHelperHost = False
            ClientLog.Write("config.load.failed.original-file-retained", ex)
        End Try
        Return New ClientConfig()
    End Function

    Public Shared Function Clone(config As ClientConfig) As ClientConfig
        Return System.Text.Json.JsonSerializer.Deserialize(Of ClientConfig)(System.Text.Json.JsonSerializer.Serialize(config))
    End Function

    Public Shared Sub Save(config As ClientConfig)
        System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(ConfigPath))
        Dim temporary As String = ConfigPath & "." & System.Guid.NewGuid().ToString("N") & ".tmp"
        Try
            System.IO.File.WriteAllText(temporary, System.Text.Json.JsonSerializer.Serialize(config, New System.Text.Json.JsonSerializerOptions With {.WriteIndented = True}), New System.Text.UTF8Encoding(False))
            System.IO.File.Move(temporary, ConfigPath, True)
        Finally
            Try
                If System.IO.File.Exists(temporary) Then System.IO.File.Delete(temporary)
            Catch cleanupError As System.Exception
                ClientLog.Write("config.temporary.cleanup.failed", cleanupError)
            End Try
        End Try
    End Sub

    Public Shared Function Protect(clearValue As String) As String
        If String.IsNullOrEmpty(clearValue) Then Return ""
        Return System.Convert.ToBase64String(System.Security.Cryptography.ProtectedData.Protect(System.Text.Encoding.UTF8.GetBytes(clearValue), Entropy, System.Security.Cryptography.DataProtectionScope.CurrentUser))
    End Function

    Public Shared Function Reveal(encryptedValue As String) As String
        If String.IsNullOrEmpty(encryptedValue) Then Return ""
        Return System.Text.Encoding.UTF8.GetString(System.Security.Cryptography.ProtectedData.Unprotect(System.Convert.FromBase64String(encryptedValue), Entropy, System.Security.Cryptography.DataProtectionScope.CurrentUser))
    End Function
End Class
