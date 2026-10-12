' Part of "Red Ink" (Ask Inky)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: SettingsForm.vb
' Purpose: Ask Inky desktop functionality in the modular Windows client.
' Architecture: Windows native host with embedded Office Web-based interface.
' Notes: Red Ink License and Ask Inky third-party attributions are embedded in LICENSE.txt.
' =============================================================================
' Part of Red Ink (AskInky). Office Web connection settings plus desktop shortcut configuration.
Friend NotInheritable Class SettingsForm
    Inherits WebDialog
    Private ReadOnly context As AssistantContext
    Private draft As ClientConfig
    Private previewLanguage As String
    Private working As Boolean
    Private pendingCredentialSave As Boolean
    Private showRegistrationHelp As Boolean
    Private ReadOnly registryManaged As Boolean
    Private ReadOnly initialFeedback As String
    Private Shared ReadOnly UsedInvites As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.Ordinal)
    Public Property IsRecordingHotkey As Boolean
    Public Property RecordingOcrHotkey As Boolean

    Public Sub New(parent As AssistantContext, config As ClientConfig, Optional needsRegistration As Boolean = False, Optional startupFeedback As String = "", Optional mustSaveCredentials As Boolean = False)
        MyBase.New("settings.html", "Ask Inky – " & UiStrings.GetText("settings", config.UiLanguage), New System.Drawing.Size(480, 620))
        context = parent
        draft = ConfigStore.Clone(config)
        registryManaged = RegistryRegistration.IsAppliedLinkPresent(draft.AppliedRegistryLinkFingerprint)
        previewLanguage = draft.UiLanguage
        showRegistrationHelp = needsRegistration
        initialFeedback = startupFeedback
        ' The link has already been consumed: never let a window close discard
        ' protected credentials that could not yet be saved to config.json.
        pendingCredentialSave = mustSaveCredentials
        MinimizeBox = False
    End Sub

    Protected Overrides ReadOnly Property Language As String
        Get
            Return previewLanguage
        End Get
    End Property

    Protected Overrides Sub PageReady()
        PublishSettings(initialFeedback)
    End Sub

    Friend Shared Function TryReserveInvite(link As String) As Boolean
        Dim fingerprint As String = System.Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(link)))
        Return UsedInvites.Add(fingerprint)
    End Function

    Private Sub PublishSettings(Optional feedback As String = "", Optional failure As Boolean = False)
        Dim key As String = ""
        Dim keyReadable As Boolean = True
        Try
            key = ConfigStore.Reveal(draft.ApiKeyProtected)
        Catch ex As System.Exception
            keyReadable = False
            ClientLog.Write("settings.apikey.decrypt.failed", ex)
        End Try
        previewLanguage = UiStrings.NormalizeLanguage(draft.UiLanguage)
        Text = "Ask Inky – " & UiStrings.GetText("settings", previewLanguage)
        SendView(New With {.[type] = "settings", .strings = UiStrings.GetBundle(previewLanguage), .bundles = UiStrings.GetAllBundles(),
                           .language = previewLanguage, .darkMode = draft.DarkMode, .fontSize = draft.FontSize,
                           .values = New With {.helperUrl = draft.HelperUrl, .apiKey = key, .apiKeyReadable = keyReadable,
                               .licenseUser = draft.LicenseUser, .language = previewLanguage, .darkMode = draft.DarkMode,
                               .fontSize = draft.FontSize, .temperature = draft.Temperature, .model = draft.Model, .systemPrompt = draft.SystemPrompt,
                               .showLlmDebug = draft.ShowLlmDebug, .ignoreRegistryHelperLink = draft.IgnoreRegistryHelperLink,
                               .registryManaged = registryManaged, .hotkeyModifiers = draft.HotkeyModifiers, .hotkeyKey = draft.HotkeyKey,
                               .ocrHotkeyModifiers = draft.OcrHotkeyModifiers, .ocrHotkeyKey = draft.OcrHotkeyKey,
                               .registered = Not String.IsNullOrEmpty(draft.UserId) AndAlso Not String.IsNullOrEmpty(draft.DeviceTokenProtected), .userId = draft.UserId},
                           .debugSent = HelperClient.LastRequest, .debugReceived = HelperClient.LastResponse, .feedback = feedback, .failure = failure,
                           .needsRegistration = showRegistrationHelp})
    End Sub

    Private Function ReadCandidate(message As System.Text.Json.JsonElement) As ClientConfig
        Dim values As System.Text.Json.JsonElement
        If Not message.TryGetProperty("values", values) OrElse values.ValueKind <> System.Text.Json.JsonValueKind.Object Then Throw New HelperFailure("desktop.invalidSettings")
        Dim input As SettingsValues = System.Text.Json.JsonSerializer.Deserialize(Of SettingsValues)(values.GetRawText(), New System.Text.Json.JsonSerializerOptions With {.PropertyNameCaseInsensitive = True})
        If input Is Nothing OrElse Not UiStrings.IsSupported(input.Language) OrElse Not HotkeyService.IsValid(input.HotkeyModifiers, input.HotkeyKey) OrElse Not HotkeyService.IsValid(input.OcrHotkeyModifiers, input.OcrHotkeyKey) Then Throw New HelperFailure("desktop.invalidSettings")
        If input.FontSize < 6 OrElse input.FontSize > 16 OrElse System.Double.IsNaN(input.Temperature) OrElse input.Temperature < 0 OrElse input.Temperature > 1 Then Throw New HelperFailure("desktop.invalidSettings")
        If input.ApiKey Is Nothing OrElse input.ApiKey.Length > 16384 OrElse input.ApiKey.Contains(vbCr) OrElse input.ApiKey.Contains(vbLf) Then Throw New HelperFailure("desktop.invalidSettings")
        If input.LicenseUser Is Nothing OrElse input.LicenseUser.Length > 1024 OrElse input.LicenseUser.Contains(vbCr) OrElse input.LicenseUser.Contains(vbLf) Then Throw New HelperFailure("desktop.invalidSettings")
        If input.Model Is Nothing OrElse input.Model.Length > 256 OrElse input.Model.Contains(vbCr) OrElse input.Model.Contains(vbLf) Then Throw New HelperFailure("desktop.invalidSettings")
        If input.SystemPrompt Is Nothing OrElse input.SystemPrompt.Length > 200000 Then Throw New HelperFailure("desktop.invalidSettings")
        Dim result As ClientConfig = ConfigStore.Clone(draft)
        result.IgnoreRegistryHelperLink = input.IgnoreRegistryHelperLink
        If registryManaged AndAlso Not input.IgnoreRegistryHelperLink Then
            ' Registry credentials are authoritative while the matching HelperLink is present.
            ' Do not allow a stale or manipulated WebView draft to overwrite them.
            result.HelperUrl = draft.HelperUrl
            result.CredentialOrigin = draft.CredentialOrigin
            result.UserId = draft.UserId
            result.DeviceTokenProtected = draft.DeviceTokenProtected
            result.ApiKeyProtected = draft.ApiKeyProtected
            result.LicenseUser = draft.LicenseUser
        Else
            result.HelperUrl = HelperClient.NormalizeHelperUrl(input.HelperUrl)
            Dim origin As String = HelperClient.GetOrigin(result.HelperUrl)
            If Not String.Equals(origin, draft.CredentialOrigin, System.StringComparison.OrdinalIgnoreCase) Then
                result.UserId = ""
                result.DeviceTokenProtected = ""
                result.ApiKeyProtected = ""
                ClientLog.Write("settings.server.changed.old-credentials-not-forwarded")
            End If
            If input.ApiKeyChanged Then result.ApiKeyProtected = ConfigStore.Protect(input.ApiKey.Trim())
            result.CredentialOrigin = origin
            result.LicenseUser = input.LicenseUser.Trim()
            ' If the user deliberately replaces managed credentials while opted out,
            ' the old one-time registry link can no longer be claimed as applied.
            If input.IgnoreRegistryHelperLink AndAlso registryManaged AndAlso
               (Not System.String.Equals(result.HelperUrl, draft.HelperUrl, System.StringComparison.OrdinalIgnoreCase) OrElse
                input.ApiKeyChanged OrElse
                Not System.String.Equals(result.LicenseUser, draft.LicenseUser, System.StringComparison.Ordinal)) Then
                result.AppliedRegistryLinkFingerprint = ""
            End If
        End If
        result.UiLanguage = UiStrings.NormalizeLanguage(input.Language)
        result.DarkMode = input.DarkMode
        result.FontSize = input.FontSize
        result.Temperature = input.Temperature
        result.Model = input.Model.Trim()
        result.SystemPrompt = input.SystemPrompt
        result.ShowLlmDebug = input.ShowLlmDebug
        result.HotkeyModifiers = input.HotkeyModifiers
        result.HotkeyKey = input.HotkeyKey
        result.OcrHotkeyModifiers = input.OcrHotkeyModifiers
        result.OcrHotkeyKey = input.OcrHotkeyKey
        If result.HotkeyModifiers = result.OcrHotkeyModifiers AndAlso result.HotkeyKey = result.OcrHotkeyKey Then Throw New HelperFailure("desktop.hotkeyConflict")
        Return result
    End Function

    Protected Overrides Async Function HandleMessageAsync(message As System.Text.Json.JsonElement) As System.Threading.Tasks.Task
        Dim action As String = HelperClient.TextValue(message, "action")
        If action = "preview" Then
            Dim value As String = HelperClient.TextValue(message, "language")
            If UiStrings.IsSupported(value) Then
                previewLanguage = UiStrings.NormalizeLanguage(value)
                Text = "Ask Inky – " & UiStrings.GetText("settings", previewLanguage)
            End If
            Return
        End If
        If action = "loadModels" Then
            Try
                Dim candidate As ClientConfig = ReadCandidate(message)
                Dim choices As System.Collections.Generic.List(Of HelperModelOption) = Await HelperClient.GetModelsAsync(candidate, Lifetime.Token)
                SendView(New With {.[type] = "models", .models = choices.Select(Function(m) New With {.id = m.Id, .name = m.Name}).ToArray(), .error = False})
            Catch ex As System.Exception
                ClientLog.Write("settings.models.failed", ex)
                SendView(New With {.[type] = "models", .models = New System.Collections.Generic.List(Of HelperModelOption)(), .error = True})
            End Try
            Return
        End If
        If action = "openLicenses" Then
            Using licenses As New LicenseForm(previewLanguage)
                licenses.ShowDialog(Me)
            End Using
            Return
        End If
        If action = "openWebsite" Then
            Try
                System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo("https://redink.ai") With {.UseShellExecute = True})
            Catch ex As System.Exception
                ClientLog.Write("settings.website.failed", ex)
            End Try
            Return
        End If
        If action = "recording" Then
            Dim active As System.Text.Json.JsonElement
            IsRecordingHotkey = Not working AndAlso message.TryGetProperty("active", active) AndAlso active.ValueKind = System.Text.Json.JsonValueKind.True
            RecordingOcrHotkey = IsRecordingHotkey AndAlso HelperClient.TextValue(message, "target") = "ocr"
            Return
        End If
        If working Then Return
        If action = "cancel" Then
            Close()
            Return
        End If
        If action <> "save" AndAlso action <> "test" AndAlso action <> "register" AndAlso action <> "unregister" Then Return
        working = True
        IsRecordingHotkey = False
        SendView(New With {.[type] = "busy", .busy = True})
        Dim receivedCredentials As Boolean = False
        Try
            Dim candidate As ClientConfig = ReadCandidate(message)
            If registryManaged AndAlso Not candidate.IgnoreRegistryHelperLink AndAlso (action = "register" OrElse action = "unregister") Then
                Throw New HelperFailure("settings.registryManagedHint")
            End If
            Select Case action
                Case "save"
                    If Not context.TryCommit(candidate) Then Throw New HelperFailure("desktop.hotkeyConflict")
                    pendingCredentialSave = False
                    draft = candidate
                    working = False
                    Close()
                Case "test"
                    context.CommitCredentials(candidate)
                    draft = candidate
                    pendingCredentialSave = False
                    Dim probe As HelperProbe = Await HelperClient.ProbeAsync(candidate, Lifetime.Token)
                    PublishSettings()
                    SendView(New With {.[type] = "probe", .connected = True, .authorized = probe.Authorized, .version = probe.Version,
                                       .licenseKey = probe.LicenseKey, .licenseProduct = probe.LicenseProduct, .httpStatus = probe.HttpStatus})
                Case "register"
                    Dim link As String = HelperClient.TextValue(message, "link").Trim()
                    If link.Length = 0 OrElse link.Length > 32768 Then Throw New HelperFailure("desktop.invalidInvite")
                    Dim isInvite As Boolean = HelperClient.IsDeviceInvite(link)
                    If isInvite AndAlso Not TryReserveInvite(link) Then Throw New HelperFailure("desktop.inviteAlreadyUsed")
                    ' Configuration URLs only populate fields. Auth invites use the existing single-use GET.
                    If isInvite Then
                        draft = Await HelperClient.RegisterAsync(link, candidate, Lifetime.Token)
                    Else
                        draft = HelperClient.ApplyConnectionLink(link, candidate)
                    End If
                    draft.AppliedRegistryLinkFingerprint = ""
                    receivedCredentials = True
                    pendingCredentialSave = True
                    context.CommitCredentials(draft)
                    pendingCredentialSave = False
                    showRegistrationHelp = False
                    PublishSettings(UiStrings.GetText(If(isInvite, "settings.registrationSuccess", "settings.configurationLinkApplied"), draft.UiLanguage))
                    SendView(New With {.[type] = "clearInvite"})
                Case "unregister"
                    candidate.UserId = ""
                    candidate.DeviceTokenProtected = ""
                    candidate.AppliedRegistryLinkFingerprint = ""
                    context.CommitCredentials(candidate)
                    draft = candidate
                    pendingCredentialSave = False
                    PublishSettings(UiStrings.GetText("settings.unregisterDone", previewLanguage))
            End Select
        Catch ex As System.Exception
            ClientLog.Write("settings." & action & ".failed", ex)
            Dim feedback As String = HelperClient.Describe(ex, previewLanguage)
            If action = "register" Then
                feedback &= " " & UiStrings.GetText(If(receivedCredentials, "desktop.registrationPendingSave", "desktop.inviteCaution"), previewLanguage)
                If receivedCredentials Then PublishSettings(feedback, True)
                SendView(New With {.[type] = "clearInvite"})
            End If
            SendView(New With {.[type] = "error", .message = feedback})
        Finally
            working = False
            SendView(New With {.[type] = "busy", .busy = False})
        End Try
    End Function

    Friend Sub RecordCurrentHotkey(modifiers As UInteger, key As Integer, isOcr As Boolean)
        If Not IsRecordingHotkey Then Return
        IsRecordingHotkey = False
        SendView(New With {.[type] = "hotkey", .modifiers = modifiers, .key = key, .ocr = isOcr})
    End Sub

    Protected Overrides Sub OnFormClosing(e As System.Windows.Forms.FormClosingEventArgs)
        If e.CloseReason = System.Windows.Forms.CloseReason.UserClosing AndAlso (working OrElse pendingCredentialSave) Then
            e.Cancel = True
            If pendingCredentialSave Then SendView(New With {.[type] = "error", .message = UiStrings.GetText("desktop.registrationPendingSave", previewLanguage)})
        End If
        IsRecordingHotkey = False
        MyBase.OnFormClosing(e)
    End Sub

    Private NotInheritable Class SettingsValues
        Public Sub New()
        End Sub
        Public Property HelperUrl As String = ""
        Public Property ApiKey As String = ""
        Public Property ApiKeyChanged As Boolean
        Public Property LicenseUser As String = ""
        Public Property Language As String = "de"
        Public Property DarkMode As Boolean
        Public Property FontSize As Integer = 12
        Public Property Temperature As Double = 0.7
        Public Property Model As String = ""
        Public Property SystemPrompt As String = ""
        Public Property ShowLlmDebug As Boolean
        Public Property IgnoreRegistryHelperLink As Boolean
        Public Property HotkeyModifiers As UInteger = 3UI
        Public Property HotkeyKey As Integer = 75
        Public Property OcrHotkeyModifiers As UInteger = 3UI
        Public Property OcrHotkeyKey As Integer = 79
    End Class
End Class
