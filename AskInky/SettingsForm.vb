' Part of "Red Ink" (Ask Inky)
' Copyright (c) LawDigital AG, Switzerland. All rights reserved. For license to use see https://redink.ai.
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
    ' 2.0.4 (DARO UID 577): guided first-run sign-in (only the sign-in steps, then straight to the question window).
    Private ReadOnly welcomeMode As Boolean
    Private Shared ReadOnly UsedInvites As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.Ordinal)
    Public Property IsRecordingHotkey As Boolean
    Public Property RecordingOcrHotkey As Boolean

    Public Sub New(parent As AssistantContext, config As ClientConfig, Optional needsRegistration As Boolean = False, Optional startupFeedback As String = "", Optional mustSaveCredentials As Boolean = False, Optional welcome As Boolean = False)
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
        welcomeMode = welcome
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
        Text = "Ask Inky – " & UiStrings.GetText(If(welcomeMode, "welcome.windowTitle", "settings"), previewLanguage)
        SendView(New With {.[type] = "settings", .strings = UiStrings.GetBundle(previewLanguage), .bundles = UiStrings.GetAllBundles(),
                           .language = previewLanguage, .darkMode = draft.DarkMode, .fontSize = draft.FontSize,
                           .values = New With {.helperUrl = draft.HelperUrl, .apiKey = key, .apiKeyReadable = keyReadable,
                               .licenseUser = draft.LicenseUser, .language = previewLanguage, .darkMode = draft.DarkMode,
                               .fontSize = draft.FontSize, .temperature = draft.Temperature, .model = draft.Model, .systemPrompt = draft.SystemPrompt,
                               .showLlmDebug = draft.ShowLlmDebug, .enableReplace = draft.EnableReplace, .startWithWindows = If(draft.StartWithWindows.HasValue, draft.StartWithWindows.Value, Autostart.IsEnabled()), .ignoreRegistryHelperLink = draft.IgnoreRegistryHelperLink,
                               .registryManaged = registryManaged, .hotkeyModifiers = draft.HotkeyModifiers, .hotkeyKey = draft.HotkeyKey,
                               .ocrHotkeyModifiers = draft.OcrHotkeyModifiers, .ocrHotkeyKey = draft.OcrHotkeyKey,
                               .registered = Not String.IsNullOrEmpty(draft.UserId) AndAlso Not String.IsNullOrEmpty(draft.DeviceTokenProtected), .userId = draft.UserId},
                           .debugSent = HelperClient.LastRequest, .debugReceived = HelperClient.LastResponse, .feedback = feedback, .failure = failure,
                           .needsRegistration = showRegistrationHelp AndAlso Not welcomeMode, .welcome = welcomeMode})
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
        result.EnableReplace = input.EnableReplace
        result.StartWithWindows = input.StartWithWindows
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
                Text = "Ask Inky – " & UiStrings.GetText(If(welcomeMode, "welcome.windowTitle", "settings"), previewLanguage)
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
        If action = "pairCancel" Then
            pairing?.Cancel()
            Return
        End If
        If action = "signInOffer" Then
            Try
                Dim url As String = HelperClient.NormalizeHelperUrl(HelperClient.TextValue(message, "helperUrl"))
                ' P18 W367: nothing configured yet -> offer the e-mail sign-in; the e-mail domain finds the server.
                If HelperModern.IsUnconfigured(url) Then
                    SendView(New With {.[type] = "signInOffer", .mailPairing = True, .microsoft = False, .discover = True})
                    Return
                End If
                Dim offer As SignInOffer = Await HelperModern.GetSignInOfferAsync(url, Lifetime.Token)
                SendView(New With {.[type] = "signInOffer", .mailPairing = offer.MailPairing, .microsoft = offer.Microsoft})
            Catch ex As System.Exception
                SendView(New With {.[type] = "signInOffer", .mailPairing = False, .microsoft = False})
            End Try
            Return
        End If
        If working Then Return
        ' P18 W367 2.0.3 (DARO UID 572): find the server from the e-mail address first, then offer Microsoft sign-in
        ' (if the server supports desktop SSO) before the confirmation mail.
        If action = "discover" Then
            working = True
            SendView(New With {.[type] = "busy", .busy = True})
            Try
                Dim email As String = HelperClient.TextValue(message, "email").Trim()
                If email.Length < 3 OrElse email.Length > 320 OrElse Not email.Contains("@"c) Then Throw New HelperFailure("desktop.emailInvalid")
                Dim found As String = Await HelperModern.DiscoverHelperAsync(email, Lifetime.Token)
                If found.Length = 0 Then Throw New HelperFailure("desktop.helperNotFound")
                SendView(New With {.[type] = "helperDiscovered", .helperUrl = found})
                Dim offer As SignInOffer = Await HelperModern.GetSignInOfferAsync(found, Lifetime.Token)
                SendView(New With {.[type] = "signInOffer", .mailPairing = offer.MailPairing, .microsoft = offer.Microsoft})
            Catch ex As System.Exception
                ClientLog.Write("settings.discover.failed", ex)
                SendView(New With {.[type] = "error", .message = HelperClient.Describe(ex, previewLanguage)})
            Finally
                working = False
                SendView(New With {.[type] = "busy", .busy = False})
            End Try
            Return
        End If
        If action = "cancel" Then
            Close()
            Return
        End If
        If action = "pairStart" OrElse action = "microsoft" Then
            Await SignInAsync(action, message)
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
                    If candidate.StartWithWindows.HasValue AndAlso Not Autostart.Apply(candidate.StartWithWindows.Value) Then ClientLog.Write("settings.autostart.apply.failed")
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
                    SendView(New With {.[type] = "clearInvite"})
                    If isInvite AndAlso Not String.IsNullOrEmpty(draft.DeviceTokenProtected) Then
                        CompleteSignIn(draft)
                    Else
                        PublishSettings(UiStrings.GetText(If(isInvite, "settings.registrationSuccess", "settings.configurationLinkApplied"), draft.UiLanguage))
                        SendView(New With {.[type] = "requestOffer"})
                    End If
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

    ' P18 W367 2.0.4 (DARO UID 577): after a successful sign-in everything is saved at once (no "Speichern" needed),
    ' the window shows a clear success state and then closes; the question window opens ready to use.
    Private finishing As Boolean
    Private Sub CompleteSignIn(signedIn As ClientConfig)
        Dim saved As Boolean = False
        Try
            saved = context.TryCommit(signedIn)
            If saved AndAlso signedIn.StartWithWindows.HasValue Then Autostart.Apply(signedIn.StartWithWindows.Value)
        Catch ex As System.Exception
            ClientLog.Write("settings.signin.autosave.failed", ex)
        End Try
        If Not saved Then
            ' Credentials are stored already; other changed settings still need "Speichern" (e.g. a shortcut conflict).
            PublishSettings(UiStrings.GetText("desktop.signedIn", previewLanguage))
            Return
        End If
        draft = signedIn
        finishing = True
        SendView(New With {.[type] = "signInSucceeded", .user = If(String.IsNullOrEmpty(signedIn.UserId), signedIn.LicenseUser, signedIn.UserId)})
        FinishAfterDelayAsync()
    End Sub

    Private Async Sub FinishAfterDelayAsync()
        Try
            Await System.Threading.Tasks.Task.Delay(1800)
            If IsDisposed Then Return
            working = False
            pendingCredentialSave = False
            Close()
            context.ShowAfterSignIn()
        Catch ex As System.Exception
            ClientLog.Write("settings.signin.finish.failed", ex)
        End Try
    End Sub
    ' P18 W367: "Anmelden ohne Link" (mail confirmation, like Office Web) and Microsoft sign-in.
    Private pairing As System.Threading.CancellationTokenSource

    Private Async Function SignInAsync(action As String, message As System.Text.Json.JsonElement) As System.Threading.Tasks.Task
        working = True
        IsRecordingHotkey = False
        SendView(New With {.[type] = "busy", .busy = True})
        Dim receivedCredentials As Boolean = False
        pairing = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(Lifetime.Token)
        Try
            Dim candidate As ClientConfig = ReadCandidate(message)
            If registryManaged AndAlso Not candidate.IgnoreRegistryHelperLink Then Throw New HelperFailure("settings.registryManagedHint")
            Dim signedIn As ClientConfig
            If action = "pairStart" Then
                Dim email As String = HelperClient.TextValue(message, "email").Trim()
                If email.Length < 3 OrElse email.Length > 320 OrElse Not email.Contains("@"c) Then Throw New HelperFailure("desktop.emailInvalid")
                If HelperModern.IsUnconfigured(candidate.HelperUrl) Then
                    SendView(New With {.[type] = "status", .message = UiStrings.GetText("desktop.helperSearching", previewLanguage)})
                    Dim found As String = Await HelperModern.DiscoverHelperAsync(email, pairing.Token)
                    If found.Length = 0 Then Throw New HelperFailure("desktop.helperNotFound")
                    candidate.HelperUrl = found
                    candidate.CredentialOrigin = HelperClient.GetOrigin(found)
                    candidate.UserId = "" : candidate.DeviceTokenProtected = "" : candidate.ApiKeyProtected = ""
                    SendView(New With {.[type] = "helperDiscovered", .helperUrl = found})
                End If
                Dim started As PairingStart = Await HelperModern.StartPairingAsync(candidate.HelperUrl, email, previewLanguage, pairing.Token)
                SendView(New With {.[type] = "pairingCode", .code = started.Code})
                signedIn = Await HelperModern.CompletePairingAsync(candidate.HelperUrl, started, candidate, pairing.Token)
            Else
                SendView(New With {.[type] = "status", .message = UiStrings.GetText("desktop.microsoftWaiting", previewLanguage)})
                signedIn = Await HelperModern.SignInWithMicrosoftAsync(candidate.HelperUrl, candidate, pairing.Token)
            End If
            receivedCredentials = True
            pendingCredentialSave = True
            signedIn.AppliedRegistryLinkFingerprint = ""
            context.CommitCredentials(signedIn)
            draft = signedIn
            pendingCredentialSave = False
            showRegistrationHelp = False
            SendView(New With {.[type] = "pairingDone"})
            CompleteSignIn(signedIn)
        Catch ex As System.OperationCanceledException When Not Lifetime.IsCancellationRequested
            SendView(New With {.[type] = "pairingDone"})
            SendView(New With {.[type] = "error", .message = UiStrings.GetText("desktop.signInCancelled", previewLanguage)})
        Catch ex As System.Exception
            ClientLog.Write("settings." & action & ".failed", ex)
            Dim feedback As String = HelperClient.Describe(ex, previewLanguage)
            If receivedCredentials Then
                feedback &= " " & UiStrings.GetText("desktop.registrationPendingSave", previewLanguage)
                PublishSettings(feedback, True)
            End If
            SendView(New With {.[type] = "pairingDone"})
            SendView(New With {.[type] = "error", .message = feedback})
        Finally
            pairing.Dispose()
            pairing = Nothing
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
        Public Property EnableReplace As Boolean
        Public Property StartWithWindows As Boolean = True
        Public Property IgnoreRegistryHelperLink As Boolean
        Public Property HotkeyModifiers As UInteger = 3UI
        Public Property HotkeyKey As Integer = 75
        Public Property OcrHotkeyModifiers As UInteger = 3UI
        Public Property OcrHotkeyKey As Integer = 79
    End Class
End Class
