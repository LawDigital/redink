' Part of Red Ink (AskInky). Office Web Helper contracts; transport stays outside the WebView.
Friend NotInheritable Class HelperFailure
    Inherits System.Exception
    Public ReadOnly Property ResourceKey As String
    Public ReadOnly Property HttpStatus As Integer
    Public Sub New(key As String, Optional status As Integer = 0)
        MyBase.New(key)
        ResourceKey = key
        HttpStatus = status
    End Sub
End Class

Friend NotInheritable Class HelperProbe
    Public Property Version As String = ""
    Public Property LicenseKey As String = "desktop.licenseUnknown"
    Public Property LicenseProduct As String = ""
    Public Property Authorized As Boolean
    Public Property HttpStatus As Integer
End Class

Public NotInheritable Class HelperClient
    Private Shared ReadOnly Http As New System.Net.Http.HttpClient(New System.Net.Http.HttpClientHandler With {.AllowAutoRedirect = False, .UseCookies = False}) With {.Timeout = System.TimeSpan.FromSeconds(120)}
    Private Shared ReadOnly SessionId As String = System.Guid.NewGuid().ToString("N")
    Friend Shared Property LastRequest As String = ""
    Friend Shared Property LastResponse As String = ""

    Public Shared Function NormalizeHelperUrl(value As String) As String
        Dim raw As String = If(value, "").Trim()
        If raw.Length = 0 Then raw = "http://localhost"
        If Not raw.Contains("://", System.StringComparison.Ordinal) Then
            Dim candidate As System.Uri = Nothing
            If Not System.Uri.TryCreate("http://" & raw, System.UriKind.Absolute, candidate) Then Throw New HelperFailure("desktop.invalidUrl")
            raw = If(candidate.IsLoopback, "http://", "https://") & raw
        End If
        Dim result As System.Uri = Nothing
        If Not System.Uri.TryCreate(raw, System.UriKind.Absolute, result) Then Throw New HelperFailure("desktop.invalidUrl")
        If result.Scheme <> System.Uri.UriSchemeHttps AndAlso Not (result.Scheme = System.Uri.UriSchemeHttp AndAlso result.IsLoopback) Then Throw New HelperFailure("desktop.httpsRequired")
        If result.Host.Length = 0 OrElse result.UserInfo.Length <> 0 OrElse result.Query.Length <> 0 OrElse result.Fragment.Length <> 0 Then Throw New HelperFailure("desktop.invalidUrl")
        Return result.GetLeftPart(System.UriPartial.Path).TrimEnd("/"c)
    End Function

    Public Shared Function GetOrigin(value As String) As String
        Return New System.Uri(NormalizeHelperUrl(value)).GetLeftPart(System.UriPartial.Authority)
    End Function

    Private Shared Function CreateRequest(config As ClientConfig, method As System.Net.Http.HttpMethod, path As String) As System.Net.Http.HttpRequestMessage
        Dim root As String = NormalizeHelperUrl(config.HelperUrl)
        Dim req As New System.Net.Http.HttpRequestMessage(method, New System.Uri(New System.Uri(root & "/"), path))
        Try
            req.Headers.Accept.Add(New System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"))
            req.Headers.Add("X-RedInk-Session", SessionId)
            If Not String.IsNullOrWhiteSpace(config.LicenseUser) Then req.Headers.Add("X-License-User", config.LicenseUser.Trim())
            Dim hasSecrets As Boolean = Not String.IsNullOrEmpty(config.ApiKeyProtected) OrElse Not String.IsNullOrEmpty(config.DeviceTokenProtected)
            If hasSecrets AndAlso Not String.Equals(config.CredentialOrigin, GetOrigin(root), System.StringComparison.OrdinalIgnoreCase) Then Throw New HelperFailure("desktop.credentialsDifferentServer")
            If Not String.IsNullOrWhiteSpace(config.UserId) AndAlso Not String.IsNullOrEmpty(config.DeviceTokenProtected) Then
                req.Headers.Add("X-User-Id", config.UserId)
                req.Headers.Add("X-Device-Token", ConfigStore.Reveal(config.DeviceTokenProtected))
            End If
            Dim apiKey As String = ConfigStore.Reveal(config.ApiKeyProtected)
            If Not String.IsNullOrWhiteSpace(apiKey) Then req.Headers.Authorization = New System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey.Trim())
            Return req
        Catch ex As System.Exception
            req.Dispose()
            Throw
        End Try
    End Function

    Friend Shared Function TextValue(root As System.Text.Json.JsonElement, name As String) As String
        Dim value As System.Text.Json.JsonElement
        If root.ValueKind = System.Text.Json.JsonValueKind.Object AndAlso root.TryGetProperty(name, value) AndAlso value.ValueKind = System.Text.Json.JsonValueKind.String Then Return If(value.GetString(), "")
        Return ""
    End Function

    Private Shared Function IsTrue(root As System.Text.Json.JsonElement, name As String) As Boolean
        Dim value As System.Text.Json.JsonElement
        Return root.ValueKind = System.Text.Json.JsonValueKind.Object AndAlso root.TryGetProperty(name, value) AndAlso value.ValueKind = System.Text.Json.JsonValueKind.True
    End Function

    Private Shared Function StatusFailure(status As System.Net.HttpStatusCode) As HelperFailure
        Dim key As String = "desktop.httpError"
        Select Case CInt(status)
            Case 401, 403 : key = "desktop.authRequired"
            Case 402 : key = "desktop.licenseRequired"
            Case 429 : key = "desktop.rateLimited"
        End Select
        Return New HelperFailure(key, CInt(status))
    End Function

    Public Shared Async Function RegisterAsync(link As String, current As ClientConfig, cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of ClientConfig)
        Dim raw As String = If(link, "").Trim()
        Dim invite As System.Uri = Nothing
        If Not System.Uri.TryCreate(raw, System.UriKind.Absolute, invite) Then Throw New HelperFailure("desktop.invalidInvite")
        If invite.UserInfo.Length <> 0 OrElse invite.Fragment.Length <> 0 OrElse (invite.Scheme <> System.Uri.UriSchemeHttps AndAlso Not (invite.Scheme = System.Uri.UriSchemeHttp AndAlso invite.IsLoopback)) Then Throw New HelperFailure("desktop.invalidInvite")
        ' Exact single-use URL; no probes, automatic redirects, cookies or retries.
        Using req As New System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Get, raw)
            req.Headers.Accept.Add(New System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"))
            Using response As System.Net.Http.HttpResponseMessage = Await Http.SendAsync(req, cancellation)
                If Not response.IsSuccessStatusCode Then Throw New HelperFailure("desktop.registrationFailed", CInt(response.StatusCode))
                If response.Content.Headers.ContentType Is Nothing OrElse Not String.Equals(response.Content.Headers.ContentType.MediaType, "application/json", System.StringComparison.OrdinalIgnoreCase) Then Throw New HelperFailure("desktop.registrationFailed")
                Dim json As String = Await response.Content.ReadAsStringAsync(cancellation)
                Using document As System.Text.Json.JsonDocument = System.Text.Json.JsonDocument.Parse(json)
                    Dim payload As System.Text.Json.JsonElement = document.RootElement
                    Dim user As String = TextValue(payload, "userId")
                    Dim token As String = TextValue(payload, "deviceToken")
                    If TextValue(payload, "status") <> "ok" OrElse Not IsTrue(payload, "registered") OrElse String.IsNullOrWhiteSpace(user) OrElse String.IsNullOrWhiteSpace(token) Then Throw New HelperFailure("desktop.registrationFailed")
                    Dim result As ClientConfig = ConfigStore.Clone(current)
                    Dim origin As String = invite.GetLeftPart(System.UriPartial.Authority)
                    If Not String.Equals(current.CredentialOrigin, origin, System.StringComparison.OrdinalIgnoreCase) Then result.ApiKeyProtected = ""
                    result.HelperUrl = origin
                    result.CredentialOrigin = origin
                    result.UserId = user
                    result.DeviceTokenProtected = ConfigStore.Protect(token)
                    result.LicenseUser = user
                    For Each item As String In invite.Query.TrimStart("?"c).Split("&"c, System.StringSplitOptions.RemoveEmptyEntries)
                        Dim pair As String() = item.Split("="c, 2)
                        Dim name As String = System.Uri.UnescapeDataString(pair(0).Replace("+", " "))
                        Dim value As String = If(pair.Length = 2, System.Uri.UnescapeDataString(pair(1).Replace("+", " ")), "")
                        Select Case name.ToLowerInvariant()
                            Case "api_key"
                                If value.Trim().Length > 0 Then result.ApiKeyProtected = ConfigStore.Protect(value.Trim())
                            Case "uid"
                                If value.Trim().Length > 0 Then result.LicenseUser = value.Trim()
                            Case "lang"
                                If UiStrings.IsSupported(value) Then result.UiLanguage = UiStrings.NormalizeLanguage(value)
                        End Select
                    Next
                    Return result
                End Using
            End Using
        End Using
    End Function

    Friend Shared Async Function ProbeAsync(config As ClientConfig, cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of HelperProbe)
        Dim result As New HelperProbe()
        Using timeout As System.Threading.CancellationTokenSource = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(cancellation)
            timeout.CancelAfter(System.TimeSpan.FromSeconds(15))
            Using req As System.Net.Http.HttpRequestMessage = CreateRequest(config, System.Net.Http.HttpMethod.Get, "health")
                Using res As System.Net.Http.HttpResponseMessage = Await Http.SendAsync(req, timeout.Token)
                    If Not res.IsSuccessStatusCode Then Throw StatusFailure(res.StatusCode)
                    Dim json As String = Await res.Content.ReadAsStringAsync(timeout.Token)
                    Using doc As System.Text.Json.JsonDocument = System.Text.Json.JsonDocument.Parse(json)
                        If TextValue(doc.RootElement, "status") <> "ok" Then Throw New HelperFailure("desktop.invalidResponse")
                        result.Version = TextValue(doc.RootElement, "version")
                    End Using
                End Using
            End Using
            ' Like Office Web: /health proves connectivity, /config checks access and licensing.
            Using req As System.Net.Http.HttpRequestMessage = CreateRequest(config, System.Net.Http.HttpMethod.Get, "config")
                Using res As System.Net.Http.HttpResponseMessage = Await Http.SendAsync(req, timeout.Token)
                    result.HttpStatus = CInt(res.StatusCode)
                    If Not res.IsSuccessStatusCode Then
                        result.LicenseKey = If(CInt(res.StatusCode) = 402, "desktop.licenseRequired", If(CInt(res.StatusCode) = 401 OrElse CInt(res.StatusCode) = 403, "desktop.authRequired", "desktop.configUnavailable"))
                        Return result
                    End If
                    Dim json As String = Await res.Content.ReadAsStringAsync(timeout.Token)
                    Using doc As System.Text.Json.JsonDocument = System.Text.Json.JsonDocument.Parse(json)
                        result.Authorized = True
                        Dim settings As System.Text.Json.JsonElement
                        Dim license As System.Text.Json.JsonElement
                        If doc.RootElement.TryGetProperty("config", settings) AndAlso settings.ValueKind = System.Text.Json.JsonValueKind.Object AndAlso settings.TryGetProperty("license", license) AndAlso license.ValueKind = System.Text.Json.JsonValueKind.Object Then
                            If Not IsTrue(license, "enabled") Then
                                result.LicenseKey = "settings.licenseDisabled"
                            Else
                                result.LicenseKey = If(IsTrue(license, "ok"), "settings.licenseOk", "settings.licenseNotConfirmed")
                                Dim product As System.Text.Json.JsonElement
                                If license.TryGetProperty("product_id", product) Then result.LicenseProduct = product.ToString()
                            End If
                        End If
                    End Using
                End Using
            End Using
        End Using
        Return result
    End Function

    Public Shared Async Function AskAsync(config As ClientConfig, prompt As String, text As String, cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of String)
        ' Helper decides whether device registration/licensing is necessary; no local prompt.
        Dim userContent As String = If(String.IsNullOrWhiteSpace(text), prompt, "Ausgewählter Text:" & vbLf & text & vbLf & vbLf & "Frage:" & vbLf & prompt)
        Dim body As String = System.Text.Json.JsonSerializer.Serialize(New With {.prompt_system = config.SystemPrompt, .prompt_user = userContent, .temperature = config.Temperature})
        LastRequest = If(config.ShowLlmDebug, body, "")
        LastResponse = ""
        Using req As System.Net.Http.HttpRequestMessage = CreateRequest(config, System.Net.Http.HttpMethod.Post, "llm/chat")
            req.Content = New System.Net.Http.StringContent(body, System.Text.Encoding.UTF8, "application/json")
            Using res As System.Net.Http.HttpResponseMessage = Await Http.SendAsync(req, cancellation)
                If Not res.IsSuccessStatusCode Then Throw StatusFailure(res.StatusCode)
                Dim response As String = Await res.Content.ReadAsStringAsync(cancellation)
                If config.ShowLlmDebug Then LastResponse = response
                Using doc As System.Text.Json.JsonDocument = System.Text.Json.JsonDocument.Parse(response)
                    Dim payload As System.Text.Json.JsonElement = doc.RootElement
                    Dim content As System.Text.Json.JsonElement
                    If Not IsTrue(payload, "success") OrElse Not payload.TryGetProperty("content", content) OrElse content.ValueKind <> System.Text.Json.JsonValueKind.String Then Throw New HelperFailure("desktop.invalidResponse")
                    Return If(content.GetString(), "")
                End Using
            End Using
        End Using
    End Function

    Friend Shared Function Describe(failure As System.Exception, language As String) As String
        Dim known As HelperFailure = TryCast(failure, HelperFailure)
        If known IsNot Nothing Then Return UiStrings.GetText(known.ResourceKey, language) & If(known.HttpStatus > 0, " (HTTP " & known.HttpStatus.ToString(System.Globalization.CultureInfo.InvariantCulture) & ")", "")
        If TypeOf failure Is System.OperationCanceledException Then Return UiStrings.GetText("desktop.timeout", language)
        If TypeOf failure Is System.Net.Http.HttpRequestException Then Return UiStrings.GetText("desktop.networkError", language)
        If TypeOf failure Is System.Security.Cryptography.CryptographicException Then Return UiStrings.GetText("desktop.decryptFailed", language)
        If TypeOf failure Is System.Text.Json.JsonException Then Return UiStrings.GetText("desktop.invalidResponse", language)
        Return UiStrings.GetText("desktop.error", language)
    End Function
End Class
