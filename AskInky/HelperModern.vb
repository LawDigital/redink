' Part of "Red Ink" (Ask Inky)
' Copyright (c) LawDigital AG, Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: HelperModern.vb
' Purpose: Ask Inky 2 features of newer Helpers: capability negotiation, agent mode, thinking level,
'          web search, sign-in without a link, Microsoft sign-in, file text and self-update.
' Architecture: Windows native host with embedded Office Web-based interface.
' Notes: Red Ink License and Ask Inky third-party attributions are embedded in LICENSE.txt.
' =============================================================================
' Part of Red Ink (AskInky). P18 W367 compatibility rule: every new request field is sent ONLY when the
' Helper announced it (capabilities desktopClientVersion >= 1, discovery, /auth/entra/status). Older and
' "mini" Helpers keep the unchanged /llm/chat path; they never receive an unknown field.
Public NotInheritable Class HelperModelCapability
    Public Property Key As String = ""
    Public Property ThinkingLevels As New System.Collections.Generic.List(Of String)()
    Public Property ThinkingStart As String = ""
    Public Property WebGrounding As Boolean
End Class

Public NotInheritable Class HelperCapabilities
    Public Property DesktopVersion As Integer
    Public Property AgentAvailable As Boolean
    Public Property WebSearchAvailable As Boolean
    Public Property DefaultModelKey As String = ""
    Public Property Models As New System.Collections.Generic.List(Of HelperModelCapability)()
    Public Function ModelFor(key As String) As HelperModelCapability
        Dim wanted As String = If(String.IsNullOrWhiteSpace(key), DefaultModelKey, key.Trim())
        Return Models.FirstOrDefault(Function(m) String.Equals(m.Key, wanted, System.StringComparison.Ordinal))
    End Function
End Class

Public NotInheritable Class AskOptions
    Public Property Agent As Boolean
    Public Property ThinkingLevel As String = ""
    Public Property Web As Boolean
    Public Property ConversationId As String = ""
    Public Property ThreadId As String = ""
End Class

Public NotInheritable Class AskOutcome
    Public Property Text As String = ""
    Public Property ThreadId As String = ""
    Public Property Status As String = ""
End Class

Public NotInheritable Class PairingStart
    Public Property Handle As String = ""
    Public Property Code As String = ""
    Public Property PollAfterSeconds As Integer = 3
    Public Property ExpiresIn As Integer = 900
End Class

Public NotInheritable Class SignInOffer
    Public Property MailPairing As Boolean
    Public Property Microsoft As Boolean
End Class

Public NotInheritable Class ClientUpdateInfo
    Public Property Version As String = ""
    Public Property FileName As String = ""
    Public Property Sha256 As String = ""
    Public Property Size As Long
    Public Property SignerOrganization As String = ""
    Public Property DownloadPath As String = ""
End Class

Public NotInheritable Class HelperModern
    Public Const ClientVersion As String = "2.0.4"
    Private Shared ReadOnly TerminalStates As New System.Collections.Generic.HashSet(Of String)({"COMPLETED", "FAILED", "BLOCKED", "CANCELLED"}, System.StringComparer.Ordinal)
    ' Organisations whose Authenticode signature Ask Inky accepts for a self-update.
    Private Shared ReadOnly AllowedSigners As String() = {"VISCHER AG", "LawDigital AG"}

    ' ---------------------------------------------------------------- capabilities

    ' Nothing = old or mini Helper (no interactive channel, or no desktop host kind): stay on /llm/chat.
    Public Shared Async Function GetCapabilitiesAsync(config As ClientConfig, cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of HelperCapabilities)
        Try
            Using timeout As System.Threading.CancellationTokenSource = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(cancellation)
                timeout.CancelAfter(System.TimeSpan.FromSeconds(20))
                Using req As System.Net.Http.HttpRequestMessage = HelperClient.CreateRequest(config, System.Net.Http.HttpMethod.Get, "runtime/v1/interactive/capabilities?channel=office_web")
                    Using res As System.Net.Http.HttpResponseMessage = Await HelperClient.Http.SendAsync(req, timeout.Token)
                        If Not res.IsSuccessStatusCode Then Return Nothing
                        Using doc As System.Text.Json.JsonDocument = System.Text.Json.JsonDocument.Parse(Await res.Content.ReadAsStringAsync(timeout.Token))
                            Return ParseCapabilities(doc.RootElement)
                        End Using
                    End Using
                End Using
            End Using
        Catch ex As System.OperationCanceledException When Not cancellation.IsCancellationRequested
            ClientLog.Write("helper.capabilities.timeout")
            Return Nothing
        Catch ex As HelperFailure
            Throw
        Catch ex As System.Exception
            ClientLog.Write("helper.capabilities.unavailable", ex)
            Return Nothing
        End Try
    End Function

    Friend Shared Function ParseCapabilities(root As System.Text.Json.JsonElement) As HelperCapabilities
        If root.ValueKind <> System.Text.Json.JsonValueKind.Object Then Return Nothing
        Dim version As System.Text.Json.JsonElement
        Dim desktop As Integer = 0
        If root.TryGetProperty("desktopClientVersion", version) AndAlso version.ValueKind = System.Text.Json.JsonValueKind.Number Then version.TryGetInt32(desktop)
        If desktop < 1 Then Return Nothing
        Dim result As New HelperCapabilities With {.DesktopVersion = desktop, .DefaultModelKey = HelperClient.TextValue(root, "default_model_key")}
        Dim modes As System.Text.Json.JsonElement
        If root.TryGetProperty("supported_interaction_modes", modes) AndAlso modes.ValueKind = System.Text.Json.JsonValueKind.Array Then
            For Each mode As System.Text.Json.JsonElement In modes.EnumerateArray()
                If mode.ValueKind = System.Text.Json.JsonValueKind.String AndAlso mode.GetString() = "AGENT" Then result.AgentAvailable = True
            Next
        End If
        result.WebSearchAvailable = HelperClient.IsTrue(root, "webSearchAvailable")
        Dim models As System.Text.Json.JsonElement
        If root.TryGetProperty("available_models", models) AndAlso models.ValueKind = System.Text.Json.JsonValueKind.Array Then
            For Each item As System.Text.Json.JsonElement In models.EnumerateArray()
                Dim key As String = HelperClient.TextValue(item, "key")
                If key.Length = 0 OrElse key.Length > 256 Then Continue For
                Dim model As New HelperModelCapability With {.Key = key, .WebGrounding = HelperClient.IsTrue(item, "webGrounding"), .ThinkingStart = HelperClient.TextValue(item, "thinkingStart")}
                Dim levels As System.Text.Json.JsonElement
                If item.TryGetProperty("thinkingLevels", levels) AndAlso levels.ValueKind = System.Text.Json.JsonValueKind.Array Then
                    For Each level As System.Text.Json.JsonElement In levels.EnumerateArray()
                        If level.ValueKind = System.Text.Json.JsonValueKind.String AndAlso level.GetString().Length <= 24 Then model.ThinkingLevels.Add(level.GetString())
                    Next
                End If
                result.Models.Add(model)
            Next
        End If
        Return result
    End Function

    ' ---------------------------------------------------------------- asking (interactive runs)

    Private Shared Sub AddDesktopHeaders(req As System.Net.Http.HttpRequestMessage, conversationId As String)
        req.Headers.Add("X-RedInk-Office-Host", "desktop")
        If Not String.IsNullOrEmpty(conversationId) Then req.Headers.Add("X-RedInk-Office-Conversation-Document", conversationId)
    End Sub

    Public Shared Function NewConversationId() As String
        Return "askinky-" & System.Guid.NewGuid().ToString("N")
    End Function

    Public Shared Async Function AskAsync(config As ClientConfig, caps As HelperCapabilities, prompt As String, text As String, options As AskOptions,
                                          progress As System.Action(Of String), cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of AskOutcome)
        Dim userContent As String = If(String.IsNullOrWhiteSpace(text), prompt, "Ausgewählter Text:" & vbLf & text & vbLf & vbLf & "Frage:" & vbLf & prompt)
        Dim body As New System.Collections.Generic.Dictionary(Of String, Object) From {
            {"channel", "office_web"},
            {"interaction_mode", If(options.Agent AndAlso caps.AgentAvailable, "AGENT", "STANDARD")},
            {"message", userContent},
            {"current_user_text", prompt},
            {"interface_language", config.UiLanguage},
            {"client_request_id", System.Guid.NewGuid().ToString()},
            {"host_context", New System.Collections.Generic.Dictionary(Of String, Object) From {
                {"host_kind", "desktop"}, {"client_version", "askinky-" & ClientVersion}, {"conversation_document_id", options.ConversationId}}}}
        If Not options.Agent AndAlso Not String.IsNullOrWhiteSpace(config.Model) Then body("model_selection") = config.Model.Trim()
        If Not String.IsNullOrEmpty(options.ThreadId) Then body("thread_id") = options.ThreadId
        Dim model As HelperModelCapability = caps.ModelFor(config.Model)
        If Not String.IsNullOrEmpty(options.ThinkingLevel) AndAlso options.ThinkingLevel <> "default" AndAlso model IsNot Nothing AndAlso model.ThinkingLevels.Contains(options.ThinkingLevel) Then
            body("thinking_level") = options.ThinkingLevel
        End If
        If options.Web AndAlso Not options.Agent Then body("web_grounding") = True
        Dim json As String = System.Text.Json.JsonSerializer.Serialize(body)
        HelperClient.LastRequest = If(config.ShowLlmDebug, json, "")
        HelperClient.LastResponse = ""
        Dim runId As String
        Dim outcome As New AskOutcome()
        Using req As System.Net.Http.HttpRequestMessage = HelperClient.CreateRequest(config, System.Net.Http.HttpMethod.Post, "runtime/v1/interactive/runs")
            AddDesktopHeaders(req, options.ConversationId)
            req.Content = New System.Net.Http.StringContent(json, System.Text.Encoding.UTF8, "application/json")
            Using res As System.Net.Http.HttpResponseMessage = Await HelperClient.Http.SendAsync(req, cancellation)
                Dim answer As String = Await res.Content.ReadAsStringAsync(cancellation)
                If Not res.IsSuccessStatusCode Then
                    If CInt(res.StatusCode) = 413 Then Throw New HelperFailure("desktop.historyTooLong", 413)
                    Throw HelperClient.StatusFailure(res.StatusCode, config)
                End If
                Using doc As System.Text.Json.JsonDocument = System.Text.Json.JsonDocument.Parse(answer)
                    runId = HelperClient.TextValue(doc.RootElement, "run_id")
                    outcome.ThreadId = HelperClient.TextValue(doc.RootElement, "thread_id")
                End Using
            End Using
        End Using
        If Not System.Text.RegularExpressions.Regex.IsMatch(runId, "^[0-9a-fA-F-]{36}$") Then Throw New HelperFailure("desktop.invalidResponse")
        Dim started As System.DateTime = System.DateTime.UtcNow
        Dim delay As Integer = 700
        Try
            Do
                Await System.Threading.Tasks.Task.Delay(delay, cancellation)
                delay = System.Math.Min(2000, delay + 200)
                Using req As System.Net.Http.HttpRequestMessage = HelperClient.CreateRequest(config, System.Net.Http.HttpMethod.Get, "runtime/v1/interactive/runs/" & runId)
                    AddDesktopHeaders(req, options.ConversationId)
                    Using res As System.Net.Http.HttpResponseMessage = Await HelperClient.Http.SendAsync(req, cancellation)
                        If Not res.IsSuccessStatusCode Then Throw HelperClient.StatusFailure(res.StatusCode, config)
                        Dim content As String = Await res.Content.ReadAsStringAsync(cancellation)
                        Using doc As System.Text.Json.JsonDocument = System.Text.Json.JsonDocument.Parse(content)
                            Dim status As String = HelperClient.TextValue(doc.RootElement, "status")
                            If TerminalStates.Contains(status) Then
                                If config.ShowLlmDebug Then HelperClient.LastResponse = content
                                outcome.Status = status
                                outcome.Text = TerminalText(doc.RootElement, status)
                                Return outcome
                            End If
                        End Using
                    End Using
                End Using
                If progress IsNot Nothing Then progress((System.DateTime.UtcNow - started).TotalSeconds.ToString("0", System.Globalization.CultureInfo.InvariantCulture))
            Loop While (System.DateTime.UtcNow - started).TotalMinutes < 30
        Catch ex As System.OperationCanceledException
            CancelQuietly(config, runId, options.ConversationId)
            Throw
        End Try
        Throw New System.TimeoutException("run did not finish")
    End Function

    Private Shared Sub CancelQuietly(config As ClientConfig, runId As String, conversationId As String)
        Try
            Using req As System.Net.Http.HttpRequestMessage = HelperClient.CreateRequest(config, System.Net.Http.HttpMethod.Post, "runtime/v1/interactive/runs/" & runId & "/cancel")
                AddDesktopHeaders(req, conversationId)
                req.Content = New System.Net.Http.StringContent("{}", System.Text.Encoding.UTF8, "application/json")
                HelperClient.Http.SendAsync(req).Wait(5000)
            End Using
        Catch ex As System.Exception
            ClientLog.Write("helper.run.cancel.failed", ex)
        End Try
    End Sub

    ' The Helper's terminal projection: 'result' is the answer; failures carry a user 'presentation'.
    Friend Shared Function TerminalText(root As System.Text.Json.JsonElement, status As String) As String
        Dim answer As String = HelperClient.TextValue(root, "result")
        Dim presentation As String = ""
        Dim value As System.Text.Json.JsonElement
        If root.TryGetProperty("presentation", value) Then
            If value.ValueKind = System.Text.Json.JsonValueKind.String Then
                presentation = value.GetString()
            ElseIf value.ValueKind = System.Text.Json.JsonValueKind.Object Then
                For Each name As String In {"text", "markdown", "content", "message"}
                    Dim candidate As String = HelperClient.TextValue(value, name)
                    If candidate.Length > 0 Then presentation = candidate : Exit For
                Next
            End If
        End If
        If status = "COMPLETED" AndAlso answer.Trim().Length > 0 Then Return answer
        If answer.Trim().Length > 0 AndAlso presentation.Trim().Length > 0 AndAlso Not presentation.Contains(answer.Trim()) Then Return presentation & vbLf & vbLf & answer
        If presentation.Trim().Length > 0 Then Return presentation
        If answer.Trim().Length > 0 Then Return answer
        Return "*" & If(status = "CANCELLED", "Abgebrochen.", "Keine Antwort erhalten (" & status & ").") & "*"
    End Function

    ' ---------------------------------------------------------------- finding the Helper from the e-mail address

    ' P18 W367 (DARO UID 551): the same zero-configuration discovery as the Office add-in (W248).
    ' Candidates: https://<domain>/.well-known/redink-helper (a pointer {helper} or a discovery document),
    ' https://redink.<domain>, https://inky.<domain>. A candidate counts only when ITS OWN discovery document
    ' says product redink-helper AND lists the e-mail domain (so a foreign server cannot claim the domain).
    Public Shared Async Function DiscoverHelperAsync(email As String, cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of String)
        Dim at As Integer = If(email, "").LastIndexOf("@"c)
        If at < 1 Then Return ""
        Dim domain As String = email.Substring(at + 1).Trim().ToLowerInvariant()
        If Not System.Text.RegularExpressions.Regex.IsMatch(domain, "^[a-z0-9]([a-z0-9-]{0,61}[a-z0-9])?(\.[a-z0-9]([a-z0-9-]{0,61}[a-z0-9])?)+$") Then Return ""
        Dim candidates As New System.Collections.Generic.List(Of String)()
        Using timeout As System.Threading.CancellationTokenSource = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(cancellation)
            timeout.CancelAfter(System.TimeSpan.FromSeconds(25))
            Dim pointer As System.Text.Json.JsonElement? = Await GetJsonAsync("https://" & domain & "/.well-known/redink-helper", timeout.Token)
            If pointer.HasValue Then
                Dim helper As String = HelperClient.TextValue(pointer.Value, "helper")
                If helper.Length = 0 Then helper = HelperClient.TextValue(pointer.Value, "origin")
                If helper.Length > 0 Then candidates.Add(helper)
            End If
            candidates.Add("https://redink." & domain)
            candidates.Add("https://inky." & domain)
            For Each candidate As String In candidates.Distinct(System.StringComparer.OrdinalIgnoreCase)
                Dim root As String
                Try
                    root = HelperClient.NormalizeHelperUrl(candidate)
                Catch
                    Continue For
                End Try
                If Not root.StartsWith("https://", System.StringComparison.OrdinalIgnoreCase) Then Continue For
                Dim doc As System.Text.Json.JsonElement? = Await GetJsonAsync(root & "/.well-known/redink-helper", timeout.Token)
                If Not doc.HasValue OrElse HelperClient.TextValue(doc.Value, "product") <> "redink-helper" Then Continue For
                Dim domains As System.Text.Json.JsonElement
                If doc.Value.TryGetProperty("domains", domains) AndAlso domains.ValueKind = System.Text.Json.JsonValueKind.Array AndAlso
                   domains.EnumerateArray().Any(Function(d) d.ValueKind = System.Text.Json.JsonValueKind.String AndAlso String.Equals(d.GetString(), domain, System.StringComparison.OrdinalIgnoreCase)) Then
                    Dim origin As String = HelperClient.TextValue(doc.Value, "origin")
                    ClientLog.Write("helper.discovery.found")
                    Return If(origin.StartsWith("https://", System.StringComparison.OrdinalIgnoreCase), HelperClient.NormalizeHelperUrl(origin), root)
                End If
            Next
        End Using
        ClientLog.Write("helper.discovery.none")
        Return ""
    End Function

    Private Shared Async Function GetJsonAsync(url As String, cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of System.Text.Json.JsonElement?)
        Try
            Using res As System.Net.Http.HttpResponseMessage = Await HelperClient.Http.GetAsync(url, cancellation)
                If Not res.IsSuccessStatusCode Then Return Nothing
                Dim text As String = Await res.Content.ReadAsStringAsync(cancellation)
                If text.Length > 65536 Then Return Nothing
                Using doc As System.Text.Json.JsonDocument = System.Text.Json.JsonDocument.Parse(text)
                    If doc.RootElement.ValueKind <> System.Text.Json.JsonValueKind.Object Then Return Nothing
                    Return doc.RootElement.Clone()
                End Using
            End Using
        Catch ex As System.OperationCanceledException When Not cancellation.IsCancellationRequested
            Return Nothing
        Catch ex As System.Exception When Not (TypeOf ex Is System.OperationCanceledException)
            Return Nothing
        End Try
    End Function

    ' True for the factory default address (nothing configured yet).
    Public Shared Function IsUnconfigured(helperUrl As String) As Boolean
        Dim value As String = If(helperUrl, "").Trim().TrimEnd("/"c)
        Return value.Length = 0 OrElse String.Equals(value, "http://localhost", System.StringComparison.OrdinalIgnoreCase)
    End Function

    ' ---------------------------------------------------------------- sign-in offers

    ' Unauthenticated discovery on the candidate Helper. Old Helpers answer 404: nothing offered.
    Public Shared Async Function GetSignInOfferAsync(helperUrl As String, cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of SignInOffer)
        Dim offer As New SignInOffer()
        Dim root As String = HelperClient.NormalizeHelperUrl(helperUrl)
        Try
            Using timeout As System.Threading.CancellationTokenSource = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(cancellation)
                timeout.CancelAfter(System.TimeSpan.FromSeconds(10))
                Using res As System.Net.Http.HttpResponseMessage = Await HelperClient.Http.GetAsync(root & "/.well-known/redink-helper", timeout.Token)
                    If res.IsSuccessStatusCode Then
                        Using doc As System.Text.Json.JsonDocument = System.Text.Json.JsonDocument.Parse(Await res.Content.ReadAsStringAsync(timeout.Token))
                            Dim signIn, pairing As System.Text.Json.JsonElement
                            offer.MailPairing = doc.RootElement.TryGetProperty("signIn", signIn) AndAlso signIn.TryGetProperty("mailPairing", pairing) AndAlso HelperClient.IsTrue(pairing, "enabled")
                        End Using
                    End If
                End Using
            End Using
        Catch ex As System.Exception
            ClientLog.Write("signin.offer.discovery.unavailable", ex)
        End Try
        ' 2.0.3: the Microsoft check runs even when the discovery document is missing (separate request).
        Try
            Using timeout As System.Threading.CancellationTokenSource = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(cancellation)
                timeout.CancelAfter(System.TimeSpan.FromSeconds(10))
                Using res As System.Net.Http.HttpResponseMessage = Await HelperClient.Http.GetAsync(root & "/auth/entra/status", timeout.Token)
                    If res.IsSuccessStatusCode Then
                        Using doc As System.Text.Json.JsonDocument = System.Text.Json.JsonDocument.Parse(Await res.Content.ReadAsStringAsync(timeout.Token))
                            Dim desktop As System.Text.Json.JsonElement
                            offer.Microsoft = HelperClient.IsTrue(doc.RootElement, "enabled") AndAlso doc.RootElement.TryGetProperty("desktopSso", desktop) AndAlso desktop.ValueKind = System.Text.Json.JsonValueKind.Number
                        End Using
                    End If
                End Using
            End Using
        Catch ex As System.Exception
            ClientLog.Write("signin.offer.unavailable", ex)
        End Try
        Return offer
    End Function

    Private Shared Async Function PostJsonAsync(url As String, payload As Object, cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of System.Tuple(Of Integer, String))
        Using req As New System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Post, url)
            req.Headers.Accept.Add(New System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"))
            req.Content = New System.Net.Http.StringContent(System.Text.Json.JsonSerializer.Serialize(payload), System.Text.Encoding.UTF8, "application/json")
            Using res As System.Net.Http.HttpResponseMessage = Await HelperClient.Http.SendAsync(req, cancellation)
                Return System.Tuple.Create(CInt(res.StatusCode), Await res.Content.ReadAsStringAsync(cancellation))
            End Using
        End Using
    End Function

    Public Shared Async Function StartPairingAsync(helperUrl As String, email As String, language As String, cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of PairingStart)
        Dim root As String = HelperClient.NormalizeHelperUrl(helperUrl)
        Dim answer As System.Tuple(Of Integer, String) = Await PostJsonAsync(root & "/auth/office-pairing/start", New With {.email = email, .hostApp = "askinky", .language = language}, cancellation)
        If answer.Item1 = 429 Then Throw New HelperFailure("desktop.rateLimited", 429)
        If answer.Item1 = 404 Then Throw New HelperFailure("desktop.pairingUnavailable", 404)
        If answer.Item1 <> 202 Then Throw New HelperFailure("desktop.pairingFailed", answer.Item1)
        Using doc As System.Text.Json.JsonDocument = System.Text.Json.JsonDocument.Parse(answer.Item2)
            Dim result As New PairingStart With {.Handle = HelperClient.TextValue(doc.RootElement, "handle"), .Code = HelperClient.TextValue(doc.RootElement, "code")}
            Dim value As System.Text.Json.JsonElement
            If doc.RootElement.TryGetProperty("pollAfterSeconds", value) AndAlso value.ValueKind = System.Text.Json.JsonValueKind.Number Then result.PollAfterSeconds = System.Math.Max(2, System.Math.Min(30, value.GetInt32()))
            If doc.RootElement.TryGetProperty("expiresIn", value) AndAlso value.ValueKind = System.Text.Json.JsonValueKind.Number Then result.ExpiresIn = System.Math.Max(60, System.Math.Min(3600, value.GetInt32()))
            If result.Handle.Length = 0 OrElse result.Code.Length = 0 Then Throw New HelperFailure("desktop.invalidResponse")
            Return result
        End Using
    End Function

    ' Polls until the person confirmed the mail (or it expired). Returns the configuration with credentials.
    Public Shared Async Function CompletePairingAsync(helperUrl As String, pairing As PairingStart, current As ClientConfig, cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of ClientConfig)
        Dim root As String = HelperClient.NormalizeHelperUrl(helperUrl)
        Dim deadline As System.DateTime = System.DateTime.UtcNow.AddSeconds(pairing.ExpiresIn + 30)
        Do While System.DateTime.UtcNow < deadline
            Await System.Threading.Tasks.Task.Delay(pairing.PollAfterSeconds * 1000, cancellation)
            Dim answer As System.Tuple(Of Integer, String) = Await PostJsonAsync(root & "/auth/office-pairing/poll", New With {.handle = pairing.Handle}, cancellation)
            If answer.Item1 = 202 Then Continue Do
            If answer.Item1 = 402 Then Throw New HelperFailure("desktop.licenseRequired", 402)
            If answer.Item1 = 410 Then
                Throw New HelperFailure(If(answer.Item2.Contains("DECLINED"), "desktop.pairingDeclined", "desktop.pairingExpired"), 410)
            End If
            If answer.Item1 = 429 Then Continue Do
            If answer.Item1 <> 200 Then Throw New HelperFailure("desktop.pairingFailed", answer.Item1)
            Return ApplyIssuedDevice(answer.Item2, root, current)
        Loop
        Throw New HelperFailure("desktop.pairingExpired")
    End Function

    Friend Shared Function ApplyIssuedDevice(json As String, root As String, current As ClientConfig) As ClientConfig
        Using doc As System.Text.Json.JsonDocument = System.Text.Json.JsonDocument.Parse(json)
            Dim payload As System.Text.Json.JsonElement = doc.RootElement
            Dim user As String = HelperClient.TextValue(payload, "userId").Trim()
            Dim token As String = HelperClient.TextValue(payload, "deviceToken").Trim()
            If HelperClient.TextValue(payload, "status") <> "ok" OrElse Not HelperClient.IsTrue(payload, "registered") OrElse user.Length = 0 OrElse token.Length = 0 Then Throw New HelperFailure("desktop.registrationFailed")
            Dim result As ClientConfig = ConfigStore.Clone(current)
            Dim origin As String = New System.Uri(root).GetLeftPart(System.UriPartial.Authority)
            If Not String.Equals(current.CredentialOrigin, origin, System.StringComparison.OrdinalIgnoreCase) Then result.ApiKeyProtected = ""
            result.HelperUrl = root
            result.CredentialOrigin = origin
            result.UserId = user
            result.DeviceTokenProtected = ConfigStore.Protect(token)
            Dim apiKey As String = HelperClient.TextValue(payload, "apiKey").Trim()
            If apiKey.Length > 0 Then result.ApiKeyProtected = ConfigStore.Protect(apiKey)
            If String.IsNullOrEmpty(result.LicenseUser) Then result.LicenseUser = user
            Return result
        End Using
    End Function

    ' ---------------------------------------------------------------- Microsoft sign-in (RFC 8252)

    ' Opens the system browser, receives the one-time code on a loopback port and exchanges it with the
    ' PKCE verifier. The browser never sees the device token; the code is useless without the verifier.
    Public Shared Async Function SignInWithMicrosoftAsync(helperUrl As String, current As ClientConfig, cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of ClientConfig)
        Dim root As String = HelperClient.NormalizeHelperUrl(helperUrl)
        Dim verifier As String = Base64Url(System.Security.Cryptography.RandomNumberGenerator.GetBytes(48))
        Dim challenge As String = Base64Url(System.Security.Cryptography.SHA256.HashData(System.Text.Encoding.ASCII.GetBytes(verifier)))
        Dim state As String = Base64Url(System.Security.Cryptography.RandomNumberGenerator.GetBytes(24))
        Dim listener As New System.Net.Sockets.TcpListener(System.Net.IPAddress.Loopback, 0)
        listener.Start()
        Try
            Dim port As Integer = DirectCast(listener.LocalEndpoint, System.Net.IPEndPoint).Port
            Dim redirect As String = "http://127.0.0.1:" & port.ToString(System.Globalization.CultureInfo.InvariantCulture) & "/askinky/callback"
            Dim start As String = root & "/auth/entra/start?purpose=desktop&lang=" & System.Uri.EscapeDataString(current.UiLanguage) &
                "&desktop_redirect=" & System.Uri.EscapeDataString(redirect) & "&desktop_challenge=" & challenge & "&desktop_state=" & state
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(start) With {.UseShellExecute = True})
            Using timeout As System.Threading.CancellationTokenSource = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(cancellation)
                timeout.CancelAfter(System.TimeSpan.FromMinutes(5))
                Do
                    Using client As System.Net.Sockets.TcpClient = Await listener.AcceptTcpClientAsync(timeout.Token)
                        Dim stream As System.Net.Sockets.NetworkStream = client.GetStream()
                        Dim buffer(16383) As Byte
                        Dim read As Integer = Await stream.ReadAsync(buffer, 0, buffer.Length, timeout.Token)
                        Dim requestLine As String = System.Text.Encoding.ASCII.GetString(buffer, 0, read).Split({vbCrLf}, System.StringSplitOptions.None)(0)
                        Dim parts As String() = requestLine.Split(" "c)
                        If parts.Length < 2 OrElse parts(0) <> "GET" OrElse Not parts(1).StartsWith("/askinky/callback?", System.StringComparison.Ordinal) Then
                            Await WriteHtmlAsync(stream, 404, "<!doctype html><title>404</title>")
                            Continue Do
                        End If
                        Dim query As New System.Collections.Generic.Dictionary(Of String, String)(System.StringComparer.Ordinal)
                        For Each pair As String In parts(1).Substring(parts(1).IndexOf("?"c) + 1).Split("&"c)
                            Dim kv As String() = pair.Split("="c, 2)
                            If kv.Length = 2 Then query(System.Uri.UnescapeDataString(kv(0))) = System.Uri.UnescapeDataString(kv(1).Replace("+", " "))
                        Next
                        Dim received As String = ""
                        If Not query.TryGetValue("state", received) OrElse Not System.Security.Cryptography.CryptographicOperations.FixedTimeEquals(System.Text.Encoding.ASCII.GetBytes(received), System.Text.Encoding.ASCII.GetBytes(state)) OrElse Not query.ContainsKey("code") Then
                            Await WriteHtmlAsync(stream, 400, RedInkPage.SignInResult(current.UiLanguage, False))
                            Throw New HelperFailure("desktop.microsoftFailed")
                        End If
                        Await WriteHtmlAsync(stream, 200, RedInkPage.SignInResult(current.UiLanguage, True))
                        Dim answer As System.Tuple(Of Integer, String) = Await PostJsonAsync(root & "/auth/entra/desktop-exchange", New With {.code = query("code"), .code_verifier = verifier}, cancellation)
                        If answer.Item1 = 402 Then Throw New HelperFailure("desktop.licenseRequired", 402)
                        If answer.Item1 = 403 Then Throw New HelperFailure("desktop.microsoftNotLinked", 403)
                        If answer.Item1 <> 200 Then Throw New HelperFailure("desktop.microsoftFailed", answer.Item1)
                        Return ApplyIssuedDevice(answer.Item2, root, current)
                    End Using
                Loop
            End Using
        Finally
            listener.Stop()
        End Try
    End Function

    Private Shared Function Base64Url(bytes As Byte()) As String
        Return System.Convert.ToBase64String(bytes).TrimEnd("="c).Replace("+"c, "-"c).Replace("/"c, "_"c)
    End Function

    ' 2.0.4: html is a complete page (RedInkPage, Red Ink design); no script, inline style and data-URL logo only.
    Private Shared Async Function WriteHtmlAsync(stream As System.Net.Sockets.NetworkStream, status As Integer, html As String) As System.Threading.Tasks.Task
        Dim body As Byte() = System.Text.Encoding.UTF8.GetBytes(html)
        Dim head As String = "HTTP/1.1 " & status.ToString(System.Globalization.CultureInfo.InvariantCulture) & If(status = 200, " OK", " Error") & vbCrLf &
            "Content-Type: text/html; charset=utf-8" & vbCrLf & "Cache-Control: no-store" & vbCrLf & "Referrer-Policy: no-referrer" & vbCrLf &
            "Content-Security-Policy: default-src 'none'; img-src data:; style-src 'unsafe-inline'; base-uri 'none'; form-action 'none'; frame-ancestors 'none'" & vbCrLf & "X-Content-Type-Options: nosniff" & vbCrLf &
            "Content-Length: " & body.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) & vbCrLf & "Connection: close" & vbCrLf & vbCrLf
        Dim headBytes As Byte() = System.Text.Encoding.ASCII.GetBytes(head)
        Await stream.WriteAsync(headBytes, 0, headBytes.Length)
        Await stream.WriteAsync(body, 0, body.Length)
        Await stream.FlushAsync()
    End Function

    ' ---------------------------------------------------------------- file drop

    ' Same route as the screenshot OCR (/files/upload), also present on older Helpers.
    Public Shared Async Function FileTextAsync(config As ClientConfig, fileName As String, bytes As Byte(), cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of String)
        If bytes Is Nothing OrElse bytes.Length = 0 OrElse bytes.Length > 25000000 Then Throw New HelperFailure("desktop.fileTooLarge")
        Using timeout As System.Threading.CancellationTokenSource = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(cancellation)
            timeout.CancelAfter(System.TimeSpan.FromMinutes(5))
            Using req As System.Net.Http.HttpRequestMessage = HelperClient.CreateRequest(config, System.Net.Http.HttpMethod.Post, "files/upload")
                Using form As New System.Net.Http.MultipartFormDataContent()
                    Dim content As New System.Net.Http.ByteArrayContent(bytes)
                    content.Headers.ContentType = New System.Net.Http.Headers.MediaTypeHeaderValue("application/octet-stream")
                    form.Add(content, "file", System.IO.Path.GetFileName(fileName))
                    form.Add(New System.Net.Http.StringContent("true"), "allow_ocr")
                    req.Content = form
                    Using res As System.Net.Http.HttpResponseMessage = Await HelperClient.Http.SendAsync(req, timeout.Token)
                        If Not res.IsSuccessStatusCode Then Throw HelperClient.StatusFailure(res.StatusCode, config)
                        Using doc As System.Text.Json.JsonDocument = System.Text.Json.JsonDocument.Parse(Await res.Content.ReadAsStringAsync(timeout.Token))
                            Dim success As System.Text.Json.JsonElement
                            If doc.RootElement.TryGetProperty("success", success) AndAlso success.ValueKind = System.Text.Json.JsonValueKind.False Then Throw New HelperFailure("desktop.fileUnreadable")
                            Dim text As String = HelperClient.TextValue(doc.RootElement, "content")
                            If String.IsNullOrWhiteSpace(text) Then Throw New HelperFailure("desktop.fileUnreadable")
                            Return text
                        End Using
                    End Using
                End Using
            End Using
        End Using
    End Function

    ' ---------------------------------------------------------------- self-update

    ' Only when the admin switched it on (404 otherwise, also on every older Helper).
    Public Shared Async Function CheckUpdateAsync(config As ClientConfig, cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of ClientUpdateInfo)
        Try
            Using timeout As System.Threading.CancellationTokenSource = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(cancellation)
                timeout.CancelAfter(System.TimeSpan.FromSeconds(20))
                Using req As System.Net.Http.HttpRequestMessage = HelperClient.CreateRequest(config, System.Net.Http.HttpMethod.Get, "askinky/update.json")
                    Using res As System.Net.Http.HttpResponseMessage = Await HelperClient.Http.SendAsync(req, timeout.Token)
                        If Not res.IsSuccessStatusCode Then Return Nothing
                        Using doc As System.Text.Json.JsonDocument = System.Text.Json.JsonDocument.Parse(Await res.Content.ReadAsStringAsync(timeout.Token))
                            Return ParseUpdate(doc.RootElement)
                        End Using
                    End Using
                End Using
            End Using
        Catch ex As System.Exception
            ClientLog.Write("update.check.failed", ex)
            Return Nothing
        End Try
    End Function

    Friend Shared Function ParseUpdate(root As System.Text.Json.JsonElement) As ClientUpdateInfo
        Dim info As New ClientUpdateInfo With {
            .Version = HelperClient.TextValue(root, "version"),
            .FileName = HelperClient.TextValue(root, "file"),
            .Sha256 = HelperClient.TextValue(root, "sha256").ToLowerInvariant(),
            .SignerOrganization = HelperClient.TextValue(root, "signerOrganization"),
            .DownloadPath = HelperClient.TextValue(root, "downloadPath")}
        If info.FileName.Length = 0 Then info.FileName = HelperClient.TextValue(root, "fileName")
        Dim size As System.Text.Json.JsonElement
        If root.TryGetProperty("size", size) AndAlso size.ValueKind = System.Text.Json.JsonValueKind.Number Then info.Size = size.GetInt64()
        If Not System.Text.RegularExpressions.Regex.IsMatch(info.Version, "^\d{1,4}\.\d{1,4}\.\d{1,6}$") OrElse
           Not System.Text.RegularExpressions.Regex.IsMatch(info.FileName, "^[A-Za-z0-9][A-Za-z0-9._-]{0,119}\.msi$") OrElse
           Not System.Text.RegularExpressions.Regex.IsMatch(info.Sha256, "^[0-9a-f]{64}$") OrElse info.Size <= 0 OrElse info.Size > 512L * 1024 * 1024 Then Return Nothing
        If info.DownloadPath.Length = 0 Then info.DownloadPath = "askinky/download/" & info.FileName
        If Not System.Text.RegularExpressions.Regex.IsMatch(info.DownloadPath, "^/?askinky/download/[A-Za-z0-9._-]{1,120}$") Then Return Nothing
        Return info
    End Function

    Public Shared Function IsNewer(candidate As String, installed As String) As Boolean
        Dim a As System.Version = Nothing
        Dim b As System.Version = Nothing
        Return System.Version.TryParse(candidate, a) AndAlso System.Version.TryParse(installed, b) AndAlso a > b
    End Function

    ' Downloads, checks size + SHA-256 + Authenticode signer, and returns the verified local path.
    Public Shared Async Function DownloadUpdateAsync(config As ClientConfig, info As ClientUpdateInfo, cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of String)
        Dim folder As String = System.IO.Path.Combine(StoragePaths.Root, "updates")
        System.IO.Directory.CreateDirectory(folder)
        Dim target As String = System.IO.Path.Combine(folder, info.FileName)
        Dim temporary As String = target & ".part"
        Using req As System.Net.Http.HttpRequestMessage = HelperClient.CreateRequest(config, System.Net.Http.HttpMethod.Get, info.DownloadPath.TrimStart("/"c))
            Using res As System.Net.Http.HttpResponseMessage = Await HelperClient.Http.SendAsync(req, System.Net.Http.HttpCompletionOption.ResponseHeadersRead, cancellation)
                If Not res.IsSuccessStatusCode Then Throw HelperClient.StatusFailure(res.StatusCode, config)
                Using input As System.IO.Stream = Await res.Content.ReadAsStreamAsync(cancellation)
                    Using output As New System.IO.FileStream(temporary, System.IO.FileMode.Create, System.IO.FileAccess.Write)
                        Dim buffer(81919) As Byte
                        Dim total As Long = 0
                        Do
                            Dim read As Integer = Await input.ReadAsync(buffer, 0, buffer.Length, cancellation)
                            If read = 0 Then Exit Do
                            total += read
                            If total > info.Size Then Throw New HelperFailure("desktop.updateInvalid")
                            Await output.WriteAsync(buffer, 0, read, cancellation)
                        Loop
                    End Using
                End Using
            End Using
        End Using
        Dim actual As String
        Using stream As System.IO.FileStream = System.IO.File.OpenRead(temporary)
            actual = System.Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(stream)).ToLowerInvariant()
        End Using
        If New System.IO.FileInfo(temporary).Length <> info.Size OrElse actual <> info.Sha256 Then
            System.IO.File.Delete(temporary)
            Throw New HelperFailure("desktop.updateInvalid")
        End If
        Dim signer As String = AuthenticodeVerifier.TrustedSignerOrganization(temporary)
        If signer.Length = 0 OrElse Not AllowedSigners.Contains(signer, System.StringComparer.Ordinal) OrElse
           (info.SignerOrganization.Length > 0 AndAlso Not String.Equals(info.SignerOrganization, signer, System.StringComparison.Ordinal)) Then
            System.IO.File.Delete(temporary)
            ClientLog.Write("update.signature.refused")
            Throw New HelperFailure("desktop.updateUnsigned")
        End If
        System.IO.File.Move(temporary, target, True)
        Return target
    End Function

    Public Shared Sub StartInstaller(msiPath As String)
        ' MajorUpgrade in the MSI replaces the running version; Ask Inky exits right after.
        Dim info As New System.Diagnostics.ProcessStartInfo("msiexec.exe") With {.UseShellExecute = False}
        For Each argument As String In {"/i", msiPath, "/passive", "AUTOSTARTAFTERINSTALL=1"}
            info.ArgumentList.Add(argument)
        Next
        System.Diagnostics.Process.Start(info)
    End Sub
End Class

' WinVerifyTrust (full chain, revocation by the system policy) plus the O= of the signing certificate.
Friend NotInheritable Class AuthenticodeVerifier
    <System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential, CharSet:=System.Runtime.InteropServices.CharSet.Unicode)>
    Private Structure WINTRUST_FILE_INFO
        Public cbStruct As Integer
        Public pcwszFilePath As String
        Public hFile As System.IntPtr
        Public pgKnownSubject As System.IntPtr
    End Structure

    <System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)>
    Private Structure WINTRUST_DATA
        Public cbStruct As Integer
        Public pPolicyCallbackData As System.IntPtr
        Public pSIPClientData As System.IntPtr
        Public dwUIChoice As Integer
        Public fdwRevocationChecks As Integer
        Public dwUnionChoice As Integer
        Public pFile As System.IntPtr
        Public dwStateAction As Integer
        Public hWVTStateData As System.IntPtr
        Public pwszURLReference As System.IntPtr
        Public dwProvFlags As Integer
        Public dwUIContext As Integer
        Public pSignatureSettings As System.IntPtr
    End Structure

    <System.Runtime.InteropServices.DllImport("wintrust.dll", CharSet:=System.Runtime.InteropServices.CharSet.Unicode)>
    Private Shared Function WinVerifyTrust(hwnd As System.IntPtr, <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPStruct)> action As System.Guid, ByRef data As WINTRUST_DATA) As Integer
    End Function

    Private Shared ReadOnly GenericVerifyV2 As New System.Guid("00AAC56B-CD44-11d0-8CC2-00C04FC295EE")

    Public Shared Function TrustedSignerOrganization(path As String) As String
        Dim fileInfo As New WINTRUST_FILE_INFO With {.cbStruct = System.Runtime.InteropServices.Marshal.SizeOf(Of WINTRUST_FILE_INFO)(), .pcwszFilePath = path}
        Dim filePtr As System.IntPtr = System.Runtime.InteropServices.Marshal.AllocHGlobal(fileInfo.cbStruct)
        Try
            System.Runtime.InteropServices.Marshal.StructureToPtr(fileInfo, filePtr, False)
            ' UI none, revocation whole chain, choice file, verify, revocation check by cache only offline-tolerant.
            Dim data As New WINTRUST_DATA With {.cbStruct = System.Runtime.InteropServices.Marshal.SizeOf(Of WINTRUST_DATA)(), .dwUIChoice = 2, .fdwRevocationChecks = 1, .dwUnionChoice = 1, .pFile = filePtr, .dwStateAction = 0, .dwProvFlags = &H1000}
            Dim status As Integer = WinVerifyTrust(New System.IntPtr(-1), GenericVerifyV2, data)
            If status <> 0 Then
                ClientLog.Write("update.authenticode.status=0x" & status.ToString("X8", System.Globalization.CultureInfo.InvariantCulture))
                Return ""
            End If
        Finally
            System.Runtime.InteropServices.Marshal.FreeHGlobal(filePtr)
        End Try
        Try
            Using certificate As New System.Security.Cryptography.X509Certificates.X509Certificate2(System.Security.Cryptography.X509Certificates.X509Certificate.CreateFromSignedFile(path))
                Return OrganizationOf(certificate)
            End Using
        Catch ex As System.Exception
            ClientLog.Write("update.authenticode.certificate.unreadable", ex)
            Return ""
        End Try
    End Function

    Private Shared Function OrganizationOf(certificate As System.Security.Cryptography.X509Certificates.X509Certificate2) As String
        For Each part As String In certificate.SubjectName.Decode(System.Security.Cryptography.X509Certificates.X500DistinguishedNameFlags.UseNewLines).Split({vbCr, vbLf}, System.StringSplitOptions.RemoveEmptyEntries)
            Dim item As String = part.Trim()
            If item.StartsWith("O=", System.StringComparison.Ordinal) Then Return item.Substring(2).Trim(""""c)
        Next
        Return ""
    End Function
End Class
