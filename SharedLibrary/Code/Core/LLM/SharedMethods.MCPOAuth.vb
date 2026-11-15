' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SharedMethods.MCPOAuth.vb
' Purpose: Provides OAuth helpers for MCP-protected resources, including bearer
'          challenge discovery, authorization-server metadata lookup, dynamic
'          client registration, device flow, authorization-code flow, and token
'          caching for MCP tool execution.
'
' Architecture:
'  - Discovery:
'      - Probes protected resources for `WWW-Authenticate` bearer challenges.
'      - Resolves resource metadata and authorization server metadata.
'  - Client setup:
'      - Reuses configured client credentials when available.
'      - Falls back to dynamic client registration when required by the server.
'  - Authentication flows:
'      - Uses device authorization flow when advertised.
'      - Otherwise runs interactive authorization-code flow.
'      - Requires HTTPS for all MCP OAuth server endpoints; loopback redirect URIs remain exempt.
'  - Token lifecycle:
'      - Caches acquired tokens by client and endpoint configuration.
'      - Exposes helpers for lookup, storage, and debug logging.
'
' Types:
'  - `MCPProtectedResourceOAuthResult`
'  - `MCPCachedToken`
' =============================================================================


Option Strict On
Option Explicit On

Imports System.Diagnostics
Imports System.Net
Imports System.Net.Http
Imports System.Net.Sockets
Imports System.Security.Cryptography
Imports System.Text
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Collections.Generic
Imports System.Linq
Imports System.Web

Namespace SharedLibrary

    Partial Public Class SharedMethods

        Public Class MCPProtectedResourceOAuthResult
            Public Property Issuer As System.String = ""
            Public Property RegistrationEndpoint As System.String = ""
            Public Property RedirectUri As System.String = ""
            Public Property RefreshToken As System.String = ""
            Public Property SupportsRefresh As System.Boolean = False
            Public Property AccessToken As String = ""
            Public Property ClientId As String = ""
            Public Property ClientSecret As String = ""
            Public Property AuthorizationEndpoint As String = ""
            Public Property TokenEndpoint As String = ""
            Public Property DeviceAuthorizationEndpoint As String = ""
            Public Property Scope As String = ""
            Public Property Resource As String = ""
            Public Property ExpiresIn As Integer = 3600
            Public Property UseDeviceFlow As Boolean = False
        End Class

        Public Class MCPCachedToken
            Public Property AccessToken As String = ""
            Public Property ExpiryUtc As DateTime = DateTime.MinValue
        End Class

        Private Shared _mcpTokenCache As New Dictionary(Of String, MCPCachedToken)(StringComparer.OrdinalIgnoreCase)
        Private Shared ReadOnly _mcpTokenCacheLock As New Object()
        Private Shared _mcpOAuthDebugEnabled As Boolean = False

        Public Shared Sub SetMCPOAuthDebugEnabled(enabled As Boolean)
            _mcpOAuthDebugEnabled = enabled
        End Sub

        Public Shared Function BuildMCPTokenCacheKey(clientId As String, oauthEndpointConfig As String) As String
            Return (If(clientId, "") & "|" & If(oauthEndpointConfig, "")).ToLowerInvariant()
        End Function

        Private Shared Function RequireMcpOAuthHttpsEndpoint(endpointUrl As String, endpointName As String) As String
            Dim trimmedEndpoint As String = If(endpointUrl, String.Empty).Trim()
            Dim parsedEndpoint As System.Uri = Nothing

            If String.IsNullOrWhiteSpace(trimmedEndpoint) OrElse
               Not System.Uri.TryCreate(trimmedEndpoint, System.UriKind.Absolute, parsedEndpoint) OrElse
               Not String.Equals(parsedEndpoint.Scheme, System.Uri.UriSchemeHttps, System.StringComparison.OrdinalIgnoreCase) Then
                Throw New System.Exception("MCP OAuth " & endpointName & " must use an absolute https:// URL.")
            End If

            Return trimmedEndpoint
        End Function

        Public Shared Function GetCachedMCPToken(cacheKey As String) As MCPCachedToken
            SyncLock _mcpTokenCacheLock
                Dim entry As MCPCachedToken = Nothing
                If _mcpTokenCache.TryGetValue(cacheKey, entry) Then
                    Return entry
                End If
            End SyncLock
            Return Nothing
        End Function

        Public Shared Sub StoreCachedMCPToken(cacheKey As String, accessToken As String, expiryUtc As DateTime)
            SyncLock _mcpTokenCacheLock
                _mcpTokenCache(cacheKey) = New MCPCachedToken With {
                    .AccessToken = accessToken,
                    .ExpiryUtc = expiryUtc
                }
            End SyncLock
        End Sub

        ' Automatic MCP OAuth keeps personal credentials outside distributable INI files.
        Private Class MCPAutomaticOAuthState
            Public Property Version As System.Int32 = 1
            Public Property Endpoint As System.String = ""
            Public Property ConfiguredScope As System.String = ""
            Public Property ExpiryUtc As System.DateTime = System.DateTime.MinValue
            Public Property Connection As MCPProtectedResourceOAuthResult
        End Class

        Public Shared Function IsAutomaticMCPOAuth(endpointConfig As System.String) As System.Boolean
            Return System.String.Equals(If(endpointConfig, "").Trim(), "auto", System.StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function AutomaticMCPResource(endpoint As System.String) As System.String
            Dim value As System.String = If(endpoint, "").Trim()
            For Each prefix As System.String In New System.String() {MCP_SSE_PREFIX, MCP_STREAMABLE_PREFIX}
                If value.StartsWith(prefix, System.StringComparison.OrdinalIgnoreCase) Then
                    value = value.Substring(prefix.Length)
                    Exit For
                End If
            Next
            value = RequireMcpOAuthHttpsEndpoint(value, "protected resource")
            Dim uri As New System.Uri(value)
            If uri.UserInfo.Length > 0 OrElse uri.Fragment.Length > 0 Then
                Throw New System.ArgumentException("Use an MCP endpoint without user information or a fragment.")
            End If
            Return uri.AbsoluteUri
        End Function

        Private Shared Function AutomaticMCPCachePath(endpoint As System.String, scope As System.String) As System.String
            Dim identity As System.String = endpoint & Microsoft.VisualBasic.vbLf & scope
            Dim digest As System.String
            Using sha As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                digest = System.BitConverter.ToString(sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(identity))).Replace("-", "").ToLowerInvariant()
            End Using
            Dim directory As System.String = System.IO.Path.Combine(
                System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData), "RedInk", "McpOAuth")
            System.IO.Directory.CreateDirectory(directory)
            Return System.IO.Path.Combine(directory, digest & ".bin")
        End Function

        Private Shared Async Function LockAutomaticMCPCacheAsync(path As System.String, cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of System.IO.FileStream)
            Dim deadline As System.DateTime = System.DateTime.UtcNow.AddMinutes(10)
            Do
                cancellationToken.ThrowIfCancellationRequested()
                Try
                    Return New System.IO.FileStream(path & ".lock", System.IO.FileMode.OpenOrCreate,
                        System.IO.FileAccess.ReadWrite, System.IO.FileShare.None)
                Catch ex As System.IO.IOException
                    If System.DateTime.UtcNow >= deadline Then
                        Throw New System.TimeoutException("Another Office process is still using this MCP connection. Finish its sign-in and retry.")
                    End If
                End Try
                Await System.Threading.Tasks.Task.Delay(200, cancellationToken).ConfigureAwait(False)
            Loop
        End Function

        Private Shared Function ReadAutomaticMCPState(path As System.String, endpoint As System.String, scope As System.String) As MCPAutomaticOAuthState
            If Not System.IO.File.Exists(path) Then Return Nothing
            Try
                Dim plain As System.Byte() = System.Security.Cryptography.ProtectedData.Unprotect(
                    System.IO.File.ReadAllBytes(path), System.Text.Encoding.UTF8.GetBytes("Red Ink MCP OAuth v1"),
                    System.Security.Cryptography.DataProtectionScope.CurrentUser)
                Dim state As MCPAutomaticOAuthState = Newtonsoft.Json.JsonConvert.DeserializeObject(Of MCPAutomaticOAuthState)(System.Text.Encoding.UTF8.GetString(plain))
                If state Is Nothing OrElse state.Version <> 1 OrElse state.Connection Is Nothing OrElse
                   Not System.String.Equals(state.Endpoint, endpoint, System.StringComparison.Ordinal) OrElse
                   Not System.String.Equals(state.ConfiguredScope, scope, System.StringComparison.Ordinal) Then
                    Throw New System.IO.InvalidDataException("The MCP connection cache does not match this configuration.")
                End If
                Return state
            Catch ex As System.Exception
                ' Do not silently replace a damaged or unreadable personal connection.
                Throw New System.InvalidOperationException("Red Ink could not read the personal MCP connection. Open the MCP importer, select Automatic OAuth and Forget saved OAuth connection, then reconnect.", ex)
            End Try
        End Function

        Private Shared Sub WriteAutomaticMCPState(path As System.String, state As MCPAutomaticOAuthState)
            Dim plain As System.Byte() = System.Text.Encoding.UTF8.GetBytes(Newtonsoft.Json.JsonConvert.SerializeObject(state))
            Dim protectedBytes As System.Byte() = System.Security.Cryptography.ProtectedData.Protect(plain,
                System.Text.Encoding.UTF8.GetBytes("Red Ink MCP OAuth v1"), System.Security.Cryptography.DataProtectionScope.CurrentUser)
            Dim temporary As System.String = path & ".tmp-" & System.Guid.NewGuid().ToString("N")
            Try
                System.IO.File.WriteAllBytes(temporary, protectedBytes)
                If System.IO.File.Exists(path) Then
                    System.IO.File.Replace(temporary, path, Nothing)
                Else
                    System.IO.File.Move(temporary, path)
                End If
            Finally
                If System.IO.File.Exists(temporary) Then System.IO.File.Delete(temporary)
            End Try
        End Sub

        Private Shared Function NewAutomaticMCPHttpClient() As System.Net.Http.HttpClient
            Dim handler As New System.Net.Http.HttpClientHandler() With {
                .AllowAutoRedirect = False,
                .AutomaticDecompression = System.Net.DecompressionMethods.GZip Or System.Net.DecompressionMethods.Deflate
            }
            handler.UseProxy = True
            handler.Proxy = System.Net.WebRequest.DefaultWebProxy
            If handler.Proxy IsNot Nothing Then handler.Proxy.Credentials = System.Net.CredentialCache.DefaultCredentials
            Dim client As New System.Net.Http.HttpClient(handler) With {.Timeout = System.TimeSpan.FromSeconds(30)}
            client.DefaultRequestHeaders.Accept.ParseAdd("application/json")
            Return client
        End Function

        Private Shared Async Function GetAutomaticMCPJsonAsync(url As System.String, cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of Newtonsoft.Json.Linq.JObject)
            url = RequireMcpOAuthHttpsEndpoint(url, "metadata endpoint")
            Using client As System.Net.Http.HttpClient = NewAutomaticMCPHttpClient()
                Using response As System.Net.Http.HttpResponseMessage = Await client.GetAsync(url, cancellationToken).ConfigureAwait(False)
                    response.EnsureSuccessStatusCode()
                    Return Newtonsoft.Json.Linq.JObject.Parse(Await response.Content.ReadAsStringAsync().ConfigureAwait(False))
                End Using
            End Using
        End Function

        Private Shared Async Function DiscoverAutomaticMCPAsync(endpoint As System.String, configuredScope As System.String, cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of MCPProtectedResourceOAuthResult)
            ' Some Streamable HTTP servers only challenge POST and answer GET with 405.
            Dim challenge As System.Collections.Generic.Dictionary(Of System.String, System.String) = Await GetBearerChallengeAsync(endpoint).ConfigureAwait(False)
            Dim metadata As Newtonsoft.Json.Linq.JObject = Nothing
            Dim advertisedMetadata As System.String = GetChallengeValue(challenge, "resource_metadata")
            If Not System.String.IsNullOrWhiteSpace(advertisedMetadata) Then
                metadata = Await GetAutomaticMCPJsonAsync(advertisedMetadata, cancellationToken).ConfigureAwait(False)
            Else
                For Each candidate As System.String In BuildProtectedResourceMetadataCandidates(endpoint)
                    Try
                        metadata = Await GetAutomaticMCPJsonAsync(candidate, cancellationToken).ConfigureAwait(False)
                    Catch ex As System.Net.Http.HttpRequestException
                        LogMCPOAuth("Automatic resource metadata candidate unavailable", candidate)
                    End Try
                    If metadata IsNot Nothing Then Exit For
                Next
            End If
            If metadata Is Nothing Then Throw New System.InvalidOperationException("The MCP server did not publish usable OAuth protected-resource metadata.")
            Dim resource As System.String = RequireMcpOAuthHttpsEndpoint(If(metadata("resource")?.ToString(), ""), "resource")
            If Not System.String.Equals(New System.Uri(resource).AbsoluteUri, endpoint, System.StringComparison.Ordinal) Then
                Throw New System.InvalidOperationException("The OAuth resource metadata identifies a different MCP endpoint.")
            End If
            Dim servers As Newtonsoft.Json.Linq.JArray = TryCast(metadata("authorization_servers"), Newtonsoft.Json.Linq.JArray)
            If servers Is Nothing OrElse servers.Count = 0 Then Throw New System.InvalidOperationException("The MCP server did not advertise an OAuth authorization server.")
            Dim issuer As System.String = RequireMcpOAuthHttpsEndpoint(servers(0).ToString(), "issuer")
            Dim issuerUri As New System.Uri(issuer)
            If issuerUri.Query.Length > 0 OrElse issuerUri.Fragment.Length > 0 OrElse issuerUri.UserInfo.Length > 0 Then
                Throw New System.InvalidOperationException("The OAuth issuer must not contain credentials, a query or a fragment.")
            End If
            Dim origin As System.String = issuerUri.GetLeftPart(System.UriPartial.Authority)
            Dim path As System.String = issuerUri.AbsolutePath.TrimEnd("/"c)
            Dim serverMetadata As Newtonsoft.Json.Linq.JObject = Nothing
            Dim candidates As New System.Collections.Generic.List(Of System.String) From {
                origin & "/.well-known/oauth-authorization-server" & path,
                issuer.TrimEnd("/"c) & "/.well-known/openid-configuration"
            }
            For Each candidate As System.String In candidates.Distinct(System.StringComparer.Ordinal)
                Try
                    serverMetadata = Await GetAutomaticMCPJsonAsync(candidate, cancellationToken).ConfigureAwait(False)
                Catch ex As System.Net.Http.HttpRequestException
                    LogMCPOAuth("Automatic authorization metadata candidate unavailable", candidate)
                End Try
                If serverMetadata IsNot Nothing Then Exit For
            Next
            If serverMetadata Is Nothing OrElse Not System.String.Equals(If(serverMetadata("issuer")?.ToString(), ""), issuer, System.StringComparison.Ordinal) Then
                Throw New System.InvalidOperationException("The OAuth authorization metadata does not match the advertised issuer.")
            End If
            Dim scope As System.String = configuredScope
            If scope.Length = 0 Then scope = GetChallengeValue(challenge, "scope")
            If scope.Length = 0 Then
                Dim scopes As Newtonsoft.Json.Linq.JArray = TryCast(metadata("scopes_supported"), Newtonsoft.Json.Linq.JArray)
                If scopes IsNot Nothing Then scope = System.String.Join(" ", scopes.Select(Function(item) item.ToString()))
            End If
            Dim result As New MCPProtectedResourceOAuthResult With {
                .Issuer = issuer, .Resource = resource, .Scope = scope,
                .AuthorizationEndpoint = RequireMcpOAuthHttpsEndpoint(If(serverMetadata("authorization_endpoint")?.ToString(), ""), "authorization_endpoint"),
                .TokenEndpoint = RequireMcpOAuthHttpsEndpoint(If(serverMetadata("token_endpoint")?.ToString(), ""), "token_endpoint"),
                .RegistrationEndpoint = If(serverMetadata("registration_endpoint")?.ToString(), "")
            }
            Dim pkce As Newtonsoft.Json.Linq.JArray = TryCast(serverMetadata("code_challenge_methods_supported"), Newtonsoft.Json.Linq.JArray)
            If pkce Is Nothing OrElse Not pkce.Any(Function(item) item.ToString() = "S256") Then
                Throw New System.InvalidOperationException("The OAuth server must advertise PKCE S256 for automatic MCP sign-in.")
            End If
            Dim grants As Newtonsoft.Json.Linq.JArray = TryCast(serverMetadata("grant_types_supported"), Newtonsoft.Json.Linq.JArray)
            If grants IsNot Nothing AndAlso Not grants.Any(Function(item) item.ToString() = "authorization_code") Then
                Throw New System.InvalidOperationException("The OAuth server does not support the authorization-code flow required by automatic MCP sign-in.")
            End If
            result.SupportsRefresh = grants Is Nothing OrElse grants.Any(Function(item) item.ToString() = "refresh_token")
            Return result
        End Function

        Private Shared Async Function RegisterAutomaticMCPAsync(connection As MCPProtectedResourceOAuthResult, cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            Dim endpoint As System.String = RequireMcpOAuthHttpsEndpoint(connection.RegistrationEndpoint, "registration_endpoint")
            ' Register only the grants used by this native client, never an unadvertised device grant.
            Dim grants As New Newtonsoft.Json.Linq.JArray("authorization_code")
            If connection.SupportsRefresh Then grants.Add("refresh_token")
            For Each redirect As System.String In New System.String() {"http://127.0.0.1/callback", "https://localhost/callback"}
                Dim payload As New Newtonsoft.Json.Linq.JObject From {
                    {"client_name", "Red Ink"}, {"application_type", "native"},
                    {"redirect_uris", New Newtonsoft.Json.Linq.JArray(redirect)},
                    {"grant_types", grants.DeepClone()}, {"response_types", New Newtonsoft.Json.Linq.JArray("code")},
                    {"token_endpoint_auth_method", "none"}
                }
                Using client As System.Net.Http.HttpClient = NewAutomaticMCPHttpClient()
                    Using content As New System.Net.Http.StringContent(payload.ToString(Newtonsoft.Json.Formatting.None), System.Text.Encoding.UTF8, "application/json")
                        Using response As System.Net.Http.HttpResponseMessage = Await client.PostAsync(endpoint, content, cancellationToken).ConfigureAwait(False)
                            Dim json As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(Await response.Content.ReadAsStringAsync().ConfigureAwait(False))
                            If Not response.IsSuccessStatusCode Then
                                Dim errorCode As System.String = If(json("error")?.ToString(), "")
                                If response.StatusCode = System.Net.HttpStatusCode.BadRequest AndAlso
                                   (errorCode = "invalid_redirect_uri" OrElse errorCode = "invalid_client_metadata") AndAlso redirect.StartsWith("http:", System.StringComparison.Ordinal) Then
                                    LogMCPOAuth("Automatic registration redirect rejected", "Retrying with HTTPS loopback.")
                                    Continue For
                                End If
                                Throw New System.InvalidOperationException("MCP OAuth client registration failed (HTTP " & CInt(response.StatusCode).ToString() & ").")
                            End If
                            connection.ClientId = If(json("client_id")?.ToString(), "")
                            connection.ClientSecret = ""
                            connection.RedirectUri = redirect
                            If System.String.IsNullOrWhiteSpace(connection.ClientId) Then Throw New System.InvalidOperationException("OAuth registration returned no client ID.")
                            Dim method As System.String = If(json("token_endpoint_auth_method")?.ToString(), "none")
                            If method <> "none" Then Throw New System.InvalidOperationException("The OAuth registration did not accept public-client authentication.")
                            Dim redirects As Newtonsoft.Json.Linq.JArray = TryCast(json("redirect_uris"), Newtonsoft.Json.Linq.JArray)
                            If redirects IsNot Nothing AndAlso Not redirects.Any(Function(item) item.ToString() = redirect) Then
                                Throw New System.InvalidOperationException("OAuth registration did not accept the requested callback address.")
                            End If
                            Return
                        End Using
                    End Using
                End Using
            Next
            Throw New System.InvalidOperationException("The OAuth server rejected the native client's callback addresses.")
        End Function

        Private Shared Sub ApplyAutomaticMCPToken(connection As MCPProtectedResourceOAuthResult, json As Newtonsoft.Json.Linq.JObject)
            Dim token As System.String = If(json("access_token")?.ToString(), "")
            Dim tokenType As System.String = If(json("token_type")?.ToString(), "")
            Dim seconds As System.Int32
            If System.String.IsNullOrWhiteSpace(token) OrElse token.Any(AddressOf System.Char.IsControl) OrElse
               Not System.String.Equals(tokenType, "Bearer", System.StringComparison.OrdinalIgnoreCase) OrElse
               Not System.Int32.TryParse(If(json("expires_in")?.ToString(), ""), seconds) OrElse seconds <= 0 Then
                Throw New System.InvalidOperationException("The OAuth server returned an incomplete or unsupported access token.")
            End If
            If json("refresh_token") IsNot Nothing AndAlso System.String.IsNullOrWhiteSpace(json("refresh_token").ToString()) Then
                Throw New System.InvalidOperationException("The OAuth server returned an empty refresh token.")
            End If
            If json("scope") IsNot Nothing Then
                Dim granted As New System.Collections.Generic.HashSet(Of System.String)(
                    json("scope").ToString().Split(New System.Char() {" "c}, System.StringSplitOptions.RemoveEmptyEntries), System.StringComparer.Ordinal)
                If connection.Scope.Split(New System.Char() {" "c}, System.StringSplitOptions.RemoveEmptyEntries).Any(Function(item) Not granted.Contains(item)) Then
                    Throw New System.InvalidOperationException("The OAuth server did not grant all requested MCP permissions. Reconnect and approve the required access.")
                End If
            End If
            connection.AccessToken = token
            connection.ExpiresIn = seconds
            If json("refresh_token") IsNot Nothing Then connection.RefreshToken = json("refresh_token").ToString()
            If json("scope") IsNot Nothing Then connection.Scope = json("scope").ToString()
        End Sub

        Public Shared Async Function GetAutomaticMCPAuthorizationAsync(mcpEndpoint As System.String,
                Optional configuredScope As System.String = "", Optional forceRefresh As System.Boolean = False,
                Optional rejectedToken As System.String = "", Optional silent As System.Boolean = False,
                Optional cancellationToken As System.Threading.CancellationToken = Nothing,
                Optional reconnect As System.Boolean = False) As System.Threading.Tasks.Task(Of MCPProtectedResourceOAuthResult)
            Dim endpoint As System.String = AutomaticMCPResource(mcpEndpoint)
            Dim scope As System.String = System.String.Join(" ", If(configuredScope, "").Split(New System.Char() {" "c, Microsoft.VisualBasic.ControlChars.Tab}, System.StringSplitOptions.RemoveEmptyEntries).Distinct(System.StringComparer.Ordinal).OrderBy(Function(item) item, System.StringComparer.Ordinal))
            Dim path As System.String = AutomaticMCPCachePath(endpoint, scope)
            Using cacheLock As System.IO.FileStream = Await LockAutomaticMCPCacheAsync(path, cancellationToken).ConfigureAwait(False)
                If reconnect Then
                    RequireInteractiveExecution("Reset personal MCP authorization", "noninteractive_auth_required")
                    If System.IO.File.Exists(path) Then System.IO.File.Delete(path)
                End If
                Dim state As MCPAutomaticOAuthState = ReadAutomaticMCPState(path, endpoint, scope)
                If state IsNot Nothing AndAlso state.ExpiryUtc > System.DateTime.UtcNow AndAlso
                   Not System.String.IsNullOrWhiteSpace(state.Connection.AccessToken) AndAlso
                   (Not forceRefresh OrElse (rejectedToken.Length > 0 AndAlso Not System.String.Equals(rejectedToken, state.Connection.AccessToken, System.StringComparison.Ordinal))) Then
                    Return state.Connection
                End If
                cancellationToken.ThrowIfCancellationRequested()
                If state IsNot Nothing Then
                    state.Connection.AccessToken = ""
                    state.ExpiryUtc = System.DateTime.MinValue
                    WriteAutomaticMCPState(path, state)
                End If
                Dim discovered As MCPProtectedResourceOAuthResult = Await DiscoverAutomaticMCPAsync(endpoint, scope, cancellationToken).ConfigureAwait(False)
                If state Is Nothing Then state = New MCPAutomaticOAuthState With {.Endpoint = endpoint, .ConfiguredScope = scope}
                Dim saved As MCPProtectedResourceOAuthResult = state.Connection
                Dim sameAuthority As System.Boolean = saved IsNot Nothing AndAlso
                    System.String.Equals(saved.Issuer, discovered.Issuer, System.StringComparison.Ordinal) AndAlso
                    System.String.Equals(saved.Resource, discovered.Resource, System.StringComparison.Ordinal) AndAlso
                    System.String.Equals(saved.TokenEndpoint, discovered.TokenEndpoint, System.StringComparison.Ordinal)
                If sameAuthority Then
                    discovered.ClientId = saved.ClientId
                    discovered.ClientSecret = saved.ClientSecret
                    discovered.RedirectUri = saved.RedirectUri
                    discovered.RefreshToken = saved.RefreshToken
                    If scope.Length = 0 Then discovered.Scope = saved.Scope
                End If
                ' Persist invalidation before any network refresh or interactive operation.
                ' A timeout or cancellation must never resurrect a token rejected by the server.
                state.Connection = discovered
                state.ExpiryUtc = System.DateTime.MinValue
                WriteAutomaticMCPState(path, state)
                If discovered.RefreshToken.Length > 0 AndAlso discovered.ClientId.Length > 0 Then
                    Dim form As New System.Collections.Generic.Dictionary(Of System.String, System.String) From {
                        {"grant_type", "refresh_token"}, {"client_id", discovered.ClientId},
                        {"refresh_token", discovered.RefreshToken}, {"resource", discovered.Resource}
                    }
                    If discovered.ClientSecret.Length > 0 Then form("client_secret") = discovered.ClientSecret
                    Using client As System.Net.Http.HttpClient = NewAutomaticMCPHttpClient()
                        Using content As New System.Net.Http.FormUrlEncodedContent(form)
                            Using response As System.Net.Http.HttpResponseMessage = Await client.PostAsync(discovered.TokenEndpoint, content, cancellationToken).ConfigureAwait(False)
                                Dim json As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(Await response.Content.ReadAsStringAsync().ConfigureAwait(False))
                                If response.IsSuccessStatusCode Then
                                    ApplyAutomaticMCPToken(discovered, json)
                                    state.ExpiryUtc = AutomaticMCPExpiry(discovered.ExpiresIn)
                                    WriteAutomaticMCPState(path, state)
                                    Return discovered
                                End If
                                Dim errorCode As System.String = If(json("error")?.ToString(), "")
                                If errorCode <> "invalid_grant" AndAlso errorCode <> "invalid_client" Then
                                    Throw New System.InvalidOperationException("MCP OAuth token renewal failed (HTTP " & CInt(response.StatusCode).ToString() & "). Retry later.")
                                End If
                                LogMCPOAuth("Automatic token renewal requires sign-in", errorCode)
                                discovered.RefreshToken = ""
                                If errorCode = "invalid_client" Then discovered.ClientId = ""
                                WriteAutomaticMCPState(path, state)
                            End Using
                        End Using
                    End Using
                End If
                RequireInteractiveExecution("MCP OAuth sign-in: connect this server interactively before retrying", "noninteractive_auth_required")
                If silent Then Throw New System.InvalidOperationException("This MCP server requires interactive sign-in. Connect it in Freestyle and retry.")
                cancellationToken.ThrowIfCancellationRequested()
                If discovered.ClientId.Length = 0 Then
                    Await RegisterAutomaticMCPAsync(discovered, cancellationToken).ConfigureAwait(False)
                    WriteAutomaticMCPState(path, state)
                End If
                Dim authorized As MCPProtectedResourceOAuthResult = Await RunOobAuthorizationCodeFlowAsync(
                    discovered.ClientId, discovered.ClientSecret, discovered.AuthorizationEndpoint, discovered.TokenEndpoint,
                    discovered.RedirectUri, discovered.Scope, discovered.Resource, False, True, cancellationToken, discovered.Issuer).ConfigureAwait(False)
                discovered.AccessToken = authorized.AccessToken
                discovered.RefreshToken = authorized.RefreshToken
                discovered.ExpiresIn = authorized.ExpiresIn
                discovered.Scope = authorized.Scope
                state.ExpiryUtc = AutomaticMCPExpiry(discovered.ExpiresIn)
                WriteAutomaticMCPState(path, state)
                Return discovered
            End Using
        End Function

        Private Shared Function AutomaticMCPExpiry(seconds As System.Int32) As System.DateTime
            Return System.DateTime.UtcNow.AddSeconds(seconds - System.Math.Min(300, System.Math.Max(1, seconds \ 10)))
        End Function

        Public Shared Async Function AcquireMCPProtectedResourceOAuthAsync(
                protectedResourceUrl As String,
                Optional configuredScope As String = "",
                Optional existingClientId As String = "",
                Optional existingClientSecret As String = "",
                Optional silent As Boolean = False) As Task(Of MCPProtectedResourceOAuthResult)

            RequireInteractiveExecution("MCP protected-resource authorization", "noninteractive_auth_required")

            LogMCPOAuth("AcquireMCPProtectedResourceOAuthAsync: starting", "protectedResourceUrl=" & protectedResourceUrl)

            Dim challenge = Await GetBearerChallengeAsync(protectedResourceUrl).ConfigureAwait(False)
            LogMCPOAuth("Bearer challenge", If(challenge Is Nothing, "(none)", String.Join(", ", challenge.Select(Function(kv) kv.Key & "=" & kv.Value))))

            Dim resourceMetadataUrl As String = ""
            Dim authorizationServerHint As String = ""
            Dim protectedResourceMetadata As JObject = Nothing

            If challenge IsNot Nothing Then
                resourceMetadataUrl = GetChallengeValue(challenge, "resource_metadata")
                authorizationServerHint = GetChallengeValue(challenge, "authorization_uri")

                If Not String.IsNullOrWhiteSpace(resourceMetadataUrl) Then
                    protectedResourceMetadata = Await GetJsonAsync(resourceMetadataUrl).ConfigureAwait(False)
                End If
            End If

            If protectedResourceMetadata Is Nothing Then
                protectedResourceMetadata = Await GetProtectedResourceMetadataAsync(protectedResourceUrl).ConfigureAwait(False)
            End If

            LogMCPOAuth("Protected resource metadata", If(protectedResourceMetadata Is Nothing, "(none)", protectedResourceMetadata.ToString()))

            If challenge Is Nothing AndAlso protectedResourceMetadata Is Nothing Then
                Return Nothing
            End If

            Dim resource As String = ""
            If protectedResourceMetadata IsNot Nothing Then
                resource = If(protectedResourceMetadata("resource")?.ToString(), "")
            End If
            If String.IsNullOrWhiteSpace(resource) Then
                resource = protectedResourceUrl
            End If

            Dim authorizationServer As String = authorizationServerHint

            If String.IsNullOrWhiteSpace(authorizationServer) AndAlso protectedResourceMetadata IsNot Nothing Then
                Dim servers As JArray = TryCast(protectedResourceMetadata("authorization_servers"), JArray)
                If servers IsNot Nothing AndAlso servers.Count > 0 Then
                    authorizationServer = servers(0).ToString()
                End If
            End If

            If String.IsNullOrWhiteSpace(authorizationServer) Then
                Throw New Exception("The MCP server requires OAuth, but no authorization server could be discovered.")
            End If

            authorizationServer = RequireMcpOAuthHttpsEndpoint(authorizationServer, "authorization server")
            LogMCPOAuth("Authorization server candidate", authorizationServer)

            Dim authorizationServerMetadata = Await GetAuthorizationServerMetadataAsync(authorizationServer).ConfigureAwait(False)

            LogMCPOAuth("Authorization server metadata", authorizationServerMetadata.ToString())

            Dim authorizationEndpoint As String = If(authorizationServerMetadata("authorization_endpoint")?.ToString(), "")
            Dim tokenEndpoint As String = If(authorizationServerMetadata("token_endpoint")?.ToString(), "")
            Dim registrationEndpoint As String = If(authorizationServerMetadata("registration_endpoint")?.ToString(), "")
            Dim deviceAuthorizationEndpoint As String = If(authorizationServerMetadata("device_authorization_endpoint")?.ToString(), "")

            If String.IsNullOrWhiteSpace(tokenEndpoint) Then
                Throw New Exception("The OAuth authorization server metadata is missing token_endpoint.")
            End If

            ' token_endpoint is always used. Optional endpoints are validated immediately before use
            ' so an unused legacy metadata field cannot break an otherwise secure OAuth flow.
            tokenEndpoint = RequireMcpOAuthHttpsEndpoint(tokenEndpoint, "token_endpoint")

            Dim scope As String = configuredScope
            If String.IsNullOrWhiteSpace(scope) Then
                scope = GetChallengeValue(challenge, "scope")
            End If

            Dim clientId As String = existingClientId
            Dim clientSecret As String = existingClientSecret
            Dim httpsLoopbackRedirectUri As String = "https://localhost/callback"
            Dim oobRedirectUri As String = "urn:ietf:wg:oauth:2.0:oob"
            Dim acceptedRedirectUri As String = ""

            If String.IsNullOrWhiteSpace(clientId) Then
                If String.IsNullOrWhiteSpace(registrationEndpoint) Then
                    Throw New Exception("The MCP OAuth server requires a client id, but did not advertise dynamic client registration.")
                End If

                Dim registration As JObject = Nothing
                Dim lastRegError As String = ""

                ' Attempt 1: no redirect_uris (servers that don't require them)
                Try
                    registration = Await RegisterOAuthCombinedClientAsync(registrationEndpoint, Nothing).ConfigureAwait(False)
                    acceptedRedirectUri = ""
                Catch ex As Exception
                    lastRegError = ex.Message
                End Try

                ' Attempt 2: HTTPS loopback (preferred for HTTPS-only servers)
                If registration Is Nothing Then
                    Try
                        registration = Await RegisterOAuthCombinedClientAsync(registrationEndpoint, httpsLoopbackRedirectUri).ConfigureAwait(False)
                        acceptedRedirectUri = httpsLoopbackRedirectUri
                    Catch ex As Exception
                        lastRegError = ex.Message
                    End Try
                End If

                ' Attempt 3: OOB (last resort)
                If registration Is Nothing Then
                    Try
                        registration = Await RegisterOAuthCombinedClientAsync(registrationEndpoint, oobRedirectUri).ConfigureAwait(False)
                        acceptedRedirectUri = oobRedirectUri
                    Catch ex As Exception
                        lastRegError = ex.Message
                    End Try
                End If

                If registration Is Nothing Then
                    Throw New Exception($"Dynamic OAuth client registration failed. Last error: {lastRegError}")
                End If

                clientId = If(registration("client_id")?.ToString(), "")
                clientSecret = If(registration("client_secret")?.ToString(), "")

                If String.IsNullOrWhiteSpace(deviceAuthorizationEndpoint) Then
                    deviceAuthorizationEndpoint = If(registration("device_authorization_endpoint")?.ToString(), "")
                    If Not String.IsNullOrWhiteSpace(deviceAuthorizationEndpoint) Then
                        deviceAuthorizationEndpoint = RequireMcpOAuthHttpsEndpoint(deviceAuthorizationEndpoint, "device_authorization_endpoint")
                    End If
                End If

                LogMCPOAuth("Registration succeeded", $"client_id={clientId}; acceptedRedirectUri={acceptedRedirectUri}")

                If String.IsNullOrWhiteSpace(clientId) Then
                    Throw New Exception("Dynamic OAuth client registration did not return a client_id.")
                End If
            End If

            Dim useDeviceFlow As Boolean = Not String.IsNullOrWhiteSpace(deviceAuthorizationEndpoint)

            If useDeviceFlow Then
                Return Await RunDeviceAuthorizationFlowAsync(
                    clientId, clientSecret, deviceAuthorizationEndpoint,
                    tokenEndpoint, scope, resource, silent).ConfigureAwait(False)
            End If

            If String.IsNullOrWhiteSpace(authorizationEndpoint) Then
                Throw New Exception("The OAuth authorization server does not advertise a device_authorization_endpoint and no authorization_endpoint is available.")
            End If

            ' Fall back to interactive auth code flow with the redirect URI the server actually accepted.
            Dim flowRedirectUri As String = If(String.IsNullOrWhiteSpace(acceptedRedirectUri), httpsLoopbackRedirectUri, acceptedRedirectUri)

            Return Await RunOobAuthorizationCodeFlowAsync(
                clientId, clientSecret, authorizationEndpoint,
                tokenEndpoint, flowRedirectUri, scope, resource, silent).ConfigureAwait(False)
        End Function

        Public Shared Async Function GetInteractiveMCPAccessTokenAsync(
                clientId As String,
                scope As String,
                clientSecret As String,
                oauthEndpointConfig As String,
                silent As Boolean) As Task(Of MCPProtectedResourceOAuthResult)

            RequireInteractiveExecution("OAuth authorization", "noninteractive_auth_required")
            Dim parts = oauthEndpointConfig.Split(New String() {"¦"}, StringSplitOptions.None)
            If parts.Length < 2 Then
                Throw New Exception("OAuth2Endpoint must be '<authorization endpoint>¦<token endpoint>[¦resource]' or 'device:<device endpoint>¦<token endpoint>[¦resource]'.")
            End If

            If parts(0).StartsWith("device:", StringComparison.OrdinalIgnoreCase) Then
                Dim deviceAuthorizationEndpoint As String = parts(0).Substring("device:".Length).Trim()
                Dim deviceTokenEndpoint As String = parts(1).Trim()
                Dim deviceResource As String = If(parts.Length >= 3, parts(2).Trim(), "")
                deviceAuthorizationEndpoint = RequireMcpOAuthHttpsEndpoint(deviceAuthorizationEndpoint, "device_authorization_endpoint")
                deviceTokenEndpoint = RequireMcpOAuthHttpsEndpoint(deviceTokenEndpoint, "token_endpoint")

                Return Await RunDeviceAuthorizationFlowAsync(
                    clientId, clientSecret, deviceAuthorizationEndpoint,
                    deviceTokenEndpoint, scope, deviceResource, silent).ConfigureAwait(False)
            End If

            Dim authorizationEndpoint As String = parts(0).Trim()
            Dim authCodeTokenEndpoint As String = parts(1).Trim()
            Dim authCodeResource As String = If(parts.Length >= 3, parts(2).Trim(), "")
            authorizationEndpoint = RequireMcpOAuthHttpsEndpoint(authorizationEndpoint, "authorization_endpoint")
            authCodeTokenEndpoint = RequireMcpOAuthHttpsEndpoint(authCodeTokenEndpoint, "token_endpoint")

            ' Use OOB flow with https://localhost/callback (the redirect URI accepted at registration time).
            Return Await RunOobAuthorizationCodeFlowAsync(
                clientId,
                clientSecret,
                authorizationEndpoint,
                authCodeTokenEndpoint,
                "https://localhost/callback",
                scope,
                authCodeResource,
                silent).ConfigureAwait(False)
        End Function

        Private Shared Async Function GetBearerChallengeAsync(protectedResourceUrl As String) As Task(Of Dictionary(Of String, String))
            Using handler As New HttpClientHandler()
                handler.AutomaticDecompression = DecompressionMethods.GZip Or DecompressionMethods.Deflate
                handler.AllowAutoRedirect = False
                handler.UseProxy = True
                handler.Proxy = WebRequest.DefaultWebProxy
                If handler.Proxy IsNot Nothing Then
                    handler.Proxy.Credentials = CredentialCache.DefaultCredentials
                End If

                Using client As New HttpClient(handler)
                    client.Timeout = TimeSpan.FromSeconds(30)
                    client.DefaultRequestHeaders.Accept.Clear()
                    client.DefaultRequestHeaders.Accept.Add(New Headers.MediaTypeWithQualityHeaderValue("text/event-stream"))
                    client.DefaultRequestHeaders.Accept.Add(New Headers.MediaTypeWithQualityHeaderValue("application/json"))

                    Using request As New HttpRequestMessage(HttpMethod.Get, protectedResourceUrl)
                        Using response = Await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(False)
                            Dim allHeaders As String = String.Join(Environment.NewLine,
                                response.Headers.SelectMany(Function(h) h.Value.Select(Function(v) h.Key & ": " & v)))
                            LogMCPOAuth($"Bearer probe GET {protectedResourceUrl} -> HTTP {CInt(response.StatusCode)}", allHeaders)

                            If response.StatusCode <> HttpStatusCode.Unauthorized Then
                                Return Nothing
                            End If

                            Dim headerValues As IEnumerable(Of String) = Nothing
                            If Not response.Headers.TryGetValues("WWW-Authenticate", headerValues) Then
                                Return Nothing
                            End If

                            For Each headerValue In headerValues
                                If headerValue.TrimStart().StartsWith("Bearer", StringComparison.OrdinalIgnoreCase) Then
                                    Return ParseWwwAuthenticateBearer(headerValue)
                                End If
                            Next
                        End Using
                    End Using
                End Using
            End Using

            Return Nothing
        End Function

        Private Shared Function ParseWwwAuthenticateBearer(headerValue As String) As Dictionary(Of String, String)
            Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim value As String = headerValue.Trim()

            If value.StartsWith("Bearer", StringComparison.OrdinalIgnoreCase) Then
                value = value.Substring("Bearer".Length).Trim()
            End If

            Dim parts As New List(Of String)()
            Dim current As New StringBuilder()
            Dim inQuotes As Boolean = False

            For Each ch As Char In value
                If ch = """"c Then
                    inQuotes = Not inQuotes
                    current.Append(ch)
                ElseIf ch = ","c AndAlso Not inQuotes Then
                    parts.Add(current.ToString())
                    current.Clear()
                Else
                    current.Append(ch)
                End If
            Next

            If current.Length > 0 Then
                parts.Add(current.ToString())
            End If

            For Each part In parts
                Dim idx = part.IndexOf("="c)
                If idx <= 0 Then Continue For

                Dim key = part.Substring(0, idx).Trim()
                Dim raw = part.Substring(idx + 1).Trim()

                If raw.StartsWith("""") AndAlso raw.EndsWith("""") AndAlso raw.Length >= 2 Then
                    raw = raw.Substring(1, raw.Length - 2)
                End If

                If key <> "" Then
                    result(key) = raw
                End If
            Next

            Return result
        End Function

        Private Shared Function GetChallengeValue(challenge As Dictionary(Of String, String), key As String) As String
            If challenge Is Nothing Then Return ""
            Dim value As String = ""
            If challenge.TryGetValue(key, value) Then Return value
            Return ""
        End Function

        Private Shared Async Function GetProtectedResourceMetadataAsync(protectedResourceUrl As String) As Task(Of JObject)
            Dim candidates As List(Of String) = BuildProtectedResourceMetadataCandidates(protectedResourceUrl)

            For Each candidate As String In candidates
                Try
                    Dim json As JObject = Await GetJsonAsync(candidate).ConfigureAwait(False)

                    If json IsNot Nothing AndAlso
                       (json("authorization_servers") IsNot Nothing OrElse
                        json("resource") IsNot Nothing OrElse
                        json("jwks_uri") IsNot Nothing) Then
                        Return json
                    End If
                Catch
                End Try
            Next

            Return Nothing
        End Function

        Private Shared Function BuildProtectedResourceMetadataCandidates(protectedResourceUrl As String) As List(Of String)
            Dim result As New List(Of String)()

            Dim uri As New Uri(protectedResourceUrl)
            Dim origin As String = uri.GetLeftPart(UriPartial.Authority)
            Dim path As String = uri.AbsolutePath.Trim("/"c)

            If path <> "" Then
                result.Add(origin & "/.well-known/oauth-protected-resource/" & path)
            End If

            result.Add(origin & "/.well-known/oauth-protected-resource")

            Return result.
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToList()
        End Function

        Private Shared Async Function GetAuthorizationServerMetadataAsync(authorizationServer As String) As Task(Of JObject)
            authorizationServer = RequireMcpOAuthHttpsEndpoint(authorizationServer, "authorization server")

            Dim candidates As New List(Of String) From {
                authorizationServer.TrimEnd("/"c),
                authorizationServer.TrimEnd("/"c) & "/.well-known/oauth-authorization-server",
                authorizationServer.TrimEnd("/"c) & "/.well-known/openid-configuration"
            }

            Dim lastError As Exception = Nothing

            For Each candidate In candidates.Distinct(StringComparer.OrdinalIgnoreCase)
                LogMCPOAuth("Trying auth server metadata candidate", candidate)

                Try
                    Dim json = Await GetJsonAsync(candidate).ConfigureAwait(False)
                    LogMCPOAuth("Candidate response", If(json Is Nothing, "(null)", json.ToString()))

                    If json("token_endpoint") IsNot Nothing AndAlso
                       (json("authorization_endpoint") IsNot Nothing OrElse
                        json("device_authorization_endpoint") IsNot Nothing) Then
                        LogMCPOAuth("Accepted candidate", candidate)
                        Return json
                    End If

                    LogMCPOAuth("Rejected candidate (missing required endpoints)", candidate)
                Catch ex As Exception
                    LogMCPOAuth("Candidate threw", $"{candidate}: {ex.Message}")
                    lastError = ex
                End Try
            Next

            Throw New Exception("Could not discover OAuth authorization server metadata.", lastError)
        End Function

        Private Shared Async Function GetJsonAsync(url As String) As Task(Of JObject)
            Using client As New HttpClient()
                client.Timeout = TimeSpan.FromSeconds(30)
                client.DefaultRequestHeaders.Accept.Clear()
                client.DefaultRequestHeaders.Accept.Add(New Headers.MediaTypeWithQualityHeaderValue("application/json"))

                Dim text = Await client.GetStringAsync(url).ConfigureAwait(False)
                Return JObject.Parse(text)
            End Using
        End Function

        ''' <summary>
        ''' Registers a native OAuth client supporting both authorization_code and device_code grant types.
        ''' If redirectUri is Nothing, no redirect_uris field is sent (ideal for device-only servers).
        ''' If redirectUri is provided, it is included (required by servers that enforce redirect_uris).
        ''' </summary>
        Private Shared Async Function RegisterOAuthCombinedClientAsync(
                registrationEndpoint As String,
                redirectUri As String) As Task(Of JObject)

            registrationEndpoint = RequireMcpOAuthHttpsEndpoint(registrationEndpoint, "registration_endpoint")

            Dim payload As New JObject From {
                {"client_name", "Red Ink"},
                {"application_type", "native"},
                {"grant_types", New JArray(
                    "authorization_code",
                    "urn:ietf:params:oauth:grant-type:device_code")},
                {"response_types", New JArray("code")},
                {"token_endpoint_auth_method", "none"}
            }

            If Not String.IsNullOrWhiteSpace(redirectUri) Then
                payload("redirect_uris") = New JArray(redirectUri)
            End If

            LogMCPOAuth("Registration request -> " & registrationEndpoint, payload.ToString())

            Using client As New HttpClient()
                client.Timeout = TimeSpan.FromSeconds(30)

                Dim content As New StringContent(payload.ToString(Formatting.None), Encoding.UTF8, "application/json")
                Dim response = Await client.PostAsync(registrationEndpoint, content).ConfigureAwait(False)
                Dim responseText = Await response.Content.ReadAsStringAsync().ConfigureAwait(False)

                LogMCPOAuth($"Registration response HTTP {CInt(response.StatusCode)}", responseText)

                If Not response.IsSuccessStatusCode Then
                    Throw New Exception($"HTTP {CInt(response.StatusCode)} {responseText}")
                End If

                Return JObject.Parse(responseText)
            End Using
        End Function

        ''' <summary>
        ''' OOB (out-of-band) authorization code flow for servers that reject HTTP loopback redirect URIs.
        ''' Opens the browser to the authorization endpoint, then asks the user to paste the redirect URL.
        ''' </summary>
        Private Shared Async Function RunOobAuthorizationCodeFlowAsync(
                clientId As String,
                clientSecret As String,
                authorizationEndpoint As String,
                tokenEndpoint As String,
                redirectUri As String,
                scope As String,
                resource As String,
                silent As Boolean,
                Optional strictOAuth As System.Boolean = False,
                Optional cancellationToken As System.Threading.CancellationToken = Nothing,
                Optional expectedIssuer As System.String = "") As System.Threading.Tasks.Task(Of MCPProtectedResourceOAuthResult)

            authorizationEndpoint = RequireMcpOAuthHttpsEndpoint(authorizationEndpoint, "authorization_endpoint")
            tokenEndpoint = RequireMcpOAuthHttpsEndpoint(tokenEndpoint, "token_endpoint")

            Dim state As String = GenerateOAuthRandomValue()
            Dim codeVerifier As String = GenerateOAuthRandomValue()
            Dim codeChallenge As String = CreateOAuthCodeChallenge(codeVerifier)

            Dim authUrl As New StringBuilder()
            authUrl.Append(authorizationEndpoint)
            authUrl.Append(If(authorizationEndpoint.Contains("?"), "&", "?"))
            authUrl.Append("response_type=code")
            authUrl.Append("&client_id=").Append(Uri.EscapeDataString(clientId))
            authUrl.Append("&redirect_uri=").Append(Uri.EscapeDataString(redirectUri))
            authUrl.Append("&state=").Append(Uri.EscapeDataString(state))
            authUrl.Append("&code_challenge=").Append(Uri.EscapeDataString(codeChallenge))
            authUrl.Append("&code_challenge_method=S256")

            If Not String.IsNullOrWhiteSpace(scope) Then
                authUrl.Append("&scope=").Append(Uri.EscapeDataString(scope))
            End If

            If Not String.IsNullOrWhiteSpace(resource) Then
                authUrl.Append("&resource=").Append(Uri.EscapeDataString(resource))
            End If

            Dim capturedRedirect As String = ""

            If silent Then
                ' Headless mode: no UI possible, fall back to opening the system browser
                ' and waiting for a paste via ShowCustomInputBox is not possible. Treat as cancelled.
                Throw New Exception("OAuth authentication required but UI is suppressed.")
            End If

            capturedRedirect = RunOnStaThread(Of System.String)(
                Function()
                    Using browserForm As New MCPOAuthBrowserForm(authUrl.ToString(), redirectUri)
                        cancellationToken.ThrowIfCancellationRequested()
                        Dim cancellationHandle As System.IntPtr = browserForm.Handle
                        Using cancelRegistration As System.Threading.CancellationTokenRegistration = cancellationToken.Register(
                            Sub()
                                Try
                                    If browserForm.IsHandleCreated AndAlso Not browserForm.IsDisposed Then
                                        browserForm.BeginInvoke(New System.Action(Sub() browserForm.Close()))
                                    End If
                                Catch ex As System.InvalidOperationException
                                    ' The STA may already have disposed the window.
                                End Try
                            End Sub)
                            Dim owner As System.Windows.Forms.IWin32Window = Nothing
                            If IsOnHostUiThread Then owner = ResolveSameThreadDialogOwner()
                            Dim result As System.Windows.Forms.DialogResult = If(owner Is Nothing, browserForm.ShowDialog(), browserForm.ShowDialog(owner))
                            If result <> System.Windows.Forms.DialogResult.OK OrElse System.String.IsNullOrWhiteSpace(browserForm.CapturedRedirectUrl) Then
                                Throw New System.Exception("OAuth authentication was cancelled or did not complete.")
                            End If
                            cancellationToken.ThrowIfCancellationRequested()
                            Return browserForm.CapturedRedirectUrl
                        End Using
                    End Using
                End Function, "Red Ink MCP OAuth")

            ' Extract the authorization code from the captured redirect URL.
            Dim code As String = ""
            Try
                Dim uri As New Uri(capturedRedirect)
                Dim qs = System.Web.HttpUtility.ParseQueryString(uri.Query)
                If Not IsMCPRedirectCallback(capturedRedirect, redirectUri) OrElse
                   Not System.String.Equals(qs("state"), state, System.StringComparison.Ordinal) Then
                    Throw New System.Exception("OAuth callback validation failed. Start sign-in again.")
                End If

                If expectedIssuer.Length > 0 AndAlso qs("iss") IsNot Nothing AndAlso
                   Not System.String.Equals(qs("iss"), expectedIssuer, System.StringComparison.Ordinal) Then
                    Throw New System.Exception("OAuth callback identifies a different authorization server.")
                End If
                Dim errorText As String = If(qs("error"), "")
                If Not String.IsNullOrWhiteSpace(errorText) Then
                    Throw New Exception("OAuth authorization failed: " & errorText)
                End If

                code = If(qs("code"), "")
            Catch ex As UriFormatException
                Throw New Exception("OAuth callback URL could not be parsed: " & ex.Message)
            End Try

            If String.IsNullOrWhiteSpace(code) Then
                Throw New Exception("OAuth callback did not contain an authorization code.")
            End If

            Dim formData As New Dictionary(Of String, String) From {
                {"grant_type", "authorization_code"},
                {"client_id", clientId},
                {"code", code},
                {"redirect_uri", redirectUri},
                {"code_verifier", codeVerifier}
            }

            If Not String.IsNullOrWhiteSpace(clientSecret) Then
                formData("client_secret") = clientSecret
            End If

            If Not String.IsNullOrWhiteSpace(resource) Then
                formData("resource") = resource
            End If

            Using client As System.Net.Http.HttpClient = If(strictOAuth, NewAutomaticMCPHttpClient(), New System.Net.Http.HttpClient())
                client.Timeout = TimeSpan.FromSeconds(60)

                Dim response = Await client.PostAsync(tokenEndpoint, New FormUrlEncodedContent(formData), cancellationToken).ConfigureAwait(False)
                Dim responseText = Await response.Content.ReadAsStringAsync().ConfigureAwait(False)

                If Not response.IsSuccessStatusCode Then
                    Throw New System.Exception(If(strictOAuth,
                        "OAuth token exchange failed (HTTP " & CInt(response.StatusCode).ToString() & "). Start sign-in again.",
                        $"OAuth token exchange failed: HTTP {CInt(response.StatusCode)} {responseText}"))
                End If

                Dim tokenJson = JObject.Parse(responseText)
                Dim accessToken As String = If(tokenJson("access_token")?.ToString(), "")

                If String.IsNullOrWhiteSpace(accessToken) Then
                    Throw New Exception("OAuth token response did not contain access_token.")
                End If

                Dim expiresIn As Integer = 3600
                If tokenJson("expires_in") IsNot Nothing Then
                    Integer.TryParse(tokenJson("expires_in").ToString(), expiresIn)
                    If expiresIn <= 0 Then expiresIn = 3600
                End If

                Dim result As New MCPProtectedResourceOAuthResult With {
                    .AccessToken = accessToken,
                    .ClientId = clientId,
                    .ClientSecret = clientSecret,
                    .AuthorizationEndpoint = authorizationEndpoint,
                    .TokenEndpoint = tokenEndpoint,
                    .Scope = scope,
                    .Resource = resource,
                    .ExpiresIn = expiresIn,
                    .UseDeviceFlow = False,
                    .RedirectUri = redirectUri
                }
                If strictOAuth Then ApplyAutomaticMCPToken(result, tokenJson)
                Return result
            End Using
        End Function

        Private Shared Async Function RunDeviceAuthorizationFlowAsync(
                clientId As String,
                clientSecret As String,
                deviceAuthorizationEndpoint As String,
                tokenEndpoint As String,
                scope As String,
                resource As String,
                silent As Boolean) As Task(Of MCPProtectedResourceOAuthResult)

            deviceAuthorizationEndpoint = RequireMcpOAuthHttpsEndpoint(deviceAuthorizationEndpoint, "device_authorization_endpoint")
            tokenEndpoint = RequireMcpOAuthHttpsEndpoint(tokenEndpoint, "token_endpoint")

            If String.IsNullOrWhiteSpace(clientId) Then
                Throw New Exception("OAuth client_id is missing.")
            End If

            If String.IsNullOrWhiteSpace(deviceAuthorizationEndpoint) Then
                Throw New Exception("OAuth device_authorization_endpoint is missing.")
            End If

            Using client As New HttpClient()
                client.Timeout = TimeSpan.FromSeconds(60)

                Dim deviceRequest As New Dictionary(Of String, String) From {
                    {"client_id", clientId}
                }

                If Not String.IsNullOrWhiteSpace(clientSecret) Then
                    deviceRequest("client_secret") = clientSecret
                End If

                If Not String.IsNullOrWhiteSpace(scope) Then
                    deviceRequest("scope") = scope
                End If

                If Not String.IsNullOrWhiteSpace(resource) Then
                    deviceRequest("resource") = resource
                End If

                Dim deviceResponse = Await client.PostAsync(
                    deviceAuthorizationEndpoint,
                    New FormUrlEncodedContent(deviceRequest)).ConfigureAwait(False)

                Dim deviceResponseText = Await deviceResponse.Content.ReadAsStringAsync().ConfigureAwait(False)

                If Not deviceResponse.IsSuccessStatusCode Then
                    Throw New Exception($"OAuth device authorization request failed: HTTP {CInt(deviceResponse.StatusCode)} {deviceResponseText}")
                End If

                Dim deviceJson As JObject = JObject.Parse(deviceResponseText)

                Dim deviceCode As String = If(deviceJson("device_code")?.ToString(), "")
                Dim userCode As String = If(deviceJson("user_code")?.ToString(), "")
                Dim verificationUri As String = If(deviceJson("verification_uri")?.ToString(), "")
                Dim verificationUriComplete As String = If(deviceJson("verification_uri_complete")?.ToString(), "")
                Dim openUrl As String = If(Not String.IsNullOrWhiteSpace(verificationUriComplete), verificationUriComplete, verificationUri)

                If String.IsNullOrWhiteSpace(deviceCode) OrElse String.IsNullOrWhiteSpace(openUrl) Then
                    Throw New Exception("OAuth device authorization response did not contain the required device_code or verification URI.")
                End If

                openUrl = RequireMcpOAuthHttpsEndpoint(openUrl, "verification_uri")
                Process.Start(New ProcessStartInfo(openUrl) With {.UseShellExecute = True})

                If String.IsNullOrWhiteSpace(verificationUriComplete) AndAlso
                   Not String.IsNullOrWhiteSpace(userCode) AndAlso
                   Not silent Then

                    ShowCustomMessageBox(
                        "Complete sign-in in the browser that just opened." & vbCrLf & vbCrLf &
                        "If prompted for a user code, enter:" & vbCrLf & vbCrLf &
                        userCode)
                End If

                Dim intervalSeconds As Integer = 5
                If deviceJson("interval") IsNot Nothing Then
                    Integer.TryParse(deviceJson("interval").ToString(), intervalSeconds)
                    If intervalSeconds <= 0 Then intervalSeconds = 5
                End If

                Dim expiresInSeconds As Integer = 900
                If deviceJson("expires_in") IsNot Nothing Then
                    Integer.TryParse(deviceJson("expires_in").ToString(), expiresInSeconds)
                    If expiresInSeconds <= 0 Then expiresInSeconds = 900
                End If

                Dim deadline As DateTime = DateTime.UtcNow.AddSeconds(expiresInSeconds)

                Do While DateTime.UtcNow < deadline
                    Await Task.Delay(TimeSpan.FromSeconds(intervalSeconds)).ConfigureAwait(False)

                    Dim tokenRequest As New Dictionary(Of String, String) From {
                        {"grant_type", "urn:ietf:params:oauth:grant-type:device_code"},
                        {"device_code", deviceCode},
                        {"client_id", clientId}
                    }

                    If Not String.IsNullOrWhiteSpace(clientSecret) Then
                        tokenRequest("client_secret") = clientSecret
                    End If

                    If Not String.IsNullOrWhiteSpace(resource) Then
                        tokenRequest("resource") = resource
                    End If

                    Dim tokenResponse = Await client.PostAsync(
                        tokenEndpoint,
                        New FormUrlEncodedContent(tokenRequest)).ConfigureAwait(False)

                    Dim tokenResponseText = Await tokenResponse.Content.ReadAsStringAsync().ConfigureAwait(False)

                    If tokenResponse.IsSuccessStatusCode Then
                        Dim tokenJson As JObject = JObject.Parse(tokenResponseText)
                        Dim accessToken As String = If(tokenJson("access_token")?.ToString(), "")

                        If String.IsNullOrWhiteSpace(accessToken) Then
                            Throw New Exception("OAuth token response did not contain access_token.")
                        End If

                        Dim tokenExpiresIn As Integer = 3600
                        If tokenJson("expires_in") IsNot Nothing Then
                            Integer.TryParse(tokenJson("expires_in").ToString(), tokenExpiresIn)
                            If tokenExpiresIn <= 0 Then tokenExpiresIn = 3600
                        End If

                        Return New MCPProtectedResourceOAuthResult With {
                            .AccessToken = accessToken,
                            .ClientId = clientId,
                            .ClientSecret = clientSecret,
                            .TokenEndpoint = tokenEndpoint,
                            .DeviceAuthorizationEndpoint = deviceAuthorizationEndpoint,
                            .Scope = scope,
                            .Resource = resource,
                            .ExpiresIn = tokenExpiresIn,
                            .UseDeviceFlow = True
                        }
                    End If

                    Dim errorCode As String = ""
                    Try
                        Dim errorJson As JObject = JObject.Parse(tokenResponseText)
                        errorCode = If(errorJson("error")?.ToString(), "")
                    Catch
                    End Try

                    Select Case errorCode
                        Case "authorization_pending"
                            Continue Do

                        Case "slow_down"
                            intervalSeconds += 5
                            Continue Do

                        Case "access_denied"
                            Throw New Exception("OAuth authorization was denied by the user.")

                        Case "expired_token"
                            Throw New Exception("OAuth device authorization expired before completion.")

                        Case Else
                            Throw New Exception($"OAuth device token polling failed: HTTP {CInt(tokenResponse.StatusCode)} {tokenResponseText}")
                    End Select
                Loop
            End Using

            Throw New Exception("OAuth device authorization timed out.")
        End Function

        Private Shared Async Function RunInteractiveMcpOAuthCodeFlowAsync(
                clientId As String,
                clientSecret As String,
                authorizationEndpoint As String,
                tokenEndpoint As String,
                scope As String,
                resource As String,
                silent As Boolean) As Task(Of MCPProtectedResourceOAuthResult)

            authorizationEndpoint = RequireMcpOAuthHttpsEndpoint(authorizationEndpoint, "authorization_endpoint")
            tokenEndpoint = RequireMcpOAuthHttpsEndpoint(tokenEndpoint, "token_endpoint")

            If String.IsNullOrWhiteSpace(clientId) Then
                Throw New Exception("OAuth client_id is missing.")
            End If

            Dim state As String = GenerateOAuthRandomValue()
            Dim codeVerifier As String = GenerateOAuthRandomValue()
            Dim codeChallenge As String = CreateOAuthCodeChallenge(codeVerifier)
            Dim port As Integer = GetFreeLoopbackPort()
            Dim redirectUri As String = $"http://127.0.0.1:{port}/callback/"

            Using listener As New HttpListener()
                listener.Prefixes.Add(redirectUri)
                listener.Start()

                Dim authUrl As New StringBuilder()
                authUrl.Append(authorizationEndpoint)
                authUrl.Append(If(authorizationEndpoint.Contains("?"), "&", "?"))
                authUrl.Append("response_type=code")
                authUrl.Append("&client_id=").Append(Uri.EscapeDataString(clientId))
                authUrl.Append("&redirect_uri=").Append(Uri.EscapeDataString(redirectUri))
                authUrl.Append("&state=").Append(Uri.EscapeDataString(state))
                authUrl.Append("&code_challenge=").Append(Uri.EscapeDataString(codeChallenge))
                authUrl.Append("&code_challenge_method=S256")

                If Not String.IsNullOrWhiteSpace(scope) Then
                    authUrl.Append("&scope=").Append(Uri.EscapeDataString(scope))
                End If

                If Not String.IsNullOrWhiteSpace(resource) Then
                    authUrl.Append("&resource=").Append(Uri.EscapeDataString(resource))
                End If

                Process.Start(New ProcessStartInfo(authUrl.ToString()) With {.UseShellExecute = True})

                Dim contextTask = listener.GetContextAsync()
                Dim completed = Await Task.WhenAny(contextTask, Task.Delay(TimeSpan.FromMinutes(3))).ConfigureAwait(False)

                If completed IsNot contextTask Then
                    Throw New Exception("OAuth sign-in timed out.")
                End If

                Dim httpContext = Await contextTask.ConfigureAwait(False)

                Dim callbackState As String = httpContext.Request.QueryString("state")
                Dim code As String = httpContext.Request.QueryString("code")
                Dim errorText As String = httpContext.Request.QueryString("error")

                Dim browserMessage = "<html><body><h2>Authentication complete.</h2><p>You may close this window and return to Red Ink.</p></body></html>"
                Dim browserBytes = Encoding.UTF8.GetBytes(browserMessage)

                httpContext.Response.ContentType = "text/html; charset=utf-8"
                httpContext.Response.ContentLength64 = browserBytes.Length
                Await httpContext.Response.OutputStream.WriteAsync(browserBytes, 0, browserBytes.Length).ConfigureAwait(False)
                httpContext.Response.Close()

                If Not String.IsNullOrWhiteSpace(errorText) Then
                    Throw New Exception("OAuth authorization failed: " & errorText)
                End If

                If String.IsNullOrWhiteSpace(code) OrElse Not String.Equals(callbackState, state, StringComparison.Ordinal) Then
                    Throw New Exception("OAuth callback validation failed.")
                End If

                Dim formData As New Dictionary(Of String, String) From {
                    {"grant_type", "authorization_code"},
                    {"client_id", clientId},
                    {"code", code},
                    {"redirect_uri", redirectUri},
                    {"code_verifier", codeVerifier}
                }

                If Not String.IsNullOrWhiteSpace(clientSecret) Then
                    formData("client_secret") = clientSecret
                End If

                If Not String.IsNullOrWhiteSpace(resource) Then
                    formData("resource") = resource
                End If

                Using client As New HttpClient()
                    client.Timeout = TimeSpan.FromSeconds(60)

                    Dim response = Await client.PostAsync(tokenEndpoint, New FormUrlEncodedContent(formData)).ConfigureAwait(False)
                    Dim responseText = Await response.Content.ReadAsStringAsync().ConfigureAwait(False)

                    If Not response.IsSuccessStatusCode Then
                        Throw New Exception($"OAuth token exchange failed: HTTP {CInt(response.StatusCode)} {responseText}")
                    End If

                    Dim tokenJson = JObject.Parse(responseText)
                    Dim accessToken As String = If(tokenJson("access_token")?.ToString(), "")
                    If String.IsNullOrWhiteSpace(accessToken) Then
                        Throw New Exception("OAuth token response did not contain access_token.")
                    End If

                    Dim expiresIn As Integer = 3600
                    If tokenJson("expires_in") IsNot Nothing Then
                        Integer.TryParse(tokenJson("expires_in").ToString(), expiresIn)
                    End If

                    Return New MCPProtectedResourceOAuthResult With {
                        .AccessToken = accessToken,
                        .ClientId = clientId,
                        .ClientSecret = clientSecret,
                        .AuthorizationEndpoint = authorizationEndpoint,
                        .TokenEndpoint = tokenEndpoint,
                        .Scope = scope,
                        .Resource = resource,
                        .ExpiresIn = expiresIn
                    }
                End Using
            End Using
        End Function

        Private Shared ReadOnly _mcpOAuthLogLock As New Object()

        Private Shared Sub LogMCPOAuth(label As String, content As String)
            If Not _mcpOAuthDebugEnabled Then
                Return
            End If

            Try
                Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                Dim path As String = IO.Path.Combine(desktopPath, "RI_MCPOAuth_Log.txt")
                Dim stamp As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")
                Dim entry As String = $"[{stamp}] {label}{Environment.NewLine}{If(content, "(null)")}{Environment.NewLine}{New String("-"c, 60)}{Environment.NewLine}"

                SyncLock _mcpOAuthLogLock
                    IO.File.AppendAllText(path, entry, Encoding.UTF8)
                End SyncLock
            Catch
                ' Never let logging itself fail OAuth
            End Try
        End Sub

        Public Shared Function IsMCPRedirectCallback(callback As System.String, expected As System.String) As System.Boolean
            Dim actualUri As System.Uri = Nothing
            Dim expectedUri As System.Uri = Nothing
            If Not System.Uri.TryCreate(callback, System.UriKind.Absolute, actualUri) OrElse
               Not System.Uri.TryCreate(expected, System.UriKind.Absolute, expectedUri) Then Return False
            Return System.String.Equals(actualUri.Scheme, expectedUri.Scheme, System.StringComparison.OrdinalIgnoreCase) AndAlso
                System.String.Equals(actualUri.Host, expectedUri.Host, System.StringComparison.OrdinalIgnoreCase) AndAlso
                actualUri.Port = expectedUri.Port AndAlso actualUri.UserInfo.Length = 0 AndAlso actualUri.Fragment.Length = 0 AndAlso
                System.String.Equals(actualUri.AbsolutePath, expectedUri.AbsolutePath, System.StringComparison.Ordinal)
        End Function

        Private Shared Function GenerateOAuthRandomValue() As String
            Dim bytes(31) As Byte
            Using rng = RandomNumberGenerator.Create()
                rng.GetBytes(bytes)
            End Using
            Return OAuthBase64UrlEncode(bytes)
        End Function

        Private Shared Function CreateOAuthCodeChallenge(codeVerifier As String) As String
            Using sha = SHA256.Create()
                Return OAuthBase64UrlEncode(sha.ComputeHash(Encoding.ASCII.GetBytes(codeVerifier)))
            End Using
        End Function

        Private Shared Function OAuthBase64UrlEncode(bytes As Byte()) As String
            Return System.Convert.ToBase64String(bytes).TrimEnd("="c).Replace("+"c, "-"c).Replace("/"c, "_"c)
        End Function

        Private Shared Function GetFreeLoopbackPort() As Integer
            Dim listener As New TcpListener(IPAddress.Loopback, 0)
            listener.Start()
            Dim port = DirectCast(listener.LocalEndpoint, IPEndPoint).Port
            listener.Stop()
            Return port
        End Function

    End Class

End Namespace
