' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SharedMethods.GoogleOAuthHelper.vb
' Purpose: Builds and signs a Google-style OAuth 2.0 JWT bearer assertion
'          (RS256) and exchanges it for an access token at the configured
'          token endpoint.
'
' Architecture:
'  - Configuration inputs:
'      - `client_email`, `private_key`, `scopes`, `token_uri`, `token_life`
'  - JWT generation:
'      - Builds a compact JWT with standard header and OAuth assertion claims.
'      - Uses Base64Url encoding for header, payload, and signature.
'  - Signing:
'      - Parses PEM RSA keys via BouncyCastle.
'      - Signs the assertion using SHA256 with RSA.
'  - Token exchange:
'      - Posts the signed assertion to the configured token endpoint.
'      - Returns the received `access_token` from the JSON response.
'
' Notes:
'  - Intended for service-account style OAuth flows used by shared LLM helpers.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Text
Imports System.IO
Imports System.Net.Http
Imports Newtonsoft.Json
Imports Org.BouncyCastle.Crypto.Parameters
Imports Org.BouncyCastle.Security

Namespace SharedLibrary
    Partial Public Class SharedMethods

        ''' <summary>
        ''' Helper for generating an RS256-signed JWT assertion and exchanging it for an OAuth access token.
        ''' </summary>
        Public Class GoogleOAuthHelper
            ' Public variables

            ''' <summary>
            ''' Service account email used as the JWT issuer (`iss`).
            ''' </summary>
            Public Shared client_email As String = ""

            ''' <summary>
            ''' PEM-encoded RSA private key used to sign the JWT (expected to be readable by BouncyCastle's PEM reader).
            ''' </summary>
            Public Shared private_key As String = ""

            ''' <summary>
            ''' OAuth scope string placed into the JWT payload (`scope`).
            ''' </summary>
            Public Shared scopes As String = ""

            ''' <summary>
            ''' OAuth token endpoint URI used as audience (`aud`) and POST destination for token exchange.
            ''' </summary>
            Public Shared token_uri As String = ""

            ''' <summary>
            ''' Token lifetime in seconds.
            ''' </summary>
            ''' <remarks>
            ''' This value is currently not used by `GenerateJWT`, which uses a fixed 3600 second expiry.
            ''' </remarks>
            Public Shared token_life As Long = 0

            ' Base64Url encoding

            ''' <summary>
            ''' Base64Url-encodes a UTF-8 string (no padding) per JWT requirements.
            ''' </summary>
            ''' <param name="input">Text to encode using UTF-8.</param>
            ''' <returns>Base64Url-encoded string without padding characters.</returns>
            Private Shared Function Base64UrlEncode(input As String) As String
                Return System.Convert.ToBase64String(Encoding.UTF8.GetBytes(input)).
                Replace("+", "-").
                Replace("/", "_").
                Replace("=", "")
            End Function

            ''' <summary>
            ''' Base64Url-encodes a byte array (no padding) per JWT requirements.
            ''' </summary>
            ''' <param name="inputBytes">Bytes to encode.</param>
            ''' <returns>Base64Url-encoded string without padding characters.</returns>
            Private Shared Function Base64UrlEncode(inputBytes As Byte()) As String
                Return System.Convert.ToBase64String(inputBytes).
                Replace("+", "-").
                Replace("/", "_").
                Replace("=", "")
            End Function

            ' Sign data using BouncyCastle

            ''' <summary>
            ''' Signs the provided data with the configured RSA private key using SHA256withRSA (RS256).
            ''' </summary>
            ''' <param name="data">Data to sign.</param>
            ''' <returns>Signature bytes.</returns>
            Private Shared Function SignData(data As Byte(), signingKey As System.String) As Byte()
                Dim rsaKey As RsaPrivateCrtKeyParameters

                ' Normalize line endings for BouncyCastle's PEM reader:
                Dim formattedPrivateKey As String = signingKey _
                    .Replace(vbCrLf, vbLf) _
                    .Replace(vbCr, vbLf) _
                    .Replace("\n", vbLf) _
                    .Replace(vbLf, Environment.NewLine)

                Using reader As New StringReader(formattedPrivateKey)
                    Dim pemReader = New Org.BouncyCastle.OpenSsl.PemReader(reader)
                    rsaKey = DirectCast(pemReader.ReadObject(), RsaPrivateCrtKeyParameters)
                End Using

                ' Explicitly specify PKCS1 padding for RS256
                Dim signer = SignerUtilities.GetSigner("SHA256WITHRSAENCRYPTION")
                signer.Init(True, rsaKey)
                signer.BlockUpdate(data, 0, data.Length)
                Return signer.GenerateSignature()
            End Function

            ' Generate JWT
            ''' <summary>
            ''' Generates a compact serialized JWT signed with RS256 containing `iss`, `scope`, `aud`, `exp`, and `iat`.
            ''' </summary>
            ''' <returns>Compact JWT string (`Base64Url(header).Base64Url(payload).Base64Url(signature)`).</returns>
            Public Shared Function GenerateJWT() As String
                Return GenerateJWT(client_email, private_key, scopes, token_uri, token_life)
            End Function

            Public Shared Function GenerateJWT(clientEmail As System.String, signingKey As System.String,
                                               requestedScopes As System.String, tokenEndpoint As System.String,
                                               lifetime As System.Int64) As System.String
                Dim issuedAt As Long = DateTimeOffset.UtcNow.ToUnixTimeSeconds()
                Dim lifetimeSeconds As Long = If(lifetime > 0, lifetime, 3600)
                Dim expiry As Long = issuedAt + lifetimeSeconds

                Dim header = New With {.alg = "RS256", .typ = "JWT"}
                Dim payload = New With {
                                        .iss = clientEmail,
                                        .scope = requestedScopes,
                                        .aud = tokenEndpoint,
                                        .exp = expiry,
                                        .iat = issuedAt
                                    }

                Dim headerBase64 = Base64UrlEncode(JsonConvert.SerializeObject(header))
                Dim payloadBase64 = Base64UrlEncode(JsonConvert.SerializeObject(payload))
                Dim unsignedToken = $"{headerBase64}.{payloadBase64}"
                Dim signature = SignData(Encoding.UTF8.GetBytes(unsignedToken), signingKey)
                Dim signatureBase64 = Base64UrlEncode(signature)

                Return $"{unsignedToken}.{signatureBase64}"
            End Function


            ' Get Access Token

            ''' <summary>
            ''' Requests an OAuth access token by exchanging a signed JWT assertion at the configured token endpoint.
            ''' </summary>
            ''' <returns>Access token string on success; otherwise an empty string.</returns>
            Public Shared Async Function GetAccessToken() As System.Threading.Tasks.Task(Of System.String)
                Return Await GetAccessToken(client_email, private_key, scopes, token_uri, token_life, False).ConfigureAwait(False)
            End Function

            ' The common implementation takes immutable per-call inputs. Isolated callers never
            ' assign the public legacy credential fields, even while another chat refreshes OAuth.
            Public Shared Async Function GetAccessToken(clientEmail As System.String, signingKey As System.String,
                                                        requestedScopes As System.String, tokenEndpoint As System.String,
                                                        lifetime As System.Int64, silent As System.Boolean,
                                                        Optional cancellationToken As System.Threading.CancellationToken = Nothing) As System.Threading.Tasks.Task(Of System.String)
                cancellationToken.ThrowIfCancellationRequested()
                Try
                    ' Validate configuration before attempting request
                    If String.IsNullOrWhiteSpace(clientEmail) Then
                        ReportOAuthFailure("OAuth configuration error: client_email is not configured.", silent)
                        Return ""
                    End If

                    If String.IsNullOrWhiteSpace(signingKey) Then
                        ReportOAuthFailure("OAuth configuration error: private_key is not configured.", silent)
                        Return ""
                    End If

                    If String.IsNullOrWhiteSpace(tokenEndpoint) Then
                        ReportOAuthFailure("OAuth configuration error: token_uri is not configured.", silent)
                        Return ""
                    End If

                    Dim jwt As String
                    Try
                        jwt = GenerateJWT(clientEmail, signingKey, requestedScopes, tokenEndpoint, lifetime)
                    Catch ex As System.Exception
                        ReportOAuthFailure($"Error generating OAuth JWT token:{vbCrLf}{vbCrLf}" &
                                           $"This usually indicates a problem with the private key format.{vbCrLf}{vbCrLf}" &
                                           $"Details: {ex.Message}", silent)
                        Return ""
                    End Try

                    ' Google's token endpoint expects form-encoded data, not JSON.
                    Dim formData As New Dictionary(Of String, String) From {
                        {"grant_type", "urn:ietf:params:oauth:grant-type:jwt-bearer"},
                        {"assertion", jwt}
                    }

                    Using client As New HttpClient()
                        client.Timeout = TimeSpan.FromSeconds(30)

                        Dim content As New FormUrlEncodedContent(formData)
                        Dim response = Await client.PostAsync(tokenEndpoint, content, cancellationToken)

                        Dim responseBody = Await response.Content.ReadAsStringAsync()

                        If response.IsSuccessStatusCode Then
                            Try
                                Dim tokenData = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(responseBody)
                                If tokenData IsNot Nothing AndAlso tokenData.ContainsKey("access_token") Then
                                    Return tokenData("access_token")?.ToString()
                                Else
                                    ReportOAuthFailure("OAuth error: The token response did not contain an access_token.", silent)
                                    Return ""
                                End If
                            Catch ex As System.Exception
                                ReportOAuthFailure($"OAuth error: Failed to parse token response.{vbCrLf}{vbCrLf}Details: {ex.Message}", silent)
                                Return ""
                            End Try
                        Else
                            ' Try to extract error details from Google's error response
                            Dim errorMessage = BuildOAuthErrorMessage(response.StatusCode, response.ReasonPhrase, responseBody)
                            ReportOAuthFailure(errorMessage, silent)
                            Return ""
                        End If
                    End Using

                Catch ex As System.OperationCanceledException When cancellationToken.IsCancellationRequested
                    Throw

                Catch ex As System.Net.Http.HttpRequestException
                    ReportOAuthFailure($"Network error while requesting OAuth token:{vbCrLf}{vbCrLf}" &
                                       $"Unable to connect to the authentication server. Please check your internet connection.{vbCrLf}{vbCrLf}" &
                                       $"Details: {ex.Message}", silent)
                    Return ""

                Catch ex As System.Threading.Tasks.TaskCanceledException
                    ReportOAuthFailure("OAuth request timed out." & vbCrLf & vbCrLf &
                                       "The authentication server did not respond in time. Please try again later.", silent)
                    Return ""

                Catch ex As System.Exception
                    ReportOAuthFailure($"Unexpected error during OAuth authentication:{vbCrLf}{vbCrLf}{ex.Message}", silent)
                    Return ""
                End Try
            End Function

            Private Shared Sub ReportOAuthFailure(message As System.String, silent As System.Boolean)
                If silent Then
                    Throw New System.InvalidOperationException("The isolated OAuth token request failed; verify the configured credentials and authentication endpoint.")
                End If
                ShowCustomMessageBox(message)
            End Sub

            ''' <summary>
            ''' Builds a user-friendly error message from an OAuth error response.
            ''' </summary>
            ''' <param name="statusCode">HTTP status code.</param>
            ''' <param name="reasonPhrase">HTTP reason phrase.</param>
            ''' <param name="responseBody">Response body that may contain JSON error details.</param>
            ''' <returns>Formatted error message for display.</returns>
            Private Shared Function BuildOAuthErrorMessage(statusCode As Net.HttpStatusCode,
                                                           reasonPhrase As String,
                                                           responseBody As String) As String
                Dim sb As New StringBuilder()
                sb.AppendLine("OAuth authentication failed.")
                sb.AppendLine()

                ' Try to parse Google's error response format
                Dim errorDescription As String = Nothing
                Dim errorCode As String = Nothing

                Try
                    If Not String.IsNullOrWhiteSpace(responseBody) Then
                        Dim errorData = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(responseBody)
                        If errorData IsNot Nothing Then
                            If errorData.ContainsKey("error") Then
                                errorCode = errorData("error")?.ToString()
                            End If
                            If errorData.ContainsKey("error_description") Then
                                errorDescription = errorData("error_description")?.ToString()
                            End If
                        End If
                    End If
                Catch
                    ' Ignore parse errors - we'll fall back to generic message
                End Try

                ' Provide context based on status code
                Select Case CInt(statusCode)
                    Case 400
                        sb.AppendLine("The authentication request was invalid.")
                        If Not String.IsNullOrWhiteSpace(errorDescription) Then
                            sb.AppendLine($"Reason: {errorDescription}")
                        Else
                            sb.AppendLine("This may indicate an issue with the service account configuration.")
                        End If

                    Case 401
                        sb.AppendLine("Authentication credentials are invalid or expired.")
                        sb.AppendLine("Please verify your service account credentials are correct.")

                    Case 403
                        sb.AppendLine("Access denied.")
                        sb.AppendLine("The service account may not have the required permissions, or the API may not be enabled.")

                    Case 404
                        sb.AppendLine("The authentication endpoint was not found.")
                        sb.AppendLine("Please verify the token_uri configuration.")

                    Case 429
                        sb.AppendLine("Too many authentication requests.")
                        sb.AppendLine("Please wait a moment and try again.")

                    Case >= 500
                        sb.AppendLine("The authentication server is temporarily unavailable.")
                        sb.AppendLine("This is usually a temporary issue. Please try again later.")

                    Case Else
                        sb.AppendLine($"HTTP {CInt(statusCode)}: {reasonPhrase}")
                        If Not String.IsNullOrWhiteSpace(errorDescription) Then
                            sb.AppendLine($"Details: {errorDescription}")
                        End If
                End Select

                ' Add error code if available
                If Not String.IsNullOrWhiteSpace(errorCode) AndAlso errorDescription Is Nothing Then
                    sb.AppendLine()
                    sb.AppendLine($"Error code: {errorCode}")
                End If

                Return sb.ToString()
            End Function
        End Class

    End Class

End Namespace