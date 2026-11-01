' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SharedMethods.License.Domain.vb
' Purpose: Implements signed offline-domain Professional license keys that can be
'          verified locally without any server call.
'
' Architecture:
'  - Offline key format:
'     - Prefix: `RI-DOM-1:`
'     - Payload contains:
'         - marker
'         - Product ID
'         - UTC end date
'         - one or more allowed network/domain identifiers
'     - Payload is signed with Ed25519 using the same key format and signature
'       approach already used by `SharedMethods.UpdateIni.vb`.
'  - Verification flow:
'     - Detects the special offline-domain key format.
'     - Verifies the Ed25519 signature using a hardcoded public key.
'     - Verifies Product ID, expiry date, and whether at least one local network
'       identifier matches an embedded identifier exactly or, for DNS-style
'       domains, is a true child domain of that identifier.
'     - Returns a synthetic `LicenseApiResponse` so the existing Pro-license
'       pipeline can continue to work with minimal changes.
'  - Generator UI:
'     - Provides a DPI-aware WinForms dialog for generating offline-domain keys
'       from Product ID, allowed identifiers, end date, and a supplied private key.
'     - Shows current local network identifiers to simplify key creation.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Globalization
Imports System.Net
Imports System.Net.NetworkInformation
Imports System.Text
Imports System.Windows.Forms
Imports Org.BouncyCastle.Crypto.Parameters
Imports Org.BouncyCastle.Crypto.Signers

Namespace SharedLibrary
    Partial Public Class SharedMethods

#Region "Offline Domain License"

        Private Const OfflineDomainLicensePrefix As String = "RI-DOM-1:"
        Private Const OfflineDomainLicensePayloadMarker As String = "RI-DOM-1"
        Private Const OfflineDomainLicenseSyntheticUserIdValue As String = "offline-domain-license"

        ' IMPORTANT:
        ' Replace this once with the Base64 public key matching the private key used for generation.
        Private Const OfflineDomainLicensePublicKeyBase64 As String = "KP2qbVGWKdOZLjF1CAcLawzf/kSEtj0KT1IWWv//Jlo="

        Private Class OfflineDomainLicenseInfo
            Public Property ProductId As String = ""
            Public Property ValidUntilUtc As Date = Date.MinValue
            Public Property AllowedDomains As New List(Of String)()
            Public Property MatchedDomain As String = ""
            Public Property CanonicalPayload As String = ""
        End Class

        Public Shared Function IsOfflineDomainLicenseKey(licenseKey As String) As Boolean
            Return Not String.IsNullOrWhiteSpace(licenseKey) AndAlso
                   licenseKey.Trim().StartsWith(OfflineDomainLicensePrefix, StringComparison.Ordinal)
        End Function

        Private Shared Function GetOfflineDomainLicenseSyntheticUserId() As String
            Return OfflineDomainLicenseSyntheticUserIdValue
        End Function

        Private Shared Function TryExtractOfflineDomainLicenseProductId(licenseKey As String, ByRef productId As String) As Boolean
            productId = ""

            Try
                Dim payloadBase64Url As String = ""
                Dim signatureBase64Url As String = ""

                If Not TrySplitOfflineDomainLicenseKey(licenseKey, payloadBase64Url, signatureBase64Url) Then
                    Return False
                End If

                Dim payloadBytes = Base64UrlDecode(payloadBase64Url)
                Dim payloadText = Encoding.UTF8.GetString(payloadBytes)
                Dim parts = payloadText.Split("|"c)

                If parts.Length <> 4 Then Return False
                If Not parts(0).Equals(OfflineDomainLicensePayloadMarker, StringComparison.Ordinal) Then Return False

                productId = parts(1).Trim()
                Return Not String.IsNullOrWhiteSpace(productId)

            Catch
                Return False
            End Try
        End Function

        Public Shared Function GenerateOfflineDomainLicenseKey(productId As String,
                                                               allowedDomains As String,
                                                               validUntil As Date,
                                                               base64PrivateKey As String) As String
            Return GenerateOfflineDomainLicenseKeyInternal(productId, ParseOfflineDomainList(allowedDomains), validUntil, base64PrivateKey)
        End Function

        Private Shared Function GenerateOfflineDomainLicenseKeyInternal(productId As String,
                                                                       allowedDomains As List(Of String),
                                                                       validUntil As Date,
                                                                       base64PrivateKey As String) As String
            If String.IsNullOrWhiteSpace(productId) Then
                Throw New ArgumentException("Product ID is required.", NameOf(productId))
            End If

            If allowedDomains Is Nothing OrElse allowedDomains.Count = 0 Then
                Throw New ArgumentException("At least one allowed domain / network identifier is required.", NameOf(allowedDomains))
            End If

            If String.IsNullOrWhiteSpace(base64PrivateKey) Then
                Throw New ArgumentException("Private key is required.", NameOf(base64PrivateKey))
            End If

            Dim normalizedProductId = productId.Trim()
            Dim normalizedDomains As New List(Of String)(allowedDomains)
            Dim canonicalPayload = $"{OfflineDomainLicensePayloadMarker}|{normalizedProductId}|{validUntil.ToUniversalTime().ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)}|{String.Join(";", normalizedDomains)}"
            Dim payloadBytes = Encoding.UTF8.GetBytes(canonicalPayload)

            Dim privateKeyBytes = System.Convert.FromBase64String(base64PrivateKey.Trim())
            Dim privateKey As New Ed25519PrivateKeyParameters(privateKeyBytes, 0)

            Dim signer As New Ed25519Signer()
            signer.Init(True, privateKey)
            signer.BlockUpdate(payloadBytes, 0, payloadBytes.Length)
            Dim signatureBytes = signer.GenerateSignature()

            If Not IsOfflineDomainLicensePublicKeyConfigured() Then
                Throw New InvalidOperationException("The hardcoded offline-domain public key has not yet been configured.")
            End If

            If Not VerifyEd25519Signature(payloadBytes, System.Convert.ToBase64String(signatureBytes), OfflineDomainLicensePublicKeyBase64) Then
                Throw New InvalidOperationException("The supplied private key does not match the hardcoded offline-domain public key.")
            End If

            Return OfflineDomainLicensePrefix &
                   Base64UrlEncode(payloadBytes) &
                   "." &
                   Base64UrlEncode(signatureBytes)
        End Function

        Private Shared Function TryCreateOfflineDomainLicenseResponse(action As String,
                                                                      productId As String,
                                                                      licenseKey As String,
                                                                      userId As String,
                                                                      ByRef response As LicenseApiResponse) As Boolean
            response = Nothing

            If Not IsOfflineDomainLicenseKey(licenseKey) Then
                Return False
            End If

            response = New LicenseApiResponse() With {
                .ProductTitle = "Offline Domain License",
                .ProductId = productId,
                .RawJson = "{""mode"":""offline-domain""}"
            }

            If action.Equals("deactivate", StringComparison.OrdinalIgnoreCase) Then
                response.Success = True
                response.Activated = False
                response.RawJson = "{""mode"":""offline-domain"",""deactivated"":true}"
                Return True
            End If

            Dim info As OfflineDomainLicenseInfo = Nothing
            Dim failureReason As String = ""

            If Not TryParseAndVerifyOfflineDomainLicense(productId, licenseKey, info, failureReason) Then
                response.Success = False
                response.Activated = False
                response.StatusCheck = "invalid"
                response.ErrorMessage = failureReason
                response.RawJson = "{""mode"":""offline-domain"",""success"":false}"
                Return True
            End If

            response.Success = True
            response.Activated = True
            response.StatusCheck = "active"
            response.ProductTitle = $"Offline Domain License ({info.MatchedDomain})"
            response.ProductId = info.ProductId

            ' Offline-domain licenses are locally verified and are not tied to
            ' server-side activation-slot accounting.
            response.TotalActivations = 0
            response.TotalActivationsPurchased = 0
            response.ActivationsRemaining = 0

            response.RawJson = "{""mode"":""offline-domain"",""success"":true}"

            Return True
        End Function

        Private Shared Function TryParseAndVerifyOfflineDomainLicense(expectedProductId As String,
                                                                      licenseKey As String,
                                                                      ByRef info As OfflineDomainLicenseInfo,
                                                                      ByRef failureReason As String) As Boolean
            info = Nothing
            failureReason = ""

            Try
                If Not IsOfflineDomainLicensePublicKeyConfigured() Then
                    failureReason = "Offline-domain license verification public key has not been configured."
                    Return False
                End If

                Dim payloadBase64Url As String = ""
                Dim signatureBase64Url As String = ""

                If Not TrySplitOfflineDomainLicenseKey(licenseKey, payloadBase64Url, signatureBase64Url) Then
                    failureReason = "The offline-domain license key format is invalid."
                    Return False
                End If

                Dim payloadBytes = Base64UrlDecode(payloadBase64Url)
                Dim signatureBytes = Base64UrlDecode(signatureBase64Url)

                If Not VerifyEd25519Signature(payloadBytes, System.Convert.ToBase64String(signatureBytes), OfflineDomainLicensePublicKeyBase64) Then
                    failureReason = "The offline-domain license signature is invalid."
                    Return False
                End If

                Dim payloadText = Encoding.UTF8.GetString(payloadBytes)
                Dim parts = payloadText.Split("|"c)

                If parts.Length <> 4 Then
                    failureReason = "The offline-domain license payload is invalid."
                    Return False
                End If

                If Not parts(0).Equals(OfflineDomainLicensePayloadMarker, StringComparison.Ordinal) Then
                    failureReason = "The offline-domain license payload marker is invalid."
                    Return False
                End If

                Dim parsedProductId = parts(1).Trim()
                If String.IsNullOrWhiteSpace(parsedProductId) Then
                    failureReason = "The offline-domain license does not contain a Product ID."
                    Return False
                End If

                If Not String.IsNullOrWhiteSpace(expectedProductId) AndAlso
                   Not parsedProductId.Equals(expectedProductId.Trim(), StringComparison.Ordinal) Then
                    failureReason = $"The offline-domain license Product ID '{parsedProductId}' does not match the configured Product ID '{expectedProductId}'."
                    Return False
                End If

                Dim validUntilUtc As Date
                If Not Date.TryParseExact(parts(2).Trim(),
                                          "yyyy-MM-dd",
                                          CultureInfo.InvariantCulture,
                                          DateTimeStyles.AssumeUniversal Or DateTimeStyles.AdjustToUniversal,
                                          validUntilUtc) Then
                    failureReason = "The offline-domain license end date is invalid."
                    Return False
                End If

                If Date.UtcNow.Date > validUntilUtc.Date Then
                    failureReason = $"The offline-domain license expired on {validUntilUtc:yyyy-MM-dd} UTC."
                    Return False
                End If

                Dim allowedDomains = ParseOfflineDomainList(parts(3))
                If allowedDomains.Count = 0 Then
                    failureReason = "The offline-domain license does not contain any allowed domains."
                    Return False
                End If

                Dim needsTenant As System.Boolean = allowedDomains.Exists(
                    Function(value As System.String) value.StartsWith(OfflineTenantIdentifierPrefix, System.StringComparison.OrdinalIgnoreCase))
                Dim currentCandidates = GetCurrentOfflineLicenseIdentitySnapshot(allowedDomains, includeTenant:=needsTenant).Identifiers
                Dim matchedDomain As String = ""

                For Each candidate In currentCandidates
                    For Each allowedDomain In allowedDomains
                        If IsOfflineDomainIdentifierMatch(candidate, allowedDomain) Then
                            matchedDomain = candidate
                            Exit For
                        End If
                    Next

                    If Not String.IsNullOrWhiteSpace(matchedDomain) Then
                        Exit For
                    End If
                Next

                If String.IsNullOrWhiteSpace(matchedDomain) Then
                    failureReason = "The offline-domain license is validly signed, but none of the local network identifiers match." &
                                    vbCrLf & vbCrLf &
                                    $"Allowed identifiers: {String.Join(", ", allowedDomains)}" &
                                    vbCrLf &
                                    $"Local identifiers: {String.Join(", ", currentCandidates)}"
                    Return False
                End If

                info = New OfflineDomainLicenseInfo() With {
                    .ProductId = parsedProductId,
                    .ValidUntilUtc = validUntilUtc,
                    .AllowedDomains = allowedDomains,
                    .MatchedDomain = matchedDomain,
                    .CanonicalPayload = payloadText
                }

                Return True

            Catch ex As System.Exception
                failureReason = $"Failed to parse the offline-domain license: {ex.Message}"
                Return False
            End Try
        End Function

        Private Shared Function ParseOfflineDomainList(value As String) As List(Of String)
            Dim results As New List(Of String)()

            If String.IsNullOrWhiteSpace(value) Then
                Return results
            End If

            Dim flattened = value.Replace(vbCrLf, vbLf).
                                  Replace(vbCr, vbLf).
                                  Replace(";", vbLf).
                                  Replace(",", vbLf)

            For Each rawEntry In flattened.Split({ControlChars.Lf}, StringSplitOptions.RemoveEmptyEntries)
                Dim typedInput As System.String = rawEntry.Trim()
                If typedInput.StartsWith(OfflineTenantIdentifierPrefix, System.StringComparison.OrdinalIgnoreCase) OrElse
                   typedInput.StartsWith("*." & OfflineTenantIdentifierPrefix, System.StringComparison.OrdinalIgnoreCase) Then
                    Dim canonicalTenant As System.String = System.String.Empty
                    If Not TryNormalizeOfflineTenantIdentifier(typedInput, canonicalTenant) Then
                        Throw New System.FormatException("Tenant IDs must be tenant:<non-empty GUID>, without a wildcard or domain suffix.")
                    End If
                End If
                AddNormalizedDomain(results, rawEntry)
            Next

            Return results
        End Function

        Private Shared Function GetCurrentNetworkDomainCandidates() As List(Of String)
            Dim results As New List(Of String)()

            Try
                Dim ipProperties = IPGlobalProperties.GetIPGlobalProperties()
                If ipProperties IsNot Nothing Then
                    AddNormalizedDomain(results, ipProperties.DomainName)
                    AddNormalizedDomain(results, ipProperties.HostName)
                End If
            Catch
            End Try

            Try
                AddNormalizedDomain(results, Environment.GetEnvironmentVariable("USERDNSDOMAIN"))
            Catch
            End Try

            Try
                AddNormalizedDomain(results, Environment.UserDomainName)
            Catch
            End Try

            Try
                Dim hostEntry = Dns.GetHostEntry(Dns.GetHostName())
                If hostEntry IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(hostEntry.HostName) Then
                    Dim hostName = hostEntry.HostName.Trim()
                    Dim dotIndex = hostName.IndexOf("."c)
                    If dotIndex >= 0 AndAlso dotIndex < hostName.Length - 1 Then
                        AddNormalizedDomain(results, hostName.Substring(dotIndex + 1))
                    End If
                End If
            Catch
            End Try

            If results.Count = 0 Then
                results.Add("unknown")
            End If

            Return results
        End Function

        Private Shared Sub AddNormalizedDomain(results As List(Of String), rawValue As String)
            Dim normalized = NormalizeOfflineDomainIdentifier(rawValue)
            If String.IsNullOrWhiteSpace(normalized) Then Return

            For Each existing In results
                If existing.Equals(normalized, StringComparison.OrdinalIgnoreCase) Then
                    Return
                End If
            Next

            results.Add(normalized)
        End Sub

        Private Shared Function NormalizeOfflineDomainIdentifier(value As String) As String
            If System.String.IsNullOrWhiteSpace(value) Then
                Return ""
            End If

            Dim normalized As System.String = value.Trim().ToLowerInvariant()
            If normalized.StartsWith(OfflineTenantIdentifierPrefix, System.StringComparison.Ordinal) Then
                Dim canonicalTenant As System.String = System.String.Empty
                If TryNormalizeOfflineTenantIdentifier(normalized, canonicalTenant) Then Return canonicalTenant
                Return System.String.Empty
            End If

            ' Accept the common wildcard notation as input, but store and compare
            ' the canonical DNS suffix.  The match boundary is enforced separately.
            If normalized.StartsWith("*.", System.StringComparison.Ordinal) Then
                normalized = normalized.Substring(2)
            End If

            normalized = normalized.Trim("."c)

            If System.String.IsNullOrWhiteSpace(normalized) OrElse normalized.IndexOf("*"c) >= 0 Then
                Return ""
            End If

            Return normalized
        End Function

        Private Shared Function IsOfflineDomainIdentifierMatch(candidate As System.String,
                                                                 allowedDomain As System.String) As System.Boolean
            If System.String.IsNullOrWhiteSpace(candidate) OrElse System.String.IsNullOrWhiteSpace(allowedDomain) Then
                Return False
            End If

            ' Typed tenant identities can only match an identical native tenant GUID.
            ' They never enter the legacy DNS suffix matching path.
            If candidate.StartsWith(OfflineTenantIdentifierPrefix, System.StringComparison.OrdinalIgnoreCase) OrElse
               allowedDomain.StartsWith(OfflineTenantIdentifierPrefix, System.StringComparison.OrdinalIgnoreCase) Then
                Dim canonicalCandidate As System.String = System.String.Empty
                Dim canonicalAllowed As System.String = System.String.Empty
                Return TryNormalizeOfflineTenantIdentifier(candidate, canonicalCandidate) AndAlso
                       TryNormalizeOfflineTenantIdentifier(allowedDomain, canonicalAllowed) AndAlso
                       System.String.Equals(canonicalCandidate, canonicalAllowed, System.StringComparison.Ordinal)
            End If

            If candidate.Equals(allowedDomain, System.StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If

            ' Every allowed identifier covers itself and true child domains,
            ' including single-label suffixes such as "abc" in "xy.cd.wa.abc".
            ' Require a complete trailing label boundary; arbitrary substrings
            ' such as "notabc" or "abc.example" must not match "abc".
            Return candidate.EndsWith("." & allowedDomain, System.StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function IsOfflineDomainLicensePublicKeyConfigured() As Boolean
            Return Not String.IsNullOrWhiteSpace(OfflineDomainLicensePublicKeyBase64) AndAlso
                   OfflineDomainLicensePublicKeyBase64.IndexOf("PASTE_YOUR_", StringComparison.OrdinalIgnoreCase) < 0
        End Function

        Private Shared Function TrySplitOfflineDomainLicenseKey(licenseKey As String,
                                                                ByRef payloadBase64Url As String,
                                                                ByRef signatureBase64Url As String) As Boolean
            payloadBase64Url = ""
            signatureBase64Url = ""

            If Not IsOfflineDomainLicenseKey(licenseKey) Then
                Return False
            End If

            Dim remainder = licenseKey.Trim().Substring(OfflineDomainLicensePrefix.Length)
            Dim separatorIndex = remainder.IndexOf("."c)

            If separatorIndex <= 0 OrElse separatorIndex >= remainder.Length - 1 Then
                Return False
            End If

            payloadBase64Url = remainder.Substring(0, separatorIndex)
            signatureBase64Url = remainder.Substring(separatorIndex + 1)

            Return Not String.IsNullOrWhiteSpace(payloadBase64Url) AndAlso
                   Not String.IsNullOrWhiteSpace(signatureBase64Url)
        End Function

        Private Shared Function Base64UrlEncode(value As Byte()) As String
            Dim base64 = System.Convert.ToBase64String(value)
            Return base64.TrimEnd("="c).Replace("+"c, "-"c).Replace("/"c, "_"c)
        End Function

        Private Shared Function Base64UrlDecode(value As String) As Byte()
            Dim base64 = value.Replace("-"c, "+"c).Replace("_"c, "/"c)

            Select Case base64.Length Mod 4
                Case 0
                Case 2
                    base64 &= "=="
                Case 3
                    base64 &= "="
                Case Else
                    Throw New FormatException("Invalid Base64Url value.")
            End Select

            Return System.Convert.FromBase64String(base64)
        End Function

#End Region

#Region "Offline Domain License Generator UI"

        Private Shared Function TryReadOfflineDomainLicenseMetadata(licenseKey As String,
                                                                    ByRef validUntilUtc As Date,
                                                                    ByRef allowedDomains As List(Of String),
                                                                    ByRef productId As String) As Boolean
            validUntilUtc = Date.MinValue
            allowedDomains = New List(Of String)()
            productId = ""

            Try
                Dim payloadBase64Url As String = ""
                Dim signatureBase64Url As String = ""

                If Not TrySplitOfflineDomainLicenseKey(licenseKey, payloadBase64Url, signatureBase64Url) Then
                    Return False
                End If

                Dim payloadBytes = Base64UrlDecode(payloadBase64Url)
                Dim signatureBytes = Base64UrlDecode(signatureBase64Url)

                If Not VerifyEd25519Signature(payloadBytes, System.Convert.ToBase64String(signatureBytes), OfflineDomainLicensePublicKeyBase64) Then
                    Return False
                End If

                Dim payloadText = Encoding.UTF8.GetString(payloadBytes)
                Dim parts = payloadText.Split("|"c)

                If parts.Length <> 4 Then Return False
                If Not parts(0).Equals(OfflineDomainLicensePayloadMarker, StringComparison.Ordinal) Then Return False

                productId = parts(1).Trim()

                If Not Date.TryParseExact(parts(2).Trim(),
                                          "yyyy-MM-dd",
                                          CultureInfo.InvariantCulture,
                                          DateTimeStyles.AssumeUniversal Or DateTimeStyles.AdjustToUniversal,
                                          validUntilUtc) Then
                    Return False
                End If

                allowedDomains = ParseOfflineDomainList(parts(3))
                Return True

            Catch
                Return False
            End Try
        End Function

        Public Shared Sub ShowOfflineDomainLicenseGeneratorDialog()
            If Not IsOnHostUiThread AndAlso Not IsOnHostNeutralSta Then
                RunOnStaThread(AddressOf ShowOfflineDomainLicenseGeneratorDialog, "Red Ink offline license", alwaysUseWorker:=True)
                Return
            End If
            Using form As New Form(), hints As New System.Windows.Forms.ToolTip() With {
                .AutoPopDelay = 30000, .InitialDelay = 350, .ReshowDelay = 100, .ShowAlways = True}
                form.AutoScaleDimensions = New SizeF(96.0F, 96.0F)
                form.AutoScaleMode = AutoScaleMode.Dpi

                form.Text = $"{AN} - Offline License Manager"
                form.StartPosition = FormStartPosition.CenterScreen
                form.FormBorderStyle = FormBorderStyle.Sizable
                form.MaximizeBox = True
                form.MinimizeBox = False
                form.ShowInTaskbar = True
                form.TopMost = True
                form.MinimumSize = New System.Drawing.Size(840, 540)
                form.Size = New System.Drawing.Size(1160, 820)
                AddHandler form.Load,
                    Sub()
                        Dim area As System.Drawing.Rectangle = System.Windows.Forms.Screen.FromControl(form).WorkingArea
                        form.MinimumSize = New System.Drawing.Size(System.Math.Min(form.MinimumSize.Width, area.Width),
                                                                  System.Math.Min(form.MinimumSize.Height, area.Height))
                        form.Size = New System.Drawing.Size(System.Math.Min(form.Width, area.Width),
                                                           System.Math.Min(form.Height, area.Height))
                    End Sub
                form.Font = New System.Drawing.Font("Segoe UI", 9.5F)
                form.AutoScroll = False

                Try
                    Global.SharedLibrary.SharedLibrary.SharedMethods.ApplyStandardDialogIcon(form)
                Catch
                End Try

                Dim mainLayout As New TableLayoutPanel() With {
                    .AutoScroll = True,
                    .Dock = DockStyle.Fill,
                    .Padding = New Padding(20),
                    .ColumnCount = 2,
                    .RowCount = 8
                }

                mainLayout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
                mainLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))

                For i As Integer = 0 To 7
                    mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
                Next

                ' Keep actions outside the scrollable fields, including on small/high-DPI screens.
                Dim dialogLayout As New System.Windows.Forms.TableLayoutPanel() With {
                    .Dock = System.Windows.Forms.DockStyle.Fill,
                    .ColumnCount = 1, .RowCount = 2, .Margin = New System.Windows.Forms.Padding(0)}
                dialogLayout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
                dialogLayout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
                dialogLayout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
                mainLayout.Margin = New System.Windows.Forms.Padding(0)
                dialogLayout.Controls.Add(mainLayout, 0, 0)
                form.Controls.Add(dialogLayout)

                Dim lblTitle As New Label() With {
                    .Text = "Generate Offline License",
                    .Font = New Font("Segoe UI", 12.0F, FontStyle.Bold),
                    .AutoSize = True,
                    .Margin = New Padding(0, 0, 0, 5)
                }
                mainLayout.Controls.Add(lblTitle, 0, 0)
                mainLayout.SetColumnSpan(lblTitle, 2)

                Dim lblDescription As New Label() With {
                    .Text = "Domains, computer names and Entra tenant IDs. Hover over fields for details.",
                    .AutoSize = True,
                    .MaximumSize = New Size(820, 0),
                    .Margin = New Padding(0, 0, 0, 15)
                }
                mainLayout.Controls.Add(lblDescription, 0, 1)
                mainLayout.SetColumnSpan(lblDescription, 2)
                hints.SetToolTip(lblDescription, "Creates a signed offline license for domains, network/host names, computer/client names and Microsoft Entra (Azure AD) tenant IDs. Domain/network IDs retain their existing exact-or-child-domain matching; tenant:<GUID> matches only that exact tenant. At least ONE allowed identifier must match (OR, not AND). A computer/client ID is the Windows computer name, not a hardware UUID.")

                Dim lblProductId As New Label() With {
                    .Text = "Product ID:",
                    .AutoSize = True,
                    .Anchor = AnchorStyles.Left,
                    .Margin = New Padding(0, 0, 10, 0)
                }
                mainLayout.Controls.Add(lblProductId, 0, 2)

                Dim txtProductId As New TextBox() With {
                    .Text = StoredProductId,
                    .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
                    .Margin = New Padding(0, 3, 0, 3)
                }
                mainLayout.Controls.Add(txtProductId, 1, 2)

                Dim lblValidUntil As New Label() With {
                    .Text = "Valid Until (UTC):",
                    .AutoSize = True,
                    .Anchor = AnchorStyles.Left,
                    .Margin = New Padding(0, 0, 10, 0)
                }
                mainLayout.Controls.Add(lblValidUntil, 0, 3)

                Dim dtpValidUntil As New DateTimePicker() With {
                    .Format = DateTimePickerFormat.Custom,
                    .CustomFormat = "yyyy-MM-dd",
                    .Value = Date.Today.AddYears(1),
                    .Anchor = AnchorStyles.Left,
                    .Margin = New Padding(0, 3, 0, 3),
                    .Width = 160
                }
                mainLayout.Controls.Add(dtpValidUntil, 1, 3)

                Dim lblDomains As New Label() With {
                    .Text = "Allowed IDs:",
                    .AutoSize = True,
                    .Anchor = AnchorStyles.Left Or AnchorStyles.Top,
                    .Margin = New Padding(0, 6, 10, 0)
                }
                mainLayout.Controls.Add(lblDomains, 0, 4)

                Dim txtDomains As New TextBox() With {
                    .Multiline = True,
                    .AcceptsReturn = True,
                    .ScrollBars = ScrollBars.Vertical,
                    .WordWrap = False,
                    .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
                    .Height = 64,
                    .Margin = New Padding(0, 3, 0, 3)
                }
                Dim allowedIdsPanel As New System.Windows.Forms.TableLayoutPanel() With {
                    .Dock = System.Windows.Forms.DockStyle.Fill, .AutoSize = True, .ColumnCount = 1, .RowCount = 2}
                allowedIdsPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
                allowedIdsPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
                allowedIdsPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
                allowedIdsPanel.Controls.Add(txtDomains, 0, 0)
                Dim lblAllowedIdsHelp As New System.Windows.Forms.Label() With {
                    .AutoSize = True, .MaximumSize = New System.Drawing.Size(820, 0),
                    .Text = "One per line, or separated by ; or , — one matching ID is sufficient (OR)."}
                hints.SetToolTip(lblAllowedIdsHelp, "Multiple values: one per line, or separated by semicolons or commas. Example: example.org; OFFICE-PC01; tenant:01234567-89ab-cdef-0123-456789abcdef. Only one matching value is required. Remove IDs that should not independently authorize this license.")
                hints.SetToolTip(txtDomains, "Multiple values: one per line, or separated by semicolons or commas. Example: example.org; OFFICE-PC01; tenant:01234567-89ab-cdef-0123-456789abcdef. Only one matching value is required. Remove IDs that should not independently authorize this license.")
                hints.SetToolTip(lblDomains, "Multiple values: one per line, or separated by semicolons or commas. Example: example.org; OFFICE-PC01; tenant:01234567-89ab-cdef-0123-456789abcdef. Only one matching value is required. Remove IDs that should not independently authorize this license.")
                allowedIdsPanel.Controls.Add(lblAllowedIdsHelp, 0, 1)
                mainLayout.Controls.Add(allowedIdsPanel, 1, 4)

                Dim lblCurrentDomains As New Label() With {
                    .Text = "Detected Local IDs:",
                    .AutoSize = True,
                    .Anchor = AnchorStyles.Left Or AnchorStyles.Top,
                    .Margin = New Padding(0, 6, 10, 0)
                }
                mainLayout.Controls.Add(lblCurrentDomains, 0, 5)

                Dim txtCurrentDomains As New TextBox() With {
                    .Multiline = True,
                    .ReadOnly = True,
                    .ScrollBars = ScrollBars.Vertical,
                    .WordWrap = False,
                    .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
                    .Height = 64,
                    .Margin = New Padding(0, 3, 0, 3)
                }
                Dim detectedIdsPanel As New System.Windows.Forms.TableLayoutPanel() With {
                    .Dock = System.Windows.Forms.DockStyle.Fill, .AutoSize = True, .ColumnCount = 1, .RowCount = 2}
                detectedIdsPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
                detectedIdsPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
                detectedIdsPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
                detectedIdsPanel.Controls.Add(txtCurrentDomains, 0, 0)
                Dim lblDetectedIdsDetails As New System.Windows.Forms.Label() With {
                    .AutoSize = True, .MaximumSize = New System.Drawing.Size(820, 0)}
                detectedIdsPanel.Controls.Add(lblDetectedIdsDetails, 0, 1)
                mainLayout.Controls.Add(detectedIdsPanel, 1, 5)

                Dim lblPrivateKey As New Label() With {
                    .Text = "Private Key (Base64):",
                    .AutoSize = True,
                    .Anchor = AnchorStyles.Left Or AnchorStyles.Top,
                    .Margin = New Padding(0, 6, 10, 0)
                }
                mainLayout.Controls.Add(lblPrivateKey, 0, 6)

                Dim privateKeyPanel As New TableLayoutPanel() With {
                    .AutoSize = True,
                    .Dock = DockStyle.Fill,
                    .ColumnCount = 1,
                    .RowCount = 3,
                    .Margin = New Padding(0, 3, 0, 3)
                }
                privateKeyPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
                privateKeyPanel.RowStyles.Add(New RowStyle(SizeType.AutoSize))
                privateKeyPanel.RowStyles.Add(New RowStyle(SizeType.AutoSize))
                privateKeyPanel.RowStyles.Add(New RowStyle(SizeType.AutoSize))

                Dim txtPrivateKey As New TextBox() With {
                    .Multiline = True,
                    .ScrollBars = ScrollBars.Vertical,
                    .WordWrap = False,
                    .Dock = DockStyle.Fill,
                    .Height = 64,
                    .Margin = New Padding(0, 0, 0, 8)
                }
                privateKeyPanel.Controls.Add(txtPrivateKey, 0, 0)

                Dim lblPublicKey As New Label() With {
                    .Text = "Hardcoded verification public key:",
                    .AutoSize = True,
                    .Margin = New Padding(0, 0, 0, 4)
                }
                privateKeyPanel.Controls.Add(lblPublicKey, 0, 1)

                Dim txtPublicKey As New TextBox() With {
                    .Text = OfflineDomainLicensePublicKeyBase64,
                    .Multiline = True,
                    .ReadOnly = True,
                    .ScrollBars = ScrollBars.Vertical,
                    .WordWrap = False,
                    .Dock = DockStyle.Fill,
                    .Height = 32
                }
                privateKeyPanel.Controls.Add(txtPublicKey, 0, 2)

                mainLayout.Controls.Add(privateKeyPanel, 1, 6)

                Dim lblResult As New Label() With {
                    .Text = "Generated License Key:",
                    .AutoSize = True,
                    .Anchor = AnchorStyles.Left Or AnchorStyles.Top,
                    .Margin = New Padding(0, 6, 10, 0)
                }
                mainLayout.Controls.Add(lblResult, 0, 7)

                Dim resultPanel As New TableLayoutPanel() With {
                    .AutoSize = True,
                    .Dock = DockStyle.Fill,
                    .ColumnCount = 1,
                    .RowCount = 1,
                    .Margin = New Padding(0, 3, 0, 3)
                }
                resultPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
                resultPanel.RowStyles.Add(New RowStyle(SizeType.AutoSize))

                Dim txtLicenseKey As New TextBox() With {
                    .Multiline = True,
                    .ReadOnly = True,
                    .ScrollBars = ScrollBars.Vertical,
                    .WordWrap = False,
                    .Dock = DockStyle.Fill,
                    .Height = 64,
                    .Margin = New Padding(0, 0, 0, 10)
                }
                resultPanel.Controls.Add(txtLicenseKey, 0, 0)

                Dim buttonPanel As New FlowLayoutPanel() With {
                    .AutoSize = True,
                    .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    .FlowDirection = FlowDirection.LeftToRight,
                    .Anchor = AnchorStyles.Left,
                    .WrapContents = True,
                    .Margin = New Padding(0)
                }

                Dim btnUseCurrentDomains As New Button() With {
                    .Text = "Use Current IDs",
                    .AutoSize = True,
                    .Margin = New Padding(0, 0, 8, 0)
                }
                buttonPanel.Controls.Add(btnUseCurrentDomains)

                Dim btnRefreshIds As New System.Windows.Forms.Button() With {
                    .Text = "Refresh / Check IDs", .AutoSize = True,
                    .Margin = New System.Windows.Forms.Padding(0, 0, 8, 0)}
                buttonPanel.Controls.Add(btnRefreshIds)

                Dim btnGenerate As New Button() With {
                    .Text = "Generate",
                    .AutoSize = True,
                    .Margin = New Padding(0, 0, 8, 0)
                }
                buttonPanel.Controls.Add(btnGenerate)

                Dim btnCopy As New Button() With {
                    .Text = "Copy Key",
                    .AutoSize = True,
                    .Margin = New Padding(0, 0, 8, 0)
                }
                buttonPanel.Controls.Add(btnCopy)

                Dim btnClose As New Button() With {
                    .Text = "Close",
                    .AutoSize = True,
                    .Margin = New Padding(0)
                }
                buttonPanel.Controls.Add(btnClose)

                buttonPanel.Dock = System.Windows.Forms.DockStyle.Fill
                buttonPanel.Padding = New System.Windows.Forms.Padding(20, 8, 20, 14)
                dialogLayout.Controls.Add(buttonPanel, 0, 1)
                mainLayout.Controls.Add(resultPanel, 1, 7)

                Dim currentSnapshot As OfflineLicenseIdentitySnapshot = Nothing
                Dim refreshIds As System.Action =
                    Sub()
                        Dim requested As System.Collections.Generic.List(Of System.String) = ParseOfflineDomainList(txtDomains.Text)
                        currentSnapshot = GetCurrentOfflineLicenseIdentitySnapshot(requested)
                        txtCurrentDomains.Text = System.String.Join(vbCrLf, currentSnapshot.Identifiers)
                        lblDetectedIdsDetails.Text = "Detection details: hover here. Check another tenant via Refresh / Check IDs."
                        Dim details As System.String = System.String.Join(vbCrLf, currentSnapshot.Details) & vbCrLf &
                            "Windows may return only one default work account. To check another tenant, enter tenant:<GUID> above and click Refresh / Check IDs. Only Windows-confirmed registrations appear as detected IDs; an M365 configuration value alone does not qualify."
                        hints.SetToolTip(lblDetectedIdsDetails, details)
                        hints.SetToolTip(txtCurrentDomains, details)
                        hints.SetToolTip(btnRefreshIds, details)
                    End Sub
                refreshIds.Invoke()
                AddHandler btnRefreshIds.Click,
                    Sub()
                        Try
                            refreshIds.Invoke()
                        Catch ex As System.Exception
                            Using PushDialogOwner(form)
                                ShowCustomMessageBox(ex.Message, $"{AN} - Offline License Manager")
                            End Using
                        End Try
                    End Sub

                If String.IsNullOrWhiteSpace(txtDomains.Text) Then
                    txtDomains.Text = txtCurrentDomains.Text
                End If

                AddHandler btnUseCurrentDomains.Click,
                    Sub()
                        txtDomains.Text = System.String.Join(vbCrLf, currentSnapshot.Identifiers)
                    End Sub

                AddHandler btnGenerate.Click,
                    Sub()
                        Try
                            Dim normalizedAllowedDomains As System.Collections.Generic.List(Of System.String) = ParseOfflineDomainList(txtDomains.Text)
                            txtDomains.Text = System.String.Join(vbCrLf, normalizedAllowedDomains)

                            txtLicenseKey.Text = GenerateOfflineDomainLicenseKeyInternal(
                                txtProductId.Text.Trim(),
                                normalizedAllowedDomains,
                                dtpValidUntil.Value.Date,
                                txtPrivateKey.Text.Trim())

                        Catch ex As System.Exception
                            ShowCustomMessageBox($"Failed to generate offline-domain license key: {ex.Message}", $"{AN} - License Generator")
                        End Try
                    End Sub

                AddHandler btnCopy.Click,
                    Sub()
                        If String.IsNullOrWhiteSpace(txtLicenseKey.Text) Then
                            ShowCustomMessageBox("No license key has been generated yet.", $"{AN} - License Generator")
                            Return
                        End If

                        Try
                            Clipboard.SetText(txtLicenseKey.Text)
                        Catch ex As System.Exception
                            ShowCustomMessageBox($"Could not copy the license key: {ex.Message}", $"{AN} - License Generator")
                        End Try
                    End Sub

                AddHandler btnClose.Click,
                    Sub()
                        form.Close()
                    End Sub

                form.AcceptButton = btnGenerate
                Dim __safeDialogOwner796 As System.Windows.Forms.IWin32Window = Global.SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
                If __safeDialogOwner796 IsNot Nothing Then
                    form.ShowDialog(__safeDialogOwner796)
                Else
                    form.ShowDialog()
                End If
            End Using
        End Sub

#End Region

    End Class
End Namespace
