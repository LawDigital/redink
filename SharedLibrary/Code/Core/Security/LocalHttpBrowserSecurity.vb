' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: LocalHttpBrowserSecurity.vb
' Purpose:
'   Provides host-agnostic security primitives for Red Ink localhost HTTP surfaces.
'
' Security model:
'   - Local chat uses a host-only HttpOnly SameSite=Strict cookie containing a
'     signed, time-bounded browser credential.
'   - The signing key is generated randomly and, where the host supplies working
'     setting delegates, persisted DPAPI-protected for the current Windows user.
'   - If persistence is unavailable, the generated key remains valid in memory
'     for the current Office process only. A restart therefore invalidates prior
'     browser credentials rather than weakening authentication.
'   - State-changing same-origin requests use an HMAC-derived CSRF token.
'   - HTTP request bodies are read through a byte-counted bounded stream so
'     declared and chunked requests cannot exceed the configured transport limit.
'   - External Chromium-extension origins are explicitly approved and persisted
'     as an allow-list. Arbitrary http/https web origins are never eligible for
'     this external-origin approval path.
'
' Architecture:
'   This class intentionally has no Office-, tool-, model-, template- or
'   organization-specific knowledge. Hosts provide persistence delegates and
'   decide when/how to ask the user for approval; hosts use the standard Red Ink
'   custom authorization UI rather than framework message boxes.
' =============================================================================

Option Explicit On
Option Strict On

Namespace SharedLibrary

    Public NotInheritable Class LocalHttpBrowserSecurity

        Public Const DefaultAuthCookieName As System.String = "RedInkLocalAuth"

        Public Const MaxStandardRequestBytes As System.Int64 = 25L * 1024L * 1024L
        Public Const MaxUploadRequestBytes As System.Int64 = 250L * 1024L * 1024L
        Public Const MaxSingleUploadBytes As System.Int64 = 150L * 1024L * 1024L
        Public Const MaxUploadTotalBytes As System.Int64 = 1024L * 1024L * 1024L
        Public Const MaxUploadFileCount As System.Int32 = 50

        Private Const TokenVersion As System.String = "v1"
        Private Const TokenLifetimeDays As System.Int32 = 365
        Private Const MasterKeyByteLength As System.Int32 = 32
        Private Const ClientNonceByteLength As System.Int32 = 32
        Private Const DpapiPrefix As System.String = "dpapi-v1:"

        Private Shared ReadOnly DpapiEntropy As System.Byte() =
            System.Text.Encoding.UTF8.GetBytes("RedInk.LocalHttpBrowserSecurity.v1")

        Private ReadOnly _masterKeySettingName As System.String
        Private ReadOnly _approvedOriginsSettingName As System.String
        Private ReadOnly _readSetting As System.Func(Of System.String, System.String)
        Private ReadOnly _writeSetting As System.Func(Of System.String, System.String, System.Boolean)
        Private ReadOnly _logger As System.Action(Of System.String)
        Private ReadOnly _gate As New System.Object()
        Private ReadOnly _approvedOrigins As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)

        Private _masterKey As System.Byte() = Nothing
        Private _approvedOriginsLoaded As System.Boolean = False

        Public Sub New(
            ByVal masterKeySettingName As System.String,
            ByVal approvedOriginsSettingName As System.String,
            ByVal readSetting As System.Func(Of System.String, System.String),
            ByVal writeSetting As System.Func(Of System.String, System.String, System.Boolean),
            Optional ByVal logger As System.Action(Of System.String) = Nothing
        )
            _masterKeySettingName = If(masterKeySettingName, System.String.Empty)
            _approvedOriginsSettingName = If(approvedOriginsSettingName, System.String.Empty)
            _readSetting = readSetting
            _writeSetting = writeSetting
            _logger = logger
        End Sub

        Public Shared Async Function ReadUtf8RequestBodyBoundedAsync(
            ByVal request As System.Net.HttpListenerRequest,
            ByVal maxBytes As System.Int64
        ) As System.Threading.Tasks.Task(Of System.String)
            If request Is Nothing Then Throw New System.ArgumentNullException(NameOf(request))
            If maxBytes <= 0L Then Throw New System.ArgumentOutOfRangeException(NameOf(maxBytes))

            If request.ContentLength64 > maxBytes Then
                Throw New LocalHttpRequestTooLargeException(maxBytes)
            End If

            Using bounded As New BoundedReadStream(request.InputStream, maxBytes)
                Using rdr As New System.IO.StreamReader(
                    bounded,
                    System.Text.Encoding.UTF8,
                    detectEncodingFromByteOrderMarks:=True,
                    bufferSize:=8192,
                    leaveOpen:=False)

                    Return Await rdr.ReadToEndAsync().ConfigureAwait(False)
                End Using
            End Using
        End Function

        Public NotInheritable Class LocalHttpRequestTooLargeException
            Inherits System.IO.IOException

            Public ReadOnly Property MaxBytes As System.Int64

            Public Sub New(ByVal maxBytes As System.Int64)
                MyBase.New("HTTP request body exceeds the configured maximum of " &
                           maxBytes.ToString(System.Globalization.CultureInfo.InvariantCulture) & " bytes.")
                Me.MaxBytes = maxBytes
            End Sub
        End Class

        Private NotInheritable Class BoundedReadStream
            Inherits System.IO.Stream

            Private ReadOnly _inner As System.IO.Stream
            Private ReadOnly _maxBytes As System.Int64
            Private _bytesRead As System.Int64

            Public Sub New(ByVal inner As System.IO.Stream, ByVal maxBytes As System.Int64)
                If inner Is Nothing Then Throw New System.ArgumentNullException(NameOf(inner))
                _inner = inner
                _maxBytes = maxBytes
            End Sub

            Public Overrides ReadOnly Property CanRead As System.Boolean
                Get
                    Return _inner.CanRead
                End Get
            End Property

            Public Overrides ReadOnly Property CanSeek As System.Boolean
                Get
                    Return False
                End Get
            End Property

            Public Overrides ReadOnly Property CanWrite As System.Boolean
                Get
                    Return False
                End Get
            End Property

            Public Overrides ReadOnly Property Length As System.Int64
                Get
                    Throw New System.NotSupportedException()
                End Get
            End Property

            Public Overrides Property Position As System.Int64
                Get
                    Return _bytesRead
                End Get
                Set(ByVal value As System.Int64)
                    Throw New System.NotSupportedException()
                End Set
            End Property

            Public Overrides Sub Flush()
            End Sub

            Public Overrides Function Read(
                ByVal buffer() As System.Byte,
                ByVal offset As System.Int32,
                ByVal count As System.Int32
            ) As System.Int32
                Dim allowed As System.Int32 = GetAllowedReadCount(count)
                Dim bytesReadNow As System.Int32 = _inner.Read(buffer, offset, allowed)
                TrackRead(bytesReadNow)
                Return bytesReadNow
            End Function

            Public Overrides Async Function ReadAsync(
                ByVal buffer() As System.Byte,
                ByVal offset As System.Int32,
                ByVal count As System.Int32,
                ByVal cancellationToken As System.Threading.CancellationToken
            ) As System.Threading.Tasks.Task(Of System.Int32)
                Dim allowed As System.Int32 = GetAllowedReadCount(count)
                Dim bytesReadNow As System.Int32 = Await _inner.ReadAsync(buffer, offset, allowed, cancellationToken).ConfigureAwait(False)
                TrackRead(bytesReadNow)
                Return bytesReadNow
            End Function

            Private Function GetAllowedReadCount(ByVal requested As System.Int32) As System.Int32
                Dim remaining As System.Int64 = _maxBytes - _bytesRead
                If remaining < 0L Then Throw New LocalHttpRequestTooLargeException(_maxBytes)

                ' Permit one additional byte so chunked/unknown-length requests are
                ' detected instead of appearing to end exactly at the configured cap.
                Dim probeAllowance As System.Int64 = remaining + 1L
                Return System.Convert.ToInt32(System.Math.Min(System.Convert.ToInt64(requested), probeAllowance))
            End Function

            Private Sub TrackRead(ByVal read As System.Int32)
                If read <= 0 Then Return
                _bytesRead += read
                If _bytesRead > _maxBytes Then Throw New LocalHttpRequestTooLargeException(_maxBytes)
            End Sub

            Public Overrides Function Seek(ByVal offset As System.Int64, ByVal origin As System.IO.SeekOrigin) As System.Int64
                Throw New System.NotSupportedException()
            End Function

            Public Overrides Sub SetLength(ByVal value As System.Int64)
                Throw New System.NotSupportedException()
            End Sub

            Public Overrides Sub Write(ByVal buffer() As System.Byte, ByVal offset As System.Int32, ByVal count As System.Int32)
                Throw New System.NotSupportedException()
            End Sub

            Protected Overrides Sub Dispose(ByVal disposing As System.Boolean)
                If disposing Then _inner.Dispose()
                MyBase.Dispose(disposing)
            End Sub
        End Class

        Public Function IsLoopbackRequest(ByVal request As System.Net.HttpListenerRequest) As System.Boolean
            If request Is Nothing Then Return False

            Try
                If request.RemoteEndPoint Is Nothing OrElse
                   Not System.Net.IPAddress.IsLoopback(request.RemoteEndPoint.Address) Then
                    Return False
                End If

                If request.Url Is Nothing Then Return False

                Dim host As System.String = If(request.Url.Host, System.String.Empty)
                If host.Equals("localhost", System.StringComparison.OrdinalIgnoreCase) Then Return True

                Dim address As System.Net.IPAddress = Nothing
                If System.Net.IPAddress.TryParse(host, address) Then
                    Return System.Net.IPAddress.IsLoopback(address)
                End If
            Catch ex As System.Exception
                Return False
            End Try

            Return False
        End Function

        Public Function IsSameOrigin(
            ByVal origin As System.String,
            ByVal requestUrl As System.Uri
        ) As System.Boolean
            If System.String.IsNullOrWhiteSpace(origin) OrElse requestUrl Is Nothing Then Return False

            Dim normalizedOrigin As System.String = Nothing
            If Not TryNormalizeOrigin(origin, normalizedOrigin) Then Return False

            Dim expectedOrigin As System.String =
                requestUrl.Scheme & "://" & requestUrl.Host & ":" &
                requestUrl.Port.ToString(System.Globalization.CultureInfo.InvariantCulture)

            Dim normalizedExpected As System.String = Nothing
            If Not TryNormalizeOrigin(expectedOrigin, normalizedExpected) Then Return False

            Return normalizedOrigin.Equals(normalizedExpected, System.StringComparison.OrdinalIgnoreCase)
        End Function

        Public Function IsBrowserExtensionOrigin(ByVal origin As System.String) As System.Boolean
            Dim normalizedOrigin As System.String = Nothing
            If Not TryNormalizeOrigin(origin, normalizedOrigin) Then Return False

            Dim uri As System.Uri = Nothing
            If Not System.Uri.TryCreate(normalizedOrigin, System.UriKind.Absolute, uri) Then Return False

            Dim scheme As System.String = If(uri.Scheme, System.String.Empty)
            Return scheme.Equals("chrome-extension", System.StringComparison.OrdinalIgnoreCase) OrElse
                   scheme.Equals("edge-extension", System.StringComparison.OrdinalIgnoreCase) OrElse
                   scheme.Equals("moz-extension", System.StringComparison.OrdinalIgnoreCase)
        End Function

        Public Function IsApprovedExternalOrigin(ByVal origin As System.String) As System.Boolean
            Dim normalizedOrigin As System.String = Nothing
            If Not TryNormalizeOrigin(origin, normalizedOrigin) Then Return False

            SyncLock _gate
                EnsureApprovedOriginsLoadedLocked()
                Return _approvedOrigins.Contains(normalizedOrigin)
            End SyncLock
        End Function

        Public Function ApproveExternalOrigin(ByVal origin As System.String) As System.Boolean
            If Not IsBrowserExtensionOrigin(origin) Then Return False

            Dim normalizedOrigin As System.String = Nothing
            If Not TryNormalizeOrigin(origin, normalizedOrigin) Then Return False

            SyncLock _gate
                EnsureApprovedOriginsLoadedLocked()
                _approvedOrigins.Add(normalizedOrigin)
                PersistApprovedOriginsLocked()
            End SyncLock

            Return True
        End Function

        Public Function IssueBrowserCredential() As System.String
            Dim issuedTicks As System.Int64 = System.DateTime.UtcNow.Ticks
            Dim nonce(ClientNonceByteLength - 1) As System.Byte

            Using rng As System.Security.Cryptography.RandomNumberGenerator =
                System.Security.Cryptography.RandomNumberGenerator.Create()
                rng.GetBytes(nonce)
            End Using

            Dim payloadText As System.String =
                issuedTicks.ToString(System.Globalization.CultureInfo.InvariantCulture) & "|" & Base64UrlEncode(nonce)

            Dim payloadBytes() As System.Byte = System.Text.Encoding.UTF8.GetBytes(payloadText)
            Dim payload As System.String = Base64UrlEncode(payloadBytes)
            Dim signature As System.String = Base64UrlEncode(ComputeHmac(TokenVersion & "." & payload))

            Return TokenVersion & "." & payload & "." & signature
        End Function

        Public Function TryValidateBrowserCredential(
            ByVal cookieHeader As System.String,
            ByRef browserCredential As System.String
        ) As System.Boolean
            browserCredential = System.String.Empty

            Dim candidate As System.String = GetCookieValue(cookieHeader, DefaultAuthCookieName)
            If System.String.IsNullOrWhiteSpace(candidate) Then Return False

            If Not ValidateCredentialToken(candidate) Then Return False

            browserCredential = candidate
            Return True
        End Function

        Public Function CreateAuthCookieHeader(ByVal browserCredential As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(browserCredential) Then
                Throw New System.ArgumentException("Browser credential is required.", NameOf(browserCredential))
            End If

            Dim expiresUtc As System.DateTime = System.DateTime.UtcNow.AddDays(TokenLifetimeDays)

            Return DefaultAuthCookieName & "=" & browserCredential &
                   "; Path=/; HttpOnly; SameSite=Strict; Max-Age=" &
                   (TokenLifetimeDays * 24 * 60 * 60).ToString(System.Globalization.CultureInfo.InvariantCulture) &
                   "; Expires=" & expiresUtc.ToString("R", System.Globalization.CultureInfo.InvariantCulture)
        End Function

        Public Function CreateCsrfToken(ByVal browserCredential As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(browserCredential) Then Return System.String.Empty
            Return Base64UrlEncode(ComputeHmac("csrf|" & browserCredential))
        End Function

        Public Function ValidateCsrfToken(
            ByVal browserCredential As System.String,
            ByVal csrfToken As System.String
        ) As System.Boolean
            If System.String.IsNullOrWhiteSpace(browserCredential) OrElse
               System.String.IsNullOrWhiteSpace(csrfToken) Then
                Return False
            End If

            Dim expected As System.Byte() = ComputeHmac("csrf|" & browserCredential)
            Dim supplied As System.Byte() = Nothing

            Try
                supplied = Base64UrlDecode(csrfToken)
            Catch ex As System.Exception
                Return False
            End Try

            Return FixedTimeEquals(expected, supplied)
        End Function

        Public Shared Function TryNormalizeOrigin(
            ByVal origin As System.String,
            ByRef normalizedOrigin As System.String
        ) As System.Boolean
            normalizedOrigin = System.String.Empty
            If System.String.IsNullOrWhiteSpace(origin) Then Return False

            Dim uri As System.Uri = Nothing
            If Not System.Uri.TryCreate(origin.Trim(), System.UriKind.Absolute, uri) Then Return False
            If Not System.String.IsNullOrEmpty(uri.UserInfo) Then Return False
            If System.String.IsNullOrWhiteSpace(uri.Scheme) OrElse System.String.IsNullOrWhiteSpace(uri.Host) Then Return False

            Dim scheme As System.String = uri.Scheme.ToLowerInvariant()
            Dim host As System.String = uri.Host.ToLowerInvariant()

            Select Case scheme
                Case "http", "https"
                    normalizedOrigin = scheme & "://" & host & ":" &
                        uri.Port.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    Return True

                Case "chrome-extension", "edge-extension", "moz-extension"
                    normalizedOrigin = scheme & "://" & host
                    Return True
            End Select

            Return False
        End Function

        Private Function ValidateCredentialToken(ByVal token As System.String) As System.Boolean
            Dim parts() As System.String = token.Split("."c)
            If parts.Length <> 3 Then Return False
            If Not parts(0).Equals(TokenVersion, System.StringComparison.Ordinal) Then Return False

            Dim suppliedSignature() As System.Byte
            Try
                suppliedSignature = Base64UrlDecode(parts(2))
            Catch ex As System.Exception
                Return False
            End Try

            Dim expectedSignature() As System.Byte = ComputeHmac(parts(0) & "." & parts(1))
            If Not FixedTimeEquals(expectedSignature, suppliedSignature) Then Return False

            Dim payloadBytes() As System.Byte
            Try
                payloadBytes = Base64UrlDecode(parts(1))
            Catch ex As System.Exception
                Return False
            End Try

            Dim payloadText As System.String = System.Text.Encoding.UTF8.GetString(payloadBytes)
            Dim separatorIndex As System.Int32 = payloadText.IndexOf("|"c)
            If separatorIndex <= 0 Then Return False

            Dim ticksText As System.String = payloadText.Substring(0, separatorIndex)
            Dim issuedTicks As System.Int64
            If Not System.Int64.TryParse(
                ticksText,
                System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture,
                issuedTicks) Then
                Return False
            End If

            Dim issuedUtc As System.DateTime
            Try
                issuedUtc = New System.DateTime(issuedTicks, System.DateTimeKind.Utc)
            Catch ex As System.Exception
                Return False
            End Try

            Dim nowUtc As System.DateTime = System.DateTime.UtcNow
            If issuedUtc > nowUtc.AddMinutes(5) Then Return False
            If nowUtc - issuedUtc > System.TimeSpan.FromDays(TokenLifetimeDays) Then Return False

            Dim nonceText As System.String = payloadText.Substring(separatorIndex + 1)
            If System.String.IsNullOrWhiteSpace(nonceText) Then Return False

            Try
                Dim nonce() As System.Byte = Base64UrlDecode(nonceText)
                If nonce.Length <> ClientNonceByteLength Then Return False
            Catch ex As System.Exception
                Return False
            End Try

            Return True
        End Function

        Private Function ComputeHmac(ByVal text As System.String) As System.Byte()
            Dim key() As System.Byte = GetMasterKey()
            Using hmac As New System.Security.Cryptography.HMACSHA256(key)
                Return hmac.ComputeHash(System.Text.Encoding.UTF8.GetBytes(If(text, System.String.Empty)))
            End Using
        End Function

        Private Function GetMasterKey() As System.Byte()
            SyncLock _gate
                If _masterKey IsNot Nothing AndAlso _masterKey.Length = MasterKeyByteLength Then
                    Return CType(_masterKey.Clone(), System.Byte())
                End If

                Dim persisted As System.String = ReadSettingSafe(_masterKeySettingName)
                Dim restored() As System.Byte = TryUnprotectMasterKey(persisted)

                If restored IsNot Nothing AndAlso restored.Length = MasterKeyByteLength Then
                    _masterKey = restored
                    Return CType(_masterKey.Clone(), System.Byte())
                End If

                Dim generated(MasterKeyByteLength - 1) As System.Byte
                Using rng As System.Security.Cryptography.RandomNumberGenerator =
                    System.Security.Cryptography.RandomNumberGenerator.Create()
                    rng.GetBytes(generated)
                End Using

                _masterKey = generated

                Dim protectedValue As System.String = TryProtectMasterKey(generated)
                If System.String.IsNullOrWhiteSpace(protectedValue) Then
                    LogMessage("DPAPI protection was unavailable; browser authorization will persist only for this Office process.")
                ElseIf Not WriteSettingSafe(_masterKeySettingName, protectedValue) Then
                    LogMessage("My.Settings persistence was unavailable; browser authorization will persist only for this Office process.")
                End If

                Return CType(_masterKey.Clone(), System.Byte())
            End SyncLock
        End Function

        Private Function TryProtectMasterKey(ByVal key As System.Byte()) As System.String
            If key Is Nothing OrElse key.Length = 0 Then Return System.String.Empty

            Try
                Dim protectedBytes() As System.Byte =
                    System.Security.Cryptography.ProtectedData.Protect(
                        key,
                        DpapiEntropy,
                        System.Security.Cryptography.DataProtectionScope.CurrentUser)

                Return DpapiPrefix & System.Convert.ToBase64String(protectedBytes)
            Catch ex As System.Exception
                Return System.String.Empty
            End Try
        End Function

        Private Function TryUnprotectMasterKey(ByVal persisted As System.String) As System.Byte()
            If System.String.IsNullOrWhiteSpace(persisted) OrElse
               Not persisted.StartsWith(DpapiPrefix, System.StringComparison.Ordinal) Then
                Return Nothing
            End If

            Try
                Dim protectedBytes() As System.Byte =
                    System.Convert.FromBase64String(persisted.Substring(DpapiPrefix.Length))

                Return System.Security.Cryptography.ProtectedData.Unprotect(
                    protectedBytes,
                    DpapiEntropy,
                    System.Security.Cryptography.DataProtectionScope.CurrentUser)
            Catch ex As System.Exception
                Return Nothing
            End Try
        End Function

        Private Sub EnsureApprovedOriginsLoadedLocked()
            If _approvedOriginsLoaded Then Return
            _approvedOriginsLoaded = True

            Dim persisted As System.String = ReadSettingSafe(_approvedOriginsSettingName)
            If System.String.IsNullOrWhiteSpace(persisted) Then Return

            Dim lines() As System.String = persisted.Replace(Microsoft.VisualBasic.Constants.vbCr, System.String.Empty).Split(Microsoft.VisualBasic.ControlChars.Lf)
            For Each line As System.String In lines
                Dim normalized As System.String = Nothing
                If IsBrowserExtensionOrigin(line) AndAlso TryNormalizeOrigin(line, normalized) Then
                    _approvedOrigins.Add(normalized)
                End If
            Next
        End Sub

        Private Sub PersistApprovedOriginsLocked()
            If _approvedOrigins.Count = 0 Then
                WriteSettingSafe(_approvedOriginsSettingName, System.String.Empty)
                Return
            End If

            Dim values As New System.Collections.Generic.List(Of System.String)(_approvedOrigins)
            values.Sort(System.StringComparer.OrdinalIgnoreCase)
            If Not WriteSettingSafe(_approvedOriginsSettingName, System.String.Join(Microsoft.VisualBasic.Constants.vbLf, values)) Then
                LogMessage("My.Settings persistence was unavailable; approved browser-extension origins will persist only for this Office process.")
            End If
        End Sub

        Private Function ReadSettingSafe(ByVal settingName As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(settingName) OrElse _readSetting Is Nothing Then
                Return System.String.Empty
            End If

            Try
                Return If(_readSetting.Invoke(settingName), System.String.Empty)
            Catch ex As System.Exception
                Return System.String.Empty
            End Try
        End Function

        Private Function WriteSettingSafe(
            ByVal settingName As System.String,
            ByVal settingValue As System.String
        ) As System.Boolean
            If System.String.IsNullOrWhiteSpace(settingName) OrElse _writeSetting Is Nothing Then Return False

            Try
                Return _writeSetting.Invoke(settingName, If(settingValue, System.String.Empty))
            Catch ex As System.Exception
                Return False
            End Try
        End Function

        Private Sub LogMessage(ByVal message As System.String)
            If _logger Is Nothing OrElse System.String.IsNullOrWhiteSpace(message) Then Return
            Try
                _logger.Invoke(message)
            Catch ex As System.Exception
            End Try
        End Sub

        Private Shared Function GetCookieValue(
            ByVal cookieHeader As System.String,
            ByVal cookieName As System.String
        ) As System.String
            If System.String.IsNullOrWhiteSpace(cookieHeader) OrElse
               System.String.IsNullOrWhiteSpace(cookieName) Then
                Return System.String.Empty
            End If

            Dim pieces() As System.String = cookieHeader.Split(";"c)
            For Each piece As System.String In pieces
                Dim separatorIndex As System.Int32 = piece.IndexOf("="c)
                If separatorIndex <= 0 Then Continue For

                Dim name As System.String = piece.Substring(0, separatorIndex).Trim()
                If name.Equals(cookieName, System.StringComparison.Ordinal) Then
                    Return piece.Substring(separatorIndex + 1).Trim()
                End If
            Next

            Return System.String.Empty
        End Function

        Private Shared Function Base64UrlEncode(ByVal data As System.Byte()) As System.String
            If data Is Nothing OrElse data.Length = 0 Then Return System.String.Empty

            Return System.Convert.ToBase64String(data).
                TrimEnd("="c).
                Replace("+"c, "-"c).
                Replace("/"c, "_"c)
        End Function

        Private Shared Function Base64UrlDecode(ByVal value As System.String) As System.Byte()
            If System.String.IsNullOrWhiteSpace(value) Then Return New System.Byte() {}

            Dim text As System.String = value.Replace("-"c, "+"c).Replace("_"c, "/"c)
            Select Case text.Length Mod 4
                Case 2
                    text &= "=="
                Case 3
                    text &= "="
                Case 1
                    Throw New System.FormatException("Invalid base64url value.")
            End Select

            Return System.Convert.FromBase64String(text)
        End Function

        Private Shared Function FixedTimeEquals(
            ByVal left As System.Byte(),
            ByVal right As System.Byte()
        ) As System.Boolean
            If left Is Nothing OrElse right Is Nothing OrElse left.Length <> right.Length Then Return False

            Dim diff As System.Int32 = 0
            For i As System.Int32 = 0 To left.Length - 1
                diff = diff Or (CInt(left(i)) Xor CInt(right(i)))
            Next

            Return diff = 0
        End Function

    End Class

End Namespace
