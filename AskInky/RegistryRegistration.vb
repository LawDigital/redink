' Part of "Red Ink" (Ask Inky)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: RegistryRegistration.vb
' Purpose: Read a Red Ink registration link for unconfigured Ask Inky installations.
' Architecture: HKCU-only discovery with single-use, privacy-preserving attempt tracking.
' Notes: Neither the link nor its secret parameters are written to diagnostics.
' =============================================================================
Friend NotInheritable Class RegistryRegistration
    Private Sub New()
    End Sub

    Private Shared ReadOnly AttemptPath As String = System.IO.Path.Combine(StoragePaths.Root, "registry-registration-attempt.sha256")

    Friend Shared Function FindLink() As String
        Try
            Using key As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\Red Ink", False)
                If key Is Nothing Then Return ""
                Dim names As String() = key.GetValueNames()
                ' HelperLink is the explicit Red Ink registry contract. The existing
                ' registration/invite names remain compatible with older installations.
                ' A root URL with api_key (and optional uid/lang) is valid by name.
                For Each name As String In {"HelperLink", "RegistrationLink", "RegistrationUrl", "RegistrationURI", "DeviceRegistrationLink", "InviteLink", "InviteUrl", "RegisterLink", "Registration"}
                    For Each actual As String In names
                        If String.Equals(actual, name, System.StringComparison.OrdinalIgnoreCase) Then
                            Dim candidate As String = TryValue(key, actual)
                            If IsAllowedLink(candidate) Then Return candidate
                        End If
                    Next
                Next
                ' Some installations use a different name: accept only a URL that clearly
                ' identifies registration through its path; other Red Ink values are ignored.
                For Each name As String In names
                    Dim candidate As String = TryValue(key, name)
                    Dim uri As System.Uri = Nothing
                    If IsAllowedLink(candidate) AndAlso System.Uri.TryCreate(candidate, System.UriKind.Absolute, uri) AndAlso
                       (uri.AbsolutePath.IndexOf("register", System.StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                        uri.AbsolutePath.IndexOf("invite", System.StringComparison.OrdinalIgnoreCase) >= 0) Then Return candidate
                Next
            End Using
        Catch ex As System.Exception
            ClientLog.Write("registry.registration.read.failed", ex)
        End Try
        Return ""
    End Function

    Private Shared Function TryValue(key As Microsoft.Win32.RegistryKey, name As String) As String
        Try
            Dim kind As Microsoft.Win32.RegistryValueKind = key.GetValueKind(name)
            If kind <> Microsoft.Win32.RegistryValueKind.String AndAlso kind <> Microsoft.Win32.RegistryValueKind.ExpandString Then Return ""
            Return If(TryCast(key.GetValue(name, Nothing, Microsoft.Win32.RegistryValueOptions.DoNotExpandEnvironmentNames), String), "").Trim()
        Catch ex As System.Exception
            ClientLog.Write("registry.registration.value-read.failed", ex)
            Return ""
        End Try
    End Function

    Private Shared Function IsAllowedLink(value As String) As Boolean
        If String.IsNullOrWhiteSpace(value) OrElse value.Length > 32768 Then Return False
        Dim uri As System.Uri = Nothing
        If Not System.Uri.TryCreate(value, System.UriKind.Absolute, uri) Then Return False
        Return uri.UserInfo.Length = 0 AndAlso uri.Fragment.Length = 0 AndAlso
               (uri.Scheme = System.Uri.UriSchemeHttps OrElse (uri.Scheme = System.Uri.UriSchemeHttp AndAlso uri.IsLoopback))
    End Function

    Friend Shared Function Fingerprint(link As String) As String
        Return System.Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(link)))
    End Function

    Friend Shared Function IsAppliedLinkPresent(fingerprint As String) As Boolean
        If String.IsNullOrEmpty(fingerprint) Then Return False
        Dim link As String = FindLink()
        Return link.Length > 0 AndAlso System.String.Equals(RegistryRegistration.Fingerprint(link), fingerprint, System.StringComparison.Ordinal)
    End Function

    ' Mark BEFORE sending the GET. A lost HTTP response must not cause a second
    ' automatic request for the same one-time registration link on next startup.
    Friend Shared Function TryMarkAttempt(link As String) As Boolean
        Dim fingerprint As String = RegistryRegistration.Fingerprint(link)
        Try
            ' Coordinate first runs in two concurrent Ask Inky processes for this session.
            Using mutex As New System.Threading.Mutex(False, "Local\RedInk.AskInky.RegistryRegistrationMarker")
                Dim acquired As Boolean = False
                Try
                    Try
                        acquired = mutex.WaitOne(System.TimeSpan.FromSeconds(5))
                    Catch ex As System.Threading.AbandonedMutexException
                        acquired = True
                    End Try
                    If Not acquired Then
                        ClientLog.Write("registry.registration.attempt-marker.busy")
                        Return False
                    End If
                    Dim known As New System.Collections.Generic.List(Of String)()
                    If System.IO.File.Exists(AttemptPath) Then
                        Dim info As New System.IO.FileInfo(AttemptPath)
                        If info.Length > 131072 Then
                            ClientLog.Write("registry.registration.attempt-marker.oversize")
                            Return False
                        End If
                        For Each line As String In System.IO.File.ReadAllLines(AttemptPath, System.Text.Encoding.ASCII)
                            Dim existing As String = line.Trim()
                            If String.Equals(existing, fingerprint, System.StringComparison.Ordinal) Then
                                ClientLog.Write("registry.registration.automatic-retry-blocked")
                                Return False
                            End If
                            If existing.Length = 64 AndAlso System.Text.RegularExpressions.Regex.IsMatch(existing, "\A[0-9A-F]{64}\z") Then known.Add(existing)
                        Next
                    End If
                    ' Keep a bounded record of attempted one-time URLs, never their text.
                    If known.Count > 1023 Then known.RemoveRange(0, known.Count - 1023)
                    known.Add(fingerprint)
                    System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(AttemptPath))
                    Dim temporary As String = AttemptPath & "." & System.Guid.NewGuid().ToString("N") & ".tmp"
                    Try
                        System.IO.File.WriteAllLines(temporary, known, System.Text.Encoding.ASCII)
                        System.IO.File.Move(temporary, AttemptPath, True)
                    Finally
                        If System.IO.File.Exists(temporary) Then System.IO.File.Delete(temporary)
                    End Try
                    Return True
                Finally
                    If acquired Then mutex.ReleaseMutex()
                End Try
            End Using
        Catch ex As System.Exception
            ClientLog.Write("registry.registration.attempt-marker.failed", ex)
            Return False
        End Try
    End Function

End Class
