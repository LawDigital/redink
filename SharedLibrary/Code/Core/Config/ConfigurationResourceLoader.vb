' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' Generic read-only resolver for configuration resources. File-system sources retain
' their existing semantics; HTTPS sources are downloaded through SharedMethods' existing
' HTTP stack and materialized as process-local read snapshots.

Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.IO

Namespace SharedLibrary

    Public Enum ConfigurationSourceKind
        FileSystem = 0
        Https = 1
        UnsupportedUri = 2
        Invalid = 3
    End Enum

    Public NotInheritable Class ConfigurationResourceLoader

        Private Shared ReadOnly _sync As New Object()
        Private Shared ReadOnly _sessionSnapshots As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        Private Sub New()
        End Sub

        Public Shared Function ClassifyConfigurationSource(ByVal source As String) As ConfigurationSourceKind
            Dim value As String = If(source, "").Trim()
            If String.IsNullOrWhiteSpace(value) Then Return ConfigurationSourceKind.Invalid

            Dim uri As System.Uri = Nothing
            If System.Uri.TryCreate(value, UriKind.Absolute, uri) AndAlso uri IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(uri.Scheme) Then
                If String.Equals(uri.Scheme, System.Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) Then
                    Return ConfigurationSourceKind.Https
                End If

                If value.IndexOf("://", StringComparison.Ordinal) >= 0 Then
                    Return ConfigurationSourceKind.UnsupportedUri
                End If
            ElseIf value.IndexOf("://", StringComparison.Ordinal) >= 0 Then
                Return ConfigurationSourceKind.UnsupportedUri
            End If

            Return ConfigurationSourceKind.FileSystem
        End Function

        Public Shared Function CanResolve(ByVal source As String) As Boolean
            Dim kind As ConfigurationSourceKind = ClassifyConfigurationSource(source)
            Return kind = ConfigurationSourceKind.FileSystem OrElse kind = ConfigurationSourceKind.Https
        End Function

        Public Shared Function ResolveForRead(ByVal source As String,
                                              Optional ByVal configurationKind As String = "configuration") As String
            Dim kind As ConfigurationSourceKind = ClassifyConfigurationSource(source)

            Select Case kind
                Case ConfigurationSourceKind.FileSystem
                    Dim expanded As String = SharedMethods.ExpandEnvironmentVariables(If(source, "").Trim())
                    If String.IsNullOrWhiteSpace(expanded) Then
                        Throw New System.IO.FileNotFoundException("The configuration source is empty.")
                    End If
                    Return expanded

                Case ConfigurationSourceKind.Https
                    Return ResolveHttpsForRead(source.Trim(), configurationKind)

                Case ConfigurationSourceKind.UnsupportedUri
                    Throw New System.NotSupportedException("Only file-system paths and HTTPS configuration URLs are supported. HTTP and other URI schemes are not supported.")

                Case Else
                    Throw New System.ArgumentException("The configuration source is invalid.", NameOf(source))
            End Select
        End Function

        Private Shared Function ResolveHttpsForRead(ByVal source As String, ByVal configurationKind As String) As String
            SyncLock _sync
                Dim existing As String = Nothing
                If _sessionSnapshots.TryGetValue(source, existing) AndAlso
                   Not String.IsNullOrWhiteSpace(existing) AndAlso
                   System.IO.File.Exists(existing) Then
                    Return existing
                End If
            End SyncLock

            Dim request As New SharedMethods.SharedHttpRequest() With {
                .Url = source,
                .Method = "GET",
                .TimeoutMs = 30000,
                .Accept = "text/plain, application/octet-stream;q=0.9, */*;q=0.1",
                .StackPreference = SharedMethods.HttpStackPreference.PreferConfiguredDefault,
                .UseAutomaticClientCertificate = True,
                .AllowAutoRedirect = False
            }

            Dim started As System.DateTime = System.DateTime.UtcNow
            Dim response As SharedMethods.SharedHttpResponse = Nothing

            Try
                response = SharedMethods.SendHttpRequestAsync(request).GetAwaiter().GetResult()
            Catch ex As System.Exception
                Throw New System.IO.IOException("TLS/authentication or network failure while loading remote " & configurationKind & ".", ex)
            End Try

            Dim elapsedMs As Long = CLng((System.DateTime.UtcNow - started).TotalMilliseconds)
            System.Diagnostics.Debug.WriteLine("Remote configuration GET: kind=" & configurationKind &
                                               "; source=https; stack=" & If(response.UsedStack, "") &
                                               "; status=" & response.StatusCode.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                                               "; durationMs=" & elapsedMs.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                                               "; url=" & GetSafeSourceIdentifier(source))

            If response.StatusCode < 200 OrElse response.StatusCode > 299 Then
                Select Case response.StatusCode
                    Case 401, 403
                        Throw New System.UnauthorizedAccessException("Access denied while loading remote " & configurationKind & " (HTTP " & response.StatusCode.ToString(System.Globalization.CultureInfo.InvariantCulture) & ").")
                    Case 404
                        Throw New System.IO.FileNotFoundException("Remote " & configurationKind & " was not found (HTTP 404).")
                    Case 300 To 399
                        Throw New System.IO.IOException("Remote " & configurationKind & " returned a redirect. Configuration URLs must address the final HTTPS resource directly.")
                    Case 500 To 599
                        Throw New System.IO.IOException("Remote server is unavailable while loading " & configurationKind & " (HTTP " & response.StatusCode.ToString(System.Globalization.CultureInfo.InvariantCulture) & ").")
                    Case Else
                        Throw New System.IO.IOException("Remote " & configurationKind & " request failed with HTTP " & response.StatusCode.ToString(System.Globalization.CultureInfo.InvariantCulture) & ".")
                End Select
            End If

            Dim bytes As Byte() = If(response.BodyBytes, New Byte() {})
            If bytes.Length = 0 Then
                Throw New System.IO.InvalidDataException("Remote " & configurationKind & " is empty.")
            End If

            Dim sessionDirectory As String = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "RedInk", "Configuration", System.Diagnostics.Process.GetCurrentProcess().Id.ToString(System.Globalization.CultureInfo.InvariantCulture))
            System.IO.Directory.CreateDirectory(sessionDirectory)

            Dim extension As String = ".ini"
            Dim finalPath As String = System.IO.Path.Combine(sessionDirectory, System.Guid.NewGuid().ToString("N") & extension)
            Dim tempPath As String = finalPath & ".download"

            Try
                System.IO.File.WriteAllBytes(tempPath, bytes)
                System.IO.File.Move(tempPath, finalPath)
            Finally
                If System.IO.File.Exists(tempPath) Then
                    Try
                        System.IO.File.Delete(tempPath)
                    Catch
                    End Try
                End If
            End Try

            SyncLock _sync
                Dim priorPath As String = Nothing
                If _sessionSnapshots.TryGetValue(source, priorPath) AndAlso System.IO.File.Exists(priorPath) Then
                    Try
                        System.IO.File.Delete(finalPath)
                    Catch
                    End Try
                    Return priorPath
                End If
                _sessionSnapshots(source) = finalPath
            End SyncLock

            System.Diagnostics.Debug.WriteLine("Remote configuration materialized: kind=" & configurationKind & "; path=" & finalPath)
            Return finalPath
        End Function

        Public Shared Function GetSafeSourceIdentifier(ByVal source As String) As String
            Dim value As String = If(source, "").Trim()
            If ClassifyConfigurationSource(value) <> ConfigurationSourceKind.Https Then Return value

            Try
                Dim uri As New System.Uri(value)
                Dim builder As New System.UriBuilder(uri) With {.Query = ""}
                Return builder.Uri.GetLeftPart(System.UriPartial.Path)
            Catch
                Return "https://[invalid-url]"
            End Try
        End Function

    End Class

End Namespace
