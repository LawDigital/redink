' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SharedMethods.SafeExternalLinks.vb
' Purpose: Central security boundary for opening external links from browser-rendered or otherwise untrusted content.
'
' Security:
'  - Accepts absolute HTTP, HTTPS and MAILTO URIs only.
'  - Rejects file paths, UNC paths, relative paths, javascript/data URIs and custom protocol handlers.
'  - Performs the shell launch only after validation.
'
' Architecture:
'  - Host-agnostic shared helper; callers keep their existing click/navigation behavior.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary
    Partial Public Class SharedMethods

        Public Shared Function SafeOpenExternalLink(ByVal rawUrl As System.String) As System.Boolean
            Try
                If System.String.IsNullOrWhiteSpace(rawUrl) Then Return False

                Dim parsedUri As System.Uri = Nothing
                If Not System.Uri.TryCreate(rawUrl.Trim(), System.UriKind.Absolute, parsedUri) Then Return False

                Dim scheme As System.String = parsedUri.Scheme.ToLowerInvariant()
                If scheme <> System.Uri.UriSchemeHttp AndAlso
                   scheme <> System.Uri.UriSchemeHttps AndAlso
                   scheme <> System.Uri.UriSchemeMailto Then
                    Return False
                End If

                Dim startInfo As New System.Diagnostics.ProcessStartInfo(parsedUri.AbsoluteUri) With {
                    .UseShellExecute = True
                }
                System.Diagnostics.Process.Start(startInfo)
                Return True
            Catch ex As System.Exception
                Return False
            End Try
        End Function

    End Class
End Namespace
