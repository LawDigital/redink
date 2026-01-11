' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.Signing.vb
' Purpose: Provides RSA-based JWT signing functionality using BouncyCastle cryptography.
'
' Architecture:
'  - PEM Parsing: Accepts private keys in PEM format or raw Base64. If raw Base64 is provided,
'    wraps it with standard PEM headers/footers before parsing.
'  - Key Format Handling: Supports both AsymmetricCipherKeyPair and RsaPrivateCrtKeyParameters
'    formats for flexibility with different key sources.
'  - Signature Algorithm: Uses SHA256withRSA for signing, producing a Base64-encoded signature.
'  - PEM Formatting: Helper function formats raw Base64 keys into 64-character line chunks
'    per PEM specification (RFC 7468).
'  - External Dependencies: Org.BouncyCastle.Crypto for cryptographic operations and PEM parsing.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Text
Imports Org.BouncyCastle.Crypto
Imports Org.BouncyCastle.Crypto.Parameters
Imports Org.BouncyCastle.Security

Namespace SharedLibrary
    Partial Public Class SharedMethods

        ''' <summary>
        ''' Signs an unsigned JWT string using an RSA private key and returns the Base64-encoded signature.
        ''' </summary>
        ''' <param name="jwtUnsigned">The unsigned JWT payload to sign (header.payload format).</param>
        ''' <param name="privateKeyPem">The RSA private key in PEM format or raw Base64 encoding.</param>
        ''' <returns>A Base64-encoded RSA-SHA256 signature.</returns>
        ''' <exception cref="ApplicationException">Thrown when signing fails or the private key format is invalid.</exception>
        Public Shared Function SignJWT(jwtUnsigned As String, privateKeyPem As String) As String
            Try

                'Dim privateKey As AsymmetricCipherKeyPair
                If Left(privateKeyPem, 3) <> "---" Then
                    privateKeyPem = "-----BEGIN PRIVATE KEY-----" & vbLf & ConvertToPemFormat(privateKeyPem) & vbLf & "-----END PRIVATE KEY-----"
                End If

                ' Read the private key using BouncyCastle PEM reader
                Dim privateKeyObject As Object
                Using reader As New StringReader(privateKeyPem)
                    Dim pemReader = New Org.BouncyCastle.OpenSsl.PemReader(reader)
                    privateKeyObject = pemReader.ReadObject()
                End Using

                ' Extract RSA key parameters based on the parsed object type
                Dim privateKeyParams As RsaKeyParameters
                If TypeOf privateKeyObject Is AsymmetricCipherKeyPair Then
                    Dim keyPair = CType(privateKeyObject, AsymmetricCipherKeyPair)
                    privateKeyParams = CType(keyPair.Private, RsaKeyParameters)
                ElseIf TypeOf privateKeyObject Is RsaPrivateCrtKeyParameters Then
                    privateKeyParams = CType(privateKeyObject, RsaPrivateCrtKeyParameters)
                Else
                    Throw New ApplicationException("Invalid private key format.")
                End If

                ' Convert unsigned JWT to bytes
                Dim unsignedDataBytes = Encoding.UTF8.GetBytes(jwtUnsigned)

                ' Create and initialize the SHA256withRSA signer
                Dim signer = SignerUtilities.GetSigner("SHA256withRSA")
                signer.Init(True, privateKeyParams)
                signer.BlockUpdate(unsignedDataBytes, 0, unsignedDataBytes.Length)
                Dim signatureBytes = signer.GenerateSignature()

                ' Return Base64-encoded signature
                Dim base64Signature = System.Convert.ToBase64String(signatureBytes)

                Return base64Signature
            Catch ex As Exception
                Throw New ApplicationException("Error signing JWT: " & ex.Message, ex)
            End Try
        End Function

        ''' <summary>
        ''' Converts a raw Base64 key string into PEM-formatted lines of 64 characters each.
        ''' </summary>
        ''' <param name="rawKey">The raw Base64-encoded key without PEM headers or line breaks.</param>
        ''' <returns>The key formatted with 64-character line breaks per PEM specification.</returns>
        Private Shared Function ConvertToPemFormat(rawKey As String) As String
            Dim sb As New StringBuilder()
            Dim index As Integer = 0
            While index < rawKey.Length
                Dim chunk As String = rawKey.Substring(index, Math.Min(64, rawKey.Length - index))
                sb.AppendLine(chunk)
                index += 64
            End While
            Return sb.ToString().Trim()
        End Function



    End Class
End Namespace