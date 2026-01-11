' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: MimeHelper.vb
' Purpose: Provides MIME type detection and Base64 encoding for files using the Windows
'          urlmon.dll `FindMimeFromData` API, which inspects file content (magic bytes)
'          rather than relying solely on file extensions.
'
' Architecture:
'  - P/Invoke Declaration: Wraps the unmanaged `FindMimeFromData` function from urlmon.dll.
'  - MIME Detection: `GetMimeType` reads the first 256 bytes of a file and passes them to
'    `FindMimeFromData` for content-based MIME type sniffing.
'  - Combined Helper: `GetFileMimeTypeAndBase64` combines MIME detection with Base64 encoding
'    of the full file content, returning both values as a tuple.
'  - Error Handling: Throws exceptions on file read errors or API call failures (non-zero HRESULT).
'  - Memory Management: Uses `Marshal.PtrToStringUni` and `Marshal.FreeCoTaskMem` to safely handle
'    the unmanaged string pointer returned by `FindMimeFromData`.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Runtime.InteropServices

Namespace SharedLibrary

    ''' <summary>
    ''' Provides helper methods for MIME type detection and file encoding using Windows APIs.
    ''' </summary>
    Public Module MimeHelper

        ''' <summary>
        ''' P/Invoke declaration for the urlmon.dll `FindMimeFromData` function, which determines
        ''' the MIME type of data by inspecting its content.
        ''' </summary>
        ''' <param name="pBC">Bind context pointer (typically <c>IntPtr.Zero</c>).</param>
        ''' <param name="pwzUrl">Optional URL hint (can be file path or <c>Nothing</c>).</param>
        ''' <param name="pBuffer">Byte buffer containing the data to inspect.</param>
        ''' <param name="cbSize">Size of the buffer.</param>
        ''' <param name="pwzMimeProposed">Optional proposed MIME type (can be <c>Nothing</c>).</param>
        ''' <param name="dwMimeFlags">Flags controlling MIME detection behavior.</param>
        ''' <param name="ppwzMimeOut">Receives a pointer to the detected MIME type string (must be freed with <c>Marshal.FreeCoTaskMem</c>).</param>
        ''' <param name="dwReserved">Reserved parameter (must be 0).</param>
        ''' <returns>HRESULT value; 0 indicates success.</returns>
        <DllImport("urlmon.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
        Private Function FindMimeFromData(
    ByVal pBC As IntPtr,
    <MarshalAs(UnmanagedType.LPWStr)> ByVal pwzUrl As String,
    <MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I1, SizeParamIndex:=3)> ByVal pBuffer As Byte(),
    ByVal cbSize As UInteger,
    <MarshalAs(UnmanagedType.LPWStr)> ByVal pwzMimeProposed As String,
    ByVal dwMimeFlags As UInteger,
    ByRef ppwzMimeOut As IntPtr,
    ByVal dwReserved As UInteger
) As Integer
        End Function

        ''' <summary>
        ''' Determines the MIME type of a file and encodes its content as a Base64 string.
        ''' </summary>
        ''' <param name="filePath">Path to the file to process.</param>
        ''' <returns>A tuple containing the detected MIME type and the Base64-encoded file content.</returns>
        ''' <exception cref="System.Exception">Thrown when file reading or MIME detection fails.</exception>
        Public Function GetFileMimeTypeAndBase64(
    ByVal filePath As String
) As (MimeType As String, EncodedData As String)
            Try
                ' 1) Sniff the MIME type.
                Dim mime As String = GetMimeType(filePath)

                ' 2) Read and Base64-encode the full file content.
                Dim bytes As Byte() = File.ReadAllBytes(filePath)
                Dim b64 As String = System.Convert.ToBase64String(bytes)

                Return (mime, b64)
            Catch ex As System.Exception
                Throw New System.Exception("Error determining MIME type or encoding data: " & ex.Message, ex)
            End Try
        End Function

        ''' <summary>
        ''' Uses <c>FindMimeFromData</c> to inspect the first 256 bytes of a file and return its MIME type.
        ''' </summary>
        ''' <param name="filePath">Path to the file to inspect.</param>
        ''' <returns>Detected MIME type string (e.g., "application/pdf").</returns>
        ''' <exception cref="System.Exception">Thrown when the file cannot be read or the API call fails.</exception>
        Private Function GetMimeType(ByVal filePath As String) As String
            Dim buffer(255) As Byte
            Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read)
                Dim read As Integer = fs.Read(buffer, 0, buffer.Length)
                If read = 0 Then
                    Throw New System.Exception("Unable to read from file: " & filePath)
                End If
            End Using

            Dim mimePtr As IntPtr = IntPtr.Zero
            Dim hr As Integer = FindMimeFromData(
        IntPtr.Zero,
        filePath,
        buffer,
        CUInt(buffer.Length),
        Nothing,
        0,
        mimePtr,
        0
    )

            If hr <> 0 Then
                Throw New System.Exception($"FindMimeFromData failed with HRESULT 0x{hr:X8}")
            End If

            Dim mime As String = Marshal.PtrToStringUni(mimePtr)
            Marshal.FreeCoTaskMem(mimePtr)

            Return mime
        End Function

    End Module

End Namespace