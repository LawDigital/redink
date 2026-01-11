' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ImageDecoder.vb
' Purpose: Extracts base64-encoded image data embedded in a JSON object (typically an LLM response),
'          determines its MIME type, and saves the decoded image to a uniquely named file on the user's Desktop.
'
' Architecture:
'  - Recursive Search: `FindImageData` traverses the JSON token tree and attempts to interpret string tokens
'    as base64-encoded image bytes.
'  - Validation: `TryGetImageData` validates candidate base64 content by attempting to load it as a
'    `System.Drawing.Image` from a `MemoryStream`.
'  - MIME Type Resolution: Attempts to read a sibling property named `mime_type`; if absent, falls back to
'    magic-byte detection for PNG/JPEG/GIF.
'  - File Output: `DecodeAndSaveImage` saves the decoded image as `AI_Image_NNN.<ext>` on the Desktop,
'    incrementing `NNN` to avoid overwriting existing files.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary

''' <summary>
''' Provides helper methods to locate base64-encoded image payloads in a JSON object and persist them as image files.
''' </summary>
Public Class ImageDecoder

    ''' <summary>
    ''' Recursively searches the provided JSON token for a string value containing a valid base64-encoded image.
    ''' On success, returns the decoded bytes and a resolved MIME type.
    ''' </summary>
    ''' <param name="token">Root token to search.</param>
    ''' <param name="imageBytes">Receives the decoded image bytes when found.</param>
    ''' <param name="mimeType">Receives the resolved MIME type when found.</param>
    ''' <returns><c>True</c> if a valid image was found; otherwise <c>False</c>.</returns>
    Private Shared Function FindImageData(token As JToken, ByRef imageBytes As Byte(), ByRef mimeType As String) As Boolean
        If token.Type = JTokenType.String Then
            If TryGetImageData(token, imageBytes, mimeType) Then
                Return True
            End If
        End If

        If token.HasValues Then
            For Each child In token.Children()
                If FindImageData(child, imageBytes, mimeType) Then
                    Return True
                End If
            Next
        End If

        Return False
    End Function

    ''' <summary>
    ''' Attempts to decode the token string as base64 and validate that it can be loaded as an image.
    ''' If validation succeeds, also attempts to resolve the MIME type from the token context or content.
    ''' </summary>
    ''' <param name="token">String token to interpret as base64.</param>
    ''' <param name="imageBytes">Receives the decoded image bytes.</param>
    ''' <param name="mimeType">Receives the resolved MIME type.</param>
    ''' <returns><c>True</c> if the token represents a valid image; otherwise <c>False</c>.</returns>
    Private Shared Function TryGetImageData(token As JToken, ByRef imageBytes As Byte(), ByRef mimeType As String) As Boolean
        Dim base64Str As String = token.ToString()
        Try
            Dim bytes As Byte() = System.Convert.FromBase64String(base64Str)
            ' Validate that the byte array represents a valid image.
            Using ms As New MemoryStream(bytes)
                Using img As Image = Image.FromStream(ms)
                    ' Successfully loaded image.
                End Using
            End Using

            imageBytes = bytes

            ' Try to get the MIME type from a nearby property.
            mimeType = GetMimeTypeFromParent(token)
            If String.IsNullOrEmpty(mimeType) Then
                mimeType = DetectMimeType(bytes)
            End If

            Return True

        Catch ex As Exception
            ' Not a valid base64 image.
            Debug.WriteLine("Decoding error: system.exception: " & ex.Message)
        End Try

        Return False
    End Function

    ''' <summary>
    ''' Attempts to read the MIME type from a sibling JSON property named <c>mime_type</c>.
    ''' </summary>
    ''' <param name="token">A token whose parent chain is inspected for a sibling <c>mime_type</c> property.</param>
    ''' <returns>The MIME type string if found; otherwise an empty string.</returns>
    Private Shared Function GetMimeTypeFromParent(token As JToken) As String
        If token.Parent IsNot Nothing AndAlso TypeOf token.Parent Is JProperty Then
            Dim parentProp As JProperty = CType(token.Parent, JProperty)
            Dim parentObj As JObject = TryCast(parentProp.Parent, JObject)
            If parentObj IsNot Nothing Then
                For Each prop As JProperty In parentObj.Properties()
                    If String.Equals(prop.Name, "mime_type", StringComparison.OrdinalIgnoreCase) Then
                        Return prop.Value.ToString()
                    End If
                Next
            End If
        End If
        Return String.Empty
    End Function

    ''' <summary>
    ''' Detects the MIME type of a byte array by checking common image magic bytes (PNG, JPEG, GIF).
    ''' </summary>
    ''' <param name="bytes">Image file bytes.</param>
    ''' <returns>The detected MIME type or an empty string if unknown.</returns>
    Private Shared Function DetectMimeType(bytes As Byte()) As String
        If bytes Is Nothing OrElse bytes.Length < 4 Then Return String.Empty

        ' Check for PNG (89 50 4E 47 0D 0A 1A 0A)
        If bytes.Length >= 8 AndAlso bytes(0) = &H89 AndAlso bytes(1) = &H50 AndAlso bytes(2) = &H4E AndAlso bytes(3) = &H47 Then
            Return "image/png"
        End If

        ' Check for JPEG (FF D8)
        If bytes(0) = &HFF AndAlso bytes(1) = &HD8 Then
            Return "image/jpeg"
        End If

        ' Check for GIF (GIF87a or GIF89a)
        If bytes.Length >= 6 Then
            Dim header As String = System.Text.Encoding.ASCII.GetString(bytes, 0, 6)
            If header = "GIF87a" OrElse header = "GIF89a" Then
                Return "image/gif"
            End If
        End If

        Return String.Empty
    End Function

    ''' <summary>
    ''' Maps a MIME type to a file extension used when saving to disk.
    ''' </summary>
    ''' <param name="mimeType">MIME type string (or short format name such as "png").</param>
    ''' <returns>File extension including a leading dot (e.g., ".png"), or empty string if unsupported.</returns>
    Private Shared Function GetExtensionFromMimeType(mimeType As String) As String
        Select Case mimeType.ToLower()
            Case "image/jpeg", "jpeg"
                Return ".jpg"
            Case "image/png", "png"
                Return ".png"
            Case "image/gif", "gif"
                Return ".gif"
            Case Else
                Return String.Empty
        End Select
    End Function


    ''' <summary>
    ''' Searches the provided JSON object for an embedded base64 image and, if found, saves it to a uniquely named file
    ''' on the current user's Desktop.
    ''' </summary>
    ''' <param name="jsonData">JSON object to search for base64-encoded image data.</param>
    ''' <returns>Full path of the saved image file, or an empty string if no supported image was found or saving failed.</returns>
    Public Shared Function DecodeAndSaveImage(jsonData As JObject) As String
        Dim imageBytes As Byte() = Nothing
        Dim mimeType As String = String.Empty

        ' Recursively search for a valid image in the JSON data.
        If Not FindImageData(jsonData, imageBytes, mimeType) Then
            Return ""
        End If

        Dim ext As String = GetExtensionFromMimeType(mimeType)
        If String.IsNullOrEmpty(ext) Then
            SharedMethods.ShowCustomMessageBox("The LLM returned an image or other object to your response, but the MIME type (i.e. the format) is not supported: " & mimeType)
            Return ""
        End If

        ' Determine the desktop path and generate a unique filename.
        Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        Dim fileNumber As Integer = 1
        Dim saveFilePath As String = String.Empty

        Do
            Dim fileName As String = "AI_Image_" & fileNumber.ToString("D3") & ext
            saveFilePath = Path.Combine(desktopPath, fileName)
            If Not File.Exists(saveFilePath) Then
                Exit Do
            End If
            fileNumber += 1
        Loop

        ' Save the image to the file.
        Try
            Using ms As New MemoryStream(imageBytes)
                Using img As Image = Image.FromStream(ms)
                    Select Case mimeType.ToLower()
                        Case "image/jpeg", "jpeg"
                            img.Save(saveFilePath, ImageFormat.Jpeg)
                        Case "image/png", "png"
                            img.Save(saveFilePath, ImageFormat.Png)
                        Case "image/gif", "gif"
                            img.Save(saveFilePath, ImageFormat.Gif)
                        Case Else
                            SharedMethods.ShowCustomMessageBox("The LLM returned an image or other object to your response, but the MIME type (i.e. the format) is not supported: " & mimeType)
                            Return ""
                    End Select
                End Using
            End Using

            Debug.WriteLine("Image saved to: " & saveFilePath)
            Return saveFilePath
        Catch ex As Exception
            Debug.WriteLine("Error saving image: system.exception: " & ex.Message)
            Return ""
        End Try
    End Function

End Class