' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ClipboardHelper.vb
' Purpose: Reads the current Windows clipboard content and converts the first supported
'          clipboard payload into a `(mimeType, base64)` pair.
'
' Architecture:
'  - STA access: Clipboard APIs are invoked on a dedicated STA thread to satisfy
'    Windows Forms clipboard threading requirements.
'  - Format precedence (first match wins):
'     1) Outlook attachment: "FileGroupDescriptorW"/"FileGroupDescriptor" + "FileContents"
'     2) Explorer file drop list: file path -> MimeHelper.GetFileMimeTypeAndBase64
'     3) Audio stream: Clipboard.GetAudioStream (assumed WAV)
'     4) Rich text: TextDataFormat.Rtf
'     5) HTML: TextDataFormat.Html
'     6) CSV: TextDataFormat.CommaSeparatedValue
'     7) Plain text: Clipboard.GetText
'     8) Bitmap image: Clipboard.GetImage (re-encoded as PNG)
'     9) Enhanced Metafile (EMF): CF_ENHMETAFILE -> Metafile -> Bitmap -> PNG
'  - Output: Base64 encoding is used for all supported payloads.
'  - Resource handling: Releases COM wrappers and native EMF handles to avoid holding
'    clipboard data objects longer than necessary.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary

    Friend Module ClipboardHelper

        ''' <summary>
        ''' Releases a COM object reference (if any) to avoid holding clipboard data objects
        ''' alive longer than necessary.
        ''' </summary>
        ''' <param name="obj">Candidate object that may be a COM wrapper.</param>
        Private Sub SafeReleaseCom(obj As Object)
            Try
                If obj IsNot Nothing AndAlso System.Runtime.InteropServices.Marshal.IsComObject(obj) Then
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(obj)
                End If
            Catch
                ' Ignore release failures; clipboard retrieval should not fail because cleanup did not succeed.
            End Try
        End Sub

        ''' <summary>
        ''' Sniffs the MIME type from raw bytes using the urlmon FindMimeFromData API.
        ''' Falls back to "application/octet-stream" on failure.
        ''' </summary>
        ''' <param name="data">The raw file bytes.</param>
        ''' <param name="fileNameHint">Optional file name hint (with extension) to assist detection.</param>
        ''' <returns>Detected MIME type string.</returns>
        Private Function SniffMimeFromBytes(data() As Byte, Optional fileNameHint As String = Nothing) As String
            Try
                Dim sniffSize As Integer = Math.Min(data.Length, 256)
                Dim buffer(sniffSize - 1) As Byte
                Array.Copy(data, buffer, sniffSize)

                ' Write bytes to a temporary file so MimeHelper.GetFileMimeTypeAndBase64 can use
                ' FindMimeFromData with a file-path hint (which improves detection for container formats
                ' like M4A/MP4 that share the same ftyp magic bytes).
                Dim tempPath As String = Nothing
                Try
                    Dim ext As String = ""
                    If Not String.IsNullOrWhiteSpace(fileNameHint) Then
                        ext = System.IO.Path.GetExtension(fileNameHint)
                    End If
                    If String.IsNullOrEmpty(ext) Then ext = ".tmp"
                    tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                                                       System.Guid.NewGuid().ToString("N") & ext)
                    System.IO.File.WriteAllBytes(tempPath, data)
                    Dim result = MimeHelper.GetFileMimeTypeAndBase64(tempPath)
                    Dim mime As String = result.MimeType.Trim()
                    If Not String.IsNullOrWhiteSpace(mime) AndAlso
                       Not mime.Equals("application/octet-stream", StringComparison.OrdinalIgnoreCase) Then
                        Return mime
                    End If
                Catch
                    ' Fall through to extension-based mapping below.
                Finally
                    Try
                        If tempPath IsNot Nothing AndAlso System.IO.File.Exists(tempPath) Then
                            System.IO.File.Delete(tempPath)
                        End If
                    Catch
                    End Try
                End Try
            Catch
                ' Ignore sniffing failures.
            End Try
            Return "application/octet-stream"
        End Function

        ''' <summary>
        ''' Returns a MIME type from a file extension. Returns Nothing if the extension is not recognized.
        ''' </summary>
        Private Function MimeFromExtension(ext As String) As String
            If String.IsNullOrWhiteSpace(ext) Then Return Nothing
            Select Case ext.ToLowerInvariant()
                Case ".wav" : Return "audio/wav"
                Case ".mp3" : Return "audio/mpeg"
                Case ".m4a", ".mp4a" : Return "audio/mp4"
                Case ".aac" : Return "audio/aac"
                Case ".ogg", ".oga" : Return "audio/ogg"
                Case ".flac" : Return "audio/flac"
                Case ".wma" : Return "audio/x-ms-wma"
                Case ".opus" : Return "audio/opus"
                Case ".webm" : Return "audio/webm"
                Case ".mp4" : Return "video/mp4"
                Case ".avi" : Return "video/x-msvideo"
                Case ".mov" : Return "video/quicktime"
                Case ".mkv" : Return "video/x-matroska"
                Case ".wmv" : Return "video/x-ms-wmv"
                Case ".txt" : Return "text/plain"
                Case ".png" : Return "image/png"
                Case ".jpg", ".jpeg" : Return "image/jpeg"
                Case ".gif" : Return "image/gif"
                Case ".bmp" : Return "image/bmp"
                Case ".tif", ".tiff" : Return "image/tiff"
                Case ".webp" : Return "image/webp"
                Case ".svg" : Return "image/svg+xml"
                Case ".pdf" : Return "application/pdf"
                Case ".doc" : Return "application/msword"
                Case ".docx" : Return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                Case ".xls" : Return "application/vnd.ms-excel"
                Case ".xlsx" : Return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                Case ".ppt" : Return "application/vnd.ms-powerpoint"
                Case ".pptx" : Return "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                Case ".csv" : Return "text/csv"
                Case ".json" : Return "application/json"
                Case ".xml" : Return "application/xml"
                Case ".html", ".htm" : Return "text/html"
                Case ".rtf" : Return "application/rtf"
                Case ".zip" : Return "application/zip"
                Case Else : Return Nothing
            End Select
        End Function

        ''' <summary>
        ''' Tries to read the current clipboard content and returns the first supported payload as a
        ''' MIME type and Base64-encoded content.
        ''' </summary>
        ''' <param name="mimeType">On success: MIME type of the extracted clipboard payload.</param>
        ''' <param name="base64">On success: Base64-encoded payload bytes (or UTF-8 bytes for text formats).</param>
        ''' <returns><see langword="True"/> if a supported clipboard payload was found; otherwise <see langword="False"/>.</returns>
        Friend Function TryGetClipboardObject(ByRef mimeType As String, ByRef base64 As String) As Boolean
            Dim succeeded As Boolean = False
            Dim localMimeType As String = Nothing
            Dim localBase64 As String = Nothing

            ' Clipboard APIs require an STA thread. All reads are performed inside this dedicated STA thread.
            Dim t As New System.Threading.Thread(
Sub()
    Try
        ' 1) Outlook attachment (FileGroupDescriptorW / FileGroupDescriptor + FileContents)
        Dim hasW = System.Windows.Forms.Clipboard.ContainsData("FileGroupDescriptorW")
        Dim hasA = System.Windows.Forms.Clipboard.ContainsData("FileGroupDescriptor")
        If hasW OrElse hasA Then
            Dim fmt = If(hasW, "FileGroupDescriptorW", "FileGroupDescriptor")
            Dim fgObj = System.Windows.Forms.Clipboard.GetData(fmt)
            Dim fgStream = TryCast(fgObj, System.IO.MemoryStream)
            Try
                If fgStream IsNot Nothing Then
                    Using reader As New System.IO.BinaryReader(fgStream, System.Text.Encoding.Unicode, leaveOpen:=False)
                        ' Read file name from FILEGROUPDESCRIPTOR structure (first item only).
                        reader.ReadInt32() ' itemCount (UINT cItems)

                        ' FILEDESCRIPTORW layout before cFileName:
                        '   dwFlags          4  bytes (DWORD)
                        '   clsid           16  bytes (CLSID)
                        '   sizel            8  bytes (SIZE)
                        '   pointl           8  bytes (POINTL)
                        '   dwFileAttributes 4  bytes (DWORD)
                        '   ftCreationTime   8  bytes (FILETIME)
                        '   ftLastAccessTime 8  bytes (FILETIME)
                        '   ftLastWriteTime  8  bytes (FILETIME)
                        '   nFileSizeHigh    4  bytes (DWORD)
                        '   nFileSizeLow     4  bytes (DWORD)
                        '                   ── total = 72 bytes
                        reader.BaseStream.Seek(72, System.IO.SeekOrigin.Current)

                        ' Read filename (up to 260 WCHARs).
                        Dim nameChars As New System.Collections.Generic.List(Of Char)
                        For i = 0 To 259
                            Dim ch As Char = reader.ReadChar()
                            If ch = ChrW(0) Then Exit For
                            nameChars.Add(ch)
                        Next
                        Dim fileName As String = New String(nameChars.ToArray())

                        ' Pull the raw attachment bytes.
                        Dim contentObj = System.Windows.Forms.Clipboard.GetData("FileContents")
                        Dim contentStream = TryCast(contentObj, System.IO.Stream)
                        Try
                            If contentStream IsNot Nothing Then
                                Using ms As New System.IO.MemoryStream()
                                    contentStream.CopyTo(ms)
                                    Dim bytes() As Byte = ms.ToArray()

                                    ' Prefer identifying WAV by RIFF/WAVE headers where applicable.
                                    If bytes.Length >= 12 AndAlso
                                       System.Text.Encoding.ASCII.GetString(bytes, 0, 4) = "RIFF" AndAlso
                                       System.Text.Encoding.ASCII.GetString(bytes, 8, 4) = "WAVE" Then

                                        localMimeType = "audio/wav"
                                    Else
                                        ' 1st: try extension-based mapping from the extracted file name.
                                        Dim ext = System.IO.Path.GetExtension(fileName).ToLowerInvariant()
                                        localMimeType = MimeFromExtension(ext)

                                        ' 2nd: if extension was empty or unrecognized, sniff from content bytes.
                                        If String.IsNullOrWhiteSpace(localMimeType) Then
                                            localMimeType = SniffMimeFromBytes(bytes, fileName)
                                        End If
                                    End If

                                    localBase64 = System.Convert.ToBase64String(bytes)
                                    succeeded = True
                                    Exit Sub
                                End Using
                            End If
                        Finally
                            ' Ensure we drop references that can keep the clipboard data object alive.
                            If contentStream IsNot Nothing Then contentStream.Dispose()
                            SafeReleaseCom(contentObj)
                        End Try
                    End Using
                End If
            Finally
                ' BinaryReader.Dispose closes fgStream; also release COM wrapper if any.
                SafeReleaseCom(fgObj)
            End Try
        End If

        ' 2) File-drop (Explorer copy)
        If System.Windows.Forms.Clipboard.ContainsFileDropList() Then
            Dim files = System.Windows.Forms.Clipboard.GetFileDropList()
            If files.Count > 0 Then
                Dim path = files(0)
                Dim mresult = MimeHelper.GetFileMimeTypeAndBase64(path)
                localMimeType = mresult.MimeType.Trim()
                localBase64 = mresult.EncodedData.Trim()
                succeeded = True
                Exit Sub
            End If
        End If

        ' 3) Raw WAV stream
        If System.Windows.Forms.Clipboard.ContainsAudio() Then
            Using audioStream As System.IO.Stream = System.Windows.Forms.Clipboard.GetAudioStream()
                Using ms As New System.IO.MemoryStream()
                    audioStream.CopyTo(ms)
                    localBase64 = System.Convert.ToBase64String(ms.ToArray())
                    localMimeType = "audio/wav"
                    succeeded = True
                    Exit Sub
                End Using
            End Using
        End If

        ' 4) RTF
        If System.Windows.Forms.Clipboard.ContainsText(System.Windows.Forms.TextDataFormat.Rtf) Then
            localMimeType = "application/rtf"
            localBase64 = System.Convert.ToBase64String(
                                System.Text.Encoding.UTF8.GetBytes(
                                    System.Windows.Forms.Clipboard.GetText(System.Windows.Forms.TextDataFormat.Rtf)))
            succeeded = True : Exit Sub
        End If

        ' 5) HTML
        If System.Windows.Forms.Clipboard.ContainsText(System.Windows.Forms.TextDataFormat.Html) Then
            localMimeType = "text/html"
            localBase64 = System.Convert.ToBase64String(
                                System.Text.Encoding.UTF8.GetBytes(
                                    System.Windows.Forms.Clipboard.GetText(System.Windows.Forms.TextDataFormat.Html)))
            succeeded = True : Exit Sub
        End If

        ' 6) CSV
        If System.Windows.Forms.Clipboard.ContainsText(System.Windows.Forms.TextDataFormat.CommaSeparatedValue) Then
            localMimeType = "text/csv"
            localBase64 = System.Convert.ToBase64String(
                                System.Text.Encoding.UTF8.GetBytes(
                                    System.Windows.Forms.Clipboard.GetText(System.Windows.Forms.TextDataFormat.CommaSeparatedValue)))
            succeeded = True : Exit Sub
        End If

        ' 7) Plain text
        If System.Windows.Forms.Clipboard.ContainsText() Then
            localMimeType = "text/plain"
            localBase64 = System.Convert.ToBase64String(
                                System.Text.Encoding.UTF8.GetBytes(
                                    System.Windows.Forms.Clipboard.GetText()))
            succeeded = True : Exit Sub
        End If

        ' 8) Image (Bitmap → PNG)
        If System.Windows.Forms.Clipboard.ContainsImage() Then
            Using img As System.Drawing.Image = System.Windows.Forms.Clipboard.GetImage()
                Using ms As New System.IO.MemoryStream()
                    img.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
                    localMimeType = "image/png"
                    localBase64 = System.Convert.ToBase64String(ms.ToArray())
                    succeeded = True : Exit Sub
                End Using
            End Using
        End If

        ' 9) EMF → Bitmap → PNG
        If NativeClipboardX.OpenClipboard(IntPtr.Zero) Then
            Try
                If NativeClipboardX.IsClipboardFormatAvailable(NativeClipboardX.CF_ENHMETAFILE) Then
                    Dim src As IntPtr = NativeClipboardX.GetClipboardData(NativeClipboardX.CF_ENHMETAFILE)
                    If src <> IntPtr.Zero Then
                        ' Copy the metafile handle so we can safely create a Metafile instance.
                        Dim clone As IntPtr = NativeClipboardX.CopyEnhMetaFile(src, Nothing)
                        Try
                            Using emf As New System.Drawing.Imaging.Metafile(clone, False)
                                Using bmp As New System.Drawing.Bitmap(emf.Width, emf.Height)
                                    Using g As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(bmp)
                                        g.DrawImage(emf, 0, 0)
                                        Using out As New System.IO.MemoryStream()
                                            bmp.Save(out, System.Drawing.Imaging.ImageFormat.Png)
                                            localMimeType = "image/png"
                                            localBase64 = System.Convert.ToBase64String(out.ToArray())
                                            succeeded = True
                                        End Using
                                    End Using
                                End Using
                            End Using
                        Finally
                            ' Always free the duplicated handle.
                            NativeClipboardX.DeleteEnhMetaFile(clone)
                        End Try
                        If succeeded Then Exit Sub
                    End If
                End If
            Finally
                NativeClipboardX.CloseClipboard()
            End Try
        End If

    Catch
        ' Suppress all exceptions to keep clipboard probing non-fatal for callers.
    End Try
End Sub)

            t.SetApartmentState(System.Threading.ApartmentState.STA)
            t.IsBackground = True
            t.Start()

            ' Wait up to 5 seconds; if the clipboard is locked, treat as failure.
            If Not t.Join(5000) Then
                Return False
            End If

            If succeeded Then
                mimeType = localMimeType
                base64 = localBase64
            End If

            Return succeeded
        End Function

    End Module
End Namespace