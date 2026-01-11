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
                        reader.ReadInt32() ' itemCount

                        ' Skip fixed-size fields up to the start of cFileName.
                        reader.BaseStream.Seek(4 + 16 + 8 + 8 + 8 + 4 + 4, System.IO.SeekOrigin.Current)

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
                                        ' Fallback to extension-based MIME mapping derived from the extracted file name.
                                        Dim ext = System.IO.Path.GetExtension(fileName).ToLowerInvariant()
                                        Select Case ext
                                            Case ".wav" : localMimeType = "audio/wav"
                                            Case ".mp3" : localMimeType = "audio/mpeg"
                                            Case ".txt" : localMimeType = "text/plain"
                                            Case ".png" : localMimeType = "image/png"
                                            Case ".jpg", ".jpeg" : localMimeType = "image/jpeg"
                                            Case Else : localMimeType = "application/octet-stream"
                                        End Select
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