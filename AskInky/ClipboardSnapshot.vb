' Part of "Red Ink" (Ask Inky)
' Copyright (c) LawDigital AG, Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: ClipboardSnapshot.vb
' Purpose: Byte-exact clipboard snapshot and restore around a synthetic copy or paste.
' Architecture: Windows native host with embedded Office Web-based interface.
' Notes: Red Ink License and Ask Inky third-party attributions are embedded in LICENSE.txt.
' =============================================================================
' Part of Red Ink (AskInky). P18 W367: the lazy IDataObject of the old code could not restore delayed
' formats. This snapshot copies every HGLOBAL format as bytes before Ask Inky touches the clipboard and
' writes them back only when nobody else changed the clipboard meanwhile. Content is never logged.
Friend NotInheritable Class ClipboardSnapshot
    Private Const MaxFormatBytes As Long = 32L * 1024 * 1024
    Private Const MaxTotalBytes As Long = 96L * 1024 * 1024
    ' GDI handles (bitmap, metafile, palette, owner display) are not HGLOBAL; Windows synthesizes
    ' CF_TEXT/CF_OEMTEXT/CF_LOCALE/CF_DIBV5 from the stored CF_UNICODETEXT/CF_DIB.
    Private Shared ReadOnly Skipped As New System.Collections.Generic.HashSet(Of UInteger)({2UI, 3UI, 9UI, 14UI, &H80UI, &H82UI, &H83UI, &H8EUI})
    Private ReadOnly formats As New System.Collections.Generic.List(Of System.Collections.Generic.KeyValuePair(Of UInteger, Byte()))()
    Public ReadOnly Property Complete As Boolean
    Public ReadOnly Property IsEmpty As Boolean
        Get
            Return formats.Count = 0
        End Get
    End Property

    Private Sub New(complete As Boolean)
        Me.Complete = complete
    End Sub

    Friend Shared Function OpenWithRetry(Optional owner As System.IntPtr = Nothing) As Boolean
        For attempt As Integer = 0 To 19
            If NativeCapture.OpenClipboard(owner) Then Return True
            System.Threading.Thread.Sleep(25)
        Next
        Return False
    End Function

    Public Shared Function Capture() As ClipboardSnapshot
        If Not OpenWithRetry() Then Return New ClipboardSnapshot(False)
        Dim result As ClipboardSnapshot = Nothing
        Dim complete As Boolean = True
        Dim collected As New System.Collections.Generic.List(Of System.Collections.Generic.KeyValuePair(Of UInteger, Byte()))()
        Try
            Dim seen As New System.Collections.Generic.HashSet(Of UInteger)()
            Dim present As New System.Collections.Generic.List(Of UInteger)()
            Dim format As UInteger = NativeCapture.EnumClipboardFormats(0UI)
            While format <> 0UI AndAlso present.Count < 256
                If seen.Add(format) Then present.Add(format)
                format = NativeCapture.EnumClipboardFormats(format)
            End While
            Dim hasUnicode As Boolean = present.Contains(13UI)
            Dim hasDib As Boolean = present.Contains(8UI)
            Dim total As Long = 0
            For Each item As UInteger In present
                If Skipped.Contains(item) Then Continue For
                If hasUnicode AndAlso (item = 1UI OrElse item = 7UI OrElse item = 16UI) Then Continue For
                If hasDib AndAlso item = 17UI Then Continue For
                Dim handle As System.IntPtr = NativeCapture.GetClipboardData(item)
                If handle = System.IntPtr.Zero Then complete = False : Continue For
                Dim size As Long = CLng(NativeCapture.GlobalSize(handle).ToUInt64())
                If size <= 0 OrElse size > MaxFormatBytes OrElse total + size > MaxTotalBytes Then complete = False : Continue For
                Dim pointer As System.IntPtr = NativeCapture.GlobalLock(handle)
                If pointer = System.IntPtr.Zero Then complete = False : Continue For
                Try
                    Dim bytes(CInt(size) - 1) As Byte
                    System.Runtime.InteropServices.Marshal.Copy(pointer, bytes, 0, CInt(size))
                    collected.Add(New System.Collections.Generic.KeyValuePair(Of UInteger, Byte())(item, bytes))
                    total += size
                Finally
                    NativeCapture.GlobalUnlock(handle)
                End Try
            Next
        Catch ex As System.Exception
            complete = False
            ClientLog.Write("clipboard.snapshot.failed", ex)
        Finally
            NativeCapture.CloseClipboard()
        End Try
        result = New ClipboardSnapshot(complete)
        result.formats.AddRange(collected)
        Return result
    End Function

    ' Writes the snapshot back. The marker format asks clipboard history/monitors to ignore this restore,
    ' because the same content is already there from before the copy.
    Public Function Restore(owner As System.IntPtr) As Boolean
        If Not OpenWithRetry(owner) Then Return False
        Try
            If Not NativeCapture.EmptyClipboard() Then Return False
            For Each pair As System.Collections.Generic.KeyValuePair(Of UInteger, Byte()) In formats
                SetBytes(pair.Key, pair.Value)
            Next
            Dim ignore As UInteger = NativeCapture.RegisterClipboardFormat("ExcludeClipboardContentFromMonitorProcessing")
            If ignore <> 0UI Then SetBytes(ignore, New Byte() {0})
            Return True
        Catch ex As System.Exception
            ClientLog.Write("clipboard.restore.failed", ex)
            Return False
        Finally
            NativeCapture.CloseClipboard()
        End Try
    End Function

    Friend Shared Sub SetBytes(format As UInteger, bytes As Byte())
        Dim handle As System.IntPtr = NativeCapture.GlobalAlloc(NativeCapture.GMEM_MOVEABLE, New System.UIntPtr(CULng(System.Math.Max(1, bytes.Length))))
        If handle = System.IntPtr.Zero Then Return
        Dim pointer As System.IntPtr = NativeCapture.GlobalLock(handle)
        If pointer = System.IntPtr.Zero Then NativeCapture.GlobalFree(handle) : Return
        Try
            If bytes.Length > 0 Then System.Runtime.InteropServices.Marshal.Copy(bytes, 0, pointer, bytes.Length)
        Finally
            NativeCapture.GlobalUnlock(handle)
        End Try
        ' On success the system owns the memory; free it only when SetClipboardData failed.
        If NativeCapture.SetClipboardData(format, handle) = System.IntPtr.Zero Then NativeCapture.GlobalFree(handle)
    End Sub

    ' Reads CF_UNICODETEXT without creating a lazy managed data object.
    Public Shared Function ReadUnicodeText() As String
        If Not OpenWithRetry() Then Return Nothing
        Try
            Dim handle As System.IntPtr = NativeCapture.GetClipboardData(13UI)
            If handle = System.IntPtr.Zero Then Return Nothing
            Dim pointer As System.IntPtr = NativeCapture.GlobalLock(handle)
            If pointer = System.IntPtr.Zero Then Return Nothing
            Try
                Return System.Runtime.InteropServices.Marshal.PtrToStringUni(pointer)
            Finally
                NativeCapture.GlobalUnlock(handle)
            End Try
        Finally
            NativeCapture.CloseClipboard()
        End Try
    End Function

    ' Puts plain text (and nothing else) on the clipboard for a paste-back, marked as not for history.
    Public Shared Function WriteUnicodeText(value As String, owner As System.IntPtr) As Boolean
        If Not OpenWithRetry(owner) Then Return False
        Try
            If Not NativeCapture.EmptyClipboard() Then Return False
            SetBytes(13UI, System.Text.Encoding.Unicode.GetBytes(If(value, "") & ChrW(0)))
            For Each marker As String In {"ExcludeClipboardContentFromMonitorProcessing"}
                Dim id As UInteger = NativeCapture.RegisterClipboardFormat(marker)
                If id <> 0UI Then SetBytes(id, New Byte() {0})
            Next
            For Each marker As String In {"CanIncludeInClipboardHistory", "CanUploadToCloudClipboard"}
                Dim id As UInteger = NativeCapture.RegisterClipboardFormat(marker)
                If id <> 0UI Then SetBytes(id, System.BitConverter.GetBytes(0))
            Next
            Return True
        Finally
            NativeCapture.CloseClipboard()
        End Try
    End Function
End Class
