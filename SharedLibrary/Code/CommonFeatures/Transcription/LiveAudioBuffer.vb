' Part of Red Ink. Session-local PCM with no named backing file.
Option Explicit On
Option Strict On

Namespace Transcription
    Friend NotInheritable Class LiveAudioBuffer
        Implements System.IDisposable

        Private Const PageSize As System.Int32 = 1048576
        Private ReadOnly pages As New System.Collections.Generic.Dictionary(Of System.Int64, System.IO.MemoryMappedFiles.MemoryMappedFile)()
        Private firstPage As System.Int64
        Private writer As System.IO.MemoryMappedFiles.MemoryMappedViewAccessor
        Private reader As System.IO.MemoryMappedFiles.MemoryMappedViewAccessor
        Private readPage As System.Int64 = -1
        Private lengthBytes As System.Int64

        ' All access is serialized by the session. Only the two active views occupy
        ' address space, even when an outage spans hours in 32-bit Word.
        Public ReadOnly Property Length As System.Int64
            Get
                Return lengthBytes
            End Get
        End Property

        Public Sub Append(data As System.Byte(), count As System.Int32)
            Dim offset As System.Int32 = 0
            While offset < count
                Dim position As System.Int32 = CInt(lengthBytes Mod PageSize)
                If position = 0 Then
                    Dim page = System.IO.MemoryMappedFiles.MemoryMappedFile.CreateNew(Nothing, PageSize)
                    Try
                        Dim view = page.CreateViewAccessor()
                        If writer IsNot Nothing Then writer.Dispose()
                        writer = view
                        pages.Add(lengthBytes \ PageSize, page)
                    Catch
                        page.Dispose()
                        Throw
                    End Try
                End If
                Dim n = System.Math.Min(count - offset, PageSize - position)
                writer.WriteArray(position, data, offset, n)
                lengthBytes += n
                offset += n
            End While
        End Sub

        Public Function Read(position As System.Int64, maximum As System.Int32) As System.Byte()
            Dim count = CInt(System.Math.Min(lengthBytes - position, CLng(maximum)))
            If count = 0 Then Return System.Array.Empty(Of System.Byte)()
            Dim result(count - 1) As System.Byte
            Dim offset As System.Int32 = 0
            While offset < count
                Dim page = position \ PageSize
                If readPage <> page Then
                    If reader IsNot Nothing Then reader.Dispose()
                    reader = pages(page).CreateViewAccessor(0, PageSize, System.IO.MemoryMappedFiles.MemoryMappedFileAccess.Read)
                    readPage = page
                End If
                Dim start = CInt(position Mod PageSize)
                Dim n = System.Math.Min(count - offset, PageSize - start)
                reader.ReadArray(start, result, offset, n)
                offset += n
                position += n
            End While
            Return result
        End Function

        Public Sub ReleaseBefore(position As System.Int64)
            While firstPage < position \ PageSize
                pages(firstPage).Dispose()
                pages.Remove(firstPage)
                firstPage += 1
            End While
        End Sub

        Public Sub Dispose() Implements System.IDisposable.Dispose
            If reader IsNot Nothing Then reader.Dispose()
            If writer IsNot Nothing Then writer.Dispose()
            For Each page In pages.Values
                page.Dispose()
            Next
            pages.Clear()
        End Sub
    End Class
End Namespace
