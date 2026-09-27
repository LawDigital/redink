' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' File-backed text input and publication. Hashes always cover the exact file bytes,
' including a BOM. Character offsets are UTF-16 code units, not bytes or graphemes.
' No cache, content normalization, OCR, document renderer or organization policy here.
Option Strict On
Option Explicit On
Option Infer On

Namespace Agents

    Public NotInheritable Class TextFileInputException
        Inherits System.IO.IOException

        Public ReadOnly Property ErrorCode As System.String
        Public ReadOnly Property SizeBytes As System.Int64

        Public Sub New(code As System.String, message As System.String,
                       Optional actualSize As System.Int64 = 0)
            MyBase.New(message)
            ErrorCode = code
            SizeBytes = actualSize
        End Sub
    End Class

    Public NotInheritable Class TextFileSnapshot
        Public ReadOnly Property SourcePath As System.String
        Public ReadOnly Property Content As System.String
        Public ReadOnly Property SizeBytes As System.Int64
        Public ReadOnly Property Sha256 As System.String

        Private Sub New(path As System.String, text As System.String,
                        size As System.Int64, hash As System.String)
            SourcePath = path
            Content = text
            SizeBytes = size
            Sha256 = hash
        End Sub

        ''' <summary>
        ''' Reads once under the existing PathPolicy size/permission limits. FileShare.Read
        ''' prevents concurrent replacement/writing while bytes and their hash are captured.
        ''' Legacy callers keep BOM detection and replacement decoding; new document inputs
        ''' request strict Unicode decoding so invalid bytes are never silently substituted.
        ''' </summary>
        Public Shared Function Read(source As System.String,
                                    Optional strictDecoding As System.Boolean = False,
                                    Optional expectedSha256 As System.String = Nothing) As TextFileSnapshot
            ValidateExpectedHash(expectedSha256)
            Dim path As System.String = PathPolicy.Resolve(source, PathAccess.Read)
            If Not System.IO.File.Exists(path) Then
                Throw New TextFileInputException("not_found", "Text source file was not found.")
            End If

            Dim bytes As System.Byte()
            Using stream As New System.IO.FileStream(path, System.IO.FileMode.Open,
                                                     System.IO.FileAccess.Read, System.IO.FileShare.Read)
                If stream.Length > PathPolicy.MaxFileSizeBytes OrElse stream.Length > System.Int32.MaxValue Then
                    Throw New TextFileInputException("file_too_large", "Text source exceeds the configured size limit.", stream.Length)
                End If
                Using buffer As New System.IO.MemoryStream()
                    stream.CopyTo(buffer)
                    bytes = buffer.ToArray()
                End Using
            End Using

            Dim hash As System.String = ComputeHash(bytes)
            If Not System.String.IsNullOrWhiteSpace(expectedSha256) AndAlso
               Not System.String.Equals(hash, expectedSha256.Trim(), System.StringComparison.OrdinalIgnoreCase) Then
                Throw New TextFileInputException("source_hash_mismatch", "The text source no longer matches the expected SHA-256 snapshot.", bytes.LongLength)
            End If

            Dim text As System.String
            Try
                If strictDecoding Then
                    text = DecodeStrictUnicode(bytes)
                Else
                    ' Matches the previous File.ReadAllText(path, Encoding.UTF8) decoding.
                    Using buffer As New System.IO.MemoryStream(bytes, writable:=False)
                        Using reader As New System.IO.StreamReader(buffer, System.Text.Encoding.UTF8, detectEncodingFromByteOrderMarks:=True)
                            text = reader.ReadToEnd()
                        End Using
                    End Using
                End If
            Catch ex As System.Text.DecoderFallbackException
                Throw New TextFileInputException("invalid_text_encoding", "The source is not valid UTF-8 or BOM-marked Unicode text.", bytes.LongLength)
            End Try
            Return New TextFileSnapshot(path, text, bytes.LongLength, hash)
        End Function

        Private Shared Function DecodeStrictUnicode(bytes As System.Byte()) As System.String
            Dim encoding As System.Text.Encoding = New System.Text.UTF8Encoding(False, True)
            Dim skip As System.Int32 = 0
            ' UTF-32 LE must be checked before UTF-16 LE because their prefixes overlap.
            If bytes.Length >= 4 AndAlso bytes(0) = &HFF AndAlso bytes(1) = &HFE AndAlso bytes(2) = 0 AndAlso bytes(3) = 0 Then
                encoding = New System.Text.UTF32Encoding(False, True, True) : skip = 4
            ElseIf bytes.Length >= 4 AndAlso bytes(0) = 0 AndAlso bytes(1) = 0 AndAlso bytes(2) = &HFE AndAlso bytes(3) = &HFF Then
                encoding = New System.Text.UTF32Encoding(True, True, True) : skip = 4
            ElseIf bytes.Length >= 3 AndAlso bytes(0) = &HEF AndAlso bytes(1) = &HBB AndAlso bytes(2) = &HBF Then
                skip = 3
            ElseIf bytes.Length >= 2 AndAlso bytes(0) = &HFF AndAlso bytes(1) = &HFE Then
                encoding = New System.Text.UnicodeEncoding(False, True, True) : skip = 2
            ElseIf bytes.Length >= 2 AndAlso bytes(0) = &HFE AndAlso bytes(1) = &HFF Then
                encoding = New System.Text.UnicodeEncoding(True, True, True) : skip = 2
            End If
            Return encoding.GetString(bytes, skip, bytes.Length - skip)
        End Function

        Public Shared Sub ValidateExpectedHash(expectedSha256 As System.String)
            If System.String.IsNullOrWhiteSpace(expectedSha256) Then Return
            If Not System.Text.RegularExpressions.Regex.IsMatch(expectedSha256.Trim(), "\A[0-9a-fA-F]{64}\z") Then
                Throw New TextFileInputException("invalid_source_hash", "An expected SHA-256 must contain exactly 64 hexadecimal characters.")
            End If
        End Sub

        Public Shared Function ComputeHash(bytes As System.Byte()) As System.String
            Using algorithm As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                Return FormatHash(algorithm.ComputeHash(bytes))
            End Using
        End Function

        ''' <summary>Streaming source identity for large binary inputs; no text-size cap.</summary>
        Public Shared Function ComputeFileHash(source As System.String) As System.String
            Dim path As System.String = PathPolicy.Resolve(source, PathAccess.Read)
            Using stream As New System.IO.FileStream(path, System.IO.FileMode.Open,
                                                     System.IO.FileAccess.Read, System.IO.FileShare.Read)
                Using algorithm As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                    Return FormatHash(algorithm.ComputeHash(stream))
                End Using
            End Using
        End Function

        Private Shared Function FormatHash(bytes As System.Byte()) As System.String
            Return System.BitConverter.ToString(bytes).Replace("-", System.String.Empty).ToLowerInvariant()
        End Function

        ''' <summary>
        ''' Publishes complete UTF-8 text using a sibling temporary file and Move/Replace.
        ''' Existing output survives failures. There is deliberately no Delete+Move fallback.
        ''' The BOM matches the legacy Encoding.UTF8 writer. No new cap is imposed on exports.
        ''' </summary>
        Public Shared Function WriteUtf8Atomic(destination As System.String, content As System.String,
                                               overwrite As System.Boolean) As TextFileSnapshot
            Dim path As System.String = PathPolicy.Resolve(destination, PathAccess.Write)
            Dim encoding As New System.Text.UTF8Encoding(True, True)
            Dim size As System.Int64
            Dim hash As System.String
            Dim directory As System.String = System.IO.Path.GetDirectoryName(path)
            System.IO.Directory.CreateDirectory(directory)
            Dim temporaryPath As System.String = PathPolicy.Resolve(
                System.IO.Path.Combine(directory, ".ri-text-" & System.Guid.NewGuid().ToString("N") & ".tmp"), PathAccess.Write)
            Try
                Using stream As New System.IO.FileStream(temporaryPath, System.IO.FileMode.CreateNew,
                                                         System.IO.FileAccess.ReadWrite, System.IO.FileShare.None)
                    ' Stream large exports without allocating a second full encoded byte array.
                    Using writer As New System.IO.StreamWriter(stream, encoding, 4096, leaveOpen:=True)
                        writer.Write(If(content, System.String.Empty))
                    End Using
                    stream.Flush(True)
                    size = stream.Length
                    stream.Position = 0
                    Using algorithm As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                        hash = FormatHash(algorithm.ComputeHash(stream))
                    End Using
                End Using
                If overwrite AndAlso System.IO.File.Exists(path) Then
                    System.IO.File.Replace(temporaryPath, path, Nothing)
                Else
                    ' Also protects overwrite=false against a competing writer after the precheck.
                    System.IO.File.Move(temporaryPath, path)
                End If
                Return New TextFileSnapshot(path, If(content, System.String.Empty), size, hash)
            Finally
                If System.IO.File.Exists(temporaryPath) Then
                    Try
                        System.IO.File.Delete(temporaryPath)
                    Catch ex As System.Exception
                        System.Diagnostics.Debug.WriteLine("Text snapshot temporary-file cleanup failed: " & ex.GetType().FullName)
                    End Try
                End If
            End Try
        End Function
    End Class
End Namespace
