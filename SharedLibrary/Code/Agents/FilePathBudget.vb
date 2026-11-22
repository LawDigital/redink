' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' Allocate bounded physical copy names; callers retain original names and identities separately.
Option Explicit On
Option Strict On
Option Infer On

Namespace Agents
    Public NotInheritable Class FilePathBudget
        ' Office, shell and third-party readers do not share the CLR's long-path capability.
        ' Apply this conservative budget to generated paths only, never to readable originals.
        Public Const MaximumGeneratedPathLength As System.Int32 = 240
        Public Const MaximumComponentLength As System.Int32 = 255
        Public Const PathTooLongMessage As System.String = "Windows or Office cannot access this path at its current length. Check the configured storage folder or the explicitly selected location."

        ' Mixed Office tools can hand a working file to Excel, whose complete path limit is 218.
        Public Const ToolWorkingReserve As System.Int32 = MaximumGeneratedPathLength - 218

        Private Sub New()
        End Sub

        Public Shared Function GetFullPath(path As System.String) As System.String
            Try
                Return System.IO.Path.GetFullPath(path)
            Catch ex As System.IO.PathTooLongException
                Throw New System.IO.PathTooLongException(PathTooLongMessage, ex)
            End Try
        End Function

        Public Shared Function FitsGeneratedPath(path As System.String, Optional reservedSuffixLength As System.Int32 = 0) As System.Boolean
            If System.String.IsNullOrWhiteSpace(path) OrElse reservedSuffixLength < 0 Then Return False
            Try
                Dim full As System.String = System.IO.Path.GetFullPath(path)
                If CLng(full.Length) + reservedSuffixLength > MaximumGeneratedPathLength Then Return False
                For Each component As System.String In full.Split(New System.Char() {System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar}, System.StringSplitOptions.RemoveEmptyEntries)
                    If component.Length > MaximumComponentLength Then Return False
                Next
                Return True
            Catch ex As System.IO.PathTooLongException
                Return False
            End Try
        End Function

        Public Shared Sub RequireGeneratedPath(path As System.String, Optional reservedSuffixLength As System.Int32 = 0)
            If Not FitsGeneratedPath(path, reservedSuffixLength) Then Throw New System.IO.PathTooLongException(PathTooLongMessage)
        End Sub

        Public Shared Function FileNameBudget(directory As System.String, Optional reservedSuffixLength As System.Int32 = 0) As System.Int32
            If reservedSuffixLength < 0 Then Throw New System.ArgumentOutOfRangeException(NameOf(reservedSuffixLength))
            Dim root As System.String = GetFullPath(directory)
            Dim prefixLength As System.Int32 = System.IO.Path.Combine(root, "x").Length - 1
            Return CInt(System.Math.Min(MaximumComponentLength - CLng(reservedSuffixLength), MaximumGeneratedPathLength - CLng(prefixLength) - reservedSuffixLength))
        End Function

        ' Per-run scratch only. If TEMP itself is too deep, use this user's local
        ' application-data Temp directory. Never move explicit workspace or durable roots.
        Public Shared Function TemporaryWorkingRoot(prefix As System.String) As System.String
            Dim normal As System.String = System.IO.Path.GetTempPath()
            ' Allow a full session GUID, a short working filename and a writer's temporary suffix.
            Dim required As System.Int32 = prefix.Length + 32 + 1 + 48
            If FitsGeneratedPath(normal, required + 1) Then Return normal
            Dim local As System.String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData)
            If Not System.String.IsNullOrWhiteSpace(local) Then
                Dim alternate As System.String = System.IO.Path.Combine(local, "Temp")
                If FitsGeneratedPath(alternate, required + 1) Then
                    System.Diagnostics.Trace.TraceInformation("Using the per-user temporary directory for bounded tool paths: " & alternate)
                    Return alternate
                End If
            End If
            Return normal
        End Function

        ' Preserve the caller's overwrite/versioning policy for usable names. A shortened
        ' output gets its own physical leaf (including on parallel calls). Callers return
        ' this actual filename to path-addressed tools and keep source identities separately.
        Public Shared Function GeneratedFilePath(directory As System.String, fileName As System.String,
                                                 Optional reservedSuffixLength As System.Int32 = ToolWorkingReserve) As System.String
            Dim root As System.String = GetFullPath(directory)
            Dim budget As System.Int32 = FileNameBudget(root, reservedSuffixLength)
            Dim fitted As System.String = FitFileName(fileName, budget)
            If System.String.Equals(fitted, fileName, System.StringComparison.Ordinal) Then
                Return System.IO.Path.Combine(root, fileName)
            End If

            ' Keep a readable tail for derived names such as "_compare" and "_processed".
            Dim extension As System.String = System.IO.Path.GetExtension(fitted)
            Dim stem As System.String = fileName.TrimEnd("."c, " "c)
            stem = stem.Substring(0, stem.Length - extension.Length)
            Dim room As System.Int32 = budget - extension.Length
            Dim nonce As System.String = System.Guid.NewGuid().ToString("N")
            Dim proposed As System.String
            If room <= 17 Then
                proposed = nonce.Substring(0, room) & extension
            Else
                Dim tailLength As System.Int32 = System.Math.Min(24, System.Math.Min(stem.Length, (room - 17) \ 2))
                Dim start As System.Int32 = stem.Length - tailLength
                If tailLength > 0 AndAlso System.Char.IsLowSurrogate(stem(start)) Then start += 1
                Dim suffix As System.String = "~" & nonce.Substring(0, 16) & stem.Substring(start)
                proposed = FitFileName(fileName, budget, suffix)
            End If
            Dim result As System.String = UniqueCopyPath(root, proposed, reservedSuffixLength)
            System.Diagnostics.Trace.TraceInformation("Generated filename shortened: " & fileName & " -> " & result)
            Return result
        End Function

        ' Persistent stores overwrite by logical name. The owner persists the physical-to-
        ' logical map with its metadata; empty maps also support pre-existing plain files.
        Public Shared Function PersistentCopyPath(directory As System.String, logicalName As System.String,
                                                  physicalNames As System.Collections.Generic.IDictionary(Of System.String, System.String)) As System.String
            For Each pair In physicalNames
                If System.String.Equals(pair.Value, logicalName, System.StringComparison.OrdinalIgnoreCase) Then
                    If Not System.String.Equals(FitFileName(pair.Key, MaximumComponentLength), pair.Key, System.StringComparison.Ordinal) Then
                        Throw New System.IO.InvalidDataException("Invalid stored working-copy filename.")
                    End If
                    Return System.IO.Path.Combine(directory, pair.Key)
                End If
            Next

            If logicalName.Length <= MaximumComponentLength AndAlso
               System.String.Equals(FitFileName(logicalName, MaximumComponentLength), logicalName, System.StringComparison.Ordinal) AndAlso
               Not physicalNames.ContainsKey(logicalName) Then
                Dim existing As System.String = System.IO.Path.Combine(directory, logicalName)
                If System.IO.File.Exists(existing) Then Return existing
            End If
            Dim path As System.String = GeneratedFilePath(directory, logicalName, reservedSuffixLength:=0)
            Dim leaf As System.String = System.IO.Path.GetFileName(path)
            ' A literal new name can itself equal another document's private alias.
            If physicalNames.ContainsKey(leaf) Then
                path = UniqueCopyPath(directory, System.Guid.NewGuid().ToString("N") & System.IO.Path.GetExtension(logicalName))
                leaf = System.IO.Path.GetFileName(path)
            End If
            If Not System.String.Equals(leaf, logicalName, System.StringComparison.Ordinal) Then physicalNames.Add(leaf, logicalName)
            Return path
        End Function

        ' Only host-owned copies opt into compaction. Never apply this to an explicit
        ' user destination, a committed logical output name, or an original source path.
        Public Shared Function UniqueCopyPath(directory As System.String, fileName As System.String,
                                               Optional reservedSuffixLength As System.Int32 = 0) As System.String
            Dim logicalFileName As System.String = fileName
            Return UniqueCopyPath(directory, fileName, logicalFileName, reservedSuffixLength)
        End Function

        ' Callers may sanitize an untrusted physical leaf first. Preserve its complete
        ' logical name separately, including the collision suffix that the host allocates.
        Public Shared Function UniqueCopyPath(directory As System.String, fileName As System.String,
                                               ByRef logicalFileName As System.String,
                                               Optional reservedSuffixLength As System.Int32 = 0,
                                               Optional occupiedLogicalNames As System.Collections.Generic.ISet(Of System.String) = Nothing) As System.String
            Dim originalLogicalName As System.String = If(System.String.IsNullOrWhiteSpace(logicalFileName), fileName, logicalFileName)
            Dim root As System.String = GetFullPath(directory)
            Dim budget As System.Int32 = FileNameBudget(root, reservedSuffixLength)
            For counter As System.Int32 = 0 To 99999
                Dim suffix As System.String = If(counter = 0, System.String.Empty, "_" & counter.ToString(System.Globalization.CultureInfo.InvariantCulture))
                Dim candidateName As System.String = FitFileName(fileName, budget, suffix)
                Dim candidate As System.String = System.IO.Path.Combine(root, candidateName)
                RequireGeneratedPath(candidate, reservedSuffixLength)
                If Not System.IO.File.Exists(candidate) AndAlso Not System.IO.Directory.Exists(candidate) Then
                    ' Logical metadata can contain characters that the physical-name sanitizer removed.
                    Dim extensionIndex As System.Int32 = originalLogicalName.LastIndexOf("."c)
                    Dim separatorIndex As System.Int32 = System.Math.Max(originalLogicalName.LastIndexOf("/"c), originalLogicalName.LastIndexOf("\"c))
                    Dim logicalExtension As System.String = If(extensionIndex > separatorIndex AndAlso extensionIndex < originalLogicalName.Length - 1, originalLogicalName.Substring(extensionIndex), System.String.Empty)
                    Dim candidateLogicalName As System.String = originalLogicalName.Substring(0, originalLogicalName.Length - logicalExtension.Length) & suffix & logicalExtension
                    If occupiedLogicalNames IsNot Nothing Then
                        If Not occupiedLogicalNames.Contains(originalLogicalName) Then candidateLogicalName = originalLogicalName
                        If Not occupiedLogicalNames.Add(candidateLogicalName) Then Continue For
                    End If
                    logicalFileName = candidateLogicalName
                    If Not System.String.Equals(candidateName, logicalFileName, System.StringComparison.Ordinal) Then
                        System.Diagnostics.Trace.TraceInformation("The host allocated a bounded physical copy and retained its logical filename: " & logicalFileName & " -> " & candidate)
                    End If
                    Return candidate
                End If
            Next
            Throw New System.IO.IOException("The host cannot allocate another distinct working-copy name. Use a new session or job and retry.")
        End Function

        Public Shared Function FitFileName(fileName As System.String, maximumLength As System.Int32,
                                            Optional collisionSuffix As System.String = "") As System.String
            If System.String.IsNullOrWhiteSpace(fileName) OrElse fileName = "." OrElse fileName = ".." OrElse
               fileName.IndexOfAny(System.IO.Path.GetInvalidFileNameChars()) >= 0 OrElse
               fileName.IndexOf("/"c) >= 0 OrElse fileName.IndexOf("\"c) >= 0 Then
                Throw New System.ArgumentException("A working-copy name must be a filename without directories.", NameOf(fileName))
            End If
            fileName = fileName.TrimEnd("."c, " "c)
            If fileName.Length = 0 Then Throw New System.ArgumentException("A working-copy name cannot consist only of periods or spaces.", NameOf(fileName))
            Dim suffix As System.String = If(collisionSuffix, System.String.Empty)
            If suffix.IndexOfAny(System.IO.Path.GetInvalidFileNameChars()) >= 0 Then Throw New System.ArgumentException("Invalid filename suffix.", NameOf(collisionSuffix))
            maximumLength = System.Math.Min(maximumLength, MaximumComponentLength)
            If suffix.Length = 0 AndAlso fileName.Length <= maximumLength AndAlso Not IsReservedFileName(fileName) Then Return fileName
            Dim extension As System.String = System.IO.Path.GetExtension(fileName)
            Dim stem As System.String = fileName.Substring(0, fileName.Length - extension.Length)
            Dim stemBudget As System.Int32 = maximumLength - extension.Length - suffix.Length
            If stemBudget < 1 Then Throw New System.IO.PathTooLongException("The host working directory leaves too little space for a file with the required extension.")
            stem = UnicodePrefix(stem, stemBudget).TrimEnd("."c, " "c)
            If stem.Length = 0 Then stem = UnicodePrefix("File", stemBudget)
            If stem.Length = 0 Then Throw New System.IO.PathTooLongException(PathTooLongMessage)
            If IsReservedFileName(stem & suffix & extension) Then stem = UnicodePrefix("_" & stem, stemBudget)
            Return stem & suffix & extension
        End Function

        Private Shared Function IsReservedFileName(value As System.String) As System.Boolean
            Return System.Text.RegularExpressions.Regex.IsMatch(value, "\A(?:CON|PRN|AUX|NUL|COM[1-9¹²³]|LPT[1-9¹²³])(?:\.|\z)",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.CultureInvariant)
        End Function

        ' The host must authorize the original before creating this private reader copy.
        ' Keep source identity, provenance and extraction-cache keys on the original path.
        Public Shared Function PrepareReaderInput(sourcePath As System.String) As ReaderInput
            Return New ReaderInput(sourcePath)
        End Function

        Public NotInheritable Class ReaderInput
            Implements System.IDisposable
            Public ReadOnly Property Path As System.String
            Private ReadOnly _directory As System.String
            Private _disposed As System.Boolean

            Friend Sub New(sourcePath As System.String)
                Dim full As System.String = GetFullPath(sourcePath)
                If FitsGeneratedPath(full, ToolWorkingReserve) Then
                    Path = full
                    Return
                End If
                Dim directory As System.String = System.IO.Path.Combine(TemporaryWorkingRoot("RIr_"), "RIr_" & System.Guid.NewGuid().ToString("N"))
                Dim extension As System.String = System.IO.Path.GetExtension(full)
                RequireGeneratedPath(System.IO.Path.Combine(directory, "Input" & extension))
                _directory = directory
                Try
                    System.IO.Directory.CreateDirectory(_directory)
                    Path = UniqueCopyPath(_directory, System.IO.Path.GetFileName(full), ToolWorkingReserve)
                    System.IO.File.Copy(full, Path, False)
                    System.Diagnostics.Trace.TraceInformation("The host uses a bounded reader copy and retains the original source identity: " & full & " -> " & Path)
                Catch
                    Dispose()
                    Throw
                End Try
            End Sub

            Public Sub Dispose() Implements System.IDisposable.Dispose
                If _disposed Then Return
                _disposed = True
                If System.String.IsNullOrWhiteSpace(_directory) Then Return
                Try
                    If System.IO.Directory.Exists(_directory) Then System.IO.Directory.Delete(_directory, True)
                Catch ex As System.Exception
                    System.Diagnostics.Trace.TraceWarning("The host could not remove its private reader copy: " & ex.Message)
                End Try
            End Sub
        End Class

        Public Shared Function UnicodePrefix(value As System.String, maximumLength As System.Int32) As System.String
            If maximumLength < 0 Then Throw New System.ArgumentOutOfRangeException(NameOf(maximumLength))
            Dim text As System.String = If(value, System.String.Empty)
            Dim count As System.Int32 = System.Math.Min(text.Length, maximumLength)
            If count > 0 AndAlso count < text.Length AndAlso System.Char.IsHighSurrogate(text(count - 1)) AndAlso System.Char.IsLowSurrogate(text(count)) Then count -= 1
            Return text.Substring(0, count)
        End Function
    End Class
End Namespace
