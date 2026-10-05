' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' Exact, durable output exclusions shared by scanners and publishers. Register before
' publishing; a directory basename is never a reason to exclude unrelated input.
Option Strict On
Option Explicit On
Option Infer On

Namespace SharedLibrary

    Public NotInheritable Partial Class GeneratedOutputRegistry

        Private Sub New()
        End Sub

        Private NotInheritable Class RegistryDocument
            Public Property Version As System.Int32 = 1
            Public Property Entries As New System.Collections.Generic.List(Of RegistryEntry)()
        End Class

        Private NotInheritable Class RegistryEntry
            Public Property RootPath As System.String = ""
            Public Property OwnerKey As System.String = ""
        End Class

        Private Shared ReadOnly Gate As New System.Object()
        Private Shared ReadOnly JsonSettings As New Newtonsoft.Json.JsonSerializerSettings With {
            .TypeNameHandling = Newtonsoft.Json.TypeNameHandling.None,
            .MetadataPropertyHandling = Newtonsoft.Json.MetadataPropertyHandling.Ignore,
            .MaxDepth = 16,
            .CheckAdditionalContent = True
        }
        Private Shared _cachedDocument As RegistryDocument = Nothing
        Private Shared _cachedWriteUtc As System.DateTime = System.DateTime.MinValue
        Private Shared _cachedLength As System.Int64 = -1

        Private Shared ReadOnly Property RegistryPath As System.String
            Get
                Dim applicationData As System.String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData)
                If System.String.IsNullOrWhiteSpace(applicationData) Then
                    Throw New System.IO.IOException("The per-user generated-output registry location is unavailable.")
                End If
                Dim path As System.String = System.IO.Path.Combine(applicationData, "RedInk", "GeneratedOutputs", "generated-output-roots.json")
                ' Include the immutable temporary registry name, its lock and marker.
                SemanticArchivePathGuard.RequireWindowsCompatiblePath(path, 37)
                Return path
            End Get
        End Property

        ''' <summary>
        ''' Durably excludes one exact generated root (or file) for all participating
        ''' processes of this user. Failure is surfaced before the publisher writes.
        ''' </summary>
        Public Shared Sub Register(rootPath As System.String, ownerKey As System.String)
            Dim normalized As System.String = NormalizePhysical(rootPath)
            If System.String.IsNullOrWhiteSpace(ownerKey) Then Throw New System.ArgumentException("An output owner is required.", NameOf(ownerKey))
            SyncLock Gate
                Dim path As System.String = RegistryPath
                Dim snapshot As RegistryDocument = ReadCurrent(path)
                If Contains(snapshot, normalized, ownerKey) Then
                    EnsureInitializationMarker(path)
                    Return
                End If
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path))
                Using fileLock As System.IO.FileStream = AcquireLock(path & ".lock")
                    snapshot = ReadDocument(path)
                    If Contains(snapshot, normalized, ownerKey) Then
                        Cache(snapshot, path)
                        EnsureInitializationMarker(path)
                        Return
                    End If
                    snapshot.Entries.Add(New RegistryEntry With {.RootPath = normalized, .OwnerKey = ownerKey})
                    WriteDocumentAtomic(path, snapshot)
                    Cache(snapshot, path)
                    EnsureInitializationMarker(path)
                End Using
            End SyncLock
        End Sub

        ''' <summary>
        ''' Maintenance only: call after the owner's generated files are reclaimed.
        ''' Unregistering an archive by itself must NOT unregister its output roots.
        ''' </summary>
        Public Shared Sub UnregisterOwner(ownerKey As System.String)
            If System.String.IsNullOrWhiteSpace(ownerKey) Then Throw New System.ArgumentException("An output owner is required.", NameOf(ownerKey))
            SyncLock Gate
                Dim path As System.String = RegistryPath
                If Not System.IO.File.Exists(path) Then Return
                Using fileLock As System.IO.FileStream = AcquireLock(path & ".lock")
                    Dim snapshot As RegistryDocument = ReadDocument(path)
                    Dim removed As System.Int32 = snapshot.Entries.RemoveAll(Function(entry) System.String.Equals(entry.OwnerKey, ownerKey, System.StringComparison.Ordinal))
                    If removed > 0 Then WriteDocumentAtomic(path, snapshot)
                    Cache(snapshot, path)
                End Using
            End SyncLock
        End Sub

        ''' <summary>
        ''' Path-segment containment, without excluding similarly named siblings.
        ''' A corrupt/unreadable registry defers ingestion; it never silently grants it.
        ''' </summary>
        Public Shared Function IsGeneratedPath(candidatePath As System.String) As System.Boolean
            If System.String.IsNullOrWhiteSpace(candidatePath) Then Return False
            Try
                Return IsGeneratedPathVerified(candidatePath)
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine("GeneratedOutputRegistry: ingestion deferred because exact output exclusions could not be validated: " & ex.Message)
                Return True
            End Try
        End Function

        ''' <summary>Physical exclusion/placement matching only; grants no source access.</summary>
        Public Shared Function IsPhysicalPathAtOrBelow(rootPath As System.String, candidatePath As System.String) As System.Boolean
            Return IsAtOrBelow(NormalizePhysical(candidatePath), NormalizePhysical(rootPath))
        End Function

        ''' <summary>
        ''' Returns a conclusive exclusion decision or throws when it is unknown.
        ''' Archive reconciliation must preserve this distinction: an unreadable
        ''' path/registry is not evidence that a previously seen source was removed.
        ''' </summary>
        Public Shared Function IsGeneratedPathVerified(candidatePath As System.String) As System.Boolean
            Dim normalized As System.String = NormalizePhysical(candidatePath)
            If IsAtOrBelow(normalized, NormalizePhysical(System.IO.Path.GetDirectoryName(RegistryPath))) Then Return True
            If IsMarkedNamespacePathCore(normalized) Then Return True
            SyncLock Gate
                For Each entry As RegistryEntry In ReadCurrent(RegistryPath).Entries
                    If IsAtOrBelow(normalized, entry.RootPath) Then Return True
                Next
            End SyncLock
            Return False
        End Function

        ' Registration may precede creation. Resolve the closest existing handle,
        ' then append only the missing suffix. Existing path aliases (mapped drives,
        ' UNC paths and substituted roots) consequently use the same physical name.
        ' This is exclusion matching only; it grants no source or filesystem access.
        Private Shared Function NormalizePhysical(path As System.String) As System.String
            Dim current As System.String = Normalize(path)
            If System.Environment.OSVersion.Platform <> System.PlatformID.Win32NT Then Return current
            Dim missingComponents As New System.Collections.Generic.Stack(Of System.String)()
            Do
                Dim openError As System.Int32
                Using handle As Microsoft.Win32.SafeHandles.SafeFileHandle = OpenPathForAttributes(
                    current, &H80UI, &H7UI, System.IntPtr.Zero, 3UI, &H2000000UI, System.IntPtr.Zero)
                    If handle IsNot Nothing AndAlso Not handle.IsInvalid Then
                        If missingComponents.Count > 0 Then
                            Dim information As NativeFileInformation
                            If Not GetFileInformationByHandle(handle, information) OrElse
                                (information.Attributes And CUInt(System.IO.FileAttributes.Directory)) = 0UI Then
                                Throw New System.IO.IOException("The generated-output path has no verifiable existing directory ancestor.")
                            End If
                        End If
                        Dim resolved As System.String = PhysicalHandlePath(handle)
                        While missingComponents.Count > 0
                            resolved = System.IO.Path.Combine(resolved, missingComponents.Pop())
                        End While
                        Return Normalize(resolved, False)
                    End If
                    openError = System.Runtime.InteropServices.Marshal.GetLastWin32Error()
                End Using
                ' Access, sharing, offline network and malformed-name failures are
                ' not treated as absent files. Callers defer ingestion or publication.
                If openError <> 2 AndAlso openError <> 3 Then
                    Throw New System.IO.IOException("The physical generated-output location could not be verified.", New System.ComponentModel.Win32Exception(openError))
                End If
                Dim parent As System.String = System.IO.Path.GetDirectoryName(current)
                If System.String.IsNullOrEmpty(parent) OrElse System.String.Equals(parent, current, System.StringComparison.OrdinalIgnoreCase) Then
                    Throw New System.IO.IOException("The generated-output path has no accessible existing filesystem root.", New System.ComponentModel.Win32Exception(openError))
                End If
                Dim component As System.String = System.IO.Path.GetFileName(current)
                If System.String.IsNullOrEmpty(component) OrElse component.EndsWith(".", System.StringComparison.Ordinal) OrElse component.EndsWith(" ", System.StringComparison.Ordinal) Then
                    Throw New System.IO.IOException("An unresolved generated-output path component has an ambiguous Win32 name.")
                End If
                missingComponents.Push(component)
                current = parent
            Loop
        End Function

        Private Shared Function PhysicalHandlePath(handle As Microsoft.Win32.SafeHandles.SafeFileHandle) As System.String
            Dim buffer As New System.Text.StringBuilder(32768)
            Dim length As System.UInt32 = GetFinalPathNameByHandle(handle, buffer, CUInt(buffer.Capacity), 0UI)
            If length = 0UI OrElse length >= CUInt(buffer.Capacity) Then
                Throw New System.IO.IOException("The filesystem did not return a verifiable generated-output name.", New System.ComponentModel.Win32Exception(System.Runtime.InteropServices.Marshal.GetLastWin32Error()))
            End If
            Dim result As System.String = buffer.ToString()
            If result.StartsWith("\\?\UNC\", System.StringComparison.OrdinalIgnoreCase) Then
                result = "\\" & result.Substring(8)
            ElseIf result.StartsWith("\\?\", System.StringComparison.Ordinal) Then
                result = result.Substring(4)
            End If
            If Not System.IO.Path.IsPathRooted(result) Then Throw New System.IO.IOException("The filesystem returned an unsupported generated-output path name.")
            Return Normalize(result, False)
        End Function

        <System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)>
        Private Structure NativeFileInformation
            Public Attributes As System.UInt32
            Public CreationTime As System.Runtime.InteropServices.ComTypes.FILETIME
            Public LastAccessTime As System.Runtime.InteropServices.ComTypes.FILETIME
            Public LastWriteTime As System.Runtime.InteropServices.ComTypes.FILETIME
            Public VolumeSerialNumber As System.UInt32
            Public FileSizeHigh As System.UInt32
            Public FileSizeLow As System.UInt32
            Public NumberOfLinks As System.UInt32
            Public FileIndexHigh As System.UInt32
            Public FileIndexLow As System.UInt32
        End Structure

        <System.Runtime.InteropServices.DllImport("kernel32.dll", EntryPoint:="CreateFileW", CharSet:=System.Runtime.InteropServices.CharSet.Unicode, ExactSpelling:=True, SetLastError:=True)>
        Private Shared Function OpenPathForAttributes(path As System.String, desiredAccess As System.UInt32, shareMode As System.UInt32,
                                                     securityAttributes As System.IntPtr, creationDisposition As System.UInt32,
                                                     flagsAndAttributes As System.UInt32, templateHandle As System.IntPtr) As Microsoft.Win32.SafeHandles.SafeFileHandle
        End Function

        <System.Runtime.InteropServices.DllImport("kernel32.dll", ExactSpelling:=True, SetLastError:=True)>
        Private Shared Function GetFileInformationByHandle(handle As Microsoft.Win32.SafeHandles.SafeFileHandle, ByRef information As NativeFileInformation) As System.Boolean
        End Function

        <System.Runtime.InteropServices.DllImport("kernel32.dll", EntryPoint:="GetFinalPathNameByHandleW", CharSet:=System.Runtime.InteropServices.CharSet.Unicode, ExactSpelling:=True, SetLastError:=True)>
        Private Shared Function GetFinalPathNameByHandle(handle As Microsoft.Win32.SafeHandles.SafeFileHandle, path As System.Text.StringBuilder, capacity As System.UInt32, flags As System.UInt32) As System.UInt32
        End Function

        Private Shared Function Normalize(path As System.String, Optional expandEnvironment As System.Boolean = True) As System.String
            If System.String.IsNullOrWhiteSpace(path) Then Throw New System.ArgumentException("An exact output path is required.", NameOf(path))
            Dim input As System.String = path.Trim()
            If expandEnvironment Then
                input = SharedMethods.ExpandEnvironmentVariables(input)
                If System.String.IsNullOrWhiteSpace(input) Then Throw New System.ArgumentException("The output path could not be expanded.", NameOf(path))
                If System.Text.RegularExpressions.Regex.IsMatch(input, "%[^%]+%", System.Text.RegularExpressions.RegexOptions.CultureInvariant, System.TimeSpan.FromMilliseconds(100)) Then
                    Throw New System.ArgumentException("The output path contains an unresolved environment variable or placeholder.", NameOf(path))
                End If
            End If
            Dim result As System.String = System.IO.Path.GetFullPath(input)
            Dim root As System.String = System.IO.Path.GetPathRoot(result)
            If System.String.Equals(result.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar),
                                    root.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar), System.StringComparison.OrdinalIgnoreCase) Then
                Return root
            End If
            Return result.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar)
        End Function

        Private Shared Function IsAtOrBelow(path As System.String, root As System.String) As System.Boolean
            If System.String.Equals(path, root, System.StringComparison.OrdinalIgnoreCase) Then Return True
            Dim prefix As System.String = root
            If Not prefix.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString(), System.StringComparison.Ordinal) AndAlso
               Not prefix.EndsWith(System.IO.Path.AltDirectorySeparatorChar.ToString(), System.StringComparison.Ordinal) Then
                prefix &= System.IO.Path.DirectorySeparatorChar
            End If
            Return path.StartsWith(prefix, System.StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function Contains(document As RegistryDocument, root As System.String, owner As System.String) As System.Boolean
            For Each entry As RegistryEntry In document.Entries
                If System.String.Equals(entry.RootPath, root, System.StringComparison.OrdinalIgnoreCase) AndAlso
                   System.String.Equals(entry.OwnerKey, owner, System.StringComparison.Ordinal) Then Return True
            Next
            Return False
        End Function

        Private Shared Function ReadCurrent(path As System.String) As RegistryDocument
            Dim exists As System.Boolean = True
            Try
                System.IO.File.GetAttributes(path)
            Catch ex As System.IO.FileNotFoundException
                exists = False
            Catch ex As System.IO.DirectoryNotFoundException
                exists = False
            End Try
            If Not exists Then
                If (_cachedDocument IsNot Nothing AndAlso _cachedDocument.Entries.Count > 0) OrElse
                   ExistsChecked(path & ".initialized") Then
                    Throw New System.IO.IOException("The previously registered output-exclusion file is missing.")
                End If
                Return New RegistryDocument()
            End If
            Dim info As New System.IO.FileInfo(path)
            If _cachedDocument IsNot Nothing AndAlso info.Length = _cachedLength AndAlso info.LastWriteTimeUtc = _cachedWriteUtc Then Return _cachedDocument
            ' The snapshot and its pathname metadata are paired while the same
            ' cross-process lock excludes atomic replacement. Tagging old bytes with
            ' a newer file's metadata would otherwise miss another host's exclusions.
            Using fileLock As System.IO.FileStream = AcquireLock(path & ".lock")
                Dim snapshot As RegistryDocument = ReadDocument(path)
                Cache(snapshot, path)
                Return snapshot
            End Using
        End Function

        Private Shared Function ReadDocument(path As System.String) As RegistryDocument
            Try
                System.IO.File.GetAttributes(path)
            Catch ex As System.IO.FileNotFoundException
                If ExistsChecked(path & ".initialized") Then Throw New System.IO.IOException("The initialized output-exclusion registry is missing.", ex)
                Return New RegistryDocument()
            Catch ex As System.IO.DirectoryNotFoundException
                Return New RegistryDocument()
            End Try
            Using stream As New System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.Read,
                                                     System.IO.FileShare.ReadWrite Or System.IO.FileShare.Delete)
                If stream.Length > 16L * 1024L * 1024L Then Throw New System.IO.InvalidDataException("The generated-output registry exceeds its safety limit.")
                Using reader As New System.IO.StreamReader(stream, New System.Text.UTF8Encoding(False, True), True)
                    Dim snapshot As RegistryDocument = Newtonsoft.Json.JsonConvert.DeserializeObject(Of RegistryDocument)(reader.ReadToEnd(), JsonSettings)
                    If snapshot Is Nothing OrElse snapshot.Version <> 1 OrElse snapshot.Entries Is Nothing Then
                        Throw New System.IO.InvalidDataException("The generated-output registry has an unsupported or invalid format.")
                    End If
                    For Each entry As RegistryEntry In snapshot.Entries
                        If entry Is Nothing OrElse System.String.IsNullOrWhiteSpace(entry.OwnerKey) Then Throw New System.IO.InvalidDataException("The generated-output registry contains an invalid entry.")
                        ' Persisted values are already physical names. Refreshing
                        ' the registry must not touch every recorded volume/share.
                        entry.RootPath = Normalize(entry.RootPath, False)
                    Next
                    Return snapshot
                End Using
            End Using
        End Function

        Private Shared Function ExistsChecked(path As System.String) As System.Boolean
            Try
                System.IO.File.GetAttributes(path)
                Return True
            Catch ex As System.IO.FileNotFoundException
                Return False
            Catch ex As System.IO.DirectoryNotFoundException
                Return False
            End Try
        End Function

        Private Shared Sub EnsureInitializationMarker(path As System.String)
            Dim marker As System.String = path & ".initialized"
            If System.IO.File.Exists(marker) Then Return
            Try
                Using stream As New System.IO.FileStream(marker, System.IO.FileMode.CreateNew, System.IO.FileAccess.Write, System.IO.FileShare.Read)
                    Dim bytes As System.Byte() = System.Text.Encoding.ASCII.GetBytes("generated-output-registry-v1")
                    stream.Write(bytes, 0, bytes.Length)
                    stream.Flush(True)
                End Using
            Catch ex As System.IO.IOException
                If Not System.IO.File.Exists(marker) Then Throw
            End Try
        End Sub

        Private Shared Sub Cache(snapshot As RegistryDocument, path As System.String)
            Dim info As New System.IO.FileInfo(path)
            _cachedDocument = snapshot
            _cachedWriteUtc = info.LastWriteTimeUtc
            _cachedLength = info.Length
        End Sub

        Private Shared Function AcquireLock(path As System.String) As System.IO.FileStream
            Dim timer As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew()
            Do
                Try
                    Return New System.IO.FileStream(path, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.ReadWrite, System.IO.FileShare.None)
                Catch ex As System.IO.IOException
                    If timer.Elapsed > System.TimeSpan.FromSeconds(5) Then Throw New System.IO.IOException("The generated-output registry is busy; publication was deferred.", ex)
                    System.Threading.Thread.Sleep(25)
                End Try
            Loop
        End Function

        Private Shared Sub WriteDocumentAtomic(path As System.String, document As RegistryDocument)
            Dim temporary As System.String = path & "." & System.Guid.NewGuid().ToString("N") & ".tmp"
            Try
                Dim bytes As System.Byte() = New System.Text.UTF8Encoding(False, True).GetBytes(Newtonsoft.Json.JsonConvert.SerializeObject(document, Newtonsoft.Json.Formatting.Indented, JsonSettings))
                Using stream As New System.IO.FileStream(temporary, System.IO.FileMode.CreateNew, System.IO.FileAccess.Write, System.IO.FileShare.None)
                    stream.Write(bytes, 0, bytes.Length)
                    stream.Flush(True)
                End Using
                If System.IO.File.Exists(path) Then
                    System.IO.File.Replace(temporary, path, Nothing)
                Else
                    System.IO.File.Move(temporary, path)
                End If
            Finally
                If System.IO.File.Exists(temporary) Then
                    Try
                        System.IO.File.Delete(temporary)
                    Catch ex As System.Exception
                        System.Diagnostics.Trace.WriteLine("GeneratedOutputRegistry: temporary registry cleanup deferred: " & ex.Message)
                    End Try
                End If
            End Try
        End Sub

    End Class

End Namespace
