' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
Option Strict On
Option Explicit On

Namespace SharedLibrary
    Public NotInheritable Class SemanticArchiveArtifactLocation
        Public Property SourcePath As System.String = ""
        Public Property SourceIdentity As System.String = ""
        Public Property ManifestPath As System.String = ""
        Public Property ArtifactDirectory As System.String = ""
        Public Property NamespaceDirectory As System.String = ""
        Public Property IsShared As System.Boolean
        Public Property Diagnostic As System.String = ""
        Public Property SourcePermissionSignature As System.String = ""
        Public Property ProtectionMode As System.String = "private"
        Friend Property SourceRoot As System.String = ""
    End Class

    ''' <summary>
    ''' Shared artifacts contain only one original's derivatives. Their readers are
    ''' projected from that original; the producer owns generated writes. Shared
    ''' contributors are trusted writers: content hashes are not authentication.
    ''' Personal aggregate catalogs never use this shared ACL policy.
    ''' </summary>
    Public NotInheritable Partial Class SemanticArchiveArtifactPlanner
        Private Sub New()
        End Sub

        Public Shared Function Plan(binding As SemanticArchiveSourceBinding, sourcePath As System.String) As SemanticArchiveArtifactLocation
            If binding Is Nothing Then Throw New System.ArgumentNullException(NameOf(binding))
            Dim source As System.String = SemanticArchivePathGuard.RequireWindowsSourcePath(sourcePath)
            Dim identity As System.String = SemanticArchivePathGuard.GetVerifiedSourceIdentity(binding.RootPath, source)
            Dim result As New SemanticArchiveArtifactLocation With {.SourcePath = source, .SourceRoot = binding.RootPath, .SourceIdentity = identity}
            Dim slot As System.String = SemanticArchiveIdentity.HashBytes(System.Text.Encoding.UTF8.GetBytes(identity)).Substring(0, 32)
            If binding.ArtifactPlacementMode <> "auto" AndAlso binding.ArtifactPlacementMode <> "private" Then Throw New System.IO.InvalidDataException("Unknown artifact placement mode.")
            If binding.ArtifactPlacementMode = "auto" Then
                Try
                    Dim parent As System.String = If(System.String.IsNullOrWhiteSpace(binding.SharedArtifactRoot), binding.RootPath, binding.SharedArtifactRoot)
                    result.NamespaceDirectory = SemanticArchivePathGuard.RequireWindowsCompatiblePath(System.IO.Path.Combine(parent, ".redink-sa"), 100)
                    SemanticArchivePathGuard.ValidateContainedPath(parent, result.NamespaceDirectory, False)
                    Dim domainProof As System.String = GetProtectionDomainSignature(identity, parent)
                    Dim snapshot As PermissionSnapshot = ReadSourcePermissions(result, domainProof)
                    result.ArtifactDirectory = System.IO.Path.Combine(result.NamespaceDirectory, slot.Substring(0, 2), slot)
                    SemanticArchivePathGuard.RequireWindowsCompatiblePath(result.ArtifactDirectory, 64)
                    result.ManifestPath = System.IO.Path.Combine(result.ArtifactDirectory, ".redink-sa.json")
                    result.SourcePermissionSignature = snapshot.Signature
                    result.IsShared = True
                    result.ProtectionMode = "source-readers-producer-writer-v1"
                    Return result
                Catch failure As System.Exception When TypeOf failure Is System.UnauthorizedAccessException OrElse TypeOf failure Is System.IO.IOException OrElse TypeOf failure Is System.NotSupportedException
                    result.Diagnostic = "Shared placement unavailable; private shadow selected: " & failure.Message
                End Try
            End If
            result.NamespaceDirectory = ""
            result.ArtifactDirectory = SemanticArchivePathGuard.RequireWindowsCompatiblePath(System.IO.Path.Combine(SemanticArchiveStore.GetDefaultShadowRoot(binding), "c", slot.Substring(0, 16)), 64)
            result.ManifestPath = System.IO.Path.Combine(result.ArtifactDirectory, ".redink-sa.json")
            result.IsShared = False
            result.ProtectionMode = "private"
            Return result
        End Function

        ''' <summary>Read-only candidates: current root first, then an existing former file-adjacent slot.</summary>
        Public Shared Function GetReadLocations(binding As SemanticArchiveSourceBinding, sourcePath As System.String) As System.Collections.Generic.IReadOnlyList(Of SemanticArchiveArtifactLocation)
            If binding Is Nothing Then Throw New System.ArgumentNullException(NameOf(binding))
            Dim results As New System.Collections.Generic.List(Of SemanticArchiveArtifactLocation)()
            If binding.ArtifactPlacementMode = "private" Then Return results.AsReadOnly()
            Dim preferred As SemanticArchiveArtifactLocation = Plan(binding, sourcePath)
            If preferred.IsShared Then results.Add(preferred)
            Return results.AsReadOnly()
        End Function

        Private Shared Function CloneBindingForSharedRoot(binding As SemanticArchiveSourceBinding, sharedRoot As System.String) As SemanticArchiveSourceBinding
            Dim copy As SemanticArchiveSourceBinding = Newtonsoft.Json.JsonConvert.DeserializeObject(Of SemanticArchiveSourceBinding)(Newtonsoft.Json.JsonConvert.SerializeObject(binding))
            If copy Is Nothing Then Throw New System.IO.InvalidDataException("A source binding could not be copied for a verified artifact location.")
            copy.SharedArtifactRoot = sharedRoot
            Return copy
        End Function

        Public Shared Function CreatePrivateVersionDirectory(binding As SemanticArchiveSourceBinding, sourceIdentity As System.String,
                                                              versionId As System.String, Optional requiredChildPathLength As System.Int32 = 20) As System.String
            If requiredChildPathLength < 0 OrElse requiredChildPathLength > 240 Then Throw New System.ArgumentOutOfRangeException(NameOf(requiredChildPathLength))
            Dim preferred As System.String = SemanticArchiveStore.GetDerivedRoot(binding)
            Dim fallback As System.String = SemanticArchiveStore.GetDefaultShadowRoot(binding)
            Dim lastFailure As System.Exception = Nothing
            For Each root As System.String In New System.String() {preferred, fallback}
                Try
                    Dim destination As System.String = SemanticArchivePathGuard.RequireWindowsCompatiblePath(System.IO.Path.Combine(root, "versions", ShortHash(sourceIdentity), ShortHash(versionId)), requiredChildPathLength)
                    GeneratedOutputRegistry.Register(root, "SemanticArchive:shadow:" & binding.BindingId)
                    SemanticArchiveStore.CreatePrivateDirectory(destination)
                    Return destination
                Catch failure As System.Exception When TypeOf failure Is System.UnauthorizedAccessException OrElse TypeOf failure Is System.IO.IOException OrElse TypeOf failure Is System.NotSupportedException
                    lastFailure = failure
                End Try
            Next
            Throw New System.IO.IOException("No writable private shadow location satisfies the Windows path and access policy.", lastFailure)
        End Function

        Public Shared Function IsRecognizedArtifactNamespace(path As System.String) As System.Boolean
            Return GeneratedOutputRegistry.IsMarkedNamespacePath(path)
        End Function

        Public Shared Sub CreateArtifactDirectory(location As SemanticArchiveArtifactLocation, path As System.String)
            Dim full As System.String = CheckLocationPath(location, path, False)
            If Not location.IsShared Then
                RequireCurrentSource(location)
                GeneratedOutputRegistry.Register(System.IO.Path.GetDirectoryName(System.IO.Path.GetDirectoryName(location.ArtifactDirectory)), "SemanticArchive:private-cooperative")
                SemanticArchiveStore.CreatePrivateDirectory(full)
                Return
            End If
            Dim snapshot As PermissionSnapshot = RequireCurrentPermissions(location)
            RequireStableSharedAncestors(System.IO.Path.GetDirectoryName(location.NamespaceDirectory), snapshot)
            If Not EntryExists(location.NamespaceDirectory) Then System.IO.Directory.CreateDirectory(location.NamespaceDirectory)
            GeneratedOutputRegistry.EnsureNamespaceMarker(location.NamespaceDirectory)
            ' Buckets contain only opaque identities. Each source slot and descendant
            ' receives its own protected ACL; inherited namespace readers never apply.
            Dim bucket As System.String = System.IO.Path.GetDirectoryName(location.ArtifactDirectory)
            If Not EntryExists(bucket) Then System.IO.Directory.CreateDirectory(bucket)
            SemanticArchivePathGuard.ValidateContainedPath(location.NamespaceDirectory, bucket, True)
            RequireStableSharedAncestors(bucket, snapshot)
            Dim producer As System.Security.Principal.SecurityIdentifier = CurrentSid()
            Dim current As System.String = location.ArtifactDirectory
            CreateSharedDirectory(current, snapshot, producer)
            EnsureSourceMarker(location, snapshot, producer)
            Dim relative As System.String = full.Substring(location.ArtifactDirectory.Length).TrimStart(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar)
            For Each component As System.String In relative.Split(New System.Char() {System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar}, System.StringSplitOptions.RemoveEmptyEntries)
                current = System.IO.Path.Combine(current, component)
                CreateSharedDirectory(current, snapshot, producer)
            Next
            RequireCurrentPermissions(location)
        End Sub

        Public Shared Function CreateArtifactFile(location As SemanticArchiveArtifactLocation, path As System.String) As System.IO.FileStream
            Dim full As System.String = CheckLocationPath(location, path, False)
            CreateArtifactDirectory(location, System.IO.Path.GetDirectoryName(full))
            If Not location.IsShared Then Return SemanticArchiveStore.CreatePrivateFile(full)
            Dim snapshot As PermissionSnapshot = RequireCurrentPermissions(location)
            Dim security As System.Security.AccessControl.FileSecurity = DirectCast(ProjectSecurity(snapshot, CurrentSid(), False), System.Security.AccessControl.FileSecurity)
            ' Explicit creation security prevents disclosure through inherited ACLs
            ' even if a writable ancestor is replaced between validation and create.
            Dim stream As New System.IO.FileStream(full, System.IO.FileMode.CreateNew, System.Security.AccessControl.FileSystemRights.Read Or System.Security.AccessControl.FileSystemRights.Write, System.IO.FileShare.None, 65536, System.IO.FileOptions.WriteThrough, security)
            Try
                SemanticArchivePathGuard.ValidateContainedPath(location.ArtifactDirectory, full, True)
                RequireCurrentPermissions(location)
                Return stream
            Catch
                stream.Dispose()
                Throw
            End Try
        End Function

        Public Shared Sub RequireArtifact(location As SemanticArchiveArtifactLocation, path As System.String)
            Dim full As System.String = CheckLocationPath(location, path, True)
            If Not location.IsShared Then
                RequireCurrentSource(location)
                SemanticArchiveStore.RequirePrivateArtifact(full)
                Return
            End If
            Dim snapshot As PermissionSnapshot = RequireCurrentPermissions(location)
            RequireStableSharedAncestors(System.IO.Path.GetDirectoryName(location.ArtifactDirectory), snapshot)
            RequireSourceMarker(location)
            RequireProjectedSecurity(full, snapshot)
            RequireCurrentPermissions(location)
        End Sub

        Public Shared Function OpenClaim(location As SemanticArchiveArtifactLocation, path As System.String) As System.IO.FileStream
            Dim full As System.String = CheckLocationPath(location, path, False)
            CreateArtifactDirectory(location, System.IO.Path.GetDirectoryName(full))
            If Not EntryExists(full) Then
                Try
                    Return CreateArtifactFile(location, full)
                Catch failure As System.IO.IOException When (failure.HResult And &HFFFF) = 80 OrElse (failure.HResult And &HFFFF) = 183
                    ' Another producer created it. The existing lock is validated
                    ' below; it is never deleted or taken over.
                End Try
            End If
            RequireArtifact(location, full)
            If location.IsShared Then RequireWritableArtifactOwner(full, RequireCurrentPermissions(location))
            Dim stream As New System.IO.FileStream(full, System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite, System.IO.FileShare.None, 4096, System.IO.FileOptions.WriteThrough)
            Try
                SemanticArchivePathGuard.ValidateContainedPath(location.ArtifactDirectory, full, True)
                RequireCurrentSource(location)
                Return stream
            Catch
                stream.Dispose()
                Throw
            End Try
        End Function

        Public Shared Sub AtomicWriteJson(location As SemanticArchiveArtifactLocation, path As System.String, value As System.Object)
            Dim full As System.String = CheckLocationPath(location, path, False)
            Dim staging As System.String = System.IO.Path.Combine(location.ArtifactDirectory, ".staging")
            CreateArtifactDirectory(location, staging)
            Dim temporary As System.String = CheckLocationPath(location, System.IO.Path.Combine(staging, System.Guid.NewGuid().ToString("N") & ".tmp"), False)
            Dim bytes As System.Byte() = New System.Text.UTF8Encoding(False, True).GetBytes(Newtonsoft.Json.JsonConvert.SerializeObject(value, Newtonsoft.Json.Formatting.None))
            Try
                Using stream As System.IO.FileStream = CreateArtifactFile(location, temporary)
                    stream.Write(bytes, 0, bytes.Length)
                    stream.Flush(True)
                End Using
                RequireArtifact(location, temporary)
                If EntryExists(full) Then
                    RequireArtifact(location, full)
                    If location.IsShared Then RequireWritableArtifactOwner(full, RequireCurrentPermissions(location))
                    System.IO.File.Replace(temporary, full, Nothing, True)
                Else
                    System.IO.File.Move(temporary, full)
                End If
                RequireArtifact(location, full)
            Finally
                If EntryExists(temporary) Then System.IO.File.Delete(temporary)
            End Try
        End Sub

        Private Shared Sub EnsureSourceMarker(location As SemanticArchiveArtifactLocation, snapshot As PermissionSnapshot, producer As System.Security.Principal.SecurityIdentifier)
            Dim marker As System.String = System.IO.Path.Combine(location.ArtifactDirectory, ".source.json")
            If Not EntryExists(marker) Then
                Dim security As System.Security.AccessControl.FileSecurity = DirectCast(ProjectSecurity(snapshot, producer, False), System.Security.AccessControl.FileSecurity)
                Dim bytes As System.Byte() = New System.Text.UTF8Encoding(False, True).GetBytes(Newtonsoft.Json.JsonConvert.SerializeObject(SemanticArchiveIdentity.HashBytes(System.Text.Encoding.UTF8.GetBytes(location.SourceIdentity))))
                Try
                    Using stream As New System.IO.FileStream(marker, System.IO.FileMode.CreateNew, System.Security.AccessControl.FileSystemRights.Write, System.IO.FileShare.None, 4096, System.IO.FileOptions.WriteThrough, security)
                        stream.Write(bytes, 0, bytes.Length)
                        stream.Flush(True)
                    End Using
                Catch failure As System.IO.IOException When (failure.HResult And &HFFFF) = 80 OrElse (failure.HResult And &HFFFF) = 183
                End Try
            End If
            RequireSourceMarker(location)
        End Sub

        Private Shared Sub RequireSourceMarker(location As SemanticArchiveArtifactLocation)
            Dim marker As System.String = System.IO.Path.Combine(location.ArtifactDirectory, ".source.json")
            Using stream As System.IO.FileStream = SemanticArchivePathGuard.OpenContainedRead(location.ArtifactDirectory, marker)
                If stream.Length <= 0 OrElse stream.Length > 16384 Then Throw New System.IO.InvalidDataException("Invalid source-slot identity marker.")
                Using reader As New System.IO.StreamReader(stream, New System.Text.UTF8Encoding(False, True), False)
                    Dim identity As System.String = Newtonsoft.Json.JsonConvert.DeserializeObject(Of System.String)(reader.ReadToEnd())
                    If Not System.String.Equals(identity, SemanticArchiveIdentity.HashBytes(System.Text.Encoding.UTF8.GetBytes(location.SourceIdentity)), System.StringComparison.Ordinal) Then Throw New System.UnauthorizedAccessException("Source-slot identity collision or substitution.")
                End Using
            End Using
        End Sub

        Private Shared Function CheckLocationPath(location As SemanticArchiveArtifactLocation, path As System.String, mustExist As System.Boolean) As System.String
            If location Is Nothing Then Throw New System.ArgumentNullException(NameOf(location))
            Dim full As System.String = SemanticArchivePathGuard.RequireWindowsCompatiblePath(path)
            Return SemanticArchivePathGuard.ValidateContainedPath(location.ArtifactDirectory, full, mustExist)
        End Function

        Private Shared Sub RequireCurrentSource(location As SemanticArchiveArtifactLocation)
            Dim actual As System.String = SemanticArchivePathGuard.GetVerifiedSourceIdentity(location.SourceRoot, location.SourcePath)
            If Not System.String.Equals(actual, location.SourceIdentity, System.StringComparison.Ordinal) Then Throw New System.UnauthorizedAccessException("The original source identity changed.")
        End Sub

        Private Shared Function ShortHash(value As System.String) As System.String
            Return SemanticArchiveIdentity.HashBytes(System.Text.Encoding.UTF8.GetBytes(If(value, ""))).Substring(0, 16)
        End Function

        Private Shared Function EntryExists(path As System.String) As System.Boolean
            Try
                System.IO.File.GetAttributes(path)
                Return True
            Catch failure As System.IO.FileNotFoundException
                Return False
            Catch failure As System.IO.DirectoryNotFoundException
                Return False
            End Try
        End Function
    End Class
End Namespace
