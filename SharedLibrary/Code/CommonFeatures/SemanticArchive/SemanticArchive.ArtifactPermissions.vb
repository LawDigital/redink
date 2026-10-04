' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
Option Strict On
Option Explicit On

Namespace SharedLibrary
    Public NotInheritable Class SemanticArchivePermissionRepairResult
        Public Property Status As System.String = "unavailable"
        Public Property CheckedArtifacts As System.Int32
        Public Property RepairedArtifacts As System.Int32
        Public Property QuarantinedArtifacts As System.Int32
        Public Property ContinuationToken As System.String = ""
        Public Property SourcePermissionSignature As System.String = ""
        Public Property Diagnostic As System.String = ""
    End Class

    Public NotInheritable Partial Class SemanticArchiveArtifactPlanner
        Private NotInheritable Class RepairFrame
            Public Path As System.String
            Public Enumerator As System.Collections.Generic.IEnumerator(Of System.String)
        End Class

        Private NotInheritable Class RepairSession
            Implements System.IDisposable
            Public Location As SemanticArchiveArtifactLocation
            Public Frames As New System.Collections.Generic.Stack(Of RepairFrame)()
            Public ExpiresUtc As System.DateTimeOffset
            Public RootPending As System.Boolean = True
            Public Quarantine As System.Boolean
            Public BindingSignature As System.String
            Public IdentityVerified As System.Boolean
            Public QuarantineControlsPending As System.Boolean = True
            Public MutationRequired As System.Boolean
            Public ManifestRevision As System.Int64 = -1
            Public Sub Dispose() Implements System.IDisposable.Dispose
                While Frames.Count > 0
                    Dim frame As RepairFrame = Frames.Pop()
                    If frame.Enumerator IsNot Nothing Then frame.Enumerator.Dispose()
                End While
            End Sub
        End Class

        Private Shared ReadOnly RepairGate As New System.Object()
        Private Shared ReadOnly RepairSessions As New System.Collections.Generic.Dictionary(Of System.String, RepairSession)(System.StringComparer.Ordinal)

        ''' <summary>
        ''' Bounded, model-free reconciliation. Tokens are opaque, process-local and
        ''' expire; restart safely repeats idempotent ACL comparisons. Original ACLs
        ''' are never changed. A failed WRITE_DAC is reported, not called revocation.
        ''' </summary>
        Private Shared Function ReconcilePermissionsAtLocation(binding As SemanticArchiveSourceBinding, sourcePath As System.String, Optional continuationToken As System.String = "", Optional maximumArtifacts As System.Int32 = 64, Optional cancellationToken As System.Threading.CancellationToken = Nothing) As SemanticArchivePermissionRepairResult
            If maximumArtifacts < 1 OrElse maximumArtifacts > 1024 Then Throw New System.ArgumentOutOfRangeException(NameOf(maximumArtifacts))
            cancellationToken.ThrowIfCancellationRequested()
            Dim result As New SemanticArchivePermissionRepairResult()
            Dim bindingSignature As System.String = SemanticArchiveIdentity.HashBytes(System.Text.Encoding.UTF8.GetBytes(Newtonsoft.Json.JsonConvert.SerializeObject(binding)))
            SyncLock RepairGate
                Dim token As System.String = continuationToken
                Dim session As RepairSession = Nothing
                Try
                    ExpireRepairSessions()
                    If Not System.String.IsNullOrWhiteSpace(token) Then
                        If Not RepairSessions.TryGetValue(token, session) Then
                            result.Status = "restart_required"
                            result.Diagnostic = "The process-local rights cursor expired; restart this source's bounded reconciliation."
                            Return result
                        End If
                        If session.BindingSignature <> bindingSignature Then
                            FinishRepairSession(token)
                            result.Status = "restart_required"
                            result.Diagnostic = "The source binding changed; restart rights reconciliation."
                            Return result
                        End If
                        If Not System.String.Equals(session.Location.SourcePath, SemanticArchivePathGuard.CanonicalPath(sourcePath), System.StringComparison.OrdinalIgnoreCase) OrElse Not System.String.Equals(session.Location.SourceRoot, SemanticArchivePathGuard.CanonicalPath(binding.RootPath), System.StringComparison.OrdinalIgnoreCase) Then Throw New System.UnauthorizedAccessException("Rights cursor source scope mismatch.")
                    Else
                        Dim location As SemanticArchiveArtifactLocation = Nothing
                        Dim quarantine As System.Boolean = False
                        Try
                            location = Plan(binding, sourcePath)
                        Catch failure As System.UnauthorizedAccessException
                            Dim identity As System.String = SemanticArchivePathGuard.GetVerifiedSourceIdentityForMaintenance(binding.RootPath, sourcePath)
                            location = New SemanticArchiveArtifactLocation With {.SourcePath = SemanticArchivePathGuard.CanonicalPath(sourcePath), .SourceRoot = binding.RootPath, .SourceIdentity = identity, .IsShared = False}
                            quarantine = True
                        End Try
                        Dim historicalPrivateMode As System.Boolean = Not location.IsShared AndAlso binding.ArtifactPlacementMode = "private"
                        If Not location.IsShared Then
                            ' Changing future placement to private does not retire the
                            ' original's existing shared artifacts from rights upkeep.
                            ' Resolve only the exact slot in the currently configured
                            ' location; no unrelated or historical roots are searched.
                            Dim parent As System.String = If(System.String.IsNullOrWhiteSpace(binding.SharedArtifactRoot), binding.RootPath, binding.SharedArtifactRoot)
                            location.NamespaceDirectory = System.IO.Path.Combine(parent, ".redink-sa")
                            Dim slot As System.String = SemanticArchiveIdentity.HashBytes(System.Text.Encoding.UTF8.GetBytes(location.SourceIdentity)).Substring(0, 32)
                            location.ArtifactDirectory = System.IO.Path.Combine(location.NamespaceDirectory, slot.Substring(0, 2), slot)
                            If location.ArtifactDirectory.Length > 240 Then
                                ' Such a slot cannot have been published by this layout.
                                result.Status = If(historicalPrivateMode, "private", "no_artifacts")
                                Return result
                            End If
                            SemanticArchivePathGuard.RequireWindowsCompatiblePath(location.ArtifactDirectory)
                            location.ManifestPath = System.IO.Path.Combine(location.ArtifactDirectory, ".redink-sa.json")
                            location.IsShared = True
                            If Not historicalPrivateMode Then quarantine = True
                        End If
                        If Not EntryExists(location.ArtifactDirectory) Then
                            result.Status = If(historicalPrivateMode, "private", "no_artifacts")
                            Return result
                        End If
                        If Not GeneratedOutputRegistry.IsMarkedNamespacePath(location.ArtifactDirectory) Then Throw New System.UnauthorizedAccessException("An unmarked artifact namespace cannot be repaired.")
                        RequireSourceMarker(location)
                        If historicalPrivateMode AndAlso Not quarantine Then
                            Try
                                Dim currentPolicy As PermissionSnapshot = ReadSourcePermissions(location)
                                RequireSameProtectionDomain(location.SourceIdentity, System.IO.Path.GetDirectoryName(location.NamespaceDirectory))
                                location.SourcePermissionSignature = currentPolicy.Signature
                                location.ProtectionMode = "source-readers-producer-writer-v1"
                            Catch failure As System.Exception When TypeOf failure Is System.UnauthorizedAccessException OrElse TypeOf failure Is System.IO.IOException OrElse TypeOf failure Is System.NotSupportedException
                                quarantine = True
                            End Try
                        End If
                        session = New RepairSession With {.Location = location, .ExpiresUtc = System.DateTimeOffset.UtcNow.AddMinutes(5), .Quarantine = quarantine, .BindingSignature = bindingSignature, .IdentityVerified = True}
                        If Not quarantine Then
                            Try
                                Dim initialManifest As SemanticArchiveCooperativeManifest = ReadRepairManifest(location)
                                If initialManifest IsNot Nothing Then
                                    session.ManifestRevision = initialManifest.Revision
                                    session.MutationRequired = initialManifest.SourcePermissionSignature <> location.SourcePermissionSignature
                                End If
                            Catch failure As System.UnauthorizedAccessException
                                ' A prior quarantine intentionally blocks content. The
                                ' opaque source marker is already verified; restoration
                                ' still requires current source authorization and owner
                                ' authority on each generated object under the claim.
                                session.MutationRequired = True
                            End Try
                        Else
                            session.MutationRequired = True
                        End If
                        If RepairSessions.Count >= 16 Then
                            result.Status = "deferred"
                            result.Diagnostic = "The bounded rights cursor capacity is in use."
                            Return result
                        End If
                        token = System.Guid.NewGuid().ToString("N")
                        RepairSessions.Add(token, session)
                    End If
                    Dim snapshot As PermissionSnapshot = Nothing
                    If Not session.Quarantine Then snapshot = ReadSourcePermissions(session.Location)
                    If snapshot IsNot Nothing Then
                        If session.Location.SourcePermissionSignature <> snapshot.Signature Then Throw New System.UnauthorizedAccessException("Source permissions changed during reconciliation; restart with a current snapshot.")
                        result.SourcePermissionSignature = snapshot.Signature
                    Else
                        RequireMaintenanceIdentity(session.Location)
                    End If
                    If Not session.Quarantine Then RequireSourceMarker(session.Location)
                    Dim claimPath As System.String = System.IO.Path.Combine(session.Location.ArtifactDirectory, "claim.lock")
                    If session.MutationRequired AndAlso Not EntryExists(claimPath) Then
                        result.Status = "deferred"
                        result.Diagnostic = "A source claim must be initialized before shared rights maintenance."
                        FinishRepairSession(token)
                        Return result
                    End If
                    If session.MutationRequired Then SemanticArchivePathGuard.ValidateContainedPath(session.Location.ArtifactDirectory, claimPath, True)
                    ' Existing stale ACLs are precisely what is being repaired. The
                    ' source marker is checked first; opening the normal claim with
                    ' Share.None serializes repair with all cooperative publishers.
                    Using claim As System.IO.FileStream = If(session.MutationRequired, OpenRepairClaim(claimPath), Nothing)
                        While result.CheckedArtifacts < maximumArtifacts
                            cancellationToken.ThrowIfCancellationRequested()
                            Dim nextPath As System.String = NextRepairPath(session)
                            If nextPath Is Nothing Then
                                If session.Quarantine AndAlso session.QuarantineControlsPending Then
                                    ' Fixed control files are completed together while the
                                    ' existing exclusive handle is still usable. At most
                                    ' three additional ACL checks supplement the payload
                                    ' budget; no original or content bytes are read.
                                    For Each control As System.String In New System.String() {System.IO.Path.Combine(session.Location.ArtifactDirectory, ".source.json"), claimPath, session.Location.ArtifactDirectory}
                                        result.CheckedArtifacts += 1
                                        RepairOneArtifact(session, Nothing, control, result, If(control = claimPath, claim, Nothing))
                                    Next
                                    session.QuarantineControlsPending = False
                                End If
                                If Not session.Quarantine AndAlso Not session.MutationRequired Then
                                    Dim finalManifest As SemanticArchiveCooperativeManifest = ReadRepairManifest(session.Location)
                                    Dim finalRevision As System.Int64 = If(finalManifest Is Nothing, -1, finalManifest.Revision)
                                    If finalRevision <> session.ManifestRevision OrElse (finalManifest IsNot Nothing AndAlso finalManifest.SourcePermissionSignature <> session.Location.SourcePermissionSignature) Then
                                        FinishRepairSession(token)
                                        result.Status = "restart_required"
                                        result.Diagnostic = "The canonical manifest changed during read-only rights verification."
                                        Return result
                                    End If
                                End If
                                If Not session.Quarantine AndAlso session.MutationRequired AndAlso EntryExists(session.Location.ManifestPath) Then
                                    Using cooperative As New SemanticArchiveCooperativeStore(session.Location)
                                        cooperative.RebindPermissionsAfterRepair(claim, cancellationToken)
                                    End Using
                                End If
                                If Not session.Quarantine Then RequireCurrentPermissions(session.Location)
                                result.Status = If(session.Quarantine, "quarantined", "complete")
                                FinishRepairSession(token)
                                Return result
                            End If
                            result.CheckedArtifacts += 1
                            If Not RepairOneArtifact(session, snapshot, nextPath, result, If(System.String.Equals(nextPath, claimPath, System.StringComparison.OrdinalIgnoreCase), claim, Nothing)) Then
                                ' Restart under the normal exclusive source claim only
                                ' when comparison found an actual policy difference.
                                session.Dispose()
                                session.RootPending = True
                                session.MutationRequired = True
                                session.ExpiresUtc = System.DateTimeOffset.UtcNow.AddMinutes(5)
                                result.Status = "partial"
                                result.ContinuationToken = token
                                Return result
                            End If
                        End While
                    End Using
                    session.ExpiresUtc = System.DateTimeOffset.UtcNow.AddMinutes(5)
                    result.Status = "partial"
                    result.ContinuationToken = token
                    Return result
                Catch failure As System.OperationCanceledException
                    If Not System.String.IsNullOrWhiteSpace(token) Then FinishRepairSession(token)
                    Throw
                Catch failure As System.Exception When TypeOf failure Is System.IO.IOException OrElse TypeOf failure Is System.UnauthorizedAccessException OrElse TypeOf failure Is System.NotSupportedException OrElse TypeOf failure Is System.Security.SecurityException
                    If Not System.String.IsNullOrWhiteSpace(token) Then FinishRepairSession(token)
                    result.Status = "unrepairable"
                    result.Diagnostic = "Generated rights were not fully reconciled; shared reuse remains denied until protection is verified. " & failure.Message
                    Return result
                End Try
            End SyncLock
        End Function

        Private Shared Function OpenRepairClaim(path As System.String) As System.IO.FileStream
            Return New System.IO.FileStream(path, System.IO.FileMode.Open, System.Security.AccessControl.FileSystemRights.Read Or System.Security.AccessControl.FileSystemRights.Write Or System.Security.AccessControl.FileSystemRights.ChangePermissions, System.IO.FileShare.None, 4096, System.IO.FileOptions.None, Nothing)
        End Function

        Private Shared Function ReadRepairManifest(location As SemanticArchiveArtifactLocation) As SemanticArchiveCooperativeManifest
            If Not EntryExists(location.ManifestPath) Then Return Nothing
            Using stream As System.IO.FileStream = SemanticArchivePathGuard.OpenContainedRead(location.ArtifactDirectory, location.ManifestPath)
                If stream.Length <= 0 OrElse stream.Length > 1024 * 1024 Then Throw New System.IO.InvalidDataException("Invalid rights manifest size.")
                Using reader As New System.IO.StreamReader(stream, New System.Text.UTF8Encoding(False, True), False)
                    Dim manifest As SemanticArchiveCooperativeManifest = Newtonsoft.Json.JsonConvert.DeserializeObject(Of SemanticArchiveCooperativeManifest)(reader.ReadToEnd())
                    If manifest Is Nothing OrElse manifest.Format <> "redink-semantic-artifacts" OrElse manifest.SchemaVersion <> 1 OrElse manifest.SourceIdentity <> location.SourceIdentity OrElse manifest.Revision < 0 OrElse manifest.Entries Is Nothing OrElse manifest.Entries.Count > 256 Then Throw New System.IO.InvalidDataException("Invalid rights manifest identity or schema.")
                    Return manifest
                End Using
            End Using
        End Function

        Private Shared Function NextRepairPath(session As RepairSession) As System.String
            If session.RootPending Then
                session.RootPending = False
                session.Frames.Push(New RepairFrame With {.Path = session.Location.ArtifactDirectory})
                If Not session.Quarantine Then Return session.Location.ArtifactDirectory
            End If
            While session.Frames.Count > 0
                Dim frame As RepairFrame = session.Frames.Peek()
                If frame.Enumerator Is Nothing Then frame.Enumerator = System.IO.Directory.EnumerateFileSystemEntries(frame.Path).GetEnumerator()
                If Not frame.Enumerator.MoveNext() Then
                    session.Frames.Pop().Enumerator.Dispose()
                    If Not session.Quarantine OrElse System.String.Equals(frame.Path, session.Location.ArtifactDirectory, System.StringComparison.OrdinalIgnoreCase) Then Continue While
                    Return frame.Path
                End If
                Dim path As System.String = SemanticArchivePathGuard.ValidateContainedPath(session.Location.ArtifactDirectory, frame.Enumerator.Current, True)
                If (System.IO.File.GetAttributes(path) And System.IO.FileAttributes.Directory) <> 0 Then
                    session.Frames.Push(New RepairFrame With {.Path = path})
                    If Not session.Quarantine Then Return path
                    Continue While
                End If
                If session.Quarantine AndAlso (System.String.Equals(path, System.IO.Path.Combine(session.Location.ArtifactDirectory, ".source.json"), System.StringComparison.OrdinalIgnoreCase) OrElse System.String.Equals(path, System.IO.Path.Combine(session.Location.ArtifactDirectory, "claim.lock"), System.StringComparison.OrdinalIgnoreCase)) Then Continue While
                Return path
            End While
            Return Nothing
        End Function

        Private Shared Sub RequireMaintenanceIdentity(location As SemanticArchiveArtifactLocation)
            If SemanticArchivePathGuard.GetVerifiedSourceIdentityForMaintenance(location.SourceRoot, location.SourcePath) <> location.SourceIdentity Then Throw New System.UnauthorizedAccessException("The source identity changed during restrictive maintenance.")
        End Sub

        Private Shared Function RepairOneArtifact(session As RepairSession, snapshot As PermissionSnapshot, path As System.String, result As SemanticArchivePermissionRepairResult, claim As System.IO.FileStream) As System.Boolean
            If snapshot IsNot Nothing Then
                Dim current As PermissionSnapshot = ReadSourcePermissions(session.Location)
                If current.Signature <> snapshot.Signature Then Throw New System.UnauthorizedAccessException("Source permissions changed during the repair batch.")
            Else
                RequireMaintenanceIdentity(session.Location)
            End If
            Dim directory As System.Boolean
            Dim actual As System.Security.AccessControl.FileSystemSecurity = If(claim Is Nothing, ArtifactSecurity(path, directory), DirectCast(claim.GetAccessControl(), System.Security.AccessControl.FileSystemSecurity))
            Dim producer As System.Security.Principal.SecurityIdentifier = TryCast(actual.GetOwner(GetType(System.Security.Principal.SecurityIdentifier)), System.Security.Principal.SecurityIdentifier)
            If producer Is Nothing Then Throw New System.UnauthorizedAccessException("Generated artifact owner is unavailable.")
            Dim desired As System.Security.AccessControl.FileSystemSecurity
            If session.Quarantine Then
                desired = PrivateQuarantineSecurity(producer, directory, System.String.Equals(path, System.IO.Path.Combine(session.Location.ArtifactDirectory, ".source.json"), System.StringComparison.OrdinalIgnoreCase) OrElse System.String.Equals(path, System.IO.Path.Combine(session.Location.ArtifactDirectory, "claim.lock"), System.StringComparison.OrdinalIgnoreCase))
            Else
                desired = ProjectSecurity(snapshot, producer, directory)
            End If
            Dim sections As System.Security.AccessControl.AccessControlSections = System.Security.AccessControl.AccessControlSections.Access Or System.Security.AccessControl.AccessControlSections.Owner
            If EquivalentSecurity(actual, desired) Then Return True
            If Not session.MutationRequired Then Return False
            If snapshot Is Nothing AndAlso Not producer.Equals(CurrentSid()) AndAlso Not producer.IsWellKnown(System.Security.Principal.WellKnownSidType.LocalSystemSid) AndAlso Not producer.IsWellKnown(System.Security.Principal.WellKnownSidType.BuiltinAdministratorsSid) Then Throw New System.UnauthorizedAccessException("The historical producer cannot be authorized for restrictive maintenance.")
            If snapshot IsNot Nothing AndAlso Not TrustedProducer(snapshot, producer) Then Throw New System.UnauthorizedAccessException("Foreign artifact owner cannot be authorized for rights repair.")
            If claim IsNot Nothing Then
                claim.SetAccessControl(DirectCast(desired, System.Security.AccessControl.FileSecurity))
            ElseIf directory Then
                System.IO.Directory.SetAccessControl(path, DirectCast(desired, System.Security.AccessControl.DirectorySecurity))
            Else
                System.IO.File.SetAccessControl(path, DirectCast(desired, System.Security.AccessControl.FileSecurity))
            End If
            Dim verified As System.Security.AccessControl.FileSystemSecurity = If(claim Is Nothing, ArtifactSecurity(path, directory), DirectCast(claim.GetAccessControl(), System.Security.AccessControl.FileSystemSecurity))
            If Not EquivalentSecurity(verified, desired) Then Throw New System.UnauthorizedAccessException("Generated ACL repair did not verify.")
            If session.Quarantine Then
                result.QuarantinedArtifacts += 1
            Else
                result.RepairedArtifacts += 1
            End If
            Return True
        End Function

        Private Shared Function PrivateQuarantineSecurity(producer As System.Security.Principal.SecurityIdentifier, directory As System.Boolean, opaqueControl As System.Boolean) As System.Security.AccessControl.FileSystemSecurity
            Dim security As System.Security.AccessControl.FileSystemSecurity = If(directory, DirectCast(New System.Security.AccessControl.DirectorySecurity(), System.Security.AccessControl.FileSystemSecurity), New System.Security.AccessControl.FileSecurity())
            security.SetAccessRuleProtection(True, False)
            security.SetOwner(producer)
            Dim inheritance As System.Security.AccessControl.InheritanceFlags = If(directory, System.Security.AccessControl.InheritanceFlags.ContainerInherit Or System.Security.AccessControl.InheritanceFlags.ObjectInherit, System.Security.AccessControl.InheritanceFlags.None)
            ' Quarantine removes ordinary-reader grants. It cannot claw back copies
            ' or restrict a privileged owner who can deliberately rewrite the DACL.
            For Each sid As System.Security.Principal.SecurityIdentifier In New System.Security.Principal.SecurityIdentifier() {New System.Security.Principal.SecurityIdentifier(System.Security.Principal.WellKnownSidType.LocalSystemSid, Nothing), New System.Security.Principal.SecurityIdentifier(System.Security.Principal.WellKnownSidType.BuiltinAdministratorsSid, Nothing)}
                security.AddAccessRule(New System.Security.AccessControl.FileSystemAccessRule(sid, System.Security.AccessControl.FileSystemRights.FullControl, inheritance, System.Security.AccessControl.PropagationFlags.None, System.Security.AccessControl.AccessControlType.Allow))
            Next
            Dim producerWrites As System.Security.AccessControl.FileSystemRights = System.Security.AccessControl.FileSystemRights.Write Or System.Security.AccessControl.FileSystemRights.ReadAttributes Or System.Security.AccessControl.FileSystemRights.ReadPermissions Or System.Security.AccessControl.FileSystemRights.ChangePermissions Or System.Security.AccessControl.FileSystemRights.Synchronize
            ' The fixed controls contain only a source-identity hash / random claim
            ' ID. Keeping their producer readable permits authorized restoration;
            ' no content, source path, title or semantic metadata is retained there.
            If opaqueControl Then
                producerWrites = producerWrites Or System.Security.AccessControl.FileSystemRights.Read
            Else
                security.AddAccessRule(New System.Security.AccessControl.FileSystemAccessRule(CurrentSid(), System.Security.AccessControl.FileSystemRights.ReadData, inheritance, System.Security.AccessControl.PropagationFlags.None, System.Security.AccessControl.AccessControlType.Deny))
            End If
            security.AddAccessRule(New System.Security.AccessControl.FileSystemAccessRule(CurrentSid(), producerWrites, inheritance, System.Security.AccessControl.PropagationFlags.None, System.Security.AccessControl.AccessControlType.Allow))
            Return security
        End Function

        Private Shared Sub FinishRepairSession(token As System.String)
            Dim session As RepairSession = Nothing
            If RepairSessions.TryGetValue(token, session) Then
                RepairSessions.Remove(token)
                session.Dispose()
            End If
        End Sub

        Private Shared Sub ExpireRepairSessions()
            Dim expired As New System.Collections.Generic.List(Of System.String)()
            For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, RepairSession) In RepairSessions
                If pair.Value.ExpiresUtc <= System.DateTimeOffset.UtcNow Then expired.Add(pair.Key)
            Next
            For Each token As System.String In expired
                FinishRepairSession(token)
            Next
        End Sub
    End Class
End Namespace
