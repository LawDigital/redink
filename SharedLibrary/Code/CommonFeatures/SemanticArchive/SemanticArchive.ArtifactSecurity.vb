' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
Option Strict On
Option Explicit On

Namespace SharedLibrary
    Public NotInheritable Partial Class SemanticArchiveArtifactPlanner
        Private NotInheritable Class PermissionSnapshot
            Public Signature As System.String
            Public Rules As System.Security.AccessControl.AuthorizationRuleCollection
            Public SourceOwner As System.Security.Principal.SecurityIdentifier
        End Class

        Private Shared Function CurrentSid() As System.Security.Principal.SecurityIdentifier
            If System.Environment.OSVersion.Platform <> System.PlatformID.Win32NT Then Throw New System.PlatformNotSupportedException("Artifact ACL projection requires Windows.")
            Using identity As System.Security.Principal.WindowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent()
                If identity.User Is Nothing Then Throw New System.UnauthorizedAccessException("The artifact producer identity is unavailable.")
                Return New System.Security.Principal.SecurityIdentifier(identity.User.Value)
            End Using
        End Function

        Private Shared Function ReadSourcePermissions(location As SemanticArchiveArtifactLocation, Optional capturedDomainProof As System.String = Nothing) As PermissionSnapshot
            CurrentSid()
            RequireCurrentSource(location)
            Using source As System.IO.FileStream = SemanticArchivePathGuard.OpenContainedRead(location.SourceRoot, location.SourcePath)
                Dim attributes As System.IO.FileAttributes = System.IO.File.GetAttributes(location.SourcePath)
                If (attributes And System.IO.FileAttributes.Encrypted) <> 0 Then Throw New System.NotSupportedException("EFS rights cannot be represented by an ordinary artifact DACL.")
                Dim security As System.Security.AccessControl.FileSecurity = source.GetAccessControl()
                Dim supplemental As System.Byte() = ReadSupplementalPolicy(source.SafeFileHandle)
                Dim ancestorPolicy As System.String = ReadSourceAncestorPolicy(System.IO.Path.GetDirectoryName(location.SourceIdentity))
                Dim binary As System.Byte() = security.GetSecurityDescriptorBinaryForm()
                Dim raw As New System.Security.AccessControl.RawSecurityDescriptor(binary, 0)
                If raw.DiscretionaryAcl Is Nothing OrElse (raw.ControlFlags And System.Security.AccessControl.ControlFlags.DiscretionaryAclPresent) = 0 Then Throw New System.NotSupportedException("An absent source DACL cannot authorize shared generated storage.")
                For Each ace As System.Security.AccessControl.GenericAce In raw.DiscretionaryAcl
                    Dim common As System.Security.AccessControl.CommonAce = TryCast(ace, System.Security.AccessControl.CommonAce)
                    If common Is Nothing OrElse common.IsCallback OrElse (common.AceQualifier <> System.Security.AccessControl.AceQualifier.AccessAllowed AndAlso common.AceQualifier <> System.Security.AccessControl.AceQualifier.AccessDenied) Then Throw New System.NotSupportedException("Conditional, object, or unknown source ACEs require private artifact storage.")
                    If common.SecurityIdentifier Is Nothing OrElse common.SecurityIdentifier.Value.StartsWith("S-1-3-", System.StringComparison.Ordinal) Then Throw New System.NotSupportedException("Creator and owner-relative source ACEs require private artifact storage.")
                Next
                Dim signature As System.String = SemanticArchiveIdentity.HashBytes(System.Text.Encoding.UTF8.GetBytes(location.SourceIdentity & "|" & security.GetSecurityDescriptorSddlForm(System.Security.AccessControl.AccessControlSections.Access Or System.Security.AccessControl.AccessControlSections.Owner Or System.Security.AccessControl.AccessControlSections.Group) & "|" & System.Convert.ToBase64String(supplemental) & "|" & ancestorPolicy))
                Dim domainProof As System.String = If(capturedDomainProof, If(System.String.IsNullOrWhiteSpace(location.NamespaceDirectory), "", GetProtectionDomainSignature(location.SourceIdentity, System.IO.Path.GetDirectoryName(location.NamespaceDirectory))))
                ' The empty same-share proof preserves existing artifact signatures.
                ' Cross-share proof is freshly recomputed by every normal read/write
                ' permission check; no time-based allow cache grants disclosure.
                If domainProof.Length > 0 Then signature = SemanticArchiveIdentity.HashBytes(System.Text.Encoding.UTF8.GetBytes(signature & "|" & domainProof))
                Return New PermissionSnapshot With {
                    .Signature = signature,
                    .SourceOwner = raw.Owner,
                    .Rules = security.GetAccessRules(True, True, GetType(System.Security.Principal.SecurityIdentifier))
                }
            End Using
        End Function

        Private Shared Function RequireCurrentPermissions(location As SemanticArchiveArtifactLocation) As PermissionSnapshot
            Dim snapshot As PermissionSnapshot = ReadSourcePermissions(location)
            If Not System.String.Equals(snapshot.Signature, location.SourcePermissionSignature, System.StringComparison.Ordinal) Then Throw New System.UnauthorizedAccessException("source_permissions_changed: Shared publication and reuse wait for rights reconciliation.")
            Return snapshot
        End Function

        Private Shared Sub RequireSameProtectionDomain(sourceIdentity As System.String, destinationParent As System.String)
            GetProtectionDomainSignature(sourceIdentity, destinationParent)
        End Sub

        Private Shared Function ShareDomain(path As System.String) As System.String
            If Not path.StartsWith("\\", System.StringComparison.Ordinal) Then Return "local"
            Dim parts As System.String() = path.Substring(2).Split("\"c)
            If parts.Length < 2 OrElse parts(0).Length = 0 OrElse parts(1).Length = 0 Then Throw New System.NotSupportedException("The source share protection domain is unknown.")
            Return "\\" & parts(0) & "\" & parts(1)
        End Function

        Private Shared Function ProjectSecurity(snapshot As PermissionSnapshot, producer As System.Security.Principal.SecurityIdentifier, directory As System.Boolean) As System.Security.AccessControl.FileSystemSecurity
            Dim security As System.Security.AccessControl.FileSystemSecurity = If(directory, DirectCast(New System.Security.AccessControl.DirectorySecurity(), System.Security.AccessControl.FileSystemSecurity), New System.Security.AccessControl.FileSecurity())
            security.SetAccessRuleProtection(True, False)
            security.SetOwner(producer)
            Dim inheritance As System.Security.AccessControl.InheritanceFlags = If(directory, System.Security.AccessControl.InheritanceFlags.ContainerInherit Or System.Security.AccessControl.InheritanceFlags.ObjectInherit, System.Security.AccessControl.InheritanceFlags.None)
            ' Read grants and read denies alone are projected. Source write bits are
            ' never reinterpreted as directory-create or generated-file write grants.
            Dim readMask As System.Security.AccessControl.FileSystemRights = System.Security.AccessControl.FileSystemRights.ReadData Or System.Security.AccessControl.FileSystemRights.ReadAttributes Or System.Security.AccessControl.FileSystemRights.ReadExtendedAttributes Or System.Security.AccessControl.FileSystemRights.ReadPermissions
            For Each item As System.Security.AccessControl.AuthorizationRule In snapshot.Rules
                Dim rule As System.Security.AccessControl.FileSystemAccessRule = TryCast(item, System.Security.AccessControl.FileSystemAccessRule)
                If rule Is Nothing OrElse (rule.PropagationFlags And System.Security.AccessControl.PropagationFlags.InheritOnly) <> 0 Then Continue For
                Dim sid As System.Security.Principal.SecurityIdentifier = TryCast(rule.IdentityReference, System.Security.Principal.SecurityIdentifier)
                If sid Is Nothing Then Throw New System.NotSupportedException("Unresolved source reader rule.")
                Dim include As System.Boolean = If(rule.AccessControlType = System.Security.AccessControl.AccessControlType.Deny, (rule.FileSystemRights And readMask) <> 0, (rule.FileSystemRights And System.Security.AccessControl.FileSystemRights.ReadData) <> 0)
                If Not include Then Continue For
                Dim rights As System.Security.AccessControl.FileSystemRights = If(directory, System.Security.AccessControl.FileSystemRights.ReadAndExecute, System.Security.AccessControl.FileSystemRights.Read)
                security.AddAccessRule(New System.Security.AccessControl.FileSystemAccessRule(sid, rights, inheritance, System.Security.AccessControl.PropagationFlags.None, rule.AccessControlType))
            Next
            Dim sourceWriteMask As System.Security.AccessControl.FileSystemRights = System.Security.AccessControl.FileSystemRights.WriteData Or System.Security.AccessControl.FileSystemRights.AppendData Or System.Security.AccessControl.FileSystemRights.Delete
            Dim generatedWrites As System.Security.AccessControl.FileSystemRights = System.Security.AccessControl.FileSystemRights.Write Or System.Security.AccessControl.FileSystemRights.Delete Or System.Security.AccessControl.FileSystemRights.Synchronize
            For Each item As System.Security.AccessControl.AuthorizationRule In snapshot.Rules
                Dim rule As System.Security.AccessControl.FileSystemAccessRule = TryCast(item, System.Security.AccessControl.FileSystemAccessRule)
                ' Source mutation DENY rules apply to the original, not to derivatives.
                ' Only source-read DENYs above are transferred. Actual destination
                ' permissions still control whether this producer can create anything.
                If rule Is Nothing OrElse rule.AccessControlType <> System.Security.AccessControl.AccessControlType.Allow OrElse
                    (rule.PropagationFlags And System.Security.AccessControl.PropagationFlags.InheritOnly) <> 0 OrElse (rule.FileSystemRights And sourceWriteMask) = 0 Then Continue For
                Dim sid As System.Security.Principal.SecurityIdentifier = TryCast(rule.IdentityReference, System.Security.Principal.SecurityIdentifier)
                If sid Is Nothing Then Throw New System.NotSupportedException("Unresolved source writer rule.")
                If rule.AccessControlType = System.Security.AccessControl.AccessControlType.Allow AndAlso Not SourcePrincipalMayRead(snapshot, sid) Then Continue For
                ' Content writers can contribute and lock, but this projection never
                ' adds ReadData, ChangePermissions, or TakeOwnership.
                security.AddAccessRule(New System.Security.AccessControl.FileSystemAccessRule(sid, generatedWrites, inheritance, System.Security.AccessControl.PropagationFlags.None, rule.AccessControlType))
            Next
            ' A producer who can currently read the original may update derivatives.
            ' Preserve all read denies above, including any applying through groups.
            Dim producerWrites As System.Security.AccessControl.FileSystemRights = System.Security.AccessControl.FileSystemRights.Write Or System.Security.AccessControl.FileSystemRights.Delete Or System.Security.AccessControl.FileSystemRights.ReadPermissions Or System.Security.AccessControl.FileSystemRights.ChangePermissions Or System.Security.AccessControl.FileSystemRights.Synchronize
            security.AddAccessRule(New System.Security.AccessControl.FileSystemAccessRule(producer, producerWrites, inheritance, System.Security.AccessControl.PropagationFlags.None, System.Security.AccessControl.AccessControlType.Allow))
            For Each sid As System.Security.Principal.SecurityIdentifier In New System.Security.Principal.SecurityIdentifier() {New System.Security.Principal.SecurityIdentifier(System.Security.Principal.WellKnownSidType.LocalSystemSid, Nothing), New System.Security.Principal.SecurityIdentifier(System.Security.Principal.WellKnownSidType.BuiltinAdministratorsSid, Nothing)}
                security.AddAccessRule(New System.Security.AccessControl.FileSystemAccessRule(sid, System.Security.AccessControl.FileSystemRights.FullControl, inheritance, System.Security.AccessControl.PropagationFlags.None, System.Security.AccessControl.AccessControlType.Allow))
            Next
            Return security
        End Function

        Private Shared Function ArtifactSecurity(path As System.String, ByRef directory As System.Boolean) As System.Security.AccessControl.FileSystemSecurity
            directory = (System.IO.File.GetAttributes(path) And System.IO.FileAttributes.Directory) <> 0
            If directory Then Return System.IO.Directory.GetAccessControl(path, System.Security.AccessControl.AccessControlSections.Access Or System.Security.AccessControl.AccessControlSections.Owner)
            Return System.IO.File.GetAccessControl(path, System.Security.AccessControl.AccessControlSections.Access Or System.Security.AccessControl.AccessControlSections.Owner)
        End Function

        Private Shared Sub RequireProjectedSecurity(path As System.String, snapshot As PermissionSnapshot)
            Dim directory As System.Boolean
            Dim actual As System.Security.AccessControl.FileSystemSecurity = ArtifactSecurity(path, directory)
            Dim producer As System.Security.Principal.SecurityIdentifier = TryCast(actual.GetOwner(GetType(System.Security.Principal.SecurityIdentifier)), System.Security.Principal.SecurityIdentifier)
            If producer Is Nothing Then Throw New System.UnauthorizedAccessException("An artifact producer owner is unavailable.")
            Dim expected As System.Security.AccessControl.FileSystemSecurity = ProjectSecurity(snapshot, producer, directory)
            Dim sections As System.Security.AccessControl.AccessControlSections = System.Security.AccessControl.AccessControlSections.Access Or System.Security.AccessControl.AccessControlSections.Owner
            If Not EquivalentSecurity(actual, expected) Then Throw New System.UnauthorizedAccessException("artifact_permissions_stale: Generated protection does not match the current original reader policy.")
        End Sub

        Private Shared Function EquivalentSecurity(actual As System.Security.AccessControl.FileSystemSecurity, expected As System.Security.AccessControl.FileSystemSecurity) As System.Boolean
            If Not actual.AreAccessRulesProtected OrElse Not actual.GetOwner(GetType(System.Security.Principal.SecurityIdentifier)).Equals(expected.GetOwner(GetType(System.Security.Principal.SecurityIdentifier))) Then Return False
            Return SecurityRuleKey(actual) = SecurityRuleKey(expected)
        End Function

        Private Shared Function SecurityRuleKey(security As System.Security.AccessControl.FileSystemSecurity) As System.String
            Dim rules As New System.Collections.Generic.List(Of System.String)()
            For Each item As System.Security.AccessControl.AuthorizationRule In security.GetAccessRules(True, True, GetType(System.Security.Principal.SecurityIdentifier))
                Dim rule As System.Security.AccessControl.FileSystemAccessRule = TryCast(item, System.Security.AccessControl.FileSystemAccessRule)
                If rule Is Nothing OrElse rule.IsInherited Then Return "invalid-inherited-rule"
                rules.Add(rule.IdentityReference.Value & "|" & CInt(rule.AccessControlType).ToString(System.Globalization.CultureInfo.InvariantCulture) & "|" & CInt(rule.FileSystemRights).ToString(System.Globalization.CultureInfo.InvariantCulture) & "|" & CInt(rule.InheritanceFlags).ToString(System.Globalization.CultureInfo.InvariantCulture) & "|" & CInt(rule.PropagationFlags).ToString(System.Globalization.CultureInfo.InvariantCulture))
            Next
            rules.Sort(System.StringComparer.Ordinal)
            Return System.String.Join(System.Environment.NewLine, rules)
        End Function

        Private Shared Sub RequireWritableArtifactOwner(path As System.String, snapshot As PermissionSnapshot)
            Dim directory As System.Boolean
            Dim security As System.Security.AccessControl.FileSystemSecurity = ArtifactSecurity(path, directory)
            Dim owner As System.Security.Principal.SecurityIdentifier = TryCast(security.GetOwner(GetType(System.Security.Principal.SecurityIdentifier)), System.Security.Principal.SecurityIdentifier)
            If owner Is Nothing OrElse Not TrustedProducer(snapshot, owner) Then Throw New System.UnauthorizedAccessException("The existing generated owner is not a verified source producer; use private writable storage.")
        End Sub

        Private Shared Function TrustedProducer(snapshot As PermissionSnapshot, owner As System.Security.Principal.SecurityIdentifier) As System.Boolean
            If owner.Equals(CurrentSid()) OrElse owner.IsWellKnown(System.Security.Principal.WellKnownSidType.LocalSystemSid) OrElse owner.IsWellKnown(System.Security.Principal.WellKnownSidType.BuiltinAdministratorsSid) OrElse (snapshot.SourceOwner IsNot Nothing AndAlso owner.Equals(snapshot.SourceOwner)) Then Return True
            ' A foreign explicit SID grant alone cannot prove that a group deny
            ' does not apply to that producer's token. Fall back to private writes.
            For Each item As System.Security.AccessControl.AuthorizationRule In snapshot.Rules
                Dim denied As System.Security.AccessControl.FileSystemAccessRule = TryCast(item, System.Security.AccessControl.FileSystemAccessRule)
                If denied IsNot Nothing AndAlso denied.AccessControlType = System.Security.AccessControl.AccessControlType.Deny AndAlso (denied.PropagationFlags And System.Security.AccessControl.PropagationFlags.InheritOnly) = 0 AndAlso (denied.FileSystemRights And System.Security.AccessControl.FileSystemRights.ReadData) <> 0 Then Return False
            Next
            Dim allowed As System.Boolean = False
            For Each item As System.Security.AccessControl.AuthorizationRule In snapshot.Rules
                Dim rule As System.Security.AccessControl.FileSystemAccessRule = TryCast(item, System.Security.AccessControl.FileSystemAccessRule)
                If rule Is Nothing OrElse (rule.PropagationFlags And System.Security.AccessControl.PropagationFlags.InheritOnly) <> 0 OrElse Not owner.Equals(rule.IdentityReference) Then Continue For
                If (rule.FileSystemRights And System.Security.AccessControl.FileSystemRights.ReadData) <> 0 Then
                    If rule.AccessControlType = System.Security.AccessControl.AccessControlType.Deny Then Return False
                    allowed = True
                End If
            Next
            Return allowed
        End Function

        Private Shared Function SourcePrincipalMayRead(snapshot As PermissionSnapshot, identity As System.Security.Principal.SecurityIdentifier) As System.Boolean
            Dim allowed As System.Boolean = False
            For Each item As System.Security.AccessControl.AuthorizationRule In snapshot.Rules
                Dim rule As System.Security.AccessControl.FileSystemAccessRule = TryCast(item, System.Security.AccessControl.FileSystemAccessRule)
                If rule Is Nothing OrElse (rule.PropagationFlags And System.Security.AccessControl.PropagationFlags.InheritOnly) <> 0 Then Continue For
                If rule.AccessControlType = System.Security.AccessControl.AccessControlType.Deny AndAlso (rule.FileSystemRights And System.Security.AccessControl.FileSystemRights.ReadData) <> 0 Then Return False
                If rule.AccessControlType = System.Security.AccessControl.AccessControlType.Allow AndAlso rule.IdentityReference.Equals(identity) AndAlso (rule.FileSystemRights And System.Security.AccessControl.FileSystemRights.ReadData) <> 0 Then allowed = True
            Next
            Return allowed
        End Function

        Private Shared Sub RequireStableSharedAncestors(path As System.String, snapshot As PermissionSnapshot)
            Dim permitted As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            permitted.Add(CurrentSid().Value)
            permitted.Add("S-1-5-18")
            permitted.Add("S-1-5-32-544")
            If snapshot.SourceOwner IsNot Nothing Then permitted.Add(snapshot.SourceOwner.Value)
            Dim hasWriterDeny As System.Boolean = False
            For Each item As System.Security.AccessControl.AuthorizationRule In snapshot.Rules
                Dim rule As System.Security.AccessControl.FileSystemAccessRule = TryCast(item, System.Security.AccessControl.FileSystemAccessRule)
                If rule IsNot Nothing AndAlso rule.AccessControlType = System.Security.AccessControl.AccessControlType.Deny AndAlso (rule.PropagationFlags And System.Security.AccessControl.PropagationFlags.InheritOnly) = 0 AndAlso (rule.FileSystemRights And (System.Security.AccessControl.FileSystemRights.WriteData Or System.Security.AccessControl.FileSystemRights.AppendData)) <> 0 Then hasWriterDeny = True
            Next
            If Not hasWriterDeny Then
                For Each item As System.Security.AccessControl.AuthorizationRule In snapshot.Rules
                    Dim rule As System.Security.AccessControl.FileSystemAccessRule = TryCast(item, System.Security.AccessControl.FileSystemAccessRule)
                    If rule IsNot Nothing AndAlso rule.AccessControlType = System.Security.AccessControl.AccessControlType.Allow AndAlso (rule.PropagationFlags And System.Security.AccessControl.PropagationFlags.InheritOnly) = 0 AndAlso (rule.FileSystemRights And (System.Security.AccessControl.FileSystemRights.WriteData Or System.Security.AccessControl.FileSystemRights.AppendData)) <> 0 AndAlso SourcePrincipalMayRead(snapshot, DirectCast(rule.IdentityReference, System.Security.Principal.SecurityIdentifier)) Then permitted.Add(rule.IdentityReference.Value)
                Next
            End If
            Dim current As New System.IO.DirectoryInfo(path)
            While current IsNot Nothing
                If EntryExists(current.FullName) Then
                    ' Use the same physical-root proof as personal storage. A drive
                    ' letter/SUBST alias alone is never a local-volume-root proof.
                    Dim verifiedLocalRoot As System.Boolean = SemanticArchiveStore.IsVerifiedLocalVolumeRoot(current.FullName)
                    Dim verifiedWindowsRoot As System.Boolean = SemanticArchiveStore.IsVerifiedWindowsVolumeRoot(current.FullName)
                    Dim security As System.Security.AccessControl.DirectorySecurity = current.GetAccessControl(
                        System.Security.AccessControl.AccessControlSections.Access Or System.Security.AccessControl.AccessControlSections.Owner)
                    RequireStableSharedAncestorSecurity(current.FullName, security, permitted, verifiedLocalRoot, verifiedWindowsRoot)
                End If
                current = current.Parent
            End While
        End Sub

        Friend Shared Sub RequireStableSharedAncestorSecurity(path As System.String,
                      security As System.Security.AccessControl.DirectorySecurity,
                      permitted As System.Collections.Generic.HashSet(Of System.String),
                      verifiedLocalRoot As System.Boolean, verifiedWindowsRoot As System.Boolean)
            Dim trusted As New System.Collections.Generic.HashSet(Of System.String)(permitted, System.StringComparer.Ordinal)
            If verifiedWindowsRoot AndAlso verifiedLocalRoot Then
                trusted.Add("S-1-5-80-956008885-3418522649-1831038044-1853292631-2271478464")
            End If
            Dim owner As System.Security.Principal.SecurityIdentifier = TryCast(security.GetOwner(GetType(System.Security.Principal.SecurityIdentifier)), System.Security.Principal.SecurityIdentifier)
            If owner Is Nothing OrElse Not trusted.Contains(owner.Value) Then
                Throw SharedAncestorFailure(path, If(owner Is Nothing, "unresolved", owner.Value), "owner", "The ancestor owner is not a verified source/storage principal.")
            End If
            Dim raw As New System.Security.AccessControl.RawSecurityDescriptor(security.GetSecurityDescriptorBinaryForm(), 0)
            If raw.DiscretionaryAcl Is Nothing Then Throw SharedAncestorFailure(path, "Everyone", "NULL DACL", "The ancestor grants unrestricted access.")
            For Each ace As System.Security.AccessControl.GenericAce In raw.DiscretionaryAcl
                If (ace.AceFlags And System.Security.AccessControl.AceFlags.InheritOnly) <> 0 Then Continue For
                Dim common As System.Security.AccessControl.CommonAce = TryCast(ace, System.Security.AccessControl.CommonAce)
                If common Is Nothing OrElse common.IsCallback OrElse (common.AceQualifier <> System.Security.AccessControl.AceQualifier.AccessAllowed AndAlso
                    common.AceQualifier <> System.Security.AccessControl.AceQualifier.AccessDenied) Then
                    Throw SharedAncestorFailure(path, "unresolved", "complex ACE", "The ancestor policy cannot be safely evaluated.")
                End If
            Next
            Dim destructive As System.Security.AccessControl.FileSystemRights = System.Security.AccessControl.FileSystemRights.DeleteSubdirectoriesAndFiles Or
                System.Security.AccessControl.FileSystemRights.ChangePermissions Or System.Security.AccessControl.FileSystemRights.TakeOwnership
            ' DELETE targets the object itself, not its children. Only a physically
            ' verified local volume root is exempt; child deletion and ACL changes
            ' are never exempt. Normal directories and all UNC roots retain DELETE.
            If Not verifiedLocalRoot Then destructive = destructive Or System.Security.AccessControl.FileSystemRights.Delete
            For Each item As System.Security.AccessControl.AuthorizationRule In security.GetAccessRules(True, True, GetType(System.Security.Principal.SecurityIdentifier))
                Dim rule As System.Security.AccessControl.FileSystemAccessRule = TryCast(item, System.Security.AccessControl.FileSystemAccessRule)
                If rule Is Nothing OrElse (rule.PropagationFlags And System.Security.AccessControl.PropagationFlags.InheritOnly) <> 0 Then Continue For
                If rule.AccessControlType = System.Security.AccessControl.AccessControlType.Deny AndAlso (rule.FileSystemRights And System.Security.AccessControl.FileSystemRights.Traverse) <> 0 Then
                    Throw SharedAncestorFailure(path, rule.IdentityReference.Value, "Traverse denial", "Explicit ancestor traversal restrictions require private storage.")
                End If
                Dim blocking As System.Security.AccessControl.FileSystemRights = rule.FileSystemRights And destructive
                If rule.AccessControlType = System.Security.AccessControl.AccessControlType.Allow AndAlso blocking <> 0 AndAlso Not trusted.Contains(rule.IdentityReference.Value) Then
                    Throw SharedAncestorFailure(path, rule.IdentityReference.Value, blocking.ToString(),
                        "An unverified principal can replace shared artifact ancestry. Applicable Allow entry: " & rule.FileSystemRights.ToString() &
                        "; scope: " & If(verifiedLocalRoot, "verified local volume root", "ordinary directory or UNC/unverified root") & ".")
                End If
            Next
        End Sub

        Private Shared Function SharedAncestorFailure(path As System.String, principal As System.String,
                                                       rights As System.String, detail As System.String) As System.UnauthorizedAccessException
            Return New System.UnauthorizedAccessException("shared_ancestor_unsafe: " & detail & " Path: '" & path &
                "'; principal: " & principal & "; blocking rights: " & rights &
                ". Use a stable writable shared artifact root or private storage; existing ACLs were not changed.")
        End Function

        Private Shared Function ReadSourceAncestorPolicy(path As System.String) As System.String
            Dim policy As New System.Text.StringBuilder()
            Dim current As New System.IO.DirectoryInfo(path)
            While current IsNot Nothing
                Dim full As System.String = SemanticArchivePathGuard.ValidateContainedPath(current.FullName, current.FullName, True)
                Using handle As Microsoft.Win32.SafeHandles.SafeFileHandle = OpenSecurityDirectory(full, &H20080UI, &H3UI, System.IntPtr.Zero, 3UI, &H2200000UI, System.IntPtr.Zero)
                    If handle Is Nothing OrElse handle.IsInvalid Then Throw New System.NotSupportedException("Source ancestor security could not be verified.")
                    Dim supplemental As System.Byte() = ReadSupplementalPolicy(handle)
                    Dim security As System.Security.AccessControl.DirectorySecurity = current.GetAccessControl(System.Security.AccessControl.AccessControlSections.Access)
                    Dim raw As New System.Security.AccessControl.RawSecurityDescriptor(security.GetSecurityDescriptorBinaryForm(), 0)
                    If raw.DiscretionaryAcl Is Nothing Then Throw New System.NotSupportedException("An absent source ancestor DACL requires private storage.")
                    For Each ace As System.Security.AccessControl.GenericAce In raw.DiscretionaryAcl
                        If (ace.AceFlags And System.Security.AccessControl.AceFlags.InheritOnly) <> 0 Then Continue For
                        Dim common As System.Security.AccessControl.CommonAce = TryCast(ace, System.Security.AccessControl.CommonAce)
                        If common Is Nothing OrElse common.IsCallback OrElse (common.AceQualifier <> System.Security.AccessControl.AceQualifier.AccessAllowed AndAlso common.AceQualifier <> System.Security.AccessControl.AceQualifier.AccessDenied) Then Throw New System.NotSupportedException("Complex source ancestor policy requires private storage.")
                        If common.AceQualifier = System.Security.AccessControl.AceQualifier.AccessDenied AndAlso (common.AccessMask And CInt(System.Security.AccessControl.FileSystemRights.Traverse)) <> 0 Then Throw New System.NotSupportedException("Source ancestor traversal denial cannot be relocated safely.")
                    Next
                    policy.Append(full.ToUpperInvariant()).Append("|").Append(security.GetSecurityDescriptorSddlForm(System.Security.AccessControl.AccessControlSections.Access)).Append("|").Append(System.Convert.ToBase64String(supplemental)).AppendLine()
                End Using
                current = current.Parent
            End While
            Return SemanticArchiveIdentity.HashBytes(System.Text.Encoding.UTF8.GetBytes(policy.ToString()))
        End Function

        <System.Runtime.InteropServices.DllImport("kernel32.dll", EntryPoint:="CreateFileW", CharSet:=System.Runtime.InteropServices.CharSet.Unicode, SetLastError:=True)>
        Private Shared Function OpenSecurityDirectory(path As System.String, access As System.UInt32, share As System.UInt32, security As System.IntPtr, disposition As System.UInt32, flags As System.UInt32, template As System.IntPtr) As Microsoft.Win32.SafeHandles.SafeFileHandle
        End Function

        Private Shared Function ReadSupplementalPolicy(handle As Microsoft.Win32.SafeHandles.SafeFileHandle) As System.Byte()
            Dim owner As System.IntPtr = System.IntPtr.Zero
            Dim group As System.IntPtr = System.IntPtr.Zero
            Dim dacl As System.IntPtr = System.IntPtr.Zero
            Dim sacl As System.IntPtr = System.IntPtr.Zero
            Dim descriptor As System.IntPtr = System.IntPtr.Zero
            Try
                ' LABEL, ATTRIBUTE and SCOPE require READ_CONTROL, not audit-SACL
                ' privilege. Failure is unknown protection and selects private output.
                Dim status As System.UInt32 = GetSecurityInfo(handle, 1, &H70UI, owner, group, dacl, sacl, descriptor)
                If status <> 0UI OrElse descriptor = System.IntPtr.Zero Then Throw New System.NotSupportedException("Supplemental source security policy could not be verified (" & status.ToString(System.Globalization.CultureInfo.InvariantCulture) & ").")
                Dim length As System.UInt32 = GetSecurityDescriptorLength(descriptor)
                If length = 0UI OrElse length > 65536UI Then Throw New System.NotSupportedException("Unsupported supplemental security descriptor.")
                Dim bytes(CInt(length) - 1) As System.Byte
                System.Runtime.InteropServices.Marshal.Copy(descriptor, bytes, 0, bytes.Length)
                Dim raw As New System.Security.AccessControl.RawSecurityDescriptor(bytes, 0)
                If raw.SystemAcl IsNot Nothing Then
                    For Each ace As System.Security.AccessControl.GenericAce In raw.SystemAcl
                        If CInt(ace.AceType) <> &H11 Then Throw New System.NotSupportedException("Resource attributes or central access policy require private artifact storage.")
                        Dim aceBytes(ace.BinaryLength - 1) As System.Byte
                        ace.GetBinaryForm(aceBytes, 0)
                        If aceBytes.Length < 20 OrElse System.BitConverter.ToInt32(aceBytes, 4) <> 1 OrElse New System.Security.Principal.SecurityIdentifier(aceBytes, 8).Value <> "S-1-16-8192" Then Throw New System.NotSupportedException("Nondefault mandatory source security requires private artifact storage.")
                    Next
                End If
                Return bytes
            Finally
                If descriptor <> System.IntPtr.Zero Then LocalFree(descriptor)
            End Try
        End Function

        <System.Runtime.InteropServices.DllImport("advapi32.dll", SetLastError:=True)>
        Private Shared Function GetSecurityInfo(handle As Microsoft.Win32.SafeHandles.SafeFileHandle, objectType As System.Int32, information As System.UInt32, ByRef owner As System.IntPtr, ByRef group As System.IntPtr, ByRef dacl As System.IntPtr, ByRef sacl As System.IntPtr, ByRef descriptor As System.IntPtr) As System.UInt32
        End Function

        <System.Runtime.InteropServices.DllImport("advapi32.dll", SetLastError:=True)>
        Private Shared Function GetSecurityDescriptorLength(descriptor As System.IntPtr) As System.UInt32
        End Function

        <System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True)>
        Private Shared Function LocalFree(memory As System.IntPtr) As System.IntPtr
        End Function

        Private Shared Sub CreateSharedDirectory(path As System.String, snapshot As PermissionSnapshot, producer As System.Security.Principal.SecurityIdentifier)
            SemanticArchivePathGuard.RequireWindowsCompatiblePath(path)
            If EntryExists(path) Then
                If (System.IO.File.GetAttributes(path) And System.IO.FileAttributes.Directory) = 0 Then Throw New System.IO.IOException("The artifact directory is occupied by a file.")
                RequireProjectedSecurity(path, snapshot)
                RequireWritableArtifactOwner(path, snapshot)
            Else
                System.IO.Directory.CreateDirectory(path, DirectCast(ProjectSecurity(snapshot, producer, True), System.Security.AccessControl.DirectorySecurity))
                RequireProjectedSecurity(path, snapshot)
            End If
        End Sub
    End Class
End Namespace
