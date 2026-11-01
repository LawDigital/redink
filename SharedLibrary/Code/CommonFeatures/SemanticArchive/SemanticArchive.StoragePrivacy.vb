' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' First-deployment access domain: private generated plaintext for one Windows user.

' =============================================================================
' File: SemanticArchive.StoragePrivacy.vb
' Purpose:
'   Private catalog/artifact location validation, creation and Windows access
'   protection.
'
' Architecture / Function:
'   Requires verified private storage rather than assuming a directory name or user-
'   local path implies safe permissions.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary
    Public NotInheritable Partial Class SemanticArchiveStore

        ''' <summary>Returns the suggested personal catalog path without creating or changing storage.</summary>
        Public Shared Function GetSuggestedPrivateCatalogDirectory() As System.String
            Dim local As System.String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData)
            If System.String.IsNullOrWhiteSpace(local) Then Throw New System.InvalidOperationException("A personal Semantic Archive catalog requires an available LocalAppData directory.")
            Return SemanticArchivePathGuard.RequireWindowsCompatiblePath(System.IO.Path.Combine(local, "RedInk", "SA"))
        End Function

        ''' <summary>Read-only path and ACL validation; does not create storage, register outputs, or change ACLs.</summary>
        Public Shared Sub ValidatePrivateCatalogLocation(path As System.String)
            Dim full As System.String = SemanticArchivePathGuard.RequireWindowsCompatiblePath(path)
            Dim catalog As System.String = System.IO.Path.Combine(full, CatalogFileName)
            RequireAtomicWritePathBudget(catalog)
            RequireAtomicWritePathBudget(System.IO.Path.Combine(full, DataDirectoryName, "catalog.lock"))
            SemanticArchivePathGuard.ValidateContainedPath(full, full, False)
            RequireStableAncestors(full, GetStorageOwner())
            If ExistsChecked(full) Then
                If (System.IO.File.GetAttributes(full) And System.IO.FileAttributes.Directory) = 0 Then Throw New System.IO.IOException("The personal catalog directory is occupied by a file: '" & full & "'.")
                ' Existing catalog containers may allow listing/read access; the
                ' catalog and generated children must still be private individually.
                For Each existing As System.String In New System.String() {catalog, System.IO.Path.Combine(full, DataDirectoryName)}
                    If ExistsChecked(existing) Then RequirePrivateArtifact(existing)
                Next
            End If
        End Sub

        ''' <summary>
        ''' Creates only generated directories. Existing source/output directory ACLs
        ''' are never rewritten. A wider existing generated directory requires an
        ''' explicit access-domain decision outside this first deployment policy.
        ''' </summary>
        Public Shared Sub CreatePrivateDirectory(path As System.String)
            Dim full As System.String = SemanticArchivePathGuard.RequireWindowsCompatiblePath(path)
            SemanticArchivePathGuard.ValidateContainedPath(full, full, False)
            Dim owner As System.Security.Principal.SecurityIdentifier = GetStorageOwner()
            RequireStableAncestors(System.IO.Path.GetDirectoryName(full), owner)
            If ExistsChecked(full) Then
                If (System.IO.File.GetAttributes(full) And System.IO.FileAttributes.Directory) = 0 Then Throw New System.IO.IOException("A generated directory path is occupied by a file.")
                RequirePrivateArtifact(full)
                Return
            End If
            Dim security As New System.Security.AccessControl.DirectorySecurity()
            security.SetAccessRuleProtection(True, False)
            security.SetOwner(owner)
            Dim identities As System.Collections.Generic.IEnumerable(Of System.Security.Principal.SecurityIdentifier) = StorageIdentities(owner)
            For Each identity As System.Security.Principal.SecurityIdentifier In identities
                security.AddAccessRule(New System.Security.AccessControl.FileSystemAccessRule(identity,
                    System.Security.AccessControl.FileSystemRights.FullControl,
                    System.Security.AccessControl.InheritanceFlags.ContainerInherit Or System.Security.AccessControl.InheritanceFlags.ObjectInherit,
                    System.Security.AccessControl.PropagationFlags.None,
                    System.Security.AccessControl.AccessControlType.Allow))
            Next
            System.IO.Directory.CreateDirectory(full, security)
            RequirePrivateArtifact(full)
        End Sub

        Public Shared Sub RequirePrivateArtifact(path As System.String)
            Dim full As System.String = SemanticArchivePathGuard.CanonicalPath(path)
            SemanticArchivePathGuard.ValidateContainedPath(System.IO.Path.GetDirectoryName(full), full, True)
            Dim owner As System.Security.Principal.SecurityIdentifier = GetStorageOwner()
            Dim allowed As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            For Each identity As System.Security.Principal.SecurityIdentifier In StorageIdentities(owner)
                allowed.Add(identity.Value)
            Next
            Dim security As System.Security.AccessControl.FileSystemSecurity
            If (System.IO.File.GetAttributes(full) And System.IO.FileAttributes.Directory) <> 0 Then
                security = System.IO.Directory.GetAccessControl(full, System.Security.AccessControl.AccessControlSections.Access Or System.Security.AccessControl.AccessControlSections.Owner)
            Else
                security = System.IO.File.GetAccessControl(full, System.Security.AccessControl.AccessControlSections.Access Or System.Security.AccessControl.AccessControlSections.Owner)
            End If
            Dim actualOwner As System.Security.Principal.SecurityIdentifier = TryCast(security.GetOwner(GetType(System.Security.Principal.SecurityIdentifier)), System.Security.Principal.SecurityIdentifier)
            If actualOwner Is Nothing OrElse Not allowed.Contains(actualOwner.Value) Then Throw PrivateStorageFailure(full, "The generated artifact owner is outside the current-user/SYSTEM/Administrators domain: " & If(actualOwner Is Nothing, "unresolved owner", actualOwner.Value) & ".")
            Dim userCanRead As System.Boolean = False
            Dim readOrMutate As System.Security.AccessControl.FileSystemRights =
                System.Security.AccessControl.FileSystemRights.ReadData Or
                System.Security.AccessControl.FileSystemRights.WriteData Or
                System.Security.AccessControl.FileSystemRights.AppendData Or
                System.Security.AccessControl.FileSystemRights.ChangePermissions Or
                System.Security.AccessControl.FileSystemRights.TakeOwnership Or
                System.Security.AccessControl.FileSystemRights.Delete Or
                System.Security.AccessControl.FileSystemRights.DeleteSubdirectoriesAndFiles
            Dim rules As System.Security.AccessControl.AuthorizationRuleCollection = security.GetAccessRules(True, True, GetType(System.Security.Principal.SecurityIdentifier))
            For Each item As System.Security.AccessControl.AuthorizationRule In rules
                Dim rule As System.Security.AccessControl.FileSystemAccessRule = TryCast(item, System.Security.AccessControl.FileSystemAccessRule)
                If rule Is Nothing OrElse rule.AccessControlType <> System.Security.AccessControl.AccessControlType.Allow OrElse (rule.PropagationFlags And System.Security.AccessControl.PropagationFlags.InheritOnly) <> 0 Then Continue For
                Dim identity As System.Security.Principal.SecurityIdentifier = TryCast(rule.IdentityReference, System.Security.Principal.SecurityIdentifier)
                If identity Is Nothing Then Throw PrivateStorageFailure(full, "A generated artifact access-rule SID could not be resolved.")
                If Not allowed.Contains(identity.Value) AndAlso (rule.FileSystemRights And readOrMutate) <> 0 Then Throw PrivateStorageFailure(full, "Existing generated storage grants principal " & identity.Value & " broader access: " & rule.FileSystemRights.ToString() & ".")
                If identity.Value = owner.Value AndAlso (rule.FileSystemRights And System.Security.AccessControl.FileSystemRights.ReadData) <> 0 Then userCanRead = True
            Next
            If Not userCanRead Then Throw PrivateStorageFailure(full, "Current user " & owner.Value & " does not have a verifiable explicit read grant.")
        End Sub

        Public Shared Function CreatePrivateFile(path As System.String) As System.IO.FileStream
            Dim full As System.String = SemanticArchivePathGuard.RequireWindowsCompatiblePath(path)
            Dim parent As System.String = System.IO.Path.GetDirectoryName(full)
            SemanticArchivePathGuard.ValidateContainedPath(parent, full, False)
            Dim owner As System.Security.Principal.SecurityIdentifier = GetStorageOwner()
            RequireStableAncestors(parent, owner)
            RequirePrivateArtifact(parent)
            Dim security As New System.Security.AccessControl.FileSecurity()
            security.SetAccessRuleProtection(True, False)
            security.SetOwner(owner)
            For Each identity As System.Security.Principal.SecurityIdentifier In StorageIdentities(owner)
                security.AddAccessRule(New System.Security.AccessControl.FileSystemAccessRule(identity, System.Security.AccessControl.FileSystemRights.FullControl, System.Security.AccessControl.AccessControlType.Allow))
            Next
            Return New System.IO.FileStream(full, System.IO.FileMode.CreateNew,
                System.Security.AccessControl.FileSystemRights.Write, System.IO.FileShare.None,
                65536, System.IO.FileOptions.WriteThrough, security)
        End Function

        Private Shared Sub RequireStableAncestors(path As System.String, owner As System.Security.Principal.SecurityIdentifier)
            If System.String.IsNullOrWhiteSpace(path) Then Return
            Dim current As New System.IO.DirectoryInfo(path)
            While current IsNot Nothing
                If ExistsChecked(current.FullName) Then
                    Dim security As System.Security.AccessControl.DirectorySecurity = current.GetAccessControl(System.Security.AccessControl.AccessControlSections.Access Or System.Security.AccessControl.AccessControlSections.Owner)
                    RequireStableAncestorSecurity(current.FullName, security, owner, IsVerifiedWindowsVolumeRoot(current.FullName), IsVerifiedLocalVolumeRoot(current.FullName))
                End If
                current = current.Parent
            End While
        End Sub

        Private Shared Sub RequireStableAncestorSecurity(path As System.String, security As System.Security.AccessControl.DirectorySecurity,
                                                          owner As System.Security.Principal.SecurityIdentifier, verifiedWindowsVolumeRoot As System.Boolean, verifiedLocalVolumeRoot As System.Boolean)
            Dim allowed As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            For Each identity As System.Security.Principal.SecurityIdentifier In StorageIdentities(owner)
                allowed.Add(identity.Value)
            Next
            ' Windows Resource Protection uses this exact service SID. Its trust is
            ' limited to the verified local Windows volume root, never generated ACLs.
            If verifiedWindowsVolumeRoot Then allowed.Add("S-1-5-80-956008885-3418522649-1831038044-1853292631-2271478464")
            Dim actualOwner As System.Security.Principal.SecurityIdentifier = TryCast(security.GetOwner(GetType(System.Security.Principal.SecurityIdentifier)), System.Security.Principal.SecurityIdentifier)
            If actualOwner Is Nothing OrElse Not allowed.Contains(actualOwner.Value) Then Throw PrivateStorageFailure(path, "Ancestor owner " & If(actualOwner Is Nothing, "unresolved", actualOwner.Value) & " is outside the trusted storage domain; owner control of the DACL cannot be established as safe.")
            Dim raw As New System.Security.AccessControl.RawSecurityDescriptor(security.GetSecurityDescriptorBinaryForm(), 0)
            If raw.DiscretionaryAcl Is Nothing Then Throw PrivateStorageFailure(path, "The ancestor has a NULL DACL, which permits access to every principal.")
            Dim canReplace As System.Security.AccessControl.FileSystemRights = System.Security.AccessControl.FileSystemRights.DeleteSubdirectoriesAndFiles Or System.Security.AccessControl.FileSystemRights.ChangePermissions Or System.Security.AccessControl.FileSystemRights.TakeOwnership
            ' DELETE applies to the directory itself. A verified physical volume
            ' root cannot be replaced this way; deleting its children and changing
            ' its security remain dangerous. Every ordinary ancestor keeps DELETE.
            If Not verifiedLocalVolumeRoot Then canReplace = canReplace Or System.Security.AccessControl.FileSystemRights.Delete
            For Each item As System.Security.AccessControl.AuthorizationRule In security.GetAccessRules(True, True, GetType(System.Security.Principal.SecurityIdentifier))
                Dim rule As System.Security.AccessControl.FileSystemAccessRule = TryCast(item, System.Security.AccessControl.FileSystemAccessRule)
                If rule Is Nothing OrElse rule.AccessControlType <> System.Security.AccessControl.AccessControlType.Allow OrElse (rule.PropagationFlags And System.Security.AccessControl.PropagationFlags.InheritOnly) <> 0 Then Continue For
                Dim identity As System.Security.Principal.SecurityIdentifier = TryCast(rule.IdentityReference, System.Security.Principal.SecurityIdentifier)
                If identity Is Nothing Then Throw PrivateStorageFailure(path, "An ancestor access-rule SID could not be resolved.")
                Dim dangerousRights As System.Security.AccessControl.FileSystemRights = rule.FileSystemRights And canReplace
                If Not allowed.Contains(identity.Value) AndAlso dangerousRights <> 0 Then Throw PrivateStorageFailure(path, "Ancestor access policy contains a replacement-capable Allow entry for principal " & identity.Value & "; applicable Allow entry rights: " & rule.FileSystemRights.ToString() & "; blocking rights: " & dangerousRights.ToString() & "; scope: " & If(verifiedLocalVolumeRoot, "verified local volume root", "directory or unverified volume root") & ".")
            Next
        End Sub

        Friend Shared Function IsVerifiedLocalVolumeRoot(path As System.String) As System.Boolean
            Dim full As System.String = SemanticArchivePathGuard.CanonicalPath(path)
            Dim lexicalRoot As System.String = System.IO.Path.GetPathRoot(full)
            If full.StartsWith("\\", System.StringComparison.Ordinal) OrElse Not System.String.Equals(full, lexicalRoot, System.StringComparison.OrdinalIgnoreCase) Then Return False
            Dim drive As New System.IO.DriveInfo(full)
            Dim driveType As System.IO.DriveType = drive.DriveType
            If driveType <> System.IO.DriveType.Fixed AndAlso driveType <> System.IO.DriveType.Removable AndAlso driveType <> System.IO.DriveType.Ram Then Return False
            ' Drive letters alone are not proof: mapped shares and SUBST aliases
            ' must not turn an ordinary directory into a trusted physical root.
            Dim physical As System.String = SemanticArchivePathGuard.GetVerifiedDirectoryIdentity(full)
            If physical.StartsWith("\\", System.StringComparison.Ordinal) OrElse Not System.String.Equals(physical, System.IO.Path.GetPathRoot(physical), System.StringComparison.OrdinalIgnoreCase) Then Return False
            Return System.String.Equals(full, physical, System.StringComparison.OrdinalIgnoreCase)
        End Function

        Friend Shared Function IsVerifiedWindowsVolumeRoot(path As System.String) As System.Boolean
            Dim full As System.String = SemanticArchivePathGuard.CanonicalPath(path)
            Dim lexicalRoot As System.String = System.IO.Path.GetPathRoot(full)
            If full.StartsWith("\\", System.StringComparison.Ordinal) OrElse Not System.String.Equals(full, lexicalRoot, System.StringComparison.OrdinalIgnoreCase) Then Return False
            Dim windows As System.String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Windows)
            If System.String.IsNullOrWhiteSpace(windows) Then Return False
            Dim physicalWindows As System.String = SemanticArchivePathGuard.GetVerifiedDirectoryIdentity(windows)
            If physicalWindows.StartsWith("\\", System.StringComparison.Ordinal) Then Return False
            Dim physical As System.String = SemanticArchivePathGuard.GetVerifiedDirectoryIdentity(full)
            If physical.StartsWith("\\", System.StringComparison.Ordinal) OrElse Not System.String.Equals(physical, System.IO.Path.GetPathRoot(physical), System.StringComparison.OrdinalIgnoreCase) Then Return False
            Return System.String.Equals(physical, System.IO.Path.GetPathRoot(physicalWindows), System.StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function PrivateStorageFailure(path As System.String, detail As System.String) As System.UnauthorizedAccessException
            Return New System.UnauthorizedAccessException("storage_access_domain_required: " & detail & " Path: '" & path & "'. Select a private personal catalog folder (for example %LOCALAPPDATA%\RedInk\SA). Existing files and ACLs were left unchanged.")
        End Function

        Private Shared Function GetStorageOwner() As System.Security.Principal.SecurityIdentifier
            If System.Environment.OSVersion.Platform <> System.PlatformID.Win32NT Then Throw New System.PlatformNotSupportedException("storage_access_domain_required: Private archive storage requires Windows ACL validation in this deployment.")
            Using identity As System.Security.Principal.WindowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent()
                If identity.User Is Nothing Then Throw New System.UnauthorizedAccessException("storage_access_domain_required: The current Windows storage identity is unavailable.")
                Return New System.Security.Principal.SecurityIdentifier(identity.User.Value)
            End Using
        End Function

        Private Shared Iterator Function StorageIdentities(owner As System.Security.Principal.SecurityIdentifier) As System.Collections.Generic.IEnumerable(Of System.Security.Principal.SecurityIdentifier)
            Yield owner
            Yield New System.Security.Principal.SecurityIdentifier(System.Security.Principal.WellKnownSidType.LocalSystemSid, Nothing)
            Yield New System.Security.Principal.SecurityIdentifier(System.Security.Principal.WellKnownSidType.BuiltinAdministratorsSid, Nothing)
        End Function
    End Class
End Namespace
