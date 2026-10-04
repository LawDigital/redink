' Part of "Red Ink" (SharedLibrary)
' ACL-protected descriptor publication. No source ACL is altered and no indexed text is published here.
Option Strict On
Option Explicit On

Namespace SharedLibrary
    Public NotInheritable Partial Class SemanticArchiveLibrary
        Private Shared Function CurrentPublisherSid() As System.String
            Using identity As System.Security.Principal.WindowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent()
                If identity.User Is Nothing Then Throw New System.UnauthorizedAccessException("The publisher's Windows identity is unavailable.")
                Return identity.User.Value
            End Using
        End Function

        Private Shared Sub RequireLibraryDirectory(directory As System.String)
            SemanticArchivePathGuard.GetVerifiedDirectoryIdentity(directory)
            Dim attributes As System.IO.FileAttributes = System.IO.File.GetAttributes(directory)
            If (attributes And System.IO.FileAttributes.Directory) = 0 OrElse (attributes And System.IO.FileAttributes.ReparsePoint) <> 0 Then Throw New System.UnauthorizedAccessException("The library must be an existing ordinary directory, provisioned by an administrator.")
        End Sub

        Private Shared Function LibraryOwner(directory As System.String) As System.String
            Dim security As System.Security.AccessControl.DirectorySecurity = System.IO.Directory.GetAccessControl(directory,
                System.Security.AccessControl.AccessControlSections.Owner Or System.Security.AccessControl.AccessControlSections.Access)
            Dim owner As System.Security.Principal.SecurityIdentifier = TryCast(security.GetOwner(GetType(System.Security.Principal.SecurityIdentifier)), System.Security.Principal.SecurityIdentifier)
            If owner Is Nothing Then Throw New System.UnauthorizedAccessException("The configured library's owner cannot be verified.")
            Return owner.Value
        End Function

        Private Shared Function TrustedWriters(publisher As System.String, directory As System.String) As System.Collections.Generic.HashSet(Of System.String)
            ' The configured library directory is the administrative trust anchor.
            ' Other publishers may create new files but must not replace existing entries.
            Return New System.Collections.Generic.HashSet(Of System.String)(New System.String() {publisher, LibraryOwner(directory), "S-1-5-18", "S-1-5-32-544"}, System.StringComparer.Ordinal)
        End Function

        Private Shared Sub RequireStableLibraryAncestors(directory As System.String, publisher As System.String)
            Dim permitted As System.Collections.Generic.HashSet(Of System.String) = TrustedWriters(publisher, directory)
            Dim current As New System.IO.DirectoryInfo(directory)
            While current IsNot Nothing
                Dim security As System.Security.AccessControl.DirectorySecurity = current.GetAccessControl(
                    System.Security.AccessControl.AccessControlSections.Access Or System.Security.AccessControl.AccessControlSections.Owner)
                SemanticArchiveArtifactPlanner.RequireStableSharedAncestorSecurity(current.FullName, security, permitted,
                    SemanticArchiveStore.IsVerifiedLocalVolumeRoot(current.FullName), SemanticArchiveStore.IsVerifiedWindowsVolumeRoot(current.FullName))
                current = current.Parent
            End While
        End Sub

        Private Shared Sub ValidateDescriptorSecurity(security As System.Security.AccessControl.FileSecurity, publisher As System.String, directory As System.String)
            Dim raw As New System.Security.AccessControl.RawSecurityDescriptor(security.GetSecurityDescriptorBinaryForm(), 0)
            If raw.Owner Is Nothing OrElse raw.Owner.Value <> publisher OrElse raw.DiscretionaryAcl Is Nothing Then Throw New System.UnauthorizedAccessException("library_descriptor_untrusted: The descriptor owner does not match its publisher or has no explicit ACL.")
            Dim trusted As System.Collections.Generic.HashSet(Of System.String) = TrustedWriters(publisher, directory)
            Dim mutation As System.Security.AccessControl.FileSystemRights = System.Security.AccessControl.FileSystemRights.Write Or
                System.Security.AccessControl.FileSystemRights.Delete Or System.Security.AccessControl.FileSystemRights.ChangePermissions Or System.Security.AccessControl.FileSystemRights.TakeOwnership
            Dim mutationMask As System.Int32 = CInt(mutation) Or &H40000000 Or &H10000000 ' GENERIC_WRITE / GENERIC_ALL
            For Each ace As System.Security.AccessControl.GenericAce In raw.DiscretionaryAcl
                If (ace.AceFlags And System.Security.AccessControl.AceFlags.InheritOnly) <> 0 Then Continue For
                Dim common As System.Security.AccessControl.CommonAce = TryCast(ace, System.Security.AccessControl.CommonAce)
                If common Is Nothing OrElse common.IsCallback OrElse
                    (common.AceQualifier <> System.Security.AccessControl.AceQualifier.AccessAllowed AndAlso common.AceQualifier <> System.Security.AccessControl.AceQualifier.AccessDenied) Then
                    Throw New System.UnauthorizedAccessException("library_acl_unsupported: A descriptor has an unsupported access rule.")
                End If
                If common.AceQualifier = System.Security.AccessControl.AceQualifier.AccessAllowed AndAlso
                    (common.AccessMask And mutationMask) <> 0 AndAlso Not trusted.Contains(common.SecurityIdentifier.Value) Then
                    Throw New System.UnauthorizedAccessException("library_descriptor_writable: Another principal can change this publication. Restrict its write/replace permissions before use.")
                End If
            Next
        End Sub

        Private Shared Function InitialDescriptorSecurity(directory As System.String, publisher As System.String) As System.Security.AccessControl.FileSecurity
            Dim parent As System.Security.AccessControl.DirectorySecurity = System.IO.Directory.GetAccessControl(directory,
                System.Security.AccessControl.AccessControlSections.Access Or System.Security.AccessControl.AccessControlSections.Owner)
            Dim raw As New System.Security.AccessControl.RawSecurityDescriptor(parent.GetSecurityDescriptorBinaryForm(), 0)
            If raw.DiscretionaryAcl Is Nothing Then Throw New System.UnauthorizedAccessException("The library needs an explicit inherited reader policy.")
            For Each ace As System.Security.AccessControl.GenericAce In raw.DiscretionaryAcl
                Dim common As System.Security.AccessControl.CommonAce = TryCast(ace, System.Security.AccessControl.CommonAce)
                If common Is Nothing OrElse common.IsCallback Then Throw New System.UnauthorizedAccessException("The library's inherited reader policy contains an unsupported access rule.")
            Next
            Dim security As New System.Security.AccessControl.FileSecurity()
            security.SetAccessRuleProtection(True, False)
            security.SetOwner(New System.Security.Principal.SecurityIdentifier(publisher))
            For Each item As System.Security.AccessControl.AuthorizationRule In parent.GetAccessRules(True, True, GetType(System.Security.Principal.SecurityIdentifier))
                Dim rule As System.Security.AccessControl.FileSystemAccessRule = DirectCast(item, System.Security.AccessControl.FileSystemAccessRule)
                If (rule.InheritanceFlags And System.Security.AccessControl.InheritanceFlags.ObjectInherit) = 0 Then Continue For
                Dim sid As System.Security.Principal.SecurityIdentifier = DirectCast(rule.IdentityReference, System.Security.Principal.SecurityIdentifier)
                If sid.Value = "S-1-3-0" OrElse sid.Value = "S-1-3-1" Then Continue For
                Dim rights As System.Security.AccessControl.FileSystemRights = rule.FileSystemRights And System.Security.AccessControl.FileSystemRights.Read
                Dim rawRights As System.Int32 = CInt(rule.FileSystemRights)
                If (rawRights And (&H80000000 Or &H10000000)) <> 0 Then rights = rights Or System.Security.AccessControl.FileSystemRights.Read
                If (rawRights And &H20000000) <> 0 Then rights = rights Or System.Security.AccessControl.FileSystemRights.ReadAttributes Or System.Security.AccessControl.FileSystemRights.ReadPermissions
                If (rawRights And &H40000000) <> 0 Then rights = rights Or System.Security.AccessControl.FileSystemRights.ReadPermissions
                If rights = 0 Then Continue For
                security.AddAccessRule(New System.Security.AccessControl.FileSystemAccessRule(sid, rights, rule.AccessControlType))
            Next
            For Each sid As System.String In New System.String() {publisher, "S-1-5-18", "S-1-5-32-544"}
                security.AddAccessRule(New System.Security.AccessControl.FileSystemAccessRule(New System.Security.Principal.SecurityIdentifier(sid),
                    System.Security.AccessControl.FileSystemRights.FullControl, System.Security.AccessControl.AccessControlType.Allow))
            Next
            Return security
        End Function

        Private Shared Function ReadEntry(directory As System.String, path As System.String) As SemanticArchiveLibraryEntry
            RequireLibraryDirectory(directory)
            Using stream As System.IO.FileStream = SemanticArchivePathGuard.OpenContainedRead(directory, path)
                If stream.Length < 2 OrElse stream.Length > SharedMethods.DEFAULT_SEMANTICARCHIVE_LIBRARY_MAX_DESCRIPTOR_BYTES Then Throw New System.IO.InvalidDataException("The library descriptor exceeds its size budget.")
                Using reader As New System.IO.StreamReader(stream, New System.Text.UTF8Encoding(False, True), True, 4096, True)
                    Using jsonReader As New Newtonsoft.Json.JsonTextReader(reader) With {.MaxDepth = 40, .DateParseHandling = Newtonsoft.Json.DateParseHandling.None}
                        Dim json As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Load(jsonReader,
                            New Newtonsoft.Json.Linq.JsonLoadSettings With {.DuplicatePropertyNameHandling = Newtonsoft.Json.Linq.DuplicatePropertyNameHandling.Error})
                        If jsonReader.Read() Then Throw New System.IO.InvalidDataException("Unexpected trailing library descriptor data.")
                        Dim names As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
                        For Each propertyValue As Newtonsoft.Json.Linq.JProperty In json.Properties()
                            If Not names.Add(propertyValue.Name) Then Throw New System.IO.InvalidDataException("Ambiguous library descriptor member.")
                        Next
                        Dim serializer As Newtonsoft.Json.JsonSerializer = Newtonsoft.Json.JsonSerializer.Create(New Newtonsoft.Json.JsonSerializerSettings With {
                            .TypeNameHandling = Newtonsoft.Json.TypeNameHandling.None, .MaxDepth = 40, .MissingMemberHandling = Newtonsoft.Json.MissingMemberHandling.Error})
                        Dim entry As SemanticArchiveLibraryEntry = json.ToObject(Of SemanticArchiveLibraryEntry)(serializer)
                        If entry Is Nothing OrElse entry.SchemaVersion <> 1 OrElse entry.Revision < 1 OrElse entry.Definition Is Nothing OrElse entry.Definition.Library IsNot Nothing Then Throw New System.IO.InvalidDataException("Invalid library descriptor.")
                        If Not System.String.Equals(EntryPath(directory, entry.EntryId), path, System.StringComparison.OrdinalIgnoreCase) OrElse entry.Definition.ArchiveId <> entry.EntryId Then Throw New System.IO.InvalidDataException("Library filename and archive identity disagree.")
                        ValidateDescriptorSecurity(stream.GetAccessControl(), entry.PublisherSid, directory)
                        RequireStableLibraryAncestors(directory, entry.PublisherSid)
                        SemanticArchiveStore.ValidateLibraryDefinition(entry.Definition)
                        If entry.Definition.Roots.Count = 0 Then Throw New System.IO.InvalidDataException("A published archive must contain a source folder.")
                        For Each binding As SemanticArchiveSourceBinding In entry.Definition.Roots
                            If Not System.String.IsNullOrEmpty(binding.ShadowArtifactRoot) Then Throw New System.IO.InvalidDataException("A library descriptor cannot prescribe a user's private derivative folder.")
                            If directory.StartsWith("\\", System.StringComparison.Ordinal) AndAlso
                                (Not binding.RootPath.StartsWith("\\", System.StringComparison.Ordinal) OrElse
                                 (binding.SharedArtifactRoot.Length > 0 AndAlso Not binding.SharedArtifactRoot.StartsWith("\\", System.StringComparison.Ordinal))) Then
                                Throw New System.IO.InvalidDataException("Network library sources and shared outputs must be portable UNC paths.")
                            End If
                        Next
                        Return entry
                    End Using
                End Using
            End Using
        End Function

        Private Shared Sub WriteEntry(directory As System.String, path As System.String, entry As SemanticArchiveLibraryEntry,
                    replace As System.Boolean, initialSecurity As System.Security.AccessControl.FileSecurity)
            Dim security As System.Security.AccessControl.FileSecurity = If(replace,
                System.IO.File.GetAccessControl(path, System.Security.AccessControl.AccessControlSections.Access Or System.Security.AccessControl.AccessControlSections.Owner), initialSecurity)
            ValidateDescriptorSecurity(security, entry.PublisherSid, directory)
            RequireStableLibraryAncestors(directory, entry.PublisherSid)
            Dim bytes As System.Byte() = New System.Text.UTF8Encoding(False, True).GetBytes(Newtonsoft.Json.JsonConvert.SerializeObject(entry, Newtonsoft.Json.Formatting.Indented))
            If bytes.Length > SharedMethods.DEFAULT_SEMANTICARCHIVE_LIBRARY_MAX_DESCRIPTOR_BYTES Then Throw New System.IO.InvalidDataException("The library descriptor is too large.")
            Dim temporary As System.String = SemanticArchivePathGuard.RequireWindowsCompatiblePath(System.IO.Path.Combine(directory, "." & System.Guid.NewGuid().ToString("N") & ".tmp"))
            Try
                ' Security is supplied at creation, not after confidential metadata was written.
                Using stream As New System.IO.FileStream(temporary, System.IO.FileMode.CreateNew,
                        System.Security.AccessControl.FileSystemRights.Read Or System.Security.AccessControl.FileSystemRights.Write,
                        System.IO.FileShare.None, 4096, System.IO.FileOptions.WriteThrough, security)
                    ValidateDescriptorSecurity(stream.GetAccessControl(), entry.PublisherSid, directory)
                    stream.Write(bytes, 0, bytes.Length)
                    stream.Flush(True)
                End Using
                RequireStableLibraryAncestors(directory, entry.PublisherSid)
                If replace Then
                    System.IO.File.Replace(temporary, path, Nothing, False)
                Else
                    System.IO.File.Move(temporary, path)
                End If
                Dim verified As SemanticArchiveLibraryEntry = ReadEntry(directory, path)
                If verified.Revision <> entry.Revision OrElse verified.Withdrawn <> entry.Withdrawn OrElse DefinitionHash(verified.Definition) <> DefinitionHash(entry.Definition) Then Throw New System.IO.IOException("Library publication verification failed.")
            Finally
                ' Only our random temporary file can be removed. The published definition is
                ' never deleted as a fallback when atomic replacement or its ACL check fails.
                If System.IO.File.Exists(temporary) Then System.IO.File.Delete(temporary)
            End Try
        End Sub
    End Class
End Namespace
