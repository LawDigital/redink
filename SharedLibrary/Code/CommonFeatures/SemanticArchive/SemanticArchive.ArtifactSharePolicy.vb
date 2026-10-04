' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
Option Strict On
Option Explicit On

Namespace SharedLibrary
    Public NotInheritable Partial Class SemanticArchiveArtifactPlanner
        Private NotInheritable Class ShareProtection
            Public Descriptor As System.String
            Public Flags As System.UInt32
        End Class

        ''' <summary>
        ''' Same-share access retains its existing policy and requires no share RPC.
        ''' Another share is accepted only on the identical verified UNC server and
        ''' with exactly equivalent supported share policy. Equal file ACLs alone
        ''' never authorize transfer to another server or between UNC and local IO.
        ''' </summary>
        Private Shared Function GetProtectionDomainSignature(sourceIdentity As System.String, destinationParent As System.String) As System.String
            Dim target As System.String = SemanticArchivePathGuard.GetVerifiedDirectoryIdentity(destinationParent)
            Dim sourceDomain As System.String = ShareDomain(sourceIdentity)
            Dim targetDomain As System.String = ShareDomain(target)
            If System.String.Equals(sourceDomain, targetDomain, System.StringComparison.OrdinalIgnoreCase) Then Return ""
            If sourceDomain = "local" OrElse targetDomain = "local" Then Throw New System.NotSupportedException("shared_location_unverified: A transfer between a UNC share and local storage has a different administrator and group authority. Private output will be used.")
            Dim sourceParts As System.String() = sourceDomain.Substring(2).Split("\"c)
            Dim targetParts As System.String() = targetDomain.Substring(2).Split("\"c)
            If Not System.String.Equals(sourceParts(0), targetParts(0), System.StringComparison.OrdinalIgnoreCase) Then Throw New System.NotSupportedException("shared_location_unverified: Different UNC servers cannot be authorized from matching file ACLs. Private output will be used.")
            Dim sourcePolicy As ShareProtection = ReadShareProtection(sourceParts(0), sourceParts(1))
            Dim targetPolicy As ShareProtection = ReadShareProtection(targetParts(0), targetParts(1))
            Return CompareShareProtection(sourceDomain, targetDomain, sourcePolicy, targetPolicy)
        End Function

        Private Shared Function CompareShareProtection(sourceDomain As System.String, targetDomain As System.String, sourcePolicy As ShareProtection, targetPolicy As ShareProtection) As System.String
            If sourcePolicy.Flags <> targetPolicy.Flags OrElse Not System.String.Equals(sourcePolicy.Descriptor, targetPolicy.Descriptor, System.StringComparison.Ordinal) Then Throw New System.NotSupportedException("shared_location_unverified: The source and target share security descriptors or share flags differ. Private output will be used.")
            Return SemanticArchiveIdentity.HashBytes(System.Text.Encoding.UTF8.GetBytes("same-server-share-policy-v1|" & sourceDomain.ToUpperInvariant() & "|" & targetDomain.ToUpperInvariant() & "|" & sourcePolicy.Descriptor & "|" & sourcePolicy.Flags.ToString(System.Globalization.CultureInfo.InvariantCulture)))
        End Function

        Private Shared Function ReadShareProtection(server As System.String, share As System.String) As ShareProtection
            Dim information As System.IntPtr = System.IntPtr.Zero
            Dim flagsInformation As System.IntPtr = System.IntPtr.Zero
            Try
                ' Level 503 also identifies scoped shares. A scoped/cluster/DFS
                ' authority is not inferred from the apparent UNC server spelling.
                Dim status As System.UInt32 = NetShareGetInfo("\\" & server, share, 503UI, information)
                If status <> 0UI OrElse information = System.IntPtr.Zero Then Throw New System.NotSupportedException("shared_policy_unavailable: Share security could not be queried (Windows status " & status.ToString(System.Globalization.CultureInfo.InvariantCulture) & "). Private output will be used.")
                Dim entry As NativeShareInfo503 = DirectCast(System.Runtime.InteropServices.Marshal.PtrToStructure(information, GetType(NativeShareInfo503)), NativeShareInfo503)
                Dim reportedName As System.String = System.Runtime.InteropServices.Marshal.PtrToStringUni(entry.NetName)
                Dim scopedServer As System.String = System.Runtime.InteropServices.Marshal.PtrToStringUni(entry.ServerName)
                If entry.ShareType <> 0UI OrElse entry.Reserved <> 0UI OrElse Not System.String.Equals(reportedName, share, System.StringComparison.OrdinalIgnoreCase) OrElse scopedServer <> "*" Then Throw New System.NotSupportedException("shared_policy_unsupported: Only ordinary, unscoped disk shares can be compared automatically.")
                If entry.SecurityDescriptor = System.IntPtr.Zero OrElse Not IsValidSecurityDescriptor(entry.SecurityDescriptor) Then Throw New System.NotSupportedException("shared_policy_unsupported: The share has no valid explicit security descriptor.")
                Dim descriptorLength As System.UInt32 = GetSecurityDescriptorLength(entry.SecurityDescriptor)
                If descriptorLength = 0UI OrElse descriptorLength > 65536UI Then Throw New System.NotSupportedException("shared_policy_unsupported: The share security descriptor exceeds its validation bound.")
                Dim descriptor(CInt(descriptorLength) - 1) As System.Byte
                System.Runtime.InteropServices.Marshal.Copy(entry.SecurityDescriptor, descriptor, 0, descriptor.Length)
                Dim normalized As System.String = ValidateShareDescriptor(descriptor)
                status = NetShareGetInfo("\\" & server, share, 1005UI, flagsInformation)
                If status <> 0UI OrElse flagsInformation = System.IntPtr.Zero Then Throw New System.NotSupportedException("shared_policy_unavailable: Share transport/caching policy could not be queried (Windows status " & status.ToString(System.Globalization.CultureInfo.InvariantCulture) & ").")
                Dim flags As System.UInt32 = System.BitConverter.ToUInt32(System.BitConverter.GetBytes(System.Runtime.InteropServices.Marshal.ReadInt32(flagsInformation)), 0)
                ValidateShareFlags(flags)
                Return New ShareProtection With {.Descriptor = normalized, .Flags = flags}
            Catch failure As System.ArgumentException
                Throw New System.NotSupportedException("shared_policy_unsupported: The returned share security descriptor could not be interpreted safely.", failure)
            Finally
                If flagsInformation <> System.IntPtr.Zero Then NetApiBufferFree(flagsInformation)
                If information <> System.IntPtr.Zero Then NetApiBufferFree(information)
            End Try
        End Function

        Private Shared Function ValidateShareDescriptor(bytes As System.Byte()) As System.String
            Dim descriptor As New System.Security.AccessControl.RawSecurityDescriptor(bytes, 0)
            If descriptor.DiscretionaryAcl Is Nothing OrElse (descriptor.ControlFlags And System.Security.AccessControl.ControlFlags.DiscretionaryAclPresent) = 0 Then Throw New System.NotSupportedException("shared_policy_unsupported: An absent share DACL cannot establish an equivalent access boundary.")
            If descriptor.SystemAcl IsNot Nothing AndAlso descriptor.SystemAcl.Count <> 0 Then Throw New System.NotSupportedException("shared_policy_unsupported: Supplemental share policy requires an explicit protection-domain integration.")
            For Each ace As System.Security.AccessControl.GenericAce In descriptor.DiscretionaryAcl
                Dim common As System.Security.AccessControl.CommonAce = TryCast(ace, System.Security.AccessControl.CommonAce)
                If common Is Nothing OrElse common.IsCallback OrElse (common.AceQualifier <> System.Security.AccessControl.AceQualifier.AccessAllowed AndAlso common.AceQualifier <> System.Security.AccessControl.AceQualifier.AccessDenied) OrElse common.AceFlags <> System.Security.AccessControl.AceFlags.None Then Throw New System.NotSupportedException("shared_policy_unsupported: Complex or inheritable share ACEs cannot be compared automatically.")
                If common.SecurityIdentifier Is Nothing OrElse common.SecurityIdentifier.Value.StartsWith("S-1-3-", System.StringComparison.Ordinal) Then Throw New System.NotSupportedException("shared_policy_unsupported: Owner-relative share security cannot be projected.")
            Next
            Return descriptor.GetSddlForm(System.Security.AccessControl.AccessControlSections.All)
        End Function

        Private Shared Sub ValidateShareFlags(flags As System.UInt32)
            ' Exact equality includes caching, ABE, oplock, peer-cache and encryption
            ' flags. DFS, forced deletion, restricted exclusive opens, continuous
            ' availability/scoped handles and unknown future bits are unsupported.
            Const supported As System.UInt32 = &HBC30UI
            If (flags And Not supported) <> 0UI Then Throw New System.NotSupportedException("shared_policy_unsupported: DFS, cluster, exclusive-open or unknown share flags prevent safe cooperative publication.")
        End Sub

        <System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)>
        Private Structure NativeShareInfo503
            Public NetName As System.IntPtr
            Public ShareType As System.UInt32
            Public Remark As System.IntPtr
            Public Permissions As System.UInt32
            Public MaximumUses As System.UInt32
            Public CurrentUses As System.UInt32
            Public Path As System.IntPtr
            Public Password As System.IntPtr
            Public ServerName As System.IntPtr
            Public Reserved As System.UInt32
            Public SecurityDescriptor As System.IntPtr
        End Structure

        <System.Runtime.InteropServices.DllImport("netapi32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Unicode, ExactSpelling:=True)>
        Private Shared Function NetShareGetInfo(server As System.String, share As System.String, level As System.UInt32, ByRef buffer As System.IntPtr) As System.UInt32
        End Function

        <System.Runtime.InteropServices.DllImport("netapi32.dll")>
        Private Shared Function NetApiBufferFree(buffer As System.IntPtr) As System.UInt32
        End Function

        <System.Runtime.InteropServices.DllImport("advapi32.dll")>
        Private Shared Function IsValidSecurityDescriptor(descriptor As System.IntPtr) As System.Boolean
        End Function
    End Class
End Namespace
