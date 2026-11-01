' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.License.Identifiers.vb
' PURPOSE
'   Resolves canonical offline-license identities, including Windows joined tenant identity.
'
' KEY ARCHITECTURE
'   • Identity: Normalizes typed offline-license and Windows tenant identifiers.
'   • Trust: Tenant identity comes from the Windows join API, not cache/config text.
' =============================================================================

'
' Local offline-license identities. Tenant identity is returned by the documented
' Windows NetGetAadJoinInformation API, never inferred from configuration/cache text.
' References: lmjoin.h / DSREG_JOIN_INFO / DSREG_JOIN_TYPE on Microsoft Learn.
Option Explicit On
Option Strict On

Namespace SharedLibrary
    Partial Public Class SharedMethods
        Private Const OfflineTenantIdentifierPrefix As System.String = "tenant:"

        Private NotInheritable Class OfflineLicenseIdentitySnapshot
            Public ReadOnly Property Identifiers As New System.Collections.Generic.List(Of System.String)()
            Public ReadOnly Property Details As New System.Collections.Generic.List(Of System.String)()
        End Class

        Private NotInheritable Class OfflineTenantIdentity
            Public Property TenantId As System.String
            Public Property JoinType As System.Int32
        End Class

        Private Shared Function TryNormalizeOfflineTenantIdentifier(value As System.String,
                                                                     ByRef canonical As System.String) As System.Boolean
            canonical = System.String.Empty
            Dim raw As System.String = If(value, System.String.Empty).Trim()
            If Not raw.StartsWith(OfflineTenantIdentifierPrefix, System.StringComparison.OrdinalIgnoreCase) Then Return False
            Dim id As System.Guid
            If Not System.Guid.TryParse(raw.Substring(OfflineTenantIdentifierPrefix.Length).Trim(), id) OrElse id = System.Guid.Empty Then Return False
            canonical = OfflineTenantIdentifierPrefix & id.ToString("D").ToLowerInvariant()
            Return True
        End Function

        Private Shared Function GetCurrentOfflineLicenseIdentitySnapshot(
            requestedIdentifiers As System.Collections.Generic.IEnumerable(Of System.String),
            Optional includeTenant As System.Boolean = True) As OfflineLicenseIdentitySnapshot

            Dim snapshot As New OfflineLicenseIdentitySnapshot()
            For Each value As System.String In GetCurrentNetworkDomainCandidates()
                ' Environment/DNS strings must never impersonate a typed tenant identity.
                If value.StartsWith(OfflineTenantIdentifierPrefix, System.StringComparison.OrdinalIgnoreCase) Then Continue For
                AddNormalizedDomain(snapshot.Identifiers, value)
            Next
            Dim computerName As System.String = GetCurrentClientIdentifier()
            If Not If(computerName, System.String.Empty).StartsWith(OfflineTenantIdentifierPrefix, System.StringComparison.OrdinalIgnoreCase) Then
                AddNormalizedDomain(snapshot.Identifiers, computerName)
            End If
            snapshot.Details.Add("Network / host IDs: existing Windows and DNS detection. Computer / client ID: " & If(computerName, System.String.Empty))
            If Not includeTenant Then Return snapshot

            Dim queried As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            Dim identity As OfflineTenantIdentity = Nothing
            Dim diagnostic As System.String = System.String.Empty
            If TryReadOfflineTenantIdentity(Nothing, identity, diagnostic) Then
                AddOfflineTenantIdentity(snapshot, identity)
                queried.Add(identity.TenantId)
            Else
                snapshot.Details.Add("Windows default tenant: " & diagnostic)
            End If
            If requestedIdentifiers IsNot Nothing Then
                For Each value As System.String In requestedIdentifiers
                    Dim canonical As System.String = System.String.Empty
                    If Not TryNormalizeOfflineTenantIdentifier(value, canonical) Then Continue For
                    Dim tenantId As System.String = canonical.Substring(OfflineTenantIdentifierPrefix.Length)
                    If Not queried.Add(tenantId) Then Continue For
                    If TryReadOfflineTenantIdentity(tenantId, identity, diagnostic) Then
                        AddOfflineTenantIdentity(snapshot, identity)
                    Else
                        snapshot.Details.Add(canonical & ": " & diagnostic)
                    End If
                Next
            End If
            Return snapshot
        End Function

        Private Shared Sub AddOfflineTenantIdentity(snapshot As OfflineLicenseIdentitySnapshot, identity As OfflineTenantIdentity)
            Dim value As System.String = OfflineTenantIdentifierPrefix & identity.TenantId
            If snapshot.Identifiers.Contains(value) Then Return
            AddNormalizedDomain(snapshot.Identifiers, value)
            snapshot.Details.Add(value & " — " & If(identity.JoinType = 1, "Windows Entra device join (including hybrid join)", "Windows work-account registration for the current user"))
        End Sub

        Private Shared Function TryReadOfflineTenantIdentity(requestedTenantId As System.String,
                                                             ByRef identity As OfflineTenantIdentity,
                                                             ByRef diagnostic As System.String) As System.Boolean
            identity = Nothing
            diagnostic = System.String.Empty
            Dim buffer As System.IntPtr = System.IntPtr.Zero
            Try
                Dim result As System.Int32 = OfflineTenantNative.NetGetAadJoinInformation(requestedTenantId, buffer)
                If result <> 0 Then
                    diagnostic = "Windows query failed (HRESULT 0x" & result.ToString("X8", System.Globalization.CultureInfo.InvariantCulture) & ")."
                    Return False
                End If
                If buffer = System.IntPtr.Zero Then
                    diagnostic = "No local Windows Entra join/work-account registration was returned."
                    Return False
                End If
                Dim info As OfflineTenantNative.JoinInfo = System.Runtime.InteropServices.Marshal.PtrToStructure(Of OfflineTenantNative.JoinInfo)(buffer)
                If info.JoinType <> 1 AndAlso info.JoinType <> 2 Then
                    diagnostic = "Windows returned no supported join type."
                    Return False
                End If
                Dim raw As System.String = System.Runtime.InteropServices.Marshal.PtrToStringUni(info.TenantId)
                Dim tenant As System.Guid
                If Not System.Guid.TryParse(raw, tenant) OrElse tenant = System.Guid.Empty Then
                    diagnostic = "Windows returned an invalid or empty tenant ID."
                    Return False
                End If
                Dim tenantId As System.String = tenant.ToString("D").ToLowerInvariant()
                If Not System.String.IsNullOrWhiteSpace(requestedTenantId) AndAlso
                   Not System.String.Equals(tenantId, requestedTenantId, System.StringComparison.OrdinalIgnoreCase) Then
                    diagnostic = "Windows returned a different tenant; the requested identity was not accepted."
                    Return False
                End If
                identity = New OfflineTenantIdentity With {.TenantId = tenantId, .JoinType = info.JoinType}
                Return True
            Catch ex As System.Exception
                diagnostic = "Windows Entra identity is unavailable (" & ex.GetType().Name & ")."
                System.Diagnostics.Trace.TraceWarning("Offline-license Windows tenant query unavailable: " & ex.Message)
                Return False
            Finally
                If buffer <> System.IntPtr.Zero Then
                    OfflineTenantNative.NetFreeAadJoinInformation(buffer)
                End If
            End Try
        End Function

        Private NotInheritable Class OfflineTenantNative
            ' Native layout: enum (32-bit), then eleven pointer-sized members.
            ' Pointer fields preserve the SDK's x86/x64 padding and native ownership.
            <System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)>
            Friend Structure JoinInfo
                Public JoinType As System.Int32
                Public JoinCertificate As System.IntPtr
                Public DeviceId As System.IntPtr
                Public IdpDomain As System.IntPtr
                Public TenantId As System.IntPtr
                Public JoinUserEmail As System.IntPtr
                Public TenantDisplayName As System.IntPtr
                Public MdmEnrollmentUrl As System.IntPtr
                Public MdmTermsOfUseUrl As System.IntPtr
                Public MdmComplianceUrl As System.IntPtr
                Public UserSettingSyncUrl As System.IntPtr
                Public UserInfo As System.IntPtr
            End Structure

            <System.Runtime.InteropServices.DllImport("Netapi32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Unicode, ExactSpelling:=True)>
            Friend Shared Function NetGetAadJoinInformation(
                <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr)> tenantId As System.String,
                ByRef joinInfo As System.IntPtr) As System.Int32
            End Function

            <System.Runtime.InteropServices.DllImport("Netapi32.dll", ExactSpelling:=True)>
            Friend Shared Sub NetFreeAadJoinInformation(joinInfo As System.IntPtr)
            End Sub
        End Class
    End Class
End Namespace
