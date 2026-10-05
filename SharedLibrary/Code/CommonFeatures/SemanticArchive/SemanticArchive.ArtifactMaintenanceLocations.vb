' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
Option Strict On
Option Explicit On

Namespace SharedLibrary
    Public NotInheritable Partial Class SemanticArchiveArtifactPlanner
        Private NotInheritable Class PermissionRoute
            Public BindingSignature As System.String
            Public SourcePath As System.String
            Public Roots As System.Collections.Generic.List(Of System.String)
            Public Position As System.Int32
            Public ChildToken As System.String = ""
            Public ExpiresUtc As System.DateTimeOffset
            Public FailureStatus As System.String = ""
            Public Diagnostics As New System.Collections.Generic.List(Of System.String)()
            Public AnySharedArtifacts As System.Boolean
            Public AnyQuarantined As System.Boolean
        End Class

        Private Shared ReadOnly PermissionRoutes As New System.Collections.Generic.Dictionary(Of System.String, PermissionRoute)(System.StringComparer.Ordinal)

        ''' <summary>
        ''' Maintains only the current configured shared artifact root with one
        ''' bounded opaque cursor. Development-era artifact locations are not searched.
        ''' The original source binding is retained for all identity/authority checks.
        ''' </summary>
        Public Shared Function ReconcilePermissions(binding As SemanticArchiveSourceBinding, sourcePath As System.String, Optional continuationToken As System.String = "", Optional maximumArtifacts As System.Int32 = 64, Optional cancellationToken As System.Threading.CancellationToken = Nothing) As SemanticArchivePermissionRepairResult
            If binding Is Nothing Then Throw New System.ArgumentNullException(NameOf(binding))
            If maximumArtifacts < 1 OrElse maximumArtifacts > 1024 Then Throw New System.ArgumentOutOfRangeException(NameOf(maximumArtifacts))
            cancellationToken.ThrowIfCancellationRequested()
            Dim full As System.String = SemanticArchivePathGuard.RequireWindowsSourcePath(sourcePath)
            Dim signature As System.String = SemanticArchiveIdentity.HashBytes(System.Text.Encoding.UTF8.GetBytes(Newtonsoft.Json.JsonConvert.SerializeObject(binding)))
            Dim result As New SemanticArchivePermissionRepairResult()
            SyncLock RepairGate
                ExpirePermissionRoutes()
                Dim token As System.String = continuationToken
                Dim route As PermissionRoute = Nothing
                Try
                    If System.String.IsNullOrWhiteSpace(token) Then
                        If PermissionRoutes.Count >= 16 Then Return New SemanticArchivePermissionRepairResult With {.Status = "deferred", .Diagnostic = "The bounded rights-location cursor capacity is in use."}
                        Dim roots As New System.Collections.Generic.List(Of System.String)()
                        Dim current As System.String = SemanticArchivePathGuard.CanonicalPath(If(System.String.IsNullOrWhiteSpace(binding.SharedArtifactRoot), binding.RootPath, binding.SharedArtifactRoot))
                        roots.Add(current)
                        route = New PermissionRoute With {.BindingSignature = signature, .SourcePath = full, .Roots = roots, .ExpiresUtc = System.DateTimeOffset.UtcNow.AddMinutes(5)}
                        token = System.Guid.NewGuid().ToString("N")
                        PermissionRoutes.Add(token, route)
                    ElseIf Not PermissionRoutes.TryGetValue(token, route) Then
                        Return New SemanticArchivePermissionRepairResult With {.Status = "restart_required", .Diagnostic = "The process-local rights-location cursor expired; restart this source."}
                    End If
                    If route.BindingSignature <> signature OrElse Not System.String.Equals(route.SourcePath, full, System.StringComparison.OrdinalIgnoreCase) Then
                        FinishPermissionRoute(token)
                        Return New SemanticArchivePermissionRepairResult With {.Status = "restart_required", .Diagnostic = "The source or placement binding changed; restart rights maintenance."}
                    End If
                    Do While route.Position < route.Roots.Count
                        cancellationToken.ThrowIfCancellationRequested()
                        Dim scopedBinding As SemanticArchiveSourceBinding = CloneBindingForSharedRoot(binding, route.Roots(route.Position))
                        Dim remaining As System.Int32 = System.Math.Max(1, maximumArtifacts - result.CheckedArtifacts)
                        Dim batch As SemanticArchivePermissionRepairResult = ReconcilePermissionsAtLocation(scopedBinding, full, route.ChildToken, remaining, cancellationToken)
                        result.CheckedArtifacts += batch.CheckedArtifacts
                        result.RepairedArtifacts += batch.RepairedArtifacts
                        result.QuarantinedArtifacts += batch.QuarantinedArtifacts
                        result.SourcePermissionSignature = batch.SourcePermissionSignature
                        If Not System.String.IsNullOrWhiteSpace(batch.Diagnostic) AndAlso route.Diagnostics.Count < 8 Then route.Diagnostics.Add(batch.Diagnostic)
                        route.ExpiresUtc = System.DateTimeOffset.UtcNow.AddMinutes(5)
                        If batch.Status = "partial" OrElse batch.Status = "restart_required" Then
                            route.ChildToken = If(batch.Status = "partial", batch.ContinuationToken, "")
                            result.Status = "partial"
                            result.ContinuationToken = token
                            result.Diagnostic = System.String.Join(System.Environment.NewLine, route.Diagnostics)
                            Return result
                        End If
                        route.ChildToken = ""
                        Select Case batch.Status
                            Case "complete"
                                route.AnySharedArtifacts = True
                            Case "quarantined"
                                route.AnySharedArtifacts = True
                                route.AnyQuarantined = True
                            Case "private", "no_artifacts"
                                ' This exact candidate has no generated artifacts.
                            Case Else
                                ' Continue the other known location even when one is
                                ' offline or unrepairable. Failure remains visible and
                                ' prevents reporting the whole source fully reconciled.
                                route.FailureStatus = If(batch.Status = "deferred" AndAlso route.FailureStatus = "", "deferred", "unrepairable")
                        End Select
                        route.Position += 1
                        If route.Position < route.Roots.Count AndAlso result.CheckedArtifacts >= maximumArtifacts Then
                            result.Status = "partial"
                            result.ContinuationToken = token
                            result.Diagnostic = System.String.Join(System.Environment.NewLine, route.Diagnostics)
                            Return result
                        End If
                    Loop
                    result.Status = If(route.FailureStatus <> "", route.FailureStatus, If(route.AnyQuarantined, "quarantined", If(route.AnySharedArtifacts, "complete", If(binding.ArtifactPlacementMode = "private", "private", "no_artifacts"))))
                    result.Diagnostic = System.String.Join(System.Environment.NewLine, route.Diagnostics)
                    FinishPermissionRoute(token)
                    Return result
                Catch
                    If Not System.String.IsNullOrWhiteSpace(token) Then FinishPermissionRoute(token)
                    Throw
                End Try
            End SyncLock
        End Function

        Private Shared Sub FinishPermissionRoute(token As System.String)
            Dim route As PermissionRoute = Nothing
            If PermissionRoutes.TryGetValue(token, route) Then
                PermissionRoutes.Remove(token)
                If Not System.String.IsNullOrWhiteSpace(route.ChildToken) Then FinishRepairSession(route.ChildToken)
            End If
        End Sub

        Private Shared Sub ExpirePermissionRoutes()
            Dim expired As New System.Collections.Generic.List(Of System.String)()
            For Each entry As System.Collections.Generic.KeyValuePair(Of System.String, PermissionRoute) In PermissionRoutes
                If entry.Value.ExpiresUtc <= System.DateTimeOffset.UtcNow Then expired.Add(entry.Key)
            Next
            For Each token As System.String In expired
                FinishPermissionRoute(token)
            Next
        End Sub
    End Class
End Namespace
