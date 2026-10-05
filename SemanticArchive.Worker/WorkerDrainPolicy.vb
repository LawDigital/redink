' Part of "Red Ink" (Red Ink Semantic Archive Worker)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: WorkerDrainPolicy.vb
' Purpose:
'   Determines whether another worker batch can make progress without looping
'   indefinitely on deferred work.
'
' Architecture / Function:
'   Continues only while runnable work and observed progress remain; cancellation,
'   writer contention and required selection stop draining.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SemanticArchiveWorker
    Friend NotInheritable Class WorkerDrainPolicy
        Private Sub New()
        End Sub

        Friend Shared Function ShouldContinue(result As Global.SharedLibrary.SharedLibrary.SemanticArchiveBuildResult) As System.Boolean
            If result Is Nothing OrElse result.Cancelled OrElse result.WriterLeaseDeferred OrElse result.SelectionRequired Then Return False
            ' Deferred source jobs do not block later runnable work. The builder advances
            ' discovery and stores ACL retries separately; only retry-only backoff stops us.
            Dim runnable As System.Boolean = result.PendingFiles > 0 OrElse result.DiscoveryPending OrElse
                (result.PermissionsPending AndAlso Not result.PermissionsDeferred)
            Dim progressed As System.Boolean = result.ProcessedFiles > 0 OrElse result.ReusedFiles > 0 OrElse
                result.DiscoveryEntriesInspected > 0 OrElse result.PermissionSourcesChecked > 0 OrElse result.PermissionArtifactsChecked > 0
            Return runnable AndAlso progressed
        End Function
    End Class
End Namespace
