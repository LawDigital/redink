' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' A bounded, private last-batch record. Source-specific details are authorized again before display.

' =============================================================================
' File: SemanticArchiveBuilder.Diagnostics.vb
' Purpose:
'   Bounded operation diagnostics, stage/progress records and protected persistence.
'
' Architecture / Function:
'   Records structured engineering evidence alongside build outcomes without making
'   diagnostics authoritative retrieval content.
' =============================================================================

Option Strict On
Option Explicit On
Option Infer On

Namespace SharedLibrary
    Friend NotInheritable Class SemanticArchiveOperationDiagnostic
        Public Property SchemaVersion As System.Int32 = 1
        Public Property ArchiveId As System.String = ""
        Public Property RunId As System.String = ""
        Public Property Operation As System.String = "refresh"
        Public Property SelectedScope As System.Boolean
        Public Property SelectedDocumentCount As System.Int32
        Public Property IsBackground As System.Boolean
        Public Property Status As System.String = "started"
        Public Property Stage As System.String = "starting"
        Public Property StartedUtc As System.DateTime
        Public Property UpdatedUtc As System.DateTime
        Public Property CompletedUtc As System.Nullable(Of System.DateTime)
        Public Property CountsComplete As System.Boolean
        Public Property ProcessedFiles As System.Int32
        Public Property ReusedFiles As System.Int32
        Public Property ExtractsReused As System.Int32
        Public Property CardsRebuilt As System.Int32
        Public Property SectionIndexesRebuilt As System.Int32
        Public Property RoutingGroupsBuilt As System.Int32
        Public Property DocumentsRequiringExtraction As System.Int32
        Public Property FailedFiles As System.Int32
        Public Property CoverageExcludedFiles As System.Int32
        Public Property PendingFiles As System.Int32
        Public Property DeferredFiles As System.Int32
        Public Property DiscoveryEntriesInspected As System.Int32
        Public Property DiscoveryPending As System.Boolean
        Public Property PermissionsPending As System.Boolean
        Public Property PermissionsDeferred As System.Boolean
        Public Property PermissionSourcesChecked As System.Int32
        Public Property PermissionArtifactsChecked As System.Int32
        Public Property Published As System.Boolean
        Public Property ErrorCode As System.String = ""
        Public Property ErrorType As System.String = ""
        ' Opaque source IDs and bounded reader/host diagnostics only. Rendering must
        ' resolve the source and recheck current access before exposing any detail.
        Public Property SourceDiagnostics As New System.Collections.Generic.List(Of SemanticArchiveSourceDiagnostic)()
        Public Property OmittedDiagnostics As System.Int32
    End Class

    Public NotInheritable Partial Class SemanticArchiveBuilder
        Friend Function CreateOperationDiagnostic(archiveId As System.String, options As SemanticArchiveBuildOptions) As SemanticArchiveOperationDiagnostic
            If options Is Nothing Then Throw New System.ArgumentNullException(NameOf(options))
            Return New SemanticArchiveOperationDiagnostic With {
                .ArchiveId = SemanticArchiveIdentity.ValidateId(archiveId, NameOf(archiveId)),
                .RunId = System.Guid.NewGuid().ToString("N"),
                .Operation = If(options.ReconcilePermissionsOnly, "permissions", If(options.ForceReextract, "extract",
                    If(options.IndexOnlyRebuild OrElse options.RebuildSemanticMetadata, "reindex", If(options.RetryFailures, "retry", "refresh")))),
                .SelectedScope = options.SelectedDocumentIds IsNot Nothing,
                .SelectedDocumentCount = If(options.SelectedDocumentIds Is Nothing, 0, options.SelectedDocumentIds.Count),
                .IsBackground = options.IsBackground, .StartedUtc = System.DateTime.UtcNow, .UpdatedUtc = System.DateTime.UtcNow}
        End Function

        Friend Function CreateOperationProgress(record As SemanticArchiveOperationDiagnostic,
                                                 progress As System.IProgress(Of SemanticArchiveBuildProgress)) As System.IProgress(Of SemanticArchiveBuildProgress)
            If record Is Nothing Then Throw New System.ArgumentNullException(NameOf(record))
            Return New OperationDiagnosticProgress(record, progress)
        End Function

        ''' <summary>
        ''' Writes only the start and completion of the latest bounded batch. Progress updates
        ''' its in-memory stage and typed bounded source events, never a raw progress/model response.
        ''' A diagnostic failure returns a separate warning and cannot replace a build failure.
        ''' </summary>
        Friend Function SaveOperationDiagnostic(record As SemanticArchiveOperationDiagnostic,
                                                result As SemanticArchiveBuildResult, failure As System.Exception,
                                                completed As System.Boolean) As System.String
            Try
                If record Is Nothing Then Throw New System.ArgumentNullException(NameOf(record))
                SyncLock record
                    record.UpdatedUtc = System.DateTime.UtcNow
                    record.CompletedUtc = If(completed, CType(record.UpdatedUtc, System.Nullable(Of System.DateTime)), Nothing)
                    If result IsNot Nothing Then
                        record.CountsComplete = True
                        record.ProcessedFiles = result.ProcessedFiles
                        record.ReusedFiles = result.ReusedFiles
                        record.ExtractsReused = result.ExtractsReused
                        record.CardsRebuilt = result.CardsRebuilt
                        record.SectionIndexesRebuilt = result.SectionIndexesRebuilt
                        record.RoutingGroupsBuilt = result.RoutingGroupsBuilt
                        record.DocumentsRequiringExtraction = result.DocumentsRequiringExtraction
                        record.FailedFiles = result.FailedFiles
                        record.PendingFiles = result.PendingFiles
                        record.DeferredFiles = result.DeferredFiles
                        record.DiscoveryEntriesInspected = result.DiscoveryEntriesInspected
                        record.DiscoveryPending = result.DiscoveryPending
                        record.PermissionsPending = result.PermissionsPending
                        record.PermissionsDeferred = result.PermissionsDeferred
                        record.PermissionSourcesChecked = result.PermissionSourcesChecked
                        record.PermissionArtifactsChecked = result.PermissionArtifactsChecked
                        record.Published = result.Published
                        For Each message As System.String In result.Diagnostics
                            CaptureDiagnostic(record, ParseSourceDiagnostic(message))
                        Next
                    End If
                    record.Status = OperationStatus(result, failure, completed)
                    record.ErrorCode = OperationErrorCode(failure)
                    record.ErrorType = If(failure Is Nothing, "", SafeExceptionType(failure))
                    ValidateOperationDiagnostic(record, record.ArchiveId)
                    Dim serialized As System.String = Newtonsoft.Json.JsonConvert.SerializeObject(record, Newtonsoft.Json.Formatting.None)
                    While System.Text.Encoding.UTF8.GetByteCount(serialized) > SharedMethods.DEFAULT_SEMANTICARCHIVE_DIAGNOSTIC_MAXIMUM_BYTES
                        If record.SourceDiagnostics.Count = 0 Then Throw New System.IO.InvalidDataException("The operation diagnostic exceeds its bounded size.")
                        ' Retain actionable events before technical routing events.
                        Dim removeIndex As System.Int32 = record.SourceDiagnostics.FindLastIndex(Function(entry) IsTechnicalDiagnostic(RawDiagnostic(entry)))
                        If removeIndex < 0 Then removeIndex = record.SourceDiagnostics.Count - 1
                        record.SourceDiagnostics.RemoveAt(removeIndex)
                        If record.OmittedDiagnostics < System.Int32.MaxValue Then record.OmittedDiagnostics += 1
                        serialized = Newtonsoft.Json.JsonConvert.SerializeObject(record, Newtonsoft.Json.Formatting.None)
                    End While
                    If _store.GetArchive(record.ArchiveId) Is Nothing Then
                        Return "operation_diagnostic_write_failed: archive_not_registered. The build result remains authoritative."
                    End If
                    Dim directory As System.String = _store.GetWorkDirectory(record.ArchiveId)
                    SemanticArchiveStore.CreatePrivateDirectory(directory)
                    Dim path As System.String = OperationDiagnosticPath(_store, record.ArchiveId)
                    SemanticArchiveStore.AtomicWriteJson(path, record)
                End SyncLock
                Return ""
            Catch diagnosticFailure As System.Exception
                Dim warning As System.String = "operation_diagnostic_write_failed: " & SafeExceptionType(diagnosticFailure) &
                    ". The build result or original processing failure remains authoritative."
                System.Diagnostics.Trace.TraceWarning(warning)
                Return warning
            End Try
        End Function

        Friend Shared Function ReadOperationDiagnostics(store As SemanticArchiveStore, archiveId As System.String,
                        Optional includeTechnical As System.Boolean = SharedMethods.DEFAULT_SEMANTICARCHIVE_SHOW_TECHNICAL_DIAGNOSTICS) As System.Collections.Generic.List(Of System.String)
            If store Is Nothing Then Throw New System.ArgumentNullException(NameOf(store))
            SemanticArchiveIdentity.ValidateId(archiveId, NameOf(archiveId))
            Dim values As New System.Collections.Generic.List(Of System.String)()
            Try
                Dim record As SemanticArchiveOperationDiagnostic = SemanticArchiveWorkQueue.ReadPrivateControl(Of SemanticArchiveOperationDiagnostic)(
                    store.GetWorkDirectory(archiveId), OperationDiagnosticPath(store, archiveId), SharedMethods.DEFAULT_SEMANTICARCHIVE_DIAGNOSTIC_MAXIMUM_BYTES)
                If record Is Nothing Then Return values
                ValidateOperationDiagnostic(record, archiveId)
                Dim recorded As New System.Collections.Generic.List(Of System.String)()
                For Each entry As SemanticArchiveSourceDiagnostic In record.SourceDiagnostics
                    recorded.Add(RawDiagnostic(entry) & If(entry.Occurrences > 1, " (observed " & entry.Occurrences.ToString(System.Globalization.CultureInfo.InvariantCulture) & " times)", ""))
                Next
                values.AddRange(PresentDiagnostics(store, archiveId, recorded, includeTechnical))
                If record.OmittedDiagnostics > 0 Then values.Add(record.OmittedDiagnostics.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                    " diagnostic events omitted from the bounded last-run record; per-document coverage remains in the published source records.")
                If includeTechnical Then
                    values.Add("Last recorded batch: " & record.Operation & "; status: " & record.Status & "; stage: " & record.Stage &
                        "; started UTC: " & record.StartedUtc.ToString("O", System.Globalization.CultureInfo.InvariantCulture) &
                        "; updated UTC: " & record.UpdatedUtc.ToString("O", System.Globalization.CultureInfo.InvariantCulture) & ".")
                    values.Add("Latest bounded batch counts (" & If(record.CountsComplete, "final", "observed before completion") & "): processed: " &
                        Number(record.ProcessedFiles) & "; reused: " & Number(record.ReusedFiles) & "; extracts reused: " & Number(record.ExtractsReused) &
                        "; cards rebuilt: " & Number(record.CardsRebuilt) & "; section indexes rebuilt: " & Number(record.SectionIndexesRebuilt) &
                        "; routing groups built: " & Number(record.RoutingGroupsBuilt) & "; requires extraction: " & Number(record.DocumentsRequiringExtraction) &
                        "; failed: " & Number(record.FailedFiles) & "; excluded from search: " & Number(record.CoverageExcludedFiles) &
                        "; ready: " & Number(record.PendingFiles) & "; deferred: " & Number(record.DeferredFiles) &
                        "; discovery entries: " & Number(record.DiscoveryEntriesInspected) & "; published: " & record.Published.ToString() & ".")
                    If record.DiscoveryPending OrElse record.PermissionsPending OrElse record.PermissionsDeferred Then
                        values.Add("Remaining work: discovery pending: " & record.DiscoveryPending.ToString() &
                            "; permissions pending: " & record.PermissionsPending.ToString() &
                            "; permissions deferred: " & record.PermissionsDeferred.ToString() & ".")
                    End If
                ElseIf record.Status = "pending" OrElse record.Status = "deferred" OrElse record.DiscoveryPending OrElse record.PermissionsPending OrElse record.PermissionsDeferred Then
                    values.Add("Maintenance is not finished yet. Remaining work is checkpointed and can continue without restarting completed documents.")
                ElseIf record.Status = "completed_with_requirements" Then
                    values.Add("The semantic index rebuild completed for reusable text. Some documents still require a separate extraction/OCR action before they can be indexed.")
                ElseIf record.Status = "failed" OrElse record.Status = "completed_with_failures" OrElse record.ErrorCode.Length > 0 Then
                    values.Add("The last maintenance run needs attention. Enable Show technical details for the recorded failure information.")
                End If
                If record.Status = "started" Then values.Add("The last maintenance run did not record a completion. It may still be running or may have been interrupted.")
                If includeTechnical AndAlso record.Status = "deferred" Then values.Add("The last batch yielded. Queue diagnostics below distinguish retry backoff, a required host reader and writer contention.")
                If includeTechnical AndAlso record.Status = "completed" AndAlso record.ProcessedFiles = 0 AndAlso Not record.Published Then
                    values.Add("The last batch had no runnable changes. Check the configured source roots and the ready/deferred queue counts before retrying.")
                End If
                If record.CoverageExcludedFiles > 0 Then values.Add("Some documents are not searchable yet. Their filenames and recommended action are listed above when current access permits it.")
                If includeTechnical AndAlso record.ErrorCode.Length > 0 Then values.Add("Last operation failure: " & record.ErrorCode & "; type: " & record.ErrorType &
                    ". Source-specific failure details are shown below only when current source access permits them.")
            Catch diagnosticFailure As System.Exception
                values.Add("operation_diagnostic_read_failed: " & SafeExceptionType(diagnosticFailure) &
                    ". The stored last-batch record is unavailable; source queue diagnostics remain separate.")
            End Try
            Return values
        End Function

        Private Shared Function OperationDiagnosticPath(store As SemanticArchiveStore, archiveId As System.String) As System.String
            Dim directory As System.String = store.GetWorkDirectory(archiveId)
            Dim path As System.String = SemanticArchivePathGuard.RequireWindowsCompatiblePath(System.IO.Path.Combine(directory, "last-run.json"))
            Return SemanticArchivePathGuard.ValidateContainedPath(directory, path, False)
        End Function

        Private Shared Function Number(value As System.Int32) As System.String
            Return value.ToString(System.Globalization.CultureInfo.InvariantCulture)
        End Function

        Private Shared Function OperationStatus(result As SemanticArchiveBuildResult, failure As System.Exception, completed As System.Boolean) As System.String
            If Not completed Then Return "started"
            If TypeOf failure Is System.OperationCanceledException OrElse (result IsNot Nothing AndAlso result.Cancelled) Then Return "cancelled"
            If failure IsNot Nothing Then Return "failed"
            If result Is Nothing Then Return "incomplete"
            If result.SelectionRequired Then Return "selection_required"
            If result.WriterLeaseDeferred Then Return "deferred"
            If result.FailedFiles > 0 Then Return "completed_with_failures"
            If result.DocumentsRequiringExtraction > 0 Then Return "completed_with_requirements"
            If result.DeferredFiles > 0 OrElse result.PermissionsDeferred Then Return "deferred"
            If result.PendingFiles > 0 OrElse result.DiscoveryPending OrElse result.PermissionsPending Then Return "pending"
            Return "completed"
        End Function

        Private Shared Function OperationErrorCode(failure As System.Exception) As System.String
            If failure Is Nothing Then Return ""
            Dim interaction As SharedMethods.HeadlessInteractionRequiredException = TryCast(failure, SharedMethods.HeadlessInteractionRequiredException)
            If interaction IsNot Nothing Then
                Return If(interaction.Code = "noninteractive_auth_required", "noninteractive_auth_required", "headless_interaction_required")
            End If
            If TypeOf failure Is System.OperationCanceledException Then Return "cancelled"
            If TypeOf failure Is System.IO.PathTooLongException Then Return "path_too_long"
            If TypeOf failure Is System.UnauthorizedAccessException Then Return "access_denied"
            If TypeOf failure Is System.TimeoutException Then Return "operation_timeout"
            If TypeOf failure Is System.IO.InvalidDataException Then Return "invalid_data"
            If TypeOf failure Is System.IO.IOException Then Return "io_error"
            If TypeOf failure Is System.Configuration.ConfigurationErrorsException OrElse TypeOf failure Is System.FormatException OrElse
                TypeOf failure Is System.ArgumentException Then Return "invalid_configuration"
            Return "operation_failed"
        End Function

        Private Shared Function SafeExceptionType(failure As System.Exception) As System.String
            If failure Is Nothing Then Return ""
            Dim name As System.String = failure.GetType().FullName
            If Not IsExceptionTypeToken(name) Then Return "System.Exception"
            Return name
        End Function

        Private Shared Function IsExceptionTypeToken(value As System.String) As System.Boolean
            If System.String.IsNullOrWhiteSpace(value) OrElse value.Length > 256 Then Return False
            For Each character As System.Char In value
                If Not ((character >= "a"c AndAlso character <= "z"c) OrElse (character >= "A"c AndAlso character <= "Z"c) OrElse
                    (character >= "0"c AndAlso character <= "9"c) OrElse character = "."c OrElse character = "_"c OrElse character = "+"c OrElse character = "`"c) Then Return False
            Next
            Return True
        End Function

        Private Shared Function KnownStage(value As System.String) As System.String
            Select Case value
                Case "starting", "scanning", "validating_extracts", "processing", "permissions", "hierarchy", "routing", "publishing", "failed", "pending_host", "needs_extraction", "coverage_excluded", "shared_claim_deferred", "operation_failed", "operation_cancelled"
                    Return value
                Case Else
                    Return "other"
            End Select
        End Function

        Private Shared Sub ValidateOperationDiagnostic(record As SemanticArchiveOperationDiagnostic, archiveId As System.String)
            If record.SourceDiagnostics Is Nothing OrElse record.SourceDiagnostics.Count > SharedMethods.DEFAULT_SEMANTICARCHIVE_DIAGNOSTIC_MAXIMUM_EVENTS OrElse record.OmittedDiagnostics < 0 Then
                Throw New System.IO.InvalidDataException("The source diagnostic list is invalid.")
            End If
            For Each entry As SemanticArchiveSourceDiagnostic In record.SourceDiagnostics
                If entry Is Nothing OrElse entry.Code Is Nothing OrElse entry.DocumentId Is Nothing OrElse entry.Detail Is Nothing OrElse
                    entry.Detail.Length > SharedMethods.DEFAULT_SEMANTICARCHIVE_DIAGNOSTIC_EVENT_CHARACTERS + 32 OrElse entry.Occurrences < 1 Then
                    Throw New System.IO.InvalidDataException("The source diagnostic entry exceeds its bounded contract.")
                End If
                Dim checked As SemanticArchiveSourceDiagnostic = ParseSourceDiagnostic(RawDiagnostic(entry))
                If checked Is Nothing OrElse checked.Code <> entry.Code OrElse checked.DocumentId <> entry.DocumentId Then
                    Throw New System.IO.InvalidDataException("The source diagnostic identity is invalid.")
                End If
            Next
            If record.SchemaVersion <> 1 OrElse record.ArchiveId <> archiveId Then Throw New System.IO.InvalidDataException("The operation diagnostic identity is invalid.")
            SemanticArchiveIdentity.ValidateId(record.ArchiveId, NameOf(record.ArchiveId))
            Dim runId As System.Guid
            If Not System.Guid.TryParseExact(record.RunId, "N", runId) Then Throw New System.IO.InvalidDataException("The operation diagnostic run identity is invalid.")
            If record.StartedUtc = System.DateTime.MinValue OrElse record.UpdatedUtc = System.DateTime.MinValue Then Throw New System.IO.InvalidDataException("The operation diagnostic timestamps are invalid.")
            If record.Stage <> "other" AndAlso KnownStage(record.Stage) <> record.Stage Then Throw New System.IO.InvalidDataException("The operation diagnostic stage is invalid.")
            Select Case record.Operation
                Case "refresh", "permissions", "extract", "reindex", "retry"
                Case Else : Throw New System.IO.InvalidDataException("The operation diagnostic command is invalid.")
            End Select
            Select Case record.Status
                Case "started", "cancelled", "failed", "incomplete", "selection_required", "deferred", "completed_with_failures", "completed_with_requirements", "pending", "completed"
                Case Else : Throw New System.IO.InvalidDataException("The operation diagnostic status is invalid.")
            End Select
            Select Case record.ErrorCode
                Case "", "noninteractive_auth_required", "headless_interaction_required", "cancelled", "path_too_long", "access_denied", "operation_timeout", "invalid_data", "io_error", "invalid_configuration", "operation_failed"
                Case Else : Throw New System.IO.InvalidDataException("The operation diagnostic error code is invalid.")
            End Select
            If record.ErrorType Is Nothing OrElse (record.ErrorType.Length > 0 AndAlso Not IsExceptionTypeToken(record.ErrorType)) Then Throw New System.IO.InvalidDataException("The operation diagnostic error type is invalid.")
            For Each count As System.Int32 In New System.Int32() {record.SelectedDocumentCount, record.ProcessedFiles, record.ReusedFiles,
                record.ExtractsReused, record.CardsRebuilt, record.SectionIndexesRebuilt, record.RoutingGroupsBuilt, record.DocumentsRequiringExtraction,
                record.FailedFiles, record.CoverageExcludedFiles, record.PendingFiles, record.DeferredFiles, record.DiscoveryEntriesInspected,
                record.PermissionSourcesChecked, record.PermissionArtifactsChecked}
                If count < 0 Then Throw New System.IO.InvalidDataException("The operation diagnostic count is invalid.")
            Next
        End Sub

        Private NotInheritable Class OperationDiagnosticProgress
            Implements System.IProgress(Of SemanticArchiveBuildProgress)
            Private ReadOnly _record As SemanticArchiveOperationDiagnostic
            Private ReadOnly _inner As System.IProgress(Of SemanticArchiveBuildProgress)

            Public Sub New(record As SemanticArchiveOperationDiagnostic, inner As System.IProgress(Of SemanticArchiveBuildProgress))
                _record = record
                _inner = inner
            End Sub

            Public Sub Report(value As SemanticArchiveBuildProgress) Implements System.IProgress(Of SemanticArchiveBuildProgress).Report
                If value Is Nothing Then Return
                SyncLock _record
                    _record.Stage = KnownStage(value.Stage)
                    CaptureDiagnostic(_record, value.SourceDiagnostic)
                    _record.ProcessedFiles = System.Math.Max(0, value.CompletedFiles)
                    _record.PendingFiles = System.Math.Max(0, value.PendingFiles)
                    _record.DeferredFiles = System.Math.Max(0, value.DeferredFiles)
                    If value.Stage = "failed" Then
                        If _record.FailedFiles < System.Int32.MaxValue Then _record.FailedFiles += 1
                    End If
                    If value.Stage = "coverage_excluded" AndAlso _record.CoverageExcludedFiles < System.Int32.MaxValue Then _record.CoverageExcludedFiles += 1
                End SyncLock
                If _inner IsNot Nothing Then _inner.Report(value)
            End Sub
        End Class
    End Class
End Namespace
