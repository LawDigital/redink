' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.

Option Strict On
Option Explicit On
Option Infer On

Namespace SharedLibrary

    ''' <summary>
    ''' Manual and background archive ingestion share this provider-neutral build service.
    ''' Enumeration and storage work always run off the Office UI thread; only an explicitly
    ''' supplied host-reader dispatcher may return an existing COM reader to its owner.
    ''' </summary>
    Public NotInheritable Partial Class SemanticArchiveBuilder
        Private ReadOnly _context As SharedContext.ISharedContext
        Private ReadOnly _store As SemanticArchiveStore

        Public Sub New(context As SharedContext.ISharedContext, store As SemanticArchiveStore)
            If context Is Nothing Then Throw New System.ArgumentNullException(NameOf(context))
            If store Is Nothing Then Throw New System.ArgumentNullException(NameOf(store))
            _context = context
            _store = store
        End Sub

        Public Function BuildAsync(
            archiveId As String,
            Optional options As SemanticArchiveBuildOptions = Nothing,
            Optional progress As System.IProgress(Of SemanticArchiveBuildProgress) = Nothing,
            Optional cancellationToken As System.Threading.CancellationToken = Nothing
        ) As System.Threading.Tasks.Task(Of SemanticArchiveBuildResult)
            If Not SemanticArchiveHostIntegration.IsConfigured(_context) Then Throw New System.InvalidOperationException("semantic_archive_disabled: SemanticArchiveCatalogPathLocal is empty.")
            Dim effectiveOptions As SemanticArchiveBuildOptions = If(options, New SemanticArchiveBuildOptions())
            If effectiveOptions.MaximumFilesPerBatch < 1 OrElse effectiveOptions.MaximumFilesPerBatch > 4096 Then Throw New System.ArgumentOutOfRangeException(NameOf(options), "Archive batches must contain 1 to 4,096 sources.")
            Return System.Threading.Tasks.Task.Run(
                Function() BuildWithoutInteractiveUiAsync(archiveId, effectiveOptions, progress, cancellationToken), System.Threading.CancellationToken.None)
        End Function

        Private Async Function BuildWithoutInteractiveUiAsync(archiveId As System.String, options As SemanticArchiveBuildOptions,
                                                              progress As System.IProgress(Of SemanticArchiveBuildProgress),
                                                              cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of SemanticArchiveBuildResult)
            ' A foreground console does not make its pool-thread build an interactive UI
            ' operation. Reuse the async-flow guard used by the worker: a nested reader or
            ' model helper must report required setup, never wait behind the Office window.
            Dim diagnostic As SemanticArchiveOperationDiagnostic = CreateOperationDiagnostic(archiveId, options)
            Dim diagnosticProgress As System.IProgress(Of SemanticArchiveBuildProgress) = CreateOperationProgress(diagnostic, progress)
            Using interactionScope As SharedMethods.HeadlessExecutionScope = SharedMethods.BeginHeadlessExecution()
                Dim result As SemanticArchiveBuildResult = Nothing
                Dim startWarning As System.String = SaveOperationDiagnostic(diagnostic, Nothing, Nothing, False)
                ReportOperationDiagnosticWarning(progress, Nothing, startWarning)
                Try
                    interactionScope.ThrowIfInteractionRequested()
                    result = Await BuildCoreAsync(archiveId, options, diagnosticProgress, cancellationToken, interactionScope).ConfigureAwait(False)
                    interactionScope.ThrowIfInteractionRequested()
                    If startWarning.Length > 0 Then result.Diagnostics.Add(startWarning)
                    Dim warning As System.String = SaveOperationDiagnostic(diagnostic, result, Nothing, True)
                    ReportOperationDiagnosticWarning(progress, result, warning)
                    interactionScope.ThrowIfInteractionRequested()
                    Return result
                Catch failure As System.Exception
                    Dim warning As System.String = SaveOperationDiagnostic(diagnostic, result, failure, True)
                    ReportOperationDiagnosticWarning(progress, Nothing, warning)
                    Throw
                End Try
            End Using
        End Function

        Private Shared Sub ReportOperationDiagnosticWarning(progress As System.IProgress(Of SemanticArchiveBuildProgress),
                                                             result As SemanticArchiveBuildResult, warning As System.String)
            If System.String.IsNullOrEmpty(warning) Then Return
            If result IsNot Nothing Then result.Diagnostics.Add(warning)
            Try
                If progress IsNot Nothing Then progress.Report(New SemanticArchiveBuildProgress With {.Stage = "diagnostic_warning", .Message = warning})
            Catch reportingFailure As System.Exception
                ' A status callback must never replace the original build failure.
                System.Diagnostics.Trace.TraceWarning("operation_diagnostic_warning_display_failed: " & reportingFailure.GetType().FullName)
            End Try
        End Sub

        Private Async Function BuildCoreAsync(archiveId As String, options As SemanticArchiveBuildOptions,
                                              progress As System.IProgress(Of SemanticArchiveBuildProgress),
                                              cancellationToken As System.Threading.CancellationToken,
                                              interactionScope As SharedMethods.HeadlessExecutionScope) As System.Threading.Tasks.Task(Of SemanticArchiveBuildResult)
            Dim result As New SemanticArchiveBuildResult() With {.ArchiveId = archiveId}
            Dim queue As SemanticArchiveWorkQueue = Nothing
            Dim coverageExcludedFiles As System.Int32 = 0
            Try
                cancellationToken.ThrowIfCancellationRequested()
                If options.SelectedDocumentIds IsNot Nothing AndAlso options.SelectedDocumentIds.Count = 0 Then
                    result.SelectionRequired = True
                    result.Diagnostics.Add("empty_selection: No documents were selected; the operation was not expanded to the whole archive.")
                    Return result
                End If
                Dim lease As SemanticArchiveWriterLease = Nothing
                Dim maximumLeaseWait As System.TimeSpan? = options.MaximumWriterLeaseWait
                If Not maximumLeaseWait.HasValue AndAlso options.IsBackground Then maximumLeaseWait = System.TimeSpan.FromMilliseconds(SharedMethods.DEFAULT_SEMANTICARCHIVE_BACKGROUND_WRITER_LEASE_WAIT_MILLISECONDS)
                Try
                    lease = _store.AcquireWriterLease(archiveId, cancellationToken, maximumLeaseWait)
                Catch ex As System.TimeoutException When options.IsBackground OrElse options.MaximumWriterLeaseWait.HasValue
                    cancellationToken.ThrowIfCancellationRequested()
                    ' Contention occurred before work/metadata/queue initialization. Keep this
                    ' distinct from user cancellation, failed source jobs and model failures.
                    result.WriterLeaseDeferred = True
                    result.Diagnostics.Add("writer_lease_deferred: Another writer owns this archive. Automatic work will retry after its normal idle backoff; no lease was taken over.")
                    Return result
                End Try
                Using lease
                    Dim archive As SemanticArchiveDefinition = _store.GetArchive(archiveId)
                    SemanticArchiveLibrary.RequireCurrentDefinition(_context, archive)
                    If archive Is Nothing Then Throw New System.InvalidOperationException("The archive is no longer registered.")
                    If archive.Roots.Count = 0 Then result.Diagnostics.Add("source_roots_empty: No source directories are configured for this archive. Add a source root before refreshing documents.")
                    If options.IsBackground AndAlso Not options.ReconcilePermissionsOnly AndAlso (Not archive.Enabled OrElse (Not archive.BackgroundEnabled AndAlso Not SemanticArchiveLibrary.IsSubscriber(archive))) Then
                        result.Diagnostics.Add("Automatic processing is disabled for this archive.")
                        Return result
                    End If
                    queue = New SemanticArchiveWorkQueue(_store, archiveId)
                    Dim previous As SemanticArchiveGenerationManifest = _store.PinGenerationForMaintenance(lease)
                    If previous IsNot Nothing Then
                        result.GenerationId = previous.GenerationId
                        result.Inventory = previous.Inventory
                    End If
                    If options.ReconcilePermissionsOnly Then
                        ScanSources(archive, previous, lease, Nothing, queue, options, Nothing, "permissions",
                            ScanConfigurationSignature(archive, "permissions"), progress, result, cancellationToken)
                        Return result
                    End If
                    Dim extractionContext As SharedContext.ISharedContext = SharedMethods.CreateIsolatedModelCallContext(_context)
                    Dim resolved = SharedMethods.ResolveIsolatedSpecialTaskModel(extractionContext, "Indexer")
                    interactionScope.ThrowIfInteractionRequested()
                    Dim semanticSignature As System.String = SemanticArchiveIndexPolicy.CreateSemanticSignature(resolved.Signature,
                        SharedMethods.SemanticSearchDefaultGeneratorVersion, archive.SemanticProfileVersion,
                        archive.SectionIndexThresholdBytes, archive.AllowPartialSearch)
                    result.Diagnostics.Add("Indexer model: " & resolved.ModelName & "; configuration " & resolved.Signature &
                        If(resolved.UsedPrimaryFallback, "; primary configuration used because no accessible Indexer assignment exists.", "."))
                    Dim generation As SemanticArchiveGenerationManifest = _store.NewGeneration(lease)
                    Dim hierarchy As New SemanticArchiveHierarchy(_store, archive, resolved.Context, previous, generation, result.Diagnostics)
                    Dim scan As SemanticArchiveScanCheckpoint = queue.LoadScan()
                    Dim scanSignature As String = ScanConfigurationSignature(archive, semanticSignature)
                    If options.ForceScan OrElse scan.InProgress OrElse scan.ConfigurationSignature <> scanSignature OrElse
                        System.DateTime.UtcNow.Subtract(scan.LastAttemptUtc).TotalMinutes >= 15 Then
                        ScanSources(archive, previous, lease, hierarchy, queue, options, extractionContext, semanticSignature, scanSignature, progress, result, cancellationToken)
                    End If
                    interactionScope.ThrowIfInteractionRequested()
                    If options.FullIntegrityAudit AndAlso previous IsNot Nothing Then
                        _store.AuditGeneration(previous)
                        result.Diagnostics.Add("Full immutable-generation integrity audit completed.")
                    End If
                    Dim ready As New System.Collections.Generic.List(Of SemanticArchiveWorkItem)()
                    Dim batch As System.Collections.Generic.List(Of SemanticArchiveWorkItem) = queue.GetBatch(options.MaximumFilesPerBatch, False, options.IsBackground, options.SelectedDocumentIds)
                    For Each item As SemanticArchiveWorkItem In batch
                        cancellationToken.ThrowIfCancellationRequested()
                        Report(progress, result, item.SourcePath, "processing", "Processing the next durable source job.")
                        interactionScope.ThrowIfInteractionRequested()
                        Try
                            Dim document As SemanticArchiveDocumentRecord
                            If item.Remove Then
                                document = If(item.CachedDocument, If(previous Is Nothing, Nothing, _store.LoadDocument(previous, item.DocumentId)))
                                If document Is Nothing Then
                                    queue.Complete(item)
                                    Continue For
                                End If
                                document.Active = False
                                document.ProcessingStatus = "removed"
                                document.Diagnostic = "The source was confirmed removed from the configured archive scope."
                                _store.SaveValidity(lease, document, False, document.Diagnostic, hierarchy.Ancestors(document.DocumentId))
                            Else
                                document = Await ProcessDocumentAsync(archive, item, previous, queue, extractionContext, resolved.Context, resolved.ModelName,
                                    semanticSignature, options, result, cancellationToken, interactionScope).ConfigureAwait(False)
                                interactionScope.ThrowIfInteractionRequested()
                                _store.SaveValidity(lease, document, document.Active, document.Diagnostic, hierarchy.Ancestors(document.DocumentId))
                            End If
                            hierarchy.Apply(document)
                            item.CachedDocument = document
                            item.State = "Ready"
                            item.NextAttemptUtc = System.DateTime.MinValue
                            queue.Save(item)
                            ready.Add(item)
                            result.ProcessedFiles += 1
                            If Not document.Active AndAlso Not item.Remove Then
                                coverageExcludedFiles += 1
                                ReportSourceDiagnostic(progress, result, archive, item, "coverage_excluded",
                                    "status: " & document.ProcessingStatus & "; " & BoundProcessingDiagnostic(document.Diagnostic))
                            End If
                        Catch ex As System.OperationCanceledException When cancellationToken.IsCancellationRequested
                            Throw
                        Catch ex As SharedMethods.HeadlessInteractionRequiredException
                            ' A dispatched host reader can own a separate async-flow scope.
                            ' Its typed setup requirement still stops this batch once.
                            Throw
                        Catch ex As SemanticArchiveCooperativeBusyException
                            ' A live peer's processing claim says nothing about this
                            ' source's validity. Defer the job without replacing an
                            ' existing authorized document or suppressing its ancestors.
                            interactionScope.ThrowIfInteractionRequested()
                            Dim reason As System.String = BoundProcessingDiagnostic("shared_claim_deferred: " & ex.Message)
                            queue.Fail(item, reason, False)
                            ReportSourceDiagnostic(progress, result, archive, item, "shared_claim_deferred", reason)
                        Catch ex As System.Exception
                            ' Legacy adapters can catch an interaction exception and return
                            ' a string/result. The sticky guard must abort that operation
                            ' before accepting a failed job or publishing partial metadata.
                            interactionScope.ThrowIfInteractionRequested()
                            Dim hostRequired As Boolean = TypeOf ex Is SemanticArchiveHostRequiredException
                            Dim failureDetail As System.String = FormatProcessingFailure(ex)
                            queue.Fail(item, failureDetail, hostRequired)
                            Dim failed As SemanticArchiveDocumentRecord = FailedRecord(item, If(hostRequired, "pending_host", "failed"), failureDetail)
                            item.CachedDocument = failed
                            queue.Save(item)
                            _store.SaveValidity(lease, failed, False, failed.Diagnostic, hierarchy.Ancestors(failed.DocumentId))
                            hierarchy.Apply(failed)
                            If Not hostRequired Then result.FailedFiles += 1
                            ReportSourceDiagnostic(progress, result, archive, item, failed.ProcessingStatus, failureDetail)
                        End Try
                    Next
                    If hierarchy.HasChanges Then
                        Report(progress, result, "", "hierarchy", "Updating affected bounded card shards and their ancestors.")
                        Await hierarchy.WriteChangesAsync(cancellationToken).ConfigureAwait(False)
                        cancellationToken.ThrowIfCancellationRequested()
                        Report(progress, result, "", "publishing", "Validating and activating the immutable generation.")
                        interactionScope.ThrowIfInteractionRequested()
                        _store.PublishGeneration(lease, generation)
                        result.Published = True
                        result.GenerationId = generation.GenerationId
                        result.Inventory = generation.Inventory
                        For Each item As SemanticArchiveWorkItem In ready
                            queue.Complete(item)
                        Next
                    End If
                    result.PendingFiles = queue.PendingCount(True, options.IsBackground, options.SelectedDocumentIds)
                    result.DeferredFiles = queue.PendingCount(False, False, options.SelectedDocumentIds) - result.PendingFiles
                    If result.DeferredFiles > 0 Then result.Diagnostics.Add(result.DeferredFiles.ToString(System.Globalization.CultureInfo.InvariantCulture) & " source job(s) await retry backoff or an appropriate host reader. Refresh status shows the stored reason and next attempt; Retry failed makes selected failures due immediately.")
                End Using
            Catch ex As System.OperationCanceledException When cancellationToken.IsCancellationRequested
                result.Cancelled = True
                result.Diagnostics.Add("Cancelled at a safe boundary. Completed extraction/index checkpoints remain available for resume; the active generation was not replaced by a partial build.")
                If queue IsNot Nothing Then
                    result.PendingFiles = queue.PendingCount(True, options.IsBackground, options.SelectedDocumentIds)
                    result.DeferredFiles = queue.PendingCount(False, False, options.SelectedDocumentIds) - result.PendingFiles
                End If
            End Try
            If coverageExcludedFiles > 0 Then result.Diagnostics.Add("coverage_excluded: " & coverageExcludedFiles.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                " processed source(s) are not searchable because their extraction is empty or the current archive policy excludes incomplete/unknown coverage. Their completed extraction checkpoints remain available.")
            Return result
        End Function

        Private Async Function ProcessDocumentAsync(archive As SemanticArchiveDefinition, item As SemanticArchiveWorkItem,
                                                    previous As SemanticArchiveGenerationManifest, queue As SemanticArchiveWorkQueue,
                                                    extractionContext As SharedContext.ISharedContext,
                                                    modelContext As SharedContext.ISharedContext, modelName As String,
                                                    semanticSignature As System.String,
                                                    options As SemanticArchiveBuildOptions, result As SemanticArchiveBuildResult,
                                                    cancellationToken As System.Threading.CancellationToken,
                                                    interactionScope As SharedMethods.HeadlessExecutionScope) As System.Threading.Tasks.Task(Of SemanticArchiveDocumentRecord)
            Dim binding As SemanticArchiveSourceBinding = FindBinding(archive, item.PrimaryBindingId)
            If binding Is Nothing Then Throw New System.InvalidOperationException("The source binding was removed; refresh the archive scope.")
            If item.SemanticSignature <> semanticSignature Then
                ' Bounded discovery may not have revisited every old queued job yet.
                ' Its durable target must match the isolated model used in this batch.
                item.SemanticSignature = semanticSignature
                item.State = "Queued"
                queue.Save(item)
                result.Diagnostics.Add("semantic_configuration_changed: " & item.DocumentId & "; the processing target was refreshed before shared lookup or model generation.")
            End If
            Dim sourcePath As String = SemanticArchivePathGuard.ValidateContainedPath(binding.RootPath, item.SourcePath, True)
            If IsExcluded(binding, sourcePath) OrElse Not SemanticArchiveStore.IsSupportedSource(binding, sourcePath) Then Throw New System.UnauthorizedAccessException("The source is now excluded from archive intake.")
            If Not System.String.Equals(SemanticArchivePathGuard.GetVerifiedSourceIdentity(binding.RootPath, sourcePath), item.CanonicalSourceKey, System.StringComparison.Ordinal) Then
                Throw New System.InvalidOperationException("source_identity_changed: The queued source now resolves to a different physical original. Refresh discovery before processing it.")
            End If
            Dim exportOptions As New Global.SharedLibrary.Agents.TextExportOptions() With {
                .OcrPdf = binding.EnableOcr, .OcrBatchPages = binding.OcrBatchPages, .Overwrite = False,
                .HostReaderDispatcher = If(options.IsBackground, Nothing, options.HostReaderDispatcher)
            }
            Dim expectedExporterSignature As String = Global.SharedLibrary.Agents.TextExportService.GetProcessingSignature(extractionContext, sourcePath, exportOptions)
            Dim effectiveExtractionSignature As String = ComposeExtractionSignature(archive, binding, expectedExporterSignature)
            If item.ExtractionSignature <> effectiveExtractionSignature Then
                item.ExtractionSignature = effectiveExtractionSignature
                item.State = "Queued"
                queue.Save(item)
                result.Diagnostics.Add("extraction_configuration_changed: " & item.DocumentId & "; the processing target was refreshed before extraction.")
            End If
            Dim current As SemanticArchiveSourceFingerprint = ReadSourceFingerprint(binding.RootPath, sourcePath, cancellationToken)
            If Not System.String.IsNullOrEmpty(item.SourceHash) AndAlso Not System.String.Equals(current.Sha256, item.SourceHash, System.StringComparison.OrdinalIgnoreCase) Then
                item.SourceHash = current.Sha256
                item.SourceLength = current.Length
                item.SourceWriteTicks = current.LastWriteUtcTicks
                item.State = "Queued"
                item.CachedDocument = Nothing
                queue.Save(item)
            Else
                item.SourceHash = current.Sha256
                item.SourceLength = current.Length
                item.SourceWriteTicks = current.LastWriteUtcTicks
            End If
            Dim document As SemanticArchiveDocumentRecord = If(item.CachedDocument, If(previous Is Nothing, Nothing, _store.LoadDocument(previous, item.DocumentId)))
            If document Is Nothing Then document = New SemanticArchiveDocumentRecord()
            document = SemanticArchiveMetadata.Clone(document)
            document.DocumentId = item.DocumentId
            document.SourceItemId = item.SourceItemId
            document.CanonicalSourceKey = item.CanonicalSourceKey
            document.PartitionKey = item.PartitionKey
            document.SourcePath = sourcePath
            document.RelativePath = item.RelativePath
            document.DisplayName = System.IO.Path.GetFileName(sourcePath)
            document.BindingIds = New System.Collections.Generic.List(Of String)(item.BindingIds)
            document.Fingerprint = current
            document.ExtractionSignature = item.ExtractionSignature
            Dim preparation As CooperativePreparation = Await PrepareCooperativeAsync(archive, binding, item, document, queue, options, result, cancellationToken).ConfigureAwait(False)
            document = preparation.Document
            BindDocumentToSource(document, item, sourcePath, current)
            Using cooperative As SemanticArchiveCooperativeStore = preparation.Store
            Dim canReuseExtraction As Boolean = Not item.ForceExtractionRebuild AndAlso ValidateReusableRepresentation(binding, document.Representation, item.SourceHash, item.ExtractionSignature)
            If Not canReuseExtraction Then
                document.Representation = Nothing
                document.Index = Nothing
                document.Card = Nothing
                Dim outputPath As String = CreateDerivedArtifactPath(binding, item.RelativePath, ".txt")
                Dim export As Global.SharedLibrary.Agents.TextExportResult
                Using authorizedPaths As System.IDisposable = CreateExportPathScope(binding, sourcePath, item.CanonicalSourceKey, outputPath)
                    export = Await Global.SharedLibrary.Agents.TextExportService.ExportFileAsync(
                        extractionContext, sourcePath, outputPath, exportOptions, cancellationToken).ConfigureAwait(False)
                End Using
                interactionScope.ThrowIfInteractionRequested()
                If export.ErrorCode = "requires_sta" Then Throw New SemanticArchiveHostRequiredException(export.Message)
                If export.Status = Global.SharedLibrary.Agents.TextExportStatus.Failed OrElse export.Status = Global.SharedLibrary.Agents.TextExportStatus.Unsupported Then
                    Throw New System.IO.InvalidDataException("text_export_failed (" & export.ErrorCode & "): " & If(System.String.IsNullOrWhiteSpace(export.Message), "The existing exporter could not produce a readable representation.", export.Message))
                End If
                If Not System.String.Equals(export.ProcessingSignature, expectedExporterSignature, System.StringComparison.Ordinal) Then
                    Throw New System.IO.IOException("The effective extraction configuration changed before the export was pinned; no representation was accepted under a stale processing signature.")
                End If
                If export.Status = Global.SharedLibrary.Agents.TextExportStatus.Empty Then
                    document.Active = False
                    document.ProcessingStatus = "empty"
                    document.Diagnostic = "The existing extractor returned no readable text; no document content was invented. " & DescribeExtraction(CaptureExtractionSourceMap(export))
                    document.SemanticSignature = item.SemanticSignature
                    document.SemanticModelIdentity = modelName
                    item.CachedDocument = document
                    queue.Save(item)
                    Return document
                End If
                If Not export.HasReadableText OrElse Not export.SourceAssociationVerified OrElse
                    Not System.String.Equals(export.SourceSha256, item.SourceHash, System.StringComparison.OrdinalIgnoreCase) Then
                    Throw New System.IO.InvalidDataException("The extraction source association is not verified or changed during extraction.")
                End If
                Dim textHash As String = SemanticArchiveIdentity.ComputeFileHash(export.OutputPath)
                If Not System.String.Equals(textHash, export.OutputSha256, System.StringComparison.OrdinalIgnoreCase) Then Throw New System.IO.InvalidDataException("The exported text-file hash does not match the complete published bytes.")
                document.Representation = New SemanticArchiveRepresentation() With {
                    .RepresentationId = SemanticArchiveIdentity.NewId(),
                    .SourceHash = export.SourceSha256,
                    .ExtractorVersion = export.ExtractorId & ":" & export.ExtractorVersion,
                    .OptionsSignature = item.ExtractionSignature,
                    .TextPath = export.OutputPath,
                    .TextFileHash = textHash,
                    .TextByteLength = New System.IO.FileInfo(export.OutputPath).Length,
                    .EncodingName = export.EncodingName,
                    .Completeness = If(export.ExtractionComplete.HasValue, If(export.ExtractionComplete.Value, "complete", "incomplete"), "unknown"),
                    .SourceMapJson = CaptureExtractionSourceMap(export)
                }
                document.Active = False
                document.ProcessingStatus = "extracted"
                item.ForceExtractionRebuild = False
                item.CachedDocument = document
                item.State = "Extracted"
                queue.Save(item)
                PublishCooperativeCheckpoint(cooperative, document, SemanticArchiveCooperativeStore.ExtractedStage, item.SemanticSignature, result, cancellationToken)
            Else
                result.ReusedFiles += 1
            End If
            Dim representation As SemanticArchiveRepresentation = document.Representation
            If representation.Completeness <> "complete" AndAlso Not archive.AllowPartialSearch Then
                document.Active = False
                document.Card = Nothing
                document.Index = Nothing
                document.SemanticSignature = item.SemanticSignature
                document.SemanticModelIdentity = modelName
                document.ProcessingStatus = representation.Completeness
                document.Diagnostic = "Extraction completeness is " & representation.Completeness & "; partial/unknown extraction search is disabled for this archive. The immutable extraction is retained for an explicit policy change or retry." & " " & DescribeExtraction(representation.SourceMapJson)
                item.CachedDocument = document
                queue.Save(item)
                Return document
            End If
            Dim text As String = System.IO.File.ReadAllText(representation.TextPath, New System.Text.UTF8Encoding(False, True))
            If System.String.IsNullOrWhiteSpace(text) Then
                document.Active = False
                document.Card = Nothing
                document.Index = Nothing
                document.ProcessingStatus = "empty"
                document.Diagnostic = "The exported representation contains no readable non-whitespace text."
                Return document
            End If
            Dim payloadHash As String = SemanticArchiveIdentity.HashBytes(New System.Text.UTF8Encoding(False, True).GetBytes(text))
            Dim needsIndex As System.Boolean = SemanticArchiveIndexPolicy.RequiresDocumentSectionIndex(current.Length, archive.SectionIndexThresholdBytes)
            Dim entries As System.Collections.Generic.List(Of SharedMethods.SemanticSearchIndexEntry) = Nothing
            If needsIndex Then
                Dim cached As SharedMethods.SemanticSearchIndexCacheItem = Nothing
                If Not item.ForceSemanticRebuild AndAlso document.Index IsNot Nothing AndAlso
                    document.Index.RepresentationId = representation.RepresentationId AndAlso
                    document.Index.ProfileVersion = archive.SemanticProfileVersion AndAlso
                    document.Index.ModelIdentity = item.SemanticSignature AndAlso
                    ValidateReusableIndex(binding, document.Index) Then
                    cached = Await SharedMethods.TryGetSemanticSearchIndexAsync(document.Index.Path, cancellationToken).ConfigureAwait(False)
                    If cached IsNot Nothing AndAlso cached.IndexDocument.ContentSha256 <> payloadHash Then cached = Nothing
                End If
                If cached Is Nothing Then
                    Dim indexPath As String = CreateDerivedArtifactPath(binding, item.RelativePath, ".indexed.txt")
                    Dim generated As SharedMethods.SemanticSearchIndexGenerationResult = Await SharedMethods.CreateSemanticSearchIndexedTextFileAsync(
                        representation.TextPath, indexPath, modelContext, SemanticArchiveMetadata.GeneratorOptions(), Nothing, cancellationToken).ConfigureAwait(False)
                    interactionScope.ThrowIfInteractionRequested()
                    If Not System.String.Equals(generated.ContentSha256, payloadHash, System.StringComparison.OrdinalIgnoreCase) Then Throw New System.IO.InvalidDataException("The section index does not describe the extracted text's BOM-free UTF-8 payload.")
                    document.Index = New SemanticArchiveIndexDescriptor() With {
                        .IndexId = SemanticArchiveIdentity.NewId(), .RepresentationId = representation.RepresentationId,
                        .Path = indexPath, .FileHash = SemanticArchiveIdentity.ComputeFileHash(indexPath), .PayloadHash = generated.ContentSha256,
                        .FormatVersion = SharedMethods.SemanticSearchCurrentFormatVersion,
                        .GeneratorVersion = SharedMethods.SemanticSearchDefaultGeneratorVersion,
                        .ProfileVersion = archive.SemanticProfileVersion,
                        .ModelIdentity = item.SemanticSignature,
                        .EntryCount = generated.SegmentCount
                    }
                    entries = generated.IndexDocument.Entries
                    document.Card = Nothing
                    item.ForceSemanticRebuild = False
                    item.CachedDocument = document
                    item.State = "Indexed"
                    queue.Save(item)
                    PublishCooperativeCheckpoint(cooperative, document, SemanticArchiveCooperativeStore.IndexedStage, item.SemanticSignature, result, cancellationToken)
                Else
                    entries = cached.IndexDocument.Entries
                End If
            Else
                document.Index = Nothing
            End If
            If item.ForceSemanticRebuild OrElse document.Card Is Nothing OrElse document.SemanticSignature <> item.SemanticSignature Then
                document.Card = Await SemanticArchiveMetadata.DescribeDocumentAsync(modelContext, document, text, entries, result.Diagnostics, cancellationToken).ConfigureAwait(False)
            End If
            interactionScope.ThrowIfInteractionRequested()
            If Not System.String.Equals(SemanticArchivePathGuard.GetVerifiedSourceIdentity(binding.RootPath, sourcePath), item.CanonicalSourceKey, System.StringComparison.Ordinal) Then
                Throw New System.InvalidOperationException("source_identity_changed: The original source mapping changed during processing; the personal document was not activated.")
            End If
            Dim afterProcessing As SemanticArchiveSourceFingerprint = ReadSourceFingerprint(binding.RootPath, sourcePath, cancellationToken)
            If Not System.String.Equals(afterProcessing.Sha256, document.Fingerprint.Sha256, System.StringComparison.OrdinalIgnoreCase) Then
                Throw New System.IO.IOException("The source changed while semantic metadata was being generated; its completed extraction remains checkpointed and the current source will be retried.")
            End If
            document.Fingerprint = afterProcessing
            document.SemanticSignature = item.SemanticSignature
            document.SemanticModelIdentity = modelName
            document.Active = True
            document.ProcessingStatus = If(representation.Completeness = "complete", "indexed", representation.Completeness)
            document.Diagnostic = If(representation.Completeness = "complete", "", "Searchable under the explicit partial-search policy; extraction completeness is " & representation.Completeness & ". " & DescribeExtraction(representation.SourceMapJson))
            item.ForceSemanticRebuild = False
            item.CachedDocument = document
            queue.Save(item)
            PublishCooperativeCheckpoint(cooperative, document, SemanticArchiveCooperativeStore.CompleteStage, item.SemanticSignature, result, cancellationToken)
            Return document
            End Using
        End Function

        Private Shared Function FormatProcessingFailure(failure As System.Exception) As System.String
            Dim message As System.String = failure.GetType().FullName & ": " & failure.Message
            Dim cause As System.Exception = failure.GetBaseException()
            If Not System.Object.ReferenceEquals(cause, failure) Then
                message &= " | Caused by " & cause.GetType().FullName & ": " & cause.Message
            End If
            Return BoundProcessingDiagnostic(message)
        End Function

        Private Shared Function BoundProcessingDiagnostic(message As System.String) As System.String
            If message Is Nothing Then Return ""
            Dim maximum As System.Int32 = SharedMethods.DEFAULT_SEMANTICARCHIVE_MAXIMUM_DIAGNOSTIC_CHARACTERS
            If message.Length <= maximum Then Return message
            Dim suffix As System.String = " [diagnostic truncated]"
            Dim length As System.Int32 = maximum - suffix.Length
            If System.Char.IsHighSurrogate(message(length - 1)) AndAlso System.Char.IsLowSurrogate(message(length)) Then length -= 1
            Return message.Substring(0, length) & suffix
        End Function

        Private Shared Function FailedRecord(item As SemanticArchiveWorkItem, status As String, diagnostic As String) As SemanticArchiveDocumentRecord
            Dim record As SemanticArchiveDocumentRecord = If(item.CachedDocument, New SemanticArchiveDocumentRecord())
            record.DocumentId = item.DocumentId
            record.SourceItemId = item.SourceItemId
            record.CanonicalSourceKey = item.CanonicalSourceKey
            record.SourcePath = item.SourcePath
            record.RelativePath = item.RelativePath
            record.DisplayName = System.IO.Path.GetFileName(item.SourcePath)
            record.PartitionKey = item.PartitionKey
            record.ExtractionSignature = item.ExtractionSignature
            record.BindingIds = New System.Collections.Generic.List(Of String)(item.BindingIds)
            record.Fingerprint = New SemanticArchiveSourceFingerprint() With {.Length = item.SourceLength, .LastWriteUtcTicks = item.SourceWriteTicks, .Sha256 = item.SourceHash}
            record.Active = False
            record.ProcessingStatus = status
            record.Diagnostic = diagnostic
            If record.Representation IsNot Nothing AndAlso record.Representation.SourceHash <> item.SourceHash Then
                record.Representation = Nothing
                record.Index = Nothing
                record.Card = Nothing
            End If
            Return record
        End Function

        Private Shared Function FindBinding(archive As SemanticArchiveDefinition, bindingId As String) As SemanticArchiveSourceBinding
            Return archive.Roots.Find(Function(binding As SemanticArchiveSourceBinding) System.String.Equals(binding.BindingId, bindingId, System.StringComparison.Ordinal))
        End Function

        Private Shared Function ValidateReusableRepresentation(binding As SemanticArchiveSourceBinding, representation As SemanticArchiveRepresentation, sourceHash As String, extractionSignature As String) As Boolean
            If representation Is Nothing OrElse representation.SourceHash <> sourceHash OrElse representation.OptionsSignature <> extractionSignature Then Return False
            Try
                SemanticArchiveStore.RequirePrivateDerivedArtifact(binding, representation.TextPath)
                Return ValidateFileHash(representation.TextPath, representation.TextFileHash)
            Catch ex As System.IO.IOException
                Return False
            Catch ex As System.UnauthorizedAccessException
                Return False
            End Try
        End Function

        Private Shared Function ValidateReusableIndex(binding As SemanticArchiveSourceBinding, descriptor As SemanticArchiveIndexDescriptor) As Boolean
            If descriptor Is Nothing Then Return False
            Try
                SemanticArchiveStore.RequirePrivateDerivedArtifact(binding, descriptor.Path)
                Return ValidateFileHash(descriptor.Path, descriptor.FileHash)
            Catch ex As System.IO.IOException
                Return False
            Catch ex As System.UnauthorizedAccessException
                Return False
            End Try
        End Function

        Private Shared Function ValidateFileHash(path As String, hash As String) As Boolean
            If System.String.IsNullOrWhiteSpace(path) OrElse System.String.IsNullOrWhiteSpace(hash) Then Return False
            Try
                Return System.String.Equals(SemanticArchiveIdentity.ComputeFileHash(path), hash, System.StringComparison.OrdinalIgnoreCase)
            Catch ex As System.IO.IOException
                Return False
            Catch ex As System.UnauthorizedAccessException
                Return False
            End Try
        End Function

        Private Shared Function CreateDerivedArtifactPath(binding As SemanticArchiveSourceBinding, relativeSourcePath As String, suffix As String) As String
            Dim fileName As System.String = "content" & suffix
            Dim requiredChildPathLength As System.Int32 = Global.SharedLibrary.Agents.TextExportService.GetRequiredOutputChildPathLength(relativeSourcePath, fileName)
            If suffix = ".indexed.txt" Then
                ' The existing semantic generator appends its GUID temporary suffix to
                ' the full indexed filename. Reserve that sibling before choosing a root.
                requiredChildPathLength = System.Math.Max(requiredChildPathLength, 1 + fileName.Length + 1 + System.Guid.Empty.ToString("N").Length + ".tmp".Length)
            End If
            Dim versionDirectory As System.String = SemanticArchiveArtifactPlanner.CreatePrivateVersionDirectory(binding,
                relativeSourcePath, SemanticArchiveIdentity.NewId(), requiredChildPathLength)
            Return SemanticArchivePathGuard.RequireWindowsCompatiblePath(System.IO.Path.Combine(versionDirectory, fileName))
        End Function

        Private Shared Function CreateExportPathScope(binding As SemanticArchiveSourceBinding, sourcePath As System.String,
                                                       sourceIdentity As System.String, outputPath As System.String) As System.IDisposable
            Dim source As System.String = SemanticArchivePathGuard.RequireWindowsSourcePath(sourcePath)
            Dim outputDirectory As System.String = System.IO.Path.GetDirectoryName(SemanticArchivePathGuard.RequireWindowsCompatiblePath(outputPath))
            If SemanticArchivePathGuard.IsContainedPath(outputDirectory, source) Then Throw New System.UnauthorizedAccessException("An extraction output directory cannot contain the original source.")
            SemanticArchiveStore.RequirePrivateDerivedArtifact(binding, outputDirectory)
            SemanticArchiveStore.RequirePrivateArtifact(outputDirectory)
            If Not System.String.Equals(SemanticArchivePathGuard.GetVerifiedSourceIdentity(binding.RootPath, source), sourceIdentity, System.StringComparison.OrdinalIgnoreCase) Then
                Throw New System.UnauthorizedAccessException("The original source identity changed before its export capability was created.")
            End If
            Return Global.SharedLibrary.Agents.PathPolicy.BeginHostFileOperationScope(
                Function(requestedPath As System.String, access As Global.SharedLibrary.Agents.PathAccess) As System.Boolean
                    Dim full As System.String = SemanticArchivePathGuard.CanonicalPath(requestedPath)
                    If System.String.Equals(full, source, System.StringComparison.OrdinalIgnoreCase) Then
                        If access <> Global.SharedLibrary.Agents.PathAccess.Read Then Return False
                        Return System.String.Equals(SemanticArchivePathGuard.GetVerifiedSourceIdentity(binding.RootPath, full), sourceIdentity, System.StringComparison.OrdinalIgnoreCase)
                    End If
                    If access <> Global.SharedLibrary.Agents.PathAccess.Read AndAlso access <> Global.SharedLibrary.Agents.PathAccess.Write Then Return False
                    If Not SemanticArchivePathGuard.IsContainedPath(outputDirectory, full) Then Return False
                    SemanticArchivePathGuard.RequireWindowsCompatiblePath(full)
                    SemanticArchivePathGuard.ValidateContainedPath(outputDirectory, full, False)
                    SemanticArchiveStore.RequirePrivateArtifact(outputDirectory)
                    If System.IO.File.Exists(full) OrElse System.IO.Directory.Exists(full) Then SemanticArchiveStore.RequirePrivateArtifact(full)
                    Return True
                End Function)
        End Function

        Private Shared Function ReadSourceFingerprint(root As String, path As String, cancellationToken As System.Threading.CancellationToken) As SemanticArchiveSourceFingerprint
            cancellationToken.ThrowIfCancellationRequested()
            Dim before As New System.IO.FileInfo(path)
            Dim beforeLength As Long = before.Length
            Dim beforeWrite As Long = before.LastWriteTimeUtc.Ticks
            Dim hash As String
            Using stream As System.IO.FileStream = SemanticArchivePathGuard.OpenContainedRead(root, path)
                Using hasher As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                    Dim buffer(81919) As Byte
                    Do
                        cancellationToken.ThrowIfCancellationRequested()
                        Dim count As Integer = stream.Read(buffer, 0, buffer.Length)
                        If count = 0 Then Exit Do
                        hasher.TransformBlock(buffer, 0, count, buffer, 0)
                    Loop
                    hasher.TransformFinalBlock(New Byte() {}, 0, 0)
                    hash = System.BitConverter.ToString(hasher.Hash).Replace("-", "").ToLowerInvariant()
                End Using
            End Using
            Dim after As New System.IO.FileInfo(path)
            If beforeLength <> after.Length OrElse beforeWrite <> after.LastWriteTimeUtc.Ticks Then Throw New System.IO.IOException("The source changed during fingerprinting; retry is required.")
            Return New SemanticArchiveSourceFingerprint() With {.Length = beforeLength, .LastWriteUtcTicks = beforeWrite, .Sha256 = hash}
        End Function

        Private Shared Sub Report(progress As System.IProgress(Of SemanticArchiveBuildProgress), result As SemanticArchiveBuildResult, path As String, stage As String, message As String)
            If progress IsNot Nothing Then progress.Report(New SemanticArchiveBuildProgress() With {
                .ArchiveId = result.ArchiveId, .SourcePath = path, .Stage = stage,
                .CompletedFiles = result.ProcessedFiles, .PendingFiles = result.PendingFiles, .DeferredFiles = result.DeferredFiles, .Message = message
            })
        End Sub

        Private NotInheritable Class SemanticArchiveHostRequiredException
            Inherits System.InvalidOperationException
            Public Sub New(message As String)
                MyBase.New(If(System.String.IsNullOrWhiteSpace(message), "This source requires an appropriate Office host reader.", message))
            End Sub
        End Class
    End Class
End Namespace
