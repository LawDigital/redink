' Part of "Red Ink" (SharedLibrary)
' Bounded resumable source discovery and independent permission reconciliation.

' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: SemanticArchiveBuilder.Scan.vb
' Purpose:
'   Bounded source discovery, scan checkpoints and eligibility/coverage reconciliation.
'
' Architecture / Function:
'   Advances resumable directory work and excludes generated outputs while retaining
'   explicit source-access failures.
' =============================================================================

Option Strict On
Option Explicit On
Option Infer On

Namespace SharedLibrary
    Public NotInheritable Partial Class SemanticArchiveBuilder
        Private NotInheritable Class LiveDirectoryCursor
            Implements System.IDisposable
            Public Property Key As System.String
            Public Property Position As System.Int64
            Public Property Iterator As System.Collections.Generic.IEnumerator(Of System.String)
            Public Sub Dispose() Implements System.IDisposable.Dispose
                If Iterator IsNot Nothing Then Iterator.Dispose()
                Iterator = Nothing
            End Sub
        End Class

        Private Shared ReadOnly DiscoveryCursorGate As New System.Object()
        Private Shared ReadOnly DiscoveryCursors As New System.Collections.Generic.Dictionary(Of System.String, LiveDirectoryCursor)(System.StringComparer.Ordinal)

        Public Shared Function ReadDiagnostics(store As SemanticArchiveStore, archiveId As System.String,
                        Optional includeTechnical As System.Boolean = SharedMethods.DEFAULT_SEMANTICARCHIVE_SHOW_TECHNICAL_DIAGNOSTICS) As System.Collections.Generic.List(Of System.String)
            If store Is Nothing Then Throw New System.ArgumentNullException(NameOf(store))
            SemanticArchiveIdentity.ValidateId(archiveId, NameOf(archiveId))
            Dim values As New System.Collections.Generic.List(Of System.String)()
            values.AddRange(ReadOperationDiagnostics(store, archiveId, includeTechnical))
            Dim workDirectory As System.String = store.GetWorkDirectory(archiveId)
            Dim state As SemanticArchiveQueueState = Nothing
            Try
                state = SemanticArchiveQueueIndex.ReadState(workDirectory)
            Catch ex As System.Exception
                values.Add("queue_state_unavailable: " & ex.GetType().FullName & ". The last recorded batch and other checkpoints remain separate.")
            End Try
            Dim archive As SemanticArchiveDefinition = store.GetArchive(archiveId)
            Dim access As SemanticArchiveAccessContext = SemanticArchiveAccessContext.CreateForCurrentUser()
            If state IsNot Nothing Then
                values.Add("Source jobs: " & state.TotalItems.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                    "; ready: " & state.ReadyItems.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                    "; deferred: " & state.DeferredItems.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                    "; host required: " & state.HostRequiredItems.ToString(System.Globalization.CultureInfo.InvariantCulture) & ".")
                If state.Dirty Then values.Add("Queue checkpoint recovery is pending before the next writer operation.")
                If state.DiscoveryPending Then values.Add("Source discovery has a durable continuation.")
                If state.PermissionsDiscoveryPending Then values.Add("Permission discovery or deferred artifact repair remains incomplete.")
                For Each id As System.String In state.ErrorDocumentIds
                    SemanticArchiveIdentity.ValidateId(id, "diagnostic document identity")
                    ' Only the bounded error index is opened; its large CachedDocument
                    ' payload is not read and this read-only path never opens a writer.
                    Try
                        Dim path As System.String = System.IO.Path.Combine(workDirectory, "items", id & ".json")
                        SemanticArchivePathGuard.ValidateContainedPath(workDirectory, path, True)
                        Dim item As SemanticArchiveWorkItem = SemanticArchiveQueueIndex.ReadHeader(path)
                        If Not System.String.Equals(item.DocumentId, id, System.StringComparison.Ordinal) Then Throw New System.IO.InvalidDataException("The diagnostic checkpoint identity does not match its filename.")
                        If CanDiscloseQueuedSource(item, archive, access) Then
                            values.Add("Pending source diagnostic: " & id & "; source: " & item.SourcePath &
                                "; state: " & item.State & "; attempts: " & item.Attempts.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                                "; next attempt UTC: " & item.NextAttemptUtc.ToString("O", System.Globalization.CultureInfo.InvariantCulture) &
                                "; " & If(System.String.IsNullOrWhiteSpace(item.LastError), "No failure detail was stored.", BoundProcessingDiagnostic(item.LastError)))
                        Else
                            values.Add("Pending source diagnostic: " & id & "; source access is unavailable. Cached paths and failure details remain hidden.")
                        End If
                    Catch ex As System.Exception
                        ' A concurrent writer may consume or replace this one checkpoint.
                        ' Do not lose every other persisted diagnostic because it changed.
                        values.Add("Pending source diagnostic: " & id & "; diagnostic_checkpoint_unavailable: " & ex.GetType().FullName & ". Refresh status to retry.")
                    End Try
                Next
            End If
            For Each name As System.String In New System.String() {"scan.json", "scan-permissions.json"}
                Try
                    Dim checkpoint As SemanticArchiveScanCheckpoint = SemanticArchiveWorkQueue.ReadPrivateControl(Of SemanticArchiveScanCheckpoint)(workDirectory, System.IO.Path.Combine(workDirectory, name), 4194304)
                    If checkpoint IsNot Nothing AndAlso checkpoint.Diagnostics IsNot Nothing Then
                        values.Add(name & ": filtered unsupported/disallowed files: " & checkpoint.FilteredFiles.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                            "; excluded entries/directories: " & checkpoint.ExcludedEntries.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                            ". Excluded directories are not traversed or counted file by file.")
                        For index As System.Int32 = 0 To System.Math.Min(99, checkpoint.Diagnostics.Count - 1)
                            values.Add(BoundProcessingDiagnostic(checkpoint.Diagnostics(index)))
                        Next
                    End If
                Catch ex As System.Exception
                    values.Add("scan_checkpoint_unavailable: " & name & "; " & ex.GetType().FullName & ". Other diagnostics remain available.")
                End Try
            Next
            Return values
        End Function

        Private Shared Function CanDiscloseQueuedSource(item As SemanticArchiveWorkItem, archive As SemanticArchiveDefinition,
                                                        access As SemanticArchiveAccessContext) As System.Boolean
            If item Is Nothing OrElse archive Is Nothing OrElse access Is Nothing OrElse item.BindingIds Is Nothing Then Return False
            For Each binding As SemanticArchiveSourceBinding In archive.Roots
                If Not item.BindingIds.Contains(binding.BindingId) Then Continue For
                Try
                    If SemanticArchiveStore.IsSourceExcluded(binding, item.SourcePath) OrElse
                       Not SemanticArchiveStore.IsSupportedSource(binding, item.SourcePath) Then Continue For
                    Dim identity As System.String = SemanticArchivePathGuard.GetVerifiedSourceIdentity(binding.RootPath, item.SourcePath)
                    If Not System.String.Equals(identity, item.CanonicalSourceKey, System.StringComparison.OrdinalIgnoreCase) Then Continue For
                    If access.CanReadSource(item.SourcePath) Then Return True
                Catch ex As System.Exception
                    ' Current original-source authorization is required even for an old error.
                End Try
            Next
            Return False
        End Function

        Private Shared Function ScanConfigurationSignature(archive As SemanticArchiveDefinition, semanticSignature As System.String) As System.String
            Return SemanticArchiveIdentity.StableId("scan", Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                .roots = archive.Roots, .extraction_profile = archive.ExtractionProfileVersion,
                .semantic_signature = semanticSignature, .partial_search = archive.AllowPartialSearch,
                .max_children = archive.MaxChildrenPerNode, .max_routing = archive.MaxRoutingCharacters}))
        End Function

        Private Sub ScanSources(archive As SemanticArchiveDefinition, previous As SemanticArchiveGenerationManifest,
                                lease As SemanticArchiveWriterLease, hierarchy As SemanticArchiveHierarchy,
                                queue As SemanticArchiveWorkQueue, options As SemanticArchiveBuildOptions,
                                extractionContext As SharedContext.ISharedContext,
                                semanticSignature As System.String, scanSignature As System.String,
                                progress As System.IProgress(Of SemanticArchiveBuildProgress), result As SemanticArchiveBuildResult,
                                cancellationToken As System.Threading.CancellationToken)
            If options.MaxDiscoveryEntries < 1 OrElse options.MaxDiscoveryEntries > 10000 OrElse options.MaxDiscoverySeconds < 1 OrElse options.MaxDiscoverySeconds > 120 Then
                Throw New System.ArgumentOutOfRangeException(NameOf(options), "Discovery batches permit 1..10000 entries and 1..120 seconds.")
            End If
            If options.SelectedDocumentIds IsNot Nothing AndAlso options.SelectedDocumentIds.Count = 0 Then
                result.SelectionRequired = True
                result.Diagnostics.Add("selection_required: No document was selected; an empty selected scope is never expanded to All.")
                Return
            End If
            Dim selected As System.Collections.Generic.List(Of System.String) = Nothing
            If options.SelectedDocumentIds IsNot Nothing Then
                selected = New System.Collections.Generic.List(Of System.String)(New System.Collections.Generic.HashSet(Of System.String)(options.SelectedDocumentIds, System.StringComparer.Ordinal))
            ElseIf options.IndexOnlyRebuild Then
                ' Reindex operates on the published logical inventory. It validates each current source
                ' during processing but deliberately does not discover new files or turn a semantic
                ' rebuild into a refresh/extraction operation.
                selected = New System.Collections.Generic.List(Of System.String)()
                If previous IsNot Nothing Then
                    For Each document As SemanticArchiveDocumentRecord In _store.EnumerateDocuments(previous)
                        If document IsNot Nothing AndAlso Not System.String.IsNullOrWhiteSpace(document.DocumentId) Then selected.Add(document.DocumentId)
                    Next
                End If
                result.Diagnostics.Add("reindex_inventory_bound: semantic rebuild is bound to the currently published logical source inventory; Refresh archive discovers additions/removals.")
            End If
            If selected IsNot Nothing Then
                selected = New System.Collections.Generic.List(Of System.String)(New System.Collections.Generic.HashSet(Of System.String)(selected, System.StringComparer.Ordinal))
                selected.Sort(System.StringComparer.Ordinal)
                For Each documentId As System.String In selected
                    SemanticArchiveIdentity.ValidateId(documentId, NameOf(options.SelectedDocumentIds))
                Next
            End If
            If options.OperationId Is Nothing Then options.OperationId = ""
            Dim scopeSignature As System.String = SemanticArchiveIdentity.StableId("scope", Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                .selected_ids = selected, .permissions_only = options.ReconcilePermissionsOnly,
                .force_extraction = options.ForceReextract, .force_semantic = options.RebuildSemanticMetadata, .index_only = options.IndexOnlyRebuild, .retry = options.RetryFailures,
                .integrity_audit = options.FullIntegrityAudit, .explicit_refresh = options.ForceScan AndAlso Not options.IsBackground}))
            Dim checkpoint As SemanticArchiveScanCheckpoint = queue.LoadScan(options.ReconcilePermissionsOnly)
            Dim timer As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew()
            Dim sameRequest As System.Boolean = options.OperationId.Length > 0 AndAlso checkpoint.RequestId = options.OperationId
            If sameRequest AndAlso checkpoint.ScopeSignature <> scopeSignature Then Throw New System.InvalidOperationException("A resumed archive command cannot change its selected scope. Start a new operation ID.")
            If sameRequest AndAlso Not checkpoint.InProgress AndAlso checkpoint.ConfigurationSignature = scanSignature Then
                If options.ReconcilePermissionsOnly Then
                    ProcessPermissionRetries(archive, previous, lease, queue, checkpoint, options, selected, result, timer, cancellationToken)
                    queue.SaveScan(checkpoint, True)
                    result.PermissionsPending = queue.HasPermissionRetries(selected)
                    If result.PermissionRetryReady Then result.PermissionsDeferred = False
                End If
                result.PendingFiles = queue.PendingCount(True, options.IsBackground, selected)
                Return
            End If
            If Not checkpoint.InProgress OrElse checkpoint.ConfigurationSignature <> scanSignature OrElse checkpoint.ScopeSignature <> scopeSignature OrElse
                (options.OperationId.Length > 0 AndAlso checkpoint.RequestId <> options.OperationId) Then
                checkpoint = New SemanticArchiveScanCheckpoint With {
                    .InProgress = True, .CycleId = System.Guid.NewGuid().ToString("N"), .RequestId = options.OperationId,
                    .ScopeSignature = scopeSignature, .ConfigurationSignature = scanSignature,
                    .ForceSemanticRebuild = options.RebuildSemanticMetadata OrElse options.IndexOnlyRebuild, .ForceExtractionRebuild = options.ForceReextract, .IndexOnlyRebuild = options.IndexOnlyRebuild,
                    .RetryFailures = options.RetryFailures, .ExplicitRefresh = options.ForceScan AndAlso Not options.IsBackground,
                    .Phase = If(selected Is Nothing, "discover", "selected")}
                queue.SaveDiscoveryBase(previous, options.ReconcilePermissionsOnly)
                If selected Is Nothing Then
                    Dim bindings As New System.Collections.Generic.List(Of SemanticArchiveSourceBinding)(archive.Roots)
                    bindings.Sort(Function(left, right) System.StringComparer.Ordinal.Compare(left.BindingId, right.BindingId))
                    For Each binding As SemanticArchiveSourceBinding In bindings
                        Try
                            queue.EnqueueDirectory(checkpoint, binding.BindingId, binding.RootPath, options.ReconcilePermissionsOnly)
                        Catch ex As System.Exception
                            AddScanFailure(checkpoint, binding.BindingId, "source_root_unavailable", ex)
                        End Try
                    Next
                End If
                queue.SaveScan(checkpoint, options.ReconcilePermissionsOnly)
            End If
            checkpoint.LastAttemptUtc = System.DateTime.UtcNow
            Dim initialDiagnostics As System.Int32 = checkpoint.Diagnostics.Count
            Report(progress, result, "", If(options.ReconcilePermissionsOnly, "permissions", "scanning"), "Resuming a bounded source/permission discovery checkpoint.")
            Try
                If options.ReconcilePermissionsOnly Then ProcessPermissionRetries(archive, previous, lease, queue, checkpoint, options, selected, result, timer, cancellationToken)
                While checkpoint.InProgress AndAlso result.DiscoveryEntriesInspected < options.MaxDiscoveryEntries AndAlso
                    timer.Elapsed.TotalSeconds < options.MaxDiscoverySeconds
                    cancellationToken.ThrowIfCancellationRequested()
                    If Not options.ReconcilePermissionsOnly AndAlso queue.PendingCount(False, options.IsBackground, selected) >= System.Math.Max(1024, options.MaximumFilesPerBatch * 4) Then
                        result.Diagnostics.Add("discovery_backpressure: Discovery paused while its bounded source-processing queue drains.")
                        Exit While
                    End If
                    Select Case checkpoint.Phase
                        Case "selected"
                            If checkpoint.SelectedPosition >= selected.Count Then
                                checkpoint.InProgress = False
                                Exit While
                            End If
                            Dim id As System.String = selected(checkpoint.SelectedPosition)
                            Dim document As SemanticArchiveDocumentRecord = If(previous Is Nothing, Nothing, _store.LoadDocument(previous, id))
                            If document Is Nothing Then
                                result.Diagnostics.Add("selected_document_unavailable: " & id & "; the selection was not expanded.")
                                checkpoint.SelectedPosition += 1
                                result.DiscoveryEntriesInspected += 1
                                Continue While
                            End If
                            Dim binding As SemanticArchiveSourceBinding = BindingForDocument(archive, document)
                            If binding Is Nothing Then
                                If Not options.ReconcilePermissionsOnly Then QueueRemoval(document, lease, queue, semanticSignature)
                                checkpoint.SelectedPosition += 1
                                result.DiscoveryEntriesInspected += 1
                                Continue While
                            End If
                            result.DiscoveryEntriesInspected += 1
                            If ProcessDiscoveredSource(archive, binding, document.SourcePath, previous, lease, queue, checkpoint, options, extractionContext, semanticSignature, result, cancellationToken, document.DocumentId) Then
                                checkpoint.SelectedPosition += 1
                            Else
                                Exit While
                            End If
                        Case "discover"
                            If checkpoint.DirectoryHead >= checkpoint.DirectoryTail Then
                                queue.SaveDiscoveryBase(previous, options.ReconcilePermissionsOnly)
                                checkpoint.Phase = "reconcile"
                                Continue While
                            End If
                            If Not DiscoverDirectoryStep(archive, previous, lease, queue, checkpoint, options, extractionContext, semanticSignature, result, cancellationToken) Then Exit While
                        Case "reconcile"
                            If Not ReconcileDiscoveryStep(archive, previous, lease, queue, checkpoint, options, extractionContext, semanticSignature, result, cancellationToken) Then
                                checkpoint.Phase = "queue"
                                checkpoint.QueueAfterId = ""
                            End If
                        Case "queue"
                            If options.ReconcilePermissionsOnly OrElse Not ReconcileQueuedStep(archive, lease, queue, checkpoint, options, result, cancellationToken) Then checkpoint.InProgress = False
                        Case Else
                            Throw New System.IO.InvalidDataException("The discovery phase is invalid.")
                    End Select
                End While
                If Not checkpoint.InProgress Then
                    checkpoint.LastCompletedUtc = If(checkpoint.IncompleteBindingIds.Count = 0, System.DateTime.UtcNow, System.DateTime.MinValue)
                    If checkpoint.IncompleteBindingIds.Count > 0 Then result.Diagnostics.Add("discovery_incomplete: Unavailable or changing directories remain unknown. They were not interpreted as source deletion; a later refresh retries them.")
                End If
            Finally
                queue.SaveScan(checkpoint, options.ReconcilePermissionsOnly)
                result.DiscoveryPending = checkpoint.InProgress
                result.PermissionsPending = options.ReconcilePermissionsOnly AndAlso (checkpoint.InProgress OrElse queue.HasPermissionRetries(selected))
                If result.PermissionRetryReady Then result.PermissionsDeferred = False
                result.PendingFiles = queue.PendingCount(True, options.IsBackground, selected)
                For index As System.Int32 = initialDiagnostics To checkpoint.Diagnostics.Count - 1
                    result.Diagnostics.Add(checkpoint.Diagnostics(index))
                Next
            End Try
        End Sub

        Private Function DiscoverDirectoryStep(archive As SemanticArchiveDefinition, previous As SemanticArchiveGenerationManifest,
                            lease As SemanticArchiveWriterLease, queue As SemanticArchiveWorkQueue, checkpoint As SemanticArchiveScanCheckpoint,
                            options As SemanticArchiveBuildOptions, extractionContext As SharedContext.ISharedContext, semanticSignature As System.String,
                            result As SemanticArchiveBuildResult, cancellationToken As System.Threading.CancellationToken) As System.Boolean
            Dim work As SemanticArchiveDirectoryWork = queue.LoadDirectory(checkpoint, options.ReconcilePermissionsOnly)
            Dim binding As SemanticArchiveSourceBinding = FindBinding(archive, work.BindingId)
            Dim cursorKey As System.String = queue.DirectoryPath & "|" & options.ReconcilePermissionsOnly.ToString()
            If binding Is Nothing Then
                FinishDirectory(checkpoint, cursorKey)
                Return True
            End If
            Try
                SemanticArchivePathGuard.RequireWindowsSourcePath(work.SourcePath)
                SemanticArchivePathGuard.ValidateContainedPath(binding.RootPath, work.SourcePath, True)
                If IsExcluded(binding, work.SourcePath) Then
                    If System.String.Equals(work.SourcePath, binding.RootPath, System.StringComparison.OrdinalIgnoreCase) Then Throw New System.IO.InvalidDataException("A source root is a generated artifact namespace.")
                    FinishDirectory(checkpoint, cursorKey)
                    Return True
                End If
                Dim writeTicks As System.Int64 = System.IO.Directory.GetLastWriteTimeUtc(work.SourcePath).Ticks
                If checkpoint.DirectoryWriteTicks = 0 Then checkpoint.DirectoryWriteTicks = writeTicks
                If writeTicks <> checkpoint.DirectoryWriteTicks AndAlso Not checkpoint.IncompleteBindingIds.Contains(binding.BindingId) Then
                    checkpoint.IncompleteBindingIds.Add(binding.BindingId)
                    checkpoint.Diagnostics.Add("directory_changed_during_discovery: Its source coverage remains unknown until a later complete pass.")
                End If
                Dim cursor As LiveDirectoryCursor = GetDirectoryCursor(cursorKey, checkpoint, work)
                ' After a process restart a filesystem enumerator has no portable seek
                ' cookie. Re-seeking is itself metered and resumes across worker ticks.
                If checkpoint.PendingEntryPath.Length = 0 AndAlso cursor.Position < checkpoint.DirectoryOffset Then
                    If cursor.Position = 0 AndAlso checkpoint.DirectoryOffset >= options.MaxDiscoveryEntries Then result.Diagnostics.Add("directory_resume_requires_persistent_worker: This large directory needs metered enumerator replay after restart; use the persistent worker loop so repeated short one-shot processes can reach new entries.")
                    cancellationToken.ThrowIfCancellationRequested()
                    result.DiscoveryEntriesInspected += 1
                    If Not cursor.Iterator.MoveNext() Then
                        AddScanFailure(checkpoint, binding.BindingId, "directory_resume_changed", Nothing)
                        FinishDirectory(checkpoint, cursorKey)
                    Else
                        cursor.Position += 1L
                    End If
                    Return True
                End If
                If checkpoint.PendingEntryPath.Length = 0 Then
                    cancellationToken.ThrowIfCancellationRequested()
                    If Not cursor.Iterator.MoveNext() Then
                        If System.IO.Directory.GetLastWriteTimeUtc(work.SourcePath).Ticks <> checkpoint.DirectoryWriteTicks Then AddScanFailure(checkpoint, binding.BindingId, "directory_changed_during_discovery", Nothing)
                        FinishDirectory(checkpoint, cursorKey)
                        Return True
                    End If
                    cursor.Position += 1L
                    checkpoint.PendingEntryPath = cursor.Iterator.Current
                End If
                Dim path As System.String = checkpoint.PendingEntryPath
                result.DiscoveryEntriesInspected += 1
                cancellationToken.ThrowIfCancellationRequested()
                SemanticArchivePathGuard.RequireWindowsSourcePath(path)
                If Not IsExcluded(binding, path) Then
                    Dim attributes As System.IO.FileAttributes = System.IO.File.GetAttributes(path)
                    If (attributes And System.IO.FileAttributes.ReparsePoint) <> 0 Then
                        AddScanFailure(checkpoint, binding.BindingId, "reparse_path_excluded", Nothing)
                    ElseIf (attributes And System.IO.FileAttributes.Directory) <> 0 Then
                        If binding.Recursive Then queue.EnqueueDirectory(checkpoint, binding.BindingId, path, options.ReconcilePermissionsOnly)
                    ElseIf SemanticArchiveStore.IsSupportedSource(binding, path) Then
                        If Not ProcessDiscoveredSource(archive, binding, path, previous, lease, queue, checkpoint, options, extractionContext, semanticSignature, result, cancellationToken) Then Return False
                    Else
                        checkpoint.FilteredFiles += 1L
                    End If
                Else
                    checkpoint.ExcludedEntries += 1L
                End If
                checkpoint.PendingEntryPath = ""
                checkpoint.PermissionContinuationToken = ""
                checkpoint.DirectoryOffset += 1L
                Return True
            Catch ex As System.OperationCanceledException When cancellationToken.IsCancellationRequested
                Throw
            Catch ex As System.Exception
                AddScanFailure(checkpoint, binding.BindingId, "scan_path_unavailable", ex)
                ' Failure of the current entry does not discard the directory's remaining
                ' entries. A failure opening/enumerating the directory retires it as unknown.
                If checkpoint.PendingEntryPath.Length > 0 Then
                    checkpoint.PendingEntryPath = ""
                    checkpoint.PermissionContinuationToken = ""
                    checkpoint.DirectoryOffset += 1L
                Else
                    FinishDirectory(checkpoint, cursorKey)
                End If
                Return True
            End Try
        End Function

        Private Function ProcessDiscoveredSource(archive As SemanticArchiveDefinition, binding As SemanticArchiveSourceBinding, sourcePath As System.String,
                            previous As SemanticArchiveGenerationManifest, lease As SemanticArchiveWriterLease, queue As SemanticArchiveWorkQueue,
                            checkpoint As SemanticArchiveScanCheckpoint, options As SemanticArchiveBuildOptions, extractionContext As SharedContext.ISharedContext,
                            semanticSignature As System.String, result As SemanticArchiveBuildResult, cancellationToken As System.Threading.CancellationToken, Optional knownDocumentId As System.String = Nothing) As System.Boolean
            Dim full As System.String = SemanticArchivePathGuard.RequireWindowsSourcePath(sourcePath)
            Dim sourceKey As System.String = ""
            Dim declaredPolicy As System.String = SemanticArchiveIdentity.StableId("policy", Newtonsoft.Json.JsonConvert.SerializeObject(New With {.ocr = binding.EnableOcr, .options = binding.ExtractionOptionsSignature}))
            Dim documentId As System.String = ""
            If options.ReconcilePermissionsOnly Then
                Try
                    sourceKey = SemanticArchivePathGuard.GetVerifiedSourceIdentityForMaintenance(binding.RootPath, full)
                    documentId = SemanticArchiveIdentity.StableId("doc", sourceKey)
                Catch ex As System.Exception
                    ' A denied data-read must not prevent restrictive artifact repair.
                    ' This private retry identity grants no source-content authority.
                    documentId = SemanticArchiveIdentity.StableId("permission", binding.BindingId & "|" & full)
                End Try
                If knownDocumentId IsNot Nothing Then documentId = knownDocumentId
            Else
                full = SemanticArchivePathGuard.ValidateContainedPath(binding.RootPath, full, True)
                sourceKey = SemanticArchivePathGuard.GetVerifiedSourceIdentity(binding.RootPath, full)
                documentId = SemanticArchiveIdentity.StableId("doc", sourceKey)
                If knownDocumentId IsNot Nothing AndAlso documentId <> knownDocumentId Then
                    Dim superseded As SemanticArchiveDocumentRecord = If(previous Is Nothing, Nothing, _store.LoadDocument(previous, knownDocumentId))
                    If superseded IsNot Nothing AndAlso superseded.ProcessingStatus <> "removed" Then QueueRemoval(superseded, lease, queue, semanticSignature)
                    result.Diagnostics.Add("selected_source_identity_changed: " & knownDocumentId & "; the superseded source identity was queued for removal. The verified source will use its stable source identity on this refresh.")
                    Return True
                End If
            End If
            Dim prior As SemanticArchiveDocumentRecord = If(previous Is Nothing, Nothing, _store.LoadDocument(previous, documentId))
            Dim seen As SemanticArchiveDiscoverySeen = queue.ReadSeen(checkpoint, documentId, options.ReconcilePermissionsOnly)
            If seen IsNot Nothing AndAlso seen.BindingIds.Contains(binding.BindingId) Then Return True
            Dim bindings As New System.Collections.Generic.List(Of System.String)()
            If seen IsNot Nothing Then bindings.AddRange(seen.BindingIds)
            If Not bindings.Contains(binding.BindingId) Then bindings.Add(binding.BindingId)
            If options.ReconcilePermissionsOnly Then
                Dim existing As SemanticArchivePermissionRetryRecord = queue.FindPermissionRetry(documentId)
                If existing Is Nothing Then
                    Dim repairItem As New SemanticArchiveWorkItem With {.DocumentId = documentId, .SourcePath = full,
                        .PrimaryBindingId = binding.BindingId, .BindingIds = bindings, .State = "PermissionRetry"}
                    ExecutePermissionRepair(archive, prior, binding, lease, queue, repairItem, options, result, cancellationToken)
                ElseIf existing.Item.NextAttemptUtc > System.DateTime.UtcNow Then
                    result.PermissionsDeferred = True
                End If
                ' An attempted source is discoverable even when its repair is pending.
                ' Its independently durable retry record prevents a false completion.
                queue.MarkSeen(checkpoint, documentId, bindings, True)
                Return True
            End If
            Dim info As New System.IO.FileInfo(full)
            Dim relative As System.String = RelativeSourcePath(binding.RootPath, full)
            Dim item As New SemanticArchiveWorkItem With {
                .DocumentId = documentId, .SourceItemId = SemanticArchiveIdentity.StableId("source", sourceKey), .CanonicalSourceKey = sourceKey,
                .SourcePath = full, .RelativePath = relative, .PrimaryBindingId = binding.BindingId, .BindingIds = bindings,
                .PartitionKey = binding.BindingId & ":" & NormalizeRelativeKey(relative) & "|" & documentId,
                .SourceLength = info.Length, .SourceWriteTicks = info.LastWriteTimeUtc.Ticks,
                .ExtractionSignature = ExtractionSignature(archive, binding, full, extractionContext), .SemanticSignature = semanticSignature,
                .ForceSemanticRebuild = checkpoint.ForceSemanticRebuild OrElse checkpoint.ForceExtractionRebuild,
                .ForceExtractionRebuild = checkpoint.ForceExtractionRebuild, .IndexOnlyRebuild = checkpoint.IndexOnlyRebuild}
            If prior IsNot Nothing AndAlso prior.PartitionKey.StartsWith(binding.BindingId & ":", System.StringComparison.Ordinal) Then item.PartitionKey = prior.PartitionKey
            If checkpoint.RetryFailures AndAlso RetryRequiresFreshExtraction(prior) Then
                ' Explicit retry is stage-aware: extraction-level failures must bypass
                ' an intact-but-incomplete representation, while semantic/index failures
                ' with a verified complete extract continue to reuse that extract.
                item.ForceExtractionRebuild = True
                item.ForceSemanticRebuild = True
            End If
            Dim unchanged As System.Boolean = prior IsNot Nothing AndAlso prior.Fingerprint IsNot Nothing AndAlso
                prior.Fingerprint.Length = item.SourceLength AndAlso prior.Fingerprint.LastWriteUtcTicks = item.SourceWriteTicks AndAlso
                Not System.String.IsNullOrWhiteSpace(prior.Fingerprint.Sha256)
            Try
                Dim audited As System.Boolean = options.FullIntegrityAudit OrElse AuditSelected(documentId)
                ' Discovery never streams original content or derived payloads. Each
                ' queued job fingerprints the current original before any reuse and
                ' validates immutable representations before restoring eligibility.
                ' Even an unchanged length/time is only a historical hint. Do not
                ' let an older published SHA replace a newer durable extraction.
                item.SourceHash = ""
                Dim validity As SemanticArchiveValidityRecord = If(prior Is Nothing, Nothing, _store.ReadValidity(archive.ArchiveId, documentId))
                Dim sharingRetry As System.Boolean = prior IsNot Nothing AndAlso checkpoint.ExplicitRefresh AndAlso
                    (prior.CooperativeState = "local_only" OrElse prior.CooperativeState = "contribution_pending")
                Dim changed As System.Boolean = prior Is Nothing OrElse Not unchanged OrElse
                    prior.ExtractionSignature <> item.ExtractionSignature OrElse prior.SemanticSignature <> semanticSignature OrElse
                    item.ForceSemanticRebuild OrElse item.ForceExtractionRebuild OrElse audited OrElse sharingRetry OrElse
                    Not SameStringSet(prior.BindingIds, bindings) OrElse prior.PartitionKey <> item.PartitionKey OrElse prior.SourcePath <> full OrElse
                    prior.ProcessingStatus = "removed" OrElse prior.ProcessingStatus = "unavailable" OrElse validity Is Nothing OrElse Not validity.Valid
                If changed Then
                    item.CachedDocument = prior
                    If prior IsNot Nothing AndAlso (Not unchanged OrElse audited) Then
                        _store.SaveValidity(lease, prior, False, "Source/version or scheduled integrity validation is pending; the prior projection remains suppressed until validation completes.", Nothing)
                    End If
                    queue.Enqueue(item, replaceOperationIntent:=checkpoint.ExplicitRefresh OrElse checkpoint.RetryFailures OrElse checkpoint.ForceSemanticRebuild OrElse checkpoint.ForceExtractionRebuild)
                End If
                If checkpoint.RetryFailures AndAlso seen Is Nothing Then queue.ResetRetry(documentId)
                queue.MarkSeen(checkpoint, documentId, bindings, False)
                Return True
            Catch ex As System.OperationCanceledException When cancellationToken.IsCancellationRequested
                Throw
            Catch ex As System.Exception
                If prior IsNot Nothing Then _store.SaveValidity(lease, prior, False, "Source validation is unavailable; deletion is not established.", Nothing)
                AddScanFailure(checkpoint, binding.BindingId, "source_validation_unknown", ex)
                result.FailedFiles += 1
                Return True
            End Try
        End Function

        Private Sub ExecutePermissionRepair(archive As SemanticArchiveDefinition, prior As SemanticArchiveDocumentRecord,
                            binding As SemanticArchiveSourceBinding, lease As SemanticArchiveWriterLease, queue As SemanticArchiveWorkQueue,
                            item As SemanticArchiveWorkItem, options As SemanticArchiveBuildOptions, result As SemanticArchiveBuildResult,
                            cancellationToken As System.Threading.CancellationToken, Optional maximumArtifactChecks As System.Int32 = 64)
            cancellationToken.ThrowIfCancellationRequested()
            result.PermissionSourcesChecked += 1
            Try
                SemanticArchivePathGuard.RequireWindowsSourcePath(item.SourcePath)
                If Not SemanticArchivePathGuard.IsContainedPath(binding.RootPath, item.SourcePath) OrElse IsExcluded(binding, item.SourcePath) OrElse
                    Not SemanticArchiveStore.IsSupportedSource(binding, item.SourcePath) Then Throw New System.IO.InvalidDataException("The pending permission source is outside its declared binding.")
                Dim repair As SemanticArchivePermissionRepairResult = SemanticArchiveArtifactPlanner.ReconcilePermissions(binding, item.SourcePath,
                    item.PermissionContinuationToken, System.Math.Max(1, System.Math.Min(maximumArtifactChecks, options.MaxDiscoveryEntries - result.DiscoveryEntriesInspected + 1)), cancellationToken)
                result.DiscoveryEntriesInspected += System.Math.Max(0, repair.CheckedArtifacts - 1)
                result.PermissionArtifactsChecked += repair.CheckedArtifacts
                result.PermissionArtifactsRepaired += repair.RepairedArtifacts
                result.PermissionArtifactsQuarantined += repair.QuarantinedArtifacts
                If Not System.String.IsNullOrWhiteSpace(repair.Diagnostic) Then result.Diagnostics.Add("permission_reconciliation: " & item.DocumentId & "; " & repair.Diagnostic)
                Select Case repair.Status
                    Case "complete", "private", "no_artifacts", "quarantined"
                        If prior IsNot Nothing Then ReconcileSourceEligibility(prior, binding, lease, result, cancellationToken)
                        queue.CompletePermissionRetry(item.DocumentId)
                        Return
                    Case "partial", "restart_required"
                        result.PermissionRetryReady = True
                        item.PermissionContinuationToken = If(repair.Status = "partial", If(repair.ContinuationToken, ""), "")
                        item.NextAttemptUtc = System.DateTime.MinValue
                        item.LastError = "permission_reconciliation_pending: " & repair.Status
                    Case Else
                        item.PermissionContinuationToken = ""
                        item.Attempts += 1
                        item.NextAttemptUtc = System.DateTime.UtcNow.AddSeconds(System.Math.Min(3600.0, 60.0 * System.Math.Pow(2.0, System.Math.Min(6, item.Attempts - 1))))
                        item.LastError = "permission_reconciliation_deferred: " & repair.Status
                        result.PermissionsDeferred = True
                        If prior IsNot Nothing Then _store.SaveValidity(lease, prior, False, "Generated artifact rights could not be fully reconciled; access remains suppressed.", Nothing)
                End Select
            Catch ex As System.OperationCanceledException When cancellationToken.IsCancellationRequested
                Throw
            Catch ex As System.Exception
                item.PermissionContinuationToken = ""
                item.Attempts += 1
                item.NextAttemptUtc = System.DateTime.UtcNow.AddSeconds(System.Math.Min(3600.0, 60.0 * System.Math.Pow(2.0, System.Math.Min(6, item.Attempts - 1))))
                item.LastError = "permission_reconciliation_unknown: " & ex.GetType().Name
                result.PermissionsDeferred = True
                result.Diagnostics.Add(item.LastError & "; this source remains queued while later sources can be checked.")
                If prior IsNot Nothing Then _store.SaveValidity(lease, prior, False, "Permission reconciliation is unknown; original source deletion was not inferred.", Nothing)
            End Try
            queue.EnqueuePermissionRetry(item)
        End Sub

        Private Sub ProcessPermissionRetries(archive As SemanticArchiveDefinition, previous As SemanticArchiveGenerationManifest,
                            lease As SemanticArchiveWriterLease, queue As SemanticArchiveWorkQueue, checkpoint As SemanticArchiveScanCheckpoint,
                            options As SemanticArchiveBuildOptions, selected As System.Collections.Generic.List(Of System.String),
                            result As SemanticArchiveBuildResult, timer As System.Diagnostics.Stopwatch, cancellationToken As System.Threading.CancellationToken)
            Dim retryBudget As System.Int32 = If(checkpoint.InProgress, options.MaxDiscoveryEntries \ 4, options.MaxDiscoveryEntries)
            If retryBudget < 1 Then Return
            Dim maximum As System.Int32 = System.Math.Min(32, retryBudget)
            If selected Is Nothing Then maximum = CInt(System.Math.Min(CLng(maximum), queue.PermissionRetrySlotCount()))
            If selected Is Nothing AndAlso checkpoint.PermissionRetryRoundEnd <= queue.PermissionRetryHead Then checkpoint.PermissionRetryRoundEnd = queue.PermissionRetryTail
            Dim initialEntries As System.Int32 = result.DiscoveryEntriesInspected
            Dim inspected As System.Int32 = 0
            While inspected < maximum AndAlso result.DiscoveryEntriesInspected - initialEntries < retryBudget AndAlso result.DiscoveryEntriesInspected < options.MaxDiscoveryEntries AndAlso timer.Elapsed.TotalSeconds < options.MaxDiscoverySeconds
                cancellationToken.ThrowIfCancellationRequested()
                Dim record As SemanticArchivePermissionRetryRecord = Nothing
                If selected IsNot Nothing Then
                    If selected.Count = 0 OrElse inspected >= selected.Count Then Exit While
                    checkpoint.PermissionRetrySelectionPosition = checkpoint.PermissionRetrySelectionPosition Mod selected.Count
                    record = queue.FindPermissionRetry(selected(checkpoint.PermissionRetrySelectionPosition))
                    checkpoint.PermissionRetrySelectionPosition += 1
                Else
                    If Not queue.HasPermissionRetries() OrElse queue.PermissionRetryHead >= checkpoint.PermissionRetryRoundEnd Then Exit While
                    record = queue.PeekPermissionRetry()
                End If
                inspected += 1
                result.DiscoveryEntriesInspected += 1
                If record IsNot Nothing Then
                    Dim item As SemanticArchiveWorkItem = record.Item
                    If item.NextAttemptUtc > System.DateTime.UtcNow Then
                        result.PermissionsDeferred = True
                        If selected Is Nothing Then queue.EnqueuePermissionRetry(item)
                    Else
                        Dim binding As SemanticArchiveSourceBinding = FindBinding(archive, item.PrimaryBindingId)
                        Dim prior As SemanticArchiveDocumentRecord = If(previous Is Nothing, Nothing, _store.LoadDocument(previous, item.DocumentId))
                        If binding Is Nothing Then
                            If prior IsNot Nothing Then _store.SaveValidity(lease, prior, False, "The permission source is outside this archive's current scope.", Nothing)
                            queue.CompletePermissionRetry(item.DocumentId)
                        Else
                            ExecutePermissionRepair(archive, prior, binding, lease, queue, item, options, result, cancellationToken, System.Math.Max(1, System.Math.Min(64, retryBudget - (result.DiscoveryEntriesInspected - initialEntries) + 1)))
                        End If
                    End If
                End If
                If selected Is Nothing Then queue.AdvancePermissionRetry()
            End While
            If selected Is Nothing Then
                If queue.PermissionRetryHead < checkpoint.PermissionRetryRoundEnd Then
                    ' Later slots have not yet been inspected in this durable round.
                    ' Continue bounded batches before concluding that only backoff remains.
                    result.PermissionsDeferred = False
                Else
                    checkpoint.PermissionRetryRoundEnd = 0
                End If
            ElseIf checkpoint.PermissionRetrySelectionPosition < selected.Count Then
                result.PermissionsDeferred = False
            Else
                checkpoint.PermissionRetrySelectionPosition = 0
            End If
        End Sub

        Private Sub ReconcileSourceEligibility(document As SemanticArchiveDocumentRecord, binding As SemanticArchiveSourceBinding,
                            lease As SemanticArchiveWriterLease, result As SemanticArchiveBuildResult,
                            cancellationToken As System.Threading.CancellationToken)
            cancellationToken.ThrowIfCancellationRequested()
            Dim access As SemanticArchiveAccessContext = SemanticArchiveAccessContext.CreateForCurrentUser()
            If Not access.CanReadSource(document.SourcePath) Then
                _store.SaveValidity(lease, document, False, "The local user's original-source access is denied or unknown.", Nothing)
                result.Diagnostics.Add("source_authorization_unknown: Original source access failed closed; no source ACL was changed.")
                Return
            End If
            Dim info As New System.IO.FileInfo(document.SourcePath)
            If info.Length <> document.Fingerprint.Length OrElse info.LastWriteTimeUtc.Ticks <> document.Fingerprint.LastWriteUtcTicks Then
                _store.SaveValidity(lease, document, False, "Source version changed; permission repair does not rebuild source content.", Nothing)
                Return
            End If
            Dim prior As SemanticArchiveValidityRecord = _store.ReadValidity(lease.ArchiveId, document.DocumentId)
            If prior IsNot Nothing AndAlso prior.Valid Then Return
            ' ACL reconciliation does not certify source bytes or perform an
            ' unbounded content hash. Leave the projection suppressed for normal
            ' content refresh, where the immutable source/version is verified.
            _store.SaveValidity(lease, document, False, "Content refresh is required to revalidate this previously suppressed private projection.", Nothing)
            result.Diagnostics.Add("content_refresh_required: " & document.DocumentId & "; artifact rights were checked, but content eligibility remains suppressed until source/version validation.")
        End Sub

        Private Function ReconcileDiscoveryStep(archive As SemanticArchiveDefinition, previous As SemanticArchiveGenerationManifest,
                            lease As SemanticArchiveWriterLease, queue As SemanticArchiveWorkQueue, checkpoint As SemanticArchiveScanCheckpoint,
                            options As SemanticArchiveBuildOptions, extractionContext As SharedContext.ISharedContext, semanticSignature As System.String,
                            result As SemanticArchiveBuildResult, cancellationToken As System.Threading.CancellationToken) As System.Boolean
            Dim basis As SemanticArchiveGenerationManifest = queue.LoadDiscoveryBase(archive.ArchiveId, options.ReconcilePermissionsOnly)
            If basis Is Nothing Then Return False
            While checkpoint.ReconcileShard < basis.DocumentShards.Count
                Dim shard As SemanticArchiveDocumentShard = _store.LoadDocumentShard(basis, basis.DocumentShards(checkpoint.ReconcileShard))
                If checkpoint.ReconcileDocument >= shard.Documents.Count Then
                    checkpoint.ReconcileShard += 1
                    checkpoint.ReconcileDocument = 0
                    Continue While
                End If
                cancellationToken.ThrowIfCancellationRequested()
                Dim document As SemanticArchiveDocumentRecord = shard.Documents(checkpoint.ReconcileDocument)
                result.DiscoveryEntriesInspected += 1
                If System.String.Equals(document.ProcessingStatus, "removed", System.StringComparison.Ordinal) Then
                    checkpoint.ReconcileDocument += 1
                    Return True
                End If
                Dim seen As SemanticArchiveDiscoverySeen = queue.ReadSeen(checkpoint, document.DocumentId, options.ReconcilePermissionsOnly)
                If seen IsNot Nothing Then
                    checkpoint.ReconcileDocument += 1
                    Return True
                End If
                Dim binding As SemanticArchiveSourceBinding = BindingForDocument(archive, document)
                If binding Is Nothing Then
                    _store.SaveValidity(lease, document, False, "The original source is no longer in this archive's declared scope.", Nothing)
                    If Not options.ReconcilePermissionsOnly AndAlso document.ProcessingStatus <> "removed" Then QueueRemoval(document, lease, queue, semanticSignature)
                    checkpoint.ReconcileDocument += 1
                    Return True
                End If
                Try
                    SemanticArchivePathGuard.RequireWindowsSourcePath(document.SourcePath)
                    SemanticArchivePathGuard.ValidateContainedPath(binding.RootPath, binding.RootPath, True)
                    Dim attributes As System.IO.FileAttributes = System.IO.File.GetAttributes(document.SourcePath)
                    If (attributes And System.IO.FileAttributes.ReparsePoint) <> 0 OrElse (attributes And System.IO.FileAttributes.Directory) <> 0 Then Throw New System.IO.IOException("The original file became an unsupported physical entry.")
                    If Not ProcessDiscoveredSource(archive, binding, document.SourcePath, previous, lease, queue, checkpoint, options, extractionContext, semanticSignature, result, cancellationToken, document.DocumentId) Then Return True
                Catch ex As System.OperationCanceledException When cancellationToken.IsCancellationRequested
                    Throw
                Catch ex As System.IO.FileNotFoundException
                    ReconcileMissingSource(document, binding, lease, queue, checkpoint, options, semanticSignature)
                Catch ex As System.IO.DirectoryNotFoundException
                    ReconcileMissingSource(document, binding, lease, queue, checkpoint, options, semanticSignature)
                Catch ex As System.Exception
                    _store.SaveValidity(lease, document, False, "Discovery/permissions are unavailable; original source deletion is not established.", Nothing)
                    AddScanFailure(checkpoint, binding.BindingId, "source_validity_unknown", ex)
                End Try
                checkpoint.ReconcileDocument += 1
                Return True
            End While
            Return False
        End Function

        Private Sub ReconcileMissingSource(document As SemanticArchiveDocumentRecord, binding As SemanticArchiveSourceBinding,
                            lease As SemanticArchiveWriterLease, queue As SemanticArchiveWorkQueue, checkpoint As SemanticArchiveScanCheckpoint,
                            options As SemanticArchiveBuildOptions, semanticSignature As System.String)
            _store.SaveValidity(lease, document, False, "The original source is missing or its discovery is incomplete.", Nothing)
            If Not options.ReconcilePermissionsOnly AndAlso Not checkpoint.IncompleteBindingIds.Contains(binding.BindingId) AndAlso document.ProcessingStatus <> "removed" Then
                QueueRemoval(document, lease, queue, semanticSignature)
            Else
                checkpoint.Diagnostics.Add("source_validity_unknown: Missing or inaccessible coverage does not establish source removal.")
            End If
        End Sub

        Private Sub QueueRemoval(document As SemanticArchiveDocumentRecord, lease As SemanticArchiveWriterLease, queue As SemanticArchiveWorkQueue, semanticSignature As System.String)
            Dim item As SemanticArchiveWorkItem = WorkFromDocument(document)
            item.Remove = True
            item.SemanticSignature = semanticSignature
            item.CachedDocument = document
            _store.SaveValidity(lease, document, False, "The original source was confirmed outside the current archive scope.", Nothing)
            queue.Enqueue(item)
        End Sub

        Private Function ReconcileQueuedStep(archive As SemanticArchiveDefinition, lease As SemanticArchiveWriterLease,
                            queue As SemanticArchiveWorkQueue, checkpoint As SemanticArchiveScanCheckpoint, options As SemanticArchiveBuildOptions,
                            result As SemanticArchiveBuildResult, cancellationToken As System.Threading.CancellationToken) As System.Boolean
            Dim ids As System.Collections.Generic.List(Of System.String) = queue.DocumentIdsAfter(checkpoint.QueueAfterId, 1)
            If ids.Count = 0 Then Return False
            cancellationToken.ThrowIfCancellationRequested()
            Dim item As SemanticArchiveWorkItem = queue.Load(ids(0))
            result.DiscoveryEntriesInspected += 1
            If item IsNot Nothing Then
                Dim binding As SemanticArchiveSourceBinding = FindBinding(archive, item.PrimaryBindingId)
                If binding Is Nothing OrElse Not SemanticArchivePathGuard.IsContainedPath(binding.RootPath, item.SourcePath) OrElse
                    SemanticArchiveStore.IsSourceExcluded(binding, item.SourcePath) OrElse Not SemanticArchiveStore.IsSupportedSource(binding, item.SourcePath) Then
                    If Not item.Remove Then queue.Complete(item)
                ElseIf checkpoint.RetryFailures Then
                    queue.ResetRetry(item.DocumentId)
                End If
            End If
            checkpoint.QueueAfterId = ids(0)
            Return True
        End Function

        Private Shared Function GetDirectoryCursor(key As System.String, checkpoint As SemanticArchiveScanCheckpoint, work As SemanticArchiveDirectoryWork) As LiveDirectoryCursor
            Dim identity As System.String = checkpoint.CycleId & "|" & work.Position.ToString(System.Globalization.CultureInfo.InvariantCulture) & "|" & work.SourcePath
            SyncLock DiscoveryCursorGate
                Dim cursor As LiveDirectoryCursor = Nothing
                If DiscoveryCursors.TryGetValue(key, cursor) Then
                    If cursor.Key = identity AndAlso (cursor.Position <= checkpoint.DirectoryOffset OrElse checkpoint.PendingEntryPath.Length > 0) Then Return cursor
                    cursor.Dispose()
                    DiscoveryCursors.Remove(key)
                End If
                If DiscoveryCursors.Count >= 16 Then
                    Dim stale As System.String = Nothing
                    For Each candidate As System.String In DiscoveryCursors.Keys
                        stale = candidate
                        Exit For
                    Next
                    If stale IsNot Nothing Then
                        DiscoveryCursors(stale).Dispose()
                        DiscoveryCursors.Remove(stale)
                    End If
                End If
                cursor = New LiveDirectoryCursor With {.Key = identity, .Iterator = System.IO.Directory.EnumerateFileSystemEntries(work.SourcePath).GetEnumerator()}
                DiscoveryCursors.Add(key, cursor)
                Return cursor
            End SyncLock
        End Function

        Private Shared Sub FinishDirectory(checkpoint As SemanticArchiveScanCheckpoint, cursorKey As System.String)
            SyncLock DiscoveryCursorGate
                Dim cursor As LiveDirectoryCursor = Nothing
                If DiscoveryCursors.TryGetValue(cursorKey, cursor) Then
                    cursor.Dispose()
                    DiscoveryCursors.Remove(cursorKey)
                End If
            End SyncLock
            checkpoint.DirectoryHead += 1L
            checkpoint.DirectoryOffset = 0
            checkpoint.DirectoryWriteTicks = 0
            checkpoint.PendingEntryPath = ""
            checkpoint.PermissionContinuationToken = ""
        End Sub

        Private Shared Sub AddScanFailure(checkpoint As SemanticArchiveScanCheckpoint, bindingId As System.String, code As System.String, failure As System.Exception)
            If Not checkpoint.IncompleteBindingIds.Contains(bindingId) Then checkpoint.IncompleteBindingIds.Add(bindingId)
            Dim message As System.String = code & ": Source coverage is unknown; no deletion was inferred."
            If failure IsNot Nothing Then message &= " " & failure.GetType().Name
            If checkpoint.Diagnostics.Count < 100 AndAlso Not checkpoint.Diagnostics.Contains(message) Then checkpoint.Diagnostics.Add(message)
        End Sub

        Private Shared Function BindingForDocument(archive As SemanticArchiveDefinition, document As SemanticArchiveDocumentRecord) As SemanticArchiveSourceBinding
            For Each bindingId As System.String In document.BindingIds
                Dim binding As SemanticArchiveSourceBinding = FindBinding(archive, bindingId)
                If binding IsNot Nothing AndAlso SemanticArchivePathGuard.IsContainedPath(binding.RootPath, document.SourcePath) AndAlso
                    Not SemanticArchiveStore.IsSourceExcluded(binding, document.SourcePath) AndAlso SemanticArchiveStore.IsSupportedSource(binding, document.SourcePath) Then Return binding
            Next
            Return Nothing
        End Function

        Private Function ExtractionSignature(archive As SemanticArchiveDefinition, binding As SemanticArchiveSourceBinding, sourcePath As System.String, extractionContext As SharedContext.ISharedContext) As System.String
            Dim signature As System.String = Global.SharedLibrary.Agents.TextExportService.GetProcessingSignature(extractionContext, sourcePath,
                New Global.SharedLibrary.Agents.TextExportOptions With {.OcrPdf = binding.EnableOcr, .OcrBatchPages = binding.OcrBatchPages, .Overwrite = False})
            Return ComposeExtractionSignature(archive, binding, signature)
        End Function

        Private Shared Function ComposeExtractionSignature(archive As SemanticArchiveDefinition, binding As SemanticArchiveSourceBinding, signature As System.String, Optional ocrBatchPages As System.Int32 = -1) As System.String
            ' Semantic compatibility is independent of private/shared placement and
            ' user-visible path aliases. Structured encoding prevents delimiter aliases.
            Return SemanticArchiveIdentity.StableId("extraction", Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                .schema = 2, .exporter = signature, .profile = archive.ExtractionProfileVersion,
                .options = binding.ExtractionOptionsSignature, .ocr = binding.EnableOcr, .ocrBatchPages = If(ocrBatchPages < 0, binding.OcrBatchPages, ocrBatchPages)}))
        End Function

        Private Shared Function RelativeSourcePath(root As System.String, path As System.String) As System.String
            Dim prefix As System.String = SemanticArchivePathGuard.CanonicalPath(root).TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar) & System.IO.Path.DirectorySeparatorChar
            Dim full As System.String = SemanticArchivePathGuard.CanonicalPath(path)
            If Not SemanticArchivePathGuard.IsContainedPath(root, full) OrElse full.Length <= prefix.Length Then Throw New System.IO.InvalidDataException("The source has no valid root-relative file name.")
            Return full.Substring(prefix.Length)
        End Function

        Private Shared Function NormalizeRelativeKey(relative As System.String) As System.String
            Dim value As System.String = relative.Replace(System.IO.Path.DirectorySeparatorChar, "/"c).Replace(System.IO.Path.AltDirectorySeparatorChar, "/"c)
            Return If(System.Environment.OSVersion.Platform = System.PlatformID.Win32NT, value.ToUpperInvariant(), value)
        End Function

        Private Shared Function IsExcluded(binding As SemanticArchiveSourceBinding, path As System.String) As System.Boolean
            Return SemanticArchiveStore.IsSourceExcluded(binding, path)
        End Function

        Private Shared Function RetryRequiresFreshExtraction(document As SemanticArchiveDocumentRecord) As System.Boolean
            If document Is Nothing OrElse document.Representation Is Nothing Then Return True
            Select Case If(document.ProcessingStatus, System.String.Empty)
                Case "empty", "incomplete", "unknown"
                    Return True
            End Select
            Return Not System.String.Equals(If(document.Representation.Completeness, System.String.Empty), "complete", System.StringComparison.Ordinal) AndAlso Not document.Active
        End Function

        Private Shared Function SameStringSet(left As System.Collections.Generic.IEnumerable(Of System.String), right As System.Collections.Generic.IEnumerable(Of System.String)) As System.Boolean
            Dim values As New System.Collections.Generic.HashSet(Of System.String)(left, System.StringComparer.Ordinal)
            Return values.SetEquals(right)
        End Function

        Private Shared Function AuditSelected(documentId As System.String) As System.Boolean
            Dim material As System.String = SemanticArchiveIdentity.StableId("audit", documentId & System.DateTime.UtcNow.ToString("yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture))
            Return System.Convert.ToInt32(material.Substring(material.Length - 2), 16) Mod 32 = 0
        End Function

        Private Shared Function WorkFromDocument(document As SemanticArchiveDocumentRecord) As SemanticArchiveWorkItem
            Return New SemanticArchiveWorkItem With {
                .DocumentId = document.DocumentId, .SourceItemId = document.SourceItemId, .CanonicalSourceKey = document.CanonicalSourceKey,
                .SourcePath = document.SourcePath, .RelativePath = document.RelativePath, .PartitionKey = document.PartitionKey,
                .PrimaryBindingId = If(document.BindingIds.Count > 0, document.BindingIds(0), ""),
                .BindingIds = New System.Collections.Generic.List(Of System.String)(document.BindingIds),
                .SourceLength = document.Fingerprint.Length, .SourceWriteTicks = document.Fingerprint.LastWriteUtcTicks,
                .SourceHash = document.Fingerprint.Sha256, .ExtractionSignature = document.ExtractionSignature,
                .SemanticSignature = document.SemanticSignature, .CachedDocument = document}
        End Function
    End Class
End Namespace
