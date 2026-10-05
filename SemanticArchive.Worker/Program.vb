' Part of "Red Ink" (Red Ink Semantic Archive Worker)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: Program.vb
' Purpose:
'   Headless console entry point for bounded/resumable Semantic Archive operations and
'   JSON lifecycle logging.
'
' Architecture / Function:
'   Loads normal configuration/licensing, invokes the shared builder and reports append-
'   only stderr progress with cancellation.
' =============================================================================

Option Strict On
Option Explicit On
Option Infer On

Imports SharedLibrary.SharedLibrary

Namespace SemanticArchiveWorker
    Public Module Program
        Public Function Main(arguments As System.String()) As System.Int32
            Dim options As WorkerOptions
            Try
                options = WorkerOptions.Parse(arguments)
            Catch ex As System.Exception
                WriteConsoleEvent("invalid_arguments", New With {.exceptionType = ex.GetType().Name})
                System.Console.Error.WriteLine("Red Ink Semantic Archive Worker: use redink-sa-worker.exe --help for explicit archive/document scope and supported options.")
                Return 2
            End Try
            If options.ShowHelp Then
                ShowHelp()
                Return 0
            End If
            If System.Environment.OSVersion.Platform <> System.PlatformID.Win32NT Then
                WriteConsoleEvent("windows_required", Nothing)
                Return 1
            End If
            Try
                Return RunAsync(options).GetAwaiter().GetResult()
            Catch ex As System.OperationCanceledException
                WriteConsoleEvent("cancelled", New With {.operationId = options.OperationId})
                Return 130
            Catch ex As SharedMethods.HeadlessInteractionRequiredException
                WriteConsoleEvent(ex.Code, New With {.operationId = options.OperationId})
                Return 1
            Catch ex As WorkerLogInitializationException
                WriteConsoleEvent("log_initialization_failed", New With {.operationId = options.OperationId})
                Return 1
            Catch ex As System.Exception
                ' Exception messages can contain provider responses, URLs or credentials.
                WriteConsoleEvent("worker_failed", New With {.operationId = options.OperationId, .exceptionType = ex.GetType().Name})
                Return 1
            End Try
        End Function

        Private Async Function RunAsync(options As WorkerOptions) As System.Threading.Tasks.Task(Of System.Int32)
            Using headless As SharedMethods.HeadlessExecutionScope = SharedMethods.BeginHeadlessExecution(),
                  cancellation As New System.Threading.CancellationTokenSource()
                Dim consoleProgress As New WorkerConsoleProgress()
                Dim onCancel As System.ConsoleCancelEventHandler =
                    Sub(sender As System.Object, args As System.ConsoleCancelEventArgs)
                        args.Cancel = True
                        System.Console.Error.WriteLine("Cancellation requested. Finishing the current safe checkpoint; completed work will remain reusable.")
                        cancellation.Cancel()
                    End Sub
                AddHandler System.Console.CancelKeyPress, onCancel
                Try
                    If options.MaximumSeconds > 0 Then cancellation.CancelAfter(System.TimeSpan.FromSeconds(options.MaximumSeconds))
                    ' Preserve the established version-bearing license identity across the display/executable rename.
                    Dim context As New SharedContext() With {.RDV = "ArchiveWorker (V.031026)"}
                    Dim permissionsOnly As System.Boolean = options.Operation = "permissions"
                    SharedMethods.InitializeConfigHeadless(context, options.ConfigurationSource, requireModelConfiguration:=Not permissionsOnly)
                    context.INI_APIDebug = False
                    headless.ThrowIfInteractionRequested()
                    cancellation.Token.ThrowIfCancellationRequested()
                    If System.String.IsNullOrWhiteSpace(context.INI_SemanticArchiveCatalogPathLocal) Then Throw New System.InvalidOperationException("semantic_archive_disabled")
                    If SemanticArchiveLibrary.IsConfigured(context) Then
                        Try
                            Dim synchronized As System.Int32 = Await SemanticArchiveLibrary.SynchronizeAsync(context, force:=True, cancellationToken:=cancellation.Token).ConfigureAwait(False)
                            WriteConsoleEvent("library_sync", New With {.changed = synchronized})
                        Catch ex As System.OperationCanceledException
                            Throw
                        Catch ex As System.Exception
                            ' Personal/publisher archives can still be processed while a central library is temporarily unavailable.
                            ' Subscribed definitions are independently revalidated by the Semantic Archive services before use.
                            WriteConsoleEvent("library_sync_failed", New With {.exceptionType = ex.GetType().Name})
                        End Try
                    End If
                    Dim store As New SemanticArchiveStore(context.INI_SemanticArchiveCatalogPathLocal)
                    Dim catalog As SemanticArchiveCatalog = store.LoadCatalog()
                    Dim selected As System.Collections.Generic.List(Of System.String) = SelectArchives(catalog, options)
                    headless.ThrowIfInteractionRequested()
                    Using log As New WorkerLog(options.LogPath, catalog, store.DirectoryPath)
                        log.Write("started", New With {.operationId = options.OperationId, .operation = options.Operation,
                            .archives = selected, .allDocuments = options.AllDocuments, .selectedDocumentCount = options.DocumentIds.Count,
                            .mode = If(options.LoopContinuously, "loop", "once")})
                        Dim buildOptions As New SemanticArchiveBuildOptions With {
                            .OperationId = options.OperationId,
                            .SelectedDocumentIds = If(options.AllDocuments, Nothing, New System.Collections.Generic.List(Of System.String)(options.DocumentIds)),
                            .ReconcilePermissionsOnly = permissionsOnly,
                            .ForceReextract = options.Operation = "extract",
                            .RebuildSemanticMetadata = options.Operation = "reindex" OrElse options.Operation = "repair",
                            .IndexOnlyRebuild = options.Operation = "reindex",
                            .RetryFailures = options.Operation = "retry" OrElse options.Operation = "repair",
                            .MaximumFilesPerBatch = options.MaximumFiles,
                            .MaximumWriterLeaseWait = System.TimeSpan.FromMilliseconds(Global.SharedLibrary.SharedLibrary.SharedMethods.DEFAULT_SEMANTICARCHIVE_WORKER_WRITER_LEASE_WAIT_MILLISECONDS),
                            .MaxDiscoveryEntries = options.DiscoveryEntries,
                            .MaxDiscoverySeconds = options.DiscoverySeconds,
                            .ForceScan = True,
                            .IsBackground = False,
                            .HostReaderDispatcher = Nothing
                        }
                        Dim builder As New SemanticArchiveBuilder(context, store)
                        Dim remaining As New System.Collections.Generic.List(Of System.String)(selected)
                        Dim incomplete As System.Boolean = False
                        Dim cycle As System.Int32 = 0
                        Do
                            cancellation.Token.ThrowIfCancellationRequested()
                            cycle += 1
                            Dim nextCycle As New System.Collections.Generic.List(Of System.String)()
                            For Each archiveId As System.String In remaining
                                cancellation.Token.ThrowIfCancellationRequested()
                                Dim result As SemanticArchiveBuildResult
                                Try
                                    result = Await builder.BuildAsync(archiveId, buildOptions, consoleProgress, cancellation.Token).ConfigureAwait(False)
                                    headless.ThrowIfInteractionRequested()
                                Catch ex As System.OperationCanceledException
                                    Throw
                                Catch ex As SharedMethods.HeadlessInteractionRequiredException
                                    Throw
                                Catch ex As System.Exception
                                    headless.ThrowIfInteractionRequested()
                                    log.Write("archive_failed", New With {.operationId = options.OperationId, .archiveId = archiveId,
                                        .cycle = cycle, .exceptionType = ex.GetType().Name})
                                    incomplete = True
                                    Continue For
                                End Try
                                log.Write("batch", New With {.operationId = options.OperationId, .archiveId = archiveId,
                                    .cycle = cycle, .processed = result.ProcessedFiles, .reused = result.ReusedFiles,
                                    .extractsReused = result.ExtractsReused, .cardsRebuilt = result.CardsRebuilt,
                                    .sectionIndexesRebuilt = result.SectionIndexesRebuilt, .routingGroupsBuilt = result.RoutingGroupsBuilt,
                                    .documentsRequiringExtraction = result.DocumentsRequiringExtraction, .coverageExcluded = result.CoverageExcludedFiles,
                                    .failed = result.FailedFiles, .pending = result.PendingFiles, .deferred = result.DeferredFiles,
                                    .published = result.Published, .discoveryPending = result.DiscoveryPending,
                                    .discoveryEntries = result.DiscoveryEntriesInspected, .permissionsPending = result.PermissionsPending,
                                    .permissionsDeferred = result.PermissionsDeferred, .permissionSources = result.PermissionSourcesChecked,
                                    .permissionArtifacts = result.PermissionArtifactsChecked, .permissionRepaired = result.PermissionArtifactsRepaired,
                                    .permissionQuarantined = result.PermissionArtifactsQuarantined, .writerLeaseDeferred = result.WriterLeaseDeferred,
                                    .selectionRequired = result.SelectionRequired, .diagnosticCodes = SafeDiagnosticCodes(result.Diagnostics)})
                                If result.Cancelled Then Throw New System.OperationCanceledException(cancellation.Token)
                                Dim hasPending As System.Boolean = result.PendingFiles > 0 OrElse result.DiscoveryPending OrElse result.PermissionsPending
                                If result.PermissionsDeferred OrElse result.DeferredFiles > 0 OrElse result.FailedFiles > 0 Then incomplete = True
                                If WorkerDrainPolicy.ShouldContinue(result) Then
                                    nextCycle.Add(archiveId)
                                ElseIf result.WriterLeaseDeferred OrElse result.SelectionRequired OrElse result.PermissionsDeferred OrElse result.DeferredFiles > 0 Then
                                    incomplete = True
                                    log.Write("deferred", New With {.operationId = options.OperationId, .archiveId = archiveId})
                                ElseIf hasPending Then
                                    incomplete = True
                                    log.Write("no_progress", New With {.operationId = options.OperationId, .archiveId = archiveId})
                                End If
                            Next
                            remaining = nextCycle
                            If remaining.Count = 0 Then Exit Do
                            If Not options.LoopContinuously OrElse (options.MaximumCycles > 0 AndAlso cycle >= options.MaximumCycles) Then
                                incomplete = True
                                Exit Do
                            End If
                            Await System.Threading.Tasks.Task.Delay(System.TimeSpan.FromSeconds(options.IntervalSeconds), cancellation.Token).ConfigureAwait(False)
                        Loop
                        headless.ThrowIfInteractionRequested()
                        If Not permissionsOnly Then
                            For Each archiveId As System.String In selected
                                cancellation.Token.ThrowIfCancellationRequested()
                                If Not VerifyPublishedScope(store, archiveId, options, log, cancellation.Token) Then incomplete = True
                            Next
                        End If
                        log.Write(If(incomplete, "incomplete", "complete"), New With {.operationId = options.OperationId, .cycles = cycle})
                        Return If(incomplete, 3, 0)
                    End Using
                Finally
                    RemoveHandler System.Console.CancelKeyPress, onCancel
                End Try
            End Using
        End Function

        Private Function VerifyPublishedScope(store As SemanticArchiveStore, archiveId As System.String,
                                              options As WorkerOptions, log As WorkerLog,
                                              cancellationToken As System.Threading.CancellationToken) As System.Boolean
            Try
                Dim generation As SemanticArchiveGenerationManifest = store.PinGeneration(archiveId)
                If generation Is Nothing Then Throw New System.IO.InvalidDataException("generation_unavailable")
                ' A completed queue is not proof of a usable published semantic index.
                store.LoadRoutingGraph(generation)
                Dim selectedIds As New System.Collections.Generic.HashSet(Of System.String)(options.DocumentIds, System.StringComparer.Ordinal)
                Dim records As New System.Collections.Generic.List(Of SemanticArchiveDocumentRecord)()
                Dim needsExtraction As System.Int32 = 0
                Dim unavailable As System.Int32 = 0
                Dim access As SemanticArchiveAccessContext = SemanticArchiveAccessContext.CreateForCurrentUser()
                For Each document As SemanticArchiveDocumentRecord In store.EnumerateDocuments(generation)
                    cancellationToken.ThrowIfCancellationRequested()
                    If Not options.AllDocuments AndAlso Not selectedIds.Contains(document.DocumentId) Then Continue For
                    records.Add(document)
                    selectedIds.Remove(document.DocumentId)
                    If document.ProcessingStatus = "needs_extraction" Then needsExtraction += 1
                    If SemanticArchiveInventory.IsSearchable(document) AndAlso Not store.CanReadDocument(access, generation, document) Then unavailable += 1
                Next
                Dim inventory As SemanticArchiveInventory = SemanticArchiveInventory.FromDocuments(records)
                Dim missingSelected As System.Int32 = If(options.AllDocuments, 0, selectedIds.Count)
                Dim unresolved As System.Int32 = inventory.CurrentSources - inventory.EmptySources - inventory.SearchableDocuments
                Dim complete As System.Boolean = unresolved = 0 AndAlso unavailable = 0 AndAlso missingSelected = 0
                log.Write("published_scope_verified", New With {.operationId = options.OperationId, .archiveId = archiveId,
                    .generationId = generation.GenerationId, .inventory = inventory, .needsExtraction = needsExtraction,
                    .unresolved = unresolved, .sourcesUnavailable = unavailable, .missingSelectedDocuments = missingSelected, .complete = complete})
                Return complete
            Catch ex As System.OperationCanceledException
                Throw
            Catch ex As System.Exception
                log.Write("published_scope_verification_failed", New With {.operationId = options.OperationId,
                    .archiveId = archiveId, .exceptionType = ex.GetType().Name})
                Return False
            End Try
        End Function

        Private Function SelectArchives(catalog As SemanticArchiveCatalog, options As WorkerOptions) As System.Collections.Generic.List(Of System.String)
            Dim result As New System.Collections.Generic.List(Of System.String)()
            If options.AllArchives Then
                For Each archive As SemanticArchiveDefinition In catalog.Archives
                    If Not result.Contains(archive.ArchiveId) Then result.Add(archive.ArchiveId)
                Next
            Else
                For Each selector As System.String In options.ArchiveSelectors
                    Dim selectedId As System.String = ""
                    For Each archive As SemanticArchiveDefinition In catalog.Archives
                        If System.String.Equals(archive.ArchiveId, selector, System.StringComparison.Ordinal) Then
                            selectedId = archive.ArchiveId
                            Exit For
                        End If
                    Next
                    If selectedId.Length = 0 Then
                        Dim matches As New System.Collections.Generic.List(Of SemanticArchiveDefinition)()
                        For Each archive As SemanticArchiveDefinition In catalog.Archives
                            If System.String.Equals(If(archive.Name, "").Trim(), selector.Trim(), System.StringComparison.OrdinalIgnoreCase) Then matches.Add(archive)
                        Next
                        If matches.Count = 0 Then Throw New System.InvalidOperationException("archive_not_found: " & selector)
                        If matches.Count > 1 Then Throw New System.InvalidOperationException("archive_name_ambiguous: " & selector & "; use the stable archive ID instead.")
                        selectedId = matches(0).ArchiveId
                    End If
                    If Not result.Contains(selectedId) Then result.Add(selectedId)
                Next
            End If
            If result.Count = 0 Then Throw New System.InvalidOperationException("selection_required")
            result.Sort(System.StringComparer.Ordinal)
            Return result
        End Function

        Private Function SafeDiagnosticCodes(diagnostics As System.Collections.Generic.IEnumerable(Of System.String)) As System.Collections.Generic.List(Of System.String)
            Dim result As New System.Collections.Generic.List(Of System.String)()
            Dim allowed As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal) From {
                "pending_host", "failed", "shared_claim_deferred", "empty_selection", "selection_required",
                "permissions_deferred", "permission_denied", "path_too_long", "restart_required", "noninteractive_auth_required",
                "coverage_excluded", "requires_extraction", "extraction_contract_mismatch", "extraction_configuration_changed", "extraction_batch_policy_compatible", "routing_rebuild_required", "routing_assignment_fallback", "routing_partition_fallback", "routing_summary_fallback", "cooperative_local_only", "cooperative_contribution_pending", "cooperative_contributed", "cooperative_reuse"}
            If diagnostics Is Nothing Then Return result
            For Each diagnostic As System.String In diagnostics
                Dim separator As System.Int32 = If(diagnostic, "").IndexOf(":"c)
                If separator <= 0 Then Continue For
                Dim code As System.String = diagnostic.Substring(0, separator)
                If allowed.Contains(code) AndAlso Not result.Contains(code) Then result.Add(code)
            Next
            Return result
        End Function

        Private Sub WriteConsoleEvent(code As System.String, details As System.Object)
            System.Console.WriteLine(Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                .utc = System.DateTime.UtcNow.ToString("o", System.Globalization.CultureInfo.InvariantCulture), .code = code, .details = details}))
        End Sub

        Private Sub ShowHelp()
            System.Console.WriteLine("Red Ink Semantic Archive Worker")
            System.Console.WriteLine("Quick use: redink-sa-worker.exe refresh ""Archive Name""")
            System.Console.WriteLine("           redink-sa-worker.exe repair ""Archive A"" ""Archive B""")
            System.Console.WriteLine("           redink-sa-worker.exe reindex ""Archive Name""")
            System.Console.WriteLine("The quick form uses Red Ink's normal redink.ini resolution, all documents and drains one or more named archives to completion.")
            System.Console.WriteLine("retry is stage-aware: missing/empty/incomplete/unknown extraction is freshly extracted/OCRed, while a verified complete extract is reused for downstream retries.")
            System.Console.WriteLine("repair performs that stage-aware repair and also rebuilds semantic cards, section indexes and current routing in the same operation.")
            System.Console.WriteLine("reindex is index-only: it reuses validated extracted text and never invokes extraction/OCR; invalid or missing extracts are reported as requiring extraction.")
            System.Console.WriteLine("")
            System.Console.WriteLine("Advanced: [--ini <path-or-HTTPS-URL>] [--once|--loop]")
            System.Console.WriteLine("  (--archive <name-or-stable-id> [--archive <name-or-id> ...]|--all-archives)")
            System.Console.WriteLine("  [--document <stable-id> [--document <id> ...]|--all-documents]")
            System.Console.WriteLine("  [--operation refresh|retry|repair|reindex|extract|permissions] [--operation-id <GUID>]")
            System.Console.WriteLine("  [--batch-files 1..4096] [--discovery-entries 1..10000] [--discovery-seconds 1..120]")
            System.Console.WriteLine("  [--interval-seconds 1..86400] [--max-cycles 1..1000000] [--max-seconds 1..604800]")
            System.Console.WriteLine("  [--log <private-file-outside-source-roots>]")
            System.Console.WriteLine("")
            System.Console.WriteLine("Defaults for the foreground worker: 64 files per processing batch, up to 5000 discovery entries / 60 seconds per discovery pass, and 1 second between continuation cycles.")
            System.Console.WriteLine("Use --batch-files to throttle model/OCR pressure; e.g. --batch-files 8. Use --discovery-seconds to limit long directory scans per cycle.")
            System.Console.WriteLine("Without --ini the worker resolves the same active redink.ini source as Red Ink. Without document options it uses all documents; without --once it drains the operation like --loop.")
            System.Console.WriteLine("A configured SemanticArchiveCatalogLibraryPath is synchronized before archive selection.")
            System.Console.WriteLine("Progress is appended as separate lines on stderr, including when redirected; window resizing needs no cursor recovery. stdout remains JSON-lines; --log stores the same JSON events.")
            System.Console.WriteLine("Ctrl+C requests graceful cancellation. Durable discovery/extraction checkpoints and the last validated published generation are retained; rerun the same command to continue.")
            System.Console.WriteLine("If a source needs an Office host reader it can remain pending; the worker does not automate Word/Outlook UI.")
            System.Console.WriteLine("Exit codes: 0 complete, 1 configuration/auth/runtime error, 2 arguments, 3 incomplete/deferred, 130 cancelled.")
        End Sub

        Private NotInheritable Class WorkerConsoleProgress
            Implements System.IProgress(Of SemanticArchiveBuildProgress)

            Private ReadOnly _sync As New System.Object()
            Private _lastText As System.String = ""

            Public Sub Report(value As SemanticArchiveBuildProgress) Implements System.IProgress(Of SemanticArchiveBuildProgress).Report
                If value Is Nothing Then Return
                Dim sourceName As System.String = ""
                If Not System.String.IsNullOrWhiteSpace(value.SourcePath) Then
                    Try
                        sourceName = System.IO.Path.GetFileName(value.SourcePath)
                    Catch
                        sourceName = value.SourcePath
                    End Try
                End If
                Dim stage As System.String = FriendlyStage(value.Stage)
                Dim counters As New System.Text.StringBuilder()
                If value.CompletedFiles > 0 OrElse value.PendingFiles > 0 OrElse value.DeferredFiles > 0 Then
                    counters.Append(value.CompletedFiles.ToString(System.Globalization.CultureInfo.InvariantCulture)).Append(" completed")
                    If value.PendingFiles > 0 Then counters.Append("; ").Append(value.PendingFiles.ToString(System.Globalization.CultureInfo.InvariantCulture)).Append(" queued")
                    If value.DeferredFiles > 0 Then counters.Append("; ").Append(value.DeferredFiles.ToString(System.Globalization.CultureInfo.InvariantCulture)).Append(" deferred")
                End If
                Dim line As System.String = stage
                If sourceName.Length > 0 Then line &= " — " & SingleLine(sourceName)
                If counters.Length > 0 Then line &= " | " & counters.ToString()
                WriteProgressLine(line, If(value.SourcePath, ""))
            End Sub

            Private Sub WriteProgressLine(text As System.String, sourcePath As System.String)
                ' Complete lines need no cursor position, width, padding or resize recovery.
                ' Suppress only identical consecutive events; equal filenames in different
                ' directories must not hide a different source's progress.
                Dim line As System.String = SingleLine(text)
                Dim key As System.String = sourcePath & Microsoft.VisualBasic.ChrW(0) & line
                SyncLock _sync
                    If System.String.Equals(key, _lastText, System.StringComparison.Ordinal) Then Return
                    System.Console.Error.WriteLine(line)
                    _lastText = key
                End SyncLock
            End Sub

            Private Shared Function SingleLine(value As System.String) As System.String
                Dim result As New System.Text.StringBuilder()
                For Each character As System.Char In If(value, "")
                    ' Filenames/stage text must never inject terminal controls or extra lines.
                    If System.Char.IsControl(character) OrElse character = Microsoft.VisualBasic.ChrW(&H2028) OrElse character = Microsoft.VisualBasic.ChrW(&H2029) Then
                        result.Append(" "c)
                    Else
                        result.Append(character)
                    End If
                Next
                Return result.ToString()
            End Function

            Private Shared Function FriendlyStage(stage As System.String) As System.String
                Select Case If(stage, "").Trim().ToLowerInvariant()
                    Case "scanning" : Return "Scanning sources"
                    Case "validating_extracts" : Return "Validating existing extracts"
                    Case "permissions" : Return "Checking permissions"
                    Case "processing" : Return "Processing source"
                    Case "extracting" : Return "Extracting text"
                    Case "indexing" : Return "Building semantic metadata"
                    Case "hierarchy" : Return "Updating archive index"
                    Case "routing" : Return "Building routing groups"
                    Case "publishing" : Return "Publishing generation"
                    Case "coverage_excluded" : Return "Coverage check"
                    Case "shared_claim_deferred" : Return "Shared artifact deferred"
                    Case "failed" : Return "Source failed"
                    Case Else
                        If System.String.IsNullOrWhiteSpace(stage) Then Return "Working"
                        Return stage.Replace("_", " ")
                End Select
            End Function
        End Class

        Private NotInheritable Class WorkerLogInitializationException
            Inherits System.IO.IOException

            Public Sub New(inner As System.Exception)
                MyBase.New("The private worker log could not be initialized.", inner)
            End Sub
        End Class

        Private NotInheritable Class WorkerLog
            Implements System.IDisposable
            Private ReadOnly _writer As System.IO.StreamWriter

            Public Sub New(path As System.String, catalog As SemanticArchiveCatalog, controlDirectory As System.String)
                If System.String.IsNullOrWhiteSpace(path) Then Return
                Try
                    Dim full As System.String = SemanticArchivePathGuard.RequireWindowsCompatiblePath(path)
                    If GeneratedOutputRegistry.IsPhysicalPathAtOrBelow(controlDirectory, full) Then Throw New System.ArgumentException("The log must be outside the archive control directory.")
                    For Each archive As SemanticArchiveDefinition In catalog.Archives
                        For Each binding As SemanticArchiveSourceBinding In archive.Roots
                            If GeneratedOutputRegistry.IsPhysicalPathAtOrBelow(binding.RootPath, full) Then Throw New System.ArgumentException("The log must be outside configured source roots.")
                        Next
                    Next
                    Dim parent As System.String = System.IO.Path.GetDirectoryName(full)
                    SemanticArchiveStore.CreatePrivateDirectory(parent)
                    ' CreateNew is deliberate: a mistyped log path must never append to a source
                    ' or an existing control artifact, including through a filesystem alias.
                    Dim stream As System.IO.FileStream = SemanticArchiveStore.CreatePrivateFile(full)
                    _writer = New System.IO.StreamWriter(stream, New System.Text.UTF8Encoding(False)) With {.AutoFlush = True}
                Catch ex As System.Exception
                    Throw New WorkerLogInitializationException(ex)
                End Try
            End Sub

            Public Sub Write(code As System.String, details As System.Object)
                Dim line As System.String = Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                    .utc = System.DateTime.UtcNow.ToString("o", System.Globalization.CultureInfo.InvariantCulture), .code = code, .details = details})
                System.Console.WriteLine(line)
                If _writer IsNot Nothing Then _writer.WriteLine(line)
            End Sub

            Public Sub Dispose() Implements System.IDisposable.Dispose
                If _writer IsNot Nothing Then _writer.Dispose()
            End Sub
        End Class
    End Module
End Namespace
