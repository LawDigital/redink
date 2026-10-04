' Part of "Red Ink" (SharedLibrary)
' Bounded, generation-pinned archive search. Models select metadata; the host resolves IDs.
Option Strict On
Option Explicit On

Namespace SharedLibrary
    Public NotInheritable Class SemanticArchiveCoverage
        ' Counts describe published records, not permissions or proof of complete corpus coverage.
        Public Property PublishedInventory As New System.Collections.Generic.Dictionary(Of System.String, SemanticArchiveInventory)(System.StringComparer.Ordinal)
        Public Property Generations As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.Ordinal)
        Public Property RetrievalStrategy As System.String = "hierarchical"
        Public Property DocumentMetadataSelectionComplete As System.Boolean
        Public Property NodesVisited As System.Int32
        Public Property ModelCalls As System.Int32
        Public Property TotalNodesVisited As System.Int64
        Public Property TotalModelCalls As System.Int64
        Public Property TotalExactMetadataRecordsInspected As System.Int64
        Public Property SourceAuthorizationChecks As System.Int64
        Public Property LiteralTextQuery As System.String
        Public Property LiteralDocumentsInspected As System.Int32
        Public Property LiteralDocumentsSkippedByBudget As System.Int32
        Public Property LiteralByteBudgetUsed As System.Int64
        Public Property TotalLiteralByteBudgetUsed As System.Int64
        Public Property LiteralScanCoverageComplete As System.Boolean
        Public Property CardsConsidered As System.Int32
        Public Property ExactMetadataRecordsInspected As System.Int32
        Public Property FilesInspected As System.Int32
        Public Property SourcesUnavailable As System.Int32
        Public Property IncompleteExtractions As System.Int32
        Public Property NeutralNavigationBranches As System.Int32
        Public Property EvidenceBytes As System.Int64
        Public Property SectionsConsidered As System.Int32
        Public Property TotalSections As System.Int32
        Public Property IndexedFilesInspected As System.Int32
        Public Property IndexedFilesWithoutSelection As System.Int32
        Public Property IndexedReadFailures As System.Int32
        Public Property ExactMetadataCoverageComplete As System.Boolean
        Public Property TraversalComplete As System.Boolean
        Public Property BudgetEndedSearch As System.Boolean
        Public Property Exhaustive As System.Boolean = False
        Public Property ElapsedMilliseconds As System.Int64
        Public Property Diagnostics As New System.Collections.Generic.List(Of System.String)()
        Public Property CoverageStatement As System.String = "Results cover the explored metadata and loaded excerpts only. A missing match is not an exhaustive negative finding. Exact lookup covers names and retained metadata, not every identifier in original text."
    End Class

    Public NotInheritable Class SemanticArchiveSearchHit
        Public Property HitReference As System.String
        Public Property ArchiveId As System.String
        Public Property GenerationId As System.String
        Public Property DocumentId As System.String
        Public Property SourcePath As System.String
        Public Property SourceUri As System.String
        Public Property DisplayName As System.String
        Public Property Title As System.String
        Public Property Summary As System.String
        Public Property Relevance As System.Double
        Public Property Reason As System.String
        Public Property Channel As System.String
        Public Property RepresentationId As System.String
        Public Property SourceSha256 As System.String
        Public Property ExtractionCompleteness As System.String
        Public Property SummaryIsEvidence As System.Boolean = False
        Public Property LiteralMatchStartByte As System.Nullable(Of System.Int64)
        Public Property LiteralMatchOffsetBase As System.String
    End Class

    Public NotInheritable Class SemanticArchiveSearchResult
        Public Property Status As System.String = "ok"
        Public Property Message As System.String = ""
        Public Property Query As System.String = ""
        Public Property Hits As New System.Collections.Generic.List(Of SemanticArchiveSearchHit)()
        Public Property ContinuationReference As System.String
        Public Property Coverage As New SemanticArchiveCoverage()
    End Class

    Friend NotInheritable Class SemanticArchiveSearchState
        Friend Property Query As System.String
        Friend Property Budgets As SemanticArchiveRetrievalBudgets
        Friend Property TotalNodes As System.Int64
        Friend Property TotalCalls As System.Int64
        Friend Property TotalExactRecords As System.Int64
        Friend Property SearchCalls As System.Int64
        Friend Property SourceChecksAtCallStart As System.Int64
        Friend Property HadTraversalFailures As System.Boolean
        Friend Property HadExactFailures As System.Boolean
        Friend Property ArchiveIds As New System.Collections.Generic.List(Of System.String)()
        Friend Property Queue As New System.Collections.Generic.List(Of SemanticArchiveNodeWork)()
        Friend Property SeenNodes As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
        Friend Property Candidates As New System.Collections.Generic.Dictionary(Of System.String, SemanticArchiveCandidate)(System.StringComparer.Ordinal)
        Friend Property Returned As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
        Friend Property ExactEnumerators As New System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.IEnumerator(Of SemanticArchiveDocumentRecord))(System.StringComparer.Ordinal)
        Friend Property ExactArchivePosition As System.Int32
        Friend Property ExactComplete As System.Boolean
        Friend Property FlatEligible As System.Boolean = True
        Friend Property FlatAttempted As System.Boolean
        Friend Property FlatDocuments As New System.Collections.Generic.List(Of SemanticArchiveFlatDocument)()
        Friend Property Sequence As System.Int64
        Friend Property NeedsWidening As System.Boolean
        Friend Property LiteralScan As SemanticArchiveLiteralScanState
    End Class

    Friend NotInheritable Class SemanticArchiveFlatDocument
        Friend Property ArchiveId As System.String
        Friend Property Document As SemanticArchiveDocumentRecord
    End Class

    Friend NotInheritable Class SemanticArchiveNodeWork
        Friend Property ArchiveId As System.String
        Friend Property NodeId As System.String
        Friend Property Priority As System.Double
        Friend Property Sequence As System.Int64
        Friend Property ProcessedCardIds As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
    End Class

    Friend NotInheritable Class SemanticArchiveCandidate
        Friend Property ArchiveId As System.String
        Friend Property DocumentId As System.String
        Friend Property Score As System.Double
        Friend Property Reason As System.String
        Friend Property Channel As System.String
        Friend Property LiteralMatchStartByte As System.Nullable(Of System.Int64)
    End Class

    ''' <summary>
    ''' One implementation for direct/agentic Word and Outlook retrieval. Selection,
    ''' continuations and evidence authority stay in the host's run scope. Generations
    ''' pin bytes, but every metadata/read disclosure rechecks current source access.
    ''' </summary>
    Partial Public NotInheritable Class SemanticArchiveSearchService
        Private ReadOnly _store As SemanticArchiveStore
        Private ReadOnly _scope As SemanticArchiveRunScope
        Private ReadOnly _context As Global.SharedLibrary.SharedLibrary.SharedContext.ISharedContext
        Private Const MaximumQueryCharacters As System.Int32 = 12000
        Private Const MaximumRunEvidenceBytes As System.Int64 = 512L * 1024L

        Public Sub New(store As SemanticArchiveStore, scope As SemanticArchiveRunScope,
                       context As Global.SharedLibrary.SharedLibrary.SharedContext.ISharedContext)
            If store Is Nothing Then Throw New System.ArgumentNullException(NameOf(store))
            If scope Is Nothing Then Throw New System.ArgumentNullException(NameOf(scope))
            _store = store
            _scope = scope
            _context = context
        End Sub

        Private Shared Function BuildSourceUri(sourcePath As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(sourcePath) Then Return System.String.Empty
            Try
                Dim fullPath As System.String = SemanticArchivePathGuard.RequireWindowsSourcePath(sourcePath)
                Return (New System.Uri(fullPath)).AbsoluteUri
            Catch ex As System.Exception
                Return System.String.Empty
            End Try
        End Function

        Public Async Function SearchAsync(query As System.String,
                        Optional narrowedArchiveIds As System.Collections.Generic.IEnumerable(Of System.String) = Nothing,
                        Optional limit As System.Int32 = SharedMethods.DEFAULT_SEMANTICARCHIVE_SEARCH_LIMIT,
                        Optional continuationReference As System.String = Nothing,
                        Optional cancellationToken As System.Threading.CancellationToken = Nothing,
                        Optional literalText As System.String = Nothing) As System.Threading.Tasks.Task(Of SemanticArchiveSearchResult)
            If Not SemanticArchiveHostIntegration.IsConfigured(_context) Then Return New SemanticArchiveSearchResult With {.Status = "not_configured", .Message = "Semantic Archive is disabled: the local catalog path is empty."}
            If Not System.String.IsNullOrWhiteSpace(_scope.AccessContext.DenialCode) Then
                Return New SemanticArchiveSearchResult() With {.Status = _scope.AccessContext.DenialCode,
                    .Message = If(System.String.IsNullOrWhiteSpace(_scope.AccessContext.DenialMessage), "The host could not verify the requesting principal's source authorization.", _scope.AccessContext.DenialMessage)}
            End If
            If _scope.ResolutionStatus.Length > 0 Then
                Return New SemanticArchiveSearchResult() With {.Status = _scope.ResolutionStatus, .Message = _scope.ResolutionMessage}
            End If
            Dim activeScope As SemanticArchiveRunScope
            Try
                activeScope = _scope.Narrow(narrowedArchiveIds)
            Catch ex As System.UnauthorizedAccessException
                Return New SemanticArchiveSearchResult() With {.Status = "scope_rejected", .Message = ex.Message}
            End Try
            If activeScope.SelectedArchiveIds.Count = 0 Then Return New SemanticArchiveSearchResult() With {.Status = "selection_required", .Message = "The host must select an archive before search."}
            If limit < 1 OrElse limit > 50 Then Return New SemanticArchiveSearchResult() With {.Status = "invalid_limit", .Message = "The result limit must be between 1 and 50."}
            Await activeScope.State.OperationGate.WaitAsync(cancellationToken).ConfigureAwait(False)
            Try
                Using priority As System.IDisposable = BackgroundMaintenanceCoordinator.EnterInteractiveWork()
                    activeScope.State.EnsureUsable()
                    activeScope.State.BindDirectory(_store.DirectoryPath)
                    Return Await System.Threading.Tasks.Task.Run(Function() SearchCoreAsync(activeScope, query, limit, continuationReference, cancellationToken, literalText), cancellationToken).ConfigureAwait(False)
                End Using
            Finally
                activeScope.State.OperationGate.Release()
            End Try
        End Function

        Private Async Function SearchCoreAsync(scope As SemanticArchiveRunScope, query As System.String, limit As System.Int32,
                        continuationReference As System.String,
                        cancellationToken As System.Threading.CancellationToken,
                        literalText As System.String) As System.Threading.Tasks.Task(Of SemanticArchiveSearchResult)
            SemanticArchiveLibrary.ValidateScope(_context, _store, scope.SelectedArchiveIds)
            Dim result As New SemanticArchiveSearchResult() With {.Query = If(query, "")}
            Dim timer As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew()
            Dim sourceChecksBefore As System.Int64 = scope.AccessContext.SourceChecks
            Dim budgets As SemanticArchiveRetrievalBudgets = Nothing
            Dim state As SemanticArchiveSearchState = Nothing
            Dim continuation As System.String = continuationReference
            If Not System.String.IsNullOrWhiteSpace(continuation) Then
                If Not scope.State.Searches.TryGetValue(continuation, state) Then
                    result.Status = "invalid_continuation"
                    result.Message = "The continuation is unknown or belongs to another run."
                    Return result
                End If
                If Not System.String.IsNullOrWhiteSpace(query) AndAlso Not System.String.Equals(query.Trim(), state.Query, System.StringComparison.Ordinal) Then
                    result.Status = "continuation_query_mismatch"
                    result.Message = "A continuation retains its original query. Start a new search for a changed query."
                    Return result
                End If
                For Each archiveId As System.String In state.ArchiveIds
                    If Not scope.ContainsArchive(archiveId) Then
                        result.Status = "scope_rejected"
                        result.Message = "The continuation includes archives outside the narrowed scope. Start a search within that scope."
                        Return result
                    End If
                Next
                query = state.Query
                If literalText IsNot Nothing AndAlso (state.LiteralScan Is Nothing OrElse Not System.String.Equals(literalText, state.LiteralScan.LiteralText, System.StringComparison.Ordinal)) Then
                    result.Status = "continuation_literal_mismatch"
                    result.Message = "A continuation retains its original literal-text inspection. Start a new search to add or change the literal."
                    Return result
                End If
            Else
                If System.String.IsNullOrWhiteSpace(query) Then
                    result.Status = "missing_query"
                    result.Message = "Provide a nonempty archive query."
                    Return result
                End If
                query = query.Trim()
                If query.Length > MaximumQueryCharacters Then
                    result.Status = "query_budget_exceeded"
                    result.Message = "The SA query exceeds the complete request budget. Submit a shorter, explicitly scoped question."
                    Return result
                End If
                If literalText IsNot Nothing Then
                    If literalText.Length = 0 OrElse literalText.Length > 4096 Then
                        result.Status = "invalid_literal"
                        result.Message = "An explicit literal must contain between 1 and 4096 UTF-16 code units."
                        Return result
                    End If
                    Try
                        Dim ignored As System.Int32 = New System.Text.UTF8Encoding(False, True).GetByteCount(literalText)
                    Catch ex As System.Text.EncoderFallbackException
                        result.Status = "invalid_literal"
                        result.Message = "The literal contains an incomplete Unicode character."
                        Return result
                    End Try
                End If
                If scope.State.Searches.Count >= SemanticArchiveRunState.MaximumContinuations Then
                    result.Status = "run_state_budget"
                    result.Message = "This run has reached its bounded continuation capacity. Start a fresh host retrieval run."
                    Return result
                End If
                budgets = ResolveBudgets(scope)
                state = New SemanticArchiveSearchState() With {.Query = query, .Budgets = budgets}
                If literalText IsNot Nothing Then state.LiteralScan = New SemanticArchiveLiteralScanState() With {.LiteralText = literalText}
                continuation = "sac_" & System.Guid.NewGuid().ToString("N")
                For Each archiveId As System.String In scope.SelectedArchiveIds
                    Try
                        Dim definition As SemanticArchiveDefinition = _store.GetArchive(archiveId)
                        If definition Is Nothing OrElse Not definition.Enabled Then
                            result.Coverage.SourcesUnavailable += 1
                            Continue For
                        End If
                        Dim generation As SemanticArchiveGenerationManifest = Nothing
                        If Not scope.State.Generations.TryGetValue(archiveId, generation) Then
                            generation = _store.PinGeneration(archiveId)
                            If generation IsNot Nothing Then scope.State.Generations.Add(archiveId, generation)
                        End If
                        If generation Is Nothing Then
                            result.Coverage.SourcesUnavailable += 1
                            Continue For
                        End If
                        state.ArchiveIds.Add(archiveId)
                        Enqueue(state, archiveId, generation.RootNodeId, 100.0R)
                        state.ExactEnumerators.Add(archiveId, _store.EnumerateDocuments(generation).GetEnumerator())
                        If state.LiteralScan IsNot Nothing Then state.LiteralScan.Enumerators.Add(archiveId, _store.EnumerateDocuments(generation).GetEnumerator())
                    Catch ex As System.Exception
                        System.Diagnostics.Debug.WriteLine("SA generation pin failed: " & ex.GetType().FullName)
                        result.Coverage.SourcesUnavailable += 1
                        result.Coverage.Diagnostics.Add("generation_unavailable: An authorized archive has no usable published generation.")
                    End Try
                Next
                scope.State.Searches.Add(continuation, state)
            End If
            If budgets Is Nothing Then budgets = ResolveBudgets(scope)
            state.Budgets = budgets
            state.SourceChecksAtCallStart = sourceChecksBefore
            result.Query = query
            For Each archiveId As System.String In state.ArchiveIds
                result.Coverage.Generations(archiveId) = scope.State.Generations(archiveId).GenerationId
                Dim inventory As SemanticArchiveInventory = scope.State.Generations(archiveId).Inventory
                result.Coverage.PublishedInventory(archiveId) = SemanticArchiveMetadata.Clone(inventory)
                If inventory.ExcludedUnknown > 0 OrElse inventory.ExcludedIncomplete > 0 Then
                    result.Coverage.Diagnostics.Add("extraction_coverage_exclusions: Published records are not searchable because their coverage is unknown/incomplete or their validity is suppressed. See PublishedInventory; query reformulation cannot make excluded records searchable.")
                End If
                If inventory.SearchableDocuments = 0 Then result.Coverage.Diagnostics.Add("no_searchable_documents: The selected published generation contains no searchable document cards. Inspect the archive build diagnostics before another search.")
            Next
            If state.ArchiveIds.Count = 0 Then
                result.Status = "generation_unavailable"
                result.Message = "No selected archive has a readable published generation."
                scope.State.Searches.Remove(continuation)
                Return result
            End If

            Using deadline As System.Threading.CancellationTokenSource = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(cancellationToken)
            deadline.CancelAfter(CInt(System.Math.Max(1L, CLng(budgets.MaxElapsedSeconds) * 1000L - timer.ElapsedMilliseconds)))
            ' Independent exact channel: this does not determine semantic eligibility.
            ' The host retains the sharded enumerators for subsequent bounded calls.
            Try
                If state.LiteralScan IsNot Nothing Then RunLiteralTextChannel(scope, state, result.Coverage, timer, deadline.Token)
                RunExactMetadataChannel(scope, state, result.Coverage, timer, deadline.Token)
            Catch ex As System.OperationCanceledException When Not cancellationToken.IsCancellationRequested
                result.Coverage.Diagnostics.Add("elapsed_budget: Literal inspection or exact metadata lookup stopped at the configured operation deadline.")
            End Try
            Dim exactMetadataFastPath As System.Boolean = state.SearchCalls = 0 AndAlso CanUseExactMetadataFastPath(state)
            Dim returnMetadataPage As System.Boolean = False
            If exactMetadataFastPath Then
                ' Keep the hierarchy available on continuation; do not report it as traversed.
                result.Coverage.RetrievalStrategy = "exact_metadata_first_page"
                returnMetadataPage = True
            ElseIf Not state.FlatAttempted AndAlso state.FlatEligible AndAlso state.ExactComplete AndAlso Not state.HadExactFailures AndAlso state.LiteralScan Is Nothing Then
                state.FlatAttempted = True
                Try
                    returnMetadataPage = Await SearchSmallCatalogAsync(scope, state, result.Coverage, deadline.Token).ConfigureAwait(False)
                Catch ex As System.OperationCanceledException When Not cancellationToken.IsCancellationRequested
                    result.Coverage.Diagnostics.Add("elapsed_budget: Document-card selection stopped at the configured deadline. The hierarchy remains available on continuation.")
                Catch ex As System.OperationCanceledException
                    Throw
                Catch ex As System.Exception
                    result.Coverage.Diagnostics.Add("flat_metadata_unavailable: Direct document-card selection could not be completed; normal hierarchy navigation remains available.")
                    System.Diagnostics.Trace.WriteLine("[SemanticArchive] Flat metadata selection: " & ex.GetType().FullName)
                End Try
            End If
            If Not state.ExactComplete Then state.FlatEligible = False
            state.FlatDocuments.Clear()
            While Not returnMetadataPage AndAlso state.Queue.Count > 0 AndAlso result.Coverage.NodesVisited < budgets.MaxNodesVisited AndAlso
                  result.Coverage.ModelCalls < budgets.MaxModelCalls AndAlso timer.ElapsedMilliseconds < CLng(budgets.MaxElapsedSeconds) * 1000L AndAlso
                  (state.Candidates.Count < budgets.MaxCandidateFiles OrElse result.Coverage.NodesVisited = 0)
                cancellationToken.ThrowIfCancellationRequested()
                If CountUnreturnedCandidates(state) >= limit * 3 AndAlso result.Coverage.NodesVisited > 0 AndAlso Not state.NeedsWidening Then Exit While
                state.Queue.Sort(Function(left As SemanticArchiveNodeWork, right As SemanticArchiveNodeWork)
                                     Dim priority As System.Int32 = right.Priority.CompareTo(left.Priority)
                                     Return If(priority <> 0, priority, left.Sequence.CompareTo(right.Sequence))
                                 End Function)
                Dim work As SemanticArchiveNodeWork = state.Queue(0)
                state.Queue.RemoveAt(0)
                Dim key As System.String = Identity(work.ArchiveId, work.NodeId)
                If Not state.SeenNodes.Add(key) Then Continue While
                result.Coverage.NodesVisited += 1
                Dim previousProgress As New System.Collections.Generic.HashSet(Of System.String)(work.ProcessedCardIds, System.StringComparer.OrdinalIgnoreCase)
                Try
                    Await SearchNodeAsync(scope, state, work, result.Coverage, deadline.Token).ConfigureAwait(False)
                Catch ex As System.OperationCanceledException When Not cancellationToken.IsCancellationRequested
                    work.ProcessedCardIds = previousProgress
                    state.SeenNodes.Remove(key)
                    If Not state.Queue.Contains(work) Then state.Queue.Add(work)
                    result.Coverage.Diagnostics.Add("elapsed_budget: Model selection stopped at the configured deadline; unfinished node cards remain eligible on continuation.")
                    Exit While
                Catch ex As System.OperationCanceledException
                    work.ProcessedCardIds = previousProgress
                    state.SeenNodes.Remove(key)
                    If Not state.Queue.Contains(work) Then state.Queue.Add(work)
                    Throw
                Catch ex As System.Exception
                    state.HadTraversalFailures = True
                    System.Diagnostics.Debug.WriteLine("SA node selection failed: " & ex.GetType().FullName)
                    result.Coverage.Diagnostics.Add("node_or_selection_failed: A visited node could not be validated or selected. Its omission is explicit; no unvalidated summary was used.")
                    result.Coverage.SourcesUnavailable += 1
                End Try
            End While
            Dim ranked As New System.Collections.Generic.List(Of SemanticArchiveCandidate)(state.Candidates.Values)
            ranked.Sort(Function(left As SemanticArchiveCandidate, right As SemanticArchiveCandidate)
                            Dim score As System.Int32 = right.Score.CompareTo(left.Score)
                            Return If(score <> 0, score, System.StringComparer.Ordinal.Compare(Identity(left.ArchiveId, left.DocumentId), Identity(right.ArchiveId, right.DocumentId)))
                        End Function)
            For Each candidate As SemanticArchiveCandidate In ranked
                If result.Hits.Count >= limit Then Exit For
                Dim key As System.String = Identity(candidate.ArchiveId, candidate.DocumentId)
                If state.Returned.Contains(key) Then Continue For
                Dim generation As SemanticArchiveGenerationManifest = scope.State.Generations(candidate.ArchiveId)
                Dim document As SemanticArchiveDocumentRecord = Nothing
                Try
                    document = _store.LoadDocument(generation, candidate.DocumentId)
                Catch ex As System.Exception
                    System.Diagnostics.Debug.WriteLine("SA candidate validation failed: " & ex.GetType().FullName)
                    state.HadTraversalFailures = True
                    result.Coverage.Diagnostics.Add("candidate_unavailable: One selected file record could not be validated; other independently authorized candidates remain available.")
                End Try
                ' Recheck after model calls and before disclosing a title, summary or path.
                If document Is Nothing OrElse Not _store.CanReadDocument(scope.AccessContext, generation, document) Then
                    result.Coverage.SourcesUnavailable += 1
                    state.Returned.Add(key)
                    state.Candidates.Remove(key)
                    Continue For
                End If
                If scope.State.Hits.Count >= SemanticArchiveRunState.MaximumEvidenceReferences Then
                    result.Coverage.Diagnostics.Add("run_state_budget: The run has reached its evidence-reference capacity.")
                    Exit For
                End If
                Dim hitReference As System.String = "sah_" & System.Guid.NewGuid().ToString("N")
                scope.State.Hits.Add(hitReference, New SemanticArchiveEvidenceReference() With {
                    .ArchiveId = candidate.ArchiveId, .DocumentId = candidate.DocumentId, .Generation = generation, .Query = query,
                    .LiteralMatchStartByte = candidate.LiteralMatchStartByte,
                    .LiteralText = If(candidate.LiteralMatchStartByte.HasValue AndAlso state.LiteralScan IsNot Nothing, state.LiteralScan.LiteralText, Nothing)})
                Dim card As SemanticArchiveCard = document.Card
                result.Hits.Add(New SemanticArchiveSearchHit() With {
                    .HitReference = hitReference, .ArchiveId = candidate.ArchiveId, .GenerationId = generation.GenerationId,
                    .DocumentId = document.DocumentId, .SourcePath = document.SourcePath, .SourceUri = BuildSourceUri(document.SourcePath), .DisplayName = document.DisplayName,
                    .Title = If(card Is Nothing, document.DisplayName, card.Title), .Summary = If(card Is Nothing, "", card.Summary),
                    .Relevance = candidate.Score, .Reason = candidate.Reason, .Channel = candidate.Channel,
                    .RepresentationId = document.Representation.RepresentationId, .SourceSha256 = document.Representation.SourceHash,
                    .ExtractionCompleteness = document.Representation.Completeness, .LiteralMatchStartByte = candidate.LiteralMatchStartByte,
                    .LiteralMatchOffsetBase = If(candidate.LiteralMatchStartByte.HasValue, "exported_text_content_without_utf8_bom", Nothing)})
                state.Returned.Add(key)
                state.Candidates.Remove(key)
                If Not System.String.Equals(document.Representation.Completeness, "complete", System.StringComparison.OrdinalIgnoreCase) Then result.Coverage.IncompleteExtractions += 1
            Next
            result.Coverage.ExactMetadataCoverageComplete = state.ExactComplete AndAlso Not state.HadExactFailures AndAlso budgets.MaxExactLookupDocuments > 0
            result.Coverage.TraversalComplete = result.Coverage.RetrievalStrategy = "hierarchical" AndAlso state.Queue.Count = 0 AndAlso Not state.HadTraversalFailures
            Dim hasMore As System.Boolean = state.Queue.Count > 0 OrElse Not state.ExactComplete OrElse CountUnreturnedCandidates(state) > 0 OrElse
                (state.LiteralScan IsNot Nothing AndAlso Not state.LiteralScan.Completed)
            result.Coverage.BudgetEndedSearch = hasMore AndAlso Not returnMetadataPage
            If hasMore Then
                ' Rotate a successfully consumed page token. Legitimate continuation calls
                ' therefore do not look like identical failed retries to generic tooling.
                Dim nextContinuation As System.String = "sac_" & System.Guid.NewGuid().ToString("N")
                scope.State.Searches.Add(nextContinuation, state)
                scope.State.Searches.Remove(continuation)
                continuation = nextContinuation
                result.ContinuationReference = continuation
            Else
                DisposeExactEnumerators(state)
                scope.State.Searches.Remove(continuation)
            End If
            If state.HadExactFailures OrElse state.HadTraversalFailures Then result.Coverage.Diagnostics.Add("retained_omissions: This continuation retains explicitly failed metadata or traversal branches; it does not claim complete coverage.")
            If state.LiteralScan IsNot Nothing AndAlso state.LiteralScan.HadOmissions Then result.Coverage.Diagnostics.Add("literal_retained_omissions: This search retains access, integrity, extraction, or complete-validation byte omissions; literal coverage is not exhaustive.")
            If result.Coverage.IncompleteExtractions > 0 Then result.Coverage.Diagnostics.Add("incomplete_extraction: At least one returned source has partial or unknown extraction completeness; omitted original content cannot be ruled out.")
            If hasMore OrElse result.Coverage.Diagnostics.Count > 0 OrElse result.Coverage.SourcesUnavailable > 0 Then result.Status = "partial"
            If result.Hits.Count = 0 Then
                result.Message = If(hasMore,
                    "No matching file was returned on this page. Continue with ContinuationReference and the unchanged query to inspect remaining work within the budgets.",
                    "No matching file was returned from the inspected searchable records. This search has no continuation. Consult PublishedInventory and Coverage.Diagnostics for excluded or unavailable sources; this is not proof that the original documents contain no answer.")
            End If
            state.TotalNodes += result.Coverage.NodesVisited
            state.SearchCalls += 1
            state.TotalCalls += result.Coverage.ModelCalls
            state.TotalExactRecords += result.Coverage.ExactMetadataRecordsInspected
            result.Coverage.TotalNodesVisited = state.TotalNodes
            result.Coverage.TotalModelCalls = state.TotalCalls
            result.Coverage.TotalExactMetadataRecordsInspected = state.TotalExactRecords
            If state.LiteralScan IsNot Nothing Then
                result.Coverage.LiteralTextQuery = state.LiteralScan.LiteralText
                result.Coverage.TotalLiteralByteBudgetUsed = state.LiteralScan.TotalByteBudgetUsed
                result.Coverage.LiteralScanCoverageComplete = state.LiteralScan.Completed AndAlso Not state.LiteralScan.HadOmissions
            End If
            result.Coverage.SourceAuthorizationChecks = scope.AccessContext.SourceChecks - sourceChecksBefore
            result.Coverage.ElapsedMilliseconds = timer.ElapsedMilliseconds
            If Not System.String.IsNullOrWhiteSpace(scope.AccessContext.DenialCode) Then
                For Each hit As SemanticArchiveSearchHit In result.Hits
                    scope.State.Hits.Remove(hit.HitReference)
                Next
                result.Hits.Clear()
                result.ContinuationReference = Nothing
                result.Status = scope.AccessContext.DenialCode
                result.Message = If(System.String.IsNullOrWhiteSpace(scope.AccessContext.DenialMessage), "The requesting principal's source authorization is no longer verified.", scope.AccessContext.DenialMessage)
                DisposeExactEnumerators(state)
                scope.State.Searches.Remove(continuation)
            End If
            SemanticArchiveLibrary.ValidateScope(_context, _store, scope.SelectedArchiveIds)
            Return result
            End Using
        End Function

        Private Sub RunExactMetadataChannel(scope As SemanticArchiveRunScope, state As SemanticArchiveSearchState,
                      coverage As SemanticArchiveCoverage, timer As System.Diagnostics.Stopwatch,
                      cancellationToken As System.Threading.CancellationToken)
            If state.Budgets.MaxExactLookupDocuments = 0 Then
                state.ExactComplete = True
                coverage.Diagnostics.Add("exact_lookup_disabled: The selected archive policy assigns no exact-metadata lookup budget.")
                Return
            End If
            ' A one-file policy alternates the independent channels across calls.
            ' Larger policies reserve capacity for semantic candidates in every call.
            Dim exactCapacity As System.Int32 = If(state.Budgets.MaxCandidateFiles = 1,
                If((state.SearchCalls Mod 2L) <> 0L OrElse state.Queue.Count = 0, 1, 0), System.Math.Max(1, state.Budgets.MaxCandidateFiles \ 2))
            If state.LiteralScan IsNot Nothing AndAlso Not state.LiteralScan.Completed Then
                ' The explicitly requested text inspection gets its reserved slots;
                ' semantic candidates keep at least one slot where the policy permits.
                exactCapacity = If(state.Budgets.MaxCandidateFiles < 3, 0,
                    System.Math.Min(state.Budgets.MaxCandidateFiles - 1, state.Candidates.Count + System.Math.Max(1, state.Budgets.MaxCandidateFiles \ 3)))
            End If
            While state.ExactArchivePosition < state.ArchiveIds.Count AndAlso coverage.ExactMetadataRecordsInspected < state.Budgets.MaxExactLookupDocuments AndAlso
                  timer.ElapsedMilliseconds < CLng(state.Budgets.MaxElapsedSeconds) * 1000L \ 3 AndAlso state.Candidates.Count < exactCapacity
                cancellationToken.ThrowIfCancellationRequested()
                Dim archiveId As System.String = state.ArchiveIds(state.ExactArchivePosition)
                Dim iterator As System.Collections.Generic.IEnumerator(Of SemanticArchiveDocumentRecord) = state.ExactEnumerators(archiveId)
                Dim moved As System.Boolean
                Try
                    moved = iterator.MoveNext()
                Catch ex As System.Exception
                    state.HadExactFailures = True
                    coverage.Diagnostics.Add("exact_metadata_unavailable: A metadata shard could not be inspected.")
                    coverage.SourcesUnavailable += 1
                    moved = False
                End Try
                If Not moved Then
                    iterator.Dispose()
                    state.ExactArchivePosition += 1
                    Continue While
                End If
                coverage.ExactMetadataRecordsInspected += 1
                Dim document As SemanticArchiveDocumentRecord = iterator.Current
                Dim generation As SemanticArchiveGenerationManifest = scope.State.Generations(archiveId)
                If state.FlatEligible Then
                    If coverage.ExactMetadataRecordsInspected + state.TotalExactRecords > SharedMethods.DEFAULT_SEMANTICARCHIVE_FLAT_METADATA_DOCUMENTS Then
                        state.FlatEligible = False
                        state.FlatDocuments.Clear()
                    ElseIf SemanticArchiveInventory.IsSearchable(document) Then
                        state.FlatDocuments.Add(New SemanticArchiveFlatDocument With {.ArchiveId = archiveId, .Document = document})
                    End If
                End If
                If Not SemanticArchiveInventory.IsSearchable(document) Then Continue While
                ' Local metadata matching reveals nothing. Source checks occur before a
                ' matching title/identifier can reach a model, candidate, or tool result.
                Dim score As System.Double = ExactMetadataScore(state.Query, document)
                If score <= 0 Then Continue While
                If Not _store.CanReadDocument(scope.AccessContext, generation, document) Then
                    coverage.SourcesUnavailable += 1
                    Continue While
                End If
                coverage.FilesInspected += 1
                AddCandidate(state, archiveId, document.DocumentId, System.Math.Min(0.95R, 0.5R + score * 0.1R), "Exact name or retained identifier/term match.", "exact_metadata")
            End While
            state.ExactComplete = state.ExactArchivePosition >= state.ArchiveIds.Count
        End Sub

        ''' <summary>Small catalogs rank their complete authorized document cards directly. No keyword-only prefilter.</summary>
        Private Async Function SearchSmallCatalogAsync(scope As SemanticArchiveRunScope, state As SemanticArchiveSearchState,
                    coverage As SemanticArchiveCoverage, cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of System.Boolean)
            If state.FlatDocuments.Count = 0 Then Return False
            Dim entries As New System.Collections.Generic.List(Of SharedMethods.SemanticSearchIndexEntry)()
            Dim documents As New System.Collections.Generic.Dictionary(Of System.String, SemanticArchiveFlatDocument)(System.StringComparer.OrdinalIgnoreCase)
            For Each item As SemanticArchiveFlatDocument In state.FlatDocuments
                cancellationToken.ThrowIfCancellationRequested()
                Dim document As SemanticArchiveDocumentRecord = item.Document
                Dim generation As SemanticArchiveGenerationManifest = scope.State.Generations(item.ArchiveId)
                ' Recheck before exposing metadata, not merely after selection.
                If document Is Nothing OrElse document.Card Is Nothing OrElse Not _store.CanReadDocument(scope.AccessContext, generation, document) Then
                    coverage.SourcesUnavailable += 1
                    Continue For
                End If
                Dim id As System.String = "S" & (entries.Count + 1).ToString("0000", System.Globalization.CultureInfo.InvariantCulture)
                entries.Add(EntryFromCard(document.Card, id))
                documents.Add(id, item)
            Next
            If entries.Count = 0 Then Return False
            coverage.RetrievalStrategy = "flat_document_cards"
            Dim options As SharedMethods.SemanticSearchRetrievalOptions = SelectionOptions(System.Math.Min(entries.Count, System.Math.Min(state.Budgets.MaxSectionCandidates, System.Math.Min(50, state.Budgets.MaxCandidateFiles))), state.Budgets, coverage)
            ' The existing selector handles complete-record character/token batching and the configured model-call cap.
            Dim selection As SharedMethods.SemanticSearchSelectionResult = Await SharedMethods.SelectSemanticSearchEntriesAsync(
                _context, state.Query, entries, options, cancellationToken).ConfigureAwait(False)
            coverage.CardsConsidered += selection.CandidatesConsidered
            Dim allRetained As System.Boolean = True
            For Each choice As SharedMethods.SemanticSearchSelectedEntryResult In selection.SelectedEntries
                Dim item As SemanticArchiveFlatDocument = Nothing
                If Not documents.TryGetValue(choice.Id, item) Then Throw New System.IO.InvalidDataException("Document-card selection returned an unknown ID.")
                Dim generation As SemanticArchiveGenerationManifest = scope.State.Generations(item.ArchiveId)
                coverage.FilesInspected += 1
                If Not _store.CanReadDocument(scope.AccessContext, generation, item.Document) Then
                    coverage.SourcesUnavailable += 1
                    Continue For
                End If
                If Not AddCandidate(state, item.ArchiveId, item.Document.DocumentId, choice.Relevance, choice.Reason, "flat_semantic") Then allRetained = False
            Next
            coverage.DocumentMetadataSelectionComplete = selection.CoverageComplete AndAlso Not selection.PotentiallyMissingInformation AndAlso allRetained
            If coverage.DocumentMetadataSelectionComplete Then
                ' Ranking every document substitutes for container navigation, not for exact evidence reading.
                state.Queue.Clear()
            Else
                coverage.Diagnostics.Add("flat_metadata_partial: Additional semantic coverage is available through the retained hierarchy continuation.")
            End If
            Return True
        End Function

        Private Async Function SearchNodeAsync(scope As SemanticArchiveRunScope, state As SemanticArchiveSearchState,
                        work As SemanticArchiveNodeWork, coverage As SemanticArchiveCoverage,
                        cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            Dim generation As SemanticArchiveGenerationManifest = scope.State.Generations(work.ArchiveId)
            Dim node As SemanticArchiveNode = _store.LoadNode(generation, work.NodeId)
            Dim entries As New System.Collections.Generic.List(Of SharedMethods.SemanticSearchIndexEntry)()
            Dim targets As New System.Collections.Generic.Dictionary(Of System.String, SemanticArchiveCard)(System.StringComparer.OrdinalIgnoreCase)
            Dim neutralIds As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            Dim cardPosition As System.Int32 = 0
            For Each card As SemanticArchiveCard In node.Cards
                cancellationToken.ThrowIfCancellationRequested()
                If card Is Nothing Then Throw New System.IO.InvalidDataException("A node contains an invalid card.")
                cardPosition += 1
                Dim id As System.String = "S" & cardPosition.ToString("0000", System.Globalization.CultureInfo.InvariantCulture)
                If work.ProcessedCardIds.Contains(id) Then Continue For
                Dim isContainer As System.Boolean = System.String.Equals(card.Level, "CONTAINER", System.StringComparison.OrdinalIgnoreCase)
                Dim remainingParentChecks As System.Int64 = System.Math.Max(0L, 2048L - (scope.AccessContext.SourceChecks - state.SourceChecksAtCallStart))
                Dim mayExpose As System.Boolean = _store.CanExposeCard(scope.AccessContext, generation, card, CInt(System.Math.Min(256L, remainingParentChecks \ 2L)))
                If Not mayExpose AndAlso Not isContainer Then
                    coverage.SourcesUnavailable += 1
                    work.ProcessedCardIds.Add(id)
                    Continue For
                End If
                Dim entry As SharedMethods.SemanticSearchIndexEntry
                If Not mayExpose Then
                    ' A host-resolved structural edge can remain traversable while its
                    ' revoked descendant summaries are never passed to a model.
                    entry = New SharedMethods.SemanticSearchIndexEntry() With {.Id = id, .Title = "Navigation branch", .Summary = "Permission-neutral branch; inspect authorized children for relevant evidence."}
                    neutralIds.Add(id)
                    coverage.NeutralNavigationBranches += 1
                    If Not coverage.Diagnostics.Contains("permission_neutral_navigation: Some branch summaries were suppressed because current descendant authorization was denied, unknown, or exceeded its bounded check budget. Widening remains available; semantic coverage is reduced.") Then
                        coverage.Diagnostics.Add("permission_neutral_navigation: Some branch summaries were suppressed because current descendant authorization was denied, unknown, or exceeded its bounded check budget. Widening remains available; semantic coverage is reduced.")
                    End If
                Else
                    entry = EntryFromCard(card, id)
                End If
                entries.Add(entry)
                targets.Add(id, card)
            Next
            If entries.Count = 0 Then Return
            Dim groups As System.Collections.Generic.List(Of System.Collections.Generic.List(Of SharedMethods.SemanticSearchIndexEntry)) = CreateGroups(entries, state.Budgets)
            Dim selected As New System.Collections.Generic.Dictionary(Of System.String, SharedMethods.SemanticSearchSelectedEntryResult)(System.StringComparer.OrdinalIgnoreCase)
            Dim widenLeaf As System.Boolean = False
            For Each group As System.Collections.Generic.List(Of SharedMethods.SemanticSearchIndexEntry) In groups
                If coverage.ModelCalls >= state.Budgets.MaxModelCalls Then
                    state.SeenNodes.Remove(Identity(work.ArchiveId, work.NodeId))
                    state.Queue.Add(work)
                    coverage.Diagnostics.Add("model_call_budget: Additional node-card groups remain eligible on continuation.")
                    Exit For
                End If
                Dim selection As SharedMethods.SemanticSearchSelectionResult = Await SharedMethods.SelectSemanticSearchEntriesAsync(
                    _context, state.Query, group, SelectionOptions(If(node.IsLeaf, 8, state.Budgets.InitialBranches), state.Budgets, coverage), cancellationToken).ConfigureAwait(False)
                coverage.CardsConsidered += selection.CandidatesConsidered
                Dim chosenIds As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
                For Each choice As SharedMethods.SemanticSearchSelectedEntryResult In selection.SelectedEntries
                    chosenIds.Add(choice.Id)
                Next
                For considered As System.Int32 = 0 To System.Math.Min(selection.CandidatesConsidered, group.Count) - 1
                    If node.IsLeaf AndAlso selection.PotentiallyMissingInformation AndAlso chosenIds.Count > 0 AndAlso Not chosenIds.Contains(group(considered).Id) Then
                        ' A result limit cannot permanently discard the remaining
                        ' relevant file cards from a leaf. Revisit without prior hits.
                        widenLeaf = True
                    Else
                        work.ProcessedCardIds.Add(group(considered).Id)
                    End If
                Next
                state.NeedsWidening = state.NeedsWidening OrElse selection.PotentiallyMissingInformation
                If Not selection.CoverageComplete Then
                    coverage.Diagnostics.Add("selection_coverage_incomplete: Not every candidate in a metadata group was evaluated; remaining cards stay eligible on continuation.")
                    state.SeenNodes.Remove(Identity(work.ArchiveId, work.NodeId))
                    state.Queue.Add(work)
                End If
                For Each choice As SharedMethods.SemanticSearchSelectedEntryResult In selection.SelectedEntries
                    If targets.ContainsKey(choice.Id) Then selected(choice.Id) = choice
                Next
                If Not selection.CoverageComplete Then Exit For
            Next
            For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, SemanticArchiveCard) In targets
                Dim card As SemanticArchiveCard = pair.Value
                Dim choice As SharedMethods.SemanticSearchSelectedEntryResult = Nothing
                selected.TryGetValue(pair.Key, choice)
                If System.String.Equals(card.Level, "CONTAINER", System.StringComparison.OrdinalIgnoreCase) Then
                    ' Selected branches are expanded first. Every other valid branch
                    ' remains eligible for deterministic bounded widening.
                    Enqueue(state, work.ArchiveId, card.TargetId, If(choice Is Nothing, If(neutralIds.Contains(pair.Key), 0.25R, 0.0R), 10.0R + choice.Relevance))
                ElseIf System.String.Equals(card.Level, "DOCUMENT", System.StringComparison.OrdinalIgnoreCase) AndAlso choice IsNot Nothing Then
                    Dim document As SemanticArchiveDocumentRecord = _store.LoadDocument(generation, card.TargetId)
                    coverage.FilesInspected += 1
                    If document IsNot Nothing AndAlso _store.CanReadDocument(scope.AccessContext, generation, document) Then
                        If Not AddCandidate(state, work.ArchiveId, card.TargetId, choice.Relevance, choice.Reason, "semantic") Then
                            ' Candidate capacity never discards a selected file card.
                            ' Retain its host-owned leaf work for the next call.
                            work.ProcessedCardIds.Remove(pair.Key)
                            widenLeaf = True
                            state.NeedsWidening = True
                        End If
                    Else
                        coverage.SourcesUnavailable += 1
                    End If
                End If
            Next
            If widenLeaf Then
                Dim alreadyQueued As System.Boolean = False
                For Each pending As SemanticArchiveNodeWork In state.Queue
                    If System.Object.ReferenceEquals(pending, work) Then alreadyQueued = True
                Next
                If Not alreadyQueued Then
                    state.SeenNodes.Remove(Identity(work.ArchiveId, work.NodeId))
                    work.Priority = -1.0R
                    state.Sequence += 1
                    work.Sequence = state.Sequence
                    state.Queue.Add(work)
                End If
            End If
        End Function

        Private Shared Function EntryFromCard(card As SemanticArchiveCard, id As System.String) As SharedMethods.SemanticSearchIndexEntry
            Dim entry As SharedMethods.SemanticSearchIndexEntry = Nothing
            If card.Metadata IsNot Nothing Then
                entry = Newtonsoft.Json.JsonConvert.DeserializeObject(Of SharedMethods.SemanticSearchIndexEntry)(Newtonsoft.Json.JsonConvert.SerializeObject(card.Metadata))
            End If
            If entry Is Nothing Then
                entry = New SharedMethods.SemanticSearchIndexEntry() With {
                    .Title = card.Title, .Summary = card.Summary,
                    .Topics = New System.Collections.Generic.List(Of System.String)(card.Topics),
                    .UserIntents = New System.Collections.Generic.List(Of System.String)(card.UserIntents),
                    .Identifiers = New System.Collections.Generic.List(Of System.String)(card.Identifiers),
                    .ExactTerms = New System.Collections.Generic.List(Of System.String)(card.ExactTerms)}
            End If
            entry.Id = id
            entry.StableId = card.CardId
            Return entry
        End Function

        Private Shared Function CreateGroups(entries As System.Collections.Generic.IEnumerable(Of SharedMethods.SemanticSearchIndexEntry), budgets As SemanticArchiveRetrievalBudgets) As System.Collections.Generic.List(Of System.Collections.Generic.List(Of SharedMethods.SemanticSearchIndexEntry))
            Dim groups As New System.Collections.Generic.List(Of System.Collections.Generic.List(Of SharedMethods.SemanticSearchIndexEntry))()
            Dim current As New System.Collections.Generic.List(Of SharedMethods.SemanticSearchIndexEntry)()
            Dim characters As System.Int32 = 0
            For Each entry As SharedMethods.SemanticSearchIndexEntry In entries
                Dim size As System.Int32 = SharedMethods.BuildCompactSemanticSearchIndex(New SharedMethods.SemanticSearchIndexEntry() {entry}).Length
                If size > budgets.MaxPromptCharacters Then Throw New System.IO.InvalidDataException("oversized_card: A preserved card cannot fit the complete routing budget.")
                If current.Count > 0 AndAlso (current.Count >= budgets.MaxSectionCandidates OrElse characters + size > budgets.MaxPromptCharacters) Then
                    groups.Add(current)
                    current = New System.Collections.Generic.List(Of SharedMethods.SemanticSearchIndexEntry)()
                    characters = 0
                End If
                current.Add(entry)
                characters += size
            Next
            If current.Count > 0 Then groups.Add(current)
            Return groups
        End Function

        Private Shared Function SelectionOptions(maximumSelected As System.Int32, budgets As SemanticArchiveRetrievalBudgets, coverage As SemanticArchiveCoverage) As SharedMethods.SemanticSearchRetrievalOptions
            Return New SharedMethods.SemanticSearchRetrievalOptions() With {
                .CandidateSelectionMode = SharedMethods.SemanticSearchCandidateSelectionMode.AllAuthorizedEntries,
                .MinimumSelectedSegments = 1, .MaximumSelectedSegments = maximumSelected, .MaximumTotalSegments = System.Math.Max(8, maximumSelected),
                .MaximumCandidateEntries = budgets.MaxSectionCandidates, .MaximumCompactIndexCharacters = budgets.MaxPromptCharacters,
                .MaximumRequestTokens = budgets.MaxRequestTokens,
                .MaximumSelectionModelCalls = System.Math.Max(1, budgets.MaxModelCalls - coverage.ModelCalls),
                .SelectionModelCallObserver = Sub() coverage.ModelCalls += 1,
                .MaximumConversationCharacters = 1, .MaximumLlmAttempts = 1, .IncludePreviouslyUsedIds = False,
                .IncludeAdjacentToPreviouslyUsedIds = False, .EnableFullScanFallback = False, .SpecialTaskName = "SemanticSearch"}
        End Function

        Private Shared Sub Enqueue(state As SemanticArchiveSearchState, archiveId As System.String, nodeId As System.String, priority As System.Double)
            If System.String.IsNullOrWhiteSpace(nodeId) OrElse state.SeenNodes.Contains(Identity(archiveId, nodeId)) Then Return
            For Each existing As SemanticArchiveNodeWork In state.Queue
                If System.String.Equals(existing.ArchiveId, archiveId, System.StringComparison.Ordinal) AndAlso System.String.Equals(existing.NodeId, nodeId, System.StringComparison.Ordinal) Then
                    existing.Priority = System.Math.Max(existing.Priority, priority)
                    Return
                End If
            Next
            If state.Queue.Count >= 8192 Then Throw New System.InvalidOperationException("navigation_state_budget: The bounded retained navigation queue is full.")
            state.Sequence += 1
            state.Queue.Add(New SemanticArchiveNodeWork() With {.ArchiveId = archiveId, .NodeId = nodeId, .Priority = priority, .Sequence = state.Sequence})
        End Sub

        Private Shared Function AddCandidate(state As SemanticArchiveSearchState, archiveId As System.String, documentId As System.String,
                        score As System.Double, reason As System.String, channel As System.String,
                        Optional literalMatchStartByte As System.Nullable(Of System.Int64) = Nothing) As System.Boolean
            Dim key As System.String = Identity(archiveId, documentId)
            Dim candidate As SemanticArchiveCandidate = Nothing
            If state.Returned.Contains(key) Then Return True
            If state.Candidates.TryGetValue(key, candidate) Then
                If score > candidate.Score Then candidate.Reason = reason
                candidate.Score = System.Math.Max(candidate.Score, score)
                If literalMatchStartByte.HasValue Then candidate.LiteralMatchStartByte = literalMatchStartByte
                If Not candidate.Channel.Contains(channel) Then candidate.Channel &= "+" & channel
                Return True
            End If
            If state.Candidates.Count >= state.Budgets.MaxCandidateFiles Then Return False
            state.Candidates.Add(key, New SemanticArchiveCandidate() With {
                .ArchiveId = archiveId, .DocumentId = documentId, .Score = score, .Reason = reason, .Channel = channel,
                .LiteralMatchStartByte = literalMatchStartByte})
            Return True
        End Function

        Private Shared Function ExactMetadataScore(query As System.String, document As SemanticArchiveDocumentRecord) As System.Double
            Dim fields As New System.Collections.Generic.List(Of System.String) From {document.DisplayName, document.RelativePath}
            If document.Card IsNot Nothing Then
                fields.Add(document.Card.Title)
                fields.AddRange(document.Card.Identifiers)
                fields.AddRange(document.Card.ExactTerms)
                If document.Card.Metadata IsNot Nothing Then
                    fields.AddRange(document.Card.Metadata.NamedEntities)
                    fields.AddRange(document.Card.Metadata.DefinedTerms)
                End If
            End If
            Dim tokens As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            For Each match As System.Text.RegularExpressions.Match In System.Text.RegularExpressions.Regex.Matches(query, "[\p{L}\p{N}][\p{L}\p{N}_./:-]{1,}", System.Text.RegularExpressions.RegexOptions.CultureInvariant, System.TimeSpan.FromMilliseconds(100))
                tokens.Add(match.Value)
            Next
            Dim score As System.Double = 0
            For Each field As System.String In fields
                If System.String.IsNullOrWhiteSpace(field) Then Continue For
                If field.IndexOf(query, System.StringComparison.OrdinalIgnoreCase) >= 0 Then score += 2.0R
                For Each token As System.String In tokens
                    If System.String.Equals(field, token, System.StringComparison.OrdinalIgnoreCase) Then
                        score += 1.0R
                    ElseIf field.IndexOf(token, System.StringComparison.OrdinalIgnoreCase) >= 0 Then
                        score += 0.25R
                    End If
                Next
            Next
            Return score
        End Function

        Private Shared Function CanUseExactMetadataFastPath(state As SemanticArchiveSearchState) As System.Boolean
            If state Is Nothing OrElse Not state.ExactComplete OrElse state.HadExactFailures OrElse state.LiteralScan IsNot Nothing Then Return False
            If state.Candidates.Count <> 1 Then Return False
            For Each candidate As SemanticArchiveCandidate In state.Candidates.Values
                Return candidate IsNot Nothing AndAlso
                    candidate.Score >= 0.75R AndAlso
                    candidate.Channel.IndexOf("exact_metadata", System.StringComparison.OrdinalIgnoreCase) >= 0
            Next
            Return False
        End Function

        Private Shared Function CountUnreturnedCandidates(state As SemanticArchiveSearchState) As System.Int32
            Dim count As System.Int32 = 0
            For Each key As System.String In state.Candidates.Keys
                If Not state.Returned.Contains(key) Then count += 1
            Next
            Return count
        End Function

        ''' <summary>One aggregate per-call budget, bounded by every selected archive and hard host limits.</summary>
        Private Function ResolveBudgets(scope As SemanticArchiveRunScope) As SemanticArchiveRetrievalBudgets
            Dim result As New SemanticArchiveRetrievalBudgets() With {
                .MaxNodesVisited = 256, .MaxModelCalls = 64, .MaxElapsedSeconds = 180,
                .MaxCandidateFiles = 512, .MaxEvidenceBytes = 262144, .InitialBranches = 16,
                .MaxExactLookupDocuments = 10000, .MaxSectionCandidates = 64, .MaxPromptCharacters = 64000, .MaxRequestTokens = 262144,
                .MaxLiteralScanBytes = 67108864, .MaxLiteralScanDocuments = 1024}
            For Each archiveId As System.String In scope.SelectedArchiveIds
                Dim definition As SemanticArchiveDefinition = _store.GetArchive(archiveId)
                If definition Is Nothing OrElse definition.RetrievalBudgets Is Nothing Then Throw New System.IO.InvalidDataException("A selected archive has no valid retrieval policy.")
                Dim source As SemanticArchiveRetrievalBudgets = definition.RetrievalBudgets
                result.MaxNodesVisited = LowerBudget(result.MaxNodesVisited, source.MaxNodesVisited)
                result.MaxModelCalls = LowerBudget(result.MaxModelCalls, source.MaxModelCalls)
                result.MaxElapsedSeconds = LowerBudget(result.MaxElapsedSeconds, source.MaxElapsedSeconds)
                result.MaxCandidateFiles = LowerBudget(result.MaxCandidateFiles, source.MaxCandidateFiles)
                result.MaxEvidenceBytes = LowerBudget(result.MaxEvidenceBytes, source.MaxEvidenceBytes)
                result.InitialBranches = LowerBudget(result.InitialBranches, source.InitialBranches)
                If source.MaxExactLookupDocuments < 0 Then Throw New System.IO.InvalidDataException("The exact lookup budget cannot be negative.")
                result.MaxExactLookupDocuments = System.Math.Min(result.MaxExactLookupDocuments, source.MaxExactLookupDocuments)
                result.MaxSectionCandidates = LowerBudget(result.MaxSectionCandidates, source.MaxSectionCandidates)
                result.MaxPromptCharacters = LowerBudget(result.MaxPromptCharacters, source.MaxPromptCharacters)
                result.MaxRequestTokens = LowerBudget(result.MaxRequestTokens, source.MaxRequestTokens)
                result.MaxLiteralScanBytes = LowerBudget(result.MaxLiteralScanBytes, source.MaxLiteralScanBytes)
                result.MaxLiteralScanDocuments = LowerBudget(result.MaxLiteralScanDocuments, source.MaxLiteralScanDocuments)
            Next
            Return result
        End Function

        Private Shared Function LowerBudget(current As System.Int32, configured As System.Int32) As System.Int32
            If configured < 1 Then Throw New System.IO.InvalidDataException("A configured archive retrieval budget must be positive.")
            Return System.Math.Min(current, configured)
        End Function

        Private Shared Function Identity(archiveId As System.String, localId As System.String) As System.String
            Return archiveId & ":" & localId
        End Function

        Private Shared Sub DisposeExactEnumerators(state As SemanticArchiveSearchState)
            For Each iterator As System.Collections.Generic.IEnumerator(Of SemanticArchiveDocumentRecord) In state.ExactEnumerators.Values
                iterator.Dispose()
            Next
            If state.LiteralScan IsNot Nothing Then
                For Each iterator As System.Collections.Generic.IEnumerator(Of SemanticArchiveDocumentRecord) In state.LiteralScan.Enumerators.Values
                    iterator.Dispose()
                Next
            End If
        End Sub
    End Class
End Namespace
