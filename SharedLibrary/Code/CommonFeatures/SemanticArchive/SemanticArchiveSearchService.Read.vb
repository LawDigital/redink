' Part of "Red Ink" (SharedLibrary)
' Exact original extracted evidence, with independent current-source authorization.

' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: SemanticArchiveSearchService.Read.vb
' Purpose:
'   Exact evidence excerpts, resumable document reads and disclosure-time authorization
'   revalidation.
'
' Architecture / Function:
'   Consumes opaque hit references from the scoped search and does not treat derived
'   visibility as a fresh source grant.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary
    Public NotInheritable Class SemanticArchiveExcerpt
        Public Property HitReference As System.String
        Public Property ArchiveId As System.String
        Public Property GenerationId As System.String
        Public Property DocumentId As System.String
        Public Property SourcePath As System.String
        Public Property SourceUri As System.String
        Public Property DisplayName As System.String
        Public Property RepresentationId As System.String
        Public Property OriginalSourceSha256 As System.String
        Public Property ExportedTextFileSha256 As System.String
        Public Property IndexedPayloadSha256 As System.String
        Public Property ExtractionCompleteness As System.String
        Public Property OffsetUnit As System.String = "utf-8 byte"
        Public Property OffsetBase As System.String
        Public Property StartByte As System.Int64
        Public Property LengthBytes As System.Int64
        Public Property EntryIds As New System.Collections.Generic.List(Of System.String)()
        Public Property OriginalSourceMapJson As System.String
        Public Property Text As System.String
        Public Property HasMore As System.Boolean
        Public Property IsExactExtractedText As System.Boolean = True
    End Class

    Public NotInheritable Class SemanticArchiveReadResult
        Public Property Status As System.String = "ok"
        Public Property Message As System.String = ""
        Public Property Excerpts As New System.Collections.Generic.List(Of SemanticArchiveExcerpt)()
        Public Property MoreHitReferences As New System.Collections.Generic.List(Of System.String)()
        Public Property Coverage As New SemanticArchiveCoverage()
    End Class

    Friend NotInheritable Class SemanticArchiveDocumentReadState
        Friend Property NextSectionIndex As System.Int32
        Friend Property PendingSections As New System.Collections.Generic.List(Of SharedMethods.SemanticSearchSelectedEntryResult)()
        Friend Property LoadedSections As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
        Friend Property SectionOffsets As New System.Collections.Generic.Dictionary(Of System.String, System.Int64)(System.StringComparer.OrdinalIgnoreCase)
        Friend Property DeferredSectionIds As New System.Collections.Generic.List(Of System.String)()
        Friend Property DeferredSectionSet As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
        Friend Property SmallFileNextByte As System.Int64
        Friend Property SmallFileComplete As System.Boolean
    End Class

    Partial Public NotInheritable Class SemanticArchiveSearchService
        Public Async Function ReadAsync(hitReferences As System.Collections.Generic.IEnumerable(Of System.String),
                        Optional maximumBytes As System.Int32 = SharedMethods.DEFAULT_SEMANTICARCHIVE_READ_MAXIMUM_BYTES,
                        Optional cancellationToken As System.Threading.CancellationToken = Nothing,
                        Optional maximumExcerpts As System.Int32 = SharedMethods.DEFAULT_SEMANTICARCHIVE_READ_MAXIMUM_EXCERPTS) As System.Threading.Tasks.Task(Of SemanticArchiveReadResult)
            If Not SemanticArchiveHostIntegration.IsConfigured(_context) Then Return New SemanticArchiveReadResult With {.Status = "not_configured", .Message = "Semantic Archive is disabled: the local catalog path is empty."}
            If Not System.String.IsNullOrWhiteSpace(_scope.AccessContext.DenialCode) Then
                Return New SemanticArchiveReadResult() With {.Status = _scope.AccessContext.DenialCode,
                    .Message = If(System.String.IsNullOrWhiteSpace(_scope.AccessContext.DenialMessage), "The host could not verify the requesting principal's source authorization.", _scope.AccessContext.DenialMessage)}
            End If
            If _scope.ResolutionStatus.Length > 0 Then
                Return New SemanticArchiveReadResult() With {.Status = _scope.ResolutionStatus, .Message = _scope.ResolutionMessage}
            End If
            If maximumBytes < 128 OrElse maximumBytes > 262144 Then
                Return New SemanticArchiveReadResult() With {.Status = "invalid_evidence_budget", .Message = "The exact-text budget must be between 128 and 262144 UTF-8 bytes."}
            End If
            If maximumExcerpts < 1 OrElse maximumExcerpts > 64 Then
                Return New SemanticArchiveReadResult() With {.Status = "invalid_excerpt_budget", .Message = "The provenance-bearing excerpt limit must be between 1 and 64."}
            End If
            Dim references As New System.Collections.Generic.List(Of System.String)()
            Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            If hitReferences IsNot Nothing Then
                For Each reference As System.String In hitReferences
                    If System.String.IsNullOrWhiteSpace(reference) Then Return New SemanticArchiveReadResult() With {.Status = "invalid_hit", .Message = "An opaque hit reference is required."}
                    If seen.Add(reference) Then references.Add(reference)
                    If references.Count > 16 Then Return New SemanticArchiveReadResult() With {.Status = "hit_limit", .Message = "Read at most 16 references per bounded call."}
                Next
            End If
            If references.Count = 0 Then Return New SemanticArchiveReadResult() With {.Status = "missing_hit", .Message = "Search first, then read one or more returned hit references."}
            Await _scope.State.OperationGate.WaitAsync(cancellationToken).ConfigureAwait(False)
            Try
                Using priority As System.IDisposable = BackgroundMaintenanceCoordinator.EnterInteractiveWork()
                    _scope.State.EnsureUsable()
                    _scope.State.BindDirectory(_store.DirectoryPath)
                    Dim checkpoints As New System.Collections.Generic.Dictionary(Of System.String, SemanticArchiveDocumentReadState)(System.StringComparer.Ordinal)
                    For Each reference As System.String In references
                        Dim hit As SemanticArchiveEvidenceReference = Nothing
                        If Not _scope.State.Hits.TryGetValue(reference, hit) OrElse hit Is Nothing OrElse Not _scope.ContainsArchive(hit.ArchiveId) Then
                            Return New SemanticArchiveReadResult() With {.Status = "invalid_hit", .Message = "A reference is unknown, belongs to another run, or lies outside the selected archive scope."}
                        End If
                        checkpoints.Add(reference, hit.ReadState)
                    Next
                    Dim loadedBytesBefore As System.Int64 = _scope.State.LoadedEvidenceBytes
                    Dim sourceChecksBefore As System.Int64 = _scope.AccessContext.SourceChecks
                    Dim timer As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew()
                    Dim budgets As SemanticArchiveRetrievalBudgets = ResolveBudgets(_scope)
                    Dim result As New SemanticArchiveReadResult()
                    Using deadline As System.Threading.CancellationTokenSource = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(cancellationToken)
                        deadline.CancelAfter(System.TimeSpan.FromSeconds(budgets.MaxElapsedSeconds))
                        Try
                            Dim completed As SemanticArchiveReadResult = Await System.Threading.Tasks.Task.Run(
                                Function() ReadCoreAsync(references, maximumBytes, maximumExcerpts, budgets, result, deadline.Token), deadline.Token).ConfigureAwait(False)
                            deadline.Token.ThrowIfCancellationRequested()
                            If Not System.String.IsNullOrWhiteSpace(_scope.AccessContext.DenialCode) Then
                                RestoreReadCheckpoints(checkpoints, loadedBytesBefore)
                                completed.Excerpts.Clear()
                                completed.MoreHitReferences.Clear()
                                completed.Coverage.EvidenceBytes = 0
                                completed.Status = _scope.AccessContext.DenialCode
                                completed.Message = If(System.String.IsNullOrWhiteSpace(_scope.AccessContext.DenialMessage), "The requesting principal's source authorization is no longer verified.", _scope.AccessContext.DenialMessage)
                            End If
                            Await System.Threading.Tasks.Task.Run(Sub() SemanticArchiveLibrary.ValidateScope(_context, _store, _scope.SelectedArchiveIds), cancellationToken).ConfigureAwait(False)
                            Return completed
                        Catch ex As System.OperationCanceledException When deadline.IsCancellationRequested AndAlso Not cancellationToken.IsCancellationRequested
                            ' No part of an interrupted call has been returned to its caller.
                            ' Restore every requested hit, including documents completed before
                            ' the deadline, so a bounded retry cannot skip undelivered evidence.
                            RestoreReadCheckpoints(checkpoints, loadedBytesBefore)
                            result.Excerpts.Clear()
                            result.MoreHitReferences.Clear()
                            result.MoreHitReferences.AddRange(references)
                            result.Status = "partial"
                            result.Message = "The elapsed-time budget ended this read. No excerpt from the interrupted call was consumed; read the retained hit references again to continue."
                            result.Coverage.EvidenceBytes = 0
                            result.Coverage.BudgetEndedSearch = True
                            result.Coverage.ElapsedMilliseconds = timer.ElapsedMilliseconds
                            result.Coverage.SourceAuthorizationChecks = _scope.AccessContext.SourceChecks - sourceChecksBefore
                            result.Coverage.Diagnostics.Add("read_elapsed_budget: The complete read deadline cancelled pending model work; undelivered evidence checkpoints and byte charges were restored.")
                            Return result
                        Catch
                            RestoreReadCheckpoints(checkpoints, loadedBytesBefore)
                            Throw
                        End Try
                    End Using
                End Using
            Finally
                _scope.State.OperationGate.Release()
            End Try
        End Function

        Private Sub RestoreReadCheckpoints(checkpoints As System.Collections.Generic.Dictionary(Of System.String, SemanticArchiveDocumentReadState),
                                           loadedBytesBefore As System.Int64)
            For Each checkpoint As System.Collections.Generic.KeyValuePair(Of System.String, SemanticArchiveDocumentReadState) In checkpoints
                Dim hit As SemanticArchiveEvidenceReference = Nothing
                If _scope.State.Hits.TryGetValue(checkpoint.Key, hit) AndAlso hit IsNot Nothing Then hit.ReadState = checkpoint.Value
            Next
            _scope.State.LoadedEvidenceBytes = loadedBytesBefore
        End Sub

        Private Async Function ReadCoreAsync(references As System.Collections.Generic.List(Of System.String), maximumBytes As System.Int32, maximumExcerpts As System.Int32,
                        budgets As SemanticArchiveRetrievalBudgets, result As SemanticArchiveReadResult,
                        cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of SemanticArchiveReadResult)
            SemanticArchiveLibrary.ValidateScope(_context, _store, _scope.SelectedArchiveIds)
            Dim timer As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew()
            Dim sourceChecksBefore As System.Int64 = _scope.AccessContext.SourceChecks
            Dim allowed As New System.Collections.Generic.List(Of System.Collections.Generic.KeyValuePair(Of System.String, SemanticArchiveEvidenceReference))()
            For Each reference As System.String In references
                Dim hit As SemanticArchiveEvidenceReference = Nothing
                If Not _scope.State.Hits.TryGetValue(reference, hit) OrElse hit Is Nothing OrElse Not _scope.ContainsArchive(hit.ArchiveId) Then
                    result.Status = "invalid_hit"
                    result.Message = "A reference is unknown, belongs to another run, or lies outside the selected archive scope."
                    Return result
                End If
                allowed.Add(New System.Collections.Generic.KeyValuePair(Of System.String, SemanticArchiveEvidenceReference)(reference, hit))
            Next
            maximumBytes = System.Math.Min(maximumBytes, budgets.MaxEvidenceBytes)
            If maximumBytes < 128 Then
                result.Status = "archive_evidence_budget"
                result.Message = "The selected archive policy permits fewer than the minimum 128 exact-text bytes per read. Increase that explicit policy before requesting evidence."
                result.Coverage.BudgetEndedSearch = True
                Return result
            End If
            Dim available As System.Int64 = System.Math.Min(CLng(maximumBytes), MaximumRunEvidenceBytes - _scope.State.LoadedEvidenceBytes)
            If available < 128 Then
                result.Status = "run_evidence_budget"
                result.Message = "This run has reached its aggregate evidence-byte budget. Start a new explicit retrieval run for additional evidence."
                result.Coverage.BudgetEndedSearch = True
                Return result
            End If
            Dim position As System.Int32 = 0
            Dim checkpoints As New System.Collections.Generic.Dictionary(Of System.String, SemanticArchiveDocumentReadState)(System.StringComparer.Ordinal)
            For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, SemanticArchiveEvidenceReference) In allowed
                cancellationToken.ThrowIfCancellationRequested()
                Dim hit As SemanticArchiveEvidenceReference = pair.Value
                If available < 128 OrElse result.Excerpts.Count >= maximumExcerpts OrElse timer.ElapsedMilliseconds >= CLng(budgets.MaxElapsedSeconds) * 1000L Then
                    result.MoreHitReferences.Add(pair.Key)
                    result.Coverage.BudgetEndedSearch = True
                    position += 1
                    Continue For
                End If
                Dim priorReadState As SemanticArchiveDocumentReadState = hit.ReadState
                Dim indexedReadRequired As System.Boolean = False
                checkpoints(pair.Key) = priorReadState
                Try
                    Dim generation As SemanticArchiveGenerationManifest = hit.Generation
                    Dim document As SemanticArchiveDocumentRecord = _store.LoadDocument(generation, hit.DocumentId)
                    ' The service account is never substituted for the scope's requester.
                    If document Is Nothing OrElse Not _store.CanReadDocument(_scope.AccessContext, generation, document, verifySourceHash:=True) Then
                        result.Coverage.SourcesUnavailable += 1
                        result.Coverage.Diagnostics.Add("source_unavailable: A requested source is no longer authorized, active, or at the pinned version.")
                        Continue For
                    End If
                    indexedReadRequired = document.Index IsNot Nothing
                    result.Coverage.Generations(hit.ArchiveId) = generation.GenerationId
                    result.Coverage.FilesInspected += 1
                    If Not System.String.Equals(document.Representation.Completeness, "complete", System.StringComparison.OrdinalIgnoreCase) Then result.Coverage.IncompleteExtractions += 1
                    hit.ReadState = CopyReadState(priorReadState)
                    If priorReadState Is Nothing AndAlso document.Index Is Nothing AndAlso hit.LiteralMatchStartByte.HasValue Then
                        hit.ReadState.SmallFileNextByte = hit.LiteralMatchStartByte.Value
                    End If
                    Dim allowance As System.Int32 = CInt(System.Math.Max(128L, available \ System.Math.Max(1, allowed.Count - position)))
                    Dim excerpts As System.Collections.Generic.List(Of SemanticArchiveExcerpt)
                    If document.Index Is Nothing Then
                        excerpts = ReadSmallFile(pair.Key, hit, document, allowance, result.Coverage)
                    Else
                        Dim excerptAllowance As System.Int32 = System.Math.Max(1, (maximumExcerpts - result.Excerpts.Count) \ System.Math.Max(1, allowed.Count - position))
                        excerpts = Await ReadIndexedFileAsync(pair.Key, hit, document, allowance, excerptAllowance, result.Coverage, budgets, timer, cancellationToken).ConfigureAwait(False)
                    End If
                    ' A revoke during model selection or artifact IO removes all content,
                    ' including names and paths, from this document's returned block.
                    If Not _store.CanReadDocument(_scope.AccessContext, generation, document, verifySourceHash:=True) Then
                        hit.ReadState = priorReadState
                        result.Coverage.SourcesUnavailable += 1
                        result.Coverage.Diagnostics.Add("source_changed_during_read: A source was suppressed before evidence publication.")
                        Continue For
                    End If
                    ' Validate the whole document result before advancing any shared
                    ' byte counters. A later invalid excerpt must not leave partial charges
                    ' or returned fragments paired with a restored document checkpoint.
                    Dim documentBytes As System.Int64 = 0
                    For Each excerpt As SemanticArchiveExcerpt In excerpts
                        If excerpt Is Nothing OrElse excerpt.LengthBytes <= 0 OrElse excerpt.LengthBytes > available - documentBytes Then Throw New System.IO.InvalidDataException("An exact excerpt exceeded the remaining evidence budget.")
                        documentBytes += excerpt.LengthBytes
                    Next
                    If excerpts.Count > 0 Then
                        SyncLock _scope.State.UsedSourcePaths
                            _scope.State.UsedSourcePaths.Add(document.SourcePath)
                        End SyncLock
                    End If
                    For Each excerpt As SemanticArchiveExcerpt In excerpts
                        result.Excerpts.Add(excerpt)
                        available -= excerpt.LengthBytes
                        _scope.State.LoadedEvidenceBytes += excerpt.LengthBytes
                        result.Coverage.EvidenceBytes += excerpt.LengthBytes
                        If excerpt.HasMore AndAlso Not result.MoreHitReferences.Contains(pair.Key) Then result.MoreHitReferences.Add(pair.Key)
                    Next
                    If excerpts.Count = 0 AndAlso document.Index IsNot Nothing AndAlso
                       (hit.ReadState.NextSectionIndex < document.Index.EntryCount OrElse hit.ReadState.PendingSections.Count > 0 OrElse hit.ReadState.DeferredSectionIds.Count > 0) Then result.MoreHitReferences.Add(pair.Key)
                Catch ex As System.OperationCanceledException
                    hit.ReadState = priorReadState
                    Throw
                Catch ex As System.Exception
                    hit.ReadState = priorReadState
                    System.Diagnostics.Debug.WriteLine("SA exact read failed: " & ex.GetType().FullName)
                    result.Coverage.SourcesUnavailable += 1
                    If indexedReadRequired Then
                        result.Coverage.IndexedReadFailures += 1
                        result.Coverage.Diagnostics.Add("indexed_read_failed: The required document index, section selection, or selected byte range could not be validated. No plaintext fallback was returned.")
                    Else
                        result.Coverage.Diagnostics.Add("read_failed: A source, artifact, or range could not be validated; no substitute text was returned.")
                    End If
                Finally
                    position += 1
                End Try
            Next
            RevalidateReadForDisclosure(result, checkpoints)
            result.Coverage.SourceAuthorizationChecks = _scope.AccessContext.SourceChecks - sourceChecksBefore
            result.Coverage.ElapsedMilliseconds = timer.ElapsedMilliseconds
            result.Coverage.BudgetEndedSearch = result.Coverage.BudgetEndedSearch OrElse result.MoreHitReferences.Count > 0
            If result.Coverage.IncompleteExtractions > 0 Then result.Coverage.Diagnostics.Add("incomplete_extraction: The returned bytes are exact extracted text, but at least one source has partial or unknown extraction completeness.")
            If result.Coverage.Diagnostics.Count > 0 OrElse result.Coverage.SourcesUnavailable > 0 Then result.Status = "partial"
            If result.Excerpts.Count = 0 Then result.Message = "No exact excerpt was loaded in this call. Inspect coverage and retained hit references before drawing a conclusion."
            Return result
        End Function

        Private Function ReadSmallFile(reference As System.String, hit As SemanticArchiveEvidenceReference, document As SemanticArchiveDocumentRecord,
                         allowance As System.Int32, coverage As SemanticArchiveCoverage) As System.Collections.Generic.List(Of SemanticArchiveExcerpt)
            Dim result As New System.Collections.Generic.List(Of SemanticArchiveExcerpt)()
            If hit.ReadState.SmallFileComplete Then Return result
            Dim encodingName As System.String = If(document.Representation.EncodingName, System.String.Empty).Trim()
            Dim requiresPreamble As System.Boolean = System.String.Equals(encodingName, "utf-8-bom", System.StringComparison.OrdinalIgnoreCase)
            If Not requiresPreamble AndAlso Not System.String.Equals(encodingName, "utf-8", System.StringComparison.OrdinalIgnoreCase) AndAlso
               Not System.String.Equals(encodingName, "utf8", System.StringComparison.OrdinalIgnoreCase) Then Throw New System.IO.InvalidDataException("The pinned representation is not UTF-8.")
            Dim path As System.String = _store.ValidateTextPath(hit.Generation, document, _scope.AccessContext)
            Dim start As System.Int64 = hit.ReadState.SmallFileNextByte
            Dim text As System.String = ""
            Dim byteLength As System.Int32
            Dim payloadLength As System.Int64
            Using stream As New System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                ' Verify exact file bytes while holding the same read handle used below.
                Using hasher As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                    Dim actual As System.String = System.BitConverter.ToString(hasher.ComputeHash(stream)).Replace("-", "").ToLowerInvariant()
                    If Not System.String.Equals(actual, document.Representation.TextFileHash, System.StringComparison.OrdinalIgnoreCase) Then Throw New System.IO.InvalidDataException("The immutable exported text hash differs from its representation.")
                End Using
                stream.Position = 0
                Dim preamble As System.Int32 = 0
                If stream.Length >= 3 Then
                    If stream.ReadByte() = &HEF AndAlso stream.ReadByte() = &HBB AndAlso stream.ReadByte() = &HBF Then preamble = 3
                End If
                If requiresPreamble AndAlso preamble <> 3 Then Throw New System.IO.InvalidDataException("The pinned UTF-8 BOM representation is missing its declared preamble.")
                payloadLength = stream.Length - preamble
                If start < 0 OrElse start > payloadLength Then Throw New System.IO.InvalidDataException("The stored read offset is invalid.")
                stream.Position = preamble + start
                Dim count As System.Int32 = CInt(System.Math.Min(CLng(allowance) + 3L, payloadLength - start))
                Dim bytes(count - 1) As System.Byte
                Dim read As System.Int32 = 0
                While read < count
                    Dim amount As System.Int32 = stream.Read(bytes, read, count - read)
                    If amount = 0 Then Throw New System.IO.EndOfStreamException("The immutable text ended inside a validated range.")
                    read += amount
                End While
                byteLength = System.Math.Min(allowance, read)
                While byteLength > 0 AndAlso byteLength < read AndAlso (bytes(byteLength) And &HC0) = &H80
                    byteLength -= 1
                End While
                text = New System.Text.UTF8Encoding(False, True).GetString(bytes, 0, byteLength)
            End Using
            hit.ReadState.SmallFileNextByte += byteLength
            hit.ReadState.SmallFileComplete = hit.ReadState.SmallFileNextByte >= payloadLength
            If byteLength = 0 Then Return result
            Dim excerpt As SemanticArchiveExcerpt = MakeExcerpt(reference, hit, document, coverage)
            excerpt.OffsetBase = "exported_text_content_without_utf8_bom"
            excerpt.StartByte = start
            excerpt.LengthBytes = byteLength
            excerpt.Text = text
            excerpt.HasMore = Not hit.ReadState.SmallFileComplete
            result.Add(excerpt)
            Return result
        End Function

        Private Async Function ReadIndexedFileAsync(reference As System.String, hit As SemanticArchiveEvidenceReference,
                         document As SemanticArchiveDocumentRecord, allowance As System.Int32, maximumExcerpts As System.Int32, coverage As SemanticArchiveCoverage,
                         budgets As SemanticArchiveRetrievalBudgets, timer As System.Diagnostics.Stopwatch,
                         cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of System.Collections.Generic.List(Of SemanticArchiveExcerpt))
            Dim result As New System.Collections.Generic.List(Of SemanticArchiveExcerpt)()
            coverage.IndexedFilesInspected += 1
            coverage.SectionIndexesQueried += 1
            Dim indexPath As System.String = _store.ValidateIndexPath(hit.Generation, document, _scope.AccessContext)
            Dim cache As SharedMethods.SemanticSearchIndexCacheItem = Await SharedMethods.TryGetSemanticSearchIndexAsync(indexPath, cancellationToken).ConfigureAwait(False)
            If cache Is Nothing OrElse cache.IndexDocument Is Nothing OrElse cache.OrderedEntries.Count <> document.Index.EntryCount OrElse
                Not System.String.Equals(cache.IndexDocument.ContentSha256, document.Index.PayloadHash, System.StringComparison.OrdinalIgnoreCase) Then
                Throw New System.IO.InvalidDataException("The section index payload is not the pinned validated payload.")
            End If
            coverage.TotalSections += cache.OrderedEntries.Count
            If hit.LiteralMatchStartByte.HasValue Then
                Dim literalLength As System.Int64 = LiteralLengthBytes(hit)
                If hit.LiteralMatchStartByte.Value < 0 OrElse hit.LiteralMatchStartByte.Value > cache.ContentByteLength - literalLength Then
                    Throw New System.IO.InvalidDataException("The literal relevance hint lies outside the pinned indexed payload.")
                End If
            End If
            Dim readState As SemanticArchiveDocumentReadState = hit.ReadState
            While (readState.NextSectionIndex < cache.OrderedEntries.Count OrElse readState.DeferredSectionIds.Count > 0) AndAlso coverage.ModelCalls < budgets.MaxModelCalls AndAlso
                  timer.ElapsedMilliseconds < CLng(budgets.MaxElapsedSeconds) * 1000L AndAlso readState.PendingSections.Count < System.Math.Min(24, maximumExcerpts)
                cancellationToken.ThrowIfCancellationRequested()
                If Not _store.CanReadDocument(_scope.AccessContext, hit.Generation, document) Then Throw New System.UnauthorizedAccessException("Source access was revoked before section metadata exposure.")
                Dim group As New System.Collections.Generic.List(Of SharedMethods.SemanticSearchIndexEntry)()
                Dim characters As System.Int32 = 0
                Dim fromDeferred As System.Boolean = readState.NextSectionIndex >= cache.OrderedEntries.Count
                Dim cursor As System.Int32 = If(fromDeferred, 0, readState.NextSectionIndex)
                Dim availableEntries As System.Int32 = If(fromDeferred, readState.DeferredSectionIds.Count, cache.OrderedEntries.Count)
                While cursor < availableEntries AndAlso group.Count < budgets.MaxSectionCandidates
                    Dim entry As SharedMethods.SemanticSearchIndexEntry = If(fromDeferred, cache.EntriesById(readState.DeferredSectionIds(cursor)), cache.OrderedEntries(cursor))
                    Dim size As System.Int32 = SharedMethods.BuildCompactSemanticSearchIndex(New SharedMethods.SemanticSearchIndexEntry() {entry}).Length
                    If size > budgets.MaxPromptCharacters Then Throw New System.IO.InvalidDataException("oversized_section_card: A complete section card exceeds the request budget.")
                    If group.Count > 0 AndAlso characters + size > budgets.MaxPromptCharacters Then Exit While
                    group.Add(entry)
                    characters += size
                    cursor += 1
                End While
                Dim options As SharedMethods.SemanticSearchRetrievalOptions = SelectionOptions(8, budgets, coverage)
                options.MaximumSelectionModelCalls = 1
                Dim selection As SharedMethods.SemanticSearchSelectionResult = Await SharedMethods.SelectSemanticSearchEntriesAsync(_context, SectionSelectionQuestion(hit, group), group, options, cancellationToken).ConfigureAwait(False)
                coverage.SectionsConsidered += selection.CandidatesConsidered
                If fromDeferred Then
                    For considered As System.Int32 = 0 To selection.CandidatesConsidered - 1
                        readState.DeferredSectionSet.Remove(readState.DeferredSectionIds(0))
                        readState.DeferredSectionIds.RemoveAt(0)
                    Next
                Else
                    readState.NextSectionIndex += selection.CandidatesConsidered
                End If
                Dim chosenIds As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
                For Each chosen As SharedMethods.SemanticSearchSelectedEntryResult In selection.SelectedEntries
                    chosenIds.Add(chosen.Id)
                    If cache.EntriesById.ContainsKey(chosen.Id) AndAlso Not readState.LoadedSections.Contains(chosen.Id) Then readState.PendingSections.Add(chosen)
                Next
                If selection.PotentiallyMissingInformation OrElse Not selection.CoverageComplete Then
                    For considered As System.Int32 = 0 To System.Math.Min(selection.CandidatesConsidered, group.Count) - 1
                        Dim id As System.String = group(considered).Id
                        If Not chosenIds.Contains(id) AndAlso Not selection.RejectedEntryIds.Contains(id) AndAlso Not readState.LoadedSections.Contains(id) AndAlso readState.DeferredSectionSet.Add(id) Then readState.DeferredSectionIds.Add(id)
                    Next
                End If
                If selection.CandidatesConsidered = 0 OrElse (selection.PotentiallyMissingInformation AndAlso chosenIds.Count = 0 AndAlso selection.RejectedEntryIds.Count = 0) OrElse
                   (fromDeferred AndAlso readState.PendingSections.Count = 0 AndAlso selection.RejectedEntryIds.Count = 0) Then
                    ' An incomplete empty selection is not an irrelevant-section proof.
                    ' Defer it without repeating the identical request in this read.
                    coverage.Diagnostics.Add("section_selection_budget: Remaining section metadata is retained for a subsequent read.")
                    Exit While
                End If
            End While
            readState.PendingSections.Sort(
                Function(left As SharedMethods.SemanticSearchSelectedEntryResult, right As SharedMethods.SemanticSearchSelectedEntryResult) As System.Int32
                    Dim relevance As System.Int32 = right.Relevance.CompareTo(left.Relevance)
                    If relevance <> 0 Then Return relevance
                    Dim sourceOrder As System.Int32 = cache.EntriesById(left.Id).StartByte.CompareTo(cache.EntriesById(right.Id).StartByte)
                    If sourceOrder <> 0 Then Return sourceOrder
                    Return System.StringComparer.OrdinalIgnoreCase.Compare(left.Id, right.Id)
                End Function)
            Dim remaining As System.Int32 = allowance
            While readState.PendingSections.Count > 0 AndAlso remaining >= 128 AndAlso result.Count < maximumExcerpts
                cancellationToken.ThrowIfCancellationRequested()
                Dim chosen As SharedMethods.SemanticSearchSelectedEntryResult = readState.PendingSections(0)
                Dim entry As SharedMethods.SemanticSearchIndexEntry = cache.EntriesById(chosen.Id)
                Dim consumed As System.Int64 = 0
                If Not readState.SectionOffsets.TryGetValue(entry.Id, consumed) AndAlso hit.LiteralMatchStartByte.HasValue AndAlso
                    LiteralOverlapsEntry(hit, entry) Then
                    ' The selector has already authorized this existing section ID.
                    ' Center its bounded exact read on the known literal position;
                    ' no unselected neighbor or arbitrary plaintext range is added.
                    consumed = System.Math.Max(0L, hit.LiteralMatchStartByte.Value - entry.StartByte)
                End If
                Dim source As SharedMethods.SemanticSearchLoadedSourceSegment = Nothing
                ' The shared exact-range adapter uses the original validated v1 UTF-8
                ' reader. It also handles sections spanning document-wrapper records
                ' without assuming one S entry corresponds to one source segment.
                source = Await SharedMethods.LoadSemanticSearchByteRangeAsync(indexPath, entry.StartByte + consumed,
                    CInt(System.Math.Min(CLng(remaining), entry.LengthBytes - consumed)), cancellationToken).ConfigureAwait(False)
                If source IsNot Nothing Then source.EntryIds = New System.Collections.Generic.List(Of System.String) From {entry.Id}
                If source Is Nothing OrElse source.LengthBytes <= 0 OrElse source.LengthBytes > remaining OrElse
                   source.RelativeStartByte < entry.StartByte OrElse source.RelativeStartByte + source.LengthBytes > entry.StartByte + entry.LengthBytes Then
                    Throw New System.IO.InvalidDataException("The loaded section lies outside the authorized bounded range.")
                End If
                consumed = source.RelativeStartByte + source.LengthBytes - entry.StartByte
                readState.SectionOffsets(entry.Id) = consumed
                If consumed >= entry.LengthBytes Then
                    readState.LoadedSections.Add(entry.Id)
                    readState.PendingSections.RemoveAt(0)
                    readState.SectionOffsets.Remove(entry.Id)
                End If
                Dim excerpt As SemanticArchiveExcerpt = MakeExcerpt(reference, hit, document, coverage)
                excerpt.OffsetBase = "indexed_text_content"
                excerpt.StartByte = source.RelativeStartByte
                excerpt.LengthBytes = source.LengthBytes
                excerpt.Text = source.Text
                excerpt.EntryIds = New System.Collections.Generic.List(Of System.String)(source.EntryIds)
                excerpt.HasMore = readState.NextSectionIndex < cache.OrderedEntries.Count OrElse readState.PendingSections.Count > 0 OrElse readState.DeferredSectionIds.Count > 0
                result.Add(excerpt)
                remaining -= CInt(source.LengthBytes)
            End While
            If readState.NextSectionIndex < cache.OrderedEntries.Count Then coverage.Diagnostics.Add("section_coverage_bounded: Additional section cards remain eligible; read this hit again to continue.")
            Dim hasMore As System.Boolean = readState.NextSectionIndex < cache.OrderedEntries.Count OrElse readState.PendingSections.Count > 0 OrElse readState.DeferredSectionIds.Count > 0
            If hit.LiteralMatchStartByte.HasValue AndAlso readState.NextSectionIndex >= cache.OrderedEntries.Count AndAlso readState.DeferredSectionIds.Count = 0 Then
                Dim selectedIds As New System.Collections.Generic.HashSet(Of System.String)(readState.LoadedSections, System.StringComparer.OrdinalIgnoreCase)
                For Each pending As SharedMethods.SemanticSearchSelectedEntryResult In readState.PendingSections
                    selectedIds.Add(pending.Id)
                Next
                For Each entry As SharedMethods.SemanticSearchIndexEntry In cache.OrderedEntries
                    If LiteralOverlapsEntry(hit, entry) AndAlso Not selectedIds.Contains(entry.Id) Then
                        coverage.Diagnostics.Add("literal_section_not_selected: At least one section overlapping the literal was not selected. Only the selected section ranges can be returned as evidence.")
                        Exit For
                    End If
                Next
            End If
            If Not hasMore AndAlso result.Count = 0 AndAlso readState.LoadedSections.Count = 0 Then
                coverage.IndexedFilesWithoutSelection += 1
                coverage.Diagnostics.Add("section_selection_empty: No eligible document section was selected. The document index was respected and no whole-text fallback was loaded.")
            End If
            If hasMore AndAlso result.Count >= maximumExcerpts Then coverage.Diagnostics.Add("excerpt_count_budget: Additional exact sections remain available by reading the retained hit reference again.")
            For Each excerpt As SemanticArchiveExcerpt In result
                excerpt.HasMore = hasMore
            Next
            Return result
        End Function

        Private Shared Function LiteralLengthBytes(hit As SemanticArchiveEvidenceReference) As System.Int64
            If System.String.IsNullOrEmpty(hit.LiteralText) Then Return 1L
            Return New System.Text.UTF8Encoding(False, True).GetByteCount(hit.LiteralText)
        End Function

        Private Shared Function LiteralOverlapsEntry(hit As SemanticArchiveEvidenceReference, entry As SharedMethods.SemanticSearchIndexEntry) As System.Boolean
            If Not hit.LiteralMatchStartByte.HasValue Then Return False
            Dim start As System.Int64 = hit.LiteralMatchStartByte.Value
            Dim length As System.Int64 = LiteralLengthBytes(hit)
            If start < 0 OrElse length > System.Int64.MaxValue - start Then Throw New System.IO.InvalidDataException("The literal byte-range hint is invalid.")
            Return entry.StartByte < start + length AndAlso start < entry.StartByte + entry.LengthBytes
        End Function

        Private Shared Function SectionSelectionQuestion(hit As SemanticArchiveEvidenceReference,
                         entries As System.Collections.Generic.IEnumerable(Of SharedMethods.SemanticSearchIndexEntry)) As System.String
            If Not hit.LiteralMatchStartByte.HasValue Then Return hit.Query
            Dim mapped As New System.Collections.Generic.List(Of System.String)()
            For Each entry As SharedMethods.SemanticSearchIndexEntry In entries
                If LiteralOverlapsEntry(hit, entry) Then mapped.Add(entry.Id)
            Next
            ' Existing v1 generation preserves the decoded UTF-8 input byte-for-byte
            ' without its BOM. Literal and entry offsets therefore share that payload
            ' coordinate system, including any document-wrapper bytes. The structured
            ' hint influences relevance only; every entry still needs model selection.
            Return Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                .question = hit.Query,
                .literal_relevance_hint = New With {
                    .text = If(hit.LiteralText, System.String.Empty),
                    .offset_base = "indexed_text_content_utf8_without_bom",
                    .start_byte = hit.LiteralMatchStartByte.Value, .length_bytes = LiteralLengthBytes(hit),
                    .overlapping_candidate_ids = mapped, .requires_section_selection = True}})
        End Function

        Private Shared Function MakeExcerpt(reference As System.String, hit As SemanticArchiveEvidenceReference,
                          document As SemanticArchiveDocumentRecord, coverage As SemanticArchiveCoverage) As SemanticArchiveExcerpt
            Dim sourceMap As System.String = document.Representation.SourceMapJson
            If sourceMap IsNot Nothing AndAlso sourceMap.Length > 4096 Then
                coverage.Diagnostics.Add("source_map_budget: The complete source map exceeds this response budget; exact extracted-text offsets are supplied.")
                sourceMap = ""
            End If
            Return New SemanticArchiveExcerpt() With {
                .HitReference = reference, .ArchiveId = hit.ArchiveId, .GenerationId = hit.Generation.GenerationId,
                .DocumentId = document.DocumentId, .SourcePath = document.SourcePath, .SourceUri = BuildSourceUri(document.SourcePath), .DisplayName = document.DisplayName,
                .RepresentationId = document.Representation.RepresentationId, .OriginalSourceSha256 = document.Representation.SourceHash,
                .ExportedTextFileSha256 = document.Representation.TextFileHash,
                .IndexedPayloadSha256 = If(document.Index Is Nothing, "", document.Index.PayloadHash),
                .ExtractionCompleteness = document.Representation.Completeness, .OriginalSourceMapJson = If(sourceMap, "")}
        End Function

        ''' <summary>
        ''' Inline resolution can perform more work after a search/read completes. Check
        ''' all blocks again immediately before building the final model context.
        ''' </summary>
        Friend Async Function RevalidateForDisclosureAsync(search As SemanticArchiveSearchResult, read As SemanticArchiveReadResult,
                         cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            Await _scope.State.OperationGate.WaitAsync(cancellationToken).ConfigureAwait(False)
            Try
                _scope.State.EnsureUsable()
                Await System.Threading.Tasks.Task.Run(Sub() SemanticArchiveLibrary.ValidateScope(_context, _store, _scope.SelectedArchiveIds), cancellationToken).ConfigureAwait(False)
                If search IsNot Nothing Then
                    For position As System.Int32 = search.Hits.Count - 1 To 0 Step -1
                        Dim hit As SemanticArchiveSearchHit = search.Hits(position)
                        Dim document As SemanticArchiveDocumentRecord = Nothing
                        If Not TryResolveCurrentDocument(hit.HitReference, hit.ArchiveId, hit.GenerationId, hit.DocumentId, document) Then
                            search.Hits.RemoveAt(position)
                            search.Coverage.SourcesUnavailable += 1
                            search.Coverage.Diagnostics.Add("source_revoked_before_context: A file reference and its metadata were removed before model exposure.")
                            search.Status = "partial"
                        End If
                    Next
                End If
                If read IsNot Nothing Then RevalidateReadForDisclosure(read)
            Finally
                _scope.State.OperationGate.Release()
            End Try
        End Function

        Private Sub RevalidateReadForDisclosure(read As SemanticArchiveReadResult,
                         Optional checkpoints As System.Collections.Generic.Dictionary(Of System.String, SemanticArchiveDocumentReadState) = Nothing)
            Dim authorized As New System.Collections.Generic.Dictionary(Of System.String, System.Boolean)(System.StringComparer.Ordinal)
            For position As System.Int32 = read.Excerpts.Count - 1 To 0 Step -1
                Dim excerpt As SemanticArchiveExcerpt = read.Excerpts(position)
                Dim document As SemanticArchiveDocumentRecord = Nothing
                Dim mayExpose As System.Boolean
                If Not authorized.TryGetValue(excerpt.HitReference, mayExpose) Then
                    mayExpose = TryResolveCurrentDocument(excerpt.HitReference, excerpt.ArchiveId, excerpt.GenerationId, excerpt.DocumentId, document)
                    authorized(excerpt.HitReference) = mayExpose
                End If
                If Not mayExpose Then
                    read.Excerpts.RemoveAt(position)
                    read.MoreHitReferences.Remove(excerpt.HitReference)
                    Dim stored As SemanticArchiveEvidenceReference = Nothing
                    If _scope.State.Hits.TryGetValue(excerpt.HitReference, stored) Then
                        Dim checkpoint As SemanticArchiveDocumentReadState = Nothing
                        If checkpoints IsNot Nothing Then checkpoints.TryGetValue(excerpt.HitReference, checkpoint)
                        stored.ReadState = checkpoint
                    End If
                    _scope.State.LoadedEvidenceBytes = System.Math.Max(0L, _scope.State.LoadedEvidenceBytes - excerpt.LengthBytes)
                    read.Coverage.EvidenceBytes = System.Math.Max(0L, read.Coverage.EvidenceBytes - excerpt.LengthBytes)
                    read.Coverage.SourcesUnavailable += 1
                    read.Coverage.Diagnostics.Add("source_revoked_before_context: An exact excerpt and its source metadata were removed before exposure.")
                    read.Status = "partial"
                End If
            Next
        End Sub

        Private Shared Function CopyReadState(source As SemanticArchiveDocumentReadState) As SemanticArchiveDocumentReadState
            If source Is Nothing Then Return New SemanticArchiveDocumentReadState()
            Return New SemanticArchiveDocumentReadState() With {
                .NextSectionIndex = source.NextSectionIndex,
                .PendingSections = New System.Collections.Generic.List(Of SharedMethods.SemanticSearchSelectedEntryResult)(source.PendingSections),
                .LoadedSections = New System.Collections.Generic.HashSet(Of System.String)(source.LoadedSections, System.StringComparer.OrdinalIgnoreCase),
                .SectionOffsets = New System.Collections.Generic.Dictionary(Of System.String, System.Int64)(source.SectionOffsets, System.StringComparer.OrdinalIgnoreCase),
                .DeferredSectionIds = New System.Collections.Generic.List(Of System.String)(source.DeferredSectionIds),
                .DeferredSectionSet = New System.Collections.Generic.HashSet(Of System.String)(source.DeferredSectionSet, System.StringComparer.OrdinalIgnoreCase),
                .SmallFileNextByte = source.SmallFileNextByte, .SmallFileComplete = source.SmallFileComplete}
        End Function

        Private Function TryResolveCurrentDocument(reference As System.String, archiveId As System.String, generationId As System.String,
                         documentId As System.String, ByRef document As SemanticArchiveDocumentRecord) As System.Boolean
            Try
                Dim hit As SemanticArchiveEvidenceReference = Nothing
                If Not _scope.State.Hits.TryGetValue(reference, hit) OrElse Not _scope.ContainsArchive(hit.ArchiveId) OrElse
                   Not System.String.Equals(hit.ArchiveId, archiveId, System.StringComparison.Ordinal) OrElse
                   Not System.String.Equals(hit.Generation.GenerationId, generationId, System.StringComparison.Ordinal) OrElse
                   Not System.String.Equals(hit.DocumentId, documentId, System.StringComparison.Ordinal) Then Return False
                document = _store.LoadDocument(hit.Generation, hit.DocumentId)
                Return document IsNot Nothing AndAlso _store.CanReadDocument(_scope.AccessContext, hit.Generation, document, verifySourceHash:=True)
            Catch ex As System.Exception
                Return False
            End Try
        End Function
    End Class
End Namespace
