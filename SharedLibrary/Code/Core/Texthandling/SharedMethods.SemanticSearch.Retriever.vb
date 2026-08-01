' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.SemanticSearch.Retriever.vb
' Purpose: Provides reusable readers, validation, caching, semantic selection,
'          byte-range loading, full-scan fallback, and response-verification
'          utilities for generic self-indexed text sources. The concept is referred
'          to as "flat semantic search" or FSS.
'
' Architecture:
'  - Detection/Validation: Reads only the index prefix, validates metadata and all
'    byte ranges, and verifies the SHA-256 hash before accepting an index.
'  - Cache: Maintains one thread-safe lazy load per normalized path and invalidates
'    entries when file length or last-write time changes.
'  - Retrieval: Builds a standalone query, selects only existing IDs through an
'    LLM, merges adjacent ranges, and reads original UTF-8 bytes with FileStream.
'  - Context/Reuse: Adds configurable surrounding bytes, preserves UTF-8 character
'    boundaries, reuses exact segment text, and considers prior/adjacent sources.
'  - Fallback/Verification: Can inspect every segment independently and can verify
'    a generated response before requesting further original source segments.
' =============================================================================

Option Strict On
Option Explicit On
Option Infer On

Imports System.Linq
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary

    Partial Public Class SharedMethods


        Public Class SemanticSearchIndexCacheItem
            Public Property FilePath As String = ""
            Public Property FileLength As Long
            Public Property LastWriteTimeUtc As System.DateTime
            Public Property ContentStartByte As Long
            Public Property ContentByteLength As Long
            Public Property IndexDocument As SemanticSearchIndexDocument = Nothing
            Public Property EntriesById As New System.Collections.Generic.Dictionary(Of String, SemanticSearchIndexEntry)(System.StringComparer.OrdinalIgnoreCase)
            Public Property OrderedEntries As New System.Collections.Generic.List(Of SemanticSearchIndexEntry)()
            Public Property LoadedSegments As New System.Collections.Concurrent.ConcurrentDictionary(
                Of String,
                System.Lazy(Of System.Threading.Tasks.Task(Of String)))(System.StringComparer.OrdinalIgnoreCase)
        End Class

        Public Class SemanticSearchQueryPreparationResult
            Public Property StandaloneQuestion As String = ""
            Public Property SearchIntents As New System.Collections.Generic.List(Of String)()
            Public Property ImportantTerms As New System.Collections.Generic.List(Of String)()
            Public Property RelatedConcepts As New System.Collections.Generic.List(Of String)()
        End Class

        Public Class SemanticSearchSelectedEntryResult
            Public Property Id As String = ""
            Public Property Relevance As Double
            Public Property Reason As String = ""
        End Class

        Public Class SemanticSearchSelectionResult
            Public Property SelectedEntries As New System.Collections.Generic.List(Of SemanticSearchSelectedEntryResult)()
            Public Property PotentiallyMissingInformation As Boolean
            Public Property SuggestedRelatedIds As New System.Collections.Generic.List(Of String)()
        End Class

        Public Class SemanticSearchSegmentScanResult
            Public Property Id As String = ""
            Public Property Relevant As Boolean
            Public Property Relevance As Double
            Public Property Evidence As New System.Collections.Generic.List(Of String)()
            Public Property ReferencedSections As New System.Collections.Generic.List(Of String)()
        End Class

        Public Class SemanticSearchLoadedSourceSegment
            Public Property EntryIds As New System.Collections.Generic.List(Of String)()
            Public Property AbsoluteStartByte As Long
            Public Property RelativeStartByte As Long
            Public Property LengthBytes As Long
            Public Property Text As String = ""
        End Class

        Public Class SemanticSearchResponseVerificationResult
            Public Property Supported As Boolean
            Public Property UnsupportedClaims As New System.Collections.Generic.List(Of String)()
            Public Property MissingDetails As New System.Collections.Generic.List(Of String)()
            Public Property RequiresMoreSources As Boolean
            Public Property AdditionalEntryIds As New System.Collections.Generic.List(Of String)()
            Public Property RevisedSearchIntent As String = ""
        End Class

        Public Class SemanticSearchRetrievalOptions
            Public Property MinimumSelectedSegments As Integer = 1
            Public Property MaximumSelectedSegments As Integer = 8
            Public Property MaximumTotalSegments As Integer = 24
            Public Property ContextBytesBefore As Integer = 2048
            Public Property ContextBytesAfter As Integer = 2048
            Public Property MergeGapBytes As Integer = 0
            Public Property SpecialTaskName As String = "SemanticSearch"
            Public Property IncludePreviouslyUsedIds As Boolean = True
            Public Property MaximumPreviouslyUsedIds As Integer = 4
            Public Property IncludeAdjacentToPreviouslyUsedIds As Boolean = True
            Public Property EnableFullScanFallback As Boolean = True
            Public Property ForceFullScan As Boolean = False
            Public Property FallbackWhenPotentiallyMissing As Boolean = True
            Public Property MinimumSelectionRelevance As Double = 0.35R
            Public Property FullScanMinimumRelevance As Double = 0.5R
            Public Property MaximumFullScanSegments As Integer = 8
            Public Property MaximumReloadRounds As Integer = 2
            Public Property MaximumLlmAttempts As Integer = 2
            Public Property MaximumConversationCharacters As Integer = 12000
        End Class

        Public Class SemanticSearchRetrievalResult
            Public Property IsIndexed As Boolean
            Public Property SearchPreparation As SemanticSearchQueryPreparationResult = Nothing
            Public Property Selection As SemanticSearchSelectionResult = Nothing
            Public Property SelectedEntryIds As New System.Collections.Generic.List(Of String)()
            Public Property LoadedSources As New System.Collections.Generic.List(Of SemanticSearchLoadedSourceSegment)()
            Public Property ReducedSourceText As String = ""
            Public Property FullScanResults As New System.Collections.Generic.List(Of SemanticSearchSegmentScanResult)()
            Public Property UsedFallback As Boolean
            Public Property DiagnosticMessage As String = ""
        End Class

        ''' <summary>
        ''' Reusable state for continuing semantic retrieval across related questions.
        ''' </summary>
        Public Class SemanticSearchConversationState
            Public Property LastUsedEntryIds As New System.Collections.Generic.List(Of String)()
            Public Property LastLoadedSources As New System.Collections.Generic.List(Of SemanticSearchLoadedSourceSegment)()
            Public Property CurrentTopic As String = ""
            Public Property IntentSummary As String = ""
        End Class

        Private Class SemanticSearchByteRange
            Public Property StartByte As Long
            Public Property EndByteExclusive As Long
            Public Property EntryIds As New System.Collections.Generic.List(Of String)()
        End Class

        Private Class SemanticSearchDecodedByteRange
            Public Property RelativeStartByte As Long
            Public Property LengthBytes As Long
            Public Property Text As String = ""
        End Class

        Private Class SemanticSearchIndexHeaderReadResult
            Public Property Json As String = ""
            Public Property ContentStartByte As Long
        End Class

        Private Shared ReadOnly SemanticSearchIndexCache As New System.Collections.Concurrent.ConcurrentDictionary(
            Of String,
            System.Lazy(Of System.Threading.Tasks.Task(Of SemanticSearchIndexCacheItem)))(System.StringComparer.OrdinalIgnoreCase)

        Private Const SemanticSearchMaximumIndexByteLength As Integer = 64 * 1024 * 1024

        Public Shared Function IsPotentiallySemanticSearchIndexedTextFile(path As String) As Boolean
            If String.IsNullOrWhiteSpace(path) OrElse Not System.IO.File.Exists(path) Then
                Return False
            End If

            Dim extension As String = System.IO.Path.GetExtension(path)
            If Not String.Equals(extension, ".txt", System.StringComparison.OrdinalIgnoreCase) AndAlso
               Not String.Equals(extension, ".md", System.StringComparison.OrdinalIgnoreCase) AndAlso
               Not String.Equals(extension, ".log", System.StringComparison.OrdinalIgnoreCase) Then
                Return False
            End If

            Try
                Dim markerBytes As Byte() = SemanticSearchUtf8NoBom.GetBytes(SemanticSearchIndexStartMarker)
                Dim maximumProbeLength As Integer = markerBytes.Length + 3

                Using stream As New System.IO.FileStream(
                    path,
                    System.IO.FileMode.Open,
                    System.IO.FileAccess.Read,
                    System.IO.FileShare.ReadWrite)

                    Dim probeLength As Integer = CInt(System.Math.Min(CLng(maximumProbeLength), stream.Length))
                    If probeLength < markerBytes.Length Then
                        Return False
                    End If

                    Dim buffer(probeLength - 1) As Byte
                    Dim readCount As Integer = stream.Read(buffer, 0, buffer.Length)
                    Dim offset As Integer = 0

                    If readCount >= markerBytes.Length + 3 AndAlso
                       buffer(0) = &HEF AndAlso buffer(1) = &HBB AndAlso buffer(2) = &HBF Then
                        offset = 3
                    End If

                    If readCount - offset < markerBytes.Length Then
                        Return False
                    End If

                    For markerIndex As Integer = 0 To markerBytes.Length - 1
                        If buffer(offset + markerIndex) <> markerBytes(markerIndex) Then
                            Return False
                        End If
                    Next

                    Return True
                End Using
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine(ex.Message)
                Return False
            End Try
        End Function

        Public Shared Async Function TryGetSemanticSearchIndexAsync(
            path As String,
            Optional cancellationToken As System.Threading.CancellationToken = Nothing
        ) As System.Threading.Tasks.Task(Of SemanticSearchIndexCacheItem)

            If String.IsNullOrWhiteSpace(path) OrElse Not System.IO.File.Exists(path) Then
                Return Nothing
            End If

            cancellationToken.ThrowIfCancellationRequested()

            Dim fullPath As String = System.IO.Path.GetFullPath(path)
            Dim currentInfo As New System.IO.FileInfo(fullPath)
            Dim existing As System.Lazy(Of System.Threading.Tasks.Task(Of SemanticSearchIndexCacheItem)) = Nothing

            If SemanticSearchIndexCache.TryGetValue(fullPath, existing) AndAlso existing IsNot Nothing Then
                Try
                    Dim cachedItem As SemanticSearchIndexCacheItem = Await existing.Value.ConfigureAwait(False)
                    cancellationToken.ThrowIfCancellationRequested()

                    currentInfo.Refresh()
                    If cachedItem IsNot Nothing AndAlso
                       cachedItem.FileLength = currentInfo.Length AndAlso
                       cachedItem.LastWriteTimeUtc = currentInfo.LastWriteTimeUtc Then
                        Return cachedItem
                    End If
                Catch ex As System.OperationCanceledException
                    Throw
                Catch ex As System.Exception
                    System.Diagnostics.Debug.WriteLine(ex.Message)
                End Try

                Dim removedExisting As System.Lazy(Of System.Threading.Tasks.Task(Of SemanticSearchIndexCacheItem)) = Nothing
                SemanticSearchIndexCache.TryRemove(fullPath, removedExisting)
            End If

            Dim created As New System.Lazy(Of System.Threading.Tasks.Task(Of SemanticSearchIndexCacheItem))(
                Function() System.Threading.Tasks.Task.Run(Function() LoadSemanticSearchIndexCore(fullPath)),
                System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

            Dim selected As System.Lazy(Of System.Threading.Tasks.Task(Of SemanticSearchIndexCacheItem)) = SemanticSearchIndexCache.GetOrAdd(fullPath, created)

            Try
                Dim loadedItem As SemanticSearchIndexCacheItem = Await selected.Value.ConfigureAwait(False)
                cancellationToken.ThrowIfCancellationRequested()

                If loadedItem Is Nothing Then
                    Dim removedInvalid As System.Lazy(Of System.Threading.Tasks.Task(Of SemanticSearchIndexCacheItem)) = Nothing
                    SemanticSearchIndexCache.TryRemove(fullPath, removedInvalid)
                End If

                Return loadedItem
            Catch ex As System.OperationCanceledException
                Throw
            Catch ex As System.Exception
                Dim removedFailed As System.Lazy(Of System.Threading.Tasks.Task(Of SemanticSearchIndexCacheItem)) = Nothing
                SemanticSearchIndexCache.TryRemove(fullPath, removedFailed)
                System.Diagnostics.Debug.WriteLine(ex.Message)
                Return Nothing
            End Try
        End Function

        Public Shared Sub InvalidateSemanticSearchIndexCache(Optional path As String = Nothing)
            If String.IsNullOrWhiteSpace(path) Then
                SemanticSearchIndexCache.Clear()
                Return
            End If

            Dim removed As System.Lazy(Of System.Threading.Tasks.Task(Of SemanticSearchIndexCacheItem)) = Nothing
            SemanticSearchIndexCache.TryRemove(System.IO.Path.GetFullPath(path), removed)
        End Sub

        Public Shared Async Function RetrieveSemanticSearchAsync(
            path As String,
            context As ISharedContext,
            currentQuestion As String,
            conversation As String,
            Optional previouslyUsedIds As System.Collections.Generic.IEnumerable(Of String) = Nothing,
            Optional options As SemanticSearchRetrievalOptions = Nothing,
            Optional cancellationToken As System.Threading.CancellationToken = Nothing
        ) As System.Threading.Tasks.Task(Of SemanticSearchRetrievalResult)

            If context Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(context))
            End If
            If String.IsNullOrWhiteSpace(currentQuestion) Then
                Throw New System.ArgumentException("A current question is required.", NameOf(currentQuestion))
            End If

            Dim effectiveOptions As SemanticSearchRetrievalOptions = If(options, New SemanticSearchRetrievalOptions())
            ValidateSemanticSearchRetrievalOptions(effectiveOptions)

            Dim result As New SemanticSearchRetrievalResult()
            Dim cacheItem As SemanticSearchIndexCacheItem = Await TryGetSemanticSearchIndexAsync(path, cancellationToken).ConfigureAwait(False)

            If cacheItem Is Nothing Then
                result.IsIndexed = False
                result.DiagnosticMessage = "No valid semantic index was found."
                Return result
            End If

            result.IsIndexed = True
            result.SearchPreparation = Await BuildSemanticSearchQueryPreparationAsync(
                context,
                effectiveOptions.SpecialTaskName,
                currentQuestion,
                conversation,
                effectiveOptions.MaximumConversationCharacters,
                effectiveOptions.MaximumLlmAttempts,
                cancellationToken).ConfigureAwait(False)

            Dim selectedIds As New System.Collections.Generic.List(Of String)()
            Dim runFullScan As Boolean = effectiveOptions.ForceFullScan

            If Not effectiveOptions.ForceFullScan Then
                Dim compactIndex As String = BuildCompactSemanticSearchIndex(cacheItem.OrderedEntries)
                result.Selection = Await SelectSemanticSearchEntriesAsync(
                    context,
                    effectiveOptions.SpecialTaskName,
                    result.SearchPreparation,
                    compactIndex,
                    effectiveOptions.MaximumSelectedSegments,
                    effectiveOptions.MaximumLlmAttempts,
                    cancellationToken).ConfigureAwait(False)

                selectedIds = ValidateAndExpandSemanticSearchSelectedIds(
                    cacheItem,
                    result.Selection,
                    previouslyUsedIds,
                    effectiveOptions)

                Dim highestRelevance As Double = GetHighestSemanticSearchRelevance(result.Selection)
                runFullScan = selectedIds.Count < effectiveOptions.MinimumSelectedSegments OrElse
                              highestRelevance < effectiveOptions.MinimumSelectionRelevance OrElse
                              (effectiveOptions.FallbackWhenPotentiallyMissing AndAlso
                               result.Selection IsNot Nothing AndAlso
                               result.Selection.PotentiallyMissingInformation)
            End If

            If runFullScan AndAlso effectiveOptions.EnableFullScanFallback Then
                result.FullScanResults = Await ScanAllSemanticSearchSegmentsAsync(
                    cacheItem,
                    context,
                    effectiveOptions.SpecialTaskName,
                    result.SearchPreparation,
                    effectiveOptions.FullScanMinimumRelevance,
                    effectiveOptions.MaximumLlmAttempts,
                    cancellationToken).ConfigureAwait(False)

                Dim scanIds As System.Collections.Generic.IEnumerable(Of String) = result.FullScanResults.
                    Where(Function(scanResult As SemanticSearchSegmentScanResult) scanResult.Relevant).
                    OrderByDescending(Function(scanResult As SemanticSearchSegmentScanResult) scanResult.Relevance).
                    Select(Function(scanResult As SemanticSearchSegmentScanResult) scanResult.Id).
                    Take(effectiveOptions.MaximumFullScanSegments)

                selectedIds = MergeValidSemanticSearchIds(
                    cacheItem,
                    scanIds,
                    selectedIds,
                    effectiveOptions.MaximumSelectedSegments)

                result.UsedFallback = True
            End If

            If selectedIds.Count = 0 Then
                result.DiagnosticMessage = "No relevant indexed segments were found."
                Return result
            End If

            result.SelectedEntryIds = selectedIds
            result.LoadedSources = Await LoadSemanticSearchSourcesAsync(
                cacheItem,
                selectedIds,
                effectiveOptions.ContextBytesBefore,
                effectiveOptions.ContextBytesAfter,
                effectiveOptions.MergeGapBytes,
                cancellationToken).ConfigureAwait(False)

            result.ReducedSourceText = BuildReducedSemanticSearchSourceText(cacheItem, result.LoadedSources)
            Return result
        End Function

        ''' <summary>
        ''' Retrieves source segments while reusing and updating a caller-owned conversation state.
        ''' </summary>
        Public Shared Async Function RetrieveSemanticSearchAsync(
            path As String,
            context As ISharedContext,
            currentQuestion As String,
            conversation As String,
            conversationState As SemanticSearchConversationState,
            Optional options As SemanticSearchRetrievalOptions = Nothing,
            Optional cancellationToken As System.Threading.CancellationToken = Nothing
        ) As System.Threading.Tasks.Task(Of SemanticSearchRetrievalResult)

            If conversationState Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(conversationState))
            End If

            Dim searchConversation As String = If(conversation, "")
            If Not String.IsNullOrWhiteSpace(conversationState.CurrentTopic) OrElse
               Not String.IsNullOrWhiteSpace(conversationState.IntentSummary) Then

                searchConversation &= vbCrLf & vbCrLf &
                    "Previous semantic topic: " & If(conversationState.CurrentTopic, "") & vbCrLf &
                    "Previous intent summary: " & If(conversationState.IntentSummary, "")
            End If

            Dim retrieval As SemanticSearchRetrievalResult = Await RetrieveSemanticSearchAsync(
                path,
                context,
                currentQuestion,
                searchConversation,
                conversationState.LastUsedEntryIds,
                options,
                cancellationToken).ConfigureAwait(False)

            UpdateSemanticSearchConversationState(conversationState, retrieval)
            Return retrieval
        End Function

        ''' <summary>
        ''' Updates reusable conversation state from a completed retrieval operation.
        ''' </summary>
        Public Shared Sub UpdateSemanticSearchConversationState(
            conversationState As SemanticSearchConversationState,
            retrieval As SemanticSearchRetrievalResult
        )
            If conversationState Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(conversationState))
            End If
            If retrieval Is Nothing OrElse Not retrieval.IsIndexed Then
                Return
            End If

            conversationState.LastUsedEntryIds = New System.Collections.Generic.List(Of String)(
                If(retrieval.SelectedEntryIds, New System.Collections.Generic.List(Of String)()))

            conversationState.LastLoadedSources = CloneSemanticSearchLoadedSources(retrieval.LoadedSources)

            If retrieval.SearchPreparation IsNot Nothing Then
                conversationState.CurrentTopic = If(retrieval.SearchPreparation.StandaloneQuestion, "").Trim()

                Dim intents As System.Collections.Generic.List(Of String) =
                    NormalizeSemanticSearchStringList(retrieval.SearchPreparation.SearchIntents)
                conversationState.IntentSummary = String.Join(" | ", intents)
            End If
        End Sub

        ''' <summary>Clears all reusable semantic conversation state.</summary>
        Public Shared Sub ResetSemanticSearchConversationState(conversationState As SemanticSearchConversationState)
            If conversationState Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(conversationState))
            End If

            conversationState.LastUsedEntryIds = New System.Collections.Generic.List(Of String)()
            conversationState.LastLoadedSources = New System.Collections.Generic.List(Of SemanticSearchLoadedSourceSegment)()
            conversationState.CurrentTopic = ""
            conversationState.IntentSummary = ""
        End Sub

        Public Shared Async Function VerifySemanticSearchResponseAsync(
            path As String,
            context As ISharedContext,
            specialTaskName As String,
            currentQuestion As String,
            conversation As String,
            retrieval As SemanticSearchRetrievalResult,
            responseText As String,
            Optional cancellationToken As System.Threading.CancellationToken = Nothing,
            Optional maximumLlmAttempts As Integer = 2,
            Optional maximumConversationCharacters As Integer = 12000
        ) As System.Threading.Tasks.Task(Of SemanticSearchResponseVerificationResult)

            If context Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(context))
            End If
            If retrieval Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(retrieval))
            End If
            If String.IsNullOrWhiteSpace(specialTaskName) Then
                specialTaskName = "SemanticSearch"
            End If
            If maximumLlmAttempts < 1 OrElse maximumLlmAttempts > 5 Then
                Throw New System.ArgumentOutOfRangeException(NameOf(maximumLlmAttempts))
            End If
            If maximumConversationCharacters < 1 Then
                Throw New System.ArgumentOutOfRangeException(NameOf(maximumConversationCharacters))
            End If

            Dim cacheItem As SemanticSearchIndexCacheItem = Await TryGetSemanticSearchIndexAsync(path, cancellationToken).ConfigureAwait(False)
            If cacheItem Is Nothing Then
                Return New SemanticSearchResponseVerificationResult() With {
                    .Supported = False,
                    .RequiresMoreSources = False,
                    .MissingDetails = New System.Collections.Generic.List(Of String) From {
                        "The semantic index is no longer available or valid."
                    }
                }
            End If

            Dim compactIndex As String = BuildCompactSemanticSearchIndex(cacheItem.OrderedEntries)
            Dim standaloneQuestion As String = currentQuestion
            If retrieval.SearchPreparation IsNot Nothing AndAlso
               Not String.IsNullOrWhiteSpace(retrieval.SearchPreparation.StandaloneQuestion) Then
                standaloneQuestion = retrieval.SearchPreparation.StandaloneQuestion
            End If

            Dim systemPrompt As String =
                "Verify a generated response against supplied original source excerpts. Do not write a replacement response. " &
                "Return only JSON with Supported, UnsupportedClaims, MissingDetails, RequiresMoreSources, AdditionalEntryIds and RevisedSearchIntent. " &
                "AdditionalEntryIds may contain only IDs visible in the compact index."

            Dim userPrompt As String =
                "Current question:" & vbCrLf & currentQuestion & vbCrLf & vbCrLf &
                "Standalone search question:" & vbCrLf & standaloneQuestion & vbCrLf & vbCrLf &
                "Conversation:" & vbCrLf & LimitSemanticSearchTextFromEnd(conversation, maximumConversationCharacters) & vbCrLf & vbCrLf &
                "Original source excerpts:" & vbCrLf & retrieval.ReducedSourceText & vbCrLf & vbCrLf &
                "Response to verify:" & vbCrLf & responseText & vbCrLf & vbCrLf &
                "Compact semantic index:" & vbCrLf & compactIndex

            Dim verification As SemanticSearchResponseVerificationResult = Await CallSemanticSearchStructuredLlmAsync(Of SemanticSearchResponseVerificationResult)(
                context,
                specialTaskName,
                systemPrompt,
                userPrompt,
                maximumLlmAttempts,
                cancellationToken).ConfigureAwait(False)

            NormalizeSemanticSearchVerification(verification, cacheItem)
            Return verification
        End Function

        Public Shared Async Function RetrieveAdditionalSemanticSearchSourcesAsync(
            path As String,
            context As ISharedContext,
            currentQuestion As String,
            conversation As String,
            previousRetrieval As SemanticSearchRetrievalResult,
            verification As SemanticSearchResponseVerificationResult,
            Optional options As SemanticSearchRetrievalOptions = Nothing,
            Optional cancellationToken As System.Threading.CancellationToken = Nothing
        ) As System.Threading.Tasks.Task(Of SemanticSearchRetrievalResult)

            If context Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(context))
            End If
            If previousRetrieval Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(previousRetrieval))
            End If
            If verification Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(verification))
            End If

            Dim effectiveOptions As SemanticSearchRetrievalOptions = If(options, New SemanticSearchRetrievalOptions())
            ValidateSemanticSearchRetrievalOptions(effectiveOptions)

            Dim result As New SemanticSearchRetrievalResult() With {
                .SearchPreparation = previousRetrieval.SearchPreparation
            }

            Dim cacheItem As SemanticSearchIndexCacheItem = Await TryGetSemanticSearchIndexAsync(path, cancellationToken).ConfigureAwait(False)
            If cacheItem Is Nothing Then
                Return result
            End If

            result.IsIndexed = True
            Dim previousIds As System.Collections.Generic.List(Of String) = If(
                previousRetrieval.SelectedEntryIds,
                New System.Collections.Generic.List(Of String)())
            Dim existingCount As Integer = previousIds.
                Where(Function(id As String) Not String.IsNullOrWhiteSpace(id)).
                Distinct(System.StringComparer.OrdinalIgnoreCase).
                Count()
            Dim remainingCapacity As Integer = System.Math.Max(0, effectiveOptions.MaximumTotalSegments - existingCount)
            Dim maximumAdditionalCount As Integer = System.Math.Min(effectiveOptions.MaximumSelectedSegments, remainingCapacity)
            Dim candidateIds As New System.Collections.Generic.List(Of String)()

            If maximumAdditionalCount = 0 Then
                result.DiagnosticMessage = "The configured total segment limit has already been reached."
                Return result
            End If

            AddValidSemanticSearchIds(cacheItem, candidateIds, verification.AdditionalEntryIds, maximumAdditionalCount)
            AddSemanticSearchNeighbourAndRelatedIds(cacheItem, candidateIds, previousIds, maximumAdditionalCount)
            RemoveExistingSemanticSearchIds(candidateIds, previousIds)

            If Not String.IsNullOrWhiteSpace(verification.RevisedSearchIntent) AndAlso
               candidateIds.Count < maximumAdditionalCount Then

                Dim revisedPreparation As New SemanticSearchQueryPreparationResult() With {
                    .StandaloneQuestion = verification.RevisedSearchIntent,
                    .SearchIntents = New System.Collections.Generic.List(Of String) From {verification.RevisedSearchIntent}
                }

                Dim revisedSelection As SemanticSearchSelectionResult = Await SelectSemanticSearchEntriesAsync(
                    context,
                    effectiveOptions.SpecialTaskName,
                    revisedPreparation,
                    BuildCompactSemanticSearchIndex(cacheItem.OrderedEntries),
                    maximumAdditionalCount,
                    effectiveOptions.MaximumLlmAttempts,
                    cancellationToken).ConfigureAwait(False)

                Dim revisedIds As New System.Collections.Generic.List(Of String)()
                If revisedSelection IsNot Nothing AndAlso revisedSelection.SelectedEntries IsNot Nothing Then
                    revisedIds.AddRange(
                        revisedSelection.SelectedEntries.
                            OrderByDescending(Function(entryResult As SemanticSearchSelectedEntryResult) entryResult.Relevance).
                            Select(Function(entryResult As SemanticSearchSelectedEntryResult) entryResult.Id))
                End If

                AddValidSemanticSearchIds(cacheItem, candidateIds, revisedIds, maximumAdditionalCount)
            End If

            RemoveExistingSemanticSearchIds(candidateIds, previousIds)

            If candidateIds.Count = 0 AndAlso effectiveOptions.EnableFullScanFallback Then
                result.FullScanResults = Await ScanAllSemanticSearchSegmentsAsync(
                    cacheItem,
                    context,
                    effectiveOptions.SpecialTaskName,
                    If(previousRetrieval.SearchPreparation, New SemanticSearchQueryPreparationResult() With {.StandaloneQuestion = currentQuestion}),
                    effectiveOptions.FullScanMinimumRelevance,
                    effectiveOptions.MaximumLlmAttempts,
                    cancellationToken).ConfigureAwait(False)

                Dim scanIds As System.Collections.Generic.IEnumerable(Of String) = result.FullScanResults.
                    Where(Function(scanResult As SemanticSearchSegmentScanResult) scanResult.Relevant).
                    OrderByDescending(Function(scanResult As SemanticSearchSegmentScanResult) scanResult.Relevance).
                    Select(Function(scanResult As SemanticSearchSegmentScanResult) scanResult.Id)

                AddValidSemanticSearchIds(
                    cacheItem,
                    candidateIds,
                    scanIds,
                    System.Math.Min(effectiveOptions.MaximumFullScanSegments, maximumAdditionalCount))
                RemoveExistingSemanticSearchIds(candidateIds, previousIds)
                result.UsedFallback = True
            End If

            If candidateIds.Count = 0 Then
                result.DiagnosticMessage = "No additional indexed segments were found."
                Return result
            End If

            result.SelectedEntryIds = candidateIds
            result.LoadedSources = Await LoadSemanticSearchSourcesAsync(
                cacheItem,
                candidateIds,
                effectiveOptions.ContextBytesBefore,
                effectiveOptions.ContextBytesAfter,
                effectiveOptions.MergeGapBytes,
                cancellationToken).ConfigureAwait(False)

            result.ReducedSourceText = BuildReducedSemanticSearchSourceText(cacheItem, result.LoadedSources)
            Return result
        End Function

        Public Shared Async Function LoadAdditionalSemanticSearchSourcesAsync(
            path As String,
            ids As System.Collections.Generic.IEnumerable(Of String),
            Optional options As SemanticSearchRetrievalOptions = Nothing,
            Optional cancellationToken As System.Threading.CancellationToken = Nothing
        ) As System.Threading.Tasks.Task(Of SemanticSearchRetrievalResult)

            Dim effectiveOptions As SemanticSearchRetrievalOptions = If(options, New SemanticSearchRetrievalOptions())
            ValidateSemanticSearchRetrievalOptions(effectiveOptions)

            Dim result As New SemanticSearchRetrievalResult()
            Dim cacheItem As SemanticSearchIndexCacheItem = Await TryGetSemanticSearchIndexAsync(path, cancellationToken).ConfigureAwait(False)
            If cacheItem Is Nothing Then
                Return result
            End If

            result.IsIndexed = True
            AddValidSemanticSearchIds(cacheItem, result.SelectedEntryIds, ids, effectiveOptions.MaximumTotalSegments)

            If result.SelectedEntryIds.Count = 0 Then
                Return result
            End If

            result.LoadedSources = Await LoadSemanticSearchSourcesAsync(
                cacheItem,
                result.SelectedEntryIds,
                effectiveOptions.ContextBytesBefore,
                effectiveOptions.ContextBytesAfter,
                effectiveOptions.MergeGapBytes,
                cancellationToken).ConfigureAwait(False)

            result.ReducedSourceText = BuildReducedSemanticSearchSourceText(cacheItem, result.LoadedSources)
            Return result
        End Function

        Public Shared Function BuildCompactSemanticSearchIndex(entries As System.Collections.Generic.IEnumerable(Of SemanticSearchIndexEntry)) As String
            Dim builder As New System.Text.StringBuilder()
            If entries Is Nothing Then
                Return builder.ToString()
            End If

            For Each entry As SemanticSearchIndexEntry In entries.OrderBy(Function(value As SemanticSearchIndexEntry) value.Order)
                builder.AppendLine("ID: " & entry.Id)
                builder.AppendLine("Title: " & entry.Title)
                builder.AppendLine("Summary: " & entry.Summary)
                AppendSemanticSearchCompactList(builder, "Topics", entry.Topics)
                AppendSemanticSearchCompactList(builder, "UserIntents", entry.UserIntents)
                AppendSemanticSearchCompactList(builder, "ExactTerms", entry.ExactTerms)
                AppendSemanticSearchCompactList(builder, "Actions", entry.Actions)
                AppendSemanticSearchCompactList(builder, "Constraints", entry.Constraints)
                AppendSemanticSearchCompactList(builder, "CrossReferences", entry.CrossReferences)
                builder.AppendLine()
            Next

            Return builder.ToString()
        End Function

        ''' <summary>
        ''' Removes the internal &lt;documentN name="..."&gt; / &lt;/documentN&gt; combine wrappers from
        ''' excerpt text before it is shown to the model. The wrapper number is an internal
        ''' segmentation aid only; document attribution is provided separately through the "Source"
        ''' line, so the raw tags (and their internal numbers) must not leak into model-visible text.
        ''' </summary>
        Private Shared Function StripSemanticSearchDocumentWrappers(text As String) As String
            If String.IsNullOrEmpty(text) Then
                Return If(text, "")
            End If

            Return SemanticSearchDocumentWrapperEventRegex.Replace(text, "")
        End Function

        Private Shared Async Function CallSanitizedSemanticSearchStructuredLlmAsync(Of TResult As Class)(
            context As ISharedContext,
            specialTaskName As String,
            systemPrompt As String,
            userPrompt As String,
            maximumAttempts As Integer,
            cancellationToken As System.Threading.CancellationToken
        ) As System.Threading.Tasks.Task(Of TResult)

            If context Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(context))
            End If
            If maximumAttempts < 1 Then
                Throw New System.ArgumentOutOfRangeException(NameOf(maximumAttempts))
            End If
            If String.IsNullOrWhiteSpace(specialTaskName) Then
                specialTaskName = "SemanticSearch"
            End If

            Dim backupConfig As ModelConfig = Nothing
            Dim specialTaskConfig As ModelConfig = Nothing
            Dim restoreRequired As Boolean = False
            Dim useSecondApi As Boolean = False
            Dim lastException As System.Exception = Nothing
            Dim lastRawResponse As String = ""

            Try
                If Not String.IsNullOrWhiteSpace(context.INI_AlternateModelPath) AndAlso
                   TryGetSpecialTaskModelConfig(context, context.INI_AlternateModelPath, specialTaskName, specialTaskConfig) Then

                    backupConfig = GetCurrentConfig(context)

                    Dim errorFlag As Boolean = False
                    ApplyModelConfig(context, specialTaskConfig, errorFlag)
                    If errorFlag Then
                        Throw New System.InvalidOperationException(
                            "Failed to apply the special task model for '" & specialTaskName & "'.")
                    End If

                    restoreRequired = True
                    useSecondApi = True
                End If

                For attempt As Integer = 1 To maximumAttempts
                    cancellationToken.ThrowIfCancellationRequested()

                    Try
                        Dim rawResponse As String = Await LLM(
                            context,
                            systemPrompt,
                            userPrompt,
                            UseSecondAPI:=useSecondApi,
                            Hidesplash:=True,
                            cancellationToken:=cancellationToken).ConfigureAwait(False)

                        rawResponse = WebAgentInterpreter.SanitizeLlmResult(rawResponse)
                        lastRawResponse = If(rawResponse, "")

                        If String.IsNullOrWhiteSpace(rawResponse) Then
                            Throw New System.InvalidOperationException("The LLM returned an empty structured response.")
                        End If

                        Dim result As TResult = DeserializeSemanticSearchJson(Of TResult)(rawResponse)
                        If result IsNot Nothing Then
                            Return result
                        End If

                        Throw New System.InvalidOperationException(
                            "The sanitized structured response could not be deserialized.")
                    Catch ex As System.OperationCanceledException
                        Throw
                    Catch ex As System.Exception
                        lastException = ex
                        System.Diagnostics.Debug.WriteLine(
                            "Semantic search structured task '" & specialTaskName & "' attempt " &
                            attempt.ToString(System.Globalization.CultureInfo.InvariantCulture) & " failed: " & ex.ToString())
                    End Try
                Next
            Finally
                If restoreRequired AndAlso backupConfig IsNot Nothing Then
                    RestoreDefaults(context, backupConfig)
                End If
            End Try

            Throw New System.InvalidOperationException(
                "The structured LLM task failed after " &
                maximumAttempts.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                " attempts." &
                BuildSemanticSearchStructuredFailureDetail(specialTaskName, lastException, lastRawResponse),
                lastException)
        End Function

        Private Shared Function LoadSemanticSearchIndexCore(path As String) As SemanticSearchIndexCacheItem
            Dim initialInfo As New System.IO.FileInfo(path)

            Using stream As New System.IO.FileStream(
                path,
                System.IO.FileMode.Open,
                System.IO.FileAccess.Read,
                System.IO.FileShare.Read)

                Dim header As SemanticSearchIndexHeaderReadResult = ReadSemanticSearchIndexHeader(stream)
                If header Is Nothing Then
                    Return Nothing
                End If

                Dim indexDocument As SemanticSearchIndexDocument = DeserializeSemanticSearchJson(Of SemanticSearchIndexDocument)(header.Json)
                Dim contentLength As Long = stream.Length - header.ContentStartByte

                If Not ValidateSemanticSearchIndexDocument(indexDocument, contentLength) Then
                    Return Nothing
                End If

                Dim actualHash As String = ComputeSemanticSearchContentSha256Hex(stream, header.ContentStartByte)
                If Not String.Equals(actualHash, indexDocument.ContentSha256, System.StringComparison.OrdinalIgnoreCase) Then
                    Return Nothing
                End If

                Dim finalInfo As New System.IO.FileInfo(path)
                finalInfo.Refresh()
                If finalInfo.Length <> initialInfo.Length OrElse
                   finalInfo.LastWriteTimeUtc <> initialInfo.LastWriteTimeUtc Then
                    Return Nothing
                End If

                Dim cacheItem As New SemanticSearchIndexCacheItem() With {
                    .FilePath = path,
                    .FileLength = finalInfo.Length,
                    .LastWriteTimeUtc = finalInfo.LastWriteTimeUtc,
                    .ContentStartByte = header.ContentStartByte,
                    .ContentByteLength = contentLength,
                    .IndexDocument = indexDocument,
                    .OrderedEntries = indexDocument.Entries.
                        OrderBy(Function(entry As SemanticSearchIndexEntry) entry.Order).
                        ToList()
                }

                For Each entry As SemanticSearchIndexEntry In cacheItem.OrderedEntries
                    cacheItem.EntriesById.Add(entry.Id, entry)
                Next

                Return cacheItem
            End Using
        End Function

        Private Shared Function ReadSemanticSearchIndexHeader(stream As System.IO.FileStream) As SemanticSearchIndexHeaderReadResult
            stream.Seek(0L, System.IO.SeekOrigin.Begin)

            Using accumulated As New System.IO.MemoryStream()
                Dim chunk(8191) As Byte

                While accumulated.Length < SemanticSearchMaximumIndexByteLength
                    Dim remaining As Integer = CInt(System.Math.Min(chunk.Length, SemanticSearchMaximumIndexByteLength - accumulated.Length))
                    Dim readCount As Integer = stream.Read(chunk, 0, remaining)
                    If readCount <= 0 Then
                        Exit While
                    End If

                    accumulated.Write(chunk, 0, readCount)
                    Dim bytes As Byte() = accumulated.ToArray()
                    Dim markerStart As Integer = -1
                    Dim contentStart As Integer = -1

                    If TryLocateSemanticSearchContentMarker(bytes, markerStart, contentStart) Then
                        Dim headerByteCount As Integer = markerStart
                        While headerByteCount > 0 AndAlso
                              (bytes(headerByteCount - 1) = 10 OrElse bytes(headerByteCount - 1) = 13)
                            headerByteCount -= 1
                        End While

                        Dim headerText As String = SemanticSearchUtf8NoBom.GetString(bytes, 0, headerByteCount)
                        If headerText.Length > 0 AndAlso headerText(0) = ChrW(&HFEFF) Then
                            headerText = headerText.Substring(1)
                        End If

                        Dim firstLineFeed As Integer = headerText.IndexOf(ControlChars.Lf)
                        If firstLineFeed < 0 Then
                            Return Nothing
                        End If

                        Dim markerLine As String = headerText.Substring(0, firstLineFeed).TrimEnd(ControlChars.Cr)
                        If Not String.Equals(markerLine, SemanticSearchIndexStartMarker, System.StringComparison.Ordinal) Then
                            Return Nothing
                        End If

                        Dim json As String = headerText.Substring(firstLineFeed + 1).Trim()
                        If String.IsNullOrWhiteSpace(json) Then
                            Return Nothing
                        End If

                        Return New SemanticSearchIndexHeaderReadResult() With {
                            .Json = json,
                            .ContentStartByte = contentStart
                        }
                    End If
                End While
            End Using

            Return Nothing
        End Function

        Private Shared Function TryLocateSemanticSearchContentMarker(
            bytes As Byte(),
            ByRef markerStart As Integer,
            ByRef contentStart As Integer
        ) As Boolean

            markerStart = -1
            contentStart = -1
            Dim markerBytes As Byte() = SemanticSearchUtf8NoBom.GetBytes(SemanticSearchContentStartMarker)

            For byteIndex As Integer = 0 To bytes.Length - markerBytes.Length
                If byteIndex > 0 AndAlso bytes(byteIndex - 1) <> 10 Then
                    Continue For
                End If

                Dim matches As Boolean = True
                For markerIndex As Integer = 0 To markerBytes.Length - 1
                    If bytes(byteIndex + markerIndex) <> markerBytes(markerIndex) Then
                        matches = False
                        Exit For
                    End If
                Next

                If Not matches Then
                    Continue For
                End If

                Dim afterMarker As Integer = byteIndex + markerBytes.Length
                If afterMarker < bytes.Length AndAlso bytes(afterMarker) = 10 Then
                    markerStart = byteIndex
                    contentStart = afterMarker + 1
                    Return True
                End If
                If afterMarker + 1 < bytes.Length AndAlso
                   bytes(afterMarker) = 13 AndAlso bytes(afterMarker + 1) = 10 Then
                    markerStart = byteIndex
                    contentStart = afterMarker + 2
                    Return True
                End If
            Next

            Return False
        End Function

        Private Shared Function ValidateSemanticSearchIndexDocument(indexDocument As SemanticSearchIndexDocument, contentLength As Long) As Boolean
            If indexDocument Is Nothing OrElse indexDocument.FormatVersion <> SemanticSearchCurrentFormatVersion Then
                Return False
            End If
            If Not String.Equals(indexDocument.Encoding, "utf-8", System.StringComparison.OrdinalIgnoreCase) Then
                Return False
            End If
            If Not String.Equals(indexDocument.OffsetUnit, "byte", System.StringComparison.OrdinalIgnoreCase) Then
                Return False
            End If
            If Not String.Equals(indexDocument.OffsetBase, "content", System.StringComparison.OrdinalIgnoreCase) Then
                Return False
            End If
            If String.IsNullOrWhiteSpace(indexDocument.GeneratorVersion) OrElse
               String.IsNullOrWhiteSpace(indexDocument.CreatedUtc) Then
                Return False
            End If

            Dim createdUtc As System.DateTime
            If Not System.DateTime.TryParse(
                indexDocument.CreatedUtc,
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.RoundtripKind,
                createdUtc) OrElse createdUtc.Kind <> System.DateTimeKind.Utc Then
                Return False
            End If
            If Not System.Text.RegularExpressions.Regex.IsMatch(If(indexDocument.ContentSha256, ""), "\A[0-9a-fA-F]{64}\z") Then
                Return False
            End If
            If indexDocument.Entries Is Nothing OrElse indexDocument.SegmentCount <> indexDocument.Entries.Count Then
                Return False
            End If

            If contentLength = 0 Then
                Return indexDocument.Entries.Count = 0
            End If
            If indexDocument.Entries.Count = 0 Then
                Return False
            End If

            Dim ids As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.OrdinalIgnoreCase)
            Dim orders As New System.Collections.Generic.HashSet(Of Integer)()
            Dim orderedEntries As System.Collections.Generic.List(Of SemanticSearchIndexEntry) = indexDocument.Entries.
                OrderBy(Function(entry As SemanticSearchIndexEntry) entry.Order).
                ToList()
            Dim expectedStart As Long = 0

            For entryIndex As Integer = 0 To orderedEntries.Count - 1
                Dim entry As SemanticSearchIndexEntry = orderedEntries(entryIndex)

                If entry Is Nothing Then
                    Return False
                End If
                If String.IsNullOrWhiteSpace(entry.Id) OrElse
                   Not System.Text.RegularExpressions.Regex.IsMatch(entry.Id, "\AS\d{4,}\z", System.Text.RegularExpressions.RegexOptions.IgnoreCase) OrElse
                   Not ids.Add(entry.Id) Then
                    Return False
                End If
                If entry.Order <> entryIndex + 1 OrElse Not orders.Add(entry.Order) Then
                    Return False
                End If
                If entry.StartByte <> expectedStart OrElse entry.LengthBytes <= 0 Then
                    Return False
                End If
                If entry.StartByte > contentLength - entry.LengthBytes Then
                    Return False
                End If
                If String.IsNullOrWhiteSpace(entry.Title) OrElse
                   String.IsNullOrWhiteSpace(entry.Summary) OrElse
                   entry.Topics Is Nothing OrElse
                   entry.UserIntents Is Nothing OrElse
                   entry.ExactTerms Is Nothing OrElse
                   entry.Actions Is Nothing OrElse
                   entry.Constraints Is Nothing OrElse
                   entry.CrossReferences Is Nothing OrElse
                   entry.RelatedIds Is Nothing Then
                    Return False
                End If

                Dim expectedPreviousId As String = If(entryIndex > 0, orderedEntries(entryIndex - 1).Id, Nothing)
                Dim expectedNextId As String = If(entryIndex < orderedEntries.Count - 1, orderedEntries(entryIndex + 1).Id, Nothing)

                If Not String.Equals(entry.PreviousId, expectedPreviousId, System.StringComparison.OrdinalIgnoreCase) Then
                    Return False
                End If
                If Not String.Equals(entry.NextId, expectedNextId, System.StringComparison.OrdinalIgnoreCase) Then
                    Return False
                End If

                Dim relatedIds As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.OrdinalIgnoreCase)
                For Each relatedId As String In entry.RelatedIds
                    If String.IsNullOrWhiteSpace(relatedId) OrElse
                       String.Equals(relatedId, entry.Id, System.StringComparison.OrdinalIgnoreCase) OrElse
                       Not relatedIds.Add(relatedId) Then
                        Return False
                    End If
                Next

                expectedStart += entry.LengthBytes
            Next

            If expectedStart <> contentLength Then
                Return False
            End If

            For Each entry As SemanticSearchIndexEntry In orderedEntries
                If entry.RelatedIds IsNot Nothing Then
                    For Each relatedId As String In entry.RelatedIds
                        If Not ids.Contains(relatedId) Then
                            Return False
                        End If
                    Next
                End If
            Next

            Return True
        End Function

        Private Shared Function ComputeSemanticSearchContentSha256Hex(stream As System.IO.FileStream, contentStartByte As Long) As String
            stream.Seek(contentStartByte, System.IO.SeekOrigin.Begin)

            Using sha256 As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                Dim hash As Byte() = sha256.ComputeHash(stream)
                Dim builder As New System.Text.StringBuilder(hash.Length * 2)
                For Each value As Byte In hash
                    builder.Append(value.ToString("x2", System.Globalization.CultureInfo.InvariantCulture))
                Next
                Return builder.ToString()
            End Using
        End Function

        Private Shared Async Function BuildSemanticSearchQueryPreparationAsync(
            context As ISharedContext,
            specialTaskName As String,
            question As String,
            conversation As String,
            maximumConversationCharacters As Integer,
            maximumAttempts As Integer,
            cancellationToken As System.Threading.CancellationToken
        ) As System.Threading.Tasks.Task(Of SemanticSearchQueryPreparationResult)

            Dim systemPrompt As String =
                "Rewrite the current question as a standalone search query using the conversation only to resolve references. " &
                "Do not inspect source content and do not answer the question. Return only JSON with StandaloneQuestion, SearchIntents, ImportantTerms and RelatedConcepts."

            Dim userPrompt As String =
                "Conversation:" & vbCrLf & LimitSemanticSearchTextFromEnd(conversation, maximumConversationCharacters) & vbCrLf & vbCrLf &
                "Current question:" & vbCrLf & question

            Dim preparation As SemanticSearchQueryPreparationResult = Await CallSemanticSearchStructuredLlmAsync(Of SemanticSearchQueryPreparationResult)(
                context,
                specialTaskName,
                systemPrompt,
                userPrompt,
                maximumAttempts,
                cancellationToken).ConfigureAwait(False)
            If String.IsNullOrWhiteSpace(preparation.StandaloneQuestion) Then
                preparation.StandaloneQuestion = question
            End If

            preparation.SearchIntents = NormalizeSemanticSearchStringList(preparation.SearchIntents)
            preparation.ImportantTerms = NormalizeSemanticSearchStringList(preparation.ImportantTerms)
            preparation.RelatedConcepts = NormalizeSemanticSearchStringList(preparation.RelatedConcepts)
            Return preparation
        End Function

        Private Shared Async Function SelectSemanticSearchEntriesAsync(
            context As ISharedContext,
            specialTaskName As String,
            preparation As SemanticSearchQueryPreparationResult,
            compactIndex As String,
            maximumEntries As Integer,
            maximumAttempts As Integer,
            cancellationToken As System.Threading.CancellationToken
        ) As System.Threading.Tasks.Task(Of SemanticSearchSelectionResult)

            Dim systemPrompt As String =
                "Select relevant existing segment IDs from a semantic index. Consider meaning, synonyms and indirect relationships. " &
                "Do not answer the question, invent IDs or output byte positions. Return only a raw JSON object with exactly these fields: " &
                "SelectedEntries, PotentiallyMissingInformation, and SuggestedRelatedIds. " &
                "SelectedEntries must be an array of objects with Id, Relevance, and Reason. " &
                "PotentiallyMissingInformation must be a JSON boolean value only (true or false), never a list, object, string, or explanation. " &
                "SuggestedRelatedIds must be an array of existing IDs. " &
                "Do not wrap the JSON in markdown or code fences. " &
                "Select at most " &
                maximumEntries.ToString(System.Globalization.CultureInfo.InvariantCulture) & " entries."

            Dim userPrompt As String =
                "Search preparation:" & vbCrLf & SerializeSemanticSearchJson(preparation) & vbCrLf & vbCrLf &
                "Compact semantic index:" & vbCrLf & compactIndex

            Dim selection As SemanticSearchSelectionResult = Await CallSanitizedSemanticSearchStructuredLlmAsync(Of SemanticSearchSelectionResult)(
                context,
                specialTaskName,
                systemPrompt,
                userPrompt,
                maximumAttempts,
                cancellationToken).ConfigureAwait(False)
            Dim normalizedSelectedEntries As New System.Collections.Generic.List(Of SemanticSearchSelectedEntryResult)()
            If selection.SelectedEntries IsNot Nothing Then
                For Each item As SemanticSearchSelectedEntryResult In selection.SelectedEntries
                    If item Is Nothing OrElse String.IsNullOrWhiteSpace(item.Id) Then
                        Continue For
                    End If

                    item.Id = item.Id.Trim()
                    If System.Double.IsNaN(item.Relevance) OrElse System.Double.IsInfinity(item.Relevance) Then
                        item.Relevance = 0.0R
                    Else
                        item.Relevance = System.Math.Max(0.0R, System.Math.Min(1.0R, item.Relevance))
                    End If
                    item.Reason = If(item.Reason, "").Trim()
                    normalizedSelectedEntries.Add(item)
                Next
            End If
            selection.SelectedEntries = normalizedSelectedEntries
            selection.SuggestedRelatedIds = NormalizeSemanticSearchStringList(selection.SuggestedRelatedIds)
            Return selection
        End Function

        Private Shared Async Function ScanAllSemanticSearchSegmentsAsync(
            cacheItem As SemanticSearchIndexCacheItem,
            context As ISharedContext,
            specialTaskName As String,
            preparation As SemanticSearchQueryPreparationResult,
            minimumRelevance As Double,
            maximumAttempts As Integer,
            cancellationToken As System.Threading.CancellationToken
        ) As System.Threading.Tasks.Task(Of System.Collections.Generic.List(Of SemanticSearchSegmentScanResult))

            Dim results As New System.Collections.Generic.List(Of SemanticSearchSegmentScanResult)()

            For Each entry As SemanticSearchIndexEntry In cacheItem.OrderedEntries
                cancellationToken.ThrowIfCancellationRequested()
                Dim segmentText As String = Await LoadExactSemanticSearchSegmentTextAsync(cacheItem, entry, cancellationToken).ConfigureAwait(False)
                Dim scanResult As SemanticSearchSegmentScanResult = Await ScanSemanticSearchSegmentAsync(
                    context,
                    specialTaskName,
                    preparation,
                    entry,
                    segmentText,
                    maximumAttempts,
                    cancellationToken).ConfigureAwait(False)

                scanResult.Id = entry.Id
                If System.Double.IsNaN(scanResult.Relevance) OrElse System.Double.IsInfinity(scanResult.Relevance) Then
                    scanResult.Relevance = 0.0R
                Else
                    scanResult.Relevance = System.Math.Max(0.0R, System.Math.Min(1.0R, scanResult.Relevance))
                End If
                scanResult.Evidence = NormalizeSemanticSearchStringList(scanResult.Evidence)
                scanResult.ReferencedSections = NormalizeSemanticSearchStringList(scanResult.ReferencedSections)
                scanResult.Relevant = scanResult.Relevant AndAlso scanResult.Relevance >= minimumRelevance
                results.Add(scanResult)
            Next

            Return results
        End Function

        Private Shared Async Function ScanSemanticSearchSegmentAsync(
            context As ISharedContext,
            specialTaskName As String,
            preparation As SemanticSearchQueryPreparationResult,
            entry As SemanticSearchIndexEntry,
            segmentText As String,
            maximumAttempts As Integer,
            cancellationToken As System.Threading.CancellationToken
        ) As System.Threading.Tasks.Task(Of SemanticSearchSegmentScanResult)

            Dim systemPrompt As String =
                "Evaluate one original source segment for relevance to a search question. Do not write a final response. " &
                "Return only JSON with Relevant, Relevance, Evidence and ReferencedSections. " &
                "Relevant must be a JSON boolean (true or false). " &
                "Relevance must be a JSON number between 0 and 1 (for example 0.0), never a string or explanation. " &
                "Evidence and ReferencedSections must be arrays of concise strings present in the segment. " &
                "Do not wrap the JSON in markdown or code fences."

            Dim userPrompt As String =
                "Search preparation:" & vbCrLf & SerializeSemanticSearchJson(preparation) & vbCrLf & vbCrLf &
                "Segment ID: " & entry.Id & vbCrLf &
                "Segment title: " & entry.Title & vbCrLf & vbCrLf &
                "<SEGMENT>" & vbCrLf & segmentText & vbCrLf & "</SEGMENT>"

            Return Await CallSemanticSearchStructuredLlmAsync(Of SemanticSearchSegmentScanResult)(
                context,
                specialTaskName,
                systemPrompt,
                userPrompt,
                maximumAttempts,
                cancellationToken).ConfigureAwait(False)
        End Function

        Private Shared Function ValidateAndExpandSemanticSearchSelectedIds(
            cacheItem As SemanticSearchIndexCacheItem,
            selection As SemanticSearchSelectionResult,
            previouslyUsedIds As System.Collections.Generic.IEnumerable(Of String),
            options As SemanticSearchRetrievalOptions
        ) As System.Collections.Generic.List(Of String)

            Dim result As New System.Collections.Generic.List(Of String)()

            ' Follow-up questions first reuse prior sources, then direct neighbours/relations,
            ' and only then add newly selected index entries.
            If options.IncludePreviouslyUsedIds AndAlso previouslyUsedIds IsNot Nothing Then
                Dim previousIds As System.Collections.Generic.List(Of String) = previouslyUsedIds.
                    Where(Function(id As String) Not String.IsNullOrWhiteSpace(id)).
                    Distinct(System.StringComparer.OrdinalIgnoreCase).
                    Take(options.MaximumPreviouslyUsedIds).
                    ToList()

                AddValidSemanticSearchIds(cacheItem, result, previousIds, options.MaximumSelectedSegments)

                If options.IncludeAdjacentToPreviouslyUsedIds Then
                    AddSemanticSearchNeighbourAndRelatedIds(cacheItem, result, previousIds, options.MaximumSelectedSegments)
                End If
            End If

            If selection IsNot Nothing AndAlso selection.SelectedEntries IsNot Nothing Then
                AddValidSemanticSearchIds(
                    cacheItem,
                    result,
                    selection.SelectedEntries.
                        OrderByDescending(Function(selectedEntry As SemanticSearchSelectedEntryResult) selectedEntry.Relevance).
                        Select(Function(selectedEntry As SemanticSearchSelectedEntryResult) selectedEntry.Id),
                    options.MaximumSelectedSegments)
            End If

            If selection IsNot Nothing Then
                AddValidSemanticSearchIds(cacheItem, result, selection.SuggestedRelatedIds, options.MaximumSelectedSegments)
            End If

            Return result
        End Function

        Private Shared Function MergeValidSemanticSearchIds(
            cacheItem As SemanticSearchIndexCacheItem,
            primaryIds As System.Collections.Generic.IEnumerable(Of String),
            secondaryIds As System.Collections.Generic.IEnumerable(Of String),
            maximumCount As Integer
        ) As System.Collections.Generic.List(Of String)

            Dim result As New System.Collections.Generic.List(Of String)()
            AddValidSemanticSearchIds(cacheItem, result, primaryIds, maximumCount)
            AddValidSemanticSearchIds(cacheItem, result, secondaryIds, maximumCount)
            Return result
        End Function

        Private Shared Sub AddValidSemanticSearchIds(
            cacheItem As SemanticSearchIndexCacheItem,
            target As System.Collections.Generic.List(Of String),
            ids As System.Collections.Generic.IEnumerable(Of String),
            maximumCount As Integer
        )

            If ids Is Nothing Then
                Return
            End If

            For Each id As String In ids
                If target.Count >= maximumCount Then
                    Exit For
                End If
                If Not String.IsNullOrWhiteSpace(id) AndAlso
                   cacheItem.EntriesById.ContainsKey(id) AndAlso
                   Not target.Contains(id, System.StringComparer.OrdinalIgnoreCase) Then
                    target.Add(id)
                End If
            Next
        End Sub

        Private Shared Sub AddSemanticSearchNeighbourAndRelatedIds(
            cacheItem As SemanticSearchIndexCacheItem,
            target As System.Collections.Generic.List(Of String),
            baseIds As System.Collections.Generic.IEnumerable(Of String),
            maximumCount As Integer
        )

            If baseIds Is Nothing Then
                Return
            End If

            For Each baseId As String In baseIds
                If target.Count >= maximumCount Then
                    Exit For
                End If

                Dim entry As SemanticSearchIndexEntry = Nothing
                If Not cacheItem.EntriesById.TryGetValue(baseId, entry) Then
                    Continue For
                End If

                AddValidSemanticSearchIds(cacheItem, target, New String() {entry.PreviousId, entry.NextId}, maximumCount)
                AddValidSemanticSearchIds(cacheItem, target, entry.RelatedIds, maximumCount)
                AddSemanticSearchCrossReferenceIds(cacheItem, target, entry.CrossReferences, maximumCount)
            Next
        End Sub

        Private Shared Sub AddSemanticSearchCrossReferenceIds(
            cacheItem As SemanticSearchIndexCacheItem,
            target As System.Collections.Generic.List(Of String),
            crossReferences As System.Collections.Generic.IEnumerable(Of String),
            maximumCount As Integer
        )

            If crossReferences Is Nothing Then
                Return
            End If

            For Each crossReference As String In crossReferences
                If target.Count >= maximumCount Then
                    Exit For
                End If
                If String.IsNullOrWhiteSpace(crossReference) Then
                    Continue For
                End If

                Dim matches As System.Text.RegularExpressions.MatchCollection =
                    System.Text.RegularExpressions.Regex.Matches(
                        crossReference,
                        "\bS\d{4,}\b",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase)

                For Each match As System.Text.RegularExpressions.Match In matches
                    AddValidSemanticSearchIds(cacheItem, target, New String() {match.Value}, maximumCount)
                    If target.Count >= maximumCount Then
                        Exit For
                    End If
                Next
            Next
        End Sub

        Private Shared Sub RemoveExistingSemanticSearchIds(
            candidateIds As System.Collections.Generic.List(Of String),
            existingIds As System.Collections.Generic.IEnumerable(Of String)
        )

            If existingIds Is Nothing Then
                Return
            End If

            Dim existingSet As New System.Collections.Generic.HashSet(Of String)(existingIds, System.StringComparer.OrdinalIgnoreCase)
            candidateIds.RemoveAll(Function(id As String) existingSet.Contains(id))
        End Sub

        Private Shared Function GetHighestSemanticSearchRelevance(selection As SemanticSearchSelectionResult) As Double
            If selection Is Nothing OrElse selection.SelectedEntries Is Nothing OrElse selection.SelectedEntries.Count = 0 Then
                Return 0.0R
            End If
            Return selection.SelectedEntries.Max(Function(entry As SemanticSearchSelectedEntryResult) entry.Relevance)
        End Function

        Private Shared Async Function LoadExactSemanticSearchSegmentTextAsync(
            cacheItem As SemanticSearchIndexCacheItem,
            entry As SemanticSearchIndexEntry,
            cancellationToken As System.Threading.CancellationToken
        ) As System.Threading.Tasks.Task(Of String)

            Dim created As New System.Lazy(Of System.Threading.Tasks.Task(Of String))(
                Function() ReadExactSemanticSearchSegmentTextAsync(cacheItem, entry),
                System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

            Dim selected As System.Lazy(Of System.Threading.Tasks.Task(Of String)) = cacheItem.LoadedSegments.GetOrAdd(entry.Id, created)

            Try
                Dim value As String = Await selected.Value.ConfigureAwait(False)
                cancellationToken.ThrowIfCancellationRequested()
                Return value
            Catch ex As System.Exception
                Dim removed As System.Lazy(Of System.Threading.Tasks.Task(Of String)) = Nothing
                cacheItem.LoadedSegments.TryRemove(entry.Id, removed)
                Throw
            End Try
        End Function

        Private Shared Async Function ReadExactSemanticSearchSegmentTextAsync(
            cacheItem As SemanticSearchIndexCacheItem,
            entry As SemanticSearchIndexEntry
        ) As System.Threading.Tasks.Task(Of String)

            If entry.LengthBytes > Integer.MaxValue Then
                Throw New System.IO.IOException("A source segment is too large to load into memory.")
            End If

            Dim data(CInt(entry.LengthBytes) - 1) As Byte
            Using stream As New System.IO.FileStream(
                cacheItem.FilePath,
                System.IO.FileMode.Open,
                System.IO.FileAccess.Read,
                System.IO.FileShare.Read,
                81920,
                True)

                ValidateOpenSemanticSearchFile(cacheItem, stream)
                stream.Seek(cacheItem.ContentStartByte + entry.StartByte, System.IO.SeekOrigin.Begin)
                Await ReadSemanticSearchExactlyAsync(stream, data, System.Threading.CancellationToken.None).ConfigureAwait(False)
                ValidateOpenSemanticSearchFile(cacheItem, stream)
            End Using

            Return SemanticSearchUtf8NoBom.GetString(data)
        End Function

        Private Shared Async Function LoadSemanticSearchSourcesAsync(
            cacheItem As SemanticSearchIndexCacheItem,
            ids As System.Collections.Generic.IEnumerable(Of String),
            beforeBytes As Integer,
            afterBytes As Integer,
            mergeGapBytes As Integer,
            cancellationToken As System.Threading.CancellationToken
        ) As System.Threading.Tasks.Task(Of System.Collections.Generic.List(Of SemanticSearchLoadedSourceSegment))

            Dim ranges As New System.Collections.Generic.List(Of SemanticSearchByteRange)()

            For Each id As String In ids
                Dim entry As SemanticSearchIndexEntry = Nothing
                If cacheItem.EntriesById.TryGetValue(id, entry) Then
                    ranges.Add(New SemanticSearchByteRange() With {
                        .StartByte = System.Math.Max(0L, entry.StartByte - CLng(beforeBytes)),
                        .EndByteExclusive = AddSemanticSearchBytesClamped(
                            entry.StartByte + entry.LengthBytes,
                            afterBytes,
                            cacheItem.ContentByteLength),
                        .EntryIds = New System.Collections.Generic.List(Of String) From {entry.Id}
                    })
                End If
            Next

            Dim mergedRanges As System.Collections.Generic.List(Of SemanticSearchByteRange) = MergeSemanticSearchRanges(ranges, mergeGapBytes)
            Dim result As New System.Collections.Generic.List(Of SemanticSearchLoadedSourceSegment)()

            Using stream As New System.IO.FileStream(
                cacheItem.FilePath,
                System.IO.FileMode.Open,
                System.IO.FileAccess.Read,
                System.IO.FileShare.Read,
                81920,
                True)

                ValidateOpenSemanticSearchFile(cacheItem, stream)

                For Each range As SemanticSearchByteRange In mergedRanges
                    cancellationToken.ThrowIfCancellationRequested()
                    Dim decoded As SemanticSearchDecodedByteRange = Await ReadAndDecodeSemanticSearchRangeAsync(
                        stream,
                        cacheItem.ContentStartByte,
                        range.StartByte,
                        range.EndByteExclusive,
                        cancellationToken).ConfigureAwait(False)

                    result.Add(New SemanticSearchLoadedSourceSegment() With {
                        .EntryIds = SortSemanticSearchEntryIds(cacheItem, range.EntryIds),
                        .AbsoluteStartByte = cacheItem.ContentStartByte + decoded.RelativeStartByte,
                        .RelativeStartByte = decoded.RelativeStartByte,
                        .LengthBytes = decoded.LengthBytes,
                        .Text = decoded.Text
                    })
                Next

                ValidateOpenSemanticSearchFile(cacheItem, stream)
            End Using

            Return result
        End Function

        Private Shared Sub ValidateOpenSemanticSearchFile(
            cacheItem As SemanticSearchIndexCacheItem,
            stream As System.IO.FileStream
        )

            Dim fileInfo As New System.IO.FileInfo(cacheItem.FilePath)
            fileInfo.Refresh()

            If stream.Length <> cacheItem.FileLength OrElse
               fileInfo.Length <> cacheItem.FileLength OrElse
               fileInfo.LastWriteTimeUtc <> cacheItem.LastWriteTimeUtc Then

                InvalidateSemanticSearchIndexCache(cacheItem.FilePath)
                Throw New System.IO.IOException("The indexed source changed while it was being accessed. Retry the operation with the reloaded index.")
            End If
        End Sub

        Private Shared Async Function ReadAndDecodeSemanticSearchRangeAsync(
            stream As System.IO.FileStream,
            contentStartByte As Long,
            rangeStart As Long,
            rangeEndExclusive As Long,
            cancellationToken As System.Threading.CancellationToken
        ) As System.Threading.Tasks.Task(Of SemanticSearchDecodedByteRange)

            Dim length As Long = rangeEndExclusive - rangeStart
            If length <= 0 Then
                Return New SemanticSearchDecodedByteRange() With {.RelativeStartByte = rangeStart}
            End If
            If length > Integer.MaxValue Then
                Throw New System.IO.IOException("A selected source range is too large to load into memory.")
            End If

            Dim data(CInt(length) - 1) As Byte
            stream.Seek(contentStartByte + rangeStart, System.IO.SeekOrigin.Begin)
            Await ReadSemanticSearchExactlyAsync(stream, data, cancellationToken).ConfigureAwait(False)

            Dim leadingTrim As Integer = 0
            While leadingTrim < data.Length AndAlso IsSemanticSearchUtf8ContinuationByte(data(leadingTrim))
                leadingTrim += 1
            End While

            Dim availableLength As Integer = data.Length - leadingTrim
            For trailingTrim As Integer = 0 To System.Math.Min(3, availableLength)
                Dim decodeLength As Integer = availableLength - trailingTrim
                Try
                    Dim text As String = SemanticSearchUtf8NoBom.GetString(data, leadingTrim, decodeLength)
                    Return New SemanticSearchDecodedByteRange() With {
                        .RelativeStartByte = rangeStart + leadingTrim,
                        .LengthBytes = decodeLength,
                        .Text = text
                    }
                Catch ex As System.Text.DecoderFallbackException
                    If trailingTrim = System.Math.Min(3, availableLength) Then
                        Throw
                    End If
                End Try
            Next

            Throw New System.Text.DecoderFallbackException("The selected source range is not valid UTF-8.")
        End Function

        Private Shared Async Function ReadSemanticSearchExactlyAsync(
            stream As System.IO.FileStream,
            data As Byte(),
            cancellationToken As System.Threading.CancellationToken
        ) As System.Threading.Tasks.Task

            Dim offset As Integer = 0
            While offset < data.Length
                Dim readCount As Integer = Await stream.ReadAsync(
                    data,
                    offset,
                    data.Length - offset,
                    cancellationToken).ConfigureAwait(False)

                If readCount = 0 Then
                    Throw New System.IO.EndOfStreamException("The indexed source ended before the selected byte range was fully read.")
                End If
                offset += readCount
            End While
        End Function

        Private Shared Function IsSemanticSearchUtf8ContinuationByte(value As Byte) As Boolean
            Return (value And &HC0) = &H80
        End Function

        Private Shared Function AddSemanticSearchBytesClamped(
            value As Long,
            additionalBytes As Integer,
            maximumValue As Long
        ) As Long
            If additionalBytes <= 0 Then
                Return System.Math.Min(value, maximumValue)
            End If
            If value >= maximumValue - CLng(additionalBytes) Then
                Return maximumValue
            End If
            Return value + CLng(additionalBytes)
        End Function

        Private Shared Function MergeSemanticSearchRanges(
            ranges As System.Collections.Generic.IEnumerable(Of SemanticSearchByteRange),
            mergeGapBytes As Integer
        ) As System.Collections.Generic.List(Of SemanticSearchByteRange)

            Dim orderedRanges As System.Collections.Generic.List(Of SemanticSearchByteRange) = ranges.
                OrderBy(Function(range As SemanticSearchByteRange) range.StartByte).
                ToList()
            Dim result As New System.Collections.Generic.List(Of SemanticSearchByteRange)()

            For Each currentRange As SemanticSearchByteRange In orderedRanges
                If result.Count = 0 OrElse
                   currentRange.StartByte > AddSemanticSearchBytesClamped(
                       result(result.Count - 1).EndByteExclusive,
                       mergeGapBytes,
                       Long.MaxValue) Then

                    result.Add(New SemanticSearchByteRange() With {
                        .StartByte = currentRange.StartByte,
                        .EndByteExclusive = currentRange.EndByteExclusive,
                        .EntryIds = New System.Collections.Generic.List(Of String)(currentRange.EntryIds)
                    })
                Else
                    Dim lastRange As SemanticSearchByteRange = result(result.Count - 1)
                    lastRange.EndByteExclusive = System.Math.Max(lastRange.EndByteExclusive, currentRange.EndByteExclusive)
                    For Each id As String In currentRange.EntryIds
                        If Not lastRange.EntryIds.Contains(id, System.StringComparer.OrdinalIgnoreCase) Then
                            lastRange.EntryIds.Add(id)
                        End If
                    Next
                End If
            Next

            Return result
        End Function

        Private Shared Function SortSemanticSearchEntryIds(
            cacheItem As SemanticSearchIndexCacheItem,
            ids As System.Collections.Generic.IEnumerable(Of String)
        ) As System.Collections.Generic.List(Of String)

            Return ids.
                Where(Function(id As String) cacheItem.EntriesById.ContainsKey(id)).
                Distinct(System.StringComparer.OrdinalIgnoreCase).
                OrderBy(Function(id As String) cacheItem.EntriesById(id).Order).
                ToList()
        End Function

        Private Shared Function BuildReducedSemanticSearchSourceText(
            cacheItem As SemanticSearchIndexCacheItem,
            sources As System.Collections.Generic.IEnumerable(Of SemanticSearchLoadedSourceSegment)
        ) As String

            Dim builder As New System.Text.StringBuilder()
            builder.AppendLine("The following blocks contain original excerpts selected from the configured source. When referring to evidence, cite the ""Source"" document name shown for each block (never any internal identifier or number).")
            builder.AppendLine()

            Dim blockNumber As Integer = 0
            For Each source As SemanticSearchLoadedSourceSegment In sources
                blockNumber += 1
                Dim titles As New System.Collections.Generic.List(Of String)()
                Dim sourceDocuments As New System.Collections.Generic.List(Of String)()

                For Each id As String In source.EntryIds
                    Dim entry As SemanticSearchIndexEntry = Nothing
                    If cacheItem.EntriesById.TryGetValue(id, entry) Then
                        titles.Add(entry.Title)
                        If entry.SourceDocuments IsNot Nothing Then
                            For Each documentName As String In entry.SourceDocuments
                                If Not String.IsNullOrWhiteSpace(documentName) AndAlso
                                   Not sourceDocuments.Contains(documentName) Then
                                    sourceDocuments.Add(documentName)
                                End If
                            Next
                        End If
                    End If
                Next

                builder.AppendLine("<<<EXCERPT " & blockNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) & ">>>")
                If sourceDocuments.Count > 0 Then
                    builder.AppendLine("Source: " & String.Join(" | ", sourceDocuments))
                End If
                builder.AppendLine(StripSemanticSearchDocumentWrappers(source.Text))
                builder.AppendLine("<<<END EXCERPT " & blockNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) & ">>>")
                builder.AppendLine()
            Next

            Return builder.ToString()
        End Function

        Private Shared Function CloneSemanticSearchLoadedSources(
            sources As System.Collections.Generic.IEnumerable(Of SemanticSearchLoadedSourceSegment)
        ) As System.Collections.Generic.List(Of SemanticSearchLoadedSourceSegment)

            Dim result As New System.Collections.Generic.List(Of SemanticSearchLoadedSourceSegment)()
            If sources Is Nothing Then
                Return result
            End If

            For Each source As SemanticSearchLoadedSourceSegment In sources
                If source Is Nothing Then
                    Continue For
                End If

                result.Add(New SemanticSearchLoadedSourceSegment() With {
                    .EntryIds = New System.Collections.Generic.List(Of String)(
                        If(source.EntryIds, New System.Collections.Generic.List(Of String)())),
                    .AbsoluteStartByte = source.AbsoluteStartByte,
                    .RelativeStartByte = source.RelativeStartByte,
                    .LengthBytes = source.LengthBytes,
                    .Text = If(source.Text, "")
                })
            Next

            Return result
        End Function

        Private Shared Sub AppendSemanticSearchCompactList(
            builder As System.Text.StringBuilder,
            name As String,
            values As System.Collections.Generic.IEnumerable(Of String)
        )

            If values Is Nothing Then
                Return
            End If

            Dim list As System.Collections.Generic.List(Of String) = values.
                Where(Function(value As String) Not String.IsNullOrWhiteSpace(value)).
                ToList()

            If list.Count > 0 Then
                builder.AppendLine(name & ": " & String.Join(" | ", list))
            End If
        End Sub

        Private Shared Function LimitSemanticSearchTextFromEnd(value As String, maximumCharacters As Integer) As String
            Dim text As String = If(value, "")
            If maximumCharacters <= 0 OrElse text.Length <= maximumCharacters Then
                Return text
            End If

            Dim startIndex As Integer = text.Length - maximumCharacters
            If startIndex > 0 AndAlso startIndex < text.Length AndAlso
               System.Char.IsLowSurrogate(text(startIndex)) AndAlso
               System.Char.IsHighSurrogate(text(startIndex - 1)) Then
                startIndex += 1
            End If

            Return text.Substring(startIndex)
        End Function

        Private Shared Sub NormalizeSemanticSearchVerification(
            verification As SemanticSearchResponseVerificationResult,
            cacheItem As SemanticSearchIndexCacheItem
        )

            verification.UnsupportedClaims = NormalizeSemanticSearchStringList(verification.UnsupportedClaims)
            verification.MissingDetails = NormalizeSemanticSearchStringList(verification.MissingDetails)
            verification.AdditionalEntryIds = NormalizeSemanticSearchStringList(verification.AdditionalEntryIds).
                Where(Function(id As String) cacheItem.EntriesById.ContainsKey(id)).
                Distinct(System.StringComparer.OrdinalIgnoreCase).
                ToList()
            verification.RevisedSearchIntent = If(verification.RevisedSearchIntent, "").Trim()
        End Sub

        Private Shared Sub ValidateSemanticSearchRetrievalOptions(options As SemanticSearchRetrievalOptions)
            If options.MinimumSelectedSegments < 1 Then
                Throw New System.ArgumentOutOfRangeException(NameOf(options.MinimumSelectedSegments))
            End If
            If options.MaximumSelectedSegments < 1 Then
                Throw New System.ArgumentOutOfRangeException(NameOf(options.MaximumSelectedSegments))
            End If
            If options.MinimumSelectedSegments > options.MaximumSelectedSegments Then
                Throw New System.ArgumentException("MinimumSelectedSegments cannot exceed MaximumSelectedSegments.")
            End If
            If options.MaximumTotalSegments < options.MaximumSelectedSegments Then
                Throw New System.ArgumentException("MaximumTotalSegments must be at least MaximumSelectedSegments.")
            End If
            If options.ContextBytesBefore < 0 Then
                Throw New System.ArgumentOutOfRangeException(NameOf(options.ContextBytesBefore))
            End If
            If options.ContextBytesAfter < 0 Then
                Throw New System.ArgumentOutOfRangeException(NameOf(options.ContextBytesAfter))
            End If
            If options.MergeGapBytes < 0 Then
                Throw New System.ArgumentOutOfRangeException(NameOf(options.MergeGapBytes))
            End If
            If options.MaximumPreviouslyUsedIds < 0 Then
                Throw New System.ArgumentOutOfRangeException(NameOf(options.MaximumPreviouslyUsedIds))
            End If
            If options.ForceFullScan AndAlso Not options.EnableFullScanFallback Then
                Throw New System.ArgumentException("ForceFullScan requires EnableFullScanFallback=True.")
            End If
            If options.MinimumSelectionRelevance < 0.0R OrElse options.MinimumSelectionRelevance > 1.0R Then
                Throw New System.ArgumentOutOfRangeException(NameOf(options.MinimumSelectionRelevance))
            End If
            If options.FullScanMinimumRelevance < 0.0R OrElse options.FullScanMinimumRelevance > 1.0R Then
                Throw New System.ArgumentOutOfRangeException(NameOf(options.FullScanMinimumRelevance))
            End If
            If options.MaximumFullScanSegments < 1 Then
                Throw New System.ArgumentOutOfRangeException(NameOf(options.MaximumFullScanSegments))
            End If
            If options.MaximumReloadRounds < 0 OrElse options.MaximumReloadRounds > 5 Then
                Throw New System.ArgumentOutOfRangeException(NameOf(options.MaximumReloadRounds))
            End If
            If options.MaximumLlmAttempts < 1 OrElse options.MaximumLlmAttempts > 5 Then
                Throw New System.ArgumentOutOfRangeException(NameOf(options.MaximumLlmAttempts))
            End If
            If options.MaximumConversationCharacters < 1 Then
                Throw New System.ArgumentOutOfRangeException(NameOf(options.MaximumConversationCharacters))
            End If
            If String.IsNullOrWhiteSpace(options.SpecialTaskName) Then
                options.SpecialTaskName = "SemanticSearch"
            End If
        End Sub


    End Class

End Namespace
