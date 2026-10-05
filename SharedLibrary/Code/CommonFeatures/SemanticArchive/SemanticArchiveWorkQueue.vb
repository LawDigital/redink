' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.

Option Strict On
Option Explicit On
Option Infer On

Namespace SharedLibrary

    Public NotInheritable Class SemanticArchiveBuildOptions
        Public Property RebuildSemanticMetadata As System.Boolean = SharedMethods.DEFAULT_SEMANTICARCHIVE_REBUILD_SEMANTIC_METADATA
        ' Hard index-only contract: semantic derivatives may be rebuilt, but text extraction/OCR is never invoked.
        Public Property IndexOnlyRebuild As System.Boolean
        Public Property ForceReextract As System.Boolean = SharedMethods.DEFAULT_SEMANTICARCHIVE_FORCE_REEXTRACT
        Public Property ReconcilePermissionsOnly As System.Boolean = SharedMethods.DEFAULT_SEMANTICARCHIVE_RECONCILE_PERMISSIONS_ONLY
        Public Property SelectedDocumentIds As System.Collections.Generic.List(Of System.String) = Nothing
        Public Property OperationId As System.String = ""
        Public Property MaxDiscoveryEntries As System.Int32 = SharedMethods.DEFAULT_SEMANTICARCHIVE_MAX_DISCOVERY_ENTRIES
        Public Property MaxDiscoverySeconds As System.Int32 = SharedMethods.DEFAULT_SEMANTICARCHIVE_MAX_DISCOVERY_SECONDS
        Public Property RetryFailures As System.Boolean = SharedMethods.DEFAULT_SEMANTICARCHIVE_RETRY_FAILURES
        Public Property MaximumFilesPerBatch As System.Int32 = SharedMethods.DEFAULT_SEMANTICARCHIVE_MAXIMUM_FILES_PER_BATCH
        Public Property MaximumWriterLeaseWait As System.Nullable(Of System.TimeSpan) = Nothing
        Public Property ForceScan As System.Boolean = SharedMethods.DEFAULT_SEMANTICARCHIVE_FORCE_SCAN
        Public Property IsBackground As System.Boolean = SharedMethods.DEFAULT_SEMANTICARCHIVE_IS_BACKGROUND
        Public Property FullIntegrityAudit As System.Boolean = SharedMethods.DEFAULT_SEMANTICARCHIVE_FULL_INTEGRITY_AUDIT
        Public Property HostReaderDispatcher As System.Func(Of System.Func(Of String), System.Threading.CancellationToken, System.Threading.Tasks.Task(Of String))

        Friend Function Snapshot() As SemanticArchiveBuildOptions
            Dim copy As SemanticArchiveBuildOptions = DirectCast(Me.MemberwiseClone(), SemanticArchiveBuildOptions)
            If Me.SelectedDocumentIds IsNot Nothing Then copy.SelectedDocumentIds = New System.Collections.Generic.List(Of System.String)(Me.SelectedDocumentIds)
            Return copy
        End Function
    End Class

    Public NotInheritable Class SemanticArchiveBuildProgress
        Public Property SourceDiagnostic As SemanticArchiveSourceDiagnostic
        Public Property ArchiveId As String = ""
        Public Property SourcePath As String = ""
        Public Property Stage As String = ""
        Public Property CompletedFiles As Integer
        Public Property PendingFiles As Integer
        Public Property DeferredFiles As Integer
        Public Property Message As String = ""
    End Class

    Public NotInheritable Class SemanticArchiveBuildResult
        Public Property Inventory As SemanticArchiveInventory
        Public Property ArchiveId As String = ""
        Public Property GenerationId As String = ""
        Public Property ProcessedFiles As Integer
        Public Property ReusedFiles As Integer
        Public Property ExtractsReused As Integer
        Public Property CardsRebuilt As Integer
        Public Property SectionIndexesRebuilt As Integer
        Public Property RoutingGroupsBuilt As Integer
        Public Property DocumentsRequiringExtraction As Integer
        Public Property CoverageExcludedFiles As Integer
        Public Property FailedFiles As Integer
        Public Property PendingFiles As Integer
        Public Property DeferredFiles As Integer
        Public Property Published As Boolean
        Public Property Cancelled As Boolean
        Public Property WriterLeaseDeferred As Boolean
        Public Property DiscoveryPending As System.Boolean
        Public Property DiscoveryEntriesInspected As System.Int32
        Public Property PermissionSourcesChecked As System.Int32
        Public Property PermissionArtifactsChecked As System.Int32
        Public Property PermissionArtifactsRepaired As System.Int32
        Public Property PermissionArtifactsQuarantined As System.Int32
        Public Property PermissionsPending As System.Boolean
        Public Property PermissionsDeferred As System.Boolean
        Friend Property PermissionRetryReady As System.Boolean
        Public Property SelectionRequired As System.Boolean
        Public Property Diagnostics As New System.Collections.Generic.List(Of String)()
    End Class

    ''' <summary>
    ''' One deduplicated source job. A completed extraction is durable before any model call.
    ''' Ready records remain queued until their generation has been activated successfully.
    ''' </summary>
    Friend NotInheritable Class SemanticArchiveWorkItem
        Public Property DocumentId As String = ""
        Public Property SourceItemId As String = ""
        Public Property CanonicalSourceKey As String = ""
        Public Property SourcePath As String = ""
        Public Property RelativePath As String = ""
        Public Property PartitionKey As String = ""
        Public Property BindingIds As New System.Collections.Generic.List(Of String)()
        Public Property PrimaryBindingId As String = ""
        Public Property SourceLength As Long
        Public Property SourceWriteTicks As Long
        Public Property SourceHash As String = ""
        Public Property ExtractionSignature As String = ""
        Public Property SemanticSignature As String = ""
        Public Property State As String = "Queued"
        Public Property Remove As Boolean
        Public Property ForceSemanticRebuild As Boolean
        Public Property ForceExtractionRebuild As System.Boolean
        ' Durable operation intent: background drain/resume may not silently enable OCR.
        Public Property IndexOnlyRebuild As System.Boolean
        Public Property Attempts As Integer
        Public Property NextAttemptUtc As System.DateTime = System.DateTime.MinValue
        Public Property UpdatedUtc As System.DateTime = System.DateTime.UtcNow
        Public Property LastError As String = ""
        Public Property PermissionContinuationToken As System.String = ""
        Public Property CachedDocument As SemanticArchiveDocumentRecord
    End Class

    Friend NotInheritable Class SemanticArchiveScanCheckpoint
        Public Property FilteredFiles As System.Int64
        Public Property ExcludedEntries As System.Int64
        Public Property SchemaVersion As System.Int32 = 2
        Public Property InProgress As System.Boolean
        Public Property CycleId As System.String = ""
        Public Property RequestId As System.String = ""
        Public Property ScopeSignature As System.String = ""
        Public Property Phase As System.String = "discover"
        Public Property ForceSemanticRebuild As System.Boolean
        Public Property ForceExtractionRebuild As System.Boolean
        ' Durable operation intent: background drain/resume may not silently enable OCR.
        Public Property IndexOnlyRebuild As System.Boolean
        Public Property RetryFailures As System.Boolean
        Public Property ExplicitRefresh As System.Boolean
        Public Property DirectoryHead As System.Int64
        Public Property DirectoryTail As System.Int64
        Public Property DirectoryOffset As System.Int64
        Public Property DirectoryWriteTicks As System.Int64
        Public Property PendingEntryPath As System.String = ""
        Public Property PermissionContinuationToken As System.String = ""
        Public Property PermissionRetryRoundEnd As System.Int64
        Public Property PermissionRetrySelectionPosition As System.Int32
        Public Property SelectedPosition As System.Int32
        Public Property ReconcileShard As System.Int32
        Public Property ReconcileDocument As System.Int32
        Public Property QueueAfterId As System.String = ""
        Public Property LastCompletedUtc As System.DateTime = System.DateTime.MinValue
        Public Property LastAttemptUtc As System.DateTime = System.DateTime.MinValue
        Public Property ConfigurationSignature As String = ""
        Public Property IncompleteBindingIds As New System.Collections.Generic.List(Of String)()
        Public Property Diagnostics As New System.Collections.Generic.List(Of String)()
    End Class

    ''' <summary>
    ''' Storage-owned queue; callers hold the archive writer lease. Each source has a small
    ''' independent checkpoint, so a file completion never rewrites the complete archive.
    ''' Retry state survives Office restarts. There is deliberately no delete/move overwrite
    ''' fallback: a failed atomic checkpoint is a failed unit of work.
    ''' </summary>
    Friend NotInheritable Partial Class SemanticArchiveWorkQueue
        Private ReadOnly _directory As String
        Private ReadOnly _itemsDirectory As String
        Private ReadOnly _index As SemanticArchiveQueueIndex

        Public Sub New(store As SemanticArchiveStore, archiveId As String)
            _directory = store.GetWorkDirectory(archiveId)
            _itemsDirectory = System.IO.Path.Combine(_directory, "items")
            SemanticArchiveStore.CreatePrivateDirectory(_directory)
            SemanticArchiveStore.CreatePrivateDirectory(_itemsDirectory)
            _index = SemanticArchiveQueueIndex.Open(_directory, _itemsDirectory)
        End Sub

        Public ReadOnly Property DirectoryPath As String
            Get
                Return _directory
            End Get
        End Property

        Private Function ItemPath(documentId As String) As String
            If System.String.IsNullOrWhiteSpace(documentId) OrElse
                documentId.IndexOfAny(System.IO.Path.GetInvalidFileNameChars()) >= 0 OrElse
                documentId.Contains("..") OrElse documentId.Contains("/") OrElse documentId.Contains("\") Then
                Throw New System.IO.InvalidDataException("Invalid archive queue identity.")
            End If
            Return System.IO.Path.Combine(_itemsDirectory, documentId & ".json")
        End Function

        Public Function Load(documentId As String) As SemanticArchiveWorkItem
            Dim path As String = ItemPath(documentId)
            Dim result As SemanticArchiveWorkItem = ReadPrivateControl(Of SemanticArchiveWorkItem)(_directory, path, 32L * 1024L * 1024L)
            If result Is Nothing Then Return Nothing
            If result Is Nothing OrElse Not System.String.Equals(result.DocumentId, documentId, System.StringComparison.Ordinal) Then
                Throw New System.IO.InvalidDataException("The archive work checkpoint is invalid.")
            End If
            Return result
        End Function

        Public Sub Save(item As SemanticArchiveWorkItem)
            If item Is Nothing Then Throw New System.ArgumentNullException(NameOf(item))
            item.UpdatedUtc = System.DateTime.UtcNow
            Try
                _index.BeginMutation()
                SemanticArchiveStore.AtomicWriteJson(ItemPath(item.DocumentId), item)
                _index.Put(item)
                _index.CommitMutation()
            Catch ex As System.Exception
                _index.Invalidate()
                Throw
            End Try
        End Sub

        Public Function Enqueue(item As SemanticArchiveWorkItem, Optional replaceOperationIntent As System.Boolean = False) As Boolean
            If _index.Matches(item) Then Return False
            Dim previous As SemanticArchiveWorkItem = Load(item.DocumentId)
            ' Only background continuation preserves an older no-extraction intent.
            ' An explicitly authorized replacement command supplies its own operation policy.
            If Not replaceOperationIntent AndAlso previous IsNot Nothing AndAlso previous.IndexOnlyRebuild AndAlso Not item.ForceExtractionRebuild Then item.IndexOnlyRebuild = True
            Dim sameSource As System.Boolean = previous IsNot Nothing AndAlso
                System.String.Equals(previous.CanonicalSourceKey, item.CanonicalSourceKey, System.StringComparison.Ordinal)
            Dim compatibleHash As System.Boolean = previous IsNot Nothing AndAlso (System.String.IsNullOrEmpty(item.SourceHash) OrElse
                System.String.Equals(previous.SourceHash, item.SourceHash, System.StringComparison.OrdinalIgnoreCase))
            If previous IsNot Nothing AndAlso Not item.ForceSemanticRebuild AndAlso Not item.ForceExtractionRebuild AndAlso previous.Remove = item.Remove AndAlso previous.IndexOnlyRebuild = item.IndexOnlyRebuild AndAlso
                sameSource AndAlso compatibleHash AndAlso System.String.Equals(previous.SourcePath, item.SourcePath, System.StringComparison.Ordinal) AndAlso
                previous.SourceLength = item.SourceLength AndAlso previous.SourceWriteTicks = item.SourceWriteTicks AndAlso
                System.String.Equals(previous.ExtractionSignature, item.ExtractionSignature, System.StringComparison.Ordinal) AndAlso
                System.String.Equals(previous.SemanticSignature, item.SemanticSignature, System.StringComparison.Ordinal) AndAlso
                System.String.Equals(previous.PartitionKey, item.PartitionKey, System.StringComparison.Ordinal) AndAlso
                SameBindings(previous.BindingIds, item.BindingIds) Then
                Return False
            End If
            If Not item.ForceExtractionRebuild AndAlso previous IsNot Nothing AndAlso previous.CachedDocument IsNot Nothing AndAlso
                sameSource AndAlso compatibleHash AndAlso
                System.String.Equals(previous.ExtractionSignature, item.ExtractionSignature, System.StringComparison.Ordinal) Then
                ' A metadata-only scan may not know the source hash yet. Preserve
                ' durable extraction/index work; ProcessDocumentAsync hashes the live
                ' original before deciding whether this checkpoint can be reused.
                item.CachedDocument = previous.CachedDocument
                item.State = "Extracted"
            End If
            Save(item)
            Return True
        End Function

        Private Shared Function SameBindings(left As System.Collections.Generic.List(Of String), right As System.Collections.Generic.List(Of String)) As Boolean
            Dim values As New System.Collections.Generic.HashSet(Of String)(If(left, New System.Collections.Generic.List(Of String)()), System.StringComparer.Ordinal)
            Return values.SetEquals(If(right, New System.Collections.Generic.List(Of String)()))
        End Function

        Public Iterator Function EnumerateItems() As System.Collections.Generic.IEnumerable(Of SemanticArchiveWorkItem)
            For Each documentId As String In _index.AllIds()
                Dim item As SemanticArchiveWorkItem = Load(documentId)
                If item IsNot Nothing Then Yield item
            Next
        End Function

        Public Function DocumentIds() As System.Collections.Generic.List(Of String)
            Return _index.AllIds()
        End Function

        Public Function DocumentIdsAfter(afterId As System.String, maximum As System.Int32) As System.Collections.Generic.List(Of System.String)
            Return _index.IdsAfter(afterId, maximum)
        End Function

        Public Sub ResetRetry(documentId As System.String)
            Dim item As SemanticArchiveWorkItem = Load(documentId)
            If item Is Nothing OrElse (item.State <> "Retry" AndAlso item.State <> "PendingHost") Then Return
            item.NextAttemptUtc = System.DateTime.MinValue
            item.State = "Queued"
            Save(item)
        End Sub

        Public Function GetBatch(maximum As Integer, retryFailures As Boolean, isBackground As Boolean, Optional selectedDocumentIds As System.Collections.Generic.IEnumerable(Of System.String) = Nothing) As System.Collections.Generic.List(Of SemanticArchiveWorkItem)
            Dim result As New System.Collections.Generic.List(Of SemanticArchiveWorkItem)()
            For Each documentId As String In _index.GetRunnableIds(maximum, retryFailures, isBackground, selectedDocumentIds)
                Dim item As SemanticArchiveWorkItem = Load(documentId)
                If item Is Nothing Then Throw New System.IO.InvalidDataException("An indexed source checkpoint is missing; queue reconciliation is required.")
                result.Add(item)
            Next
            Return result
        End Function

        Public Function PendingCount(Optional runnableOnly As Boolean = False, Optional isBackground As Boolean = False, Optional selectedDocumentIds As System.Collections.Generic.IEnumerable(Of System.String) = Nothing) As Integer
            Return _index.PendingCount(runnableOnly, isBackground, selectedDocumentIds)
        End Function

        Public Sub Fail(item As SemanticArchiveWorkItem, diagnostic As String, Optional requiresHost As Boolean = False, Optional retryOnlyWhenExplicit As System.Boolean = False)
            item.Attempts += 1
            item.State = If(requiresHost, "PendingHost", "Retry")
            item.LastError = If(diagnostic, "Archive processing failed.")
            ' Shared policy at the job boundary, independent of source format or model provider.
            ' Deterministic failures can be held until an explicit Retry failed action rather
            ' than executing an identical automatic retry loop.
            If retryOnlyWhenExplicit Then
                item.NextAttemptUtc = System.DateTime.MaxValue
            Else
                Dim delaySeconds As Double = System.Math.Min(3600.0, 15.0 * System.Math.Pow(2.0, System.Math.Min(8, item.Attempts - 1)))
                item.NextAttemptUtc = System.DateTime.UtcNow.AddSeconds(delaySeconds)
            End If
            Save(item)
        End Sub

        Public Sub Complete(item As SemanticArchiveWorkItem)
            ' Called only after the active-generation pointer has been verified by publication.
            Try
                _index.BeginMutation()
                System.IO.File.Delete(ItemPath(item.DocumentId))
                _index.Remove(item.DocumentId)
                _index.CommitMutation()
            Catch ex As System.Exception
                _index.Invalidate()
                Throw
            End Try
        End Sub

        Public Function LoadScan(Optional permissionsOnly As System.Boolean = False) As SemanticArchiveScanCheckpoint
            Dim path As String = System.IO.Path.Combine(_directory, If(permissionsOnly, "scan-permissions.json", "scan.json"))
            Dim checkpoint As SemanticArchiveScanCheckpoint = ReadPrivateControl(Of SemanticArchiveScanCheckpoint)(_directory, path, 4L * 1024L * 1024L)
            If checkpoint Is Nothing Then Return New SemanticArchiveScanCheckpoint()
            If checkpoint Is Nothing Then Throw New System.IO.InvalidDataException("The archive scan checkpoint is invalid.")
            Return checkpoint
        End Function

        Public Sub SaveScan(checkpoint As SemanticArchiveScanCheckpoint, Optional permissionsOnly As System.Boolean = False)
            While checkpoint.Diagnostics.Count > 100
                checkpoint.Diagnostics.RemoveAt(0)
            End While
            SemanticArchiveStore.AtomicWriteJson(System.IO.Path.Combine(_directory, If(permissionsOnly, "scan-permissions.json", "scan.json")), checkpoint)
            _index.SetDiscoveryPending(checkpoint.InProgress OrElse (permissionsOnly AndAlso HasPermissionRetries()), permissionsOnly)
        End Sub
    End Class
End Namespace
