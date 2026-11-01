' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' Exact named catalog, immutable generations, durable atomic publication and fencing.

' =============================================================================
' File: SemanticArchive.Storage.vb
' Purpose:
'   Named-catalog persistence, immutable generations, writer fencing and atomic
'   publication.
'
' Architecture / Function:
'   Validates the configured catalog and committed generations; descriptor-only reads do
'   not initialize or traverse archive storage.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary
    Public NotInheritable Class SemanticArchiveCatalogConflictException
        Inherits System.InvalidOperationException
        Public Sub New()
            MyBase.New("The Semantic Archives catalog changed in another host. Reload it before saving.")
        End Sub
    End Class


    Public NotInheritable Partial Class SemanticArchiveStore
        Public Const CatalogFileName As System.String = "redink-sa-catalog.json"
        Public Const DataDirectoryName As System.String = "sa-archives"
        Public Const GenerationSchemaVersion As System.Int32 = 2
        Private ReadOnly _directoryPath As System.String
        Private Shared ReadOnly JsonSettings As New Newtonsoft.Json.JsonSerializerSettings With {
            .TypeNameHandling = Newtonsoft.Json.TypeNameHandling.None,
            .DateParseHandling = Newtonsoft.Json.DateParseHandling.DateTimeOffset,
            .MaxDepth = 80
        }

        Public Sub New(configuredDirectory As System.String)
            If System.String.IsNullOrWhiteSpace(configuredDirectory) Then Throw New System.InvalidOperationException("Semantic Archives is not configured. Set SemanticArchiveCatalogPathLocal first.")
            _directoryPath = SemanticArchivePathGuard.RequireWindowsCompatiblePath(configuredDirectory)
            RequireAtomicWritePathBudget(CatalogPath)
            SemanticArchivePathGuard.ValidateContainedPath(_directoryPath, _directoryPath, False)
            RegisterCentralOutputs()
        End Sub

        Public ReadOnly Property DirectoryPath As System.String
            Get
                Return _directoryPath
            End Get
        End Property

        Public ReadOnly Property CatalogPath As System.String
            Get
                Return System.IO.Path.Combine(_directoryPath, CatalogFileName)
            End Get
        End Property

        ' Read-only consumers require an existing named catalog. Do not use the
        ' administrator's missing-catalog initialization path for tool discovery.
        ''' <summary>Descriptor-only read for UI snapshots: no constructor/output registration, writes or document traversal.</summary>
        Public Shared Function ReadCatalogOverview(configuredDirectory As System.String) As SemanticArchiveCatalog
            Dim directory As System.String = SemanticArchivePathGuard.RequireWindowsCompatiblePath(configuredDirectory)
            Dim path As System.String = SemanticArchivePathGuard.ValidateContainedPath(directory, System.IO.Path.Combine(directory, CatalogFileName), True)
            Dim catalog As SemanticArchiveCatalog = ReadJson(Of SemanticArchiveCatalog)(path, SharedMethods.DEFAULT_SEMANTICARCHIVE_CATALOG_MAXIMUM_BYTES)
            ValidateCatalog(catalog)
            Return catalog
        End Function

        Public Function LoadExistingCatalog() As SemanticArchiveCatalog
            Dim catalog As SemanticArchiveCatalog = ReadJson(Of SemanticArchiveCatalog)(CatalogPath, 8 * 1024 * 1024)
            ValidateCatalog(catalog)
            RegisterOutputs(catalog)
            Return catalog
        End Function

        Public Function LoadCatalog() As SemanticArchiveCatalog
            Dim path As System.String = SemanticArchivePathGuard.ValidateContainedPath(_directoryPath, CatalogPath, False)
            Dim catalog As SemanticArchiveCatalog
            If Not ExistsChecked(path) Then
                catalog = New SemanticArchiveCatalog()
            Else
                catalog = ReadJson(Of SemanticArchiveCatalog)(path, 8 * 1024 * 1024)
            End If
            ValidateCatalog(catalog)
            RegisterOutputs(catalog)
            Return catalog
        End Function

        ''' <summary>
        ''' A revision conflict never overwrites another host's edits. Removing a root
        ''' or archive changes search eligibility immediately; no source is deleted.
        ''' </summary>
        Public Sub SaveCatalog(catalog As SemanticArchiveCatalog, expectedRevision As System.Int64)
            ValidateCatalog(catalog)
            ValidatePrivateCatalogLocation(_directoryPath)
            If Not ExistsChecked(_directoryPath) Then CreatePrivateDirectory(_directoryPath)
            CreatePrivateDirectory(System.IO.Path.Combine(_directoryPath, DataDirectoryName))
            Using mutationLock As SemanticArchiveStorageLock = AcquireStorageLock(System.IO.Path.Combine(_directoryPath, DataDirectoryName, "catalog.lock"), System.Threading.CancellationToken.None)
                Dim previous As SemanticArchiveCatalog = LoadCatalog()
                If previous.Revision <> expectedRevision Then Throw New SemanticArchiveCatalogConflictException()
                If previous.Revision = System.Int64.MaxValue Then Throw New System.InvalidOperationException("The catalog revision is exhausted.")
                catalog.Revision = previous.Revision + 1
                ' Source paths and eligibility are checked against this live catalog
                ' on every read, including reads pinned to older generations.
                AtomicWriteJson(CatalogPath, catalog)
                Dim saved As SemanticArchiveCatalog = ReadJson(Of SemanticArchiveCatalog)(CatalogPath, 8 * 1024 * 1024)
                If saved.Revision <> catalog.Revision Then Throw New System.IO.IOException("The catalog publication could not be verified.")
                RegisterOutputs(saved)
            End Using
        End Sub

        Public Function GetArchive(archiveId As System.String) As SemanticArchiveDefinition
            SemanticArchiveIdentity.ValidateId(archiveId, NameOf(archiveId))
            For Each definition As SemanticArchiveDefinition In LoadCatalog().Archives
                If System.String.Equals(definition.ArchiveId, archiveId, System.StringComparison.Ordinal) Then Return definition
            Next
            Return Nothing
        End Function

        Public Sub RemoveArchive(archiveId As System.String)
            Dim catalog As SemanticArchiveCatalog = LoadCatalog()
            Dim revision As System.Int64 = catalog.Revision
            Dim selected As SemanticArchiveDefinition = catalog.Archives.Find(Function(item As SemanticArchiveDefinition) item.ArchiveId = archiveId)
            If SemanticArchiveLibrary.IsSubscriber(selected) Then
                selected.Library.OptOut = True
                selected.Enabled = False
            Else
                If selected IsNot Nothing AndAlso selected.Library IsNot Nothing AndAlso selected.Library.State = "available" Then Throw New System.InvalidOperationException("Withdraw the published definition before unregistering its local publisher archive.")
                catalog.Archives.RemoveAll(Function(item As SemanticArchiveDefinition) System.String.Equals(item.ArchiveId, archiveId, System.StringComparison.Ordinal))
            End If
            catalog.DefaultArchiveIds.RemoveAll(Function(item As System.String) System.String.Equals(item, archiveId, System.StringComparison.Ordinal))
            SaveCatalog(catalog, revision)
        End Sub

        Public Sub RemoveRoot(archiveId As System.String, bindingId As System.String)
            Dim catalog As SemanticArchiveCatalog = LoadCatalog()
            Dim revision As System.Int64 = catalog.Revision
            For Each definition As SemanticArchiveDefinition In catalog.Archives
                If System.String.Equals(definition.ArchiveId, archiveId, System.StringComparison.Ordinal) Then
                    If SemanticArchiveLibrary.IsSubscriber(definition) Then Throw New System.InvalidOperationException("Subscribed source folders are managed by their publisher.")
                    definition.Roots.RemoveAll(Function(binding As SemanticArchiveSourceBinding) System.String.Equals(binding.BindingId, bindingId, System.StringComparison.Ordinal))
                End If
            Next
            SaveCatalog(catalog, revision)
        End Sub

        Public Function GetArchiveDirectory(archiveId As System.String) As System.String
            Return System.IO.Path.Combine(_directoryPath, DataDirectoryName, SemanticArchiveIdentity.ValidateId(archiveId, NameOf(archiveId)))
        End Function

        Public Function GetWorkDirectory(archiveId As System.String) As System.String
            Return System.IO.Path.Combine(GetArchiveDirectory(archiveId), "work")
        End Function

        Public Function GetGenerationDirectory(archiveId As System.String, generationId As System.String) As System.String
            Return System.IO.Path.Combine(GetArchiveDirectory(archiveId), "generations", SemanticArchiveIdentity.ValidateId(generationId, NameOf(generationId)))
        End Function

        ''' <summary>Acquires an exclusive storage-backed writer lease. The optional wait budget
        ''' bounds contention retries only (thirty seconds by default); zero makes one attempt.
        ''' Cancellation wakes retry waits. Synchronous filesystem calls must return safely,
        ''' and elapsed time never expires or replaces another writer's lease.</summary>
        Public Function AcquireWriterLease(archiveId As System.String,
                                           cancellationToken As System.Threading.CancellationToken,
                                           Optional maximumWait As System.TimeSpan? = Nothing) As SemanticArchiveWriterLease
            Dim definition As SemanticArchiveDefinition = GetArchive(archiveId)
            If definition Is Nothing Then Throw New System.InvalidOperationException("The requested archive is not registered.")
            RequireArchiveWritePathBudget(archiveId)
            RegisterOutputs(LoadCatalog())
            Dim directory As System.String = GetArchiveDirectory(archiveId)
            CreatePrivateDirectory(directory)
            Dim storageLock As SemanticArchiveStorageLock = AcquireStorageLock(System.IO.Path.Combine(directory, "writer.lock"), cancellationToken, maximumWait)
            Try
                definition = GetArchive(archiveId)
                If definition Is Nothing Then Throw New System.InvalidOperationException("The archive was removed while its writer lease was pending.")
                Dim configurationSignature As System.String = GetConfigurationSignature(definition)
                Dim fencePath As System.String = System.IO.Path.Combine(directory, "writer-fence.json")
                Dim previous As SemanticArchiveFenceRecord = If(ExistsChecked(fencePath), ReadJson(Of SemanticArchiveFenceRecord)(fencePath, 16384), New SemanticArchiveFenceRecord())
                If previous.Token < 0 OrElse previous.Token = System.Int64.MaxValue Then Throw New System.IO.InvalidDataException("Invalid archive fencing counter.")
                Dim fence As New SemanticArchiveFenceRecord With {.Token = previous.Token + 1, .LeaseId = SemanticArchiveIdentity.NewId()}
                AtomicWriteJson(fencePath, fence)
                Dim current As SemanticArchiveGenerationPointer = ReadPointer(archiveId)
                Return New SemanticArchiveWriterLease(Me, storageLock, archiveId, fence.Token, fence.LeaseId, If(current Is Nothing, "", current.GenerationId), configurationSignature)
            Catch ex As System.Exception
                storageLock.Dispose()
                Throw
            End Try
        End Function

        Public Function NewGeneration(lease As SemanticArchiveWriterLease) As SemanticArchiveGenerationManifest
            AssertLease(lease)
            Dim definition As SemanticArchiveDefinition = GetArchive(lease.ArchiveId)
            If definition Is Nothing Then Throw New System.InvalidOperationException("The archive is no longer registered.")
            Dim manifest As New SemanticArchiveGenerationManifest With {
                .ArchiveId = lease.ArchiveId,
                .GenerationId = SemanticArchiveIdentity.NewId(),
                .PreviousGenerationId = lease.BaseGenerationId,
                .ConfigurationSignature = lease.ConfigurationSignature,
                .MaxChildrenPerNode = definition.MaxChildrenPerNode,
                .MaxRoutingCharacters = definition.MaxRoutingCharacters,
                .FenceToken = lease.FenceToken
            }
            CreatePrivateDirectory(GetGenerationDirectory(manifest.ArchiveId, manifest.GenerationId))
            Return manifest
        End Function

        Public Function PinGeneration(archiveId As System.String) As SemanticArchiveGenerationManifest
            Dim archive As SemanticArchiveDefinition = GetArchive(archiveId)
            If archive Is Nothing OrElse Not archive.Enabled Then Return Nothing
            Return ReadPinnedGeneration(archiveId)
        End Function

        Public Function PinGenerationForAdministration(archiveId As System.String) As SemanticArchiveGenerationManifest
            If GetArchive(archiveId) Is Nothing Then Return Nothing
            Return ReadPinnedGeneration(archiveId)
        End Function

        Public Function PinGenerationForMaintenance(lease As SemanticArchiveWriterLease) As SemanticArchiveGenerationManifest
            AssertLease(lease)
            Return ReadPinnedGeneration(lease.ArchiveId)
        End Function

        Private Function ReadPinnedGeneration(archiveId As System.String) As SemanticArchiveGenerationManifest
            Dim pointer As SemanticArchiveGenerationPointer = ReadPointer(archiveId)
            If pointer Is Nothing Then Return Nothing
            Dim manifest As SemanticArchiveGenerationManifest = ReadArtifactJson(Of SemanticArchiveGenerationManifest)(pointer.Manifest, archiveId, 64 * 1024 * 1024)
            ValidateManifestHeader(manifest)
            If Not System.String.Equals(manifest.ArchiveId, archiveId, System.StringComparison.Ordinal) OrElse Not System.String.Equals(manifest.GenerationId, pointer.GenerationId, System.StringComparison.Ordinal) OrElse manifest.FenceToken <> pointer.FenceToken OrElse manifest.ValidationStatus <> "validated" Then Throw New System.IO.InvalidDataException("The active generation does not match its validated pointer.")
            Return manifest
        End Function

        Public Function LoadNode(manifest As SemanticArchiveGenerationManifest, nodeId As System.String) As SemanticArchiveNode
            ValidateManifestHeader(manifest)
            Dim descriptor As SemanticArchiveNodeDescriptor = Nothing
            If Not manifest.Nodes.TryGetValue(nodeId, descriptor) OrElse descriptor Is Nothing Then Throw New System.IO.InvalidDataException("The node is not a member of the pinned generation.")
            Dim node As SemanticArchiveNode = ReadArtifactJson(Of SemanticArchiveNode)(descriptor.Artifact, manifest.ArchiveId, 8 * 1024 * 1024)
            If node Is Nothing OrElse node.Cards Is Nothing OrElse node.NodeId <> nodeId OrElse node.GenerationId <> descriptor.GenerationId OrElse node.Cards.Count <> descriptor.ChildCount Then Throw New System.IO.InvalidDataException("The immutable node record does not match its descriptor.")
            If descriptor.IndexArtifact Is Nothing Then Throw New System.IO.InvalidDataException("A navigation node has no indexed-text artifact.")
            ValidateArtifact(descriptor.IndexArtifact, manifest.ArchiveId)
            node.IndexPath = ResolveArtifactPath(descriptor.IndexArtifact)
            node.ParentNodeId = descriptor.ParentNodeId
            Return node
        End Function

        Public Function LoadDocumentShard(manifest As SemanticArchiveGenerationManifest, shardId As System.String) As SemanticArchiveDocumentShard
            ValidateManifestHeader(manifest)
            For Each descriptor As SemanticArchiveDocumentShardDescriptor In manifest.DocumentShards
                If System.String.Equals(descriptor.ShardId, shardId, System.StringComparison.Ordinal) Then Return LoadDocumentShard(manifest, descriptor)
            Next
            Throw New System.IO.InvalidDataException("The document shard is not a member of the pinned generation.")
        End Function

        Public Function LoadDocumentShard(manifest As SemanticArchiveGenerationManifest, descriptor As SemanticArchiveDocumentShardDescriptor) As SemanticArchiveDocumentShard
            If manifest Is Nothing OrElse descriptor Is Nothing OrElse Not manifest.DocumentShards.Contains(descriptor) Then Throw New System.IO.InvalidDataException("The document shard is not a member of the pinned generation.")
            Dim shard As SemanticArchiveDocumentShard = ReadArtifactJson(Of SemanticArchiveDocumentShard)(descriptor.Artifact, manifest.ArchiveId, 32 * 1024 * 1024)
            If shard Is Nothing OrElse shard.Documents Is Nothing OrElse shard.ShardId <> descriptor.ShardId OrElse shard.Documents.Count <> descriptor.RecordCount OrElse shard.MinKey <> descriptor.MinKey OrElse shard.MaxKey <> descriptor.MaxKey Then Throw New System.IO.InvalidDataException("The immutable document shard does not match its descriptor.")
            Dim identities As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            For Each document As SemanticArchiveDocumentRecord In shard.Documents
                ValidateDocumentRecord(document)
                If Not identities.Add(document.DocumentId) OrElse Not descriptor.DocumentIds.Contains(document.DocumentId) Then Throw New System.IO.InvalidDataException("Invalid document membership in immutable shard.")
            Next
            If identities.Count <> descriptor.DocumentIds.Count Then Throw New System.IO.InvalidDataException("Incomplete immutable document membership.")
            If Not SemanticArchiveInventory.FromDocuments(shard.Documents).SameAs(descriptor.Inventory) Then Throw New System.IO.InvalidDataException("Document shard inventory does not match its records.")
            Return shard
        End Function

        Public Iterator Function EnumerateDocuments(manifest As SemanticArchiveGenerationManifest) As System.Collections.Generic.IEnumerable(Of SemanticArchiveDocumentRecord)
            ValidateManifestHeader(manifest)
            For Each descriptor As SemanticArchiveDocumentShardDescriptor In manifest.DocumentShards
                For Each document As SemanticArchiveDocumentRecord In LoadDocumentShard(manifest, descriptor).Documents
                    Yield document
                Next
            Next
        End Function

        Public Function LoadDocument(manifest As SemanticArchiveGenerationManifest, documentId As System.String) As SemanticArchiveDocumentRecord
            SemanticArchiveIdentity.ValidateId(documentId, NameOf(documentId))
            ValidateManifestHeader(manifest)
            Dim lookup As System.Collections.Generic.Dictionary(Of System.String, SemanticArchiveDocumentShardDescriptor) = DocumentShardLookups.GetValue(manifest, AddressOf BuildDocumentShardLookup)
            Dim found As SemanticArchiveDocumentShardDescriptor = Nothing
            lookup.TryGetValue(documentId, found)
            If found Is Nothing Then Return Nothing
            For Each document As SemanticArchiveDocumentRecord In LoadDocumentShard(manifest, found).Documents
                If System.String.Equals(document.DocumentId, documentId, System.StringComparison.Ordinal) Then Return document
            Next
            Throw New System.IO.InvalidDataException("The pinned document record is missing.")
        End Function

        Private Shared ReadOnly DocumentShardLookups As New System.Runtime.CompilerServices.ConditionalWeakTable(Of SemanticArchiveGenerationManifest, System.Collections.Generic.Dictionary(Of System.String, SemanticArchiveDocumentShardDescriptor))()

        Private Shared Function BuildDocumentShardLookup(manifest As SemanticArchiveGenerationManifest) As System.Collections.Generic.Dictionary(Of System.String, SemanticArchiveDocumentShardDescriptor)
            Dim lookup As New System.Collections.Generic.Dictionary(Of System.String, SemanticArchiveDocumentShardDescriptor)(System.StringComparer.Ordinal)
            For Each shard As SemanticArchiveDocumentShardDescriptor In manifest.DocumentShards
                If shard Is Nothing OrElse shard.DocumentIds Is Nothing OrElse shard.RecordCount <> shard.DocumentIds.Count Then Throw New System.IO.InvalidDataException("Invalid immutable document shard membership.")
                For Each identity As System.String In shard.DocumentIds
                    SemanticArchiveIdentity.ValidateId(identity, NameOf(identity))
                    If lookup.ContainsKey(identity) Then Throw New System.IO.InvalidDataException("A document belongs to more than one immutable shard.")
                    lookup.Add(identity, shard)
                Next
            Next
            Return lookup
        End Function

        Public Function WriteNode(manifest As SemanticArchiveGenerationManifest, node As SemanticArchiveNode) As SemanticArchiveNodeDescriptor
            ValidateManifestHeader(manifest)
            If node Is Nothing OrElse node.Cards Is Nothing Then Throw New System.ArgumentException("A complete node is required.", NameOf(node))
            SemanticArchiveIdentity.ValidateId(node.NodeId, NameOf(node.NodeId))
            node.GenerationId = manifest.GenerationId
            Dim generationDirectory As System.String = GetGenerationDirectory(manifest.ArchiveId, manifest.GenerationId)
            Dim indexPath As System.String = SemanticArchivePathGuard.ValidateContainedPath(generationDirectory, node.IndexPath, True)
            Dim path As System.String = System.IO.Path.Combine(generationDirectory, "nodes", node.NodeId & ".json")
            Dim descriptor As New SemanticArchiveNodeDescriptor With {
                .NodeId = node.NodeId, .GenerationId = manifest.GenerationId, .ParentNodeId = node.ParentNodeId,
                .MinKey = node.MinKey, .MaxKey = node.MaxKey, .IsLeaf = node.IsLeaf,
                .ChildCount = node.Cards.Count, .IndexArtifact = MakeArtifact(indexPath)
            }
            Dim ids As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            For Each card As SemanticArchiveCard In node.Cards
                ValidateCard(card)
                If Not ids.Add(card.CardId) Then Throw New System.IO.InvalidDataException("Duplicate card identity in node.")
                If card.Level = "CONTAINER" Then descriptor.ChildNodeIds.Add(card.TargetId)
            Next
            WriteImmutableJson(path, node)
            descriptor.Artifact = MakeArtifact(path)
            descriptor.ArtifactPath = descriptor.Artifact.RelativePath
            descriptor.Sha256 = descriptor.Artifact.Sha256
            manifest.Nodes(node.NodeId) = descriptor
            Return descriptor
        End Function

        Public Function WriteDocumentShard(manifest As SemanticArchiveGenerationManifest, shard As SemanticArchiveDocumentShard) As SemanticArchiveDocumentShardDescriptor
            ValidateManifestHeader(manifest)
            If shard Is Nothing OrElse shard.Documents Is Nothing Then Throw New System.ArgumentException("A complete document shard is required.", NameOf(shard))
            SemanticArchiveIdentity.ValidateId(shard.ShardId, NameOf(shard.ShardId))
            Dim path As System.String = System.IO.Path.Combine(GetGenerationDirectory(manifest.ArchiveId, manifest.GenerationId), "documents", shard.ShardId & ".json")
            Dim descriptor As New SemanticArchiveDocumentShardDescriptor With {.ShardId = shard.ShardId, .GenerationId = manifest.GenerationId, .MinKey = shard.MinKey, .MaxKey = shard.MaxKey, .RecordCount = shard.Documents.Count}
            Dim ids As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            For Each document As SemanticArchiveDocumentRecord In shard.Documents
                ValidateDocumentRecord(document)
                If Not ids.Add(document.DocumentId) Then Throw New System.IO.InvalidDataException("Duplicate document identity in shard.")
                descriptor.DocumentIds.Add(document.DocumentId)
            Next
            descriptor.Inventory = SemanticArchiveInventory.FromDocuments(shard.Documents)
            WriteImmutableJson(path, shard)
            descriptor.Artifact = MakeArtifact(path)
            descriptor.ArtifactPath = descriptor.Artifact.RelativePath
            descriptor.Sha256 = descriptor.Artifact.Sha256
            manifest.DocumentShards.RemoveAll(Function(existing As SemanticArchiveDocumentShardDescriptor) existing.ShardId = shard.ShardId)
            manifest.DocumentShards.Add(descriptor)
            Return descriptor
        End Function

        ''' <summary>
        ''' Activates only a closed, checksum-validated generation. A storage failure
        ''' is propagated. There is intentionally no delete-then-move replacement.
        ''' </summary>
        Public Sub PublishGeneration(lease As SemanticArchiveWriterLease, manifest As SemanticArchiveGenerationManifest)
            AssertLease(lease)
            ValidateManifestHeader(manifest)
            If manifest.ArchiveId <> lease.ArchiveId OrElse manifest.FenceToken <> lease.FenceToken OrElse manifest.PreviousGenerationId <> lease.BaseGenerationId Then Throw New System.InvalidOperationException("The generation is not owned by the current writer lease.")
            Dim current As SemanticArchiveGenerationPointer = ReadPointer(lease.ArchiveId)
            If If(current Is Nothing, "", current.GenerationId) <> lease.BaseGenerationId Then Throw New System.InvalidOperationException("A newer archive generation has already been published.")
            Dim definition As SemanticArchiveDefinition = GetArchive(lease.ArchiveId)
            If definition Is Nothing Then Throw New System.InvalidOperationException("The archive was removed during the build.")
            If manifest.ConfigurationSignature <> lease.ConfigurationSignature OrElse GetConfigurationSignature(definition) <> lease.ConfigurationSignature Then Throw New System.InvalidOperationException("Archive settings changed during the build. Refresh with the current configuration.")
            ValidateGeneration(manifest, definition)
            manifest.ValidationStatus = "validated"
            Dim manifestPath As System.String = System.IO.Path.Combine(GetGenerationDirectory(manifest.ArchiveId, manifest.GenerationId), "manifest.json")
            WriteImmutableJson(manifestPath, manifest)
            Dim pointer As New SemanticArchiveGenerationPointer With {.ArchiveId = manifest.ArchiveId, .GenerationId = manifest.GenerationId, .Manifest = MakeArtifact(manifestPath), .FenceToken = lease.FenceToken}
            Using catalogMutationLock As SemanticArchiveStorageLock = AcquireStorageLock(System.IO.Path.Combine(_directoryPath, DataDirectoryName, "catalog.lock"), System.Threading.CancellationToken.None)
                AssertLease(lease)
                Dim currentDefinition As SemanticArchiveDefinition = GetArchive(lease.ArchiveId)
                If currentDefinition Is Nothing OrElse GetConfigurationSignature(currentDefinition) <> lease.ConfigurationSignature Then Throw New System.InvalidOperationException("Archive settings changed before activation. Refresh with the current configuration.")
                current = ReadPointer(lease.ArchiveId)
                If If(current Is Nothing, "", current.GenerationId) <> lease.BaseGenerationId Then Throw New System.InvalidOperationException("The writer generation is obsolete.")
                AtomicWriteJson(System.IO.Path.Combine(GetArchiveDirectory(lease.ArchiveId), "current.json"), pointer)
                Dim saved As SemanticArchiveGenerationPointer = ReadPointer(lease.ArchiveId)
                If saved Is Nothing OrElse saved.GenerationId <> manifest.GenerationId OrElse saved.FenceToken <> lease.FenceToken OrElse saved.Manifest.Sha256 <> pointer.Manifest.Sha256 Then Throw New System.IO.IOException("Archive activation could not be verified.")
                lease.MarkPublished(manifest.GenerationId)
            End Using
            ' Retain every prior generation. Automatic deletion would need cross-host
            ' reader retention accounting and is intentionally not inferred here.
        End Sub

        Public Function ResolveArtifactPath(reference As SemanticArchiveArtifactReference) As System.String
            If reference Is Nothing OrElse System.String.IsNullOrWhiteSpace(reference.RelativePath) OrElse System.IO.Path.IsPathRooted(reference.RelativePath) Then Throw New System.IO.InvalidDataException("An immutable artifact must have a catalog-relative path.")
            Return SemanticArchivePathGuard.ValidateContainedPath(_directoryPath, System.IO.Path.Combine(_directoryPath, reference.RelativePath), True)
        End Function

        Public Function MakeArtifact(path As System.String) As SemanticArchiveArtifactReference
            Dim full As System.String = SemanticArchivePathGuard.ValidateContainedPath(_directoryPath, path, True)
            Using stream As System.IO.FileStream = SemanticArchivePathGuard.OpenContainedRead(_directoryPath, full)
                Using hasher As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                    Dim hash As System.String = System.BitConverter.ToString(hasher.ComputeHash(stream)).Replace("-", "").ToLowerInvariant()
                    Return New SemanticArchiveArtifactReference With {.RelativePath = full.Substring(_directoryPath.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar).Length + 1), .Sha256 = hash, .Length = stream.Length}
                End Using
            End Using
        End Function

        Friend Sub AssertLease(lease As SemanticArchiveWriterLease)
            If lease Is Nothing OrElse Not System.Object.ReferenceEquals(lease.Store, Me) Then Throw New System.InvalidOperationException("An archive writer lease is required.")
            lease.AssertHandle()
            Dim fence As SemanticArchiveFenceRecord = ReadJson(Of SemanticArchiveFenceRecord)(System.IO.Path.Combine(GetArchiveDirectory(lease.ArchiveId), "writer-fence.json"), 16384)
            If fence.Token <> lease.FenceToken OrElse fence.LeaseId <> lease.LeaseId Then Throw New System.InvalidOperationException("This archive writer has been fenced out.")
        End Sub

        Private Function ReadPointer(archiveId As System.String) As SemanticArchiveGenerationPointer
            Dim path As System.String = System.IO.Path.Combine(GetArchiveDirectory(archiveId), "current.json")
            If Not ExistsChecked(path) Then Return Nothing
            Dim pointer As SemanticArchiveGenerationPointer = ReadJson(Of SemanticArchiveGenerationPointer)(path, 65536)
            If pointer Is Nothing OrElse pointer.SchemaVersion <> GenerationSchemaVersion OrElse pointer.ArchiveId <> archiveId OrElse pointer.FenceToken <= 0 OrElse pointer.Manifest Is Nothing Then Throw New System.IO.InvalidDataException("Unsupported archive generation pointer. This storage revision requires a clean archive rebuild; old generations are not migrated.")
            SemanticArchiveIdentity.ValidateId(pointer.GenerationId, NameOf(pointer.GenerationId))
            Dim expected As System.String = System.IO.Path.Combine(GetGenerationDirectory(archiveId, pointer.GenerationId), "manifest.json")
            If Not System.String.Equals(ResolveArtifactPath(pointer.Manifest), expected, System.StringComparison.OrdinalIgnoreCase) Then Throw New System.IO.InvalidDataException("The active pointer does not address its exact immutable manifest.")
            Return pointer
        End Function

        Private Sub RegisterCentralOutputs()
            GeneratedOutputRegistry.Register(CatalogPath, "SemanticArchive:" & _directoryPath)
            GeneratedOutputRegistry.Register(System.IO.Path.Combine(_directoryPath, DataDirectoryName), "SemanticArchive:" & _directoryPath)
        End Sub

        Private Sub RegisterOutputs(catalog As SemanticArchiveCatalog)
            RegisterCentralOutputs()
            For Each definition As SemanticArchiveDefinition In catalog.Archives
                Dim owner As System.String = "SemanticArchive:" & _directoryPath & ":" & definition.ArchiveId
                GeneratedOutputRegistry.Register(GetWorkDirectory(definition.ArchiveId), owner)
                For Each binding As SemanticArchiveSourceBinding In definition.Roots
                    GeneratedOutputRegistry.Register(GetDefaultShadowRoot(binding), owner)
                    ' Optional old/configured roots may be offline. They are registered
                    ' before any new write; existing durable registrations remain.
                    For Each root As System.String In New System.String() {binding.ShadowArtifactRoot}
                        If System.String.IsNullOrWhiteSpace(root) Then Continue For
                        Try
                            GeneratedOutputRegistry.Register(root, owner)
                        Catch failure As System.Exception When TypeOf failure Is System.IO.IOException OrElse TypeOf failure Is System.UnauthorizedAccessException
                            System.Diagnostics.Trace.TraceWarning("Semantic archive output registration deferred: " & failure.Message)
                        End Try
                    Next
                Next
            Next
        End Sub

        Public Shared Function GetDerivedRoot(binding As SemanticArchiveSourceBinding) As System.String
            If binding Is Nothing Then Throw New System.ArgumentNullException(NameOf(binding))
            If Not System.String.IsNullOrWhiteSpace(binding.ShadowArtifactRoot) Then Return SemanticArchivePathGuard.CanonicalPath(binding.ShadowArtifactRoot)
            Return GetDefaultShadowRoot(binding)
        End Function

        Public Shared Function GetDefaultShadowRoot(binding As SemanticArchiveSourceBinding) As System.String
            If binding Is Nothing Then Throw New System.ArgumentNullException(NameOf(binding))
            Dim local As System.String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData)
            If System.String.IsNullOrWhiteSpace(local) Then Throw New System.IO.IOException("A private local application-data directory is unavailable.")
            Return SemanticArchivePathGuard.CanonicalPath(System.IO.Path.Combine(local, "RedInk", "SA", "shadow", SemanticArchiveIdentity.HashBytes(System.Text.Encoding.UTF8.GetBytes(binding.BindingId & "|" & binding.RootPath)).Substring(0, 16)))
        End Function


        Public Shared Iterator Function GetPrivateDerivedRoots(binding As SemanticArchiveSourceBinding) As System.Collections.Generic.IEnumerable(Of System.String)
            Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            For Each root As System.String In New System.String() {GetDerivedRoot(binding), GetDefaultShadowRoot(binding)}
                If seen.Add(root) Then Yield root
            Next
        End Function

        Public Shared Function RequirePrivateDerivedArtifact(binding As SemanticArchiveSourceBinding, path As System.String) As System.String
            For Each root As System.String In GetPrivateDerivedRoots(binding)
                Dim versions As System.String = System.IO.Path.Combine(root, "versions")
                If SemanticArchivePathGuard.IsContainedPath(versions, path) Then
                    Dim full As System.String = SemanticArchivePathGuard.ValidateContainedPath(versions, path, True)
                    RequirePrivateArtifact(full)
                    Return full
                End If
            Next
            Throw New System.UnauthorizedAccessException("The private artifact is outside registered immutable version locations.")
        End Function

        Private Shared Function GetConfigurationSignature(definition As SemanticArchiveDefinition) As System.String
            Dim configuration As SemanticArchiveDefinition = SemanticArchiveMetadata.Clone(definition)
            configuration.Library = Nothing
            Return SemanticArchiveIdentity.HashBytes(System.Text.Encoding.UTF8.GetBytes(Newtonsoft.Json.JsonConvert.SerializeObject(configuration, Newtonsoft.Json.Formatting.None, JsonSettings)))
        End Function

        Private Shared Sub EnsureDirectory(path As System.String)
            Dim full As System.String = SemanticArchivePathGuard.CanonicalPath(path)
            SemanticArchivePathGuard.ValidateContainedPath(full, full, False)
            System.IO.Directory.CreateDirectory(full)
            SemanticArchivePathGuard.ValidateContainedPath(full, full, True)
        End Sub

        Private Shared Function ExistsChecked(path As System.String) As System.Boolean
            Try
                System.IO.File.GetAttributes(path)
                Return True
            Catch ex As System.IO.FileNotFoundException
                Return False
            Catch ex As System.IO.DirectoryNotFoundException
                Return False
            End Try
        End Function

        Private Shared Function ReadJson(Of T As Class)(path As System.String, maximumBytes As System.Int64) As T
            Dim parent As System.String = System.IO.Path.GetDirectoryName(path)
            SemanticArchivePathGuard.ValidateContainedPath(parent, path, True)
            RequirePrivateArtifact(path)
            ' Delete sharing lets a reader pin the complete old or complete new file
            ' while an atomic replacement is taking place.
            Using stream As New System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read Or System.IO.FileShare.Delete)
                If stream.Length <= 0 OrElse stream.Length > maximumBytes Then Throw New System.IO.InvalidDataException("An archive control record has an invalid size.")
                Using reader As New System.IO.StreamReader(stream, New System.Text.UTF8Encoding(False, True), True)
                    Dim value As T = Newtonsoft.Json.JsonConvert.DeserializeObject(Of T)(reader.ReadToEnd(), JsonSettings)
                    If value Is Nothing Then Throw New System.IO.InvalidDataException("An archive control record is empty.")
                    Return value
                End Using
            End Using
        End Function

        Private Shared Sub WriteImmutableJson(path As System.String, value As System.Object)
            Dim bytes As System.Byte() = New System.Text.UTF8Encoding(False, True).GetBytes(Newtonsoft.Json.JsonConvert.SerializeObject(value, Newtonsoft.Json.Formatting.None, JsonSettings))
            CreatePrivateDirectory(System.IO.Path.GetDirectoryName(path))
            If ExistsChecked(path) Then
                RequirePrivateArtifact(path)
                If SemanticArchiveIdentity.ComputeFileHash(path) = SemanticArchiveIdentity.HashBytes(bytes) Then Return
                Throw New System.IO.IOException("An immutable archive artifact already exists with different bytes.")
            End If
            Using stream As System.IO.FileStream = CreatePrivateFile(path)
                stream.Write(bytes, 0, bytes.Length)
                stream.Flush(True)
            End Using
            If SemanticArchiveIdentity.ComputeFileHash(path) <> SemanticArchiveIdentity.HashBytes(bytes) Then Throw New System.IO.IOException("The immutable archive write failed verification.")
            RequirePrivateArtifact(path)
        End Sub

        ''' <summary>For host-owned mutable work/control records only.</summary>
        Public Shared Sub AtomicWriteJson(path As System.String, value As System.Object)
            If value Is Nothing Then Throw New System.ArgumentNullException(NameOf(value))
            Dim full As System.String = SemanticArchivePathGuard.RequireWindowsCompatiblePath(path)
            RequireAtomicWritePathBudget(full)
            Dim parent As System.String = System.IO.Path.GetDirectoryName(full)
            EnsureDirectory(parent)
            SemanticArchivePathGuard.ValidateContainedPath(parent, full, False)
            Dim bytes As System.Byte() = New System.Text.UTF8Encoding(False, True).GetBytes(Newtonsoft.Json.JsonConvert.SerializeObject(value, Newtonsoft.Json.Formatting.None, JsonSettings))
            If ExistsChecked(full) Then RequirePrivateArtifact(full)
            ' A protected sibling temporary file stays on the target filesystem.
            ' No permanent staging directory is required for an atomic replacement.
            Dim temporary As System.String = System.IO.Path.Combine(parent, ".sa-" & SemanticArchiveIdentity.NewId() & ".tmp")
            Try
                Using stream As System.IO.FileStream = CreatePrivateFile(temporary)
                    stream.Write(bytes, 0, bytes.Length)
                    stream.Flush(True)
                End Using
                SemanticArchivePathGuard.ValidateContainedPath(parent, full, False)
                If ExistsChecked(full) Then
                    System.IO.File.Replace(temporary, full, Nothing, True)
                Else
                    System.IO.File.Move(temporary, full)
                End If
                If SemanticArchiveIdentity.ComputeFileHash(full) <> SemanticArchiveIdentity.HashBytes(bytes) Then Throw New System.IO.IOException("Atomic archive publication failed verification.")
                RequirePrivateArtifact(full)
            Finally
                ' Only the privately generated temporary artifact is removable.
                If ExistsChecked(temporary) Then System.IO.File.Delete(temporary)
            End Try
        End Sub

        Private Shared Function AcquireStorageLock(path As System.String,
                                                   cancellationToken As System.Threading.CancellationToken,
                                                   Optional maximumWait As System.TimeSpan? = Nothing) As SemanticArchiveStorageLock
            Dim waitBudget As System.TimeSpan = If(maximumWait.HasValue, maximumWait.Value, System.TimeSpan.FromSeconds(SharedMethods.DEFAULT_SEMANTICARCHIVE_WRITER_LEASE_WAIT_SECONDS))
            If waitBudget < System.TimeSpan.Zero OrElse waitBudget > System.TimeSpan.FromSeconds(30) Then
                Throw New System.ArgumentOutOfRangeException(NameOf(maximumWait), "A storage-lock wait must be between zero and thirty seconds.")
            End If
            SemanticArchivePathGuard.RequireWindowsCompatiblePath(path)
            EnsureDirectory(System.IO.Path.GetDirectoryName(path))
            Dim elapsed As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew()
            Do
                cancellationToken.ThrowIfCancellationRequested()
                Try
                    Return New SemanticArchiveStorageLock(path)
                Catch ex As System.IO.IOException When IsStorageLockContention(ex)
                    cancellationToken.ThrowIfCancellationRequested()
                    If elapsed.Elapsed >= waitBudget Then
                        Throw New System.TimeoutException("The storage writer is busy; the bounded acquisition attempt yielded without acquiring or replacing its lease.", ex)
                    End If
                    ' Bound only the retry wait. A synchronous filesystem call must finish safely;
                    ' it is never abandoned and a timeout never authorizes lock takeover.
                    Dim remainingMilliseconds As System.Int32 = CInt(System.Math.Ceiling((waitBudget - elapsed.Elapsed).TotalMilliseconds))
                    cancellationToken.WaitHandle.WaitOne(System.Math.Min(100, System.Math.Max(1, remainingMilliseconds)))
                End Try
            Loop
        End Function

        Private Shared Function IsStorageLockContention(failure As System.IO.IOException) As System.Boolean
            ' Win32 ERROR_SHARING_VIOLATION / ERROR_LOCK_VIOLATION. Filesystem, permission,
            ' path and unsupported-lock errors must remain visible failures, not busy retries.
            Dim nativeError As System.Int32 = failure.HResult And &HFFFF
            Return nativeError = 32 OrElse nativeError = 33
        End Function
    End Class

    Friend NotInheritable Class SemanticArchiveFenceRecord
        Public Property Token As System.Int64
        Public Property LeaseId As System.String = ""
    End Class

    Friend NotInheritable Class SemanticArchiveStorageLock
        Implements System.IDisposable
        Private _handle As System.IO.FileStream

        Public Sub New(path As System.String)
            SemanticArchivePathGuard.ValidateContainedPath(System.IO.Path.GetDirectoryName(path), path, False)
            Dim stream As New System.IO.FileStream(path, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.ReadWrite, System.IO.FileShare.None, 4096, System.IO.FileOptions.WriteThrough)
            Try
                ' No timestamp-based takeover exists. The filesystem must release
                ' this exclusive OS/share lock before another lease can be issued.
                stream.Lock(0, 1)
                stream.Flush(True)
                Dim conflictingHandle As System.IO.FileStream = Nothing
                Try
                    conflictingHandle = New System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite, System.IO.FileShare.ReadWrite)
                Catch ex As System.IO.IOException
                    ' The existing exclusive share/OS lock must reject this open.
                End Try
                If conflictingHandle IsNot Nothing Then
                    conflictingHandle.Dispose()
                    Throw New System.IO.IOException("The filesystem did not enforce exclusive archive writer sharing.")
                End If
                _handle = stream
            Catch ex As System.Exception
                stream.Dispose()
                Throw
            End Try
        End Sub

        Public Sub AssertHandle()
            If _handle Is Nothing OrElse Not _handle.CanWrite Then Throw New System.InvalidOperationException("The archive writer lease has ended.")
            _handle.Flush(True)
        End Sub

        Public Sub Dispose() Implements System.IDisposable.Dispose
            Dim handle As System.IO.FileStream = _handle
            _handle = Nothing
            If handle IsNot Nothing Then handle.Dispose()
        End Sub
    End Class

    Public NotInheritable Class SemanticArchiveWriterLease
        Implements System.IDisposable
        Private ReadOnly _lock As SemanticArchiveStorageLock
        Private _disposed As System.Boolean
        Private _baseGenerationId As System.String
        Friend ReadOnly Property Store As SemanticArchiveStore
        Public ReadOnly Property ArchiveId As System.String
        Public ReadOnly Property FenceToken As System.Int64
        Public ReadOnly Property LeaseId As System.String
        Public ReadOnly Property ConfigurationSignature As System.String
        Public ReadOnly Property BaseGenerationId As System.String
            Get
                Return _baseGenerationId
            End Get
        End Property

        Friend Sub New(owner As SemanticArchiveStore, storageLock As SemanticArchiveStorageLock, archiveId As System.String, fenceToken As System.Int64, leaseId As System.String, baseGenerationId As System.String, configurationSignature As System.String)
            Store = owner
            _lock = storageLock
            Me.ArchiveId = archiveId
            Me.FenceToken = fenceToken
            Me.LeaseId = leaseId
            Me.ConfigurationSignature = configurationSignature
            _baseGenerationId = baseGenerationId
        End Sub

        Friend Sub AssertHandle()
            If _disposed Then Throw New System.ObjectDisposedException(NameOf(SemanticArchiveWriterLease))
            _lock.AssertHandle()
        End Sub

        Friend Sub MarkPublished(generationId As System.String)
            _baseGenerationId = generationId
        End Sub

        Public Sub Dispose() Implements System.IDisposable.Dispose
            If _disposed Then Return
            _disposed = True
            _lock.Dispose()
        End Sub
    End Class
End Namespace
