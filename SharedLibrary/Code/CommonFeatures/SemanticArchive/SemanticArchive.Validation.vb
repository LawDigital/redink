' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' Strict immutable/reference validation and generation-independent live validity.
Option Strict On
Option Explicit On

Namespace SharedLibrary

    Public NotInheritable Partial Class SemanticArchiveStore
        Private Shared ReadOnly CacheGate As New System.Object()
        Private Shared ReadOnly ImmutableJsonCache As New System.Collections.Generic.Dictionary(Of System.String, System.Byte())(System.StringComparer.Ordinal)
        Private Shared ReadOnly ImmutableCacheOrder As New System.Collections.Generic.Queue(Of System.String)()
        Private Shared ImmutableCacheBytes As System.Int64
        Private Const MaximumCachedJsonBytes As System.Int32 = 1024 * 1024
        Private Const MaximumCacheBytes As System.Int64 = 16L * 1024L * 1024L
        Private Const MaximumCacheEntries As System.Int32 = 128

        Friend Shared Sub ValidateLibraryDefinition(definition As SemanticArchiveDefinition)
            Dim catalog As New SemanticArchiveCatalog()
            catalog.Archives.Add(definition)
            ValidateCatalog(catalog)
        End Sub

        Private Shared Sub ValidateCatalog(catalog As SemanticArchiveCatalog)
            If catalog Is Nothing OrElse catalog.SchemaVersion <> 1 OrElse catalog.Revision < 0 Then Throw New System.IO.InvalidDataException("Unsupported or invalid Semantic Archives catalog.")
            If catalog.Archives Is Nothing Then catalog.Archives = New System.Collections.Generic.List(Of SemanticArchiveDefinition)()
            If catalog.DefaultArchiveIds Is Nothing Then catalog.DefaultArchiveIds = New System.Collections.Generic.List(Of System.String)()
            Dim archives As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            Dim exactArchiveIds As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            For Each archive As SemanticArchiveDefinition In catalog.Archives
                If archive Is Nothing Then Throw New System.IO.InvalidDataException("Null archive definition.")
                SemanticArchiveIdentity.ValidateId(archive.ArchiveId, NameOf(archive.ArchiveId))
                If Not archives.Add(archive.ArchiveId) Then Throw New System.IO.InvalidDataException("Duplicate stable archive identity.")
                exactArchiveIds.Add(archive.ArchiveId)
                If System.String.IsNullOrWhiteSpace(archive.Name) Then Throw New System.IO.InvalidDataException("An archive display name is required.")
                archive.Name = archive.Name.Trim()
                If archive.Roots Is Nothing Then archive.Roots = New System.Collections.Generic.List(Of SemanticArchiveSourceBinding)()
                If archive.RetrievalBudgets Is Nothing Then archive.RetrievalBudgets = New SemanticArchiveRetrievalBudgets()
                If archive.MaxChildrenPerNode < 2 OrElse archive.MaxChildrenPerNode > 256 OrElse archive.MaxRoutingCharacters < 4096 OrElse archive.MaxRoutingCharacters > 120000 OrElse archive.SectionIndexThresholdBytes < 0 Then Throw New System.IO.InvalidDataException("Archive hierarchy limits are invalid.")
                Dim budgets As SemanticArchiveRetrievalBudgets = archive.RetrievalBudgets
                If budgets.MaxNodesVisited < 1 OrElse budgets.MaxModelCalls < 1 OrElse budgets.MaxElapsedSeconds < 1 OrElse budgets.MaxCandidateFiles < 1 OrElse budgets.MaxEvidenceBytes < 1 OrElse budgets.InitialBranches < 1 OrElse budgets.MaxExactLookupDocuments < 0 OrElse budgets.MaxSectionCandidates < 1 OrElse budgets.MaxPromptCharacters < 1024 OrElse budgets.MaxRequestTokens < 1 OrElse budgets.MaxRequestTokens > 262144 OrElse budgets.MaxLiteralScanBytes < 1 OrElse budgets.MaxLiteralScanBytes > 67108864 OrElse budgets.MaxLiteralScanDocuments < 1 OrElse budgets.MaxLiteralScanDocuments > 1024 Then Throw New System.IO.InvalidDataException("Archive retrieval budgets are invalid.")
                Dim bindings As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
                For Each binding As SemanticArchiveSourceBinding In archive.Roots
                    If binding Is Nothing Then Throw New System.IO.InvalidDataException("Null source binding.")
                    SemanticArchiveIdentity.ValidateId(binding.BindingId, NameOf(binding.BindingId))
                    If Not bindings.Add(binding.BindingId) Then Throw New System.IO.InvalidDataException("Duplicate source binding identity.")
                    binding.RootPath = SemanticArchivePathGuard.CanonicalPath(binding.RootPath)
                    If binding.OcrBatchPages < 1 OrElse binding.OcrBatchPages > 75 Then Throw New System.IO.InvalidDataException("SemanticArchiveOcrBatchPages must be between 1 and 75.")
                    If binding.ArtifactPlacementMode <> "auto" AndAlso binding.ArtifactPlacementMode <> "private" Then Throw New System.IO.InvalidDataException("Unknown artifact placement mode.")
                    If Not System.String.IsNullOrWhiteSpace(binding.SharedArtifactRoot) Then binding.SharedArtifactRoot = SemanticArchivePathGuard.CanonicalPath(binding.SharedArtifactRoot)
                    If Not System.String.IsNullOrWhiteSpace(binding.ShadowArtifactRoot) Then binding.ShadowArtifactRoot = SemanticArchivePathGuard.CanonicalPath(binding.ShadowArtifactRoot)
                    If binding.Exclusions Is Nothing Then binding.Exclusions = New System.Collections.Generic.List(Of System.String)()
                    If binding.SupportedExtensions Is Nothing Then binding.SupportedExtensions = New System.Collections.Generic.List(Of System.String)()
                    If binding.ScopeTags Is Nothing Then binding.ScopeTags = New System.Collections.Generic.List(Of System.String)()
                    Dim normalizedExtensions As New System.Collections.Generic.List(Of System.String)()
                    For Each extension As System.String In binding.SupportedExtensions
                        Dim normalized As System.String = NormalizeSourceExtension(extension)
                        If normalized.Length > 0 AndAlso Not normalizedExtensions.Contains(normalized) Then normalizedExtensions.Add(normalized)
                    Next
                    binding.SupportedExtensions = normalizedExtensions
                    If SemanticArchivePathGuard.IsContainedPath(GetDerivedRoot(binding), binding.RootPath) Then Throw New System.IO.InvalidDataException("A generated-output directory may not contain its source root.")
                Next
            Next
            Dim defaults As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            For Each archiveId As System.String In catalog.DefaultArchiveIds
                If Not exactArchiveIds.Contains(archiveId) OrElse Not defaults.Add(archiveId) Then Throw New System.IO.InvalidDataException("Invalid or duplicate default archive scope.")
            Next
        End Sub

        Private Shared Sub ValidateManifestHeader(manifest As SemanticArchiveGenerationManifest)
            If manifest Is Nothing OrElse manifest.SchemaVersion <> GenerationSchemaVersion OrElse manifest.Nodes Is Nothing OrElse manifest.DocumentShards Is Nothing OrElse manifest.TermShards Is Nothing OrElse manifest.Inventory Is Nothing OrElse manifest.FenceToken <= 0 Then Throw New System.IO.InvalidDataException("Invalid archive generation descriptor.")
            If manifest.ValidationStatus = "validated" AndAlso manifest.RoutingGraphArtifact Is Nothing Then
                Throw New System.IO.InvalidDataException("unsupported_semantic_index: This published archive has no current routing graph. Create a new archive index; legacy index migration is not supported.")
            End If
            SemanticArchiveIdentity.ValidateId(manifest.ArchiveId, NameOf(manifest.ArchiveId))
            SemanticArchiveIdentity.ValidateId(manifest.GenerationId, NameOf(manifest.GenerationId))
            manifest.Inventory.Validate()
            If manifest.ValidationStatus = "validated" AndAlso
                (manifest.Inventory.CurrentSources + manifest.Inventory.RemovedSources <> manifest.TotalDocumentCount OrElse
                 manifest.Inventory.SearchableDocuments <> manifest.DocumentCount OrElse manifest.Inventory.FailedSources <> manifest.FailureCount) Then
                Throw New System.IO.InvalidDataException("Published generation counts are inconsistent; rebuild the archive.")
            End If
        End Sub

        Private Shared Sub ValidateCard(card As SemanticArchiveCard)
            If card Is Nothing Then Throw New System.IO.InvalidDataException("A semantic card is missing.")
            SemanticArchiveIdentity.ValidateId(card.CardId, NameOf(card.CardId))
            SemanticArchiveIdentity.ValidateId(card.TargetId, NameOf(card.TargetId))
            If card.Level <> "CONTAINER" AndAlso card.Level <> "DOCUMENT" AndAlso card.Level <> "SECTION" Then Throw New System.IO.InvalidDataException("Unknown semantic card level.")
            If card.StartByte < 0 OrElse card.LengthBytes < 0 Then Throw New System.IO.InvalidDataException("A semantic card has an invalid content-relative byte range.")
            If card.Topics Is Nothing Then card.Topics = New System.Collections.Generic.List(Of System.String)()
            If card.UserIntents Is Nothing Then card.UserIntents = New System.Collections.Generic.List(Of System.String)()
            If card.Identifiers Is Nothing Then card.Identifiers = New System.Collections.Generic.List(Of System.String)()
            If card.ExactTerms Is Nothing Then card.ExactTerms = New System.Collections.Generic.List(Of System.String)()
        End Sub

        Private Shared Sub ValidateDocumentRecord(document As SemanticArchiveDocumentRecord)
            If document Is Nothing OrElse document.Fingerprint Is Nothing OrElse document.BindingIds Is Nothing Then Throw New System.IO.InvalidDataException("A document provenance record is incomplete.")
            SemanticArchiveIdentity.ValidateId(document.DocumentId, NameOf(document.DocumentId))
            If System.String.IsNullOrWhiteSpace(document.CanonicalSourceKey) OrElse System.String.IsNullOrWhiteSpace(document.SourcePath) OrElse document.BindingIds.Count = 0 OrElse document.Fingerprint.Length < 0 Then Throw New System.IO.InvalidDataException("A document has no authoritative source identity.")
            SemanticArchivePathGuard.CanonicalPath(document.SourcePath)
            For Each bindingId As System.String In document.BindingIds
                SemanticArchiveIdentity.ValidateId(bindingId, NameOf(bindingId))
            Next
            If document.Representation IsNot Nothing Then
                SemanticArchiveIdentity.ValidateId(document.Representation.RepresentationId, NameOf(document.Representation.RepresentationId))
                If Not IsHash(document.Representation.SourceHash) OrElse Not IsHash(document.Representation.TextFileHash) OrElse document.Representation.TextByteLength < 0 OrElse System.String.IsNullOrWhiteSpace(document.Representation.TextPath) Then Throw New System.IO.InvalidDataException("Invalid immutable representation provenance.")
                If Not System.String.Equals(document.Representation.SourceHash, document.Fingerprint.Sha256, System.StringComparison.OrdinalIgnoreCase) Then Throw New System.IO.InvalidDataException("The representation source hash is not the source snapshot hash.")
            End If
            If document.Index IsNot Nothing Then
                If document.Representation Is Nothing OrElse document.Index.RepresentationId <> document.Representation.RepresentationId OrElse document.Index.FormatVersion <> 1 OrElse Not IsHash(document.Index.FileHash) OrElse Not IsHash(document.Index.PayloadHash) OrElse document.Index.EntryCount < 1 Then Throw New System.IO.InvalidDataException("Invalid per-document semantic index descriptor.")
            End If
            If document.Card IsNot Nothing Then
                ValidateCard(document.Card)
                If document.Card.Level <> "DOCUMENT" OrElse document.Card.TargetId <> document.DocumentId Then Throw New System.IO.InvalidDataException("A document card does not resolve to its authoritative document.")
            End If
        End Sub

        Private Shared Function IsHash(value As System.String) As System.Boolean
            If value Is Nothing OrElse value.Length <> 64 Then Return False
            For Each character As System.Char In value
                If Not ((character >= "0"c AndAlso character <= "9"c) OrElse (character >= "a"c AndAlso character <= "f"c) OrElse (character >= "A"c AndAlso character <= "F"c)) Then Return False
            Next
            Return True
        End Function

        Private Sub ValidateArtifact(reference As SemanticArchiveArtifactReference, archiveId As System.String)
            ValidateArtifactLocation(reference, archiveId)
            Dim path As System.String = ResolveArtifactPath(reference)
            Using stream As System.IO.FileStream = SemanticArchivePathGuard.OpenContainedRead(GetArchiveDirectory(archiveId), path)
                If stream.Length <> reference.Length Then Throw New System.IO.InvalidDataException("An immutable artifact length does not match its descriptor.")
                Using hasher As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                    Dim hash As System.String = System.BitConverter.ToString(hasher.ComputeHash(stream)).Replace("-", "").ToLowerInvariant()
                    If Not System.String.Equals(hash, reference.Sha256, System.StringComparison.OrdinalIgnoreCase) Then Throw New System.IO.InvalidDataException("An immutable archive artifact checksum is invalid.")
                End Using
            End Using
        End Sub

        Private Sub ValidateArtifactLocation(reference As SemanticArchiveArtifactReference, archiveId As System.String)
            If reference Is Nothing OrElse Not IsHash(reference.Sha256) OrElse reference.Length < 0 Then Throw New System.IO.InvalidDataException("An immutable artifact reference is incomplete.")
            Dim path As System.String = ResolveArtifactPath(reference)
            SemanticArchivePathGuard.ValidateContainedPath(System.IO.Path.Combine(GetArchiveDirectory(archiveId), "generations"), path, True)
            RequirePrivateArtifact(path)
        End Sub

        Private Function ReadArtifactJson(Of T As Class)(reference As SemanticArchiveArtifactReference, archiveId As System.String, maximumBytes As System.Int32) As T
            ValidateArtifactLocation(reference, archiveId)
            If reference.Length <= 0 OrElse reference.Length > maximumBytes Then Throw New System.IO.InvalidDataException("An immutable JSON artifact exceeds its record budget.")
            Dim path As System.String = ResolveArtifactPath(reference)
            Dim key As System.String = path & "|" & reference.Sha256.ToLowerInvariant() & "|" & reference.Length.ToString(System.Globalization.CultureInfo.InvariantCulture)
            Dim bytes As System.Byte() = Nothing
            SyncLock CacheGate
                ImmutableJsonCache.TryGetValue(key, bytes)
            End SyncLock
            If bytes Is Nothing Then
                Using stream As System.IO.FileStream = SemanticArchivePathGuard.OpenContainedRead(GetArchiveDirectory(archiveId), path)
                    If stream.Length <> reference.Length Then Throw New System.IO.InvalidDataException("Immutable JSON length mismatch.")
                    bytes = New System.Byte(CInt(stream.Length) - 1) {}
                    Dim offset As System.Int32 = 0
                    While offset < bytes.Length
                        Dim count As System.Int32 = stream.Read(bytes, offset, bytes.Length - offset)
                        If count = 0 Then Throw New System.IO.EndOfStreamException("Incomplete immutable JSON artifact.")
                        offset += count
                    End While
                End Using
                If Not System.String.Equals(SemanticArchiveIdentity.HashBytes(bytes), reference.Sha256, System.StringComparison.OrdinalIgnoreCase) Then Throw New System.IO.InvalidDataException("Immutable JSON checksum mismatch.")
                If bytes.Length <= MaximumCachedJsonBytes Then
                    SyncLock CacheGate
                        If Not ImmutableJsonCache.ContainsKey(key) Then
                            While ImmutableJsonCache.Count >= MaximumCacheEntries OrElse ImmutableCacheBytes + bytes.Length > MaximumCacheBytes
                                Dim oldest As System.String = ImmutableCacheOrder.Dequeue()
                                Dim removed As System.Byte() = ImmutableJsonCache(oldest)
                                ImmutableJsonCache.Remove(oldest)
                                ImmutableCacheBytes -= removed.Length
                            End While
                            ImmutableJsonCache.Add(key, bytes)
                            ImmutableCacheOrder.Enqueue(key)
                            ImmutableCacheBytes += bytes.Length
                        End If
                    End SyncLock
                End If
            End If
            ' Freshly deserialize the verified bytes: a builder may edit the returned
            ' object without mutating a pinned reader's cache entry.
            Dim value As T = Newtonsoft.Json.JsonConvert.DeserializeObject(Of T)(New System.Text.UTF8Encoding(False, True).GetString(bytes), JsonSettings)
            If value Is Nothing Then Throw New System.IO.InvalidDataException("An immutable JSON record is empty.")
            Return value
        End Function

        Private Sub ValidateGeneration(manifest As SemanticArchiveGenerationManifest, definition As SemanticArchiveDefinition)
            If System.String.IsNullOrWhiteSpace(manifest.RootNodeId) OrElse Not manifest.Nodes.ContainsKey(manifest.RootNodeId) Then Throw New System.IO.InvalidDataException("The generation has no root navigation node.")
            If manifest.MaxChildrenPerNode <> definition.MaxChildrenPerNode OrElse manifest.MaxRoutingCharacters <> definition.MaxRoutingCharacters Then Throw New System.IO.InvalidDataException("The generation routing bounds do not match its current configuration.")
            Dim previous As SemanticArchiveGenerationManifest = ReadPinnedGeneration(manifest.ArchiveId)
            Dim documentIds As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            Dim shardIds As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            Dim changedDocuments As New System.Collections.Generic.Dictionary(Of System.String, SemanticArchiveDocumentRecord)(System.StringComparer.Ordinal)
            Dim inventory As New SemanticArchiveInventory()
            For Each shard As SemanticArchiveDocumentShardDescriptor In manifest.DocumentShards
                If shard Is Nothing OrElse shard.DocumentIds Is Nothing OrElse Not shardIds.Add(shard.ShardId) OrElse shard.DocumentIds.Count <> shard.RecordCount OrElse shard.RecordCount > definition.MaxChildrenPerNode Then Throw New System.IO.InvalidDataException("Invalid or duplicate document shard descriptor.")
                SemanticArchiveIdentity.ValidateId(shard.ShardId, NameOf(shard.ShardId))
                ValidateArtifactLocation(shard.Artifact, manifest.ArchiveId)
                If Not manifest.Nodes.ContainsKey(shard.ShardId) OrElse Not manifest.Nodes(shard.ShardId).IsLeaf Then Throw New System.IO.InvalidDataException("A document shard must identify its structural leaf node.")
                For Each documentId As System.String In shard.DocumentIds
                    SemanticArchiveIdentity.ValidateId(documentId, NameOf(documentId))
                    If Not documentIds.Add(documentId) Then Throw New System.IO.InvalidDataException("A document identity occurs in more than one shard.")
                Next
                If Not IsRetainedShard(previous, shard) Then
                    Dim data As SemanticArchiveDocumentShard = LoadDocumentShard(manifest, shard)
                    For Each document As SemanticArchiveDocumentRecord In data.Documents
                        changedDocuments.Add(document.DocumentId, document)
                        If document.Active AndAlso LiveValidityMatches(manifest, document) AndAlso document.Representation IsNot Nothing Then
                            ValidateDerivedPath(definition, document, document.Representation.TextPath, document.Representation.TextFileHash, document.Representation.TextByteLength)
                            If document.Index IsNot Nothing Then ValidateDerivedPath(definition, document, document.Index.Path, document.Index.FileHash, -1)
                        End If
                    Next
                End If
                inventory.Add(shard.Inventory)
                If shard.Inventory.CurrentSources + shard.Inventory.RemovedSources <> shard.RecordCount Then Throw New System.IO.InvalidDataException("Document shard inventory size is inconsistent.")
            Next
            Dim changedNodes As New System.Collections.Generic.List(Of SemanticArchiveNodeDescriptor)()
            Dim physicalNodeIds As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, SemanticArchiveNodeDescriptor) In manifest.Nodes
                Dim descriptor As SemanticArchiveNodeDescriptor = pair.Value
                If descriptor Is Nothing OrElse descriptor.NodeId <> pair.Key OrElse Not physicalNodeIds.Add(descriptor.NodeId) OrElse descriptor.ChildNodeIds Is Nothing OrElse descriptor.ChildCount < 0 OrElse descriptor.ChildCount > definition.MaxChildrenPerNode Then Throw New System.IO.InvalidDataException("Invalid navigation node descriptor.")
                SemanticArchiveIdentity.ValidateId(descriptor.NodeId, NameOf(descriptor.NodeId))
                ValidateArtifactLocation(descriptor.Artifact, manifest.ArchiveId)
                ValidateArtifactLocation(descriptor.IndexArtifact, manifest.ArchiveId)
                For Each child As System.String In descriptor.ChildNodeIds
                    If Not manifest.Nodes.ContainsKey(child) Then Throw New System.IO.InvalidDataException("The generation has a dangling child-node reference.")
                    If manifest.Nodes(child).ParentNodeId <> descriptor.NodeId Then Throw New System.IO.InvalidDataException("A structural child has an inconsistent parent reference.")
                Next
                If Not IsRetainedNode(previous, descriptor) Then
                    changedNodes.Add(descriptor)
                    Dim node As SemanticArchiveNode = LoadNode(manifest, descriptor.NodeId)
                    ValidateNavigationIndex(manifest, node, descriptor)
                    Dim renderedCharacters As System.Int64 = 0
                    Dim childIds As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
                    For Each card As SemanticArchiveCard In node.Cards
                        ValidateCard(card)
                        renderedCharacters += If(card.RetrievalText, "").Length
                        If card.Level = "CONTAINER" Then
                            If Not manifest.Nodes.ContainsKey(card.TargetId) OrElse Not descriptor.ChildNodeIds.Contains(card.TargetId) OrElse Not childIds.Add(card.TargetId) Then Throw New System.IO.InvalidDataException("A container card is not a valid unique node edge.")
                        ElseIf card.Level = "DOCUMENT" Then
                            If Not documentIds.Contains(card.TargetId) Then Throw New System.IO.InvalidDataException("A document card has a dangling target.")
                            Dim document As SemanticArchiveDocumentRecord = Nothing
                            If Not changedDocuments.TryGetValue(card.TargetId, document) Then document = LoadDocument(manifest, card.TargetId)
                            If document Is Nothing OrElse document.Representation Is Nothing OrElse card.RepresentationId <> document.Representation.RepresentationId OrElse Not System.String.Equals(card.SourceVersion, document.Fingerprint.Sha256, System.StringComparison.OrdinalIgnoreCase) Then Throw New System.IO.InvalidDataException("A document card's representation or source version is stale.")
                        Else
                            Throw New System.IO.InvalidDataException("Archive navigation nodes may only route to containers or documents.")
                        End If
                    Next
                    If childIds.Count <> descriptor.ChildNodeIds.Count OrElse renderedCharacters > definition.MaxRoutingCharacters Then Throw New System.IO.InvalidDataException("Navigation card mapping or routing size exceeds its contract.")
                End If
            Next
            If Not System.String.IsNullOrEmpty(manifest.Nodes(manifest.RootNodeId).ParentNodeId) Then Throw New System.IO.InvalidDataException("The root navigation node has an unexpected parent.")
            Dim visited As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            Dim visiting As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            ValidateAcyclic(manifest, manifest.RootNodeId, visiting, visited)
            If visited.Count <> manifest.Nodes.Count Then Throw New System.IO.InvalidDataException("The generation contains unreachable navigation nodes.")
            For Each reference As SemanticArchiveArtifactReference In manifest.TermShards
                ValidateArtifact(reference, manifest.ArchiveId)
            Next
            ' Current routing is mandatory. LoadRoutingGraph validates the graph once;
            ' there is no storage-tree fallback or second graph traversal.
            ValidateArtifact(manifest.RoutingGraphArtifact, manifest.ArchiveId)
            LoadRoutingGraph(manifest)
            inventory.Validate()
            manifest.Inventory = inventory
            manifest.DocumentCount = inventory.SearchableDocuments
            manifest.FailureCount = inventory.FailedSources
            manifest.TotalDocumentCount = documentIds.Count
            If manifest.DocumentCount < 0 OrElse manifest.DocumentCount > manifest.TotalDocumentCount Then Throw New System.IO.InvalidDataException("Invalid searchable document count.")
            RefreshDocumentAncestors(manifest, changedDocuments)
            RefreshNodeValidity(manifest, changedNodes)
        End Sub

        ''' <summary>
        ''' Explicit full integrity audit. Ordinary publication reuses known immutable
        ''' descriptors; this operation deliberately rehashes every retained artifact.
        ''' </summary>
        Public Sub AuditGeneration(manifest As SemanticArchiveGenerationManifest)
            ValidateManifestHeader(manifest)
            Dim definition As SemanticArchiveDefinition = GetArchive(manifest.ArchiveId)
            If definition Is Nothing Then Throw New System.InvalidOperationException("The archive is no longer registered.")
            For Each descriptor As SemanticArchiveNodeDescriptor In manifest.Nodes.Values
                ValidateArtifact(descriptor.Artifact, manifest.ArchiveId)
                ValidateArtifact(descriptor.IndexArtifact, manifest.ArchiveId)
                ValidateNavigationIndex(manifest, LoadNode(manifest, descriptor.NodeId), descriptor)
            Next
            For Each descriptor As SemanticArchiveDocumentShardDescriptor In manifest.DocumentShards
                ValidateArtifact(descriptor.Artifact, manifest.ArchiveId)
                For Each document As SemanticArchiveDocumentRecord In LoadDocumentShard(manifest, descriptor).Documents
                    If document.Active AndAlso document.Representation IsNot Nothing Then
                        ValidateDerivedPath(definition, document, document.Representation.TextPath, document.Representation.TextFileHash, document.Representation.TextByteLength)
                        If document.Index IsNot Nothing Then ValidateDerivedPath(definition, document, document.Index.Path, document.Index.FileHash, -1)
                    End If
                Next
            Next
            For Each reference As SemanticArchiveArtifactReference In manifest.TermShards
                ValidateArtifact(reference, manifest.ArchiveId)
            Next
            ' Current routing is mandatory. LoadRoutingGraph validates the graph once;
            ' there is no storage-tree fallback or second graph traversal.
            ValidateArtifact(manifest.RoutingGraphArtifact, manifest.ArchiveId)
            LoadRoutingGraph(manifest)
            Dim visited As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            ValidateAcyclic(manifest, manifest.RootNodeId, New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal), visited)
            If visited.Count <> manifest.Nodes.Count Then Throw New System.IO.InvalidDataException("The audited generation contains unreachable nodes.")
        End Sub

        Private Sub ValidateNavigationIndex(manifest As SemanticArchiveGenerationManifest, node As SemanticArchiveNode, descriptor As SemanticArchiveNodeDescriptor)
            Dim path As System.String = ResolveArtifactPath(descriptor.IndexArtifact)
            Dim index As SharedMethods.SemanticSearchIndexCacheItem = SharedMethods.TryGetSemanticSearchIndexAsync(path).GetAwaiter().GetResult()
            If index Is Nothing OrElse index.IndexDocument Is Nothing OrElse index.IndexDocument.FormatVersion <> 1 OrElse index.IndexDocument.OffsetBase <> "content" OrElse index.IndexDocument.OffsetUnit <> "byte" OrElse Not System.String.Equals(index.IndexDocument.Encoding, "utf-8", System.StringComparison.OrdinalIgnoreCase) Then Throw New System.IO.InvalidDataException("An archive navigation file does not satisfy the existing version-1 indexed-text contract.")
            Dim endOfPrevious As System.Int64 = 0
            Using stream As System.IO.FileStream = SemanticArchivePathGuard.OpenContainedRead(GetArchiveDirectory(manifest.ArchiveId), path)
                For Each card As SemanticArchiveCard In node.Cards
                    If card.LengthBytes <= 0 OrElse card.StartByte < endOfPrevious OrElse card.StartByte > index.ContentByteLength OrElse card.LengthBytes > index.ContentByteLength - card.StartByte Then Throw New System.IO.InvalidDataException("A complete card lies outside its local navigation payload or overlaps another card.")
                    Dim endByte As System.Int64 = card.StartByte + card.LengthBytes
                    For Each boundary As System.Int64 In New System.Int64() {card.StartByte, endByte}
                        If boundary > 0 AndAlso boundary < index.ContentByteLength Then
                            stream.Position = index.ContentStartByte + boundary
                            Dim nextByte As System.Int32 = stream.ReadByte()
                            If nextByte >= &H80 AndAlso nextByte <= &HBF Then Throw New System.IO.InvalidDataException("A navigation card boundary cuts a UTF-8 sequence.")
                        End If
                    Next
                    endOfPrevious = endByte
                Next
            End Using
        End Sub

        Private Shared Function SameArtifact(first As SemanticArchiveArtifactReference, second As SemanticArchiveArtifactReference) As System.Boolean
            Return first IsNot Nothing AndAlso second IsNot Nothing AndAlso first.RelativePath = second.RelativePath AndAlso first.Sha256 = second.Sha256 AndAlso first.Length = second.Length
        End Function

        Private Shared Function IsRetainedShard(previous As SemanticArchiveGenerationManifest, shard As SemanticArchiveDocumentShardDescriptor) As System.Boolean
            If previous Is Nothing Then Return False
            For Each old As SemanticArchiveDocumentShardDescriptor In previous.DocumentShards
                If old.ShardId = shard.ShardId AndAlso SameArtifact(old.Artifact, shard.Artifact) AndAlso old.GenerationId = shard.GenerationId AndAlso old.RecordCount = shard.RecordCount AndAlso old.Inventory IsNot Nothing AndAlso old.Inventory.SameAs(shard.Inventory) AndAlso System.Linq.Enumerable.SequenceEqual(old.DocumentIds, shard.DocumentIds, System.StringComparer.Ordinal) Then Return True
            Next
            Return False
        End Function

        Private Shared Function IsRetainedNode(previous As SemanticArchiveGenerationManifest, node As SemanticArchiveNodeDescriptor) As System.Boolean
            If previous Is Nothing Then Return False
            Dim old As SemanticArchiveNodeDescriptor = Nothing
            Return previous.Nodes.TryGetValue(node.NodeId, old) AndAlso SameArtifact(old.Artifact, node.Artifact) AndAlso SameArtifact(old.IndexArtifact, node.IndexArtifact) AndAlso old.GenerationId = node.GenerationId AndAlso old.ChildCount = node.ChildCount AndAlso System.Linq.Enumerable.SequenceEqual(old.ChildNodeIds, node.ChildNodeIds, System.StringComparer.Ordinal)
        End Function

        Private Shared Sub ValidateAcyclic(manifest As SemanticArchiveGenerationManifest, nodeId As System.String, visiting As System.Collections.Generic.HashSet(Of System.String), visited As System.Collections.Generic.HashSet(Of System.String))
            If visited.Contains(nodeId) Then Return
            If Not visiting.Add(nodeId) Then Throw New System.IO.InvalidDataException("A navigation generation contains a cycle.")
            For Each child As System.String In manifest.Nodes(nodeId).ChildNodeIds
                ValidateAcyclic(manifest, child, visiting, visited)
            Next
            visiting.Remove(nodeId)
            visited.Add(nodeId)
        End Sub

        Public Sub SaveValidity(lease As SemanticArchiveWriterLease, document As SemanticArchiveDocumentRecord, valid As System.Boolean, reason As System.String, ancestorNodeIds As System.Collections.Generic.IEnumerable(Of System.String))
            AssertLease(lease)
            ValidateDocumentRecord(document)
            Dim prior As SemanticArchiveValidityRecord = ReadValidity(lease.ArchiveId, document.DocumentId)
            Dim ancestors As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            If prior IsNot Nothing AndAlso prior.AncestorNodeIds IsNot Nothing Then ancestors.UnionWith(prior.AncestorNodeIds)
            If ancestorNodeIds IsNot Nothing Then ancestors.UnionWith(ancestorNodeIds)
            For Each nodeId As System.String In ancestors
                SemanticArchiveIdentity.ValidateId(nodeId, NameOf(nodeId))
            Next
            Dim record As New SemanticArchiveValidityRecord With {
                .ArchiveId = lease.ArchiveId, .DocumentId = document.DocumentId,
                .SourcePath = document.SourcePath, .SourceHash = document.Fingerprint.Sha256,
                .RepresentationId = If(document.Representation Is Nothing, "", document.Representation.RepresentationId),
                .Valid = valid AndAlso document.Active AndAlso document.Representation IsNot Nothing,
                .Reason = If(reason, ""), .AncestorNodeIds = New System.Collections.Generic.List(Of System.String)(ancestors),
                .FenceToken = lease.FenceToken
            }
            AtomicWriteJson(GetValidityPath(lease.ArchiveId, document.DocumentId), record)
            Dim changedVersion As System.Boolean = prior IsNot Nothing AndAlso (Not System.String.Equals(prior.SourceHash, record.SourceHash, System.StringComparison.OrdinalIgnoreCase) OrElse prior.RepresentationId <> record.RepresentationId)
            If Not record.Valid OrElse changedVersion Then
                For Each nodeId As System.String In ancestors
                    AtomicWriteJson(GetSuppressionPath(lease.ArchiveId, nodeId), New SemanticArchiveSuppressionRecord With {.ArchiveId = lease.ArchiveId, .NodeId = nodeId, .Suppressed = True, .Reason = If(System.String.IsNullOrWhiteSpace(reason), "Source validity or version changed.", reason)})
                Next
            End If
        End Sub

        Public Function ReadValidity(archiveId As System.String, documentId As System.String) As SemanticArchiveValidityRecord
            Dim path As System.String = GetValidityPath(archiveId, documentId)
            If Not ExistsChecked(path) Then Return Nothing
            Dim record As SemanticArchiveValidityRecord = ReadJson(Of SemanticArchiveValidityRecord)(path, 1024 * 1024)
            If record Is Nothing OrElse record.SchemaVersion <> 1 OrElse record.ArchiveId <> archiveId OrElse record.DocumentId <> documentId OrElse record.AncestorNodeIds Is Nothing OrElse record.FenceToken <= 0 Then Throw New System.IO.InvalidDataException("Invalid live source validity record.")
            Return record
        End Function

        Private Function GetValidityPath(archiveId As System.String, documentId As System.String) As System.String
            SemanticArchiveIdentity.ValidateId(documentId, NameOf(documentId))
            Dim bucket As System.String = SemanticArchiveIdentity.HashBytes(System.Text.Encoding.UTF8.GetBytes(documentId)).Substring(0, 2)
            Return System.IO.Path.Combine(GetArchiveDirectory(archiveId), "state", "documents", bucket, documentId & ".json")
        End Function

        Private Function GetSuppressionPath(archiveId As System.String, nodeId As System.String) As System.String
            Return System.IO.Path.Combine(GetArchiveDirectory(archiveId), "state", "nodes", SemanticArchiveIdentity.ValidateId(nodeId, NameOf(nodeId)) & ".json")
        End Function

        Private Function LiveValidityMatches(manifest As SemanticArchiveGenerationManifest, document As SemanticArchiveDocumentRecord) As System.Boolean
            If document Is Nothing OrElse Not document.Active OrElse document.Representation Is Nothing OrElse document.Fingerprint Is Nothing Then Return False
            Dim record As SemanticArchiveValidityRecord = ReadValidity(manifest.ArchiveId, document.DocumentId)
            Return record IsNot Nothing AndAlso record.Valid AndAlso record.RepresentationId = document.Representation.RepresentationId AndAlso System.String.Equals(record.SourceHash, document.Fingerprint.Sha256, System.StringComparison.OrdinalIgnoreCase) AndAlso System.String.Equals(record.SourcePath, document.SourcePath, System.StringComparison.OrdinalIgnoreCase)
        End Function

        Private Sub RefreshDocumentAncestors(manifest As SemanticArchiveGenerationManifest, changedDocuments As System.Collections.Generic.Dictionary(Of System.String, SemanticArchiveDocumentRecord))
            For Each descriptor As SemanticArchiveDocumentShardDescriptor In manifest.DocumentShards
                For Each documentId As System.String In descriptor.DocumentIds
                    If Not changedDocuments.ContainsKey(documentId) Then Continue For
                    Dim record As SemanticArchiveValidityRecord = ReadValidity(manifest.ArchiveId, documentId)
                    If record Is Nothing Then Continue For
                    Dim ancestors As New System.Collections.Generic.HashSet(Of System.String)(record.AncestorNodeIds, System.StringComparer.Ordinal)
                    Dim nodeId As System.String = descriptor.ShardId
                    Dim visited As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
                    While Not System.String.IsNullOrEmpty(nodeId)
                        If Not visited.Add(nodeId) OrElse Not manifest.Nodes.ContainsKey(nodeId) Then Throw New System.IO.InvalidDataException("Invalid document-to-ancestor mapping.")
                        ancestors.Add(nodeId)
                        nodeId = manifest.Nodes(nodeId).ParentNodeId
                    End While
                    record.AncestorNodeIds = New System.Collections.Generic.List(Of System.String)(ancestors)
                    record.FenceToken = manifest.FenceToken
                    AtomicWriteJson(GetValidityPath(manifest.ArchiveId, documentId), record)
                Next
            Next
        End Sub

        Private Sub RefreshNodeValidity(manifest As SemanticArchiveGenerationManifest, changedNodes As System.Collections.Generic.IEnumerable(Of SemanticArchiveNodeDescriptor))
            Dim changed As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            For Each descriptor As SemanticArchiveNodeDescriptor In changedNodes
                changed.Add(descriptor.NodeId)
            Next
            Dim validity As New System.Collections.Generic.Dictionary(Of System.String, System.Boolean)(System.StringComparer.Ordinal)
            For Each descriptor As SemanticArchiveNodeDescriptor In changedNodes
                Dim valid As System.Boolean = HasValidDescendants(manifest, descriptor.NodeId, changed, validity, New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal))
                AtomicWriteJson(GetSuppressionPath(manifest.ArchiveId, descriptor.NodeId), New SemanticArchiveSuppressionRecord With {
                    .ArchiveId = manifest.ArchiveId, .NodeId = descriptor.NodeId,
                    .Suppressed = Not valid,
                    .AllowedArtifactHash = If(valid, descriptor.Artifact.Sha256, ""),
                    .Reason = If(valid, "Validated immutable node replacement.", "One or more descendant sources lack current validity.")
                })
            Next
        End Sub

        Private Function HasValidDescendants(manifest As SemanticArchiveGenerationManifest, nodeId As System.String, changed As System.Collections.Generic.HashSet(Of System.String), validity As System.Collections.Generic.Dictionary(Of System.String, System.Boolean), visited As System.Collections.Generic.HashSet(Of System.String)) As System.Boolean
            Dim known As System.Boolean
            If validity.TryGetValue(nodeId, known) Then Return known
            If Not visited.Add(nodeId) Then Return False
            ' Unchanged immutable subtrees carry a validated live summary state. Do
            ' not replay every document record when one unrelated leaf changes.
            If Not changed.Contains(nodeId) Then
                known = IsNodeUnsuppressed(manifest, nodeId)
                validity(nodeId) = known
                Return known
            End If
            Dim node As SemanticArchiveNode = LoadNode(manifest, nodeId)
            For Each card As SemanticArchiveCard In node.Cards
                If card.Level = "CONTAINER" Then
                    If Not HasValidDescendants(manifest, card.TargetId, changed, validity, visited) Then
                        validity(nodeId) = False
                        Return False
                    End If
                ElseIf Not LiveValidityMatches(manifest, LoadDocument(manifest, card.TargetId)) Then
                    validity(nodeId) = False
                    Return False
                End If
            Next
            validity(nodeId) = True
            Return True
        End Function

        Private Function IsNodeUnsuppressed(manifest As SemanticArchiveGenerationManifest, nodeId As System.String) As System.Boolean
            Dim path As System.String = GetSuppressionPath(manifest.ArchiveId, nodeId)
            If Not ExistsChecked(path) Then Return False
            Dim record As SemanticArchiveSuppressionRecord = ReadJson(Of SemanticArchiveSuppressionRecord)(path, 65536)
            Dim descriptor As SemanticArchiveNodeDescriptor = Nothing
            Return record IsNot Nothing AndAlso record.SchemaVersion = 1 AndAlso record.ArchiveId = manifest.ArchiveId AndAlso record.NodeId = nodeId AndAlso Not record.Suppressed AndAlso manifest.Nodes.TryGetValue(nodeId, descriptor) AndAlso descriptor.Artifact IsNot Nothing AndAlso System.String.Equals(record.AllowedArtifactHash, descriptor.Artifact.Sha256, System.StringComparison.OrdinalIgnoreCase)
        End Function

        Public Function CanReadDocument(access As SemanticArchiveAccessContext, manifest As SemanticArchiveGenerationManifest, document As SemanticArchiveDocumentRecord, Optional verifySourceHash As System.Boolean = False) As System.Boolean
            Try
                If access Is Nothing OrElse manifest Is Nothing OrElse document Is Nothing OrElse Not LiveValidityMatches(manifest, document) Then Return False
                Dim definition As SemanticArchiveDefinition = GetArchive(manifest.ArchiveId)
                If definition Is Nothing OrElse Not definition.Enabled Then Return False
                Dim completeness As System.String = If(document.Representation.Completeness, "unknown").Trim().ToLowerInvariant()
                If completeness <> "complete" AndAlso Not definition.AllowPartialSearch Then Return False
                Dim sourceRoot As System.String = ResolveSourceRoot(definition, document)
                If sourceRoot Is Nothing OrElse Not access.CanReadSource(document.SourcePath) Then Return False
                Using stream As System.IO.FileStream = SemanticArchivePathGuard.OpenContainedRead(sourceRoot, document.SourcePath)
                    Dim info As New System.IO.FileInfo(document.SourcePath)
                    If stream.Length <> document.Fingerprint.Length OrElse info.LastWriteTimeUtc.Ticks <> document.Fingerprint.LastWriteUtcTicks Then Return False
                    If verifySourceHash Then
                        Using hasher As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                            Dim hash As System.String = System.BitConverter.ToString(hasher.ComputeHash(stream)).Replace("-", "").ToLowerInvariant()
                            If Not System.String.Equals(hash, document.Fingerprint.Sha256, System.StringComparison.OrdinalIgnoreCase) Then Return False
                        End Using
                    End If
                End Using
                Return LiveValidityMatches(manifest, document) AndAlso access.CanReadSource(document.SourcePath)
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine("SemanticArchive source authorization failed closed: " & ex.GetType().FullName)
                Return False
            End Try
        End Function

        Public Function CanExposeCard(access As SemanticArchiveAccessContext, manifest As SemanticArchiveGenerationManifest, card As SemanticArchiveCard, Optional maximumDescendantChecks As System.Int32 = 256) As System.Boolean
            Try
                If access Is Nothing OrElse manifest Is Nothing OrElse card Is Nothing OrElse System.String.IsNullOrWhiteSpace(access.PrincipalId) Then Return False
                If card.Level = "DOCUMENT" OrElse card.Level = "SECTION" Then
                    Dim document As SemanticArchiveDocumentRecord = LoadDocument(manifest, card.TargetId)
                    If document Is Nothing OrElse Not CanReadDocument(access, manifest, document) Then Return False
                    Return (System.String.IsNullOrWhiteSpace(card.SourceVersion) OrElse System.String.Equals(card.SourceVersion, document.Fingerprint.Sha256, System.StringComparison.OrdinalIgnoreCase)) AndAlso (System.String.IsNullOrWhiteSpace(card.RepresentationId) OrElse card.RepresentationId = document.Representation.RepresentationId)
                End If
                If card.Level <> "CONTAINER" OrElse Not manifest.Nodes.ContainsKey(card.TargetId) Then Return False
                If Not IsNodeUnsuppressed(manifest, card.TargetId) Then Return False
                ' A container summary can contain information from any descendant.
                ' Current requester access must be established for every one.
                Dim remaining As System.Int32 = System.Math.Min(256, System.Math.Max(0, maximumDescendantChecks))
                Return CanReadDescendants(access, manifest, card.TargetId, New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal), remaining)
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine("SemanticArchive card authorization failed closed: " & ex.GetType().FullName)
                Return False
            End Try
        End Function

        Private Function CanReadDescendants(access As SemanticArchiveAccessContext, manifest As SemanticArchiveGenerationManifest, nodeId As System.String, visited As System.Collections.Generic.HashSet(Of System.String), ByRef remaining As System.Int32) As System.Boolean
            If Not visited.Add(nodeId) OrElse Not IsNodeUnsuppressed(manifest, nodeId) Then Return False
            Dim node As SemanticArchiveNode = LoadNode(manifest, nodeId)
            For Each child As SemanticArchiveCard In node.Cards
                If child.Level = "CONTAINER" Then
                    If Not CanReadDescendants(access, manifest, child.TargetId, visited, remaining) Then Return False
                Else
                    If remaining <= 0 Then Return False
                    remaining -= 1
                    If Not CanReadDocument(access, manifest, LoadDocument(manifest, child.TargetId)) Then Return False
                End If
            Next
            Return True
        End Function

        Private Function ResolveSourceRoot(definition As SemanticArchiveDefinition, document As SemanticArchiveDocumentRecord) As System.String
            If document.BindingIds Is Nothing Then Return Nothing
            For Each binding As SemanticArchiveSourceBinding In definition.Roots
                If document.BindingIds.Contains(binding.BindingId) AndAlso SemanticArchivePathGuard.IsContainedPath(binding.RootPath, document.SourcePath) Then
                    Dim relative As System.String = document.SourcePath.Substring(binding.RootPath.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar).Length).TrimStart(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar)
                    If Not binding.Recursive AndAlso relative.IndexOfAny(New System.Char() {System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar}) >= 0 Then Continue For
                    If IsSourceExcluded(binding, document.SourcePath) OrElse Not IsSupportedSource(binding, document.SourcePath) Then Continue For
                    SemanticArchivePathGuard.ValidateContainedPath(binding.RootPath, document.SourcePath, True)
                    Return binding.RootPath
                End If
            Next
            Return Nothing
        End Function

        Public Shared Function IsExcluded(binding As SemanticArchiveSourceBinding, path As System.String, relativePath As System.String) As System.Boolean
            Return IsSourceExcluded(binding, path)
        End Function

        Public Shared Function IsSourceExcluded(binding As SemanticArchiveSourceBinding, path As System.String) As System.Boolean
            If binding Is Nothing Then Throw New System.ArgumentNullException(NameOf(binding))
            ' This provider reserves its own metadata namespace even when its catalog
            ' is not configured in this host. File-type filtering must never admit it.
            If KnowledgeStoreCatalog.IsGeneratedMetadataPath(path) Then Return True
            If GeneratedOutputRegistry.IsGeneratedPathVerified(path) Then Return True
            Dim root As System.String = SemanticArchivePathGuard.CanonicalPath(binding.RootPath)
            Dim full As System.String = SemanticArchivePathGuard.CanonicalPath(path)
            If Not SemanticArchivePathGuard.IsContainedPath(root, full) Then Return True
            Dim relative As System.String = full.Substring(root.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar).Length).TrimStart(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar).Replace("\", "/")
            Dim comparison As System.StringComparison = If(System.Environment.OSVersion.Platform = System.PlatformID.Win32NT, System.StringComparison.OrdinalIgnoreCase, System.StringComparison.Ordinal)
            For Each rawPattern As System.String In binding.Exclusions
                If System.String.IsNullOrWhiteSpace(rawPattern) Then Continue For
                Dim pattern As System.String = SharedMethods.ExpandEnvironmentVariables(rawPattern.Trim())
                If System.String.IsNullOrWhiteSpace(pattern) Then Throw New System.ArgumentException("An archive exclusion could not be expanded.")
                If System.Text.RegularExpressions.Regex.IsMatch(pattern, "%[^%]+%", System.Text.RegularExpressions.RegexOptions.CultureInvariant, System.TimeSpan.FromMilliseconds(100)) Then
                    Throw New System.ArgumentException("An archive exclusion contains an unresolved environment variable or placeholder.")
                End If
                If System.IO.Path.IsPathRooted(pattern) Then
                    If SemanticArchivePathGuard.IsContainedPath(pattern, full) Then Return True
                    Continue For
                End If
                pattern = pattern.Replace("\", "/").Trim("/"c)
                If pattern.IndexOfAny(New System.Char() {"*"c, "?"c}) >= 0 Then
                    Dim target As System.String = If(pattern.Contains("/"), relative, System.IO.Path.GetFileName(full))
                    Dim expression As System.String = "\A" & System.Text.RegularExpressions.Regex.Escape(pattern).Replace("\*", ".*").Replace("\?", ".") & "\z"
                    Dim flags As System.Text.RegularExpressions.RegexOptions = System.Text.RegularExpressions.RegexOptions.CultureInvariant
                    If comparison = System.StringComparison.OrdinalIgnoreCase Then flags = flags Or System.Text.RegularExpressions.RegexOptions.IgnoreCase
                    If System.Text.RegularExpressions.Regex.IsMatch(target, expression, flags, System.TimeSpan.FromSeconds(1)) Then Return True
                ElseIf System.String.Equals(relative, pattern, comparison) OrElse relative.StartsWith(pattern & "/", comparison) Then
                    Return True
                End If
            Next
            Return False
        End Function

        Public Shared Function NormalizeSourceExtension(value As System.String) As System.String
            Dim extension As System.String = If(value, "").Trim().TrimStart("*"c)
            If extension.Length = 0 Then Return ""
            If Not extension.StartsWith(".", System.StringComparison.Ordinal) Then extension = "." & extension
            If extension.IndexOfAny(New System.Char() {"/"c, "\"c, "*"c, "?"c, ":"c}) >= 0 Then Throw New System.ArgumentException("Only source filename extensions are supported in SupportedExtensions.", NameOf(value))
            Return extension.ToLowerInvariant()
        End Function

        ''' <summary>A blank enabled list means the central office/image defaults, not every type.
        ''' Disabling the filter retains the list and still requires an installed converter.</summary>
        Public Shared Function GetSourceFilterExtensions(binding As SemanticArchiveSourceBinding) As System.Collections.Generic.IEnumerable(Of System.String)
            If binding Is Nothing Then Throw New System.ArgumentNullException(NameOf(binding))
            If binding.SupportedExtensions IsNot Nothing AndAlso binding.SupportedExtensions.Count > 0 Then Return binding.SupportedExtensions
            Return SharedMethods.DEFAULT_SEMANTICARCHIVE_SUPPORTED_EXTENSIONS.Split(";"c)
        End Function

        Public Shared Function IsSupportedSource(binding As SemanticArchiveSourceBinding, path As System.String) As System.Boolean
            If binding Is Nothing Then Throw New System.ArgumentNullException(NameOf(binding))
            Dim extension As System.String = System.IO.Path.GetExtension(path)
            Dim converterSupports As System.Boolean = False
            For Each supported As System.String In Global.SharedLibrary.Agents.TextExportService.GetSupportedExtensions()
                If System.String.Equals(extension, supported, System.StringComparison.OrdinalIgnoreCase) Then
                    converterSupports = True
                    Exit For
                End If
            Next
            If Not converterSupports Then Return False
            If Not binding.FileTypeFilterEnabled Then Return True
            For Each configured As System.String In GetSourceFilterExtensions(binding)
                If System.String.Equals(extension, NormalizeSourceExtension(configured), System.StringComparison.OrdinalIgnoreCase) Then Return True
            Next
            Return False
        End Function

        Private Function ValidateDerivedPath(definition As SemanticArchiveDefinition, document As SemanticArchiveDocumentRecord, path As System.String, expectedHash As System.String, expectedLength As System.Int64) As System.String
            If Not IsHash(expectedHash) OrElse System.String.IsNullOrWhiteSpace(path) Then Throw New System.IO.InvalidDataException("Incomplete derived artifact provenance.")
            Dim derivedRoot As System.String = Nothing
            For Each binding As SemanticArchiveSourceBinding In definition.Roots
                If document.BindingIds.Contains(binding.BindingId) Then
                    For Each registeredRoot As System.String In GetPrivateDerivedRoots(binding)
                        Dim candidate As System.String = System.IO.Path.Combine(registeredRoot, "versions")
                        If SemanticArchivePathGuard.IsContainedPath(candidate, path) Then
                            derivedRoot = candidate
                            Exit For
                        End If
                    Next
                    If derivedRoot IsNot Nothing Then Exit For
                End If
            Next
            If derivedRoot Is Nothing Then Throw New System.UnauthorizedAccessException("The derived artifact is outside registered immutable version locations.")
            Dim full As System.String = SemanticArchivePathGuard.ValidateContainedPath(derivedRoot, path, True)
            RequirePrivateArtifact(full)
            Using stream As System.IO.FileStream = SemanticArchivePathGuard.OpenContainedRead(derivedRoot, full)
                If expectedLength >= 0 AndAlso stream.Length <> expectedLength Then Throw New System.IO.InvalidDataException("Derived artifact length mismatch.")
                Using hasher As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                    Dim hash As System.String = System.BitConverter.ToString(hasher.ComputeHash(stream)).Replace("-", "").ToLowerInvariant()
                    If Not System.String.Equals(hash, expectedHash, System.StringComparison.OrdinalIgnoreCase) Then Throw New System.IO.InvalidDataException("Derived artifact hash mismatch.")
                End Using
            End Using
            Return full
        End Function

        Public Function ValidateTextPath(manifest As SemanticArchiveGenerationManifest, document As SemanticArchiveDocumentRecord, access As SemanticArchiveAccessContext) As System.String
            If Not CanReadDocument(access, manifest, document, True) Then Throw New System.UnauthorizedAccessException("The source is outside current permitted archive scope or no longer valid.")
            Dim definition As SemanticArchiveDefinition = GetArchive(manifest.ArchiveId)
            Dim full As System.String = ValidateDerivedPath(definition, document, document.Representation.TextPath, document.Representation.TextFileHash, document.Representation.TextByteLength)
            If Not CanReadDocument(access, manifest, document) Then Throw New System.UnauthorizedAccessException("The source was revoked during artifact validation.")
            Return full
        End Function

        Public Function ValidateIndexPath(manifest As SemanticArchiveGenerationManifest, document As SemanticArchiveDocumentRecord, access As SemanticArchiveAccessContext) As System.String
            If document Is Nothing OrElse document.Index Is Nothing Then Throw New System.IO.InvalidDataException("The document has no semantic index.")
            If Not CanReadDocument(access, manifest, document, True) Then Throw New System.UnauthorizedAccessException("The source is outside current permitted archive scope or no longer valid.")
            If document.Index.RepresentationId <> document.Representation.RepresentationId OrElse document.Index.FormatVersion <> 1 Then Throw New System.IO.InvalidDataException("The semantic index has a stale representation binding.")
            Dim full As System.String = ValidateDerivedPath(GetArchive(manifest.ArchiveId), document, document.Index.Path, document.Index.FileHash, -1)
            If Not CanReadDocument(access, manifest, document) Then Throw New System.UnauthorizedAccessException("The source was revoked during index validation.")
            Return full
        End Function

        Public Function ReadTextRange(manifest As SemanticArchiveGenerationManifest, document As SemanticArchiveDocumentRecord, access As SemanticArchiveAccessContext, startByte As System.Int64, maxBytes As System.Int32) As System.Byte()
            If startByte < 0 OrElse maxBytes < 1 OrElse maxBytes > 4 * 1024 * 1024 Then Throw New System.ArgumentOutOfRangeException(NameOf(maxBytes), "A positive evidence range of at most 4 MiB is required.")
            Dim path As System.String = ValidateTextPath(manifest, document, access)
            Dim result As System.Byte()
            Using stream As System.IO.FileStream = SemanticArchivePathGuard.OpenContainedRead(System.IO.Path.GetDirectoryName(path), path)
                If startByte > stream.Length Then Throw New System.ArgumentOutOfRangeException(NameOf(startByte))
                stream.Position = startByte
                Dim length As System.Int32 = CInt(System.Math.Min(CLng(maxBytes), stream.Length - startByte))
                If length = 0 Then Return New System.Byte() {}
                result = New System.Byte(length - 1) {}
                Dim offset As System.Int32 = 0
                While offset < length
                    Dim count As System.Int32 = stream.Read(result, offset, length - offset)
                    If count = 0 Then Throw New System.IO.EndOfStreamException("The derived evidence was truncated.")
                    offset += count
                End While
            End Using
            If Not CanReadDocument(access, manifest, document) Then Throw New System.UnauthorizedAccessException("The source was revoked while evidence was read.")
            Return result
        End Function

        Public Function ReadTextBytes(manifest As SemanticArchiveGenerationManifest, document As SemanticArchiveDocumentRecord, access As SemanticArchiveAccessContext) As System.Byte()
            If document Is Nothing OrElse document.Representation Is Nothing OrElse document.Representation.TextByteLength > 4L * 1024L * 1024L Then Throw New System.InvalidOperationException("Use a bounded text range or semantic sections for this document.")
            Return ReadTextRange(manifest, document, access, 0, System.Math.Max(1, CInt(document.Representation.TextByteLength)))
        End Function
    End Class
End Namespace
