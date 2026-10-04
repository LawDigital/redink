' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.

Option Strict On
Option Explicit On
Option Infer On

Namespace SharedLibrary

    ''' <summary>
    ''' Persistent range hierarchy. A mutation reads only its document shard and ancestor
    ''' path; overflow splits that existing node. Unrelated node identities, ranges, bytes,
    ''' summaries and direct immutable references remain unchanged.
    ''' </summary>
    Friend NotInheritable Class SemanticArchiveHierarchy
        Private ReadOnly _store As SemanticArchiveStore
        Private ReadOnly _archive As SemanticArchiveDefinition
        Private ReadOnly _context As SharedContext.ISharedContext
        Private ReadOnly _previous As SemanticArchiveGenerationManifest
        Private ReadOnly _manifest As SemanticArchiveGenerationManifest
        Private ReadOnly _nodes As New System.Collections.Generic.Dictionary(Of String, SemanticArchiveNode)(System.StringComparer.Ordinal)
        Private ReadOnly _shards As New System.Collections.Generic.Dictionary(Of String, SemanticArchiveDocumentShard)(System.StringComparer.Ordinal)
        Private ReadOnly _documentLeaves As New System.Collections.Generic.Dictionary(Of String, String)(System.StringComparer.Ordinal)
        Private ReadOnly _dirty As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.Ordinal)
        Private ReadOnly _diagnostics As System.Collections.Generic.List(Of String)

        Public Sub New(store As SemanticArchiveStore, archive As SemanticArchiveDefinition, context As SharedContext.ISharedContext,
                       previous As SemanticArchiveGenerationManifest, manifest As SemanticArchiveGenerationManifest,
                       diagnostics As System.Collections.Generic.List(Of String))
            _store = store
            _archive = archive
            _context = context
            _previous = previous
            _manifest = manifest
            _diagnostics = diagnostics
            _manifest.MaxChildrenPerNode = archive.MaxChildrenPerNode
            _manifest.MaxRoutingCharacters = archive.MaxRoutingCharacters
            If archive.MaxChildrenPerNode < 2 OrElse archive.MaxChildrenPerNode > 256 Then Throw New System.IO.InvalidDataException("Archive child-count limit must be between 2 and 256.")
            If archive.MaxRoutingCharacters < 4096 OrElse archive.MaxRoutingCharacters > 120000 Then Throw New System.IO.InvalidDataException("Archive routing budget must be between 4,096 and 120,000 characters.")
            If previous IsNot Nothing Then
                _manifest.RootNodeId = previous.RootNodeId
                _manifest.Nodes = SemanticArchiveMetadata.Clone(previous.Nodes)
                _manifest.DocumentShards = SemanticArchiveMetadata.Clone(previous.DocumentShards)
                _manifest.DocumentCount = previous.DocumentCount
                _manifest.FailureCount = previous.FailureCount
                For Each shard As SemanticArchiveDocumentShardDescriptor In _manifest.DocumentShards
                    For Each documentId As String In shard.DocumentIds
                        If _documentLeaves.ContainsKey(documentId) Then Throw New System.IO.InvalidDataException("A document belongs to more than one archive record shard.")
                        _documentLeaves.Add(documentId, shard.ShardId)
                    Next
                Next
                PrunePreviouslyRemovedRecords()
                If previous.MaxChildrenPerNode <> archive.MaxChildrenPerNode OrElse previous.MaxRoutingCharacters <> archive.MaxRoutingCharacters Then
                    For Each nodeId As String In New System.Collections.Generic.List(Of String)(_manifest.Nodes.Keys)
                        Dim node As SemanticArchiveNode = GetNode(nodeId)
                        node.GenerationId = _manifest.GenerationId
                        _dirty.Add(nodeId)
                        If node.IsLeaf Then RebuildLeafCards(nodeId)
                    Next
                    _diagnostics.Add("routing_policy_changed: navigation is rebuilt from persisted metadata; extraction is reused.")
                End If
            Else
                Dim root As SemanticArchiveNode = NewNode(True, 0, "")
                _manifest.RootNodeId = root.NodeId
                _shards(root.NodeId) = New SemanticArchiveDocumentShard() With {.ShardId = root.NodeId}
            End If
        End Sub

        Private Sub PrunePreviouslyRemovedRecords()
            If _previous Is Nothing OrElse _previous.Inventory Is Nothing OrElse _previous.Inventory.RemovedSources <= 0 Then Return
            Dim retiredIds As New System.Collections.Generic.List(Of System.String)()
            For Each descriptor As SemanticArchiveDocumentShardDescriptor In _manifest.DocumentShards
                Dim shard As SemanticArchiveDocumentShard = GetShard(descriptor.ShardId)
                For Each document As SemanticArchiveDocumentRecord In shard.Documents
                    If document IsNot Nothing AndAlso System.String.Equals(document.ProcessingStatus, "removed", System.StringComparison.Ordinal) Then retiredIds.Add(document.DocumentId)
                Next
            Next
            For Each documentId As System.String In retiredIds
                Remove(documentId)
            Next
            If retiredIds.Count > 0 Then
                _diagnostics.Add("retired_records_pruned: " & retiredIds.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) & " obsolete removed record(s) were dropped from the active generation.")
            End If
        End Sub

        Public ReadOnly Property HasChanges As Boolean
            Get
                Return _dirty.Count > 0
            End Get
        End Property

        Public Function Ancestors(documentId As String) As System.Collections.Generic.List(Of String)
            Dim result As New System.Collections.Generic.List(Of String)()
            Dim nodeId As String = Nothing
            If Not _documentLeaves.TryGetValue(documentId, nodeId) Then Return result
            Dim seen As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.Ordinal)
            While Not System.String.IsNullOrEmpty(nodeId)
                If Not seen.Add(nodeId) Then Throw New System.IO.InvalidDataException("The archive hierarchy contains a parent cycle.")
                result.Add(nodeId)
                Dim node As SemanticArchiveNode = Nothing
                If _nodes.TryGetValue(nodeId, node) Then
                    nodeId = node.ParentNodeId
                Else
                    nodeId = _manifest.Nodes(nodeId).ParentNodeId
                End If
            End While
            Return result
        End Function

        Public Sub Apply(document As SemanticArchiveDocumentRecord)
            If document Is Nothing Then Throw New System.ArgumentNullException(NameOf(document))
            If System.String.Equals(document.ProcessingStatus, "removed", System.StringComparison.Ordinal) Then
                Remove(document.DocumentId)
                Return
            End If
            Dim existingLeaf As String = Nothing
            Dim oldRecord As SemanticArchiveDocumentRecord = Nothing
            If _documentLeaves.TryGetValue(document.DocumentId, existingLeaf) Then
                Dim oldShard As SemanticArchiveDocumentShard = GetShard(existingLeaf)
                oldRecord = oldShard.Documents.Find(Function(value As SemanticArchiveDocumentRecord) System.String.Equals(value.DocumentId, document.DocumentId, System.StringComparison.Ordinal))
                oldShard.Documents.RemoveAll(Function(value As SemanticArchiveDocumentRecord) System.String.Equals(value.DocumentId, document.DocumentId, System.StringComparison.Ordinal))
                RebuildLeafCards(existingLeaf)
                MarkDirtyPath(existingLeaf)
            End If
            If oldRecord IsNot Nothing Then
                If IsSearchable(oldRecord) Then _manifest.DocumentCount -= 1
                If IsFailure(oldRecord) Then _manifest.FailureCount -= 1
            End If
            Dim targetLeaf As String = existingLeaf
            If System.String.IsNullOrEmpty(targetLeaf) OrElse
                (oldRecord IsNot Nothing AndAlso Not System.String.Equals(oldRecord.PartitionKey, document.PartitionKey, System.StringComparison.Ordinal)) Then
                targetLeaf = FindLeaf(document.PartitionKey)
            End If
            Dim shard As SemanticArchiveDocumentShard = GetShard(targetLeaf)
            shard.Documents.Add(document)
            _documentLeaves(document.DocumentId) = targetLeaf
            RebuildLeafCards(targetLeaf)
            MarkDirtyPath(targetLeaf)
            If IsSearchable(document) Then _manifest.DocumentCount += 1
            If IsFailure(document) Then _manifest.FailureCount += 1
        End Sub


        Public Sub Remove(documentId As System.String)
            If System.String.IsNullOrWhiteSpace(documentId) Then Return
            Dim existingLeaf As System.String = Nothing
            If Not _documentLeaves.TryGetValue(documentId, existingLeaf) Then Return
            Dim shard As SemanticArchiveDocumentShard = GetShard(existingLeaf)
            Dim oldRecord As SemanticArchiveDocumentRecord = shard.Documents.Find(Function(value As SemanticArchiveDocumentRecord) System.String.Equals(value.DocumentId, documentId, System.StringComparison.Ordinal))
            If oldRecord Is Nothing Then
                _documentLeaves.Remove(documentId)
                Return
            End If
            shard.Documents.RemoveAll(Function(value As SemanticArchiveDocumentRecord) System.String.Equals(value.DocumentId, documentId, System.StringComparison.Ordinal))
            _documentLeaves.Remove(documentId)
            If IsSearchable(oldRecord) Then _manifest.DocumentCount -= 1
            If IsFailure(oldRecord) Then _manifest.FailureCount -= 1
            RebuildLeafCards(existingLeaf)
            MarkDirtyPath(existingLeaf)
        End Sub

        Private Shared Function IsSearchable(document As SemanticArchiveDocumentRecord) As Boolean
            Return document.Active AndAlso document.Card IsNot Nothing AndAlso document.Representation IsNot Nothing AndAlso
                document.Representation.Completeness <> "empty"
        End Function

        Private Shared Function IsFailure(document As SemanticArchiveDocumentRecord) As Boolean
            Return document.ProcessingStatus = "failed" OrElse document.ProcessingStatus = "unavailable"
        End Function

        Private Function FindLeaf(key As String) As String
            Dim nodeId As String = _manifest.RootNodeId
            Dim visited As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.Ordinal)
            Do
                If Not visited.Add(nodeId) Then Throw New System.IO.InvalidDataException("The archive hierarchy contains a child cycle.")
                Dim node As SemanticArchiveNode = GetNode(nodeId)
                If node.IsLeaf Then Return nodeId
                If node.Cards.Count = 0 Then Throw New System.IO.InvalidDataException("A non-leaf archive node has no child references.")
                Dim nextId As String = node.Cards(node.Cards.Count - 1).TargetId
                For Each card As SemanticArchiveCard In node.Cards
                    Dim maximum As String = GetNodeMaxKey(card.TargetId)
                    If System.String.IsNullOrEmpty(maximum) OrElse System.StringComparer.Ordinal.Compare(key, maximum) <= 0 Then
                        nextId = card.TargetId
                        Exit For
                    End If
                Next
                nodeId = nextId
            Loop
        End Function

        Private Function GetNodeMaxKey(nodeId As String) As String
            Dim node As SemanticArchiveNode = Nothing
            If _nodes.TryGetValue(nodeId, node) Then Return node.MaxKey
            Return _manifest.Nodes(nodeId).MaxKey
        End Function

        Private Function GetNode(nodeId As String) As SemanticArchiveNode
            Dim node As SemanticArchiveNode = Nothing
            If _nodes.TryGetValue(nodeId, node) Then Return node
            If _previous Is Nothing OrElse Not _manifest.Nodes.ContainsKey(nodeId) Then Throw New System.IO.InvalidDataException("An archive node reference is missing.")
            node = _store.LoadNode(_previous, nodeId)
            ' Descriptor parent identity is authoritative; an immutable child can acquire
            ' a new parent without rewriting its own payload or metadata artifact.
            node.ParentNodeId = _manifest.Nodes(nodeId).ParentNodeId
            _nodes.Add(nodeId, node)
            Return node
        End Function

        Private Function GetShard(nodeId As String) As SemanticArchiveDocumentShard
            Dim shard As SemanticArchiveDocumentShard = Nothing
            If _shards.TryGetValue(nodeId, shard) Then Return shard
            shard = _store.LoadDocumentShard(_previous, nodeId)
            _shards.Add(nodeId, shard)
            Return shard
        End Function

        Private Function NewNode(isLeaf As Boolean, level As Integer, parentNodeId As String) As SemanticArchiveNode
            Dim node As New SemanticArchiveNode() With {
                .NodeId = SemanticArchiveIdentity.NewId(), .GenerationId = _manifest.GenerationId,
                .ParentNodeId = parentNodeId, .IsLeaf = isLeaf, .Level = level, .IsPermissionNeutral = False
            }
            _nodes.Add(node.NodeId, node)
            _dirty.Add(node.NodeId)
            Return node
        End Function

        Private Sub MarkDirtyPath(nodeId As String)
            Dim seen As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.Ordinal)
            While Not System.String.IsNullOrEmpty(nodeId)
                If Not seen.Add(nodeId) Then Throw New System.IO.InvalidDataException("The archive hierarchy contains a parent cycle.")
                _dirty.Add(nodeId)
                Dim node As SemanticArchiveNode = GetNode(nodeId)
                node.GenerationId = _manifest.GenerationId
                nodeId = node.ParentNodeId
            End While
        End Sub

        Private Sub RebuildLeafCards(nodeId As String)
            Dim shard As SemanticArchiveDocumentShard = GetShard(nodeId)
            shard.Documents.Sort(Function(left As SemanticArchiveDocumentRecord, right As SemanticArchiveDocumentRecord) System.StringComparer.Ordinal.Compare(left.PartitionKey, right.PartitionKey))
            Dim node As SemanticArchiveNode = GetNode(nodeId)
            node.Cards.Clear()
            For Each document As SemanticArchiveDocumentRecord In shard.Documents
                If Not IsSearchable(document) Then Continue For
                Dim card As SemanticArchiveCard = SemanticArchiveMetadata.BoundRoutingCard(document.Card, System.Math.Min(8000, _archive.MaxRoutingCharacters - 1024), _diagnostics)
                card.PartitionKey = document.PartitionKey
                node.Cards.Add(card)
            Next
            If shard.Documents.Count > 0 Then
                shard.MinKey = shard.Documents(0).PartitionKey
                shard.MaxKey = shard.Documents(shard.Documents.Count - 1).PartitionKey
                node.MinKey = shard.MinKey
                node.MaxKey = shard.MaxKey
            Else
                shard.MinKey = ""
                shard.MaxKey = ""
                node.MinKey = ""
                node.MaxKey = ""
            End If
        End Sub

        ''' <summary>
        ''' Finalizes dirty ancestors once per batch, splitting locally after actual complete
        ''' routing metadata is known. Index files are emitted only after the root is final.
        ''' </summary>
        Public Async Function WriteChangesAsync(cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            If _dirty.Count = 0 Then Return
            Dim roots As System.Collections.Generic.List(Of SemanticArchiveCard) = Await FinalizeNodeAsync(_manifest.RootNodeId, cancellationToken).ConfigureAwait(False)
            While roots.Count > 1
                Dim oldLevel As Integer = GetNode(roots(0).TargetId).Level
                Dim newRoot As SemanticArchiveNode = NewNode(False, oldLevel + 1, "")
                newRoot.Cards = roots
                SetChildParents(newRoot)
                _manifest.RootNodeId = newRoot.NodeId
                roots = Await FinalizeLocalNodeAsync(newRoot, cancellationToken).ConfigureAwait(False)
            End While
            _manifest.RootNodeId = roots(0).TargetId
            Dim finalRoot As SemanticArchiveNode = GetNode(_manifest.RootNodeId)
            finalRoot.ParentNodeId = ""
            For Each nodeId As String In _dirty
                cancellationToken.ThrowIfCancellationRequested()
                Dim node As SemanticArchiveNode = _nodes(nodeId)
                Dim generationDirectory As String = _store.GetGenerationDirectory(_manifest.ArchiveId, _manifest.GenerationId)
                Dim path As String = If(System.String.Equals(nodeId, _manifest.RootNodeId, System.StringComparison.Ordinal),
                    System.IO.Path.Combine(generationDirectory, "root.indexed.txt"),
                    System.IO.Path.Combine(generationDirectory, "nodes", nodeId & ".indexed.txt"))
                Await SharedMethods.WriteSemanticArchiveNavigationIndexAsync(node, path, cancellationToken).ConfigureAwait(False)
                Dim descriptor As SemanticArchiveNodeDescriptor = _store.WriteNode(_manifest, node)
                _manifest.Nodes(nodeId) = descriptor
                If node.IsLeaf Then
                    Dim shard As SemanticArchiveDocumentShard = GetShard(nodeId)
                    Dim shardDescriptor As SemanticArchiveDocumentShardDescriptor = _store.WriteDocumentShard(_manifest, shard)
                    _manifest.DocumentShards.RemoveAll(Function(value As SemanticArchiveDocumentShardDescriptor) System.String.Equals(value.ShardId, nodeId, System.StringComparison.Ordinal))
                    _manifest.DocumentShards.Add(shardDescriptor)
                End If
            Next
            ' Parent changes on immutable, otherwise unchanged child descriptors are only
            ' manifest changes. Queries follow direct artifact references, never delta chains.
            For Each node As SemanticArchiveNode In _nodes.Values
                If _manifest.Nodes.ContainsKey(node.NodeId) Then _manifest.Nodes(node.NodeId).ParentNodeId = node.ParentNodeId
            Next
            PruneUnreachableDescriptors()
            _manifest.Diagnostics.AddRange(_diagnostics)
        End Function

        Private Sub PruneUnreachableDescriptors()
            Dim reachable As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            Dim pending As New System.Collections.Generic.Stack(Of System.String)()
            pending.Push(_manifest.RootNodeId)
            While pending.Count > 0
                Dim nodeId As System.String = pending.Pop()
                If Not reachable.Add(nodeId) Then Continue While
                Dim descriptor As SemanticArchiveNodeDescriptor = Nothing
                If Not _manifest.Nodes.TryGetValue(nodeId, descriptor) OrElse descriptor Is Nothing OrElse descriptor.ChildNodeIds Is Nothing Then Continue While
                For Each childId As System.String In descriptor.ChildNodeIds
                    pending.Push(childId)
                Next
            End While
            For Each nodeId As System.String In New System.Collections.Generic.List(Of System.String)(_manifest.Nodes.Keys)
                If Not reachable.Contains(nodeId) Then _manifest.Nodes.Remove(nodeId)
            Next
            _manifest.DocumentShards.RemoveAll(Function(value As SemanticArchiveDocumentShardDescriptor) value Is Nothing OrElse Not reachable.Contains(value.ShardId))
        End Sub

        Private Async Function FinalizeNodeAsync(nodeId As String, cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of System.Collections.Generic.List(Of SemanticArchiveCard))
            Dim node As SemanticArchiveNode = GetNode(nodeId)
            If node.IsLeaf AndAlso GetShard(nodeId).Documents.Count = 0 AndAlso Not System.String.Equals(nodeId, _manifest.RootNodeId, System.StringComparison.Ordinal) Then
                Return New System.Collections.Generic.List(Of SemanticArchiveCard)()
            End If
            If Not node.IsLeaf Then
                Dim children As New System.Collections.Generic.List(Of SemanticArchiveCard)()
                For Each card As SemanticArchiveCard In node.Cards
                    If _dirty.Contains(card.TargetId) Then
                        children.AddRange(Await FinalizeNodeAsync(card.TargetId, cancellationToken).ConfigureAwait(False))
                    Else
                        children.Add(card)
                    End If
                Next
                node.Cards = children
                SetChildParents(node)
                If node.Cards.Count = 0 AndAlso Not System.String.Equals(nodeId, _manifest.RootNodeId, System.StringComparison.Ordinal) Then
                    Return New System.Collections.Generic.List(Of SemanticArchiveCard)()
                End If
            End If
            Return Await FinalizeLocalNodeAsync(node, cancellationToken).ConfigureAwait(False)
        End Function

        Private Function FinalizeLocalNodeAsync(node As SemanticArchiveNode, cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of System.Collections.Generic.List(Of SemanticArchiveCard))
            cancellationToken.ThrowIfCancellationRequested()
            Dim parts As System.Collections.Generic.List(Of SemanticArchiveNode) = SplitLocal(node)
            Dim results As New System.Collections.Generic.List(Of SemanticArchiveCard)()
            For Each part As SemanticArchiveNode In parts
                cancellationToken.ThrowIfCancellationRequested()
                part.GenerationId = _manifest.GenerationId
                UpdateRange(part)
                If Not part.IsLeaf Then SetChildParents(part)
                Dim summary As SemanticArchiveCard = SemanticArchiveMetadata.ComposeContainerCard(part)
                results.Add(SemanticArchiveMetadata.BoundRoutingCard(summary, System.Math.Min(8000, _archive.MaxRoutingCharacters - 1024), _diagnostics))
            Next
            Return System.Threading.Tasks.Task.FromResult(results)
        End Function

        Private Function SplitLocal(node As SemanticArchiveNode) As System.Collections.Generic.List(Of SemanticArchiveNode)
            Dim count As Integer = If(node.IsLeaf, GetShard(node.NodeId).Documents.Count, node.Cards.Count)
            If count <= _archive.MaxChildrenPerNode AndAlso RoutingSize(node) <= _archive.MaxRoutingCharacters Then
                Return New System.Collections.Generic.List(Of SemanticArchiveNode) From {node}
            End If
            If count <= 1 Then Throw New System.IO.InvalidDataException("oversized_card: a single routing record cannot fit a bounded node.")
            Dim midpoint As Integer = count \ 2
            Dim sibling As SemanticArchiveNode = NewNode(node.IsLeaf, node.Level, node.ParentNodeId)
            If node.IsLeaf Then
                Dim shard As SemanticArchiveDocumentShard = GetShard(node.NodeId)
                Dim right As New SemanticArchiveDocumentShard() With {.ShardId = sibling.NodeId}
                right.Documents.AddRange(shard.Documents.GetRange(midpoint, count - midpoint))
                shard.Documents.RemoveRange(midpoint, count - midpoint)
                _shards.Add(sibling.NodeId, right)
                For Each document As SemanticArchiveDocumentRecord In right.Documents
                    _documentLeaves(document.DocumentId) = sibling.NodeId
                Next
                RebuildLeafCards(node.NodeId)
                RebuildLeafCards(sibling.NodeId)
            Else
                sibling.Cards.AddRange(node.Cards.GetRange(midpoint, count - midpoint))
                node.Cards.RemoveRange(midpoint, count - midpoint)
                SetChildParents(node)
                SetChildParents(sibling)
            End If
            _dirty.Add(node.NodeId)
            _diagnostics.Add("local_split: " & node.NodeId & " -> " & node.NodeId & ", " & sibling.NodeId & ".")
            Dim result As System.Collections.Generic.List(Of SemanticArchiveNode) = SplitLocal(node)
            result.AddRange(SplitLocal(sibling))
            Return result
        End Function

        Private Shared Function RoutingSize(node As SemanticArchiveNode) As Integer
            Dim length As Integer = 0
            For Each card As SemanticArchiveCard In node.Cards
                length += SemanticArchiveMetadata.RenderCard(card).Length + 1
            Next
            Return length
        End Function

        Private Sub SetChildParents(node As SemanticArchiveNode)
            For Each card As SemanticArchiveCard In node.Cards
                Dim child As SemanticArchiveNode = Nothing
                If _nodes.TryGetValue(card.TargetId, child) Then child.ParentNodeId = node.NodeId
                If _manifest.Nodes.ContainsKey(card.TargetId) Then _manifest.Nodes(card.TargetId).ParentNodeId = node.NodeId
            Next
        End Sub

        Private Sub UpdateRange(node As SemanticArchiveNode)
            If node.IsLeaf Then
                Dim shard As SemanticArchiveDocumentShard = GetShard(node.NodeId)
                node.MinKey = shard.MinKey
                node.MaxKey = shard.MaxKey
            ElseIf node.Cards.Count > 0 Then
                node.MinKey = node.Cards(0).PartitionKey
                node.MaxKey = GetNodeMaxKey(node.Cards(node.Cards.Count - 1).TargetId)
            End If
        End Sub
    End Class
End Namespace
