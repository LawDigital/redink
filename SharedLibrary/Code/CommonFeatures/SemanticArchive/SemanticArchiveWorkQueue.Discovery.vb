' Part of "Red Ink" (SharedLibrary)
' Durable bounded discovery state; no source contents or model summaries are retained here.
Option Strict On
Option Explicit On
Option Infer On

Namespace SharedLibrary
    Friend NotInheritable Class SemanticArchiveDirectoryWork
        Public Property CycleId As System.String = ""
        Public Property Position As System.Int64
        Public Property BindingId As System.String = ""
        Public Property SourcePath As System.String = ""
    End Class

    Friend NotInheritable Class SemanticArchiveDiscoverySeen
        Public Property CycleId As System.String = ""
        Public Property DocumentId As System.String = ""
        Public Property BindingIds As New System.Collections.Generic.List(Of System.String)()
    End Class

    Friend NotInheritable Class SemanticArchivePermissionRetryState
        Public Property Head As System.Int64
        Public Property Tail As System.Int64
    End Class

    Friend NotInheritable Class SemanticArchivePermissionRetryRecord
        Public Property Position As System.Int64
        Public Property Completed As System.Boolean
        Public Property Item As SemanticArchiveWorkItem
    End Class

    Friend NotInheritable Partial Class SemanticArchiveWorkQueue
        Private ReadOnly _discoveryBases As New System.Collections.Generic.Dictionary(Of System.Boolean, SemanticArchiveGenerationManifest)()
        Private _permissionRetryState As SemanticArchivePermissionRetryState

        Private Function PermissionRetryState() As SemanticArchivePermissionRetryState
            If _permissionRetryState Is Nothing Then
                _permissionRetryState = ReadDiscoveryRecord(Of SemanticArchivePermissionRetryState)(System.IO.Path.Combine(DiscoveryRoot(True), "retry-state.json"), True)
                If _permissionRetryState Is Nothing Then _permissionRetryState = New SemanticArchivePermissionRetryState()
                If _permissionRetryState.Head < 0 OrElse _permissionRetryState.Tail < _permissionRetryState.Head Then Throw New System.IO.InvalidDataException("The permission retry frontier is invalid.")
            End If
            Return _permissionRetryState
        End Function

        Private Sub SavePermissionRetryState()
            SemanticArchiveStore.AtomicWriteJson(System.IO.Path.Combine(DiscoveryRoot(True), "retry-state.json"), PermissionRetryState())
        End Sub

        Private Function PermissionRetrySlot(position As System.Int64) As System.String
            If position < 0 Then Throw New System.IO.InvalidDataException("A permission retry position cannot be negative.")
            Return SemanticArchivePathGuard.RequireWindowsCompatiblePath(System.IO.Path.Combine(DiscoveryRoot(True), "retries",
                (position \ 256L).ToString("x", System.Globalization.CultureInfo.InvariantCulture), position.ToString("x", System.Globalization.CultureInfo.InvariantCulture) & ".json"))
        End Function

        Public Function FindPermissionRetry(documentId As System.String) As SemanticArchivePermissionRetryRecord
            SemanticArchiveIdentity.ValidateId(documentId, NameOf(documentId))
            Dim record As SemanticArchivePermissionRetryRecord = ReadDiscoveryRecord(Of SemanticArchivePermissionRetryRecord)(DiscoveryRecordPath("retry-keys", documentId, True), True)
            If record Is Nothing Then Return Nothing
            If record.Item Is Nothing OrElse record.Item.DocumentId <> documentId OrElse record.Position < 0 OrElse record.Position >= PermissionRetryState().Tail Then Throw New System.IO.InvalidDataException("The permission retry source identity is invalid.")
            If record.Completed Then Return Nothing
            Return record
        End Function

        Public Sub EnqueuePermissionRetry(item As SemanticArchiveWorkItem)
            SemanticArchiveIdentity.ValidateId(item.DocumentId, NameOf(item.DocumentId))
            item.CachedDocument = Nothing
            Dim state As SemanticArchivePermissionRetryState = PermissionRetryState()
            ' The slot is committed before the tail and deduplication reference. A crash
            ' can replay an idempotent repair but cannot silently lose accepted work.
            Dim record As New SemanticArchivePermissionRetryRecord With {.Position = state.Tail, .Item = item}
            SemanticArchiveStore.AtomicWriteJson(PermissionRetrySlot(record.Position), record)
            state.Tail += 1L
            SavePermissionRetryState()
            SemanticArchiveStore.AtomicWriteJson(DiscoveryRecordPath("retry-keys", item.DocumentId, True), record)
        End Sub

        Public Function PeekPermissionRetry() As SemanticArchivePermissionRetryRecord
            Dim state As SemanticArchivePermissionRetryState = PermissionRetryState()
            If state.Head >= state.Tail Then Return Nothing
            Dim record As SemanticArchivePermissionRetryRecord = ReadDiscoveryRecord(Of SemanticArchivePermissionRetryRecord)(PermissionRetrySlot(state.Head), True)
            If record Is Nothing OrElse record.Position <> state.Head OrElse record.Item Is Nothing Then Throw New System.IO.InvalidDataException("The permission retry frontier is incomplete; remaining rights are unknown.")
            SemanticArchiveIdentity.ValidateId(record.Item.DocumentId, "permission retry identity")
            Dim current As SemanticArchivePermissionRetryRecord = ReadDiscoveryRecord(Of SemanticArchivePermissionRetryRecord)(DiscoveryRecordPath("retry-keys", record.Item.DocumentId, True), True)
            If current IsNot Nothing Then
                If current.Item Is Nothing OrElse current.Item.DocumentId <> record.Item.DocumentId OrElse current.Position < 0 OrElse current.Position >= state.Tail Then Throw New System.IO.InvalidDataException("The permission retry reference has an invalid identity.")
                If current.Completed OrElse current.Position > record.Position Then Return Nothing
            End If
            Return record
        End Function

        Public Sub AdvancePermissionRetry()
            Dim state As SemanticArchivePermissionRetryState = PermissionRetryState()
            If state.Head >= state.Tail Then Return
            Dim oldPath As System.String = PermissionRetrySlot(state.Head)
            state.Head += 1L
            SavePermissionRetryState()
            System.IO.File.Delete(oldPath)
        End Sub

        Public Sub CompletePermissionRetry(documentId As System.String)
            Dim existing As SemanticArchivePermissionRetryRecord = FindPermissionRetry(documentId)
            If existing Is Nothing Then Return
            existing.Completed = True
            SemanticArchiveStore.AtomicWriteJson(DiscoveryRecordPath("retry-keys", documentId, True), existing)
        End Sub

        Public ReadOnly Property PermissionRetryHead As System.Int64
            Get
                Return PermissionRetryState().Head
            End Get
        End Property

        Public ReadOnly Property PermissionRetryTail As System.Int64
            Get
                Return PermissionRetryState().Tail
            End Get
        End Property

        Public Function PermissionRetrySlotCount() As System.Int64
            Dim state As SemanticArchivePermissionRetryState = PermissionRetryState()
            Return state.Tail - state.Head
        End Function

        Public Function HasPermissionRetries(Optional selectedDocumentIds As System.Collections.Generic.IEnumerable(Of System.String) = Nothing) As System.Boolean
            If selectedDocumentIds IsNot Nothing Then
                For Each id As System.String In selectedDocumentIds
                    If FindPermissionRetry(id) IsNot Nothing Then Return True
                Next
                Return False
            End If
            Dim state As SemanticArchivePermissionRetryState = PermissionRetryState()
            Return state.Head < state.Tail
        End Function

        Friend Shared Function ReadPrivateControl(Of T As Class)(root As System.String, path As System.String, maximumBytes As System.Int64) As T
            SemanticArchivePathGuard.ValidateContainedPath(root, path, False)
            Try
                System.IO.File.GetAttributes(path)
            Catch missing As System.IO.FileNotFoundException
                Return Nothing
            Catch missing As System.IO.DirectoryNotFoundException
                Return Nothing
            End Try
            SemanticArchiveStore.RequirePrivateArtifact(path)
            Using stream As System.IO.FileStream = SemanticArchivePathGuard.OpenContainedRead(root, path)
                If stream.Length > maximumBytes Then Throw New System.IO.InvalidDataException("A private control record exceeds its bounded size.")
                Using bytes As New System.IO.MemoryStream()
                    Dim buffer(8191) As System.Byte
                    While True
                        Dim count As System.Int32 = stream.Read(buffer, 0, buffer.Length)
                        If count = 0 Then Exit While
                        If bytes.Length + count > maximumBytes Then Throw New System.IO.InvalidDataException("A private control record grew beyond its bounded size.")
                        bytes.Write(buffer, 0, count)
                    End While
                    Dim text As System.String = New System.Text.UTF8Encoding(False, True).GetString(bytes.ToArray())
                    If text.Length > 0 AndAlso text(0) = System.Char.ConvertFromUtf32(&HFEFF)(0) Then text = text.Substring(1)
                    Dim settings As New Newtonsoft.Json.JsonSerializerSettings With {
                        .CheckAdditionalContent = True, .TypeNameHandling = Newtonsoft.Json.TypeNameHandling.None,
                        .MetadataPropertyHandling = Newtonsoft.Json.MetadataPropertyHandling.Ignore, .MaxDepth = 96}
                    Dim value As T = Newtonsoft.Json.JsonConvert.DeserializeObject(Of T)(text, settings)
                    If value Is Nothing Then Throw New System.IO.InvalidDataException("A private control record is incomplete.")
                    Return value
                End Using
            End Using
        End Function
        Private Function DiscoveryRoot(permissionsOnly As System.Boolean) As System.String
            Dim path As System.String = System.IO.Path.Combine(_directory, If(permissionsOnly, "permissions", "discovery"))
            SemanticArchivePathGuard.RequireWindowsCompatiblePath(path, 112)
            SemanticArchiveStore.CreatePrivateDirectory(path)
            Return path
        End Function

        Private Function DiscoveryRecordPath(kind As System.String, key As System.String, permissionsOnly As System.Boolean) As System.String
            Dim digest As System.String = SemanticArchiveIdentity.HashBytes(System.Text.Encoding.UTF8.GetBytes(key))
            Dim path As System.String = System.IO.Path.Combine(DiscoveryRoot(permissionsOnly), kind, digest.Substring(0, 2), digest & ".json")
            Return SemanticArchivePathGuard.RequireWindowsCompatiblePath(path)
        End Function

        Private Function FrontierPath(position As System.Int64, permissionsOnly As System.Boolean) As System.String
            If position < 0 Then Throw New System.IO.InvalidDataException("A directory cursor cannot be negative.")
            Dim path As System.String = System.IO.Path.Combine(DiscoveryRoot(permissionsOnly), "dirs",
                (position \ 256L).ToString("x", System.Globalization.CultureInfo.InvariantCulture),
                position.ToString("x", System.Globalization.CultureInfo.InvariantCulture) & ".json")
            Return SemanticArchivePathGuard.RequireWindowsCompatiblePath(path)
        End Function

        Private Function ReadDiscoveryRecord(Of T As Class)(path As System.String, permissionsOnly As System.Boolean, Optional maximumBytes As System.Int64 = 131072) As T
            Return ReadPrivateControl(Of T)(DiscoveryRoot(permissionsOnly), path, maximumBytes)
        End Function

        Public Sub EnqueueDirectory(checkpoint As SemanticArchiveScanCheckpoint, bindingId As System.String, sourcePath As System.String, permissionsOnly As System.Boolean)
            SemanticArchiveIdentity.ValidateId(bindingId, NameOf(bindingId))
            Dim source As System.String = SemanticArchivePathGuard.RequireWindowsSourcePath(sourcePath)
            Dim identity As System.String = bindingId & "|" & source
            Dim recordPath As System.String = DiscoveryRecordPath("directory-keys", identity, permissionsOnly)
            Dim existing As SemanticArchiveDirectoryWork = ReadDiscoveryRecord(Of SemanticArchiveDirectoryWork)(recordPath, permissionsOnly)
            If existing IsNot Nothing AndAlso existing.CycleId = checkpoint.CycleId Then
                ' Restore a frontier link if a writer stopped between its two private
                ' records and the scan checkpoint commit. No source is removed here.
                SemanticArchiveStore.AtomicWriteJson(FrontierPath(existing.Position, permissionsOnly), existing)
                checkpoint.DirectoryTail = System.Math.Max(checkpoint.DirectoryTail, existing.Position + 1L)
                Return
            End If
            While True
                Dim occupied As SemanticArchiveDirectoryWork = ReadDiscoveryRecord(Of SemanticArchiveDirectoryWork)(FrontierPath(checkpoint.DirectoryTail, permissionsOnly), permissionsOnly)
                If occupied Is Nothing OrElse occupied.CycleId <> checkpoint.CycleId Then Exit While
                checkpoint.DirectoryTail += 1L
            End While
            Dim work As New SemanticArchiveDirectoryWork With {
                .CycleId = checkpoint.CycleId, .Position = checkpoint.DirectoryTail, .BindingId = bindingId, .SourcePath = source}
            SemanticArchiveStore.AtomicWriteJson(recordPath, work)
            SemanticArchiveStore.AtomicWriteJson(FrontierPath(checkpoint.DirectoryTail, permissionsOnly), work)
            checkpoint.DirectoryTail += 1L
        End Sub

        Public Function LoadDirectory(checkpoint As SemanticArchiveScanCheckpoint, permissionsOnly As System.Boolean) As SemanticArchiveDirectoryWork
            If checkpoint.DirectoryHead >= checkpoint.DirectoryTail Then Return Nothing
            Dim work As SemanticArchiveDirectoryWork = ReadDiscoveryRecord(Of SemanticArchiveDirectoryWork)(FrontierPath(checkpoint.DirectoryHead, permissionsOnly), permissionsOnly)
            If work Is Nothing OrElse work.CycleId <> checkpoint.CycleId OrElse work.Position <> checkpoint.DirectoryHead Then
                Throw New System.IO.InvalidDataException("The durable directory frontier is incomplete; discovery remains unknown.")
            End If
            Return work
        End Function

        Public Function ReadSeen(checkpoint As SemanticArchiveScanCheckpoint, documentId As System.String, permissionsOnly As System.Boolean) As SemanticArchiveDiscoverySeen
            SemanticArchiveIdentity.ValidateId(documentId, NameOf(documentId))
            Dim seen As SemanticArchiveDiscoverySeen = ReadDiscoveryRecord(Of SemanticArchiveDiscoverySeen)(DiscoveryRecordPath("seen", documentId, permissionsOnly), permissionsOnly)
            If seen Is Nothing OrElse seen.CycleId <> checkpoint.CycleId Then Return Nothing
            If seen.DocumentId <> documentId OrElse seen.BindingIds Is Nothing Then Throw New System.IO.InvalidDataException("A source discovery marker has an invalid identity.")
            Return seen
        End Function

        Public Sub MarkSeen(checkpoint As SemanticArchiveScanCheckpoint, documentId As System.String, bindingIds As System.Collections.Generic.IEnumerable(Of System.String), permissionsOnly As System.Boolean)
            Dim seen As SemanticArchiveDiscoverySeen = ReadSeen(checkpoint, documentId, permissionsOnly)
            If seen Is Nothing Then seen = New SemanticArchiveDiscoverySeen With {.CycleId = checkpoint.CycleId, .DocumentId = documentId}
            For Each bindingId As System.String In bindingIds
                SemanticArchiveIdentity.ValidateId(bindingId, NameOf(bindingId))
                If Not seen.BindingIds.Contains(bindingId) Then seen.BindingIds.Add(bindingId)
            Next
            SemanticArchiveStore.AtomicWriteJson(DiscoveryRecordPath("seen", documentId, permissionsOnly), seen)
        End Sub

        Public Sub SaveDiscoveryBase(generation As SemanticArchiveGenerationManifest, permissionsOnly As System.Boolean)
            Dim path As System.String = System.IO.Path.Combine(DiscoveryRoot(permissionsOnly), "base.json")
            _discoveryBases(permissionsOnly) = generation
            If generation Is Nothing Then
                Dim prior As SemanticArchiveGenerationManifest = ReadPrivateControl(Of SemanticArchiveGenerationManifest)(DiscoveryRoot(permissionsOnly), path, 64L * 1024L * 1024L)
                If prior IsNot Nothing Then System.IO.File.Delete(path)
            Else
                SemanticArchiveStore.AtomicWriteJson(path, generation)
            End If
        End Sub

        Public Function LoadDiscoveryBase(archiveId As System.String, permissionsOnly As System.Boolean) As SemanticArchiveGenerationManifest
            Dim generation As SemanticArchiveGenerationManifest = Nothing
            If Not _discoveryBases.TryGetValue(permissionsOnly, generation) Then
                generation = ReadDiscoveryRecord(Of SemanticArchiveGenerationManifest)(
                    System.IO.Path.Combine(DiscoveryRoot(permissionsOnly), "base.json"), permissionsOnly, 64L * 1024L * 1024L)
                _discoveryBases(permissionsOnly) = generation
            End If
            If generation IsNot Nothing AndAlso generation.ArchiveId <> archiveId Then Throw New System.IO.InvalidDataException("The discovery generation belongs to another archive.")
            Return generation
        End Function
    End Class
End Namespace
