' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.

Option Strict On
Option Explicit On
Option Infer On

Namespace SharedLibrary

    Friend NotInheritable Class SemanticArchiveQueueState
        Public Property SchemaVersion As Integer = 1
        Public Property Revision As Long
        Public Property Dirty As Boolean
        Public Property DiscoveryPending As System.Boolean
        Public Property PermissionsDiscoveryPending As System.Boolean
        Public Property TotalItems As Integer
        Public Property ReadyItems As Integer
        Public Property DeferredItems As Integer
        Public Property RetryItems As Integer
        Public Property HostRequiredItems As Integer
        Public Property ErrorDocumentIds As New System.Collections.Generic.List(Of String)()
        Public Property UpdatedUtc As System.DateTime = System.DateTime.UtcNow
    End Class

    ''' <summary>
    ''' Lightweight queue scheduling cache. A small durable revision/dirty marker is
    ''' committed around every checkpoint mutation under the archive writer lease.
    ''' Counts and ready selection never deserialize all queued document/card records.
    ''' Another writer or an interrupted commit causes one header-only reconciliation.
    ''' </summary>
    Friend NotInheritable Class SemanticArchiveQueueIndex
        Private NotInheritable Class Header
            Public Property DocumentId As String
            Public Property Signature As String
            Public Property State As String
            Public Property DueUtc As System.DateTime
            Public Property HasError As Boolean
            Public Property Bucket As String
        End Class

        Private NotInheritable Class DueComparer
            Implements System.Collections.Generic.IComparer(Of Header)
            Public Function Compare(left As Header, right As Header) As Integer Implements System.Collections.Generic.IComparer(Of Header).Compare
                Dim byTime As Integer = left.DueUtc.CompareTo(right.DueUtc)
                Return If(byTime <> 0, byTime, System.StringComparer.Ordinal.Compare(left.DocumentId, right.DocumentId))
            End Function
        End Class

        Private Shared ReadOnly CacheGate As New Object()
        Private Shared ReadOnly Cache As New System.Collections.Generic.Dictionary(Of String, SemanticArchiveQueueIndex)(System.StringComparer.Ordinal)
        Private ReadOnly _statePath As String
        Private ReadOnly _itemsDirectory As String
        Private ReadOnly _headers As New System.Collections.Generic.Dictionary(Of String, Header)(System.StringComparer.Ordinal)
        Private ReadOnly _all As New System.Collections.Generic.SortedSet(Of System.String)(System.StringComparer.Ordinal)
        Private ReadOnly _ready As New System.Collections.Generic.SortedSet(Of String)(System.StringComparer.Ordinal)
        Private ReadOnly _hostReady As New System.Collections.Generic.SortedSet(Of String)(System.StringComparer.Ordinal)
        Private ReadOnly _deferred As New System.Collections.Generic.SortedSet(Of Header)(New DueComparer())
        Private ReadOnly _errors As New System.Collections.Generic.SortedSet(Of String)(System.StringComparer.Ordinal)
        Private _state As SemanticArchiveQueueState
        Private _valid As Boolean = True
        Private _retryCount As Integer
        Private _hostCount As Integer

        Private Sub New(statePath As String, itemsDirectory As String, state As SemanticArchiveQueueState)
            _statePath = statePath
            _itemsDirectory = itemsDirectory
            _state = state
        End Sub

        Public Shared Function Open(workDirectory As String, itemsDirectory As String) As SemanticArchiveQueueIndex
            Dim statePath As String = System.IO.Path.Combine(workDirectory, "queue-state.json")
            Dim state As SemanticArchiveQueueState = ReadState(workDirectory)
            SyncLock CacheGate
                Dim existing As SemanticArchiveQueueIndex = Nothing
                If state IsNot Nothing AndAlso Not state.Dirty AndAlso Cache.TryGetValue(statePath, existing) AndAlso existing._valid AndAlso existing._state.Revision = state.Revision Then
                    existing.RefreshDue()
                    Return existing
                End If
            End SyncLock
            ' The caller holds this archive's storage writer lease. Header reconciliation
            ' does no filesystem I/O under the process-wide cache dictionary lock.
            Dim rebuilt As New SemanticArchiveQueueIndex(statePath, itemsDirectory, If(state, New SemanticArchiveQueueState()))
            For Each path As String In System.IO.Directory.EnumerateFiles(itemsDirectory, "*.json", System.IO.SearchOption.TopDirectoryOnly)
                Dim item As SemanticArchiveWorkItem = ReadHeader(path)
                If item.DocumentId <> System.IO.Path.GetFileNameWithoutExtension(path) Then Throw New System.IO.InvalidDataException("A queue checkpoint identity does not match its filename.")
                rebuilt.Put(item)
            Next
            If rebuilt._state.Revision = System.Int64.MaxValue Then Throw New System.IO.InvalidDataException("The archive queue revision is exhausted.")
            rebuilt._state.Revision += 1
            rebuilt._state.Dirty = False
            rebuilt.WriteState()
            SyncLock CacheGate
                If Cache.Count >= 8 AndAlso Not Cache.ContainsKey(statePath) Then
                    Dim first As String = Nothing
                    For Each key As String In Cache.Keys
                        first = key
                        Exit For
                    Next
                    If first IsNot Nothing Then Cache.Remove(first)
                End If
                Cache(statePath) = rebuilt
            End SyncLock
            Return rebuilt
        End Function

        Public Shared Function ReadState(workDirectory As String) As SemanticArchiveQueueState
            Dim path As String = System.IO.Path.Combine(workDirectory, "queue-state.json")
            Dim state As SemanticArchiveQueueState = SemanticArchiveWorkQueue.ReadPrivateControl(Of SemanticArchiveQueueState)(workDirectory, path, 131072)
            If state Is Nothing Then Return Nothing
            If state Is Nothing OrElse state.SchemaVersion <> 1 OrElse state.Revision < 0 OrElse state.TotalItems < 0 OrElse state.ErrorDocumentIds Is Nothing OrElse state.ErrorDocumentIds.Count > 100 Then Throw New System.IO.InvalidDataException("The durable queue control record is invalid.")
            Return state
        End Function

        Public Sub BeginMutation()
            If Not _valid OrElse _state.Dirty Then Throw New System.IO.InvalidDataException("The queue requires reconciliation before another mutation.")
            If _state.Revision = System.Int64.MaxValue Then Throw New System.IO.InvalidDataException("The archive queue revision is exhausted.")
            _state.Revision += 1
            _state.Dirty = True
            SemanticArchiveStore.AtomicWriteJson(_statePath, _state)
        End Sub

        Public Sub CommitMutation()
            _state.Dirty = False
            WriteState()
        End Sub

        Public Sub SetDiscoveryPending(pending As System.Boolean, permissionsOnly As System.Boolean)
            If If(permissionsOnly, _state.PermissionsDiscoveryPending, _state.DiscoveryPending) = pending Then Return
            BeginMutation()
            If permissionsOnly Then
                _state.PermissionsDiscoveryPending = pending
            Else
                _state.DiscoveryPending = pending
            End If
            CommitMutation()
        End Sub

        Public Sub Invalidate()
            _valid = False
        End Sub

        Private Sub WriteState()
            RefreshDue()
            _state.TotalItems = _headers.Count
            _state.ReadyItems = _ready.Count + _hostReady.Count
            _state.DeferredItems = _deferred.Count
            _state.RetryItems = _retryCount
            _state.HostRequiredItems = _hostCount
            _state.ErrorDocumentIds.Clear()
            For Each identity As String In _errors
                _state.ErrorDocumentIds.Add(identity)
                If _state.ErrorDocumentIds.Count >= 100 Then Exit For
            Next
            _state.UpdatedUtc = System.DateTime.UtcNow
            SemanticArchiveStore.AtomicWriteJson(_statePath, _state)
        End Sub

        Public Function Matches(item As SemanticArchiveWorkItem) As Boolean
            Dim existing As Header = Nothing
            Return Not item.ForceSemanticRebuild AndAlso Not item.ForceExtractionRebuild AndAlso _headers.TryGetValue(item.DocumentId, existing) AndAlso existing.Signature = ProcessingSignature(item)
        End Function

        Public Sub Put(item As SemanticArchiveWorkItem)
            Remove(item.DocumentId)
            Dim header As New Header() With {
                .DocumentId = item.DocumentId, .Signature = ProcessingSignature(item), .State = item.State,
                .DueUtc = item.NextAttemptUtc, .HasError = Not System.String.IsNullOrWhiteSpace(item.LastError)
            }
            _headers.Add(header.DocumentId, header)
            _all.Add(header.DocumentId)
            If header.State = "Retry" Then _retryCount += 1
            If header.State = "PendingHost" Then _hostCount += 1
            If header.HasError Then _errors.Add(header.DocumentId)
            AddBucket(header)
        End Sub

        Public Sub Remove(documentId As String)
            Dim old As Header = Nothing
            If Not _headers.TryGetValue(documentId, old) Then Return
            Select Case old.Bucket
                Case "ready" : _ready.Remove(documentId)
                Case "host" : _hostReady.Remove(documentId)
                Case "deferred" : _deferred.Remove(old)
            End Select
            If old.State = "Retry" Then _retryCount -= 1
            If old.State = "PendingHost" Then _hostCount -= 1
            _errors.Remove(documentId)
            _headers.Remove(documentId)
            _all.Remove(documentId)
        End Sub

        Private Sub AddBucket(header As Header)
            If header.DueUtc > System.DateTime.UtcNow Then
                header.Bucket = "deferred"
                _deferred.Add(header)
            ElseIf header.State = "PendingHost" Then
                header.Bucket = "host"
                _hostReady.Add(header.DocumentId)
            Else
                header.Bucket = "ready"
                _ready.Add(header.DocumentId)
            End If
        End Sub

        Private Sub RefreshDue()
            Dim now As System.DateTime = System.DateTime.UtcNow
            While _deferred.Count > 0 AndAlso _deferred.Min.DueUtc <= now
                Dim nextItem As Header = _deferred.Min
                _deferred.Remove(nextItem)
                AddBucket(nextItem)
            End While
        End Sub

        Public Function PendingCount(runnableOnly As Boolean, isBackground As Boolean, Optional selectedDocumentIds As System.Collections.Generic.IEnumerable(Of System.String) = Nothing) As Integer
            RefreshDue()
            If selectedDocumentIds IsNot Nothing Then
                Dim count As System.Int32 = 0
                Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
                For Each id As System.String In selectedDocumentIds
                    Dim item As Header = Nothing
                    If seen.Add(id) AndAlso _headers.TryGetValue(id, item) AndAlso
                        (Not runnableOnly OrElse (item.Bucket = "ready" OrElse (item.Bucket = "host" AndAlso Not isBackground))) Then count += 1
                Next
                Return count
            End If
            If Not runnableOnly Then Return _headers.Count
            Return _ready.Count + If(isBackground, 0, _hostReady.Count)
        End Function

        Public Function GetRunnableIds(maximum As Integer, retryFailures As Boolean, isBackground As Boolean, Optional selectedDocumentIds As System.Collections.Generic.IEnumerable(Of System.String) = Nothing) As System.Collections.Generic.List(Of String)
            RefreshDue()
            Dim selection As System.Collections.Generic.HashSet(Of System.String) = If(selectedDocumentIds Is Nothing, Nothing, New System.Collections.Generic.HashSet(Of System.String)(selectedDocumentIds, System.StringComparer.Ordinal))
            Dim result As New System.Collections.Generic.List(Of String)()
            For Each identity As String In _ready
                If selection IsNot Nothing AndAlso Not selection.Contains(identity) Then Continue For
                result.Add(identity)
                If result.Count >= maximum Then Return result
            Next
            If Not isBackground Then
                For Each identity As String In _hostReady
                    If selection IsNot Nothing AndAlso Not selection.Contains(identity) Then Continue For
                    result.Add(identity)
                    If result.Count >= maximum Then Return result
                Next
            End If
            If retryFailures Then
                For Each header As Header In _deferred
                    If selection IsNot Nothing AndAlso Not selection.Contains(header.DocumentId) Then Continue For
                    If isBackground AndAlso header.State = "PendingHost" Then Continue For
                    result.Add(header.DocumentId)
                    If result.Count >= maximum Then Return result
                Next
            End If
            Return result
        End Function

        Public Function IdsAfter(afterId As System.String, maximum As System.Int32) As System.Collections.Generic.List(Of System.String)
            Dim result As New System.Collections.Generic.List(Of System.String)()
            If maximum < 1 OrElse _all.Count = 0 Then Return result
            Dim lower As System.String = If(afterId, "")
            If System.StringComparer.Ordinal.Compare(lower, _all.Max) > 0 Then Return result
            For Each id As System.String In _all.GetViewBetween(lower, _all.Max)
                If System.StringComparer.Ordinal.Compare(id, lower) <= 0 Then Continue For
                result.Add(id)
                If result.Count >= maximum Then Exit For
            Next
            Return result
        End Function

        Public Function AllIds() As System.Collections.Generic.List(Of String)
            Return New System.Collections.Generic.List(Of String)(_headers.Keys)
        End Function

        Private Shared Function ProcessingSignature(item As SemanticArchiveWorkItem) As String
            Dim bindings As New System.Collections.Generic.List(Of String)(item.BindingIds)
            bindings.Sort(System.StringComparer.Ordinal)
            Return SemanticArchiveIdentity.StableId("work", Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                .canonical_source = item.CanonicalSourceKey, .source_path = item.SourcePath,
                .hash = item.SourceHash, .length = item.SourceLength, .write_ticks = item.SourceWriteTicks,
                .extraction = item.ExtractionSignature, .semantic = item.SemanticSignature,
                .partition = item.PartitionKey, .bindings = bindings, .remove = item.Remove, .index_only = item.IndexOnlyRebuild
            }))
        End Function

        ''' <summary>Reads only host job fields; the large CachedDocument property is never loaded.</summary>
        Public Shared Function ReadHeader(path As String) As SemanticArchiveWorkItem
            SemanticArchiveStore.RequirePrivateArtifact(path)
            Dim value As New Newtonsoft.Json.Linq.JObject()
            Using stream As New System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite Or System.IO.FileShare.Delete)
                Using textReader As New System.IO.StreamReader(stream, System.Text.Encoding.UTF8, True, 4096, False)
                    Using reader As New Newtonsoft.Json.JsonTextReader(textReader)
                        reader.MaxDepth = 80
                        While reader.Read()
                            If reader.TokenType <> Newtonsoft.Json.JsonToken.PropertyName OrElse reader.Depth <> 1 Then Continue While
                            Dim name As String = CStr(reader.Value)
                            If name = "CachedDocument" Then Exit While
                            If Not reader.Read() Then Throw New System.IO.InvalidDataException("A queue checkpoint header ended unexpectedly.")
                            value(name) = Newtonsoft.Json.Linq.JToken.ReadFrom(reader)
                        End While
                    End Using
                End Using
            End Using
            Dim item As SemanticArchiveWorkItem = value.ToObject(Of SemanticArchiveWorkItem)()
            If item Is Nothing OrElse item.BindingIds Is Nothing Then Throw New System.IO.InvalidDataException("A queue checkpoint header is invalid.")
            SemanticArchiveIdentity.ValidateId(item.DocumentId, NameOf(item.DocumentId))
            Return item
        End Function
    End Class
End Namespace
