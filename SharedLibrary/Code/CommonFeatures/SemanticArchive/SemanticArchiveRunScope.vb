' Part of "Red Ink" (SharedLibrary)
' Host-owned archive selections and opaque, run-local retrieval references.
Option Strict On
Option Explicit On

Namespace SharedLibrary

    ''' <summary>
    ''' Immutable authority supplied by the host. Tool arguments may narrow this set;
    ''' they cannot grant access to another archive, source path, or another run's hits.
    ''' Child agents retain the same evidence store while receiving only a subset.
    ''' </summary>
    Public NotInheritable Class SemanticArchiveRunScope
        Private ReadOnly _archiveIds As System.Collections.ObjectModel.ReadOnlyCollection(Of System.String)
        Private ReadOnly _access As SemanticArchiveAccessContext
        Private ReadOnly _state As SemanticArchiveRunState
        Public ReadOnly Property ResolutionStatus As System.String
        Public ReadOnly Property ResolutionMessage As System.String

        Public Sub New(selectedArchiveIds As System.Collections.Generic.IEnumerable(Of System.String),
                       accessContext As SemanticArchiveAccessContext)
            Me.New(selectedArchiveIds, accessContext, New SemanticArchiveRunState(), "", "")
        End Sub

        Private Sub New(selectedArchiveIds As System.Collections.Generic.IEnumerable(Of System.String),
                        accessContext As SemanticArchiveAccessContext,
                        state As SemanticArchiveRunState,
                        resolutionStatus As System.String,
                        resolutionMessage As System.String)
            If accessContext Is Nothing Then Throw New System.ArgumentNullException(NameOf(accessContext))
            Dim ids As New System.Collections.Generic.List(Of System.String)()
            Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            If selectedArchiveIds IsNot Nothing Then
                For Each candidate As System.String In selectedArchiveIds
                    If System.String.IsNullOrWhiteSpace(candidate) Then Throw New System.ArgumentException("An archive selection contains an empty ID.", NameOf(selectedArchiveIds))
                    Dim id As System.String = candidate.Trim()
                    If seen.Add(id) Then ids.Add(id)
                Next
            End If
            _archiveIds = ids.AsReadOnly()
            _access = accessContext
            _state = state
            Me.ResolutionStatus = If(resolutionStatus, "")
            Me.ResolutionMessage = If(resolutionMessage, "")
        End Sub

        Public ReadOnly Property SelectedArchiveIds As System.Collections.Generic.IReadOnlyList(Of System.String)
            Get
                Return _archiveIds
            End Get
        End Property

        Public ReadOnly Property AccessContext As SemanticArchiveAccessContext
            Get
                Return _access
            End Get
        End Property

        Public ReadOnly Property RunId As System.String
            Get
                Return _state.RunId
            End Get
        End Property

        Public Function GetUsedSourcePaths() As System.Collections.Generic.IReadOnlyList(Of System.String)
            Dim snapshot As New System.Collections.Generic.List(Of System.String)()
            SyncLock _state.UsedSourcePaths
                snapshot.AddRange(_state.UsedSourcePaths)
            End SyncLock
            Dim result As New System.Collections.Generic.List(Of System.String)()
            For Each sourcePath As System.String In snapshot
                If System.String.IsNullOrWhiteSpace(sourcePath) Then Continue For
                If _access.CanReadSource(sourcePath) AndAlso System.IO.File.Exists(sourcePath) Then result.Add(sourcePath)
            Next
            Return result.AsReadOnly()
        End Function

        Public Function ContainsArchive(archiveId As System.String) As System.Boolean
            For Each selected As System.String In _archiveIds
                If System.String.Equals(selected, archiveId, System.StringComparison.Ordinal) Then Return True
            Next
            Return False
        End Function

        Public Function Narrow(archiveIds As System.Collections.Generic.IEnumerable(Of System.String)) As SemanticArchiveRunScope
            If archiveIds Is Nothing Then Return Me
            Dim requested As New System.Collections.Generic.List(Of System.String)()
            For Each id As System.String In archiveIds
                If Not ContainsArchive(id) Then Throw New System.UnauthorizedAccessException("The requested archive is outside the host-selected run scope.")
                requested.Add(id)
            Next
            Return New SemanticArchiveRunScope(requested, _access, _state, ResolutionStatus, ResolutionMessage)
        End Function

        ' Only the authoritative inline resolver can establish a new user selection.
        ' This method is never called with a model-generated tool argument.
        Friend Function WithAuthoritativeSelection(archiveIds As System.Collections.Generic.IEnumerable(Of System.String)) As SemanticArchiveRunScope
            Return New SemanticArchiveRunScope(archiveIds, _access, _state, ResolutionStatus, ResolutionMessage)
        End Function

        ' Host-owned diagnostic state survives delegation. Model arguments cannot clear it.
        Friend Function WithResolutionFailure(code As System.String, message As System.String) As SemanticArchiveRunScope
            Return New SemanticArchiveRunScope(New System.String() {}, _access, _state, code, message)
        End Function

        Friend ReadOnly Property State As SemanticArchiveRunState
            Get
                Return _state
            End Get
        End Property
    End Class

    Friend NotInheritable Class SemanticArchiveRunState
        Friend ReadOnly RunId As System.String = System.Guid.NewGuid().ToString("N")
        Friend ReadOnly CreatedUtc As System.DateTime = System.DateTime.UtcNow
        Friend ReadOnly OperationGate As New System.Threading.SemaphoreSlim(1, 1)
        Friend ReadOnly Generations As New System.Collections.Generic.Dictionary(Of System.String, SemanticArchiveGenerationManifest)(System.StringComparer.Ordinal)
        Friend ReadOnly Hits As New System.Collections.Generic.Dictionary(Of System.String, SemanticArchiveEvidenceReference)(System.StringComparer.Ordinal)
        Friend ReadOnly Searches As New System.Collections.Generic.Dictionary(Of System.String, SemanticArchiveSearchState)(System.StringComparer.Ordinal)
        Friend ReadOnly UsedSourcePaths As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
        Friend Property LoadedEvidenceBytes As System.Int64
        Private _directoryPath As System.String

        Friend Sub BindDirectory(directoryPath As System.String)
            If _directoryPath Is Nothing Then
                _directoryPath = directoryPath
            ElseIf Not System.String.Equals(_directoryPath, directoryPath, System.StringComparison.OrdinalIgnoreCase) Then
                Throw New System.InvalidOperationException("The configured archive directory changed during this run. Start a fresh explicit archive run.")
            End If
        End Sub
        Friend Const MaximumEvidenceReferences As System.Int32 = 2048
        Friend Const MaximumContinuations As System.Int32 = 16

        Friend Sub EnsureUsable()
            If System.DateTime.UtcNow - CreatedUtc > System.TimeSpan.FromHours(8) Then
                Throw New System.InvalidOperationException("The archive run has expired. Resolve the current user request again to select a fresh generation.")
            End If
        End Sub
    End Class

    Friend NotInheritable Class SemanticArchiveEvidenceReference
        Friend Property ArchiveId As System.String
        Friend Property DocumentId As System.String
        Friend Property Generation As SemanticArchiveGenerationManifest
        Friend Property Query As System.String
        Friend Property LiteralMatchStartByte As System.Nullable(Of System.Int64)
        Friend Property LiteralText As System.String
        Friend Property ReadState As SemanticArchiveDocumentReadState
    End Class

    ''' <summary>Optional execution-context bridge; explicit scope parameters take precedence.</summary>
    Public NotInheritable Class SemanticArchiveRunContext
        Private Shared ReadOnly Slot As New System.Threading.AsyncLocal(Of SemanticArchiveRunScope)()
        Private Sub New()
        End Sub

        Public Shared ReadOnly Property Current As SemanticArchiveRunScope
            Get
                Return Slot.Value
            End Get
        End Property

        Public Shared Function Push(scope As SemanticArchiveRunScope) As System.IDisposable
            Dim previous As SemanticArchiveRunScope = Slot.Value
            Slot.Value = scope
            Return New ScopeRestorer(previous)
        End Function

        Private NotInheritable Class ScopeRestorer
            Implements System.IDisposable
            Private ReadOnly _previous As SemanticArchiveRunScope
            Private _disposed As System.Boolean
            Public Sub New(previous As SemanticArchiveRunScope)
                _previous = previous
            End Sub
            Public Sub Dispose() Implements System.IDisposable.Dispose
                If _disposed Then Return
                Slot.Value = _previous
                _disposed = True
            End Sub
        End Class
    End Class
End Namespace
