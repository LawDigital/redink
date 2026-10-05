' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
Option Strict On
Option Explicit On
Option Infer On

Namespace SharedLibrary
    ''' <summary>Independent per-user permission maintenance; never aliases or saves content-index controls.</summary>
    Public NotInheritable Class SemanticArchivePermissionMaintenanceSettings
        Public NotInheritable Class Controls
            Public ReadOnly Property Enabled As Boolean
            Public ReadOnly Property Window As String
            Public Sub New(enabled As Boolean, window As String)
                Me.Enabled = enabled
                Me.Window = If(window, SharedMethods.DEFAULT_SEMANTICARCHIVE_PERMISSION_MAINTENANCE_WINDOW)
            End Sub
        End Class

        Private Shared ReadOnly SettingsLock As System.Object = SemanticArchiveConfiguration.UserSettingsGate
        Private Sub New()
        End Sub

        Public Shared Function ReadControls(Optional context As SharedContext.ISharedContext = Nothing) As Controls
            SyncLock SettingsLock
                Dim effective As SemanticArchiveConfiguration.ControlsSnapshot = SemanticArchiveConfiguration.ReadEffectiveControls(context)
                SemanticArchiveConfiguration.ApplyControls(context, effective)
                Return New Controls(effective.PermissionEnabled, effective.PermissionWindow)
            End SyncLock
        End Function

        Public Shared Sub Save(enabled As System.Boolean, window As System.String, Optional context As SharedContext.ISharedContext = Nothing)
            SemanticArchiveConfiguration.SaveUserControls(context, New System.Collections.Generic.Dictionary(Of System.String, System.String) From {
                {SemanticArchiveConfiguration.PermissionEnabledSetting, enabled.ToString()},
                {SemanticArchiveConfiguration.PermissionWindowSetting, If(window, SharedMethods.DEFAULT_SEMANTICARCHIVE_PERMISSION_MAINTENANCE_WINDOW)}})
            SemanticArchivePermissionMaintenanceProvider.ApplyCurrentSettings()
        End Sub
    End Class

    ''' <summary>Bounded permission reconciliation is scheduled independently of all content indexing flags.
    ''' Registered and hidden archives remain eligible because generated artifacts still require maintenance.</summary>
    Public NotInheritable Class SemanticArchivePermissionMaintenanceProvider
        Implements System.IDisposable

        Private Const ProviderId As String = "semantic-archive-rights"
        Private Shared ReadOnly ProvidersLock As New Object()
        Private Shared ReadOnly Providers As New System.Collections.Generic.List(Of SemanticArchivePermissionMaintenanceProvider)()
        Private ReadOnly _context As SharedContext.ISharedContext
        Private ReadOnly _coordinator As BackgroundMaintenanceCoordinator
        Private _controls As SemanticArchivePermissionMaintenanceSettings.Controls
        Private _configuredPath As String
        Private _lastArchiveId As String = ""
        Private ReadOnly _archiveDue As New System.Collections.Generic.Dictionary(Of String, System.DateTime)(System.StringComparer.Ordinal)
        Private _paused As Integer
        Private _disposed As Integer

        Public Sub New(context As SharedContext.ISharedContext, coordinator As BackgroundMaintenanceCoordinator)
            If context Is Nothing Then Throw New System.ArgumentNullException(NameOf(context))
            If coordinator Is Nothing Then Throw New System.ArgumentNullException(NameOf(coordinator))
            _context = context
            _coordinator = coordinator
            _configuredPath = If(context.INI_SemanticArchiveCatalogPathLocal, "")
            _controls = SemanticArchivePermissionMaintenanceSettings.ReadControls(_context)
            SyncLock ProvidersLock
                Providers.Add(Me)
            End SyncLock
        End Sub

        Public Function Registration() As BackgroundMaintenanceCoordinator.Provider
            Return New BackgroundMaintenanceCoordinator.Provider() With {
                .Id = ProviderId,
                .RefreshControls = AddressOf RefreshControls,
                .IsEnabled = Function() System.Threading.Volatile.Read(_disposed) = 0 AndAlso
                    System.Threading.Volatile.Read(_paused) = 0 AndAlso (_controls.Enabled OrElse SemanticArchiveLibrary.IsConfigured(_context)) AndAlso
                    Not System.String.IsNullOrWhiteSpace(_context.INI_SemanticArchiveCatalogPathLocal),
                .CanRunNow = Function() BackgroundProcessingWindow.Allows(_controls.Window, System.DateTime.Now),
                .RunBatchAsync = AddressOf RunBatchAsync,
                .MinimumIntervalSeconds = 15,
                .IdleIntervalSeconds = 30,
                .Shutdown = AddressOf Dispose
            }
        End Function

        Private Sub RefreshControls()
            Dim nextControls = SemanticArchivePermissionMaintenanceSettings.ReadControls(_context)
            Dim currentPath = If(_context.INI_SemanticArchiveCatalogPathLocal, "")
            _controls = nextControls
            If (Not nextControls.Enabled AndAlso Not SemanticArchiveLibrary.IsConfigured(_context)) OrElse Not BackgroundProcessingWindow.Allows(nextControls.Window, System.DateTime.Now) OrElse
               Not System.String.Equals(currentPath, _configuredPath, System.StringComparison.Ordinal) Then
                _coordinator.CancelProvider(ProviderId)
            End If
            _configuredPath = currentPath
        End Sub

        Private Async Function RunBatchAsync(cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of Integer)
            cancellationToken.ThrowIfCancellationRequested()
            Dim store As New SemanticArchiveStore(_context.INI_SemanticArchiveCatalogPathLocal)
            Dim eligible As New System.Collections.Generic.List(Of SemanticArchiveDefinition)()
            For Each archive In store.LoadCatalog().Archives
                Dim due As System.DateTime = System.DateTime.MinValue
                _archiveDue.TryGetValue(store.DirectoryPath & "|" & archive.ArchiveId, due)
                If archive.Roots.Count > 0 AndAlso due <= System.DateTime.UtcNow AndAlso
                    (Not SemanticArchiveLibrary.IsSubscriber(archive) OrElse SemanticArchiveLibrary.CanUseLocally(archive, _context.INI_SemanticArchiveCatalogLibraryPath)) AndAlso
                    (_controls.Enabled OrElse SemanticArchiveLibrary.IsSubscriber(archive)) Then eligible.Add(archive)
            Next
            If eligible.Count = 0 Then Return 0
            eligible.Sort(Function(left, right) System.StringComparer.Ordinal.Compare(left.ArchiveId, right.ArchiveId))
            Dim nextIndex As Integer = 0
            For index As Integer = 0 To eligible.Count - 1
                If System.String.Equals(eligible(index).ArchiveId, _lastArchiveId, System.StringComparison.Ordinal) Then
                    nextIndex = (index + 1) Mod eligible.Count
                    Exit For
                End If
            Next
            Dim archiveId = eligible(nextIndex).ArchiveId
            _lastArchiveId = archiveId
            Dim builder As New SemanticArchiveBuilder(_context, store)
            Dim result = Await builder.BuildAsync(archiveId, New SemanticArchiveBuildOptions() With {
                .ReconcilePermissionsOnly = True, .IsBackground = True, .ForceScan = False,
                .MaximumFilesPerBatch = SharedMethods.DEFAULT_SEMANTICARCHIVE_PERMISSION_MAXIMUM_FILES_PER_BATCH, .MaxDiscoveryEntries = SharedMethods.DEFAULT_SEMANTICARCHIVE_PERMISSION_MAX_DISCOVERY_ENTRIES, .MaxDiscoverySeconds = SharedMethods.DEFAULT_SEMANTICARCHIVE_PERMISSION_MAX_DISCOVERY_SECONDS
            }, Nothing, cancellationToken).ConfigureAwait(False)
            For Each diagnostic In result.Diagnostics
                System.Diagnostics.Debug.WriteLine("Semantic Archive permissions: " & diagnostic)
            Next
            cancellationToken.ThrowIfCancellationRequested()
            If result.Cancelled Then Throw New System.OperationCanceledException(cancellationToken)
            Dim archiveKey = store.DirectoryPath & "|" & archiveId
            If result.WriterLeaseDeferred OrElse (result.PermissionsDeferred AndAlso Not result.DiscoveryPending) Then
                _archiveDue(archiveKey) = System.DateTime.UtcNow.AddSeconds(60)
                Return 0
            End If
            ' Deadlines are per archive: one completed or busy archive cannot delay
            ' another archive's first pass or durable continuation.
            Dim pending As Boolean = result.PermissionsPending OrElse result.DiscoveryPending
            _archiveDue(archiveKey) = System.DateTime.UtcNow.AddSeconds(If(pending, 15, 300))
            Return If(pending, 1, 0)
        End Function

        Public Shared Sub ApplyCurrentSettings()
            SyncLock ProvidersLock
                For Each provider In Providers
                    provider.RefreshControls()
                Next
            End SyncLock
        End Sub

        Public Shared Function PauseAutomatic(context As SharedContext.ISharedContext) As System.IDisposable
            Dim affected As New System.Collections.Generic.List(Of SemanticArchivePermissionMaintenanceProvider)()
            SyncLock ProvidersLock
                For Each provider In Providers
                    If System.Object.ReferenceEquals(context, provider._context) Then
                        System.Threading.Interlocked.Increment(provider._paused)
                        provider._coordinator.CancelProvider(ProviderId)
                        affected.Add(provider)
                    End If
                Next
            End SyncLock
            Return New PauseLease(affected)
        End Function

        Public Sub Dispose() Implements System.IDisposable.Dispose
            If System.Threading.Interlocked.Exchange(_disposed, 1) <> 0 Then Return
            SyncLock ProvidersLock
                Providers.Remove(Me)
            End SyncLock
        End Sub

        Private NotInheritable Class PauseLease
            Implements System.IDisposable
            Private ReadOnly _affected As System.Collections.Generic.List(Of SemanticArchivePermissionMaintenanceProvider)
            Private _disposed As Integer
            Public Sub New(affected As System.Collections.Generic.List(Of SemanticArchivePermissionMaintenanceProvider))
                _affected = affected
            End Sub
            Public Sub Dispose() Implements System.IDisposable.Dispose
                If System.Threading.Interlocked.Exchange(_disposed, 1) <> 0 Then Return
                For Each provider In _affected
                    System.Threading.Interlocked.Decrement(provider._paused)
                Next
            End Sub
        End Class
    End Class
End Namespace
