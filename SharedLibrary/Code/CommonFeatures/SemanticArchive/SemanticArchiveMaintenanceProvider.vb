' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SemanticArchiveMaintenanceProvider.vb
' Purpose:
'   Configured archive-content maintenance provider and per-user background processing
'   controls.
'
' Architecture / Function:
'   Registers cancellable bounded work with the generic coordinator; host wiring
'   determines automatic versus explicit use.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary

    ''' <summary>Only SA settings are persisted here. The corresponding KB settings retain their existing names and values.</summary>
    Public NotInheritable Class SemanticArchiveBackgroundSettings
        Private Shared ReadOnly SettingsLock As System.Object = SemanticArchiveConfiguration.UserSettingsGate
        Private Sub New()
        End Sub

        Public Shared Sub ReadInto(context As SharedContext.ISharedContext)
            If context Is Nothing Then Return
            SyncLock SettingsLock
                SemanticArchiveConfiguration.ApplyControls(context, SemanticArchiveConfiguration.ReadEffectiveControls(context))
            End SyncLock
        End Sub

        Public Shared Sub Save(context As SharedContext.ISharedContext, enabled As System.Boolean, window As System.String)
            If context Is Nothing Then Throw New System.ArgumentNullException(NameOf(context))
            SemanticArchiveConfiguration.SaveUserControls(context, New System.Collections.Generic.Dictionary(Of System.String, System.String) From {
                {SemanticArchiveConfiguration.BackgroundEnabledSetting, enabled.ToString()},
                {SemanticArchiveConfiguration.BackgroundWindowSetting, If(window, SharedMethods.DEFAULT_SEMANTICARCHIVE_BACKGROUND_INDEXING_WINDOW)}})
            SemanticArchiveMaintenanceProvider.ApplyCurrentSettings(context)
        End Sub
    End Class

    ''' <summary>Archive-specific control and durable-build adapter for the generic Office coordinator.</summary>
    Public NotInheritable Class SemanticArchiveMaintenanceProvider
        Implements System.IDisposable

        Private Shared ReadOnly ProvidersLock As New Object()
        Private Shared ReadOnly Providers As New System.Collections.Generic.List(Of SemanticArchiveMaintenanceProvider)()
        Private ReadOnly _context As SharedContext.ISharedContext
        Private ReadOnly _coordinator As BackgroundMaintenanceCoordinator
        Private _lastArchiveId As String = ""
        Private _configuredPath As String = ""
        Private _disposed As Integer
        Private _paused As Integer

        Public Sub New(context As SharedContext.ISharedContext, coordinator As BackgroundMaintenanceCoordinator)
            _context = context
            _coordinator = coordinator
            _configuredPath = If(context.INI_SemanticArchiveCatalogPathLocal, "")
            SyncLock ProvidersLock
                Providers.Add(Me)
            End SyncLock
        End Sub

        Public Function Registration() As BackgroundMaintenanceCoordinator.Provider
            Return New BackgroundMaintenanceCoordinator.Provider() With {
                .Id = "semantic-archive",
                .RefreshControls = AddressOf RefreshControls,
                .IsEnabled = Function() System.Threading.Volatile.Read(_disposed) = 0 AndAlso
                    System.Threading.Volatile.Read(_paused) = 0 AndAlso (_context.INI_SemanticArchiveBackgroundIndexing OrElse SemanticArchiveLibrary.IsConfigured(_context)) AndAlso
                    Not System.String.IsNullOrWhiteSpace(_context.INI_SemanticArchiveCatalogPathLocal),
                .CanRunNow = Function() BackgroundProcessingWindow.Allows(_context.INI_SemanticArchiveBackgroundIndexingWindow, System.DateTime.Now),
                .RunBatchAsync = AddressOf RunBatchAsync,
                .Shutdown = AddressOf Dispose
            }
        End Function

        Private Sub RefreshControls()
            SemanticArchiveBackgroundSettings.ReadInto(_context)
            Dim current = If(_context.INI_SemanticArchiveCatalogPathLocal, "")
            If Not System.String.Equals(current, _configuredPath, System.StringComparison.Ordinal) Then
                _coordinator.CancelProvider("semantic-archive")
                _configuredPath = current
            End If
        End Sub

        Private Async Function RunBatchAsync(cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of Integer)
            cancellationToken.ThrowIfCancellationRequested()
            Dim store As New SemanticArchiveStore(_context.INI_SemanticArchiveCatalogPathLocal)
            Dim catalog = store.LoadCatalog()
            Dim eligible As New System.Collections.Generic.List(Of SemanticArchiveDefinition)()
            For Each archive In catalog.Archives
                If SemanticArchiveLibrary.CanUseLocally(archive, _context.INI_SemanticArchiveCatalogLibraryPath) AndAlso
                   (SemanticArchiveLibrary.IsSubscriber(archive) OrElse (_context.INI_SemanticArchiveBackgroundIndexing AndAlso archive.BackgroundEnabled)) AndAlso archive.Roots.Count > 0 AndAlso
                   BackgroundProcessingWindow.Allows(archive.BackgroundWindow, System.DateTime.Now) Then eligible.Add(archive)
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
            Dim selected = eligible(nextIndex)
            _lastArchiveId = selected.ArchiveId
            Dim builder As New SemanticArchiveBuilder(_context, store)
            Dim result = Await builder.BuildAsync(selected.ArchiveId,
                New SemanticArchiveBuildOptions() With {.IsBackground = True, .MaximumFilesPerBatch = SharedMethods.DEFAULT_SEMANTICARCHIVE_BACKGROUND_MAXIMUM_FILES_PER_BATCH, .ForceScan = False},
                Nothing, cancellationToken).ConfigureAwait(False)
            For Each diagnostic In result.Diagnostics
                System.Diagnostics.Debug.WriteLine("Semantic Archive maintenance: " & diagnostic)
            Next
            cancellationToken.ThrowIfCancellationRequested()
            If result.Cancelled Then Throw New System.OperationCanceledException(cancellationToken)
            ' A yielded acquisition consumes no source job and is not a failure. Returning
            ' no work lets the generic coordinator apply its idle backoff and rotate providers.
            If result.WriterLeaseDeferred Then Return 0
            If result.FailedFiles > 0 AndAlso result.ProcessedFiles = 0 AndAlso Not result.Published Then
                Throw New System.InvalidOperationException("Archive batch failed; see archive diagnostics. The previous published generation remains active.")
            End If
            Return result.ProcessedFiles + result.ReusedFiles + If(result.PendingFiles > 0 OrElse result.DiscoveryPending, 1, 0)
        End Function

        Public Shared Sub ApplyCurrentSettings(context As SharedContext.ISharedContext)
            SyncLock ProvidersLock
                For Each provider In Providers
                    If System.Object.ReferenceEquals(provider._context, context) AndAlso
                       ((Not context.INI_SemanticArchiveBackgroundIndexing AndAlso Not SemanticArchiveLibrary.IsConfigured(context)) OrElse
                        Not BackgroundProcessingWindow.Allows(context.INI_SemanticArchiveBackgroundIndexingWindow, System.DateTime.Now)) Then
                        provider._coordinator.CancelProvider("semantic-archive")
                    End If
                Next
            End SyncLock
        End Sub

        ''' <summary>Manual work pauses only this provider; its own cancellation remains owned by the console.</summary>
        Public Shared Function PauseAutomatic(context As SharedContext.ISharedContext) As System.IDisposable
            Dim affected As New System.Collections.Generic.List(Of SemanticArchiveMaintenanceProvider)()
            SyncLock ProvidersLock
                For Each provider In Providers
                    If System.Object.ReferenceEquals(provider._context, context) Then
                        System.Threading.Interlocked.Increment(provider._paused)
                        provider._coordinator.CancelProvider("semantic-archive")
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
            Private ReadOnly _providers As System.Collections.Generic.List(Of SemanticArchiveMaintenanceProvider)
            Private _disposed As Integer
            Public Sub New(providers As System.Collections.Generic.List(Of SemanticArchiveMaintenanceProvider))
                _providers = providers
            End Sub
            Public Sub Dispose() Implements System.IDisposable.Dispose
                If System.Threading.Interlocked.Exchange(_disposed, 1) <> 0 Then Return
                For Each provider In _providers
                    System.Threading.Interlocked.Decrement(provider._paused)
                Next
            End Sub
        End Class
    End Class
End Namespace
