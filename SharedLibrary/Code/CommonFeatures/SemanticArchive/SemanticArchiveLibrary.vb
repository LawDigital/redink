' Part of "Red Ink" (SharedLibrary)
' Central definitions only. Original content, private generations and source ACLs remain separate.
Option Strict On
Option Explicit On

Namespace SharedLibrary
    Public NotInheritable Class SemanticArchiveLibraryAccessException
        Inherits System.UnauthorizedAccessException
        Public ReadOnly Property Code As System.String
        Public Sub New(code As System.String, message As System.String, Optional cause As System.Exception = Nothing)
            MyBase.New(message, cause)
            Me.Code = code
        End Sub
    End Class

    Public NotInheritable Class SemanticArchiveLibraryRegistration
        Public Property Directory As System.String = ""
        Public Property EntryId As System.String = ""
        Public Property Revision As System.Int64
        Public Property PublisherSid As System.String = ""
        Public Property DefinitionHash As System.String = ""
        Public Property PublicationHash As System.String = ""
        Public Property Role As System.String = "subscriber"
        Public Property State As System.String = "available"
        Public Property OptOut As System.Boolean
    End Class

    Public NotInheritable Class SemanticArchiveLibraryEntry
        Public Property SchemaVersion As System.Int32 = 1
        Public Property EntryId As System.String = ""
        Public Property Revision As System.Int64
        Public Property PublisherSid As System.String = ""
        Public Property Withdrawn As System.Boolean
        Public Property AutoSubscribe As System.Boolean = SharedMethods.DEFAULT_SEMANTICARCHIVE_LIBRARY_AUTO_SUBSCRIBE
        Public Property Definition As SemanticArchiveDefinition
    End Class

    Public NotInheritable Partial Class SemanticArchiveLibrary
        Private Sub New()
        End Sub

        Private NotInheritable Class SyncState
            Friend ReadOnly Gate As New System.Threading.SemaphoreSlim(1, 1)
            Friend NextCheckUtc As System.DateTime = System.DateTime.MinValue
            Friend ConfigurationKey As System.String = ""
        End Class
        Private Shared ReadOnly States As New System.Runtime.CompilerServices.ConditionalWeakTable(Of SharedContext.ISharedContext, SyncState)()

        Public Shared Function IsConfigured(context As SharedContext.ISharedContext) As System.Boolean
            Return SemanticArchiveHostIntegration.IsConfigured(context) AndAlso Not System.String.IsNullOrWhiteSpace(context.INI_SemanticArchiveCatalogLibraryPath)
        End Function

        Public Shared Function IsSubscriber(definition As SemanticArchiveDefinition) As System.Boolean
            Return definition IsNot Nothing AndAlso definition.Library IsNot Nothing AndAlso definition.Library.Role = "subscriber"
        End Function

        Public Shared Function CanUseLocally(definition As SemanticArchiveDefinition, libraryPath As System.String) As System.Boolean
            If definition Is Nothing OrElse Not definition.Enabled Then Return False
            If Not IsSubscriber(definition) Then Return True
            Return Not definition.Library.OptOut AndAlso definition.Library.State = "available" AndAlso
                Not System.String.IsNullOrWhiteSpace(libraryPath) AndAlso
                System.String.Equals(SemanticArchivePathGuard.CanonicalPath(libraryPath), definition.Library.Directory, System.StringComparison.OrdinalIgnoreCase)
        End Function

        ''' <summary>Cheap local filtering for host selection. Actual library authority is
        ''' rechecked on the worker before each search/read/list/build operation.</summary>
        Public Shared Function FilterCatalog(catalog As SemanticArchiveCatalog, libraryPath As System.String) As SemanticArchiveCatalog
            Dim result As SemanticArchiveCatalog = SemanticArchiveMetadata.Clone(catalog)
            result.Archives.RemoveAll(Function(item As SemanticArchiveDefinition) Not CanUseLocally(item, libraryPath))
            ' Keep explicit/default selections even when their archive is unavailable.
            ' ResolveSelection must report that condition, never broaden to other archives.
            Return result
        End Function

        Public Shared Async Function SynchronizeAsync(context As SharedContext.ISharedContext,
                    Optional force As System.Boolean = False,
                    Optional cancellationToken As System.Threading.CancellationToken = Nothing) As System.Threading.Tasks.Task(Of System.Int32)
            If Not IsConfigured(context) Then Return 0
            Dim state As SyncState = States.GetValue(context, Function(unused As SharedContext.ISharedContext) New SyncState())
            If force Then
                Await state.Gate.WaitAsync(cancellationToken).ConfigureAwait(False)
            ElseIf Not Await state.Gate.WaitAsync(0, cancellationToken).ConfigureAwait(False) Then
                Return 0
            End If
            Try
                Dim key As System.String = context.INI_SemanticArchiveCatalogPathLocal & Microsoft.VisualBasic.ChrW(0) & context.INI_SemanticArchiveCatalogLibraryPath
                If Not force AndAlso key = state.ConfigurationKey AndAlso System.DateTime.UtcNow < state.NextCheckUtc Then Return 0
                state.ConfigurationKey = key
                state.NextCheckUtc = System.DateTime.UtcNow.AddSeconds(SharedMethods.DEFAULT_SEMANTICARCHIVE_LIBRARY_SYNC_SECONDS)
                Return Await System.Threading.Tasks.Task.Run(Function() SynchronizeCore(context, key, cancellationToken), cancellationToken).ConfigureAwait(False)
            Finally
                state.Gate.Release()
            End Try
        End Function

        Private Shared Function SynchronizeCore(context As SharedContext.ISharedContext, configurationKey As System.String,
                    cancellationToken As System.Threading.CancellationToken) As System.Int32
            Dim directory As System.String = SemanticArchivePathGuard.RequireWindowsCompatiblePath(context.INI_SemanticArchiveCatalogLibraryPath)
            RequireLibraryDirectory(directory)
            Dim entries As New System.Collections.Generic.Dictionary(Of System.String, SemanticArchiveLibraryEntry)(System.StringComparer.Ordinal)
            Dim count As System.Int32 = 0
            ' Do not use File.Exists/Directory.Exists to interpret denied or offline storage as deletion.
            For Each path As System.String In System.IO.Directory.EnumerateFiles(directory, "*.json", System.IO.SearchOption.TopDirectoryOnly)
                cancellationToken.ThrowIfCancellationRequested()
                count += 1
                If count > SharedMethods.DEFAULT_SEMANTICARCHIVE_LIBRARY_MAX_ENTRIES Then Throw New System.IO.InvalidDataException("library_limit: The library exceeds its descriptor limit. No local subscriptions were changed.")
                Try
                    Dim entry As SemanticArchiveLibraryEntry = ReadEntry(directory, path)
                    If entries.ContainsKey(entry.EntryId) Then Throw New System.IO.InvalidDataException("Duplicate library identity.")
                    entries.Add(entry.EntryId, entry)
                Catch ex As System.OperationCanceledException
                    Throw
                Catch ex As System.Exception
                    ' Unreadable definitions never disclose their contents. Existing subscriptions
                    ' to these IDs are suspended below, not silently changed into personal archives.
                    System.Diagnostics.Trace.WriteLine("[SemanticArchiveLibrary] Descriptor skipped: " & System.IO.Path.GetFileName(path) & "; " & ex.GetType().FullName)
                End Try
            Next
            Dim store As New SemanticArchiveStore(context.INI_SemanticArchiveCatalogPathLocal)
            For attempt As System.Int32 = 1 To 3
                cancellationToken.ThrowIfCancellationRequested()
                If Not IsConfigured(context) OrElse configurationKey <> (context.INI_SemanticArchiveCatalogPathLocal & Microsoft.VisualBasic.ChrW(0) & context.INI_SemanticArchiveCatalogLibraryPath) Then
                    Throw New System.OperationCanceledException("Semantic Archive configuration changed during library synchronization.", cancellationToken)
                End If
                Dim catalog As SemanticArchiveCatalog = store.LoadCatalog()
                Dim before As System.String = Newtonsoft.Json.JsonConvert.SerializeObject(catalog)
                Dim changed As System.Int32 = MergeSubscriptions(catalog, directory, entries)
                Dim mustCreate As System.Boolean = False
                Try
                    System.IO.File.GetAttributes(store.CatalogPath)
                Catch ex As System.IO.FileNotFoundException
                    mustCreate = True
                Catch ex As System.IO.DirectoryNotFoundException
                    mustCreate = True
                End Try
                If Not mustCreate AndAlso before = Newtonsoft.Json.JsonConvert.SerializeObject(catalog) Then Return 0
                Try
                    store.SaveCatalog(catalog, catalog.Revision)
                    System.Diagnostics.Trace.WriteLine("[SemanticArchiveLibrary] Local subscriptions synchronized: " & changed.ToString(System.Globalization.CultureInfo.InvariantCulture))
                    Return System.Math.Max(1, changed)
                Catch ex As SemanticArchiveCatalogConflictException When attempt < 3
                    ' Optimistic local catalog conflict: reload without losing another host's edits.
                    System.Diagnostics.Trace.WriteLine("[SemanticArchiveLibrary] Retrying local catalog revision conflict.")
                End Try
            Next
            Throw New System.InvalidOperationException("The local catalog changed repeatedly during library synchronization.")
        End Function

        Friend Shared Function MergeSubscriptions(catalog As SemanticArchiveCatalog, directory As System.String,
                    entries As System.Collections.Generic.IDictionary(Of System.String, SemanticArchiveLibraryEntry)) As System.Int32
            Dim changed As System.Int32 = 0
            For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, SemanticArchiveLibraryEntry) In entries
                Dim entry As SemanticArchiveLibraryEntry = pair.Value
                Dim existing As SemanticArchiveDefinition = Nothing
                For Each candidate As SemanticArchiveDefinition In catalog.Archives
                    If candidate.Library IsNot Nothing AndAlso candidate.Library.EntryId = entry.EntryId AndAlso
                        System.String.Equals(candidate.Library.Directory, directory, System.StringComparison.OrdinalIgnoreCase) Then existing = candidate : Exit For
                Next
                If existing IsNot Nothing AndAlso Not IsSubscriber(existing) Then Continue For ' Publisher drafts are never overwritten by auto-sync.
                If existing Is Nothing AndAlso (entry.Withdrawn OrElse Not entry.AutoSubscribe) Then Continue For
                If existing IsNot Nothing AndAlso (entry.Revision < existing.Library.Revision OrElse entry.PublisherSid <> existing.Library.PublisherSid OrElse
                    (entry.Revision = existing.Library.Revision AndAlso PublicationHash(entry) <> existing.Library.PublicationHash)) Then
                    existing.Library.State = "revision_conflict"
                    existing.Enabled = False
                    Continue For
                End If
                If entry.Withdrawn Then
                    existing.Library.State = "withdrawn"
                    existing.Library.Revision = entry.Revision
                    existing.Library.PublicationHash = PublicationHash(entry)
                    existing.Enabled = False
                    changed += 1
                    Continue For
                End If
                Dim replacement As SemanticArchiveDefinition = SemanticArchiveMetadata.Clone(entry.Definition)
                replacement.ArchiveId = If(existing Is Nothing, SemanticArchiveIdentity.HashBytes(System.Text.Encoding.UTF8.GetBytes("library-subscription-v1|" & directory.ToUpperInvariant() & "|" & entry.EntryId)), existing.ArchiveId)
                replacement.Library = New SemanticArchiveLibraryRegistration With {.Directory = directory, .EntryId = entry.EntryId,
                    .Revision = entry.Revision, .PublisherSid = entry.PublisherSid, .DefinitionHash = DefinitionHash(entry.Definition), .PublicationHash = PublicationHash(entry),
                    .Role = "subscriber", .State = If(entry.Definition.Enabled, "available", "disabled_by_publisher"), .OptOut = existing IsNot Nothing AndAlso existing.Library.OptOut}
                replacement.Enabled = entry.Definition.Enabled AndAlso Not replacement.Library.OptOut
                If existing Is Nothing Then
                    If catalog.Archives.Exists(Function(item As SemanticArchiveDefinition) item.ArchiveId = replacement.ArchiveId) Then Throw New System.IO.InvalidDataException("A personal archive collides with a managed library identity; nothing was replaced.")
                    catalog.Archives.Add(replacement)
                Else
                    catalog.Archives(catalog.Archives.IndexOf(existing)) = replacement
                End If
                changed += 1
            Next
            For Each archive As SemanticArchiveDefinition In catalog.Archives
                If Not IsSubscriber(archive) Then Continue For
                If Not System.String.Equals(archive.Library.Directory, directory, System.StringComparison.OrdinalIgnoreCase) Then
                    archive.Library.State = "disconnected"
                    archive.Enabled = False
                ElseIf Not entries.ContainsKey(archive.Library.EntryId) Then
                    archive.Library.State = "unavailable"
                    archive.Enabled = False
                End If
            Next
            Return changed
        End Function

        ''' <summary>Library visibility is not source authority. Reopen the published
        ''' definition for each operation; then the existing source ACL/hash checks run.</summary>
        Public Shared Sub RequireCurrentDefinition(context As SharedContext.ISharedContext, definition As SemanticArchiveDefinition)
            If Not SemanticArchiveHostIntegration.IsConfigured(context) Then Throw New System.InvalidOperationException("semantic_archive_disabled: SemanticArchiveCatalogPathLocal is empty.")
            If Not IsSubscriber(definition) Then Return
            If Not CanUseLocally(definition, context.INI_SemanticArchiveCatalogLibraryPath) Then Throw New SemanticArchiveLibraryAccessException("library_subscription_unavailable", "This archive is no longer an active library subscription. Check its library status or your local subscription selection.")
            Dim entry As SemanticArchiveLibraryEntry
            Try
                entry = ReadEntry(definition.Library.Directory, EntryPath(definition.Library.Directory, definition.Library.EntryId))
            Catch ex As System.Exception
                RetrievalSourceDiscovery.RequestRefresh(context)
                Throw New SemanticArchiveLibraryAccessException("library_unavailable", "The subscribed archive's central definition is currently unavailable or its protection cannot be verified. Private data is retained, but this source will not be used until library access is restored.", ex)
            End Try
            If entry.Withdrawn OrElse Not entry.Definition.Enabled OrElse entry.Revision <> definition.Library.Revision OrElse
                entry.PublisherSid <> definition.Library.PublisherSid OrElse PublicationHash(entry) <> definition.Library.PublicationHash OrElse DefinitionHash(entry.Definition) <> definition.Library.DefinitionHash Then
                RetrievalSourceDiscovery.RequestRefresh(context)
                Throw New SemanticArchiveLibraryAccessException("library_refresh_required", "The published archive changed or was withdrawn. Automatic library synchronization has been requested; retry after it completes, or use Sync library.")
            End If
            Dim local As SemanticArchiveDefinition = SemanticArchiveMetadata.Clone(definition)
            local.ArchiveId = entry.EntryId
            local.Library = Nothing
            local.Enabled = entry.Definition.Enabled
            If DefinitionHash(local) <> DefinitionHash(entry.Definition) Then Throw New SemanticArchiveLibraryAccessException("library_definition_modified", "This subscribed definition differs from its publisher revision. Use Sync library to restore it before retrieval.")
        End Sub

        Public Shared Sub ValidateScope(context As SharedContext.ISharedContext, store As SemanticArchiveStore,
                    archiveIds As System.Collections.Generic.IEnumerable(Of System.String))
            For Each id As System.String In archiveIds
                Dim definition As SemanticArchiveDefinition = store.GetArchive(id)
                If definition Is Nothing Then Throw New System.UnauthorizedAccessException("An archive was unregistered.")
                RequireCurrentDefinition(context, definition)
            Next
        End Sub

        Public Shared Sub Publish(context As SharedContext.ISharedContext, archiveId As System.String, withdraw As System.Boolean,
                                  Optional cancellationToken As System.Threading.CancellationToken = Nothing)
            cancellationToken.ThrowIfCancellationRequested()
            If Not IsConfigured(context) Then Throw New System.InvalidOperationException("Configure SemanticArchiveCatalogPathLocal and SemanticArchiveCatalogLibraryPath in redink.ini before publishing.")
            Dim directory As System.String = SemanticArchivePathGuard.RequireWindowsCompatiblePath(context.INI_SemanticArchiveCatalogLibraryPath)
            RequireLibraryDirectory(directory)
            Dim store As New SemanticArchiveStore(context.INI_SemanticArchiveCatalogPathLocal)
            Dim catalog As SemanticArchiveCatalog = store.LoadCatalog()
            Dim archive As SemanticArchiveDefinition = catalog.Archives.Find(Function(item As SemanticArchiveDefinition) item.ArchiveId = archiveId)
            If archive Is Nothing OrElse IsSubscriber(archive) Then Throw New System.InvalidOperationException("Only a locally authored archive can be published or withdrawn by its publisher.")
            Dim registration As SemanticArchiveLibraryRegistration = archive.Library
            If registration IsNot Nothing AndAlso Not System.String.Equals(registration.Directory, directory, System.StringComparison.OrdinalIgnoreCase) Then Throw New System.InvalidOperationException("This archive was published to a different library. Restore that configured library to update or withdraw it.")
            Dim entryId As System.String = If(registration Is Nothing, archive.ArchiveId, registration.EntryId)
            Dim path As System.String = EntryPath(directory, entryId)
            Dim sid As System.String = CurrentPublisherSid()
            Dim security As System.Security.AccessControl.FileSecurity = InitialDescriptorSecurity(directory, sid)
            Dim committed As SemanticArchiveLibraryEntry = Nothing
            ' One stable, ACL-protected lock per publication; never delete another writer's lock.
            Dim lockPath As System.String = SemanticArchivePathGuard.RequireWindowsCompatiblePath(System.IO.Path.Combine(directory, entryId & ".lock"))
            RequireStableLibraryAncestors(directory, sid)
            Using publicationLock As New System.IO.FileStream(lockPath, System.IO.FileMode.OpenOrCreate,
                    System.Security.AccessControl.FileSystemRights.Read Or System.Security.AccessControl.FileSystemRights.Write,
                    System.IO.FileShare.None, 4096, System.IO.FileOptions.None, security)
                ValidateDescriptorSecurity(publicationLock.GetAccessControl(), sid, directory)
                Dim previous As SemanticArchiveLibraryEntry = Nothing
                Try
                    previous = ReadEntry(directory, path)
                Catch ex As System.IO.FileNotFoundException
                    If registration IsNot Nothing Then Throw New System.IO.IOException("The previously published definition is missing. Do not overwrite or recreate an unverified publication.", ex)
                End Try
                Dim desired As SemanticArchiveDefinition = Nothing
                If withdraw Then
                    If previous Is Nothing Then Throw New System.InvalidOperationException("This archive has not been published.")
                    desired = SemanticArchiveMetadata.Clone(previous.Definition)
                Else
                    desired = DistributedDefinition(archive, directory)
                End If
                desired.ArchiveId = entryId
                If previous IsNot Nothing Then
                    If previous.PublisherSid <> sid Then Throw New System.UnauthorizedAccessException("Only the publisher may update or withdraw this library entry.")
                    Dim identical As System.Boolean = previous.Withdrawn = withdraw AndAlso DefinitionHash(desired) = DefinitionHash(previous.Definition)
                    If identical Then
                        ' Idempotent retry, including central-commit/local-commit failure.
                        ' Adopt the already committed revision; do not publish it again.
                        committed = previous
                    ElseIf registration Is Nothing OrElse previous.Revision <> registration.Revision Then
                        Throw New System.InvalidOperationException("The library entry changed after this local draft was loaded. Restore the latest publisher catalog before publishing; another revision will not be overwritten.")
                    End If
                End If
                If committed Is Nothing Then
                    If previous IsNot Nothing AndAlso previous.Revision = System.Int64.MaxValue Then Throw New System.InvalidOperationException("Library revision exhausted.")
                    committed = New SemanticArchiveLibraryEntry With {.EntryId = entryId, .PublisherSid = sid,
                        .Revision = If(previous Is Nothing, 1L, previous.Revision + 1L), .Withdrawn = withdraw,
                        .AutoSubscribe = If(previous Is Nothing, SharedMethods.DEFAULT_SEMANTICARCHIVE_LIBRARY_AUTO_SUBSCRIBE, previous.AutoSubscribe),
                        .Definition = desired}
                    cancellationToken.ThrowIfCancellationRequested()
                    WriteEntry(directory, path, committed, previous IsNot Nothing, security)
                End If
            End Using
            archive.Library = New SemanticArchiveLibraryRegistration With {.Directory = directory, .EntryId = entryId,
                .Revision = committed.Revision, .PublisherSid = sid, .DefinitionHash = DefinitionHash(committed.Definition), .PublicationHash = PublicationHash(committed),
                .Role = "publisher", .State = If(withdraw, "withdrawn", "available")}
            Try
                store.SaveCatalog(catalog, catalog.Revision)
            Catch ex As System.Exception
                Throw New System.InvalidOperationException("The central publication committed, but updating the local publisher record failed. Reload the local catalog before retrying. Central revision: " & committed.Revision.ToString(System.Globalization.CultureInfo.InvariantCulture), ex)
            End Try
            RetrievalSourceDiscovery.RequestRefresh(context, True)
        End Sub

        Private Shared Function DistributedDefinition(archive As SemanticArchiveDefinition, directory As System.String) As SemanticArchiveDefinition
            Dim result As SemanticArchiveDefinition = SemanticArchiveMetadata.Clone(archive)
            result.Library = Nothing
            result.AdditionalConfiguration.Clear()
            If result.Roots.Count = 0 Then Throw New System.InvalidOperationException("Add a source folder before publishing.")
            For Each binding As SemanticArchiveSourceBinding In result.Roots
                binding.RootPath = SemanticArchivePathGuard.RequireWindowsSourcePath(binding.RootPath)
                ' A publisher's personal output path has no meaning on another user's machine.
                binding.ShadowArtifactRoot = SharedMethods.DEFAULT_SEMANTICARCHIVE_SHADOW_ARTIFACT_ROOT
                binding.AdditionalConfiguration.Clear()
                If directory.StartsWith("\\", System.StringComparison.Ordinal) AndAlso Not binding.RootPath.StartsWith("\\", System.StringComparison.Ordinal) Then
                    Throw New System.InvalidOperationException("A network library requires UNC source folders. Publishing does not copy originals: replace local/mapped source folders with their shared UNC paths first.")
                End If
                If Not System.String.IsNullOrWhiteSpace(binding.SharedArtifactRoot) Then
                    binding.SharedArtifactRoot = SemanticArchivePathGuard.CanonicalPath(binding.SharedArtifactRoot)
                    If directory.StartsWith("\\", System.StringComparison.Ordinal) AndAlso Not binding.SharedArtifactRoot.StartsWith("\\", System.StringComparison.Ordinal) Then Throw New System.InvalidOperationException("Shared artifact folders in a network library must also use UNC paths.")
                End If
            Next
            SemanticArchiveStore.ValidateLibraryDefinition(result)
            Return result
        End Function

        Private Shared Function PublicationHash(entry As SemanticArchiveLibraryEntry) As System.String
            Return SemanticArchiveIdentity.HashBytes(System.Text.Encoding.UTF8.GetBytes(Newtonsoft.Json.JsonConvert.SerializeObject(entry)))
        End Function

        Friend Shared Function DefinitionHash(definition As SemanticArchiveDefinition) As System.String
            Return SemanticArchiveIdentity.HashBytes(System.Text.Encoding.UTF8.GetBytes(Newtonsoft.Json.JsonConvert.SerializeObject(definition)))
        End Function

        Private Shared Function EntryPath(directory As System.String, entryId As System.String) As System.String
            Return SemanticArchivePathGuard.RequireWindowsCompatiblePath(System.IO.Path.Combine(directory, SemanticArchiveIdentity.ValidateId(entryId, NameOf(entryId)) & ".json"))
        End Function
    End Class
End Namespace
