' Part of "Red Ink" (SharedLibrary)
' Cached source-menu snapshots. Library provisioning is a separate background service.
' No model calls, document traversal or UI-thread I/O during menu enumeration.

' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: RetrievalSourceDiscovery.vb
' Purpose:
'   Independent cached Knowledge Store/Semantic Archive snapshots for responsive source
'   menus.
'
' Architecture / Function:
'   Loads descriptors in background without menu-time model calls, document traversal or
'   UI-thread source I/O; library provisioning is separate.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary
    Public NotInheritable Class RetrievalSourceDiscovery
        Private Sub New()
        End Sub

        Private NotInheritable Class SnapshotState
            Friend ReadOnly Gate As New System.Object()
            Friend Key As System.String = System.String.Empty
            Friend LoadedUtc As System.DateTime = System.DateTime.MinValue
            Friend Loading As System.Boolean
            Friend Entries As New System.Collections.Generic.List(Of SharedMethods.FreestylePromptSource)()
        End Class

        Private Shared ReadOnly States As New System.Runtime.CompilerServices.ConditionalWeakTable(Of SharedContext.ISharedContext, SnapshotState)()
        Private Shared ReadOnly KnowledgeStates As New System.Runtime.CompilerServices.ConditionalWeakTable(Of SharedContext.ISharedContext, SnapshotState)()

        Public Shared Function CreateMenuProvider(context As SharedContext.ISharedContext) As System.Func(Of System.Collections.Generic.IReadOnlyList(Of SharedMethods.FreestylePromptSource))
            RequestRefresh(context)
            Return Function() GetSnapshot(context)
        End Function

        ' Independent single-flight caches: a stalled KB share never blocks SA updates.
        Public Shared Sub RequestRefresh(context As SharedContext.ISharedContext, Optional force As System.Boolean = False)
            If context Is Nothing Then Return
            Dim archivePath As System.String = If(context.INI_SemanticArchiveCatalogPathLocal, System.String.Empty)
            Dim libraryPath As System.String = If(context.INI_SemanticArchiveCatalogLibraryPath, System.String.Empty)
            Dim centralKb As System.String = If(context.INI_KnowledgeStorePath, System.String.Empty)
            Dim localKb As System.String = If(context.INI_KnowledgeStorePathLocal, System.String.Empty)
            Dim knowledgeOwner As System.String = If(context.INI_KnowledgeStoreOwner, System.String.Empty)
            Dim archiveState As SnapshotState = States.GetValue(context, Function(unused As SharedContext.ISharedContext) New SnapshotState())
            If SemanticArchiveHostIntegration.IsConfigured(context) Then
                QueueSnapshot(archiveState, archivePath & Microsoft.VisualBasic.ChrW(0) & libraryPath,
                    Sub(entries As System.Collections.Generic.List(Of SharedMethods.FreestylePromptSource)) LoadArchives(context, archivePath, libraryPath, entries), force)
            Else
                SyncLock archiveState.Gate
                    archiveState.Key = System.String.Empty
                    archiveState.Entries.Clear()
                    archiveState.LoadedUtc = System.DateTime.UtcNow
                End SyncLock
            End If
            QueueSnapshot(KnowledgeStates.GetValue(context, Function(unused As SharedContext.ISharedContext) New SnapshotState()),
                centralKb & Microsoft.VisualBasic.ChrW(0) & localKb & Microsoft.VisualBasic.ChrW(0) & knowledgeOwner,
                Sub(entries As System.Collections.Generic.List(Of SharedMethods.FreestylePromptSource)) LoadKnowledgeStores(centralKb, localKb, knowledgeOwner, entries), force)
        End Sub

        Private Shared Sub QueueSnapshot(state As SnapshotState, key As System.String,
                readSnapshot As System.Action(Of System.Collections.Generic.List(Of SharedMethods.FreestylePromptSource)), force As System.Boolean)
            SyncLock state.Gate
                If state.Key <> key Then
                    state.Key = key
                    state.LoadedUtc = System.DateTime.MinValue
                    state.Entries.Clear()
                End If
                If state.Loading OrElse (Not force AndAlso (System.DateTime.UtcNow - state.LoadedUtc).TotalSeconds < SharedMethods.DEFAULT_RETRIEVAL_SOURCE_MENU_REFRESH_SECONDS) Then Return
                state.Loading = True
            End SyncLock
            Dim ignored As System.Threading.Tasks.Task = System.Threading.Tasks.Task.Run(
                Sub()
                    Dim entries As New System.Collections.Generic.List(Of SharedMethods.FreestylePromptSource)()
                    Try
                        readSnapshot.Invoke(entries)
                    Catch ex As System.Exception
                        System.Diagnostics.Trace.WriteLine("[RetrievalSources] Snapshot read failed: " & ex.GetType().FullName)
                    Finally
                        SyncLock state.Gate
                            If state.Key = key Then
                                state.Entries = entries
                                state.LoadedUtc = System.DateTime.UtcNow
                            End If
                            state.Loading = False
                        End SyncLock
                    End Try
                End Sub)
        End Sub

        Public Shared Function GetSnapshot(context As SharedContext.ISharedContext) As System.Collections.Generic.IReadOnlyList(Of SharedMethods.FreestylePromptSource)
            Dim result As New System.Collections.Generic.List(Of SharedMethods.FreestylePromptSource)()
            If context Is Nothing Then Return result.AsReadOnly()
            RequestRefresh(context) ' Schedules only; never waits for a file or worker.
            If SemanticArchiveHostIntegration.IsConfigured(context) Then AppendSnapshot(States.GetValue(context, Function(unused As SharedContext.ISharedContext) New SnapshotState()), "Semantic Archives", result)
            AppendSnapshot(KnowledgeStates.GetValue(context, Function(unused As SharedContext.ISharedContext) New SnapshotState()), "Knowledge Stores", result)
            If result.Count = 0 Then result.Add(New SharedMethods.FreestylePromptSource With {.Group = "Sources", .Caption = "No enabled sources are configured."})
            Return result.AsReadOnly()
        End Function

        Private Shared Sub AppendSnapshot(state As SnapshotState, group As System.String,
                result As System.Collections.Generic.List(Of SharedMethods.FreestylePromptSource))
            SyncLock state.Gate
                For Each entry As SharedMethods.FreestylePromptSource In state.Entries
                    result.Add(New SharedMethods.FreestylePromptSource With {
                        .Group = entry.Group, .Caption = entry.Caption, .Description = entry.Description, .InsertText = entry.InsertText})
                Next
                If state.Entries.Count = 0 AndAlso state.Loading Then
                    result.Add(New SharedMethods.FreestylePromptSource With {.Group = group, .Caption = "Loading source overview; open this menu again shortly."})
                End If
            End SyncLock
        End Sub

        Private Shared Sub LoadArchives(context As SharedContext.ISharedContext, path As System.String, libraryPath As System.String, entries As System.Collections.Generic.List(Of SharedMethods.FreestylePromptSource))
            If System.String.IsNullOrWhiteSpace(path) Then Return
            Try
                Try
                    SemanticArchiveLibrary.SynchronizeAsync(context).GetAwaiter().GetResult()
                Catch ex As System.Exception
                    System.Diagnostics.Trace.WriteLine("[RetrievalSources] Library synchronization unavailable: " & ex.GetType().FullName)
                End Try
                Dim catalog As SemanticArchiveCatalog = SemanticArchiveLibrary.FilterCatalog(SemanticArchiveStore.ReadCatalogOverview(path), libraryPath)
                Dim count As System.Int32 = 0
                For Each archive As SemanticArchiveDefinition In catalog.Archives
                    If archive Is Nothing OrElse Not archive.Enabled Then Continue For
                    Try
                        SemanticArchiveLibrary.RequireCurrentDefinition(context, archive)
                    Catch ex As System.Exception
                        Continue For
                    End Try
                    If count >= SharedMethods.DEFAULT_RETRIEVAL_SOURCE_MENU_MAX_ENTRIES Then
                        entries.Add(New SharedMethods.FreestylePromptSource With {.Group = "Semantic Archives", .Caption = "More archives are available in Semantic Archives administration."})
                        Exit For
                    End If
                    SemanticArchiveIdentity.ValidateId(archive.ArchiveId, "archiveId")
                    Dim selector As System.String = archive.Name
                    Dim duplicates As System.Int32 = 0
                    For Each candidate As SemanticArchiveDefinition In catalog.Archives
                        If candidate.Enabled AndAlso System.String.Equals(candidate.Name, archive.Name, System.StringComparison.OrdinalIgnoreCase) Then duplicates += 1
                    Next
                    If duplicates <> 1 OrElse selector.IndexOfAny(New System.Char() {Microsoft.VisualBasic.ChrW(34), "\"c, Microsoft.VisualBasic.ChrW(10), Microsoft.VisualBasic.ChrW(13)}) >= 0 Then selector = archive.ArchiveId
                    entries.Add(New SharedMethods.FreestylePromptSource With {
                        .Group = "Semantic Archives", .Caption = Bounded(archive.Name, SharedMethods.DEFAULT_RETRIEVAL_SOURCE_MENU_NAME_CHARACTERS),
                        .Description = Bounded(archive.Description, SharedMethods.DEFAULT_RETRIEVAL_SOURCE_MENU_DESCRIPTION_CHARACTERS) & System.Environment.NewLine & "Insert this archive as the requested source. Current source access and available index contents are checked when the request runs.",
                        .InsertText = "(sa: archive:""" & selector & """)"})
                    count += 1
                Next
                If count = 0 Then entries.Add(New SharedMethods.FreestylePromptSource With {.Group = "Semantic Archives", .Caption = "No enabled archives in the personal catalog."})
            Catch ex As System.Exception
                entries.Add(New SharedMethods.FreestylePromptSource With {.Group = "Semantic Archives", .Caption = "Catalog unavailable; check Semantic Archives administration."})
                System.Diagnostics.Trace.WriteLine("[RetrievalSources] SA descriptors unavailable: " & ex.GetType().FullName)
            End Try
        End Sub

        Private Shared Sub LoadKnowledgeStores(centralPath As System.String, localPath As System.String, knowledgeOwner As System.String,
                    entries As System.Collections.Generic.List(Of SharedMethods.FreestylePromptSource))
            If System.String.IsNullOrWhiteSpace(centralPath) AndAlso System.String.IsNullOrWhiteSpace(localPath) Then Return
            Try
                ' Use the existing read-only resolver, including direct-folder catalogs.
                ' This is a detached configuration snapshot, never the live UI context.
                Dim context As New SharedContext With {.INI_KnowledgeStorePath = centralPath, .INI_KnowledgeStorePathLocal = localPath, .INI_KnowledgeStoreOwner = knowledgeOwner}
                Dim stores As System.Collections.Generic.List(Of KnowledgeStoreCatalog.KnowledgeStoreDefinition) = KnowledgeStoreCatalog.LoadOverviewReadOnly(context)
                Dim count As System.Int32 = 0
                For Each store As KnowledgeStoreCatalog.KnowledgeStoreDefinition In stores
                    If store Is Nothing OrElse Not store.Active Then Continue For
                    If count >= SharedMethods.DEFAULT_RETRIEVAL_SOURCE_MENU_MAX_ENTRIES Then
                        entries.Add(New SharedMethods.FreestylePromptSource With {.Group = "Knowledge Stores", .Caption = "More stores are available in Knowledge Store administration."})
                        Exit For
                    End If
                    Dim name As System.String = If(store.Name, System.String.Empty)
                    Dim canInsert As System.Boolean = Not System.String.IsNullOrWhiteSpace(name) AndAlso name.IndexOfAny(New System.Char() {Microsoft.VisualBasic.ChrW(34), ")"c, "("c, Microsoft.VisualBasic.ChrW(10), Microsoft.VisualBasic.ChrW(13)}) < 0
                    entries.Add(New SharedMethods.FreestylePromptSource With {
                        .Group = "Knowledge Stores", .Caption = Bounded(KnowledgeStoreCatalog.GetDisplayLabel(store), SharedMethods.DEFAULT_RETRIEVAL_SOURCE_MENU_NAME_CHARACTERS),
                        .Description = "Configured Knowledge Store: " & Bounded(name, SharedMethods.DEFAULT_RETRIEVAL_SOURCE_MENU_DESCRIPTION_CHARACTERS) &
                            System.Environment.NewLine & "Inserts a store-name filter; stores sharing this name follow the existing KB selection rules. Availability is checked when the request runs.",
                        .InsertText = If(canInsert, "(kb: store:""" & name & """)", System.String.Empty)})
                    count += 1
                Next
                If count = 0 Then entries.Add(New SharedMethods.FreestylePromptSource With {.Group = "Knowledge Stores", .Caption = "No enabled stores are available in the configured catalogs."})
            Catch ex As System.Exception
                entries.Add(New SharedMethods.FreestylePromptSource With {.Group = "Knowledge Stores", .Caption = "A catalog is unavailable; check Knowledge Store settings."})
                System.Diagnostics.Trace.WriteLine("[RetrievalSources] KB descriptors unavailable: " & ex.GetType().FullName)
            End Try
        End Sub

        Private Shared Function Bounded(value As System.String, maximum As System.Int32) As System.String
            Dim text As System.String = If(value, System.String.Empty).Replace(Microsoft.VisualBasic.ChrW(13), " "c).Replace(Microsoft.VisualBasic.ChrW(10), " "c)
            If text.Length <= maximum Then Return text
            Dim length As System.Int32 = maximum
            If System.Char.IsHighSurrogate(text(length - 1)) Then length -= 1
            Return text.Substring(0, length) & "..."
        End Function
    End Class
End Namespace
