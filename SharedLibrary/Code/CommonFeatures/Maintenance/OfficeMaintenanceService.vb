' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: OfficeMaintenanceService.vb
' Purpose:
'   Office composition root for configured background-maintenance providers.
'
' Architecture / Function:
'   Wires providers to the generic coordinator; scheduling, cancellation and retry rules
'   remain outside this composition root.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary
    ''' <summary>Composition root only. Scheduling, retry and cancellation rules remain in the generic coordinator.</summary>
    Public NotInheritable Class OfficeMaintenanceService
        Private Sub New()
        End Sub

        Public Shared Function IsConfigured(context As SharedContext.ISharedContext,
                                            Optional includeSemanticArchiveProviders As System.Boolean = True) As Boolean
            Return context IsNot Nothing AndAlso
                (KnowledgeStoreCatalog.IsConfigured(context) OrElse
                 (includeSemanticArchiveProviders AndAlso Not System.String.IsNullOrWhiteSpace(context.INI_SemanticArchiveCatalogPathLocal)))
        End Function

        ''' <summary>Called on a worker. Semantic Archive automatic providers are hosted by Outlook only;
        ''' Word may still host Knowledge Store maintenance and all manual Semantic Archive commands.</summary>
        Public Shared Function Create(context As SharedContext.ISharedContext,
                                      Optional includeSemanticArchiveProviders As System.Boolean = True) As BackgroundMaintenanceCoordinator
            If context Is Nothing Then Throw New System.ArgumentNullException(NameOf(context))
            If includeSemanticArchiveProviders AndAlso Not System.String.IsNullOrWhiteSpace(context.INI_SemanticArchiveCatalogPathLocal) Then
                Try
                    Dim store As New SemanticArchiveStore(context.INI_SemanticArchiveCatalogPathLocal)
                    store.LoadCatalog()
                Catch ex As System.Exception
                    System.Diagnostics.Debug.WriteLine("Semantic Archive configuration: " & ex.Message)
                End Try
            End If
            Dim coordinator As New BackgroundMaintenanceCoordinator()
            Try
                If includeSemanticArchiveProviders Then
                    coordinator.Register(New BackgroundMaintenanceCoordinator.Provider With {
                        .Id = "semantic-archive-library",
                        .IsEnabled = Function() SemanticArchiveLibrary.IsConfigured(context),
                        .CanRunNow = Function() True,
                        .RunBatchAsync = Function(token As System.Threading.CancellationToken) SemanticArchiveLibrary.SynchronizeAsync(context, True, token),
                        .MinimumIntervalSeconds = SharedMethods.DEFAULT_SEMANTICARCHIVE_LIBRARY_SYNC_SECONDS,
                        .IdleIntervalSeconds = SharedMethods.DEFAULT_SEMANTICARCHIVE_LIBRARY_SYNC_SECONDS})
                    Dim archives As New SemanticArchiveMaintenanceProvider(context, coordinator)
                    coordinator.Register(archives.Registration())
                    Dim permissions As New SemanticArchivePermissionMaintenanceProvider(context, coordinator)
                    coordinator.Register(permissions.Registration())
                End If
                KnowledgeStoreIdleService.Initialize(context)
                coordinator.Register(New BackgroundMaintenanceCoordinator.Provider() With {
                    .Id = "knowledge-store",
                    .RefreshControls = Sub() KnowledgeStoreIdleService.RefreshPersistedControls(context),
                    .IsEnabled = Function() KnowledgeStoreIdleService.IsEnabled,
                    .CanRunNow = Function() KnowledgeStoreIdleService.CanRunNow(context),
                    .RunBatchAsync = Function(token) KnowledgeStoreIdleService.OnIdleTickAsync(token),
                    .MinimumIntervalSeconds = 60,
                    .Shutdown = AddressOf KnowledgeStoreIdleService.Shutdown
                })
                Return coordinator
            Catch
                coordinator.Dispose()
                Throw
            End Try
        End Function
    End Class
End Namespace
