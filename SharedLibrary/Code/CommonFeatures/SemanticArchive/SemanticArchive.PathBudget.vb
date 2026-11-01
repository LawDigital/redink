' Preflight private control paths before extraction or model work begins.

' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: SemanticArchive.PathBudget.vb
' Purpose:
'   Windows path-length budgets for archive trees and same-directory atomic write
'   temporaries.
'
' Architecture / Function:
'   Checks required suffix/temporary space before writes rather than discovering length
'   failures after publication starts.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary

    Public NotInheritable Partial Class SemanticArchiveStore

        Private Sub RequireArchiveWritePathBudget(archiveId As System.String)
            Dim archiveDirectory As System.String = GetArchiveDirectory(archiveId)
            Dim generation As System.String = New System.String("0"c, 32)
            Dim node As System.String = New System.String("0"c, 32)
            Dim document As System.String = "doc_" & New System.String("0"c, 64)
            Dim generationDirectory As System.String = GetGenerationDirectory(archiveId, generation)
            Dim targets As System.String() = {
                CatalogPath,
                System.IO.Path.Combine(_directoryPath, DataDirectoryName, "catalog.lock"),
                System.IO.Path.Combine(archiveDirectory, "writer.lock"),
                System.IO.Path.Combine(archiveDirectory, "writer-fence.json"),
                System.IO.Path.Combine(archiveDirectory, "current.json"),
                System.IO.Path.Combine(generationDirectory, "manifest.json"),
                System.IO.Path.Combine(generationDirectory, "nodes", node & ".indexed.txt"),
                System.IO.Path.Combine(generationDirectory, "documents", node & ".json"),
                System.IO.Path.Combine(archiveDirectory, "state", "documents", "00", document & ".json"),
                System.IO.Path.Combine(archiveDirectory, "state", "nodes", node & ".json"),
                System.IO.Path.Combine(GetWorkDirectory(archiveId), "items", document & ".json"),
                System.IO.Path.Combine(GetWorkDirectory(archiveId), "scan.json"),
                System.IO.Path.Combine(GetWorkDirectory(archiveId), "permissions", "directory-keys", "00", New System.String("0"c, 64) & ".json"),
                System.IO.Path.Combine(GetWorkDirectory(archiveId), "discovery", "directory-keys", "00", New System.String("0"c, 64) & ".json"),
                System.IO.Path.Combine(GetWorkDirectory(archiveId), "permissions", "retry-keys", "00", New System.String("0"c, 64) & ".json"),
                System.IO.Path.Combine(GetWorkDirectory(archiveId), "permissions", "retry-state.json"),
                System.IO.Path.Combine(GetWorkDirectory(archiveId), "permissions", "retries", "00000000000000", "0000000000000000.json")
            }
            For Each target As System.String In targets
                RequireAtomicWritePathBudget(target)
            Next
        End Sub

        Private Shared Sub RequireAtomicWritePathBudget(path As System.String)
            SemanticArchivePathGuard.RequireWindowsCompatiblePath(path)
            Dim parent As System.String = System.IO.Path.GetDirectoryName(path)
            Dim staging As System.String = System.IO.Path.Combine(parent, ".sa-" & New System.String("0"c, 32) & ".tmp")
            SemanticArchivePathGuard.RequireWindowsCompatiblePath(staging)
        End Sub

    End Class

End Namespace
