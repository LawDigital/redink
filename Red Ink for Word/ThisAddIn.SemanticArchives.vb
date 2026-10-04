' Part of "Red Ink for Word"
' Explicit local archive source selection; archive tools never derive grants from arguments.
Option Strict On
Option Explicit On

Partial Public Class ThisAddIn
    ' Nothing: no session choice yet. Empty: the user explicitly opted out.
    Private _selectedSemanticArchiveIds As System.Collections.Generic.List(Of System.String) = Nothing

    Friend Function CreateSelectedSemanticArchiveRunScope() As Global.SharedLibrary.SharedLibrary.SemanticArchiveRunScope
        If Not Global.SharedLibrary.SharedLibrary.SemanticArchiveHostIntegration.IsConfigured(_context) Then Return Nothing
        Return Global.SharedLibrary.SharedLibrary.SemanticArchiveHostIntegration.CreateRunScope(
            _context, If(_selectedSemanticArchiveIds Is Nothing, Nothing, _selectedSemanticArchiveIds.ToArray()),
            Global.SharedLibrary.SharedLibrary.SemanticArchiveAccessContext.CreateForCurrentUser(), allowEnabledCatalogFallback:=True)
    End Function

    Private Sub SelectSemanticArchiveSources(Optional owner As System.Windows.Forms.IWin32Window = Nothing)
        Dim selected As Global.SharedLibrary.SharedLibrary.SemanticArchiveRunScope =
            Global.SharedLibrary.SharedLibrary.SemanticArchiveHostIntegration.ShowArchiveScopePicker(_context, CreateSelectedSemanticArchiveRunScope(), owner)
        If selected Is Nothing Then Return
        _selectedSemanticArchiveIds = New System.Collections.Generic.List(Of System.String)(selected.SelectedArchiveIds)
    End Sub
End Class
