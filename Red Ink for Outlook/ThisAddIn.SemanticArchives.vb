' Part of "Red Ink for Outlook"
' Local source selection and requester-bound unattended Semantic Archive authorization.

' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: ThisAddIn.SemanticArchives.vb
' Purpose:
'   Outlook archive selection and requester-bound Semantic Archive authorization for
'   unattended runs.
'
' Architecture / Function:
'   Separates interactive selections from independently verified AutoPilot requester
'   grants; arguments never grant source rights.
' =============================================================================

Option Strict On
Option Explicit On

Partial Public Class ThisAddIn
    ' Nothing: no session choice yet. Empty: the user explicitly opted out.
    Private _selectedSemanticArchiveIds As System.Collections.Generic.List(Of System.String) = Nothing

    Friend Function CreateSelectedSemanticArchiveRunScope() As Global.SharedLibrary.SharedLibrary.SemanticArchiveRunScope
        If Not Global.SharedLibrary.SharedLibrary.SemanticArchiveHostIntegration.IsConfigured(_context) Then Return Nothing
        If _apActive Then
            Return CreateUnverifiedAutoPilotSemanticArchiveRunScope()
        End If
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

    ' Only trusted host initialization code may install this adapter. There is no
    ' configuration, tool argument or email command that can assert verification.
    Friend Property SemanticArchiveRequesterAuthorizer As Global.SharedLibrary.SharedLibrary.ISemanticArchiveRequesterAuthorizer

    ''' <summary>
    ''' A sender address, caller-ID mapping or owner marker never grants archive
    ''' authority. Fallback tooling roots fail before reading archive defaults.
    ''' </summary>
    Private Function CreateUnverifiedAutoPilotSemanticArchiveRunScope() As Global.SharedLibrary.SharedLibrary.SemanticArchiveRunScope
        If Not Global.SharedLibrary.SharedLibrary.SemanticArchiveHostIntegration.IsConfigured(_context) Then Return Nothing
        Return New Global.SharedLibrary.SharedLibrary.SemanticArchiveRunScope(New System.String() {},
            Global.SharedLibrary.SharedLibrary.SemanticArchiveRequesterAuthorization.CreateUnverifiedAccess())
    End Function

    ''' <summary>
    ''' Requests fresh independent authority for one mail, voicemail or scheduled
    ''' execution. Scheduled CreatedBy is an untrusted hint, not reusable proof.
    ''' No adapter is installed by default; missing, failed, expired or mismatched
    ''' grants deny before catalog, source or model work. Child agents inherit only
    ''' the resulting bounded scope and cannot invoke this identity boundary.
    ''' </summary>
    Private Async Function CreateAutoPilotSemanticArchiveRunScopeAsync(
        origin As Global.SharedLibrary.SharedLibrary.SemanticArchiveRequesterOrigin,
        sourceItemId As System.String,
        claimedIdentity As System.String,
        cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of Global.SharedLibrary.SharedLibrary.SemanticArchiveRunScope)

        If Not Global.SharedLibrary.SharedLibrary.SemanticArchiveHostIntegration.IsConfigured(_context) Then Return Nothing
        cancellationToken.ThrowIfCancellationRequested()
        Dim authorizer As Global.SharedLibrary.SharedLibrary.ISemanticArchiveRequesterAuthorizer = SemanticArchiveRequesterAuthorizer
        If authorizer Is Nothing Then Return CreateUnverifiedAutoPilotSemanticArchiveRunScope()

        Dim request As New Global.SharedLibrary.SharedLibrary.SemanticArchiveRequesterRequest(origin, sourceItemId, claimedIdentity)
        Dim grant As Global.SharedLibrary.SharedLibrary.SemanticArchiveRequesterGrant = Nothing
        Try
            grant = Await authorizer.AuthorizeAsync(request, cancellationToken)
        Catch ex As System.OperationCanceledException
            Throw
        Catch ex As System.Exception
            System.Diagnostics.Trace.WriteLine("[SemanticArchive] Requester verification failed closed: " & ex.GetType().FullName)
        End Try
        cancellationToken.ThrowIfCancellationRequested()
        Dim access As Global.SharedLibrary.SharedLibrary.SemanticArchiveAccessContext =
            Global.SharedLibrary.SharedLibrary.SemanticArchiveRequesterAuthorization.BindGrant(request, grant)
        If access.DenialCode.Length > 0 Then
            Return New Global.SharedLibrary.SharedLibrary.SemanticArchiveRunScope(New System.String() {}, access)
        End If
        Return Global.SharedLibrary.SharedLibrary.SemanticArchiveHostIntegration.CreateRunScope(_context, Nothing, access)
    End Function
End Class
