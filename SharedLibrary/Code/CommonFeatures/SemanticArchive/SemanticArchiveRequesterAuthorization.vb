' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' Trusted host extension for independently verified, request-bound remote authority.
Option Strict On
Option Explicit On

Namespace SharedLibrary

    Public Enum SemanticArchiveRequesterOrigin
        AutoPilotMail = 1
        Voicemail = 2
        ScheduledTask = 3
    End Enum

    ''' <summary>
    ''' Created afresh by the host for one execution. ClaimedIdentity is an untrusted
    ''' routing hint, never authentication evidence. A trusted adapter must verify
    ''' the original item or the scheduled task's independently established owner.
    ''' </summary>
    Public NotInheritable Class SemanticArchiveRequesterRequest
        Public ReadOnly Property RequestId As System.String
        Public ReadOnly Property Origin As SemanticArchiveRequesterOrigin
        Public ReadOnly Property SourceItemId As System.String
        Public ReadOnly Property ClaimedIdentity As System.String
        Public ReadOnly Property CreatedUtc As System.DateTimeOffset

        Public Sub New(origin As SemanticArchiveRequesterOrigin, sourceItemId As System.String, claimedIdentity As System.String)
            If Not System.Enum.IsDefined(GetType(SemanticArchiveRequesterOrigin), origin) Then Throw New System.ArgumentOutOfRangeException(NameOf(origin))
            RequestId = System.Guid.NewGuid().ToString("N")
            Me.Origin = origin
            Me.SourceItemId = If(sourceItemId, "")
            Me.ClaimedIdentity = If(claimedIdentity, "")
            CreatedUtc = System.DateTimeOffset.UtcNow
        End Sub
    End Class

    ''' <summary>
    ''' A trusted in-process host adapter is the authority issuing this grant. This
    ''' container does not authenticate mail, caller ID, an email address or a token.
    ''' Its source callback must enforce the independently verified principal's live
    ''' permissions. Never substitute the Office/service account's access policy.
    ''' </summary>
    Public NotInheritable Class SemanticArchiveRequesterGrant
        Friend ReadOnly Request As SemanticArchiveRequesterRequest
        Friend ReadOnly Access As SemanticArchiveAccessContext
        Friend ReadOnly ValidUntilUtc As System.DateTimeOffset

        Public Sub New(request As SemanticArchiveRequesterRequest, access As SemanticArchiveAccessContext, validUntilUtc As System.DateTimeOffset)
            Me.Request = request
            Me.Access = access
            Me.ValidUntilUtc = validUntilUtc
        End Sub
    End Class

    ''' <summary>
    ''' Installed only by trusted host initialization code. No adapter is installed
    ''' by default. Implementations must authenticate independently of message text,
    ''' From/Reply-To, display names, sender allowlists, owner markers or caller ID.
    ''' A scheduled execution needs fresh proof; CreatedBy alone is not proof.
    ''' </summary>
    Public Interface ISemanticArchiveRequesterAuthorizer
        Function AuthorizeAsync(request As SemanticArchiveRequesterRequest,
                                cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of SemanticArchiveRequesterGrant)
    End Interface

    Public NotInheritable Class SemanticArchiveRequesterAuthorization
        Public Const UnverifiedCode As System.String = "requester_identity_unverified"
        Public Shared ReadOnly MaximumGrantLifetime As System.TimeSpan = System.TimeSpan.FromMinutes(30)

        Private Sub New()
        End Sub

        Public Shared Function CreateUnverifiedAccess() As SemanticArchiveAccessContext
            Return New SemanticArchiveAccessContext("unverified-remote-requester", Nothing,
                UnverifiedCode,
                "Semantic Archive access is unavailable because the host has not independently verified the requester and source permissions for this execution. An email address, sender allowlist, owner marker, caller ID or scheduled CreatedBy value does not establish that authority.")
        End Function

        ''' <summary>
        ''' Accepts a grant only for this exact in-memory request object and a bounded
        ''' lifetime. It cannot be restored from a sender string or a persisted task.
        ''' Child runs inherit this returned access context and its expiry unchanged.
        ''' </summary>
        Public Shared Function BindGrant(request As SemanticArchiveRequesterRequest, grant As SemanticArchiveRequesterGrant) As SemanticArchiveAccessContext
            If request Is Nothing OrElse grant Is Nothing OrElse Not System.Object.ReferenceEquals(request, grant.Request) Then Return CreateUnverifiedAccess()
            If System.String.IsNullOrWhiteSpace(request.SourceItemId) OrElse grant.Access Is Nothing OrElse System.String.IsNullOrWhiteSpace(grant.Access.PrincipalId) Then Return CreateUnverifiedAccess()
            Dim now As System.DateTimeOffset = System.DateTimeOffset.UtcNow
            If now < request.CreatedUtc OrElse grant.ValidUntilUtc <= now OrElse grant.ValidUntilUtc > request.CreatedUtc.Add(MaximumGrantLifetime) Then Return CreateUnverifiedAccess()
            If grant.Access.DenialCode.Length > 0 Then Return grant.Access
            Return New SemanticArchiveAccessContext(grant.Access.PrincipalId,
                Function(sourcePath As System.String) As SemanticArchiveAccessDecision
                    Return If(grant.Access.CanReadSource(sourcePath), SemanticArchiveAccessDecision.Allowed, SemanticArchiveAccessDecision.Denied)
                End Function,
                validUntilUtc:=grant.ValidUntilUtc)
        End Function
    End Class
End Namespace
