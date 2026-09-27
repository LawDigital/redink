' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: LlmTransportRetryPolicy.vb
' Purpose: Defines provider-agnostic retry profiles for transient LLM transport
'          failures and an async-flow scope used by orchestrated/background runs.
'
' Architecture:
'  - Interactive preserves the existing short retry budget.
'  - Unattended allows materially longer bounded backoff without coupling the
'    shared transport to a specific host, provider, model, skill, or template.
'  - Inherit resolves from the current async-flow scope and falls back to
'    Interactive when no scope is active.
' =============================================================================
Option Strict On
Option Explicit On

Namespace SharedLibrary

    Public Enum LlmTransportRetryProfile
        Inherit = 0
        Interactive = 1
        Unattended = 2
    End Enum

    Public NotInheritable Class LlmTransportRetryPolicy

        Private ReadOnly _fallbackDelayMilliseconds As System.Int32()

        Public ReadOnly Property Profile As LlmTransportRetryProfile
        Public ReadOnly Property MaxRetries As System.Int32
        Public ReadOnly Property MaxFallbackDelayMilliseconds As System.Int32
        Public ReadOnly Property MaxRetryAfterDelayMilliseconds As System.Int32
        Public ReadOnly Property MaxCumulativeDelayMilliseconds As System.Int32
        Public ReadOnly Property MaxTransientFailureWallClockMilliseconds As System.Int32

        Private Shared ReadOnly JitterLock As New System.Object()
        Private Shared ReadOnly JitterRandom As New System.Random()

        Private Sub New(profile As LlmTransportRetryProfile,
                        maxRetries As System.Int32,
                        maxFallbackDelayMilliseconds As System.Int32,
                        maxRetryAfterDelayMilliseconds As System.Int32,
                        maxCumulativeDelayMilliseconds As System.Int32,
                        maxTransientFailureWallClockMilliseconds As System.Int32,
                        fallbackDelayMilliseconds As System.Int32())
            Me.Profile = profile
            Me.MaxRetries = System.Math.Max(0, maxRetries)
            Me.MaxFallbackDelayMilliseconds = System.Math.Max(0, maxFallbackDelayMilliseconds)
            Me.MaxRetryAfterDelayMilliseconds = System.Math.Max(0, maxRetryAfterDelayMilliseconds)
            Me.MaxCumulativeDelayMilliseconds = System.Math.Max(0, maxCumulativeDelayMilliseconds)
            Me.MaxTransientFailureWallClockMilliseconds = System.Math.Max(0, maxTransientFailureWallClockMilliseconds)
            Me._fallbackDelayMilliseconds =
                If(fallbackDelayMilliseconds Is Nothing OrElse fallbackDelayMilliseconds.Length = 0,
                   New System.Int32() {0},
                   DirectCast(fallbackDelayMilliseconds.Clone(), System.Int32()))
        End Sub

        Public Shared Function ForProfile(requestedProfile As LlmTransportRetryProfile) As LlmTransportRetryPolicy
            Dim effectiveProfile As LlmTransportRetryProfile =
                LlmTransportRetryPolicyScope.ResolveProfile(requestedProfile)

            Select Case effectiveProfile
                Case LlmTransportRetryProfile.Unattended
                    ' Bounded for unattended/background orchestration: long enough to
                    ' survive temporary quota pressure, but never an endless retry loop.
                    Return New LlmTransportRetryPolicy(
                        LlmTransportRetryProfile.Unattended,
                        maxRetries:=8,
                        maxFallbackDelayMilliseconds:=300000,
                        maxRetryAfterDelayMilliseconds:=900000,
                        maxCumulativeDelayMilliseconds:=900000,
                        maxTransientFailureWallClockMilliseconds:=900000,
                        fallbackDelayMilliseconds:=New System.Int32() {
                            5000, 10000, 20000, 40000, 80000, 160000, 300000, 300000
                        })

                Case Else
                    ' Preserve the pre-existing interactive behavior exactly.
                    Return New LlmTransportRetryPolicy(
                        LlmTransportRetryProfile.Interactive,
                        maxRetries:=3,
                        maxFallbackDelayMilliseconds:=30000,
                        maxRetryAfterDelayMilliseconds:=30000,
                        maxCumulativeDelayMilliseconds:=60000,
                        maxTransientFailureWallClockMilliseconds:=60000,
                        fallbackDelayMilliseconds:=New System.Int32() {5000, 10000, 30000})
            End Select
        End Function

        Public Function GetFallbackDelayMilliseconds(retryIndex As System.Int32) As System.Int32
            Dim safeIndex As System.Int32 =
                System.Math.Max(0, System.Math.Min(retryIndex, _fallbackDelayMilliseconds.Length - 1))
            Dim baseDelay As System.Int32 = _fallbackDelayMilliseconds(safeIndex)

            ' Preserve the historical interactive timings exactly. Unattended runs add a
            ' small bounded jitter to avoid synchronized retry waves when several workers
            ' are throttled at the same time. Retry-After values are never jittered.
            If Profile <> LlmTransportRetryProfile.Unattended OrElse baseDelay <= 0 Then
                Return baseDelay
            End If

            Dim jitterRange As System.Int32 = System.Math.Max(1, baseDelay \ 10)
            Dim jitter As System.Int32
            SyncLock JitterLock
                jitter = JitterRandom.Next(-jitterRange, jitterRange + 1)
            End SyncLock

            Return System.Math.Max(0, baseDelay + jitter)
        End Function

        ''' <summary>
        ''' Returns the host-level timeout budget for one logical model turn. Interactive
        ''' retains the historic per-call timeout plus host buffer. Unattended must not be
        ''' cancelled by the host before its bounded transport retry policy can complete.
        ''' </summary>
        Public Function GetHostOperationTimeoutMilliseconds(perAttemptTimeoutMilliseconds As System.Int32,
                                                            hostBufferMilliseconds As System.Int32) As System.Int32
            Dim safePerAttempt As System.Int64 = System.Math.Max(1, perAttemptTimeoutMilliseconds)
            Dim safeBuffer As System.Int64 = System.Math.Max(0, hostBufferMilliseconds)

            If Profile <> LlmTransportRetryProfile.Unattended Then
                Return ClampToInt32(safePerAttempt + safeBuffer)
            End If

            ' Unattended host operations are bounded by the same wall-clock ceiling
            ' as transient transport recovery. This prevents the sum of backoff plus
            ' repeated per-attempt timeouts from extending beyond the unattended cap.
            Dim boundedTotal As System.Int64 =
                System.Math.Min(
                    CLng(MaxTransientFailureWallClockMilliseconds),
                    safePerAttempt + CLng(MaxCumulativeDelayMilliseconds) + safeBuffer)

            Return ClampToInt32(boundedTotal)
        End Function

        Private Shared Function ClampToInt32(value As System.Int64) As System.Int32
            If value <= 0L Then Return 0
            If value >= System.Int32.MaxValue Then Return System.Int32.MaxValue
            Return CInt(value)
        End Function

    End Class

    ''' <summary>
    ''' Async-flow retry-profile scope. This lets nested LLM calls (including tools and
    ''' sub-agents) inherit the parent run's execution character without host/provider
    ''' special cases or signature plumbing through every helper.
    ''' </summary>
    Public NotInheritable Class LlmTransportRetryPolicyScope

        Private Shared ReadOnly CurrentProfileStorage As New System.Threading.AsyncLocal(Of System.Nullable(Of LlmTransportRetryProfile))()

        Private Sub New()
        End Sub

        Public Shared Function ResolveProfile(requestedProfile As LlmTransportRetryProfile) As LlmTransportRetryProfile
            If requestedProfile <> LlmTransportRetryProfile.Inherit Then
                Return requestedProfile
            End If

            Dim scopedProfile As System.Nullable(Of LlmTransportRetryProfile) = CurrentProfileStorage.Value
            If scopedProfile.HasValue AndAlso scopedProfile.Value <> LlmTransportRetryProfile.Inherit Then
                Return scopedProfile.Value
            End If

            Return LlmTransportRetryProfile.Interactive
        End Function

        Private Shared ReadOnly CurrentFailureBudgetStorage As New System.Threading.AsyncLocal(Of TransientFailureBudgetState)()

        Public Shared Sub NoteTransientFailure(maxWallClockMilliseconds As System.Int32)
            If ResolveProfile(LlmTransportRetryProfile.Inherit) <> LlmTransportRetryProfile.Unattended Then Return
            If maxWallClockMilliseconds <= 0 Then Return

            Dim state As TransientFailureBudgetState = CurrentFailureBudgetStorage.Value
            If state Is Nothing Then
                state = New TransientFailureBudgetState()
                CurrentFailureBudgetStorage.Value = state
            End If

            state.StartIfNeeded(maxWallClockMilliseconds)
        End Sub

        Public Shared Function GetRemainingTransientFailureBudgetMilliseconds(maxWallClockMilliseconds As System.Int32) As System.Int32
            If ResolveProfile(LlmTransportRetryProfile.Inherit) <> LlmTransportRetryProfile.Unattended Then
                Return System.Math.Max(0, maxWallClockMilliseconds)
            End If

            Dim state As TransientFailureBudgetState = CurrentFailureBudgetStorage.Value
            If state Is Nothing OrElse Not state.Started Then
                Return System.Math.Max(0, maxWallClockMilliseconds)
            End If

            Return state.GetRemainingMilliseconds()
        End Function

        Public Shared Function Push(requestedProfile As LlmTransportRetryProfile) As System.IDisposable
            Dim previousProfile As System.Nullable(Of LlmTransportRetryProfile) = CurrentProfileStorage.Value
            Dim previousFailureBudget As TransientFailureBudgetState = CurrentFailureBudgetStorage.Value
            Dim effectiveProfile As LlmTransportRetryProfile = ResolveProfile(requestedProfile)
            CurrentProfileStorage.Value = effectiveProfile

            If effectiveProfile = LlmTransportRetryProfile.Unattended AndAlso CurrentFailureBudgetStorage.Value Is Nothing Then
                CurrentFailureBudgetStorage.Value = New TransientFailureBudgetState()
            End If

            Return New Scope(previousProfile, previousFailureBudget)
        End Function

        Private NotInheritable Class Scope
            Implements System.IDisposable

            Private ReadOnly _previousProfile As System.Nullable(Of LlmTransportRetryProfile)
            Private ReadOnly _previousFailureBudget As TransientFailureBudgetState
            Private _disposed As System.Boolean

            Public Sub New(previousProfile As System.Nullable(Of LlmTransportRetryProfile),
                           previousFailureBudget As TransientFailureBudgetState)
                _previousProfile = previousProfile
                _previousFailureBudget = previousFailureBudget
            End Sub

            Public Sub Dispose() Implements System.IDisposable.Dispose
                If _disposed Then Return
                _disposed = True
                CurrentProfileStorage.Value = _previousProfile
                CurrentFailureBudgetStorage.Value = _previousFailureBudget
            End Sub
        End Class

        Private NotInheritable Class TransientFailureBudgetState
            Private ReadOnly _syncRoot As New System.Object()
            Private _stopwatch As System.Diagnostics.Stopwatch
            Private _maxWallClockMilliseconds As System.Int32

            Public ReadOnly Property Started As System.Boolean
                Get
                    SyncLock _syncRoot
                        Return _stopwatch IsNot Nothing
                    End SyncLock
                End Get
            End Property

            Public Sub StartIfNeeded(maxWallClockMilliseconds As System.Int32)
                SyncLock _syncRoot
                    If _stopwatch Is Nothing Then
                        _maxWallClockMilliseconds = System.Math.Max(0, maxWallClockMilliseconds)
                        _stopwatch = System.Diagnostics.Stopwatch.StartNew()
                    End If
                End SyncLock
            End Sub

            Public Function GetRemainingMilliseconds() As System.Int32
                SyncLock _syncRoot
                    If _stopwatch Is Nothing Then Return _maxWallClockMilliseconds

                    Dim remaining As System.Int64 =
                        CLng(_maxWallClockMilliseconds) - _stopwatch.ElapsedMilliseconds
                    If remaining <= 0L Then Return 0
                    If remaining >= System.Int32.MaxValue Then Return System.Int32.MaxValue
                    Return CInt(remaining)
                End SyncLock
            End Function
        End Class

    End Class

End Namespace
