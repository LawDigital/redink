' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: AgentGate.vb
' Purpose: Global serialization gate ensuring that no two agentic runs / LLM / MCP
'          model calls that depend on shared host state execute concurrently.
'
' Architecture:
'  - Uses SemaphoreSlim(1, 1) for mutual exclusion across independent async flows.
'  - Re-entrant support via BeginOwnedScopeAsync / EndOwnedScope for complete
'    tooling runs that contain nested LLM/MCP/sub-agent calls.
'  - AsyncLocal(Of OwnerHolder) carries one mutable holder through a logical async
'    flow. Nested depth is protected by SyncLock so ownership can safely transfer
'    to a captured Task.Run execution context without a release/acquire race.
'  - MarkCurrentFlowAsOwner / UnmarkCurrentFlowAsOwner remain compatible with the
'    SubAgentRunner EnterAsync -> Mark -> ... -> Unmark -> Release pattern.
' =============================================================================

Option Strict On
Option Explicit On

Namespace Agents

    Public NotInheritable Class AgentGate

        Private Sub New()
        End Sub

        Private Shared ReadOnly _gate As New System.Threading.SemaphoreSlim(1, 1)

        Private NotInheritable Class OwnerHolder
            Public Depth As Integer
            Public GateHeld As Boolean
        End Class

        Private Shared ReadOnly _ownerHolder As New System.Threading.AsyncLocal(Of OwnerHolder)()

        Private Shared Function EnsureHolder() As OwnerHolder
            Dim holder As OwnerHolder = _ownerHolder.Value
            If holder Is Nothing Then
                holder = New OwnerHolder()
                _ownerHolder.Value = holder
            End If
            Return holder
        End Function

        Private Shared Function IsOwner() As Boolean
            Dim holder As OwnerHolder = _ownerHolder.Value
            If holder Is Nothing Then Return False

            SyncLock holder
                Return holder.Depth > 0
            End SyncLock
        End Function

        Private Shared Function GetOwnerDepth() As Integer
            Dim holder As OwnerHolder = _ownerHolder.Value
            If holder Is Nothing Then Return 0

            SyncLock holder
                Return holder.Depth
            End SyncLock
        End Function

        ''' <summary>Acquires the global gate for a single non-owned call. Honors cancellation.</summary>
        Public Shared Async Function EnterAsync(
            Optional cancellationToken As System.Threading.CancellationToken = Nothing) As System.Threading.Tasks.Task

            Dim owner As Boolean = IsOwner()
            System.Diagnostics.Debug.WriteLine(
                $"[AGENTGATE] EnterAsync ENTER owner={owner} depth={GetOwnerDepth()} busy={IsBusy} thread={System.Threading.Thread.CurrentThread.ManagedThreadId}")

            If owner Then
                System.Diagnostics.Debug.WriteLine("[AGENTGATE] EnterAsync BYPASS (owner)")
                Return
            End If

            System.Diagnostics.Debug.WriteLine("[AGENTGATE] EnterAsync WAITING")
            Await _gate.WaitAsync(cancellationToken).ConfigureAwait(False)
            System.Diagnostics.Debug.WriteLine(
                $"[AGENTGATE] EnterAsync ACQUIRED busy={IsBusy} thread={System.Threading.Thread.CurrentThread.ManagedThreadId}")
        End Function

        ''' <summary>
        ''' Releases a gate acquired by EnterAsync. If the current logical flow still
        ''' has an owned outer scope, the call was re-entrant and no release occurs.
        ''' </summary>
        Public Shared Sub Release()
            Dim holder As OwnerHolder = _ownerHolder.Value
            Dim shouldRelease As Boolean = False

            If holder Is Nothing Then
                shouldRelease = True
            Else
                SyncLock holder
                    If holder.Depth > 0 Then
                        Return
                    End If

                    ' SubAgentRunner uses EnterAsync -> Mark -> Unmark -> Release.
                    ' GateHeld remains True after Unmark reaches depth zero so this
                    ' Release call can relinquish the physical semaphore exactly once.
                    If holder.GateHeld Then
                        holder.GateHeld = False
                    End If
                    shouldRelease = True
                End SyncLock
            End If

            If Not shouldRelease Then Return

            Try
                _gate.Release()
            Catch ex As System.Exception
                ' Defensive only. Correct call pairs must never over-release.
                System.Diagnostics.Debug.WriteLine($"[AGENTGATE] Release ERROR: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' Acquires an owned run scope. Nested scopes increment a holder-local depth;
        ''' only the outermost EndOwnedScope releases the physical semaphore.
        ''' </summary>
        Public Shared Function BeginOwnedScopeAsync(
            Optional cancellationToken As System.Threading.CancellationToken = Nothing) As System.Threading.Tasks.Task

            ' IMPORTANT: establish the AsyncLocal holder synchronously in the CALLER's
            ' execution context before returning a Task. If the holder is first assigned
            ' inside an Async Function, that AsyncLocal assignment belongs to the callee's
            ' captured execution context and is restored when the awaited method returns.
            ' The caller would then hold the physical semaphore but appear to nested
            ' LLM/MCP calls as a non-owner, causing a self-deadlock in EnterAsync().
            Dim holder As OwnerHolder = EnsureHolder()

            SyncLock holder
                System.Diagnostics.Debug.WriteLine(
                    $"[AGENTGATE] BeginOwnedScopeAsync ENTER depth={holder.Depth} held={holder.GateHeld} busy={IsBusy} thread={System.Threading.Thread.CurrentThread.ManagedThreadId}")

                If holder.Depth > 0 Then
                    holder.Depth += 1
                    System.Diagnostics.Debug.WriteLine(
                        $"[AGENTGATE] BeginOwnedScopeAsync NESTED depth={holder.Depth}")
                    Return System.Threading.Tasks.Task.CompletedTask
                End If
            End SyncLock

            Return AcquireOwnedScopeAsync(holder, cancellationToken)
        End Function

        Private Shared Async Function AcquireOwnedScopeAsync(
            holder As OwnerHolder,
            cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task

            If holder Is Nothing Then
                Throw New System.InvalidOperationException("AgentGate owner holder was not initialized.")
            End If

            System.Diagnostics.Debug.WriteLine("[AGENTGATE] BeginOwnedScopeAsync WAITING")
            Await _gate.WaitAsync(cancellationToken).ConfigureAwait(False)

            SyncLock holder
                holder.GateHeld = True
                holder.Depth += 1
                System.Diagnostics.Debug.WriteLine(
                    $"[AGENTGATE] BeginOwnedScopeAsync ACQUIRED depth={holder.Depth} held={holder.GateHeld} busy={IsBusy} thread={System.Threading.Thread.CurrentThread.ManagedThreadId}")
            End SyncLock
        End Function

        Public Shared Sub EndOwnedScope()
            Dim holder As OwnerHolder = _ownerHolder.Value
            If holder Is Nothing Then
                System.Diagnostics.Debug.WriteLine("[AGENTGATE] EndOwnedScope BYPASS (no holder)")
                Return
            End If

            Dim shouldRelease As Boolean = False

            SyncLock holder
                System.Diagnostics.Debug.WriteLine(
                    $"[AGENTGATE] EndOwnedScope ENTER depth={holder.Depth} held={holder.GateHeld} busy={IsBusy} thread={System.Threading.Thread.CurrentThread.ManagedThreadId}")

                If holder.Depth <= 0 Then
                    holder.Depth = 0
                    System.Diagnostics.Debug.WriteLine("[AGENTGATE] EndOwnedScope BYPASS (not owner)")
                    Return
                End If

                holder.Depth -= 1
                If holder.Depth > 0 Then
                    System.Diagnostics.Debug.WriteLine(
                        $"[AGENTGATE] EndOwnedScope NESTED-EXIT depth={holder.Depth}")
                    Return
                End If

                If holder.GateHeld Then
                    holder.GateHeld = False
                    shouldRelease = True
                End If
            End SyncLock

            If shouldRelease Then
                Try
                    _gate.Release()
                    System.Diagnostics.Debug.WriteLine(
                        $"[AGENTGATE] EndOwnedScope RELEASED busy={IsBusy} thread={System.Threading.Thread.CurrentThread.ManagedThreadId}")
                Catch ex As System.Exception
                    System.Diagnostics.Debug.WriteLine(
                        $"[AGENTGATE] EndOwnedScope RELEASE ERROR: {ex.Message}")
                End Try
            End If
        End Sub

        ''' <summary>
        ''' Marks ownership after a successful EnterAsync. Nested callers only increase
        ''' depth; a depth-zero transition records that this holder owns the semaphore.
        ''' </summary>
        Public Shared Sub MarkCurrentFlowAsOwner()
            Dim holder As OwnerHolder = EnsureHolder()

            SyncLock holder
                If holder.Depth = 0 Then
                    holder.GateHeld = True
                End If
                holder.Depth += 1
                System.Diagnostics.Debug.WriteLine(
                    $"[AGENTGATE] MarkCurrentFlowAsOwner depth={holder.Depth} held={holder.GateHeld} busy={IsBusy} thread={System.Threading.Thread.CurrentThread.ManagedThreadId}")
            End SyncLock
        End Sub

        ''' <summary>
        ''' Removes one SubAgentRunner ownership level. It intentionally does not release
        ''' the semaphore; the caller's paired Release() performs that step only when no
        ''' outer owned scope remains.
        ''' </summary>
        Public Shared Sub UnmarkCurrentFlowAsOwner()
            Dim holder As OwnerHolder = _ownerHolder.Value
            If holder Is Nothing Then Return

            SyncLock holder
                System.Diagnostics.Debug.WriteLine(
                    $"[AGENTGATE] UnmarkCurrentFlowAsOwner ENTER depth={holder.Depth} held={holder.GateHeld} busy={IsBusy} thread={System.Threading.Thread.CurrentThread.ManagedThreadId}")

                If holder.Depth > 0 Then
                    holder.Depth -= 1
                Else
                    holder.Depth = 0
                End If

                System.Diagnostics.Debug.WriteLine(
                    $"[AGENTGATE] UnmarkCurrentFlowAsOwner EXIT depth={holder.Depth} held={holder.GateHeld} busy={IsBusy} thread={System.Threading.Thread.CurrentThread.ManagedThreadId}")
            End SyncLock
        End Sub

        ''' <summary>True if the physical gate is currently held by some caller.</summary>
        Public Shared ReadOnly Property IsBusy As Boolean
            Get
                Return _gate.CurrentCount = 0
            End Get
        End Property

    End Class

End Namespace
