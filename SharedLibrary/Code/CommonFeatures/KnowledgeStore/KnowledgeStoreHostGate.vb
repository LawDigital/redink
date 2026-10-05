' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: KnowledgeStoreHostGate.vb
' Purpose:
'   Coordinates Knowledge Store LLM execution with the current Office host state.
'
' Responsibilities:
'   - Serialize gated Knowledge Store AI phases so only one host-aware KS LLM
'     step runs at a time.
'   - Defer Knowledge Store AI work until the registered Office host reports
'     itself idle.
'   - Surface queue, wait, resume, and start status back to callers through a
'     lightweight callback.
'   - Allow Word and Outlook host wiring to register and clear host-specific
'     idle predicates.
'
' Notes:
'   - This gate is intended for Knowledge Store maintenance and ingestion work
'     that must not compete with foreground Office AI activity.
'   - In Outlook, the registered idle predicate can give AutoPilot and other
'     active LLM or tooling work precedence over Knowledge Store AI work.
' =============================================================================

Option Strict On
Option Explicit On


Imports System.Diagnostics
Imports System.Threading
Imports System.Threading.Tasks

Namespace SharedLibrary

    ''' <summary>
    ''' Serializes Knowledge Store LLM work and delays it until the current Office host is idle.
    ''' </summary>
    Public NotInheritable Class KnowledgeStoreHostGate

        Private Shared ReadOnly _llmGate As New SemaphoreSlim(1, 1)
        Private Shared ReadOnly _registrationLock As New Object()

        Private Shared _hostDisplayName As String = "Office"
        Private Shared _hostIdleProvider As Func(Of Boolean) = Function() True

        Private Sub New()
        End Sub

        Public Shared ReadOnly Property HostDisplayName As String
            Get
                SyncLock _registrationLock
                    Return _hostDisplayName
                End SyncLock
            End Get
        End Property

        Public Shared Sub RegisterHostIdleProvider(hostDisplayName As String,
                                                   hostIdleProvider As Func(Of Boolean))
            SyncLock _registrationLock
                _hostDisplayName = If(String.IsNullOrWhiteSpace(hostDisplayName), "Office", hostDisplayName.Trim())
                _hostIdleProvider = If(hostIdleProvider, Function() True)
            End SyncLock
        End Sub

        Public Shared Sub ClearHostIdleProvider()
            SyncLock _registrationLock
                _hostDisplayName = "Office"
                _hostIdleProvider = Function() True
            End SyncLock
        End Sub

        Public Shared Async Function EnterLlmExecutionAsync(operationName As String,
                                                            Optional statusCallback As Action(Of String) = Nothing,
                                                            Optional cancellationToken As CancellationToken = Nothing) As Task(Of IDisposable)

            ' Automatic callers inherit the generic coordinator cancellation checkpoint.
            If Not cancellationToken.CanBeCanceled Then cancellationToken = BackgroundMaintenanceCoordinator.CurrentCancellationToken
            cancellationToken.ThrowIfCancellationRequested()

            Dim effectiveOperationName = If(String.IsNullOrWhiteSpace(operationName),
                                            "Knowledge Store AI step",
                                            operationName.Trim())

            Report(statusCallback, $"{effectiveOperationName}: queued.")

            Await _llmGate.WaitAsync(cancellationToken).ConfigureAwait(False)

            Try
                Dim waitStartedUtc As DateTime? = Nothing

                Do While Not IsHostIdle()
                    If Not waitStartedUtc.HasValue Then
                        waitStartedUtc = DateTime.UtcNow
                    End If

                    Dim waitedSeconds = CInt(Math.Max(1, (DateTime.UtcNow - waitStartedUtc.Value).TotalSeconds))
                    Report(
                        statusCallback,
                        $"{effectiveOperationName}: waiting ({waitedSeconds}s) — {HostDisplayName} busy.")

                    Await Task.Delay(1000, cancellationToken).ConfigureAwait(False)
                Loop

                If waitStartedUtc.HasValue Then
                    Dim waitedSeconds = CInt(Math.Max(1, (DateTime.UtcNow - waitStartedUtc.Value).TotalSeconds))
                    Report(
                        statusCallback,
                        $"{effectiveOperationName}: resumed after wait ({waitedSeconds}s).")
                End If

                Report(statusCallback, $"{effectiveOperationName}: starting AI step.")
                Return New GateLease()
            Catch
                _llmGate.Release()
                Throw
            End Try
        End Function


        Private Shared Function IsHostIdle() As Boolean
            Dim idleProvider As Func(Of Boolean) = Nothing

            SyncLock _registrationLock
                idleProvider = _hostIdleProvider
            End SyncLock

            Try
                If idleProvider Is Nothing Then
                    Return True
                End If

                Return idleProvider.Invoke()
            Catch ex As Exception
                Debug.WriteLine($"KSGate: Host idle provider failed: {ex.Message}")
                Return True
            End Try
        End Function

        Private Shared Sub Report(statusCallback As Action(Of String), statusText As String)
            Debug.WriteLine("KSGate: " & statusText)

            If statusCallback Is Nothing Then
                Return
            End If

            Try
                statusCallback.Invoke(statusText)
            Catch
            End Try
        End Sub

        Private NotInheritable Class GateLease
            Implements IDisposable

            Private _disposed As Integer = 0

            Public Sub Dispose() Implements IDisposable.Dispose
                If Interlocked.Exchange(_disposed, 1) <> 0 Then
                    Return
                End If

                _llmGate.Release()
            End Sub
        End Class

    End Class

End Namespace