' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.KnowledgeStoreWiring.vb
' Purpose:
'   Wires Knowledge Store services into the Outlook add-in lifecycle.
'
' Responsibilities:
'   - Initialize and shut down the shared Knowledge Store idle service.
'   - Drive periodic background indexing only while Outlook is genuinely idle.
'   - Register Outlook-specific host-idle logic with `KnowledgeStoreHostGate`
'     so gated Knowledge Store AI work yields to AutoPilot, chat, tooling, and
'     other active Outlook activity.
'   - Prevent background Knowledge Store work from competing with higher-priority
'     foreground automation on the same host.
'
' Lifecycle:
'   - Startup: `InitializeKnowledgeStoreService()` from delayed startup.
'   - Idle: `KsTimer_Tick` drives background indexing when Outlook is idle.
'   - Shutdown: `ShutdownKnowledgeStoreService()` clears timer and gate wiring.
' =============================================================================


Option Explicit On
Option Strict On

Imports System.Diagnostics
Imports System.Threading.Tasks
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    Private _ksTimer As System.Windows.Forms.Timer
    Private _ksInitializationState As Integer = 0 ' 0=not started, 1=running, 2=initialized
    Private _ksShutdownRequested As Integer = 0
    Private Const KS_IDLE_INTERVAL_MS As Integer = 60000

    Public Sub InitializeKnowledgeStoreService()
        Try
            KnowledgeStoreHostGate.RegisterHostIdleProvider("Outlook", Function() IsOutlookIdle())

            If Not KnowledgeStoreCatalog.IsConfigured(_context) Then Return
            If System.Threading.Interlocked.CompareExchange(_ksInitializationState, 1, 0) <> 0 Then Return

            System.Threading.Interlocked.Exchange(_ksShutdownRequested, 0)
            Dim capturedContext = _context

            ' KnowledgeStoreIdleService.Initialize can synchronously enumerate configured stores.
            ' It has no Outlook COM/WinForms dependency, so keep that I/O off Outlook's UI thread.
            System.Threading.Tasks.Task.Run(
                Sub()
                    Try
                        KnowledgeStoreIdleService.Initialize(capturedContext)
                        System.Threading.Interlocked.Exchange(_ksInitializationState, 2)

                        If System.Threading.Volatile.Read(_ksShutdownRequested) <> 0 Then
                            KnowledgeStoreIdleService.Shutdown()
                            Return
                        End If

                        Dim uiControl = mainThreadControl
                        If uiControl Is Nothing OrElse uiControl.IsDisposed Then Return

                        uiControl.BeginInvoke(
                            New System.Windows.Forms.MethodInvoker(
                                Sub()
                                    If System.Threading.Volatile.Read(_ksShutdownRequested) <> 0 Then Return
                                    EnsureKnowledgeStoreTimerStarted()
                                End Sub))
                    Catch ex As System.Exception
                        System.Threading.Interlocked.Exchange(_ksInitializationState, 0)
                        Debug.WriteLine($"KS Wiring: Background init error: {ex.Message}")
                    End Try
                End Sub)
        Catch ex As System.Exception
            System.Threading.Interlocked.Exchange(_ksInitializationState, 0)
            Debug.WriteLine($"KS Wiring: Init scheduling error: {ex.Message}")
        End Try
    End Sub

    Private Sub EnsureKnowledgeStoreTimerStarted()
        If System.Threading.Volatile.Read(_ksShutdownRequested) <> 0 OrElse _ksTimer IsNot Nothing Then Return

        _ksTimer = New System.Windows.Forms.Timer()
        _ksTimer.Interval = KS_IDLE_INTERVAL_MS
        AddHandler _ksTimer.Tick, AddressOf KsTimer_Tick
        _ksTimer.Start()
    End Sub

    Public Sub ShutdownKnowledgeStoreService()
        System.Threading.Interlocked.Exchange(_ksShutdownRequested, 1)

        Try
            If _ksTimer IsNot Nothing Then
                _ksTimer.Stop()
                RemoveHandler _ksTimer.Tick, AddressOf KsTimer_Tick
                _ksTimer.Dispose()
                _ksTimer = Nothing
            End If

            KnowledgeStoreIdleService.Shutdown()
            System.Threading.Interlocked.Exchange(_ksInitializationState, 0)
        Catch
        Finally
            KnowledgeStoreHostGate.ClearHostIdleProvider()
        End Try
    End Sub

    ''' <summary>
    ''' Returns True when AutoPilot is currently doing real work that should block
    ''' Knowledge Store background processing.
    ''' </summary>
    Private Function IsAutoPilotBusy() As Boolean
        If Not _apActive Then
            Return False
        End If

        If Not String.IsNullOrWhiteSpace(_apCurrentProcessingEntryId) Then
            Return True
        End If

        If _apMailQueue.Count > 0 Then
            Return True
        End If

        If System.Threading.Interlocked.CompareExchange(_apSchedulerCheckRunning, 0, 0) <> 0 Then
            Return True
        End If

        Return False
    End Function

    ''' <summary>
    ''' Returns True when the host is genuinely idle — no active AutoPilot mail/voicemail/scheduler
    ''' work, no chat LLM jobs, no chat agent execution, and no power transitions.
    ''' </summary>
    Private Function IsOutlookIdle() As Boolean
        If System.Threading.Interlocked.CompareExchange(powerChanging, 0, 0) <> 0 Then
            Return False
        End If

        If IsAutoPilotBusy() Then
            Return False
        End If

        If _chatAgentActive Then
            Return False
        End If

        If System.Threading.Interlocked.CompareExchange(activeJobs, 0, 0) > 0 Then
            Return False
        End If

        If _activeToolingContext IsNot Nothing Then
            Return False
        End If

        Return True
    End Function

    Private Async Sub KsTimer_Tick(sender As Object, e As EventArgs)
        Try
            If Not KnowledgeStoreIdleService.CanRunNow(_context) Then
                Debug.WriteLine("KS Wiring: Skipping tick — outside configured Knowledge Store processing window.")
                Return
            End If

            If Not IsOutlookIdle() Then
                Debug.WriteLine("KS Wiring: Skipping tick — Outlook is busy (active AutoPilot work, Chat, or Agent).")
                Return
            End If

            Debug.WriteLine("KS Wiring: Timer tick fired (60s).")
            Await KnowledgeStoreIdleService.OnIdleTickAsync().ConfigureAwait(False)
        Catch
            ' Never let timer callbacks crash the host
        End Try
    End Sub



End Class