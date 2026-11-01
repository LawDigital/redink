' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.KnowledgeStoreWiring.vb
' Purpose:
'   Hosts Knowledge Store automatic maintenance for Word. Semantic Archive automatic maintenance is Outlook-only.
'
' Responsibilities:
'   - Initialize and shut down the shared Knowledge Store idle service.
'   - Drive periodic background indexing from a WinForms timer.
'   - Register Word-specific host-idle logic with `KnowledgeStoreHostGate` so
'     gated Knowledge Store AI work waits while Word chat activity is visible.
'   - Expose the foreground Knowledge Store indexing command used by UI entry
'     points.
'
' Lifecycle:
'   - Startup: `InitializeKnowledgeStoreService()` from delayed startup.
'   - Idle: `KsTimer_Tick` drives background indexing.
'   - Shutdown: `ShutdownKnowledgeStoreService()` clears timer and gate wiring.
' =============================================================================


Option Explicit On
Option Strict On

Imports System.Diagnostics
Imports System.Threading.Tasks
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>UI activity snapshots and Word-hosted Knowledge Store maintenance.</summary>
    Private _ksTimer As System.Windows.Forms.Timer
    Private _maintenanceCoordinator As BackgroundMaintenanceCoordinator
    Private _hostIdleSnapshot As Integer = 1

    ' Catalog initialization and all recursive scans execute on workers. Only the
    ' UI timer captures Word/WinForms activity; no worker reads Office window state.
    Private _ksInitializationState As Integer = 0 ' 0=not started, 1=running, 2=initialized
    Private _ksShutdownRequested As Integer = 0

    ''' <summary>Activity snapshots every five seconds; each provider owns its processing cadence.</summary>
    Private Const KS_IDLE_INTERVAL_MS As Integer = 5000


    ''' <summary>
    ''' Idle timer tick — drives Word-hosted Knowledge Store background indexing.
    ''' </summary>
    Private Async Sub KsTimer_Tick(sender As Object, e As System.EventArgs)
        Try
            ' This is the only place that reads Office/WinForms activity. Workers see a Boolean snapshot.
            Dim hostIsIdle = IsWordIdle()
            System.Threading.Interlocked.Exchange(_hostIdleSnapshot, If(hostIsIdle, 1, 0))
            Dim coordinator = _maintenanceCoordinator
            If coordinator IsNot Nothing Then Await coordinator.OnTickAsync(hostIsIdle).ConfigureAwait(False)
        Catch ex As System.Exception
            System.Diagnostics.Debug.WriteLine("Office maintenance tick: " & ex.Message)
        End Try
    End Sub

    Private Function IsWordIdle() As Boolean
        Try
            If chatForm IsNot Nothing AndAlso
               Not chatForm.IsDisposed AndAlso
               chatForm.Visible Then
                Return False
            End If
        Catch
        End Try

        Try
            For Each openForm As System.Windows.Forms.Form In System.Windows.Forms.Application.OpenForms
                If openForm Is Nothing OrElse openForm.IsDisposed OrElse Not openForm.Visible Then
                    Continue For
                End If

                If TypeOf openForm Is frmAIChat Then
                    Return False
                End If
            Next
        Catch
        End Try

        Return True
    End Function

    Public Sub InitializeKnowledgeStoreService()
        Try
            System.Threading.Interlocked.Exchange(_hostIdleSnapshot, If(IsWordIdle(), 1, 0))
            KnowledgeStoreHostGate.RegisterHostIdleProvider("Word", Function() System.Threading.Volatile.Read(_hostIdleSnapshot) = 1)
            ' Word intentionally does not register Semantic Archive automatic providers; Outlook owns them.
            If Not OfficeMaintenanceService.IsConfigured(_context, includeSemanticArchiveProviders:=False) Then Return
            If System.Threading.Interlocked.CompareExchange(_ksInitializationState, 1, 0) <> 0 Then Return
            System.Threading.Interlocked.Exchange(_ksShutdownRequested, 0)
            Dim capturedContext = _context
            System.Threading.Tasks.Task.Run(
                Sub()
                    Try
                        Dim created = OfficeMaintenanceService.Create(capturedContext, includeSemanticArchiveProviders:=False)
                        System.Threading.Interlocked.Exchange(_maintenanceCoordinator, created)
                        System.Threading.Interlocked.Exchange(_ksInitializationState, 2)
                        If System.Threading.Volatile.Read(_ksShutdownRequested) <> 0 Then
                            Dim retired = System.Threading.Interlocked.Exchange(_maintenanceCoordinator, Nothing)
                            If retired IsNot Nothing Then retired.Dispose()
                            Return
                        End If
                        Dim uiControl = mainThreadControl
                        If uiControl Is Nothing OrElse uiControl.IsDisposed Then
                            Dim retired = System.Threading.Interlocked.Exchange(_maintenanceCoordinator, Nothing)
                            If retired IsNot Nothing Then retired.Dispose()
                            Return
                        End If
                        uiControl.BeginInvoke(New System.Windows.Forms.MethodInvoker(
                            Sub()
                                If System.Threading.Volatile.Read(_ksShutdownRequested) = 0 Then EnsureKnowledgeStoreTimerStarted()
                            End Sub))
                    Catch ex As System.Exception
                        Dim failed = System.Threading.Interlocked.Exchange(_maintenanceCoordinator, Nothing)
                        If failed IsNot Nothing Then failed.Dispose()
                        System.Threading.Interlocked.Exchange(_ksInitializationState, 0)
                        System.Diagnostics.Debug.WriteLine("Office maintenance initialization: " & ex.Message)
                    End Try
                End Sub)
        Catch ex As System.Exception
            System.Threading.Interlocked.Exchange(_ksInitializationState, 0)
            System.Diagnostics.Debug.WriteLine("Office maintenance startup: " & ex.Message)
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
            Dim retired = System.Threading.Interlocked.Exchange(_maintenanceCoordinator, Nothing)
            If retired IsNot Nothing Then retired.Dispose()
            System.Threading.Interlocked.Exchange(_ksInitializationState, 0)
        Catch ex As System.Exception
            System.Diagnostics.Debug.WriteLine("Office maintenance shutdown: " & ex.Message)
        Finally
            KnowledgeStoreHostGate.ClearHostIdleProvider()
        End Try
    End Sub

    ''' <summary>
    ''' Runs a foreground index of all (or a specific) Knowledge Store with progress bar.
    ''' Called from menu or settings UI.
    ''' </summary>
    ''' <param name="storeName">Optional store name to restrict indexing to. Empty = all stores.</param>
    ''' <param name="forceReindex">If True, re-indexes all files even if already indexed.</param>
    Public Async Function RunForegroundKnowledgeStoreIndexAsync(
            Optional storeName As String = "",
            Optional forceReindex As Boolean = False) As Task
        Try
            Dim result = Await KnowledgeStoreForegroundIndexer.RunAsync(_context, storeName, forceReindex).ConfigureAwait(False)

            Dim msg As String
            If result.WasCancelled Then
                msg = $"Indexing was cancelled. Indexed {result.IndexedFiles} of {result.TotalFiles} file(s)."
            Else
                msg = $"Indexing complete. Indexed: {result.IndexedFiles}, Skipped: {result.SkippedFiles}, Failed: {result.FailedFiles}."
            End If

            ShowCustomMessageBox(msg, $"{AN} Knowledge Store")
        Catch ex As Exception
            ShowCustomMessageBox($"Error during foreground indexing: {ex.Message}", $"{AN} Knowledge Store")
        End Try
    End Function


End Class