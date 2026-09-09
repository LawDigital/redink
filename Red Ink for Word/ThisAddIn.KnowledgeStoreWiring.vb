' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.KnowledgeStoreWiring.vb
' Purpose:
'   Wires Knowledge Store services into the Word add-in lifecycle.
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

    ''' <summary>Timer driving background Knowledge Store indexing.</summary>
    Private _ksTimer As System.Windows.Forms.Timer

    ' KnowledgeStoreIdleService.Initialize may perform a full recursive store scan when
    ' background indexing is enabled. Keep that work off Word's STA/UI thread.
    Private _ksInitializationState As Integer = 0 ' 0=not started, 1=running, 2=initialized
    Private _ksShutdownRequested As Integer = 0

    ''' <summary>Interval between idle ticks in milliseconds (60 seconds).</summary>
    Private Const KS_IDLE_INTERVAL_MS As Integer = 60000


    ''' <summary>
    ''' Idle timer tick — drives background indexing.
    ''' </summary>
    Private Async Sub KsTimer_Tick(sender As Object, e As EventArgs)
        Try
            If Not KnowledgeStoreIdleService.CanRunNow(_context) Then
                Debug.WriteLine("KS Wiring: Skipping tick — outside configured Knowledge Store processing window.")
                Return
            End If

            Debug.WriteLine("KS Wiring: Timer tick fired (60s).")
            Await KnowledgeStoreIdleService.OnIdleTickAsync().ConfigureAwait(False)
        Catch
            ' Never let timer callbacks crash the host
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
            KnowledgeStoreHostGate.RegisterHostIdleProvider("Word", Function() IsWordIdle())

            If Not KnowledgeStoreCatalog.IsConfigured(_context) Then Return
            If System.Threading.Interlocked.CompareExchange(_ksInitializationState, 1, 0) <> 0 Then Return

            System.Threading.Interlocked.Exchange(_ksShutdownRequested, 0)
            Dim capturedContext = _context

            ' IMPORTANT: Initialize() can synchronously call KnowledgeStoreWatcher.RunPeriodicScan(),
            ' which recursively enumerates all configured store directories. On network shares or
            ' large stores that can take seconds. It has no Word-COM/WinForms dependency, so perform
            ' the service/watcher initialization on a worker and marshal only the WinForms timer
            ' creation back to Word's UI thread.
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

            If System.Threading.Volatile.Read(_ksInitializationState) = 1 Then
                ' Do not make Word shutdown wait for a potentially long recursive startup scan.
                ' Shutdown() takes the same service lock and will run as soon as initialization
                ' releases it. The shutdown flag also prevents the UI timer from being created.
                System.Threading.Tasks.Task.Run(
                    Sub()
                        Try
                            KnowledgeStoreIdleService.Shutdown()
                        Catch ex As System.Exception
                        End Try
                    End Sub)
            Else
                KnowledgeStoreIdleService.Shutdown()
            End If
        Catch
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