' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.WebExtension.ListenerAndPower.vb
' Purpose: Hosts a local HttpListener for the web extension, manages suspend/resume
'          transitions, listener lifecycle, watchdog monitoring, job cancellation,
'          and recovery of stalled operations.
'
' Architecture:
' - HttpListener loop: restart-safe via a generation counter (listenerGeneration).
' - Power handling: hidden message-only window (PowerNotificationWindow) receives
'   WM_POWERBROADCAST suspend/resume messages and triggers async handlers.
' - Synchronization: SemaphoreSlim (suspendResumeGate) serializes suspend/resume sequences.
' - Cancellation: Active jobs and in-flight LLM operations are cancelled before suspend.
' - Watchdog: Timer periodically evaluates listener and request/job activity, restarting
'   listener if it is dead or faulted.
' - Recovery: Performs stuck job cancellation, listener reset (if inactive), and COM/OLE
'   message filter re-registration.
' - Resume cooldown: Prevents premature watchdog actions immediately after system resume.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Net
Imports System.Threading.Tasks

Partial Public Class ThisAddIn

    Private httpListener As HttpListener
    Private isShuttingDown As Boolean = False
    Private listenerTask As System.Threading.Tasks.Task

    Private llmOperationCts As System.Threading.CancellationTokenSource

    Private activeRequests As Integer = 0
    Private ModelTimeout As Integer = 300

    ' Threading & listener state (class level)
    Private llmSyncContext As System.Threading.SynchronizationContext
    Private llmThread As System.Threading.Thread
    Private cts As System.Threading.CancellationTokenSource

    Private wasListenerRunningBeforeSleep As System.Boolean = False
    Private wasLlmThreadAliveBeforeSleep As System.Boolean = False
    Private restartingAfterResume As System.Int32 = 0  ' 0/1 via Interlocked

    ' Guard to mute watchdog and concurrent restarts during power transitions
    Private powerChanging As System.Int32 = 0

    ' Generation protection (pre/post sleep)
    Private listenerGeneration As System.Int64 = 0

    ' Progress watchdog
    Private lastListenerProgressUtc As System.DateTime = System.DateTime.UtcNow
    Private watchdog As System.Threading.Timer

    ' Power notifications via hidden window
    Private powerWindow As PowerNotificationWindow

    ' Resume cooldown boundary
    Private resumeCooldownUntilUtc As System.DateTime = System.DateTime.MinValue
    Private inMainMenu As System.Int32 = 0

    ''' <summary>
    ''' Returns true if resume cooldown interval has not elapsed yet.
    ''' </summary>
    Private Function IsInResumeCooldown() As System.Boolean
        Return System.DateTime.UtcNow < resumeCooldownUntilUtc
    End Function

    ''' <summary>
    ''' Message-only window receiving WM_POWERBROADCAST for suspend/resume events.
    ''' </summary>
    Private NotInheritable Class PowerNotificationWindow
        Inherits System.Windows.Forms.NativeWindow
        Implements System.IDisposable

        Private Const WM_POWERBROADCAST As System.Int32 = &H218
        Private Const PBT_APMSUSPEND As System.Int32 = &H4
        Private Const PBT_APMRESUMEAUTOMATIC As System.Int32 = &H12
        Private Const PBT_APMRESUMESUSPEND As System.Int32 = &H7

        Private ReadOnly owner As ThisAddIn

        ''' <summary>
        ''' Initializes the message-only window handle used to receive WM_POWERBROADCAST notifications.
        ''' </summary>
        Public Sub New(ByVal owner As ThisAddIn)
            Me.owner = owner
            Dim cp As New System.Windows.Forms.CreateParams()
            cp.Caption = "InkyPowerWnd"
            cp.X = 0 : cp.Y = 0 : cp.Height = 0 : cp.Width = 0
            cp.Style = 0 : cp.ExStyle = 0
            ' Important: message-only window (HWND_MESSAGE)
            cp.Parent = New System.IntPtr(-3) ' HWND_MESSAGE
            Me.CreateHandle(cp)
        End Sub

        ''' <summary>
        ''' Initializes the message-only window handle used to receive WM_POWERBROADCAST notifications.
        ''' </summary>
        Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
            Const WM_POWERBROADCAST As System.Int32 = &H218
            Const PBT_APMQUERYSUSPEND As System.Int32 = &H0
            Const PBT_APMSUSPEND As System.Int32 = &H4
            Const PBT_APMRESUMEAUTOMATIC As System.Int32 = &H12
            Const PBT_APMRESUMESUSPEND As System.Int32 = &H7

            If m.Msg = WM_POWERBROADCAST Then
                Dim wp As System.Int32 = m.WParam.ToInt32()
                Select Case wp
                    Case PBT_APMQUERYSUSPEND
                        ' Immediately acknowledge and do nothing.
                        m.Result = New System.IntPtr(1)
                        Return

                    Case PBT_APMSUSPEND
                        System.Threading.ThreadPool.QueueUserWorkItem(
                            Sub(stateObj As System.Object)
                                Try : owner.HandlePowerSuspendAsync() : Catch : End Try
                            End Sub)
                        m.Result = New System.IntPtr(1)
                        Return

                    Case PBT_APMRESUMEAUTOMATIC
                        System.Threading.ThreadPool.QueueUserWorkItem(
                            Sub(stateObj As System.Object)
                                Try : owner.HandlePowerResumeAsync(userPresent:=False) : Catch : End Try
                            End Sub)
                        m.Result = New System.IntPtr(1)
                        Return

                    Case PBT_APMRESUMESUSPEND
                        System.Threading.ThreadPool.QueueUserWorkItem(
                            Sub(stateObj As System.Object)
                                Try : owner.HandlePowerResumeAsync(userPresent:=True) : Catch : End Try
                            End Sub)
                        m.Result = New System.IntPtr(1)
                        Return
                End Select
            End If

            MyBase.WndProc(m)
        End Sub

        ''' <summary>
        ''' Releases the native window handle backing the power notification window.
        ''' </summary>
        Public Sub Dispose() Implements System.IDisposable.Dispose
            If Me.Handle <> System.IntPtr.Zero Then
                Me.DestroyHandle()
            End If
        End Sub
    End Class

    ' Ensure only one suspend/resume sequence runs at a time
    Private suspendResumeGate As New System.Threading.SemaphoreSlim(1, 1)

    ''' <summary>
    ''' Asynchronously handles power suspend: cancels jobs, records state, stops listener, signals LLM thread exit.
    ''' </summary>
    Friend Sub HandlePowerSuspendAsync()
        System.Threading.Tasks.Task.Run(
            Async Function() As System.Threading.Tasks.Task
                If Not Await TryEnterGateAsync().ConfigureAwait(False) Then Return
                System.Threading.Interlocked.Exchange(powerChanging, 1)
                Try
                    ' Cancel all jobs FIRST and wait for them briefly
                    Dim cancelTasks As New List(Of System.Threading.Tasks.Task)()
                    For Each kv In jobMap
                        Try
                            kv.Value.Cts.Cancel()
                            ' Add a wait for the task to complete
                            If kv.Value.Tcs IsNot Nothing AndAlso Not kv.Value.Tcs.Task.IsCompleted Then
                                cancelTasks.Add(kv.Value.Tcs.Task.ContinueWith(Sub(t)
                                                                                   ' Swallow exceptions
                                                                               End Sub, TaskContinuationOptions.ExecuteSynchronously))
                            End If
                        Catch
                        End Try
                    Next

                    ' Wait briefly for jobs to terminate
                    If cancelTasks.Count > 0 Then
                        Try
                            Await System.Threading.Tasks.Task.WhenAny(
                                System.Threading.Tasks.Task.WhenAll(cancelTasks),
                                System.Threading.Tasks.Task.Delay(2000)
                            ).ConfigureAwait(False)
                        Catch
                            ' Ignore exceptions from cancelled tasks
                        End Try
                    End If

                    ' Clear the job map after cancellation
                    jobMap.Clear()
                    System.Threading.Interlocked.Exchange(activeJobs, 0)

                    ' Mute watchdog during suspend
                    Try : StopListenerWatchdog() : Catch : End Try

                    ' Remember state
                    Try
                        wasListenerRunningBeforeSleep =
                            (httpListener IsNot Nothing AndAlso httpListener.IsListening)
                    Catch
                        wasListenerRunningBeforeSleep = False
                    End Try
                    Try
                        wasLlmThreadAliveBeforeSleep =
                            (llmThread IsNot Nothing AndAlso llmThread.IsAlive)
                    Catch
                        wasLlmThreadAliveBeforeSleep = False
                    End Try

                    ' Proactively cancel any in-flight LLM op (prevents wake-up on dead STA)
                    Try
                        Dim inflight As System.Threading.CancellationTokenSource = llmOperationCts
                        If inflight IsNot Nothing AndAlso Not inflight.IsCancellationRequested Then
                            inflight.Cancel()
                        End If
                    Catch
                    Finally
                        ' Do not Dispose here: the owning RunLlmAsync 'Using' disposes it.
                        llmOperationCts = Nothing
                    End Try

                    ' Force any stale listener loop to exit quickly
                    System.Threading.Interlocked.Increment(listenerGeneration)

                    ' Stop listener without stopping UI STA thread
                    Try
                        Dim t As System.Threading.Tasks.Task = ShutdownHttpListener(stopUiThread:=False)
                        Await System.Threading.Tasks.Task.WhenAny(
                            t,
                            System.Threading.Tasks.Task.Delay(1000)
                        ).ConfigureAwait(False)
                    Catch
                    End Try

                    ' LLM STA: request exit without waiting (no Join while suspending)
                    Try
                        If wasLlmThreadAliveBeforeSleep Then
                            StopLlmUiThreadNonBlocking()
                        End If
                    Catch
                    End Try

                Finally
                    suspendResumeGate.Release()
                End Try
            End Function)
    End Sub

    ''' <summary>
    ''' Asynchronously handles power resume: conditional listener restart and cooldown initialization.
    ''' </summary>
    ''' <summary>
    ''' Asynchronously handles power resume: conditional listener restart and cooldown initialization.
    ''' </summary>
    Friend Sub HandlePowerResumeAsync(userPresent As Boolean)
        System.Threading.Tasks.Task.Run(
            Async Function() As System.Threading.Tasks.Task
                ' Wait (bounded) for any in-flight suspend sequence to release the gate,
                ' instead of dropping the resume and leaving powerChanging stuck at 1.
                Dim gotGate As System.Boolean = False
                Try
                    gotGate = Await suspendResumeGate.WaitAsync(System.TimeSpan.FromSeconds(10)).ConfigureAwait(False)
                Catch
                    gotGate = False
                End Try

                Try
                    If gotGate Then
                        ' Minimal delay before restart
                        Await System.Threading.Tasks.Task.Delay(500).ConfigureAwait(False)

                        isShuttingDown = False

                        ' Clean up old listener reference
                        httpListener = Nothing

                        ' Restart only if user was present pre-suspend
                        If userPresent AndAlso wasListenerRunningBeforeSleep Then
                            Try
                                StartupHttpListener(INI_WebServerBlock)
                            Catch
                            End Try
                        End If

                        ' Set resume cooldown
                        resumeCooldownUntilUtc = System.DateTime.UtcNow.AddSeconds(30)
                    End If
                Finally
                    ' CRITICAL: always clear the power-transition flag, even if the gate
                    ' could not be acquired, so requests/LLM/watchdog are not blocked forever.
                    System.Threading.Interlocked.Exchange(powerChanging, 0)
                    If gotGate Then
                        suspendResumeGate.Release()
                    End If
                End Try
            End Function)
    End Sub

    ''' <summary>
    ''' Signals UI STA thread to exit without blocking caller. Clears related references.
    ''' </summary>
    Private Sub StopLlmUiThreadNonBlocking()
        Try
            If llmSyncContext IsNot Nothing Then
                llmSyncContext.Post(Sub() System.Windows.Forms.Application.ExitThread(), Nothing)
            End If
        Catch
        End Try
        ' No Join here.
        llmScheduler = Nothing
        llmSyncContext = Nothing
        llmThread = Nothing
    End Sub

    ''' <summary>
    ''' Attempts to enter the suspend/resume gate with a short timeout.
    ''' </summary>
    Private Async Function TryEnterGateAsync() As System.Threading.Tasks.Task(Of System.Boolean)
        Try
            Return Await suspendResumeGate.WaitAsync(100).ConfigureAwait(False)
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Initializes listener generation, cancellation token source, logs, and starts listener loop.
    ''' </summary>
    Private Sub StartupHttpListener(WebServerBlock)

        ' WebServerBlock:
        '   1 = no webserver
        '   2 = no Local Chatbot / InkyPlay
        '   3 = no Edge extension receiver
        '   4 = scheduler disabled

        If WebServerBlock = 1 Then Return

        isShuttingDown = False

        ResetInkyServerLogIfNeeded()
        EnsureInkyServerExceptionHooks()

        Dim gen As System.Int64 = System.Threading.Interlocked.Increment(listenerGeneration)
        cts = New System.Threading.CancellationTokenSource()

        Dim startMessage As System.String =
            "HttpListener START gen=" &
            gen.ToString(System.Globalization.CultureInfo.InvariantCulture) &
            " at " &
            System.DateTime.UtcNow.ToString("o", System.Globalization.CultureInfo.InvariantCulture)

        AppendInkyServerLog(startMessage)
        AppendInkyServerLog(
            "StartupHttpListener state: powerChanging=" &
            System.Threading.Interlocked.CompareExchange(powerChanging, 0, 0).ToString(System.Globalization.CultureInfo.InvariantCulture) &
            "; resumeCooldown=" &
            IsInResumeCooldown().ToString() &
            "; isShuttingDown=" &
            isShuttingDown.ToString())

        System.Diagnostics.Debug.WriteLine(startMessage)

        lastListenerProgressUtc = System.DateTime.UtcNow
        listenerTask = StartHttpListener(cts.Token, gen)

        If Not _apActive Then
            EnsureLocalSchedulerTimerStarted()
        End If
    End Sub

    ''' <summary>
    ''' Shuts down listener loop, aborts outstanding contexts, disposes resources and optionally stops UI STA.
    ''' </summary>
    Private Async Function ShutdownHttpListener(
        Optional ByVal stopUiThread As System.Boolean = True
    ) As System.Threading.Tasks.Task
        isShuttingDown = True

        AppendInkyServerLog(
            "ShutdownHttpListener begin. stopUiThread=" &
            stopUiThread.ToString() &
            "; listenerGeneration=" &
            listenerGeneration.ToString(System.Globalization.CultureInfo.InvariantCulture))

        ' Cancel current loop
        Try
            If cts IsNot Nothing Then cts.Cancel()
        Catch
        End Try

        ' Force-break any pending GetContextAsync and clean up thoroughly
        Try
            If httpListener IsNot Nothing Then
                Try
                    If httpListener.IsListening Then httpListener.Stop()
                Catch
                End Try
                Try
                    httpListener.Abort() ' harsher than Close; reliably breaks GetContextAsync
                Catch
                End Try
                Try
                    If httpListener.Prefixes IsNot Nothing Then httpListener.Prefixes.Clear()
                Catch
                End Try
                Try
                    httpListener.Close()
                Catch
                End Try
            End If
        Catch
        Finally
            httpListener = Nothing
        End Try

        ' Await the running listener task to completion
        Try
            If listenerTask IsNot Nothing Then
                Await listenerTask.ConfigureAwait(False)
            End If
        Catch
        Finally
            listenerTask = Nothing
        End Try

        ' Dispose CTS after we've awaited its dependents
        Try
            If cts IsNot Nothing Then cts.Dispose()
        Catch
        Finally
            cts = Nothing
        End Try

        If Not _apActive Then
            StopLocalSchedulerRuntime()
        End If

        System.Diagnostics.Debug.WriteLine(
            "HttpListener STOP at " &
            System.DateTime.UtcNow.ToString("o", System.Globalization.CultureInfo.InvariantCulture))

        ' Stop UI STA only if requested
        If stopUiThread Then
            StopLlmUiThread()
        End If

        AppendInkyServerLog("ShutdownHttpListener end.")

    End Function

    ''' <summary>
    ''' Listener loop: accepts requests, dispatches per-request tasks, logs metrics, and self-recovers after repeated failures.
    ''' </summary>
    Private Async Function StartHttpListener(
        ByVal token As System.Threading.CancellationToken,
        ByVal gen As System.Int64
    ) As System.Threading.Tasks.Task

        Dim prefixes As System.String() = {
            "http://127.0.0.1:12333/",
            "http://localhost:12333/"
        }

        Dim consecutiveFailures As System.Int32 = 0
        Dim lastMetrics As System.DateTime = System.DateTime.UtcNow

        AppendInkyServerLog(
            "StartHttpListener enter. gen=" &
            gen.ToString(System.Globalization.CultureInfo.InvariantCulture))

        While (Not isShuttingDown) AndAlso (Not token.IsCancellationRequested)
            If gen <> listenerGeneration Then
                AppendInkyServerLog(
                    "StartHttpListener exit because generation changed. localGen=" &
                    gen.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                    "; currentGen=" &
                    listenerGeneration.ToString(System.Globalization.CultureInfo.InvariantCulture))
                Return
            End If

            Dim needShortDelay As System.Boolean = False

            Try
                If httpListener Is Nothing Then
                    httpListener = New System.Net.HttpListener()
                    httpListener.IgnoreWriteExceptions = True

                    With httpListener.TimeoutManager
                        .IdleConnection = System.TimeSpan.FromMinutes(10)
                        .HeaderWait = System.TimeSpan.FromSeconds(30)
                        .EntityBody = System.TimeSpan.FromMinutes(10)
                        .DrainEntityBody = System.TimeSpan.FromSeconds(30)
                        .MinSendBytesPerSecond = CType(0UL, System.UInt64)
                    End With

                    Try
                        httpListener.Prefixes.Clear()
                    Catch
                    End Try

                    For Each prefix As System.String In prefixes
                        If Not httpListener.Prefixes.Contains(prefix) Then
                            httpListener.Prefixes.Add(prefix)
                        End If
                        AppendInkyServerLog("HttpListener prefix (pre-start): " & prefix)
                    Next

                    Try
                        httpListener.Start()
                        AppendInkyServerLog("HttpListener started.")
                        For Each prefix As System.String In httpListener.Prefixes
                            AppendInkyServerLog("HttpListener prefix (post-start): " & prefix)
                        Next
                    Catch exStart As System.Exception
                        AppendInkyServerLog("HttpListener.Start failed.", exStart)
                        Throw
                    End Try

                    System.Diagnostics.Debug.WriteLine("HttpListener started.")
                ElseIf Not httpListener.IsListening Then
                    AppendInkyServerLog("HttpListener exists but is not listening. Recycling instance.")
                    Try : httpListener.Close() : Catch : End Try
                    httpListener = Nothing
                    Continue While
                End If

                If INI_APIDebug Then
                    AppendInkyServerLog("Waiting in GetContextAsync().")
                End If

                Dim ctx As System.Net.HttpListenerContext =
                    Await httpListener.GetContextAsync().ConfigureAwait(False)

                lastListenerProgressUtc = System.DateTime.UtcNow

                Dim requestId As System.String = NewInkyRequestId()

                If INI_APIDebug Then
                    Try
                        Dim req As System.Net.HttpListenerRequest = ctx.Request
                        Dim method As System.String = ""
                        Dim requestPath As System.String = "/"
                        Dim rawUrl As System.String = ""
                        Dim hostHeader As System.String = ""
                        Dim userHostName As System.String = ""
                        Dim contentType As System.String = ""
                        Dim contentLength As System.String = ""
                        Dim localEndPointText As System.String = ""
                        Dim remoteEndPointText As System.String = ""

                        If req IsNot Nothing Then
                            method = If(req.HttpMethod, "")
                            requestPath = GetSafeRequestPath(req)
                            rawUrl = If(req.RawUrl, "")
                            hostHeader = If(req.Headers("Host"), "")
                            userHostName = If(req.UserHostName, "")
                            contentType = If(req.ContentType, "")
                            contentLength = req.ContentLength64.ToString(System.Globalization.CultureInfo.InvariantCulture)
                            localEndPointText = GetSafeEndpointText(req.LocalEndPoint)
                            remoteEndPointText = GetSafeEndpointText(req.RemoteEndPoint)
                        End If

                        AppendInkyServerLog(
                            "REQ " & requestId &
                            " accepted: " &
                            method & " " & requestPath &
                            " RawUrl=" & rawUrl &
                            " HostHeader=" & hostHeader &
                            " UserHostName=" & userHostName &
                            " ContentType=" & contentType &
                            " ContentLength=" & contentLength &
                            " LocalEndPoint=" & localEndPointText &
                            " RemoteEndPoint=" & remoteEndPointText)
                    Catch exLog As System.Exception
                        AppendInkyServerLog("REQ " & requestId & " request logging failed.", exLog)
                    End Try
                End If

                Dim ctxLocal As System.Net.HttpListenerContext = ctx
                Dim requestIdLocal As System.String = requestId

                System.Threading.Tasks.Task.Run(
                    Async Function()
                        Dim resLocal As System.Net.HttpListenerResponse = Nothing
                        Try
                            Await HandleHttpRequest(ctxLocal, requestIdLocal).ConfigureAwait(False)
                        Catch exHandle As System.Exception
                            AppendInkyServerLog("REQ " & requestIdLocal & " unhandled request task failure.", exHandle)

                            Try
                                resLocal = ctxLocal.Response

                                Dim bufErr() As System.Byte =
                                        System.Text.Encoding.UTF8.GetBytes(
                                            "{""ok"":false,""error"":""Internal server error.""}")

                                SendBufferedHttpResponse(
                                        resLocal,
                                        500,
                                        "application/json; charset=utf-8",
                                        bufErr,
                                        requestIdLocal,
                                        "outer-fallback-500",
                                        addCors:=True)

                                AppendInkyServerLog(
                                        "REQ " & requestIdLocal &
                                        " outer fallback 500 written through SendBufferedHttpResponse.")
                            Catch exWrite As System.Exception
                                AppendInkyServerLog(
                                        "REQ " & requestIdLocal &
                                        " outer fallback 500 write failed.",
                                        exWrite)
                            End Try
                        Finally
                            lastListenerProgressUtc = System.DateTime.UtcNow
                        End Try
                    End Function)

                Dim now As System.DateTime = System.DateTime.UtcNow
                If (now - lastMetrics).TotalSeconds >= 10.0 Then
                    Dim gdi As System.UInt32 = GetGdiCount()
                    Dim usr As System.UInt32 = GetUserCount()

                    If INI_APIDebug Then
                        AppendInkyServerLog(
                            "Listener metrics: GDI=" &
                            gdi.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                            "; USER=" &
                            usr.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                            "; activeRequests=" &
                            System.Threading.Interlocked.CompareExchange(activeRequests, 0, 0).ToString(System.Globalization.CultureInfo.InvariantCulture) &
                            "; activeJobs=" &
                            System.Threading.Interlocked.CompareExchange(activeJobs, 0, 0).ToString(System.Globalization.CultureInfo.InvariantCulture))
                    End If

                    System.Diagnostics.Debug.WriteLine(
                        System.String.Format(
                            System.Globalization.CultureInfo.InvariantCulture,
                            "RES {0:HH:mm:ss}: GDI={1}  USER={2}",
                            now, gdi, usr))

                    If gdi >= 8000UI OrElse usr >= 8000UI Then
                        System.Diagnostics.Debug.WriteLine("WARN: High handle count – check for GDI/USER leaks.")
                    End If

                    lastMetrics = now
                End If

                consecutiveFailures = 0

            Catch ex As System.ObjectDisposedException
                consecutiveFailures += 1
                needShortDelay = True
                AppendInkyServerLog(
                    "Listener ObjectDisposedException. consecutiveFailures=" &
                    consecutiveFailures.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    ex)

            Catch ex As System.Net.HttpListenerException
                consecutiveFailures += 1
                needShortDelay = True
                AppendInkyServerLog(
                    "Listener HttpListenerException. consecutiveFailures=" &
                    consecutiveFailures.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    ex)

                System.Diagnostics.Debug.WriteLine("Listener error: " & ex.Message)

            Catch ex As System.Exception
                consecutiveFailures += 1
                AppendInkyServerLog(
                    "Listener error. consecutiveFailures=" &
                    consecutiveFailures.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    ex)

                System.Diagnostics.Debug.WriteLine("Listener error: " & ex.Message)
            End Try

            If needShortDelay AndAlso (Not token.IsCancellationRequested) Then
                Try
                    Await System.Threading.Tasks.Task.Delay(50, token).ConfigureAwait(False)
                Catch
                End Try
            End If

            If consecutiveFailures >= 10 AndAlso (Not isShuttingDown) AndAlso (Not token.IsCancellationRequested) Then
                System.Diagnostics.Debug.WriteLine("Restarting HttpListener after 10 failures.")
                AppendInkyServerLog("Restarting HttpListener after 10 consecutive failures.")

                Try
                    If httpListener IsNot Nothing Then
                        Try : httpListener.Abort() : Catch : End Try
                        Try : httpListener.Close() : Catch : End Try
                    End If
                Catch exReset As System.Exception
                    AppendInkyServerLog("Listener reset after repeated failures failed.", exReset)
                Finally
                    httpListener = Nothing
                End Try

                consecutiveFailures = 0

                Try
                    Await System.Threading.Tasks.Task.Delay(5000, token).ConfigureAwait(False)
                Catch
                End Try
            End If
        End While

        AppendInkyServerLog(
            "StartHttpListener exit. isShuttingDown=" &
            isShuttingDown.ToString() &
            "; tokenCancelled=" &
            token.IsCancellationRequested.ToString())
    End Function



    ''' <summary>
    ''' Starts hidden power notification window if not already active.
    ''' </summary>
    Private Sub StartPowerWatch()
        If powerWindow Is Nothing Then
            powerWindow = New PowerNotificationWindow(Me)
        End If
    End Sub

    ''' <summary>
    ''' Disposes hidden power notification window if active.
    ''' </summary>
    Private Sub StopPowerWatch()
        If powerWindow IsNot Nothing Then
            powerWindow.Dispose()
            powerWindow = Nothing
        End If
    End Sub

    ''' <summary>
    ''' Handles system power mode changes (legacy event path). Delegates to suspend/resume logic.
    ''' </summary>
    Private Sub OnPowerModeChanged(ByVal sender As System.Object,
                                   ByVal e As Microsoft.Win32.PowerModeChangedEventArgs)
        If e Is Nothing Then Return

        Select Case e.Mode
            Case Microsoft.Win32.PowerModes.Suspend
                ' Graceful listener stop in the background
                System.Threading.ThreadPool.QueueUserWorkItem(
                    Sub(state As Object)
                        Try : ShutdownHttpListener().GetAwaiter().GetResult() : Catch : End Try
                    End Sub)

            Case Microsoft.Win32.PowerModes.Resume
                ' Avoid re-entrancy; delegate to unified resume path (userPresent:=True)
                If System.Threading.Interlocked.Exchange(restartingAfterResume, 1) = 1 Then Return
                System.Threading.ThreadPool.QueueUserWorkItem(
                    Sub(state As Object)
                        Try
                            HandlePowerResumeAsync(userPresent:=True)
                        Finally
                            System.Threading.Interlocked.Exchange(restartingAfterResume, 0)
                        End Try
                    End Sub)
        End Select
    End Sub

    ''' <summary>
    ''' Starts watchdog timer that monitors listener health and restarts if dead or faulted.
    ''' </summary>
    Private Sub StartListenerWatchdog()
        If watchdog IsNot Nothing Then Return

        watchdog = New System.Threading.Timer(
            Sub(stateObj As System.Object)
                Try
                    ' Skip during power transitions
                    If System.Threading.Interlocked.CompareExchange(powerChanging, 0, 0) <> 0 Then Return

                    ' Skip if in cooldown after resume
                    If IsInResumeCooldown() Then Return

                    ' Check active work
                    Dim inFlight As Integer = Threading.Interlocked.CompareExchange(activeRequests, 0, 0)
                    Dim jobsInFlight As Integer = Threading.Interlocked.CompareExchange(activeJobs, 0, 0)
                    If inFlight > 0 OrElse jobsInFlight > 0 Then
                        lastListenerProgressUtc = System.DateTime.UtcNow
                        Return
                    End If

                    ' Assess listener health without penalizing idle
                    Dim listenerIsDead As Boolean = False
                    Try
                        listenerIsDead = (httpListener Is Nothing) OrElse (Not httpListener.IsListening)
                    Catch
                        listenerIsDead = True
                    End Try

                    ' Simple timeout check (idle is OK; only use if listener isnot listening/faulted)
                    Dim age As Double = (System.DateTime.UtcNow - lastListenerProgressUtc).TotalSeconds

                    ' If the listener is gone or loop has faulted, restart cleanly
                    Dim loopFaulted As Boolean = False
                    Try
                        loopFaulted = (listenerTask IsNot Nothing AndAlso listenerTask.IsCompleted AndAlso listenerTask.IsFaulted)
                    Catch
                        loopFaulted = True
                    End Try

                    If (listenerIsDead OrElse loopFaulted) AndAlso Not isShuttingDown Then
                        System.Threading.Tasks.Task.Run(
                            Async Function()
                                Try
                                    ' Fully stop old instance to free the port
                                    Await ShutdownHttpListener(stopUiThread:=False)
                                Catch
                                End Try
                                Try
                                    StartupHttpListener(INI_WebServerBlock)
                                Catch
                                End Try
                            End Function)
                        Return
                    End If

                    ' No restart solely due to idle

                Catch
                End Try
            End Sub,
            state:=Nothing,
            dueTime:=System.TimeSpan.FromSeconds(30),
            period:=System.TimeSpan.FromSeconds(10))
    End Sub

    ''' <summary>
    ''' Stops and disposes the listener watchdog timer.
    ''' </summary>
    Private Sub StopListenerWatchdog()
        Try
            If watchdog IsNot Nothing Then
                watchdog.Dispose()
            End If
        Catch
        Finally
            watchdog = Nothing
        End Try
    End Sub


    Private lastSuccessfulOperationUtc As System.DateTime = System.DateTime.UtcNow
    Private Const StuckDetectionMinutes As Integer = 15  ' If no successful op for 15 min, consider stuck

    Private ReadOnly inkyServerLogSync As New Object()
    Private inkyServerLogInitialized As Integer = 0
    Private inkyServerExceptionHooksInstalled As Integer = 0
    Private nextInkyRequestId As Long = 0

    Private Function GetInkyServerLogPath() As System.String
        Return System.IO.Path.Combine(System.IO.Path.GetTempPath(), "RI_InkyServer_Log.txt")
    End Function

    Private Function NewInkyRequestId() As System.String
        Dim id As Long = System.Threading.Interlocked.Increment(nextInkyRequestId)
        Return id.ToString("000000", System.Globalization.CultureInfo.InvariantCulture)
    End Function

    Private Function ClipForInkyServerLog(
        ByVal value As System.String,
        Optional ByVal maxLength As Integer = 2000
    ) As System.String

        Dim s As System.String = If(value, System.String.Empty)

        s = s.Replace(vbCr, "\r").Replace(vbLf, "\n")

        If s.Length > maxLength Then
            s = s.Substring(0, maxLength) & "... [truncated]"
        End If

        Return s
    End Function

    Private Sub ResetInkyServerLogIfNeeded()
        If Not INI_APIDebug Then Return

        If System.Threading.Interlocked.CompareExchange(inkyServerLogInitialized, 1, 0) <> 0 Then
            Return
        End If

        SyncLock inkyServerLogSync
            Dim logPath As System.String = GetInkyServerLogPath()

            Try
                If System.IO.File.Exists(logPath) Then
                    System.IO.File.Delete(logPath)
                End If
            Catch
            End Try

            Try
                Dim header As New System.Text.StringBuilder()

                header.AppendLine("==== RI_InkyServer_Log.txt ====")
                header.AppendLine("StartedLocal: " & System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", System.Globalization.CultureInfo.InvariantCulture))
                header.AppendLine("StartedUtc:   " & System.DateTime.UtcNow.ToString("o", System.Globalization.CultureInfo.InvariantCulture))
                header.AppendLine("Process:      " & System.Diagnostics.Process.GetCurrentProcess().ProcessName)
                header.AppendLine("PID:          " & System.Diagnostics.Process.GetCurrentProcess().Id.ToString(System.Globalization.CultureInfo.InvariantCulture))
                header.AppendLine("Machine:      " & System.Environment.MachineName)
                header.AppendLine("User:         " & System.Environment.UserDomainName & "\" & System.Environment.UserName)
                header.AppendLine("64BitProcess: " & System.Environment.Is64BitProcess.ToString())
                header.AppendLine("CLR:          " & System.Environment.Version.ToString())
                header.AppendLine("OS:           " & System.Environment.OSVersion.ToString())
                header.AppendLine("Culture:      " & System.Globalization.CultureInfo.CurrentCulture.Name)
                header.AppendLine("UICulture:    " & System.Globalization.CultureInfo.CurrentUICulture.Name)
                header.AppendLine()

                System.IO.File.AppendAllText(logPath, header.ToString(), System.Text.Encoding.UTF8)
            Catch
            End Try
        End SyncLock
    End Sub

    Private Sub AppendInkyServerLog(ByVal message As System.String)
        If Not INI_APIDebug Then Return

        ResetInkyServerLogIfNeeded()

        SyncLock inkyServerLogSync
            Try
                System.IO.File.AppendAllText(
                    GetInkyServerLogPath(),
                    System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", System.Globalization.CultureInfo.InvariantCulture) &
                    " [T" &
                    System.Threading.Thread.CurrentThread.ManagedThreadId.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                    "] " &
                    If(message, System.String.Empty) &
                    System.Environment.NewLine,
                    System.Text.Encoding.UTF8)
            Catch
            End Try
        End SyncLock
    End Sub

    Private Sub AppendInkyServerLog(ByVal message As System.String, ByVal ex As System.Exception)
        If ex Is Nothing Then
            AppendInkyServerLog(message)
            Return
        End If

        AppendInkyServerLog(
            If(message, System.String.Empty) &
            System.Environment.NewLine &
            ex.ToString())
    End Sub

    Private Sub EnsureInkyServerExceptionHooks()
        If Not INI_APIDebug Then Return

        If System.Threading.Interlocked.CompareExchange(inkyServerExceptionHooksInstalled, 1, 0) <> 0 Then
            Return
        End If

        AddHandler AppDomain.CurrentDomain.UnhandledException,
            Sub(sender As System.Object, e As System.UnhandledExceptionEventArgs)
                Dim exObj As System.Exception = TryCast(e.ExceptionObject, System.Exception)
                AppendInkyServerLog(
                    "AppDomain.CurrentDomain.UnhandledException. IsTerminating=" &
                    e.IsTerminating.ToString(),
                    exObj)
            End Sub

        AddHandler System.Threading.Tasks.TaskScheduler.UnobservedTaskException,
            Sub(sender As System.Object, e As System.Threading.Tasks.UnobservedTaskExceptionEventArgs)
                AppendInkyServerLog("TaskScheduler.UnobservedTaskException", e.Exception)
            End Sub
    End Sub

    Private Function GetSafeRequestPath(ByVal req As System.Net.HttpListenerRequest) As System.String
        Try
            If req Is Nothing Then Return "/"

            Dim path As System.String = Nothing

            Try
                If req.Url IsNot Nothing Then
                    path = req.Url.AbsolutePath
                End If
            Catch
                path = Nothing
            End Try

            If System.String.IsNullOrWhiteSpace(path) Then
                Try
                    path = req.RawUrl
                Catch
                    path = Nothing
                End Try
            End If

            If System.String.IsNullOrWhiteSpace(path) Then
                path = "/"
            End If

            Return path
        Catch
            Return "/"
        End Try
    End Function

    Private Function GetSafeEndpointText(ByVal value As System.Object) As System.String
        Try
            If value Is Nothing Then Return ""
            Return value.ToString()
        Catch
            Return ""
        End Try
    End Function




End Class
