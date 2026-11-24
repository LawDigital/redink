' Part of Red Ink. One worker owns engine creation, audio writes and recovery.
Option Explicit On
Option Strict On

Namespace Transcription
    Friend NotInheritable Class LiveStreamRotationException
        Inherits System.IO.IOException
        Public Sub New()
            MyBase.New("The live stream reached its session rotation point.")
        End Sub
    End Class

    Public NotInheritable Class LiveTranscriptionSession
        Private ReadOnly gate As New System.Object()
        Private ReadOnly wake As New System.Threading.SemaphoreSlim(0, 1)
        Private ReadOnly audio As New LiveAudioBuffer()
        Private ReadOnly shutdown As New System.Threading.CancellationTokenSource()
        Private ReadOnly factory As System.Func(Of System.Threading.Tasks.Task(Of ITranscriptionEngine))
        Private ReadOnly options As TranscriptionOptions
        Private ReadOnly bytesPerSecond As System.Int32
        Private engine As ITranscriptionEngine
        Private attempt As System.Threading.CancellationTokenSource
        Private worker As System.Threading.Tasks.Task
        Private bufferFailed As System.Boolean
        Private stopping As System.Boolean
        Private recovering As System.Boolean
        Private interruptions As System.Int32
        Private fault As System.String
        Private cursor As System.Int64
        Private inFlight As System.Threading.Tasks.Task
        Private retainFrom As System.Int64 = -1
        Private recoveryStart As System.Int64
        Private Const BoundarySeconds As System.Int32 = 20

        Public Event PartialResult As System.EventHandler(Of TranscriptionEventArgs)
        Public Event FinalResult As System.EventHandler(Of TranscriptionEventArgs)
        Public Event AudioBufferFailed As System.EventHandler(Of TranscriptionErrorEventArgs)
        Public Event EngineError As System.EventHandler(Of TranscriptionErrorEventArgs)
        Public Event Status As System.EventHandler(Of TranscriptionStatusEventArgs)

        Public Sub New(makeEngine As System.Func(Of System.Threading.Tasks.Task(Of ITranscriptionEngine)), settings As TranscriptionOptions, channels As System.Int32, Optional keepFullAudio As System.Boolean = False)
            factory = makeEngine
            options = settings.ForSupervisedLive()
            options.MultiChannelDiarization = (channels = 2)
            bytesPerSecond = 32000 * channels
            If keepFullAudio Then retainFrom = 0
        End Sub

        Private Sub Signal()
            ' Called with gate held. Several producer notifications need only one wakeup.
            If wake.CurrentCount = 0 Then wake.Release()
        End Sub

        Public Sub Append(data As System.Byte(), count As System.Int32)
            Dim failure As System.Exception = Nothing
            SyncLock gate
                If stopping OrElse bufferFailed OrElse count <= 0 Then Return
                Try
                    audio.Append(data, count)
                Catch ex As System.Exception
                    bufferFailed = True
                    If retainFrom < 0 Then retainFrom = recoveryStart
                    failure = ex
                End Try
                Signal()
            End SyncLock
            If failure IsNot Nothing Then
                RaiseEvent AudioBufferFailed(Me, New TranscriptionErrorEventArgs("Windows could not buffer additional live audio: " & failure.Message, failure, True))
            End If
        End Sub

        Public Async Function StartAsync(first As ITranscriptionEngine) As System.Threading.Tasks.Task
            Await OpenAsync(first).ConfigureAwait(False)
            worker = System.Threading.Tasks.Task.Run(New System.Func(Of System.Threading.Tasks.Task)(AddressOf PumpAsync))
        End Function

        Private Async Function OpenAsync(nextEngine As ITranscriptionEngine) As System.Threading.Tasks.Task
            engine = nextEngine
            attempt = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(shutdown.Token)
            AddHandler engine.PartialResult, AddressOf OnPartial
            AddHandler engine.FinalResult, AddressOf OnFinal
            AddHandler engine.Status, AddressOf OnStatus
            AddHandler engine.EngineError, AddressOf OnError
            shutdown.Token.ThrowIfCancellationRequested()
            attempt.CancelAfter(System.TimeSpan.FromSeconds(60))
            Await WaitForAttemptAsync(engine.StartLiveAsync(options, attempt.Token)).ConfigureAwait(False)
            attempt.Token.ThrowIfCancellationRequested()
            attempt.CancelAfter(System.Threading.Timeout.Infinite)
        End Function

        Private Sub OnPartial(sender As System.Object, e As TranscriptionEventArgs)
            If System.Object.ReferenceEquals(sender, engine) Then RaiseEvent PartialResult(Me, e)
        End Sub

        Private Sub OnFinal(sender As System.Object, e As TranscriptionEventArgs)
            SyncLock gate
                If Not System.Object.ReferenceEquals(sender, engine) Then Return
                ' Providers generally report finals asynchronously without a common PCM
                ' timestamp. Keep a boundary overlap and everything after the last final.
                recoveryStart = System.Math.Max(recoveryStart, cursor - CLng(bytesPerSecond) * BoundarySeconds)
                If retainFrom < 0 Then audio.ReleaseBefore(recoveryStart)
            End SyncLock
            RaiseEvent FinalResult(Me, e)
        End Sub

        Private Sub OnStatus(sender As System.Object, e As TranscriptionStatusEventArgs)
            If System.Object.ReferenceEquals(sender, engine) Then RaiseEvent Status(Me, e)
        End Sub

        Private Sub OnError(sender As System.Object, e As TranscriptionErrorEventArgs)
            Dim interruptAttempt As System.Threading.CancellationTokenSource = Nothing
            SyncLock gate
                If Not System.Object.ReferenceEquals(sender, engine) Then Return
                If e.Fatal AndAlso fault Is Nothing Then
                    fault = If(e.Message, "Live stream ended.")
                    If retainFrom < 0 Then retainFrom = recoveryStart
                    interruptAttempt = attempt
                    Signal()
                End If
            End SyncLock
            If interruptAttempt IsNot Nothing Then
                ' A transport can report its failure from inside a receive callback.
                ' Cancel outside that callback and never wait on the receiving task here.
                Dim cancelTask = System.Threading.Tasks.Task.Run(
                    Sub()
                        Try
                            interruptAttempt.Cancel()
                        Catch ex As System.ObjectDisposedException
                        End Try
                    End Sub)
            End If
            RaiseEvent EngineError(Me, e)
        End Sub

        Public Sub CancelStartup()
            shutdown.Cancel()
        End Sub

        Private Async Function WaitForAttemptAsync(operation As System.Threading.Tasks.Task) As System.Threading.Tasks.Task
            inFlight = operation
            If Not operation.IsCompleted Then
                Dim interrupted As New System.Threading.Tasks.TaskCompletionSource(Of System.Boolean)(System.Threading.Tasks.TaskCreationOptions.RunContinuationsAsynchronously)
                Using registration = attempt.Token.Register(Sub() interrupted.TrySetResult(True))
                    If Await System.Threading.Tasks.Task.WhenAny(operation, interrupted.Task).ConfigureAwait(False) IsNot operation Then
                        attempt.Token.ThrowIfCancellationRequested()
                    End If
                End Using
            End If
            Await operation.ConfigureAwait(False)
        End Function

        Private Shared Async Function CloseAttemptAsync(previous As ITranscriptionEngine, pending As System.Threading.Tasks.Task, source As System.Threading.CancellationTokenSource) As System.Threading.Tasks.Task
            Try
                If pending IsNot Nothing Then
                    Try
                        Await pending.ConfigureAwait(False)
                    Catch ex As System.Exception
                        ' The worker already handles this operation's failure.
                    End Try
                End If
                Dim stopFailure As System.Exception = Nothing
                Try
                    Await previous.StopLiveAsync().ConfigureAwait(False)
                Catch ex As System.Exception
                    stopFailure = ex
                End Try
                Await previous.DisposeAsync().ConfigureAwait(False)
                If stopFailure IsNot Nothing Then Throw stopFailure
            Finally
                source.Dispose()
            End Try
        End Function

        Private Async Function ReleaseEngineAsync(graceful As System.Boolean) As System.Threading.Tasks.Task(Of System.Boolean)
            Dim previous = engine
            If previous Is Nothing Then Return True
            Dim source = attempt
            Dim pending = inFlight
            inFlight = Nothing
            If Not graceful Then
                SyncLock gate
                    engine = Nothing
                End SyncLock
                source.Cancel()
            End If
            Dim finished As System.Boolean = False
            Dim cleanup = System.Threading.Tasks.Task.Run(Function() CloseAttemptAsync(previous, pending, source))
            Try
                If Await System.Threading.Tasks.Task.WhenAny(cleanup, System.Threading.Tasks.Task.Delay(8000)).ConfigureAwait(False) Is cleanup Then
                    Await cleanup.ConfigureAwait(False)
                    finished = True
                Else
                    Try
                        source.Cancel()
                    Catch ex As System.ObjectDisposedException
                    End Try
                    Dim observed = cleanup.ContinueWith(Sub(t) System.Diagnostics.Debug.WriteLine(t.Exception),
                        System.Threading.CancellationToken.None, System.Threading.Tasks.TaskContinuationOptions.OnlyOnFaulted,
                        System.Threading.Tasks.TaskScheduler.Default)
                    RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Live session shutdown timed out.", Nothing, False))
                End If
            Catch ex As System.Exception
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Live shutdown: " & ex.Message, ex, False))
            End Try
            SyncLock gate
                engine = Nothing
            End SyncLock
            RemoveHandler previous.PartialResult, AddressOf OnPartial
            RemoveHandler previous.FinalResult, AddressOf OnFinal
            RemoveHandler previous.Status, AddressOf OnStatus
            RemoveHandler previous.EngineError, AddressOf OnError
            attempt = Nothing
            Return finished
        End Function

        Private Async Function PumpAsync() As System.Threading.Tasks.Task
            Dim backoffSeconds As System.Int32 = 1
            Dim stableSince = System.Diagnostics.Stopwatch.StartNew()
            Do
                Dim failure As System.String = Nothing
                Dim rotation As System.Boolean = False
                Try
                    Dim frame As System.Byte()
                    SyncLock gate
                        failure = fault
                        If failure IsNot Nothing Then Throw New System.IO.IOException(failure)
                        frame = audio.Read(cursor, bytesPerSecond \ 20)
                        If stopping AndAlso frame.Length = 0 Then Exit Do
                    End SyncLock
                    If frame.Length = 0 Then
                        Await wake.WaitAsync(shutdown.Token).ConfigureAwait(False)
                    Else
                        attempt.CancelAfter(System.TimeSpan.FromSeconds(90))
                        Await WaitForAttemptAsync(engine.PushAudioAsync(frame, frame.Length, attempt.Token)).ConfigureAwait(False)
                        attempt.Token.ThrowIfCancellationRequested()
                        attempt.CancelAfter(System.Threading.Timeout.Infinite)
                        SyncLock gate
                            If fault Is Nothing Then
                                cursor += frame.Length
                            End If
                        End SyncLock
                    End If
                Catch ex As LiveStreamRotationException
                    failure = ex.Message
                    rotation = True
                Catch ex As System.Exception
                    failure = ex.Message
                End Try
                If failure Is Nothing Then Continue Do

                SyncLock gate
                    recovering = True
                    If Not rotation Then
                        interruptions += 1
                        If retainFrom < 0 Then retainFrom = recoveryStart
                    End If
                End SyncLock
                If stableSince.Elapsed.TotalSeconds >= 30 Then backoffSeconds = 1
                Do
                    Dim statusText = If(rotation, "Rotating live stream. ", "Live session interrupted. ")
                    RaiseEvent Status(Me, New TranscriptionStatusEventArgs(statusText & "Reconnecting in " & backoffSeconds.ToString() & " s. " & failure) With {.IsSessionTransition = True})
                    Dim closed = Await ReleaseEngineAsync(rotation).ConfigureAwait(False)
                    If Not rotation OrElse Not closed Then
                        SyncLock gate
                            If retainFrom < 0 Then retainFrom = recoveryStart
                            cursor = recoveryStart
                        End SyncLock
                    End If
                    rotation = False
                    SyncLock gate
                        fault = Nothing
                        If stopping Then Return
                    End SyncLock
                    Try
                        Await System.Threading.Tasks.Task.Delay(backoffSeconds * 1000, shutdown.Token).ConfigureAwait(False)
                        SyncLock gate
                            If stopping Then Return
                        End SyncLock
                        Dim candidate = Await factory().ConfigureAwait(False)
                        Await OpenAsync(candidate).ConfigureAwait(False)
                        SyncLock gate
                            If fault IsNot Nothing Then Throw New System.IO.IOException(fault)
                            recovering = False
                        End SyncLock
                        backoffSeconds = System.Math.Min(backoffSeconds * 2, 30)
                        stableSince.Restart()
                        RaiseEvent Status(Me, New TranscriptionStatusEventArgs("Live session resumed.") With {.IsSessionTransition = True})
                        Exit Do
                    Catch ex As System.Exception
                        failure = ex.Message
                        RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Live reconnect: " & failure, ex, False))
                    End Try
                    backoffSeconds = System.Math.Min(backoffSeconds * 2, 30)
                Loop
            Loop
        End Function

        Public Async Function StopAsync() As System.Threading.Tasks.Task
            Dim cancelRetry As System.Boolean
            SyncLock gate
                cancelRetry = recovering
                stopping = True
                Signal()
            End SyncLock
            ' Drain healthy sessions first. Cancel an offline retry or stalled write
            ' after a short grace period; the unsent PCM remains available for export.
            If cancelRetry Then shutdown.Cancel()
            shutdown.CancelAfter(System.TimeSpan.FromSeconds(8))
            Try
                If worker IsNot Nothing Then Await worker.ConfigureAwait(False)
            Finally
                ' Keep the final recognizer alive for its native stop/flush below.
                If Not shutdown.IsCancellationRequested Then shutdown.CancelAfter(System.Threading.Timeout.Infinite)
            End Try
            Dim graceful As System.Boolean
            SyncLock gate
                graceful = fault Is Nothing AndAlso Not shutdown.IsCancellationRequested
                If Not graceful Then cursor = recoveryStart
            End SyncLock
            Dim flushed = Await ReleaseEngineAsync(graceful).ConfigureAwait(False)
            SyncLock gate
                If Not flushed OrElse fault IsNot Nothing Then cursor = recoveryStart
                If retainFrom >= 0 Then cursor = System.Math.Min(cursor, retainFrom)
            End SyncLock
        End Function

        Public ReadOnly Property InterruptedSessions As System.Int32
            Get
                Return interruptions
            End Get
        End Property

        Public ReadOnly Property PendingBytes As System.Int64
            Get
                Return audio.Length - cursor
            End Get
        End Property

        ' Accessible only after StopAsync has completed. Advance after successful close,
        ' so a disk error leaves the part available for retry. Each WAV fits classic RIFF.
        Public Sub SavePendingPart(path As System.String)
            Dim endOffset = System.Math.Min(audio.Length, cursor + 1024L * 1024L * 1024L)
            Dim position = cursor
            Using wav As New NAudio.Wave.WaveFileWriter(path, New NAudio.Wave.WaveFormat(16000, 16, bytesPerSecond \ 32000))
                While position < endOffset
                    Dim data = audio.Read(position, CInt(System.Math.Min(65536L, endOffset - position)))
                    wav.Write(data, 0, data.Length)
                    position += data.Length
                End While
            End Using
            cursor = position
        End Sub

        Public Sub ReleaseBuffer()
            audio.Dispose()
            shutdown.Dispose()
            wake.Dispose()
        End Sub
    End Class
End Namespace
