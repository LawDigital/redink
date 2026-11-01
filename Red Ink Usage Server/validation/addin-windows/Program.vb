Option Strict On
Option Explicit On
Imports System.Threading
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary

Module Program
    Private checks As Integer
    Private Sub Require(value As Boolean, name As String)
        checks += 1
        If Not value Then Throw New Exception("Failed: " & name)
        Console.WriteLine("PASS " & name)
    End Sub

    Private Sub PumpUntil(task As System.Threading.Tasks.Task, milliseconds As Integer)
        Dim elapsed = System.Diagnostics.Stopwatch.StartNew()
        While Not task.IsCompleted AndAlso elapsed.ElapsedMilliseconds < milliseconds
            Application.DoEvents()
            Thread.Sleep(5)
        End While
        Require(task.IsCompleted, "native asynchronous operation completes within its bound")
    End Sub

    <STAThread>
    Public Sub Main()
        Application.EnableVisualStyles()
        Dim beforeContext = SynchronizationContext.Current
        Dim beforeAutoInstall = WindowsFormsSynchronizationContext.AutoInstall
        Dim rtf = "{\rtf1\ansi\ansicpg1252\uc1 First \u945?\par Second\tab End}"
        Dim expected As String
        Using original As New RichTextBox()
            original.Rtf = rtf
            expected = original.Text
        End Using
        Dim activeContext = SynchronizationContext.Current
        Require(SharedMethods.RtfToPlainText(rtf) = expected, "STA Unicode/paragraph/tab parsing matches original RichEdit")
        Require(SynchronizationContext.Current Is activeContext, "RTF preserves caller synchronization context")
        Require(WindowsFormsSynchronizationContext.AutoInstall = beforeAutoInstall, "RTF restores AutoInstall")
        Dim tableRtf = "{\rtf1\ansi\trowd\cellx1000\cellx2000\intbl A\cell B\cell\row}"
        Using tableReference As New RichTextBox()
            tableReference.Rtf = tableRtf
            Require(SharedMethods.RtfToPlainText(tableRtf) = tableReference.Text, "RTF table text matches original RichEdit")
        End Using
        Dim malformedFailed As Boolean
        Try
            SharedMethods.RtfToPlainText("invalid RTF")
        Catch ex As ArgumentException
            malformedFailed = True
        End Try
        Require(malformedFailed, "invalid RTF preserves parser exception")

        Dim workerFailure As Exception = Nothing
        Dim mtaText As String = Nothing
        Dim worker As New Thread(
            Sub()
                Try
                    Require(SynchronizationContext.Current Is Nothing, "MTA begins without UI context")
                    mtaText = SharedMethods.RtfToPlainText(rtf)
                    Require(SynchronizationContext.Current Is Nothing, "MTA caller retains no UI context")
                Catch ex As Exception
                    workerFailure = ex
                End Try
            End Sub)
        worker.SetApartmentState(ApartmentState.MTA)
        worker.Start()
        Require(worker.Join(5000), "MTA RTF conversion returns within five seconds")
        If workerFailure IsNot Nothing Then Throw workerFailure
        Require(mtaText = expected, "MTA Unicode parsing matches original RichEdit")

        Dim tempFile = IO.Path.GetTempFileName()
        Try
            IO.File.WriteAllText(tempFile, rtf, New System.Text.UTF8Encoding(False))
            Require(SharedMethods.LoadRtfFileAsPlainText(tempFile) = expected, "file parser preserves Unicode/paragraphs")
            Dim readError As String = Nothing
            Require(SharedMethods.ReadRtfAsText(tempFile, True, readError) = expected AndAlso readError = "", "legacy read result/error contract")
            IO.File.WriteAllText(tempFile, "invalid RTF")
            Dim failureResult = SharedMethods.ReadRtfAsText(tempFile, False, readError)
            Require(failureResult = "" AndAlso Not String.IsNullOrEmpty(readError), "legacy suppression preserves readError")
        Finally
            IO.File.Delete(tempFile)
        End Try

        Dim dispatcher As New WindowsFormsSynchronizationContext()
        SynchronizationContext.SetSynchronizationContext(dispatcher)
        Dim registry As New HostModelessFormRegistry()
        LogWindow.InitializeUiDispatcher(dispatcher, registry)
        Require(SharedMethods.HasHostUiDispatcher AndAlso SharedMethods.IsOnHostUiThread, "actual shared dispatcher captures this STA")
        Dim dispatched As System.Threading.Tasks.Task(Of Integer) = Nothing
        Dim hostThreadId = Thread.CurrentThread.ManagedThreadId
        Dim dispatchWorker As New Thread(Sub()
                                             dispatched = SharedMethods.PostToHostUiAsync(Function() Thread.CurrentThread.ManagedThreadId)
                                         End Sub)
        dispatchWorker.SetApartmentState(ApartmentState.MTA)
        dispatchWorker.Start()
        Require(dispatchWorker.Join(2000), "MTA host dispatch returns an asynchronous task")
        PumpUntil(dispatched, 5000)
        Require(dispatched.GetAwaiter().GetResult() = hostThreadId, "actual awaitable runs the full action on the captured host thread")
        For Each extension As String In New String() {".doc", ".rtf", ".xls", ".ppt", ".docx", ".xlsx", ".pptx", ".msg", ".eml"}
            Dim reader As New Global.NativeOfficeReaderProbe.OfficeReaderDispatchProbe()
            Dim readTask As System.Threading.Tasks.Task(Of Tuple(Of Boolean, String, String)) = Nothing
            Dim readCallerId As Integer
            Dim path = "probe" & extension
            Dim readWorker As New Thread(Sub()
                                             readCallerId = Thread.CurrentThread.ManagedThreadId
                                             readTask = reader.ReadOfficeForProbe(path)
                                         End Sub)
            readWorker.SetApartmentState(ApartmentState.MTA)
            readWorker.Start()
            Require(readWorker.Join(2000), "Office reader producer does not block: " & extension)
            PumpUntil(readTask, 5000)
            Dim requiresHost = extension = ".doc" OrElse extension = ".rtf" OrElse extension = ".xls" OrElse extension = ".ppt"
            Require(reader.LastReaderThreadId = If(requiresHost, hostThreadId, readCallerId), "Office reader dispatch preserves apartment boundary: " & extension)
            Dim readResult = readTask.GetAwaiter().GetResult()
            Require(readResult.Item1 AndAlso readResult.Item2 = "office:" & path AndAlso readResult.Item3 = "office label", "reader result tuple preserves success/text/label: " & extension)
        Next
        LogWindow.ShowLogWindow()
        LogWindow.AppendLog("first")
        Dim logWorker As New Thread(Sub()
                                       For i As Integer = 1 To 250
                                           LogWindow.AppendLog("background " & i.ToString())
                                       Next
                                   End Sub)
        logWorker.Start()
        Require(logWorker.Join(2000), "background log producer never waits for UI")
        For i As Integer = 1 To 20
            Application.DoEvents()
        Next
        Require(Application.OpenForms.Count = 1, "one host UI log form created")
        Dim callbackCount As Integer
        LogWindow.OnCloseRequested = Function()
                                         callbackCount += 1
                                         Return False
                                     End Function
        LogWindow.CloseLogWindow()
        Application.DoEvents()
        Require(callbackCount = 0, "programmatic log close does not ask to abort")
        Require(Application.OpenForms.Count = 0, "programmatic close disposes log form")
        LogWindow.ShowLogWindow()
        Application.DoEvents()
        Dim queuedActionRan As Boolean
        Dim canceledDispatch As System.Threading.Tasks.Task = Nothing
        Dim pendingWorker As New Thread(Sub()
                                            canceledDispatch = SharedMethods.PostToHostUiAsync(Sub() queuedActionRan = True)
                                        End Sub)
        pendingWorker.SetApartmentState(ApartmentState.MTA)
        pendingWorker.Start()
        Require(pendingWorker.Join(2000), "pending host action does not block its producer")
        LogWindow.Shutdown()
        registry.CloseAll()
        PumpUntil(canceledDispatch, 5000)
        Require(canceledDispatch.IsCanceled AndAlso Not queuedActionRan, "host shutdown cancels the actual queued awaitable before UI work")
        Require(Not SharedMethods.HasHostUiDispatcher, "closed host dispatcher is unavailable")
        LogWindow.Shutdown()
        LogWindow.AppendLog("ignored after shutdown")
        Application.DoEvents()
        Require(Application.OpenForms.Count = 0, "host shutdown is idempotent and does not recreate log")
        Require(callbackCount = 0, "shutdown does not invoke operator Stop callback")
        Require(RiCrashLogger.Failures.IsEmpty, "no logged UI/STA exceptions")

        Using hidden As New SplashScreenCountDown("not shown")
            hidden.Dispose()
            Require(hidden.Completion.IsCompleted, "unshown splash cleanup is immediate")
        End Using
        Dim splash As New SplashScreenCountDown("visible probe", countdownSeconds:=5)
        Dim ready = splash.ShowAsync()
        PumpUntil(ready, 12000)
        ready.GetAwaiter().GetResult()
        splash.RestartCountdown(2, "updated")
        splash.Close()
        PumpUntil(splash.Completion, 5000)
        splash.Completion.GetAwaiter().GetResult()
        Require(Not splash.RequestCancellationToken.IsCancellationRequested, "normal splash close does not cancel request")
        splash.Dispose()
        Using works As New SplashScreenWorks("owner UI countdown", countdownSeconds:=2)
            works.Show()
            Application.DoEvents()
            works.RestartCountdown(1, "restarted")
            works.Close()
            Require(works.IsDisposed, "busy splash disposes on the owner UI thread")
        End Using
        ProgressBarModule.GlobalProgressMax = 100
        ProgressBarModule.CancelOperation = False
        Using progress As New ProgressForm("Progress lifecycle probe", "Starting")
            progress.Show()
            Application.DoEvents()
            progress.Close()
            Require(progress.IsDisposed, "progress form and its owned resources dispose on the owner thread")
        End Using
        Dim ocr = SharedMethods.CreateOcrStatusForProbe()
        Dim ocrReady = SharedMethods.StartOcrStatusForProbe(ocr, CancellationToken.None)
        PumpUntil(ocrReady, 12000)
        ocrReady.GetAwaiter().GetResult()
        ocr.Dispose()
        Dim ocrCompletion = SharedMethods.OcrCompletionForProbe(ocr)
        PumpUntil(ocrCompletion, 5000)
        ocrCompletion.GetAwaiter().GetResult()
        ocr.Dispose()
        Dim canceledOcr = SharedMethods.CreateOcrStatusForProbe()
        Dim canceledReady = SharedMethods.StartOcrStatusForProbe(canceledOcr, New CancellationToken(True))
        Try
            canceledReady.GetAwaiter().GetResult()
            Throw New Exception("Canceled OCR status unexpectedly started")
        Catch ex As OperationCanceledException
            Require(True, "already canceled OCR startup creates no UI")
        Finally
            canceledOcr.Dispose()
        End Try
        Require(SharedMethods.OcrCompletionForProbe(canceledOcr).IsCompleted, "never-started OCR cleanup completes")
        Require(RiCrashLogger.Failures.IsEmpty, "all status/splash lifecycles end without logged failures")
        Console.WriteLine("PASS " & checks.ToString() & " native Windows assertions")
    End Sub
End Module
