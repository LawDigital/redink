' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' 8.1.2026
'
' The compiled version of Red Ink also ...
'
' Includes DiffPlex in unchanged form; Copyright (c) 2023 Matthew Manela; licensed under the Apache-2.0 license (http://www.apache.org/licenses/LICENSE-2.0) at GitHub (https://github.com/mmanela/diffplex).
' Includes Newtonsoft.Json in unchanged form; Copyright (c) 2023 James Newton-King; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://www.newtonsoft.com/json
' Includes HtmlAgilityPack in unchanged form; Copyright (c) 2024 ZZZ Projects, Simon Mourrier,Jeff Klawiter,Stephan Grell; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://html-agility-pack.net/
' Includes Bouncycastle.Cryptography in unchanged form; Copyright (c) 2024 Legion of the Bouncy Castle Inc.; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://www.bouncycastle.org/download/bouncy-castle-c/
' Includes PdfPig in unchanged form; Copyright (c) 2024 UglyToad, EliotJones PdfPig, BobLd; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/UglyToad/PdfPig
' Includes MarkDig in unchanged form; Copyright (c) 2024 Alexandre Mutel; licensed under the BSD 2 Clause (Simplified) license (https://licenses.nuget.org/BSD-2-Clause) at https://github.com/xoofx/markdig
' Includes NAudio and components in unchanged form; Copyright (c) 2020 Mark Heath; licensed under a proprietary open source license (https://www.nuget.org/packages/NAudio/2.2.1/license) at https://github.com/naudio/NAudio
' Includes Vosk in unchanged form; Copyright (c) 2022 Alpha Cephei Inc.; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://alphacephei.com/vosk/
' Includes Whisper.net in unchanged form; Copyright (c) 2024 Sandro Hanea; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/sandrohanea/whisper.net
' Includes Grpc.core/Grpc.net in unchanged form; Copyright (c) 2023/2025 The gRPC Authors; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/grpc/grpc
' Includes Google Speech V1 library and related API libraries in unchanged form; Copyright (c) 2024 Google LLC; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/googleapis/google-cloud-dotnet
' Includes Google Protobuf in unchanged form; Copyright (c) 2025 Google Inc.; licensed under the BSD-3-Clause license (https://licenses.nuget.org/BSD-3-Clause) at https://github.com/protocolbuffers/protobuf
' Includes Google.Api in unchanged form; Copyright (c) 2025 Google LLC; licensed under the BSD-3-Clause license (https://licenses.nuget.org/BSD-3-Clause) at https://github.com/googleapis/gax-dotnet
' Includes Google.Apis in unchanged form; Copyright (c) 2025 Google LLC; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/googleapis/google-api-dotnet-client
' Includes Google.Longrunning in unchanged form; Copyright (c) 2025 Google LLC; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/googleapis/google-cloud-dotnet
' Includes MarkdownToRTF in modified form; Copyright (c) 2025 Gustavo Hennig; original licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/GustavoHennig/MarkdownToRtf
' Includes Nito.AsyncEx in unchanged form; Copyright (c) 2021 Stephen Cleary; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/StephenCleary/AsyncEx
' Includes NetOffice libraries in unchanged form; Copyright (c) 2020 Sebastian Lange, Erika LeBlanc; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/netoffice/NetOffice-NuGet
' Includes NAudio.Lame in unchanged form; Copyright (c) 2019 Corey Murtagh; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/Corey-M/NAudio.Lame
' Includes PdfiumViewer in unchanged form; Copyright (c) 2017 Pieter van Ginkel; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/pvginkel/PdfiumViewer
' Includes PDFsharp in unchanged form; Copyright (c) 2025 PDFSharp Team; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://docs.pdfsharp.net/
' Includes System.Interactive.Async in unchanged form; Copyright (c) 2025 by .NET Foundation and Contributors; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/dotnet/reactive
' Includes also various Microsoft distributables and libraries copyrighted by Microsoft Corporation and available, among others, under the Microsoft EULA, the Visual Studio Community 2022 License and the MIT License (including Microsoft.Bcl.*, Microsoft.Extensions.*, System.*, System.Security.*, System.CodeDom, DocumentFormat.OpenXml.*, Microsoft.ml.*, CommunityToolkit.HighPerformance licensed under MIT License) (https://licenses.nuget.org/MIT); Copyright (c) 2016- Microsoft Corp.'
'
' Licenses of Red Ink and of third-party components and further legal terms/notices are available in the installation folder and via https://redink.ai.
'
' Documentation for developers: See at the end of this file, throughout the code and the manual (https://redink.ai).


Option Explicit On
Option Strict Off

Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Outlook
Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Interop.Word
Imports SharedLibrary.SharedLibrary

Partial Public Class ThisAddIn

    ' Hardcoded config values

    Public Const AN As String = "Red Ink"
    Public Const AN2 As String = "red_ink"
    Public Const AN6 As String = "Inky"

    Public Shared Version As String = "V.080126" & SharedMethods.VersionQualifier

    Public Const ShortenPercent As Integer = 20
    Public Const SummaryPercent As Integer = 20
    Private Const NetTrigger As String = "(net)"
    Private Const LibTrigger As String = "(Lib)"
    Private Const MarkupPrefix As String = "Markup:"
    Private Const MarkupPrefixDiff As String = "MarkupDiff:"
    Private Const MarkupPrefixDiffW As String = "MarkupDiffW:"
    Private Const MarkupPrefixWord As String = "MarkupWord:"
    Private Const MarkupPrefixAll As String = "Markup[Diff|DiffW|Word]:"
    Private Const ClipboardPrefix As String = "Clipboard:"
    Private Const ClipboardPrefix2 As String = "Clip:"
    Private Const InsertPrefix As String = "Insert:"
    Private Const MyStyleTrigger As String = "(mystyle)"
    Private Const NoFormatTrigger As String = "(noformat)"
    Private Const NoFormatTrigger2 As String = "(nf)"
    Private Const KFTrigger As String = "(keepformat)"
    Private Const KFTrigger2 As String = "(kf)"
    Private Const KPFTrigger As String = "(keepparaformat)"
    Private Const KPFTrigger2 As String = "(kpf)"
    Private Const InPlacePrefix As String = "Replace:"
    Private Const NewDocPrefix As String = "Newdoc:"
    Private Const ObjectTrigger2 As String = "(clip)"

    Private Const ESC_KEY As Integer = &H1B

    Private Const SecondAPICode As String = "(2nd)"

    ' Variables that are available to InterpolateAtRuntime

    Public TranslateLanguage As String = ""
    Public OtherPrompt As String = ""
    Public CurrentDate As String = "(Current Date: " & DateTime.Now.ToString("dd-MMM-yyyy", CultureInfo.GetCultureInfo("en-US")) & ")"
    Public Username As String = ""
    Public MyStyleInsert As String = ""
    Public ShortenLength, SummaryLength As Long
    Public DateTimeNow As String

    Public InspectorOpened As Boolean = False


    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function FindWindow(
                                ByVal lpClassName As String,
                                ByVal lpWindowName As String
                            ) As IntPtr
    End Function

    ''' <summary>
    ''' Indicates whether delayed startup tasks have completed (set after activation path).
    ''' </summary>
    Public StartupInitialized As Boolean = False

    ''' <summary>
    ''' Hidden control created to obtain a Windows Forms handle for marshaling to the Outlook UI thread.
    ''' </summary>
    Private mainThreadControl As New System.Windows.Forms.Control()

    ''' <summary>
    ''' Outlook Explorer instance used to hook activation events for deferred initialization.
    ''' </summary>
    Private WithEvents outlookExplorer As Outlook.Explorer

    ''' <summary>
    ''' SynchronizationContext captured at construction for potential use (unused below where _uiContext is set).
    ''' </summary>
    Private ReadOnly uiCtx As System.Threading.SynchronizationContext =
        System.Threading.SynchronizationContext.Current

    ''' <summary>
    ''' Captured UI SynchronizationContext after handle creation; used for thread marshaling.
    ''' </summary>
    Private Shared _uiContext As SynchronizationContext

    ''' <summary>
    ''' TaskScheduler corresponding to the captured UI context.
    ''' </summary>
    Private Shared _uiScheduler As TaskScheduler

    ''' <summary>
    ''' Interlocked flag ensuring DelayedStartupTasks executes only once.
    ''' </summary>
    Private delayedStartupOnce As Integer = 0

    ''' <summary>
    ''' Outlook add-in startup handler. Captures UI context, initializes UpdateHandler handles, obtains host HWND, sets explorer activation path, and restores last chat id.
    ''' </summary>
    Private explorers As Outlook.Explorers

    ''' <summary>
    ''' Fallback timer to trigger delayed startup if no Activate event occurs.
    ''' </summary>
    Private startupFallbackTimer As System.Windows.Forms.Timer

    ''' <summary>
    ''' Handles add-in startup. Initializes UI synchronization, UpdateHandler targets, host window handle, Explorer hooks, fallback timer, and restores last chat id.
    ''' </summary>
    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Try
            RemoveHandler Microsoft.Win32.SystemEvents.PowerModeChanged, AddressOf OnPowerModeChanged
        Catch
        End Try

        StartPowerWatch()

        ' Necessary for Update Handler to work correctly
        ' 1) Force the creation of the Control's handle on the Office UI thread
        Dim dummy = mainThreadControl.Handle

        ' 2) Capture synchronization context & scheduler exactly once (after handle exists)
        _uiContext = SynchronizationContext.Current
        If _uiContext Is Nothing Then
            _uiContext = New WindowsFormsSynchronizationContext()
            SynchronizationContext.SetSynchronizationContext(_uiContext)
        End If
        _uiScheduler = TaskScheduler.FromCurrentSynchronizationContext()

        ' 3) Give that Control to the UpdateHandler so it can Invoke on it
        UpdateHandler.MainControl = mainThreadControl

        ' 4) Capture the host window’s HWND (Word / Excel / Outlook)
        Dim hwnd As IntPtr
        Dim progId = Me.Application.GetType().Name.ToLowerInvariant()
        If progId.Contains("word") OrElse progId.Contains("excel") Then
            hwnd = New IntPtr(CInt(Me.Application.Hwnd))
        Else
            hwnd = FindWindow("rctrl_renwnd32", Nothing)
        End If
        UpdateHandler.HostHandle = hwnd

        ' Other tasks that need to be done at startup
        mainThreadControl.CreateControl()

        ' Hook Explorers collection early (fires before Activate if window created but not focused)
        Try
            explorers = Application.Explorers
            If explorers IsNot Nothing Then
                AddHandler explorers.NewExplorer, AddressOf Explorers_NewExplorer
            End If
        Catch
        End Try

        outlookExplorer = ComRetry(Function() Application.ActiveExplorer())
        If outlookExplorer IsNot Nothing Then
            ' If already available, keep existing Activate path but also start fallback timer in case Activate never fires
            AddHandler outlookExplorer.Activate, AddressOf Explorer_Activate
            StartStartupFallbackTimer()
        Else
            ' No explorer yet – start fallback timer + rely on NewExplorer event
            StartStartupFallbackTimer()
        End If

        Try
            activeChatId = If(My.Settings.Inky_LastChat = 2, 2, 1)
        Catch
            activeChatId = 1
        End Try
    End Sub

    ''' <summary>
    ''' Handles creation of a new Explorer window. Attaches Activate, marks initialized, runs delayed startup, and cleans handlers.
    ''' </summary>
    ''' <param name="Explorer">The new Outlook Explorer instance.</param>
    Private Sub Explorers_NewExplorer(ByVal Explorer As Outlook.Explorer)
        ' Run delayed tasks ASAP when the first Explorer object is created (before user interaction)
        Try
            If Not StartupInitialized Then
                outlookExplorer = Explorer
                Try
                    AddHandler outlookExplorer.Activate, AddressOf Explorer_Activate
                Catch
                End Try
                StartupInitialized = True
                DelayedStartupTasks()
                CleanupStartupHandlers()
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Explorer activation callback. Marks startup initialized, removes handler and executes delayed tasks.
    ''' </summary>
    Private Sub Explorer_Activate()
        StartupInitialized = True
        DelayedStartupTasks()
        CleanupStartupHandlers()
    End Sub

    ''' <summary>
    ''' Starts a fallback timer to execute delayed startup if activation does not occur.
    ''' </summary>
    ''' <param name="ms">Timer interval in milliseconds. Default is 3000.</param>
    Private Sub StartStartupFallbackTimer(Optional ms As Integer = 3000)
        If startupFallbackTimer IsNot Nothing Then Return
        startupFallbackTimer = New System.Windows.Forms.Timer() With {.Interval = ms}
        AddHandler startupFallbackTimer.Tick,
        Sub()
            Try
                startupFallbackTimer.Stop()
                startupFallbackTimer.Dispose()
            Catch
            End Try
            startupFallbackTimer = Nothing

            ' IMPORTANT: Do NOT set delayedStartupOnce here.
            ' Let DelayedStartupTasks() decide single-run via its own Interlocked check.
            Try
                StartupInitialized = True
                DelayedStartupTasks()
                CleanupStartupHandlers()
            Catch
            End Try
        End Sub
        Try
            startupFallbackTimer.Start()
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Removes activation and NewExplorer handlers and disposes the fallback timer.
    ''' </summary>
    Private Sub CleanupStartupHandlers()
        Try
            If outlookExplorer IsNot Nothing Then
                RemoveHandler outlookExplorer.Activate, AddressOf Explorer_Activate
            End If
        Catch
        End Try
        Try
            If explorers IsNot Nothing Then
                RemoveHandler explorers.NewExplorer, AddressOf Explorers_NewExplorer
            End If
        Catch
        End Try
        Try
            If startupFallbackTimer IsNot Nothing Then
                startupFallbackTimer.Stop()
                startupFallbackTimer.Dispose()
                startupFallbackTimer = Nothing
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Deferred initialization routine (single execution) performing configuration load, ribbon refresh, update check start, watchdog and HTTP listener startup.
    ''' </summary>
    Private Sub DelayedStartupTasks()
        ' Run once even if scheduled twice (e.g., event + BeginInvoke)
        If System.Threading.Interlocked.CompareExchange(delayedStartupOnce, 1, 0) <> 0 Then Return

        Try
            InitializeConfig(True, True)
            UpdateHandler.PeriodicCheckForUpdates(INI_UpdateCheckInterval, "Outlook", INI_UpdatePath, _context)
            Dim result = Globals.Ribbons.Ribbon1.UpdateRibbon()
            result = Globals.Ribbons.Ribbon2.UpdateRibbon()
            mainThreadControl.CreateControl()
            StartListenerWatchdog()
            StartupHttpListener()
        Catch ex As System.Exception
            ' Handling errors gracefully
        End Try
    End Sub

    ''' <summary>
    ''' Outlook add-in shutdown handler. Sequentially stops HTTP listener, watchdog, and power watch components.
    ''' </summary>
    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        ' 1) deterministically stop the HTTP listener (await synchronously)
        Try
            Dim t As System.Threading.Tasks.Task = ShutdownHttpListener()
            t.GetAwaiter().GetResult() ' safe: our shutdown continuations don’t capture the UI context
        Catch ex As System.Exception
            System.Diagnostics.Debug.WriteLine("ShutdownHttpListener failed: " & ex.Message)
        End Try

        ' 2) stop watchdog (if you added it)
        Try
            StopListenerWatchdog()
        Catch
        End Try

        ' 3) tear down power notifications window
        Try
            StopPowerWatch()
        Catch
        End Try

    End Sub

    ' Lightweight UI switch helper (matches Excel version)
    ''' <summary>
    ''' Ensures execution continuity on the captured UI thread by posting a no-op if current context differs.
    ''' </summary>
    Private Shared Function EnsureUIThread() As System.Threading.Tasks.Task
        If _uiContext Is Nothing OrElse SynchronizationContext.Current Is _uiContext Then
            Return System.Threading.Tasks.Task.CompletedTask
        End If
        Dim tcs As New TaskCompletionSource(Of Object)()
        _uiContext.Post(
            Sub(state As Object)
                tcs.SetResult(Nothing)
            End Sub,
            Nothing)
        Return tcs.Task
    End Function

    ''' <summary>
    ''' Schedules an Action for execution on the UI thread using control invocation and returns an awaitable Task for completion.
    ''' </summary>
    Private Function SwitchToUi(uiAction As System.Action) _
        As System.Threading.Tasks.Task

        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of Object)()

        mainThreadControl.BeginInvoke(New MethodInvoker(
        Sub()
            Try
                uiAction.Invoke()
                tcs.SetResult(Nothing)
            Catch ex As System.Exception
                tcs.SetException(ex)
            End Try
        End Sub))

        Return tcs.Task
    End Function

    ''' <summary>
    ''' Schedules a Func(Of T) on the UI thread and returns a Task(Of T) representing the result.
    ''' </summary>
    Private Function SwitchToUi(Of T)(uiFunc As System.Func(Of T)) _
        As System.Threading.Tasks.Task(Of T)

        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of T)()

        mainThreadControl.BeginInvoke(New MethodInvoker(
        Sub()
            Try
                tcs.SetResult(uiFunc.Invoke())
            Catch ex As System.Exception
                tcs.SetException(ex)
            End Try
        End Sub))

        Return tcs.Task
    End Function


    ''' <summary>
    ''' Provides temporary OLE message filtering to adjust retry behavior for transient COM call rejections while preserving any previous filter.
    ''' </summary>
    Friend NotInheritable Class OleMessageFilter
        <System.Runtime.InteropServices.DllImport("ole32.dll")>
        Private Shared Function CoRegisterMessageFilter(newFilter As IOleMessageFilter, ByRef oldFilter As IOleMessageFilter) As Integer
        End Function

        <System.Runtime.InteropServices.ComImport(),
         System.Runtime.InteropServices.Guid("00000016-0000-0000-C000-000000000046"),
         System.Runtime.InteropServices.InterfaceType(System.Runtime.InteropServices.ComInterfaceType.InterfaceIsIUnknown)>
        Private Interface IOleMessageFilter
            <System.Runtime.InteropServices.PreserveSig()>
            Function HandleInComingCall(dwCallType As Integer,
                                        hTaskCaller As IntPtr,
                                        dwTickCount As Integer,
                                        lpInterfaceInfo As IntPtr) As Integer
            <System.Runtime.InteropServices.PreserveSig()>
            Function RetryRejectedCall(hTaskCallee As IntPtr,
                                       dwTickCount As Integer,
                                       dwRejectType As Integer) As Integer
            <System.Runtime.InteropServices.PreserveSig()>
            Function MessagePending(hTaskCallee As IntPtr,
                                    dwTickCount As Integer,
                                    dwPendingType As Integer) As Integer
        End Interface

        ' Keep a reference to the filter that Outlook installed before ours,
        ' so we can forward to it and restore it later.
        Private Shared prevFilter As IOleMessageFilter = Nothing
        Private Shared registered As Boolean

        Private Class Filter
            Implements IOleMessageFilter

            Public Function HandleInComingCall(dwCallType As Integer, hTaskCaller As IntPtr, dwTickCount As Integer, lpInterfaceInfo As IntPtr) As Integer Implements IOleMessageFilter.HandleInComingCall
                If prevFilter IsNot Nothing Then
                    Try : Return prevFilter.HandleInComingCall(dwCallType, hTaskCaller, dwTickCount, lpInterfaceInfo) : Catch : End Try
                End If
                Return 0 ' SERVERCALL_ISHANDLED
            End Function

            Public Function RetryRejectedCall(hTaskCallee As IntPtr, dwTickCount As Integer, dwRejectType As Integer) As Integer Implements IOleMessageFilter.RetryRejectedCall
                ' Ask Outlook’s filter first
                Dim prevRet As Integer = -1
                If prevFilter IsNot Nothing Then
                    Try : prevRet = prevFilter.RetryRejectedCall(hTaskCallee, dwTickCount, dwRejectType) : Catch : prevRet = -1 : End Try
                End If

                ' Only adjust RETRYLATER if Outlook would cancel (-1)
                If dwRejectType = 2 Then ' SERVERCALL_RETRYLATER
                    If prevRet >= 0 Then Return prevRet
                    Return 150 ' retry after 150ms
                End If

                ' For all other cases, preserve Outlook’s behavior
                Return prevRet
            End Function

            Public Function MessagePending(hTaskCallee As IntPtr, dwTickCount As Integer, dwPendingType As Integer) As Integer Implements IOleMessageFilter.MessagePending
                If prevFilter IsNot Nothing Then
                    Try : Return prevFilter.MessagePending(hTaskCallee, dwTickCount, dwPendingType) : Catch : End Try
                End If
                Return 2 ' PENDINGMSG_WAITDEFPROCESS
            End Function
        End Class

        ''' <summary>
        ''' Registers the custom message filter and stores any existing filter (single registration guard).
        ''' </summary>
        Public Shared Sub Register()
            If registered Then Return
            Dim oldF As IOleMessageFilter = Nothing
            ' Register our filter and capture the previous (Outlook’s) filter
            CoRegisterMessageFilter(New Filter(), oldF)
            prevFilter = oldF
            registered = True
        End Sub

        ''' <summary>
        ''' Restores the previously registered Outlook filter and clears internal references.
        ''' </summary>
        Public Shared Sub Revoke()
            If Not registered Then Return
            Dim oldF As IOleMessageFilter = Nothing
            ' Restore Outlook’s original filter (do NOT set Nothing here)
            CoRegisterMessageFilter(prevFilter, oldF)
            prevFilter = Nothing
            registered = False
        End Sub
    End Class

    ''' <summary>
    ''' Enables the OLE message filter for at least 500 ms or the specified duration, then revokes it via a timer callback.
    ''' </summary>
    Private Sub EnableOleFilterFor(durationMs As Integer)
        ' must run on the Outlook UI thread
        Dim t As New System.Windows.Forms.Timer() With {.Interval = Math.Max(500, durationMs)}
        AddHandler t.Tick,
            Sub()
                Try : OleMessageFilter.Revoke() : Catch : End Try
                Try : t.Stop() : t.Dispose() : Catch : End Try
            End Sub
        Try : OleMessageFilter.Register() : Catch : End Try
        t.Start()
    End Sub

    ''' <summary>
    ''' Executes a delegate with up to three retry attempts on specific transient COM exceptions (RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER, E_FAIL).
    ''' </summary>
    Private Shared Function ComRetry(Of T)(work As System.Func(Of T)) As T
        For i As Integer = 0 To 2
            Try
                Return work()
            Catch ex As System.Runtime.InteropServices.COMException When _
            ex.HResult = &H80010001 OrElse   ' RPC_E_CALL_REJECTED
            ex.HResult = &H8001010A OrElse   ' RPC_E_SERVERCALL_RETRYLATER
            ex.HResult = &H80004005          ' E_FAIL (some busy states)
                ' Avoid Application.DoEvents here to prevent re-entrancy into COM/Ribbon
                System.Threading.Thread.Sleep(150)
            End Try
        Next
        Return work() ' last try to surface real error
    End Function

    ''' <summary>
    ''' Invokes an asynchronous function on the UI thread and propagates its completion, fault, or cancellation status to a returned Task(Of T).
    ''' </summary>
    Private Function SwitchToUiTask(Of T)(
        uiFunc As System.Func(Of System.Threading.Tasks.Task(Of T))) _
        As System.Threading.Tasks.Task(Of T)

        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of T)()

        mainThreadControl.BeginInvoke(New MethodInvoker(
        Sub()
            Try
                Dim inner As System.Threading.Tasks.Task(Of T) = uiFunc.Invoke()
                inner.ContinueWith(
                    Sub(taskObj)
                        If taskObj.IsFaulted Then
                            tcs.SetException(taskObj.Exception.InnerExceptions)
                        ElseIf taskObj.IsCanceled Then
                            tcs.SetCanceled()
                        Else
                            tcs.SetResult(taskObj.Result)
                        End If
                    End Sub,
                    System.Threading.Tasks.TaskScheduler.Default)
            Catch ex As System.Exception
                tcs.SetException(ex)
            End Try
        End Sub))

        Return tcs.Task
    End Function


    ' Bridge to SharedLibrary 

    ''' <summary>
    ''' Initializes shared configuration state via SharedMethods and sets descriptive version identifier.
    ''' </summary>
    Public Sub InitializeConfig(FirstTime As Boolean, Reload As Boolean)
        _context.InitialConfigFailed = False
        _context.RDV = "Outlook (" & Version & ")"
        SharedMethods.InitializeConfig(_context, FirstTime, Reload)
    End Sub

    ''' <summary>
    ''' Returns True if required INI values are missing according to SharedMethods.
    ''' </summary>
    Private Function INIValuesMissing() As Boolean
        Return SharedMethods.INIValuesMissing(_context)
    End Function

    ''' <summary>
    ''' Asynchronously posts correction processing to SharedMethods and returns the resulting String.
    ''' </summary>
    Public Shared Async Function PostCorrection(inputText As String, Optional ByVal UseSecondAPI As Boolean = False) As Task(Of String)
        Return Await SharedMethods.PostCorrection(_context, inputText, UseSecondAPI)
    End Function

    ''' <summary>
    ''' Asynchronously sends prompts to SharedMethods.LLM and optionally ensures UI thread affinity before returning the response.
    ''' </summary>
    Public Shared Async Function LLM(ByVal promptSystem As String, ByVal promptUser As String, Optional ByVal Model As String = "", Optional ByVal Temperature As String = "", Optional ByVal Timeout As Long = 0, Optional ByVal UseSecondAPI As Boolean = False, Optional HideSplash As Boolean = False, Optional ByVal AddUserPrompt As String = "", Optional ByVal FileObject As String = "", Optional cancellationToken As Threading.CancellationToken = Nothing, Optional EnsureUI As Boolean = True) As Task(Of String)
        Dim Response = Await SharedMethods.LLM(_context, promptSystem, promptUser, Model, Temperature, Timeout, UseSecondAPI, HideSplash, AddUserPrompt, FileObject, cancellationToken)
        If EnsureUI Then Await EnsureUIThread().ConfigureAwait(False)
        Return Response
    End Function

    ''' <summary>
    ''' Shows the settings window after ensuring configuration has been loaded; triggers delayed initialization if required.
    ''' </summary>
    Private Sub ShowSettingsWindow(Settings As Dictionary(Of String, String), SettingsTips As Dictionary(Of String, String))
        If Not INIloaded Then
            If Not StartupInitialized Then
                Try
                    DelayedStartupTasks()
                    RemoveHandler outlookExplorer.Activate, AddressOf Explorer_Activate
                Catch ex As System.Exception
                End Try
                If Not INIloaded Then Return
            Else
                InitializeConfig(False, False)
                If Not INIloaded Then
                    Return
                End If
            End If
        End If
        SharedMethods.ShowSettingsWindow(Settings, SettingsTips, _context)
    End Sub

    ''' <summary>
    ''' Opens the prompt selector via SharedMethods and returns a tuple (String, Boolean, Boolean, Boolean) with its selection result.
    ''' </summary>
    Private Function ShowPromptSelector(filePath As String, filePathlocal As String, enableMarkup As Boolean, enableBubbles As Boolean) As (String, Boolean, Boolean, Boolean)
        Return SharedMethods.ShowPromptSelector(filePath, filePathlocal, enableMarkup, enableBubbles, _context)
    End Function


End Class

''' <summary>
''' Declares GetAsyncKeyState for polling asynchronous keyboard state (per virtual key code).
''' </summary>
Module GetAsyncKeyStateModule
    ' Correct attribute declaration for DllImport
    ''' <summary>
    ''' Retrieves state for the specified virtual key; return value contains transition and key press information.
    ''' </summary>
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Function GetAsyncKeyState(ByVal vKey As Integer) As Short
    End Function
End Module

' =====================================================================================
' Red Ink for Outlook – Architectural Overview (Reviewer Documentation)
' =====================================================================================
' PURPOSE
'   AI-assisted authoring add-in for Outlook (with sibling add-ins for Word/Excel).
'   Provides translation, summarization, rewriting, style application, freestyle prompting,
'   markup/diff visualization, email-chain analysis, clipboard/object insertion, and
'   configurable prompt library usage via local and alternate LLM endpoints.
'
' CORE COMPONENTS (PARTIAL CLASS SPLIT)
'   ThisAddIn.vb
'       - Outlook/VSTO lifecycle (Startup/Shutdown) with a two-phase initialization:
'           1. Early: capture UI SynchronizationContext + TaskScheduler, host HWND, set
'              UpdateHandler.MainControl + HostHandle.
'           2. Delayed: configuration load, ribbon refresh, update polling, HTTP listener start.
'       - Global constants: product naming, version, command triggers (e.g. Markup:, Clip:, Insert:, Replace:, Newdoc:, (nf)/(kf)/(kpf)).
'       - Mutable runtime variables (TranslateLanguage, ShortenLength, SummaryLength, Username, etc.)
'       - COM robustness:
'           * OleMessageFilter: temporary registration to auto-retry RPC_E_CALL_REJECTED / RETRYLATER.
'           * ComRetry(): small exponential-ish retry for transient COM busy states.
'       - UI thread marshaling:
'           * EnsureUIThread(), SwitchToUi(Action/Func), SwitchToUiTask(Func(Of Task(Of T)))
'           * Allows asynchronous background LLM calls while keeping final UI mutation safe.
'       - External bridge to SharedLibrary.SharedMethods for:
'           * InitializeConfig(), LLM(), PostCorrection(), Settings window, Prompt selector.
'       - Startup-added watchdogs: power mode watcher, listener watchdog, update checker.
'       - Shutdown order: graceful HTTP listener stop, watchdog stop, power watcher teardown.
'
'   ThisAddIn.Commands.vb
'       - Entry command dispatcher: MainMenu(RI_Command) routes ribbon actions.
'       - Email-context resolution:
'           * Supports inline response → optional forced Inspector promotion with selection reapplication.
'           * Multi-selection aggregation for chain summarization.
'       - High-level operations:
'           * Translate / PrimLang / Correct / Summarize / Improve / NoFillers / ApplyMyStyle /
'             Friendly / Convincing / Shorten / Sumup / Answers / Freestyle / InsertClipboard.
'       - Freestyle engine:
'           * Prefix parsing for behavior flags: Markup*, Replace:, Clipboard:, Clip:, Newdoc:, (net), (Lib),
'             (mystyle), (nf)/(kf)/(kpf), (clip) object inclusion, (2nd) alternate model.
'           * Dynamic prompt augmentation (default prefix injection).
'       - Markup strategies:
'           * Method 1: Word built-in compare.
'           * Method 2: Diff (DiffPlex).
'           * Method 3: Diff rendered in window (no inline insertion).
'           * Cap enforcement (MarkupDiffCap) with user override prompt.
'       - Formatting retention:
'           * Optional HTML extraction + reinsertion (KeepFormat flags, cap INI_KeepFormatCap).
'           * Markdown conversion of inline formatting when enabled (MarkdownConvert).
'       - Clipboard insertion:
'           * Robust multi-attempt STA clipboard setter with fallback to manual window or temp file.
'           * Optional RTF conversion (MarkdownToRtf).
'       - MyStyle:
'           * Style prompt file selection, automatic AI-based profile generation and persistence.
'       - Chain parsing heuristics:
'           * GetLatestMailBody(): detects quoted sections using marker & header patterns.
'
'   ThisAddIn.Helpers.vb (naming suggests; verify implementation)
'       - Expected to host shared helper routines used across command code:
'           * Text diff rendering, CompareAndInsertText / CompareAndInsertTextCompareDocs
'           * HTML/Markdown conversion (RemoveHTML, ConvertRangeToMarkdown, InsertTextWithMarkdown,
'             GetRangeHtml).
'           * Utility wrappers for dialogs (ShowCustomMessageBox, ShowCustomWindow, Yes/No boxes).
'           * Configuration override logic (Override()) and selection length retrieval.
'       - Reviewer focus: ensure no unsafe string injection into COM automation or HTML; validate
'         any file I/O or reflection usage.
'
'   ThisAddIn.Processing.vb 
'       - Background processing concerns:
'           * HTTP listener startup (StartupHttpListener/ShutdownHttpListener), watchdog timers,
'             periodic update check (UpdateHandler.PeriodicCheckForUpdates).
'           * Post/Pre-correction pipelines (cleaning, normalization).
'       - Reviewer focus: authentication/authorization of HTTP surface, port binding, exposure risk,
'         timeout handling, cancellation token usage.
'
'   ThisAddIn.Properties.vb 
'       - Used to share configuration with SharedLibrary.
'       - Lightweight computed properties and wrappers mapping INI/config values to runtime flags.
'       - Reviewer focus: thread safety for shared state; ensure no implicit cross-thread access.
'
'   ThisAddIn.WebExtension.vb / ThisAddIn.WebExtension.FileHelpers.vb
'       - Bridge to a local web extension (e.g., for UI panel or external integration).
'       - FileHelpers: safe file access (prompt libraries, MyStyle storage, temp export).
'           * Should sanitize paths (environment variable expansion), enforce allowed directories,
'             handle large file sizes, avoid blocking UI thread.
'       - WebExtension core: request routing, minimal protocol, likely JSON/HTTP.
'       - Reviewer focus: validate no arbitrary file read/write, restrict external origin access,
'         ensure request size limits, avoid code injection via prompt content.
'
'   Resources
'       - Embedded assets: icons (ribbon), templates, prompt library defaults, language strings.
'       - Reviewer focus: check for hardcoded secrets, mutable resources used as dynamic prompts,
'         localization fallback behavior.
'
'   Ribbon1.vb (and Ribbon2 if present)
'       - UI layer: button click handlers calling MainMenu or helper methods (e.g., ShowSettings, HelpMeInky).
'       - Dynamic enable/disable based on config (INIloaded, GPTSetupError, selection context).
'       - Reviewer focus: ensure no long-running work on UI callbacks; all heavy LLM calls are async.
'
' EXTERNAL DEPENDENCIES (FOR REVIEW)
'   DiffPlex: text diff generation (verify unmodified usage, license Apache 2.0).
'   Markdig: Markdown → HTML conversion (pipeline configured with advanced extensions; review for XSS if HTML rendered in custom windows).
'   Newtonsoft.Json / Google Protobuf / gRPC libs: serialization / API calls (audit for external traffic).
'   HtmlAgilityPack: HTML parsing (ensure safe usage for user-controlled HTML).
'   Cryptography / Pdf libraries (BouncyCastle, PdfPig, PdfiumViewer): not directly shown here—confirm constrained usage.
'   Whisper.net / Vosk / NAudio: speech/transcription modules (check if loaded lazily; resource cleanup).
'   SharedLibrary.SharedMethods: central abstraction for LLM calls, config reading, UI dialogs; treat as trust boundary (audit separately).
'
' CONCURRENCY & THREADING
'   - UI thread affinity enforced via captured SynchronizationContext.
'   - Asynchronous Tasks for LLM requests; ConfigureAwait(False) used to prevent deadlocks.
'   - Interlocked guards (e.g., inMainMenu, delayedStartupOnce) prevent reentrancy.
'   - Timer-based OLE filter revocation to bound filter lifetime.
'
' SECURITY / REVIEW HOTSPOTS
'   1. Prompt Injection: Freestyle and user-supplied prompts are concatenated into system prompts—validate escaping where model calls depend on structured markup (<TEXTTOPROCESS> tags).
'   2. HTML / Markdown Rendering: Output into windows (ShowHTMLCustomMessageBox) → ensure no script execution (Markdig produces HTML; confirm viewer control neutralizes scripts).
'   3. File I/O: MyStyle prompt storage & prompt library reading—check path validation and absence of directory traversal.
'   4. HTTP Listener: Confirm authentication, port binding restrictions, and rejection of unsolicited external requests.
'   5. Clipboard: Large data insertion & RTF conversion—ensure size caps to avoid memory pressure.
'   6. COM Automation: Robust retry logic prevents crashes but could mask persistent failures—log repeated COMException patterns.
'   7. Alternate Models: When switching (2nd API / model selection), verify restoration of original config (RestoreDefaults) always succeeds.
'   8. Diff Cap (MarkupDiffCap): Performance safeguard—ensure enforced consistently for both Command_InsertAfter and Freestyle flows.
'
' QUICK TRACE POINTS
'   Startup path: ThisAddIn_Startup → (Explorer_Activate OR BeginInvoke DelayedStartupTasks) → InitializeConfig → UpdateHandler.PeriodicCheckForUpdates → StartupHttpListener.
'   Command flow (example): Ribbon click → MainMenu("Correct") → Command_InsertAfter(prompt...) → LLM() → (optional PostCorrection) → insertion + markup.
''
' =====================================================================================
' End of Reviewer Documentation
' =====================================================================================