' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.vb
' Purpose:
'   Main Outlook VSTO add-in entry point. Owns startup/shutdown, Explorer lifecycle,
'   configuration initialization, UI-thread marshaling, and host-level services.
'
' Architecture:
'   Root partial ThisAddIn lifecycle module for the Outlook host. It coordinates
'   shared configuration/resources, Outlook event hooks, COM/UI safety, and deferred
'   warm-up while specialized mail, tooling, web, M365, and AutoPilot logic lives in
'   the other ThisAddIn.* files.
' =============================================================================
'
' 23.8.2026
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
' Includes Google Speech V1/V2 library and related API libraries in unchanged form; Copyright (c) 2024 Google LLC; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/googleapis/google-cloud-dotnet
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
' Includes Microsoft.Playwright (Playwright for .NET) in unchanged form; Copyright (c) 2020 Darío Kondratiuk and other contributors, with modifications copyright (c) Microsoft Corporation where stated in the source files; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/microsoft/playwright-dotnet
' Includes also various Microsoft distributables and libraries copyrighted by Microsoft Corporation and available, among others, under the Microsoft EULA, the Visual Studio Community 2022 License, the Microsoft.Web.WebView2 License (for Microsoft.Web.WebView2, see license on https://www.nuget.org/packages/Microsoft.Web.WebView2/ and below) and the MIT License (including Microsoft.Bcl.*, Microsoft.Extensions.*, Microsoft.Identity.Client, Microsoft.Identity.Client.Extensions.Msal, System.*, System.Security.*, System.CodeDom, DocumentFormat.OpenXml.*, Microsoft.ml.*, CommunityToolkit.HighPerformance licensed under MIT License) (https://licenses.nuget.org/MIT); Copyright (c) 2016- Microsoft Corp.
'
' Licenses of Red Ink and of third-party components and further legal terms/notices are available in the installation folder and via https://redink.ai.
'
' Documentation for developers: See at the end of this file, throughout the code and the manual (https://redink.ai).


Option Explicit On
Option Strict Off

Imports System.Diagnostics
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
    Public Const AN5 As String = "RI"
    Public Const AN6 As String = "Inky"
    Public Const AN4 As String = "redink_"
    Public Const AN3 As String = "redink"

    Public Shared Version As String = "V.230826" & SharedMethods.VersionQualifier

    Public Const ShortenPercent As Integer = 20
    Public Const SummaryPercent As Integer = 20
    Private Const NetTrigger As String = "(net)"
    Private Const LibTrigger As String = "(lib)"
    Private Const MarkupPrefix As String = "Markup:"
    Private Const MarkupPrefixDiff As String = "MarkupDiff:"
    Private Const MarkupPrefixDiffW As String = "MarkupDiffW:"
    Private Const MarkupPrefixWord As String = "MarkupWord:"
    Private Const MarkupPrefixAll As String = "Markup[Diff|DiffW|Word|Approve]:"
    Private Const MarkupPrefixApprove As String = "MarkupApprove:"
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
    Private Const ToolTrigger As String = "(ag)"
    Private Const KBTrigger As String = "(kb)"
    Private Const AddmailTrigger As String = "(addmail)"

    Private Const ESC_KEY As Integer = &H1B

    Private Const SecondAPICode As String = "(2nd)"

    ' Variables that are available to InterpolateAtRuntime
    Public TranslateLanguage As String = ""
    Public SourceLanguage As String = ""
    Public OutputLanguage As String = ""
    Public OtherPrompt As String = ""
    Public Dictionary As String = ""
    Public Username As String = ""
    Public MyStyleInsert As String = ""
    Public ShortenLength, SummaryLength As Long
    Public DateTimeNow As String
    Public WebGrounding As String = ""
    Public PrivacyProtection As String = ""

    Public HostName As String = ""
    Public GuestName As String = ""
    Public TargetAudience As String = ""
    Public Duration As String = ""
    Public Language As String = ""
    Public DialogueContext As String = ""
    Public ExtraInstructions As String = ""

    Public InspectorOpened As Boolean = False

    Public ReadOnly Property CurrentDate As String
        Get
            Return "Current date (ISO 8601 yyyy-MM-dd): " &
               DateTime.Now.ToString("yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture)
        End Get
    End Property

    Public ReadOnly Property LocalTimeContext As String
        Get
            Dim localNow As System.DateTimeOffset = System.DateTimeOffset.Now
            Dim localTz As System.TimeZoneInfo = System.TimeZoneInfo.Local
            Dim localZoneName As String = If(
                localTz.IsDaylightSavingTime(System.DateTime.Now),
                localTz.DaylightName,
                localTz.StandardName)

            Return "Current local runtime time (ISO 8601): " &
                localNow.ToString("yyyy-MM-dd HH:mm:ss zzz", Globalization.CultureInfo.InvariantCulture) &
                $". Local timezone: {localZoneName} ({localTz.Id}). Unless the user explicitly asks for UTC or another timezone, use this local timezone for user-facing dates and times."
        End Get
    End Property


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
    ''' Arms the ask_user interactivity guard for Outlook. A live user is only present in
    ''' Local Chat / Web Agent runs. Unattended e-mail Scheduler and AutoPilot runs
    ''' (indicated by _apActive) must never block on ask_user. Idempotent; safe to call
    ''' more than once.
    ''' </summary>
    Private Sub ArmAskUserInteractivityGuard()
        SharedLibrary.Agents.AskUserTool.InteractivityProvider =
            Function() As Boolean
                ' Interactive only when a chat agent session is active and this is not an
                ' unattended AutoPilot / e-mail Scheduler execution.
                Return _chatAgentActive AndAlso Not _apActive
            End Function
    End Sub

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
    ''' Periodic offline license-counter timer for long-running Outlook sessions.
    ''' </summary>
    Private licenseCounterTimer As System.Threading.Timer

    ''' <summary>
    ''' Shared UI synchronization context used by the shared tooling loop.
    ''' </summary>
    Public Shared UiSyncContext As System.Threading.SynchronizationContext

    ''' <summary>
    ''' Managed thread id of the Outlook UI thread captured at startup.
    ''' </summary>
    Public Shared UiThreadId As Integer

    ''' <summary>
    ''' Handles add-in startup. Initializes UI synchronization, UpdateHandler targets, host window handle, Explorer hooks, fallback timer, and restores last chat id.
    ''' </summary>
    Private Sub ThisAddIn_Startup() Handles Me.Startup

        ' Crash diagnostics: enabled only when the user's My.Settings.CrashLog is True.
        ' This flag is reconciled from INI_Crashlog after config load, so it takes effect
        ' on the next host launch. When False, no handlers are installed (zero overhead).
        Try
            RiCrashLogger.Initialize(
                "RedInk Outlook Add-in",
                Me.GetType().Assembly,
                My.Settings.CrashLog,
                True,
                RDV)
        Catch
        End Try

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

        UiSyncContext = _uiContext
        UiThreadId = System.Threading.Thread.CurrentThread.ManagedThreadId

        ' Arm the ask_user interactivity guard so unattended e-mail Scheduler / AutoPilot
        ' runs never block on a modal (interactive only in Local Chat / Web Agent).
        ArmAskUserInteractivityGuard()

        _uiScheduler = TaskScheduler.FromCurrentSynchronizationContext()

        ' 3) Give that Control to the UpdateHandler so it can Invoke on it
        UpdateHandler.MainControl = mainThreadControl

        ' 4) Capture the host window’s HWND.
        Dim hwnd As IntPtr = FindWindow("rctrl_renwnd32", Nothing)
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

#If DEBUG Then
        RunPythonExecuteRepairAdvisorSelfTestsAtStartup()
#End If

    End Sub


#If DEBUG Then
    Private Shared _pythonExecuteRepairAdvisorSelfTestsRan As Boolean = False

    ''' <summary>
    ''' DEBUG-only: runs the PythonExecuteRepairAdvisor self-tests once per process on a background
    ''' thread and writes the outcome to the Visual Studio Output window (Debug pane). No UI is shown;
    ''' this never runs in Release builds.
    ''' </summary>
    Private Sub RunPythonExecuteRepairAdvisorSelfTestsAtStartup()
        If _pythonExecuteRepairAdvisorSelfTestsRan Then Return
        _pythonExecuteRepairAdvisorSelfTestsRan = True

        Debug.WriteLine("[Startup] Queueing PythonExecuteRepairAdvisor self-tests...")

        System.Threading.Tasks.Task.Run(
            Sub()
                Try
                    Debug.WriteLine("[Startup] Running PythonExecuteRepairAdvisor self-tests...")
                    Dim status = SharedLibrary.AgentsXX.PythonExecuteRepairAdvisorSelfTests.RunAllAndReturnStatus()
                    Debug.WriteLine("[Startup] " & status)
                Catch ex As System.Exception
                    Debug.WriteLine("[Startup] PythonExecuteRepairAdvisor self-tests failed: " & ex.ToString())
                End Try
            End Sub)
    End Sub
#End If

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
            QueueModelAndAgentResourceWarmup()

            ' Reconcile the persisted CrashLog switch with the INI parameter. Any change
            ' takes effect on the next host launch (the INI is read after ThisAddIn_Startup).
            Try
                If My.Settings.CrashLog <> INI_Crashlog Then
                    My.Settings.CrashLog = INI_Crashlog
                    My.Settings.Save()
                End If
            Catch
            End Try

            UpdateHandler.PeriodicCheckForUpdates(INI_UpdateCheckInterval, "Outlook", INI_UpdatePath, _context)

            If _context IsNot Nothing AndAlso _context.INIloaded Then
                SharedMethods.RegisterLicenseCounterUsageAndReport(_context, "OL")
                StartLicenseCounterTimer()
            End If

            Dim result = Globals.Ribbons.Ribbon1.UpdateRibbon()
            Globals.Ribbons.Ribbon1.ApplyRibbonVisibilityConfiguration()
            result = Globals.Ribbons.Ribbon2.UpdateRibbon()
            Globals.Ribbons.Ribbon2.ApplyRibbonVisibilityConfiguration()
            mainThreadControl.CreateControl()
            StartListenerWatchdog()
            StartupHttpListener(INI_WebServerBlock)

            ' Initialize Knowledge Store background indexing service
            InitializeKnowledgeStoreService()

            Try
                If System.Threading.SynchronizationContext.Current Is Nothing Then
                    System.Threading.SynchronizationContext.SetSynchronizationContext(
                        New System.Windows.Forms.WindowsFormsSynchronizationContext())
                End If
                ' Touch a Control on the UI thread so the WindowsForms sync-context is fully active.
                Using anchor As New System.Windows.Forms.Control()
                    Dim h = anchor.Handle
                End Using
                SharedLibrary.SharedLibrary.SharedMethods.CleanupOrphanedWebView2Profiles()
                SharedLibrary.Agents.WebView2JsSandbox.Initialize(
                    System.Threading.SynchronizationContext.Current)
            Catch
                ' js_run will report "sandbox_uninitialized" if this failed.
            End Try

            ' Warm the Python Agent version cache off the UI thread so the first tooling run
            ' does not pay the version-probe cost. The lazy path in ResolveAndValidateAvailability
            ' re-probes if this warm-up is skipped or the cache is cold.
            PrimePythonAgentVersionCache()

        Catch ex As System.Exception
            ' Handling errors gracefully
        End Try

        ' AutoPilot auto-start: attempt after all other startup tasks have completed
        Try
            TryAutoStartAutoPilot()
        Catch
        End Try


    End Sub

    ''' <summary>
    ''' Fire-and-forget warm-up of the Python Agent version cache. Runs off the UI thread and never
    ''' throws; on failure the on-demand tool-registration path re-probes.
    ''' </summary>
    Private Sub PrimePythonAgentVersionCache()
        System.Threading.Tasks.Task.Run(
            Sub()
                Try
                    Dim configuration = SharedLibrary.Agents.PythonExecuteToolConfig.Parse(
                        SharedMethods.ExpandEnvironmentVariables(INI_PythonAgentPath))
                    If configuration IsNot Nothing Then
                        SharedLibrary.Agents.PythonExecuteToolCore.PrimeAvailabilityCache(configuration)
                    End If
                Catch ex As System.Exception
                    ' Non-fatal: the tool-registration path will re-probe on demand.
                End Try
            End Sub)
    End Sub

    ''' <summary>
    ''' Outlook add-in shutdown handler. Sequentially stops HTTP listener, watchdog, and power watch components.
    ''' </summary>
    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

        Try
            RiCrashLogger.Shutdown("ThisAddIn_Shutdown was called.")
        Catch
        End Try

        ' Shut down Knowledge Store service
        Try
            ShutdownKnowledgeStoreService()
        Catch
        End Try

        Try
            SharedLibrary.Agents.BrowserTools.Shutdown()
        Catch ex As System.Exception
            System.Diagnostics.Debug.WriteLine("BrowserTools.Shutdown failed: " & ex.Message)
        End Try

        ' 1) deterministically stop the HTTP listener (await synchronously)
        Try
            Dim t As System.Threading.Tasks.Task = ShutdownHttpListener()
            t.GetAwaiter().GetResult() ' safe: our shutdown continuations don’t capture the UI context
        Catch ex As System.Exception
            System.Diagnostics.Debug.WriteLine("ShutdownHttpListener failed: " & ex.Message)
        End Try

        Try
            StopLicenseCounterTimer()
        Catch
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

    Private Sub StartLicenseCounterTimer()
        Try
            StopLicenseCounterTimer()

            licenseCounterTimer = New System.Threading.Timer(
                Sub(state As Object)
                    Try
                        If _context Is Nothing OrElse Not _context.INIloaded Then
                            Return
                        End If

                        SharedMethods.RegisterLicenseCounterUsageAndReport(_context, "OL")
                    Catch
                    End Try
                End Sub,
                Nothing,
                TimeSpan.FromHours(SharedMethods.LicenseCounterOutlookTimerHours),
                TimeSpan.FromHours(SharedMethods.LicenseCounterOutlookTimerHours))
        Catch
        End Try
    End Sub

    Private Sub StopLicenseCounterTimer()
        Try
            If licenseCounterTimer IsNot Nothing Then
                licenseCounterTimer.Dispose()
                licenseCounterTimer = Nothing
            End If
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
        Dim tcs As New TaskCompletionSource(Of Object)(System.Threading.Tasks.TaskCreationOptions.RunContinuationsAsynchronously)
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

        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of Object)(System.Threading.Tasks.TaskCreationOptions.RunContinuationsAsynchronously)

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

        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of T)(System.Threading.Tasks.TaskCreationOptions.RunContinuationsAsynchronously)

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

        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of T)(System.Threading.Tasks.TaskCreationOptions.RunContinuationsAsynchronously)

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


    ''' <summary>
    ''' Primes context-independent model/tool INI parsing and the skill/agent index off the
    ''' Outlook UI thread. The on-demand paths remain authoritative and retry on failure.
    ''' No Office COM or WinForms objects are touched by the background task.
    ''' </summary>
    Private Sub QueueModelAndAgentResourceWarmup()
        Try
            Dim alternateModelPath As String = SharedMethods.ExpandEnvironmentVariables(If(INI_AlternateModelPath, ""))
            Dim specialServicePath As String = SharedMethods.ExpandEnvironmentVariables(If(INI_SpecialServicePath, ""))
            Dim centralResourcePath As String = SharedMethods.ExpandEnvironmentVariables(If(INI_AgentResourcesPath, ""))
            Dim localResourcePath As String = SharedMethods.ExpandEnvironmentVariables(If(INI_AgentResourcesPathLocal, ""))

            SharedLibrary.Agents.AgentResources.SetPaths(centralResourcePath, localResourcePath)

            System.Threading.Tasks.Task.Run(
                Sub()
                    Dim stopwatch As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew()
                    Try
                        SharedMethods.WarmAlternativeModelsCache(alternateModelPath)
                        SharedMethods.WarmAlternativeModelsCache(specialServicePath)
                        SharedLibrary.Agents.AgentResources.EnsureFresh()
                    Catch ex As System.Exception
                        System.Diagnostics.Debug.WriteLine("[PERF] Outlook startup warm-up failed: " & ex.Message)
                    Finally
                        stopwatch.Stop()
                        System.Diagnostics.Debug.WriteLine(
                            "[PERF] Outlook model/tool/resource warm-up: " &
                            stopwatch.ElapsedMilliseconds.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                            " ms")
                    End Try
                End Sub)
        Catch ex As System.Exception
            ' Non-fatal. Each on-demand path performs the same work if the cache is cold.
        End Try
    End Sub

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
    Public Shared Async Function LLM(ByVal promptSystem As String, ByVal promptUser As String, Optional ByVal Model As String = "", Optional ByVal Temperature As String = "", Optional ByVal Timeout As Long = 0, Optional ByVal UseSecondAPI As Boolean = False, Optional HideSplash As Boolean = False, Optional ByVal AddUserPrompt As String = "", Optional ByVal FileObject As String = "", Optional cancellationToken As Threading.CancellationToken = Nothing, Optional EnsureUI As Boolean = True, Optional ToolExecution As Boolean = False, Optional binaryOutputDirectory As String = Nothing) As Task(Of String)
        Dim Response = Await SharedMethods.LLM(_context, promptSystem, promptUser, Model, Temperature, Timeout, UseSecondAPI, HideSplash, AddUserPrompt, FileObject, cancellationToken, ToolExecution, binaryOutputDirectory)
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

' =================================================================================================
' Red Ink for Outlook - Architecture Overview
'
' ROLE OF THIS FILE
'   ThisAddIn.vb is the Outlook VSTO composition root. It owns application startup/shutdown,
'   Explorer/Inspector lifecycle integration, shared-context/config initialization,
'   UI-thread marshaling and host-level services. Mail features, Local Chat tooling and
'   AutoPilot are implemented in the other partial ThisAddIn.* files.
'
' PRIMARY ARCHITECTURAL AREAS
'   - Outlook UI/commands: Ribbon1.vb, DragDropForm.vb, ReviewChangesDialog.vb and
'     Commands.*, Helpers, Properties and Processing partials.
'   - Local Chat tooling: Tooling.ToolExecution, ToolExecutionContext, ToolResponse,
'     Tools, PromptBuilding, UserRequestResolution, Sources, Memory, M365, PythonExecute
'     and Logging; shared registry/sequencing/finality rules live in SharedLibrary.Agents.
'   - Agent isolation: ThisAddIn.AgentHost.vb implements ISubAgentHost and runs bounded
'     isolated tooling loops under the parent's workflow/task contracts.
'   - AutoPilot mail workflow: Autopilot.vb plus Config, Cleanup, Scheduler,
'     SenderToolPolicy, ThreadRetention, UserStorage*, Doc/PDF/comment processors and
'     CompleteTables. It processes one mail/session context while retaining explicit
'     attachment/output identities.
'   - AutoPilot tools: AutoPilot.Tools.vb aggregates tool definitions; Tools.Office*,
'     Tools.PDF, Tools.Other and DataCollector own domain execution. Structured/generic
'     Word generation is OOXML-first; Office.Interop is the explicit live-Excel COM
'     compatibility boundary. OpenXmlTemplate/OpenXmlVisuals own deterministic DOCX work.
'   - Microsoft 365/knowledge: Tooling.M365, M365SearchForm, Processing.KnowledgeRAG and
'     KnowledgeStoreWiring bridge shared M365 and Knowledge Store services.
'   - Web-extension surfaces: WebExtension.vb and Agent/FileHelpers/InkyPlay/
'     ListenerAndPower partials provide the external/local browser integration boundary.
'
' DESIGN/ARTIFACT BOUNDARIES
'   Office design catalogs and active design-set selection are data/configuration, not
'   hard-coded organization logic. Generated files are registered through explicit
'   artifact/deliverable identities. Retry fidelity must preserve resolved designs and
'   other user-significant choices instead of silently degrading to a neutral result.
'
' EXECUTION FLOW
'   Outlook event/user request -> Local Chat or AutoPilot orchestration -> capability
'   routing/tool loading -> shared sequencing + host tool execution -> validated artifact
'   or mail result -> final-response/delivery contract. UI/COM work must cross the Outlook
'   UI-thread boundary explicitly; OOXML-only document generation must not start Office.
'
' MAINTENANCE HOTSPOTS
'   Outlook COM/UI responsiveness; sender/trust and AutoPilot storage isolation; tooling
'   retries/finality; design routing; OOXML package validity; explicit Excel COM boundary;
'   attachment/path containment; M365/web external data; and user-visible artifact delivery.
' =================================================================================================
