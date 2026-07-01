' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Autopilot.vb
' Purpose:
'   Inky AutoPilot — AI-assisted mail watcher for Outlook that monitors incoming
'   e-mails, applies sender/domain/subject policy filters, processes requests via
'   LLM (with optional tool-calling), and sends or stages replies.
'
' Architecture:
'  - Event-driven intake via Outlook `Application.NewMailEx`.
'  - Dual mode:
'      * AutoPilot: auto-send when policy permits.
'      * CoPilot: approval workflow for non-whitelisted senders.
'  - Trust boundary:
'      * Positive/negative filter rules (domain/sender/folder) are always enforced.
'      * Reply-To mismatch is logged as warning (processing continues).
'  - Loop prevention:
'      * Custom MAPI marker (`X-RedInk-AutoReply`) on sent replies.
'      * Subject/header auto-reply detection (OOF/DSN patterns).
'      * Per-sender cooldown, per-session reply limits, thread-depth limit.
'  - Conversation awareness:
'      * Detects prior AutoPilot participation via conversation traversal and
'        session fallback map (`ConversationID`) to support follow-up replies.
'  - Catch-up processing:
'      * On start, scans recent mailbox history since last processed UTC timestamp,
'        applies full filter pipeline, and offers operator selection dialog.
'      * Mails marked with holding-only MAPI value (`AP_LoopHeaderValueHolding`)
'        are re-eligible for full processing.
'  - Queue + timing:
'      * Central mail queue with processing pump.
'      * Attachment-based category classification (text/light/heavy-doc/heavy-pdf).
'      * Rolling timing estimates and queue position notifications for delayed jobs.
'      * Active-job progress notifications for long-running tool executions.
'  - Tooling integration:
'      * Uses existing `ExecuteToolingLoop` and AutoPilot internal tools.
'      * Supports per-mail `#model:` override with tooling-capability checks.
'  - Reply handling:
'      * Sends HTML replies with optional generated attachments.
'      * Optional approval dialog and "Sources used" footer.
'      * Moves AutoPilot replies to Sent Items\Inky Replies.
'  - Web grounding:
'      * When `EnableWebGrounding` is active, the system prompt informs the model
'        about its native web-search capability.
'  - Scheduler lifecycle:
'      * Starts/stops the scheduler timer (`StartSchedulerTimer`/`StopSchedulerTimer`)
'        alongside the AutoPilot session when `EnableScheduler` is configured.
'      * Manages a scheduler dashboard for task monitoring.
'  - Voicemail processing:
'      * When enabled, incoming voicemails are identified by sender address,
'        transcribed via the model's audio capability, and processed as instructions.
'        Responses are delivered to the mapped e-mail address from the caller ID CSV.
'
' Security Model for Attachments:
'  - Per-mail isolated temp directory under `%TEMP%` (GUID-based).
'  - Attachments and tool outputs are constrained to that directory.
'  - Result collection validates path-prefix containment before attaching outputs.
'  - Text caches are cleared after processing.
'  - Temp directory is deleted recursively in `Finally` (best effort).
'
' Licensing & Access:
'  - AutoPilot requires valid license state (`Private` or permitted `Pro` product).
'  - Additional user whitelist gate via `INI_AutoPilot` (`*` = all licensed users).
'
' Threading Model:
'  - Outlook COM access is marshaled to UI thread (`SwitchToUi`).
'  - LLM and queue operations run asynchronously with cancellation support.
'  - Dashboard updates are thread-safe through `LogWindow` invocation guards.
'  - Notification timer runs independently of processing pump for timely alerts.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Collections.Concurrent
Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Markdig
Imports Microsoft.Office.Interop.Outlook
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ' ═══════════════════════════════════════════════════════════════════════════
    '  CONSTANTS
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Const AP_LoopHeaderProperty As String =
        "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/X-RedInk-AutoReply"
    Private Const AP_LoopHeaderValue As String = "true"

    ''' <summary>Header value for holding-only notices (not substantive replies).
    ''' The catch-up scan uses this to distinguish "already fully replied" from "only holding notice sent".</summary>
    Private Const AP_LoopHeaderValueHolding As String = "holding"

    Private Const AP_CategoryName As String = "Inky AutoPilot"
    Private Const AP_SentSubfolder As String = "Inky Replies"
    Private Const AP_MaxThreadDepth As Integer = 20
    Private Const AP_DefaultCooldownSeconds As Integer = 60
    Private Const AP_DefaultMaxRepliesPerSession As Integer = 200
    Private Const AP_DefaultMaxAttachmentBytes As Long = 10 * 1024 * 1024
    Private Const AP_TempPrefix As String = AN2 & "_autopilot_"

    ''' <summary>Maximum recursion depth for unpacking nested embedded mails.</summary>
    Private Const AP_MaxEmbeddedMailDepth As Integer = 5

    ''' <summary>Maximum number of files to extract from a single zip archive (zip bomb protection).</summary>
    Private Const AP_MaxZipEntries As Integer = 100

    ''' <summary>Maximum total uncompressed size from a single zip archive (zip bomb protection, 100 MB).</summary>
    Private Const AP_MaxZipTotalBytes As Long = 100 * 1024 * 1024

    ''' <summary>Maximum recursion depth for nested archives.</summary>
    Private Const AP_MaxArchiveDepth As Integer = 3


    ''' <summary>Command prefix scanned in the first few lines of the latest e-mail body.</summary>
    Private Const AP_ModelCommandPrefix As String = "#model:"

    ''' <summary>Max lines from top of latest mail body to scan for #model: command.</summary>
    Private Const AP_ModelCommandScanLines As Integer = 5

    Private Const SP_AutoPilot_HoldingResponse As String =
        "Thank you for your message. This is an automated acknowledgement — your request has not yet been processed. " &
        "I will respond with a substantive reply once your request has been handled. " &
        "If you need immediate assistance, you can also use the " & AN & " add-in's corresponding feature to have your tasks done right away. Use 'Help me, Inky' or the chatbot on https://redink.ai if you need instructions. Last but not least, a similar 'agent mode' is available in the Local Chat feature in the Outlook add-in if configured accordingly. " &
        "— " & AN6

    Private Const AP_MaxToolIterations As Integer = 50

    ''' <summary>
    ''' Product IDs whose Pro license holders are permitted to use AutoPilot.
    ''' Add or remove IDs here to control which Pro products grant access.
    ''' A Private license always grants access regardless of product ID.
    ''' </summary>
    Private Shared ReadOnly AP_PermittedProProductIds As String() = {
        "1702",
        "1727",
        "1827",
        "2040",
        "2058",
        "2059",
        "2060",
        "2061",
        "2062",
        "2063"
    }

    '  1702 Pro Special
    '  1727 Pro Test Tec
    '  1827 Pro Test
    '  2040 Pro Support
    '  2150 Pro Special 2 -- NO
    '  1693 Pro -- NO
    '  2058-2063 Special License for AutoPilot


    ' ═══════════════════════════════════════════════════════════════════════════
    '  STATE
    ' ═══════════════════════════════════════════════════════════════════════════

    Private _apActive As Boolean = False
    Private _apConfig As AutoPilotConfig = Nothing
    Private _apCts As CancellationTokenSource = Nothing
    Private _apDashboard As LogWindow = Nothing
    Private ReadOnly _apSenderLastMailSentUtc As New ConcurrentDictionary(Of String, DateTime)(StringComparer.OrdinalIgnoreCase)
    Private _apSessionReplyCount As Integer = 0
    Private ReadOnly _apMailQueue As New ConcurrentQueue(Of String)()
    Private ReadOnly _apProcessingSemaphore As New SemaphoreSlim(1, 1)
    Private _apSelectedTools As List(Of ModelConfig) = Nothing
    Private _apUseSecondApi As Boolean = False

    ''' <summary>
    ''' Paths of knowledge-store source files copied into the temp directory.
    ''' Only these are candidates for citation-based filtering — tool outputs
    ''' from process_word_document, merge_pdfs, etc. are never filtered.
    ''' </summary>
    Private _apKnowledgeSourceCopies As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

    Private Shared ReadOnly AP_AutoReplySubjectPatterns As String() = {
        "automatic reply:", "automatische antwort:", "réponse automatique:",
        "risposta automatica:", "out of office:", "abwesenheitsnotiz:",
        "absence du bureau:", "fuori sede:", "auto-reply:", "auto reply:",
        "undeliverable:", "nicht zustellbar:", "delivery status notification",
        "mail delivery failed", "returned mail:"
    }

    ' ── Per-mail transient state ──
    Private _apCurrentTempDir As String = Nothing
    Private _apCurrentAttachments As List(Of AutoPilotAttachmentInfo) = Nothing
    Private _apCurrentMailInfo As AutoPilotMailInfo = Nothing

    ''' <summary>Tracks tool calls during processing of a single e-mail for the "Sources used:" footer.</summary>
    Private _apCurrentToolCallLog As List(Of AutoPilotToolCallEntry) = Nothing

    ''' <summary>Captured at AutoPilot startup — the model config to re-apply before every LLM call.</summary>
    Private _apBaseModelConfig As ModelConfig = Nothing

    ''' <summary>Minimum estimated queue wait (seconds) before a queue notification is sent.</summary>
    Private Const AP_QueueNotifyThresholdSeconds As Integer = 60

    ''' <summary>Interval (seconds) between repeat holding notifications for the same queued mail.</summary>
    Private Const AP_QueueNotifyRepeatIntervalSeconds As Integer = 45 * 60  ' 45 minutes

    ''' <summary>Interval (seconds) between progress notifications for the currently-processing mail.</summary>
    Private Const AP_ActiveJobNotifyIntervalSeconds As Integer = 30 * 60  ' 30 minutes

    ''' <summary>Extensions that indicate heavy document processing (doc processor, comment processor).</summary>
    Private Shared ReadOnly AP_HeavyDocExtensions As String() = {".docx", ".pptx", ".xlsx", ".xls", ".doc"}

    ''' <summary>Threshold (bytes) above which a PDF is considered heavy (likely OCR).</summary>
    Private Const AP_HeavyPdfThreshold As Long = 2 * 1024 * 1024

    ''' <summary>
    ''' Internal time seeds per mail category (seconds). Used solely to decide whether
    ''' the 60-second notification threshold is crossed — never exposed to the sender.
    ''' </summary>
    Private Const AP_SeedSeconds_TextOnly As Double = 15.0
    Private Const AP_SeedSeconds_LightAttachments As Double = 30.0
    Private Const AP_SeedSeconds_HeavyDoc As Double = 300.0
    Private Const AP_SeedSeconds_HeavyPdf As Double = 600.0

    ''' <summary>Tracks when each queued mail was first seen in the queue (EntryId → UTC enqueue time).</summary>
    Private ReadOnly _apQueueEnqueueTimes As New ConcurrentDictionary(Of String, DateTime)()

    ''' <summary>Entry IDs that have received a queue holding notification, with the UTC time of the last notification.</summary>
    Private ReadOnly _apQueueNotifiedEntryIds As New ConcurrentDictionary(Of String, DateTime)()

    ''' <summary>Rolling average processing time per mail category. Key = category tag, Value = (avg, count).</summary>
    Private ReadOnly _apCategoryTimings As New ConcurrentDictionary(Of String, (AvgSeconds As Double, SampleCount As Integer))(StringComparer.OrdinalIgnoreCase)

    ''' <summary>Stopwatch tracking the current mail's processing time.</summary>
    Private _apCurrentProcessingStopwatch As Stopwatch = Nothing

    ''' <summary>Category tag of the mail currently being processed (for timing updates).</summary>
    Private _apCurrentProcessingCategory As String = Nothing

    ''' <summary>
    ''' Periodic timer that checks for queued mails needing holding notifications.
    ''' Runs independently of the processing pump so notifications fire even during
    ''' long-running tool executions (e.g. multi-hour document processing).
    ''' </summary>
    Private _apNotificationTimer As System.Threading.Timer = Nothing

    ''' <summary>Guard to prevent overlapping notification checks.</summary>
    Private _apNotificationCheckRunning As Integer = 0

    ''' <summary>Cached HelpMeInky manual text, loaded once per AutoPilot session.</summary>
    Private _apHelpMeManualCache As String = Nothing
    Private _apHelpMeManualCacheLoaded As Boolean = False

    ''' <summary>Per-job CancellationTokenSource. Cancelling this aborts only the current job, not the entire AutoPilot session.</summary>
    Private _apCurrentJobCts As CancellationTokenSource = Nothing

    ''' <summary>Tracks entry IDs that have received only a holding notice (not a substantive reply).
    ''' Used by the catch-up scan to re-process mails that were acknowledged but never fully handled.</summary>
    Private ReadOnly _apHoldingOnlyEntryIds As New ConcurrentDictionary(Of String, Boolean)()

    ''' <summary>Entry ID of the mail currently being processed (set by the pump, cleared on completion).</summary>
    Private _apCurrentProcessingEntryId As String = Nothing

    ''' <summary>UTC time of the last active-job progress notification sent for the current mail.</summary>
    Private _apActiveJobLastNotifiedUtc As DateTime = DateTime.MinValue

    ''' <summary>Tracks entry IDs that were explicitly selected for catch-up or reprocessing.
    ''' These mails bypass cooldown because the operator (or unattended mode) explicitly chose them.</summary>
    Private ReadOnly _apCatchUpEntryIds As New ConcurrentDictionary(Of String, Boolean)()

    ' ═══════════════════════════════════════════════════════════════════════════
    '  PUBLIC ENTRY POINTS
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Stops AutoPilot processing and clears runtime state.</summary>
    Public Sub StopAutoPilot()
        If Not _apActive Then Return
        _apActive = False
        Try : RemoveHandler Application.NewMailEx, AddressOf AutoPilot_NewMailEx : Catch : End Try
        Try
            If _apDashboard IsNot Nothing Then
                RemoveHandler _apDashboard.CancelRequested, AddressOf AutoPilot_DashboardCancelRequested
                RemoveHandler _apDashboard.AbortJobRequested, AddressOf AutoPilot_DashboardAbortJobRequested
            End If
        Catch : End Try
        Try : _apNotificationTimer?.Dispose() : Catch : End Try
        _apNotificationTimer = Nothing
        StopAutoDeleteTimer()
        StopSchedulerTimer()
        StopSchedulerTimer()
        CloseSchedulerDashboard()
        CloseUserStorageDashboard()
        Try : _apCts?.Cancel() : Catch : End Try
        Try : _apCurrentJobCts?.Cancel() : Catch : End Try
        Try : _apCts?.Dispose() : Catch : End Try
        Try : _apCurrentJobCts?.Dispose() : Catch : End Try
        _apCts = Nothing
        _apCurrentJobCts = Nothing
        Dim dummy As String = Nothing
        While _apMailQueue.TryDequeue(dummy) : End While
        _apSessionReplyCount = 0
        _apSelectedTools = Nothing
        _apConfig = Nothing
        _apBaseModelConfig = Nothing
        _apSenderLastMailSentUtc.Clear()
        _apProcessedConversations.Clear()
        _apQueueNotifiedEntryIds.Clear()
        _apQueueEnqueueTimes.Clear()
        _apCategoryTimings.Clear()
        _apCurrentProcessingStopwatch = Nothing
        _apCurrentProcessingCategory = Nothing
        _apHelpMeManualCache = Nothing
        _apHelpMeManualCacheLoaded = False
        _apHoldingOnlyEntryIds.Clear()
        _apCatchUpEntryIds.Clear()
        _apCurrentProcessingEntryId = Nothing
        _apActiveJobLastNotifiedUtc = DateTime.MinValue
        _apVoicemailCallerIdMap = Nothing

        If INI_WebServerBlock <> 4 Then
            EnsureLocalSchedulerTimerStarted()
        End If

        WebGrounding = ""
        ApDashboardLog("AutoPilot stopped.", "info")
        ApDashboardMarkComplete()
        ShowCustomMessageBox($"{AN6} AutoPilot has been stopped.", AN)
    End Sub

    ''' <summary>Shows the AutoPilot dashboard if it is available.</summary>
    Public Sub ShowAutoPilotDashboard()
        If _apDashboard Is Nothing Then Return

        Dim showDashboard As System.Action =
            Sub()
                Try
                    If _apDashboard Is Nothing OrElse _apDashboard.IsDisposed Then Return
                    _apDashboard.Show()
                    _apDashboard.BringToFront()
                Catch
                End Try
            End Sub

        Try
            If UiSyncContext IsNot Nothing AndAlso
               System.Threading.Thread.CurrentThread.ManagedThreadId <> UiThreadId Then
                UiSyncContext.Post(Sub() showDashboard(), Nothing)
            Else
                showDashboard()
            End If
        Catch
        End Try
    End Sub

    ''' <summary>Starts AutoPilot using the configuration dialog and saved settings.</summary>
    Public Sub StartAutoPilot()
        ' Distinguish license failure from user whitelist failure for a clear message
        If Not IsAutoPilotLicenseValid() Then
            Dim licenseStatus = GetLicenseStatusShort()
            ShowCustomMessageBox(
                $"{AN6} AutoPilot requires a Pro or Private license." & vbCrLf &
                $"Current license: {licenseStatus}" & vbCrLf & vbCrLf &
                $"Visit {AN4} to obtain a license.",
                AN)
            Return
        End If
        If Not IsAutoPilotPermitted() Then
            ShowCustomMessageBox(
                $"{AN6} AutoPilot is not available for your user account ({Environment.UserName}). Contact your administrator to request access.",
                AN)
            Return
        End If
        If _apActive Then
            Dim choice = ShowCustomYesNoBox(
                AN6 & " AutoPilot is already running." & vbCrLf &
                "Would you like to stop it or view the dashboard?",
                "Stop AutoPilot", "Show Dashboard",
                header:=AN6 & " AutoPilot")
            If choice = 1 Then StopAutoPilot()
            If choice = 2 Then ShowAutoPilotDashboard()
            Return
        End If
        Dim config = ShowAutoPilotConfigDialog()
        If config Is Nothing Then Return
        SaveAutoPilotConfigToSettings(config)
        StartAutoPilotWithConfig(config)
    End Sub

    ''' <summary>Starts AutoPilot with an explicit configuration instance.</summary>
    Private Sub StartAutoPilotWithConfig(config As AutoPilotConfig)
        _apConfig = config
        StopLocalSchedulerRuntime()
        _apActive = True
        _apCts = New CancellationTokenSource()
        _apSessionReplyCount = 0
        _apSenderLastMailSentUtc.Clear()
        _apUseSecondApi = config.UseSecondApi

        WebGrounding = ""
        PrivacyProtection = If(_apConfig.EnablePrivacyProtection, SP_Add_PrivacyProtection, "")

        ' Load voicemail caller ID map if voicemail processing is enabled
        If config.EnableVoicemailProcessing AndAlso Not String.IsNullOrWhiteSpace(config.VoicemailCallerIdMapPath) Then
            LoadVoicemailCallerIdMap(config.VoicemailCallerIdMapPath)
        Else
            _apVoicemailCallerIdMap = Nothing
        End If

        ' Capture the current model config at startup so we can re-apply it
        ' before every LLM call, ensuring tooling templates are always pristine.
        _apBaseModelConfig = GetCurrentConfig(_context)

        ' Determine if the selected model supports tooling
        Dim modelCanCallTools As Boolean = ModelSupportsTooling(_apBaseModelConfig)

        ' Only load tools if the model can actually call them
        If modelCanCallTools Then
            _apSelectedTools = New List(Of ModelConfig)()
            _apSelectedTools.AddRange(GetAutoPilotInternalTools())

            If config.SelectedExternalTools IsNot Nothing Then
                _apSelectedTools.AddRange(
                    NormalizeAutoPilotSelectableExternalTools(config.SelectedExternalTools))
            End If

            _apSelectedTools =
                _apSelectedTools.
                    Where(Function(tool)
                              Return tool IsNot Nothing AndAlso
                                     Not SharedLibrary.Agents.MemoryTools.IsMemoryTool(tool.ToolName)
                          End Function).
                    GroupBy(Function(tool) tool.ToolName, StringComparer.OrdinalIgnoreCase).
                    Select(Function(group) group.First()).
                    ToList()
        Else
            _apSelectedTools = Nothing
        End If

        ' Load the optional per-sender tool policy for this session (additive layer).
        LoadSenderToolPolicy(config.SenderToolPolicyPath)

        Dim initializeDashboard As System.Action =
            Sub()
                _apDashboard = New LogWindow()
                _apDashboard.Text = $"{AN6} AutoPilot Dashboard"
                AddHandler _apDashboard.CancelRequested, AddressOf AutoPilot_DashboardCancelRequested
                AddHandler _apDashboard.AbortJobRequested, AddressOf AutoPilot_DashboardAbortJobRequested
                _apDashboard.ShowAbortJobButton(True)
                _apDashboard.Show()

                ' Add Scheduler dashboard button if scheduler is enabled
                If config.EnableScheduler Then
                    Try
                        Dim buttonPanel = _apDashboard.Controls.OfType(Of TableLayoutPanel)().
                            FirstOrDefault()?.Controls.OfType(Of FlowLayoutPanel)().FirstOrDefault()

                        If buttonPanel IsNot Nothing Then
                            Dim btnScheduler As New Button() With {
                                .Text = "Scheduler",
                                .AutoSize = True,
                                .Padding = New Padding(10, 5, 10, 5)
                            }
                            AddHandler btnScheduler.Click, Sub(s, e) ShowSchedulerDashboard()
                            buttonPanel.Controls.Add(btnScheduler)
                        End If
                    Catch
                    End Try
                End If

                ' Add User Storage dashboard button if memory or files are enabled
                If config.EnableUserMemory OrElse config.EnableUserFiles Then
                    Try
                        Dim buttonPanel = _apDashboard.Controls.OfType(Of TableLayoutPanel)().
                            FirstOrDefault()?.Controls.OfType(Of FlowLayoutPanel)().FirstOrDefault()

                        If buttonPanel IsNot Nothing Then
                            Dim btnUserStorage As New Button() With {
                                .Text = "User Storage",
                                .AutoSize = True,
                                .Padding = New Padding(10, 5, 10, 5)
                            }
                            AddHandler btnUserStorage.Click, Sub(s, e) ShowUserStorageDashboard()
                            buttonPanel.Controls.Add(btnUserStorage)
                        End If
                    Catch
                    End Try
                End If
            End Sub

        If UiSyncContext IsNot Nothing AndAlso
           System.Threading.Thread.CurrentThread.ManagedThreadId <> UiThreadId Then
            UiSyncContext.Send(Sub() initializeDashboard(), Nothing)
        Else
            initializeDashboard()
        End If

        Dim modelLabel As String
        If Not String.IsNullOrWhiteSpace(config.SelectedModelKey) Then
            modelLabel = config.SelectedModelKey
        ElseIf config.UseSecondApi Then
            modelLabel = INI_Model_2
        Else
            modelLabel = INI_Model
        End If

        ApDashboardLog($"AutoPilot started.", "info")
        ApDashboardLog($"Model: {modelLabel}", "info")

        If Not modelCanCallTools Then
            ApDashboardLog("⚠ Model does not support tool calling — running without tools", "warn")
        Else
            ApDashboardLog($"Filters: {config.FilterRules.Count} rule(s) active", "info")
            ApDashboardLog($"{ToolFriendlyName}: {_apSelectedTools.Count} available (incl. built-in)", "info")

            ' Log each tool so operator can verify the court database tool is loaded
            For Each tool In _apSelectedTools
                Dim toolLabel = If(Not String.IsNullOrEmpty(tool.ModelDescription), tool.ModelDescription,
                               If(Not String.IsNullOrEmpty(tool.ToolName), tool.ToolName, "(unnamed)"))
                Dim hasDefinition = Not String.IsNullOrWhiteSpace(tool.ToolDefinition)
                Dim hasInstructions = Not String.IsNullOrWhiteSpace(tool.ToolInstructionsPrompt)
                ApDashboardLog($"  • {toolLabel} [def:{hasDefinition}, instr:{hasInstructions}]", "step")
            Next

            ' Log tooling config to verify detection/extraction patterns are set
            'ApDashboardLog($"ToolCallDetectionPattern: {If(Not String.IsNullOrWhiteSpace(_context.INI_ToolCallDetectionPattern_2), "set", "EMPTY")}", "step")
            'ApDashboardLog($"ToolCallExtractionMap: {If(Not String.IsNullOrWhiteSpace(_context.INI_ToolCallExtractionMap_2), "set", "EMPTY")}", "step")
            'ApDashboardLog($"ToolInstructions template: {If(Not String.IsNullOrWhiteSpace(_context.INI_APICall_ToolInstructions_Template_2), "set (" & _context.INI_APICall_ToolInstructions_Template_2.Length.ToString() & " chars)", "EMPTY")}", "step")
            'ApDashboardLog($"ToolResponses template: {If(Not String.IsNullOrWhiteSpace(_context.INI_APICall_ToolResponses_Template_2), "set", "EMPTY")}", "step")
            'ApDashboardLog($"ToolCallPart template: {If(Not String.IsNullOrWhiteSpace(_context.INI_APICall_ToolCallPart_Template_2), "set", "EMPTY")}", "step")
            'ApDashboardLog($"APICall_2 contains toolinstructions placeholder: {If(_context.INI_APICall_2 IsNot Nothing AndAlso _context.INI_APICall_2.Contains("toolinstructions"), "YES", "NO")}", "step")
        End If

        If config.EnableWebGrounding Then
            ApDashboardLog("🌐 Web grounding enabled (model has built-in web search)", "info")
        End If

        If config.EnableVoicemailProcessing Then
            ApDashboardLog($"📞 Voicemail processing enabled (sender: {config.VoicemailSenderAddress})", "info")
            If _apVoicemailCallerIdMap IsNot Nothing Then
                ApDashboardLog($"   Caller ID map: {_apVoicemailCallerIdMap.Count} mapping(s) loaded", "info")
            Else
                ApDashboardLog("   ⚠ Caller ID map not loaded", "warn")
            End If
        End If

        If config.EnableUserMemory Then
            ApDashboardLog("🧠 Per-user memory enabled (auto-learn + manage_user_memory tool)", "info")
        End If
        If config.EnableUserFiles Then
            ApDashboardLog("📁 Per-user file storage enabled (manage_user_files tool)", "info")
        End If
        If config.EnablePrivacyProtection Then
            ApDashboardLog("🔒 Privacy protection enabled (search queries sanitized)", "info")
        Else
            ApDashboardLog("🔓 Privacy protection disabled (unrestricted search queries)", "info")
        End If

        ApDashboardLog($"Filters: {config.FilterRules.Count} rule(s) active", "info")
        ApDashboardLog($"Mode: {If(config.RequireApprovalForNonWhitelisted, "CoPilot (approval for non-whitelisted)", "AutoPilot (auto-send all)")}", "info")
        ApDashboardLog("Watching for new mail...", "info")

        CatchUpMissedMails()

        ' Start scheduler if enabled
        If config.EnableScheduler Then
            If INI_WebServerBlock = 4 Then
                ApDashboardLog("📅 Scheduler disabled because WebServerBlock = 4.", "warn")
            Else
                SchedulerCatchUp()
                StartSchedulerTimer()
            End If
        End If

        If config.AutoDeleteAfterHours > 0 Then
            ApDashboardLog($"🗑 Auto-delete enabled after {config.AutoDeleteAfterHours}h (includes Deleted Items).", "info")
            StartAutoDeleteTimer()
            RunAutoDeleteCleanupAsync(_apCts.Token, "startup")
        Else
            ApDashboardLog("🗑 Auto-delete disabled.", "info")
        End If

        AddHandler Application.NewMailEx, AddressOf AutoPilot_NewMailEx
        Task.Run(Function() AutoPilotProcessingPump(_apCts.Token))

        ' Start a periodic timer for queue holding notifications.
        ' This runs independently of the processing pump so notifications
        ' can fire even while a long tool execution blocks the pump.
        _apNotificationTimer = New System.Threading.Timer(
            AddressOf NotificationTimerCallback,
            Nothing,
            dueTime:=TimeSpan.FromSeconds(15),
            period:=TimeSpan.FromSeconds(15))

    End Sub

    ''' <summary>Handles a dashboard stop request.</summary>
    Private Sub AutoPilot_DashboardCancelRequested(sender As Object, e As EventArgs)
        StopAutoPilot()
    End Sub

    ''' <summary>Handles a dashboard abort-current-job request (does NOT stop AutoPilot).</summary>
    Private Sub AutoPilot_DashboardAbortJobRequested(sender As Object, e As EventArgs)
        Dim jobCts = _apCurrentJobCts
        If jobCts IsNot Nothing AndAlso Not jobCts.IsCancellationRequested Then
            ApDashboardLog("⛔ Operator requested abort of current job.", "warn")
            Try : jobCts.Cancel() : Catch : End Try
        Else
            ApDashboardLog("No job currently processing to abort.", "step")
        End If
    End Sub

    ''' <summary>
    ''' Scans the Inbox (or monitored folder) for mails that arrived after the last
    ''' processed timestamp, applies the full filter/trigger logic, and shows a
    ''' preview dialog where the operator can check/uncheck individual mails.
    ''' Only mails that have NOT already been processed (no AP_CategoryName tag) are shown.
    ''' Mails that have received only a holding notice (not a substantive reply) ARE shown.
    ''' When ReprocessLookbackHours > 0, already-processed mails within that window are
    ''' also shown, allowing the operator to select them for reprocessing.
    ''' </summary>
    Private Sub CatchUpMissedMails()
        Try
            Dim lastProcessed As DateTime = My.Settings.AP_LastProcessedTime
            Dim isFirstRun As Boolean = False

            If lastProcessed = DateTime.MinValue OrElse lastProcessed = New DateTime(1, 1, 1) Then
                ' No previous session timestamp — default to 24 hours ago in local time.
                ' We use local time directly because MailItem.ReceivedTime is always local,
                ' and we want exactly 24 hours of local-time coverage.
                lastProcessed = DateTime.Now.AddHours(-24)
                isFirstRun = True
                ApDashboardLog("No previous session timestamp — scanning last 24 hours for relevant mails.", "info")
            End If

            ' For persisted values: ensure we treat the stored value as UTC
            ' (it was saved as DateTime.UtcNow) and convert to local for comparison.
            ' For the first-run fallback: already local, skip conversion.
            Dim lastProcessedLocal As DateTime
            If isFirstRun Then
                lastProcessedLocal = lastProcessed
            Else
                If lastProcessed.Kind = DateTimeKind.Unspecified Then
                    lastProcessed = DateTime.SpecifyKind(lastProcessed, DateTimeKind.Utc)
                End If
                lastProcessedLocal = lastProcessed.ToLocalTime()

                ' Cap: never look back more than 24 hours even if the saved timestamp is older
                Dim maxLookback = DateTime.Now.AddHours(-24)
                If lastProcessedLocal < maxLookback Then
                    lastProcessedLocal = maxLookback
                    ApDashboardLog("Saved timestamp is older than 24 hours — capping lookback to 24h.", "info")
                End If
            End If

            ' Determine the reprocess lookback cutoff (for already-processed mails)
            Dim reprocessLookbackHours = If(_apConfig IsNot Nothing, _apConfig.ReprocessLookbackHours, 0)
            Dim reprocessCutoffLocal As DateTime = DateTime.MinValue
            If reprocessLookbackHours > 0 Then
                reprocessCutoffLocal = DateTime.Now.AddHours(-reprocessLookbackHours)
                ApDashboardLog($"Reprocess lookback: {reprocessLookbackHours}h (mails since {reprocessCutoffLocal:yyyy-MM-dd HH:mm:ss})", "info")
            End If

            ' Use the earlier of the two cutoffs for the scan window
            Dim scanCutoffLocal As DateTime
            If reprocessLookbackHours > 0 AndAlso reprocessCutoffLocal < lastProcessedLocal Then
                scanCutoffLocal = reprocessCutoffLocal
            Else
                scanCutoffLocal = lastProcessedLocal
            End If

            If isFirstRun Then
                ApDashboardLog($"Scanning for mails received after: {scanCutoffLocal:yyyy-MM-dd HH:mm:ss} (local)", "info")
            Else
                ApDashboardLog($"Last processed: {lastProcessedLocal:yyyy-MM-dd HH:mm:ss} (local) / {lastProcessed:yyyy-MM-dd HH:mm:ss} (UTC)", "info")
            End If
            ApDashboardLog("Scanning inbox for newer mails...", "info")

            Dim ns = Application.GetNamespace("MAPI")
            Dim inbox As MAPIFolder = Nothing

            ' Resolve the correct inbox when a specific mailbox is monitored
            If Not String.IsNullOrWhiteSpace(_apConfig.MonitoredMailbox) Then
                Try
                    For i As Integer = 1 To ns.Accounts.Count
                        Dim acct = ns.Accounts(i)
                        If acct.SmtpAddress.Equals(_apConfig.MonitoredMailbox, StringComparison.OrdinalIgnoreCase) Then
                            Dim deliveryStore = acct.DeliveryStore
                            If deliveryStore IsNot Nothing Then
                                inbox = deliveryStore.GetDefaultFolder(OlDefaultFolders.olFolderInbox)
                            End If
                            Exit For
                        End If
                    Next
                Catch
                End Try
            End If

            ' Fallback to default inbox
            If inbox Is Nothing Then
                inbox = ns.GetDefaultFolder(OlDefaultFolders.olFolderInbox)
            End If

            ' Sort descending so we process newest first and can stop early
            ' once we hit mails older than scanCutoffLocal.
            Dim allItems = inbox.Items
            allItems.Sort("[ReceivedTime]", Descending:=True)

            Dim candidates As New List(Of CatchUpCandidate)()
            Dim skippedAlreadyProcessed As Integer = 0
            Dim skippedOther As Integer = 0
            Dim totalScanned As Integer = 0
            Dim reprocessCandidateCount As Integer = 0

            For Each item As Object In allItems
                If Not TypeOf item Is MailItem Then Continue For
                Dim mi = DirectCast(item, MailItem)
                Try

                    ' MailItem.ReceivedTime is always local time.
                    ' Stop scanning once we've gone past the scan cutoff.
                    If mi.ReceivedTime <= scanCutoffLocal Then
                        Exit For
                    End If

                    totalScanned += 1

                    ' 1. Skip our own auto-replies (header on the mail itself)
                    Try
                        Dim prop = mi.PropertyAccessor.GetProperty(AP_LoopHeaderProperty)
                        If prop IsNot Nothing AndAlso prop.ToString() = AP_LoopHeaderValue Then
                            Continue For
                        End If
                    Catch
                    End Try

                    ' 2. Check if already processed (AP category tag or substantive reply)
                    Dim isAlreadyProcessed As Boolean = False

                    Try
                        Dim cats = mi.Categories
                        If cats IsNot Nothing AndAlso cats.IndexOf(AP_CategoryName, StringComparison.OrdinalIgnoreCase) >= 0 Then
                            isAlreadyProcessed = True
                        End If
                    Catch
                    End Try

                    If Not isAlreadyProcessed Then
                        Try
                            If HasSubstantiveAutoPilotReply(mi) Then
                                isAlreadyProcessed = True
                            End If
                        Catch
                        End Try
                    End If

                    ' If already processed: only include if within the reprocess lookback window
                    If isAlreadyProcessed Then
                        If reprocessLookbackHours > 0 AndAlso mi.ReceivedTime > reprocessCutoffLocal Then
                            ' Falls within reprocess window — allow it through as a reprocess candidate
                            ' (don't skip, let it pass the filter checks below)
                        Else
                            skippedAlreadyProcessed += 1
                            Continue For
                        End If
                    End If

                    ' Full pre-filter: same logic as ProcessIncomingMailAsync
                    Dim mailInfo = ExtractMailInfo(mi)
                    If mailInfo Is Nothing Then Continue For

                    ' ── Voicemail detection (bypasses normal filter pipeline, same as ProcessIncomingMailAsync) ──
                    Dim isVoicemail As Boolean = IsVoicemailFromRegisteredSender(mailInfo)

                    If Not isVoicemail Then
                        If mailInfo.HasAutoReplyHeader Then Continue For
                        If IsAutoReplyOrOof(mailInfo) Then
                            skippedOther += 1
                            Continue For
                        End If
                        If Not MatchesFilterRules(mailInfo) Then
                            skippedOther += 1
                            Continue For
                        End If
                        If MatchesNegativeFilters(mailInfo) Then
                            skippedOther += 1
                            Continue For
                        End If

                        ' Subject trigger word check
                        If Not String.IsNullOrWhiteSpace(_apConfig.SubjectTriggerWord) Then
                            If mailInfo.Subject.IndexOf(_apConfig.SubjectTriggerWord, StringComparison.OrdinalIgnoreCase) < 0 Then
                                skippedOther += 1
                                Continue For
                            End If
                        End If
                    End If

                    Dim displayLabel As String = Nothing
                    If isVoicemail Then
                        displayLabel = "[VOICEMAIL] " & New CatchUpCandidate() With {
                            .SenderName = mailInfo.SenderName,
                            .SenderEmail = mailInfo.SenderEmail,
                            .Subject = mailInfo.Subject,
                            .ReceivedTime = mailInfo.ReceivedTime
                        }.ToDisplayLabel()
                    ElseIf isAlreadyProcessed Then
                        reprocessCandidateCount += 1
                        displayLabel = "[REPROCESS] " & New CatchUpCandidate() With {
                            .SenderName = mailInfo.SenderName,
                            .SenderEmail = mailInfo.SenderEmail,
                            .Subject = mailInfo.Subject,
                            .ReceivedTime = mailInfo.ReceivedTime
                        }.ToDisplayLabel()
                    End If

                    candidates.Add(New CatchUpCandidate() With {
                        .EntryID = mi.EntryID,
                        .SenderName = mailInfo.SenderName,
                        .SenderEmail = mailInfo.SenderEmail,
                        .Subject = mailInfo.Subject,
                        .ReceivedTime = mailInfo.ReceivedTime,
                        .IsReprocessCandidate = isAlreadyProcessed,
                        .CustomDisplayLabel = displayLabel
                    })
                Finally
                    Try : Marshal.ReleaseComObject(mi) : Catch : End Try
                End Try
            Next

            ApDashboardLog($"Scan complete: {totalScanned} mail(s) scanned, {skippedAlreadyProcessed} already processed (outside reprocess window), {skippedOther} filtered out.", "step")
            If reprocessCandidateCount > 0 Then
                ApDashboardLog($"Reprocess candidates: {reprocessCandidateCount} already-processed mail(s) within {reprocessLookbackHours}h lookback", "info")
            End If

            If candidates.Count = 0 Then
                ApDashboardLog("No missed mails found matching filters.", "step")
                Return
            End If

            ' Re-sort candidates oldest-first for processing order
            candidates.Sort(Function(a, b) a.ReceivedTime.CompareTo(b.ReceivedTime))

            ApDashboardLog($"Found {candidates.Count} mail(s) matching filters ({candidates.Count - reprocessCandidateCount} new, {reprocessCandidateCount} reprocess).", "info")

            Dim selectedEntryIds As List(Of String) = Nothing

            If _apConfig.IsUnattended Then
                ' Unattended auto-start: skip the preview dialog, process all candidates
                ' but exclude reprocess candidates in unattended mode (only operator should choose those)
                selectedEntryIds = candidates.Where(Function(c) Not c.IsReprocessCandidate).Select(Function(c) c.EntryID).ToList()
                ApDashboardLog($"Unattended mode — auto-selecting {selectedEntryIds.Count} new catch-up mail(s) (skipping {reprocessCandidateCount} reprocess candidates).", "info")
            Else
                ' Show preview dialog using MultiModelSelectorForm pattern
                ' Pre-select only NEW (unprocessed) mails; reprocess candidates are unchecked by default
                selectedEntryIds = ShowCatchUpPreviewDialog(candidates, lastProcessed)
            End If

            If selectedEntryIds IsNot Nothing AndAlso selectedEntryIds.Count > 0 Then
                For Each entryId In selectedEntryIds
                    _apMailQueue.Enqueue(entryId)
                    _apQueueEnqueueTimes.TryAdd(entryId, DateTime.UtcNow)
                    _apCatchUpEntryIds.TryAdd(entryId, True)
                Next
                ApDashboardLog($"Queued {selectedEntryIds.Count} of {candidates.Count} mail(s) for processing.", "info")
            Else
                ApDashboardLog("Operator skipped catch-up processing.", "step")
            End If

        Catch ex As System.Exception
            ApDashboardLog($"Catch-up scan error: {ex.Message}", "warn")
            Debug.WriteLine($"[AutoPilot] CatchUpMissedMails error: {ex}")
        End Try
    End Sub

    ''' <summary>
    ''' Checks whether a mail already has a SUBSTANTIVE AutoPilot reply in its conversation.
    ''' Returns True only if a reply with header value "true" is found (not "holding").
    ''' This allows mails that received only a holding notice to be re-processed on catch-up.
    ''' </summary>
    Private Function HasSubstantiveAutoPilotReply(mi As MailItem) As Boolean
        Try
            ' Walk the conversation thread to find any mail with a substantive reply header.
            Dim conversation As Outlook.Conversation = Nothing
            Try
                conversation = mi.GetConversation()
            Catch
            End Try

            If conversation IsNot Nothing Then
                Try
                    Dim rootItems As SimpleItems = conversation.GetRootItems()
                    If rootItems IsNot Nothing Then
                        For Each rootItem As Object In rootItems
                            If CheckConversationNodeForSubstantiveReply(conversation, rootItem, 0) Then
                                Return True
                            End If
                        Next
                    End If
                Catch
                End Try
            End If
        Catch
        End Try
        Return False
    End Function

    ''' <summary>Recursively checks conversation tree nodes for a substantive AutoPilot reply header (value = "true", not "holding").</summary>
    Private Function CheckConversationNodeForSubstantiveReply(
            conv As Outlook.Conversation,
            item As Object,
            depth As Integer) As Boolean

        If depth > AP_MaxThreadDepth Then Return False

        Try
            Dim nodeMail = TryCast(item, MailItem)
            If nodeMail IsNot Nothing Then
                Try
                    Dim prop As Object = Nothing
                    Try
                        prop = nodeMail.PropertyAccessor.GetProperty(AP_LoopHeaderProperty)
                    Catch
                        ' Property not set — normal for user mails
                    End Try
                    ' Only "true" counts as substantive; "holding" does not
                    If prop IsNot Nothing AndAlso CStr(prop).Equals(AP_LoopHeaderValue, StringComparison.OrdinalIgnoreCase) Then
                        Return True
                    End If
                Catch
                End Try
            End If

            Dim children As SimpleItems = conv.GetChildren(item)
            If children IsNot Nothing Then
                For Each child As Object In children
                    If CheckConversationNodeForSubstantiveReply(conv, child, depth + 1) Then
                        Return True
                    End If
                Next
            End If
        Catch
        End Try

        Return False
    End Function

    ''' <summary>Data class for catch-up preview items.</summary>
    ''' <summary>Data class for catch-up preview items.</summary>
    Private Class CatchUpCandidate
        Public Property EntryID As String
        Public Property SenderName As String
        Public Property SenderEmail As String
        Public Property Subject As String
        Public Property ReceivedTime As DateTime

        ''' <summary>True if this mail was already processed but falls within the reprocess lookback window.</summary>
        Public Property IsReprocessCandidate As Boolean = False

        ''' <summary>Optional custom display label (e.g. prefixed with [REPROCESS]).</summary>
        Public Property CustomDisplayLabel As String = Nothing

        ''' <summary>Returns a display label for the checked list box.</summary>
        Public Function ToDisplayLabel() As String
            If Not String.IsNullOrWhiteSpace(CustomDisplayLabel) Then Return CustomDisplayLabel
            Dim timeStr = ReceivedTime.ToString("yyyy-MM-dd HH:mm")
            Dim senderStr = If(Not String.IsNullOrWhiteSpace(SenderName), SenderName, SenderEmail)
            Dim subjectStr = If(Subject.Length > 60, Subject.Substring(0, 57) & "...", Subject)
            Return $"[{timeStr}]  {senderStr}  —  {subjectStr}"
        End Function
    End Class

    ''' <summary>
    ''' Shows a preview dialog listing all catch-up candidate mails with checkboxes.
    ''' New (unprocessed) items are pre-checked. Reprocess candidates are unchecked by default.
    ''' Returns the list of EntryIDs for checked items, or Nothing if cancelled.
    ''' </summary>
    Private Function ShowCatchUpPreviewDialog(candidates As List(Of CatchUpCandidate), lastProcessed As DateTime) As List(Of String)
        ' Build ModelConfig list to feed into MultiModelSelectorForm
        ' We repurpose ModelDescription as the display label and Model as the EntryID
        Dim items As New List(Of ModelConfig)()
        For Each c In candidates
            items.Add(New ModelConfig() With {
                .ModelDescription = c.ToDisplayLabel(),
                .Model = c.EntryID
            })
        Next

        ' Pre-select only NEW (unprocessed) items — reprocess candidates are unchecked
        Dim preselectLabels = candidates.Where(Function(c) Not c.IsReprocessCandidate).
            Select(Function(c) c.ToDisplayLabel()).ToList()

        Dim reprocessCount = candidates.Where(Function(c) c.IsReprocessCandidate).Count()
        Dim instructionText = $"The following {candidates.Count} mail(s) arrived while AutoPilot was inactive. Select the mails to process:"
        If reprocessCount > 0 Then
            instructionText &= $" ({reprocessCount} already-processed mail(s) marked [REPROCESS] are unchecked by default.)"
        End If

        Using dlg As New SharedLibrary.SharedLibrary.MultiModelSelectorForm(
            items, Nothing,
            title:=$"{AN6} AutoPilot — Catch-Up ({candidates.Count} mail(s) since {lastProcessed:yyyy-MM-dd HH:mm})",
            resetChecked:=False,
            preselectMany:=preselectLabels,
            instruction:=instructionText)

            ' Center on screen since there is no parent form in this context
            dlg.StartPosition = FormStartPosition.CenterScreen

            If dlg.ShowDialog() = DialogResult.OK Then
                Dim selected = dlg.SelectedModels
                If selected IsNot Nothing AndAlso selected.Count > 0 Then
                    Return selected.Select(Function(m) m.Model).ToList()
                End If
                Return New List(Of String)()
            End If
        End Using

        Return Nothing
    End Function


    ''' <summary>Queues newly received mail entry IDs for processing.</summary>
    Private Sub AutoPilot_NewMailEx(ByVal EntryIDCollection As String)
        If Not _apActive Then Return
        For Each id In EntryIDCollection.Split(","c)
            Dim trimmed = id.Trim()
            If trimmed.Length > 0 Then
                _apMailQueue.Enqueue(trimmed)
                _apQueueEnqueueTimes.TryAdd(trimmed, DateTime.UtcNow)
            End If
        Next
    End Sub

    ''' <summary>
    ''' Timer callback that periodically checks for queued mails needing holding
    ''' notifications and sends progress updates for the actively-processing mail.
    ''' Uses Interlocked to prevent overlapping executions.
    ''' </summary>
    Private Async Sub NotificationTimerCallback(state As Object)
        ' Guard: skip if AutoPilot is not active or already checking
        If Not _apActive Then Return
        If Interlocked.CompareExchange(_apNotificationCheckRunning, 1, 0) <> 0 Then Return

        Try
            Dim ct = _apCts?.Token
            If ct Is Nothing OrElse ct.Value.IsCancellationRequested Then Return
            Await SendQueuePositionNotificationsAsync(ct.Value)
            Await SendActiveJobProgressNotificationAsync(ct.Value)
        Catch ex As OperationCanceledException
            ' Expected during shutdown
        Catch ex As System.Exception
            ApDashboardLog($"Notification timer error: {ex.Message}", "warn")
        Finally
            Interlocked.Exchange(_apNotificationCheckRunning, 0)
        End Try
    End Sub

    ''' <summary>Processes queued messages until cancellation is requested.</summary>
    Private Async Function AutoPilotProcessingPump(ct As CancellationToken) As Task
        Try
            While Not ct.IsCancellationRequested
                Dim entryId As String = Nothing
                If _apMailQueue.TryDequeue(entryId) Then
                    ' Clean up enqueue time tracking for this mail
                    Dim dummy As DateTime
                    _apQueueEnqueueTimes.TryRemove(entryId, dummy)
                    ' NOTE: Do NOT remove _apQueueNotifiedEntryIds here.
                    ' ProcessIncomingMailAsync needs to check it for the holding-notice
                    ' cooldown bypass. It is cleaned up after processing completes.

                    ' Track the active job for progress notifications
                    _apCurrentProcessingEntryId = entryId
                    _apActiveJobLastNotifiedUtc = DateTime.MinValue

                    Dim pending = _apMailQueue.Count
                    If pending > 0 Then
                        ApDashboardLog($"⏳ {pending} mail(s) queued behind current processing", "step")
                    End If

                    ' Classify and time the current mail
                    Try
                        _apCurrentProcessingCategory = Await SwitchToUi(Function() As String
                                                                            Try
                                                                                Dim ns = Application.GetNamespace("MAPI")
                                                                                Dim obj = ns.GetItemFromID(entryId)
                                                                                If TypeOf obj Is MailItem Then
                                                                                    Dim mi = DirectCast(obj, MailItem)
                                                                                    Try : Return ClassifyMailCategory(mi)
                                                                                    Finally : Marshal.ReleaseComObject(mi) : End Try
                                                                                End If
                                                                            Catch : End Try
                                                                            Return AP_Cat_TextOnly
                                                                        End Function)
                    Catch
                        _apCurrentProcessingCategory = AP_Cat_TextOnly
                    End Try

                    ' Create a per-job CancellationTokenSource linked to the session CTS
                    _apCurrentJobCts = CancellationTokenSource.CreateLinkedTokenSource(ct)
                    Dim jobCt = _apCurrentJobCts.Token

                    _apCurrentProcessingStopwatch = Stopwatch.StartNew()
                    Try
                        Await ProcessIncomingMailAsync(entryId, jobCt)
                    Catch ex As OperationCanceledException When Not ct.IsCancellationRequested
                        ' Job-level abort (not session-level) — log and continue to next mail
                        ApDashboardLog($"⛔ Job aborted by operator for: {entryId.Substring(0, Math.Min(20, entryId.Length))}...", "warn")
                    Finally
                        _apCurrentProcessingStopwatch.Stop()
                        UpdateCategoryTiming(_apCurrentProcessingCategory, _apCurrentProcessingStopwatch.Elapsed.TotalSeconds)
                        ApDashboardLog($"⏱ Completed in {_apCurrentProcessingStopwatch.Elapsed.TotalSeconds:F1}s [{_apCurrentProcessingCategory}]", "step")
                        _apCurrentProcessingStopwatch = Nothing
                        _apCurrentProcessingCategory = Nothing
                        _apCurrentProcessingEntryId = Nothing
                        _apActiveJobLastNotifiedUtc = DateTime.MinValue
                        ' Clean up notification and catch-up tracking AFTER processing is complete
                        _apQueueNotifiedEntryIds.TryRemove(entryId, dummy)
                        Dim dummyCatchUp As Boolean
                        _apCatchUpEntryIds.TryRemove(entryId, dummyCatchUp)
                        Try : _apCurrentJobCts?.Dispose() : Catch : End Try
                        _apCurrentJobCts = Nothing
                    End Try
                Else
                    Await Task.Delay(1000, ct)
                End If
            End While
        Catch ex As OperationCanceledException
            ApDashboardLog("AutoPilot cancelled.", "step")
        Catch ex As System.Exception
            ApDashboardLog("AutoPilot pump error: " & ex.Message, "error")
        End Try
    End Function


    ' ═══════════════════════════════════════════════════════════════════════════
    '  AUTOPILOT LICENSE & PERMISSION GATE
    ' ═══════════════════════════════════════════════════════════════════════════


    ''' <summary>
    ''' Checks whether the current license permits AutoPilot.
    ''' Allowed if:
    '''   (a) Private license in any active state (PrivateActive or PrivateReconfirmNeeded), OR
    '''   (b) Pro license active AND the stored Product ID is in AP_PermittedProProductIds.
    ''' Denied for: None, Legacy, PrivateExpired, ProActive with non-matching product, etc.
    ''' </summary>
    Public Shared Function IsAutoPilotLicenseValid() As Boolean
        Select Case CurrentLicenseState
            Case LicenseState.PrivateActive, LicenseState.PrivateReconfirmNeeded
                ' Private license always permits AutoPilot (non-commercial use)
                Return True

            Case LicenseState.ProActive, LicenseState.TestingProActive
                ' Pro license: only if the product ID is in the permitted list
                Dim productId As String = StoredProductId
                If String.IsNullOrWhiteSpace(productId) Then Return False
                For Each permitted In AP_PermittedProProductIds
                    If productId.Equals(permitted, StringComparison.OrdinalIgnoreCase) Then
                        Return True
                    End If
                Next
                Return False

            Case Else
                ' None, PrivateExpired, Legacy, ProOfflineGrace, or any other state → not permitted
                Return False
        End Select
    End Function

    ''' <summary>
    ''' Checks whether the current user is permitted to run AutoPilot.
    ''' Requires BOTH:
    '''   1. A valid license (checked via IsAutoPilotLicenseValid) — mandatory.
    '''   2. User whitelist (INI_AutoPilot) — "*" allows all licensed users;
    '''      otherwise %USERNAME% must appear in the comma-separated list.
    '''      Empty/whitespace means no users are permitted.
    ''' </summary>
    ''' <returns>True if the current user is allowed; False otherwise.</returns>
    Public Shared Function IsAutoPilotPermitted() As Boolean
        ' ── Diagnostics ──
        Debug.WriteLine("[AutoPilot] ═══ IsAutoPilotPermitted diagnostics ═══")
        Debug.WriteLine($"[AutoPilot] CurrentLicenseState = {CurrentLicenseState} ({CInt(CurrentLicenseState)})")
        Debug.WriteLine($"[AutoPilot] StoredProductId = '{If(StoredProductId, "(Nothing)")}'")
        Debug.WriteLine($"[AutoPilot] AP_PermittedProProductIds = {String.Join(", ", AP_PermittedProProductIds)}")
        Debug.WriteLine($"[AutoPilot] IsAutoPilotLicenseValid() = {IsAutoPilotLicenseValid()}")
        Debug.WriteLine($"[AutoPilot] _context.INI_AutoPilot = '{If(_context.INI_AutoPilot, "(Nothing)")}'")
        Debug.WriteLine($"[AutoPilot] _context.INI_AutoPilot length = {If(_context.INI_AutoPilot IsNot Nothing, _context.INI_AutoPilot.Length.ToString(), "N/A")}")
        Debug.WriteLine($"[AutoPilot] Environment USERNAME = '{Environment.GetEnvironmentVariable("USERNAME")}'")
        Debug.WriteLine($"[AutoPilot] Environment.UserName = '{Environment.UserName}'")
        Debug.WriteLine("[AutoPilot] ═══════════════════════════════════════════")

        ' Gate 1: License check (mandatory, cannot be bypassed)
        If Not IsAutoPilotLicenseValid() Then
            Debug.WriteLine("[AutoPilot] DENIED: IsAutoPilotLicenseValid() returned False")
            Return False
        End If

        ' Gate 2: User whitelist (mandatory — empty = nobody allowed)
        Dim allowedUsers As String = _context.INI_AutoPilot
        If String.IsNullOrWhiteSpace(allowedUsers) Then
            Debug.WriteLine("[AutoPilot] DENIED: INI_AutoPilot is empty/whitespace → no users permitted")
            Return False
        End If
        If allowedUsers.Trim() = "*" Then
            Debug.WriteLine("[AutoPilot] ALLOWED: INI_AutoPilot = '*' → all users permitted")
            Return True
        End If

        Dim currentUser As String = Environment.GetEnvironmentVariable("USERNAME")
        If String.IsNullOrWhiteSpace(currentUser) Then
            Debug.WriteLine("[AutoPilot] DENIED: USERNAME environment variable is empty")
            Return False
        End If

        Debug.WriteLine($"[AutoPilot] Splitting INI_AutoPilot by comma: '{allowedUsers}'")
        Dim users = allowedUsers.Split(","c)
        For Each entry In users
            Dim trimmed = entry.Trim()
            Debug.WriteLine($"[AutoPilot]   comparing '{trimmed}' with '{currentUser}' → {trimmed.Equals(currentUser, StringComparison.OrdinalIgnoreCase)}")
            If trimmed.Length > 0 AndAlso trimmed.Equals(currentUser, StringComparison.OrdinalIgnoreCase) Then
                Debug.WriteLine("[AutoPilot] ALLOWED: username match found")
                Return True
            End If
        Next

        Debug.WriteLine("[AutoPilot] DENIED: no username match in whitelist")
        Return False
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  QUEUE POSITION NOTIFICATION
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Mail weight categories used for internal time estimation.</summary>
    Private Const AP_Cat_TextOnly As String = "text"
    Private Const AP_Cat_Light As String = "light"
    Private Const AP_Cat_HeavyDoc As String = "heavydoc"
    Private Const AP_Cat_HeavyPdf As String = "heavypdf"

    ''' <summary>
    ''' Classifies a mail into a weight category based on its attachments.
    ''' Called on the UI thread (needs COM access).
    ''' </summary>
    Private Function ClassifyMailCategory(mi As MailItem) As String
        Try
            If mi.Attachments.Count = 0 Then Return AP_Cat_TextOnly

            Dim hasHeavyDoc As Boolean = False
            Dim hasHeavyPdf As Boolean = False
            Dim hasAnyProcessable As Boolean = False

            For i As Integer = 1 To mi.Attachments.Count
                Dim att = mi.Attachments(i)
                Try
                    If att.Type = OlAttachmentType.olEmbeddeditem Then Continue For
                    Dim fileName = att.FileName
                    If String.IsNullOrWhiteSpace(fileName) Then Continue For
                    Dim ext = Path.GetExtension(fileName).ToLowerInvariant()
                    Dim size As Long = att.Size

                    For Each heavyExt In AP_HeavyDocExtensions
                        If ext = heavyExt Then hasHeavyDoc = True : hasAnyProcessable = True : Exit For
                    Next
                    If ext = ".pdf" Then
                        hasAnyProcessable = True
                        If size >= AP_HeavyPdfThreshold Then hasHeavyPdf = True
                    End If
                    If Not hasAnyProcessable Then
                        Dim lightExts = {".txt", ".csv", ".html", ".xml", ".json", ".md"}
                        For Each le In lightExts
                            If ext = le Then hasAnyProcessable = True : Exit For
                        Next
                    End If
                Catch
                End Try
            Next

            If hasHeavyPdf Then Return AP_Cat_HeavyPdf
            If hasHeavyDoc Then Return AP_Cat_HeavyDoc
            If hasAnyProcessable Then Return AP_Cat_Light
        Catch
        End Try
        Return AP_Cat_TextOnly
    End Function

    ''' <summary>Returns the estimated seconds for a category, using learned data or seed defaults.</summary>
    Private Function GetCategoryEstimate(category As String) As Double
        Dim timing As (AvgSeconds As Double, SampleCount As Integer) = Nothing
        If _apCategoryTimings.TryGetValue(category, timing) AndAlso timing.SampleCount > 0 Then
            Return timing.AvgSeconds
        End If
        Select Case category
            Case AP_Cat_HeavyPdf : Return AP_SeedSeconds_HeavyPdf
            Case AP_Cat_HeavyDoc : Return AP_SeedSeconds_HeavyDoc
            Case AP_Cat_Light : Return AP_SeedSeconds_LightAttachments
            Case Else : Return AP_SeedSeconds_TextOnly
        End Select
    End Function

    ''' <summary>Updates the rolling average for a category after processing completes.</summary>
    Private Sub UpdateCategoryTiming(category As String, elapsedSeconds As Double)
        _apCategoryTimings.AddOrUpdate(
            category,
            addValueFactory:=Function(k) (elapsedSeconds, 1),
            updateValueFactory:=Function(k, existing)
                                    Dim newCount = existing.SampleCount + 1
                                    If newCount <= 2 Then
                                        Return (((existing.AvgSeconds * existing.SampleCount) + elapsedSeconds) / newCount, newCount)
                                    Else
                                        Const alpha As Double = 0.3
                                        Return (alpha * elapsedSeconds + (1.0 - alpha) * existing.AvgSeconds, newCount)
                                    End If
                                End Function)
    End Sub

    ''' <summary>
    ''' Runs the full AutoPilot pre-filter pipeline against a mail item to determine
    ''' whether it will actually be processed. This is the same set of checks performed
    ''' at the top of ProcessIncomingMailAsync, extracted here so that queue notifications
    ''' are only sent for mails that will genuinely be processed.
    ''' Must be called via SwitchToUi for COM access.
    ''' </summary>
    ''' <returns>True if the mail would pass all filters and be processed.</returns>
    Private Function WouldMailBeProcessed(mi As MailItem, Optional entryId As String = Nothing) As Boolean
        Try
            ' Mailbox filter
            If Not String.IsNullOrWhiteSpace(_apConfig.MonitoredMailbox) Then
                Dim recipientAddress As String = ""
                Try
                    Dim acct = mi.SendUsingAccount
                    If acct IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(acct.SmtpAddress) Then
                        recipientAddress = acct.SmtpAddress
                    Else
                        For Each recip As Recipient In mi.Recipients
                            If recip.Type = CInt(OlMailRecipientType.olTo) Then
                                Dim addr = recip.Address
                                If Not String.IsNullOrWhiteSpace(addr) Then recipientAddress = addr : Exit For
                            End If
                        Next
                    End If
                Catch
                End Try
                If Not _apConfig.MonitoredMailbox.Equals(recipientAddress, StringComparison.OrdinalIgnoreCase) Then Return False
            End If

            Dim mailInfo = ExtractMailInfo(mi)
            If mailInfo Is Nothing Then Return False

            ' Loop prevention
            If mailInfo.HasAutoReplyHeader Then Return False

            ' Auto-reply/OOF detection
            If IsAutoReplyOrOof(mailInfo) Then Return False

            ' Domain/sender filter rules (trust boundary)
            If Not MatchesFilterRules(mailInfo) Then Return False
            If MatchesNegativeFilters(mailInfo) Then Return False

            ' Subject trigger word — allow bypass for existing conversations
            Dim isExistingConversation As Boolean = False
            Try : isExistingConversation = IsPartOfAutoPilotConversation(mi) : Catch : End Try

            If Not String.IsNullOrWhiteSpace(_apConfig.SubjectTriggerWord) Then
                If mailInfo.Subject.IndexOf(_apConfig.SubjectTriggerWord, StringComparison.OrdinalIgnoreCase) < 0 Then
                    If Not isExistingConversation Then Return False
                End If
            End If

            ' Cooldown — bypass if we already sent a holding notice or this is a catch-up mail
            Dim mailSentUtc As DateTime = mailInfo.SentOn.ToUniversalTime()
            If IsSenderOnCooldown(mailInfo.SenderEmail, mailSentUtc) Then
                Dim hasHoldingCommitment = (entryId IsNot Nothing AndAlso
                    (_apQueueNotifiedEntryIds.ContainsKey(entryId) OrElse _apHoldingOnlyEntryIds.ContainsKey(entryId)))
                Dim isCatchUpMail = (entryId IsNot Nothing AndAlso _apCatchUpEntryIds.ContainsKey(entryId))
                If Not hasHoldingCommitment AndAlso Not isCatchUpMail Then Return False
            End If

            ' Session limit
            If _apConfig.MaxRepliesPerSession > 0 AndAlso _apSessionReplyCount >= _apConfig.MaxRepliesPerSession Then Return False

            ' Thread depth
            If mailInfo.ThreadAIReplyCount >= AP_MaxThreadDepth Then Return False

            Return True
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Checks all currently queued mails and sends a holding notification to any
    ''' that have been waiting in the queue for at least AP_QueueNotifyThresholdSeconds
    ''' of real wall-clock time. The message includes the queue position and a
    ''' qualitative description of expected wait based on attachment classification.
    ''' Only mails that pass the full filter pipeline receive a notification.
    ''' Called by the periodic notification timer.
    ''' </summary>
    Private Async Function SendQueuePositionNotificationsAsync(ct As CancellationToken) As Task
        Dim queueSnapshot = _apMailQueue.ToArray()
        If queueSnapshot.Length = 0 Then Return

        ' Classify every queued mail and build position/weight info
        Dim queueEntries As New List(Of (EntryId As String, Position As Integer, Category As String))()
        For i As Integer = 0 To queueSnapshot.Length - 1
            Dim queuedId = queueSnapshot(i)
            Dim qCat As String = AP_Cat_TextOnly
            Try
                qCat = Await SwitchToUi(Function() As String
                                            Try
                                                Dim ns = Application.GetNamespace("MAPI")
                                                Dim obj = ns.GetItemFromID(queuedId)
                                                If TypeOf obj Is MailItem Then
                                                    Dim mi = DirectCast(obj, MailItem)
                                                    Try : Return ClassifyMailCategory(mi)
                                                    Finally : Marshal.ReleaseComObject(mi) : End Try
                                                End If
                                            Catch : End Try
                                            Return AP_Cat_TextOnly
                                        End Function)
            Catch : End Try
            queueEntries.Add((queuedId, i + 1, qCat))
        Next

        ' Include the currently-processing mail's category for "ahead" description
        Dim currentIsHeavy As Boolean = (_apCurrentProcessingCategory = AP_Cat_HeavyDoc OrElse _apCurrentProcessingCategory = AP_Cat_HeavyPdf)

        Dim now = DateTime.UtcNow

        For Each entry In queueEntries
            ' Only notify mails that have been waiting >= threshold seconds
            Dim enqueueTime As DateTime = Nothing
            If Not _apQueueEnqueueTimes.TryGetValue(entry.EntryId, enqueueTime) Then Continue For
            Dim waitedSeconds = (now - enqueueTime).TotalSeconds
            If waitedSeconds < AP_QueueNotifyThresholdSeconds Then Continue For

            ' Allow repeat notifications every AP_QueueNotifyRepeatIntervalSeconds
            Dim lastNotifiedUtc As DateTime = DateTime.MinValue
            If _apQueueNotifiedEntryIds.TryGetValue(entry.EntryId, lastNotifiedUtc) Then
                Dim sinceLastNotify = (now - lastNotifiedUtc).TotalSeconds
                If sinceLastNotify < AP_QueueNotifyRepeatIntervalSeconds Then Continue For
            End If

            ' Record / update notification time
            _apQueueNotifiedEntryIds(entry.EntryId) = now

            Try
                Dim mi As MailItem = Await SwitchToUi(Function() As MailItem
                                                          Try
                                                              Dim ns = Application.GetNamespace("MAPI")
                                                              Dim obj = ns.GetItemFromID(entry.EntryId)
                                                              If TypeOf obj Is MailItem Then Return DirectCast(obj, MailItem)
                                                          Catch : End Try
                                                          Return Nothing
                                                      End Function)
                If mi Is Nothing Then Continue For

                Try
                    ' Full filter pipeline check
                    Dim wouldProcess = Await SwitchToUi(Function() WouldMailBeProcessed(mi, entry.EntryId))

                    If Not wouldProcess Then
                        ApDashboardLog($"Queue notification skipped (would not be processed): {entry.EntryId.Substring(0, Math.Min(20, entry.EntryId.Length))}...", "step")
                        Continue For
                    End If

                    Dim mailInfo = Await SwitchToUi(Function() ExtractMailInfo(mi))
                    Dim senderLabel = If(mailInfo IsNot Nothing, mailInfo.SenderEmail, "unknown")

                    ' Count heavy jobs ahead of this mail (including current processing)
                    Dim heavyAhead As Integer = If(currentIsHeavy, 1, 0)
                    For Each other In queueEntries
                        If other.Position >= entry.Position Then Exit For
                        If other.Category = AP_Cat_HeavyDoc OrElse other.Category = AP_Cat_HeavyPdf Then heavyAhead += 1
                    Next

                    ' Build position text
                    Dim positionText = If(entry.Position = 1,
                        "Your request is next in line.",
                        $"Your request is at position {entry.Position} in the queue.")

                    ' Build qualitative wait description based on what's ahead
                    Dim waitDescription As String = ""
                    If heavyAhead > 0 Then
                        If heavyAhead = 1 Then
                            waitDescription = " There is a document processing request ahead of yours which may take a while to complete."
                        Else
                            waitDescription = $" There are {heavyAhead} document processing requests ahead of yours which may each take a while to complete."
                        End If
                    ElseIf entry.Position > 3 Then
                        waitDescription = " There are several requests ahead of yours."
                    End If

                    Dim stallWarning As String =
                        " Please also note that occasional delays may occur if the system requires operator attention."

                    Dim localChatHint As String =
                        $" If you need immediate assistance, you can also use the {AN} add-in's corresponding feature to have your tasks done right away. Use 'Help me, Inky' or the chatbot on https://redink.ai if you need instructions. Last but not least, a similar 'agent mode' is available in the Local Chat feature in the Outlook add-in if configured accordingly."

                    Dim isRepeat = (lastNotifiedUtc <> DateTime.MinValue)

                    Dim holdingMessage As String
                    If isRepeat Then
                        holdingMessage =
                            $"Thank you for your continued patience. This is an automated queue status update — your request is still awaiting processing. " &
                            $"{positionText}{waitDescription}{stallWarning}{localChatHint} " &
                            $"I will get back to you with a substantive reply as soon as possible. — {AN6}"
                    Else
                        holdingMessage =
                            $"Thank you for your message. This is an automated queue status notification — your request has not yet been processed in substance. " &
                            $"{positionText}{waitDescription}{stallWarning}{localChatHint} " &
                            $"I will get back to you with a substantive reply as soon as possible. — {AN6}"
                    End If

                    Await SwitchToUi(Sub() SendReplyToSender(mi, holdingMessage, Nothing, tagAsAutoReply:=True, isHoldingOnly:=True))

                    ' Track that this mail has only received a holding notice
                    _apHoldingOnlyEntryIds.TryAdd(entry.EntryId, True)

                    Dim repeatTag = If(isRepeat, " [repeat]", "")
                    ApDashboardLog($"📨 Queue notification{repeatTag} sent to {senderLabel} (position {entry.Position}, waited {waitedSeconds:F0}s, {heavyAhead} heavy ahead)", "info")
                Finally
                    Try : Marshal.ReleaseComObject(mi) : Catch : End Try
                End Try
            Catch ex As System.Exception
                ApDashboardLog($"Queue notification error: {ex.Message}", "warn")
            End Try
        Next
    End Function

    ''' <summary>
    ''' Sends a progress notification to the sender of the mail that is currently
    ''' being processed, if it has been running for at least AP_ActiveJobNotifyIntervalSeconds
    ''' since the last notification (or since processing started).
    ''' This covers the case where a single job (e.g. heavy document processing) takes
    ''' a very long time — the sender is informed that their request is still active.
    ''' </summary>
    Private Async Function SendActiveJobProgressNotificationAsync(ct As CancellationToken) As Task
        Dim entryId = _apCurrentProcessingEntryId
        If String.IsNullOrEmpty(entryId) Then Return

        ' Check elapsed processing time
        Dim sw = _apCurrentProcessingStopwatch
        If sw Is Nothing OrElse Not sw.IsRunning Then Return
        Dim elapsedSeconds = sw.Elapsed.TotalSeconds

        ' Don't notify before the first interval has passed
        If elapsedSeconds < AP_ActiveJobNotifyIntervalSeconds Then Return

        ' Check if enough time has passed since the last notification for this job
        Dim now = DateTime.UtcNow
        If _apActiveJobLastNotifiedUtc <> DateTime.MinValue Then
            Dim sinceLastNotify = (now - _apActiveJobLastNotifiedUtc).TotalSeconds
            If sinceLastNotify < AP_ActiveJobNotifyIntervalSeconds Then Return
        End If

        ' Record notification time
        _apActiveJobLastNotifiedUtc = now

        Try
            Dim mi As MailItem = Await SwitchToUi(Function() As MailItem
                                                      Try
                                                          Dim ns = Application.GetNamespace("MAPI")
                                                          Dim obj = ns.GetItemFromID(entryId)
                                                          If TypeOf obj Is MailItem Then Return DirectCast(obj, MailItem)
                                                      Catch : End Try
                                                      Return Nothing
                                                  End Function)
            If mi Is Nothing Then Return

            Try
                Dim mailInfo = Await SwitchToUi(Function() ExtractMailInfo(mi))
                Dim senderLabel = If(mailInfo IsNot Nothing, mailInfo.SenderEmail, "unknown")

                Dim elapsedMinutes = CInt(Math.Floor(elapsedSeconds / 60))
                Dim progressMessage As String =
                    $"Thank you for your continued patience. This is an automated progress update — your request is currently being processed and has been running for approximately {elapsedMinutes} minutes. " &
                    $"Some requests, particularly those involving document processing, may take considerable time to complete. " &
                    $"I will get back to you with the final result as soon as processing is finished. " &
                    $"If you need immediate assistance, you can also use the {AN} add-in's corresponding feature to have your tasks done right away. Use 'Help me, Inky' or the chatbot on https://redink.ai if you need instructions. " &
                    $"— {AN6}"

                Await SwitchToUi(Sub() SendReplyToSender(mi, progressMessage, Nothing, tagAsAutoReply:=True, isHoldingOnly:=True))

                ' Track that this mail has only received holding notices so far
                _apHoldingOnlyEntryIds.TryAdd(entryId, True)

                ApDashboardLog($"📨 Active job progress notification sent to {senderLabel} (processing for {elapsedMinutes}min)", "info")
            Finally
                Try : Marshal.ReleaseComObject(mi) : Catch : End Try
            End Try
        Catch ex As System.Exception
            ApDashboardLog($"Active job notification error: {ex.Message}", "warn")
        End Try
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  CORE PROCESSING
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Processes a single incoming mail item by entry ID.</summary>
    Private Async Function ProcessIncomingMailAsync(entryId As String, ct As CancellationToken) As Task
        Dim mi As MailItem = Nothing
        Dim jobAborted As Boolean = False
        Dim mailInfoForAbort As AutoPilotMailInfo = Nothing
        Dim tempDirForAbort As String = Nothing
        Dim attachmentPathsForAbort As List(Of AutoPilotAttachmentInfo) = Nothing

        Try
            mi = Await SwitchToUi(Function() As MailItem
                                      Try
                                          Dim ns = Application.GetNamespace("MAPI")
                                          Dim obj = ns.GetItemFromID(entryId)
                                          If TypeOf obj Is MailItem Then Return DirectCast(obj, MailItem)
                                      Catch
                                      End Try
                                      Return Nothing
                                  End Function)
            If mi Is Nothing Then
                ApDashboardLog("SKIP (could not resolve mail item from EntryID — may have been moved or is not a MailItem)", "warn")
                Return
            End If

            ' Mailbox filter
            If Not String.IsNullOrWhiteSpace(_apConfig.MonitoredMailbox) Then
                Dim recipientAddress As String = ""
                Try
                    recipientAddress = Await SwitchToUi(Function() As String
                                                            Try
                                                                Dim acct = mi.SendUsingAccount
                                                                If acct IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(acct.SmtpAddress) Then Return acct.SmtpAddress
                                                                For Each recip As Recipient In mi.Recipients
                                                                    If recip.Type = CInt(OlMailRecipientType.olTo) Then
                                                                        Dim addr = recip.Address
                                                                        If Not String.IsNullOrWhiteSpace(addr) Then Return addr
                                                                    End If
                                                                Next
                                                            Catch
                                                            End Try
                                                            Return ""
                                                        End Function)
                Catch
                End Try
                If Not _apConfig.MonitoredMailbox.Equals(recipientAddress, StringComparison.OrdinalIgnoreCase) Then
                    ApDashboardLog("SKIP (wrong mailbox: " & If(String.IsNullOrWhiteSpace(recipientAddress), "(empty)", recipientAddress) & ")", "step")
                    Return
                End If
            End If

            Dim mailInfo = Await SwitchToUi(Function() ExtractMailInfo(mi))
            If mailInfo Is Nothing Then
                ApDashboardLog("SKIP (ExtractMailInfo returned Nothing — COM error or unexpected mail format)", "warn")
                Return
            End If

            ' ── Voicemail detection (bypasses normal filter pipeline) ──
            If IsVoicemailFromRegisteredSender(mailInfo) Then
                Await ProcessVoicemailAsync(mi, mailInfo, entryId, ct)
                Return
            End If

            ' ── Filter checks ──            

            ' Loop prevention: skip if this mail itself is an auto-reply from us
            If mailInfo.HasAutoReplyHeader Then : ApDashboardLog("SKIP (own auto-reply): " & mailInfo.Subject, "step") : Return : End If

            ' Conversation awareness: detect if this is a follow-up in an existing AutoPilot thread.
            ' This is needed so that second/subsequent user replies to an AutoPilot response are
            ' still processed, even though the user's mail itself has no X-RedInk-AutoReply header.
            Dim isExistingConversation As Boolean = False
            Try
                isExistingConversation = IsPartOfAutoPilotConversation(mi)
            Catch : End Try

            If IsAutoReplyOrOof(mailInfo) Then : ApDashboardLog("SKIP (auto-reply/OOF): " & mailInfo.Subject, "step") : Return : End If

            ' SECURITY: Filter rules (domain/sender) are NEVER bypassed — they are the trust boundary.
            ' An attacker could spoof a ConversationID (same subject line) to join an existing thread,
            ' so we must always verify the sender is permitted regardless of conversation membership.
            If Not MatchesFilterRules(mailInfo) Then : ApDashboardLog("SKIP (no filter match): " & mailInfo.SenderEmail & " — " & mailInfo.Subject, "step") : Return : End If
            If MatchesNegativeFilters(mailInfo) Then : ApDashboardLog("SKIP (negative filter): " & mailInfo.SenderEmail & " — " & mailInfo.Subject, "step") : Return : End If

            ' Subject trigger word CAN be bypassed for existing conversations — reply subjects
            ' often have "Re: Re: ..." prefixes or the trigger word gets stripped in threading.
            If Not String.IsNullOrWhiteSpace(_apConfig.SubjectTriggerWord) Then
                If Not mailInfo.Subject.IndexOf(_apConfig.SubjectTriggerWord, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    If isExistingConversation Then
                        ApDashboardLog("TRIGGER BYPASS (existing conversation): " & mailInfo.Subject, "step")
                    Else
                        ApDashboardLog("SKIP (no trigger word): " & mailInfo.Subject, "step") : Return
                    End If
                End If
            End If

            ' Convert SentOn to UTC for cooldown comparison.
            ' SentOn is the time the sender pressed Send — the only independent timestamp
            ' that correctly represents whether mails were sent in quick succession.
            Dim mailSentUtc As DateTime = mailInfo.SentOn.ToUniversalTime()

            If IsSenderOnCooldown(mailInfo.SenderEmail, mailSentUtc) Then
                ' Bypass cooldown if:
                '   (a) A holding notice was already sent for this mail (we committed to processing it), OR
                '   (b) This mail was explicitly selected in the catch-up/reprocess dialog.
                Dim hasHoldingCommitment = _apQueueNotifiedEntryIds.ContainsKey(entryId) OrElse
                                           _apHoldingOnlyEntryIds.ContainsKey(entryId)
                Dim isCatchUpMail = _apCatchUpEntryIds.ContainsKey(entryId)
                If hasHoldingCommitment Then
                    ApDashboardLog("COOLDOWN BYPASS (holding notice already sent): " & mailInfo.SenderEmail, "step")
                ElseIf isCatchUpMail Then
                    ApDashboardLog("COOLDOWN BYPASS (catch-up/reprocess mail): " & mailInfo.SenderEmail, "step")
                Else
                    ApDashboardLog($"SKIP (cooldown): {mailInfo.SenderEmail} — SentOn {mailSentUtc:HH:mm:ss} UTC", "step")
                    Return
                End If
            End If

            If _apConfig.MaxRepliesPerSession > 0 AndAlso _apSessionReplyCount >= _apConfig.MaxRepliesPerSession Then : ApDashboardLog("SKIP (global session limit " & _apConfig.MaxRepliesPerSession.ToString() & " reached — " & _apSessionReplyCount.ToString() & " replies sent across all senders): " & mailInfo.SenderEmail & " — " & mailInfo.Subject, "warn") : Return : End If

            If mailInfo.ThreadAIReplyCount >= AP_MaxThreadDepth Then : ApDashboardLog("SKIP (thread depth " & mailInfo.ThreadAIReplyCount.ToString() & " >= " & AP_MaxThreadDepth.ToString() & "): " & mailInfo.Subject, "warn") : Return : End If

            ' ── Reply-To mismatch: potential spoofing or reflection attack ──
            ' If the mail has a Reply-To that differs from the sender, log a warning
            ' but do NOT force approval — that would block the entire processing queue.
            ' The sender/domain filter rules are the trust boundary; Reply-To mismatches
            ' are common with legitimate mailing lists, shared mailboxes, and CRM systems.
            Try
                Dim replyToAddress = Await SwitchToUi(Function() As String
                                                          Try
                                                              If mi.ReplyRecipients IsNot Nothing AndAlso mi.ReplyRecipients.Count > 0 Then
                                                                  Return mi.ReplyRecipients(1).Address
                                                              End If
                                                          Catch
                                                          End Try
                                                          Return ""
                                                      End Function)
                If Not String.IsNullOrWhiteSpace(replyToAddress) AndAlso
                   Not replyToAddress.Equals(mailInfo.SenderEmail, StringComparison.OrdinalIgnoreCase) Then
                    ApDashboardLog($"⚠ Reply-To mismatch: From={mailInfo.SenderEmail}, Reply-To={replyToAddress} (processing continues)", "warn")
                End If
            Catch
            End Try

            Dim isWhitelisted As Boolean = IsSenderWhitelisted(mailInfo.SenderEmail)
            ' Existing conversations: skip approval only (sender already passed domain/sender filter above)
            Dim requiresApproval As Boolean = (_apConfig.RequireApprovalForNonWhitelisted AndAlso Not isWhitelisted AndAlso Not isExistingConversation)
            ApDashboardLog("━━━ PROCESSING ━━━", "info")
            SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "AutoPilot (Mail) invoked")
            ApDashboardLog($"From: {mailInfo.SenderName} <{mailInfo.SenderEmail}>", "info")
            ApDashboardLog($"Subject: {mailInfo.Subject}", "info")
            ApDashboardLog($"Attachments: {mailInfo.AttachmentCount}" & If(requiresApproval, " [approval required]", " [auto-send]"), "info")

            ' Check for #model: command
            Dim modelOverrideConfig As ModelConfig = Nothing
            Dim modelOverrideName As String = Nothing
            Dim strippedBody As String = mailInfo.Body
            TryParseModelOverride(mailInfo.Body, modelOverrideConfig, modelOverrideName, strippedBody)
            If modelOverrideConfig IsNot Nothing Then
                ApDashboardLog($"Model override: {modelOverrideName}", "info")
                mailInfo.Body = strippedBody ' Remove the #model: line from the body sent to the LLM
            End If

            ' ── Extract attachments to temp ──
            Dim tempDir As String = Path.Combine(Path.GetTempPath(), AP_TempPrefix & Guid.NewGuid().ToString("N"))
            Directory.CreateDirectory(tempDir)
            Dim attachmentPaths As New List(Of AutoPilotAttachmentInfo)()

            ' Store references for abort handler
            mailInfoForAbort = mailInfo
            tempDirForAbort = tempDir

            Try
                attachmentPaths = Await SwitchToUi(Function() SaveAttachmentsToTemp(mi, tempDir))
                attachmentPathsForAbort = attachmentPaths

                If attachmentPaths.Count > 0 Then
                    ApDashboardLog($"Saved {attachmentPaths.Count} attachment(s) to temp:", "info")
                    For Each att In attachmentPaths
                        Dim sizeStr = If(att.SizeBytes > 0, $"{att.SizeBytes / 1024:F0} KB", "?")
                        Dim statusStr = If(att.IsOverSizeLimit, " [OVER LIMIT]", "")
                        ApDashboardLog($"  • {att.OriginalFileName} ({att.Extension}, {sizeStr}){statusStr}", "info")
                    Next
                End If

                ' ── Early exit: ALL attachments over size limit ──
                ' If the mail has attachments and every single one exceeds the size limit,
                ' skip the LLM call entirely and send a static rejection message.
                ' This prevents the LLM from generating unwanted advice (Red Ink features,
                ' third-party tools) when the only issue is file size.
                Dim oversizedAttachments = attachmentPaths.Where(Function(a) a.IsOverSizeLimit).ToList()
                If attachmentPaths.Count > 0 AndAlso oversizedAttachments.Count = attachmentPaths.Count Then
                    Dim limitMB = _apConfig.MaxAttachmentBytes / 1024.0 / 1024.0
                    Dim oversizedSb As New StringBuilder()
                    oversizedSb.AppendLine($"I'm sorry, but I was unable to process your request because all attachments exceed the maximum permitted size of {limitMB:F0} MB:")
                    oversizedSb.AppendLine()
                    For Each att In oversizedAttachments
                        oversizedSb.AppendLine($"  - {att.OriginalFileName} ({att.SizeBytes / 1024.0 / 1024.0:F1} MB)")
                    Next
                    oversizedSb.AppendLine()
                    oversizedSb.AppendLine($"Please split large documents into smaller parts and resend, or use the {AN} add-in locally to process the file(s) directly on your computer.")
                    oversizedSb.AppendLine()
                    oversizedSb.Append($"— {AN6}")

                    Dim oversizedMessage = oversizedSb.ToString()
                    Await SwitchToUi(Sub() SendReplyToSender(mi, oversizedMessage, Nothing, tagAsAutoReply:=True))
                    Await SwitchToUi(Sub() TagOriginalMailAsProcessed(mi))
                    Interlocked.Increment(_apSessionReplyCount)
                    RecordSenderCooldown(mailInfo.SenderEmail, mailSentUtc)
                    RecordLastProcessedTime()
                    ApDashboardLog($"✉ Sent size-limit rejection to: {mailInfo.SenderEmail} (all {attachmentPaths.Count} attachment(s) over limit)", "info")
                    Return
                End If

                ' ── Partial oversized: warn about skipped attachments in the prompt ──
                ' When some attachments are processable but others exceed the limit,
                ' build the prompt using only the processable ones. The oversized files
                ' are mentioned as a brief note so the LLM can inform the sender, but
                ' WITHOUT the [OVER SIZE LIMIT] tag that causes the LLM to generate
                ' unwanted tool/service recommendations.
                Dim processableAttachments = attachmentPaths.Where(Function(a) Not a.IsOverSizeLimit).ToList()
                Dim oversizedNote As String = Nothing
                If oversizedAttachments.Count > 0 AndAlso processableAttachments.Count > 0 Then
                    Dim limitMB = _apConfig.MaxAttachmentBytes / 1024.0 / 1024.0
                    Dim noteSb As New StringBuilder()
                    noteSb.AppendLine()
                    noteSb.AppendLine("[NOTE TO ASSISTANT — include this information in your reply to the sender]")
                    noteSb.AppendLine($"The following attachment(s) could NOT be processed because they exceed the {limitMB:F0} MB size limit:")
                    For Each att In oversizedAttachments
                        noteSb.AppendLine($"  - {att.OriginalFileName} ({att.SizeBytes / 1024.0 / 1024.0:F1} MB)")
                    Next
                    noteSb.AppendLine($"Tell the sender to split these files into smaller parts and resend them, or to use the {AN} add-in locally.")
                    noteSb.AppendLine("Do NOT suggest any other tools, services, or workarounds for the oversized files.")
                    noteSb.AppendLine("[/NOTE]")
                    oversizedNote = noteSb.ToString()
                    ApDashboardLog($"Mixed attachments: {processableAttachments.Count} processable, {oversizedAttachments.Count} over limit", "info")
                End If

                ' Initialize tool call log for this e-mail
                _apCurrentToolCallLog = New List(Of AutoPilotToolCallEntry)()

                ' ── Build LLM prompt ──
                ' Use only processable attachments in the prompt so the LLM doesn't see
                ' [OVER SIZE LIMIT] tags that trigger unwanted tool recommendations.
                Dim userPrompt As String = BuildUserPromptFromMail(mailInfo, processableAttachments)
                If oversizedNote IsNot Nothing Then
                    userPrompt &= oversizedNote
                End If
                Dim systemPrompt As String = InterpolateAtRuntime(SP_AutoPilot)

                ' ── Inject per-user memory into system prompt ──
                If _apConfig.EnableUserMemory AndAlso IsUserMemoryEnabled(mailInfo.SenderEmail) Then
                    Dim userMemory = ReadUserMemory(mailInfo.SenderEmail, _context.INI_InkyMemoryCap)
                    systemPrompt &= vbLf & _context.SP_Add_InkyMemory
                    If Not String.IsNullOrWhiteSpace(userMemory) Then
                        systemPrompt &= vbLf & "<INKY_MEMORY_CURRENT>" & vbLf & userMemory & vbLf & "</INKY_MEMORY_CURRENT>"
                    End If
                End If

                ' ── Inject user home file listing into user prompt ──
                If _apConfig.EnableUserFiles AndAlso HasUserHomeFiles(mailInfo.SenderEmail) Then
                    Dim homeFiles = ListUserHomeFiles(mailInfo.SenderEmail)
                    If homeFiles.Count > 0 Then
                        Dim fileListing As New StringBuilder()
                        fileListing.AppendLine()
                        fileListing.AppendLine("[USER HOME FILES — persistent files stored by this user for use in requests]")
                        For Each f In homeFiles
                            fileListing.AppendLine($"  - {f.Name} ({f.SizeBytes / 1024:F0} KB)")
                        Next
                        fileListing.AppendLine("Use manage_user_files with action='use' to load a file into the current session for processing.")
                        fileListing.AppendLine("[/USER HOME FILES]")
                        userPrompt &= fileListing.ToString()
                    End If
                End If

                ' ── Set AutoPilot tool context ──
                _apCurrentTempDir = tempDir
                _apCurrentAttachments = attachmentPaths
                _apCurrentMailInfo = mailInfo

                Dim response As String

                ' ── Set AutoPilot-specific max tool iterations ──
                Dim previousMaxToolIterations = MaxToolIterations
                MaxToolIterations = AP_MaxToolIterations

                Try
                    ApDashboardLog("Calling AI model...", "llm")

                    ' Re-apply the base model config captured at startup.
                    ' This ensures all tooling templates, detection patterns, and API call
                    ' structures are pristine, regardless of what ExecuteToolingLoop's internal
                    ' ApplyModelConfig/RestoreDefaults cycles did during the previous mail.
                    ApplyModelConfig(_context, _apBaseModelConfig)

                    ' Apply model override on top if requested
                    Dim overrideSupportsTooling As Boolean = True
                    If modelOverrideConfig IsNot Nothing Then
                        ' Check if the override model supports tooling
                        overrideSupportsTooling = ModelSupportsTooling(modelOverrideConfig)

                        If overrideSupportsTooling Then
                            ' Override model has its own tooling config — apply the full config
                            ' so its detection patterns, extraction maps, and templates are used.
                            ApplyModelConfig(_context, modelOverrideConfig)
                            ApDashboardLog($"Applied full model override: {modelOverrideName} (tooling supported)", "info")
                        Else
                            ' Override model does NOT support tooling — apply only non-tooling
                            ' fields, and disable tools for this mail.
                            _context.INI_Model_2 = If(Not String.IsNullOrEmpty(modelOverrideConfig.Model), modelOverrideConfig.Model, _context.INI_Model_2)
                            _context.INI_Endpoint_2 = If(Not String.IsNullOrEmpty(modelOverrideConfig.Endpoint), modelOverrideConfig.Endpoint, _context.INI_Endpoint_2)
                            _context.INI_APIKey_2 = If(Not String.IsNullOrEmpty(modelOverrideConfig.APIKey), modelOverrideConfig.APIKey, _context.INI_APIKey_2)
                            _context.INI_APIKeyBack_2 = If(Not String.IsNullOrEmpty(modelOverrideConfig.APIKeyBack), modelOverrideConfig.APIKeyBack, _context.INI_APIKeyBack_2)
                            _context.INI_HeaderA_2 = If(Not String.IsNullOrEmpty(modelOverrideConfig.HeaderA), modelOverrideConfig.HeaderA, _context.INI_HeaderA_2)
                            _context.INI_HeaderB_2 = If(Not String.IsNullOrEmpty(modelOverrideConfig.HeaderB), modelOverrideConfig.HeaderB, _context.INI_HeaderB_2)
                            _context.INI_Temperature_2 = If(Not String.IsNullOrEmpty(modelOverrideConfig.Temperature), modelOverrideConfig.Temperature, _context.INI_Temperature_2)
                            _context.INI_MaxOutputToken_2 = If(modelOverrideConfig.MaxOutputToken <> 0, modelOverrideConfig.MaxOutputToken, _context.INI_MaxOutputToken_2)
                            _context.INI_Timeout_2 = If(modelOverrideConfig.Timeout <> 0, modelOverrideConfig.Timeout, _context.INI_Timeout_2)
                            _context.INI_APIEncrypted_2 = modelOverrideConfig.APIEncrypted
                            _context.INI_APIKeyPrefix_2 = If(Not String.IsNullOrEmpty(modelOverrideConfig.APIKeyPrefix), modelOverrideConfig.APIKeyPrefix, _context.INI_APIKeyPrefix_2)
                            _context.DecodedAPI_2 = If(Not String.IsNullOrEmpty(modelOverrideConfig.DecodedAPI), modelOverrideConfig.DecodedAPI, _context.DecodedAPI_2)
                            _context.INI_Response_2 = If(Not String.IsNullOrEmpty(modelOverrideConfig.Response), modelOverrideConfig.Response, _context.INI_Response_2)
                            _context.INI_APICall_2 = If(Not String.IsNullOrEmpty(modelOverrideConfig.APICall), modelOverrideConfig.APICall, _context.INI_APICall_2)
                            _context.INI_APICall_Object_2 = If(Not String.IsNullOrEmpty(modelOverrideConfig.APICall_Object), modelOverrideConfig.APICall_Object, _context.INI_APICall_Object_2)
                            ApDashboardLog($"Applied model override: {modelOverrideName} (⚠ no tooling support — running without tools for this mail)", "warn")
                        End If
                    End If

                    ' Determine whether to use tooling for this specific mail
                    Dim useToolsForThisMail As Boolean = (modelOverrideConfig Is Nothing OrElse overrideSupportsTooling)

                    If useToolsForThisMail AndAlso _apSelectedTools IsNot Nothing AndAlso _apSelectedTools.Count > 0 Then
                        ' Apply the additive per-sender tool policy (no-op when no policy is configured).
                        Dim senderTools As List(Of ModelConfig) = ResolveToolsForSender(mailInfo.SenderEmail, _apSelectedTools)
                        response = Await ExecuteToolingLoop(
                            systemPrompt, userPrompt,
                            senderTools, _apUseSecondApi,
                            hideSplash:=True, hideLogWindow:=True,
                            cancellationToken:=ct,
                            binaryOutputDirectory:=_apCurrentTempDir,
                            workflowId:=SharedLibrary.Agents.WorkflowContinuity.CreateWorkflowId(),
                            memoryGroundingMode:=SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingMode.None,
                            memoryGroundingModeIsExplicit:=True)
                    Else
                        ' Use no-tools system prompt when tooling is disabled
                        Dim effectiveSystemPrompt = If(useToolsForThisMail, systemPrompt, InterpolateAtRuntime(SP_AutoPilot_NoTools))
                        response = Await LLM(effectiveSystemPrompt, userPrompt,
                                             UseSecondAPI:=_apUseSecondApi,
                                             HideSplash:=True, EnsureUI:=False,
                                             cancellationToken:=ct,
                                             binaryOutputDirectory:=_apCurrentTempDir)
                    End If
                Catch ex As OperationCanceledException
                    ' Check if this is a job-level abort (not session-level)
                    Dim sessionCt = _apCts?.Token
                    If sessionCt.HasValue AndAlso Not sessionCt.Value.IsCancellationRequested Then
                        ' Job-level abort — handle gracefully
                        jobAborted = True
                        response = Nothing
                        ApDashboardLog("⛔ Job aborted during LLM/tooling execution.", "warn")
                    Else
                        Throw ' Session-level cancel — propagate
                    End If
                Finally
                    _apCurrentTempDir = Nothing
                    _apCurrentAttachments = Nothing
                    _apCurrentMailInfo = Nothing
                    MaxToolIterations = previousMaxToolIterations
                    ClearAttachmentCaches()
                End Try

                ' ── Detect swallowed cancellation ──
                ' ExecuteToolingLoop catches OperationCanceledException internally and returns
                ' a string like "Operation was canceled by the user." instead of re-throwing.
                ' LLM() returns "" on cancellation. In both cases, if the job CTS was triggered
                ' we must treat this as an abort, not as valid LLM output.
                If Not jobAborted AndAlso ct.IsCancellationRequested Then
                    jobAborted = True
                    response = Nothing
                    ApDashboardLog("⛔ Job abort detected (cancellation swallowed by LLM/tooling layer).", "warn")
                End If

                ' ── Handle job abort ──
                If jobAborted Then
                    Dim resultAttachmentsAbort = CollectResultAttachments(tempDir, attachmentPaths)
                    Dim hasPartialOutput = (resultAttachmentsAbort IsNot Nothing AndAlso resultAttachmentsAbort.Count > 0)

                    Dim abortChoice = Await SwitchToUi(Function() As Integer
                                                           Dim msg = "The current job has been aborted."
                                                           If hasPartialOutput Then
                                                               msg &= $" There are {resultAttachmentsAbort.Count} output file(s) generated so far." & vbCrLf & vbCrLf &
                                                                      "Would you like to send the partial output with an abort notice, or discard everything?"
                                                           Else
                                                               msg &= vbCrLf & vbCrLf & "Would you like to send an abort notice to the sender, or discard silently?"
                                                           End If
                                                           Return ShowCustomYesNoBox(msg,
                                                               If(hasPartialOutput, "Send Partial Output", "Send Abort Notice"),
                                                               "Discard",
                                                               header:=$"{AN6} AutoPilot — Job Aborted")
                                                       End Function)

                    If abortChoice = 1 Then
                        Dim abortMessage As String =
                            $"I apologize, but the processing of your request was interrupted by the operator before it could be completed. " &
                            If(hasPartialOutput,
                                $"I was able to produce some partial output, which is attached for your reference. Please note that the results may be incomplete. ",
                                "") &
                            $"Please resend your request if you would like it to be fully processed. " &
                            $"If you need immediate assistance, you can also use the {AN} add-in's Local Chat feature directly in Outlook or Word. " &
                            $"— {AN6}"

                        Dim attachmentsToSend = If(hasPartialOutput, resultAttachmentsAbort, Nothing)
                        Await SwitchToUi(Sub() SendReplyToSender(mi, abortMessage, attachmentsToSend, tagAsAutoReply:=True, isHoldingOnly:=True))
                        ' Do NOT tag as processed — allow catch-up to pick it up again
                        Interlocked.Increment(_apSessionReplyCount)
                        RecordSenderCooldown(mailInfo.SenderEmail, mailSentUtc)
                        ApDashboardLog($"✉ Abort notice sent to: {mailInfo.SenderEmail}" & If(hasPartialOutput, $" (with {resultAttachmentsAbort.Count} partial attachment(s))", ""), "info")
                    Else
                        ' Do NOT tag as processed, do NOT record last processed time
                        ApDashboardLog($"Job aborted and discarded for: {mailInfo.SenderEmail}", "step")
                    End If
                    Return
                End If

                If String.IsNullOrWhiteSpace(response) Then
                    ApDashboardLog("WARNING: LLM returned empty response for: " & mailInfo.Subject, "warn")
                    Dim errorMessage = Await GenerateHelpfulFailureResponseAsync(
                        mailInfo, attachmentPaths, "The AI model returned an empty response.", ct)
                    Await SwitchToUi(Sub() SendReplyToSender(mi, errorMessage, Nothing, tagAsAutoReply:=True))
                    Interlocked.Increment(_apSessionReplyCount)
                    RecordSenderCooldown(mailInfo.SenderEmail, mailSentUtc)
                    ApDashboardLog("Sent helpful error response to: " & mailInfo.SenderEmail, "warn")
                    Return
                End If

                ApDashboardLog($"AI response received ({response.Length} chars).", "info")

                ' ── Refined version of the memory processing block ──

                ' ── Process automatic memory learning from LLM response ──
                If _apConfig.EnableUserMemory AndAlso IsUserMemoryEnabled(mailInfo.SenderEmail) Then
                    ' Check if auto-learn is disabled for this user
                    Dim userMemoryContent = ReadUserMemory(mailInfo.SenderEmail, _context.INI_InkyMemoryCap)
                    Dim autoLearnDisabled = (userMemoryContent IsNot Nothing AndAlso
                        userMemoryContent.Contains("AUTO_LEARN_DISABLED"))

                    If Not autoLearnDisabled Then
                        Try
                            response = ProcessUserMemoryResponse(mailInfo.SenderEmail, response, _context.INI_InkyMemoryCap)
                        Catch
                            response = StripInkyMemoryBlock(response)
                        End Try
                    Else
                        ' Auto-learn off — strip the block but don't apply operations
                        response = StripInkyMemoryBlock(response)
                    End If
                Else
                    ' Always strip memory blocks from output even if memory is disabled
                    response = StripInkyMemoryBlock(response)
                End If

                ' Remove knowledge-store source files not cited by the LLM
                RemoveUncitedKnowledgeSourceCopies(response)

                ' Collect result attachments from OutputFiles
                Dim resultAttachments As List(Of String) = CollectResultAttachments(tempDir, attachmentPaths)

                If resultAttachments.Count > 0 Then
                    ApDashboardLog($"Result attachments to send: {resultAttachments.Count}", "info")
                    For Each ra In resultAttachments
                        ApDashboardLog($"  📎 {Path.GetFileName(ra)}", "info")
                    Next
                End If

                ' Build "Sources used:" footer
                Dim sourcesHtml = BuildSourcesUsedHtml(_apCurrentToolCallLog)

                ' ── Approval or auto-send ──
                If requiresApproval AndAlso Not _apConfig.IsUnattended Then
                    Await SwitchToUi(Sub() SendReplyToSender(mi, SP_AutoPilot_HoldingResponse, Nothing, tagAsAutoReply:=True, isHoldingOnly:=True))
                    _apHoldingOnlyEntryIds.TryAdd(entryId, True)
                    ApDashboardLog("Holding response sent to: " & mailInfo.SenderEmail, "step")

                    Dim approved As Boolean = Await SwitchToUi(Function() ShowApprovalDialog(mailInfo, response, resultAttachments))
                    If approved Then
                        Await SwitchToUi(Sub() SendReplyToSender(mi, response, resultAttachments, tagAsAutoReply:=True, sourcesHtml:=sourcesHtml))
                        Await SwitchToUi(Sub() TagOriginalMailAsProcessed(mi))
                        Interlocked.Increment(_apSessionReplyCount)
                        RecordSenderCooldown(mailInfo.SenderEmail, mailSentUtc)
                        RecordLastProcessedTime()
                        Dim dummyBool As Boolean
                        _apHoldingOnlyEntryIds.TryRemove(entryId, dummyBool)
                        ApDashboardLog("✓ APPROVED & SENT reply to: " & mailInfo.SenderEmail, "info")
                    Else
                        ApDashboardLog("REJECTED reply for: " & mailInfo.Subject, "step")
                    End If
                Else
                    If requiresApproval AndAlso _apConfig.IsUnattended Then
                        ApDashboardLog($"⚡ Unattended mode — auto-approving reply for non-whitelisted sender: {mailInfo.SenderEmail}", "info")
                    End If
                    Await SwitchToUi(Sub() SendReplyToSender(mi, response, resultAttachments, tagAsAutoReply:=True, sourcesHtml:=sourcesHtml))
                    Await SwitchToUi(Sub() TagOriginalMailAsProcessed(mi))
                    Interlocked.Increment(_apSessionReplyCount)
                    RecordSenderCooldown(mailInfo.SenderEmail, mailSentUtc)
                    RecordLastProcessedTime()
                    TrackAutoPilotConversation(mi)
                    Dim dummyBool As Boolean
                    _apHoldingOnlyEntryIds.TryRemove(entryId, dummyBool)
                    ApDashboardLog($"✓ SENT reply to: {mailInfo.SenderEmail} (session total: {_apSessionReplyCount})", "info")
                End If
            Finally
                ' ── SECURITY: Clean up temp directory ──
                Try
                    If Directory.Exists(tempDir) Then Directory.Delete(tempDir, recursive:=True)
                Catch
                End Try
                _apCurrentToolCallLog = Nothing
                _apKnowledgeSourceCopies.Clear()
            End Try

        Catch ex As OperationCanceledException
            Throw
        Catch ex As System.Exception
            ApDashboardLog("ERROR: " & ex.Message, "error")
            Debug.WriteLine("AutoPilot ProcessIncomingMailAsync error: " & ex.ToString())
        Finally
            If mi IsNot Nothing Then Try : Marshal.ReleaseComObject(mi) : Catch : End Try
        End Try
    End Function

    ''' <summary>
    ''' Clears all cached text data from attachment info objects for security.
    ''' Must be called after each reply is sent or when processing is abandoned.
    ''' </summary>
    Private Sub ClearAttachmentCaches()
        If _apCurrentAttachments IsNot Nothing Then
            For Each att In _apCurrentAttachments
                att.CachedText = Nothing
                att.CachedDocxHint = Nothing
            Next
        End If
    End Sub

    ''' <summary>
    ''' Tags the original incoming mail with the AP_CategoryName category so that
    ''' the catch-up scan can identify it as already processed.
    ''' Also starts the auto-delete retention clock for the whole cleanup group.
    ''' </summary>
    Private Sub TagOriginalMailAsProcessed(mi As MailItem)
        Try
            Dim existing = mi.Categories
            Dim needsSave As Boolean = False

            If String.IsNullOrWhiteSpace(existing) Then
                mi.Categories = AP_CategoryName
                needsSave = True
            ElseIf existing.IndexOf(AP_CategoryName, StringComparison.OrdinalIgnoreCase) < 0 Then
                mi.Categories = existing & ", " & AP_CategoryName
                needsSave = True
            End If

            If needsSave Then
                mi.Save()
            End If

            MarkMailGroupAsAnsweredAndEligible(mi)
        Catch ex As System.Exception
            Debug.WriteLine($"[AutoPilot] TagOriginalMailAsProcessed error: {ex.Message}")
        End Try
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  MAIL INFO EXTRACTION
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Extracts the core data needed for AutoPilot processing.</summary>
    Private Function ExtractMailInfo(mi As MailItem) As AutoPilotMailInfo
        Try
            Dim info As New AutoPilotMailInfo()
            info.EntryID = mi.EntryID
            info.Subject = If(mi.Subject, "")
            info.SenderName = If(mi.SenderName, "")
            info.ReceivedTime = mi.ReceivedTime
            Try
                info.SentOn = mi.SentOn
            Catch
                info.SentOn = mi.ReceivedTime ' Fallback if SentOn is unavailable
            End Try
            Try
                If mi.SenderEmailType = "EX" Then
                    Dim sender As AddressEntry = mi.Sender
                    If sender IsNot Nothing Then
                        Dim exUser = sender.GetExchangeUser()
                        info.SenderEmail = If(exUser?.PrimarySmtpAddress, mi.SenderEmailAddress)
                    Else
                        info.SenderEmail = mi.SenderEmailAddress
                    End If
                Else
                    info.SenderEmail = If(mi.SenderEmailAddress, "")
                End If
            Catch
                info.SenderEmail = If(mi.SenderEmailAddress, "")
            End Try
            info.Body = If(mi.Body, "")
            Try
                Dim propVal = mi.PropertyAccessor.GetProperty(AP_LoopHeaderProperty)
                info.HasAutoReplyHeader = (propVal IsNot Nothing AndAlso (propVal.ToString() = AP_LoopHeaderValue OrElse propVal.ToString() = AP_LoopHeaderValueHolding))
            Catch
                info.HasAutoReplyHeader = False
            End Try
            info.ThreadAIReplyCount = 0
            Try
                Dim conv = mi.GetConversation()
                If conv IsNot Nothing Then
                    Dim table = conv.GetTable()
                    Dim count As Integer = 0
                    While Not table.EndOfTable
                        Dim row = table.GetNextRow()
                        Try
                            Dim rowCategories = row("Categories")?.ToString()
                            If rowCategories IsNot Nothing AndAlso rowCategories.Contains(AP_CategoryName) Then count += 1
                        Catch
                        End Try
                    End While
                    info.ThreadAIReplyCount = count
                End If
            Catch
            End Try
            info.AttachmentCount = mi.Attachments.Count
            Dim attachNames As New List(Of String)()
            For i As Integer = 1 To mi.Attachments.Count
                attachNames.Add(mi.Attachments(i).FileName)
            Next
            info.AttachmentNames = attachNames
            Try
                Dim parentFolder As MAPIFolder = mi.Parent
                info.FolderPath = parentFolder.FolderPath
            Catch
                info.FolderPath = ""
            End Try
            info.MessageClass = If(mi.MessageClass, "")
            info.InternetHeaders = ""
            Try
                Dim headers = mi.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001F")
                info.InternetHeaders = If(headers?.ToString(), "")
            Catch
            End Try
            Return info
        Catch ex As System.Exception
            Debug.WriteLine($"ExtractMailInfo error: {ex.Message}")
            Return Nothing
        End Try
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  FILTER ENGINE
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Determines whether a mail appears to be an automatic reply or OOF.</summary>
    Private Function IsAutoReplyOrOof(info As AutoPilotMailInfo) As Boolean
        Dim subjectLower = info.Subject.ToLowerInvariant()
        For Each pattern In AP_AutoReplySubjectPatterns
            If subjectLower.StartsWith(pattern) Then Return True
        Next
        If info.MessageClass.IndexOf("OofTemplate", StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
        If info.MessageClass.IndexOf("Report", StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
        If info.InternetHeaders.Length > 0 Then
            Dim headersLower = info.InternetHeaders.ToLowerInvariant()
            If headersLower.Contains("auto-submitted: auto-replied") Then Return True
            If headersLower.Contains("auto-submitted: auto-generated") Then Return True
            If headersLower.Contains("auto-submitted: auto-notified") Then Return True
            If Regex.IsMatch(headersLower, "precedence:\s*(bulk|junk|list)") Then Return True
        End If
        Return False
    End Function

    ''' <summary>Checks whether the mail matches any positive filter rule.</summary>
    Private Function MatchesFilterRules(info As AutoPilotMailInfo) As Boolean
        If _apConfig.FilterRules Is Nothing OrElse _apConfig.FilterRules.Count = 0 Then Return True
        For Each rule In _apConfig.FilterRules
            If rule.IsNegative Then Continue For
            If MatchesSingleRule(info, rule) Then Return True
        Next
        Return False
    End Function

    ''' <summary>Checks whether the mail matches any negative filter rule.</summary>
    Private Function MatchesNegativeFilters(info As AutoPilotMailInfo) As Boolean
        If _apConfig.FilterRules Is Nothing Then Return False
        For Each rule In _apConfig.FilterRules
            If Not rule.IsNegative Then Continue For
            If MatchesSingleRule(info, rule) Then Return True
        Next
        Return False
    End Function

    ''' <summary>Evaluates a single filter rule against the mail info.</summary>
    Private Shared Function MatchesSingleRule(info As AutoPilotMailInfo, rule As AutoPilotFilterRule) As Boolean
        Select Case rule.RuleType
            Case AutoPilotFilterRuleType.Domain
                Dim pattern = rule.Pattern.TrimStart("*"c)
                If info.SenderEmail.EndsWith(pattern, StringComparison.OrdinalIgnoreCase) Then Return True
            Case AutoPilotFilterRuleType.Sender
                If WildcardMatch(info.SenderEmail, rule.Pattern) Then Return True
            ' SECURITY: Do NOT match on SenderName for sender-type rules.
            ' An attacker can set their display name to "john@trusted.com"
            ' which would cause a false positive match against a sender filter.
            ' Display names are unauthenticated and trivially spoofable.
            Case AutoPilotFilterRuleType.Folder
                If info.FolderPath.IndexOf(rule.Pattern, StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
        End Select
        Return False
    End Function

    ''' <summary>Performs wildcard matching using * and ? patterns.</summary>
    Private Shared Function WildcardMatch(input As String, pattern As String) As Boolean
        If String.IsNullOrEmpty(input) OrElse String.IsNullOrEmpty(pattern) Then Return False
        Dim regexPattern = "^" & Regex.Escape(pattern).Replace("\*", ".*").Replace("\?", ".") & "$"
        Return Regex.IsMatch(input, regexPattern, RegexOptions.IgnoreCase)
    End Function

    ''' <summary>Determines whether the sender is in the whitelist.</summary>
    Private Function IsSenderWhitelisted(senderEmail As String) As Boolean
        If _apConfig.WhitelistedSenders Is Nothing OrElse _apConfig.WhitelistedSenders.Count = 0 Then Return False
        ' SECURITY: Only match on the SMTP email address, never on display name.
        ' Display names can be trivially spoofed to look like email addresses.
        If String.IsNullOrWhiteSpace(senderEmail) Then Return False
        For Each pattern In _apConfig.WhitelistedSenders
            If WildcardMatch(senderEmail, pattern) Then Return True
        Next
        Return False
    End Function

    ''' <summary>
    ''' Checks whether the current mail should be skipped due to cooldown.
    ''' Cooldown is based on the mail's SentOn time compared to the SentOn time of the
    ''' last processed mail from the same sender — NOT on ReceivedTime or wall-clock time.
    ''' </summary>
    Private Function IsSenderOnCooldown(senderEmail As String, mailSentTimeUtc As DateTime) As Boolean
        Dim lastMailSentUtc As DateTime
        If _apSenderLastMailSentUtc.TryGetValue(senderEmail, lastMailSentUtc) Then
            Dim delta = (mailSentTimeUtc - lastMailSentUtc).TotalSeconds
            Return delta >= 0 AndAlso delta < _apConfig.CooldownSeconds
        End If
        Return False
    End Function

    ''' <summary>Records the SentOn time of the processed mail for cooldown tracking.</summary>
    Private Sub RecordSenderCooldown(senderEmail As String, mailSentTimeUtc As DateTime)
        ' Only advance the cooldown marker forward — never regress it.
        ' This prevents out-of-order processing from poisoning the cooldown window.
        _apSenderLastMailSentUtc.AddOrUpdate(
            senderEmail,
            addValueFactory:=Function(k) mailSentTimeUtc,
            updateValueFactory:=Function(k, existing) If(mailSentTimeUtc > existing, mailSentTimeUtc, existing))
    End Sub


    ' ═══════════════════════════════════════════════════════════════════════════
    '  ATTACHMENT HANDLING (save + improved collect)
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Saves attachments to a per-mail temporary directory.</summary>
    Private Function SaveAttachmentsToTemp(mi As MailItem, tempDir As String) As List(Of AutoPilotAttachmentInfo)
        Dim result As New List(Of AutoPilotAttachmentInfo)()
        For i As Integer = 1 To mi.Attachments.Count
            Dim att = mi.Attachments(i)
            Try
                ' ── Embedded mail items (.msg) ──
                ' Outlook skips these with olEmbeddeditem; we save via SaveAsFile
                ' and then recursively unpack the contained message.
                If att.Type = OlAttachmentType.olEmbeddeditem Then
                    Try
                        Dim embeddedFileName As String = att.FileName
                        If String.IsNullOrWhiteSpace(embeddedFileName) Then embeddedFileName = $"embedded_{i}.msg"
                        If Not embeddedFileName.EndsWith(".msg", StringComparison.OrdinalIgnoreCase) Then
                            embeddedFileName = Path.GetFileNameWithoutExtension(embeddedFileName) & ".msg"
                        End If

                        Dim savePath = Path.Combine(tempDir, embeddedFileName)
                        Dim counter = 1
                        While File.Exists(savePath)
                            savePath = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(embeddedFileName) & $"_{counter}.msg")
                            counter += 1
                        End While
                        att.SaveAsFile(savePath)

                        ' Unpack the .msg recursively into the temp directory
                        Dim unpackedInfos = UnpackEmbeddedMailFile(savePath, tempDir, Path.GetFileNameWithoutExtension(savePath), 0)
                        result.AddRange(unpackedInfos)
                    Catch ex As System.Exception
                        result.Add(New AutoPilotAttachmentInfo() With {
                            .OriginalFileName = If(att.FileName, $"embedded_{i}"),
                            .StatusMessage = $"Error unpacking embedded mail: {ex.Message}"
                        })
                    End Try
                    Continue For
                End If

                Dim fileName As String = att.FileName
                If String.IsNullOrWhiteSpace(fileName) Then Continue For
                Dim size As Long = att.Size
                Dim info As New AutoPilotAttachmentInfo() With {
                    .OriginalFileName = fileName,
                    .Extension = Path.GetExtension(fileName).ToLowerInvariant(),
                    .SizeBytes = size,
                    .IsOverSizeLimit = (size > _apConfig.MaxAttachmentBytes)
                }
                ' Capture original file dates from MAPI properties (preserved from the original attachment,
                ' unlike file system timestamps which reset to current time on SaveAsFile)
                Try
                    Dim pa = att.PropertyAccessor
                    Try
                        info.CreatedTime = CDate(pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x30070040"))
                    Catch : End Try
                    Try
                        info.LastModifiedTime = CDate(pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x30080040"))
                    Catch : End Try
                Catch
                    ' Dates are optional; ignore failures
                End Try
                If info.IsOverSizeLimit Then
                    info.TempFilePath = Nothing
                    info.StatusMessage = $"Skipped (size {size / 1024 / 1024:F1} MB exceeds limit)"
                Else
                    Dim savePath = Path.Combine(tempDir, fileName)
                    Dim counter = 1
                    While File.Exists(savePath)
                        savePath = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(fileName) & $"_{counter}" & Path.GetExtension(fileName))
                        counter += 1
                    End While
                    att.SaveAsFile(savePath)
                    info.TempFilePath = savePath
                    info.StatusMessage = "Saved"

                    ' ── Extract PDF page metadata (page count, orientation, page size) ──
                    If info.Extension = ".pdf" Then
                        Try
                            Using pdfDoc = UglyToad.PdfPig.PdfDocument.Open(savePath)
                                info.PageCount = pdfDoc.NumberOfPages
                                If pdfDoc.NumberOfPages > 0 Then
                                    Dim firstPage = pdfDoc.GetPage(1)
                                    Dim w = firstPage.Width
                                    Dim h = firstPage.Height
                                    info.PageOrientation = If(w > h, "landscape", "portrait")

                                    ' Classify common page sizes (tolerance ±3 pt)
                                    Dim pw = Math.Min(w, h) ' always compare as portrait
                                    Dim ph = Math.Max(w, h)
                                    If Math.Abs(pw - 595) < 3 AndAlso Math.Abs(ph - 842) < 3 Then
                                        info.PageSize = "A4"
                                    ElseIf Math.Abs(pw - 612) < 3 AndAlso Math.Abs(ph - 792) < 3 Then
                                        info.PageSize = "Letter"
                                    ElseIf Math.Abs(pw - 612) < 3 AndAlso Math.Abs(ph - 1008) < 3 Then
                                        info.PageSize = "Legal"
                                    ElseIf Math.Abs(pw - 420) < 3 AndAlso Math.Abs(ph - 595) < 3 Then
                                        info.PageSize = "A5"
                                    ElseIf Math.Abs(pw - 842) < 3 AndAlso Math.Abs(ph - 1191) < 3 Then
                                        info.PageSize = "A3"
                                    Else
                                        info.PageSize = $"{w:F0} × {h:F0} pt"
                                    End If
                                End If
                            End Using
                        Catch
                            ' Non-critical — leave defaults (0 / Nothing)
                        End Try
                    End If
                End If

                ' If the saved file is a .msg, .eml, or .zip, recursively unpack it
                If Not info.IsOverSizeLimit AndAlso info.TempFilePath IsNot Nothing Then
                    Dim ext = info.Extension
                    If ext = ".msg" OrElse ext = ".eml" Then
                        Try
                            Dim unpackedInfos = UnpackEmbeddedMailFile(info.TempFilePath, tempDir,
                                Path.GetFileNameWithoutExtension(info.OriginalFileName), 0)
                            result.AddRange(unpackedInfos)
                            ' Skip adding the raw .msg/.eml itself — the unpacked text + attachments replace it
                            Continue For
                        Catch ex As System.Exception
                            ' Unpacking failed — fall through and add the raw file as-is
                            info.StatusMessage = $"Saved (unpacking failed: {ex.Message})"
                        End Try
                    ElseIf ext = ".zip" Then
                        Try
                            Dim unpackedInfos = UnpackZipFile(info.TempFilePath, tempDir,
                                Path.GetFileNameWithoutExtension(info.OriginalFileName), 0)
                            result.AddRange(unpackedInfos)
                            ' Skip adding the raw .zip itself — the unpacked files replace it
                            Continue For
                        Catch ex As System.Exception
                            info.StatusMessage = $"Saved (zip unpacking failed: {ex.Message})"
                        End Try
                    End If
                End If

                result.Add(info)
            Catch ex As System.Exception
                result.Add(New AutoPilotAttachmentInfo() With {.OriginalFileName = att.FileName, .StatusMessage = $"Error: {ex.Message}"})
            End Try
        Next
        Return result
    End Function


    ''' <summary>
    ''' Unpacks a .zip file into the temp directory. Each extracted file becomes a separate
    ''' <see cref="AutoPilotAttachmentInfo"/> entry. Nested .zip/.msg/.eml files are unpacked
    ''' recursively up to <see cref="AP_MaxArchiveDepth"/>. 
    ''' 
    ''' Security:
    '''   - Zip-slip prevention: entry paths are validated to stay inside tempDir.
    '''   - Zip bomb prevention: limits on entry count and total uncompressed size.
    '''   - Per-file size limit: individual entries exceeding MaxAttachmentBytes are skipped.
    '''   - Directory entries and hidden/system files are skipped.
    ''' </summary>
    ''' <param name="zipFilePath">Path to the .zip file.</param>
    ''' <param name="tempDir">The per-mail temp directory for all outputs.</param>
    ''' <param name="baseName">Base name prefix derived from the zip filename.</param>
    ''' <param name="depth">Current archive nesting depth (0 = top level).</param>
    ''' <returns>List of AutoPilotAttachmentInfo for each extracted file.</returns>
    Private Function UnpackZipFile(
            zipFilePath As String,
            tempDir As String,
            baseName As String,
            depth As Integer) As List(Of AutoPilotAttachmentInfo)

        Dim results As New List(Of AutoPilotAttachmentInfo)()

        If depth > AP_MaxArchiveDepth Then
            results.Add(New AutoPilotAttachmentInfo() With {
                .OriginalFileName = Path.GetFileName(zipFilePath),
                .TempFilePath = zipFilePath,
                .Extension = ".zip",
                .SizeBytes = If(File.Exists(zipFilePath), New FileInfo(zipFilePath).Length, 0),
                .StatusMessage = $"Max archive nesting depth ({AP_MaxArchiveDepth}) reached — not unpacked"
            })
            Return results
        End If

        Try
            Using archive = ZipFile.OpenRead(zipFilePath)
                ' ── Zip bomb check: entry count ──
                If archive.Entries.Count > AP_MaxZipEntries Then
                    ApDashboardLog($"  ⚠ Zip has {archive.Entries.Count} entries (limit {AP_MaxZipEntries}) — not unpacked: {Path.GetFileName(zipFilePath)}", "warn")
                    results.Add(New AutoPilotAttachmentInfo() With {
                        .OriginalFileName = Path.GetFileName(zipFilePath),
                        .TempFilePath = zipFilePath,
                        .Extension = ".zip",
                        .SizeBytes = New FileInfo(zipFilePath).Length,
                        .StatusMessage = $"Skipped (zip contains {archive.Entries.Count} entries, limit is {AP_MaxZipEntries})"
                    })
                    Return results
                End If

                ' ── Zip bomb check: total uncompressed size ──
                Dim totalUncompressed As Long = 0
                For Each entry In archive.Entries
                    totalUncompressed += entry.Length
                Next
                If totalUncompressed > AP_MaxZipTotalBytes Then
                    ApDashboardLog($"  ⚠ Zip uncompressed size {totalUncompressed / 1024 / 1024:F1} MB exceeds limit — not unpacked: {Path.GetFileName(zipFilePath)}", "warn")
                    results.Add(New AutoPilotAttachmentInfo() With {
                        .OriginalFileName = Path.GetFileName(zipFilePath),
                        .TempFilePath = zipFilePath,
                        .Extension = ".zip",
                        .SizeBytes = New FileInfo(zipFilePath).Length,
                        .StatusMessage = $"Skipped (uncompressed size {totalUncompressed / 1024 / 1024:F1} MB exceeds limit)"
                    })
                    Return results
                End If

                ApDashboardLog($"  📦 Unpacking zip: {Path.GetFileName(zipFilePath)} ({archive.Entries.Count} entries)", "info")

                Dim tempDirFull = Path.GetFullPath(tempDir)

                For Each entry In archive.Entries
                    Try
                        ' Skip directory entries
                        If String.IsNullOrWhiteSpace(entry.Name) Then Continue For

                        ' Use only the filename (strip any directory structure from the entry)
                        Dim entryFileName = Path.GetFileName(entry.FullName)
                        If String.IsNullOrWhiteSpace(entryFileName) Then Continue For

                        ' Skip hidden/system files (e.g., __MACOSX, .DS_Store, Thumbs.db)
                        If entryFileName.StartsWith(".", StringComparison.Ordinal) OrElse
                           entryFileName.Equals("Thumbs.db", StringComparison.OrdinalIgnoreCase) OrElse
                           entry.FullName.Contains("__MACOSX") Then
                            Continue For
                        End If

                        ' Build safe output path and validate against zip-slip
                        Dim outputPath = Path.Combine(tempDir, entryFileName)
                        Dim outputPathFull = Path.GetFullPath(outputPath)
                        If Not outputPathFull.StartsWith(tempDirFull, StringComparison.OrdinalIgnoreCase) Then
                            ' Zip-slip attempt — skip this entry
                            Continue For
                        End If

                        ' Prevent filename collision
                        Dim counter = 1
                        While File.Exists(outputPath)
                            outputPath = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(entryFileName) &
                                $"_{counter}" & Path.GetExtension(entryFileName))
                            counter += 1
                        End While

                        Dim entryExt = Path.GetExtension(entryFileName).ToLowerInvariant()
                        Dim entrySize = entry.Length

                        ' Per-file size check
                        If entrySize > _apConfig.MaxAttachmentBytes Then
                            results.Add(New AutoPilotAttachmentInfo() With {
                                .OriginalFileName = entryFileName,
                                .Extension = entryExt,
                                .SizeBytes = entrySize,
                                .IsOverSizeLimit = True,
                                .TempFilePath = Nothing,
                                .StatusMessage = $"Skipped (from zip, size {entrySize / 1024 / 1024:F1} MB exceeds limit)"
                            })
                            Continue For
                        End If

                        ' Extract the entry
                        entry.ExtractToFile(outputPath, True)

                        Dim info As New AutoPilotAttachmentInfo() With {
                            .OriginalFileName = Path.GetFileName(outputPath),
                            .TempFilePath = outputPath,
                            .Extension = entryExt,
                            .SizeBytes = New FileInfo(outputPath).Length,
                            .IsOverSizeLimit = False,
                            .StatusMessage = $"Extracted from zip: {Path.GetFileName(zipFilePath)}"
                        }

                        ' Recursively unpack nested containers
                        If entryExt = ".msg" OrElse entryExt = ".eml" Then
                            Try
                                Dim subResults = UnpackEmbeddedMailFile(outputPath, tempDir,
                                    Path.GetFileNameWithoutExtension(entryFileName), depth + 1)
                                results.AddRange(subResults)
                                Continue For
                            Catch
                                ' Fall through and add the raw file
                            End Try
                        ElseIf entryExt = ".zip" Then
                            Try
                                Dim subResults = UnpackZipFile(outputPath, tempDir,
                                    Path.GetFileNameWithoutExtension(entryFileName), depth + 1)
                                results.AddRange(subResults)
                                Continue For
                            Catch
                                ' Fall through and add the raw file
                            End Try
                        End If

                        results.Add(info)
                    Catch ex As System.Exception
                        results.Add(New AutoPilotAttachmentInfo() With {
                            .OriginalFileName = Path.GetFileName(entry.FullName),
                            .StatusMessage = $"Error extracting from zip: {ex.Message}"
                        })
                    End Try
                Next
            End Using

            ApDashboardLog($"  ✓ Zip unpacked: {Path.GetFileName(zipFilePath)} → {results.Count} file(s)", "info")

        Catch ex As System.Exception
            ApDashboardLog($"  ⚠ Failed to unpack zip: {ex.Message}", "warn")
            results.Add(New AutoPilotAttachmentInfo() With {
                .OriginalFileName = Path.GetFileName(zipFilePath),
                .TempFilePath = zipFilePath,
                .Extension = ".zip",
                .SizeBytes = If(File.Exists(zipFilePath), New FileInfo(zipFilePath).Length, 0),
                .StatusMessage = $"Saved (zip unpack failed: {ex.Message})"
            })
        End Try

        Return results
    End Function

    ''' <summary>
    ''' Recursively unpacks a .msg or .eml file into the temp directory.
    ''' Produces:
    '''   (1) A .txt file containing the mail metadata (From, To, Subject, Date) and body text.
    '''   (2) AutoPilotAttachmentInfo entries for each attachment found inside the embedded mail.
    '''   (3) Nested .msg/.eml attachments are unpacked recursively up to AP_MaxEmbeddedMailDepth.
    ''' 
    ''' Security: All output files are constrained to tempDir (path prefix check).
    ''' </summary>
    ''' <param name="mailFilePath">Path to the .msg or .eml file.</param>
    ''' <param name="tempDir">The per-mail temp directory for all outputs.</param>
    ''' <param name="baseName">Base name prefix for generated files (e.g. "forwarded_mail").</param>
    ''' <param name="depth">Current recursion depth (0 = top level).</param>
    ''' <returns>List of AutoPilotAttachmentInfo for the text conversion and any extracted attachments.</returns>
    Private Function UnpackEmbeddedMailFile(
            mailFilePath As String,
            tempDir As String,
            baseName As String,
            depth As Integer) As List(Of AutoPilotAttachmentInfo)

        Dim results As New List(Of AutoPilotAttachmentInfo)()
        If depth > AP_MaxEmbeddedMailDepth Then
            results.Add(New AutoPilotAttachmentInfo() With {
                .OriginalFileName = Path.GetFileName(mailFilePath),
                .TempFilePath = mailFilePath,
                .Extension = Path.GetExtension(mailFilePath).ToLowerInvariant(),
                .SizeBytes = If(File.Exists(mailFilePath), New FileInfo(mailFilePath).Length, 0),
                .StatusMessage = $"Max nesting depth ({AP_MaxEmbeddedMailDepth}) reached — not unpacked"
            })
            Return results
        End If

        Dim ext = Path.GetExtension(mailFilePath).ToLowerInvariant()

        If ext = ".msg" Then
            Return UnpackMsgFile(mailFilePath, tempDir, baseName, depth)
        ElseIf ext = ".eml" Then
            Return UnpackEmlFile(mailFilePath, tempDir, baseName, depth)
        Else
            ' Unknown extension — return as-is
            results.Add(New AutoPilotAttachmentInfo() With {
                .OriginalFileName = Path.GetFileName(mailFilePath),
                .TempFilePath = mailFilePath,
                .Extension = ext,
                .SizeBytes = If(File.Exists(mailFilePath), New FileInfo(mailFilePath).Length, 0),
                .StatusMessage = "Saved"
            })
            Return results
        End If
    End Function

    ''' <summary>
    ''' Unpacks a .msg file using Outlook's OpenSharedItem API.
    ''' Extracts the mail body as a .txt file and recursively processes nested attachments.
    ''' Must be called on the UI thread (COM access required).
    ''' </summary>
    Private Function UnpackMsgFile(
            msgFilePath As String,
            tempDir As String,
            baseName As String,
            depth As Integer) As List(Of AutoPilotAttachmentInfo)

        Dim results As New List(Of AutoPilotAttachmentInfo)()
        Dim msgItem As MailItem = Nothing

        Try
            ' Open the .msg file via Outlook
            Dim ns = Application.GetNamespace("MAPI")
            Dim sharedItem = ns.OpenSharedItem(msgFilePath)
            msgItem = TryCast(sharedItem, MailItem)

            If msgItem Is Nothing Then
                ' Not a MailItem (could be a meeting request, etc.) — return the raw file
                results.Add(New AutoPilotAttachmentInfo() With {
                    .OriginalFileName = Path.GetFileName(msgFilePath),
                    .TempFilePath = msgFilePath,
                    .Extension = ".msg",
                    .SizeBytes = New FileInfo(msgFilePath).Length,
                    .StatusMessage = "Saved (not a standard mail message)"
                })
                If sharedItem IsNot Nothing Then Try : Marshal.ReleaseComObject(sharedItem) : Catch : End Try
                Return results
            End If

            ' Extract metadata + body as text
            Dim textContent = BuildEmbeddedMailText(msgItem)
            Dim textFileName = baseName & "_content.txt"
            Dim textFilePath = Path.Combine(tempDir, textFileName)
            Dim counter = 1
            While File.Exists(textFilePath)
                textFilePath = Path.Combine(tempDir, baseName & $"_content_{counter}.txt")
                textFileName = Path.GetFileName(textFilePath)
                counter += 1
            End While
            File.WriteAllText(textFilePath, textContent, Encoding.UTF8)

            results.Add(New AutoPilotAttachmentInfo() With {
                .OriginalFileName = textFileName,
                .TempFilePath = textFilePath,
                .Extension = ".txt",
                .SizeBytes = New FileInfo(textFilePath).Length,
                .IsOverSizeLimit = False,
                .StatusMessage = $"Extracted from embedded mail: {Path.GetFileName(msgFilePath)}"
            })

            ApDashboardLog($"  📧 Unpacked embedded mail: {Path.GetFileName(msgFilePath)} → {textFileName}" &
                           If(msgItem.Attachments.Count > 0, $" + {msgItem.Attachments.Count} attachment(s)", ""), "info")

            ' Extract nested attachments
            For j As Integer = 1 To msgItem.Attachments.Count
                Dim nestedAtt = msgItem.Attachments(j)
                Try
                    If nestedAtt.Type = OlAttachmentType.olEmbeddeditem Then
                        ' Nested embedded mail — save as .msg and recurse
                        Try
                            Dim nestedName = If(nestedAtt.FileName, $"nested_{depth}_{j}.msg")
                            If Not nestedName.EndsWith(".msg", StringComparison.OrdinalIgnoreCase) Then
                                nestedName = Path.GetFileNameWithoutExtension(nestedName) & ".msg"
                            End If
                            Dim nestedPath = Path.Combine(tempDir, nestedName)
                            Dim nc = 1
                            While File.Exists(nestedPath)
                                nestedPath = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(nestedName) & $"_{nc}.msg")
                                nc += 1
                            End While
                            nestedAtt.SaveAsFile(nestedPath)
                            results.AddRange(UnpackEmbeddedMailFile(nestedPath, tempDir,
                                Path.GetFileNameWithoutExtension(nestedPath), depth + 1))
                        Catch ex As System.Exception
                            results.Add(New AutoPilotAttachmentInfo() With {
                                .OriginalFileName = If(nestedAtt.FileName, $"nested_{depth}_{j}"),
                                .StatusMessage = $"Error unpacking nested embedded mail: {ex.Message}"
                            })
                        End Try
                    Else
                        ' Regular attachment inside the embedded mail
                        Dim nestedFileName = nestedAtt.FileName
                        If String.IsNullOrWhiteSpace(nestedFileName) Then Continue For
                        Dim nestedSize As Long = nestedAtt.Size
                        Dim nestedInfo As New AutoPilotAttachmentInfo() With {
                            .OriginalFileName = nestedFileName,
                            .Extension = Path.GetExtension(nestedFileName).ToLowerInvariant(),
                            .SizeBytes = nestedSize,
                            .IsOverSizeLimit = (nestedSize > _apConfig.MaxAttachmentBytes)
                        }
                        If nestedInfo.IsOverSizeLimit Then
                            nestedInfo.TempFilePath = Nothing
                            nestedInfo.StatusMessage = $"Skipped (from embedded mail, size {nestedSize / 1024 / 1024:F1} MB exceeds limit)"
                        Else
                            Dim nestedSavePath = Path.Combine(tempDir, nestedFileName)
                            Dim nc = 1
                            While File.Exists(nestedSavePath)
                                nestedSavePath = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(nestedFileName) & $"_{nc}" & Path.GetExtension(nestedFileName))
                                nc += 1
                            End While
                            nestedAtt.SaveAsFile(nestedSavePath)
                            nestedInfo.TempFilePath = nestedSavePath
                            nestedInfo.StatusMessage = $"Saved (from embedded mail: {Path.GetFileName(msgFilePath)})"

                            ' If this nested attachment is itself a .msg, .eml, or .zip, recurse
                            If nestedInfo.Extension = ".msg" OrElse nestedInfo.Extension = ".eml" Then
                                Try
                                    Dim subResults = UnpackEmbeddedMailFile(nestedSavePath, tempDir,
                                        Path.GetFileNameWithoutExtension(nestedFileName), depth + 1)
                                    results.AddRange(subResults)
                                    Continue For ' Don't add the raw file — the unpacked results replace it
                                Catch
                                    ' Fall through and add the raw file
                                End Try
                            ElseIf nestedInfo.Extension = ".zip" Then
                                Try
                                    Dim subResults = UnpackZipFile(nestedSavePath, tempDir,
                                        Path.GetFileNameWithoutExtension(nestedFileName), depth + 1)
                                    results.AddRange(subResults)
                                    Continue For
                                Catch
                                End Try
                            End If
                        End If
                        results.Add(nestedInfo)
                    End If
                Catch ex As System.Exception
                    results.Add(New AutoPilotAttachmentInfo() With {
                        .OriginalFileName = If(nestedAtt.FileName, $"attachment_{j}"),
                        .StatusMessage = $"Error: {ex.Message}"
                    })
                End Try
            Next

        Catch ex As System.Exception
            ' OpenSharedItem failed — return the raw .msg as a fallback
            ApDashboardLog($"  ⚠ Failed to unpack .msg: {ex.Message}", "warn")
            results.Add(New AutoPilotAttachmentInfo() With {
                .OriginalFileName = Path.GetFileName(msgFilePath),
                .TempFilePath = msgFilePath,
                .Extension = ".msg",
                .SizeBytes = If(File.Exists(msgFilePath), New FileInfo(msgFilePath).Length, 0),
                .StatusMessage = $"Saved (unpack failed: {ex.Message})"
            })
        Finally
            If msgItem IsNot Nothing Then Try : Marshal.ReleaseComObject(msgItem) : Catch : End Try
        End Try

        Return results
    End Function

    ''' <summary>
    ''' Unpacks a .eml file by parsing it as plain text (RFC 2822 format).
    ''' Extracts headers and body as a .txt file. Embedded attachments in .eml
    ''' are extracted via simple MIME boundary parsing when possible.
    ''' </summary>
    Private Function UnpackEmlFile(
            emlFilePath As String,
            tempDir As String,
            baseName As String,
            depth As Integer) As List(Of AutoPilotAttachmentInfo)

        Dim results As New List(Of AutoPilotAttachmentInfo)()

        Try
            ' First try: open via Outlook's OpenSharedItem (supports .eml on most Outlook versions)
            Try
                Dim ns = Application.GetNamespace("MAPI")
                Dim sharedItem = ns.OpenSharedItem(emlFilePath)
                Dim emlMail = TryCast(sharedItem, MailItem)

                If emlMail IsNot Nothing Then
                    ' Successfully opened as MailItem — use the same logic as .msg
                    Dim textContent = BuildEmbeddedMailText(emlMail)
                    Dim textFileName = baseName & "_content.txt"
                    Dim textFilePath = Path.Combine(tempDir, textFileName)
                    Dim counter = 1
                    While File.Exists(textFilePath)
                        textFilePath = Path.Combine(tempDir, baseName & $"_content_{counter}.txt")
                        textFileName = Path.GetFileName(textFilePath)
                        counter += 1
                    End While
                    File.WriteAllText(textFilePath, textContent, Encoding.UTF8)

                    results.Add(New AutoPilotAttachmentInfo() With {
                        .OriginalFileName = textFileName,
                        .TempFilePath = textFilePath,
                        .Extension = ".txt",
                        .SizeBytes = New FileInfo(textFilePath).Length,
                        .IsOverSizeLimit = False,
                        .StatusMessage = $"Extracted from embedded mail: {Path.GetFileName(emlFilePath)}"
                    })

                    ApDashboardLog($"  📧 Unpacked .eml via Outlook: {Path.GetFileName(emlFilePath)} → {textFileName}" &
                                   If(emlMail.Attachments.Count > 0, $" + {emlMail.Attachments.Count} attachment(s)", ""), "info")

                    ' Extract nested attachments (same logic as .msg)
                    For j As Integer = 1 To emlMail.Attachments.Count
                        Dim nestedAtt = emlMail.Attachments(j)
                        Try
                            If nestedAtt.Type = OlAttachmentType.olEmbeddeditem Then
                                Try
                                    Dim nestedName = If(nestedAtt.FileName, $"nested_{depth}_{j}.msg")
                                    If Not nestedName.EndsWith(".msg", StringComparison.OrdinalIgnoreCase) Then
                                        nestedName = Path.GetFileNameWithoutExtension(nestedName) & ".msg"
                                    End If
                                    Dim nestedPath = Path.Combine(tempDir, nestedName)
                                    Dim nc = 1
                                    While File.Exists(nestedPath)
                                        nestedPath = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(nestedName) & $"_{nc}.msg")
                                        nc += 1
                                    End While
                                    nestedAtt.SaveAsFile(nestedPath)
                                    results.AddRange(UnpackEmbeddedMailFile(nestedPath, tempDir,
                                        Path.GetFileNameWithoutExtension(nestedPath), depth + 1))
                                Catch ex As System.Exception
                                    results.Add(New AutoPilotAttachmentInfo() With {
                                        .OriginalFileName = If(nestedAtt.FileName, $"nested_{depth}_{j}"),
                                        .StatusMessage = $"Error unpacking nested embedded mail: {ex.Message}"
                                    })
                                End Try
                            Else
                                Dim nestedFileName = nestedAtt.FileName
                                If String.IsNullOrWhiteSpace(nestedFileName) Then Continue For
                                Dim nestedSize As Long = nestedAtt.Size
                                Dim nestedInfo As New AutoPilotAttachmentInfo() With {
                                    .OriginalFileName = nestedFileName,
                                    .Extension = Path.GetExtension(nestedFileName).ToLowerInvariant(),
                                    .SizeBytes = nestedSize,
                                    .IsOverSizeLimit = (nestedSize > _apConfig.MaxAttachmentBytes)
                                }
                                If Not nestedInfo.IsOverSizeLimit Then
                                    Dim nestedSavePath = Path.Combine(tempDir, nestedFileName)
                                    Dim nc = 1
                                    While File.Exists(nestedSavePath)
                                        nestedSavePath = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(nestedFileName) & $"_{nc}" & Path.GetExtension(nestedFileName))
                                        nc += 1
                                    End While
                                    nestedAtt.SaveAsFile(nestedSavePath)
                                    nestedInfo.TempFilePath = nestedSavePath
                                    nestedInfo.StatusMessage = $"Saved (from embedded .eml: {Path.GetFileName(emlFilePath)})"

                                    If nestedInfo.Extension = ".msg" OrElse nestedInfo.Extension = ".eml" Then
                                        Try
                                            Dim subResults = UnpackEmbeddedMailFile(nestedSavePath, tempDir,
                                                Path.GetFileNameWithoutExtension(nestedFileName), depth + 1)
                                            results.AddRange(subResults)
                                            Continue For
                                        Catch
                                        End Try
                                    ElseIf nestedInfo.Extension = ".zip" Then
                                        Try
                                            Dim subResults = UnpackZipFile(nestedSavePath, tempDir,
                                                Path.GetFileNameWithoutExtension(nestedFileName), depth + 1)
                                            results.AddRange(subResults)
                                            Continue For
                                        Catch
                                        End Try
                                    End If
                                Else
                                    nestedInfo.TempFilePath = Nothing
                                    nestedInfo.StatusMessage = $"Skipped (from embedded .eml, size exceeds limit)"
                                End If
                                results.Add(nestedInfo)
                            End If
                        Catch ex As System.Exception
                            results.Add(New AutoPilotAttachmentInfo() With {
                                .OriginalFileName = If(nestedAtt.FileName, $"attachment_{j}"),
                                .StatusMessage = $"Error: {ex.Message}"
                            })
                        End Try
                    Next

                    Try : Marshal.ReleaseComObject(emlMail) : Catch : End Try
                    Return results
                End If

                If sharedItem IsNot Nothing Then Try : Marshal.ReleaseComObject(sharedItem) : Catch : End Try
            Catch
                ' OpenSharedItem may not support .eml on all Outlook versions — fall through to text parsing
            End Try

            ' Fallback: parse .eml as plain text (RFC 2822)
            ApDashboardLog($"  📧 Parsing .eml as text: {Path.GetFileName(emlFilePath)}", "step")
            Dim emlContent = File.ReadAllText(emlFilePath, Encoding.UTF8)
            Dim textResult = ParseEmlAsText(emlContent, Path.GetFileName(emlFilePath))

            Dim fallbackTextFileName = baseName & "_content.txt"
            Dim fallbackTextFilePath = Path.Combine(tempDir, fallbackTextFileName)
            Dim fallbackCounter = 1
            While File.Exists(fallbackTextFilePath)
                fallbackTextFilePath = Path.Combine(tempDir, baseName & $"_content_{fallbackCounter}.txt")
                fallbackTextFileName = Path.GetFileName(fallbackTextFilePath)
                fallbackCounter += 1
            End While
            File.WriteAllText(fallbackTextFilePath, textResult, Encoding.UTF8)

            results.Add(New AutoPilotAttachmentInfo() With {
                .OriginalFileName = fallbackTextFileName,
                .TempFilePath = fallbackTextFilePath,
                .Extension = ".txt",
                .SizeBytes = New FileInfo(fallbackTextFilePath).Length,
                .IsOverSizeLimit = False,
                .StatusMessage = $"Extracted from .eml (text parse): {Path.GetFileName(emlFilePath)}"
            })

        Catch ex As System.Exception
            ApDashboardLog($"  ⚠ Failed to unpack .eml: {ex.Message}", "warn")
            results.Add(New AutoPilotAttachmentInfo() With {
                .OriginalFileName = Path.GetFileName(emlFilePath),
                .TempFilePath = emlFilePath,
                .Extension = ".eml",
                .SizeBytes = If(File.Exists(emlFilePath), New FileInfo(emlFilePath).Length, 0),
                .StatusMessage = $"Saved (unpack failed: {ex.Message})"
            })
        End Try

        Return results
    End Function

    ''' <summary>
    ''' Builds a structured text representation of an embedded mail item,
    ''' including headers (From, To, CC, Subject, Date) and the plain-text body.
    ''' </summary>
    Private Shared Function BuildEmbeddedMailText(mi As MailItem) As String
        Dim sb As New StringBuilder()
        sb.AppendLine("═══════════════════════════════════════════════════")
        sb.AppendLine("EMBEDDED EMAIL MESSAGE")
        sb.AppendLine("═══════════════════════════════════════════════════")
        sb.AppendLine()

        Try : sb.AppendLine($"From: {If(mi.SenderName, "")} <{If(mi.SenderEmailAddress, "")}>") : Catch : End Try
        Try : sb.AppendLine($"To: {If(CStr(mi.To), "")}") : Catch : End Try

        Try : If Not String.IsNullOrWhiteSpace(CStr(mi.CC)) Then sb.AppendLine($"CC: {mi.CC}")
        Catch : End Try

        Try : sb.AppendLine($"Subject: {If(mi.Subject, "")}") : Catch : End Try
        Try : sb.AppendLine($"Date: {mi.ReceivedTime:yyyy-MM-dd HH:mm:ss}") : Catch : End Try

        Try
            If mi.Attachments.Count > 0 Then
                sb.AppendLine($"Attachments: {mi.Attachments.Count}")
                For j As Integer = 1 To mi.Attachments.Count
                    Try
                        sb.AppendLine($"  - {mi.Attachments(j).FileName} ({mi.Attachments(j).Size / 1024:F0} KB)")
                    Catch
                    End Try
                Next
            End If
        Catch : End Try

        sb.AppendLine()
        sb.AppendLine("───────────────────────────────────────────────────")
        sb.AppendLine()

        Try
            Dim body = If(mi.Body, "")
            If body.Length > 50000 Then body = body.Substring(0, 50000) & vbCrLf & "[... body truncated at 50,000 characters ...]"
            sb.Append(body)
        Catch : End Try

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Simple text-based parser for .eml (RFC 2822) files.
    ''' Extracts common headers and the text body portion.
    ''' Does not decode MIME attachments (those require Outlook's OpenSharedItem).
    ''' </summary>
    Private Shared Function ParseEmlAsText(emlContent As String, sourceFileName As String) As String
        Dim sb As New StringBuilder()
        sb.AppendLine("═══════════════════════════════════════════════════")
        sb.AppendLine($"EMBEDDED EMAIL MESSAGE (parsed from {sourceFileName})")
        sb.AppendLine("═══════════════════════════════════════════════════")
        sb.AppendLine()

        ' Split into headers and body at the first blank line
        Dim headerEnd = emlContent.IndexOf(vbCrLf & vbCrLf)
        If headerEnd < 0 Then headerEnd = emlContent.IndexOf(vbLf & vbLf)

        Dim headerSection As String
        Dim bodySection As String

        If headerEnd >= 0 Then
            headerSection = emlContent.Substring(0, headerEnd)
            bodySection = emlContent.Substring(headerEnd).TrimStart({CChar(vbCr), CChar(vbLf)})
        Else
            headerSection = emlContent
            bodySection = ""
        End If

        ' Unfold continuation lines (RFC 2822: lines starting with whitespace are continuations)
        headerSection = Regex.Replace(headerSection, "(\r?\n)[ \t]+", " ")

        ' Extract key headers
        Dim headerLines = headerSection.Split({vbCrLf, vbLf}, StringSplitOptions.None)
        Dim importantHeaders = {"From:", "To:", "Cc:", "Subject:", "Date:", "Content-Type:"}
        For Each line In headerLines
            For Each hdr In importantHeaders
                If line.StartsWith(hdr, StringComparison.OrdinalIgnoreCase) Then
                    sb.AppendLine(line.Trim())
                    Exit For
                End If
            Next
        Next

        sb.AppendLine()
        sb.AppendLine("───────────────────────────────────────────────────")
        sb.AppendLine()

        ' For the body, strip MIME boundaries and encoding artifacts as best we can
        ' Remove common MIME boundary lines
        Dim cleanBody = Regex.Replace(bodySection, "^--[a-zA-Z0-9_\-\.=]+\r?\n", "", RegexOptions.Multiline)
        ' Remove Content-Type / Content-Transfer-Encoding header blocks within body parts
        cleanBody = Regex.Replace(cleanBody, "^Content-[A-Za-z\-]+:.*\r?\n", "", RegexOptions.Multiline)

        ' Strip HTML tags if the body appears to be HTML
        If cleanBody.IndexOf("<html", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
           cleanBody.IndexOf("<body", StringComparison.OrdinalIgnoreCase) >= 0 Then
            Try
                Dim htmlDoc As New HtmlAgilityPack.HtmlDocument()
                htmlDoc.LoadHtml(cleanBody)
                cleanBody = htmlDoc.DocumentNode.InnerText
            Catch
                ' Fallback: crude tag stripping
                cleanBody = Regex.Replace(cleanBody, "<[^>]+>", "")
            End Try
        End If

        If cleanBody.Length > 50000 Then cleanBody = cleanBody.Substring(0, 50000) & vbCrLf & "[... body truncated at 50,000 characters ...]"
        sb.Append(cleanBody.Trim())

        Return sb.ToString()
    End Function



    ''' <summary>
    ''' Collects output files produced by tools. Uses the OutputFiles lists on each
    ''' attachment info AND falls back to scanning for new files in tempDir.
    ''' Only files within tempDir are eligible (security: path prefix check).
    ''' </summary>
    Private Function CollectResultAttachments(tempDir As String, originalAttachments As List(Of AutoPilotAttachmentInfo)) As List(Of String)
        Dim results As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim tempDirFull = Path.GetFullPath(tempDir)

        ' Files already surfaced in a previous turn (or pre-existing user uploads)
        ' MUST NOT be returned again. The set is populated by
        ' ResetChatAgentDeliverableTrackingForNewTurn() at the start of each turn.
        Dim alreadySurfaced As HashSet(Of String) = _chatAgentSurfacedFiles
        If alreadySurfaced Is Nothing Then
            alreadySurfaced = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        End If

        ' 1. Collect from OutputFiles (registered by tools like process_word_document, merge_pdfs).
        '    OutputFiles for the current turn are accumulated as tools run; prior-turn
        '    entries were cleared by ResetChatAgentDeliverableTrackingForNewTurn().
        If originalAttachments IsNot Nothing Then
            For Each att In originalAttachments
                If att.OutputFiles IsNot Nothing Then
                    For Each outputPath In att.OutputFiles
                        If Not String.IsNullOrEmpty(outputPath) AndAlso File.Exists(outputPath) Then
                            Dim fullOut As String = Path.GetFullPath(outputPath)
                            ' Security: only include files inside the per-turn temp dir.
                            If Not fullOut.StartsWith(tempDirFull, StringComparison.OrdinalIgnoreCase) Then Continue For
                            ' Bleed protection: skip files already surfaced in a previous turn.
                            If alreadySurfaced.Contains(fullOut) Then Continue For
                            results.Add(fullOut)
                        End If
                    Next
                End If
            Next
        End If

        ' 2. Fallback: scan for files in tempDir that are (a) not in the original
        '    upload set, AND (b) not in the already-surfaced set from prior turns.
        If Directory.Exists(tempDir) Then
            Dim originalPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            If originalAttachments IsNot Nothing Then
                For Each att In originalAttachments
                    If att.TempFilePath IsNot Nothing Then originalPaths.Add(Path.GetFullPath(att.TempFilePath))
                Next
            End If
            For Each filePath In Directory.GetFiles(tempDir, "*.*", SearchOption.AllDirectories)
                Dim full As String = Path.GetFullPath(filePath)
                If originalPaths.Contains(full) Then Continue For
                If alreadySurfaced.Contains(full) Then Continue For
                results.Add(full)
            Next
        End If

        ' Promote everything returned this turn into the "already surfaced" set so
        ' the next turn does not pick it up again — even if the OutputFiles tracking
        ' on uploads is not perfectly cleared by the next reset.
        For Each surfaced In results
            alreadySurfaced.Add(surfaced)
        Next
        _chatAgentSurfacedFiles = alreadySurfaced

        Return results.ToList()
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  PROMPT BUILDING
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Builds the user prompt from the mail content and attachments.</summary>
    Private Shared Function BuildUserPromptFromMail(info As AutoPilotMailInfo, attachments As List(Of AutoPilotAttachmentInfo)) As String
        Dim sb As New StringBuilder()
        sb.AppendLine("[INCOMING EMAIL]")
        sb.AppendLine($"From: {info.SenderName} <{info.SenderEmail}>")
        sb.AppendLine($"Subject: {info.Subject}")
        sb.AppendLine($"Received: {info.ReceivedTime:yyyy-MM-dd HH:mm}")
        sb.AppendLine()
        sb.AppendLine("[EMAIL BODY]")
        Dim body = info.Body
        If body.Length > 25000 Then body = body.Substring(0, 25000) & vbCrLf & "[... truncated ...]"
        sb.AppendLine(body)
        sb.AppendLine("[/EMAIL BODY]")
        If attachments IsNot Nothing AndAlso attachments.Count > 0 Then
            sb.AppendLine()
            sb.AppendLine("[ATTACHMENTS]")
            For i As Integer = 0 To attachments.Count - 1
                Dim att = attachments(i)
                Dim sizeStr = If(att.SizeBytes > 0, $" ({att.SizeBytes / 1024:F0} KB)", "")
                Dim statusStr = If(att.IsOverSizeLimit,
                    $" [OVER SIZE LIMIT — max {AP_DefaultMaxAttachmentBytes / 1024 / 1024:F0} MB, cannot process]", "")
                Dim dateStr As String = ""
                If att.LastModifiedTime.HasValue Then
                    dateStr &= $", modified: {att.LastModifiedTime.Value:yyyy-MM-dd HH:mm:ss} UTC"
                End If
                If att.CreatedTime.HasValue Then
                    dateStr &= $", created: {att.CreatedTime.Value:yyyy-MM-dd HH:mm:ss} UTC"
                End If
                Dim pdfStr As String = ""
                If att.PageCount > 0 Then
                    pdfStr = $", {att.PageCount} page(s)"
                    If Not String.IsNullOrWhiteSpace(att.PageOrientation) Then
                        pdfStr &= $", {att.PageOrientation}"
                    End If
                    If Not String.IsNullOrWhiteSpace(att.PageSize) Then
                        pdfStr &= $", {att.PageSize}"
                    End If
                End If
                sb.AppendLine($"  {i + 1}. {att.OriginalFileName}{sizeStr}{pdfStr}{dateStr}{statusStr}")
            Next
            sb.AppendLine("[/ATTACHMENTS]")
            sb.AppendLine()
            sb.AppendLine("You have tools available to process these attachments if the sender requests it.")
        End If
        Return sb.ToString()
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  REPLY SENDING (sourcesHtml parameter)
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Sends a reply to the original sender, optionally with attachments and sources.</summary>
    Private Sub SendReplyToSender(originalMail As MailItem, responseText As String,
                                  resultAttachments As List(Of String),
                                  Optional tagAsAutoReply As Boolean = True,
                                  Optional sourcesHtml As String = "",
                                  Optional isHoldingOnly As Boolean = False)
        Dim reply As MailItem = Nothing
        Try
            reply = originalMail.Reply()
            reply.CC = ""
            reply.BCC = ""

            ' SECURITY: Explicitly set reply.To to the envelope sender address.
            ' This prevents Reply-To header manipulation where an attacker sets
            ' Reply-To: victim@other.com to redirect AutoPilot responses.
            ' We also clear ReplyRecipients to ensure Outlook doesn't silently
            ' re-add a Reply-To address when the mail is sent.
            Dim senderAddr As String = If(originalMail.SenderEmailAddress, "")
            If String.IsNullOrWhiteSpace(senderAddr) Then
                senderAddr = If(originalMail.SenderName, "")
            End If

            ' Detect and log Reply-To mismatch for operator awareness
            Try
                Dim replyToAddr As String = ""
                If originalMail.ReplyRecipients IsNot Nothing AndAlso originalMail.ReplyRecipients.Count > 0 Then
                    replyToAddr = originalMail.ReplyRecipients(1).Address
                End If
                If Not String.IsNullOrWhiteSpace(replyToAddr) AndAlso
                   Not replyToAddr.Equals(senderAddr, StringComparison.OrdinalIgnoreCase) Then
                    ApDashboardLog($"⚠ SECURITY: Reply-To mismatch detected! From={senderAddr}, Reply-To={replyToAddr}. Replying to From address only.", "warn")
                End If
            Catch
                ' ReplyRecipients may not be available on all mail types
            End Try

            reply.To = senderAddr

            ' Clear any Reply-To recipients that Outlook's .Reply() may have populated
            Try
                While reply.ReplyRecipients.Count > 0
                    reply.ReplyRecipients.Remove(1)
                End While
            Catch
            End Try

            reply.BodyFormat = OlBodyFormat.olFormatHTML

            Dim originalThread As String = If(reply.HTMLBody, "")

            ' Build HTML body: response + sources used footer + branded footer
            Dim htmlBody = ConvertResponseToHtml(responseText)

            ' Append "Sources used:" section
            If Not String.IsNullOrWhiteSpace(sourcesHtml) Then
                htmlBody &= sourcesHtml
            End If

            htmlBody &= BuildAutoPilotFooter()
            reply.HTMLBody = htmlBody & originalThread

            ' Add result attachments
            If resultAttachments IsNot Nothing Then
                For Each attachPath In resultAttachments
                    If File.Exists(attachPath) Then
                        reply.Attachments.Add(attachPath, OlAttachmentType.olByValue, , Path.GetFileName(attachPath))
                    End If
                Next
            End If

            If tagAsAutoReply Then
                Try
                    ' Use "holding" value for holding-only notices so the catch-up scan
                    ' can distinguish them from substantive replies (value "true").
                    Dim headerValue = If(isHoldingOnly, AP_LoopHeaderValueHolding, AP_LoopHeaderValue)
                    reply.PropertyAccessor.SetProperty(AP_LoopHeaderProperty, headerValue)
                Catch : End Try
            End If
            Try : reply.Categories = AP_CategoryName : Catch : End Try

            Dim cleanupGroupId As String = Nothing
            Dim cleanupIsEligible As Boolean = False
            Dim cleanupAnsweredUtc As DateTime? = Nothing
            Dim cleanupDeleteAfterUtc As DateTime? = Nothing

            Try
                cleanupGroupId = GetOrCreateCleanupGroupId(originalMail)
                If Not String.IsNullOrWhiteSpace(cleanupGroupId) Then
                    cleanupDeleteAfterUtc = If(isHoldingOnly, CType(Nothing, DateTime?), GetAutoDeleteCutoffUtc())
                    cleanupAnsweredUtc = If(isHoldingOnly, CType(Nothing, DateTime?), DateTime.UtcNow)
                    cleanupIsEligible = Not isHoldingOnly AndAlso cleanupDeleteAfterUtc.HasValue

                    StampCleanupMetadata(
                        reply,
                        cleanupGroupId,
                        isEligible:=cleanupIsEligible,
                        answeredUtc:=cleanupAnsweredUtc,
                        deleteAfterUtc:=cleanupDeleteAfterUtc,
                        saveItem:=False)
                End If
            Catch ex As System.Exception
                Debug.WriteLine($"[AutoPilot] Failed to stamp cleanup metadata on reply: {ex.Message}")
            End Try

            reply.Send()
            Try : MoveLastSentToInkyReplies(cleanupGroupId, cleanupIsEligible, cleanupAnsweredUtc, cleanupDeleteAfterUtc) : Catch : End Try

        Catch ex As System.Exception
            ApDashboardLog($"ERROR sending reply: {ex.Message}", "error")
        Finally
            If reply IsNot Nothing Then Try : Marshal.ReleaseComObject(reply) : Catch : End Try
        End Try
    End Sub

    ''' <summary>Builds the HTML footer appended to AutoPilot replies.</summary>
    Private Function BuildAutoPilotFooter() As String
        Dim logoBase64 As String = ""
        Try
            Dim logoBytes = GetLogoPngBytes()
            If logoBytes IsNot Nothing AndAlso logoBytes.Length > 0 Then logoBase64 = System.Convert.ToBase64String(logoBytes)
        Catch
        End Try
        Dim sb As New StringBuilder()
        sb.AppendLine("<br/><hr style='border:none;border-top:1px solid #cccccc;margin:20px 0 10px 0;'/>")
        sb.AppendLine("<table cellpadding='0' cellspacing='0' border='0' style='font-family:Arial,sans-serif;font-size:9pt;color:#888888;'>")
        sb.AppendLine("<tr>")
        If logoBase64.Length > 0 Then
            sb.AppendLine($"<td style='padding-right:8px;vertical-align:middle;'><img src='data:image/png;base64,{logoBase64}' width='24' height='24' alt='{AN}' style='display:block;'/></td>")
        End If
        sb.AppendLine($"<td style='vertical-align:middle;'>")
        Dim customFooter = _apConfig?.FooterText
        If Not String.IsNullOrWhiteSpace(customFooter) Then
            sb.AppendLine($"<span style='font-size:9pt;color:#666666;'>{System.Net.WebUtility.HtmlEncode(customFooter)}</span><br/>")
        End If
        sb.AppendLine($"<span style='font-size:8pt;color:#aaaaaa;'>Powered by <a href='https://redink.ai' style='color:#aaaaaa;text-decoration:none;'>{AN}</a></span>")
        sb.AppendLine("</td></tr></table>")
        Return sb.ToString()
    End Function

    ''' <summary>Converts a Markdown response to HTML for Outlook.</summary>
    Private Function ConvertResponseToHtml(responseMarkdown As String) As String
        Dim builder As New Markdig.MarkdownPipelineBuilder()
        builder.UsePipeTables().UseGridTables().UseSoftlineBreakAsHardlineBreak()
        builder.UseListExtras().UseFootnotes().UseDefinitionLists()
        builder.UseAbbreviations().UseAutoLinks().UseTaskLists()
        builder.UseMathematics().UseFigures()
        builder.UseAdvancedExtensions().UseGenericAttributes()
        Dim pipeline = builder.Build()
        Dim htmlBody As String = Markdig.Markdown.ToHtml(responseMarkdown, pipeline)
        Return "<div style='font-family:Arial,sans-serif;font-size:11pt;'>" & htmlBody & "</div>"
    End Function

    ''' <summary>Moves the most recent AutoPilot reply into the Inky Replies folder.</summary>
    Private Sub MoveLastSentToInkyReplies(Optional groupId As String = Nothing,
                                          Optional isEligible As Boolean = False,
                                          Optional answeredUtc As DateTime? = Nothing,
                                          Optional deleteAfterUtc As DateTime? = Nothing)
        Try
            Dim ns = Application.GetNamespace("MAPI")
            Dim sentFolder = ns.GetDefaultFolder(OlDefaultFolders.olFolderSentMail)
            Dim inkyFolder As MAPIFolder = Nothing
            Dim items As Items = Nothing

            Try
                Try
                    inkyFolder = sentFolder.Folders(AP_SentSubfolder)
                Catch
                    inkyFolder = sentFolder.Folders.Add(AP_SentSubfolder, OlDefaultFolders.olFolderSentMail)
                End Try

                items = sentFolder.Items
                items.Sort("[ReceivedTime]", True)

                For Each item As Object In items
                    Dim mi2 As MailItem = Nothing
                    Dim movedObj As Object = Nothing
                    Dim movedMail As MailItem = Nothing

                    Try
                        mi2 = TryCast(item, MailItem)
                        If mi2 Is Nothing Then Continue For
                        If mi2.Categories Is Nothing OrElse Not mi2.Categories.Contains(AP_CategoryName) Then Continue For

                        ' When a cleanup group id is supplied, bind the move/stamp to the
                        ' exact reply we just sent for this group instead of blindly taking
                        ' the most recent categorized item. This prevents an older AutoPilot
                        ' reply from being picked (and the just-sent reply from being left in
                        ' Sent Items without persisted cleanup metadata, so it is never deleted).
                        If Not String.IsNullOrWhiteSpace(groupId) AndAlso
                           Not String.Equals(GetCleanupGroupId(mi2), groupId, StringComparison.OrdinalIgnoreCase) Then
                            Continue For
                        End If

                        movedObj = mi2.Move(inkyFolder)
                        movedMail = TryCast(movedObj, MailItem)

                        If movedMail IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(groupId) Then
                            StampCleanupMetadata(
                                movedMail,
                                groupId,
                                isEligible:=isEligible,
                                answeredUtc:=answeredUtc,
                                deleteAfterUtc:=deleteAfterUtc,
                                saveItem:=True)
                        End If

                        Exit For
                    Finally
                        If movedMail IsNot Nothing Then Try : Marshal.ReleaseComObject(movedMail) : Catch : End Try
                        If movedObj IsNot Nothing AndAlso Not ReferenceEquals(movedObj, movedMail) Then
                            Try : Marshal.ReleaseComObject(movedObj) : Catch : End Try
                        End If
                        If mi2 IsNot Nothing Then Try : Marshal.ReleaseComObject(mi2) : Catch : End Try
                        If item IsNot Nothing AndAlso Not ReferenceEquals(item, mi2) Then
                            Try : Marshal.ReleaseComObject(item) : Catch : End Try
                        End If
                    End Try
                Next
            Finally
                If items IsNot Nothing Then Try : Marshal.ReleaseComObject(items) : Catch : End Try
                If inkyFolder IsNot Nothing Then Try : Marshal.ReleaseComObject(inkyFolder) : Catch : End Try
                If sentFolder IsNot Nothing Then Try : Marshal.ReleaseComObject(sentFolder) : Catch : End Try
                If ns IsNot Nothing Then Try : Marshal.ReleaseComObject(ns) : Catch : End Try
            End Try
        Catch ex As System.Exception
            Debug.WriteLine($"MoveLastSentToInkyReplies error: {ex.Message}")
        End Try
    End Sub

    ''' <summary>Shows the approval dialog and returns the user's choice.</summary>
    Private Function ShowApprovalDialog(info As AutoPilotMailInfo, draftResponse As String, resultAttachments As List(Of String)) As Boolean
        Dim attachmentInfo As String = ""
        If resultAttachments IsNot Nothing AndAlso resultAttachments.Count > 0 Then
            attachmentInfo = vbCrLf & vbCrLf & $"Attachments to include: {resultAttachments.Count}" & vbCrLf &
                String.Join(vbCrLf, resultAttachments.Select(Function(p) "  • " & Path.GetFileName(p)))
        End If
        Dim displayText As String =
            $"From: {info.SenderName} <{info.SenderEmail}>" & vbCrLf &
            $"Subject: {info.Subject}" & vbCrLf & vbCrLf &
            "── Draft Reply ──" & vbCrLf & draftResponse & attachmentInfo
        Return ShowCustomYesNoBox(displayText, "Send Reply", "Discard", header:=$"{AN6} AutoPilot — Approve Reply") = 1
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  DASHBOARD LOG — date+time, no duplicate timestamp
    ' ═══════════════════════════════════════════════════════════════════════════


    ''' <summary>
    ''' Appends a log entry to the AutoPilot dashboard.
    ''' Does NOT prepend a timestamp here because LogWindow.AppendLogInternal
    ''' already adds [HH:mm:ss]. Instead we prepend the date-only portion so the
    ''' dashboard shows [HH:mm:ss] [dd-MMM] message.
    ''' </summary>
    Private Sub ApDashboardLog(message As String, level As String)
        Debug.WriteLine($"[AutoPilot] [{level}] {message}")

        Try
            If _apDashboard IsNot Nothing Then
                Dim dateTag = DateTime.Now.ToString("dd-MMM", Globalization.CultureInfo.InvariantCulture)
                Dim taggedMessage = $"[{dateTag}] {message}"

                If _apDashboard.InvokeRequired Then
                    _apDashboard.BeginInvoke(New MethodInvoker(Sub() _apDashboard.AppendLog(taggedMessage, level)))
                Else
                    _apDashboard.AppendLog(taggedMessage, level)
                End If
            End If
        Catch
        End Try

        ' Intentionally do not mirror AutoPilot dashboard entries into the Local Chat
        ' tooling log. Local Chat already receives direct context.Log(...) messages,
        ' and mirroring here creates duplicate entries such as web_grounding progress.
    End Sub

    ''' <summary>Marks the dashboard operation as complete.</summary>
    Private Sub ApDashboardMarkComplete()
        Try
            If _apDashboard IsNot Nothing Then
                If _apDashboard.InvokeRequired Then
                    _apDashboard.BeginInvoke(New MethodInvoker(Sub() _apDashboard.MarkComplete()))
                Else
                    _apDashboard.MarkComplete()
                End If
            End If
        Catch
        End Try
    End Sub


    ''' <summary>
    ''' Records the current UTC time as the last successfully processed timestamp.
    ''' Uses wall-clock time (not the mail's ReceivedTime) to avoid losing mails
    ''' received at the exact same second as the last processed mail.
    ''' </summary>
    Private Sub RecordLastProcessedTime()
        Try
            My.Settings.AP_LastProcessedTime = DateTime.UtcNow
            My.Settings.Save()
        Catch ex As System.Exception
            Debug.WriteLine($"[AutoPilot] Failed to save last processed time: {ex.Message}")
        End Try
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  SOURCES USED HTML FOOTER
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Returns True only for tool calls that represent real information sources
    ''' for the "Sources used:" footer. Excludes helper/meta/classical tools.
    ''' </summary>
    Private Shared Function IsSourceEntryForFooter(entry As AutoPilotToolCallEntry) As Boolean
        If entry Is Nothing OrElse Not entry.WasSuccessful Then Return False

        Dim toolName As String = If(entry.ToolName, "").Trim()
        If toolName = "" Then Return False

        ' Never show loader/helper/meta tools.
        If toolName.Equals(SharedLibrary.Agents.ToolLoaderTool.LoaderToolName, StringComparison.OrdinalIgnoreCase) Then
            Return False
        End If

        ' Real built-in source tools.
        If toolName.Equals(InternalWebToolName, StringComparison.OrdinalIgnoreCase) OrElse
           toolName.Equals(InternalDownloadWebFilesToolName, StringComparison.OrdinalIgnoreCase) OrElse
           toolName.Equals(InternalSearchToolName, StringComparison.OrdinalIgnoreCase) OrElse
           toolName.StartsWith(InternalKnowledgeToolNamePrefix, StringComparison.OrdinalIgnoreCase) OrElse
           SharedLibrary.Agents.WebGroundingTool.IsWebGroundingTool(toolName) OrElse
           SharedLibrary.SharedLibrary.M365ToolService.IsM365ToolName(toolName) Then
            Return True
        End If

        ' Never show agent-layer or other classical/helper tools.
        If toolName.StartsWith("skill_", StringComparison.OrdinalIgnoreCase) OrElse
           toolName.StartsWith("agent_", StringComparison.OrdinalIgnoreCase) OrElse
           toolName.Equals(SharedLibrary.Agents.SkillInvokeTool.ToolName, StringComparison.OrdinalIgnoreCase) OrElse
           SharedLibrary.Agents.MemoryTools.IsMemoryTool(toolName) OrElse
           SharedLibrary.Agents.TextTools.IsTextTool(toolName) OrElse
           SharedLibrary.Agents.WorkspaceTools.IsWorkspaceTool(toolName) OrElse
           SharedLibrary.Agents.WordTools.IsWordTool(toolName) OrElse
           SharedLibrary.Agents.WordDocTools.IsWordDocTool(toolName) OrElse
           SharedLibrary.Agents.JsRunTool.IsJsTool(toolName) Then
            Return False
        End If

        ' External tools are treated as real sources by default
        ' (e.g. Special Services tools), unless excluded above.
        Return Not entry.IsInternalTool
    End Function

    ''' <summary>
    ''' Builds an HTML "Sources used:" section from real source calls only.
    ''' Includes source tools such as M365, Special Services, internet search,
    ''' web grounding, web retrieval, and knowledge-store lookups.
    ''' Excludes helper/meta/classical tools such as tool_loader, skills, agents,
    ''' memory/text/workspace/js tools, and other non-source internals.
    ''' </summary>
    Private Shared Function BuildSourcesUsedHtml(toolCallLog As List(Of AutoPilotToolCallEntry)) As String
        If toolCallLog Is Nothing OrElse toolCallLog.Count = 0 Then Return ""

        Dim sourceCalls = toolCallLog.Where(Function(e) IsSourceEntryForFooter(e)).ToList()
        If sourceCalls.Count = 0 Then Return ""

        Dim sb As New StringBuilder()
        sb.AppendLine("<div style='font-size:8pt;color:#999999;margin-top:16px;border-top:1px solid #eeeeee;padding-top:8px;'>")
        sb.AppendLine("<b style='color:#888888;'>Sources used:</b><br/>")

        For Each entry In sourceCalls
            Dim toolLabel = If(Not String.IsNullOrEmpty(entry.ToolDisplayName), entry.ToolDisplayName, entry.ToolName)
            Dim icon = If(entry.WasSuccessful, "✓", "✗")

            If entry.Urls IsNot Nothing AndAlso entry.Urls.Count > 0 Then
                For Each url In entry.Urls
                    Dim encodedUrl = System.Net.WebUtility.HtmlEncode(url)
                    sb.AppendLine($"&bull; {icon} <a href='{encodedUrl}' style='color:#4A90D9;text-decoration:underline;'>{encodedUrl}</a><br/>")
                Next
            Else
                Dim paramInfo = ""
                If Not String.IsNullOrEmpty(entry.ParamSummary) Then
                    paramInfo = $" — {System.Net.WebUtility.HtmlEncode(entry.ParamSummary)}"
                End If

                sb.AppendLine($"&bull; {icon} {System.Net.WebUtility.HtmlEncode(toolLabel)}{paramInfo}<br/>")
            End If
        Next

        sb.AppendLine("</div>")
        Return sb.ToString()
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL CALL LOGGING HELPER
    '  Called from ExecuteToolCall in ThisAddin.Tooling.vb and from
    '  TryExecuteAutoPilotTool in ThisAddIn.AutoPilot.Tools.vb
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Records a tool call for dashboard logging and the "Sources used:" footer.</summary>
    Friend Sub RecordAutoPilotToolCall(toolName As String, toolDisplayName As String,
                                       paramSummary As String, isInternalTool As Boolean,
                                       wasSuccessful As Boolean, resultExcerpt As String,
                                       elapsed As TimeSpan,
                                       Optional urls As List(Of String) = Nothing)
        If _apCurrentToolCallLog Is Nothing Then Return

        _apCurrentToolCallLog.Add(New AutoPilotToolCallEntry() With {
            .ToolName = toolName,
            .ToolDisplayName = toolDisplayName,
            .ParamSummary = paramSummary,
            .IsInternalTool = isInternalTool,
            .WasSuccessful = wasSuccessful,
            .ResultExcerpt = resultExcerpt,
            .Elapsed = elapsed,
            .Urls = urls
        })
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  Conversation Detection
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Checks whether the incoming mail is part of an existing AutoPilot conversation,
    ''' even if it is a second (or subsequent) user reply to a previous AutoPilot response.
    ''' This fixes the issue where a user replies to an already-responded-to mail and the
    ''' AutoPilot does not detect it as part of an active conversation.
    ''' </summary>
    ''' <param name="mi">The incoming MailItem to check.</param>
    ''' <returns>True if this mail is part of an AutoPilot-managed conversation.</returns>
    Private Function IsPartOfAutoPilotConversation(mi As MailItem) As Boolean
        Try
            ' 1. Check the incoming mail itself for the header (any value: "true" or "holding")
            If HasAutoReplyHeader(mi) Then Return True

            ' 2. Walk the conversation thread to find any mail with the AutoPilot header.
            '    This handles the case where the user replies to a previous AutoPilot reply:
            '    the user's new mail won't have the header, but the previous reply does.
            Dim conversation As Outlook.Conversation = Nothing
            Try
                conversation = mi.GetConversation()
            Catch
                ' GetConversation can fail on some store types
            End Try

            If conversation IsNot Nothing Then
                Try
                    Dim rootItems As SimpleItems = conversation.GetRootItems()
                    If rootItems IsNot Nothing Then
                        For Each rootItem As Object In rootItems
                            If CheckConversationNodeForAutoReply(conversation, rootItem, 0) Then
                                Return True
                            End If
                        Next
                    End If
                Catch
                    ' Conversation API can throw on certain store types
                End Try
            End If

            ' 3. Fallback: check the ConversationID against recently processed conversations.
            '    This handles edge cases where the Conversation API doesn't work but we
            '    have already processed a mail in this thread during this session.
            If Not String.IsNullOrWhiteSpace(mi.ConversationID) Then
                If _apProcessedConversations.ContainsKey(mi.ConversationID) Then
                    Return True
                End If
            End If

        Catch ex As System.Exception
            Debug.WriteLine($"[AutoPilot] IsPartOfAutoPilotConversation error: {ex.Message}")
        End Try

        Return False
    End Function

    ''' <summary>Recursively checks conversation tree nodes for the AutoPilot reply header (any value).</summary>
    Private Function CheckConversationNodeForAutoReply(
            conv As Outlook.Conversation,
            item As Object,
            depth As Integer) As Boolean

        If depth > AP_MaxThreadDepth Then Return False

        Try
            Dim nodeMail = TryCast(item, MailItem)
            If nodeMail IsNot Nothing Then
                If HasAutoReplyHeader(nodeMail) Then Return True
            End If

            ' Check child nodes
            Dim children As SimpleItems = conv.GetChildren(item)
            If children IsNot Nothing Then
                For Each child As Object In children
                    If CheckConversationNodeForAutoReply(conv, child, depth + 1) Then
                        Return True
                    End If
                Next
            End If
        Catch
            ' Ignore individual node errors
        End Try

        Return False
    End Function

    ''' <summary>Checks whether a mail item has the X-RedInk-AutoReply MAPI property set (any value: "true" or "holding").</summary>
    Private Function HasAutoReplyHeader(mi As MailItem) As Boolean
        Try
            Dim prop As Object = Nothing
            Try
                Dim pa As PropertyAccessor = mi.PropertyAccessor
                prop = pa.GetProperty(AP_LoopHeaderProperty)
            Catch
                ' Property not set — this is normal for user-sent mails
                Return False
            End Try
            If prop IsNot Nothing Then
                Dim val = CStr(prop)
                If val.Equals(AP_LoopHeaderValue, StringComparison.OrdinalIgnoreCase) OrElse
                   val.Equals(AP_LoopHeaderValueHolding, StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            End If
        Catch
            ' Ignore errors
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Tracks conversation IDs that AutoPilot has already processed during this session.
    ''' Used as a fallback when the Conversation API is unavailable.
    ''' Key = ConversationID, Value = UTC timestamp of last processing.
    ''' </summary>
    Private ReadOnly _apProcessedConversations As New ConcurrentDictionary(Of String, DateTime)()

    ''' <summary>Records that a conversation has been processed by AutoPilot.</summary>
    Private Sub TrackAutoPilotConversation(mi As MailItem)
        Try
            If Not String.IsNullOrWhiteSpace(mi.ConversationID) Then
                _apProcessedConversations(mi.ConversationID) = DateTime.UtcNow
            End If
        Catch
            ' Ignore
        End Try
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  MODEL OVERRIDE COMMAND
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Scans the first few lines of the latest e-mail body for "#model: ModelName".
    ''' Only the latest/top-most portion is scanned (not quoted replies).
    ''' The command must appear in the first AP_ModelCommandScanLines non-blank lines.
    ''' </summary>
    Private Sub TryParseModelOverride(body As String,
                                      ByRef modelConfig As ModelConfig,
                                      ByRef modelName As String,
                                      ByRef strippedBody As String)
        modelConfig = Nothing
        modelName = Nothing
        strippedBody = body

        If String.IsNullOrWhiteSpace(body) Then Return

        ' Use the existing helper from ThisAddIn.Commands.vb which handles
        ' more separator patterns (From:, On ... wrote:, etc.)
        Dim latestBody = GetLatestMailBody(body)
        Dim lines = latestBody.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
        Dim nonBlankCount = 0
        Dim commandLineIndex = -1

        For i As Integer = 0 To lines.Length - 1
            Dim line = lines(i).Trim()
            If String.IsNullOrWhiteSpace(line) Then Continue For

            nonBlankCount += 1
            If nonBlankCount > AP_ModelCommandScanLines Then Exit For

            If line.StartsWith(AP_ModelCommandPrefix, StringComparison.OrdinalIgnoreCase) Then
                Dim requestedModel = line.Substring(AP_ModelCommandPrefix.Length).Trim()
                If String.IsNullOrWhiteSpace(requestedModel) Then Continue For

                ' Try to find a matching model from available alternates
                Try
                    ' Use INI_AlternateModelPath (same as the config dialog and model selector),
                    ' NOT INI_LocalModelPath (which is for speech/local models).
                    Dim iniPath = INI_AlternateModelPath
                    If String.IsNullOrWhiteSpace(iniPath) Then
                        ApDashboardLog($"Model override '{requestedModel}' ignored — no alternate model path configured.", "warn")
                        Continue For
                    End If

                    iniPath = ExpandEnvironmentVariables(iniPath)
                    If Not IO.File.Exists(iniPath) Then
                        ApDashboardLog($"Model override '{requestedModel}' ignored — alternate model file not found: {iniPath}", "warn")
                        Continue For
                    End If

                    Dim alternates = LoadAlternativeModels(iniPath, _context)
                    If alternates Is Nothing OrElse alternates.Count = 0 Then
                        ApDashboardLog($"Model override '{requestedModel}' ignored — no models loaded from INI.", "warn")
                        Continue For
                    End If

                    ApDashboardLog($"Resolving model override '{requestedModel}' against {alternates.Count} available model(s)...", "step")

                    ' Match by description (friendly name) first, then by model ID
                    Dim match = alternates.FirstOrDefault(
                        Function(m) m.ModelDescription IsNot Nothing AndAlso
                                    m.ModelDescription.IndexOf(requestedModel, StringComparison.OrdinalIgnoreCase) >= 0)
                    If match Is Nothing Then
                        match = alternates.FirstOrDefault(
                            Function(m) m.Model IsNot Nothing AndAlso
                                        m.Model.IndexOf(requestedModel, StringComparison.OrdinalIgnoreCase) >= 0)
                    End If

                    ' Fallback: keyword matching — all words in the request must appear
                    ' somewhere in ModelDescription (order-independent)
                    If match Is Nothing Then
                        Dim keywords = requestedModel.Split({" "c, ","c}, StringSplitOptions.RemoveEmptyEntries)
                        If keywords.Length > 1 Then
                            match = alternates.FirstOrDefault(
                                Function(m) m.ModelDescription IsNot Nothing AndAlso
                                            keywords.All(Function(kw) m.ModelDescription.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0))
                        End If
                        If match Is Nothing AndAlso keywords.Length > 1 Then
                            match = alternates.FirstOrDefault(
                                Function(m) m.Model IsNot Nothing AndAlso
                                            keywords.All(Function(kw) m.Model.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0))
                        End If
                    End If

                    If match IsNot Nothing Then
                        modelConfig = match
                        modelName = If(match.ModelDescription, match.Model)
                        commandLineIndex = i
                        ApDashboardLog($"✓ Model override resolved: '{requestedModel}' → {modelName}", "info")
                    Else
                        ApDashboardLog($"Model override '{requestedModel}' not found in {alternates.Count} available model(s).", "warn")
                        ' Log available model names to help the user
                        For Each alt In alternates
                            Dim altLabel = If(Not String.IsNullOrWhiteSpace(alt.ModelDescription), alt.ModelDescription, alt.Model)
                            ApDashboardLog($"  available: {altLabel}", "step")
                        Next
                    End If
                Catch ex As System.Exception
                    ApDashboardLog($"Error resolving model override: {ex.Message}", "warn")
                End Try

                Exit For ' Only process the first #model: found
            End If
        Next

        ' Strip the #model: line from the body
        If commandLineIndex >= 0 Then
            Dim sbStripped As New StringBuilder()
            For j As Integer = 0 To lines.Length - 1
                If j <> commandLineIndex Then sbStripped.AppendLine(lines(j))
            Next
            ' Reconstruct: stripped latest portion + any remaining quoted thread
            If latestBody.Length < body.Length Then
                strippedBody = sbStripped.ToString().TrimEnd() & body.Substring(latestBody.Length)
            Else
                strippedBody = sbStripped.ToString().TrimEnd()
            End If
        End If
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  VOICEMAIL PROCESSING
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Audio file extensions recognized as voicemail attachments.</summary>
    Private Shared ReadOnly AP_VoicemailAudioExtensions As String() = {
        ".wav", ".mp3", ".m4a", ".aac", ".wma", ".ogg", ".flac", ".amr"
    }

    ''' <summary>
    ''' In-memory caller ID → email map loaded from the user-provided CSV.
    ''' Key = normalized phone number (digits only), Value = (Email, DisplayName).
    ''' </summary>
    Private _apVoicemailCallerIdMap As Dictionary(Of String, (Email As String, DisplayName As String)) = Nothing

    ''' <summary>
    ''' Loads the caller ID → email map from a CSV file.
    ''' Format per line: phone,email[,displayname]
    ''' Lines starting with ; or # are comments. Empty lines are skipped.
    ''' Phone numbers are normalized (stripped of +, spaces, dashes, parentheses).
    ''' Multiple numbers mapping to the same email are supported naturally.
    ''' </summary>
    Private Sub LoadVoicemailCallerIdMap(csvPath As String)
        _apVoicemailCallerIdMap = New Dictionary(Of String, (Email As String, DisplayName As String))()
        Try
            If Not IO.File.Exists(csvPath) Then
                ApDashboardLog($"⚠ Voicemail caller ID map not found: {csvPath}", "warn")
                Return
            End If

            Dim lines = IO.File.ReadAllLines(csvPath, Text.Encoding.UTF8)
            Dim loaded As Integer = 0

            For Each line In lines
                Dim trimmed = line.Trim()
                If String.IsNullOrWhiteSpace(trimmed) Then Continue For
                If trimmed.StartsWith(";") OrElse trimmed.StartsWith("#") Then Continue For

                Dim parts = trimmed.Split(","c)
                If parts.Length < 2 Then Continue For

                Dim rawPhone = parts(0).Trim()
                Dim email = parts(1).Trim()
                Dim displayName = If(parts.Length >= 3, parts(2).Trim(), "")

                If String.IsNullOrWhiteSpace(rawPhone) OrElse String.IsNullOrWhiteSpace(email) Then Continue For

                Dim normalizedPhone = NormalizePhoneNumber(rawPhone)
                If normalizedPhone.Length >= 6 Then
                    _apVoicemailCallerIdMap(normalizedPhone) = (email, displayName)
                    loaded += 1
                End If
            Next

            ApDashboardLog($"Loaded {loaded} caller ID mapping(s) from: {IO.Path.GetFileName(csvPath)}", "info")

        Catch ex As System.Exception
            ApDashboardLog($"⚠ Error loading voicemail caller ID map: {ex.Message}", "warn")
        End Try
    End Sub

    ''' <summary>
    ''' Normalizes a phone number to digits only for consistent matching.
    ''' Strips +, spaces, dashes, parentheses, dots, and leading 00 (international prefix).
    ''' Examples:
    '''   +41 76 548 32 21  → 41765483221
    '''   0041765483221     → 41765483221
    '''   (076) 548-32-21   → 765483221
    '''   41765483221       → 41765483221
    ''' </summary>
    Private Shared Function NormalizePhoneNumber(phone As String) As String
        If String.IsNullOrWhiteSpace(phone) Then Return ""
        ' Strip all non-digit characters
        Dim digitsOnly = Regex.Replace(phone, "[^\d]", "")
        ' Strip leading 00 (international dialing prefix → country code)
        If digitsOnly.StartsWith("00") AndAlso digitsOnly.Length > 6 Then
            digitsOnly = digitsOnly.Substring(2)
        End If
        Return digitsOnly
    End Function

    ''' <summary>
    ''' Extracts a caller ID (phone number) from a voicemail filename and/or email subject/body.
    ''' Tries multiple patterns:
    '''   1. Filename pattern: From_41765483221_... → 41765483221
    '''   2. Subject pattern: "COMBOX pro 0860793221038" → 0860793221038
    '''   3. Flexible: any sequence of 8+ digits in the filename
    '''   4. Flexible: any sequence of 8+ digits in the subject
    ''' Returns the normalized phone number, or empty string if no caller ID found.
    ''' </summary>
    Private Shared Function ExtractCallerIdFromVoicemail(
            audioFileName As String,
            subject As String,
            body As String) As String

        ' 1. Filename pattern: From_{digits}_...
        If Not String.IsNullOrWhiteSpace(audioFileName) Then
            Dim fileMatch = Regex.Match(audioFileName, "From[_\-]?(\d{6,})", RegexOptions.IgnoreCase)
            If fileMatch.Success Then
                Return NormalizePhoneNumber(fileMatch.Groups(1).Value)
            End If
        End If

        ' 2. Subject/body pattern: "COMBOX pro {digits}" or similar PBX patterns
        For Each textSource In {subject, body}
            If String.IsNullOrWhiteSpace(textSource) Then Continue For
            ' "COMBOX pro 0860793221038" or "Voicemail from +41765483221"
            Dim pbxMatch = Regex.Match(textSource, "(?:COMBOX|voicemail|mailbox|anruf|appel|chiamata)\s+(?:pro|from|von|de|di)?\s*\+?(\d[\d\s\-\.]{5,})", RegexOptions.IgnoreCase)
            If pbxMatch.Success Then
                Return NormalizePhoneNumber(pbxMatch.Groups(1).Value)
            End If
        Next

        ' 3. Flexible: first sequence of 8+ digits in the filename
        If Not String.IsNullOrWhiteSpace(audioFileName) Then
            Dim digitMatch = Regex.Match(audioFileName, "(\d{8,})")
            If digitMatch.Success Then
                Return NormalizePhoneNumber(digitMatch.Groups(1).Value)
            End If
        End If

        ' 4. Flexible: first sequence of 8+ digits in the subject
        If Not String.IsNullOrWhiteSpace(subject) Then
            Dim digitMatch = Regex.Match(subject, "(\d{8,})")
            If digitMatch.Success Then
                Return NormalizePhoneNumber(digitMatch.Groups(1).Value)
            End If
        End If

        Return ""
    End Function

    ''' <summary>
    ''' Looks up a normalized caller ID in the caller ID map.
    ''' Tries exact match first, then suffix matching (last N digits) to handle
    ''' country code variations (e.g., 41765483221 vs 765483221).
    ''' </summary>
    Private Function LookupCallerIdEmail(normalizedCallerId As String) As (Email As String, DisplayName As String)?
        If _apVoicemailCallerIdMap Is Nothing OrElse String.IsNullOrWhiteSpace(normalizedCallerId) Then Return Nothing

        ' Exact match
        Dim result As (Email As String, DisplayName As String) = Nothing
        If _apVoicemailCallerIdMap.TryGetValue(normalizedCallerId, result) Then
            Return result
        End If

        ' Suffix match: try matching the last N digits (N = min of caller and map entry lengths)
        ' This handles cases where the voicemail system sends "765483221" but the map has "41765483221"
        For Each kvp In _apVoicemailCallerIdMap
            Dim mapPhone = kvp.Key
            Dim shorter = Math.Min(normalizedCallerId.Length, mapPhone.Length)
            If shorter >= 7 Then ' at least 7 digits must match
                If normalizedCallerId.Length >= shorter AndAlso mapPhone.Length >= shorter Then
                    If normalizedCallerId.Substring(normalizedCallerId.Length - shorter) =
                       mapPhone.Substring(mapPhone.Length - shorter) Then
                        Return kvp.Value
                    End If
                End If
            End If
        Next

        Return Nothing
    End Function

    ''' <summary>
    ''' Determines whether a mail item is a voicemail from the registered voicemail sender.
    ''' Returns True if:
    '''   1. Voicemail processing is enabled
    '''   2. The sender matches the configured voicemail sender address
    '''   3. The mail has at least one audio attachment
    ''' </summary>
    Private Function IsVoicemailFromRegisteredSender(mailInfo As AutoPilotMailInfo) As Boolean
        If _apConfig Is Nothing OrElse Not _apConfig.EnableVoicemailProcessing Then Return False
        If String.IsNullOrWhiteSpace(_apConfig.VoicemailSenderAddress) Then Return False
        If Not mailInfo.SenderEmail.Equals(_apConfig.VoicemailSenderAddress, StringComparison.OrdinalIgnoreCase) Then Return False
        ' Must have at least one audio attachment
        If mailInfo.AttachmentNames Is Nothing OrElse mailInfo.AttachmentNames.Count = 0 Then Return False
        For Each name In mailInfo.AttachmentNames
            Dim ext = IO.Path.GetExtension(name).ToLowerInvariant()
            For Each audioExt In AP_VoicemailAudioExtensions
                If ext = audioExt Then Return True
            Next
        Next
        Return False
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  HELPFUL FAILURE RESPONSE
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Generates a meaningful failure response when AutoPilot cannot fulfill a request.
    ''' Includes: (1) attachment size limit info if relevant, (2) a brief Red Ink add-in
    ''' suggestion based on the HelpMeInky manual, and (3) optionally an Internet-based
    ''' alternative tool suggestion via the InternetResearch model.
    ''' Falls back to a static message if all LLM calls fail.
    ''' </summary>
    Private Async Function GenerateHelpfulFailureResponseAsync(
            mailInfo As AutoPilotMailInfo,
            attachments As List(Of AutoPilotAttachmentInfo),
            failureReason As String,
            ct As CancellationToken) As Task(Of String)

        Dim sb As New StringBuilder()

        ' ── Attachment size limit detail ──
        Dim hasOversized = attachments IsNot Nothing AndAlso attachments.Any(Function(a) a.IsOverSizeLimit)
        If hasOversized Then
            Dim limitMB = _apConfig.MaxAttachmentBytes / 1024.0 / 1024.0
            Dim oversizedNames = attachments.Where(Function(a) a.IsOverSizeLimit).
                Select(Function(a) $"{a.OriginalFileName} ({a.SizeBytes / 1024.0 / 1024.0:F1} MB)").ToList()
            sb.AppendLine($"I'm sorry, but I was unable to process your request because one or more attachments exceed the maximum permitted size of {limitMB:F0} MB:")
            sb.AppendLine()
            For Each name In oversizedNames
                sb.AppendLine($"  - {name}")
            Next
            sb.AppendLine()
            sb.AppendLine($"Please split large documents into smaller parts and resend, or use the {AN} add-in locally to process the file(s) directly on your computer.")
            sb.AppendLine()
            sb.Append($"— {AN6}")
            Return sb.ToString()
        End If

        ' ── Static preamble ──
        sb.AppendLine($"I'm sorry, but I was unable to fulfill your request.")
        sb.AppendLine()

        ' ── Red Ink add-in suggestion (via HelpMeInky manual) ──
        Dim redInkSuggestion As String = Nothing
        Try
            redInkSuggestion = Await GetRedInkSuggestionAsync(mailInfo, failureReason, ct)
        Catch ex As System.Exception
            ApDashboardLog($"Fallback Red Ink suggestion failed: {ex.Message}", "warn")
        End Try

        If Not String.IsNullOrWhiteSpace(redInkSuggestion) Then
            sb.AppendLine(redInkSuggestion.Trim())
            sb.AppendLine()
        End If

        ' ── Internet alternative suggestion (optional, via InternetResearch model) ──
        Dim internetSuggestion As String = Nothing
        Try
            internetSuggestion = Await GetInternetAlternativeSuggestionAsync(mailInfo, failureReason, ct)
        Catch ex As System.Exception
            ApDashboardLog($"Fallback internet suggestion failed: {ex.Message}", "warn")
        End Try

        If Not String.IsNullOrWhiteSpace(internetSuggestion) Then
            sb.AppendLine(internetSuggestion.Trim())
            sb.AppendLine()
            ' Append disclaimer directly — this path sends text to the sender without LLM rewriting
            sb.AppendLine("_Please note: Third-party services and tools may only be used if permitted by your organization's policies. " &
                          "Before using any external service or tool, ensure it meets your corporate security, confidentiality, and data protection requirements._")
            sb.AppendLine()
        End If

        ' ── Closing ──
        If String.IsNullOrWhiteSpace(redInkSuggestion) AndAlso String.IsNullOrWhiteSpace(internetSuggestion) Then
            sb.AppendLine("Please try again, rephrase your request, or contact the operator for assistance.")
        End If

        sb.Append($"— {AN6}")
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Calls the LLM with the HelpMeInky manual to suggest how the user could
    ''' accomplish the task using the Red Ink add-ins. Returns a brief suggestion
    ''' or Nothing if no manual is available or the call fails.
    ''' Uses the dedicated HelpMe model (if configured in the alternate model INI)
    ''' exactly as HelpMeInky does, falling back to the base AutoPilot model only
    ''' if no HelpMe model is available.
    ''' </summary>
    Private Async Function GetRedInkSuggestionAsync(
            mailInfo As AutoPilotMailInfo,
            failureReason As String,
            ct As CancellationToken) As Task(Of String)

        ' Load the HelpMeInky manual (cached per session)
        If Not _apHelpMeManualCacheLoaded Then
            _apHelpMeManualCacheLoaded = True
            Try
                Dim manualPath = INI_HelpMeInkyPath
                If Not String.IsNullOrWhiteSpace(manualPath) Then
                    manualPath = ExpandEnvironmentVariables(manualPath)
                    _apHelpMeManualCache = Await LoadManualTextAsync(manualPath, ct)
                    ApDashboardLog($"Loaded HelpMe manual for fallback suggestions ({If(_apHelpMeManualCache IsNot Nothing, _apHelpMeManualCache.Length.ToString() & " chars", "not available")})", "step")
                End If
            Catch ex As System.Exception
                ApDashboardLog($"Failed to load HelpMe manual: {ex.Message}", "warn")
            End Try
        End If

        If String.IsNullOrWhiteSpace(_apHelpMeManualCache) Then Return Nothing

        ' Use the SP_HelpMe system prompt (same as HelpMeInky) with an AutoPilot-specific addendum
        Dim systemPrompt As String = _context.SP_HelpMe &
            $" (This is an automated query from {AN6} AutoPilot, not an interactive chat. " &
            $"A user sent an e-mail request that could not be fulfilled automatically. " &
            $"Based on the manual, suggest in 4-10 SHORT sentences what {AN} feature could help, " &
            $"how to use it, and tell the user to open 'Help me, {AN8}' inside the add-in for more guidance. " &
            $"If no {AN} feature applies, reply with exactly one word: NONE)"

        ' Build user prompt in the same structure as HelpMeInky.SendAsync
        Dim userPrompt As New StringBuilder()
        userPrompt.AppendLine("User question:")
        userPrompt.AppendLine($"I tried to use {AN6} AutoPilot to: {mailInfo.Subject}")
        userPrompt.AppendLine($"Details: {If(mailInfo.Body.Length > 500, mailInfo.Body.Substring(0, 500), mailInfo.Body)}")
        userPrompt.AppendLine($"But it failed because: {failureReason}")
        If mailInfo.AttachmentNames IsNot Nothing AndAlso mailInfo.AttachmentNames.Count > 0 Then
            userPrompt.AppendLine($"The e-mail had these attachments: {String.Join(", ", mailInfo.AttachmentNames)}")
        End If
        userPrompt.AppendLine($"What {AN} feature could help the user accomplish this task directly?")
        userPrompt.AppendLine()
        userPrompt.AppendLine("Manual:")
        userPrompt.AppendLine(_apHelpMeManualCache)

        ' Try to use the dedicated HelpMe model (same as HelpMeInky.CallHelpMeLlmAsync)
        Dim backupConfig = GetCurrentConfig(_context)
        Dim useHelpMeModel As Boolean = False
        Dim useSecondApi As Boolean = False
        Dim timeout As Long = 0

        Try
            If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                useHelpMeModel = GetSpecialTaskModel(_context, INI_AlternateModelPath, "HelpMe")
            End If
        Catch
        End Try

        If useHelpMeModel Then
            useSecondApi = True
            timeout = If(_context.INI_Timeout_2 > 0, _context.INI_Timeout_2, _context.INI_Timeout)
            ApDashboardLog("Red Ink suggestion using dedicated HelpMe model", "step")
        Else
            ' Fallback: use the base AutoPilot model
            ApplyModelConfig(_context, _apBaseModelConfig)
            useSecondApi = _apUseSecondApi
            timeout = _context.INI_Timeout
            ApDashboardLog("Red Ink suggestion using base AutoPilot model (no HelpMe model configured)", "step")
        End If

        Try
            Dim result = Await LLM(systemPrompt, userPrompt.ToString(),
                                   UseSecondAPI:=useSecondApi, HideSplash:=True,
                                   EnsureUI:=False, Timeout:=timeout,
                                   cancellationToken:=ct)
            result = If(result, "").Trim()
            If result.Equals("NONE", StringComparison.OrdinalIgnoreCase) OrElse result.Length < 5 Then
                ApDashboardLog($"Red Ink suggestion returned NONE or too short ({result.Length} chars)", "step")
                Return Nothing
            End If
            ApDashboardLog($"Red Ink suggestion generated ({result.Length} chars)", "step")
            Return result
        Catch ex As System.Exception
            ApDashboardLog($"Red Ink suggestion LLM call failed: {ex.GetType().Name}: {ex.Message}", "warn")
            Return Nothing
        Finally
            RestoreDefaults(_context, backupConfig)
        End Try
    End Function

    ''' <summary>
    ''' Loads manual text from an HTTP(S) URL or a local file path.
    ''' Mirrors the loading logic of HelpMeInky.GetManualTextFreshAsync, supporting
    ''' plain text, PDF (with magic-byte detection for remote content), DOCX, and RTF.
    ''' </summary>
    Private Async Function LoadManualTextAsync(pathOrUrl As String, ct As CancellationToken) As Task(Of String)
        If String.IsNullOrWhiteSpace(pathOrUrl) Then Return Nothing

        Dim s = pathOrUrl.Trim()

        ' ── Remote URL ──
        If s.StartsWith("http://", StringComparison.OrdinalIgnoreCase) OrElse
           s.StartsWith("https://", StringComparison.OrdinalIgnoreCase) Then
            Try
                ' Ensure modern TLS
                Try
                    System.Net.ServicePointManager.SecurityProtocol =
                        System.Net.SecurityProtocolType.Tls12 Or CType(&HC00, System.Net.SecurityProtocolType)
                Catch : End Try

                Dim handler As New System.Net.Http.HttpClientHandler()
                handler.AllowAutoRedirect = True
                handler.AutomaticDecompression = System.Net.DecompressionMethods.GZip Or System.Net.DecompressionMethods.Deflate

                Using client As New System.Net.Http.HttpClient(handler)
                    client.Timeout = TimeSpan.FromSeconds(30)
                    Try
                        client.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", "RedInk/1.0 (+https://redink.ai)")
                        client.DefaultRequestHeaders.TryAddWithoutValidation("Accept", "application/pdf, text/*, */*")
                    Catch : End Try

                    Using resp = Await client.GetAsync(s, System.Net.Http.HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(False)
                        If Not resp.IsSuccessStatusCode Then
                            ApDashboardLog($"Manual URL returned HTTP {CInt(resp.StatusCode)}: {s}", "warn")
                            Return Nothing
                        End If

                        Dim data = Await resp.Content.ReadAsByteArrayAsync().ConfigureAwait(False)

                        ' Detect content type
                        Dim mediaType = ""
                        Try
                            If resp.Content?.Headers?.ContentType?.MediaType IsNot Nothing Then
                                mediaType = resp.Content.Headers.ContentType.MediaType.ToLowerInvariant()
                            End If
                        Catch : End Try

                        ' PDF detection: content-type, URL, or magic bytes
                        Dim isPdf = False
                        If mediaType.Contains("pdf") Then isPdf = True
                        If Not isPdf AndAlso s.IndexOf(".pdf", StringComparison.OrdinalIgnoreCase) >= 0 Then isPdf = True
                        If Not isPdf AndAlso data IsNot Nothing AndAlso data.Length >= 4 Then
                            Dim scanMax = Math.Min(data.Length - 4, 1024)
                            For i = 0 To scanMax
                                If data(i) = AscW("%"c) AndAlso data(i + 1) = AscW("P"c) AndAlso
                                   data(i + 2) = AscW("D"c) AndAlso data(i + 3) = AscW("F"c) Then
                                    isPdf = True : Exit For
                                End If
                            Next
                        End If

                        If isPdf Then
                            Dim tmpPath = IO.Path.Combine(IO.Path.GetTempPath(), "ap_manual_" & Guid.NewGuid().ToString("N") & ".pdf")
                            Try
                                IO.File.WriteAllBytes(tmpPath, data)
                                Return Await SharedMethods.ReadPdfAsText(tmpPath, True, False, False, _context).ConfigureAwait(False)
                            Finally
                                Try : IO.File.Delete(tmpPath) : Catch : End Try
                            End Try
                        End If

                        ' Decode as text
                        Dim enc As Encoding = Encoding.UTF8
                        Try
                            Dim charset = resp.Content?.Headers?.ContentType?.CharSet
                            If Not String.IsNullOrEmpty(charset) Then enc = Encoding.GetEncoding(charset)
                        Catch : End Try

                        Dim text = enc.GetString(data)

                        ' Strip HTML if detected
                        If (mediaType.Contains("html") OrElse text.TrimStart().StartsWith("<", StringComparison.Ordinal)) AndAlso
                           text.IndexOf("<html", StringComparison.OrdinalIgnoreCase) >= 0 Then
                            Try
                                Dim htmlDoc As New HtmlAgilityPack.HtmlDocument()
                                htmlDoc.LoadHtml(text)
                                Return htmlDoc.DocumentNode.InnerText
                            Catch
                                Return text
                            End Try
                        End If

                        Return text
                    End Using
                End Using
            Catch ex As System.Exception
                ApDashboardLog($"Manual URL load error: {ex.Message}", "warn")
                Return Nothing
            End Try
        End If

        ' ── Local file ──
        Try
            If Not IO.File.Exists(s) Then
                ApDashboardLog($"Manual file not found: {s}", "warn")
                Return Nothing
            End If

            Select Case IO.Path.GetExtension(s).ToLowerInvariant()
                Case ".txt", ".md", ".log"
                    Return IO.File.ReadAllText(s, Encoding.UTF8)
                Case ".pdf"
                    Return Await SharedMethods.ReadPdfAsText(s, True, False, False, _context).ConfigureAwait(False)
                Case ".docx"
                    Dim text As String = Nothing
                    Dim label As String = Nothing
                    If TryExtractOfficeText(s, text, label) AndAlso Not String.IsNullOrWhiteSpace(text) Then
                        Return text
                    End If
                    Return Nothing
                Case ".rtf"
                    Return SharedMethods.ReadRtfAsText(s)
                Case Else
                    ' Best effort: read as plain text
                    Return IO.File.ReadAllText(s, Encoding.UTF8)
            End Select
        Catch ex As System.Exception
            ApDashboardLog($"Manual file load error: {ex.Message}", "warn")
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Optionally uses the InternetResearch model (if configured) to suggest
    ''' an alternative online tool the user could use. Returns a brief suggestion
    ''' or Nothing if no InternetResearch model is available.
    ''' Does NOT include the third-party disclaimer — callers are responsible
    ''' for appending it in the appropriate form (verbatim LLM instruction for
    ''' tool responses, direct text for static failure messages).
    ''' </summary>
    Private Async Function GetInternetAlternativeSuggestionAsync(
            mailInfo As AutoPilotMailInfo,
            failureReason As String,
            ct As CancellationToken) As Task(Of String)

        ' Check if an InternetResearch model is configured
        If String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then Return Nothing

        Dim backupConfig = GetCurrentConfig(_context)
        Dim hasInternetModel As Boolean = False

        Try
            hasInternetModel = GetSpecialTaskModel(_context, INI_AlternateModelPath, "InternetResearch")
        Catch
            ' GetSpecialTaskModel may have modified originalConfig even on failure — restore
            Try : RestoreDefaults(_context, backupConfig) : Catch : End Try
            Return Nothing
        End Try

        If Not hasInternetModel Then
            ' GetSpecialTaskModel overwrites the static originalConfig/originalConfigLoaded
            ' even when it returns False — always restore to avoid corrupting subsequent calls
            Try : RestoreDefaults(_context, backupConfig) : Catch : End Try
            Return Nothing
        End If

        Try
            Dim systemPrompt =
                $"You are a concise assistant. In 1-3 SHORT sentences, suggest other practical ways " &
                $"(other than {AN}) that could help the user accomplish the task described below without the help of another person. Focus on the most popular/common approaches. " &
                $"However, do not name specific tools, products, services or websites and stay very brief. If unsure, respond with exactly: NONE"

            Dim userPrompt =
                $"Task: {mailInfo.Subject}" & vbCrLf &
                $"Details: {If(mailInfo.Body.Length > 300, mailInfo.Body.Substring(0, 300), mailInfo.Body)}" & vbCrLf &
                $"Why it failed: {failureReason}"

            Dim result = Await LLM(systemPrompt, userPrompt,
                                   UseSecondAPI:=True, HideSplash:=True,
                                   EnsureUI:=False, cancellationToken:=ct)
            result = If(result, "").Trim()
            If result.Equals("NONE", StringComparison.OrdinalIgnoreCase) OrElse result.Length < 5 Then Return Nothing
            ApDashboardLog($"Internet alternative suggestion generated ({result.Length} chars)", "step")
            Return "**Alternative:** " & result
        Catch ex As System.Exception
            ApDashboardLog($"InternetResearch suggestion error: {ex.Message}", "warn")
            Return Nothing
        Finally
            RestoreDefaults(_context, backupConfig)
        End Try
    End Function

    ''' <summary>
    ''' Processes an incoming voicemail: extracts the audio attachment, identifies the caller ID,
    ''' looks up the mapped email address, transcribes the audio via the LLM's binary input,
    ''' treats the transcription as an instruction, generates a response, and sends the reply
    ''' to the mapped email address (not to the voicemail system sender).
    ''' </summary>
    Private Async Function ProcessVoicemailAsync(
            mi As MailItem,
            mailInfo As AutoPilotMailInfo,
            entryId As String,
            ct As CancellationToken) As Task

        ApDashboardLog("━━━ VOICEMAIL ━━━", "info")
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "AutoPilot (Voicemail) invoked")
        ApDashboardLog($"From voicemail system: {mailInfo.SenderEmail}", "info")
        ApDashboardLog($"Subject: {mailInfo.Subject}", "info")

        ' ── Find the audio attachment ──
        Dim audioFileName As String = Nothing
        For Each name In mailInfo.AttachmentNames
            Dim ext = IO.Path.GetExtension(name).ToLowerInvariant()
            For Each audioExt In AP_VoicemailAudioExtensions
                If ext = audioExt Then audioFileName = name : Exit For
            Next
            If audioFileName IsNot Nothing Then Exit For
        Next

        If audioFileName Is Nothing Then
            ApDashboardLog("SKIP (voicemail: no audio attachment found)", "warn")
            Return
        End If

        ' ── Extract caller ID ──
        Dim rawCallerId = ExtractCallerIdFromVoicemail(audioFileName, mailInfo.Subject, mailInfo.Body)
        If String.IsNullOrWhiteSpace(rawCallerId) Then
            ApDashboardLog($"SKIP (voicemail: no caller ID found in '{audioFileName}' or subject)", "warn")
            Return
        End If

        ApDashboardLog($"Caller ID extracted: {rawCallerId}", "info")

        ' ── Look up caller ID in map ──
        Dim lookup = LookupCallerIdEmail(rawCallerId)
        If Not lookup.HasValue Then
            ApDashboardLog($"SKIP (voicemail: caller ID '{rawCallerId}' not found in caller ID map)", "warn")
            Return
        End If

        Dim recipientEmail = lookup.Value.Email
        Dim recipientName = If(String.IsNullOrWhiteSpace(lookup.Value.DisplayName),
                               recipientEmail, lookup.Value.DisplayName)

        ApDashboardLog($"Caller mapped to: {recipientName} <{recipientEmail}>", "info")

        ' ── Save audio to temp ──
        Dim tempDir As String = IO.Path.Combine(IO.Path.GetTempPath(), AP_TempPrefix & Guid.NewGuid().ToString("N"))
        IO.Directory.CreateDirectory(tempDir)

        Try
            Dim audioTempPath As String = Nothing
            Await SwitchToUi(Sub()
                                 For i As Integer = 1 To mi.Attachments.Count
                                     Dim att = mi.Attachments(i)
                                     Try
                                         If att.FileName = audioFileName Then
                                             audioTempPath = IO.Path.Combine(tempDir, audioFileName)
                                             att.SaveAsFile(audioTempPath)
                                             Exit For
                                         End If
                                     Catch
                                     End Try
                                 Next
                             End Sub)

            If audioTempPath Is Nothing OrElse Not IO.File.Exists(audioTempPath) Then
                ApDashboardLog("SKIP (voicemail: failed to save audio attachment to temp)", "warn")
                Return
            End If

            ApDashboardLog($"Audio saved: {audioFileName} ({New IO.FileInfo(audioTempPath).Length / 1024:F0} KB)", "info")

            ' ── Transcribe the audio via LLM binary input ──
            ApDashboardLog("Transcribing voicemail...", "llm")

            ' Re-apply base model config for clean state
            ApplyModelConfig(_context, _apBaseModelConfig)

            Dim transcription = Await LLM(
                SP_InsertClipboard, "",
                Model:="", Temperature:="",
                Timeout:=_context.INI_Timeout * 2,
                UseSecondAPI:=_apUseSecondApi,
                HideSplash:=True, AddUserPrompt:="",
                FileObject:=audioTempPath,
                cancellationToken:=ct,
                EnsureUI:=False)

            If String.IsNullOrWhiteSpace(transcription) Then
                ApDashboardLog("SKIP (voicemail: transcription returned empty)", "warn")
                Return
            End If

            transcription = transcription.Trim()

            ApDashboardLog($"Transcription received ({transcription.Length} chars): {If(transcription.Length > 200, transcription.Substring(0, 200) & "...", transcription)}", "info")

            ' ── Build a synthetic mail info from the transcription ──
            Dim voicemailMailInfo As New AutoPilotMailInfo() With {
                .EntryID = entryId,
                .Subject = $"Voicemail from {rawCallerId}",
                .SenderName = recipientName,
                .SenderEmail = recipientEmail,
                .Body = transcription,
                .ReceivedTime = mailInfo.ReceivedTime,
                .SentOn = mailInfo.SentOn,
                .HasAutoReplyHeader = False,
                .ThreadAIReplyCount = 0,
                .AttachmentCount = 0,
                .AttachmentNames = New List(Of String)(),
                .FolderPath = mailInfo.FolderPath,
                .MessageClass = mailInfo.MessageClass,
                .InternetHeaders = ""
            }

            ' ── Initialize tool call log ──
            _apCurrentToolCallLog = New List(Of AutoPilotToolCallEntry)()
            _apCurrentTempDir = tempDir
            _apCurrentAttachments = New List(Of AutoPilotAttachmentInfo)()
            _apCurrentMailInfo = voicemailMailInfo

            Dim response As String

            Dim previousMaxToolIterations = MaxToolIterations
            MaxToolIterations = AP_MaxToolIterations

            ' Save references before Finally clears them
            Dim savedAttachments = _apCurrentAttachments

            Try
                ' Build prompt with the transcription as the "email body"
                Dim userPrompt = BuildUserPromptFromMail(voicemailMailInfo, Nothing)
                Dim systemPrompt = InterpolateAtRuntime(SP_AutoPilot)

                ' ── Inject per-user memory into system prompt (keyed on recipientEmail, not voicemail system) ──
                If _apConfig.EnableUserMemory AndAlso IsUserMemoryEnabled(recipientEmail) Then
                    Dim userMemory = ReadUserMemory(recipientEmail, _context.INI_InkyMemoryCap)
                    systemPrompt &= vbLf & _context.SP_Add_InkyMemory
                    If Not String.IsNullOrWhiteSpace(userMemory) Then
                        systemPrompt &= vbLf & "<INKY_MEMORY_CURRENT>" & vbLf & userMemory & vbLf & "</INKY_MEMORY_CURRENT>"
                    End If
                End If

                ' ── Inject user home file listing into user prompt ──
                If _apConfig.EnableUserFiles AndAlso HasUserHomeFiles(recipientEmail) Then
                    Dim homeFiles = ListUserHomeFiles(recipientEmail)
                    If homeFiles.Count > 0 Then
                        Dim fileListing As New StringBuilder()
                        fileListing.AppendLine()
                        fileListing.AppendLine("[USER HOME FILES — persistent files stored by this user for use in requests]")
                        For Each f In homeFiles
                            fileListing.AppendLine($"  - {f.Name} ({f.SizeBytes / 1024:F0} KB)")
                        Next
                        fileListing.AppendLine("Use manage_user_files with action='use' to load a file into the current session for processing.")
                        fileListing.AppendLine("[/USER HOME FILES]")
                        userPrompt &= fileListing.ToString()
                    End If
                End If

                ' Re-apply base model config
                ApplyModelConfig(_context, _apBaseModelConfig)

                Dim modelCanCallTools As Boolean = ModelSupportsTooling(_apBaseModelConfig)

                If modelCanCallTools AndAlso _apSelectedTools IsNot Nothing AndAlso _apSelectedTools.Count > 0 Then
                    ' Apply the additive per-sender tool policy using the caller's email of record
                    ' (the caller-ID-mapped address), not the voicemail system sender.
                    Dim senderTools As List(Of ModelConfig) = ResolveToolsForSender(recipientEmail, _apSelectedTools)
                    response = Await ExecuteToolingLoop(
                        systemPrompt, userPrompt,
                        senderTools, _apUseSecondApi,
                        hideSplash:=True, hideLogWindow:=True,
                        cancellationToken:=ct,
                        binaryOutputDirectory:=tempDir,
                        workflowId:=SharedLibrary.Agents.WorkflowContinuity.CreateWorkflowId(),
                        memoryGroundingMode:=SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingMode.None,
                        memoryGroundingModeIsExplicit:=True)
                Else
                    Dim effectiveSystemPrompt = If(modelCanCallTools, systemPrompt, InterpolateAtRuntime(SP_AutoPilot_NoTools))
                    response = Await LLM(effectiveSystemPrompt, userPrompt,
                                         UseSecondAPI:=_apUseSecondApi,
                                         HideSplash:=True, EnsureUI:=False,
                                         cancellationToken:=ct,
                                         binaryOutputDirectory:=tempDir)
                End If
            Finally
                _apCurrentTempDir = Nothing
                _apCurrentAttachments = Nothing
                _apCurrentMailInfo = Nothing
                MaxToolIterations = previousMaxToolIterations
                ClearAttachmentCaches()
            End Try

            If String.IsNullOrWhiteSpace(response) Then
                ApDashboardLog("WARNING: LLM returned empty response for voicemail", "warn")
                Return
            End If

            ApDashboardLog($"AI response received ({response.Length} chars).", "info")

            ' ── Process automatic memory learning from LLM response (keyed on recipientEmail) ──
            If _apConfig.EnableUserMemory AndAlso IsUserMemoryEnabled(recipientEmail) Then
                Dim userMemoryContent = ReadUserMemory(recipientEmail, _context.INI_InkyMemoryCap)
                Dim autoLearnDisabled = (userMemoryContent IsNot Nothing AndAlso
                    userMemoryContent.Contains("AUTO_LEARN_DISABLED"))

                If Not autoLearnDisabled Then
                    Try
                        response = ProcessUserMemoryResponse(recipientEmail, response, _context.INI_InkyMemoryCap)
                    Catch
                        response = StripInkyMemoryBlock(response)
                    End Try
                Else
                    response = StripInkyMemoryBlock(response)
                End If
            Else
                response = StripInkyMemoryBlock(response)
            End If

            ' Collect any result attachments (use saved reference, not cleared field)
            Dim resultAttachments = CollectResultAttachments(tempDir, savedAttachments)

            ' Build "Sources used:" footer
            Dim sourcesHtml = BuildSourcesUsedHtml(_apCurrentToolCallLog)

            ' ── Build voicemail footer note ──
            Dim voicemailNote = $"This reply was generated from a voicemail received on " &
                $"{mailInfo.ReceivedTime:yyyy-MM-dd HH:mm} from caller {rawCallerId}."

            ' ── Send reply to the MAPPED email address (not the voicemail sender) ──
            Await SwitchToUi(Sub()
                                 SendVoicemailReply(mi, recipientEmail, recipientName,
                                                    voicemailMailInfo.Subject,
                                                    response, resultAttachments,
                                                    sourcesHtml, voicemailNote)
                             End Sub)

            Await SwitchToUi(Sub() TagOriginalMailAsProcessed(mi))
            Interlocked.Increment(_apSessionReplyCount)
            RecordLastProcessedTime()
            ApDashboardLog($"✓ SENT voicemail reply to: {recipientEmail} (caller: {rawCallerId})", "info")

        Finally
            Try
                If IO.Directory.Exists(tempDir) Then IO.Directory.Delete(tempDir, recursive:=True)
            Catch
            End Try
            _apCurrentToolCallLog = Nothing
        End Try
    End Function

    ''' <summary>
    ''' Sends a reply for a voicemail to a specific recipient (the mapped email address),
    ''' not to the voicemail system sender. Creates a new MailItem rather than using Reply().
    ''' </summary>
    Private Sub SendVoicemailReply(originalMail As MailItem,
                                   recipientEmail As String,
                                   recipientName As String,
                                   subject As String,
                                   responseText As String,
                                   resultAttachments As List(Of String),
                                   sourcesHtml As String,
                                   voicemailNote As String)
        Dim newMail As MailItem = Nothing
        Try
            newMail = Application.CreateItem(OlItemType.olMailItem)
            newMail.To = recipientEmail
            newMail.Subject = "Re: " & subject
            newMail.BodyFormat = OlBodyFormat.olFormatHTML

            ' Build HTML body
            Dim htmlBody = ConvertResponseToHtml(responseText)

            ' Append voicemail note
            htmlBody &= "<br/><div style='font-size:9pt;color:#888888;font-style:italic;margin-top:12px;'>" &
                         System.Net.WebUtility.HtmlEncode(voicemailNote) & "</div>"

            ' Append sources
            If Not String.IsNullOrWhiteSpace(sourcesHtml) Then
                htmlBody &= sourcesHtml
            End If

            htmlBody &= BuildAutoPilotFooter()
            newMail.HTMLBody = htmlBody

            ' Add result attachments
            If resultAttachments IsNot Nothing Then
                For Each attachPath In resultAttachments
                    If IO.File.Exists(attachPath) Then
                        newMail.Attachments.Add(attachPath, OlAttachmentType.olByValue, , IO.Path.GetFileName(attachPath))
                    End If
                Next
            End If

            ' Tag as AutoPilot reply
            Try
                newMail.PropertyAccessor.SetProperty(AP_LoopHeaderProperty, AP_LoopHeaderValue)
            Catch : End Try
            Try : newMail.Categories = AP_CategoryName : Catch : End Try

            ' Use the same sending account as the monitored mailbox
            If Not String.IsNullOrWhiteSpace(_apConfig.MonitoredMailbox) Then
                Try
                    Dim ns = Application.GetNamespace("MAPI")
                    For i As Integer = 1 To ns.Accounts.Count
                        If ns.Accounts(i).SmtpAddress.Equals(_apConfig.MonitoredMailbox, StringComparison.OrdinalIgnoreCase) Then
                            newMail.SendUsingAccount = ns.Accounts(i)
                            Exit For
                        End If
                    Next
                Catch
                End Try
            End If

            Dim cleanupGroupId As String = Nothing
            Dim cleanupIsEligible As Boolean = False
            Dim cleanupAnsweredUtc As DateTime? = Nothing
            Dim cleanupDeleteAfterUtc As DateTime? = Nothing

            Try
                cleanupGroupId = GetOrCreateCleanupGroupId(originalMail)
                If Not String.IsNullOrWhiteSpace(cleanupGroupId) Then
                    cleanupDeleteAfterUtc = GetAutoDeleteCutoffUtc()
                    cleanupIsEligible = cleanupDeleteAfterUtc.HasValue
                    cleanupAnsweredUtc = If(cleanupDeleteAfterUtc.HasValue, CType(DateTime.UtcNow, DateTime?), CType(Nothing, DateTime?))

                    StampCleanupMetadata(
                        newMail,
                        cleanupGroupId,
                        isEligible:=cleanupIsEligible,
                        answeredUtc:=cleanupAnsweredUtc,
                        deleteAfterUtc:=cleanupDeleteAfterUtc,
                        saveItem:=False)
                End If
            Catch ex As System.Exception
                Debug.WriteLine($"[AutoPilot] Failed to stamp cleanup metadata on voicemail reply: {ex.Message}")
            End Try

            newMail.Send()
            Try : MoveLastSentToInkyReplies(cleanupGroupId, cleanupIsEligible, cleanupAnsweredUtc, cleanupDeleteAfterUtc) : Catch : End Try

        Catch ex As System.Exception
            ApDashboardLog($"ERROR sending voicemail reply: {ex.Message}", "error")
        Finally
            If newMail IsNot Nothing Then Try : Marshal.ReleaseComObject(newMail) : Catch : End Try
        End Try
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  DATA CLASSES
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Represents extracted mail metadata used by AutoPilot.</summary>
    Private Class AutoPilotMailInfo
        Public Property EntryID As String
        Public Property Subject As String
        Public Property SenderName As String
        Public Property SenderEmail As String
        Public Property Body As String
        Public Property ReceivedTime As DateTime
        Public Property SentOn As DateTime
        Public Property HasAutoReplyHeader As Boolean
        Public Property ThreadAIReplyCount As Integer
        Public Property AttachmentCount As Integer
        Public Property AttachmentNames As List(Of String)
        Public Property FolderPath As String
        Public Property MessageClass As String
        Public Property InternetHeaders As String
    End Class

    ''' <summary>Describes an attachment and its per-mail processing data.</summary>
    Friend Class AutoPilotAttachmentInfo
        Public Property OriginalFileName As String
        Public Property TempFilePath As String
        Public Property SourcePath As String
        Public Property Extension As String
        Public Property SizeBytes As Long
        Public Property IsOverSizeLimit As Boolean
        Public Property StatusMessage As String
        Public Property CreatedTime As DateTime?
        Public Property LastModifiedTime As DateTime?
        Public Property OutputFiles As New List(Of String)()
        Property CachedText As String
        Property CachedDocxHint As String
        Public Property IsToolOutput As Boolean = False

        ''' <summary>Number of pages (PDF only, 0 if unknown).</summary>
        Public Property PageCount As Integer = 0

        ''' <summary>Orientation of the first page: "portrait", "landscape", or Nothing if unknown.</summary>
        Public Property PageOrientation As String = Nothing

        ''' <summary>Page size description of the first page (e.g. "A4", "Letter", "595 × 842 pt"), or Nothing if unknown.</summary>
        Public Property PageSize As String = Nothing
    End Class

    ''' <summary>Defines supported filter rule types.</summary>
    Public Enum AutoPilotFilterRuleType
        Domain
        Sender
        Folder
    End Enum

    ''' <summary>Represents a single sender or folder filter rule.</summary>
    Public Class AutoPilotFilterRule
        Public Property RuleType As AutoPilotFilterRuleType
        Public Property Pattern As String
        Public Property IsNegative As Boolean
    End Class

    ''' <summary>Tracks a single tool call for logging and the "Sources used:" footer.</summary>
    Friend Class AutoPilotToolCallEntry
        Public Property ToolName As String
        Public Property ToolDisplayName As String
        Public Property ParamSummary As String
        Public Property IsInternalTool As Boolean
        Public Property WasSuccessful As Boolean
        Public Property ResultExcerpt As String
        Public Property Elapsed As TimeSpan

        Public Property Urls As List(Of String)
    End Class

End Class
