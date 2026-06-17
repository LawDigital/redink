' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.Scheduler.vb
' Purpose:
'   AutoPilot Scheduler — manages a persistent list of scheduled tasks stored
'   in a JSON file alongside the redink.ini configuration file. Tasks are
'   created, queried, updated, and deleted via natural language commands
'   processed by the LLM through the manage_scheduled_tasks internal tool.
'
' Key Features:
'  - Persistent Storage:
'      * Tasks persisted in autopilot_schedule.json in the same directory as redink.ini
'      * Re-read and re-written on every access to support live manual edits
'      * Per-task file attachments stored in schedule_tasks\{taskId}\
'      * Workspace files (persisted across executions) in schedule_tasks\{taskId}\workspace\
'      * Automatic quota enforcement (100 MB per task workspace) with oldest-file pruning
'
'  - Scheduling & Recurrence:
'      * Each task has a nextDueUtc field monitored by a periodic timer (30-second intervals)
'      * Recurrence expressed as iCalendar-style RRULE strings (FREQ=DAILY, WEEKLY, MONTHLY)
'      * Lightweight evaluator supports daily, weekly, monthly patterns with intervals,
'        weekday selection, nth-weekday-in-month, and count-limited occurrences
'      * LLM translates natural language schedules to structured RRULE + timeOfDayLocal fields
'      * One-shot (no RRULE) and recurring tasks both supported
'
'  - Task Execution Modes:
'      * Email mode: results delivered as e-mail to deliverTo addresses
'      * Browser prompt mode: Local Agent opens browser for interactive approval + execution
'      * Local Agent tasks may be queued, run, snoozed, or skipped by user
'      * Task instructions filtered to exclude manage_scheduled_tasks tool (no recursive task creation)
'
'  - Execution Pipeline:
'      * Due tasks processed through LLM + tooling with instruction as user prompt
'      * Input attachments and prior workspace files loaded into temp directory
'      * A persistent task-specific workspace is available for optional cross-run state
'      * Generated outputs automatically persisted to workspace for next execution
'      * Result e-mail sent to authorized recipients with HTML formatting and sources
'
'  - Catch-up & Reliability:
'      * On AutoPilot start, overdue tasks executed immediately
'      * Orphaned task directories cleaned up automatically
'      * Failed task execution does NOT advance schedule (will retry on next cycle)
'      * Snoozed Local Agent tasks can be resumed or skipped interactively
'
'  - Security & Control:
'      * Task management restricted to whitelisted/approval-required senders (AutoPilot filter rules)
'      * Local Chat access controlled via INI_AutoPilotSchedulerLocalChat setting
'      * Task recipients validated against creator to prevent unauthorized delivery
'      * Dashboard UI provides pause/resume, delete, and JSON editing
'
'  - Threading & Marshaling:
'      * File I/O synchronous (single-process, no multi-instance locking)
'      * Timer callback protected by Interlocked flag to prevent overlapping checks
'      * Outlook COM access marshaled to UI thread via SwitchToUi
'      * Separate CancellationTokenSource for Local Agent scheduler runtime
'
'  - Dashboard & Monitoring:
'      * DataGridView form shows all tasks with status, schedule, execution mode, snooz state, next due time
'      * Color-coded task status (active=green, paused=orange, completed=gray, failed=red)
'      * In-place JSON editor for manual task creation/modification
'      * Pause/resume and delete operations refresh display immediately
'
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Outlook
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ' ═══════════════════════════════════════════════════════════════════════════
    '  SCHEDULER CONSTANTS
    ' ═══════════════════════════════════════════════════════════════════════════

    Private AP_ScheduleFileName As String = AN4 & "autopilot_schedule.json"
    Private Const AP_ScheduleAttachmentDir As String = "schedule_tasks"
    Private Const AP_ScheduleWorkspaceSubdir As String = "workspace"
    Private Const AP_ScheduleWorkspaceMaxBytes As Long = 100 * 1024 * 1024  ' 100 MB per task
    Private Const AP_SchedulerCheckIntervalSeconds As Integer = 30
    Private Const AP_SchedulerEmailFailureRetryLimit As Integer = 2
    Private Const AP_SchedulerLocalChatSnoozeMinutes As Integer = 15
    Private Const AP_SchedulerLocalChatUrl As String = "http://localhost:12333/inky"
    Private Const AP_SchedulerLocalChatId As Integer = 2
    Private Const AP_TaskExecutionModeEmail As String = "email"
    Private Const AP_TaskExecutionModeBrowserPrompt As String = "browser_prompt"
    Private Const AP_TaskLocalExecutionStateQueued As String = "queued"
    Private Const AP_TaskLocalExecutionStateRunning As String = "running"

    ' ═══════════════════════════════════════════════════════════════════════════
    '  SCHEDULER STATE
    ' ═══════════════════════════════════════════════════════════════════════════

    Private _apSchedulerTimer As System.Threading.Timer = Nothing
    Private _apSchedulerCheckRunning As Integer = 0
    Private _apLocalSchedulerCts As CancellationTokenSource = Nothing
    Private _apScheduledWorkspaceTaskId As String = ""
    Private _apScheduledWorkspaceRoot As String = ""


    ' ═══════════════════════════════════════════════════════════════════════════
    '  DATA MODEL
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Represents a single scheduled task persisted to JSON.</summary>
    Public Class ScheduledTask
        ''' <summary>Unique identifier (GUID).</summary>
        Public Property Id As String = ""

        ''' <summary>The instruction / task description to execute.</summary>
        Public Property Instruction As String = ""

        ''' <summary>E-mail addresses to deliver the result to.</summary>
        Public Property DeliverTo As New List(Of String)()

        ''' <summary>Human-readable schedule description (e.g. "every Monday at 08:00").</summary>
        Public Property ScheduleDescription As String = ""

        ''' <summary>iCalendar-style RRULE string for recurrence (empty = one-shot).</summary>
        Public Property Rrule As String = ""

        ''' <summary>Local time-of-day for execution (HH:mm). Used together with RRULE.</summary>
        Public Property TimeOfDayLocal As String = ""

        ''' <summary>UTC timestamp for the next scheduled execution.</summary>
        Public Property NextDueUtc As DateTime = DateTime.MaxValue

        ''' <summary>UTC timestamp of the last successful execution (Nothing/MinValue if never).</summary>
        Public Property LastExecutedUtc As DateTime = DateTime.MinValue

        ''' <summary>UTC end date for the recurrence (Nothing/MaxValue = no end).</summary>
        Public Property EndDateUtc As DateTime = DateTime.MaxValue

        ''' <summary>Remaining occurrences (0 = unlimited, -1 = exhausted).</summary>
        Public Property RemainingOccurrences As Integer = 0

        ''' <summary>Relative path to the task's attachment subdirectory.</summary>
        Public Property AttachmentDir As String = ""

        ''' <summary>Filenames stored in the attachment directory (original input files).</summary>
        Public Property AttachmentFiles As New List(Of String)()

        ''' <summary>Filenames stored in the workspace directory (persisted across executions).</summary>
        Public Property WorkspaceFiles As New List(Of String)()

        ''' <summary>UTC timestamp when the task was created.</summary>
        Public Property CreatedUtc As DateTime = DateTime.UtcNow

        ''' <summary>E-mail address of the user who created the task.</summary>
        Public Property CreatedBy As String = ""

        ''' <summary>Task status: "active", "paused", "completed", "failed".</summary>
        Public Property Status As String = "active"

        ''' <summary>Brief summary of the last execution result.</summary>
        Public Property LastResult As String = Nothing

        ''' <summary>Subject line for the result e-mail.</summary>
        Public Property Subject As String = ""

        ''' <summary>
        ''' Execution mode: "email" = classic AutoPilot e-mail delivery,
        ''' "browser_prompt" = Local Agent popup approval + browser execution.
        ''' </summary>
        Public Property ExecutionMode As String = "email"

        ''' <summary>UTC snooze-until timestamp for Local Agent prompting (MinValue = not snoozed).</summary>
        Public Property SnoozeUntilUtc As DateTime = DateTime.MinValue

        ''' <summary>Local Agent execution state: "", "queued", or "running".</summary>
        Public Property LocalExecutionState As String = ""

        ''' <summary>Number of failed e-mail execution retries for the current due occurrence.</summary>
        Public Property FailureRetryCount As Integer = 0
    End Class

    ''' <summary>Root object for the schedule JSON file.</summary>
    Private Class ScheduleFile
        Public Property Tasks As New List(Of ScheduledTask)()
    End Class

    ' ═══════════════════════════════════════════════════════════════════════════
    '  FILE I/O — always re-read / re-write for live manual editing support
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Returns the full path to the scheduler JSON file.</summary>
    Private Function GetScheduleFilePath() As String
        Dim iniPath = GetDefaultINIPath(_context.RDV)
        Dim iniDir = Path.GetDirectoryName(iniPath)
        Return Path.Combine(iniDir, AP_ScheduleFileName)
    End Function

    ''' <summary>Returns the base directory for task attachment subdirectories.</summary>
    Private Function GetScheduleAttachmentBaseDir() As String
        Return Path.Combine(Path.GetDirectoryName(GetScheduleFilePath()), AP_ScheduleAttachmentDir)
    End Function

    ''' <summary>Returns the full path to the workspace directory for a given task.</summary>
    Private Function GetTaskWorkspaceDir(taskId As String) As String
        Return Path.Combine(Path.GetDirectoryName(GetScheduleFilePath()),
                            AP_ScheduleAttachmentDir, taskId, AP_ScheduleWorkspaceSubdir)
    End Function

    ''' <summary>Returns the full path to the input (attachment) directory for a given task.</summary>
    Private Function GetTaskInputDir(taskId As String) As String
        Return Path.Combine(Path.GetDirectoryName(GetScheduleFilePath()),
                            AP_ScheduleAttachmentDir, taskId)
    End Function

    ''' <summary>Reads and deserializes the schedule file. Returns empty schedule if file missing or corrupt.</summary>
    Private Function ReadScheduleFile() As ScheduleFile
        Try
            Dim filePath = GetScheduleFilePath()
            If Not File.Exists(filePath) Then Return New ScheduleFile()
            Dim json = File.ReadAllText(filePath, Encoding.UTF8)
            If String.IsNullOrWhiteSpace(json) Then Return New ScheduleFile()

            Dim sf = JsonConvert.DeserializeObject(Of ScheduleFile)(json)
            sf = If(sf, New ScheduleFile())

            If sf.Tasks Is Nothing Then sf.Tasks = New List(Of ScheduledTask)()

            For Each task In sf.Tasks
                NormalizeScheduledTask(task)
            Next

            Return sf
        Catch ex As System.Exception
            ApDashboardLog($"📅 ERROR reading schedule file: {ex.Message}", "error")
            Return New ScheduleFile()
        End Try
    End Function

    ''' <summary>Serializes and writes the schedule file atomically (write-to-temp then move).</summary>
    Private Sub WriteScheduleFile(schedule As ScheduleFile)
        Try
            Dim filePath = GetScheduleFilePath()
            Dim dir = Path.GetDirectoryName(filePath)
            If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)

            Dim json = JsonConvert.SerializeObject(schedule, Formatting.Indented)
            Dim tempPath = filePath & ".tmp"
            File.WriteAllText(tempPath, json, Encoding.UTF8)

            ' Atomic replace
            If File.Exists(filePath) Then File.Delete(filePath)
            File.Move(tempPath, filePath)

            ApDashboardLog($"📅 Schedule saved to: {filePath} ({schedule.Tasks.Count} task(s))", "step")
        Catch ex As System.Exception
            ApDashboardLog($"📅 ERROR writing schedule file: {ex.Message}", "error")
        End Try
    End Sub

    Private Sub NormalizeScheduledTask(task As ScheduledTask)
        If task Is Nothing Then Return

        If task.DeliverTo Is Nothing Then task.DeliverTo = New List(Of String)()
        If task.AttachmentFiles Is Nothing Then task.AttachmentFiles = New List(Of String)()
        If task.WorkspaceFiles Is Nothing Then task.WorkspaceFiles = New List(Of String)()

        If String.IsNullOrWhiteSpace(task.Status) Then task.Status = "active"
        If String.IsNullOrWhiteSpace(task.Subject) Then task.Subject = "Scheduled Task Result"

        task.ExecutionMode = NormalizeScheduledTaskExecutionMode(task.ExecutionMode)

        If task.SnoozeUntilUtc = DateTime.MaxValue Then
            task.SnoozeUntilUtc = DateTime.MinValue
        End If

        If task.LocalExecutionState Is Nothing Then
            task.LocalExecutionState = ""
        End If

        If task.FailureRetryCount < 0 Then
            task.FailureRetryCount = 0
        End If
    End Sub

    Private Shared Function NormalizeScheduledTaskExecutionMode(mode As String) As String
        If String.IsNullOrWhiteSpace(mode) Then Return AP_TaskExecutionModeEmail

        If mode.Equals(AP_TaskExecutionModeBrowserPrompt, StringComparison.OrdinalIgnoreCase) Then
            Return AP_TaskExecutionModeBrowserPrompt
        End If

        Return AP_TaskExecutionModeEmail
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TASK CRUD (called by the manage_scheduled_tasks tool executor)
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Creates a new scheduled task and persists it.</summary>
    Friend Function SchedulerCreateTask(task As ScheduledTask) As String
        If String.IsNullOrWhiteSpace(task.Id) Then task.Id = Guid.NewGuid().ToString("N")
        If task.CreatedUtc = DateTime.MinValue Then task.CreatedUtc = DateTime.UtcNow
        If String.IsNullOrWhiteSpace(task.Status) Then task.Status = "active"
        If String.IsNullOrWhiteSpace(task.Subject) Then task.Subject = "Scheduled Task Result"

        ' Create attachment directory if files are referenced
        If task.AttachmentFiles IsNot Nothing AndAlso task.AttachmentFiles.Count > 0 Then
            task.AttachmentDir = Path.Combine(AP_ScheduleAttachmentDir, task.Id)
            Dim fullDir = Path.Combine(Path.GetDirectoryName(GetScheduleFilePath()), task.AttachmentDir)
            If Not Directory.Exists(fullDir) Then Directory.CreateDirectory(fullDir)
        End If

        Dim schedule = ReadScheduleFile()
        schedule.Tasks.Add(task)
        WriteScheduleFile(schedule)

        ApDashboardLog($"📅 Scheduler: Created task {task.Id.Substring(0, 8)}... — ""{Truncate(task.Instruction, 60)}"" next due: {task.NextDueUtc:yyyy-MM-dd HH:mm} UTC", "info")
        RefreshSchedulerDashboard()
        Return task.Id
    End Function

    ''' <summary>Lists all tasks, optionally filtered by status.</summary>
    Friend Function SchedulerListTasks(Optional statusFilter As String = Nothing) As List(Of ScheduledTask)
        Dim schedule = ReadScheduleFile()
        If String.IsNullOrWhiteSpace(statusFilter) Then Return schedule.Tasks
        Return schedule.Tasks.Where(Function(t) t.Status.Equals(statusFilter, StringComparison.OrdinalIgnoreCase)).ToList()
    End Function

    ''' <summary>Finds a task by ID or by instruction substring match.</summary>
    Friend Function SchedulerFindTask(idOrQuery As String) As ScheduledTask
        Dim schedule = ReadScheduleFile()
        ' Try exact ID match first
        Dim byId = schedule.Tasks.FirstOrDefault(
            Function(t) t.Id.Equals(idOrQuery, StringComparison.OrdinalIgnoreCase) OrElse
                         t.Id.StartsWith(idOrQuery, StringComparison.OrdinalIgnoreCase))
        If byId IsNot Nothing Then Return byId

        ' Fuzzy match on instruction text
        Return schedule.Tasks.FirstOrDefault(
            Function(t) t.Instruction.IndexOf(idOrQuery, StringComparison.OrdinalIgnoreCase) >= 0)
    End Function

    ''' <summary>Updates an existing task (replaces the task with matching ID).</summary>
    Friend Function SchedulerUpdateTask(updated As ScheduledTask) As Boolean
        Dim schedule = ReadScheduleFile()
        Dim idx = schedule.Tasks.FindIndex(Function(t) t.Id.Equals(updated.Id, StringComparison.OrdinalIgnoreCase))
        If idx < 0 Then Return False
        schedule.Tasks(idx) = updated
        WriteScheduleFile(schedule)
        ApDashboardLog($"📅 Scheduler: Updated task {updated.Id.Substring(0, 8)}...", "info")
        RefreshSchedulerDashboard()
        Return True
    End Function

    ''' <summary>Deletes a task by ID. Also removes its attachment directory.</summary>
    Friend Function SchedulerDeleteTask(taskId As String) As Boolean
        Dim schedule = ReadScheduleFile()
        Dim task = schedule.Tasks.FirstOrDefault(Function(t) t.Id.Equals(taskId, StringComparison.OrdinalIgnoreCase))
        If task Is Nothing Then Return False

        schedule.Tasks.Remove(task)
        WriteScheduleFile(schedule)

        ' Clean up the full task directory, including its persistent workspace
        Try
            DeactivateScheduledTaskWorkspace(taskId)
        Catch
        End Try

        Try
            Dim fullDir = GetTaskInputDir(taskId)
            If Directory.Exists(fullDir) Then Directory.Delete(fullDir, recursive:=True)
        Catch
        End Try

        ApDashboardLog($"📅 Scheduler: Deleted task {taskId.Substring(0, Math.Min(8, taskId.Length))}...", "info")
        RefreshSchedulerDashboard()
        Return True
    End Function

    ''' <summary>Stores an attachment file for a task from the current AutoPilot temp directory.</summary>
    Friend Function SchedulerStoreAttachment(taskId As String, sourceFilePath As String) As Boolean
        Try
            Dim schedule = ReadScheduleFile()
            Dim task = schedule.Tasks.FirstOrDefault(Function(t) t.Id.Equals(taskId, StringComparison.OrdinalIgnoreCase))
            If task Is Nothing Then Return False

            If String.IsNullOrWhiteSpace(task.AttachmentDir) Then
                task.AttachmentDir = Path.Combine(AP_ScheduleAttachmentDir, task.Id)
            End If

            Dim fullDir = Path.Combine(Path.GetDirectoryName(GetScheduleFilePath()), task.AttachmentDir)
            If Not Directory.Exists(fullDir) Then Directory.CreateDirectory(fullDir)

            Dim destName = Path.GetFileName(sourceFilePath)
            Dim destPath = Path.Combine(fullDir, destName)
            File.Copy(sourceFilePath, destPath, overwrite:=True)

            If Not task.AttachmentFiles.Contains(destName, StringComparer.OrdinalIgnoreCase) Then
                task.AttachmentFiles.Add(destName)
            End If

            WriteScheduleFile(schedule)
            Return True
        Catch ex As System.Exception
            Debug.WriteLine($"[Scheduler] Error storing attachment: {ex.Message}")
            Return False
        End Try
    End Function


    ''' <summary>
    ''' Persists a file from the execution temp directory into the task's workspace.
    ''' Enforces a per-task size quota. If the quota would be exceeded, the oldest
    ''' workspace files are pruned until the new file fits.
    ''' </summary>
    Friend Function SchedulerStoreWorkspaceFile(taskId As String, sourceFilePath As String) As Boolean
        Try
            Dim schedule = ReadScheduleFile()
            Dim task = schedule.Tasks.FirstOrDefault(Function(t) t.Id.Equals(taskId, StringComparison.OrdinalIgnoreCase))
            If task Is Nothing Then Return False

            Dim wsDir = GetTaskWorkspaceDir(taskId)
            If Not Directory.Exists(wsDir) Then Directory.CreateDirectory(wsDir)

            Dim destName = Path.GetFileName(sourceFilePath)
            Dim destPath = Path.Combine(wsDir, destName)
            Dim newFileSize As Long = New FileInfo(sourceFilePath).Length

            ' Enforce workspace quota — prune oldest files until we fit
            EnforceWorkspaceQuota(wsDir, newFileSize)

            File.Copy(sourceFilePath, destPath, overwrite:=True)

            If task.WorkspaceFiles Is Nothing Then task.WorkspaceFiles = New List(Of String)()
            If Not task.WorkspaceFiles.Contains(destName, StringComparer.OrdinalIgnoreCase) Then
                task.WorkspaceFiles.Add(destName)
            End If

            WriteScheduleFile(schedule)
            ApDashboardLog($"📅 Workspace: stored '{destName}' for task {taskId.Substring(0, 8)}...", "step")
            Return True
        Catch ex As System.Exception
            Debug.WriteLine($"[Scheduler] Error storing workspace file: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Removes the oldest files in the workspace directory until there is room for
    ''' <paramref name="requiredBytes"/> additional bytes within the quota.
    ''' </summary>
    Private Shared Sub EnforceWorkspaceQuota(wsDir As String, requiredBytes As Long)
        If Not Directory.Exists(wsDir) Then Return
        Try
            Dim files = New DirectoryInfo(wsDir).GetFiles().OrderBy(Function(f) f.LastWriteTimeUtc).ToList()
            Dim currentSize As Long = files.Sum(Function(f) f.Length)

            While currentSize + requiredBytes > AP_ScheduleWorkspaceMaxBytes AndAlso files.Count > 0
                Dim oldest = files(0)
                currentSize -= oldest.Length
                oldest.Delete()
                files.RemoveAt(0)
            End While
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Scans the schedule_tasks base directory and deletes any subdirectory whose
    ''' name does not match an existing task ID. Called on scheduler start-up to
    ''' remove orphaned data from manually deleted tasks.
    ''' </summary>
    Friend Sub SchedulerPurgeOrphanDirectories()
        Try
            Dim baseDir = GetScheduleAttachmentBaseDir()
            If Not Directory.Exists(baseDir) Then Return

            Dim schedule = ReadScheduleFile()
            Dim activeIds As New HashSet(Of String)(
                schedule.Tasks.Select(Function(t) t.Id),
                StringComparer.OrdinalIgnoreCase)

            Dim purgedCount = 0
            For Each subDir In Directory.GetDirectories(baseDir)
                Dim dirName = Path.GetFileName(subDir)
                If Not activeIds.Contains(dirName) Then
                    Try
                        Directory.Delete(subDir, recursive:=True)
                        purgedCount += 1
                    Catch
                    End Try
                End If
            Next

            If purgedCount > 0 Then
                ApDashboardLog($"📅 Scheduler: purged {purgedCount} orphan task director(ies).", "info")
            End If
        Catch ex As System.Exception
            ApDashboardLog($"📅 Orphan purge error: {ex.Message}", "warn")
        End Try
    End Sub

    ''' <summary>
    ''' Synchronizes the task's WorkspaceFiles list with what is actually on disk.
    ''' Removes stale entries and adds any files found on disk but not in the list.
    ''' </summary>
    Private Sub SyncWorkspaceFileList(task As ScheduledTask)
        Dim wsDir = GetTaskWorkspaceDir(task.Id)
        If task.WorkspaceFiles Is Nothing Then task.WorkspaceFiles = New List(Of String)()

        If Not Directory.Exists(wsDir) Then
            task.WorkspaceFiles.Clear()
            Return
        End If

        Dim onDisk As New HashSet(Of String)(
            Directory.GetFiles(wsDir).Select(Function(f) Path.GetFileName(f)),
            StringComparer.OrdinalIgnoreCase)

        ' Remove entries that no longer exist on disk
        task.WorkspaceFiles.RemoveAll(Function(f) Not onDisk.Contains(f))

        ' Add any files on disk not yet tracked
        For Each f In onDisk
            If Not task.WorkspaceFiles.Contains(f, StringComparer.OrdinalIgnoreCase) Then
                task.WorkspaceFiles.Add(f)
            End If
        Next
    End Sub

    Private Function HasActiveScheduledTaskWorkspace() As Boolean
        Return Not String.IsNullOrWhiteSpace(_apScheduledWorkspaceTaskId) AndAlso
               Not String.IsNullOrWhiteSpace(_apScheduledWorkspaceRoot) AndAlso
               Directory.Exists(_apScheduledWorkspaceRoot)
    End Function

    Private Sub ActivateScheduledTaskWorkspace(taskId As String)
        If String.IsNullOrWhiteSpace(taskId) Then Return

        Dim wsDir = GetTaskWorkspaceDir(taskId)
        If Not Directory.Exists(wsDir) Then Directory.CreateDirectory(wsDir)

        _apScheduledWorkspaceTaskId = taskId.Trim()
        _apScheduledWorkspaceRoot = Path.GetFullPath(wsDir)

        SharedLibrary.Agents.PathPolicy.SetWorkspaceRoot(_apScheduledWorkspaceRoot)
        SharedLibrary.Agents.WorkspaceTools.SetActive(
            New SharedLibrary.Agents.WorkspaceState() With {
                .RootPath = _apScheduledWorkspaceRoot,
                .PersistUntilRevoked = False,
                .AllowRead = True,
                .AllowWrite = True,
                .AllowMoveCopyRename = True,
                .AllowDelete = True,
                .IncludeHiddenSystem = False
            })
    End Sub

    Private Sub DeactivateScheduledTaskWorkspace(Optional taskId As String = Nothing)
        If Not String.IsNullOrWhiteSpace(taskId) AndAlso
           Not String.Equals(_apScheduledWorkspaceTaskId, taskId.Trim(), StringComparison.OrdinalIgnoreCase) Then
            Return
        End If

        _apScheduledWorkspaceTaskId = ""
        _apScheduledWorkspaceRoot = ""

        If _chatAgentActive AndAlso Not _apActive Then
            SyncWorkspaceToPathPolicy(includeSessionTempFallback:=True)
        Else
            SharedLibrary.Agents.PathPolicy.SetWorkspaceRoot(Nothing)
            SharedLibrary.Agents.WorkspaceTools.SetActive(New SharedLibrary.Agents.WorkspaceState())
        End If
    End Sub

    Private Sub RefreshScheduledTaskWorkspaceTracking(taskId As String)
        If String.IsNullOrWhiteSpace(taskId) Then Return

        Dim task = SchedulerFindTask(taskId)
        If task Is Nothing Then Return

        Dim wsDir = GetTaskWorkspaceDir(taskId)
        If Directory.Exists(wsDir) Then
            EnforceWorkspaceQuota(wsDir, 0)
        End If

        SyncWorkspaceFileList(task)
        SchedulerUpdateTask(task)
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  SCHEDULER TIMER — periodic check for due tasks
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Starts the scheduler timer. Called from StartAutoPilotWithConfig when scheduler is enabled.</summary>
    Friend Sub StartSchedulerTimer()
        _apSchedulerTimer?.Dispose()
        _apSchedulerTimer = New System.Threading.Timer(
            AddressOf SchedulerTimerCallback,
            Nothing,
            dueTime:=TimeSpan.FromSeconds(5),
            period:=TimeSpan.FromSeconds(AP_SchedulerCheckIntervalSeconds))
        ApDashboardLog("📅 Scheduler timer started.", "info")
    End Sub

    ''' <summary>Starts the local scheduler runtime when Local Agent scheduling is enabled and the listener is live.</summary>
    Friend Sub EnsureLocalSchedulerTimerStarted()
        If Not IsLocalSchedulerEnabled() Then
            StopLocalSchedulerRuntime()
            Return
        End If

        If _apLocalSchedulerCts Is Nothing OrElse _apLocalSchedulerCts.IsCancellationRequested Then
            Try : _apLocalSchedulerCts?.Dispose() : Catch : End Try
            _apLocalSchedulerCts = New CancellationTokenSource()
        End If

        SchedulerPurgeOrphanDirectories()
        StartSchedulerTimer()
    End Sub

    ''' <summary>Stops the scheduler timer. Called from StopAutoPilot.</summary>
    Friend Sub StopSchedulerTimer()
        Try : _apSchedulerTimer?.Dispose() : Catch : End Try
        _apSchedulerTimer = Nothing
    End Sub

    ''' <summary>Stops the Local Agent scheduler runtime without affecting the AutoPilot runtime.</summary>
    Friend Sub StopLocalSchedulerRuntime()
        Try : _apLocalSchedulerCts?.Cancel() : Catch : End Try
        Try : _apLocalSchedulerCts?.Dispose() : Catch : End Try
        _apLocalSchedulerCts = Nothing

        If Not _apActive Then
            StopSchedulerTimer()
        End If
    End Sub

    Private Function IsLocalSchedulerEnabled() As Boolean
        If INI_WebServerBlock = 2 OrElse INI_WebServerBlock = 4 Then Return False
        If _apActive Then Return False

        Try
            Return httpListener IsNot Nothing AndAlso httpListener.IsListening
        Catch
            Return False
        End Try
    End Function

    Private Function IsSchedulerRuntimeAvailable() As Boolean
        If INI_WebServerBlock = 4 Then Return False
        If _apActive Then Return True
        Return IsLocalSchedulerEnabled()
    End Function

    ''' <summary>Timer callback that checks for and executes due tasks.</summary>
    Private Async Sub SchedulerTimerCallback(state As Object)
        If Not IsSchedulerRuntimeAvailable() Then
            If Not _apActive Then
                StopLocalSchedulerRuntime()
            End If
            Return
        End If
        If Interlocked.CompareExchange(_apSchedulerCheckRunning, 1, 0) <> 0 Then Return

        Try
            Dim effectiveCt As CancellationToken

            If _apActive Then
                If _apCts Is Nothing OrElse _apCts.IsCancellationRequested Then Return
                effectiveCt = _apCts.Token
            Else
                If _apLocalSchedulerCts Is Nothing OrElse _apLocalSchedulerCts.IsCancellationRequested Then Return
                effectiveCt = _apLocalSchedulerCts.Token
            End If

            Await CheckAndExecuteDueTasks(effectiveCt)
        Catch ex As OperationCanceledException
            ' Expected during shutdown
        Catch ex As System.Exception
            ApDashboardLog($"📅 Scheduler timer error: {ex.Message}", "warn")
        Finally
            Interlocked.Exchange(_apSchedulerCheckRunning, 0)
        End Try
    End Sub

    ''' <summary>Performs catch-up on AutoPilot start — executes any overdue tasks.</summary>
    Friend Sub SchedulerCatchUp()
        Try
            ' Purge orphan directories from manually deleted tasks
            SchedulerPurgeOrphanDirectories()

            Dim schedule = ReadScheduleFile()
            Dim now = DateTime.UtcNow
            Dim overdueCount = 0

            For Each task In schedule.Tasks
                If task.Status <> "active" Then Continue For
                If task.NextDueUtc <= now AndAlso task.LastExecutedUtc < task.NextDueUtc Then
                    overdueCount += 1
                End If
            Next

            If overdueCount > 0 Then
                ApDashboardLog($"📅 Scheduler: {overdueCount} overdue task(s) found — will execute on next timer cycle.", "info")
            Else
                Dim activeCount = schedule.Tasks.Where(Function(t) t.Status = "active").Count()
                ApDashboardLog($"📅 Scheduler: {activeCount} active task(s), none overdue.", "info")
            End If
        Catch ex As System.Exception
            ApDashboardLog($"📅 Scheduler catch-up error: {ex.Message}", "warn")
        End Try
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TASK EXECUTION
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Checks for due tasks and executes them sequentially.</summary>
    Private Async Function CheckAndExecuteDueTasks(ct As CancellationToken) As Task
        If Not IsSchedulerRuntimeAvailable() Then Return

        Dim schedule = ReadScheduleFile()
        Dim now = DateTime.UtcNow

        For Each task In schedule.Tasks.ToList()
            ct.ThrowIfCancellationRequested()

            If task.Status <> "active" Then Continue For
            If task.SnoozeUntilUtc > now Then Continue For
            If task.NextDueUtc > now Then Continue For

            If IsLocalInteractiveScheduledTask(task) Then
                If task.LocalExecutionState.Equals(AP_TaskLocalExecutionStateQueued, StringComparison.OrdinalIgnoreCase) Then Continue For
                If task.LocalExecutionState.Equals(AP_TaskLocalExecutionStateRunning, StringComparison.OrdinalIgnoreCase) Then Continue For

                Await HandleDueLocalInteractiveTaskAsync(task, ct)
                Continue For
            End If

            If task.LastExecutedUtc >= task.NextDueUtc Then Continue For

            ' This task is due — execute it
            ApDashboardLog($"📅 Executing scheduled task: {task.Id.Substring(0, 8)}... — ""{Truncate(task.Instruction, 60)}""", "info")

            Try
                Await ExecuteScheduledTask(task, ct)

                ' Mark as executed and advance schedule
                task.LastExecutedUtc = DateTime.UtcNow
                task.LastResult = "Completed successfully at " & DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss UTC")

                AdvanceTaskSchedule(task)
                SchedulerUpdateTask(task)

                ApDashboardLog($"📅 Task {task.Id.Substring(0, 8)}... completed. Next due: {If(task.Status = "completed", "(none — completed)", task.NextDueUtc.ToString("yyyy-MM-dd HH:mm UTC"))}", "info")

            Catch ex As OperationCanceledException
                Throw
            Catch ex As System.Exception
                HandleScheduledTaskExecutionFailure(task, ex)
            End Try
        Next
    End Function

    Private Sub HandleScheduledTaskExecutionFailure(task As ScheduledTask, ex As System.Exception)
        If task Is Nothing Then Return

        task.FailureRetryCount += 1

        If task.FailureRetryCount <= Math.Max(0, AP_SchedulerEmailFailureRetryLimit) Then
            task.LastResult =
                $"Failed at {DateTime.UtcNow:yyyy-MM-dd HH:mm:ss UTC}: {ex.Message} " &
                $"— retry {task.FailureRetryCount} of {AP_SchedulerEmailFailureRetryLimit} will be attempted"

            ApDashboardLog(
                $"📅 Task {task.Id.Substring(0, 8)}... FAILED: {ex.Message} — retry {task.FailureRetryCount} of {AP_SchedulerEmailFailureRetryLimit} will be attempted.",
                "warn")

            SchedulerUpdateTask(task)
            Return
        End If

        task.LastResult =
            $"Failed at {DateTime.UtcNow:yyyy-MM-dd HH:mm:ss UTC}: {ex.Message} " &
            $"— retry limit reached after {AP_SchedulerEmailFailureRetryLimit} retr{If(AP_SchedulerEmailFailureRetryLimit = 1, "y", "ies")}"

        AdvanceTaskSchedule(task)
        SchedulerUpdateTask(task)

        ApDashboardLog(
            $"📅 Task {task.Id.Substring(0, 8)}... FAILED: {ex.Message} — retry limit reached. Next due: {If(task.Status = "completed", "(none — completed)", task.NextDueUtc.ToString("yyyy-MM-dd HH:mm UTC"))}",
            "error")
    End Sub

    ''' <summary>
    ''' Executes a single scheduled task: sends instruction to LLM with tooling,
    ''' collects result, persists workspace files, and delivers result by e-mail.
    ''' </summary>
    Private Async Function ExecuteScheduledTask(task As ScheduledTask, ct As CancellationToken) As Task
        Dim tempDir As String = Nothing
        Dim previousMaxToolIterations = MaxToolIterations
        Dim executionState As InkyState = Nothing
        Dim executionModelConfig As ModelConfig = Nothing
        Dim executionSelectedTools As List(Of ModelConfig) = Nothing
        Dim executionUseSecondApi As Boolean = _apUseSecondApi
        Dim restoreExecutionConfig As Boolean = False
        Dim originalExecutionConfig As ModelConfig = Nothing

        Try
            ' Create isolated temp directory
            tempDir = Path.Combine(Path.GetTempPath(), AP_TempPrefix & "sched_" & Guid.NewGuid().ToString("N"))
            Directory.CreateDirectory(tempDir)

            ActivateScheduledTaskWorkspace(task.Id)

            ' ── Load input (original) attachments into temp dir ──
            Dim attachments As New List(Of AutoPilotAttachmentInfo)()
            Dim inputFileNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            If Not String.IsNullOrWhiteSpace(task.AttachmentDir) Then
                Dim fullAttDir = Path.Combine(Path.GetDirectoryName(GetScheduleFilePath()), task.AttachmentDir)
                If Directory.Exists(fullAttDir) Then
                    ' Only load files directly in the task dir (not workspace subdir)
                    For Each filePath In Directory.GetFiles(fullAttDir)
                        Dim destPath = Path.Combine(tempDir, Path.GetFileName(filePath))
                        File.Copy(filePath, destPath, overwrite:=True)
                        inputFileNames.Add(Path.GetFileName(filePath))
                        attachments.Add(New AutoPilotAttachmentInfo() With {
                            .OriginalFileName = Path.GetFileName(filePath),
                            .TempFilePath = destPath,
                            .Extension = Path.GetExtension(filePath).ToLowerInvariant(),
                            .SizeBytes = New FileInfo(destPath).Length,
                            .IsOverSizeLimit = False,
                            .StatusMessage = "OK"
                        })
                    Next
                End If
            End If

            ' ── Load workspace files into temp dir ──
            Dim workspaceFileNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim wsDir = GetTaskWorkspaceDir(task.Id)
            If Directory.Exists(wsDir) Then
                For Each filePath In Directory.GetFiles(wsDir)
                    Dim fileName = Path.GetFileName(filePath)
                    ' Avoid collision with input files — workspace files get a prefix
                    Dim destName = fileName
                    If inputFileNames.Contains(fileName) Then
                        destName = "ws_" & fileName
                    End If
                    Dim destPath = Path.Combine(tempDir, destName)
                    File.Copy(filePath, destPath, overwrite:=True)
                    workspaceFileNames.Add(destName)
                    attachments.Add(New AutoPilotAttachmentInfo() With {
                        .OriginalFileName = destName,
                        .TempFilePath = destPath,
                        .Extension = Path.GetExtension(filePath).ToLowerInvariant(),
                        .SizeBytes = New FileInfo(destPath).Length,
                        .IsOverSizeLimit = False,
                        .StatusMessage = "OK (workspace)"
                    })
                Next
            End If

            ' Set transient state so internal tools can access attachments
            _apCurrentTempDir = tempDir
            _apCurrentAttachments = attachments
            _apCurrentMailInfo = New AutoPilotMailInfo() With {
                .EntryID = "scheduler_" & task.Id,
                .Subject = If(task.Subject, "Scheduled Task"),
                .SenderName = "Scheduler",
                .SenderEmail = If(task.CreatedBy, ""),
                .ReceivedTime = DateTime.Now,
                .Body = task.Instruction
            }
            _apCurrentToolCallLog = New List(Of AutoPilotToolCallEntry)()
            MaxToolIterations = AP_MaxToolIterations

            ' Build prompts — tell the LLM it is executing a scheduled task
            Dim userPrompt As StringBuilder =
                BuildScheduledTaskExecutionPrompt(
                    task,
                    inputFileNames,
                    workspaceFileNames)

            Dim systemPrompt = InterpolateAtRuntime(SP_AutoPilot)

            ResolveScheduledTaskExecutionContext(
                task,
                executionState,
                executionModelConfig,
                executionSelectedTools,
                executionUseSecondApi,
                restoreExecutionConfig,
                originalExecutionConfig)

            If _apActive AndAlso executionModelConfig IsNot Nothing Then
                ApplyModelConfig(_context, executionModelConfig)
            End If

            ' Execute with tooling
            Dim response As String
            Dim modelCanCallTools As Boolean =
                ScheduledTaskExecutionSupportsTooling(
                    executionState,
                    executionModelConfig,
                    executionUseSecondApi)

            If modelCanCallTools AndAlso executionSelectedTools IsNot Nothing AndAlso executionSelectedTools.Count > 0 Then
                ' Filter out manage_scheduled_tasks — the LLM must not create/modify
                ' tasks while executing inside a scheduled task
                Dim schedulerTools = executionSelectedTools.Where(
                    Function(t) t IsNot Nothing AndAlso
                                Not String.IsNullOrWhiteSpace(t.ToolName) AndAlso
                                Not t.ToolName.Equals(AP_Tool_ManageScheduledTasks, StringComparison.OrdinalIgnoreCase)).ToList()

                If schedulerTools.Count > 0 Then
                    response = Await ExecuteToolingLoop(
                        systemPrompt, userPrompt.ToString(),
                        schedulerTools, executionUseSecondApi,
                        hideSplash:=True, hideLogWindow:=True,
                        cancellationToken:=ct, binaryOutputDirectory:=tempDir)
                Else
                    response = Await LLM(systemPrompt, userPrompt.ToString(),
                                         UseSecondAPI:=executionUseSecondApi,
                                         HideSplash:=True, EnsureUI:=False,
                                         cancellationToken:=ct,
                                         binaryOutputDirectory:=tempDir)
                End If
            Else
                Dim effectiveSystemPrompt = If(modelCanCallTools, systemPrompt, InterpolateAtRuntime(SP_AutoPilot_NoTools))
                response = Await LLM(effectiveSystemPrompt, userPrompt.ToString(),
                                     UseSecondAPI:=executionUseSecondApi,
                                     HideSplash:=True, EnsureUI:=False,
                                     cancellationToken:=ct,
                                     binaryOutputDirectory:=tempDir)
            End If

            If String.IsNullOrWhiteSpace(response) Then
                Throw New InvalidOperationException("LLM returned empty response for scheduled task")
            End If

            Dim taskStatus = SharedLibrary.Agents.TaskStatusFooterParser.Parse(response)

            If taskStatus IsNot Nothing Then
                Select Case taskStatus.Kind
                    Case SharedLibrary.Agents.TaskStatusKind.Blocked
                        Throw New InvalidOperationException(
                            "Scheduled task execution was blocked" &
                            If(String.IsNullOrWhiteSpace(taskStatus.Reason), ".", ": " & taskStatus.Reason))

                    Case SharedLibrary.Agents.TaskStatusKind.Invalid
                        Throw New InvalidOperationException(
                            "Scheduled task execution returned an invalid TASK_STATUS footer" &
                            If(String.IsNullOrWhiteSpace(taskStatus.InvalidDetail), ".", ": " & taskStatus.InvalidDetail))
                End Select
            End If

            response = SharedLibrary.Agents.TaskStatusFooterParser.Strip(response)

            ' Collect result attachments
            Dim resultAttachments = CollectResultAttachments(tempDir, attachments)
            Dim sourcesHtml = BuildSourcesUsedHtml(_apCurrentToolCallLog)

            ' ── Persist new tool outputs to workspace ──
            If resultAttachments IsNot Nothing AndAlso resultAttachments.Count > 0 Then
                PersistResultsToWorkspace(task.Id, tempDir, resultAttachments, inputFileNames, workspaceFileNames)
                ' Re-sync the workspace file list
                SyncWorkspaceFileList(task)
            End If

            ' Send result e-mail
            Await SwitchToUi(Sub() SendScheduledTaskResult(task, response, resultAttachments, sourcesHtml))

            Interlocked.Increment(_apSessionReplyCount)
            RecordLastProcessedTime()

        Finally
            _apCurrentTempDir = Nothing
            _apCurrentAttachments = Nothing
            _apCurrentMailInfo = Nothing
            _apCurrentToolCallLog = Nothing

            Try
                If restoreExecutionConfig AndAlso originalExecutionConfig IsNot Nothing Then
                    RestoreDefaults(_context, originalExecutionConfig)
                End If
            Catch
            End Try

            MaxToolIterations = previousMaxToolIterations
            ClearAttachmentCaches()

            Try
                RefreshScheduledTaskWorkspaceTracking(task.Id)
            Catch
            End Try

            DeactivateScheduledTaskWorkspace(task.Id)

            Try
                If tempDir IsNot Nothing AndAlso Directory.Exists(tempDir) Then
                    Directory.Delete(tempDir, recursive:=True)
                End If
            Catch
            End Try
        End Try
    End Function

    Private Sub EnsureScheduledTaskLocalChatToolSelection(st As InkyState,
                                                          Optional executionMode As String = AP_TaskExecutionModeEmail)
        If st Is Nothing Then Return

        MigrateLocalChatToolSelections(st, includeInteractiveM365Tools:=True)
        st.ToolingEnabled = True
        st.AgentModeEnabled = True

        Dim allowedMainToolNames As New HashSet(Of String)(
            GetScheduledTaskAllMainSelectableToolNames(executionMode),
            StringComparer.OrdinalIgnoreCase)

        Dim allowedAdvancedToolNames As New HashSet(Of String)(
            GetScheduledTaskAllAdvancedExecutionTools(executionMode).
                Where(Function(t) t IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(t.ToolName)).
                Select(Function(t) t.ToolName.Trim()),
            StringComparer.OrdinalIgnoreCase)

        st.SelectedMainToolNames =
            If(st.SelectedMainToolNames, New List(Of String)()).
                Where(Function(name) Not String.IsNullOrWhiteSpace(name) AndAlso
                                     allowedMainToolNames.Contains(name.Trim())).
                Select(Function(name) name.Trim()).
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToList()

        st.SelectedAdvancedToolNames =
            ResolveLocalChatAdvancedToolNamesForEnabledState(
                st.SelectedAdvancedToolNames,
                includeInteractiveM365Tools:=True).
                    Where(Function(name) Not String.IsNullOrWhiteSpace(name) AndAlso
                                         allowedAdvancedToolNames.Contains(name.Trim())).
                    Select(Function(name) name.Trim()).
                    Distinct(StringComparer.OrdinalIgnoreCase).
                    ToList()

        Dim effectiveTools =
            FilterScheduledTaskExecutableTools(
                GetLocalChatEffectiveSelection(st, includeInteractiveM365Tools:=True),
                executionMode)

        If (effectiveTools Is Nothing OrElse effectiveTools.Count = 0) AndAlso
           Not IsChatAgentWorkspaceConnected() Then

            st.SelectedMainToolNames = GetScheduledTaskAllMainSelectableToolNames(executionMode)

            st.SelectedAdvancedToolNames =
                GetScheduledTaskAllAdvancedExecutionTools(executionMode).
                    Select(Function(t) t.ToolName.Trim()).
                    Distinct(StringComparer.OrdinalIgnoreCase).
                    ToList()

            effectiveTools =
                FilterScheduledTaskExecutableTools(
                    GetLocalChatEffectiveSelection(st, includeInteractiveM365Tools:=True),
                    executionMode)
        End If

        st.SelectedToolNames =
            If(effectiveTools, New List(Of ModelConfig)()).
                Where(Function(t) t IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(t.ToolName)).
                Select(Function(t) t.ToolName.Trim()).
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToList()
    End Sub

    Private Function GetScheduledTaskExecutionState(Optional executionMode As String = AP_TaskExecutionModeEmail) As InkyState
        If _apActive Then Return Nothing

        Dim sourceChatId As Integer = 1
        If activeChatId > 0 AndAlso activeChatId <> AP_SchedulerLocalChatId Then
            sourceChatId = activeChatId
        End If

        Dim st As InkyState = Nothing

        Try
            st = LoadInkyState(sourceChatId)
        Catch
        End Try

        If st Is Nothing Then
            Try
                st = LoadInkyState()
            Catch
                Return Nothing
            End Try
        End If

        If st Is Nothing Then Return Nothing

        EnsureScheduledTaskLocalChatToolSelection(st, executionMode)

        Return st
    End Function

    Private Function IsScheduledTaskEmailExecution(task As ScheduledTask) As Boolean
        If task Is Nothing Then Return True
        Return Not IsLocalInteractiveScheduledTask(task)
    End Function

    Private Function IsScheduledTaskToolExecutableForMode(tool As ModelConfig,
                                                          executionMode As String) As Boolean
        If tool Is Nothing OrElse String.IsNullOrWhiteSpace(tool.ToolName) Then Return False

        Dim toolName = tool.ToolName.Trim()
        Dim normalizedExecutionMode = NormalizeScheduledTaskExecutionMode(executionMode)

        If toolName.Equals(AP_Tool_ManageScheduledTasks, StringComparison.OrdinalIgnoreCase) Then
            Return False
        End If

        If toolName.Equals(SharedLibrary.Agents.ToolLoaderTool.LoaderToolName, StringComparison.OrdinalIgnoreCase) Then
            Return True
        End If

        If SharedLibrary.Agents.AgentToolRouter.IsAgentLayerTool(toolName) Then
            Return True
        End If

        If toolName.StartsWith("skill_", StringComparison.OrdinalIgnoreCase) Then
            Return True
        End If

        If toolName.Equals(InternalWebToolName, StringComparison.OrdinalIgnoreCase) OrElse
           toolName.Equals(InternalDownloadWebFilesToolName, StringComparison.OrdinalIgnoreCase) OrElse
           toolName.Equals(InternalSearchToolName, StringComparison.OrdinalIgnoreCase) OrElse
           IsInternalKnowledgeToolName(toolName) Then
            Return True
        End If

        If SharedLibrary.SharedLibrary.M365ToolService.IsM365ToolName(toolName) Then
            Return Not _apActive
        End If

        If IsAutoPilotInternalTool(toolName) Then
            Return _apActive OrElse
                   normalizedExecutionMode.Equals(AP_TaskExecutionModeBrowserPrompt, StringComparison.OrdinalIgnoreCase)
        End If

        If Not String.IsNullOrWhiteSpace(tool.ToolAPICall) OrElse
           Not String.IsNullOrWhiteSpace(tool.APICall) Then
            Return True
        End If

        Return SharedLibrary.Agents.ToolExecutorRegistry.IsAdvertisable(
            SharedLibrary.Agents.ToolingHostKind.Outlook,
            toolName)
    End Function

    Private Function FilterScheduledTaskExecutableTools(tools As IEnumerable(Of ModelConfig),
                                                        executionMode As String) As List(Of ModelConfig)
        Return DeduplicateToolsByName(
            If(tools, Enumerable.Empty(Of ModelConfig)()).
                Where(Function(t) IsScheduledTaskToolExecutableForMode(t, executionMode)))
    End Function

    Private Function GetScheduledTaskAllMainSelectableToolNames(executionMode As String) As List(Of String)
        Return FilterScheduledTaskExecutableTools(
            GetLocalChatMainSelectableTools(includeInteractiveM365Tools:=True),
            executionMode).
                Where(Function(t) t IsNot Nothing AndAlso
                                  Not String.IsNullOrWhiteSpace(t.ToolName) AndAlso
                                  Not t.ToolName.Equals(AP_Tool_ManageScheduledTasks, StringComparison.OrdinalIgnoreCase)).
                Select(Function(t) t.ToolName.Trim()).
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToList()
    End Function

    Private Function GetScheduledTaskWorkspaceTools() As List(Of ModelConfig)
        Dim allowedNames As New HashSet(Of String)(
            New String() {
                SharedLibrary.Agents.WorkspaceTools.ToolGet,
                SharedLibrary.Agents.WorkspaceTools.ToolInventory,
                SharedLibrary.Agents.WorkspaceTools.ToolRead,
                SharedLibrary.Agents.WorkspaceTools.ToolWrite,
                SharedLibrary.Agents.WorkspaceTools.ToolSearch,
                SharedLibrary.Agents.WorkspaceTools.ToolCopy,
                SharedLibrary.Agents.WorkspaceTools.ToolMove,
                SharedLibrary.Agents.WorkspaceTools.ToolRename,
                SharedLibrary.Agents.WorkspaceTools.ToolDelete,
                SharedLibrary.Agents.WorkspaceTools.ToolMakeDir
            },
            StringComparer.OrdinalIgnoreCase)

        Return SharedLibrary.Agents.WorkspaceTools.BuildAll().
            Where(Function(t) t IsNot Nothing AndAlso
                              Not String.IsNullOrWhiteSpace(t.ToolName) AndAlso
                              allowedNames.Contains(t.ToolName.Trim())).
            ToList()
    End Function

    Private Function GetScheduledTaskAllAdvancedExecutionTools(executionMode As String) As List(Of ModelConfig)
        Dim tools As New List(Of ModelConfig)()

        tools.AddRange(GetLocalChatAdvancedSelectableTools(includeInteractiveM365Tools:=True))
        tools.AddRange(GetScheduledTaskWorkspaceTools())

        Return FilterScheduledTaskExecutableTools(
            tools,
            executionMode).
                Where(Function(t) t IsNot Nothing AndAlso
                                  Not String.IsNullOrWhiteSpace(t.ToolName) AndAlso
                                  Not t.ToolName.Equals(AP_Tool_ManageScheduledTasks, StringComparison.OrdinalIgnoreCase)).
                ToList()
    End Function

    Private Sub EnsureScheduledTaskLocalChatEmailMainToolSelection(st As InkyState)
        If st Is Nothing Then Return

        MigrateLocalChatToolSelections(st, includeInteractiveM365Tools:=True)

        Dim hasMainSelection As Boolean =
            st.SelectedMainToolNames IsNot Nothing AndAlso
            st.SelectedMainToolNames.Any(Function(name) Not String.IsNullOrWhiteSpace(name))

        If hasMainSelection Then Return

        st.SelectedMainToolNames = GetScheduledTaskAllMainSelectableToolNames(AP_TaskExecutionModeEmail)
    End Sub

    Private Function MergeScheduledTaskExecutionTools(selectedTools As IEnumerable(Of ModelConfig),
                                                      includeAllAdvancedTools As Boolean,
                                                      executionMode As String) As List(Of ModelConfig)
        Dim merged As New List(Of ModelConfig)()

        If selectedTools IsNot Nothing Then
            merged.AddRange(
                selectedTools.Where(
                    Function(t) t IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(t.ToolName)))
        End If

        If includeAllAdvancedTools Then
            merged.AddRange(GetScheduledTaskAllAdvancedExecutionTools(executionMode))
        End If

        Return FilterScheduledTaskExecutableTools(merged, executionMode)
    End Function

    Private Sub ResolveScheduledTaskExecutionContext(task As ScheduledTask,
                                                     ByRef executionState As InkyState,
                                                     ByRef executionModelConfig As ModelConfig,
                                                     ByRef executionSelectedTools As List(Of ModelConfig),
                                                     ByRef executionUseSecondApi As Boolean,
                                                     ByRef restoreExecutionConfig As Boolean,
                                                     ByRef originalExecutionConfig As ModelConfig)
        executionState = Nothing
        executionModelConfig = _apBaseModelConfig

        Dim executionMode As String = AP_TaskExecutionModeEmail
        If task IsNot Nothing Then
            executionMode = NormalizeScheduledTaskExecutionMode(task.ExecutionMode)
        End If

        executionSelectedTools =
            MergeScheduledTaskExecutionTools(
                If(_apSelectedTools IsNot Nothing,
                   New List(Of ModelConfig)(_apSelectedTools),
                   New List(Of ModelConfig)()),
                includeAllAdvancedTools:=True,
                executionMode:=executionMode)
        executionUseSecondApi = _apUseSecondApi
        restoreExecutionConfig = False
        originalExecutionConfig = Nothing

        If _apActive Then Return

        executionState = GetScheduledTaskExecutionState(executionMode)
        If executionState Is Nothing Then Return

        Dim isEmailExecution As Boolean = IsScheduledTaskEmailExecution(task)

        If isEmailExecution Then
            EnsureScheduledTaskLocalChatEmailMainToolSelection(executionState)
        End If

        executionUseSecondApi = executionState.UseSecondApi
        executionSelectedTools =
            MergeScheduledTaskExecutionTools(
                GetLocalChatEffectiveSelection(executionState, includeInteractiveM365Tools:=True),
                includeAllAdvancedTools:=isEmailExecution,
                executionMode:=executionMode)

        If executionSelectedTools Is Nothing Then
            executionSelectedTools = New List(Of ModelConfig)()
        End If

        If Not executionUseSecondApi Then Return

        originalExecutionConfig = GetCurrentConfig(_context)

        Dim selectedModelKey = executionState.SelectedModelKey

        If String.IsNullOrWhiteSpace(selectedModelKey) Then
            executionModelConfig = originalExecutionConfig
        Else
            Try
                Dim alts = LoadAlternativeModels(INI_AlternateModelPath, _context)
                Dim selectedModel = alts?.FirstOrDefault(
                    Function(m)
                        If m Is Nothing Then Return False
                        If Not String.IsNullOrWhiteSpace(m.ModelDescription) AndAlso
                           String.Equals(m.ModelDescription, selectedModelKey, StringComparison.OrdinalIgnoreCase) Then Return True
                        If Not String.IsNullOrWhiteSpace(m.Model) AndAlso
                           String.Equals(m.Model, selectedModelKey, StringComparison.OrdinalIgnoreCase) Then Return True
                        Return False
                    End Function)

                executionModelConfig = If(selectedModel, originalExecutionConfig)
            Catch
                executionModelConfig = originalExecutionConfig
            End Try
        End If

        If executionModelConfig IsNot Nothing Then
            ApplyModelConfig(_context, executionModelConfig)
            restoreExecutionConfig = True
        End If
    End Sub

    Private Function ScheduledTaskExecutionSupportsTooling(executionState As InkyState,
                                                           executionModelConfig As ModelConfig,
                                                           executionUseSecondApi As Boolean) As Boolean
        If _apActive Then
            Return executionModelConfig IsNot Nothing AndAlso ModelSupportsTooling(executionModelConfig)
        End If

        If executionState IsNot Nothing Then
            Return CurrentModelSupportsTooling(executionState)
        End If

        If executionUseSecondApi Then
            Return executionModelConfig IsNot Nothing AndAlso ModelSupportsTooling(executionModelConfig)
        End If

        Return False
    End Function

    ''' <summary>
    ''' Copies new tool-generated files from the temp directory into the task's
    ''' persistent workspace. Files that were original input or already in workspace
    ''' are skipped (only genuinely new outputs are persisted).
    ''' </summary>
    Private Sub PersistResultsToWorkspace(taskId As String,
                                           tempDir As String,
                                           resultAttachments As List(Of String),
                                           inputFileNames As HashSet(Of String),
                                           workspaceFileNames As HashSet(Of String))
        Try
            For Each resultPath In resultAttachments
                If Not File.Exists(resultPath) Then Continue For
                Dim fileName = Path.GetFileName(resultPath)

                ' Skip files that are original input attachments (they already live
                ' in the task's attachment directory and should not be duplicated)
                If inputFileNames.Contains(fileName) Then Continue For

                ' Skip files that were loaded from workspace (avoid re-copying unchanged files;
                ' if a tool modified the file in-place, the overwrite in Copy handles it)
                ' We DO want to update workspace files that were modified, so we only skip
                ' files whose name AND size match what was loaded.

                SchedulerStoreWorkspaceFile(taskId, resultPath)
            Next
        Catch ex As System.Exception
            ApDashboardLog($"📅 Error persisting workspace files: {ex.Message}", "warn")
        End Try
    End Sub

    Private Function BuildScheduledTaskExecutionPrompt(task As ScheduledTask,
                                                       inputFileNames As IEnumerable(Of String),
                                                       workspaceFileNames As IEnumerable(Of String)) As StringBuilder
        Dim prompt As New StringBuilder()
        Dim shortTaskId As String = ""

        If task IsNot Nothing Then
            shortTaskId = If(task.Id, "")
            If shortTaskId.Length > 8 Then
                shortTaskId = shortTaskId.Substring(0, 8)
            End If
        End If

        Dim normalizedInputFiles As List(Of String) =
            If(inputFileNames, Enumerable.Empty(Of String)()).
                Where(Function(name) Not String.IsNullOrWhiteSpace(name)).
                Select(Function(name) name.Trim()).
                Distinct(StringComparer.OrdinalIgnoreCase).
                OrderBy(Function(name) name, StringComparer.OrdinalIgnoreCase).
                ToList()

        Dim normalizedWorkspaceFiles As List(Of String) =
            If(workspaceFileNames, Enumerable.Empty(Of String)()).
                Where(Function(name) Not String.IsNullOrWhiteSpace(name)).
                Select(Function(name) name.Trim()).
                Distinct(StringComparer.OrdinalIgnoreCase).
                OrderBy(Function(name) name, StringComparer.OrdinalIgnoreCase).
                ToList()

        prompt.AppendLine("[SCHEDULED TASK EXECUTION]")
        prompt.AppendLine($"You are executing a previously scheduled task (ID: {shortTaskId}...).")
        prompt.AppendLine($"Schedule: {If(task?.ScheduleDescription, "one-time")}")
        prompt.AppendLine("This task is already scheduled — do NOT call manage_scheduled_tasks to create, update, or re-schedule it.")
        prompt.AppendLine("Just perform the work described below and return the result.")

        If task IsNot Nothing AndAlso task.LastExecutedUtc > DateTime.MinValue Then
            prompt.AppendLine($"Last executed: {task.LastExecutedUtc.ToLocalTime():yyyy-MM-dd HH:mm} (local)")
        End If

        prompt.AppendLine()
        prompt.AppendLine("A persistent task-specific workspace is available during this run. Files stored there remain available to future executions of the same task.")

        If normalizedInputFiles.Count > 0 Then
            prompt.AppendLine($"Input files: {String.Join(", ", normalizedInputFiles)}")
        End If

        If normalizedWorkspaceFiles.Count > 0 Then
            prompt.AppendLine($"Existing workspace files: {String.Join(", ", normalizedWorkspaceFiles)}")
        End If

        prompt.AppendLine()
        prompt.AppendLine("[TASK INSTRUCTION]")
        prompt.AppendLine(If(task?.Instruction, ""))
        prompt.AppendLine("[/TASK INSTRUCTION]")

        Return prompt
    End Function

    ''' <summary>
    ''' Sends the result of a scheduled task as a new e-mail to the task's deliverTo addresses.
    ''' Must be called on the UI thread.
    ''' </summary>
    Private Sub SendScheduledTaskResult(task As ScheduledTask,
                                         responseText As String,
                                         resultAttachments As List(Of String),
                                         sourcesHtml As String)
        Dim newMail As MailItem = Nothing
        Try
            newMail = Application.CreateItem(OlItemType.olMailItem)

            Dim safeRecipients As New List(Of String)()

            If Not _apActive Then
                If Not INI_AutoPilotSchedulerLocalChat Then
                    Throw New InvalidOperationException("Local Chat scheduled e-mail delivery is disabled.")
                End If

                Dim currentMailboxAddress As String = GetLocalChatPrimaryMailboxSmtpAddress()

                If String.IsNullOrWhiteSpace(currentMailboxAddress) Then
                    Throw New InvalidOperationException("Could not determine the current mailbox address for Local Chat scheduled e-mail delivery.")
                End If

                safeRecipients.Add(currentMailboxAddress.Trim())
            Else
                If task.DeliverTo IsNot Nothing Then
                    safeRecipients = task.DeliverTo.
                        Where(Function(addr) Not String.IsNullOrWhiteSpace(addr) AndAlso
                                             Not String.IsNullOrWhiteSpace(task.CreatedBy) AndAlso
                                             addr.Trim().Equals(task.CreatedBy.Trim(), StringComparison.OrdinalIgnoreCase)).
                        Select(Function(addr) addr.Trim()).
                        Distinct(StringComparer.OrdinalIgnoreCase).
                        ToList()
                End If

                If safeRecipients.Count = 0 AndAlso Not String.IsNullOrWhiteSpace(task.CreatedBy) Then
                    safeRecipients.Add(task.CreatedBy.Trim())
                End If
            End If

            If safeRecipients.Count = 0 Then
                Throw New InvalidOperationException("Scheduled task has no authorized delivery recipient.")
            End If

            task.DeliverTo = safeRecipients

            newMail.To = String.Join("; ", safeRecipients)
            newMail.Subject = If(Not String.IsNullOrWhiteSpace(task.Subject),
                                 $"{AN6} Scheduled Task: {task.Subject}",
                                 $"{AN6} Scheduled Task Result")
            newMail.BodyFormat = OlBodyFormat.olFormatHTML

            ' Build HTML body
            Dim htmlBody = ConvertResponseToHtml(responseText)

            ' Add schedule info header
            Dim schedNote = $"<div style='font-size:9pt;color:#888888;font-style:italic;margin-bottom:12px;'>" &
                            $"Scheduled task executed at {DateTime.Now:yyyy-MM-dd HH:mm} — " &
                            $"Schedule: {System.Net.WebUtility.HtmlEncode(If(task.ScheduleDescription, "one-time"))}</div>"
            htmlBody = schedNote & htmlBody

            ' Append sources
            If Not String.IsNullOrWhiteSpace(sourcesHtml) Then
                htmlBody &= sourcesHtml
            End If

            htmlBody &= BuildAutoPilotFooter()
            newMail.HTMLBody = htmlBody

            ' Add result attachments
            If resultAttachments IsNot Nothing Then
                For Each attachPath In resultAttachments
                    If File.Exists(attachPath) Then
                        newMail.Attachments.Add(attachPath, OlAttachmentType.olByValue, , Path.GetFileName(attachPath))
                    End If
                Next
            End If

            ' Tag as AutoPilot reply for loop prevention
            Try
                newMail.PropertyAccessor.SetProperty(AP_LoopHeaderProperty, AP_LoopHeaderValue)
            Catch : End Try
            Try : newMail.Categories = AP_CategoryName : Catch : End Try

            ' Use the same sending account as the monitored mailbox
            If _apConfig IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(_apConfig.MonitoredMailbox) Then
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

            newMail.Send()
            Try : MoveLastSentToInkyReplies() : Catch : End Try
            ApDashboardLog($"📅 Result e-mail sent to: {String.Join(", ", task.DeliverTo)}", "info")

        Catch ex As System.Exception
            ApDashboardLog($"📅 ERROR sending scheduled task result: {ex.Message}", "error")
        Finally
            If newMail IsNot Nothing Then Try : Marshal.ReleaseComObject(newMail) : Catch : End Try
        End Try
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  LOCAL AGENT SCHEDULER HELPERS
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Function IsLocalInteractiveScheduledTask(task As ScheduledTask) As Boolean
        If task Is Nothing Then Return False
        Return NormalizeScheduledTaskExecutionMode(task.ExecutionMode).
            Equals(AP_TaskExecutionModeBrowserPrompt, StringComparison.OrdinalIgnoreCase)
    End Function

    Private Function ResolveScheduledTaskExecutionMode(isLocalChatOrigin As Boolean,
                                                       deliverTo As IEnumerable(Of String)) As String
        If Not isLocalChatOrigin Then
            Return AP_TaskExecutionModeEmail
        End If

        Dim recipients As List(Of String) =
            If(deliverTo, Enumerable.Empty(Of String)()).
                Where(Function(addr) Not String.IsNullOrWhiteSpace(addr)).
                Select(Function(addr) addr.Trim()).
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToList()

        If recipients.Count > 0 Then
            Return AP_TaskExecutionModeEmail
        End If

        Return AP_TaskExecutionModeBrowserPrompt
    End Function

    Private Async Function HandleDueLocalInteractiveTaskAsync(task As ScheduledTask, ct As CancellationToken) As Task
        If task Is Nothing Then Return
        If Not IsLocalSchedulerEnabled() Then Return
        If HasPendingScheduledBrowserTask() Then Return
        If _chatAgentActive Then Return
        If Interlocked.CompareExchange(activeJobs, 0, 0) > 0 Then Return

        Dim choice As Integer =
            Await SwitchToUi(Function() PromptForLocalScheduledTaskAction(task))

        If choice = 0 Then choice = 2

        Select Case choice
            Case 1
                If QueueScheduledTaskForLocalBrowserRun(task) Then
                    task.LocalExecutionState = AP_TaskLocalExecutionStateQueued
                    task.SnoozeUntilUtc = DateTime.MinValue
                    task.LastResult = $"Queued for Local Agent browser execution at {DateTime.UtcNow:yyyy-MM-dd HH:mm:ss UTC}"
                    SchedulerUpdateTask(task)
                    ApDashboardLog($"📅 Local Agent task queued: {task.Id.Substring(0, 8)}...", "info")
                Else
                    task.LocalExecutionState = ""
                    task.SnoozeUntilUtc = DateTime.UtcNow.AddMinutes(AP_SchedulerLocalChatSnoozeMinutes)
                    task.LastResult = $"Failed to queue Local Agent execution at {DateTime.UtcNow:yyyy-MM-dd HH:mm:ss UTC} — snoozed for {AP_SchedulerLocalChatSnoozeMinutes} minutes"
                    SchedulerUpdateTask(task)
                    ApDashboardLog($"📅 Local Agent queue failed — snoozed task {task.Id.Substring(0, 8)}...", "warn")
                End If

            Case 2
                task.LocalExecutionState = ""
                task.SnoozeUntilUtc = DateTime.UtcNow.AddMinutes(AP_SchedulerLocalChatSnoozeMinutes)
                task.LastResult = $"Snoozed by user at {DateTime.UtcNow:yyyy-MM-dd HH:mm:ss UTC} for {AP_SchedulerLocalChatSnoozeMinutes} minutes"
                SchedulerUpdateTask(task)
                ApDashboardLog($"📅 Local Agent task snoozed: {task.Id.Substring(0, 8)}...", "step")

            Case 3
                task.LocalExecutionState = ""
                task.SnoozeUntilUtc = DateTime.MinValue
                task.LastExecutedUtc = DateTime.UtcNow
                task.LastResult = $"Skipped by user at {DateTime.UtcNow:yyyy-MM-dd HH:mm:ss UTC}"
                AdvanceTaskSchedule(task)
                SchedulerUpdateTask(task)
                ApDashboardLog($"📅 Local Agent task skipped: {task.Id.Substring(0, 8)}...", "step")
        End Select
    End Function

    Private Function PromptForLocalScheduledTaskAction(task As ScheduledTask) As Integer
        Dim options As New List(Of SelectionItem) From {
            New SelectionItem("Run now — open browser and execute task", 1),
            New SelectionItem($"Snooze {AP_SchedulerLocalChatSnoozeMinutes} minutes", 2),
            New SelectionItem("Skip this occurrence", 3)
        }

        Dim body As New StringBuilder()
        body.AppendLine("A scheduled Local Agent task is due.")
        body.AppendLine()
        body.AppendLine($"Instruction: {Truncate(task.Instruction, 200)}")
        body.AppendLine($"Schedule: {If(task.ScheduleDescription, "one-time")}")
        body.AppendLine($"Due: {task.NextDueUtc.ToLocalTime():yyyy-MM-dd HH:mm} (local)")
        body.AppendLine()
        body.AppendLine("Choose whether the browser may be opened and the task may run now.")

        Return SelectValue(
            options,
            1,
            prompt:=body.ToString().TrimEnd(),
            header:=$"{AN6} Scheduler")
    End Function

    Private Function QueueScheduledTaskForLocalBrowserRun(task As ScheduledTask) As Boolean
        If task Is Nothing Then Return False
        If HasPendingScheduledBrowserTask() Then Return False
        If Not PrepareScheduledTaskForLocalChat(task) Then Return False

        QueuePendingScheduledBrowserTask(
            task.Id,
            BuildScheduledTaskExecutionPrompt(
                task,
                task.AttachmentFiles,
                task.WorkspaceFiles).ToString(),
            AP_SchedulerLocalChatId)

        If Not OpenLocalSchedulerBrowser() Then
            ClearPendingScheduledBrowserTask(task.Id)
            DeactivateScheduledTaskWorkspace(task.Id)
            Return False
        End If

        Return True
    End Function

    Private Function PrepareScheduledTaskForLocalChat(task As ScheduledTask) As Boolean
        Try
            ChatAgentClearFiles()

            If String.IsNullOrWhiteSpace(_chatAgentTempDir) OrElse Not Directory.Exists(_chatAgentTempDir) Then
                _chatAgentTempDir = Path.Combine(Path.GetTempPath(), CA_TempPrefix & Guid.NewGuid().ToString("N"))
                Directory.CreateDirectory(_chatAgentTempDir)
            End If

            Dim preparedFiles As New List(Of AutoPilotAttachmentInfo)()
            Dim inputFileNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            If Not String.IsNullOrWhiteSpace(task.AttachmentDir) Then
                Dim fullAttDir = Path.Combine(Path.GetDirectoryName(GetScheduleFilePath()), task.AttachmentDir)
                If Directory.Exists(fullAttDir) Then
                    For Each filePath In Directory.GetFiles(fullAttDir)
                        Dim att = CopyScheduledTaskFileToChatAgentSession(
                            filePath,
                            Path.GetFileName(filePath),
                            "OK (scheduled input)")
                        If att IsNot Nothing Then
                            preparedFiles.Add(att)
                            inputFileNames.Add(att.OriginalFileName)
                        End If
                    Next
                End If
            End If

            Dim wsDir = GetTaskWorkspaceDir(task.Id)
            If Directory.Exists(wsDir) Then
                For Each filePath In Directory.GetFiles(wsDir)
                    Dim fileName = Path.GetFileName(filePath)
                    Dim desiredName = fileName

                    If inputFileNames.Contains(desiredName) Then
                        desiredName = "ws_" & desiredName
                    End If

                    Dim att = CopyScheduledTaskFileToChatAgentSession(
                        filePath,
                        desiredName,
                        "OK (scheduled workspace)")
                    If att IsNot Nothing Then
                        preparedFiles.Add(att)
                    End If
                Next
            End If

            _chatAgentFiles = preparedFiles
            ResetChatAgentDeliverableTrackingForNewTurn()
            ActivateScheduledTaskWorkspace(task.Id)
            PrepareLocalScheduledChatState(AP_SchedulerLocalChatId)

            Return True
        Catch ex As System.Exception
            DeactivateScheduledTaskWorkspace(task.Id)
            ApDashboardLog($"📅 Failed to prepare Local Agent task files: {ex.Message}", "warn")
            Return False
        End Try
    End Function

    Private Sub PrepareLocalScheduledChatState(chatId As Integer)
        Dim sourceChatId = If(activeChatId = chatId, 1, activeChatId)
        Dim sourceState = LoadInkyState(sourceChatId)
        Dim targetState = LoadInkyState(chatId)

        targetState.History = New List(Of ChatTurn)()
        targetState.LastAssistantText = ""
        targetState.DarkMode = True
        targetState.UseSecondApi = sourceState.UseSecondApi
        targetState.SelectedModelKey = sourceState.SelectedModelKey
        targetState.ToolingEnabled = True
        targetState.AgentModeEnabled = True
        targetState.AgentModelActive = sourceState.AgentModelActive
        targetState.PreAgentModelKey = sourceState.PreAgentModelKey
        targetState.PreAgentUseSecondApi = sourceState.PreAgentUseSecondApi
        targetState.SelectedToolNames = New List(Of String)(If(sourceState.SelectedToolNames, New List(Of String)()))
        targetState.SelectedMainToolNames = New List(Of String)(If(sourceState.SelectedMainToolNames, New List(Of String)()))
        targetState.SelectedAdvancedToolNames = New List(Of String)(If(sourceState.SelectedAdvancedToolNames, New List(Of String)()))

        EnsureScheduledTaskLocalChatToolSelection(
            targetState,
            AP_TaskExecutionModeBrowserPrompt)

        Dim scheduledBrowserTools =
            MergeScheduledTaskExecutionTools(
                GetLocalChatEffectiveSelection(targetState, includeInteractiveM365Tools:=True),
                includeAllAdvancedTools:=True,
                executionMode:=AP_TaskExecutionModeBrowserPrompt)

        targetState.SelectedToolNames =
            scheduledBrowserTools.
                Where(Function(t) t IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(t.ToolName)).
                Select(Function(t) t.ToolName.Trim()).
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToList()

        Try
            targetState.SupportsFileUploads = ComputeSupportsFiles(targetState.UseSecondApi, targetState.SelectedModelKey)
        Catch
            targetState.SupportsFileUploads = False
        End Try

        SaveInkyState(targetState, chatId)

        activeChatId = chatId
        _chatAdvancedToolsEnabled = targetState.AgentModeEnabled
        _selectedToolsForChat = scheduledBrowserTools

        Try
            My.Settings.Inky_LastChat = chatId
            My.Settings.Save()
        Catch
        End Try
    End Sub

    Private Function CopyScheduledTaskFileToChatAgentSession(sourceFilePath As String,
                                                             desiredFileName As String,
                                                             statusMessage As String) As AutoPilotAttachmentInfo
        If String.IsNullOrWhiteSpace(sourceFilePath) OrElse Not File.Exists(sourceFilePath) Then Return Nothing

        Dim destName = desiredFileName
        If String.IsNullOrWhiteSpace(destName) Then
            destName = Path.GetFileName(sourceFilePath)
        End If

        Dim destPath = Path.Combine(_chatAgentTempDir, destName)
        Dim counter = 1

        While File.Exists(destPath)
            destPath = Path.Combine(
                _chatAgentTempDir,
                Path.GetFileNameWithoutExtension(destName) & $"_{counter}" & Path.GetExtension(destName))
            counter += 1
        End While

        File.Copy(sourceFilePath, destPath, overwrite:=True)

        Dim fi = New FileInfo(destPath)
        Return New AutoPilotAttachmentInfo() With {
            .OriginalFileName = Path.GetFileName(destPath),
            .TempFilePath = destPath,
            .SourcePath = sourceFilePath,
            .Extension = fi.Extension.ToLowerInvariant(),
            .SizeBytes = fi.Length,
            .IsOverSizeLimit = False,
            .StatusMessage = statusMessage,
            .CreatedTime = fi.CreationTimeUtc,
            .LastModifiedTime = fi.LastWriteTimeUtc,
            .OutputFiles = New List(Of String)(),
            .IsToolOutput = False
        }
    End Function

    Private Function OpenLocalSchedulerBrowser() As Boolean
        Try
            Dim startInfo As New ProcessStartInfo(AP_SchedulerLocalChatUrl) With {
                .UseShellExecute = True
            }
            Process.Start(startInfo)
            Return True
        Catch ex As System.Exception
            ApDashboardLog($"📅 Failed to open Local Agent browser: {ex.Message}", "warn")
            Return False
        End Try
    End Function

    Friend Sub SchedulerMarkLocalTaskRunning(taskId As String)
        Dim task = SchedulerFindTask(taskId)
        If task Is Nothing Then Return

        task.LocalExecutionState = AP_TaskLocalExecutionStateRunning
        task.SnoozeUntilUtc = DateTime.MinValue
        task.LastResult = $"Local Agent execution started at {DateTime.UtcNow:yyyy-MM-dd HH:mm:ss UTC}"
        SchedulerUpdateTask(task)
    End Sub

    Friend Sub SchedulerPersistLocalBrowserOutputs(taskId As String, resultFiles As IEnumerable(Of String))
        If resultFiles Is Nothing Then Return

        Dim storedAny As Boolean = False

        For Each filePath In resultFiles
            If String.IsNullOrWhiteSpace(filePath) Then Continue For
            If Not File.Exists(filePath) Then Continue For

            If SchedulerStoreWorkspaceFile(taskId, filePath) Then
                storedAny = True
            End If
        Next

        If storedAny Then
            Dim task = SchedulerFindTask(taskId)
            If task IsNot Nothing Then
                SyncWorkspaceFileList(task)
                SchedulerUpdateTask(task)
            End If
        End If
    End Sub

    Friend Sub SchedulerCompleteLocalBrowserTask(taskId As String, resultSummary As String)
        Dim task = SchedulerFindTask(taskId)
        If task Is Nothing Then Return

        task.LocalExecutionState = ""
        task.SnoozeUntilUtc = DateTime.MinValue
        task.LastExecutedUtc = DateTime.UtcNow
        task.LastResult = $"Completed via Local Agent browser at {DateTime.UtcNow:yyyy-MM-dd HH:mm:ss UTC}: {Truncate(If(resultSummary, ""), 160)}"

        AdvanceTaskSchedule(task)
        SchedulerUpdateTask(task)

        RefreshScheduledTaskWorkspaceTracking(taskId)
        DeactivateScheduledTaskWorkspace(taskId)
    End Sub

    Friend Sub SchedulerFailLocalBrowserTask(taskId As String,
                                             failureSummary As String,
                                             Optional snoozeMinutes As Integer = AP_SchedulerLocalChatSnoozeMinutes)
        Dim task = SchedulerFindTask(taskId)
        If task Is Nothing Then Return

        task.LocalExecutionState = ""
        task.SnoozeUntilUtc = DateTime.UtcNow.AddMinutes(snoozeMinutes)
        task.LastResult = $"Local Agent execution failed at {DateTime.UtcNow:yyyy-MM-dd HH:mm:ss UTC}: {Truncate(If(failureSummary, ""), 160)}"
        SchedulerUpdateTask(task)

        RefreshScheduledTaskWorkspaceTracking(taskId)
        DeactivateScheduledTaskWorkspace(taskId)
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  RECURRENCE — advance nextDueUtc after execution
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Advances a task's NextDueUtc based on its recurrence rule.
    ''' If the task has no recurrence or has reached its limit, marks it as completed.
    ''' </summary>
    Private Sub AdvanceTaskSchedule(task As ScheduledTask)
        If task Is Nothing Then Return

        task.FailureRetryCount = 0

        ' Decrement remaining occurrences if count-limited
        If task.RemainingOccurrences > 0 Then
            task.RemainingOccurrences -= 1
            If task.RemainingOccurrences <= 0 Then
                task.Status = "completed"
                task.NextDueUtc = DateTime.MaxValue
                Return
            End If
        End If

        ' No RRULE means one-shot
        If String.IsNullOrWhiteSpace(task.Rrule) Then
            task.Status = "completed"
            task.NextDueUtc = DateTime.MaxValue
            Return
        End If

        ' Parse RRULE and compute next occurrence
        Dim nextUtc = ComputeNextOccurrence(task.Rrule, task.NextDueUtc, task.TimeOfDayLocal)
        If nextUtc Is Nothing OrElse nextUtc.Value >= task.EndDateUtc Then
            task.Status = "completed"
            task.NextDueUtc = DateTime.MaxValue
            Return
        End If

        task.NextDueUtc = nextUtc.Value
    End Sub

    ''' <summary>
    ''' Lightweight RRULE evaluator. Supports:
    '''   FREQ=DAILY;INTERVAL=n
    '''   FREQ=WEEKLY;INTERVAL=n;BYDAY=MO,TU,...
    '''   FREQ=MONTHLY;INTERVAL=n;BYMONTHDAY=d
    '''   FREQ=MONTHLY;INTERVAL=n;BYDAY=1SU,2MO,...  (nth weekday)
    ''' </summary>
    Private Shared Function ComputeNextOccurrence(rrule As String, currentDueUtc As DateTime, timeOfDayLocal As String) As DateTime?
        Try
            Dim parts As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            For Each part In rrule.Split(";"c)
                Dim kv = part.Split("="c)
                If kv.Length = 2 Then parts(kv(0).Trim()) = kv(1).Trim()
            Next

            Dim freq As String = ""
            parts.TryGetValue("FREQ", freq)

            Dim interval As Integer = 1
            If parts.ContainsKey("INTERVAL") Then Integer.TryParse(parts("INTERVAL"), interval)
            If interval < 1 Then interval = 1

            ' Resolve time-of-day in local timezone
            Dim currentLocal = currentDueUtc.ToLocalTime()
            Dim tod As TimeSpan = currentLocal.TimeOfDay
            If Not String.IsNullOrWhiteSpace(timeOfDayLocal) Then
                Dim parsed As DateTime
                If DateTime.TryParse(timeOfDayLocal, parsed) Then
                    tod = parsed.TimeOfDay
                End If
            End If

            Dim nextLocal As DateTime

            Select Case freq.ToUpperInvariant()
                Case "MINUTELY"
                    nextLocal = currentLocal.AddMinutes(interval)

                Case "HOURLY"
                    nextLocal = currentLocal.AddHours(interval)

                Case "DAILY"
                    nextLocal = currentLocal.Date.AddDays(interval).Add(tod)

                Case "WEEKLY"
                    Dim byDay As String = ""
                    parts.TryGetValue("BYDAY", byDay)

                    If String.IsNullOrWhiteSpace(byDay) Then
                        ' Simple weekly recurrence
                        nextLocal = currentLocal.Date.AddDays(7 * interval).Add(tod)
                    Else
                        ' Find next matching day of week
                        Dim targetDays = ParseByDay(byDay)
                        nextLocal = FindNextMatchingDay(currentLocal.Date, targetDays, interval, tod)
                    End If

                Case "MONTHLY"
                    Dim byMonthDay As String = ""
                    Dim byDayMonthly As String = ""
                    parts.TryGetValue("BYMONTHDAY", byMonthDay)
                    parts.TryGetValue("BYDAY", byDayMonthly)

                    If Not String.IsNullOrWhiteSpace(byMonthDay) Then
                        Dim day As Integer
                        If Integer.TryParse(byMonthDay, day) Then
                            Dim nextMonth = currentLocal.Date.AddMonths(interval)
                            Dim daysInMonth = DateTime.DaysInMonth(nextMonth.Year, nextMonth.Month)
                            day = Math.Min(day, daysInMonth)
                            nextLocal = New DateTime(nextMonth.Year, nextMonth.Month, day).Add(tod)
                        Else
                            Return Nothing
                        End If
                    ElseIf Not String.IsNullOrWhiteSpace(byDayMonthly) Then
                        ' e.g. "1SU" = first Sunday, "2MO" = second Monday
                        nextLocal = FindNthWeekdayInMonth(currentLocal.Date.AddMonths(interval), byDayMonthly, tod)
                    Else
                        nextLocal = currentLocal.Date.AddMonths(interval).Add(tod)
                    End If

                Case Else
                    Return Nothing
            End Select

            ' If the computed next time is still in the past (e.g. catch-up),
            ' keep advancing until we reach the future
            Dim nowLocal = DateTime.Now
            Dim safetyCounter = 0
            While nextLocal <= nowLocal AndAlso safetyCounter < 1000
                Select Case freq.ToUpperInvariant()
                    Case "MINUTELY" : nextLocal = nextLocal.AddMinutes(interval)
                    Case "HOURLY" : nextLocal = nextLocal.AddHours(interval)
                    Case "DAILY" : nextLocal = nextLocal.AddDays(interval)
                    Case "WEEKLY" : nextLocal = nextLocal.AddDays(7 * interval)
                    Case "MONTHLY" : nextLocal = nextLocal.AddMonths(interval)
                    Case Else : Return Nothing
                End Select
                safetyCounter += 1
            End While

            Return nextLocal.ToUniversalTime()

        Catch ex As System.Exception
            Debug.WriteLine($"[Scheduler] RRULE parse error: {ex.Message}")
            Return Nothing
        End Try
    End Function

    ''' <summary>Parses BYDAY values like "MO,TU,WE" into DayOfWeek list.</summary>
    Private Shared Function ParseByDay(byDay As String) As List(Of DayOfWeek)
        Dim result As New List(Of DayOfWeek)()
        For Each token In byDay.Split(","c)
            Dim t = token.Trim().ToUpperInvariant()
            ' Strip numeric prefix (e.g. "1SU" → "SU") for simple day matching
            If t.Length > 2 Then t = t.Substring(t.Length - 2)
            Select Case t
                Case "MO" : result.Add(DayOfWeek.Monday)
                Case "TU" : result.Add(DayOfWeek.Tuesday)
                Case "WE" : result.Add(DayOfWeek.Wednesday)
                Case "TH" : result.Add(DayOfWeek.Thursday)
                Case "FR" : result.Add(DayOfWeek.Friday)
                Case "SA" : result.Add(DayOfWeek.Saturday)
                Case "SU" : result.Add(DayOfWeek.Sunday)
            End Select
        Next
        Return result
    End Function

    ''' <summary>Finds the next date matching one of the target days, with week interval.</summary>
    Private Shared Function FindNextMatchingDay(fromDate As DateTime, targetDays As List(Of DayOfWeek),
                                                 weekInterval As Integer, tod As TimeSpan) As DateTime
        ' Start from next day after fromDate
        Dim candidate = fromDate.AddDays(1)
        Dim endSearch = fromDate.AddDays(7 * weekInterval + 7)
        While candidate <= endSearch
            If targetDays.Contains(candidate.DayOfWeek) Then
                Return candidate.Add(tod)
            End If
            candidate = candidate.AddDays(1)
        End While
        ' Fallback: just add the interval in weeks
        Return fromDate.AddDays(7 * weekInterval).Add(tod)
    End Function

    ''' <summary>
    ''' Finds the nth weekday in a given month. E.g. "1SU" = first Sunday.
    ''' </summary>
    Private Shared Function FindNthWeekdayInMonth(monthStart As DateTime, byDaySpec As String, tod As TimeSpan) As DateTime
        Dim spec = byDaySpec.Trim().ToUpperInvariant()
        Dim nth As Integer = 1
        Dim dayStr As String = spec

        ' Parse leading digit(s)
        If spec.Length > 2 AndAlso Char.IsDigit(spec(0)) Then
            nth = Integer.Parse(spec.Substring(0, spec.Length - 2))
            dayStr = spec.Substring(spec.Length - 2)
        End If

        Dim targetDay As DayOfWeek
        Select Case dayStr
            Case "MO" : targetDay = DayOfWeek.Monday
            Case "TU" : targetDay = DayOfWeek.Tuesday
            Case "WE" : targetDay = DayOfWeek.Wednesday
            Case "TH" : targetDay = DayOfWeek.Thursday
            Case "FR" : targetDay = DayOfWeek.Friday
            Case "SA" : targetDay = DayOfWeek.Saturday
            Case "SU" : targetDay = DayOfWeek.Sunday
            Case Else : Return monthStart.Add(tod)
        End Select

        Dim d = New DateTime(monthStart.Year, monthStart.Month, 1)
        Dim count = 0
        While d.Month = monthStart.Month
            If d.DayOfWeek = targetDay Then
                count += 1
                If count = nth Then Return d.Add(tod)
            End If
            d = d.AddDays(1)
        End While

        ' Fallback to last occurrence if nth exceeds available
        Return New DateTime(monthStart.Year, monthStart.Month, DateTime.DaysInMonth(monthStart.Year, monthStart.Month)).Add(tod)
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  HELPER
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Truncates a string to the given max length with ellipsis.</summary>
    Private Shared Function Truncate(s As String, maxLen As Integer) As String
        If String.IsNullOrEmpty(s) Then Return ""
        If s.Length <= maxLen Then Return s
        Return s.Substring(0, maxLen - 3) & "..."
    End Function

    ''' <summary>
    ''' Formats the task list into a human-readable text summary for the LLM to return.
    ''' </summary>
    Friend Function FormatTaskListForDisplay(tasks As List(Of ScheduledTask)) As String
        If tasks Is Nothing OrElse tasks.Count = 0 Then Return "No scheduled tasks found."

        Dim sb As New StringBuilder()
        sb.AppendLine($"=== Scheduled Tasks ({tasks.Count}) ===")
        sb.AppendLine()

        For Each t In tasks.OrderBy(Function(x) x.NextDueUtc)
            Dim shortId = If(t.Id.Length > 8, t.Id.Substring(0, 8), t.Id)
            sb.AppendLine($"ID: {shortId}...")
            sb.AppendLine($"  Instruction: {Truncate(t.Instruction, 120)}")
            sb.AppendLine($"  Schedule: {If(t.ScheduleDescription, "(not set)")}")
            sb.AppendLine($"  Deliver to: {String.Join(", ", t.DeliverTo)}")
            sb.AppendLine($"  Status: {t.Status}")
            If t.NextDueUtc < DateTime.MaxValue Then
                sb.AppendLine($"  Next due: {t.NextDueUtc.ToLocalTime():yyyy-MM-dd HH:mm} (local)")
            End If
            If t.LastExecutedUtc > DateTime.MinValue Then
                sb.AppendLine($"  Last executed: {t.LastExecutedUtc.ToLocalTime():yyyy-MM-dd HH:mm} (local)")
            End If
            If t.AttachmentFiles IsNot Nothing AndAlso t.AttachmentFiles.Count > 0 Then
                sb.AppendLine($"  Input files: {String.Join(", ", t.AttachmentFiles)}")
            End If
            If t.WorkspaceFiles IsNot Nothing AndAlso t.WorkspaceFiles.Count > 0 Then
                sb.AppendLine($"  Workspace files: {String.Join(", ", t.WorkspaceFiles)}")
            End If
            If Not String.IsNullOrWhiteSpace(t.LastResult) Then
                sb.AppendLine($"  Last result: {Truncate(t.LastResult, 100)}")
            End If
            If Not String.IsNullOrWhiteSpace(t.Rrule) Then
                sb.AppendLine($"  RRULE: {t.Rrule}")
            End If
            If t.RemainingOccurrences > 0 Then
                sb.AppendLine($"  Remaining: {t.RemainingOccurrences} occurrence(s)")
            End If
            sb.AppendLine()
        Next

        Return sb.ToString().TrimEnd()
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  SCHEDULER DASHBOARD — standalone form with DataGridView
    ' ═══════════════════════════════════════════════════════════════════════════

    Private _apSchedulerDashboard As Form = Nothing

    ''' <summary>Shows the scheduler dashboard. Creates it if not yet open.</summary>
    Friend Sub ShowSchedulerDashboard()
        If _apSchedulerDashboard IsNot Nothing AndAlso Not _apSchedulerDashboard.IsDisposed Then
            RefreshSchedulerDashboard()
            _apSchedulerDashboard.Show()
            _apSchedulerDashboard.BringToFront()
            Return
        End If

        _apSchedulerDashboard = CreateSchedulerDashboardForm()
        RefreshSchedulerDashboard()
        _apSchedulerDashboard.Show()
    End Sub

    ''' <summary>Creates the scheduler dashboard form with a DataGridView and action buttons.</summary>
    Private Function CreateSchedulerDashboardForm() As Form
        Dim frm As New Form() With {
            .Text = $"{AN6} AutoPilot — Scheduled Tasks",
            .Width = 1200,
            .Height = 450,
            .StartPosition = FormStartPosition.CenterScreen,
            .FormBorderStyle = FormBorderStyle.Sizable,
            .MinimumSize = New Drawing.Size(700, 300),
            .TopMost = False,
            .MaximizeBox = True,
            .MinimizeBox = True,
            .ShowInTaskbar = True,
            .AutoScaleMode = AutoScaleMode.Font,
            .Font = New Drawing.Font("Segoe UI", 9.0F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point)
        }

        Try
            frm.Icon = Drawing.Icon.FromHandle(
                (New Drawing.Bitmap(GetLogoBitmap(LogoType.Standard))).GetHicon())
        Catch
        End Try

        Dim mainPanel As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill, .ColumnCount = 1, .RowCount = 2, .Padding = New Padding(8)
        }
        mainPanel.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        mainPanel.RowStyles.Add(New RowStyle(SizeType.AutoSize))

        Dim dgv As New DataGridView() With {
            .Dock = DockStyle.Fill,
            .Name = "dgvTasks",
            .ReadOnly = True,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .AllowUserToResizeRows = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .MultiSelect = False,
            .RowHeadersVisible = False,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            .BackgroundColor = Drawing.SystemColors.Window,
            .BorderStyle = BorderStyle.None,
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
            .DefaultCellStyle = New DataGridViewCellStyle() With {
                .Font = New Drawing.Font("Segoe UI", 8.5F),
                .WrapMode = DataGridViewTriState.False
            },
            .AlternatingRowsDefaultCellStyle = New DataGridViewCellStyle() With {
                .BackColor = Drawing.Color.FromArgb(245, 245, 250)
            }
        }

        dgv.Columns.Add("colId", "ID")
        dgv.Columns.Add("colInstruction", "Instruction")
        dgv.Columns.Add("colSchedule", "Schedule")
        dgv.Columns.Add("colStatus", "Status")
        dgv.Columns.Add("colExecutionMode", "Execution Mode")
        dgv.Columns.Add("colLocalState", "Local State")
        dgv.Columns.Add("colSnoozeUntil", "Snoozed Until")
        dgv.Columns.Add("colNextDue", "Next Due")
        dgv.Columns.Add("colLastExec", "Last Executed")
        dgv.Columns.Add("colDeliverTo", "Deliver To")
        dgv.Columns.Add("colLastResult", "Last Result")

        dgv.Columns("colId").FillWeight = 7
        dgv.Columns("colInstruction").FillWeight = 20
        dgv.Columns("colSchedule").FillWeight = 11
        dgv.Columns("colStatus").FillWeight = 7
        dgv.Columns("colExecutionMode").FillWeight = 10
        dgv.Columns("colLocalState").FillWeight = 8
        dgv.Columns("colSnoozeUntil").FillWeight = 11
        dgv.Columns("colNextDue").FillWeight = 10
        dgv.Columns("colLastExec").FillWeight = 10
        dgv.Columns("colDeliverTo").FillWeight = 11
        dgv.Columns("colLastResult").FillWeight = 15

        mainPanel.Controls.Add(dgv, 0, 0)

        ' Button panel
        Dim buttonPanel As New FlowLayoutPanel() With {
            .Dock = DockStyle.Fill, .FlowDirection = FlowDirection.RightToLeft,
            .AutoSize = True, .Padding = New Padding(0, 4, 0, 0)
        }

        Dim btnRefresh As New Button() With {
            .Text = "Refresh", .AutoSize = True, .Padding = New Padding(10, 5, 10, 5)
        }
        AddHandler btnRefresh.Click, Sub(s, e) RefreshSchedulerDashboard()

        Dim btnDelete As New Button() With {
            .Text = "Delete Selected", .AutoSize = True, .Padding = New Padding(10, 5, 10, 5)
        }
        AddHandler btnDelete.Click, Sub(s, e) DeleteSelectedSchedulerTask(dgv)

        Dim btnPause As New Button() With {
            .Text = "Pause/Resume", .AutoSize = True, .Padding = New Padding(10, 5, 10, 5)
        }
        AddHandler btnPause.Click, Sub(s, e) TogglePauseSelectedTask(dgv)

        Dim btnClose As New Button() With {
            .Text = "Close", .AutoSize = True, .Padding = New Padding(10, 5, 10, 5)
        }
        AddHandler btnClose.Click, Sub(s, e) frm.Hide()

        Dim btnOpenFile As New Button() With {
            .Text = "Open JSON", .AutoSize = True, .Padding = New Padding(10, 5, 10, 5)
        }
        AddHandler btnOpenFile.Click, Sub(s, e)
                                          Try
                                              Dim schedPath = GetScheduleFilePath()
                                              If Not File.Exists(schedPath) Then
                                                  ShowCustomMessageBox("Schedule file does not exist yet.")
                                                  Return
                                              End If
                                              ShowTextFileEditor(schedPath, $"{AN6} AutoPilot — Schedule File:", ForceJson:=True, _context:=_context)
                                              RefreshSchedulerDashboard()
                                          Catch ex As System.Exception
                                              ShowCustomMessageBox($"Cannot open file: {ex.Message}")
                                          End Try
                                      End Sub

        buttonPanel.Controls.Add(btnClose)
        buttonPanel.Controls.Add(btnRefresh)
        buttonPanel.Controls.Add(btnDelete)
        buttonPanel.Controls.Add(btnPause)
        buttonPanel.Controls.Add(btnOpenFile)
        mainPanel.Controls.Add(buttonPanel, 0, 1)

        frm.Controls.Add(mainPanel)

        ' Hide instead of closing
        AddHandler frm.FormClosing, Sub(s As Object, e As FormClosingEventArgs)
                                        If e.CloseReason = CloseReason.UserClosing Then
                                            e.Cancel = True
                                            frm.Hide()
                                        End If
                                    End Sub

        Return frm
    End Function

    ''' <summary>Refreshes the scheduler dashboard DataGridView with current task data.</summary>
    Friend Sub RefreshSchedulerDashboard()
        If _apSchedulerDashboard Is Nothing OrElse _apSchedulerDashboard.IsDisposed Then Return

        Dim dgv = _apSchedulerDashboard.Controls.Find("dgvTasks", True).
                    OfType(Of DataGridView)().FirstOrDefault()
        If dgv Is Nothing Then Return

        If dgv.InvokeRequired Then
            dgv.BeginInvoke(New MethodInvoker(Sub() RefreshSchedulerDashboardCore(dgv)))
        Else
            RefreshSchedulerDashboardCore(dgv)
        End If
    End Sub

    ''' <summary>Core logic for populating the scheduler DataGridView. Must run on the UI thread.</summary>
    Private Sub RefreshSchedulerDashboardCore(dgv As DataGridView)
        dgv.Rows.Clear()
        Dim tasks = ReadScheduleFile().Tasks.OrderBy(Function(t) t.NextDueUtc).ToList()

        For Each t In tasks
            Dim shortId = If(t.Id.Length > 8, t.Id.Substring(0, 8), t.Id)
            Dim executionMode = NormalizeScheduledTaskExecutionMode(t.ExecutionMode)
            Dim localState = If(String.IsNullOrWhiteSpace(t.LocalExecutionState), "—", t.LocalExecutionState)
            Dim snoozeUntil = If(t.SnoozeUntilUtc > DateTime.MinValue AndAlso t.SnoozeUntilUtc < DateTime.MaxValue,
                                 t.SnoozeUntilUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm"), "—")
            Dim nextDue = If(t.NextDueUtc < DateTime.MaxValue,
                             t.NextDueUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm"), "—")
            Dim lastExec = If(t.LastExecutedUtc > DateTime.MinValue,
                              t.LastExecutedUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm"), "—")
            Dim deliver = String.Join(", ", t.DeliverTo)
            Dim lastResult = If(t.LastResult, "—")
            If lastResult.Length > 80 Then lastResult = lastResult.Substring(0, 77) & "..."

            Dim rowIdx = dgv.Rows.Add(
                shortId,
                Truncate(t.Instruction, 80),
                If(t.ScheduleDescription, "one-time"),
                t.Status,
                executionMode,
                localState,
                snoozeUntil,
                nextDue,
                lastExec,
                deliver,
                lastResult)

            ' Color-code status
            Dim row = dgv.Rows(rowIdx)
            Select Case t.Status.ToLowerInvariant()
                Case "active"
                    row.Cells("colStatus").Style.ForeColor = Drawing.Color.DarkGreen
                Case "paused"
                    row.Cells("colStatus").Style.ForeColor = Drawing.Color.DarkOrange
                Case "completed"
                    row.Cells("colStatus").Style.ForeColor = Drawing.Color.Gray
                Case "failed"
                    row.Cells("colStatus").Style.ForeColor = Drawing.Color.DarkRed
            End Select

            ' Store full task ID in Tag for delete/pause operations
            row.Tag = t.Id
        Next
    End Sub

    ''' <summary>Deletes the selected task from the scheduler dashboard.</summary>
    Private Sub DeleteSelectedSchedulerTask(dgv As DataGridView)
        If dgv.SelectedRows.Count = 0 Then Return
        Dim taskId = TryCast(dgv.SelectedRows(0).Tag, String)
        If String.IsNullOrWhiteSpace(taskId) Then Return

        Dim confirm = ShowCustomYesNoBox(
            $"Delete scheduled task {taskId.Substring(0, Math.Min(8, taskId.Length))}...?",
            "Delete", "Cancel",
            header:=$"{AN6} Scheduler")
        If confirm <> 1 Then Return

        If SchedulerDeleteTask(taskId) Then
            RefreshSchedulerDashboard()
        Else
            ShowCustomMessageBox("Failed to delete task.")
        End If
    End Sub

    ''' <summary>Toggles the status of the selected task between active and paused.</summary>
    Private Sub TogglePauseSelectedTask(dgv As DataGridView)
        If dgv.SelectedRows.Count = 0 Then Return
        Dim taskId = TryCast(dgv.SelectedRows(0).Tag, String)
        If String.IsNullOrWhiteSpace(taskId) Then Return

        Dim task = SchedulerFindTask(taskId)
        If task Is Nothing Then Return

        Select Case task.Status.ToLowerInvariant()
            Case "active" : task.Status = "paused"
            Case "paused" : task.Status = "active"
            Case Else
                ShowCustomMessageBox($"Task status is '{task.Status}' — cannot pause/resume.")
                Return
        End Select

        If SchedulerUpdateTask(task) Then
            RefreshSchedulerDashboard()
            ApDashboardLog($"📅 Task {taskId.Substring(0, 8)}... status changed to: {task.Status}", "info")
        End If
    End Sub

    ''' <summary>Disposes the scheduler dashboard form.</summary>
    Friend Sub CloseSchedulerDashboard()
        Try
            If _apSchedulerDashboard IsNot Nothing AndAlso Not _apSchedulerDashboard.IsDisposed Then
                _apSchedulerDashboard.Close()
                _apSchedulerDashboard.Dispose()
            End If
        Catch
        End Try
        _apSchedulerDashboard = Nothing
    End Sub

End Class
