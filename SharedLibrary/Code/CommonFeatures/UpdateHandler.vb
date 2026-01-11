' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: UpdateHandler.vb
' Purpose: Centralizes update checking and installation for VSTO add-ins, supporting both
'          ClickOnce network deployments and local installer paths. Provides user prompting,
'          a lightweight update splash UI, and a bounded update log with trimming.
'
' Architecture:
'  - UI Bridge: Uses `MainControl`/`HostHandle` to marshal dialogs onto the UI thread and
'    to bring relevant windows to the foreground.
'  - Update Modes:
'     - Network-deployed: ClickOnce `ApplicationDeployment` update check (sync and async).
'     - Local-deployed: Constructs a `.vsto` path under the configured update folder.
'  - Retry/Backoff: Tracks per-add-in daily retry state in `My.Settings` and may prompt the user
'    to pause update checks after repeated failures.
'  - Installer: Locates `VSTOInstaller.exe`, attempts a silent install first, then falls back
'    to an interactive install that is brought to the foreground.
'  - Logging: Writes to `%AppData%\{AN2}\updater.log` and trims by size/line count.
'  - Configuration Updates: Optionally triggers INI configuration refresh via
'    `SharedMethods.CheckForIniUpdates`.
' =============================================================================


Option Strict On
Option Explicit On

Imports System.Deployment.Application
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

Namespace SharedLibrary
    ''' <summary>
    ''' Provides update checking and installation for VSTO add-ins, including ClickOnce network deployments
    ''' and local update paths. Includes UI prompts, retry tracking, and bounded logging.
    ''' </summary>
    Public Class UpdateHandler

        ''' <summary>
        ''' Controls whether a "Checking updates …" splash is displayed for async network checks.
        ''' </summary>
        Private Const ShowCheckingSplash As Boolean = False

        ''' <summary>
        ''' Maximum number of update-check failures to silently tolerate per day before prompting.
        ''' </summary>
        Private Const MaxDailyUpdateRetries As Integer = 5

        ''' <summary>Maximum log file size in bytes before trimming is initiated.</summary>
        Private Const LogMaxBytes As Integer = 120000      ' ~120 KB

        ''' <summary>Maximum number of lines allowed in the update log before trimming.</summary>
        Private Const LogMaxLines As Integer = 2000

        ''' <summary>Number of last lines to retain when trimming the update log.</summary>
        Private Const LogKeepLines As Integer = 1500       ' retain last 1.5k lines when trimming

        ''' <summary>
        ''' UI control used to marshal calls back onto the UI thread when necessary.
        ''' </summary>
        Public Shared MainControl As System.Windows.Forms.Control

        ''' <summary>
        ''' Host window handle used to request foreground focus for dialogs.
        ''' </summary>
        Public Shared HostHandle As IntPtr

        ''' <summary>
        ''' Splash screen instance shown while checking/installing updates.
        ''' </summary>
        Private Shared _splash As SplashScreen

        ''' <summary>
        ''' Win32 API to adjust z-order and visibility to keep specific windows on top.
        ''' </summary>
        <Runtime.InteropServices.DllImport("user32.dll")>
        Private Shared Function SetWindowPos(hWnd As IntPtr, hWndInsertAfter As IntPtr,
                                         X As Integer, Y As Integer, cx As Integer, cy As Integer,
                                         uFlags As UInteger) As Boolean
        End Function

        ''' <summary>Win32 z-order constant for topmost windows.</summary>
        Private Shared ReadOnly HWND_TOPMOST As IntPtr = New IntPtr(-1)

        ''' <summary>Win32 z-order constant for non-topmost windows.</summary>
        Private Shared ReadOnly HWND_NOTOPMOST As IntPtr = New IntPtr(-2)

        ''' <summary>Do not change window size when calling `SetWindowPos`.</summary>
        Private Const SWP_NOSIZE As UInteger = &H1UI

        ''' <summary>Do not change window position when calling `SetWindowPos`.</summary>
        Private Const SWP_NOMOVE As UInteger = &H2UI

        ''' <summary>Show the window when calling `SetWindowPos`.</summary>
        Private Const SWP_SHOWWINDOW As UInteger = &H40UI

        ''' <summary>
        ''' Shows (or updates) the splash window with the specified message, marshaling to the UI thread when needed.
        ''' </summary>
        Private Shared Sub ShowUpdatingSplash(message As String)
            Try
                If MainControl IsNot Nothing AndAlso MainControl.InvokeRequired Then
                    MainControl.Invoke(Sub() ShowUpdatingSplashCore(message))
                Else
                    ShowUpdatingSplashCore(message)
                End If
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' Core splash display/update logic. Keeps the window top-most and attempts to bring it to the foreground.
        ''' </summary>
        Private Shared Sub ShowUpdatingSplashCore(message As String)
            Try
                If _splash Is Nothing OrElse _splash.IsDisposed Then
                    _splash = New SplashScreen()
                    Try
                        _splash.TopMost = True
                        _splash.ShowInTaskbar = False
                    Catch
                    End Try
                End If
                _splash.UpdateMessage(message)
                _splash.Show()
                _splash.BringToFront()
                _splash.Activate()
                Try
                    NativeMethods.SetForegroundWindow(_splash.Handle)
                Catch
                End Try
                Try
                    SetWindowPos(_splash.Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)
                Catch
                End Try
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' Closes and disposes the splash window, marshaling to its UI thread when necessary.
        ''' </summary>
        Private Shared Sub CloseUpdatingSplash()
            Try
                If _splash Is Nothing Then Return
                Dim closer As Action =
                    Sub()
                        Try
                            Try
                                SetWindowPos(_splash.Handle, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)
                            Catch
                            End Try
                            Try
                                _splash.TopMost = False
                            Catch
                            End Try
                            _splash.Close()
                        Finally
                            _splash = Nothing
                        End Try
                    End Sub
                If _splash.InvokeRequired Then
                    _splash.Invoke(closer)
                Else
                    closer()
                End If
            Catch
                _splash = Nothing
            End Try
        End Sub


        ' LOGGING with trimming
        ''' <summary>
        ''' Appends an entry to the updater log under `%AppData%\{AN2}\updater.log` and triggers trimming when needed.
        ''' </summary>
        ''' <param name="message">Log message to append.</param>
        ''' <param name="ex">Optional exception details to include.</param>
        Public Shared Sub WriteUpdateLog(message As String, Optional ex As Exception = Nothing)
            Try
                Dim logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), SharedMethods.AN2)
                If Not Directory.Exists(logDir) Then Directory.CreateDirectory(logDir)
                Dim logPath = Path.Combine(logDir, "updater.log")

                ' Trim before writing if too large
                Try
                    If File.Exists(logPath) Then
                        Dim fi = New FileInfo(logPath)
                        If fi.Length > LogMaxBytes Then
                            TrimLogFile(logPath)
                        End If
                    End If
                Catch
                End Try

                Dim sb As New StringBuilder()
                sb.AppendFormat("[{0:yyyy-MM-dd HH:mm:ss}] {1}", Date.Now, message)
                If ex IsNot Nothing Then
                    sb.AppendLine()
                    sb.AppendFormat("  Exception: {0}", ex.GetType().FullName)
                    sb.AppendLine()
                    sb.AppendFormat("  Message: {0}", ex.Message)
                    If ex.HResult <> 0 Then
                        sb.AppendLine()
                        sb.AppendFormat("  HResult: 0x{0:X8}", ex.HResult)
                    End If
                    If ex.InnerException IsNot Nothing Then
                        sb.AppendLine()
                        sb.AppendFormat("  Inner: {0}: {1}", ex.InnerException.GetType().FullName, ex.InnerException.Message)
                    End If
                End If
                File.AppendAllText(logPath, sb.ToString() & Environment.NewLine, Encoding.UTF8)

                ' Secondary safeguard by line count (after append)
                Try
                    Dim fi2 = New FileInfo(logPath)
                    If fi2.Length > LogMaxBytes * 1.2 Then
                        TrimLogFile(logPath)
                    End If
                Catch
                End Try
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' Trims the given log file to retain only the most recent lines, bounded by `LogKeepLines`.
        ''' </summary>
        ''' <param name="path">Full path of the log file to trim.</param>
        Private Shared Sub TrimLogFile(path As String)
            Try
                Dim lines = File.ReadAllLines(path, Encoding.UTF8)
                If lines.Length > LogMaxLines Then
                    lines = lines.Skip(Math.Max(0, lines.Length - LogKeepLines)).ToArray()
                ElseIf New FileInfo(path).Length > LogMaxBytes Then
                    ' Fallback if line count small but file large (e.g., huge lines) – trim by tail chars
                    lines = lines.Skip(Math.Max(0, lines.Length - LogKeepLines)).ToArray()
                End If
                File.WriteAllLines(path, lines, Encoding.UTF8)
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' Shows a Yes/No prompt on the UI thread (when required) and returns the underlying dialog result.
        ''' </summary>
        ''' <param name="prompt">Prompt body text.</param>
        ''' <param name="caption">Dialog caption.</param>
        Private Shared Function UIInvokePrompt(prompt As String, caption As String) As Integer
            NativeMethods.SetForegroundWindow(HostHandle)
            If MainControl IsNot Nothing AndAlso MainControl.InvokeRequired Then
                Dim result As Integer = 0
                MainControl.Invoke(Sub() result = SharedMethods.ShowCustomYesNoBox(prompt, "Yes", "No", caption))
                Return result
            Else
                Return SharedMethods.ShowCustomYesNoBox(prompt, "Yes", "No", caption)
            End If
        End Function

        ''' <summary>
        ''' Shows a message box on the UI thread (when required).
        ''' </summary>
        ''' <param name="msg">Message body.</param>
        ''' <param name="caption">Dialog caption.</param>
        Private Shared Sub UIInvokeMessage(msg As String, caption As String)
            NativeMethods.SetForegroundWindow(HostHandle)
            If MainControl IsNot Nothing AndAlso MainControl.InvokeRequired Then
                MainControl.Invoke(Sub() SharedMethods.ShowCustomMessageBox(msg, caption))
            Else
                SharedMethods.ShowCustomMessageBox(msg, caption)
            End If
        End Sub

        ''' <summary>
        ''' Returns whether interactive UI can be shown in the current environment.
        ''' </summary>
        Private Shared Function CanShowInteractiveUi() As Boolean
            Return Environment.UserInteractive
        End Function

        ''' <summary>
        ''' Returns the .NET SecurityZone name for a URL, or "Unknown" if it cannot be determined.
        ''' </summary>
        ''' <param name="url">URL to evaluate.</param>
        Private Shared Function GetUrlZoneName(url As String) As String
            Try
                Dim z = System.Security.Policy.Zone.CreateFromUrl(url)
                Return z.SecurityZone.ToString()
            Catch
                Return "Unknown"
            End Try
        End Function

        ''' <summary>
        ''' Determines whether an exception chain indicates a ClickOnce trust-not-granted scenario.
        ''' </summary>
        ''' <param name="ex">Exception to examine.</param>
        Private Shared Function IsTrustNotGranted(ex As Exception) As Boolean
            If ex Is Nothing Then Return False
            If ex.[GetType]().FullName.EndsWith("TrustNotGrantedException", StringComparison.OrdinalIgnoreCase) Then Return True
            If ex.Message IsNot Nothing AndAlso ex.Message.IndexOf("User has refused to grant required permissions", StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
            Return IsTrustNotGranted(ex.InnerException)
        End Function


        ''' <summary>
        ''' Performs a user-initiated update check and, when available, installs the update.
        ''' Supports ClickOnce network deployments and local update path scenarios.
        ''' </summary>
        ''' <param name="appname">Add-in name used to select per-app settings keys (prefix matters: Word/Exce/Outl).</param>
        ''' <param name="LocalPath">Optional local update root path. If empty and network deployed, ClickOnce is used.</param>
        ''' <param name="context">Optional shared context used for INI configuration update checks.</param>
        Public Sub CheckAndInstallUpdates(appname As String, LocalPath As String, Optional context As ISharedContext = Nothing)
            Try
                Dim currentDate As Date = Date.Now
                If ApplicationDeployment.IsNetworkDeployed AndAlso String.IsNullOrWhiteSpace(LocalPath) Then
                    Dim deployment As ApplicationDeployment = ApplicationDeployment.CurrentDeployment
                    WriteUpdateLog($"[CheckAndInstallUpdates] network-deployed app='{appname}' url='{deployment.UpdateLocation}' zone='{GetUrlZoneName(deployment.UpdateLocation.AbsoluteUri)}'")

                    Dim updateAvailable As Boolean = False
                    Dim hasUpdateInfo As Boolean = False
                    Try
                        updateAvailable = deployment.CheckForUpdate()
                        hasUpdateInfo = True
                    Catch ex As Exception
                        WriteUpdateLog("[CheckAndInstallUpdates] CheckForUpdate failed", ex)
                        UIInvokeMessage(
                            "The update check could not complete (network or permission issue). Try again later or use a manual reinstall if problems persist.",
                            $"{SharedMethods.AN} Updater")
                    End Try

                    If hasUpdateInfo AndAlso updateAvailable Then
                        Dim vstoUrl = deployment.UpdateLocation.AbsoluteUri.Replace(".application", ".vsto")
                        Dim dialogResult As Integer = SharedMethods.ShowCustomYesNoBox(
                            $"An update is available. Install now?",
                            "Yes", "No", $"{SharedMethods.AN} Updater")

                        If dialogResult = 1 Then
                            WriteUpdateLog($"[CheckAndInstallUpdates] user accepted update vsto='{vstoUrl}'")
                            ShowUpdatingSplash("Updating …")
                            Try
                                RunVstoInstaller(vstoUrl)
                            Finally
                                CloseUpdatingSplash()
                            End Try
                        Else
                            WriteUpdateLog("[CheckAndInstallUpdates] user declined update")
                        End If
                    ElseIf hasUpdateInfo Then
                        UIInvokeMessage($"No updates are currently available ({deployment.UpdateLocation.AbsoluteUri}).", $"{SharedMethods.AN} Updater")
                    End If

                    Select Case Left(appname, 4)
                        Case "Word" : My.Settings.LastUpdateCheckWord = currentDate
                        Case "Exce" : My.Settings.LastUpdateCheckExcel = currentDate
                        Case "Outl" : My.Settings.LastUpdateCheckOutlook = currentDate
                    End Select
                    My.Settings.Save()
                Else
                    If LocalPath = "" Then
                        UIInvokeMessage(
                            $"This version of {SharedMethods.AN} has not been configured with an update path ('UpdatePath ='). Ask your administrator.",
                            $"{SharedMethods.AN} Updater")
                    Else
                        LocalPath = SharedMethods.ExpandEnvironmentVariables(LocalPath)
                        Dim dialogResult As Integer = SharedMethods.ShowCustomYesNoBox(
                            $"This will start the local installer. If a newer version exists under '{LocalPath}', it will be installed. Proceed?",
                            "Yes", "No", $"{SharedMethods.AN} Updater")
                        If dialogResult = 1 Then
                            Dim vstoFilePath As String = ""
                            Select Case Left(appname, 4)
                                Case "Word" : vstoFilePath = Path.Combine(LocalPath, $"word\{SharedMethods.AN3} for Word.vsto")
                                Case "Exce" : vstoFilePath = Path.Combine(LocalPath, $"excel\{SharedMethods.AN3} for Excel.vsto")
                                Case "Outl" : vstoFilePath = Path.Combine(LocalPath, $"outlook\{SharedMethods.AN3} for Outlook.vsto")
                            End Select
                            If File.Exists(vstoFilePath) Then
                                ShowUpdatingSplash("Updating …")
                                Try
                                    RunVstoInstaller(vstoFilePath)
                                Finally
                                    CloseUpdatingSplash()
                                End Try
                            Else
                                UIInvokeMessage(
                                    $"Installer not found: '{vstoFilePath}'. Check 'UpdatePath =' in '{SharedMethods.AN2}.ini'.",
                                    $"{SharedMethods.AN} Updater")
                            End If
                        End If
                    End If
                End If

                ' === INI Configuration Updates (manual check) ===                
                ''' <summary>
                ''' When a shared context is provided, attempts an INI configuration update check and shows a message
                ''' if no updates were applied.
                ''' </summary>
                If context IsNot Nothing Then
                    Try
                        If MainControl IsNot Nothing AndAlso MainControl.InvokeRequired Then
                            MainControl.Invoke(Sub() If Not SharedMethods.CheckForIniUpdates(context) Then UIInvokeMessage(
                                "No configuration updates available or made",
                                $"{SharedMethods.AN} INI Updater"))
                        Else
                            If Not SharedMethods.CheckForIniUpdates(context) Then UIInvokeMessage(
                                "No configuration updates available or made",
                                $"{SharedMethods.AN} INI Updater")
                        End If
                    Catch iniEx As Exception
                        WriteUpdateLog("[CheckAndInstallUpdates] INI update check failed", iniEx)
                    End Try
                End If

            Catch ex As DeploymentException
                WriteUpdateLog("[CheckAndInstallUpdates] DeploymentException", ex)
                UIInvokeMessage(
                    $"Error during update (you may try a manual install via {SharedMethods.AppsUrl}): {ex.Message}",
                    $"{SharedMethods.AN} Updater")
            Catch ex As Exception
                WriteUpdateLog("[CheckAndInstallUpdates] Unexpected Exception", ex)
                UIInvokeMessage(
                    $"Unexpected error during update: {ex.Message}",
                    $"{SharedMethods.AN} Updater")
            End Try
        End Sub

        ''' <summary>Cached add-in name used by async handlers and retry logic.</summary>
        Private Shared _appname As String

        ''' <summary>Cached local update path used by periodic checks.</summary>
        Private Shared _localPath As String

        ''' <summary>Cached update check interval in days (special -1 value implies infinite retry mode in this class).</summary>
        Private Shared _checkIntervalInDays As Integer

        ''' <summary>Cached shared context used for INI update checks.</summary>
        Private Shared _context As ISharedContext = Nothing

        ''' <summary>
        ''' Loads the per-app retry state from `My.Settings`.
        ''' </summary>
        ''' <param name="day">Stored retry day.</param>
        ''' <param name="count">Stored retry count for the day.</param>
        ''' <param name="shownToday">Whether a pause prompt has already been shown today.</param>
        Private Shared Sub GetRetryStateFromSettings(ByRef day As Date, ByRef count As Integer, ByRef shownToday As Boolean)
            day = Date.MinValue : count = 0 : shownToday = False
            Select Case Left(_appname, 4)
                Case "Word"
                    day = My.Settings.UpdateRetryDateWord
                    count = My.Settings.UpdateRetryCountWord
                    shownToday = My.Settings.UpdateRetryPromptShownWord
                Case "Exce"
                    day = My.Settings.UpdateRetryDateExcel
                    count = My.Settings.UpdateRetryCountExcel
                    shownToday = My.Settings.UpdateRetryPromptShownExcel
                Case "Outl"
                    day = My.Settings.UpdateRetryDateOutlook
                    count = My.Settings.UpdateRetryCountOutlook
                    shownToday = My.Settings.UpdateRetryPromptShownOutlook
            End Select
        End Sub

        ''' <summary>
        ''' Saves the per-app retry state into `My.Settings`.
        ''' </summary>
        ''' <param name="day">Retry state day.</param>
        ''' <param name="count">Retry failures count.</param>
        ''' <param name="shownToday">Whether the prompt has been shown for the day.</param>
        Private Shared Sub SetRetryStateToSettings(day As Date, count As Integer, shownToday As Boolean)
            Select Case Left(_appname, 4)
                Case "Word"
                    My.Settings.UpdateRetryDateWord = day
                    My.Settings.UpdateRetryCountWord = count
                    My.Settings.UpdateRetryPromptShownWord = shownToday
                Case "Exce"
                    My.Settings.UpdateRetryDateExcel = day
                    My.Settings.UpdateRetryCountExcel = count
                    My.Settings.UpdateRetryPromptShownExcel = shownToday
                Case "Outl"
                    My.Settings.UpdateRetryDateOutlook = day
                    My.Settings.UpdateRetryCountOutlook = count
                    My.Settings.UpdateRetryPromptShownOutlook = shownToday
            End Select
            Try : My.Settings.Save() : Catch : End Try
        End Sub

        ''' <summary>
        ''' Resets retry counters if the stored day differs from today.
        ''' </summary>
        ''' <param name="day">Stored retry day (updated if reset).</param>
        ''' <param name="count">Stored retry count (updated if reset).</param>
        ''' <param name="shownToday">Stored prompt flag (updated if reset).</param>
        Private Shared Sub ResetRetryIfNewDay(ByRef day As Date, ByRef count As Integer, ByRef shownToday As Boolean)
            If day.Date <> Date.Today Then
                day = Date.Today : count = 0 : shownToday = False
                SetRetryStateToSettings(day, count, shownToday)
            End If
        End Sub

        ' Version if multiple prompts per day (every 3 failures)

        ''' <summary>
        ''' Alternative retry/prompt implementation (currently unused). Records a failure and may prompt the user
        ''' after reaching the daily threshold, allowing multiple prompts per day based on the comment.
        ''' </summary>
        ''' <param name="optionalReason">Optional failure reason to log.</param>
        ''' <returns>True to indicate the caller should pause update checks; otherwise False.</returns>
        Private Shared Function _RecordCheckFailureAndMaybePrompt(optionalReason As String) As Boolean
            If _checkIntervalInDays = -1 Then
                Return False ' infinite retry mode
            End If

            Dim day As Date, count As Integer, shown As Boolean
            GetRetryStateFromSettings(day, count, shown)
            ResetRetryIfNewDay(day, count, shown)

            count += 1
            SetRetryStateToSettings(day, count, shown)
            WriteUpdateLog($"[Retry] update check failed (reason='{optionalReason}'), todayCount={count}/{MaxDailyUpdateRetries}")

            ' Still within the silent window
            If count < MaxDailyUpdateRetries Then Return False

            ' We reached the threshold: ask user
            Dim msg = $"No update check was possible at least {MaxDailyUpdateRetries} times today (e.g., network issues). Continue trying or pause for {_checkIntervalInDays} day(s)?"
            Dim choice = UIInvokeYesNo(msg, "Continue trying", "Pause", $"{SharedMethods.AN} Updater")

            If choice = 1 Then
                ' Continue: restart cycle (3 more silent tries)
                count = 0
                shown = False    ' allow prompting again if another 3 fail
                SetRetryStateToSettings(day, count, shown)
                Return False      ' do not pause
            Else
                ' Pause: mark shown to avoid re-prompting; caller will pause via returned True
                shown = True
                SetRetryStateToSettings(day, count, shown)
                Return True
            End If
        End Function


        ' Version if only one prompt per day

        ''' <summary>
        ''' Records a failed update check attempt and may prompt the user to pause checks after reaching
        ''' the daily failure threshold.
        ''' </summary>
        ''' <param name="optionalReason">Optional reason string for logging.</param>
        ''' <returns>True to indicate the caller should pause update checks; otherwise False.</returns>
        Private Shared Function RecordCheckFailureAndMaybePrompt(optionalReason As String) As Boolean
            If _checkIntervalInDays = -1 Then
                ' Infinite retry mode: never pause, never prompt
                Return False
            End If

            Dim day As Date, count As Integer, shown As Boolean
            GetRetryStateFromSettings(day, count, shown)
            ResetRetryIfNewDay(day, count, shown)
            count += 1
            SetRetryStateToSettings(day, count, shown)
            WriteUpdateLog($"[Retry] update check failed (reason='{optionalReason}'), todayCount={count}/{MaxDailyUpdateRetries}")

            If count < MaxDailyUpdateRetries Then Return False
            If Not shown Then
                Dim msg = $"No update check was possible at least {MaxDailyUpdateRetries} times today (e.g., because of network issues). Continue trying today or pause for {_checkIntervalInDays} day(s)?"
                Dim choice = UIInvokeYesNo(msg, "Continue trying", "Pause", $"{SharedMethods.AN} Updater")
                shown = True
                SetRetryStateToSettings(day, count, shown)
                Return (choice <> 1) ' True means pause
            End If
            Return False
        End Function

        ''' <summary>
        ''' Shows a two-button Yes/No style dialog with custom button labels, marshaling to the UI thread when required.
        ''' </summary>
        ''' <param name="bodyText">Dialog body text.</param>
        ''' <param name="button1Text">Text for the first button.</param>
        ''' <param name="button2Text">Text for the second button.</param>
        ''' <param name="caption">Dialog caption.</param>
        ''' <returns>Dialog result as returned by `SharedMethods.ShowCustomYesNoBox`.</returns>
        Private Shared Function UIInvokeYesNo(bodyText As String, button1Text As String, button2Text As String, caption As String) As Integer
            Try
                If MainControl IsNot Nothing AndAlso MainControl.IsHandleCreated AndAlso MainControl.InvokeRequired Then
                    Return CInt(MainControl.Invoke(New Func(Of Integer)(
                        Function() SharedMethods.ShowCustomYesNoBox(bodyText, button1Text, button2Text, caption))))
                Else
                    Return SharedMethods.ShowCustomYesNoBox(bodyText, button1Text, button2Text, caption)
                End If
            Catch
                Return SharedMethods.ShowCustomYesNoBox(bodyText, button1Text, button2Text, caption)
            End Try
        End Function

        ''' <summary>
        ''' Performs an automatic/periodic update check, honoring the check interval and daily retry state.
        ''' May run an async ClickOnce check (network deployed) or run a local `.vsto` installer check (local path).
        ''' </summary>
        ''' <param name="checkIntervalInDays">
        ''' Interval in days between checks; 0 disables checks; -1 enables infinite retry mode for failure handling in this class.
        ''' </param>
        ''' <param name="appname">Add-in name used to select per-app settings keys (prefix matters: Word/Exce/Outl).</param>
        ''' <param name="LocalPath">Optional local update root path; if empty and network deployed, ClickOnce is used.</param>
        ''' <param name="context">Optional shared context used for INI configuration update checks.</param>
        Public Shared Sub PeriodicCheckForUpdates(
            checkIntervalInDays As Integer,
            appname As String,
            LocalPath As String,
            Optional context As ISharedContext = Nothing)

            Dim splashManagedByOnCheck As Boolean = False

            Try
                If checkIntervalInDays = 0 Then Return
                _appname = appname
                _localPath = LocalPath
                _checkIntervalInDays = checkIntervalInDays
                _context = context

                Dim lastCheck As Date = If(
                Left(_appname, 4) = "Word", My.Settings.LastUpdateCheckWord,
                If(Left(_appname, 4) = "Exce", My.Settings.LastUpdateCheckExcel,
                   If(Left(_appname, 4) = "Outl", My.Settings.LastUpdateCheckOutlook, Date.MinValue)))
                Dim nowDate As Date = Date.Now
                Dim days As Double = (nowDate - lastCheck).TotalDays

                ' Allow retries within the same day even if interval not reached (if we had failures today but below threshold)
                Dim rDay As Date, rCount As Integer, rShown As Boolean
                GetRetryStateFromSettings(rDay, rCount, rShown)
                ResetRetryIfNewDay(rDay, rCount, rShown)
                Dim allowRetryToday As Boolean = (rCount > 0 AndAlso rCount < MaxDailyUpdateRetries)

                WriteUpdateLog($"[PeriodicCheck] app='{_appname}' localPath='{_localPath}' interval={_checkIntervalInDays} lastCheck='{lastCheck:yyyy-MM-dd HH:mm:ss}' ageDays={days:0.0} retryToday={rCount}")

                If days < _checkIntervalInDays AndAlso _checkIntervalInDays > 0 AndAlso Not allowRetryToday Then
                    WriteUpdateLog("[PeriodicCheck] skipped - interval not reached")
                    Return
                End If

                If ApplicationDeployment.IsNetworkDeployed AndAlso String.IsNullOrWhiteSpace(_localPath) Then
                    Dim dep = ApplicationDeployment.CurrentDeployment
                    WriteUpdateLog($"[PeriodicCheck] network-deployed url='{dep.UpdateLocation}' zone='{GetUrlZoneName(dep.UpdateLocation.AbsoluteUri)}'")

                    If ShowCheckingSplash Then
                        ShowUpdatingSplash("Checking updates …")
                        splashManagedByOnCheck = True
                    Else
                        splashManagedByOnCheck = False
                    End If

                    RemoveHandler dep.CheckForUpdateCompleted, AddressOf OnCheck
                    AddHandler dep.CheckForUpdateCompleted, AddressOf OnCheck
                    dep.CheckForUpdateAsync()
                Else
                    Dim vstoFile = Path.Combine(
                    Environment.ExpandEnvironmentVariables(_localPath),
                    $"{_appname.ToLowerInvariant()}\{SharedMethods.AN3} for {_appname}.vsto")

                    WriteUpdateLog($"[PeriodicCheck] local-deployed vsto='{vstoFile}'")

                    If File.Exists(vstoFile) Then
                        ShowUpdatingSplash("Updating …")
                        Try
                            RunVstoInstaller(vstoFile)
                        Finally
                            CloseUpdatingSplash()
                        End Try
                    Else
                        UIInvokeMessage(
                        $"The configuration asks me to check for local updates of {SharedMethods.AN}, but I have not found '{vstoFile}'. Please inform your administrator.",
                        $"{SharedMethods.AN} Updater")
                    End If

                    ' Local path execution = a successful check happened; record timestamp
                    SaveTimestamp(nowDate)
                    ' Also reset daily retries on success
                    Dim day As Date = Date.Today : Dim cnt As Integer = 0 : Dim shown As Boolean = False
                    SetRetryStateToSettings(day, cnt, shown)
                End If

                ' === INI Configuration Updates (async path) ===
                If _context IsNot Nothing Then
                    Try
                        If MainControl IsNot Nothing AndAlso MainControl.InvokeRequired Then
                            MainControl.Invoke(Sub() SharedMethods.CheckForIniUpdates(_context))
                        Else
                            SharedMethods.CheckForIniUpdates(_context)
                        End If
                    Catch iniEx As Exception
                        WriteUpdateLog("[OnCheck] INI update check failed", iniEx)
                    End Try
                End If

            Catch dex As DeploymentException
                WriteUpdateLog("[PeriodicCheck] DeploymentException", dex)
                Dim pause As Boolean = RecordCheckFailureAndMaybePrompt("DeploymentException")
                If pause AndAlso _checkIntervalInDays > 0 Then SaveTimestamp(Date.Now)

            Catch ex As Exception
                WriteUpdateLog("[PeriodicCheck] Unexpected Exception", ex)
                Dim pause As Boolean = RecordCheckFailureAndMaybePrompt(ex.Message)
                If pause AndAlso _checkIntervalInDays > 0 Then SaveTimestamp(Date.Now)

            Finally
                If Not splashManagedByOnCheck Then
                    CloseUpdatingSplash()
                End If
            End Try
        End Sub

        ''' <summary>
        ''' Async callback for ClickOnce update checks started by `PeriodicCheckForUpdates`.
        ''' Handles update availability, errors, retry tracking, and optional interactive installer fallback.
        ''' </summary>
        Private Shared Sub OnCheck(sender As Object, e As CheckForUpdateCompletedEventArgs)
            Dim dep = CType(sender, ApplicationDeployment)
            Dim nowDate As Date = Date.Now
            Dim saved As Boolean = False

            Try
                CloseUpdatingSplash()

                If e.Error IsNot Nothing Then
                    WriteUpdateLog($"[OnCheck] error url='{dep.UpdateLocation}' zone='{GetUrlZoneName(dep.UpdateLocation.AbsoluteUri)}'", e.Error)

                    If IsTrustNotGranted(e.Error) AndAlso CanShowInteractiveUi() Then
                        Dim appUrl = dep.UpdateLocation.AbsoluteUri
                        Dim vstoUrl = appUrl.Replace(".application", ".vsto")
                        WriteUpdateLog($"[OnCheck] TrustNotGranted → trying interactive VSTOInstaller on '{vstoUrl}'")

                        ShowUpdatingSplash("Updating …")
                        Try
                            RunVstoInstaller(vstoUrl)
                        Finally
                            CloseUpdatingSplash()
                        End Try

                        If _checkIntervalInDays > 0 Then SaveTimestamp(nowDate) : saved = True
                        ' Reset retries on success path (we attempted install)
                        Dim d As Date = Date.Today : Dim c As Integer = 0 : Dim s As Boolean = False
                        SetRetryStateToSettings(d, c, s)
                        Return
                    End If

                    ' Silent failure handling with daily retries
                    Dim pause As Boolean = RecordCheckFailureAndMaybePrompt(If(e.Error Is Nothing, "Unknown", e.Error.Message))
                    If pause AndAlso _checkIntervalInDays > 0 Then SaveTimestamp(nowDate) : saved = True
                    Return
                End If

                If e.UpdateAvailable Then
                    Dim localV = dep.CurrentVersion.ToString()
                    Dim remoteV = e.AvailableVersion.ToString()
                    WriteUpdateLog($"[OnCheck] update available current='{localV}' new='{remoteV}' url='{dep.UpdateLocation}'")

                    Dim prompt = $"A new version is available (current: {localV}, new: {remoteV}). Do you want to install it now?"
                    Dim choice = UIInvokePrompt(prompt, $"{SharedMethods.AN} Updater")

                    If choice = 1 Then
                        Dim appUrl = dep.UpdateLocation.AbsoluteUri
                        Dim vstoUrl = appUrl.Replace(".application", ".vsto")
                        WriteUpdateLog($"[OnCheck] user accepted → installing '{vstoUrl}'")

                        ShowUpdatingSplash("Updating …")
                        Try
                            RunVstoInstaller(vstoUrl)
                        Finally
                            CloseUpdatingSplash()
                        End Try

                        SaveTimestamp(nowDate) : saved = True
                    Else
                        If _checkIntervalInDays = -1 Then
                            SaveTimestamp(nowDate) : saved = True
                        ElseIf _checkIntervalInDays > 0 Then
                            Dim postPrompt = $"Do you want to pause update checks for {_checkIntervalInDays} days?"
                            Dim postChoice = UIInvokePrompt(postPrompt, $"{SharedMethods.AN} Updater")
                            If postChoice = 1 Then
                                SaveTimestamp(nowDate) : saved = True
                            End If
                        End If
                    End If
                Else
                    WriteUpdateLog("[OnCheck] no update available")
                    If _checkIntervalInDays > 0 Then
                        SaveTimestamp(nowDate) : saved = True
                    End If
                End If

                ' On any successful network check outcome, reset the daily retry state
                Dim day As Date = Date.Today : Dim cnt As Integer = 0 : Dim shown As Boolean = False
                SetRetryStateToSettings(day, cnt, shown)

                ' === INI Configuration Updates (async network path) ===
                If _context IsNot Nothing Then
                    Try
                        If MainControl IsNot Nothing AndAlso MainControl.InvokeRequired Then
                            MainControl.Invoke(Sub() SharedMethods.CheckForIniUpdates(_context))
                        Else
                            SharedMethods.CheckForIniUpdates(_context)
                        End If
                    Catch iniEx As Exception
                        WriteUpdateLog("[OnCheck] INI update check failed", iniEx)
                    End Try
                End If

            Catch dex As DeploymentException
                WriteUpdateLog("[OnCheck] DeploymentException", dex)
                Dim pause As Boolean = RecordCheckFailureAndMaybePrompt("DeploymentException")
                If pause AndAlso _checkIntervalInDays > 0 Then SaveTimestamp(nowDate) : saved = True

            Catch ex As Exception
                WriteUpdateLog("[OnCheck] Unexpected Exception", ex)
                Dim pause As Boolean = RecordCheckFailureAndMaybePrompt(ex.Message)
                If pause AndAlso _checkIntervalInDays > 0 Then SaveTimestamp(nowDate) : saved = True

            Finally
                CloseUpdatingSplash()
                If _checkIntervalInDays > 0 AndAlso Not saved Then
                    ' Do not force-save a timestamp here unless user chose pause.
                End If
            End Try
        End Sub


        ''' <summary>
        ''' Win32 API that permits a specific process to set the foreground window.
        ''' </summary>
        <DllImport("user32.dll")>
        Private Shared Function AllowSetForegroundWindow(dwProcessId As Integer) As Boolean
        End Function

        ''' <summary>
        ''' Win32 API to change window show state.
        ''' </summary>
        <DllImport("user32.dll")>
        Private Shared Function ShowWindow(hWnd As IntPtr, nCmdShow As Integer) As Boolean
        End Function

        ''' <summary>Win32 constant indicating a window should be shown (SW_SHOW).</summary>
        Private Const SW_SHOW As Integer = 5

        ''' <summary>
        ''' Win32 API to get the current foreground window handle.
        ''' </summary>
        <DllImport("user32.dll")>
        Private Shared Function GetForegroundWindow() As IntPtr
        End Function

        ''' <summary>
        ''' Win32 API to get the thread/process id for the owning window.
        ''' </summary>
        <DllImport("user32.dll")>
        Private Shared Function GetWindowThreadProcessId(hWnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
        End Function

        ''' <summary>
        ''' Win32 API to attach/detach input processing between threads (used to satisfy foreground focus restrictions).
        ''' </summary>
        <DllImport("user32.dll")>
        Private Shared Function AttachThreadInput(idAttach As Integer, idAttachTo As Integer, fAttach As Boolean) As Boolean
        End Function

        ''' <summary>
        ''' Win32 API to enumerate top-level windows.
        ''' </summary>
        <DllImport("user32.dll")>
        Private Shared Function EnumWindows(lpEnumFunc As EnumWindowsProc, lParam As IntPtr) As Boolean
        End Function

        ''' <summary>Delegate for `EnumWindows` callback.</summary>
        Private Delegate Function EnumWindowsProc(hWnd As IntPtr, lParam As IntPtr) As Boolean

        ''' <summary>
        ''' Win32 API that indicates whether a window is visible.
        ''' </summary>
        <DllImport("user32.dll")>
        Private Shared Function IsWindowVisible(hWnd As IntPtr) As Boolean
        End Function

        ''' <summary>
        ''' Win32 API to get window text.
        ''' </summary>
        <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
        Private Shared Function GetWindowText(hWnd As IntPtr, sb As StringBuilder, cch As Integer) As Integer
        End Function

        ''' <summary>
        ''' Win32 API to get the calling thread id.
        ''' </summary>
        <DllImport("kernel32.dll")>
        Private Shared Function GetCurrentThreadId() As Integer
        End Function

        ' === Helper: enumerate top-level windows for a process ===
        ''' <summary>
        ''' Enumerates visible top-level windows belonging to the specified process id.
        ''' </summary>
        ''' <param name="procId">Process id.</param>
        Private Shared Function EnumProcessTopLevelWindows(procId As Integer) As List(Of IntPtr)
            Dim list As New List(Of IntPtr)
            EnumWindows(Function(h, p)
                            Dim pid As Integer = 0
                            GetWindowThreadProcessId(h, pid)
                            If pid = procId AndAlso IsWindowVisible(h) Then
                                list.Add(h)
                            End If
                            Return True
                        End Function,
                IntPtr.Zero)
            Return list
        End Function

        ' === Helper: bring process window to foreground with retries ===
        ''' <summary>
        ''' Attempts to bring a newly started process window to the foreground, using retries and thread input attach
        ''' to satisfy foreground lock rules. Writes the result to the update log.
        ''' </summary>
        ''' <param name="p">Target process.</param>
        ''' <param name="logPrefix">Prefix for log entries written by this method.</param>
        ''' <returns>True if a target window was found and activated; otherwise False.</returns>
        Private Shared Function BringProcessWindowToFront(p As Process, logPrefix As String) As Boolean
            Const totalWaitMs As Integer = 5000
            Const stepMs As Integer = 300
            Dim waited As Integer = 0
            Dim brought As Boolean = False

            Try
                ' Permit the new process to steal foreground
                Try : AllowSetForegroundWindow(p.Id) : Catch : End Try

                While waited <= totalWaitMs AndAlso Not p.HasExited AndAlso Not brought
                    p.Refresh()
                    Dim targetH As IntPtr = IntPtr.Zero

                    ' 1. Prefer Process.MainWindowHandle if valid
                    If p.MainWindowHandle <> IntPtr.Zero Then
                        targetH = p.MainWindowHandle
                    Else
                        ' 2. Enumerate windows (bootstrap scenarios)
                        Dim wins = EnumProcessTopLevelWindows(p.Id)
                        If wins.Count > 0 Then
                            targetH = wins(0)
                        End If
                    End If

                    If targetH <> IntPtr.Zero Then
                        Dim thisThread = GetCurrentThreadId()
                        Dim winThread = GetWindowThreadProcessId(targetH, Nothing)

                        ' Attach input queues to overcome foreground lock
                        AttachThreadInput(winThread, thisThread, True)
                        Try
                            ShowWindow(targetH, SW_SHOW)
                            NativeMethods.SetForegroundWindow(targetH)
                            SetWindowPos(targetH, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)
                            brought = True
                        Finally
                            AttachThreadInput(winThread, thisThread, False)
                        End Try

                        ' Drop TOPMOST shortly after (async)
                        System.Threading.Tasks.Task.Run(Async Sub()
                                                            Await System.Threading.Tasks.Task.Delay(600)
                                                            Try
                                                                SetWindowPos(targetH, HWND_NOTOPMOST, 0, 0, 0, 0,
                                              SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)
                                                            Catch
                                                            End Try
                                                        End Sub)
                    End If

                    If Not brought Then
                        Thread.Sleep(stepMs)
                        waited += stepMs
                    End If
                End While

                WriteUpdateLog($"{logPrefix} bring-to-front result=" & If(brought, "Success", "Failed/TimedOut"))
                Return brought
            Catch ex As Exception
                WriteUpdateLog($"{logPrefix} bring-to-front exception", ex)
                Return False
            End Try
        End Function


        ''' <summary>
        ''' Runs `VSTOInstaller.exe` to install a `.vsto` from a local path or URL. Attempts a silent install first,
        ''' and falls back to an interactive install that is brought to the foreground.
        ''' </summary>
        ''' <param name="pathOrUrl">Local path or URL to the `.vsto` deployment manifest.</param>
        Private Shared Sub RunVstoInstaller(pathOrUrl As String)
            ' Try to locate VSTOInstaller in both x64 and x86 common locations
            Dim candidates As New List(Of String)
            Try
                Dim base1 = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles), "Microsoft Shared", "VSTO")
                Dim base2 = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFilesX86), "Microsoft Shared", "VSTO")
                If Directory.Exists(base1) Then candidates.AddRange(Directory.GetFiles(base1, "VSTOInstaller.exe", SearchOption.AllDirectories))
                If Directory.Exists(base2) Then candidates.AddRange(Directory.GetFiles(base2, "VSTOInstaller.exe", SearchOption.AllDirectories))
            Catch
            End Try
            Dim installer = candidates.FirstOrDefault()

            If installer Is Nothing Then
                WriteUpdateLog("[RunVstoInstaller] VSTOInstaller.exe not found")
                UIInvokeMessage(
                    "The update could not be completed (VSTOInstaller.exe not found). Please inform your administrator.",
                    $"{SharedMethods.AN} Updater")
                Return
            End If

            WriteUpdateLog($"[RunVstoInstaller] using='{installer}' target='{pathOrUrl}'")

            ' Ensure splash is visible (caller normally showed it, but be defensive)
            If _splash Is Nothing OrElse _splash.IsDisposed Then
                ShowUpdatingSplash("Updating …")
            Else
                Try : _splash.UpdateMessage("Updating …") : Catch : End Try
            End If

            ' 1) Silent attempt (fast, no UI)
            Dim silentOk As Boolean = False
            Try
                Dim psiSilent = New ProcessStartInfo(installer, $"/I ""{pathOrUrl}"" /S") With {
                    .UseShellExecute = False,
                    .CreateNoWindow = True
                }
                Using p = Process.Start(psiSilent)
                    p.WaitForExit()
                    WriteUpdateLog($"[RunVstoInstaller] silent exitCode={p.ExitCode}")
                    silentOk = (p.ExitCode = 0)
                End Using
            Catch ex As Exception
                WriteUpdateLog("[RunVstoInstaller] silent failed", ex)
                silentOk = False
            End Try

            If silentOk Then
                ' Success: close splash, inform user
                CloseUpdatingSplash()
                UIInvokeMessage(
                       "Update completed. It will be active the next time you restart your application.",
                      $"{SharedMethods.AN} Updater")
                Return
            End If

            ' 2) Interactive fallback (needs trust consent)
            Try
                ' Update splash text to indicate fallback; keep it till just before showing installer UI
                If _splash IsNot Nothing AndAlso Not _splash.IsDisposed Then
                    Try : _splash.UpdateMessage("Opening installer …") : Catch : End Try
                End If

                Dim psiUi = New ProcessStartInfo(installer, $"/I ""{pathOrUrl}""") With {
                    .UseShellExecute = False,
                    .CreateNoWindow = False
                }

                ' Close splash so installer window can gain foreground cleanly
                CloseUpdatingSplash()

                Using p = Process.Start(psiUi)
                    Dim settingsForm As Form = Nothing
                    Dim restoreTopMost As Boolean = False
                    Dim restoreVisible As Boolean = False

                    ' Hide any Settings window to avoid z-order conflicts
                    Try
                        settingsForm = System.Windows.Forms.Application.OpenForms.Cast(Of Form)().
                            FirstOrDefault(Function(f) f.Visible AndAlso
                                                   (f.Name.IndexOf("Setting", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                                    f.Text.IndexOf("Setting", StringComparison.OrdinalIgnoreCase) >= 0))
                        If settingsForm IsNot Nothing Then
                            restoreTopMost = settingsForm.TopMost
                            restoreVisible = settingsForm.Visible
                            settingsForm.TopMost = False
                            settingsForm.Hide()
                            WriteUpdateLog("[RunVstoInstaller] temporarily hid Settings window")
                        End If
                    Catch
                    End Try

                    Try : p.WaitForInputIdle(4000) : Catch : End Try
                    BringProcessWindowToFront(p, "[RunVstoInstaller]")

                    p.WaitForExit()
                    WriteUpdateLog($"[RunVstoInstaller] interactive exitCode={p.ExitCode}")

                    ' Restore Settings window
                    If settingsForm IsNot Nothing Then
                        Try
                            If restoreVisible Then settingsForm.Show()
                            settingsForm.TopMost = restoreTopMost
                        Catch
                        End Try
                    End If

                    If p.ExitCode = 0 Then
                        UIInvokeMessage(
                            "Update completed. It will be active the next time you restart your application.",
                            $"{SharedMethods.AN} Updater")
                    Else
                        UIInvokeMessage(
                            "The update could not be completed. A required trust confirmation may have been refused or blocked by policy. You can always try a manual install by visiting " & SharedMethods.AppsUrl & ".",
                            $"{SharedMethods.AN} Updater")
                    End If
                End Using

            Catch ex As Exception
                WriteUpdateLog("[RunVstoInstaller] interactive failed", ex)
                UIInvokeMessage(
                    $"The update could not be completed: {ex.Message}. Please inform your administrator. You can always try a manual install by visiting {SharedMethods.AppsUrl}.",
                    $"{SharedMethods.AN} Updater")
            Finally
                CloseUpdatingSplash()
            End Try
        End Sub


        ''' <summary>
        ''' Persists the last update check timestamp to per-app settings and logs the change.
        ''' </summary>
        ''' <param name="timeStamp">Timestamp to persist.</param>
        Private Shared Sub SaveTimestamp(timeStamp As Date)
            Select Case Left(_appname, 4)
                Case "Word" : My.Settings.LastUpdateCheckWord = timeStamp
                Case "Exce" : My.Settings.LastUpdateCheckExcel = timeStamp
                Case "Outl" : My.Settings.LastUpdateCheckOutlook = timeStamp
            End Select
            My.Settings.Save()
            WriteUpdateLog($"[SaveTimestamp] {_appname} -> {timeStamp:yyyy-MM-dd HH:mm:ss}")
        End Sub

    End Class


End Namespace