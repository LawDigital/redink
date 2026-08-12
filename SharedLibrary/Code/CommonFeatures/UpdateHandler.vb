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
'  - Logging: Writes to `%AppData%\{AN2}\RI_Logfile.txt` and trims by size/line count.
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
        ''' Number of cumulative periodic update failures (across days/sessions) before showing
        ''' the manual download fallback to the user.
        ''' </summary>
        Private Const ManualFallbackThreshold As Integer = 3

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
        ''' Appends an entry to the updater log under `%AppData%\{AN2}\RI_Logfile.txt` and triggers trimming when needed.
        ''' </summary>
        ''' <param name="message">Log message to append.</param>
        ''' <param name="ex">Optional exception details to include.</param>
        Public Shared Sub WriteUpdateLog(message As String, Optional ex As Exception = Nothing)
            Try
                Dim logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), SharedMethods.AN2)
                If Not Directory.Exists(logDir) Then Directory.CreateDirectory(logDir)
                Dim logPath = Path.Combine(logDir, LogFileName)

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
        ''' Shows a Yes/No prompt on the UI thread (always marshaled) and returns the underlying dialog result.
        ''' </summary>
        ''' <param name="prompt">Prompt body text.</param>
        ''' <param name="caption">Dialog caption.</param>
        Private Shared Function UIInvokePrompt(prompt As String, caption As String) As Integer
            If MainControl IsNot Nothing AndAlso MainControl.IsHandleCreated Then
                Dim result As Integer = 0
                MainControl.Invoke(
                    Sub()
                        NativeMethods.SetForegroundWindow(HostHandle)
                        result = SharedMethods.ShowCustomYesNoBox(prompt, "Yes", "No", caption)
                    End Sub)
                Return result
            Else
                ' Fallback: no UI control available — log and return "declined"
                WriteUpdateLog($"[UIInvokePrompt] MainControl unavailable, auto-declining: {prompt}")
                Return 0
            End If
        End Function

        ''' <summary>
        ''' Shows a message box on the UI thread (always marshaled).
        ''' </summary>
        ''' <param name="msg">Message body.</param>
        ''' <param name="caption">Dialog caption.</param>
        Private Shared Sub UIInvokeMessage(msg As String, caption As String)
            If MainControl IsNot Nothing AndAlso MainControl.IsHandleCreated Then
                MainControl.Invoke(
                    Sub()
                        NativeMethods.SetForegroundWindow(HostHandle)
                        SharedMethods.ShowCustomMessageBox(msg, caption)
                    End Sub)
            Else
                ' Fallback: no UI control available — log only
                WriteUpdateLog($"[UIInvokeMessage] MainControl unavailable, message suppressed: {msg}")
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
        ''' Determines whether the current add-in instance was installed via the central MSI package
        ''' rather than ClickOnce. Any ClickOnce deployment (network-deployed, whether from a local path
        ''' or a remote URL) returns False; otherwise the presence of a per-host MSI registry marker
        ''' (Preview or GA channel) indicates an MSI installation.
        ''' </summary>
        ''' <param name="appname">Add-in name used to select the host registry marker (prefix: Word/Exce/Outl).</param>
        Public Shared Function IsRunningFromMsi(appname As String) As Boolean

            ' Any ClickOnce deployment, whether installed from a local path or a remote URL, is not MSI.
            If System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed Then
                Return False
            End If

            Dim host As String
            Select Case Left(appname, 4)
                Case "Word" : host = "Word"
                Case "Exce" : host = "Excel"
                Case "Outl" : host = "Outlook"
                Case Else : host = String.Empty
            End Select

            If host = String.Empty Then Return False

            For Each channel As String In New String() {"Preview", "GA"}
                Using registryKey As Microsoft.Win32.RegistryKey =
                    Microsoft.Win32.Registry.LocalMachine.OpenSubKey(
                        $"Software\Red Ink\MSI\{channel}\{host}", False)

                    If registryKey IsNot Nothing Then
                        WriteUpdateLog($"[IsRunningFromMsi] MSI marker found for host='{host}' channel='{channel}'")
                        Return True
                    End If
                End Using
            Next

            Return False

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
                If IsRunningFromMsi(appname) Then
                    WriteUpdateLog($"[CheckAndInstallUpdates] MSI installation detected app='{appname}' — ClickOnce/local update check suppressed")
                    UIInvokeMessage(
                        $"This installation of {SharedMethods.AN} was deployed through a central software distribution (MSI). Automatic program updates are therefore disabled. Please contact your administrator for new versions.",
                        $"{SharedMethods.AN} Updater")
                ElseIf ApplicationDeployment.IsNetworkDeployed AndAlso String.IsNullOrWhiteSpace(LocalPath) Then
                    Dim deployment As ApplicationDeployment = ApplicationDeployment.CurrentDeployment
                    WriteUpdateLog($"[CheckAndInstallUpdates] network-deployed app='{appname}' url='{deployment.UpdateLocation}' zone='{GetUrlZoneName(deployment.UpdateLocation.AbsoluteUri)}'")

                    Dim updateAvailable As Boolean = False
                    Dim hasUpdateInfo As Boolean = False
                    Dim lastEx As Exception = Nothing
                    Const MaxRetries As Integer = 3

                    For attempt As Integer = 1 To MaxRetries
                        Try
                            updateAvailable = deployment.CheckForUpdate()
                            hasUpdateInfo = True
                            lastEx = Nothing
                            Exit For
                        Catch ex As Exception
                            lastEx = ex
                            WriteUpdateLog($"[CheckAndInstallUpdates] CheckForUpdate attempt {attempt}/{MaxRetries} failed", ex)
                            If attempt < MaxRetries Then
                                Thread.Sleep(1000 * CInt(Math.Pow(2, attempt - 1))) ' 1s, 2s backoff
                            End If
                        End Try
                    Next

                    If lastEx IsNot Nothing AndAlso Not hasUpdateInfo Then
                        ' CheckForUpdate failed — try lightweight manifest version check as fallback
                        WriteUpdateLog("[CheckAndInstallUpdates] CheckForUpdate failed after all retries, trying manifest fallback", lastEx)

                        Dim remoteVersion = GetRemoteVersionFromManifest(deployment.UpdateLocation.AbsoluteUri)
                        If remoteVersion IsNot Nothing AndAlso remoteVersion > deployment.CurrentVersion Then
                            ' Update IS available but ClickOnce validation failed (trust/prereq)
                            WriteUpdateLog($"[CheckAndInstallUpdates] manifest shows update: current='{deployment.CurrentVersion}' remote='{remoteVersion}'")

                            Dim dialogResult As Integer = SharedMethods.ShowCustomYesNoBox(
                                $"A new version is available (current: {deployment.CurrentVersion}, new: {remoteVersion}). Install now?" &
                                Environment.NewLine & Environment.NewLine &
                                "(Note: the standard update check encountered an issue, so an interactive install will be used.)",
                                "Yes", "No", $"{SharedMethods.AN} Updater")

                            If dialogResult = 1 Then
                                Dim vstoUrl = deployment.UpdateLocation.AbsoluteUri.Replace(".application", ".vsto")
                                WriteUpdateLog($"[CheckAndInstallUpdates] user accepted → interactive install '{vstoUrl}'")
                                ShowUpdatingSplash("Updating …")
                                Try
                                    RunVstoInstaller(vstoUrl)
                                Finally
                                    CloseUpdatingSplash()
                                End Try
                            Else
                                WriteUpdateLog("[CheckAndInstallUpdates] user declined update (manifest fallback)")
                            End If
                        ElseIf remoteVersion IsNot Nothing Then
                            ' No update needed
                            WriteUpdateLog($"[CheckAndInstallUpdates] manifest shows no update: current='{deployment.CurrentVersion}' remote='{remoteVersion}'")
                            UIInvokeMessage($"No updates are currently available (current: {deployment.CurrentVersion}, remote: {remoteVersion}).", $"{SharedMethods.AN} Updater")
                        Else
                            ' Could not fetch manifest either — offer manual download fallback
                            WriteUpdateLog("[CheckAndInstallUpdates] manifest fetch also failed, offering manual download")
                            Dim detail = If(lastEx.InnerException IsNot Nothing, lastEx.InnerException.Message, lastEx.Message)
                            Dim downloadUrl = $"{SharedMethods.AppsUrl}{SharedMethods.UpdateSubUrl}"

                            If MainControl IsNot Nothing AndAlso MainControl.IsHandleCreated Then
                                MainControl.Invoke(
                                    Sub()
                                        NativeMethods.SetForegroundWindow(HostHandle)
                                        SharedMethods.ShowCustomMessageBox(
                                            $"The update check could not complete after {MaxRetries} attempts." &
                                            Environment.NewLine & $"Detail: {detail}" &
                                            Environment.NewLine & Environment.NewLine &
                                            $"You can download and install the latest version manually.",
                                            $"{SharedMethods.AN} Updater",
                                            extraButtonText:="Open downloads page",
                                            extraButtonAction:=Sub()
                                                                   Try
                                                                       Process.Start(New ProcessStartInfo(downloadUrl) With {.UseShellExecute = True})
                                                                   Catch exBrowser As Exception
                                                                       WriteUpdateLog("[CheckAndInstallUpdates] failed to open browser", exBrowser)
                                                                   End Try
                                                               End Sub,
                                            CloseAfterExtra:=True)
                                    End Sub)
                            Else
                                WriteUpdateLog("[CheckAndInstallUpdates] MainControl unavailable, manual fallback suppressed")
                            End If
                        End If
                    End If

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
                            MainControl.Invoke(
                                Sub()
                                    If SharedMethods.CanRunIniGovernedUpdate(context, "INI configuration") Then
                                        If Not SharedMethods.CheckForIniUpdates(context, skipAuthorizationGate:=True) Then
                                            UIInvokeMessage(
                                                "No configuration updates available or made",
                                                $"{SharedMethods.AN} INI Updater")
                                        End If

                                        If Not String.IsNullOrWhiteSpace(context.INI_AgentResourcesPath) Then
                                            SharedMethods.CheckForAgentResourceUpdates(context, skipAuthorizationGate:=True)
                                        End If
                                    End If
                                End Sub)
                        Else
                            If SharedMethods.CanRunIniGovernedUpdate(context, "INI configuration") Then
                                If Not SharedMethods.CheckForIniUpdates(context, skipAuthorizationGate:=True) Then
                                    UIInvokeMessage(
                                        "No configuration updates available or made",
                                        $"{SharedMethods.AN} INI Updater")
                                End If

                                If Not String.IsNullOrWhiteSpace(context.INI_AgentResourcesPath) Then
                                    SharedMethods.CheckForAgentResourceUpdates(context, skipAuthorizationGate:=True)
                                End If
                            End If
                        End If
                    Catch iniEx As Exception
                        WriteUpdateLog("[CheckAndInstallUpdates] INI or agent resource update check failed", iniEx)
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
        ''' Shows a two-button Yes/No style dialog with custom button labels, always marshaled to the UI thread.
        ''' </summary>
        ''' <param name="bodyText">Dialog body text.</param>
        ''' <param name="button1Text">Text for the first button.</param>
        ''' <param name="button2Text">Text for the second button.</param>
        ''' <param name="caption">Dialog caption.</param>
        ''' <returns>Dialog result as returned by `SharedMethods.ShowCustomYesNoBox`.</returns>
        Private Shared Function UIInvokeYesNo(bodyText As String, button1Text As String, button2Text As String, caption As String) As Integer
            If MainControl IsNot Nothing AndAlso MainControl.IsHandleCreated Then
                Try
                    Return CInt(MainControl.Invoke(New Func(Of Integer)(
                        Function()
                            NativeMethods.SetForegroundWindow(HostHandle)
                            Return SharedMethods.ShowCustomYesNoBox(bodyText, button1Text, button2Text, caption)
                        End Function)))
                Catch ex As Exception
                    WriteUpdateLog("[UIInvokeYesNo] Invoke failed", ex)
                    Return 0
                End Try
            Else
                ' Fallback: no UI control available — log and return "declined"
                WriteUpdateLog($"[UIInvokeYesNo] MainControl unavailable, auto-declining: {bodyText}")
                Return 0
            End If
        End Function

        ''' <summary>
        ''' Initial delay in milliseconds before performing the first network update check,
        ''' giving the OS network stack time to fully initialize (VPN, proxy, Wi-Fi).
        ''' </summary>
        Private Const StartupDelayMs As Integer = 8000

        ''' <summary>
        ''' Maximum number of inline retries for the synchronous ClickOnce check fallback
        ''' within <see cref="PeriodicCheckForUpdates"/>.
        ''' </summary>
        Private Const MaxPeriodicRetries As Integer = 3

        ''' <summary>
        ''' Performs an automatic/periodic update check on a background thread, honoring the check
        ''' interval and daily retry state. Uses a synchronous ClickOnce check with retry and
        ''' exponential backoff for network-deployed add-ins, or runs a local `.vsto` installer
        ''' check for local-path deployments. All UI interactions are marshaled back to the UI
        ''' thread via <see cref="MainControl"/>.
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

            If checkIntervalInDays = 0 Then Return

            ' Offload all blocking work (startup delay, network checks, retries, installer) to a
            ' background thread so the Office UI thread is never blocked. All dialog/splash calls
            ' already marshal back via MainControl.Invoke.
            System.Threading.Tasks.Task.Run(
                Sub() PeriodicCheckForUpdatesCore(checkIntervalInDays, appname, LocalPath, context))
        End Sub

        ''' <summary>
        ''' Core implementation of the periodic update check. Runs entirely on a background thread.
        ''' </summary>
        Private Shared Sub PeriodicCheckForUpdatesCore(
            checkIntervalInDays As Integer,
            appname As String,
            LocalPath As String,
            context As ISharedContext)

            Try
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

                If IsRunningFromMsi(_appname) Then
                    WriteUpdateLog($"[PeriodicCheck] MSI installation detected app='{_appname}' — ClickOnce/local update check suppressed (silent)")
                ElseIf ApplicationDeployment.IsNetworkDeployed AndAlso String.IsNullOrWhiteSpace(_localPath) Then
                    Dim dep = ApplicationDeployment.CurrentDeployment
                    WriteUpdateLog($"[PeriodicCheck] network-deployed url='{dep.UpdateLocation}' zone='{GetUrlZoneName(dep.UpdateLocation.AbsoluteUri)}'")

                    ' --- Startup delay: give VPN/proxy/Wi-Fi time to initialize ---
                    WriteUpdateLog($"[PeriodicCheck] waiting {StartupDelayMs}ms for network readiness")
                    Thread.Sleep(StartupDelayMs)

                    ' --- Synchronous check with retry + exponential backoff ---
                    Dim updateAvailable As Boolean = False
                    Dim checkSucceeded As Boolean = False
                    Dim lastEx As Exception = Nothing

                    For attempt As Integer = 1 To MaxPeriodicRetries
                        Try
                            updateAvailable = dep.CheckForUpdate()
                            checkSucceeded = True
                            lastEx = Nothing
                            Exit For
                        Catch ex As Exception
                            lastEx = ex
                            WriteUpdateLog($"[PeriodicCheck] CheckForUpdate attempt {attempt}/{MaxPeriodicRetries} failed", ex)
                            If attempt < MaxPeriodicRetries Then
                                Dim delayMs = 1000 * CInt(Math.Pow(2, attempt - 1)) ' 1s, 2s
                                Thread.Sleep(delayMs)
                            End If
                        End Try
                    Next

                    If Not checkSucceeded Then
                        ' All retries exhausted — try lightweight manifest version check as fallback
                        WriteUpdateLog("[PeriodicCheck] all retries exhausted, trying manifest version check", lastEx)

                        Dim remoteVersion = GetRemoteVersionFromManifest(dep.UpdateLocation.AbsoluteUri)
                        If remoteVersion IsNot Nothing AndAlso remoteVersion > dep.CurrentVersion Then
                            ' Update IS available, but CheckForUpdate failed (trust/prereq issue)
                            ' Use interactive VSTOInstaller which shows trust prompt (same as clicking .vsto)
                            WriteUpdateLog($"[PeriodicCheck] manifest shows update: current='{dep.CurrentVersion}' remote='{remoteVersion}' → interactive install")

                            Dim prompt = $"A new version is available (current: {dep.CurrentVersion}, new: {remoteVersion}). Do you want to install it now?"
                            Dim choice = UIInvokePrompt(prompt, $"{SharedMethods.AN} Updater")

                            If choice = 1 Then
                                Dim vstoUrl = dep.UpdateLocation.AbsoluteUri.Replace(".application", ".vsto")
                                WriteUpdateLog($"[PeriodicCheck] user accepted → interactive install '{vstoUrl}'")
                                ShowUpdatingSplash("Updating …")
                                Try
                                    RunVstoInstaller(vstoUrl)
                                Finally
                                    CloseUpdatingSplash()
                                End Try
                                SaveTimestamp(nowDate)
                            Else
                                WriteUpdateLog("[PeriodicCheck] user declined update (manifest fallback)")
                                If _checkIntervalInDays > 0 Then SaveTimestamp(nowDate)
                            End If

                            ' Reset retry state since we successfully determined update status
                            Dim d As Date = Date.Today : Dim c As Integer = 0 : Dim s As Boolean = False
                            SetRetryStateToSettings(d, c, s)
                            SetUpdateFailureCount(0)

                        ElseIf remoteVersion IsNot Nothing Then
                            ' Remote version same or older — no update needed
                            WriteUpdateLog($"[PeriodicCheck] manifest shows no update needed: current='{dep.CurrentVersion}' remote='{remoteVersion}'")
                            If _checkIntervalInDays > 0 Then SaveTimestamp(nowDate)
                            Dim d As Date = Date.Today : Dim c As Integer = 0 : Dim s As Boolean = False
                            SetRetryStateToSettings(d, c, s)
                            SetUpdateFailureCount(0)
                        Else
                            ' Could not even fetch the manifest — real network issue or persistent failure
                            If lastEx IsNot Nothing AndAlso IsTrustNotGranted(lastEx) AndAlso CanShowInteractiveUi() Then
                                Dim appUrl = dep.UpdateLocation.AbsoluteUri
                                Dim vstoUrl = appUrl.Replace(".application", ".vsto")
                                WriteUpdateLog($"[PeriodicCheck] TrustNotGranted → trying interactive VSTOInstaller on '{vstoUrl}'")

                                ShowUpdatingSplash("Updating …")
                                Try
                                    RunVstoInstaller(vstoUrl)
                                Finally
                                    CloseUpdatingSplash()
                                End Try

                                SaveTimestamp(nowDate)
                                Dim d As Date = Date.Today : Dim c As Integer = 0 : Dim s As Boolean = False
                                SetRetryStateToSettings(d, c, s)
                                SetUpdateFailureCount(0)
                            Else
                                ' Increment cumulative failure counter (persists across days/sessions)
                                Dim failCount = GetUpdateFailureCount() + 1
                                SetUpdateFailureCount(failCount)
                                WriteUpdateLog($"[PeriodicCheck] cumulative update failure count={failCount}/{ManualFallbackThreshold}")

                                If failCount >= ManualFallbackThreshold AndAlso CanShowInteractiveUi() Then
                                    ' All automated paths exhausted after repeated sessions — show manual download
                                    WriteUpdateLog("[PeriodicCheck] cumulative threshold reached, offering manual download")
                                    ShowManualUpdateFallback()
                                    If _checkIntervalInDays > 0 Then SaveTimestamp(nowDate)
                                    ' Reset so the user is not spammed every single time going forward
                                    SetUpdateFailureCount(0)
                                Else
                                    ' Below threshold — use existing daily retry/prompt logic
                                    Dim pause As Boolean = RecordCheckFailureAndMaybePrompt(
                                        If(lastEx?.InnerException IsNot Nothing, lastEx.InnerException.Message,
                                           If(lastEx IsNot Nothing, lastEx.Message, "Unknown")))
                                    If pause AndAlso _checkIntervalInDays > 0 Then SaveTimestamp(nowDate)
                                End If
                            End If
                        End If
                    Else
                        ' --- Check succeeded ---
                        If updateAvailable Then
                            Dim localV = dep.CurrentVersion.ToString()
                            Dim remoteV = dep.CheckForDetailedUpdate().AvailableVersion.ToString()
                            WriteUpdateLog($"[PeriodicCheck] update available current='{localV}' new='{remoteV}' url='{dep.UpdateLocation}'")

                            Dim prompt = $"A new version is available (current: {localV}, new: {remoteV}). Do you want to install it now?"
                            Dim choice = UIInvokePrompt(prompt, $"{SharedMethods.AN} Updater")

                            If choice = 1 Then
                                Dim appUrl = dep.UpdateLocation.AbsoluteUri
                                Dim vstoUrl = appUrl.Replace(".application", ".vsto")
                                WriteUpdateLog($"[PeriodicCheck] user accepted → installing '{vstoUrl}'")

                                ShowUpdatingSplash("Updating …")
                                Try
                                    RunVstoInstaller(vstoUrl)
                                Finally
                                    CloseUpdatingSplash()
                                End Try

                                SaveTimestamp(nowDate)
                            Else
                                WriteUpdateLog("[PeriodicCheck] user declined update")
                                If _checkIntervalInDays = -1 Then
                                    SaveTimestamp(nowDate)
                                ElseIf _checkIntervalInDays > 0 Then
                                    Dim postPrompt = $"Do you want to pause update checks for {_checkIntervalInDays} days?"
                                    Dim postChoice = UIInvokePrompt(postPrompt, $"{SharedMethods.AN} Updater")
                                    If postChoice = 1 Then
                                        SaveTimestamp(nowDate)
                                    End If
                                End If
                            End If
                        Else
                            WriteUpdateLog("[PeriodicCheck] no update available")
                            If _checkIntervalInDays > 0 Then
                                SaveTimestamp(nowDate)
                            End If
                        End If

                        ' Reset daily retry state on any successful check
                        Dim day As Date = Date.Today : Dim cnt As Integer = 0 : Dim shown As Boolean = False
                        SetRetryStateToSettings(day, cnt, shown)
                        SetUpdateFailureCount(0)
                    End If

                Else
                    ' --- Local path deployment ---
                    Dim vstoFile = Path.Combine(
                    Environment.ExpandEnvironmentVariables(_localPath),
                    $"{_appname.ToLowerInvariant()}\{SharedMethods.AN3} for {_appname}.vsto")

                    WriteUpdateLog($"[PeriodicCheck] local-deployed vsto='{vstoFile}'")

                    If File.Exists(vstoFile) Then
                        ' Compare local .vsto manifest version against currently installed version
                        Dim remoteVersion = GetVersionFromLocalManifest(vstoFile)
                        Dim currentVersion As Version = Nothing
                        Try
                            If ApplicationDeployment.IsNetworkDeployed Then
                                currentVersion = ApplicationDeployment.CurrentDeployment.CurrentVersion
                            End If
                        Catch
                        End Try

                        If remoteVersion IsNot Nothing AndAlso currentVersion IsNot Nothing AndAlso remoteVersion <= currentVersion Then
                            WriteUpdateLog($"[PeriodicCheck] local vsto version={remoteVersion} is not newer than current={currentVersion}, skipping install")
                        Else
                            Dim remoteStr As String = If(remoteVersion IsNot Nothing, remoteVersion.ToString(), "unknown")
                            Dim currentStr As String = If(currentVersion IsNot Nothing, currentVersion.ToString(), "unknown")
                            WriteUpdateLog($"[PeriodicCheck] local vsto version={remoteStr} current={currentStr} → running installer")
                            ShowUpdatingSplash("Updating …")
                            Try
                                RunVstoInstaller(vstoFile)
                            Finally
                                CloseUpdatingSplash()
                            End Try
                        End If
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
                    SetUpdateFailureCount(0)
                End If

                ' === INI Configuration Updates ===
                ' Only run after a confirmed successful network/local check
                If _context IsNot Nothing Then
                    Try
                        If MainControl IsNot Nothing AndAlso MainControl.InvokeRequired Then
                            MainControl.Invoke(Sub() SharedMethods.CheckForIniUpdates(_context))
                        Else
                            SharedMethods.CheckForIniUpdates(_context)
                        End If
                    Catch iniEx As Exception
                        WriteUpdateLog("[PeriodicCheck] INI update check failed", iniEx)
                    End Try
                End If

            Catch dex As DeploymentException
                WriteUpdateLog("[PeriodicCheck] DeploymentException", dex)
                Dim failCount = GetUpdateFailureCount() + 1
                SetUpdateFailureCount(failCount)
                WriteUpdateLog($"[PeriodicCheck] cumulative update failure count={failCount}/{ManualFallbackThreshold}")
                If failCount >= ManualFallbackThreshold AndAlso CanShowInteractiveUi() Then
                    ShowManualUpdateFallback()
                    SetUpdateFailureCount(0)
                    If _checkIntervalInDays > 0 Then SaveTimestamp(Date.Now)
                Else
                    Dim pause As Boolean = RecordCheckFailureAndMaybePrompt("DeploymentException")
                    If pause AndAlso _checkIntervalInDays > 0 Then SaveTimestamp(Date.Now)
                End If

            Catch ex As Exception
                WriteUpdateLog("[PeriodicCheck] Unexpected Exception", ex)
                Dim failCount = GetUpdateFailureCount() + 1
                SetUpdateFailureCount(failCount)
                WriteUpdateLog($"[PeriodicCheck] cumulative update failure count={failCount}/{ManualFallbackThreshold}")
                If failCount >= ManualFallbackThreshold AndAlso CanShowInteractiveUi() Then
                    ShowManualUpdateFallback()
                    SetUpdateFailureCount(0)
                    If _checkIntervalInDays > 0 Then SaveTimestamp(Date.Now)
                Else
                    Dim pause As Boolean = RecordCheckFailureAndMaybePrompt(ex.Message)
                    If pause AndAlso _checkIntervalInDays > 0 Then SaveTimestamp(Date.Now)
                End If

            Finally
                CloseUpdatingSplash()
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
                        SetUpdateFailureCount(0)
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
                SetUpdateFailureCount(0)

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
        ''' Parses the assemblyIdentity version from a local .vsto manifest file on disk.
        ''' </summary>
        ''' <param name="vstoFilePath">Full path to the local .vsto file.</param>
        ''' <returns>The parsed version, or Nothing if it could not be determined.</returns>
        Private Shared Function GetVersionFromLocalManifest(vstoFilePath As String) As Version
            Try
                Dim doc As New System.Xml.XmlDocument()
                doc.Load(vstoFilePath)

                Dim node As System.Xml.XmlNode = Nothing

                ' Strategy 1: asmv1 namespace
                Dim nsMgr As New System.Xml.XmlNamespaceManager(doc.NameTable)
                nsMgr.AddNamespace("asmv1", "urn:schemas-microsoft-com:asm.v1")
                node = doc.SelectSingleNode("/asmv1:assembly/asmv1:assemblyIdentity", nsMgr)

                ' Strategy 2: default namespace alias
                If node Is Nothing Then
                    nsMgr.AddNamespace("d", "urn:schemas-microsoft-com:asm.v1")
                    node = doc.SelectSingleNode("/d:assembly/d:assemblyIdentity", nsMgr)
                End If

                ' Strategy 3: asmv2 namespace
                If node Is Nothing Then
                    nsMgr.AddNamespace("asmv2", "urn:schemas-microsoft-com:asm.v2")
                    node = doc.SelectSingleNode("/asmv2:assembly/asmv2:assemblyIdentity", nsMgr)
                End If

                ' Strategy 4: namespace-agnostic
                If node Is Nothing Then
                    For Each child As System.Xml.XmlNode In doc.DocumentElement.ChildNodes
                        If child.LocalName = "assemblyIdentity" Then
                            node = child
                            Exit For
                        End If
                    Next
                End If

                If node IsNot Nothing Then
                    Dim verAttr = node.Attributes("version")
                    If verAttr IsNot Nothing Then
                        Dim v As Version = Nothing
                        If Version.TryParse(verAttr.Value, v) Then
                            WriteUpdateLog($"[GetVersionFromLocalManifest] parsed version='{v}' from '{vstoFilePath}'")
                            Return v
                        End If
                    End If
                End If

                WriteUpdateLog($"[GetVersionFromLocalManifest] could not find version in '{vstoFilePath}'")
            Catch ex As Exception
                WriteUpdateLog("[GetVersionFromLocalManifest] failed to parse manifest", ex)
            End Try
            Return Nothing
        End Function

        ''' <summary>
        ''' Downloads the remote ClickOnce .vsto manifest and extracts the version string,
        ''' bypassing the full ClickOnce trust/prerequisite validation that CheckForUpdate() performs.
        ''' If the input URL ends with .application, it is converted to .vsto since only .vsto
        ''' manifests are deployed on the server.
        ''' </summary>
        ''' <param name="applicationUrl">URL to the .application or .vsto manifest file.</param>
        ''' <returns>The remote version string, or Nothing if it could not be determined.</returns>
        Private Shared Function GetRemoteVersionFromManifest(applicationUrl As String) As Version
            ' Only .vsto manifests exist on the server — convert .application URLs
            Dim vstoUrl As String = applicationUrl
            If vstoUrl.EndsWith(".application", StringComparison.OrdinalIgnoreCase) Then
                vstoUrl = vstoUrl.Substring(0, vstoUrl.Length - ".application".Length) & ".vsto"
            End If
            Return TryGetVersionFromManifestUrl(vstoUrl)
        End Function

        ''' <summary>
        ''' Attempts to download and parse a single manifest URL (either .application or .vsto)
        ''' to extract the version from the assemblyIdentity element.
        ''' </summary>
        ''' <param name="manifestUrl">URL to the manifest file.</param>
        ''' <returns>The parsed version, or Nothing if it could not be determined.</returns>
        Private Shared Function TryGetVersionFromManifestUrl(manifestUrl As String) As Version
            Try
                ' .NET Framework defaults to TLS 1.0/1.1 which most modern servers reject.
                ' Ensure TLS 1.2 (and 1.3 if available) is enabled for this request.
                System.Net.ServicePointManager.SecurityProtocol =
                    System.Net.ServicePointManager.SecurityProtocol Or
                    System.Net.SecurityProtocolType.Tls12 Or
                    CType(12288, System.Net.SecurityProtocolType) ' Tls13 = 12288 (not defined in older .NET FW)

                Using client As New System.Net.Http.HttpClient()
                    client.Timeout = TimeSpan.FromSeconds(15)

                    ' Use GetAsync to inspect status code and content type before parsing
                    Dim response = client.GetAsync(manifestUrl).Result
                    Dim statusCode = CInt(response.StatusCode)
                    Dim contentType = response.Content?.Headers?.ContentType?.MediaType
                    Dim xml = response.Content.ReadAsStringAsync().Result

                    WriteUpdateLog($"[GetRemoteVersion] fetched {xml.Length} chars from '{manifestUrl}' " &
                                   $"status={statusCode} contentType='{contentType}' " &
                                   $"preview='{If(xml.Length > 200, xml.Substring(0, 200), xml).Replace(vbCrLf, " ").Replace(vbLf, " ")}'")

                    ' Guard: if server returned error or HTML, don't try to parse as XML
                    If statusCode < 200 OrElse statusCode >= 300 Then
                        WriteUpdateLog($"[GetRemoteVersion] non-success HTTP status {statusCode}, skipping")
                        Return Nothing
                    End If

                    If contentType IsNot Nothing AndAlso contentType.Contains("html") Then
                        WriteUpdateLog("[GetRemoteVersion] server returned HTML instead of XML manifest, skipping")
                        Return Nothing
                    End If

                    If Not xml.TrimStart().StartsWith("<", StringComparison.Ordinal) Then
                        WriteUpdateLog("[GetRemoteVersion] response does not start with '<', not valid XML, skipping")
                        Return Nothing
                    End If

                    Dim doc As New System.Xml.XmlDocument()
                    doc.LoadXml(xml)

                    Dim node As System.Xml.XmlNode = Nothing

                    ' Strategy 1: Try with asmv1 namespace prefix
                    Dim nsMgr As New System.Xml.XmlNamespaceManager(doc.NameTable)
                    nsMgr.AddNamespace("asmv1", "urn:schemas-microsoft-com:asm.v1")
                    node = doc.SelectSingleNode("/asmv1:assembly/asmv1:assemblyIdentity", nsMgr)

                    ' Strategy 2: Try with default namespace alias
                    If node Is Nothing Then
                        nsMgr.AddNamespace("d", "urn:schemas-microsoft-com:asm.v1")
                        node = doc.SelectSingleNode("/d:assembly/d:assemblyIdentity", nsMgr)
                    End If

                    ' Strategy 3: Try asmv2 namespace
                    If node Is Nothing Then
                        nsMgr.AddNamespace("asmv2", "urn:schemas-microsoft-com:asm.v2")
                        node = doc.SelectSingleNode("/asmv2:assembly/asmv2:assemblyIdentity", nsMgr)
                    End If

                    ' Strategy 4: Namespace-agnostic — first assemblyIdentity child of root
                    If node Is Nothing Then
                        For Each child As System.Xml.XmlNode In doc.DocumentElement.ChildNodes
                            If child.LocalName = "assemblyIdentity" Then
                                node = child
                                Exit For
                            End If
                        Next
                    End If

                    If node IsNot Nothing Then
                        Dim verStr = node.Attributes("version")?.Value
                        If verStr IsNot Nothing Then
                            Dim v As Version = Nothing
                            If Version.TryParse(verStr, v) Then
                                WriteUpdateLog($"[GetRemoteVersion] parsed version='{v}' from '{manifestUrl}'")
                                Return v
                            End If
                        End If
                    End If

                    WriteUpdateLog($"[GetRemoteVersion] could not find version in manifest from '{manifestUrl}' " &
                                   $"rootElement='{doc.DocumentElement?.Name}' rootNS='{doc.DocumentElement?.NamespaceURI}' " &
                                   $"childCount={doc.DocumentElement?.ChildNodes.Count}")
                End Using
            Catch ex As Exception
                WriteUpdateLog("[GetRemoteVersion] failed to fetch/parse manifest", ex)
            End Try
            Return Nothing
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
                       "Update completed. It will be active the next time you restart your application (until you do so, your add-in may not work correctly anymore).",
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
                            "Update completed. It will be active the next time you restart your application (until you do so, your add-in may not work correctly anymore).",
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


        ''' <summary>
        ''' Shows a message informing the user that all automatic update methods have failed,
        ''' with an extra button that opens the manual download page in the default browser.
        ''' Includes the current version and environment (Preview/Develop/GA) in the message.
        ''' </summary>
        Private Shared Sub ShowManualUpdateFallback()
            Dim downloadUrl = $"{SharedMethods.AppsUrl}{SharedMethods.UpdateSubUrl}"

            ' Build version info string including environment qualifier
            Dim versionInfo As String = ""
            Dim versionQualifier As String = SharedMethods.VersionQualifier?.Trim()
            Try
                If ApplicationDeployment.IsNetworkDeployed Then
                    versionInfo = ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString()
                End If
            Catch
            End Try
            If String.IsNullOrEmpty(versionInfo) Then
                versionInfo = "unknown"
            End If

            ' Append environment qualifier (e.g., "Preview", "Develop", or nothing for GA)
            Dim qualifier As String = SharedMethods.VersionQualifier?.Trim()
            If Not String.IsNullOrEmpty(qualifier) Then
                versionInfo &= $" (= {qualifier})"
            Else
                versionInfo &= " (= GA)"
            End If

            If String.IsNullOrEmpty(SharedMethods.VersionQualifier) Then
                versionQualifier = "GA"
            End If

            WriteUpdateLog($"[ManualUpdateFallback] showing manual download prompt url='{downloadUrl}' version='{versionInfo}'")

            If MainControl IsNot Nothing AndAlso MainControl.IsHandleCreated Then
                MainControl.Invoke(
                    Sub()
                        NativeMethods.SetForegroundWindow(HostHandle)
                        SharedMethods.ShowCustomMessageBox(
                            $"The automatic update could not be completed despite multiple attempts. " &
                            $"You should try to manually update your add-in to make sure you have the latest version." &
                            Environment.NewLine & Environment.NewLine &
                            $"Click the button below or visit (you are using the {versionQualifier} version): {downloadUrl}",
                            $"{SharedMethods.AN} Updater",
                            extraButtonText:="Open downloads page",
                            extraButtonAction:=Sub()
                                                   Try
                                                       Process.Start(New ProcessStartInfo(downloadUrl) With {.UseShellExecute = True})
                                                   Catch ex As Exception
                                                       WriteUpdateLog("[ManualUpdateFallback] failed to open browser", ex)
                                                   End Try
                                               End Sub,
                            CloseAfterExtra:=True)
                    End Sub)
            Else
                WriteUpdateLog("[ManualUpdateFallback] MainControl unavailable, fallback suppressed")
            End If
        End Sub


        ''' <summary>
        ''' Gets the cumulative update failure count for the current app from `My.Settings`.
        ''' </summary>
        Private Shared Function GetUpdateFailureCount() As Integer
            Select Case Left(_appname, 4)
                Case "Word" : Return My.Settings.UpdateFailureCountWord
                Case "Exce" : Return My.Settings.UpdateFailureCountExcel
                Case "Outl" : Return My.Settings.UpdateFailureCountOutlook
                Case Else : Return 0
            End Select
        End Function

        ''' <summary>
        ''' Sets the cumulative update failure count for the current app in `My.Settings`.
        ''' </summary>
        Private Shared Sub SetUpdateFailureCount(count As Integer)
            Select Case Left(_appname, 4)
                Case "Word" : My.Settings.UpdateFailureCountWord = count
                Case "Exce" : My.Settings.UpdateFailureCountExcel = count
                Case "Outl" : My.Settings.UpdateFailureCountOutlook = count
            End Select
            Try : My.Settings.Save() : Catch : End Try
        End Sub

        Public Shared Sub TestManualFallback()
            ' Ensure _appname is set so GetUpdateFailureCount/SetUpdateFailureCount don't fail
            If String.IsNullOrEmpty(_appname) Then _appname = "Word"
            ShowManualUpdateFallback()
        End Sub

        ''' <summary>
        ''' Runs the manifest version check test on a background thread (safe to call from a UI button click).
        ''' </summary>
        Public Shared Sub TestManifestVersionCheckAsync()
            If String.IsNullOrEmpty(_appname) Then _appname = "Word"
            System.Threading.Tasks.Task.Run(Sub() TestManifestVersionCheck())
        End Sub

        ''' <summary>
        ''' DEBUG: Tests GetRemoteVersionFromManifest against the live manifest URLs for all three environments.
        ''' Call from Immediate Window: SharedLibrary.SharedLibrary.UpdateHandler.TestManifestVersionCheck()
        ''' </summary>
        Public Shared Sub TestManifestVersionCheck()
            If String.IsNullOrEmpty(_appname) Then _appname = "Word"

            ' Test against all three environment manifest URLs (.application → .vsto fallback)
            Dim urls As New Dictionary(Of String, String) From {
                {"GA", "https://redink.ai/apps/ga/word/Red%20Ink%20for%20Word.application"},
                {"Preview", "https://redink.ai/apps/preview/word/Red%20Ink%20for%20Word.application"},
                {"Develop", "https://redink.ai/apps/develop/word/Red%20Ink%20for%20Word.application"}
            }

            Dim results As New StringBuilder()
            results.AppendLine("=== GetRemoteVersionFromManifest Test Results ===")
            results.AppendLine("  (each test tries .application first, then .vsto fallback)")
            results.AppendLine()

            For Each kvp In urls
                Dim env = kvp.Key
                Dim url = kvp.Value
                Try
                    Dim version = GetRemoteVersionFromManifest(url)
                    If version IsNot Nothing Then
                        results.AppendLine($"  {env}: {version}  ✓")
                    Else
                        results.AppendLine($"  {env}: Nothing returned  ✗")
                    End If
                Catch ex As Exception
                    results.AppendLine($"  {env}: EXCEPTION — {ex.Message}  ✗")
                End Try
            Next

            ' Also test with the actual deployed URL if network-deployed
            Try
                If ApplicationDeployment.IsNetworkDeployed Then
                    Dim dep = ApplicationDeployment.CurrentDeployment
                    Dim liveUrl = dep.UpdateLocation.AbsoluteUri
                    Dim liveVersion = GetRemoteVersionFromManifest(liveUrl)
                    results.AppendLine()
                    results.AppendLine($"  Live deployment URL: {liveUrl}")
                    results.AppendLine($"  Local version:  {dep.CurrentVersion}")
                    results.AppendLine($"  Remote version: {If(liveVersion?.ToString(), "Nothing")}")
                    If liveVersion IsNot Nothing Then
                        results.AppendLine($"  Update available: {liveVersion > dep.CurrentVersion}")
                    End If
                Else
                    results.AppendLine()
                    results.AppendLine("  (Not network-deployed — skipped live deployment test)")
                End If
            Catch ex As Exception
                results.AppendLine($"  Live test error: {ex.Message}")
            End Try

            Dim output = results.ToString()
            WriteUpdateLog(output)
            UIInvokeMessage(output, $"{SharedMethods.AN} Updater — Manifest Test")
        End Sub

    End Class


End Namespace
