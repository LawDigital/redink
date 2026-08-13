' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: OfficeWindowWatchdog.vb
' Purpose:
'   Minimal, ANOMALY-ONLY Office window watchdog. During a healthy Office session
'   it produces ZERO log entries. It runs independently inside each Office add-in
'   process (OUTLOOK.EXE, WINWORD.EXE, EXCEL.EXE, POWERPNT.EXE) on a background
'   System.Threading.Timer (~1.5s) and uses Win32 ONLY from the timer thread. It
'   never calls Office COM and never repairs any condition (purely diagnostic).
'
'   It answers one key question: did code in one Office host obtain/use the HWND
'   of another Office host (particularly as a modal owner), and did that leave the
'   main Office HWND disabled?
'
'   Anomaly records are written through RiCrashLogger.LogMarker, so they honour the
'   existing diagnostics on/off switch, session metadata and file rotation.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary

    Public NotInheritable Class OfficeWindowWatchdog

        Private Sub New()
        End Sub

        ' --- Anomaly type names -------------------------------------------------

        Private Const EventMainWindowDisabled As System.String = "OFFICE_MAIN_WINDOW_DISABLED"
        Private Const EventCrossHostDialogOwner As System.String = "CROSS_HOST_DIALOG_OWNER"
        Private Const EventCrossHostWindowOwnership As System.String = "CROSS_HOST_WINDOW_OWNERSHIP"
        Private Const EventCrossProcessHwndUse As System.String = "CROSS_PROCESS_HWND_USE"
        Private Const EventMainHwndInvalid As System.String = "MAIN_HWND_INVALID"

        ' --- Polling configuration ---------------------------------------------

        Private Const PollIntervalMilliseconds As System.Int32 = 1500

        ' Optional second record if the main window stays disabled for an unusually
        ' long time (diagnostic only; still no repeated per-cycle logging).
        Private Shared ReadOnly LongDisabledThreshold As System.TimeSpan =
            System.TimeSpan.FromSeconds(20)

        ' --- Win32 constants ----------------------------------------------------

        Private Const GW_OWNER As System.UInt32 = 4UI

        ' --- State --------------------------------------------------------------

        Private Shared ReadOnly StateLock As New System.Object()
        Private Shared WatchdogTimer As System.Threading.Timer
        Private Shared Started As System.Boolean

        ' Reentrancy guard so overlapping/slow callbacks cannot pile up.
        Private Shared CallbackInProgress As System.Int32

        ' Latch: true once we have already reported the current disabled episode, so
        ' we do not log the same disabled state on every polling cycle.
        Private Shared DisabledLatched As System.Boolean
        Private Shared DisabledSinceUtc As System.DateTime
        Private Shared LongDisabledReported As System.Boolean
        Private Shared MainHwndInvalidLatched As System.Boolean

        ' =====================================================================
        ' Lifecycle
        ' =====================================================================

        ''' <summary>
        ''' Starts the background watchdog. The main Office HWND is read on each
        ''' cycle from UpdateHandler.HostHandle (captured during normal add-in
        ''' initialization), so it is safe to start before the handle is available.
        ''' </summary>
        Public Shared Sub StartWatchdog()

            SyncLock StateLock

                If Started Then
                    Return
                End If

                Try

                    WatchdogTimer =
                        New System.Threading.Timer(
                            AddressOf TimerCallback,
                            Nothing,
                            PollIntervalMilliseconds,
                            PollIntervalMilliseconds)

                    Started = True

                Catch ex As System.Exception

                    ' Never throw into the host; watchdog simply stays inactive.
                    WatchdogTimer = Nothing
                    Started = False

                End Try

            End SyncLock

        End Sub

        ''' <summary>
        ''' Stops the background watchdog and releases the timer.
        ''' </summary>
        Public Shared Sub StopWatchdog()

            Dim timerToDispose As System.Threading.Timer = Nothing

            SyncLock StateLock

                timerToDispose = WatchdogTimer
                WatchdogTimer = Nothing
                Started = False

            End SyncLock

            If timerToDispose IsNot Nothing Then

                Try
                    timerToDispose.Dispose()
                Catch ex As System.Exception
                End Try

            End If

        End Sub

        ' =====================================================================
        ' Timer callback (Win32 only; never calls Office COM; never repairs)
        ' =====================================================================

        Private Shared Sub TimerCallback(ByVal state As System.Object)

            ' Skip if a previous cycle is still running.
            If System.Threading.Interlocked.CompareExchange(
                CallbackInProgress, 1, 0) <> 0 Then

                Return

            End If

            Try

                Dim mainHwnd As System.IntPtr = UpdateHandler.HostHandle

                ' Handle not captured yet: nothing to observe. No logging.
                If mainHwnd = System.IntPtr.Zero Then
                    Return
                End If

                ' If the captured handle is no longer a valid window, report once.
                If Not IsWindow(mainHwnd) Then

                    If Not MainHwndInvalidLatched Then

                        MainHwndInvalidLatched = True

                        RiCrashLogger.LogMarker(
                            EventMainHwndInvalid,
                            "MainHWND=" & FormatHandle(mainHwnd) &
                            System.Environment.NewLine &
                            "The previously captured Office main HWND is no longer a valid window.")

                    End If

                    Return

                End If

                MainHwndInvalidLatched = False

                Dim enabled As System.Boolean = IsWindowEnabled(mainHwnd)

                If enabled Then

                    ' Healthy transition back to enabled: reset the latch so a later,
                    ' independent occurrence can be logged again. No log on recovery.
                    DisabledLatched = False
                    LongDisabledReported = False
                    Return

                End If

                ' The main Office window is disabled.
                If Not DisabledLatched Then

                    ' Transition Enabled=True -> Enabled=False: log exactly once and
                    ' capture the richer one-time snapshot.
                    DisabledLatched = True
                    DisabledSinceUtc = System.DateTime.UtcNow
                    LongDisabledReported = False

                    RiCrashLogger.LogMarker(
                        EventMainWindowDisabled,
                        BuildDisabledSnapshot(mainHwnd))

                    ' Only enumerate windows AFTER an anomaly condition is detected.
                    DetectCrossHostOwnership(mainHwnd)

                Else

                    ' Already reported this episode. Optionally emit ONE additional
                    ' record if it remains disabled for an unusually long time.
                    If Not LongDisabledReported AndAlso
                       (System.DateTime.UtcNow - DisabledSinceUtc) >= LongDisabledThreshold Then

                        LongDisabledReported = True

                        RiCrashLogger.LogMarker(
                            EventMainWindowDisabled,
                            "State=STILL_DISABLED" & System.Environment.NewLine &
                            "MainHWND=" & FormatHandle(mainHwnd) & System.Environment.NewLine &
                            "DisabledForSeconds=" &
                            CInt((System.DateTime.UtcNow - DisabledSinceUtc).TotalSeconds).
                                ToString(System.Globalization.CultureInfo.InvariantCulture) &
                            System.Environment.NewLine &
                            BuildDisabledSnapshot(mainHwnd))

                    End If

                End If

            Catch ex As System.Exception

                ' The watchdog must never interfere with the host. Fail silently.

            Finally

                System.Threading.Interlocked.Exchange(CallbackInProgress, 0)

            End Try

        End Sub

        ' =====================================================================
        ' Public inspection hooks (called from the UI thread, not the timer)
        ' =====================================================================

        ''' <summary>
        ''' Inspects the owner about to be passed to Form.ShowDialog(owner). Logs
        ''' CROSS_HOST_DIALOG_OWNER only when the owner window belongs to a different
        ''' process than the current one. Healthy same-host owners are never logged.
        ''' </summary>
        Public Shared Sub InspectDialogOwner(
            ByVal owner As System.Windows.Forms.IWin32Window,
            ByVal dialogType As System.String,
            ByVal callerMethod As System.String)

            If owner Is Nothing Then
                Return
            End If

            Try

                Dim ownerHwnd As System.IntPtr

                Try
                    ownerHwnd = owner.Handle
                Catch ex As System.Exception
                    Return
                End Try

                If ownerHwnd = System.IntPtr.Zero Then
                    Return
                End If

                Dim ownerPid As System.Int32 = 0
                GetWindowThreadProcessId(ownerHwnd, ownerPid)

                If ownerPid = 0 Then
                    Return
                End If

                Dim currentPid As System.Int32 = GetCurrentProcessIdSafe()

                ' Same host process => healthy. Do NOT log.
                If ownerPid = currentPid Then
                    Return
                End If

                Dim details As New System.Text.StringBuilder()

                details.AppendLine("CurrentProcess=" & GetProcessNameById(currentPid))
                details.AppendLine("CurrentPID=" & currentPid.ToString(
                    System.Globalization.CultureInfo.InvariantCulture))
                details.AppendLine("Dialog=" & SafeText(dialogType))
                details.AppendLine("Caller=" & SafeText(callerMethod))
                details.AppendLine("OwnerHWND=" & FormatHandle(ownerHwnd))
                details.AppendLine("OwnerPID=" & ownerPid.ToString(
                    System.Globalization.CultureInfo.InvariantCulture))
                details.AppendLine("OwnerProcess=" & GetProcessNameById(ownerPid))
                details.AppendLine("OwnerClass=" & GetWindowClass(ownerHwnd))
                details.AppendLine("OwnerTitle=" & GetWindowTitle(ownerHwnd))
                details.AppendLine("OwnerEnabled=" & IsWindowEnabled(ownerHwnd).ToString(
                    System.Globalization.CultureInfo.InvariantCulture))
                details.AppendLine("OwnerVisible=" & IsWindowVisible(ownerHwnd).ToString(
                    System.Globalization.CultureInfo.InvariantCulture))
                details.AppendLine()
                details.AppendLine("StackTrace=")
                details.AppendLine(System.Environment.StackTrace)

                RiCrashLogger.LogMarker(
                    EventCrossHostDialogOwner,
                    details.ToString())

            Catch ex As System.Exception

                ' Diagnostic hook must never interfere with the host.

            End Try

        End Sub

        ''' <summary>
        ''' Inspects a cached/stored HWND immediately before it is used as a modal
        ''' owner, parent, or argument to an ownership/modal Win32 operation. Logs
        ''' CROSS_PROCESS_HWND_USE only when the HWND belongs to a different process.
        ''' </summary>
        Public Shared Sub InspectCachedHwndUse(
            ByVal hwnd As System.IntPtr,
            ByVal intendedOperation As System.String)

            If hwnd = System.IntPtr.Zero Then
                Return
            End If

            Try

                Dim hwndPid As System.Int32 = 0
                GetWindowThreadProcessId(hwnd, hwndPid)

                If hwndPid = 0 Then
                    Return
                End If

                Dim currentPid As System.Int32 = GetCurrentProcessIdSafe()

                ' Same process => healthy. Do NOT log.
                If hwndPid = currentPid Then
                    Return
                End If

                Dim details As New System.Text.StringBuilder()

                details.AppendLine("CurrentProcess=" & GetProcessNameById(currentPid))
                details.AppendLine("CurrentPID=" & currentPid.ToString(
                    System.Globalization.CultureInfo.InvariantCulture))
                details.AppendLine("Operation=" & SafeText(intendedOperation))
                details.AppendLine("HWND=" & FormatHandle(hwnd))
                details.AppendLine("HwndPID=" & hwndPid.ToString(
                    System.Globalization.CultureInfo.InvariantCulture))
                details.AppendLine("HwndProcess=" & GetProcessNameById(hwndPid))
                details.AppendLine("HwndClass=" & GetWindowClass(hwnd))
                details.AppendLine("HwndTitle=" & GetWindowTitle(hwnd))
                details.AppendLine()
                details.AppendLine("StackTrace=")
                details.AppendLine(System.Environment.StackTrace)

                RiCrashLogger.LogMarker(
                    EventCrossProcessHwndUse,
                    details.ToString())

            Catch ex As System.Exception

                ' Diagnostic hook must never interfere with the host.

            End Try

        End Sub

        ' =====================================================================
        ' Snapshot builders (only invoked once an anomaly has been detected)
        ' =====================================================================

        Private Shared Function BuildDisabledSnapshot(
            ByVal mainHwnd As System.IntPtr) As System.String

            Dim result As New System.Text.StringBuilder()

            Try

                Dim currentPid As System.Int32 = GetCurrentProcessIdSafe()

                result.AppendLine("SnapshotUtc=" &
                    System.DateTime.UtcNow.ToString(
                        "o", System.Globalization.CultureInfo.InvariantCulture))
                result.AppendLine("Host=" & GetProcessNameById(currentPid))
                result.AppendLine("PID=" & currentPid.ToString(
                    System.Globalization.CultureInfo.InvariantCulture))
                result.AppendLine("MainHWND=" & FormatHandle(mainHwnd))
                result.AppendLine("MainClass=" & GetWindowClass(mainHwnd))
                result.AppendLine("MainTitle=" & GetWindowTitle(mainHwnd))
                result.AppendLine("MainEnabled=" & IsWindowEnabled(mainHwnd).ToString(
                    System.Globalization.CultureInfo.InvariantCulture))
                result.AppendLine("MainVisible=" & IsWindowVisible(mainHwnd).ToString(
                    System.Globalization.CultureInfo.InvariantCulture))

                ' Foreground window + process.
                Dim fg As System.IntPtr = GetForegroundWindow()
                Dim fgPid As System.Int32 = 0
                Dim fgOwner As System.IntPtr = System.IntPtr.Zero

                result.AppendLine("ForegroundHWND=" & FormatHandle(fg))

                If fg <> System.IntPtr.Zero Then
                    GetWindowThreadProcessId(fg, fgPid)
                    fgOwner = GetWindow(fg, GW_OWNER)

                    result.AppendLine("ForegroundPID=" & fgPid.ToString(
                        System.Globalization.CultureInfo.InvariantCulture))
                    result.AppendLine("ForegroundProcess=" & GetProcessNameById(fgPid))
                    result.AppendLine("ForegroundClass=" & GetWindowClass(fg))
                    result.AppendLine("ForegroundTitle=" & GetWindowTitle(fg))
                    result.AppendLine("ForegroundOwnerHWND=" & FormatHandle(fgOwner))

                    If fgOwner <> System.IntPtr.Zero Then

                        Dim fgOwnerPid As System.Int32 = 0
                        GetWindowThreadProcessId(fgOwner, fgOwnerPid)

                        result.AppendLine("ForegroundOwnerPID=" & fgOwnerPid.ToString(
                            System.Globalization.CultureInfo.InvariantCulture))
                        result.AppendLine("ForegroundOwnerProcess=" & GetProcessNameById(fgOwnerPid))
                        result.AppendLine("ForegroundOwnerClass=" & GetWindowClass(fgOwner))
                        result.AppendLine("ForegroundOwnerTitle=" & GetWindowTitle(fgOwner))

                    End If

                Else

                    result.AppendLine("ForegroundOwnerHWND=0x0")

                End If

                result.AppendLine("ForegroundOwnedByMain=" & (fgOwner = mainHwnd).ToString(
                    System.Globalization.CultureInfo.InvariantCulture))

                result.AppendLine("SameProcessForegroundWithoutMainOwnership=" &
                    (fg <> System.IntPtr.Zero AndAlso
                     fg <> mainHwnd AndAlso
                     fgPid = currentPid AndAlso
                     fgOwner <> mainHwnd).ToString(
                        System.Globalization.CultureInfo.InvariantCulture))

                ' Last active popup for the main window.
                Dim popup As System.IntPtr = GetLastActivePopup(mainHwnd)
                result.AppendLine("LastActivePopupHWND=" & FormatHandle(popup))

                If popup <> System.IntPtr.Zero AndAlso popup <> mainHwnd Then
                    Dim popupPid As System.Int32 = 0
                    GetWindowThreadProcessId(popup, popupPid)
                    result.AppendLine("LastActivePopupPID=" & popupPid.ToString(
                        System.Globalization.CultureInfo.InvariantCulture))
                    result.AppendLine("LastActivePopupProcess=" & GetProcessNameById(popupPid))
                    result.AppendLine("LastActivePopupClass=" & GetWindowClass(popup))
                    result.AppendLine("LastActivePopupTitle=" & GetWindowTitle(popup))
                    result.AppendLine("LastActivePopupEnabled=" & IsWindowEnabled(popup).ToString(
                        System.Globalization.CultureInfo.InvariantCulture))
                    result.AppendLine("LastActivePopupVisible=" & IsWindowVisible(popup).ToString(
                        System.Globalization.CultureInfo.InvariantCulture))
                End If

                ' Windows owned by the main HWND.
                result.AppendLine()
                result.AppendLine("----- WINDOWS OWNED BY MAIN HWND -----")

                For Each owned As System.IntPtr In EnumerateOwnedWindows(mainHwnd)

                    Dim ownedPid As System.Int32 = 0
                    GetWindowThreadProcessId(owned, ownedPid)

                    result.AppendLine(
                        "HWND=" & FormatHandle(owned) &
                        " | PID=" & ownedPid.ToString(
                            System.Globalization.CultureInfo.InvariantCulture) &
                        " | Process=" & GetProcessNameById(ownedPid) &
                        " | OwnerHWND=" & FormatHandle(GetWindow(owned, GW_OWNER)) &
                        " | OwnedByMain=" & (GetWindow(owned, GW_OWNER) = mainHwnd).ToString(
                            System.Globalization.CultureInfo.InvariantCulture) &
                        " | Foreground=" & (owned = fg).ToString(
                            System.Globalization.CultureInfo.InvariantCulture) &
                        " | Class=" & GetWindowClass(owned) &
                        " | Title=" & GetWindowTitle(owned) &
                        " | Enabled=" & IsWindowEnabled(owned).ToString(
                            System.Globalization.CultureInfo.InvariantCulture) &
                        " | Visible=" & IsWindowVisible(owned).ToString(
                            System.Globalization.CultureInfo.InvariantCulture))

                Next

                result.AppendLine()
                result.AppendLine("----- RELEVANT SAME-PROCESS TOP-LEVEL WINDOWS -----")

                For Each sameProcessWindow As System.IntPtr In
                    EnumerateRelevantSameProcessTopLevelWindows(
                        currentPid,
                        mainHwnd,
                        fg,
                        popup)

                    Dim sameProcessOwner As System.IntPtr =
                        GetWindow(sameProcessWindow, GW_OWNER)

                    result.AppendLine(
                        "HWND=" & FormatHandle(sameProcessWindow) &
                        " | PID=" & currentPid.ToString(
                            System.Globalization.CultureInfo.InvariantCulture) &
                        " | Process=" & GetProcessNameById(currentPid) &
                        " | OwnerHWND=" & FormatHandle(sameProcessOwner) &
                        " | OwnedByMain=" & (sameProcessOwner = mainHwnd).ToString(
                            System.Globalization.CultureInfo.InvariantCulture) &
                        " | Foreground=" & (sameProcessWindow = fg).ToString(
                            System.Globalization.CultureInfo.InvariantCulture) &
                        " | LastActivePopup=" & (sameProcessWindow = popup).ToString(
                            System.Globalization.CultureInfo.InvariantCulture) &
                        " | Enabled=" & IsWindowEnabled(sameProcessWindow).ToString(
                            System.Globalization.CultureInfo.InvariantCulture) &
                        " | Visible=" & IsWindowVisible(sameProcessWindow).ToString(
                            System.Globalization.CultureInfo.InvariantCulture) &
                        " | Class=" & GetWindowClass(sameProcessWindow) &
                        " | Title=" & GetWindowTitle(sameProcessWindow))

                Next

            Catch ex As System.Exception

                result.AppendLine("SnapshotError=" & ex.Message)

            End Try

            Return result.ToString()

        End Function

        ''' <summary>
        ''' After a disabled state is detected, enumerates candidate top-level windows
        ''' and reports CROSS_HOST_WINDOW_OWNERSHIP when the Office main HWND is owned
        ''' or blocked by a window belonging to a DIFFERENT process.
        ''' </summary>
        Private Shared Sub DetectCrossHostOwnership(ByVal mainHwnd As System.IntPtr)

            Try

                Dim currentPid As System.Int32 = GetCurrentProcessIdSafe()

                Dim candidates As New System.Collections.Generic.List(Of System.IntPtr)()

                ' The last active popup is the most likely blocker.
                Dim popup As System.IntPtr = GetLastActivePopup(mainHwnd)
                If popup <> System.IntPtr.Zero AndAlso popup <> mainHwnd Then
                    candidates.Add(popup)
                End If

                ' Any top-level window owned by the main HWND.
                For Each owned As System.IntPtr In EnumerateOwnedWindows(mainHwnd)
                    If Not candidates.Contains(owned) Then
                        candidates.Add(owned)
                    End If
                Next

                For Each candidate As System.IntPtr In candidates

                    Dim candidatePid As System.Int32 = 0
                    GetWindowThreadProcessId(candidate, candidatePid)

                    If candidatePid = 0 OrElse candidatePid = currentPid Then
                        Continue For
                    End If

                    ' Cross-process owner/popup of an Office main window: high priority.
                    Dim details As New System.Text.StringBuilder()

                    details.AppendLine("CurrentProcess=" & GetProcessNameById(currentPid))
                    details.AppendLine("CurrentPID=" & currentPid.ToString(
                        System.Globalization.CultureInfo.InvariantCulture))
                    details.AppendLine("MainHWND=" & FormatHandle(mainHwnd))
                    details.AppendLine("MainClass=" & GetWindowClass(mainHwnd))
                    details.AppendLine("MainTitle=" & GetWindowTitle(mainHwnd))
                    details.AppendLine("BlockingHWND=" & FormatHandle(candidate))
                    details.AppendLine("BlockingPID=" & candidatePid.ToString(
                        System.Globalization.CultureInfo.InvariantCulture))
                    details.AppendLine("BlockingProcess=" & GetProcessNameById(candidatePid))
                    details.AppendLine("BlockingClass=" & GetWindowClass(candidate))
                    details.AppendLine("BlockingTitle=" & GetWindowTitle(candidate))
                    details.AppendLine("BlockingOwnerHWND=" &
                        FormatHandle(GetWindow(candidate, GW_OWNER)))
                    details.AppendLine("BlockingEnabled=" & IsWindowEnabled(candidate).ToString(
                        System.Globalization.CultureInfo.InvariantCulture))
                    details.AppendLine("BlockingVisible=" & IsWindowVisible(candidate).ToString(
                        System.Globalization.CultureInfo.InvariantCulture))

                    RiCrashLogger.LogMarker(
                        EventCrossHostWindowOwnership,
                        details.ToString())

                Next

            Catch ex As System.Exception

                ' Never interfere with the host.

            End Try

        End Sub

        ''' <summary>
        ''' Enumerates top-level windows whose GW_OWNER is the supplied main HWND.
        ''' Only called after an anomaly condition has already been detected.
        ''' </summary>
        Private Shared Function EnumerateOwnedWindows(
            ByVal mainHwnd As System.IntPtr) As System.Collections.Generic.List(Of System.IntPtr)

            Dim owned As New System.Collections.Generic.List(Of System.IntPtr)()

            Try

                Dim callback As EnumWindowsProc =
                    Function(hwnd As System.IntPtr, lParam As System.IntPtr) As System.Boolean

                        Try
                            If GetWindow(hwnd, GW_OWNER) = mainHwnd Then
                                owned.Add(hwnd)
                            End If
                        Catch ex As System.Exception
                        End Try

                        Return True

                    End Function

                EnumWindows(callback, System.IntPtr.Zero)

                ' Keep the delegate alive for the duration of the call.
                System.GC.KeepAlive(callback)

            Catch ex As System.Exception
            End Try

            Return owned

        End Function

        ''' <summary>
        ''' Enumerates relevant same-process top-level windows once an anomaly has
        ''' already been detected. This is intentionally broader than
        ''' EnumerateOwnedWindows so we can diagnose same-host cases where the
        ''' foreground dialog/form is visible but not properly owned by the main
        ''' Office HWND.
        ''' </summary>
        Private Shared Function EnumerateRelevantSameProcessTopLevelWindows(
            ByVal currentPid As System.Int32,
            ByVal mainHwnd As System.IntPtr,
            ByVal foregroundHwnd As System.IntPtr,
            ByVal popupHwnd As System.IntPtr) As System.Collections.Generic.List(Of System.IntPtr)

            Dim windows As New System.Collections.Generic.List(Of System.IntPtr)()

            Try

                Dim callback As EnumWindowsProc =
                    Function(hwnd As System.IntPtr, lParam As System.IntPtr) As System.Boolean

                        Try

                            Dim hwndPid As System.Int32 = 0
                            GetWindowThreadProcessId(hwnd, hwndPid)

                            If hwndPid <> currentPid Then
                                Return True
                            End If

                            Dim owner As System.IntPtr =
                                GetWindow(hwnd, GW_OWNER)

                            Dim includeWindow As System.Boolean =
                                hwnd = mainHwnd OrElse
                                hwnd = foregroundHwnd OrElse
                                hwnd = popupHwnd OrElse
                                owner = mainHwnd OrElse
                                IsWindowVisible(hwnd)

                            If includeWindow AndAlso
                               Not windows.Contains(hwnd) Then

                                windows.Add(hwnd)

                            End If

                        Catch ex As System.Exception
                        End Try

                        Return True

                    End Function

                EnumWindows(callback, System.IntPtr.Zero)

                System.GC.KeepAlive(callback)

            Catch ex As System.Exception
            End Try

            Return windows

        End Function

        ' =====================================================================
        ' Helpers
        ' =====================================================================

        Private Shared Function GetCurrentProcessIdSafe() As System.Int32

            Try
                Return System.Diagnostics.Process.GetCurrentProcess().Id
            Catch ex As System.Exception
                Return 0
            End Try

        End Function

        Private Shared Function GetProcessNameById(ByVal pid As System.Int32) As System.String

            If pid = 0 Then
                Return System.String.Empty
            End If

            Try

                Using process As System.Diagnostics.Process =
                    System.Diagnostics.Process.GetProcessById(pid)

                    Return process.ProcessName

                End Using

            Catch ex As System.Exception
                Return "[unknown pid " & pid.ToString(
                    System.Globalization.CultureInfo.InvariantCulture) & "]"
            End Try

        End Function

        Private Shared Function GetWindowClass(ByVal hwnd As System.IntPtr) As System.String

            If hwnd = System.IntPtr.Zero Then
                Return System.String.Empty
            End If

            Try
                Dim buffer As New System.Text.StringBuilder(256)
                GetClassName(hwnd, buffer, buffer.Capacity)
                Return buffer.ToString()
            Catch ex As System.Exception
                Return System.String.Empty
            End Try

        End Function

        Private Shared Function GetWindowTitle(ByVal hwnd As System.IntPtr) As System.String

            If hwnd = System.IntPtr.Zero Then
                Return System.String.Empty
            End If

            Try
                Dim length As System.Int32 = GetWindowTextLength(hwnd)
                If length <= 0 Then
                    Return System.String.Empty
                End If
                Dim buffer As New System.Text.StringBuilder(length + 1)
                GetWindowText(hwnd, buffer, buffer.Capacity)
                Return SafeText(buffer.ToString())
            Catch ex As System.Exception
                Return System.String.Empty
            End Try

        End Function

        Private Shared Function FormatHandle(ByVal hwnd As System.IntPtr) As System.String
            Return "0x" & hwnd.ToInt64().ToString(
                "X", System.Globalization.CultureInfo.InvariantCulture)
        End Function

        Private Shared Function SafeText(ByVal value As System.String) As System.String

            If value Is Nothing Then
                Return System.String.Empty
            End If

            Return value.
                Replace(System.Convert.ToString(System.Convert.ToChar(13)), " ").
                Replace(System.Convert.ToString(System.Convert.ToChar(10)), " ").
                Replace(System.Convert.ToString(System.Convert.ToChar(9)), " ")

        End Function

        ' =====================================================================
        ' Win32 declarations (self-contained; used from the timer thread)
        ' =====================================================================

        Private Delegate Function EnumWindowsProc(
            ByVal hwnd As System.IntPtr,
            ByVal lParam As System.IntPtr) As System.Boolean

        <System.Runtime.InteropServices.DllImport("user32.dll")>
        Private Shared Function IsWindow(ByVal hWnd As System.IntPtr) As System.Boolean
        End Function

        <System.Runtime.InteropServices.DllImport("user32.dll")>
        Private Shared Function IsWindowEnabled(ByVal hWnd As System.IntPtr) As System.Boolean
        End Function

        <System.Runtime.InteropServices.DllImport("user32.dll")>
        Private Shared Function IsWindowVisible(ByVal hWnd As System.IntPtr) As System.Boolean
        End Function

        <System.Runtime.InteropServices.DllImport("user32.dll")>
        Private Shared Function GetForegroundWindow() As System.IntPtr
        End Function

        <System.Runtime.InteropServices.DllImport("user32.dll")>
        Private Shared Function GetLastActivePopup(ByVal hWnd As System.IntPtr) As System.IntPtr
        End Function

        <System.Runtime.InteropServices.DllImport("user32.dll")>
        Private Shared Function GetWindow(
            ByVal hWnd As System.IntPtr,
            ByVal uCmd As System.UInt32) As System.IntPtr
        End Function

        <System.Runtime.InteropServices.DllImport("user32.dll")>
        Private Shared Function GetWindowThreadProcessId(
            ByVal hWnd As System.IntPtr,
            ByRef lpdwProcessId As System.Int32) As System.Int32
        End Function

        <System.Runtime.InteropServices.DllImport("user32.dll")>
        Private Shared Function EnumWindows(
            ByVal lpEnumFunc As EnumWindowsProc,
            ByVal lParam As System.IntPtr) As System.Boolean
        End Function

        <System.Runtime.InteropServices.DllImport("user32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Auto)>
        Private Shared Function GetClassName(
            ByVal hWnd As System.IntPtr,
            ByVal lpClassName As System.Text.StringBuilder,
            ByVal nMaxCount As System.Int32) As System.Int32
        End Function

        <System.Runtime.InteropServices.DllImport("user32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Auto)>
        Private Shared Function GetWindowText(
            ByVal hWnd As System.IntPtr,
            ByVal lpString As System.Text.StringBuilder,
            ByVal nMaxCount As System.Int32) As System.Int32
        End Function

        <System.Runtime.InteropServices.DllImport("user32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Auto)>
        Private Shared Function GetWindowTextLength(ByVal hWnd As System.IntPtr) As System.Int32
        End Function

    End Class

End Namespace
