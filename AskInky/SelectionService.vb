' Part of "Red Ink" (Ask Inky)
' Copyright (c) LawDigital AG, Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: SelectionService.vb
' Purpose: Robust capture of the selected text in the foreground application, and optional paste-back.
' Architecture: Windows native host with embedded Office Web-based interface.
' Notes: Red Ink License and Ask Inky third-party attributions are embedded in LICENSE.txt.
' =============================================================================
' Part of Red Ink (AskInky). P18 W367 (DARO UID 550): a chain instead of one blind Ctrl+C.
'  1. UI Automation TextPattern selection of the focused element (no clipboard involved).
'  2. Classic Win32 edit controls: EM_GETSEL + WM_GETTEXT (no clipboard involved).
'  3. Clipboard: byte snapshot, SendInput Ctrl+C, adaptive wait up to 2.5 s, one Ctrl+Insert retry,
'     restore only when nobody else changed the clipboard meanwhile.
'  4. Otherwise a named reason (no selection / elevated target / copying blocked) for the user.
' Only the method and the reason are logged, never the text.
Public NotInheritable Class SelectionResult
    Public Property Text As String = ""
    Public Property Method As String = ""
    Public Property Reason As String = ""
    Public Property SourceWindow As System.IntPtr
    Public Property ProcessName As String = ""
    ' True = the source element is known editable, False = known read-only, Nothing = unknown.
    Public Property Editable As Boolean?
End Class

Public NotInheritable Class SelectionService
    Public Const ReasonNone As String = "none"
    Public Const ReasonElevated As String = "elevated"
    Public Const ReasonCopyBlocked As String = "copyBlocked"
    Public Const ReasonForeground As String = "foregroundChanged"
    Private Const MaxChars As Integer = 2000000

    Public Shared Async Function CaptureAsync(expectedWindow As System.IntPtr) As System.Threading.Tasks.Task(Of SelectionResult)
        Dim result As New SelectionResult With {.SourceWindow = expectedWindow}
        Dim processId As UInteger
        Native.GetWindowThreadProcessId(expectedWindow, processId)
        result.ProcessName = SafeProcessName(processId)
        ' A registered shortcut can fire while Ctrl/Shift/Alt are still held.
        For attempt As Integer = 0 To 39
            If expectedWindow = System.IntPtr.Zero OrElse Native.GetForegroundWindow() <> expectedWindow Then
                result.Reason = ReasonForeground
                Log(result)
                Return result
            End If
            If Not ModifiersDown() Then Exit For
            Await System.Threading.Tasks.Task.Delay(50)
        Next

        ' 1. UI Automation (bounded: a hung target must never block Ask Inky).
        Try
            Dim uia As SelectionResult = Await RunStaAsync(Function() UiaSelection(), 1500)
            If uia IsNot Nothing AndAlso uia.Text.Length > 0 Then
                uia.SourceWindow = expectedWindow : uia.ProcessName = result.ProcessName
                Log(uia)
                Return uia
            End If
            If uia IsNot Nothing Then result.Editable = uia.Editable
        Catch ex As System.Exception
            ClientLog.Write("selection.uia.failed", ex)
        End Try

        ' 2. Classic Win32 edit control.
        Try
            Dim edit As String = Win32EditSelection(expectedWindow)
            If Not String.IsNullOrEmpty(edit) Then
                result.Text = Limit(edit) : result.Method = "win32edit" : result.Editable = True
                Log(result)
                Return result
            End If
        Catch ex As System.Exception
            ClientLog.Write("selection.win32edit.failed", ex)
        End Try

        ' Synthetic input into an elevated window is silently dropped by Windows (UIPI).
        If IsElevated(processId) AndAlso Not IsElevated(CUInt(System.Environment.ProcessId)) Then
            result.Reason = ReasonElevated
            Log(result)
            Return result
        End If

        ' 3. Clipboard with snapshot and adaptive wait.
        If Native.GetForegroundWindow() <> expectedWindow OrElse ModifiersDown() Then
            result.Reason = ReasonForeground
            Log(result)
            Return result
        End If
        Dim copied As CopyOutcome = Await ClipboardCopyAsync(expectedWindow)
        If copied.Text.Length > 0 Then
            result.Text = Limit(copied.Text) : result.Method = "clipboard"
        Else
            result.Reason = If(copied.ClipboardChanged, ReasonCopyBlocked, ReasonNone)
        End If
        Log(result)
        Return result
    End Function

    Private Shared Function ModifiersDown() As Boolean
        Return (CInt(Native.GetAsyncKeyState(Native.VK_CONTROL)) And &H8000) <> 0 OrElse
               (CInt(Native.GetAsyncKeyState(Native.VK_MENU)) And &H8000) <> 0 OrElse
               (CInt(Native.GetAsyncKeyState(Native.VK_SHIFT)) And &H8000) <> 0
    End Function

    Private Shared Function Limit(value As String) As String
        Dim text As String = If(value, "")
        Return If(text.Length > MaxChars, text.Substring(0, MaxChars), text)
    End Function

    Private Shared Sub Log(result As SelectionResult)
        ClientLog.Write("selection.capture method=" & If(result.Method = "", "-", result.Method) &
                        " reason=" & If(result.Reason = "", "-", result.Reason) &
                        " app=" & If(result.ProcessName = "", "-", result.ProcessName) &
                        " chars=" & result.Text.Length.ToString(System.Globalization.CultureInfo.InvariantCulture))
    End Sub

    Private Shared Function SafeProcessName(processId As UInteger) As String
        Try
            Using p As System.Diagnostics.Process = System.Diagnostics.Process.GetProcessById(CInt(processId))
                Dim name As String = System.Text.RegularExpressions.Regex.Replace(p.ProcessName, "[^A-Za-z0-9._-]", "")
                Return name.Substring(0, System.Math.Min(40, name.Length))
            End Using
        Catch
            Return ""
        End Try
    End Function

    ' Runs a function on a fresh STA thread with a deadline (UIA and clipboard calls can hang on a frozen target).
    Friend Shared Function RunStaAsync(Of T)(work As System.Func(Of T), timeoutMs As Integer) As System.Threading.Tasks.Task(Of T)
        Dim completion As New System.Threading.Tasks.TaskCompletionSource(Of T)(System.Threading.Tasks.TaskCreationOptions.RunContinuationsAsynchronously)
        Dim worker As New System.Threading.Thread(Sub()
                                                      Try
                                                          completion.TrySetResult(work())
                                                      Catch ex As System.Exception
                                                          completion.TrySetException(ex)
                                                      End Try
                                                  End Sub) With {.IsBackground = True, .Name = "AskInky.Capture"}
        worker.SetApartmentState(System.Threading.ApartmentState.STA)
        worker.Start()
        Return WaitWithTimeout(completion.Task, timeoutMs)
    End Function

    Private Shared Async Function WaitWithTimeout(Of T)(work As System.Threading.Tasks.Task(Of T), timeoutMs As Integer) As System.Threading.Tasks.Task(Of T)
        Dim finished As System.Threading.Tasks.Task = Await System.Threading.Tasks.Task.WhenAny(work, System.Threading.Tasks.Task.Delay(timeoutMs))
        If finished IsNot work Then
            ClientLog.Write("selection.step.timeout ms=" & timeoutMs.ToString(System.Globalization.CultureInfo.InvariantCulture))
            Return Nothing
        End If
        Return Await work
    End Function

    Private Shared Function UiaSelection() As SelectionResult
        Return UiaSelectionFrom(System.Windows.Automation.AutomationElement.FocusedElement)
    End Function

    Friend Shared Function UiaSelectionFrom(focused As System.Windows.Automation.AutomationElement) As SelectionResult
        If focused Is Nothing Then Return Nothing
        Dim result As New SelectionResult With {.Method = "uia"}
        Dim element As System.Windows.Automation.AutomationElement = focused
        For level As Integer = 0 To 3
            If element Is Nothing Then Exit For
            Dim valueObject As Object = Nothing
            If result.Editable Is Nothing AndAlso element.TryGetCurrentPattern(System.Windows.Automation.ValuePattern.Pattern, valueObject) Then
                result.Editable = Not DirectCast(valueObject, System.Windows.Automation.ValuePattern).Current.IsReadOnly
            End If
            Dim textObject As Object = Nothing
            If element.TryGetCurrentPattern(System.Windows.Automation.TextPattern.Pattern, textObject) Then
                Dim pattern As System.Windows.Automation.TextPattern = DirectCast(textObject, System.Windows.Automation.TextPattern)
                Dim builder As New System.Text.StringBuilder()
                For Each range As System.Windows.Automation.Text.TextPatternRange In pattern.GetSelection()
                    Dim part As String = range.GetText(MaxChars - builder.Length)
                    If Not String.IsNullOrEmpty(part) Then
                        If builder.Length > 0 Then builder.Append(vbLf)
                        builder.Append(part)
                    End If
                    If builder.Length >= MaxChars Then Exit For
                Next
                If builder.ToString().Trim().Length > 0 Then result.Text = builder.ToString()
                Return result
            End If
            element = System.Windows.Automation.TreeWalker.ControlViewWalker.GetParent(element)
        Next
        Return result
    End Function

    Private Shared Function Win32EditSelection(foreground As System.IntPtr) As String
        Dim processId As UInteger
        Dim threadId As UInteger = Native.GetWindowThreadProcessId(foreground, processId)
        Dim info As New NativeCapture.GUITHREADINFO With {.cbSize = System.Runtime.InteropServices.Marshal.SizeOf(Of NativeCapture.GUITHREADINFO)()}
        If Not NativeCapture.GetGUIThreadInfo(threadId, info) OrElse info.hwndFocus = System.IntPtr.Zero Then Return ""
        Return Win32EditText(info.hwndFocus)
    End Function

    Friend Shared Function Win32EditText(focusWindow As System.IntPtr) As String
        Dim info As New NativeCapture.GUITHREADINFO With {.hwndFocus = focusWindow}
        Dim name As New System.Text.StringBuilder(64)
        NativeCapture.GetClassName(info.hwndFocus, name, name.Capacity)
        ' Only the plain EDIT class: RichEdit uses other line-break positions for EM_GETSEL.
        If Not String.Equals(name.ToString(), "Edit", System.StringComparison.OrdinalIgnoreCase) Then Return ""
        Dim packed As System.IntPtr
        If NativeCapture.SendMessageTimeout(info.hwndFocus, NativeCapture.EM_GETSEL, System.IntPtr.Zero, System.IntPtr.Zero, NativeCapture.SMTO_ABORTIFHUNG, 500UI, packed) = System.IntPtr.Zero Then Return ""
        Dim value As Long = packed.ToInt64()
        Dim startPos As Integer = CInt(value And &HFFFFL)
        Dim endPos As Integer = CInt((value >> 16) And &HFFFFL)
        If endPos <= startPos Then Return ""
        Dim lengthResult As System.IntPtr
        If NativeCapture.SendMessageTimeout(info.hwndFocus, NativeCapture.WM_GETTEXTLENGTH, System.IntPtr.Zero, System.IntPtr.Zero, NativeCapture.SMTO_ABORTIFHUNG, 500UI, lengthResult) = System.IntPtr.Zero Then Return ""
        Dim length As Integer = CInt(System.Math.Min(lengthResult.ToInt64(), 65535L))
        ' EM_GETSEL is 16-bit in the packed form: beyond 65535 characters the positions are unreliable.
        If length <= 0 OrElse length >= 65535 OrElse endPos > length Then Return ""
        Dim buffer As New System.Text.StringBuilder(length + 1)
        Dim copiedResult As System.IntPtr
        If NativeCapture.SendMessageTimeoutText(info.hwndFocus, NativeCapture.WM_GETTEXT, New System.IntPtr(length + 1), buffer, NativeCapture.SMTO_ABORTIFHUNG, 500UI, copiedResult) = System.IntPtr.Zero Then Return ""
        Dim text As String = buffer.ToString()
        If endPos > text.Length Then Return ""
        Return text.Substring(startPos, endPos - startPos)
    End Function

    Friend Shared Function IsElevated(processId As UInteger) As Boolean
        Dim own As Boolean = processId = CUInt(System.Environment.ProcessId)
        Dim process As System.IntPtr = If(own, NativeCapture.GetCurrentProcess(), NativeCapture.OpenProcess(NativeCapture.PROCESS_QUERY_LIMITED_INFORMATION, False, processId))
        If process = System.IntPtr.Zero Then Return False
        Dim token As System.IntPtr
        Try
            If Not NativeCapture.OpenProcessToken(process, NativeCapture.TOKEN_QUERY, token) Then
                ' Access denied to the token of an elevated process.
                Return System.Runtime.InteropServices.Marshal.GetLastWin32Error() = 5
            End If
            Dim elevation As Integer
            Dim returned As Integer
            Return NativeCapture.GetTokenInformation(token, NativeCapture.TokenElevationClass, elevation, 4, returned) AndAlso elevation <> 0
        Finally
            If token <> System.IntPtr.Zero Then NativeCapture.CloseHandle(token)
            If Not own Then NativeCapture.CloseHandle(process)
        End Try
    End Function

    Private NotInheritable Class CopyOutcome
        Public Property Text As String = ""
        Public Property ClipboardChanged As Boolean
    End Class

    Friend Shared Sub SendChord(modifier As UShort, key As UShort, extended As Boolean)
        Dim inputs(3) As NativeCapture.INPUT
        Dim extra As UInteger = If(extended, NativeCapture.KEYEVENTF_EXTENDEDKEY, 0UI)
        inputs(0).type = NativeCapture.INPUT_KEYBOARD : inputs(0).ki.wVk = modifier
        inputs(1).type = NativeCapture.INPUT_KEYBOARD : inputs(1).ki.wVk = key : inputs(1).ki.dwFlags = extra
        inputs(2).type = NativeCapture.INPUT_KEYBOARD : inputs(2).ki.wVk = key : inputs(2).ki.dwFlags = extra Or NativeCapture.KEYEVENTF_KEYUP
        inputs(3).type = NativeCapture.INPUT_KEYBOARD : inputs(3).ki.wVk = modifier : inputs(3).ki.dwFlags = NativeCapture.KEYEVENTF_KEYUP
        NativeCapture.SendInput(4UI, inputs, System.Runtime.InteropServices.Marshal.SizeOf(Of NativeCapture.INPUT)())
    End Sub

    Private Shared Async Function ClipboardCopyAsync(expectedWindow As System.IntPtr) As System.Threading.Tasks.Task(Of CopyOutcome)
        Dim outcome As New CopyOutcome()
        Dim owner As New ClipboardOwnerWindow()
        Try
            Dim snapshot As ClipboardSnapshot = Await RunStaAsync(Function() ClipboardSnapshot.Capture(), 3000)
            Dim sequenceBefore As UInteger = Native.GetClipboardSequenceNumber()
            If Native.GetForegroundWindow() <> expectedWindow Then Return outcome
            SendChord(NativeCapture.VK_CONTROL_U, NativeCapture.VK_C, False)
            Dim sequenceAfter As UInteger = sequenceBefore
            Dim watch As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew()
            Dim retried As Boolean = False
            Do While watch.ElapsedMilliseconds < 2500
                Await System.Threading.Tasks.Task.Delay(40)
                sequenceAfter = Native.GetClipboardSequenceNumber()
                If sequenceAfter <> sequenceBefore Then Exit Do
                If Not retried AndAlso watch.ElapsedMilliseconds > 1200 AndAlso Native.GetForegroundWindow() = expectedWindow Then
                    ' Some viewers ignore Ctrl+C in a pane but honour Ctrl+Insert.
                    SendChord(NativeCapture.VK_CONTROL_U, NativeCapture.VK_INSERT, True)
                    retried = True
                End If
            Loop
            If sequenceAfter = sequenceBefore Then Return outcome
            outcome.ClipboardChanged = True
            ' Delayed rendering: give the source a moment to finish writing the formats.
            Await System.Threading.Tasks.Task.Delay(60)
            Dim text As String = Await RunStaAsync(Function() ClipboardSnapshot.ReadUnicodeText(), 2000)
            outcome.Text = If(text, "")
            Dim current As UInteger = Native.GetClipboardSequenceNumber()
            If snapshot IsNot Nothing AndAlso snapshot.Complete AndAlso current = sequenceAfter Then
                Dim restored As Boolean = Await RunStaAsync(Function() snapshot.Restore(owner.Handle), 3000)
                If Not restored Then ClientLog.Write("selection.clipboard.restore.failed")
            ElseIf snapshot Is Nothing OrElse Not snapshot.Complete Then
                ClientLog.Write("selection.clipboard.restore.skipped-incomplete-snapshot")
            End If
        Catch ex As System.Exception
            ClientLog.Write("selection.clipboard.failed", ex)
        Finally
            owner.DestroyHandle()
        End Try
        Return outcome
    End Function

    ' Optional "Ersetzen": paste the answer as plain text into the original window, then restore the clipboard.
    Public Shared Async Function PasteBackAsync(target As System.IntPtr, text As String) As System.Threading.Tasks.Task(Of Boolean)
        If target = System.IntPtr.Zero OrElse String.IsNullOrEmpty(text) Then Return False
        Dim processId As UInteger
        Native.GetWindowThreadProcessId(target, processId)
        If processId = 0UI OrElse (IsElevated(processId) AndAlso Not IsElevated(CUInt(System.Environment.ProcessId))) Then Return False
        Dim owner As New ClipboardOwnerWindow()
        Try
            Dim snapshot As ClipboardSnapshot = Await RunStaAsync(Function() ClipboardSnapshot.Capture(), 3000)
            Dim written As Boolean = Await RunStaAsync(Function() ClipboardSnapshot.WriteUnicodeText(text, owner.Handle), 3000)
            If Not written Then Return False
            Dim ours As UInteger = Native.GetClipboardSequenceNumber()
            Native.SetForegroundWindow(target)
            For attempt As Integer = 0 To 19
                If Native.GetForegroundWindow() = target Then Exit For
                Await System.Threading.Tasks.Task.Delay(25)
            Next
            If Native.GetForegroundWindow() <> target Then Return False
            Dim waited As Integer = 0
            Do While ModifiersDown() AndAlso waited < 3000
                Await System.Threading.Tasks.Task.Delay(30)
                waited += 30
            Loop
            SendChord(NativeCapture.VK_CONTROL_U, NativeCapture.VK_V, False)
            ' The target reads the clipboard asynchronously; wait before restoring.
            Await System.Threading.Tasks.Task.Delay(700)
            If snapshot IsNot Nothing AndAlso snapshot.Complete AndAlso Native.GetClipboardSequenceNumber() = ours Then
                Await RunStaAsync(Function() snapshot.Restore(owner.Handle), 3000)
            End If
            ClientLog.Write("selection.paste-back.sent")
            Return True
        Catch ex As System.Exception
            ClientLog.Write("selection.paste-back.failed", ex)
            Return False
        Finally
            owner.DestroyHandle()
        End Try
    End Function
End Class

' A message-only window: EmptyClipboard + SetClipboardData need a real clipboard owner window.
Friend NotInheritable Class ClipboardOwnerWindow
    Inherits System.Windows.Forms.NativeWindow
    Public Sub New()
        CreateHandle(New System.Windows.Forms.CreateParams With {.Caption = "AskInky.ClipboardOwner", .Parent = New System.IntPtr(-3)})
    End Sub
End Class
