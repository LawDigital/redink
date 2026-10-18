' Mirrored after AskInky.SelectionService: send Ctrl+C to the original foreground
' window only, observe clipboard sequence, restore the previous IDataObject when safe.
Friend NotInheritable Class SelectionClipboardService
    Private Sub New()
    End Sub

    Private Shared Function ModifiersHeld() As Boolean
        Return (CInt(Native.GetAsyncKeyState(Native.VK_CONTROL)) And &H8000) <> 0 OrElse
               (CInt(Native.GetAsyncKeyState(Native.VK_MENU)) And &H8000) <> 0 OrElse
               (CInt(Native.GetAsyncKeyState(Native.VK_SHIFT)) And &H8000) <> 0
    End Function

    Private Shared Function MouseButtonsHeld() As Boolean
        Return (CInt(Native.GetAsyncKeyState(Native.VK_LBUTTON)) And &H8000) <> 0 OrElse
               (CInt(Native.GetAsyncKeyState(Native.VK_RBUTTON)) And &H8000) <> 0 OrElse
               (CInt(Native.GetAsyncKeyState(Native.VK_MBUTTON)) And &H8000) <> 0
    End Function

    Private Shared Async Function WaitUntilReleasedAsync(Optional includeMouse As Boolean = False) As System.Threading.Tasks.Task(Of Boolean)
        For i As Integer = 0 To 39
            If Not ModifiersHeld() AndAlso (Not includeMouse OrElse Not MouseButtonsHeld()) Then Return True
            Await System.Threading.Tasks.Task.Delay(50)
        Next
        Return False
    End Function

    Public Shared Async Function ReadSelectionAsync(window As System.IntPtr) As System.Threading.Tasks.Task(Of String)
        If window = System.IntPtr.Zero OrElse Not Native.IsWindow(window) Then Return ""
        If Not (Await WaitUntilReleasedAsync()) Then Return ""
        If Native.GetForegroundWindow() <> window Then Return ""
        Dim previous As System.Windows.Forms.IDataObject = Nothing
        Dim copySeen As Boolean = False
        Dim sequenceAfter As UInteger = 0UI
        Try
            previous = System.Windows.Forms.Clipboard.GetDataObject()
            Dim before As UInteger = Native.GetClipboardSequenceNumber()
            If Native.GetForegroundWindow() <> window Then Return ""
            System.Windows.Forms.SendKeys.Send("^c")
            Dim text As String = ""
            For i As Integer = 0 To 13
                Await System.Threading.Tasks.Task.Delay(55)
                sequenceAfter = Native.GetClipboardSequenceNumber()
                If sequenceAfter <> before Then
                    copySeen = True
                    If System.Windows.Forms.Clipboard.ContainsText() Then text = System.Windows.Forms.Clipboard.GetText()
                    Exit For
                End If
            Next
            Return text
        Catch ex As System.Exception
            AppLog.Write("selection.capture.failed", ex)
            Return ""
        Finally
            If copySeen AndAlso Native.GetClipboardSequenceNumber() = sequenceAfter Then
                Try
                    If previous IsNot Nothing Then
                        System.Windows.Forms.Clipboard.SetDataObject(previous, True)
                    Else
                        System.Windows.Forms.Clipboard.Clear()
                    End If
                Catch ex As System.Exception
                    AppLog.Write("selection.clipboard.restore.failed", ex)
                End Try
            End If
        End Try
    End Function

    Public Shared Function CaptureFocusedControl(window As System.IntPtr) As System.IntPtr
        If window = System.IntPtr.Zero OrElse Native.GetForegroundWindow() <> window Then Return System.IntPtr.Zero
        Dim processId As UInteger
        Dim threadId As UInteger = Native.GetWindowThreadProcessId(window, processId)
        If threadId = 0UI Then Return System.IntPtr.Zero
        Dim info As New Native.GuiThreadInfo With {.Size = CUInt(System.Runtime.InteropServices.Marshal.SizeOf(Of Native.GuiThreadInfo)())}
        If Not Native.GetGUIThreadInfo(threadId, info) Then Return System.IntPtr.Zero
        If info.Focus = window OrElse (info.Focus <> System.IntPtr.Zero AndAlso Native.IsChild(window, info.Focus)) Then Return info.Focus
        Return System.IntPtr.Zero
    End Function

    Private Shared Sub RestoreFocusedControl(window As System.IntPtr, originalFocus As System.IntPtr)
        If originalFocus = System.IntPtr.Zero OrElse Not Native.IsWindow(originalFocus) OrElse
           Not (originalFocus = window OrElse Native.IsChild(window, originalFocus)) OrElse
           Native.GetForegroundWindow() <> window Then Return
        If CaptureFocusedControl(window) = originalFocus Then Return
        Dim focusProcess As UInteger
        Dim focusThread As UInteger = Native.GetWindowThreadProcessId(originalFocus, focusProcess)
        Dim ownThread As UInteger = Native.GetCurrentThreadId()
        If focusThread = 0UI OrElse focusThread = ownThread Then Return
        Dim attached As Boolean = False
        Try
            attached = Native.AttachThreadInput(ownThread, focusThread, True)
            If Not attached Then
                AppLog.Write("insertion.focus.attach.failed " & System.Runtime.InteropServices.Marshal.GetLastWin32Error().ToString(System.Globalization.CultureInfo.InvariantCulture))
                Return
            End If
            If Native.GetForegroundWindow() = window Then Native.SetFocus(originalFocus)
        Catch ex As System.Exception
            AppLog.Write("insertion.focus.restore.failed", ex)
        Finally
            If attached Then Native.AttachThreadInput(ownThread, focusThread, False)
        End Try
    End Sub

    Private Shared Function KeyboardEvent(virtualKey As UShort, flags As UInteger) As Native.InputRecord
        Return New Native.InputRecord With {
            .InputType = Native.INPUT_KEYBOARD,
            .Data = New Native.InputData With {
                .Keyboard = New Native.KeyboardRecord With {.VirtualKey = virtualKey, .Flags = flags}
            }
        }
    End Function

    Private Shared Function SendPaste() As Boolean
        Dim downControl As Native.InputRecord = KeyboardEvent(CUShort(Native.VK_CONTROL), 0UI)
        Dim downV As Native.InputRecord = KeyboardEvent(Native.VK_V, 0UI)
        Dim upV As Native.InputRecord = KeyboardEvent(Native.VK_V, Native.KEYEVENTF_KEYUP)
        Dim upControl As Native.InputRecord = KeyboardEvent(CUShort(Native.VK_CONTROL), Native.KEYEVENTF_KEYUP)
        Dim input As Native.InputRecord() = {downControl, downV, upV, upControl}
        Dim emitted As UInteger = Native.SendInput(CUInt(input.Length), input, System.Runtime.InteropServices.Marshal.SizeOf(Of Native.InputRecord)())
        If emitted = CUInt(input.Length) Then Return True
        ' A partial injection must not leave a virtual Ctrl key pressed.
        Dim releases As Native.InputRecord() = {upV, upControl}
        Native.SendInput(CUInt(releases.Length), releases, System.Runtime.InteropServices.Marshal.SizeOf(Of Native.InputRecord)())
        Return False
    End Function

    Public Shared Async Function InsertAsync(window As System.IntPtr, originalFocus As System.IntPtr, content As String) As System.Threading.Tasks.Task(Of Boolean)
        If window = System.IntPtr.Zero OrElse Not Native.IsWindow(window) OrElse System.String.IsNullOrEmpty(content) Then Return False
        ' A WebView2 double-click can complete selection while the mouse button is
        ' still held. Do not close the picker and paste under a pending mouse-up:
        ' Windows may restore the picker focus and swallow Ctrl+V after the click.
        If Not (Await WaitUntilReleasedAsync(True)) Then
            AppLog.Write("insertion.aborted.input-not-released")
            Return False
        End If
        ' Let the host process the WebView2 mouse-up and modal teardown before
        ' restoring foreground. Otherwise a queued click can steal it again.
        Await System.Threading.Tasks.Task.Delay(120)
        ' Only the originally recorded external target is eligible for insertion.
        ' A foreground request can be refused by Windows; retry and verify.
        For i As Integer = 0 To 10
            If Native.GetForegroundWindow() = window Then Exit For
            Dim activated As Boolean = Native.SetForegroundWindow(window)
            AppLog.Write("insertion.activate.attempt=" & i.ToString(System.Globalization.CultureInfo.InvariantCulture) & ";accepted=" & activated.ToString())
            Await System.Threading.Tasks.Task.Delay(95)
        Next
        If Native.GetForegroundWindow() <> window Then
            AppLog.Write("insertion.aborted.foreground-not-restored")
            Return False
        End If

        ' A short settling interval covers Windows focus activation after closing a
        ' mouse-driven WebView2 dialog. Recheck the exact target before injecting.
        RestoreFocusedControl(window, originalFocus)
        Await System.Threading.Tasks.Task.Delay(180)
        If Native.GetForegroundWindow() <> window Then
            AppLog.Write("insertion.aborted.foreground-lost-after-settle")
            Return False
        End If

        Dim original As System.Windows.Forms.IDataObject = Nothing
        Dim pastedSequence As UInteger = 0UI
        Dim temporarySet As Boolean = False
        Try
            original = System.Windows.Forms.Clipboard.GetDataObject()
            System.Windows.Forms.Clipboard.SetText(content, System.Windows.Forms.TextDataFormat.UnicodeText)
            pastedSequence = Native.GetClipboardSequenceNumber()
            temporarySet = True
            If Native.GetForegroundWindow() <> window Then Return False
            ' Deliver Ctrl+V to the verified foreground window. SendInput is
            ' independent of which control previously owned the WinForms focus,
            ' unlike SendKeys.Send, which can silently target the closed picker.
            ' If Windows/UIPI refuses the injection, return False (not success).
            Dim foregroundFocus As System.IntPtr = CaptureFocusedControl(window)
            AppLog.Write("insertion.dispatch; original-focus-captured=" & (originalFocus <> System.IntPtr.Zero).ToString() & ";focus-restored=" & (originalFocus <> System.IntPtr.Zero AndAlso foregroundFocus = originalFocus).ToString() & ";target-foreground=" & (Native.GetForegroundWindow() = window).ToString())
            If Not SendPaste() Then
                AppLog.Write("insertion.sendinput.failed " & System.Runtime.InteropServices.Marshal.GetLastWin32Error().ToString(System.Globalization.CultureInfo.InvariantCulture))
                Return False
            End If
            ' WinUI/Chromium editors may read clipboard text asynchronously after
            ' Ctrl+V has been injected. Do not restore the previous clipboard
            ' immediately after SendInput reports that the events were queued.
            Await System.Threading.Tasks.Task.Delay(1600)
            AppLog.Write("insertion.keys.dispatched")
            Return True
        Catch ex As System.Exception
            AppLog.Write("insertion.failed", ex)
            Return False
        Finally
            If temporarySet AndAlso Native.GetClipboardSequenceNumber() = pastedSequence Then
                Try
                    If original IsNot Nothing Then
                        System.Windows.Forms.Clipboard.SetDataObject(original, True)
                    Else
                        System.Windows.Forms.Clipboard.Clear()
                    End If
                Catch ex As System.Exception
                    AppLog.Write("insertion.clipboard.restore.failed", ex)
                End Try
            ElseIf temporarySet Then
                AppLog.Write("insertion.clipboard.changed-by-another-process.no-restore")
            End If
        End Try
    End Function
End Class
