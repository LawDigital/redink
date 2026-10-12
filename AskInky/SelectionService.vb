' Part of "Red Ink" (Ask Inky)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: SelectionService.vb
' Purpose: Ask Inky desktop functionality in the modular Windows client.
' Architecture: Windows native host with embedded Office Web-based interface.
' Notes: Red Ink License and Ask Inky third-party attributions are embedded in LICENSE.txt.
' =============================================================================
Public NotInheritable Class SelectionService
    Public Shared Async Function CaptureAsync(expectedWindow As System.IntPtr) As System.Threading.Tasks.Task(Of String)
        ' A registered shortcut can fire while Ctrl/Shift/Alt are still held.
        ' Wait for their release so synthetic Ctrl+C is not modified.
        For attempt As Integer = 0 To 39
            If expectedWindow = System.IntPtr.Zero OrElse Native.GetForegroundWindow() <> expectedWindow Then
                ClientLog.Write("selection.capture.aborted.foreground-changed")
                Return ""
            End If
            Dim ctrlDown As Boolean = (CInt(Native.GetAsyncKeyState(Native.VK_CONTROL)) And &H8000) <> 0
            Dim altDown As Boolean = (CInt(Native.GetAsyncKeyState(Native.VK_MENU)) And &H8000) <> 0
            Dim shiftDown As Boolean = (CInt(Native.GetAsyncKeyState(Native.VK_SHIFT)) And &H8000) <> 0
            If Not ctrlDown AndAlso Not altDown AndAlso Not shiftDown Then Exit For
            Await System.Threading.Tasks.Task.Delay(50)
        Next
        If (CInt(Native.GetAsyncKeyState(Native.VK_CONTROL)) And &H8000) <> 0 OrElse (CInt(Native.GetAsyncKeyState(Native.VK_MENU)) And &H8000) <> 0 OrElse (CInt(Native.GetAsyncKeyState(Native.VK_SHIFT)) And &H8000) <> 0 Then Return ""

        Dim before As System.Windows.Forms.IDataObject = Nothing
        Try
            before = System.Windows.Forms.Clipboard.GetDataObject()
            Dim sequenceBefore As UInteger = Native.GetClipboardSequenceNumber()
            If Native.GetForegroundWindow() <> expectedWindow Then Return ""
            ' SendWait can synchronously block the WinForms message pump if the target
            ' window does not process synthetic input (e.g. after an RDP reconnect).
            ' Send posts the keystroke; the existing bounded clipboard polling below
            ' observes completion without waiting for the foreign window on this thread.
            System.Windows.Forms.SendKeys.Send("^c")
            Dim value As String = ""
            Dim sequenceAfter As UInteger = sequenceBefore
            For i As Integer = 0 To 11
                Await System.Threading.Tasks.Task.Delay(55)
                sequenceAfter = Native.GetClipboardSequenceNumber()
                If sequenceAfter <> sequenceBefore Then
                    If System.Windows.Forms.Clipboard.ContainsText() Then value = System.Windows.Forms.Clipboard.GetText()
                    Exit For
                End If
            Next
            ' Preserve clipboard formats where possible, and never overwrite a newer clipboard.
            If sequenceAfter <> sequenceBefore AndAlso before IsNot Nothing AndAlso Native.GetClipboardSequenceNumber() = sequenceAfter Then
                Try
                    System.Windows.Forms.Clipboard.SetDataObject(before, True)
                Catch ex As System.Exception
                    ClientLog.Write("selection.clipboard.restore.failed", ex)
                End Try
            End If
            Return value
        Catch ex As System.Exception
            ClientLog.Write("selection.capture.failed", ex)
            Return ""
        End Try
    End Function
End Class
