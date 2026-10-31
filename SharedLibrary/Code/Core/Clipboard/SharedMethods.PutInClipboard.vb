' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.PutInClipboard.vb
' Purpose: Writes text to the Windows clipboard on an STA thread.
'
' Architecture:
'  - STA execution: A dedicated background STA worker preserves OLE isolation.
'    Waits are bounded to five seconds; worker failures are rethrown on the caller.
'    A native call still running at timeout is not aborted or retried automatically.
'  - Format Handling: If the input appears to be RTF (prefix "{\rtf"), the
'    method places both RTF and a derived plain-text representation into a
'    single DataObject. Otherwise, it places plain text only.
'  - Plain-Text Derivation: Plain text is obtained by loading the RTF into a
'    RichTextBox and reading its Text property.
' =============================================================================


Option Strict On
Option Explicit On

Imports System.Windows.Forms

Namespace SharedLibrary
    Partial Public Class SharedMethods

        ''' <summary>
        ''' Puts the specified text into the Windows clipboard. If the text appears to be RTF (starts with "{\rtf"),
        ''' both RTF and plain text are written; otherwise only plain text is written.
        ''' </summary>
        ''' <param name="text">The text to write to the clipboard.</param>
        Public Shared Sub PutInClipboard(text As String)
            RunOnStaThread(
                Sub()
                    Dim textValue As System.String = If(text, System.String.Empty)
                    If textValue.StartsWith("{\rtf", System.StringComparison.Ordinal) Then
                        Dim plainText As System.String = RtfToPlainText(textValue)
                        Dim data As New System.Windows.Forms.DataObject()
                        data.SetData(System.Windows.Forms.DataFormats.Rtf, textValue)
                        data.SetData(System.Windows.Forms.DataFormats.Text, plainText)
                        System.Windows.Forms.Clipboard.SetDataObject(data, True)
                    Else
                        System.Windows.Forms.Clipboard.SetText(textValue)
                    End If
                End Sub, "Red Ink clipboard", waitTimeoutMilliseconds:=5000, alwaysUseWorker:=True)

        End Sub

    End Class
End Namespace