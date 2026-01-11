' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.PutInClipboard.vb
' Purpose: Writes text to the Windows clipboard on an STA thread.
'
' Architecture:
'  - STA Isolation: Clipboard APIs require an STA thread; the operation runs on
'    a dedicated STA thread and blocks the caller until completion (Join).
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
            Dim thread As New Threading.Thread(Sub()
                                                   Dim textValue As String = If(text, String.Empty)

                                                   ' Check if the text is RTF formatted
                                                   If textValue.StartsWith("{\rtf", StringComparison.Ordinal) Then
                                                       ' Set RTF content to the clipboard
                                                       'Clipboard.SetData(DataFormats.Rtf, text)

                                                       Dim plainText As String

                                                       ' Convert RTF to plain text using RichTextBox
                                                       Using rtb As New RichTextBox()
                                                           rtb.Rtf = textValue
                                                           plainText = rtb.Text
                                                       End Using

                                                       ' Set both RTF and plain text in the clipboard
                                                       Dim dataObj As New DataObject()
                                                       dataObj.SetData(DataFormats.Rtf, textValue)
                                                       dataObj.SetData(DataFormats.Text, plainText)
                                                       Clipboard.SetDataObject(dataObj, True)

                                                   Else
                                                       ' Set plain text to the clipboard
                                                       Clipboard.SetText(textValue)
                                                   End If
                                               End Sub)

            ' Ensure the thread is STA (Single Thread Apartment), as required by the clipboard
            thread.SetApartmentState(Threading.ApartmentState.STA)
            thread.Start()
            thread.Join()

        End Sub

    End Class
End Namespace