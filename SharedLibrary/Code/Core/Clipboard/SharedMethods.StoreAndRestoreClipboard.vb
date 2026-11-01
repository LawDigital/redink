' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.StoreAndRestoreClipboard.vb
' Purpose: Stores a single clipboard payload in a private shared field and
'          restores it later.
'
' Architecture:
'  - Storage: `StoreClipboard()` reads the first matching clipboard content in
'    this order: Text, Image, FileDropList, Serializable; otherwise stores Nothing.
'  - Restoration: `RestoreClipboard()` restores the stored payload based on its
'    runtime type (String, Image, StringCollection for FileDropList, otherwise Serializable).
'  - STA Isolation: Both methods run clipboard operations on a dedicated STA
'    thread, as required by System.Windows.Forms.Clipboard.
'  - Scope: Only one value is stored at a time (overwritten on each call).
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Windows.Forms

Namespace SharedLibrary

    Partial Public Class SharedMethods

        ''' <summary>
        ''' Stores the last captured clipboard payload for <see cref="StoreClipboard"/> / <see cref="RestoreClipboard"/>.
        ''' </summary>
        Private Shared clipboardData As Object = Nothing

        ''' <summary>
        ''' Captures one supported clipboard payload into a shared field.
        ''' The first matching format in the implemented check order is stored.
        ''' </summary>
        Public Shared Sub StoreClipboard()
            Dim capturedData As System.Object = Nothing
            Dim t As New Threading.Thread(
                Sub()
                    Try
                        If Clipboard.ContainsText() Then
                            capturedData = Clipboard.GetText()
                        ElseIf Clipboard.ContainsImage() Then
                            capturedData = Clipboard.GetImage()
                        ElseIf Clipboard.ContainsFileDropList() Then
                            capturedData = Clipboard.GetFileDropList()
                        ElseIf Clipboard.ContainsData(DataFormats.Serializable) Then
                            capturedData = Clipboard.GetData(DataFormats.Serializable)
                        Else
                            capturedData = Nothing ' No supported data format found
                        End If
                    Catch
                        capturedData = Nothing
                    End Try
                End Sub)
            t.SetApartmentState(Threading.ApartmentState.STA)
            t.IsBackground = True
            t.Start()
            If t.Join(5000) Then
                clipboardData = capturedData
            Else
                ' The late worker may finish its local read, but cannot overwrite shared state.
                clipboardData = Nothing
                System.Diagnostics.Trace.TraceWarning("Red Ink clipboard storage timed out; no payload will be restored.")
            End If
        End Sub

        ''' <summary>
        ''' Restores the last value captured by <see cref="StoreClipboard"/> to the clipboard.
        ''' </summary>
        Public Shared Sub RestoreClipboard()
            Dim dataToRestore As System.Object = clipboardData
            If dataToRestore Is Nothing Then Return

            Dim t As New Threading.Thread(
                Sub()
                    Try
                        If TypeOf dataToRestore Is String Then
                            Clipboard.SetText(CStr(dataToRestore))
                        ElseIf TypeOf dataToRestore Is Image Then
                            Clipboard.SetImage(CType(dataToRestore, Image))
                        ElseIf TypeOf dataToRestore Is Specialized.StringCollection Then
                            Clipboard.SetFileDropList(CType(dataToRestore, Specialized.StringCollection))
                        Else
                            Clipboard.SetData(DataFormats.Serializable, dataToRestore)
                        End If
                    Catch
                    End Try
                End Sub)
            t.SetApartmentState(Threading.ApartmentState.STA)
            t.IsBackground = True
            t.Start()
            If Not t.Join(5000) Then
                System.Diagnostics.Trace.TraceWarning("Red Ink clipboard restoration timed out; the native call was not retried.")
            End If
        End Sub

    End Class
End Namespace