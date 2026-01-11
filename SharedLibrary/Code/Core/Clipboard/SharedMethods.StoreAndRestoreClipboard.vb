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
            Dim t As New Threading.Thread(
                Sub()
                    Try
                        If Clipboard.ContainsText() Then
                            clipboardData = Clipboard.GetText()
                        ElseIf Clipboard.ContainsImage() Then
                            clipboardData = Clipboard.GetImage()
                        ElseIf Clipboard.ContainsFileDropList() Then
                            clipboardData = Clipboard.GetFileDropList()
                        ElseIf Clipboard.ContainsData(DataFormats.Serializable) Then
                            clipboardData = Clipboard.GetData(DataFormats.Serializable)
                        Else
                            clipboardData = Nothing ' No supported data format found
                        End If
                    Catch
                        clipboardData = Nothing
                    End Try
                End Sub)
            t.SetApartmentState(Threading.ApartmentState.STA)
            t.Start()
            t.Join()
        End Sub

        ''' <summary>
        ''' Restores the last value captured by <see cref="StoreClipboard"/> to the clipboard.
        ''' </summary>
        Public Shared Sub RestoreClipboard()
            If clipboardData Is Nothing Then Return

            Dim t As New Threading.Thread(
                Sub()
                    Try
                        If TypeOf clipboardData Is String Then
                            Clipboard.SetText(CStr(clipboardData))
                        ElseIf TypeOf clipboardData Is Image Then
                            Clipboard.SetImage(CType(clipboardData, Image))
                        ElseIf TypeOf clipboardData Is Specialized.StringCollection Then
                            Clipboard.SetFileDropList(CType(clipboardData, Specialized.StringCollection))
                        Else
                            Clipboard.SetData(DataFormats.Serializable, clipboardData)
                        End If
                    Catch
                    End Try
                End Sub)
            t.SetApartmentState(Threading.ApartmentState.STA)
            t.Start()
            t.Join()
        End Sub

    End Class
End Namespace