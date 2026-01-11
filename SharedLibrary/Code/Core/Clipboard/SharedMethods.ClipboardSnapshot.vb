' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.ClipboardSnapshot.vb
' Purpose: Captures a minimal, "safe" snapshot of the current Windows clipboard
'          and restores it later.
'
' Architecture:
'  - Snapshot Creation: `Capture()` reads only selected simple clipboard formats
'    (HTML, RTF, UnicodeText/Text, FileDropList) into a new `DataObject`.
'  - Safety Constraint: The capture path intentionally avoids enumerating all
'    clipboard formats to reduce the risk of triggering delayed rendering for
'    OLE/embedded objects.
'  - STA Isolation: Both capture and restore are executed on a dedicated STA
'    thread, as required by `System.Windows.Forms.Clipboard`.
'  - Best-Effort Behavior: Exceptions from clipboard access are swallowed and
'    an empty `DataObject` may be returned when capture fails.
'  - Legacy Support: `CloneData()` remains for backward compatibility but is
'    not used by the current safe-capture implementation.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary

    Partial Public Class SharedMethods

        ''' <summary>
        ''' Provides a minimal clipboard snapshot facility that captures a restricted
        ''' set of clipboard formats and can restore them later.
        ''' </summary>
        Public NotInheritable Class ClipboardSnapshot

            ''' <summary>
            ''' Prevents instantiation. This type provides shared (static) members only.
            ''' </summary>
            Private Sub New()
            End Sub

            ''' <summary>
            ''' Captures a restricted subset of clipboard formats into a new <see cref="System.Windows.Forms.DataObject"/>.
            ''' The operation is performed on an STA thread for clipboard compatibility.
            ''' </summary>
            ''' <returns>
            ''' An <see cref="System.Windows.Forms.IDataObject"/> containing the captured formats.
            ''' Returns an empty <see cref="System.Windows.Forms.DataObject"/> on failure.
            ''' </returns>
            Public Shared Function Capture() As System.Windows.Forms.IDataObject
                Dim captured As System.Windows.Forms.IDataObject = Nothing

                Dim t As New System.Threading.Thread(
                    Sub()
                        Try
                            Dim result As New System.Windows.Forms.DataObject()

                            ' Capture ONLY safe, simple formats. Avoid enumerating all formats to prevent
                            ' delayed rendering of OLE/embedded objects that can re-enter Word.
                            Try
                                If System.Windows.Forms.Clipboard.ContainsText(System.Windows.Forms.TextDataFormat.Html) Then
                                    Dim html = System.Windows.Forms.Clipboard.GetText(System.Windows.Forms.TextDataFormat.Html)
                                    result.SetData(System.Windows.Forms.DataFormats.Html, False, html)
                                End If
                            Catch
                            End Try

                            Try
                                If System.Windows.Forms.Clipboard.ContainsText(System.Windows.Forms.TextDataFormat.Rtf) Then
                                    Dim rtf = System.Windows.Forms.Clipboard.GetText(System.Windows.Forms.TextDataFormat.Rtf)
                                    result.SetData(System.Windows.Forms.DataFormats.Rtf, False, rtf)
                                End If
                            Catch
                            End Try

                            Try
                                If System.Windows.Forms.Clipboard.ContainsText(System.Windows.Forms.TextDataFormat.UnicodeText) Then
                                    Dim u = System.Windows.Forms.Clipboard.GetText(System.Windows.Forms.TextDataFormat.UnicodeText)
                                    result.SetData(System.Windows.Forms.DataFormats.UnicodeText, False, u)
                                ElseIf System.Windows.Forms.Clipboard.ContainsText(System.Windows.Forms.TextDataFormat.Text) Then
                                    Dim ttxt = System.Windows.Forms.Clipboard.GetText(System.Windows.Forms.TextDataFormat.Text)
                                    result.SetData(System.Windows.Forms.DataFormats.Text, False, ttxt)
                                End If
                            Catch
                            End Try

                            Try
                                If System.Windows.Forms.Clipboard.ContainsFileDropList() Then
                                    Dim files = System.Windows.Forms.Clipboard.GetFileDropList()
                                    If files IsNot Nothing AndAlso files.Count > 0 Then
                                        Dim copy As New System.Collections.Specialized.StringCollection()
                                        For Each f As String In files
                                            copy.Add(f)
                                        Next
                                        result.SetFileDropList(copy)
                                    End If
                                End If
                            Catch
                            End Try

                            captured = result
                        Catch
                            captured = New System.Windows.Forms.DataObject() ' empty fallback
                        End Try
                    End Sub)
                t.SetApartmentState(System.Threading.ApartmentState.STA)
                t.Start()
                t.Join()

                Return captured
            End Function

            ''' <summary>
            ''' Restores a previously captured clipboard snapshot.
            ''' The operation is performed on an STA thread for clipboard compatibility.
            ''' </summary>
            ''' <param name="snapshot">The snapshot to restore. If <c>Nothing</c>, the method returns without action.</param>
            Public Shared Sub Restore(snapshot As System.Windows.Forms.IDataObject)
                If snapshot Is Nothing Then Return

                Dim t As New System.Threading.Thread(
                    Sub()
                        Try
                            ' true => keep after app exits; also retries internally.
                            System.Windows.Forms.Clipboard.SetDataObject(snapshot, True)
                        Catch exClip As System.Runtime.InteropServices.ExternalException
                            ' Clipboard busy — best effort only.
                        Catch exAny As System.Exception
                        End Try
                    End Sub)
                t.SetApartmentState(System.Threading.ApartmentState.STA)
                t.Start()
                t.Join()
            End Sub

            ''' <summary>
            ''' Clones clipboard payload objects for a given format.
            ''' This method is retained for backward compatibility and is not used by <see cref="Capture"/>.
            ''' </summary>
            ''' <param name="fmt">The clipboard format identifier.</param>
            ''' <param name="data">The clipboard data object to clone.</param>
            ''' <returns>A cloned object where supported; otherwise the original <paramref name="data"/> reference.</returns>
            Private Shared Function CloneData(fmt As System.String, data As System.Object) As System.Object
                ' Kept for backward compatibility; not used by the safe-capture path.

                If TypeOf data Is System.IO.Stream Then
                    Dim src As System.IO.Stream = DirectCast(data, System.IO.Stream)
                    Try
                        If src.CanSeek Then src.Position = 0
                    Catch
                    End Try
                    Dim ms As New System.IO.MemoryStream()
                    src.CopyTo(ms)
                    ms.Position = 0
                    Return ms
                End If

                If TypeOf data Is System.Drawing.Bitmap Then
                    Dim bmp As System.Drawing.Bitmap = DirectCast(data, System.Drawing.Bitmap)
                    Dim rect As New System.Drawing.Rectangle(0, 0, bmp.Width, bmp.Height)
                    Return bmp.Clone(rect, bmp.PixelFormat)
                End If

                If TypeOf data Is System.Drawing.Image Then
                    Dim img As System.Drawing.Image = DirectCast(data, System.Drawing.Image)
                    Return New System.Drawing.Bitmap(img)
                End If

                If TypeOf data Is System.String() Then
                    Dim arr As System.String() = DirectCast(data, System.String())
                    Return CType(arr.Clone(), System.String())
                End If

                If TypeOf data Is System.String Then
                    Return System.String.Copy(DirectCast(data, System.String))
                End If

                Try
                    If TypeOf data Is System.ICloneable Then
                        Return DirectCast(data, System.ICloneable).Clone()
                    End If
                Catch
                End Try

                Return data
            End Function
        End Class


    End Class

End Namespace