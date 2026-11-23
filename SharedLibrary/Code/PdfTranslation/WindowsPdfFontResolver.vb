Option Explicit On
Option Strict On
Option Infer On

Namespace Agents
    Public NotInheritable Class WindowsPdfFontResolver
        Implements PdfSharp.Fonts.IFontResolver

        Private Shared ReadOnly RegistrationLock As New System.Object()

        Public Shared Function EnsureAvailable(Optional factory As System.Func(Of PdfSharp.Fonts.IFontResolver) = Nothing) As System.Boolean
            SyncLock RegistrationLock
                If PdfSharp.Fonts.GlobalFontSettings.FontResolver IsNot Nothing OrElse
                    PdfSharp.Fonts.GlobalFontSettings.FallbackFontResolver IsNot Nothing OrElse
                    PdfSharp.Fonts.GlobalFontSettings.UseWindowsFontsUnderWindows Then Return True

                Dim resolver As PdfSharp.Fonts.IFontResolver = If(factory Is Nothing, New WindowsPdfFontResolver(), factory())
                If resolver Is Nothing Then Return False
                PdfSharp.Fonts.GlobalFontSettings.FontResolver = resolver
                Return True
            End SyncLock
        End Function

        Private Shared ReadOnly _fontCache As New System.Collections.Generic.Dictionary(Of System.String, System.Byte())(System.StringComparer.OrdinalIgnoreCase)

        ''' <summary>
        ''' Resolves a requested font family/style combination to a concrete font face key.
        ''' </summary>
        Public Function ResolveTypeface(familyName As String, bold As Boolean, italic As Boolean) As PdfSharp.Fonts.FontResolverInfo _
                Implements PdfSharp.Fonts.IFontResolver.ResolveTypeface

            ' Map common family names to Windows font filenames
            Dim key As String = "arial.ttf"
            Select Case familyName.ToLowerInvariant()
                Case "arial", "helvetica"
                    If bold AndAlso italic Then
                        key = "arialbi.ttf"
                    ElseIf bold Then
                        key = "arialbd.ttf"
                    ElseIf italic Then
                        key = "ariali.ttf"
                    Else
                        key = "arial.ttf"
                    End If
                Case "times new roman", "times"
                    If bold AndAlso italic Then
                        key = "timesbi.ttf"
                    ElseIf bold Then
                        key = "timesbd.ttf"
                    ElseIf italic Then
                        key = "timesi.ttf"
                    Else
                        key = "times.ttf"
                    End If
                Case "courier new", "courier"
                    If bold AndAlso italic Then
                        key = "courbi.ttf"
                    ElseIf bold Then
                        key = "courbd.ttf"
                    ElseIf italic Then
                        key = "couri.ttf"
                    Else
                        key = "cour.ttf"
                    End If
                Case "segoe ui"
                    If bold AndAlso italic Then
                        key = "segoeuiz.ttf"
                    ElseIf bold Then
                        key = "segoeuib.ttf"
                    ElseIf italic Then
                        key = "segoeuii.ttf"
                    Else
                        key = "segoeui.ttf"
                    End If
                Case "calibri"
                    If bold AndAlso italic Then
                        key = "calibriz.ttf"
                    ElseIf bold Then
                        key = "calibrib.ttf"
                    ElseIf italic Then
                        key = "calibrii.ttf"
                    Else
                        key = "calibri.ttf"
                    End If
                Case "verdana"
                    If bold AndAlso italic Then
                        key = "verdanaz.ttf"
                    ElseIf bold Then
                        key = "verdanab.ttf"
                    ElseIf italic Then
                        key = "verdanai.ttf"
                    Else
                        key = "verdana.ttf"
                    End If
                Case "tahoma"
                    If bold Then
                        key = "tahomabd.ttf"
                    Else
                        key = "tahoma.ttf"
                    End If
                Case "georgia"
                    If bold AndAlso italic Then
                        key = "georgiaz.ttf"
                    ElseIf bold Then
                        key = "georgiab.ttf"
                    ElseIf italic Then
                        key = "georgiai.ttf"
                    Else
                        key = "georgia.ttf"
                    End If
                Case "trebuchet ms"
                    If bold AndAlso italic Then
                        key = "trebucbi.ttf"
                    ElseIf bold Then
                        key = "trebucbd.ttf"
                    ElseIf italic Then
                        key = "trebucit.ttf"
                    Else
                        key = "trebuc.ttf"
                    End If
                Case Else
                    ' Fallback to Arial
                    If bold Then
                        key = "arialbd.ttf"
                    Else
                        key = "arial.ttf"
                    End If
            End Select

            Return New PdfSharp.Fonts.FontResolverInfo(key)
        End Function

        ''' <summary>
        ''' Loads raw font bytes for a resolved face name from Windows font locations.
        ''' </summary>
        Public Function GetFont(faceName As String) As Byte() _
                Implements PdfSharp.Fonts.IFontResolver.GetFont

            SyncLock _fontCache
                If _fontCache.ContainsKey(faceName) Then Return _fontCache(faceName)
            End SyncLock

            Dim fontsDir = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Fonts)
            Dim fontPath = System.IO.Path.Combine(fontsDir, faceName)

            ' Also check Windows\Fonts directly (SpecialFolder.Fonts may return user fonts folder on some systems)
            If Not System.IO.File.Exists(fontPath) Then
                fontPath = System.IO.Path.Combine(System.Environment.GetEnvironmentVariable("SystemRoot"), "Fonts", faceName)
            End If

            If System.IO.File.Exists(fontPath) Then
                Dim data = System.IO.File.ReadAllBytes(fontPath)
                SyncLock _fontCache
                    _fontCache(faceName) = data
                End SyncLock
                Return data
            End If

            ' Last resort: return Nothing — PdfSharp will throw a descriptive error
            Return Nothing
        End Function
    End Class
End Namespace
