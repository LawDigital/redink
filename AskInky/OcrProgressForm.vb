' Part of "Red Ink" (Ask Inky)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: OcrProgressForm.vb
' Purpose: Lightweight non-blocking OCR activity indicator after screenshot selection.
' Architecture: Native WinForms window with embedded original Red Ink logo and
'               indeterminate animated progress ring (no external assets).
' Security: No screenshot pixels, credentials or OCR output are shown or stored here.
' =============================================================================
Friend NotInheritable Class OcrProgressForm
    Inherits System.Windows.Forms.Form

    Private ReadOnly animationTimer As System.Windows.Forms.Timer
    Private ReadOnly mark As System.Drawing.Image
    Private ReadOnly visibleLogoBounds As System.Drawing.Rectangle
    Private angle As Single = 0.0F

    Public Sub New()
        FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        ShowInTaskbar = False
        ShowIcon = False
        TopMost = True
        StartPosition = System.Windows.Forms.FormStartPosition.Manual
        AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        ClientSize = New System.Drawing.Size(120, 120)
        MinimumSize = New System.Drawing.Size(120, 120)
        MaximumSize = New System.Drawing.Size(120, 120)
        ' Use a simple opaque, square white progress popup.
        ' Keep the logo and animated ring centered within its client area.
        BackColor = System.Drawing.Color.White
        DoubleBuffered = True
        Try
            Using stream As System.IO.Stream = GetType(OcrProgressForm).Assembly.GetManifestResourceStream("AskInky.Assets.Red_Ink_Logo.png")
                If stream IsNot Nothing Then
                    Using original As System.Drawing.Image = System.Drawing.Image.FromStream(stream)
                        mark = New System.Drawing.Bitmap(original)
                        ' Scan the alpha bounds only once, not on every animation frame.
                        Dim bitmap As System.Drawing.Bitmap = DirectCast(mark, System.Drawing.Bitmap)
                        Dim minX As Integer = bitmap.Width, minY As Integer = bitmap.Height
                        Dim maxX As Integer = -1, maxY As Integer = -1
                        For y As Integer = 0 To bitmap.Height - 1
                            For x As Integer = 0 To bitmap.Width - 1
                                If bitmap.GetPixel(x, y).A > 8 Then
                                    minX = System.Math.Min(minX, x)
                                    minY = System.Math.Min(minY, y)
                                    maxX = System.Math.Max(maxX, x)
                                    maxY = System.Math.Max(maxY, y)
                                End If
                            Next
                        Next
                        If maxX >= minX AndAlso maxY >= minY Then
                            visibleLogoBounds = System.Drawing.Rectangle.FromLTRB(minX, minY, maxX + 1, maxY + 1)
                        Else
                            visibleLogoBounds = New System.Drawing.Rectangle(0, 0, bitmap.Width, bitmap.Height)
                        End If
                    End Using
                End If
            End Using
        Catch ex As System.Exception
            ClientLog.Write("ocr.progress.logo-unavailable", ex)
        End Try
        Dim workingArea As System.Drawing.Rectangle = System.Windows.Forms.Screen.FromPoint(System.Windows.Forms.Cursor.Position).WorkingArea
        Left = workingArea.Left + (workingArea.Width - Width) \ 2
        Top = workingArea.Top + (workingArea.Height - Height) \ 2
        animationTimer = New System.Windows.Forms.Timer With {.Interval = 55}
        AddHandler animationTimer.Tick, AddressOf AdvanceAnimation
    End Sub

    Protected Overrides Sub OnShown(e As System.EventArgs)
        MyBase.OnShown(e)
        animationTimer.Start()
    End Sub

    Private Sub AdvanceAnimation(sender As Object, e As System.EventArgs)
        angle = (angle + 12.0F) Mod 360.0F
        Invalidate()
    End Sub

    Protected Overrides Sub OnPaint(e As System.Windows.Forms.PaintEventArgs)
        MyBase.OnPaint(e)
        e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias
        ' Draw all elements in the same client-coordinate system. Width/Height include
        ' non-client sizing conventions and may diverge from e.Graphics on high-DPI VDI.
        Dim area As System.Drawing.Rectangle = Me.ClientRectangle
        Dim centerX As Single = CSng(area.Left) + CSng(area.Width) / 2.0F
        Dim centerY As Single = CSng(area.Top) + CSng(area.Height) / 2.0F
        Dim diameter As Single = System.Math.Max(40.0F, System.Math.Min(CSng(area.Width), CSng(area.Height)) - 18.0F)
        Dim ring As New System.Drawing.RectangleF(centerX - diameter / 2.0F, centerY - diameter / 2.0F, diameter, diameter)
        ' Fill the complete square client area with white (no transparency key).
        e.Graphics.Clear(System.Drawing.Color.White)
        Using track As New System.Drawing.Pen(System.Drawing.Color.FromArgb(235, 235, 235), 5.0F)
            e.Graphics.DrawEllipse(track, ring)
        End Using
        Using active As New System.Drawing.Pen(System.Drawing.Color.FromArgb(185, 19, 18), 5.0F)
            active.StartCap = System.Drawing.Drawing2D.LineCap.Round
            active.EndCap = System.Drawing.Drawing2D.LineCap.Round
            e.Graphics.DrawArc(active, ring, angle, 102.0F)
        End Using
        If mark IsNot Nothing Then
            ' Center the visible alpha content, not transparent padding in the icon.
            Dim visible As System.Drawing.Rectangle = visibleLogoBounds
            Dim ratio As Single = System.Math.Min(48.0F / CSng(visible.Width), 48.0F / CSng(visible.Height))
            Dim logoWidth As Single = visible.Width * ratio
            Dim logoHeight As Single = visible.Height * ratio
            Dim destination As New System.Drawing.RectangleF(centerX - logoWidth / 2.0F,
                                                              centerY - logoHeight / 2.0F,
                                                              logoWidth, logoHeight)
            e.Graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic
            e.Graphics.DrawImage(mark, destination, visible, System.Drawing.GraphicsUnit.Pixel)
        End If
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            animationTimer.Stop()
            animationTimer.Dispose()
            If mark IsNot Nothing Then mark.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub
End Class
