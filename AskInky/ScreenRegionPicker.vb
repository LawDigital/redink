' Part of "Red Ink" (Ask Inky)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: ScreenRegionPicker.vb
' Purpose: One-shot, user-initiated multi-monitor screenshot rectangle selection.
' Architecture: WinForms overlay; no hooks, no injected input, no Helper modifications.
' Security: Captured pixels stay in memory; Escape cancels; maximum PNG bytes enforced by caller.
' =============================================================================
Friend NotInheritable Class ScreenRegionPicker
    Inherits System.Windows.Forms.Form
    Private ReadOnly virtualBounds As System.Drawing.Rectangle
    Private ReadOnly backgroundCapture As System.Drawing.Bitmap
    Private start As System.Drawing.Point
    Private selection As System.Drawing.Rectangle
    Private dragging As Boolean

    Public Sub New()
        virtualBounds = System.Windows.Forms.SystemInformation.VirtualScreen
        If virtualBounds.Width <= 0 OrElse virtualBounds.Height <= 0 OrElse CLng(virtualBounds.Width) * CLng(virtualBounds.Height) > 100000000L Then Throw New System.InvalidOperationException("Virtual desktop dimensions unsupported")
        backgroundCapture = New System.Drawing.Bitmap(virtualBounds.Width, virtualBounds.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
        Using graphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(backgroundCapture)
            graphics.CopyFromScreen(virtualBounds.Left, virtualBounds.Top, 0, 0, virtualBounds.Size, System.Drawing.CopyPixelOperation.SourceCopy)
        End Using
        FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        ShowInTaskbar = False
        TopMost = True
        KeyPreview = True
        StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Bounds = virtualBounds
        Cursor = System.Windows.Forms.Cursors.Cross
        DoubleBuffered = True
    End Sub

    Public Function SelectedPng() As Byte()
        If selection.Width < 3 OrElse selection.Height < 3 Then Return Nothing
        ' DrawImage may scale at mixed DPI. Clamp mapped coordinates before cropping.
        Dim sourceWidth As Integer = System.Math.Max(1, ClientSize.Width)
        Dim sourceHeight As Integer = System.Math.Max(1, ClientSize.Height)
        Dim x As Integer = CInt(System.Math.Floor(selection.X * CDbl(backgroundCapture.Width) / sourceWidth))
        Dim y As Integer = CInt(System.Math.Floor(selection.Y * CDbl(backgroundCapture.Height) / sourceHeight))
        Dim right As Integer = CInt(System.Math.Ceiling(selection.Right * CDbl(backgroundCapture.Width) / sourceWidth))
        Dim bottom As Integer = CInt(System.Math.Ceiling(selection.Bottom * CDbl(backgroundCapture.Height) / sourceHeight))
        x = System.Math.Max(0, System.Math.Min(x, backgroundCapture.Width - 1))
        y = System.Math.Max(0, System.Math.Min(y, backgroundCapture.Height - 1))
        right = System.Math.Max(x + 1, System.Math.Min(right, backgroundCapture.Width))
        bottom = System.Math.Max(y + 1, System.Math.Min(bottom, backgroundCapture.Height))
        Using cropped As System.Drawing.Bitmap = backgroundCapture.Clone(New System.Drawing.Rectangle(x, y, right - x, bottom - y), backgroundCapture.PixelFormat)
            Using output As New System.IO.MemoryStream()
                cropped.Save(output, System.Drawing.Imaging.ImageFormat.Png)
                Return output.ToArray()
            End Using
        End Using
    End Function

    Protected Overrides Sub OnPaint(e As System.Windows.Forms.PaintEventArgs)
        MyBase.OnPaint(e)
        e.Graphics.DrawImage(backgroundCapture, ClientRectangle)
        Using tint As New System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(75, 0, 0, 0))
            e.Graphics.FillRectangle(tint, ClientRectangle)
        End Using
        If selection.Width > 0 AndAlso selection.Height > 0 Then
            e.Graphics.DrawImage(backgroundCapture, selection, selection, System.Drawing.GraphicsUnit.Pixel)
            Using pen As New System.Drawing.Pen(System.Drawing.Color.FromArgb(185, 19, 18), 2)
                e.Graphics.DrawRectangle(pen, selection)
            End Using
        End If
    End Sub
    Protected Overrides Sub OnMouseDown(e As System.Windows.Forms.MouseEventArgs)
        MyBase.OnMouseDown(e)
        If e.Button = System.Windows.Forms.MouseButtons.Right Then DialogResult = System.Windows.Forms.DialogResult.Cancel : Close() : Return
        If e.Button <> System.Windows.Forms.MouseButtons.Left Then Return
        start = e.Location
        dragging = True
        Capture = True
    End Sub
    Protected Overrides Sub OnMouseMove(e As System.Windows.Forms.MouseEventArgs)
        MyBase.OnMouseMove(e)
        If Not dragging Then Return
        Dim left As Integer = System.Math.Max(0, System.Math.Min(start.X, e.X))
        Dim top As Integer = System.Math.Max(0, System.Math.Min(start.Y, e.Y))
        Dim right As Integer = System.Math.Min(ClientSize.Width, System.Math.Max(start.X, e.X))
        Dim bottom As Integer = System.Math.Min(ClientSize.Height, System.Math.Max(start.Y, e.Y))
        selection = System.Drawing.Rectangle.FromLTRB(left, top, right, bottom)
        Invalidate()
    End Sub
    Protected Overrides Sub OnMouseUp(e As System.Windows.Forms.MouseEventArgs)
        MyBase.OnMouseUp(e)
        If Not dragging Then Return
        dragging = False
        Capture = False
        If selection.Width < 3 OrElse selection.Height < 3 Then Return
        DialogResult = System.Windows.Forms.DialogResult.OK
        Close()
    End Sub
    Protected Overrides Sub OnKeyDown(e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = System.Windows.Forms.Keys.Escape Then
            DialogResult = System.Windows.Forms.DialogResult.Cancel
            Close()
        End If
        MyBase.OnKeyDown(e)
    End Sub
    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then backgroundCapture.Dispose()
        MyBase.Dispose(disposing)
    End Sub
End Class
