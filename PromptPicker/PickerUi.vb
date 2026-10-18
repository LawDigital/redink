' Small shared desktop control vocabulary for Prompt Picker dialogs and native fallback.
' The line icons mirror Ask Inky's compact 32px SVG actions without external fonts.
Friend NotInheritable Class PickerUi
    Private Sub New()
    End Sub

    Public Shared ReadOnly Property Surface As System.Drawing.Color
        Get
            Return System.Drawing.Color.FromArgb(249, 249, 250)
        End Get
    End Property

    Public Shared ReadOnly Property Ink As System.Drawing.Color
        Get
            Return System.Drawing.Color.FromArgb(50, 49, 48)
        End Get
    End Property

    Public Shared ReadOnly Property Muted As System.Drawing.Color
        Get
            Return System.Drawing.Color.FromArgb(102, 99, 96)
        End Get
    End Property

    Public Shared Function Caption(value As String) As System.Windows.Forms.Label
        Return New System.Windows.Forms.Label With {
            .Text = value,
            .AutoSize = True,
            .ForeColor = Muted,
            .Font = New System.Drawing.Font("Segoe UI", 9.0F),
            .Margin = New System.Windows.Forms.Padding(0, 0, 0, 5)
        }
    End Function

    Public Shared Sub StyleTextBox(value As System.Windows.Forms.TextBox)
        value.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        value.BackColor = System.Drawing.Color.White
        value.ForeColor = Ink
        value.Font = New System.Drawing.Font("Segoe UI", 9.5F)
    End Sub

    Public Shared Function ActionButton(caption As String, iconName As String, isPrimary As Boolean, width As Integer) As System.Windows.Forms.Button
        Dim fore As System.Drawing.Color = If(isPrimary, System.Drawing.Color.White, Ink)
        Dim control As New RoundedActionButton With {
            .Text = caption,
            .Width = width,
            .Height = 34,
            .IsPrimary = isPrimary,
            .Font = New System.Drawing.Font("Segoe UI", 9.0F),
            .Margin = New System.Windows.Forms.Padding(6, 0, 0, 0),
            .TabStop = True
        }
        control.Image = DrawIcon(iconName, fore)
        AddHandler control.Disposed, Sub(sender As Object, e As System.EventArgs)
                                         If control.Image IsNot Nothing Then control.Image.Dispose()
                                     End Sub
        Return control
    End Function

    ' A native Button, not a decorative label: WinForms still provides keyboard
    ' activation, access keys, AcceptButton, CancelButton and DialogResult.
    Private NotInheritable Class RoundedActionButton
        Inherits System.Windows.Forms.Button

        Private hovering As Boolean
        Private pressing As Boolean
        Public Property IsPrimary As Boolean

        Public Sub New()
            SetStyle(System.Windows.Forms.ControlStyles.UserPaint Or
                     System.Windows.Forms.ControlStyles.AllPaintingInWmPaint Or
                     System.Windows.Forms.ControlStyles.OptimizedDoubleBuffer Or
                     System.Windows.Forms.ControlStyles.ResizeRedraw, True)
            FlatStyle = System.Windows.Forms.FlatStyle.Flat
            UseVisualStyleBackColor = False
        End Sub

        Protected Overrides Sub OnMouseEnter(e As System.EventArgs)
            MyBase.OnMouseEnter(e)
            hovering = True
            Invalidate()
        End Sub

        Protected Overrides Sub OnMouseLeave(e As System.EventArgs)
            MyBase.OnMouseLeave(e)
            hovering = False
            Invalidate()
        End Sub

        Protected Overrides Sub OnMouseDown(e As System.Windows.Forms.MouseEventArgs)
            MyBase.OnMouseDown(e)
            pressing = (e.Button = System.Windows.Forms.MouseButtons.Left)
            Invalidate()
        End Sub

        Protected Overrides Sub OnMouseUp(e As System.Windows.Forms.MouseEventArgs)
            MyBase.OnMouseUp(e)
            pressing = False
            Invalidate()
        End Sub

        Protected Overrides Sub OnGotFocus(e As System.EventArgs)
            MyBase.OnGotFocus(e)
            Invalidate()
        End Sub

        Protected Overrides Sub OnLostFocus(e As System.EventArgs)
            MyBase.OnLostFocus(e)
            pressing = False
            Invalidate()
        End Sub

        Protected Overrides Sub OnEnabledChanged(e As System.EventArgs)
            MyBase.OnEnabledChanged(e)
            Invalidate()
        End Sub

        Protected Overrides Sub OnPaint(e As System.Windows.Forms.PaintEventArgs)
            Dim painter As System.Drawing.Graphics = e.Graphics
            painter.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias
            painter.Clear(If(Parent Is Nothing, PickerUi.Surface, Parent.BackColor))
            Dim fill As System.Drawing.Color = If(IsPrimary, System.Drawing.Color.FromArgb(185, 19, 18), System.Drawing.Color.White)
            Dim edge As System.Drawing.Color = If(IsPrimary, fill, System.Drawing.Color.FromArgb(220, 218, 215))
            Dim fore As System.Drawing.Color = If(IsPrimary, System.Drawing.Color.White, PickerUi.Ink)
            If Not Enabled Then
                fill = System.Drawing.Color.FromArgb(246, 245, 244)
                edge = System.Drawing.Color.FromArgb(233, 232, 230)
                fore = System.Drawing.Color.FromArgb(169, 166, 164)
            ElseIf pressing Then
                fill = If(IsPrimary, System.Drawing.Color.FromArgb(145, 16, 15), System.Drawing.Color.FromArgb(237, 235, 233))
            ElseIf hovering Then
                fill = If(IsPrimary, System.Drawing.Color.FromArgb(157, 16, 15), System.Drawing.Color.FromArgb(247, 245, 243))
            End If
            Using shape As New System.Drawing.Drawing2D.GraphicsPath()
                Dim w As Single = CSng(System.Math.Max(0, Width - 1))
                Dim h As Single = CSng(System.Math.Max(0, Height - 1))
                Dim diameter As Single = System.Math.Min(14.0F, System.Math.Min(w, h))
                shape.AddArc(0.5F, 0.5F, diameter, diameter, 180.0F, 90.0F)
                shape.AddArc(w - diameter, 0.5F, diameter, diameter, 270.0F, 90.0F)
                shape.AddArc(w - diameter, h - diameter, diameter, diameter, 0.0F, 90.0F)
                shape.AddArc(0.5F, h - diameter, diameter, diameter, 90.0F, 90.0F)
                shape.CloseFigure()
                Using brush As New System.Drawing.SolidBrush(fill)
                    painter.FillPath(brush, shape)
                End Using
                Using stroke As New System.Drawing.Pen(edge, 1.0F)
                    painter.DrawPath(stroke, shape)
                End Using
            End Using
            Dim textFlags As System.Windows.Forms.TextFormatFlags = System.Windows.Forms.TextFormatFlags.NoPadding Or
                System.Windows.Forms.TextFormatFlags.SingleLine Or System.Windows.Forms.TextFormatFlags.NoPrefix
            Dim textSize As System.Drawing.Size = System.Windows.Forms.TextRenderer.MeasureText(Text, Font, New System.Drawing.Size(2048, Height), textFlags)
            Dim iconSize As Integer = If(Image Is Nothing, 0, 20)
            Dim space As Integer = If(iconSize > 0 AndAlso Text.Length > 0, 6, 0)
            Dim left As Integer = System.Math.Max(5, (Width - textSize.Width - iconSize - space) \ 2)
            If Image IsNot Nothing Then
                Using effects As New System.Drawing.Imaging.ImageAttributes()
                    If Not Enabled Then
                        Dim matrix As New System.Drawing.Imaging.ColorMatrix()
                        matrix.Matrix33 = 0.45F
                        effects.SetColorMatrix(matrix)
                    End If
                    painter.DrawImage(Image, New System.Drawing.Rectangle(left, (Height - 20) \ 2, 20, 20), 0, 0, Image.Width, Image.Height, System.Drawing.GraphicsUnit.Pixel, effects)
                End Using
            End If
            Dim labelRect As New System.Drawing.Rectangle(left + iconSize + space, 0, System.Math.Max(0, Width - left - iconSize - space - 4), Height)
            System.Windows.Forms.TextRenderer.DrawText(painter, Text, Font, labelRect, fore, textFlags Or System.Windows.Forms.TextFormatFlags.VerticalCenter)
            If Focused AndAlso ShowFocusCues Then
                Dim focus As New System.Drawing.Rectangle(3, 3, System.Math.Max(0, Width - 7), System.Math.Max(0, Height - 7))
                System.Windows.Forms.ControlPaint.DrawFocusRectangle(painter, focus, fore, fill)
            End If
        End Sub
    End Class

    Private Shared Function DrawIcon(name As String, inkColor As System.Drawing.Color) As System.Drawing.Bitmap
        Dim bitmap As New System.Drawing.Bitmap(20, 20)
        Using graphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(bitmap)
            graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias
            Using stroke As New System.Drawing.Pen(inkColor, 1.7F)
                stroke.StartCap = System.Drawing.Drawing2D.LineCap.Round
                stroke.EndCap = System.Drawing.Drawing2D.LineCap.Round
                stroke.LineJoin = System.Drawing.Drawing2D.LineJoin.Round
                Select Case name
                    Case "check"
                        graphics.DrawLines(stroke, {New System.Drawing.PointF(3.5F, 10.5F), New System.Drawing.PointF(8.0F, 14.5F), New System.Drawing.PointF(16.5F, 5.5F)})
                    Case "close"
                        graphics.DrawLine(stroke, 5.0F, 5.0F, 15.0F, 15.0F)
                        graphics.DrawLine(stroke, 15.0F, 5.0F, 5.0F, 15.0F)
                    Case "plus"
                        graphics.DrawLine(stroke, 10.0F, 4.0F, 10.0F, 16.0F)
                        graphics.DrawLine(stroke, 4.0F, 10.0F, 16.0F, 10.0F)
                    Case "insert"
                        graphics.DrawLine(stroke, 3.5F, 10.0F, 16.0F, 10.0F)
                        graphics.DrawLine(stroke, 11.0F, 5.0F, 16.0F, 10.0F)
                        graphics.DrawLine(stroke, 16.0F, 10.0F, 11.0F, 15.0F)
                    Case "keyboard"
                        graphics.DrawRectangle(stroke, 2.5F, 5.0F, 15.0F, 10.0F)
                        graphics.DrawLine(stroke, 6.0F, 12.5F, 14.0F, 12.5F)
                        graphics.DrawLine(stroke, 5.5F, 8.5F, 6.5F, 8.5F)
                        graphics.DrawLine(stroke, 9.5F, 8.5F, 10.5F, 8.5F)
                        graphics.DrawLine(stroke, 13.5F, 8.5F, 14.5F, 8.5F)
                End Select
            End Using
        End Using
        Return bitmap
    End Function
End Class
