' Part of "Red Ink" (Ask Inky)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: WindowPlacement.vb
' Purpose: Keeps desktop windows visible across monitor / DPI / VDI changes.
' Architecture: Native WinForms geometry helper; no server or application policy.
' =============================================================================
Friend NotInheritable Class WindowPlacement
    Private Sub New()
    End Sub

    Private Shared Function PreferredScreen() As System.Windows.Forms.Screen
        Return System.Windows.Forms.Screen.FromPoint(System.Windows.Forms.Cursor.Position)
    End Function

    Private Shared Function ClampToArea(desired As System.Drawing.Point,
                                         size As System.Drawing.Size,
                                         area As System.Drawing.Rectangle) As System.Drawing.Point
        Dim maximumX As Integer = System.Math.Max(area.Left, area.Right - size.Width)
        Dim maximumY As Integer = System.Math.Max(area.Top, area.Bottom - size.Height)
        Return New System.Drawing.Point(System.Math.Clamp(desired.X, area.Left, maximumX),
                                        System.Math.Clamp(desired.Y, area.Top, maximumY))
    End Function

    Private Shared Function ScreenForSavedBounds(bounds As System.Drawing.Rectangle) As System.Windows.Forms.Screen
        For Each monitor As System.Windows.Forms.Screen In System.Windows.Forms.Screen.AllScreens
            Dim visible As System.Drawing.Rectangle = System.Drawing.Rectangle.Intersect(bounds, monitor.WorkingArea)
            If visible.Width >= System.Math.Min(bounds.Width, 100) AndAlso
               visible.Height >= System.Math.Min(bounds.Height, 100) Then Return monitor
        Next
        Return PreferredScreen()
    End Function

    Friend Shared Function IsMeaningfullyVisible(bounds As System.Drawing.Rectangle) As Boolean
        For Each monitor As System.Windows.Forms.Screen In System.Windows.Forms.Screen.AllScreens
            Dim overlap As System.Drawing.Rectangle = System.Drawing.Rectangle.Intersect(bounds, monitor.WorkingArea)
            If overlap.Width >= System.Math.Min(bounds.Width, 100) AndAlso
               overlap.Height >= System.Math.Min(bounds.Height, 100) Then Return True
        Next
        Return False
    End Function

    Friend Shared Sub RestoreAssistant(window As System.Windows.Forms.Form, config As ClientConfig)
        window.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        If config.AssistantLeft.HasValue AndAlso config.AssistantTop.HasValue Then
            Dim saved As New System.Drawing.Point(config.AssistantLeft.Value, config.AssistantTop.Value)
            Dim savedBounds As New System.Drawing.Rectangle(saved, window.Size)
            Dim area As System.Drawing.Rectangle = ScreenForSavedBounds(savedBounds).WorkingArea
            window.Location = ClampToArea(saved, window.Size, area)
        Else
            CenterOnActiveScreen(window)
        End If
    End Sub

    Friend Shared Sub EnsureVisible(window As System.Windows.Forms.Form)
        Dim screen As System.Windows.Forms.Screen = ScreenForSavedBounds(window.Bounds)
        window.Location = ClampToArea(window.Location, window.Size, screen.WorkingArea)
    End Sub

    Friend Shared Sub CenterOnActiveScreen(window As System.Windows.Forms.Form)
        Dim area As System.Drawing.Rectangle = PreferredScreen().WorkingArea
        Dim target As New System.Drawing.Point(area.Left + (area.Width - window.Width) \ 2,
                                               area.Top + (area.Height - window.Height) \ 2)
        window.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        window.Location = ClampToArea(target, window.Size, area)
    End Sub

    Friend Shared Sub CenterOverOwner(window As System.Windows.Forms.Form, owner As System.Windows.Forms.Form)
        Dim area As System.Drawing.Rectangle = System.Windows.Forms.Screen.FromControl(owner).WorkingArea
        Dim target As New System.Drawing.Point(owner.Left + (owner.Width - window.Width) \ 2,
                                               owner.Top + (owner.Height - window.Height) \ 2)
        window.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        window.Location = ClampToArea(target, window.Size, area)
    End Sub
End Class
