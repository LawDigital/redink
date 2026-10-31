' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.SplashScreen.vb
' Purpose:
'   Provides a small, borderless WinForms splash screen showing the application
'   logo and a single message line.
'
' How it works:
'   - The constructor creates a borderless form centered on screen, applies a
'     standard font, and adds:
'       - a PictureBox showing `My.Resources.Red_Ink_Logo`
'       - a label showing the provided message text
'   - The form size is calculated to fit the content tightly.
'   - `UpdateMessage` updates the label text, re-measures size, and resizes the form.
'
' KEY ARCHITECTURE
'   • Surface: SharedMethods, SplashScreen (VB implementation).
'   • Entry points/helpers: UpdateMessage.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Windows.Forms

Namespace SharedLibrary

    Partial Public Class SharedMethods
        ''' <summary>
        ''' Borderless splash form that displays the application logo and a single message line.
        ''' </summary>
        Public Class SplashScreen

            Inherits Form

            ''' <summary>
            ''' Label used to display the current splash message.
            ''' </summary>
            Private Label As System.Windows.Forms.Label

            ''' <summary>
            ''' PictureBox for the logo.
            ''' </summary>
            Private LogoPictureBox As System.Windows.Forms.PictureBox

            ''' <summary>
            ''' Padding around content.
            ''' </summary>
            Private Const Padding As Integer = 10

            ''' <summary>
            ''' Logo size (square).
            ''' </summary>
            Private Const LogoSize As Integer = 30

            ''' <summary>
            ''' Initializes a new splash screen with a message.
            ''' </summary>
            ''' <param name="customText">Initial message shown next to the logo.</param>
            ''' <param name="formWidth">Minimum form width (optional).</param>
            ''' <param name="formHeight">Minimum form height (optional).</param>
            Public Sub New(Optional customText As String = "Please wait ...", Optional formWidth As Integer = 0, Optional formHeight As Integer = 0)
                ' Set the form properties
                Me.Text = $"{SharedMethods.AN}"
                Me.FormBorderStyle = FormBorderStyle.None
                Me.StartPosition = FormStartPosition.CenterScreen
                Me.BackColor = ColorTranslator.FromWin32(&H8000000F)

                ' Set a predefined font for consistency
                Dim standardFont As New System.Drawing.Font("Segoe UI", 10.0F, FontStyle.Regular, GraphicsUnit.Point)

                ' Create the PictureBox
                Dim bmp As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                LogoPictureBox = New PictureBox()
                LogoPictureBox.Image = bmp
                LogoPictureBox.SizeMode = PictureBoxSizeMode.Zoom
                LogoPictureBox.SetBounds(Padding, Padding, LogoSize, LogoSize)

                ' Create the Label with updated font
                Label = New System.Windows.Forms.Label()
                Label.Text = customText
                Label.Font = standardFont
                Label.AutoSize = False
                Label.TextAlign = System.Drawing.ContentAlignment.MiddleLeft

                ' Calculate sizes and layout
                Dim labelSize As Size = TextRenderer.MeasureText(customText, standardFont)

                ' Label position: right of logo with padding
                Dim labelX As Integer = LogoPictureBox.Right + Padding

                ' Calculate form dimensions - tight fit
                Dim contentWidth As Integer = labelX + labelSize.Width + Padding
                Dim contentHeight As Integer = Math.Max(LogoSize, labelSize.Height) + (Padding * 2)

                ' Apply minimum dimensions if specified
                If formWidth > 0 Then contentWidth = Math.Max(formWidth, contentWidth)
                If formHeight > 0 Then contentHeight = Math.Max(formHeight, contentHeight)

                Me.ClientSize = New System.Drawing.Size(contentWidth, contentHeight)

                ' Center logo and label vertically
                LogoPictureBox.Top = (Me.ClientSize.Height - LogoSize) \ 2
                Dim labelTop As Integer = (Me.ClientSize.Height - labelSize.Height) \ 2
                Label.SetBounds(labelX, labelTop, labelSize.Width, labelSize.Height)

                ' Add the controls to the form
                Me.Controls.Add(LogoPictureBox)
                Me.Controls.Add(Label)

                ' Adjust position slightly up
                Me.Top -= 40
            End Sub

            ''' <summary>
            ''' Updates the displayed message text, re-measures the label size, and resizes the form.
            ''' </summary>
            ''' <param name="newMessage">New message to display.</param>
            Public Sub UpdateMessage(newMessage As String)
                Label.Text = newMessage
                Dim newSize As Size = TextRenderer.MeasureText(newMessage, Label.Font)

                ' Recalculate form width to fit new text
                Dim labelX As Integer = LogoPictureBox.Right + Padding
                Dim contentWidth As Integer = labelX + newSize.Width + Padding
                Dim contentHeight As Integer = Math.Max(LogoSize, newSize.Height) + (Padding * 2)

                Me.ClientSize = New System.Drawing.Size(contentWidth, contentHeight)

                ' Re-center vertically
                LogoPictureBox.Top = (Me.ClientSize.Height - LogoSize) \ 2
                Dim labelTop As Integer = (Me.ClientSize.Height - newSize.Height) \ 2
                Label.SetBounds(labelX, labelTop, newSize.Width, newSize.Height)

                Label.Refresh()
                Me.Refresh()
            End Sub

        End Class
    End Class
End Namespace