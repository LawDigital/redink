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
'   - The form size is calculated to fit the content, honoring the provided
'     minimum width/height arguments.
'   - `UpdateMessage` updates the label text and re-measures its size.
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
            ''' Initializes a new splash screen with a message and optional sizing constraints.
            ''' </summary>
            ''' <param name="customText">Initial message shown next to the logo.</param>
            ''' <param name="formWidth">Minimum form width.</param>
            ''' <param name="formHeight">Minimum form height (note: current sizing logic uses it only as a minimum width/height input).</param>
            Public Sub New(Optional customText As String = "Please wait ...", Optional formWidth As Integer = 300, Optional formHeight As Integer = 100)
                ' Set the form properties
                Me.Text = $"{SharedMethods.AN}"
                Me.FormBorderStyle = FormBorderStyle.None
                Me.StartPosition = FormStartPosition.CenterScreen
                Me.Top -= 40
                Me.BackColor = ColorTranslator.FromWin32(&H8000000F)

                ' Set a predefined font for consistency
                Dim standardFont As New System.Drawing.Font("Segoe UI", 10.0F, FontStyle.Regular, GraphicsUnit.Point)

                ' Create the PictureBox
                Dim bmp As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                Dim pictureBox As New PictureBox()
                pictureBox.Image = bmp
                pictureBox.SizeMode = PictureBoxSizeMode.Zoom
                pictureBox.SetBounds(10, 10, 30, 30)

                ' Create the Label with updated font
                Label = New System.Windows.Forms.Label()
                Label.Text = customText
                Label.Font = standardFont
                Label.AutoSize = True
                Label.TextAlign = System.Drawing.ContentAlignment.MiddleLeft

                ' Dynamically calculate the label width
                Dim labelSize As Size = TextRenderer.MeasureText(Label.Text, standardFont)
                Label.SetBounds(pictureBox.Right + 10, 15, labelSize.Width, labelSize.Height)

                ' Adjust the form size dynamically based on the provided dimensions
                Dim contentWidth As Integer = pictureBox.Width + Label.Width + 40 ' Add padding for spacing
                Dim contentHeight As Integer = Math.Max(pictureBox.Height + 20, Label.Height + 30) ' Align to bottom of logo
                Me.ClientSize = New System.Drawing.Size(Math.Max(formWidth, contentWidth), Math.Max(formHeight, contentHeight))
                pictureBox.Top = (Me.ClientSize.Height - pictureBox.Height) \ 2

                ' Add the controls to the form
                Me.Controls.Add(pictureBox)
                Me.Controls.Add(Label)
            End Sub

            ''' <summary>
            ''' Updates the displayed message text and re-measures the label size.
            ''' </summary>
            ''' <param name="newMessage">New message to display.</param>
            Public Sub UpdateMessage(newMessage As String)
                Label.Text = newMessage
                Dim newSize As Size = TextRenderer.MeasureText(newMessage, Label.Font)
                Label.Size = newSize
                Label.Refresh()
            End Sub

        End Class
    End Class
End Namespace