' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SplashScreenWorks.vb
' Purpose: Provides a borderless WinForms splash form displaying a message (optionally
'          with a seconds countdown) and a logo.
'
' Architecture:
'  - UI Composition: A `PictureBox` (logo) and `Label` (message) arranged side-by-side.
'  - Countdown: `StartCountdown()` runs a background `Task` that delays in 1-second ticks and
'    marshals UI updates to the UI thread via `Invoke` when required.
'  - Message Updates: `UpdateMessage()` changes the label text and recalculates its size.
'  - Cancellation: ESC cancels the countdown and raises `CancelRequested` before closing the form.
'  - Window Interaction: Uses Win32 `ReleaseCapture`/`SendMessage` to allow dragging a borderless form.
' =============================================================================


Option Strict On
Option Explicit On

Namespace SharedLibrary

    ''' <summary>
    ''' Borderless splash screen form showing a logo and message with an optional countdown.
    ''' </summary>
    Public Class SplashScreenWorks
        Inherits System.Windows.Forms.Form

        ' ─── Controls & state ────────────────────────────────────────
        Private lblMessage As System.Windows.Forms.Label
        Private picLogo As System.Windows.Forms.PictureBox
        Private remainingSeconds As Integer
        Private baseText As String
        Private countdownCts As System.Threading.CancellationTokenSource

        ''' <summary>
        ''' Raised when the user presses ESC while the splash screen has focus.
        ''' </summary>
        Public Event CancelRequested As System.EventHandler

        ' ─── WinAPI for borderless dragging ───────────────────────────
        ''' <summary>
        ''' Releases the current mouse capture to allow initiating a window move drag operation.
        ''' </summary>
        <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
        Private Shared Function ReleaseCapture() As Boolean
        End Function

        ''' <summary>
        ''' Sends a window message used here to emulate dragging a caption on a borderless form.
        ''' </summary>
        <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
        Private Shared Function SendMessage(
    ByVal hWnd As IntPtr,
    ByVal wMsg As Integer,
    ByVal wParam As IntPtr,
    ByVal lParam As IntPtr
) As IntPtr
        End Function

        ''' <summary>
        ''' Window message for non-client left button down (caption drag is initiated using HTCAPTION).
        ''' </summary>
        Private Const WM_NCLBUTTONDOWN As Integer = &HA1

        ''' <summary>
        ''' Hit-test value indicating the title bar / caption area.
        ''' </summary>
        Private Const HTCAPTION As Integer = 2

        ''' <summary>
        ''' Initializes a new splash screen instance.
        ''' </summary>
        ''' <param name="customText">Prefix text shown in the label.</param>
        ''' <param name="formWidth">Client width override (0 keeps auto-size behavior).</param>
        ''' <param name="formHeight">Client height override (0 keeps auto-size behavior).</param>
        ''' <param name="countdownSeconds">Initial countdown length in seconds (0 disables countdown).</param>
        Public Sub New(
    Optional ByVal customText As String = "Please wait …",
    Optional ByVal formWidth As Integer = 0,
    Optional ByVal formHeight As Integer = 0,
    Optional ByVal countdownSeconds As Integer = 0)

            MyBase.New()

            ' ─── Form setup ───────────────────────────────────────────
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.BackColor = System.Drawing.ColorTranslator.FromWin32(&H8000000F)
            Me.KeyPreview = True

            ' ─── Logo ─────────────────────────────────────────────────
            picLogo = New System.Windows.Forms.PictureBox() With {
        .Image = New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard)),
        .SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
    }
            Me.Controls.Add(picLogo)

            ' ─── Label ───────────────────────────────────────────────
            Dim stdFont As System.Drawing.Font =
        New System.Drawing.Font("Segoe UI", 10.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
            lblMessage = New System.Windows.Forms.Label() With {
        .Font = stdFont,
        .AutoSize = True,
        .TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    }
            Me.Controls.Add(lblMessage)

            ' ─── Layout & initial text ───────────────────────────────
            baseText = customText
            remainingSeconds = countdownSeconds
            Dim initialText As String = If(countdownSeconds > 0,
                                   $"{customText} {countdownSeconds}s",
                                   customText)
            lblMessage.Text = initialText

            Dim padding As Integer = 10
            Dim textSize As System.Drawing.Size = System.Windows.Forms.TextRenderer.MeasureText(initialText, stdFont)
            lblMessage.Size = textSize

            ' Logo height matches text height for equal top/bottom padding
            Dim logoSize As Integer = textSize.Height
            picLogo.SetBounds(padding, padding, logoSize, logoSize)

            ' Center label vertically next to logo
            Dim labelX As Integer = picLogo.Right + padding
            Dim labelY As Integer = padding + (logoSize - textSize.Height) \ 2
            lblMessage.SetBounds(labelX, labelY, textSize.Width, textSize.Height)

            ' Auto-size form to content (unless overridden)
            Dim clientW As Integer = lblMessage.Right + padding
            Dim clientH As Integer = logoSize + padding * 2
            If formWidth > 0 Then clientW = formWidth
            If formHeight > 0 Then clientH = formHeight
            Me.ClientSize = New System.Drawing.Size(clientW, clientH)

            ' ─── ESC cancels ──────────────────────────────────────────
            AddHandler Me.KeyDown, AddressOf OnKeyDown

            ' ─── Start countdown if needed ───────────────────────────
            If countdownSeconds > 0 Then
                StartCountdown()
            End If
        End Sub

        ''' <summary>
        ''' Updates the label text and recalculates its size.
        ''' </summary>
        ''' <param name="newMessage">The new message to render in the label.</param>
        Public Sub UpdateMessage(ByVal newMessage As String)
            lblMessage.Text = newMessage
            Dim newSize As System.Drawing.Size = System.Windows.Forms.TextRenderer.MeasureText(newMessage, lblMessage.Font)
            lblMessage.Size = newSize
            lblMessage.Refresh()
        End Sub

        ''' <summary>
        ''' Restarts the countdown from the specified value, optionally updating the base message prefix.
        ''' </summary>
        ''' <param name="seconds">New countdown duration in seconds.</param>
        ''' <param name="newBaseText">Optional replacement for the base message prefix.</param>
        Public Sub RestartCountdown(
    ByVal seconds As Integer,
    Optional ByVal newBaseText As String = Nothing)

            If newBaseText IsNot Nothing Then
                baseText = newBaseText
            End If

            remainingSeconds = seconds
            UpdateMessage($"{baseText} {remainingSeconds}s")
            StartCountdown()
        End Sub

        ''' <summary>
        ''' Starts (or restarts) the countdown task.
        ''' Cancels any previously started countdown via `countdownCts`.
        ''' </summary>
        Private Sub StartCountdown()
            ' Cancel previous if running
            countdownCts?.Cancel()

            countdownCts = New System.Threading.CancellationTokenSource()
            Dim ct = countdownCts.Token

            System.Threading.Tasks.Task.Run(Async Function()
                                                While remainingSeconds > 0 AndAlso Not ct.IsCancellationRequested
                                                    Try
                                                        Await System.Threading.Tasks.Task.Delay(1000, ct)
                                                    Catch ex As System.Threading.Tasks.TaskCanceledException
                                                        Exit While
                                                    End Try

                                                    remainingSeconds -= 1
                                                    If remainingSeconds < 0 Then remainingSeconds = 0

                                                    ' Update on UI thread
                                                    If Not Me.IsDisposed Then
                                                        If Me.InvokeRequired Then
                                                            Me.Invoke(Sub() UpdateMessage($"{baseText} {remainingSeconds}s"))
                                                        Else
                                                            UpdateMessage($"{baseText} {remainingSeconds}s")
                                                        End If
                                                    End If
                                                End While
                                            End Function)
        End Sub

        ''' <summary>
        ''' Handles ESC key press: cancels countdown, raises `CancelRequested`, and closes the form.
        ''' </summary>
        Private Sub OnKeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
            If e.KeyCode = System.Windows.Forms.Keys.Escape Then
                countdownCts?.Cancel()
                RaiseEvent CancelRequested(Me, System.EventArgs.Empty)
                Me.Close()
            End If
        End Sub

        ''' <summary>
        ''' Enables dragging the borderless form by sending a caption drag message on left mouse down.
        ''' </summary>
        ''' <param name="e">Mouse event data.</param>
        Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
            MyBase.OnMouseDown(e)
            If e.Button = System.Windows.Forms.MouseButtons.Left Then
                ReleaseCapture()
                SendMessage(Me.Handle, WM_NCLBUTTONDOWN, CType(HTCAPTION, IntPtr), IntPtr.Zero)
            End If
        End Sub

    End Class
End Namespace