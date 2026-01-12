' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: QuickTranslateWidget.vb
' Purpose: Provides a non-modal, resizable translation widget allowing users to
'          quickly translate text using the LLM. Features debounced auto-translate
'          on typing pause, Enter key translation, and clipboard copy support.
'
' Architecture:
'  - Non-modal Form with two side-by-side panels (input TextBox, output Label)
'  - Language input field with persistence via My.Settings
'  - Debounce timer (1 second) triggers translation after user stops typing
'  - Enter key also triggers immediate translation
'  - Buttons: Clear, Copy, Close
'  - Keyboard shortcuts: Escape closes, Ctrl+C copies result, Ctrl+L clears
'  - Spinner indicator during LLM call
'  - Window position/size persisted via My.Settings
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms

Namespace SharedLibrary
    ''' <summary>
    ''' A non-modal translation widget that provides quick LLM-powered translation.
    ''' </summary>
    Public Class QuickTranslateWidget
        Inherits Form

        Private Const DEBOUNCE_MS As Integer = 1000

        ' Controls
        Private WithEvents txtInput As TextBox
        Private lblOutput As Label
        Private WithEvents txtLanguage As TextBox
        Private WithEvents btnClear As Button
        Private WithEvents btnCopy As Button
        Private WithEvents btnClose As Button
        Private lblSpinner As Label
        Private pnlOutput As Panel

        ' Debounce timer
        Private debounceTimer As System.Windows.Forms.Timer

        ' Cancellation for ongoing translation
        Private _cts As CancellationTokenSource

        ' Callback to perform translation
        Private ReadOnly _translateFunc As Func(Of String, String, CancellationToken, Task(Of String))

        ' Default language from context
        Private ReadOnly _defaultLanguage As String

        ''' <summary>
        ''' Creates a new QuickTranslateWidget.
        ''' </summary>
        ''' <param name="translateFunc">
        ''' Async function that takes (textToTranslate, targetLanguage, cancellationToken) and returns the translated text.
        ''' </param>
        ''' <param name="defaultLanguage">The default target language (from INI_Language1).</param>
        Public Sub New(translateFunc As Func(Of String, String, CancellationToken, Task(Of String)),
                       defaultLanguage As String)
            _translateFunc = translateFunc
            _defaultLanguage = If(defaultLanguage, "English")
            InitializeComponent()
            RestoreSettings()
        End Sub

        Private Sub InitializeComponent()
            Me.Text = $"{SharedMethods.AN} Translate on-the-fly"
            Me.FormBorderStyle = FormBorderStyle.Sizable
            Me.MinimizeBox = True
            Me.MaximizeBox = False
            Me.ShowInTaskbar = True
            Me.TopMost = True
            Me.StartPosition = FormStartPosition.Manual
            Me.MinimumSize = New Size(500, 180)
            Me.Size = New Size(600, 220)
            Me.KeyPreview = True

            ' Set icon
            Try
                Dim bmp As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                Me.Icon = Icon.FromHandle(bmp.GetHicon())
            Catch
            End Try

            ' Font
            Dim stdFont As New Font("Segoe UI", 10.0F, FontStyle.Regular, GraphicsUnit.Point)
            Me.Font = stdFont

            ' Main layout: 2 rows
            ' Row 0: Input/Output panels (fill)
            ' Row 1: Language field + buttons (auto-size, minimal padding)
            Dim mainTable As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 2,
                .Padding = New Padding(10, 10, 10, 5),
                .Margin = New Padding(0)
            }
            mainTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
            mainTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
            mainTable.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            mainTable.RowStyles.Add(New RowStyle(SizeType.AutoSize))

            ' Input TextBox (left)
            txtInput = New TextBox() With {
                .Multiline = True,
                .ScrollBars = ScrollBars.Vertical,
                .Dock = DockStyle.Fill,
                .Font = stdFont,
                .AcceptsReturn = False
            }
            mainTable.Controls.Add(txtInput, 0, 0)

            ' Output panel with label (right) - scrollable
            pnlOutput = New Panel() With {
                .Dock = DockStyle.Fill,
                .AutoScroll = True,
                .BorderStyle = BorderStyle.FixedSingle,
                .BackColor = SystemColors.Window
            }
            lblOutput = New Label() With {
                .AutoSize = True,
                .MaximumSize = New Size(0, 0),
                .Font = stdFont,
                .Location = New Point(3, 3),
                .Cursor = Cursors.IBeam
            }
            pnlOutput.Controls.Add(lblOutput)
            mainTable.Controls.Add(pnlOutput, 1, 0)

            ' Bottom row: use a TableLayoutPanel for left (language) and right (buttons) alignment
            Dim bottomTable As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 1,
                .Margin = New Padding(0),
                .Padding = New Padding(0, 5, 0, 0),
                .AutoSize = True
            }
            bottomTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
            bottomTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
            bottomTable.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            mainTable.Controls.Add(bottomTable, 0, 1)
            mainTable.SetColumnSpan(bottomTable, 2)

            ' Left side: Language label + textbox + spinner
            Dim leftFlow As New FlowLayoutPanel() With {
                .FlowDirection = FlowDirection.LeftToRight,
                .Dock = DockStyle.Fill,
                .AutoSize = True,
                .WrapContents = False,
                .Margin = New Padding(0),
                .Padding = New Padding(0)
            }
            bottomTable.Controls.Add(leftFlow, 0, 0)

            ' Language label - aligned with textbox baseline
            Dim lblLang As New Label() With {
                .Text = "Language:",
                .AutoSize = True,
                .Margin = New Padding(0, 4, 5, 0),
                .TextAlign = ContentAlignment.MiddleLeft
            }
            leftFlow.Controls.Add(lblLang)

            ' Language textbox
            txtLanguage = New TextBox() With {
                .Width = 120,
                .Font = stdFont,
                .Margin = New Padding(0, 2, 0, 0)
            }
            leftFlow.Controls.Add(txtLanguage)

            ' Spinner label (hidden by default)
            lblSpinner = New Label() With {
                .Text = "⏳",
                .AutoSize = True,
                .Visible = False,
                .Margin = New Padding(10, 5, 0, 0)
            }
            leftFlow.Controls.Add(lblSpinner)

            ' Right side: Buttons right-aligned
            Dim rightFlow As New FlowLayoutPanel() With {
                .FlowDirection = FlowDirection.RightToLeft,
                .Dock = DockStyle.Fill,
                .AutoSize = True,
                .WrapContents = False,
                .Margin = New Padding(0),
                .Padding = New Padding(0)
            }
            bottomTable.Controls.Add(rightFlow, 1, 0)

            ' Buttons (added in reverse order for RightToLeft flow)
            btnClose = New Button() With {
                .Text = "Close",
                .AutoSize = True,
                .Margin = New Padding(0, 0, 0, 0)
            }
            rightFlow.Controls.Add(btnClose)

            btnCopy = New Button() With {
                .Text = "Copy",
                .AutoSize = True,
                .Margin = New Padding(5, 0, 0, 0)
            }
            rightFlow.Controls.Add(btnCopy)

            btnClear = New Button() With {
                .Text = "Clear",
                .AutoSize = True,
                .Margin = New Padding(5, 0, 0, 0)
            }
            rightFlow.Controls.Add(btnClear)

            Me.Controls.Add(mainTable)

            ' Debounce timer
            debounceTimer = New System.Windows.Forms.Timer() With {.Interval = DEBOUNCE_MS}
            AddHandler debounceTimer.Tick, AddressOf OnDebounceTimerTick

            ' Update output label max width on panel resize
            AddHandler pnlOutput.Resize, Sub()
                                             lblOutput.MaximumSize = New Size(Math.Max(50, pnlOutput.ClientSize.Width - 20), 0)
                                         End Sub

            ' Enable text selection on label via mouse (copies to clipboard on click)
            AddHandler lblOutput.MouseDown, AddressOf LblOutput_MouseDown
        End Sub

        Private Sub RestoreSettings()
            Try
                ' Restore language
                Dim savedLang As String = My.Settings.QuickTranslateLanguage
                txtLanguage.Text = If(String.IsNullOrWhiteSpace(savedLang), _defaultLanguage, savedLang)

                ' Restore position/size
                Dim x As Integer = My.Settings.QuickTranslateX
                Dim y As Integer = My.Settings.QuickTranslateY
                Dim w As Integer = My.Settings.QuickTranslateWidth
                Dim h As Integer = My.Settings.QuickTranslateHeight

                If w > 0 AndAlso h > 0 Then
                    Me.Size = New Size(w, h)
                End If

                ' Validate position is on screen (use > 0 to treat 0 as unset)
                If x > 0 AndAlso y > 0 Then
                    Dim wa As Rectangle = Screen.FromPoint(New Point(x, y)).WorkingArea
                    If x >= wa.Left AndAlso x < wa.Right - 50 AndAlso
                       y >= wa.Top AndAlso y < wa.Bottom - 50 Then
                        Me.Location = New Point(x, y)
                    Else
                        PositionOnScreen()
                    End If
                Else
                    PositionOnScreen()
                End If
            Catch
                txtLanguage.Text = _defaultLanguage
                PositionOnScreen()
            End Try
        End Sub

        Private Sub PositionOnScreen()
            Dim wa As Rectangle = Screen.FromPoint(Cursor.Position).WorkingArea
            Const MARGIN As Integer = 30
            Me.Location = New Point(
                wa.Right - Me.Width - MARGIN,
                wa.Top + MARGIN
            )
        End Sub

        Private Sub SaveSettings()
            Try
                My.Settings.QuickTranslateLanguage = txtLanguage.Text
                My.Settings.QuickTranslateX = Me.Location.X
                My.Settings.QuickTranslateY = Me.Location.Y
                My.Settings.QuickTranslateWidth = Me.Size.Width
                My.Settings.QuickTranslateHeight = Me.Size.Height
                My.Settings.Save()
            Catch
            End Try
        End Sub

        Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
            SaveSettings()
            CancelOngoingTranslation()
            debounceTimer?.Stop()
            MyBase.OnFormClosing(e)
        End Sub

        Private Sub txtInput_TextChanged(sender As Object, e As EventArgs) Handles txtInput.TextChanged
            ' Reset debounce timer
            debounceTimer.Stop()
            If Not String.IsNullOrWhiteSpace(txtInput.Text) Then
                debounceTimer.Start()
            End If
        End Sub

        Private Sub txtInput_KeyDown(sender As Object, e As KeyEventArgs) Handles txtInput.KeyDown
            If e.KeyCode = Keys.Enter AndAlso Not e.Shift Then
                e.SuppressKeyPress = True
                debounceTimer.Stop()
                PerformTranslationAsync()
            End If
        End Sub

        Private Sub OnDebounceTimerTick(sender As Object, e As EventArgs)
            debounceTimer.Stop()
            PerformTranslationAsync()
        End Sub

        Private Async Sub PerformTranslationAsync()
            Dim textToTranslate As String = txtInput.Text.Trim()
            Dim targetLanguage As String = txtLanguage.Text.Trim()

            If String.IsNullOrWhiteSpace(textToTranslate) Then
                lblOutput.Text = ""
                Return
            End If

            If String.IsNullOrWhiteSpace(targetLanguage) Then
                targetLanguage = _defaultLanguage
            End If

            ' Cancel any ongoing translation
            CancelOngoingTranslation()

            _cts = New CancellationTokenSource()
            Dim token As CancellationToken = _cts.Token

            ' Show spinner
            lblSpinner.Visible = True
            lblOutput.Text = ""

            Try
                ' Run translation on background, but capture result
                Dim result As String = Await Task.Run(
                    Async Function()
                        Return Await _translateFunc(textToTranslate, targetLanguage, token)
                    End Function, token).ConfigureAwait(False)

                ' Marshal back to UI thread for control updates
                If Not token.IsCancellationRequested Then
                    If Me.InvokeRequired Then
                        Me.BeginInvoke(Sub()
                                           lblOutput.Text = If(result, "")
                                           lblSpinner.Visible = False
                                       End Sub)
                    Else
                        lblOutput.Text = If(result, "")
                        lblSpinner.Visible = False
                    End If
                End If

            Catch ex As OperationCanceledException
                ' Cancelled - ignore
            Catch ex As Exception
                If Not token.IsCancellationRequested Then
                    If Me.InvokeRequired Then
                        Me.BeginInvoke(Sub()
                                           lblOutput.Text = $"Error: {ex.Message}"
                                           lblSpinner.Visible = False
                                       End Sub)
                    Else
                        lblOutput.Text = $"Error: {ex.Message}"
                        lblSpinner.Visible = False
                    End If
                End If
            Finally
                ' Ensure spinner is hidden even if we didn't hit above paths
                If Me.InvokeRequired Then
                    Me.BeginInvoke(Sub() lblSpinner.Visible = False)
                Else
                    lblSpinner.Visible = False
                End If
            End Try
        End Sub

        Private Sub CancelOngoingTranslation()
            Try
                _cts?.Cancel()
                _cts?.Dispose()
                _cts = Nothing
            Catch
            End Try
        End Sub

        Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
            ClearAll()
        End Sub

        Private Sub ClearAll()
            CancelOngoingTranslation()
            txtInput.Text = ""
            lblOutput.Text = ""
            txtInput.Focus()
        End Sub

        Private Sub btnCopy_Click(sender As Object, e As EventArgs) Handles btnCopy.Click
            CopyResult()
        End Sub

        Private Sub CopyResult()
            Dim text As String = lblOutput.Text
            If Not String.IsNullOrEmpty(text) Then
                Try
                    Clipboard.SetText(text.Trim())
                Catch
                End Try
            End If
        End Sub

        Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
            Me.Close()
        End Sub

        Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
            MyBase.OnKeyDown(e)

            If e.KeyCode = Keys.Escape Then
                Me.Close()
                e.Handled = True
            ElseIf e.Control AndAlso e.KeyCode = Keys.C AndAlso Not txtInput.Focused Then
                CopyResult()
                e.Handled = True
            ElseIf e.Control AndAlso e.KeyCode = Keys.L Then
                ClearAll()
                e.Handled = True
            End If
        End Sub

        Private Sub LblOutput_MouseDown(sender As Object, e As MouseEventArgs)
            ' Allow selecting text by copying to clipboard on click
            If Not String.IsNullOrEmpty(lblOutput.Text) Then
                Try
                    Clipboard.SetText(lblOutput.Text)
                    ' Brief visual feedback
                    Dim original As Color = lblOutput.BackColor
                    lblOutput.BackColor = Color.LightGreen
                    Task.Delay(200).ContinueWith(Sub() Me.BeginInvoke(Sub() lblOutput.BackColor = original))
                Catch
                End Try
            End If
        End Sub

        ''' <summary>
        ''' Shows the widget non-modally. If already visible, brings to front.
        ''' </summary>
        Public Sub ShowWidget()
            If Me.Visible Then
                Me.BringToFront()
                Me.Activate()
            Else
                Me.Show()
            End If
            txtInput.Focus()
        End Sub
    End Class
End Namespace