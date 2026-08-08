' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ShowMethods.AskUser.vb
' Purpose: Shared modal dialog backing the internal ask_user tool. Presents one
'          Markdown-rendered question, optional concrete choices (large clickable
'          buttons; single- or multi-select), and an optional free-text / numeric
'          answer field. Host- and UI-agnostic. Never reached in unattended
'          e-mail Scheduler or AutoPilot runs (guarded by InteractivityProvider).
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Globalization
Imports System.Windows.Forms
Imports Markdig
Imports SharedLibrary.Agents

Namespace SharedLibrary
    Partial Public Class SharedMethods

        Public Shared Function ShowAskUserDialog(request As AskUserRequest) As AskUserResult
            If request Is Nothing Then
                Return New AskUserResult() With {.Status = "cancelled"}
            End If

            Dim result As AskUserResult = Nothing
            Dim uiThread As New System.Threading.Thread(
                Sub()
                    result = ShowAskUserDialogCore(request)
                End Sub)
            uiThread.SetApartmentState(System.Threading.ApartmentState.STA)
            uiThread.IsBackground = True
            uiThread.Start()
            uiThread.Join()

            Return If(result, New AskUserResult() With {.Status = "cancelled"})
        End Function

        Private Shared Function ShowAskUserDialogCore(request As AskUserRequest) As AskUserResult
            Dim answer As New AskUserResult() With {.Status = "cancelled"}

            Dim wa As Rectangle = Screen.FromPoint(Cursor.Position).WorkingArea
            Dim standardFont As New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            Dim contentWidth As Integer = Math.Min(560, wa.Width - 120)

            Dim inputForm As New Form() With {
                .Opacity = 0,
                .Text = "Red Ink",
                .FormBorderStyle = FormBorderStyle.Sizable,
                .StartPosition = FormStartPosition.Manual,
                .MinimizeBox = False,
                .MaximizeBox = True,
                .ShowInTaskbar = False,
                .TopMost = True,
                .AutoScaleMode = AutoScaleMode.Dpi,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Font = standardFont
            }

            Try
                Dim bmp As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                inputForm.Icon = Icon.FromHandle(bmp.GetHicon())
            Catch
            End Try

            Dim hasOptions As Boolean = request.Options IsNot Nothing AndAlso request.Options.Count > 0
            Dim allowText As Boolean = request.AllowFreeText OrElse Not hasOptions
            Dim isNumeric As Boolean = (request.InputType = "integer" OrElse request.InputType = "number")

            ' Root single-column table; explicit row management so order is guaranteed.
            Dim root As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .AutoScroll = True,
                .Padding = New Padding(16)
            }
            root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))

            Dim nextRow As Integer = 0
            Dim inputRowIndex As Integer = -1
            Dim addRow As Func(Of Control, Integer) =
                Function(c As Control) As Integer
                    root.RowStyles.Add(New RowStyle(SizeType.AutoSize))
                    root.RowCount = nextRow + 1
                    root.Controls.Add(c, 0, nextRow)
                    nextRow += 1
                    Return nextRow - 1
                End Function

            ' 1) Question (Markdown, auto-height).
            Dim questionView As New WebBrowser() With {
                .IsWebBrowserContextMenuEnabled = False,
                .AllowWebBrowserDrop = False,
                .ScriptErrorsSuppressed = True,
                .WebBrowserShortcutsEnabled = False,
                .TabStop = False,
                .Anchor = AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top,
                .Width = contentWidth,
                .Height = 44,
                .Margin = New Padding(0, 0, 0, 12)
            }
            AddHandler questionView.NewWindow,
                Sub(s As Object, e As System.ComponentModel.CancelEventArgs)
                    e.Cancel = True
                End Sub
            AddHandler questionView.DocumentCompleted,
                Sub(s As Object, e As WebBrowserDocumentCompletedEventArgs)
                    Try
                        Dim hgt As Integer = questionView.Document.Body.ScrollRectangle.Height
                        questionView.Height = Math.Max(36, Math.Min(hgt + 6, 260))
                    Catch
                    End Try
                End Sub
            questionView.DocumentText = BuildAskUserQuestionHtml(request.Question, SystemColors.Control)
            addRow(questionView)

            ' 2) Predefined answer buttons.
            Dim orderedOptions As New List(Of AskUserOption)()
            Dim optionButtons As New List(Of Button)()
            Dim selectedFlags As New List(Of Boolean)()

            If hasOptions Then
                For Each opt In request.Options
                    If opt Is Nothing Then Continue For
                    orderedOptions.Add(opt)
                Next

                Dim optionsTable As New TableLayoutPanel() With {
                    .ColumnCount = 1,
                    .AutoSize = True,
                    .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    .Anchor = AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top,
                    .Width = contentWidth,
                    .Margin = New Padding(0, 0, 0, 12)
                }
                optionsTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))

                For i As Integer = 0 To orderedOptions.Count - 1
                    Dim o As AskUserOption = orderedOptions(i)
                    Dim caption As String = If(String.IsNullOrWhiteSpace(o.Description),
                                               If(o.Label, ""),
                                               If(o.Label, "") & Environment.NewLine & o.Description)

                    ' FlatStyle.Standard so Padding is honored (left/right breathing room for text).
                    Dim b As New Button() With {
                        .Text = caption,
                        .Tag = i,
                        .Dock = DockStyle.Top,
                        .AutoSize = True,
                        .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                        .MinimumSize = New Size(0, 44),
                        .TextAlign = ContentAlignment.MiddleLeft,
                        .Padding = New Padding(14, 8, 14, 8),
                        .Margin = New Padding(0, 0, 0, 8),
                        .UseVisualStyleBackColor = True,
                        .FlatStyle = FlatStyle.Standard,
                        .Font = standardFont
                    }
                    selectedFlags.Add(False)
                    optionButtons.Add(b)
                    optionsTable.RowStyles.Add(New RowStyle(SizeType.AutoSize))
                    optionsTable.RowCount = i + 1
                    optionsTable.Controls.Add(b, 0, i)
                Next

                For i As Integer = 0 To optionButtons.Count - 1
                    Dim idx As Integer = i
                    AddHandler optionButtons(i).Click,
                        Sub(s As Object, e As EventArgs)
                            If request.MultiSelect Then
                                selectedFlags(idx) = Not selectedFlags(idx)
                                optionButtons(idx).BackColor =
                                    If(selectedFlags(idx), Color.FromArgb(210, 229, 247), SystemColors.Control)
                                optionButtons(idx).UseVisualStyleBackColor = Not selectedFlags(idx)
                            Else
                                answer.Status = "answered"
                                answer.SelectedOptionIds = New List(Of String) From {orderedOptions(idx).Id}
                                answer.FreeText = Nothing
                                inputForm.DialogResult = DialogResult.OK
                                inputForm.Close()
                            End If
                        End Sub
                Next

                addRow(optionsTable)
            End If

            ' 3) "Or type your own answer:" label + 4) input field.
            Dim txtInput As TextBox = Nothing
            If allowText Then
                Dim caption As String
                Select Case request.InputType
                    Case "integer" : caption = "Enter a whole number:"
                    Case "number" : caption = "Enter a number:"
                    Case Else : caption = If(hasOptions, "Or type your own answer:", "Your answer:")
                End Select

                Dim lblInput As New Label() With {
                    .Text = caption,
                    .Font = standardFont,
                    .AutoSize = True,
                    .Margin = New Padding(0, 0, 0, 4)
                }
                addRow(lblInput)

                txtInput = New TextBox() With {
                    .Font = standardFont,
                    .Multiline = Not isNumeric,
                    .WordWrap = True,
                    .AcceptsReturn = Not isNumeric,
                    .ScrollBars = If(isNumeric, ScrollBars.None, ScrollBars.Vertical),
                    .Width = contentWidth,
                    .Margin = New Padding(0, 0, 0, 4)
                }
                If isNumeric Then
                    ' Single-line numeric input: fixed height, stretches horizontally only.
                    txtInput.Anchor = AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top
                    txtInput.Height = TextRenderer.MeasureText("Wy", standardFont).Height + 8
                    addRow(txtInput)
                Else
                    ' Multiline free-text: fills its row so it grows/shrinks with the form.
                    txtInput.Dock = DockStyle.Fill
                    txtInput.Height = 96
                    txtInput.MinimumSize = New Size(0, 96)
                    ' Measured as AutoSize first (so the button row is included in the
                    ' shrink-wrap height); promoted to a growing row in the Shown handler.
                    inputRowIndex = addRow(txtInput)
                End If
            End If

            ' 5) OK / Cancel — flow from the right, OK rightmost.
            Dim okButton As New Button() With {.Text = "OK", .AutoSize = True, .Font = standardFont, .Margin = New Padding(8, 3, 0, 3)}
            Dim cancelButton As New Button() With {.Text = "Cancel", .AutoSize = True, .Font = standardFont, .Margin = New Padding(0, 3, 0, 3)}
            okButton.Visible = allowText OrElse request.MultiSelect

            Dim bottomFlow As New FlowLayoutPanel() With {
                .FlowDirection = FlowDirection.RightToLeft,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .WrapContents = False,
                .Dock = DockStyle.Fill,
                .Margin = New Padding(0, 14, 0, 0)
            }
            If okButton.Visible Then bottomFlow.Controls.Add(okButton)   ' rightmost
            bottomFlow.Controls.Add(cancelButton)                        ' left of OK
            addRow(bottomFlow)

            inputForm.Controls.Add(root)
            inputForm.CancelButton = cancelButton
            If okButton.Visible Then inputForm.AcceptButton = okButton

            Dim commitOk As Action =
                Sub()
                    Dim typed As String = If(txtInput IsNot Nothing, txtInput.Text.Trim(), "")
                    Dim selectedIds As New List(Of String)()

                    If request.MultiSelect Then
                        For i As Integer = 0 To orderedOptions.Count - 1
                            If i < selectedFlags.Count AndAlso selectedFlags(i) Then
                                selectedIds.Add(orderedOptions(i).Id)
                            End If
                        Next
                    End If

                    Dim hasSelection As Boolean = selectedIds.Count > 0
                    Dim hasTextAnswer As Boolean = typed.Length > 0

                    If Not hasSelection AndAlso Not hasTextAnswer Then
                        ShowCustomMessageBox("Please choose an option or enter an answer.")
                        Return
                    End If

                    If hasTextAnswer AndAlso Not hasSelection Then
                        If request.InputType = "integer" Then
                            Dim iv As Long
                            If Not Long.TryParse(typed, NumberStyles.Integer, CultureInfo.InvariantCulture, iv) Then
                                ShowCustomMessageBox("Please enter a valid whole number.")
                                Return
                            End If
                        ElseIf request.InputType = "number" Then
                            Dim dv As Double
                            If Not Double.TryParse(typed, NumberStyles.Float Or NumberStyles.AllowThousands, CultureInfo.InvariantCulture, dv) AndAlso
                               Not Double.TryParse(typed, NumberStyles.Float Or NumberStyles.AllowThousands, CultureInfo.CurrentCulture, dv) Then
                                ShowCustomMessageBox("Please enter a valid number.")
                                Return
                            End If
                        End If
                    End If

                    answer.Status = "answered"
                    answer.SelectedOptionIds = selectedIds
                    answer.FreeText = If(hasTextAnswer, typed, Nothing)
                    inputForm.DialogResult = DialogResult.OK
                    inputForm.Close()
                End Sub

            AddHandler okButton.Click, Sub(sender, e) commitOk()
            AddHandler cancelButton.Click,
                Sub(sender, e)
                    inputForm.DialogResult = DialogResult.Cancel
                    inputForm.Close()
                End Sub

            If txtInput IsNot Nothing Then
                AddHandler txtInput.KeyDown,
                    Sub(sender, e)
                        If isNumeric AndAlso e.KeyCode = Keys.Enter Then
                            e.SuppressKeyPress = True
                            commitOk()
                        ElseIf Not isNumeric AndAlso e.KeyCode = Keys.Enter AndAlso e.Modifiers = Keys.Control Then
                            e.SuppressKeyPress = True
                            commitOk()
                        ElseIf e.KeyCode = Keys.Escape Then
                            e.SuppressKeyPress = True
                            inputForm.DialogResult = DialogResult.Cancel
                            inputForm.Close()
                        End If
                    End Sub
            End If

            AddHandler inputForm.Shown,
                Sub()
                    inputForm.PerformLayout()

                    Dim maxW As Integer = wa.Width - 40
                    Dim maxH As Integer = wa.Height - 40
                    Dim w As Integer = Math.Min(inputForm.Width, maxW)
                    Dim h As Integer = Math.Min(inputForm.Height, maxH)

                    inputForm.AutoSize = False
                    root.AutoSize = False
                    ' Prevent shrinking below the fully-visible content size.
                    inputForm.MinimumSize = New Size(w, h)
                    inputForm.Size = New Size(w, h)
                    ' Now that the correct height (incl. buttons) is measured and locked,
                    ' let the multiline input row absorb extra vertical space on resize.
                    If inputRowIndex >= 0 Then
                        root.RowStyles(inputRowIndex) = New RowStyle(SizeType.Percent, 100.0F)
                    End If

                    inputForm.Location = New Point(
                        wa.X + (wa.Width - inputForm.Width) \ 2,
                        wa.Y + (wa.Height - inputForm.Height) \ 2)

                    ForceDialogToForeground(inputForm)
                    SharedMethods.AttachForeignForegroundWatchdog(inputForm)
                    inputForm.Opacity = 1

                    If txtInput IsNot Nothing Then
                        Try : txtInput.Focus() : Catch : End Try
                    End If
                End Sub

            Dim dr As DialogResult
            Try
                dr = inputForm.ShowDialog()
            Finally
                inputForm.Dispose()
            End Try

            If dr <> DialogResult.OK Then
                Return New AskUserResult() With {.Status = "cancelled"}
            End If

            Return answer
        End Function

        ''' <summary>
        ''' Renders the ask_user question as a small Markdown-enabled HTML document
        ''' whose background matches the given control color so it blends with the form.
        ''' </summary>
        Private Shared Function BuildAskUserQuestionHtml(question As String, back As Color) As String
            Dim bodyHtml As String
            Try
                Dim pipeline As MarkdownPipeline =
                    New MarkdownPipelineBuilder().
                        UseAdvancedExtensions().
                        UseSoftlineBreakAsHardlineBreak().
                        Build()
                bodyHtml = Markdown.ToHtml(If(question, ""), pipeline)
            Catch
                bodyHtml = "<p>" & System.Net.WebUtility.HtmlEncode(If(question, "")) & "</p>"
            End Try

            Dim bg As String = String.Format("#{0:X2}{1:X2}{2:X2}", back.R, back.G, back.B)

            Return "<!DOCTYPE html><html><head><meta charset=""utf-8"">" &
                   "<meta http-equiv=""X-UA-Compatible"" content=""IE=edge"">" &
                   "<style>" &
                   "html,body{margin:0;padding:0;background:" & bg & ";overflow:hidden;}" &
                   "body{font-family:'Segoe UI',sans-serif;font-size:11pt;color:#1b1b1b;line-height:1.35;}" &
                   "p{margin:0 0 6px 0;} ul,ol{margin:0 0 6px 20px;padding:0;}" &
                   "code{background:#e6e6e6;padding:1px 4px;border-radius:3px;}" &
                   "h1,h2,h3{font-size:12pt;margin:0 0 6px 0;} strong{font-weight:600;}" &
                   "</style></head><body>" & bodyHtml & "</body></html>"
        End Function

    End Class
End Namespace
