' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ShowMethods.AskUser.vb
' Purpose: Shared modal dialog backing the internal ask_user tool. Presents one
'          Markdown-rendered question, optional concrete choices (large clickable
'          Markdown-capable buttons; single- or multi-select), and an optional
'          free-text / numeric
'          answer field. Host- and UI-agnostic. Never reached in unattended
'          e-mail Scheduler or AutoPilot runs (guarded by InteractivityProvider).
'
'          Layout:
'            - The question/options area scrolls independently.
'            - Markdown height is measured from a dedicated content wrapper,
'              never from the WebBrowser viewport.
'            - Markdown height is re-measured after final width/DPI layout.
'            - The free-text area has a compact fixed logical height.
'            - The input area and footer are separate table rows.
'            - OK/Cancel therefore remain reachable regardless of prompt length.
'
'          Threading / ownership:
'            - If the caller is already STA, the dialog is shown on that thread.
'              This permits a real modal owner without blocking the owner's thread.
'            - If the caller is not STA, a dedicated STA thread is used.
'            - If an MTA caller itself owns the captured native owner window, that
'              owner is deliberately not passed across to the STA dialog thread;
'              doing so while the owner thread is blocked on Join can deadlock.
' =============================================================================

Option Strict On
Option Explicit On

Imports Markdig

Namespace SharedLibrary
    Partial Public Class SharedMethods

        Public Shared Function ShowAskUserDialog(
            request As Global.SharedLibrary.Agents.AskUserRequest
        ) As Global.SharedLibrary.Agents.AskUserResult

            If request Is Nothing Then
                Return New Global.SharedLibrary.Agents.AskUserResult() With {
                    .Status = "cancelled"
                }
            End If

            Dim ownerInfo As AskUserOwnerInfo = CaptureAskUserOwnerInfo()

            ' Preferred path: if we are already STA, stay on the caller thread.
            ' This is especially important when the caller is also the native owner
            ' thread (e.g. an Office host UI thread): ShowDialog(owner) can then run
            ' its normal nested modal loop instead of blocking that owner on Join().
            If System.Threading.Thread.CurrentThread.GetApartmentState() =
               System.Threading.ApartmentState.STA Then

                Return ShowAskUserDialogCore(request, ownerInfo.Handle)
            End If

            ' Fallback for MTA / unknown apartment callers: run the WinForms modal
            ' loop on a dedicated STA thread, as the previous implementation did.
            Dim result As Global.SharedLibrary.Agents.AskUserResult = Nothing
            Dim uiError As System.Exception = Nothing

            Dim callerNativeThreadId As System.UInt32 = AskUserGetCurrentThreadId()
            Dim ownerHandleForUiThread As System.IntPtr = ownerInfo.Handle

            ' Never pass an owner HWND to another UI thread while that HWND belongs
            ' to the caller thread that is about to block on Join(). That combination
            ' can deadlock through synchronous native window messages.
            If ownerInfo.ThreadId <> 0UI AndAlso
               ownerInfo.ThreadId = callerNativeThreadId Then

                ownerHandleForUiThread = System.IntPtr.Zero
            End If

            Dim uiThread As New System.Threading.Thread(
                Sub()
                    Try
                        result = ShowAskUserDialogCore(
                            request,
                            ownerHandleForUiThread
                        )
                    Catch ex As System.Exception
                        uiError = ex
                    End Try
                End Sub)

            uiThread.SetApartmentState(System.Threading.ApartmentState.STA)
            uiThread.IsBackground = True
            uiThread.Name = "Red Ink ask_user"
            uiThread.Start()
            uiThread.Join()

            If uiError IsNot Nothing Then
                System.Runtime.ExceptionServices.ExceptionDispatchInfo.
                    Capture(uiError).
                    Throw()
            End If

            Return If(
                result,
                New Global.SharedLibrary.Agents.AskUserResult() With {
                    .Status = "cancelled"
                }
            )
        End Function

        Private Shared Function ShowAskUserDialogCore(
            request As Global.SharedLibrary.Agents.AskUserRequest,
            ownerHandle As System.IntPtr
        ) As Global.SharedLibrary.Agents.AskUserResult

            Dim answer As New Global.SharedLibrary.Agents.AskUserResult() With {
                .Status = "cancelled"
            }

            Dim usableOwnerHandle As System.IntPtr = System.IntPtr.Zero

            Try
                If ownerHandle <> System.IntPtr.Zero AndAlso
                   AskUserIsWindow(ownerHandle) Then

                    usableOwnerHandle = ownerHandle
                End If
            Catch ex As System.Exception
                usableOwnerHandle = System.IntPtr.Zero
            End Try

            Dim targetScreen As System.Windows.Forms.Screen

            If usableOwnerHandle <> System.IntPtr.Zero Then
                targetScreen = System.Windows.Forms.Screen.FromHandle(
                    usableOwnerHandle
                )
            Else
                targetScreen = System.Windows.Forms.Screen.FromPoint(
                    System.Windows.Forms.Cursor.Position
                )
            End If

            Dim wa As System.Drawing.Rectangle = targetScreen.WorkingArea

            Using standardFont As New System.Drawing.Font(
                "Segoe UI",
                9.0F,
                System.Drawing.FontStyle.Regular,
                System.Drawing.GraphicsUnit.Point
            )

                Using inputForm As New System.Windows.Forms.Form()

                    inputForm.Opacity = 0
                    inputForm.Text = "Red Ink - Inky needs your input"
                    inputForm.FormBorderStyle =
                        System.Windows.Forms.FormBorderStyle.Sizable
                    inputForm.StartPosition =
                        System.Windows.Forms.FormStartPosition.Manual
                    inputForm.MinimizeBox = False
                    inputForm.MaximizeBox = True
                    inputForm.ShowInTaskbar = False
                    inputForm.TopMost = True
                    inputForm.AutoScaleMode =
                        System.Windows.Forms.AutoScaleMode.Dpi
                    inputForm.AutoScaleDimensions =
                        New System.Drawing.SizeF(96.0F, 96.0F)
                    inputForm.AutoSize = False
                    inputForm.Font = standardFont
                    inputForm.ClientSize = New System.Drawing.Size(760, 600)

                    Try
                        Dim bmp As New System.Drawing.Bitmap(
                            SharedMethods.GetLogoBitmap(
                                SharedMethods.LogoType.Standard
                            )
                        )
                        inputForm.Icon =
                            System.Drawing.Icon.FromHandle(bmp.GetHicon())
                    Catch ex As System.Exception
                        ' Keep the dialog usable even if the application icon
                        ' cannot be obtained.
                    End Try

                    Dim hasOptions As Boolean =
                        request.Options IsNot Nothing AndAlso
                        request.Options.Count > 0

                    Dim allowText As Boolean =
                        request.AllowFreeText OrElse Not hasOptions

                    Dim isNumeric As Boolean =
                        request.InputType = "integer" OrElse
                        request.InputType = "number"

                    ' =========================================================
                    ' Root layout
                    '
                    ' Row 0: question + options (independently scrollable)
                    ' Row 1: input caption
                    ' Row 2: input control
                    ' Row 3: fixed footer
                    '
                    ' The root itself never scrolls. Therefore the footer cannot
                    ' be pushed out of the client area by a large prompt.
                    ' =========================================================
                    Dim root As New System.Windows.Forms.TableLayoutPanel() With {
                        .Dock = System.Windows.Forms.DockStyle.Fill,
                        .ColumnCount = 1,
                        .RowCount = 4,
                        .AutoSize = False,
                        .AutoScroll = False,
                        .Padding = New System.Windows.Forms.Padding(16),
                        .Margin = New System.Windows.Forms.Padding(0)
                    }

                    root.ColumnStyles.Add(
                        New System.Windows.Forms.ColumnStyle(
                            System.Windows.Forms.SizeType.Percent,
                            100.0F
                        )
                    )

                    ' The prompt receives all remaining vertical space. The free-text
                    ' editor deliberately does NOT receive a percentage row: otherwise a
                    ' large form makes the editor huge and starves the actual question /
                    ' option area.
                    root.RowStyles.Add(
                        New System.Windows.Forms.RowStyle(
                            System.Windows.Forms.SizeType.Percent,
                            100.0F
                        )
                    )

                    root.RowStyles.Add(
                        New System.Windows.Forms.RowStyle(
                            System.Windows.Forms.SizeType.AutoSize
                        )
                    )

                    root.RowStyles.Add(
                        New System.Windows.Forms.RowStyle(
                            System.Windows.Forms.SizeType.AutoSize
                        )
                    )

                    root.RowStyles.Add(
                        New System.Windows.Forms.RowStyle(
                            System.Windows.Forms.SizeType.AutoSize
                        )
                    )

                    inputForm.Controls.Add(root)

                    ' =========================================================
                    ' Scrollable prompt/options host
                    ' =========================================================
                    Dim promptHost As New System.Windows.Forms.Panel() With {
                        .Dock = System.Windows.Forms.DockStyle.Fill,
                        .AutoScroll = True,
                        .Margin = New System.Windows.Forms.Padding(0, 0, 0, 12),
                        .Padding = New System.Windows.Forms.Padding(0),
                        .BackColor = System.Drawing.SystemColors.Control
                    }

                    root.Controls.Add(promptHost, 0, 0)

                    Dim promptLayout As New System.Windows.Forms.TableLayoutPanel() With {
                        .Dock = System.Windows.Forms.DockStyle.Top,
                        .ColumnCount = 1,
                        .RowCount = 1,
                        .AutoSize = True,
                        .AutoSizeMode =
                            System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                        .Margin = New System.Windows.Forms.Padding(0),
                        .Padding = New System.Windows.Forms.Padding(0)
                    }

                    promptLayout.ColumnStyles.Add(
                        New System.Windows.Forms.ColumnStyle(
                            System.Windows.Forms.SizeType.Percent,
                            100.0F
                        )
                    )

                    promptLayout.RowStyles.Add(
                        New System.Windows.Forms.RowStyle(
                            System.Windows.Forms.SizeType.AutoSize
                        )
                    )

                    promptHost.Controls.Add(promptLayout)

                    ' 1) Question (Markdown).
                    '
                    ' The browser itself does not scroll. It grows only to the actual
                    ' rendered Markdown wrapper height; the surrounding promptHost scrolls
                    ' the complete question + options as one unit. Keeping viewport height
                    ' out of the measurement prevents artificial blank space between the
                    ' question and the first option button.
                    Dim questionView As New System.Windows.Forms.WebBrowser() With {
                        .IsWebBrowserContextMenuEnabled = False,
                        .AllowWebBrowserDrop = False,
                        .ScriptErrorsSuppressed = True,
                        .WebBrowserShortcutsEnabled = False,
                        .ScrollBarsEnabled = False,
                        .TabStop = False,
                        .Dock = System.Windows.Forms.DockStyle.Top,
                        .Height = 44,
                        .Margin = New System.Windows.Forms.Padding(0, 0, 0, 6)
                    }

                    AddHandler questionView.NewWindow,
                        Sub(
                            sender As Object,
                            e As System.ComponentModel.CancelEventArgs
                        )
                            e.Cancel = True
                        End Sub

                    ' WebBrowser/IE can finish DocumentCompleted before the final
                    ' TableLayoutPanel width is known. When the outer vertical
                    ' scrollbar later appears, the browser becomes slightly narrower,
                    ' so the Markdown can wrap onto additional lines. Measure again
                    ' after every relevant width change and once more asynchronously.
                    '
                    ' IMPORTANT: do NOT use Body.ScrollRectangle.Height or the <html>
                    ' ScrollRectangle here. In the legacy WinForms WebBrowser those
                    ' values are frequently at least as tall as the browser viewport.
                    ' Adding padding to that value on every re-measure creates a
                    ' positive feedback loop: questionView gets taller, the reported
                    ' document height gets taller, and the option buttons are pushed
                    ' farther and farther down. Measure the dedicated content wrapper
                    ' instead; its OffsetRectangle is independent of viewport height.
                    Dim questionMeasurePending As Boolean = False
                    Dim lastQuestionWidth As Integer = -1

                    Dim measureQuestionHeight As System.Action =
                        Sub()
                            Try
                                If questionView.IsDisposed OrElse
                                   questionView.Document Is Nothing Then

                                    Return
                                End If

                                Dim contentElement As System.Windows.Forms.HtmlElement =
                                    questionView.Document.GetElementById(
                                        "ask-user-question-content"
                                    )

                                If contentElement Is Nothing Then
                                    Return
                                End If

                                Dim contentRectangle As System.Drawing.Rectangle =
                                    contentElement.OffsetRectangle

                                Dim contentHeight As Integer =
                                    contentRectangle.Height

                                ' OffsetRectangle.Height is the rendered wrapper height
                                ' and therefore does not inherit the WebBrowser viewport
                                ' height. A small DPI-scaled allowance protects the last
                                ' line against IE/WinForms rounding without accumulating
                                ' on subsequent measurements.
                                Dim newHeight As Integer =
                                    System.Math.Max(
                                        AskUserScale(inputForm, 34),
                                        contentHeight +
                                        AskUserScale(inputForm, 6)
                                    )

                                lastQuestionWidth =
                                    questionView.ClientSize.Width

                                If System.Math.Abs(
                                    questionView.Height - newHeight
                                ) > 1 Then

                                    questionView.Height = newHeight
                                    promptLayout.PerformLayout()
                                End If

                                promptHost.AutoScrollPosition =
                                    New System.Drawing.Point(0, 0)
                            Catch ex As System.Exception
                                ' Keep the current height if IE/DOM measurement is
                                ' temporarily unavailable.
                            End Try
                        End Sub

                    Dim queueQuestionHeightMeasure As System.Action =
                        Sub()
                            If questionMeasurePending OrElse
                               questionView.IsDisposed OrElse
                               Not questionView.IsHandleCreated Then

                                Return
                            End If

                            questionMeasurePending = True

                            Try
                                questionView.BeginInvoke(
                                    New System.Windows.Forms.MethodInvoker(
                                        Sub()
                                            questionMeasurePending = False
                                            measureQuestionHeight()
                                        End Sub
                                    )
                                )
                            Catch ex As System.Exception
                                questionMeasurePending = False
                            End Try
                        End Sub

                    AddHandler questionView.DocumentCompleted,
                        Sub(
                            sender As Object,
                            e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs
                        )
                            measureQuestionHeight()
                            queueQuestionHeightMeasure()
                        End Sub

                    AddHandler questionView.ClientSizeChanged,
                        Sub(
                            sender As Object,
                            e As System.EventArgs
                        )
                            If questionView.ClientSize.Width <> lastQuestionWidth Then
                                queueQuestionHeightMeasure()
                            End If
                        End Sub

                    AddHandler promptHost.ClientSizeChanged,
                        Sub(
                            sender As Object,
                            e As System.EventArgs
                        )
                            queueQuestionHeightMeasure()
                        End Sub

                    questionView.DocumentText =
                        BuildAskUserQuestionHtml(
                            request.Question,
                            System.Drawing.SystemColors.Control
                        )

                    promptLayout.Controls.Add(questionView, 0, 0)

                    ' =========================================================
                    ' 2) Predefined answer buttons
                    ' =========================================================
                    Dim orderedOptions As New System.Collections.Generic.List(
                        Of Global.SharedLibrary.Agents.AskUserOption
                    )()

                    Dim optionButtons As New System.Collections.Generic.List(
                        Of AskUserMarkdownButton
                    )()

                    Dim selectedFlags As New System.Collections.Generic.List(
                        Of Boolean
                    )()

                    Dim optionsTable As System.Windows.Forms.TableLayoutPanel =
                        Nothing

                    If hasOptions Then
                        For Each opt As Global.SharedLibrary.Agents.AskUserOption In request.Options

                            If opt Is Nothing Then
                                Continue For
                            End If

                            orderedOptions.Add(opt)
                        Next

                        optionsTable =
                            New System.Windows.Forms.TableLayoutPanel() With {
                                .ColumnCount = 1,
                                .RowCount = 0,
                                .AutoSize = True,
                                .AutoSizeMode =
                                    System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                                .Dock = System.Windows.Forms.DockStyle.Top,
                                .Margin = New System.Windows.Forms.Padding(0),
                                .Padding = New System.Windows.Forms.Padding(0)
                            }

                        optionsTable.ColumnStyles.Add(
                            New System.Windows.Forms.ColumnStyle(
                                System.Windows.Forms.SizeType.Percent,
                                100.0F
                            )
                        )

                        For i As Integer = 0 To orderedOptions.Count - 1
                            Dim optionItem As Global.SharedLibrary.Agents.AskUserOption =
                                orderedOptions(i)

                            Dim caption As String =
                                If(
                                    System.String.IsNullOrWhiteSpace(
                                        optionItem.Description
                                    ),
                                    If(optionItem.Label, ""),
                                    If(optionItem.Label, "") &
                                        System.Environment.NewLine &
                                        optionItem.Description
                                )

                            Dim optionButton As New AskUserMarkdownButton() With {
                                .Text = caption,
                                .AccessibleName = AskUserStripOptionMarkdown(caption),
                                .Tag = i,
                                .Dock = System.Windows.Forms.DockStyle.Top,
                                .AutoSize = False,
                                .MinimumSize =
                                    New System.Drawing.Size(0, 44),
                                .TextAlign =
                                    System.Drawing.ContentAlignment.MiddleLeft,
                                .Padding =
                                    New System.Windows.Forms.Padding(14, 8, 14, 8),
                                .Margin =
                                    New System.Windows.Forms.Padding(0, 0, 0, 8),
                                .UseVisualStyleBackColor = True,
                                .FlatStyle =
                                    System.Windows.Forms.FlatStyle.Standard,
                                .Font = standardFont,
                                .AutoEllipsis = False
                            }

                            selectedFlags.Add(False)
                            optionButtons.Add(optionButton)

                            optionsTable.RowStyles.Add(
                                New System.Windows.Forms.RowStyle(
                                    System.Windows.Forms.SizeType.AutoSize
                                )
                            )

                            optionsTable.RowCount = i + 1
                            optionsTable.Controls.Add(optionButton, 0, i)
                        Next

                        For i As Integer = 0 To optionButtons.Count - 1
                            Dim idx As Integer = i

                            AddHandler optionButtons(i).Click,
                                Sub(
                                    sender As Object,
                                    e As System.EventArgs
                                )
                                    If request.MultiSelect Then
                                        selectedFlags(idx) =
                                            Not selectedFlags(idx)

                                        optionButtons(idx).IsSelected =
                                            selectedFlags(idx)
                                    Else
                                        answer.Status = "answered"
                                        answer.SelectedOptionIds =
                                            New System.Collections.Generic.List(
                                                Of String
                                            ) From {
                                                orderedOptions(idx).Id
                                            }
                                        answer.FreeText = Nothing

                                        inputForm.DialogResult =
                                            System.Windows.Forms.DialogResult.OK
                                        inputForm.Close()
                                    End If
                                End Sub
                        Next

                        promptLayout.RowStyles.Add(
                            New System.Windows.Forms.RowStyle(
                                System.Windows.Forms.SizeType.AutoSize
                            )
                        )
                        promptLayout.RowCount = 2
                        promptLayout.Controls.Add(optionsTable, 0, 1)
                    End If

                    ' Re-measure option button heights against their actual current
                    ' width. This keeps long labels/descriptions readable at every
                    ' form width and DPI scale.
                    Dim updateOptionButtonHeights As System.Action =
                        Sub()
                            If optionsTable Is Nothing OrElse
                               optionButtons.Count = 0 Then

                                Return
                            End If

                            Dim availableWidth As Integer =
                                optionsTable.ClientSize.Width

                            If availableWidth <= 0 Then
                                Return
                            End If

                            For Each optionButton As AskUserMarkdownButton In optionButtons

                                Dim textWidth As Integer =
                                    availableWidth -
                                    optionButton.Margin.Horizontal -
                                    optionButton.Padding.Horizontal -
                                    AskUserScale(inputForm, 12)

                                textWidth =
                                    System.Math.Max(
                                        AskUserScale(inputForm, 80),
                                        textWidth
                                    )

                                Dim markdownTextHeight As Integer =
                                    optionButton.GetMarkdownPreferredTextHeight(
                                        textWidth
                                    )

                                optionButton.Height =
                                    System.Math.Max(
                                        AskUserScale(inputForm, 44),
                                        markdownTextHeight +
                                        optionButton.Padding.Vertical +
                                        AskUserScale(inputForm, 8)
                                    )
                            Next
                        End Sub

                    If optionsTable IsNot Nothing Then
                        AddHandler optionsTable.ClientSizeChanged,
                            Sub(
                                sender As Object,
                                e As System.EventArgs
                            )
                                updateOptionButtonHeights()
                            End Sub
                    End If

                    ' =========================================================
                    ' 3) Input caption + 4) input field
                    ' =========================================================
                    Dim txtInput As System.Windows.Forms.TextBox = Nothing
                    Dim lblInput As System.Windows.Forms.Label = Nothing

                    If allowText Then
                        Dim caption As String

                        Select Case request.InputType
                            Case "integer"
                                caption = "Enter a whole number:"
                            Case "number"
                                caption = "Enter a number:"
                            Case Else
                                caption =
                                    If(
                                        hasOptions,
                                        "Or type your own answer:",
                                        "Your answer:"
                                    )
                        End Select

                        lblInput = New System.Windows.Forms.Label() With {
                            .Text = caption,
                            .Font = standardFont,
                            .AutoSize = True,
                            .Margin =
                                New System.Windows.Forms.Padding(0, 0, 0, 4)
                        }

                        root.Controls.Add(lblInput, 0, 1)

                        txtInput = New System.Windows.Forms.TextBox() With {
                            .Font = standardFont,
                            .Multiline = Not isNumeric,
                            .WordWrap = True,
                            .AcceptsReturn = Not isNumeric,
                            .ScrollBars =
                                If(
                                    isNumeric,
                                    System.Windows.Forms.ScrollBars.None,
                                    System.Windows.Forms.ScrollBars.Vertical
                                ),
                            .Margin =
                                New System.Windows.Forms.Padding(0, 0, 0, 4)
                        }

                        If isNumeric Then
                            txtInput.Dock = System.Windows.Forms.DockStyle.Top
                            txtInput.Height =
                                System.Windows.Forms.TextRenderer.MeasureText(
                                    "Wy",
                                    standardFont
                                ).Height +
                                AskUserScale(inputForm, 8)
                        Else
                            ' A free-text answer needs room for several lines, but it
                            ' should never consume most of the dialog. Keep it at a
                            ' compact DPI-scaled height; when the user enlarges the
                            ' window, the extra room goes to question/options instead.
                            Dim freeTextHeight As Integer =
                                AskUserScale(inputForm, 112)

                            txtInput.Dock = System.Windows.Forms.DockStyle.Top
                            txtInput.Height = freeTextHeight
                            txtInput.MinimumSize =
                                New System.Drawing.Size(0, freeTextHeight)
                            txtInput.MaximumSize =
                                New System.Drawing.Size(0, freeTextHeight)
                        End If

                        root.Controls.Add(txtInput, 0, 2)
                    Else
                        root.RowStyles(0) =
                            New System.Windows.Forms.RowStyle(
                                System.Windows.Forms.SizeType.Percent,
                                100.0F
                            )

                        root.RowStyles(1) =
                            New System.Windows.Forms.RowStyle(
                                System.Windows.Forms.SizeType.Absolute,
                                0.0F
                            )

                        root.RowStyles(2) =
                            New System.Windows.Forms.RowStyle(
                                System.Windows.Forms.SizeType.Absolute,
                                0.0F
                            )
                    End If

                    ' =========================================================
                    ' 5) Fixed OK / Cancel footer
                    ' =========================================================
                    Dim okButton As New System.Windows.Forms.Button() With {
                        .Text = "OK",
                        .AutoSize = True,
                        .MinimumSize = New System.Drawing.Size(88, 34),
                        .Font = standardFont,
                        .Margin =
                            New System.Windows.Forms.Padding(8, 3, 0, 3),
                        .UseVisualStyleBackColor = True
                    }

                    Dim cancelButton As New System.Windows.Forms.Button() With {
                        .Text = "Cancel",
                        .AutoSize = True,
                        .MinimumSize = New System.Drawing.Size(88, 34),
                        .Font = standardFont,
                        .Margin =
                            New System.Windows.Forms.Padding(0, 3, 0, 3),
                        .DialogResult =
                            System.Windows.Forms.DialogResult.Cancel,
                        .UseVisualStyleBackColor = True
                    }

                    okButton.Visible =
                        allowText OrElse request.MultiSelect

                    Dim bottomFlow As New System.Windows.Forms.FlowLayoutPanel() With {
                        .FlowDirection =
                            System.Windows.Forms.FlowDirection.RightToLeft,
                        .AutoSize = True,
                        .AutoSizeMode =
                            System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                        .WrapContents = False,
                        .Dock = System.Windows.Forms.DockStyle.Fill,
                        .Margin =
                            New System.Windows.Forms.Padding(0, 10, 0, 0),
                        .Padding = New System.Windows.Forms.Padding(0)
                    }

                    If okButton.Visible Then
                        bottomFlow.Controls.Add(okButton)
                    End If

                    bottomFlow.Controls.Add(cancelButton)
                    root.Controls.Add(bottomFlow, 0, 3)

                    inputForm.CancelButton = cancelButton

                    If okButton.Visible Then
                        inputForm.AcceptButton = okButton
                    End If

                    ' =========================================================
                    ' Commit / validation
                    ' =========================================================
                    Dim commitOk As System.Action =
                        Sub()
                            Dim typed As String =
                                If(
                                    txtInput IsNot Nothing,
                                    txtInput.Text.Trim(),
                                    ""
                                )

                            Dim selectedIds As New System.Collections.Generic.List(
                                Of String
                            )()

                            If request.MultiSelect Then
                                For i As Integer = 0 To orderedOptions.Count - 1
                                    If i < selectedFlags.Count AndAlso
                                       selectedFlags(i) Then

                                        selectedIds.Add(
                                            orderedOptions(i).Id
                                        )
                                    End If
                                Next
                            End If

                            Dim hasSelection As Boolean =
                                selectedIds.Count > 0

                            Dim hasTextAnswer As Boolean =
                                typed.Length > 0

                            If Not hasSelection AndAlso
                               Not hasTextAnswer Then

                                ShowCustomMessageBox(
                                    "Please choose an option or enter an answer."
                                )
                                Return
                            End If

                            If hasTextAnswer AndAlso
                               Not hasSelection Then

                                If request.InputType = "integer" Then
                                    Dim integerValue As Long

                                    If Not Long.TryParse(
                                        typed,
                                        System.Globalization.NumberStyles.Integer,
                                        System.Globalization.CultureInfo.InvariantCulture,
                                        integerValue
                                    ) Then
                                        ShowCustomMessageBox(
                                            "Please enter a valid whole number."
                                        )
                                        Return
                                    End If

                                ElseIf request.InputType = "number" Then
                                    Dim doubleValue As Double

                                    If Not Double.TryParse(
                                        typed,
                                        System.Globalization.NumberStyles.Float Or
                                        System.Globalization.NumberStyles.AllowThousands,
                                        System.Globalization.CultureInfo.InvariantCulture,
                                        doubleValue
                                    ) AndAlso
                                       Not Double.TryParse(
                                        typed,
                                        System.Globalization.NumberStyles.Float Or
                                        System.Globalization.NumberStyles.AllowThousands,
                                        System.Globalization.CultureInfo.CurrentCulture,
                                        doubleValue
                                    ) Then
                                        ShowCustomMessageBox(
                                            "Please enter a valid number."
                                        )
                                        Return
                                    End If
                                End If
                            End If

                            answer.Status = "answered"
                            answer.SelectedOptionIds = selectedIds
                            answer.FreeText =
                                If(hasTextAnswer, typed, Nothing)

                            inputForm.DialogResult =
                                System.Windows.Forms.DialogResult.OK
                            inputForm.Close()
                        End Sub

                    AddHandler okButton.Click,
                        Sub(
                            sender As Object,
                            e As System.EventArgs
                        )
                            commitOk()
                        End Sub

                    AddHandler cancelButton.Click,
                        Sub(
                            sender As Object,
                            e As System.EventArgs
                        )
                            inputForm.DialogResult =
                                System.Windows.Forms.DialogResult.Cancel
                            inputForm.Close()
                        End Sub

                    If txtInput IsNot Nothing Then
                        AddHandler txtInput.KeyDown,
                            Sub(
                                sender As Object,
                                e As System.Windows.Forms.KeyEventArgs
                            )
                                If isNumeric AndAlso
                                   e.KeyCode =
                                   System.Windows.Forms.Keys.Enter Then

                                    e.SuppressKeyPress = True
                                    commitOk()

                                ElseIf Not isNumeric AndAlso
                                       e.KeyCode =
                                       System.Windows.Forms.Keys.Enter AndAlso
                                       e.Modifiers =
                                       System.Windows.Forms.Keys.Control Then

                                    e.SuppressKeyPress = True
                                    commitOk()

                                ElseIf e.KeyCode =
                                       System.Windows.Forms.Keys.Escape Then

                                    e.SuppressKeyPress = True

                                    inputForm.DialogResult =
                                        System.Windows.Forms.DialogResult.Cancel
                                    inputForm.Close()
                                End If
                            End Sub
                    End If

                    ' =========================================================
                    ' Shown: final DPI-aware sizing, positioning and foreground
                    ' protection.
                    ' =========================================================
                    AddHandler inputForm.Shown,
                        Sub(
                            sender As Object,
                            e As System.EventArgs
                        )
                            inputForm.PerformLayout()
                            updateOptionButtonHeights()
                            queueQuestionHeightMeasure()

                            Dim edgeMargin As Integer =
                                AskUserScale(inputForm, 20)

                            Dim maxWidth As Integer =
                                System.Math.Max(
                                    AskUserScale(inputForm, 320),
                                    wa.Width - (edgeMargin * 2)
                                )

                            Dim maxHeight As Integer =
                                System.Math.Max(
                                    AskUserScale(inputForm, 260),
                                    wa.Height - (edgeMargin * 2)
                                )

                            Dim desiredWidth As Integer =
                                System.Math.Min(
                                    inputForm.Width,
                                    maxWidth
                                )

                            Dim desiredHeight As Integer =
                                System.Math.Min(
                                    inputForm.Height,
                                    maxHeight
                                )

                            inputForm.Size =
                                New System.Drawing.Size(
                                    desiredWidth,
                                    desiredHeight
                                )

                            ' Unlike the old implementation, do NOT lock the minimum
                            ' size to the fully measured content height. A compact,
                            ' DPI-scaled minimum keeps resizing useful; long content
                            ' simply scrolls in promptHost.
                            Dim minimumWidth As Integer =
                                System.Math.Min(
                                    AskUserScale(inputForm, 520),
                                    maxWidth
                                )

                            Dim minimumHeight As Integer =
                                System.Math.Min(
                                    AskUserScale(inputForm, 360),
                                    maxHeight
                                )

                            inputForm.MinimumSize =
                                New System.Drawing.Size(
                                    minimumWidth,
                                    minimumHeight
                                )

                            inputForm.Location =
                                New System.Drawing.Point(
                                    wa.X +
                                    (wa.Width - inputForm.Width) \ 2,
                                    wa.Y +
                                    (wa.Height - inputForm.Height) \ 2
                                )

                            inputForm.PerformLayout()
                            updateOptionButtonHeights()
                            queueQuestionHeightMeasure()

                            ' Never let layout/focus side effects leave the prompt
                            ' scrolled down when the dialog first becomes visible.
                            promptHost.AutoScrollPosition =
                                New System.Drawing.Point(0, 0)

                            ForceDialogToForeground(inputForm)
                            SharedMethods.AttachForeignForegroundWatchdog(
                                inputForm
                            )

                            inputForm.Opacity = 1

                            If txtInput IsNot Nothing Then
                                Try
                                    txtInput.Focus()
                                Catch ex As System.Exception
                                End Try
                            ElseIf optionButtons.Count > 0 Then
                                Try
                                    optionButtons(0).Focus()
                                Catch ex As System.Exception
                                End Try
                            End If

                            ' Focus can trigger ScrollControlIntoView on nested layout
                            ' containers. Reset once more after the current UI message
                            ' has completed so the user always sees the question from
                            ' its first line.
                            Try
                                inputForm.BeginInvoke(
                                    New System.Windows.Forms.MethodInvoker(
                                        Sub()
                                            If Not promptHost.IsDisposed Then
                                                promptHost.AutoScrollPosition =
                                                    New System.Drawing.Point(0, 0)
                                            End If
                                        End Sub
                                    )
                                )
                            Catch ex As System.Exception
                            End Try
                        End Sub

                    Dim dialogResult As System.Windows.Forms.DialogResult

                    ' Revalidate immediately before ShowDialog because the captured
                    ' foreground/owner window may have disappeared in the meantime.
                    Dim ownerWindow As System.Windows.Forms.IWin32Window =
                        Nothing

                    Try
                        If usableOwnerHandle <> System.IntPtr.Zero AndAlso
                           AskUserIsWindow(usableOwnerHandle) Then

                            ownerWindow =
                                New AskUserWindowOwner(
                                    usableOwnerHandle
                                )
                        End If
                    Catch ex As System.Exception
                        ownerWindow = Nothing
                    End Try

                    If ownerWindow IsNot Nothing Then
                        dialogResult = inputForm.ShowDialog(ownerWindow)
                    Else
                        dialogResult = inputForm.ShowDialog()
                    End If

                    If dialogResult <>
                       System.Windows.Forms.DialogResult.OK Then

                        Return New Global.SharedLibrary.Agents.AskUserResult() With {
                            .Status = "cancelled"
                        }
                    End If

                    Return answer
                End Using
            End Using
        End Function

        ''' <summary>
        ''' Renders the ask_user question as a small Markdown-enabled HTML document
        ''' whose background matches the given control color so it blends with the form.
        ''' </summary>
        Private Shared Function BuildAskUserQuestionHtml(
            question As String,
            back As System.Drawing.Color
        ) As String

            Dim bodyHtml As String

            Try
                Dim pipeline As Markdig.MarkdownPipeline =
                    New Markdig.MarkdownPipelineBuilder().
                        UseAdvancedExtensions().
                        UseSoftlineBreakAsHardlineBreak().
                        Build()

                bodyHtml =
                    Markdig.Markdown.ToHtml(
                        If(question, ""),
                        pipeline
                    )
            Catch ex As System.Exception
                bodyHtml =
                    "<p>" &
                    System.Net.WebUtility.HtmlEncode(
                        If(question, "")
                    ) &
                    "</p>"
            End Try

            Dim background As String =
                System.String.Format(
                    "#{0:X2}{1:X2}{2:X2}",
                    back.R,
                    back.G,
                    back.B
                )

            Return "<!DOCTYPE html><html><head><meta charset=""utf-8"">" &
                   "<meta http-equiv=""X-UA-Compatible"" content=""IE=edge"">" &
                   "<style>" &
                   "html,body{margin:0;padding:0;background:" &
                   background &
                   ";overflow:hidden;max-width:100%;height:auto;min-height:0;}" &
                   "body{font-family:'Segoe UI',sans-serif;font-size:11pt;" &
                   "color:#1b1b1b;line-height:1.35;overflow-wrap:anywhere;" &
                   "word-wrap:break-word;}" &
                   "#ask-user-question-content{margin:0;padding:0;" &
                   "overflow:hidden;width:100%;box-sizing:border-box;}" &
                   "p{margin:0 0 6px 0;}" &
                   "ul,ol{margin:0 0 6px 20px;padding:0;}" &
                   "code{background:#e6e6e6;padding:1px 4px;border-radius:3px;}" &
                   "pre{white-space:pre-wrap;overflow-wrap:anywhere;}" &
                   "h1,h2,h3{font-size:12pt;margin:0 0 6px 0;}" &
                   "strong{font-weight:600;}" &
                   "</style></head><body>" &
                   "<div id=""ask-user-question-content"">" &
                   bodyHtml &
                   "</div></body></html>"
        End Function

        ' =====================================================================
        ' Markdown-capable option buttons
        '
        ' Supported inline formatting in option labels/descriptions:
        '   **bold** / __bold__
        '   *italic* / _italic_
        '   ***bold italic*** / ___bold italic___
        '   ++underline++
        '   <u>underline</u> / <ins>underline</ins>
        '
        ' Underline is not part of core Markdown, hence ++...++ and the HTML
        ' underline forms are accepted explicitly for ask_user option captions.
        ' =====================================================================
        Private NotInheritable Class AskUserMarkdownRun

            Public Sub New(
                text As String,
                style As System.Drawing.FontStyle
            )
                Me.Text = text
                Me.Style = style
            End Sub

            Public Property Text As String
            Public Property Style As System.Drawing.FontStyle
        End Class

        Private NotInheritable Class AskUserMarkdownLayoutItem

            Public Sub New(
                text As String,
                style As System.Drawing.FontStyle,
                x As Integer,
                y As Integer
            )
                Me.Text = text
                Me.Style = style
                Me.X = x
                Me.Y = y
            End Sub

            Public ReadOnly Property Text As String
            Public ReadOnly Property Style As System.Drawing.FontStyle
            Public ReadOnly Property X As Integer
            Public ReadOnly Property Y As Integer
        End Class

        Private NotInheritable Class AskUserMarkdownLayout

            Public Sub New()
                Me.Items =
                    New System.Collections.Generic.List(
                        Of AskUserMarkdownLayoutItem
                    )()
            End Sub

            Public ReadOnly Property Items As System.Collections.Generic.List(
                Of AskUserMarkdownLayoutItem
            )

            Public Property Height As Integer
        End Class

        Private NotInheritable Class AskUserMarkdownButton
            Inherits System.Windows.Forms.Button

            Private _isSelected As Boolean
            Private _mouseOver As Boolean
            Private _mouseDown As Boolean

            Public Sub New()
                MyBase.New()

                Me.SetStyle(
                    System.Windows.Forms.ControlStyles.UserPaint Or
                    System.Windows.Forms.ControlStyles.AllPaintingInWmPaint Or
                    System.Windows.Forms.ControlStyles.OptimizedDoubleBuffer Or
                    System.Windows.Forms.ControlStyles.ResizeRedraw,
                    True
                )
            End Sub

            Public Property IsSelected As Boolean
                Get
                    Return _isSelected
                End Get
                Set(value As Boolean)
                    If _isSelected = value Then
                        Return
                    End If

                    _isSelected = value
                    Me.Invalidate()
                End Set
            End Property

            Public Function GetMarkdownPreferredTextHeight(
                textWidth As Integer
            ) As Integer

                If textWidth <= 0 Then
                    Return 0
                End If

                Try
                    Using graphics As System.Drawing.Graphics =
                        Me.CreateGraphics()

                        Dim layout As AskUserMarkdownLayout =
                            CreateMarkdownLayout(
                                graphics,
                                textWidth
                            )

                        Return layout.Height
                    End Using
                Catch ex As System.Exception
                    Dim measured As System.Drawing.Size =
                        System.Windows.Forms.TextRenderer.MeasureText(
                            AskUserStripOptionMarkdown(Me.Text),
                            Me.Font,
                            New System.Drawing.Size(
                                textWidth,
                                System.Int32.MaxValue
                            ),
                            System.Windows.Forms.TextFormatFlags.WordBreak Or
                            System.Windows.Forms.TextFormatFlags.TextBoxControl Or
                            System.Windows.Forms.TextFormatFlags.NoPrefix
                        )

                    Return measured.Height
                End Try
            End Function

            Protected Overrides Sub OnMouseEnter(e As System.EventArgs)
                _mouseOver = True
                Me.Invalidate()
                MyBase.OnMouseEnter(e)
            End Sub

            Protected Overrides Sub OnMouseLeave(e As System.EventArgs)
                _mouseOver = False
                _mouseDown = False
                Me.Invalidate()
                MyBase.OnMouseLeave(e)
            End Sub

            Protected Overrides Sub OnMouseDown(
                e As System.Windows.Forms.MouseEventArgs
            )
                If e.Button = System.Windows.Forms.MouseButtons.Left Then
                    _mouseDown = True
                    Me.Invalidate()
                End If

                MyBase.OnMouseDown(e)
            End Sub

            Protected Overrides Sub OnMouseUp(
                e As System.Windows.Forms.MouseEventArgs
            )
                _mouseDown = False
                Me.Invalidate()
                MyBase.OnMouseUp(e)
            End Sub

            Protected Overrides Sub OnGotFocus(e As System.EventArgs)
                MyBase.OnGotFocus(e)
                Me.Invalidate()
            End Sub

            Protected Overrides Sub OnLostFocus(e As System.EventArgs)
                MyBase.OnLostFocus(e)
                _mouseDown = False
                Me.Invalidate()
            End Sub

            Protected Overrides Sub OnEnabledChanged(e As System.EventArgs)
                MyBase.OnEnabledChanged(e)
                Me.Invalidate()
            End Sub

            Protected Overrides Sub OnTextChanged(e As System.EventArgs)
                MyBase.OnTextChanged(e)
                Me.AccessibleName = AskUserStripOptionMarkdown(Me.Text)
                Me.Invalidate()
            End Sub

            Protected Overrides Sub OnPaint(
                e As System.Windows.Forms.PaintEventArgs
            )
                DrawMarkdownButtonBackground(e.Graphics)

                Dim borderInset As Integer =
                    System.Math.Max(
                        2,
                        CInt(
                            System.Math.Round(
                                Me.DeviceDpi / 96.0R * 2.0R,
                                System.MidpointRounding.AwayFromZero
                            )
                        )
                    )

                Dim textWidth As Integer =
                    System.Math.Max(
                        1,
                        Me.ClientSize.Width -
                        Me.Padding.Horizontal -
                        (borderInset * 2)
                    )

                Dim layout As AskUserMarkdownLayout =
                    CreateMarkdownLayout(
                        e.Graphics,
                        textWidth
                    )

                Dim availableHeight As Integer =
                    System.Math.Max(
                        0,
                        Me.ClientSize.Height -
                        Me.Padding.Vertical -
                        (borderInset * 2)
                    )

                Dim textTop As Integer =
                    Me.Padding.Top +
                    borderInset +
                    System.Math.Max(
                        0,
                        (availableHeight - layout.Height) \ 2
                    )

                Dim textLeft As Integer =
                    Me.Padding.Left + borderInset

                Dim textColor As System.Drawing.Color =
                    If(
                        Me.Enabled,
                        Me.ForeColor,
                        System.Drawing.SystemColors.GrayText
                    )

                Dim fontCache As New System.Collections.Generic.Dictionary(
                    Of System.Drawing.FontStyle,
                    System.Drawing.Font
                )()

                Try
                    For Each item As AskUserMarkdownLayoutItem In layout.Items
                        Dim drawFont As System.Drawing.Font =
                            GetMarkdownFont(
                                fontCache,
                                item.Style
                            )

                        System.Windows.Forms.TextRenderer.DrawText(
                            e.Graphics,
                            item.Text,
                            drawFont,
                            New System.Drawing.Point(
                                textLeft + item.X,
                                textTop + item.Y
                            ),
                            textColor,
                            System.Windows.Forms.TextFormatFlags.NoPadding Or
                            System.Windows.Forms.TextFormatFlags.NoPrefix Or
                            System.Windows.Forms.TextFormatFlags.SingleLine
                        )
                    Next
                Finally
                    For Each cachedFont As System.Drawing.Font In fontCache.Values
                        cachedFont.Dispose()
                    Next
                End Try

                If Me.Focused AndAlso Me.ShowFocusCues Then
                    Dim focusRectangle As System.Drawing.Rectangle =
                        Me.ClientRectangle

                    focusRectangle.Inflate(
                        -System.Math.Max(3, borderInset + 1),
                        -System.Math.Max(3, borderInset + 1)
                    )

                    System.Windows.Forms.ControlPaint.DrawFocusRectangle(
                        e.Graphics,
                        focusRectangle
                    )
                End If
            End Sub

            Private Sub DrawMarkdownButtonBackground(
                graphics As System.Drawing.Graphics
            )
                Dim bounds As System.Drawing.Rectangle = Me.ClientRectangle

                If bounds.Width <= 0 OrElse bounds.Height <= 0 Then
                    Return
                End If

                If _isSelected Then
                    Using selectedBrush As New System.Drawing.SolidBrush(
                        System.Drawing.Color.FromArgb(210, 229, 247)
                    )
                        graphics.FillRectangle(
                            selectedBrush,
                            bounds
                        )
                    End Using

                    System.Windows.Forms.ControlPaint.DrawBorder(
                        graphics,
                        bounds,
                        System.Drawing.SystemColors.Highlight,
                        System.Windows.Forms.ButtonBorderStyle.Solid
                    )

                    Return
                End If

                Dim state As System.Windows.Forms.VisualStyles.PushButtonState

                If Not Me.Enabled Then
                    state =
                        System.Windows.Forms.VisualStyles.PushButtonState.Disabled
                ElseIf _mouseDown Then
                    state =
                        System.Windows.Forms.VisualStyles.PushButtonState.Pressed
                ElseIf _mouseOver Then
                    state =
                        System.Windows.Forms.VisualStyles.PushButtonState.Hot
                ElseIf Me.Focused Then
                    state =
                        System.Windows.Forms.VisualStyles.PushButtonState.Default
                Else
                    state =
                        System.Windows.Forms.VisualStyles.PushButtonState.Normal
                End If

                ' ButtonRenderer has no IsSupported property. Unlike renderers such
                ' as TextBoxRenderer/TabRenderer, ButtonRenderer is documented to
                ' render with visual styles when available and fall back to the
                ' classic Windows appearance otherwise.
                System.Windows.Forms.ButtonRenderer.DrawButton(
                    graphics,
                    bounds,
                    state
                )
            End Sub

            Private Function CreateMarkdownLayout(
                graphics As System.Drawing.Graphics,
                maximumWidth As Integer
            ) As AskUserMarkdownLayout

                Dim result As New AskUserMarkdownLayout()

                If maximumWidth <= 0 Then
                    Return result
                End If

                Dim runs As System.Collections.Generic.List(
                    Of AskUserMarkdownRun
                ) = ParseAskUserOptionMarkdown(Me.Text)

                Dim fontCache As New System.Collections.Generic.Dictionary(
                    Of System.Drawing.FontStyle,
                    System.Drawing.Font
                )()

                Try
                    Dim regularFont As System.Drawing.Font =
                        GetMarkdownFont(
                            fontCache,
                            System.Drawing.FontStyle.Regular
                        )

                    Dim regularLineHeight As Integer =
                        GetMarkdownLineHeight(
                            graphics,
                            regularFont
                        )

                    Dim x As Integer = 0
                    Dim y As Integer = 0
                    Dim lineHeight As Integer = 0
                    Dim lineHasContent As Boolean = False

                    Dim nextLine As System.Action =
                        Sub()
                            y +=
                                System.Math.Max(
                                    regularLineHeight,
                                    lineHeight
                                )

                            x = 0
                            lineHeight = 0
                            lineHasContent = False
                        End Sub

                    For Each run As AskUserMarkdownRun In runs
                        Dim runFont As System.Drawing.Font =
                            GetMarkdownFont(
                                fontCache,
                                run.Style
                            )

                        Dim matches As System.Text.RegularExpressions.MatchCollection =
                            System.Text.RegularExpressions.Regex.Matches(
                                If(run.Text, ""),
                                "\n|[ \t]+|[^ \t\n]+"
                            )

                        For Each match As System.Text.RegularExpressions.Match In matches
                            Dim token As String = match.Value

                            If token = System.Environment.NewLine OrElse
                               token = vbLf Then

                                nextLine()
                                Continue For
                            End If

                            Dim isWhitespace As Boolean =
                                token.Trim().Length = 0

                            If isWhitespace AndAlso Not lineHasContent Then
                                Continue For
                            End If

                            Dim remaining As String = token

                            Do While remaining.Length > 0
                                Dim remainingWidth As Integer =
                                    maximumWidth - x

                                Dim remainingSize As System.Drawing.Size =
                                    MeasureMarkdownText(
                                        graphics,
                                        remaining,
                                        runFont
                                    )

                                If remainingSize.Width <= remainingWidth Then
                                    result.Items.Add(
                                        New AskUserMarkdownLayoutItem(
                                            remaining,
                                            run.Style,
                                            x,
                                            y
                                        )
                                    )

                                    x += remainingSize.Width
                                    lineHeight =
                                        System.Math.Max(
                                            lineHeight,
                                            remainingSize.Height
                                        )
                                    lineHasContent = True
                                    remaining = ""
                                    Continue Do
                                End If

                                If x > 0 Then
                                    nextLine()

                                    If isWhitespace Then
                                        remaining = ""
                                    End If

                                    Continue Do
                                End If

                                ' A single token is wider than the complete row.
                                ' Split it at the largest character prefix that fits.
                                Dim fitLength As Integer =
                                    FindMarkdownPrefixThatFits(
                                        graphics,
                                        remaining,
                                        runFont,
                                        maximumWidth
                                    )

                                If fitLength <= 0 Then
                                    fitLength = 1
                                End If

                                Dim piece As String =
                                    remaining.Substring(0, fitLength)

                                Dim pieceSize As System.Drawing.Size =
                                    MeasureMarkdownText(
                                        graphics,
                                        piece,
                                        runFont
                                    )

                                result.Items.Add(
                                    New AskUserMarkdownLayoutItem(
                                        piece,
                                        run.Style,
                                        x,
                                        y
                                    )
                                )

                                x += pieceSize.Width
                                lineHeight =
                                    System.Math.Max(
                                        lineHeight,
                                        pieceSize.Height
                                    )
                                lineHasContent = True

                                remaining =
                                    remaining.Substring(fitLength)

                                If remaining.Length > 0 Then
                                    nextLine()
                                End If
                            Loop
                        Next
                    Next

                    If lineHasContent OrElse result.Items.Count = 0 Then
                        result.Height =
                            y +
                            System.Math.Max(
                                regularLineHeight,
                                lineHeight
                            )
                    Else
                        result.Height = y
                    End If

                    Return result
                Finally
                    For Each cachedFont As System.Drawing.Font In fontCache.Values
                        cachedFont.Dispose()
                    Next
                End Try
            End Function

            Private Function GetMarkdownFont(
                cache As System.Collections.Generic.Dictionary(
                    Of System.Drawing.FontStyle,
                    System.Drawing.Font
                ),
                markdownStyle As System.Drawing.FontStyle
            ) As System.Drawing.Font

                Dim effectiveStyle As System.Drawing.FontStyle =
                    Me.Font.Style Or markdownStyle

                Dim result As System.Drawing.Font = Nothing

                If cache.TryGetValue(effectiveStyle, result) Then
                    Return result
                End If

                result =
                    New System.Drawing.Font(
                        Me.Font,
                        effectiveStyle
                    )

                cache.Add(effectiveStyle, result)
                Return result
            End Function

            Private Shared Function GetMarkdownLineHeight(
                graphics As System.Drawing.Graphics,
                font As System.Drawing.Font
            ) As Integer

                Return MeasureMarkdownText(
                    graphics,
                    "Wy",
                    font
                ).Height
            End Function

            Private Shared Function MeasureMarkdownText(
                graphics As System.Drawing.Graphics,
                text As String,
                font As System.Drawing.Font
            ) As System.Drawing.Size

                Return System.Windows.Forms.TextRenderer.MeasureText(
                    graphics,
                    If(text, ""),
                    font,
                    New System.Drawing.Size(
                        System.Int32.MaxValue,
                        System.Int32.MaxValue
                    ),
                    System.Windows.Forms.TextFormatFlags.NoPadding Or
                    System.Windows.Forms.TextFormatFlags.NoPrefix Or
                    System.Windows.Forms.TextFormatFlags.SingleLine
                )
            End Function

            Private Shared Function FindMarkdownPrefixThatFits(
                graphics As System.Drawing.Graphics,
                text As String,
                font As System.Drawing.Font,
                maximumWidth As Integer
            ) As Integer

                If System.String.IsNullOrEmpty(text) Then
                    Return 0
                End If

                Dim low As Integer = 1
                Dim high As Integer = text.Length
                Dim best As Integer = 0

                While low <= high
                    Dim middle As Integer =
                        low + ((high - low) \ 2)

                    Dim width As Integer =
                        MeasureMarkdownText(
                            graphics,
                            text.Substring(0, middle),
                            font
                        ).Width

                    If width <= maximumWidth Then
                        best = middle
                        low = middle + 1
                    Else
                        high = middle - 1
                    End If
                End While

                Return best
            End Function
        End Class

        Private Shared Function AskUserStripOptionMarkdown(
            markdown As String
        ) As String

            Dim runs As System.Collections.Generic.List(
                Of AskUserMarkdownRun
            ) = ParseAskUserOptionMarkdown(markdown)

            Dim builder As New System.Text.StringBuilder()

            For Each run As AskUserMarkdownRun In runs
                builder.Append(run.Text)
            Next

            Return builder.ToString()
        End Function

        Private Shared Function ParseAskUserOptionMarkdown(
            markdown As String
        ) As System.Collections.Generic.List(Of AskUserMarkdownRun)

            Dim result As New System.Collections.Generic.List(
                Of AskUserMarkdownRun
            )()

            Dim text As String =
                If(markdown, "").
                    Replace(vbCrLf, vbLf).
                    Replace(vbCr, vbLf)

            Dim style As System.Drawing.FontStyle =
                System.Drawing.FontStyle.Regular

            Dim buffer As New System.Text.StringBuilder()

            Dim flush As System.Action =
                Sub()
                    If buffer.Length = 0 Then
                        Return
                    End If

                    Dim value As String = buffer.ToString()
                    buffer.Clear()

                    If result.Count > 0 AndAlso
                       result(result.Count - 1).Style = style Then

                        result(result.Count - 1).Text &= value
                    Else
                        result.Add(
                            New AskUserMarkdownRun(
                                value,
                                style
                            )
                        )
                    End If
                End Sub

            Dim i As Integer = 0

            While i < text.Length
                ' Markdown-style escaping for the formatting delimiters supported
                ' by this lightweight button renderer.
                If text(i) = "\"c AndAlso i + 1 < text.Length Then
                    Dim escaped As Char = text(i + 1)

                    If escaped = "*"c OrElse
                       escaped = "_"c OrElse
                       escaped = "+"c OrElse
                       escaped = "\"c OrElse
                       escaped = "<"c OrElse
                       escaped = ">"c Then

                        buffer.Append(escaped)
                        i += 2
                        Continue While
                    End If
                End If

                If AskUserStartsWith(
                    text,
                    i,
                    "<u>"
                ) OrElse
                   AskUserStartsWith(
                    text,
                    i,
                    "<ins>"
                ) Then

                    flush()
                    style = style Or System.Drawing.FontStyle.Underline
                    i += If(AskUserStartsWith(text, i, "<u>"), 3, 5)
                    Continue While
                End If

                If AskUserStartsWith(
                    text,
                    i,
                    "</u>"
                ) OrElse
                   AskUserStartsWith(
                    text,
                    i,
                    "</ins>"
                ) Then

                    flush()
                    style = style And Not System.Drawing.FontStyle.Underline
                    i += If(AskUserStartsWith(text, i, "</u>"), 4, 6)
                    Continue While
                End If

                If AskUserTryToggleMarkdownDelimiter(
                    text,
                    i,
                    "***",
                    System.Drawing.FontStyle.Bold Or
                    System.Drawing.FontStyle.Italic,
                    style,
                    flush
                ) Then
                    i += 3
                    Continue While
                End If

                If AskUserTryToggleMarkdownDelimiter(
                    text,
                    i,
                    "___",
                    System.Drawing.FontStyle.Bold Or
                    System.Drawing.FontStyle.Italic,
                    style,
                    flush
                ) Then
                    i += 3
                    Continue While
                End If

                If AskUserTryToggleMarkdownDelimiter(
                    text,
                    i,
                    "**",
                    System.Drawing.FontStyle.Bold,
                    style,
                    flush
                ) Then
                    i += 2
                    Continue While
                End If

                If AskUserTryToggleMarkdownDelimiter(
                    text,
                    i,
                    "__",
                    System.Drawing.FontStyle.Bold,
                    style,
                    flush
                ) Then
                    i += 2
                    Continue While
                End If

                If AskUserTryToggleMarkdownDelimiter(
                    text,
                    i,
                    "++",
                    System.Drawing.FontStyle.Underline,
                    style,
                    flush
                ) Then
                    i += 2
                    Continue While
                End If

                If AskUserTryToggleMarkdownDelimiter(
                    text,
                    i,
                    "*",
                    System.Drawing.FontStyle.Italic,
                    style,
                    flush
                ) Then
                    i += 1
                    Continue While
                End If

                If AskUserTryToggleMarkdownDelimiter(
                    text,
                    i,
                    "_",
                    System.Drawing.FontStyle.Italic,
                    style,
                    flush
                ) Then
                    i += 1
                    Continue While
                End If

                buffer.Append(text(i))
                i += 1
            End While

            flush()
            Return result
        End Function

        Private Shared Function AskUserTryToggleMarkdownDelimiter(
            text As String,
            index As Integer,
            delimiter As String,
            flags As System.Drawing.FontStyle,
            ByRef style As System.Drawing.FontStyle,
            flush As System.Action
        ) As Boolean

            If Not AskUserStartsWith(
                text,
                index,
                delimiter
            ) Then
                Return False
            End If

            Dim currentlyActive As Boolean =
                (style And flags) = flags

            If Not currentlyActive Then
                Dim closingIndex As Integer =
                    text.IndexOf(
                        delimiter,
                        index + delimiter.Length,
                        System.StringComparison.Ordinal
                    )

                If closingIndex < 0 Then
                    Return False
                End If
            End If

            flush()

            If currentlyActive Then
                style = style And Not flags
            Else
                style = style Or flags
            End If

            Return True
        End Function

        Private Shared Function AskUserStartsWith(
            text As String,
            index As Integer,
            value As String
        ) As Boolean

            If text Is Nothing OrElse
               value Is Nothing OrElse
               index < 0 OrElse
               index + value.Length > text.Length Then

                Return False
            End If

            Return System.String.Compare(
                text,
                index,
                value,
                0,
                value.Length,
                System.StringComparison.OrdinalIgnoreCase
            ) = 0
        End Function

        ' =====================================================================
        ' DPI helpers
        ' =====================================================================
        Private Shared Function AskUserScale(
            control As System.Windows.Forms.Control,
            logicalPixels As Integer
        ) As Integer

            If control Is Nothing Then
                Return logicalPixels
            End If

            Dim dpi As Integer = 96

            Try
                dpi = control.DeviceDpi
            Catch ex As System.Exception
                dpi = 96
            End Try

            If dpi <= 0 Then
                dpi = 96
            End If

            Return CInt(
                System.Math.Round(
                    logicalPixels *
                    (CDbl(dpi) / 96.0R),
                    System.MidpointRounding.AwayFromZero
                )
            )
        End Function

        ' =====================================================================
        ' Owner capture / native owner wrapper
        ' =====================================================================
        Private NotInheritable Class AskUserOwnerInfo

            Public Sub New(
                handle As System.IntPtr,
                threadId As System.UInt32
            )
                Me.Handle = handle
                Me.ThreadId = threadId
            End Sub

            Public ReadOnly Property Handle As System.IntPtr
            Public ReadOnly Property ThreadId As System.UInt32
        End Class

        Private NotInheritable Class AskUserWindowOwner
            Implements System.Windows.Forms.IWin32Window

            Private ReadOnly _handle As System.IntPtr

            Public Sub New(handle As System.IntPtr)
                _handle = handle
            End Sub

            Public ReadOnly Property Handle As System.IntPtr _
                Implements System.Windows.Forms.IWin32Window.Handle
                Get
                    Return _handle
                End Get
            End Property
        End Class

        Private Shared Function CaptureAskUserOwnerInfo() As AskUserOwnerInfo
            Dim handle As System.IntPtr = System.IntPtr.Zero

            ' Prefer an active WinForms form on the caller thread. This is the
            ' strongest owner when ask_user was opened from Local Chat.
            Try
                Dim activeForm As System.Windows.Forms.Form =
                    System.Windows.Forms.Form.ActiveForm

                If activeForm IsNot Nothing AndAlso
                   Not activeForm.IsDisposed AndAlso
                   activeForm.IsHandleCreated Then

                    handle = activeForm.Handle
                End If
            Catch ex As System.Exception
                handle = System.IntPtr.Zero
            End Try

            ' Office/foreign/native host fallback.
            If handle = System.IntPtr.Zero Then
                Try
                    handle = AskUserGetForegroundWindow()
                Catch ex As System.Exception
                    handle = System.IntPtr.Zero
                End Try
            End If

            If handle = System.IntPtr.Zero Then
                Return New AskUserOwnerInfo(
                    System.IntPtr.Zero,
                    0UI
                )
            End If

            Try
                If Not AskUserIsWindow(handle) Then
                    Return New AskUserOwnerInfo(
                        System.IntPtr.Zero,
                        0UI
                    )
                End If

                Dim processId As System.UInt32 = 0UI
                Dim threadId As System.UInt32 =
                    AskUserGetWindowThreadProcessId(
                        handle,
                        processId
                    )

                Return New AskUserOwnerInfo(
                    handle,
                    threadId
                )
            Catch ex As System.Exception
                Return New AskUserOwnerInfo(
                    System.IntPtr.Zero,
                    0UI
                )
            End Try
        End Function

        <System.Runtime.InteropServices.DllImport(
            "user32.dll",
            EntryPoint:="GetForegroundWindow"
        )>
        Private Shared Function AskUserGetForegroundWindow() As System.IntPtr
        End Function

        <System.Runtime.InteropServices.DllImport(
            "user32.dll",
            EntryPoint:="IsWindow"
        )>
        Private Shared Function AskUserIsWindow(
            hWnd As System.IntPtr
        ) As Boolean
        End Function

        <System.Runtime.InteropServices.DllImport(
            "user32.dll",
            EntryPoint:="GetWindowThreadProcessId"
        )>
        Private Shared Function AskUserGetWindowThreadProcessId(
            hWnd As System.IntPtr,
            ByRef processId As System.UInt32
        ) As System.UInt32
        End Function

        <System.Runtime.InteropServices.DllImport(
            "kernel32.dll",
            EntryPoint:="GetCurrentThreadId"
        )>
        Private Shared Function AskUserGetCurrentThreadId() As System.UInt32
        End Function

    End Class
End Namespace
