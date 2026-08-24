' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ShowMethods.CustomForms.vb
' Purpose: Provides modal WinForms dialogs used across the SharedLibrary for
'          user interaction (selection lists, input boxes, message boxes, and
'          multi-parameter input forms), including sizing/layout behavior and
'          optional extra actions.
'
' Architecture:
'  - Native window integration: Uses `FindWindow` / `SendMessage` (user32) and
'    `WindowWrapper` ownership to optionally parent dialogs to Office app windows.
'  - Host-prompt visibility safeguard: several dialogs intentionally run as
'    `TopMost` so they do not disappear behind Office. In Word/Office, this can
'    create a specific deadlock-like UX when the host raises its own native prompt
'    (for example "Save changes?" while the user closes Word): the native prompt
'    gains focus but stays hidden behind our `TopMost` dialog, so the UI appears
'    stuck until the custom dialog is moved away manually.
'  - Centralized fix for that case: do not change owner behavior broadly and do
'    not rely on a single `Deactivate` event, because foreground activation is
'    transient and `GetForegroundWindow()` may briefly return zero or still point
'    to our dialog. Instead, `AttachForeignForegroundWatchdog` starts a short
'    WinForms timer for affected `TopMost` dialogs and repeatedly calls
'    `PromoteForeignForegroundDialog`, which temporarily drops only that dialog
'    out of the topmost band when a same-process host prompt takes foreground,
'    then restores normal behavior when the dialog is re-activated.
'  - Selection UI: `ShowSelectionForm` shows a fixed dialog with a ListBox and
'    OK/Cancel behavior.
'  - Text input UI: `ShowCustomInputBox` supports single-line and multi-line input,
'    optional shortcut insertion (Ctrl+P), and optional extra prefix buttons.
'  - Decision UI: `ShowCustomYesNoBox` returns an integer result for two buttons,
'    with optional auto-close and an optional extra button action.
'  - Notifications: `ShowCustomMessageBox` (plain text), `ShowRTFCustomMessageBox`
'    (RichTextBox), and `ShowHTMLCustomMessageBox` (WebBrowser on STA thread).
'  - Parameter collection: `ShowCustomVariableInputForm` builds controls from an
'    `InputParameter` array and writes back validated values when OK is pressed.
'  - Rich editor window: `ShowCustomWindow` shows editable content (optionally RTF)
'    with formatting buttons and multiple return modes.
' =============================================================================


Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Globalization
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary
    Partial Public Class SharedMethods


        ''' <summary>
        ''' Sends a message to the specified window handle (Win32).
        ''' </summary>
        ''' <param name="hWnd">Target window handle.</param>
        ''' <param name="msg">Message identifier.</param>
        ''' <param name="wParam">Additional message information (wParam).</param>
        ''' <param name="lParam">Additional message information (lParam).</param>
        ''' <returns>Message result.</returns>
        <DllImport("user32.dll", CharSet:=CharSet.Auto)>
        Private Shared Function SendMessage(
                    ByVal hWnd As IntPtr,
                    ByVal msg As Integer,
                    ByVal wParam As IntPtr,
                    ByVal lParam As IntPtr
                ) As IntPtr
        End Function

        ''' <summary>
        ''' Finds a top-level window by class name and/or window title (Win32).
        ''' </summary>
        ''' <param name="lpClassName">Window class name (e.g., "OpusApp").</param>
        ''' <param name="lpWindowName">Window title; may be <c>Nothing</c>.</param>
        ''' <returns>The window handle if found; otherwise <see cref="IntPtr.Zero"/>.</returns>
        <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
        Private Shared Function FindWindow(lpClassName As String, lpWindowName As String) As IntPtr
        End Function

        ''' <summary>
        ''' Detects an Office host application's top-level window handle for dialog ownership.
        ''' </summary>
        ''' <returns>
        ''' The window handle if a known Office application class is found; otherwise <see cref="IntPtr.Zero"/>.
        ''' </returns>
        Private Shared Function GetOfficeApplicationHwnd() As IntPtr
            ' Prefer the host window captured at add-in startup (UpdateHandler.HostHandle).
            ' It points at the actually running host (Outlook, Word or Excel). Probing by
            ' window class below would otherwise return a different Office app that merely
            ' happens to be open (e.g. Word in the background while the Outlook add-in shows
            ' a dialog), parenting the dialog to the wrong application, pushing it behind
            ' that app and stealing focus from it.
            Dim host As System.IntPtr = UpdateHandler.HostHandle
            If host <> System.IntPtr.Zero Then
                ' If the user is currently working in another top-level window of the SAME
                ' host process (e.g. an Outlook Inspector while editing an e-mail, rather
                ' than the main Explorer window), prefer that foreground window as the owner.
                ' Otherwise parenting to the main host window would bring the main window to
                ' the front and hide the Inspector the user is actually editing.
                Dim fg As System.IntPtr = NativeMethods.GetForegroundWindow()
                If fg <> System.IntPtr.Zero AndAlso fg <> host Then
                    Dim hostPid As Integer = 0
                    Dim fgPid As Integer = 0
                    NativeMethods.GetWindowThreadProcessId(host, hostPid)
                    NativeMethods.GetWindowThreadProcessId(fg, fgPid)
                    If hostPid <> 0 AndAlso fgPid = hostPid Then Return fg
                End If

                Return host
            End If

            ' Fallbacks (only used if the host handle was not captured at startup).
            ' Try Word first.
            Dim hwnd As System.IntPtr = FindWindow("OpusApp", Nothing)
            If hwnd <> System.IntPtr.Zero Then Return hwnd

            ' Try Excel.
            hwnd = FindWindow("XLMAIN", Nothing)
            If hwnd <> IntPtr.Zero Then Return hwnd

            ' Try Outlook.
            hwnd = FindWindow("rctrl_renwnd32", Nothing)
            If hwnd <> IntPtr.Zero Then Return hwnd

            Return IntPtr.Zero
        End Function


        ''' <summary>
        ''' Shows a fixed-size modal dialog with a prompt and a list of options to select from.
        ''' </summary>
        ''' <param name="prompt">Prompt text shown above the list.</param>
        ''' <param name="title">Window title.</param>
        ''' <param name="options">Options to populate the ListBox.</param>
        ''' <returns>
        ''' The selected option string, or the sentinel string <c>"ESC"</c> when canceled/closed via Escape.
        ''' </returns>
        Public Shared Function ShowSelectionForm(
                                            prompt As String,
                                            title As String,
                                            options As IEnumerable(Of String)
                                        ) As String

            Dim selectedOption As String = "ESC"

            ' Screen working area (accounts for taskbar, DPI, etc.).
            Dim wa As System.Drawing.Rectangle = System.Windows.Forms.Screen.FromPoint(System.Windows.Forms.Cursor.Position).WorkingArea

            ' Sizing constants.
            Const MIN_WIDTH As Integer = 450
            Const MIN_HEIGHT As Integer = 240
            Const SIDE_PADDING As Integer = 20
            Const LIST_CHROME As Integer = 35 ' scrollbar + borders (approx)

            ' Configure the form: resizable, DPI-aware.
            Dim inputForm As New System.Windows.Forms.Form() With {
        .Text = title,
        .FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable,
        .StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen,
        .MinimizeBox = False,
        .MaximizeBox = False,
        .ShowInTaskbar = False,
        .KeyPreview = True,
        .AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font,
        .MinimumSize = New System.Drawing.Size(MIN_WIDTH, MIN_HEIGHT)
    }
            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
            inputForm.Font = standardFont

            ' Use logo as icon.
            Dim bmp As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
            inputForm.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

            ' --- Measure content to determine optimal width ---
            Dim optionsList As String() = options.ToArray()
            Dim maxScreenWidth As Integer = CInt(wa.Width * 0.9)

            Dim measuredContentWidth As Integer = MIN_WIDTH
            Using g As System.Drawing.Graphics = inputForm.CreateGraphics()
                Dim maxItemTextWidth As Integer = 0
                For Each item As String In optionsList
                    Dim w As Integer = System.Windows.Forms.TextRenderer.MeasureText(
                        g,
                        item,
                        standardFont,
                        New System.Drawing.Size(Integer.MaxValue, Integer.MaxValue),
                        System.Windows.Forms.TextFormatFlags.SingleLine
                    ).Width
                    If w > maxItemTextWidth Then maxItemTextWidth = w
                Next

                Dim promptWidth As Integer = System.Windows.Forms.TextRenderer.MeasureText(
                    g,
                    prompt,
                    standardFont,
                    New System.Drawing.Size(Integer.MaxValue, Integer.MaxValue),
                    System.Windows.Forms.TextFormatFlags.SingleLine
                ).Width

                Dim neededClientWidth As Integer =
                    Math.Max(promptWidth + 2 * SIDE_PADDING, maxItemTextWidth + LIST_CHROME + 2 * SIDE_PADDING)

                measuredContentWidth = Math.Max(MIN_WIDTH, Math.Min(maxScreenWidth, neededClientWidth))
            End Using

            inputForm.ClientSize = New System.Drawing.Size(measuredContentWidth, 320)

            ' Main layout: prompt, ListBox, buttons.
            Dim layout As New System.Windows.Forms.TableLayoutPanel() With {
        .Dock = System.Windows.Forms.DockStyle.Fill,
        .ColumnCount = 1,
        .RowCount = 3
    }
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            inputForm.Controls.Add(layout)

            ' Prompt label with automatic wrapping.
            Dim labelPrompt As New System.Windows.Forms.Label() With {
        .Text = prompt,
        .AutoSize = True,
        .MaximumSize = New System.Drawing.Size(inputForm.ClientSize.Width - 40, 0),
        .Margin = New System.Windows.Forms.Padding(20, 20, 20, 10),
        .TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    }
            layout.Controls.Add(labelPrompt, 0, 0)

            ' ListBox with padding.
            Dim listPanel As New System.Windows.Forms.Panel() With {
        .Dock = System.Windows.Forms.DockStyle.Fill,
        .Padding = New System.Windows.Forms.Padding(20)
    }
            layout.Controls.Add(listPanel, 0, 1)

            Dim listBoxOptions As New System.Windows.Forms.ListBox() With {
        .Dock = System.Windows.Forms.DockStyle.Fill,
        .SelectionMode = System.Windows.Forms.SelectionMode.One
    }
            listBoxOptions.Items.AddRange(optionsList)
            listPanel.Controls.Add(listBoxOptions)

            ' Tooltip for truncated items.
            Dim listToolTip As New System.Windows.Forms.ToolTip() With {.ShowAlways = True}
            Dim lastToolTipIndex As Integer = -1

            AddHandler listBoxOptions.MouseMove,
                Sub(sender As Object, e As System.Windows.Forms.MouseEventArgs)
                    Dim idx As Integer = listBoxOptions.IndexFromPoint(e.Location)
                    If idx <> lastToolTipIndex Then
                        lastToolTipIndex = idx
                        If idx >= 0 AndAlso idx < listBoxOptions.Items.Count Then
                            Dim itemText As String = CStr(listBoxOptions.Items(idx))
                            Dim itemWidth As Integer = System.Windows.Forms.TextRenderer.MeasureText(itemText, listBoxOptions.Font).Width
                            Dim usableWidth As Integer = listBoxOptions.ClientSize.Width
                            If itemWidth > usableWidth Then
                                listToolTip.SetToolTip(listBoxOptions, itemText)
                            Else
                                listToolTip.SetToolTip(listBoxOptions, Nothing)
                            End If
                        Else
                            listToolTip.SetToolTip(listBoxOptions, Nothing)
                        End If
                    End If
                End Sub

            AddHandler listBoxOptions.MouseLeave,
                Sub(sender As Object, e As System.EventArgs)
                    lastToolTipIndex = -1
                    listToolTip.SetToolTip(listBoxOptions, Nothing)
                End Sub

            ' Left-aligned buttons with spacing.
            Dim panelButtons As New System.Windows.Forms.FlowLayoutPanel() With {
        .Dock = System.Windows.Forms.DockStyle.Fill,
        .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
        .Padding = New System.Windows.Forms.Padding(20, 10, 20, 20),
        .AutoSize = True,
        .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
        .WrapContents = False
    }
            layout.Controls.Add(panelButtons, 0, 2)

            ' OK button.
            Dim buttonOK As New System.Windows.Forms.Button() With {
        .Text = "OK",
        .DialogResult = System.Windows.Forms.DialogResult.OK,
        .Enabled = False,
        .AutoSize = True,
        .Padding = New System.Windows.Forms.Padding(8, 4, 5, 4),
        .Margin = New System.Windows.Forms.Padding(0, 0, 20, 0)
    }
            AddHandler buttonOK.Click, Sub()
                                           selectedOption = CStr(listBoxOptions.SelectedItem)
                                       End Sub

            ' Cancel button.
            Dim buttonCancel As New System.Windows.Forms.Button() With {
        .Text = "Cancel",
        .DialogResult = System.Windows.Forms.DialogResult.Cancel,
        .AutoSize = True,
        .Padding = New System.Windows.Forms.Padding(8, 4, 5, 4),
        .Margin = New System.Windows.Forms.Padding(0, 0, 0, 0)
    }
            AddHandler buttonCancel.Click, Sub()
                                               selectedOption = "ESC"
                                               inputForm.Close()
                                           End Sub

            panelButtons.Controls.Add(buttonOK)
            panelButtons.Controls.Add(buttonCancel)

            ' Ensure both buttons have the same height.
            Dim btnHeight As Integer = Math.Max(buttonOK.Height, buttonCancel.Height)
            buttonOK.Height = btnHeight
            buttonCancel.Height = btnHeight

            ' ListBox events.
            AddHandler listBoxOptions.SelectedIndexChanged, Sub()
                                                                buttonOK.Enabled = (listBoxOptions.SelectedItem IsNot Nothing)
                                                            End Sub
            AddHandler listBoxOptions.DoubleClick, Sub()
                                                       If listBoxOptions.SelectedItem IsNot Nothing Then
                                                           selectedOption = CStr(listBoxOptions.SelectedItem)
                                                           inputForm.DialogResult = System.Windows.Forms.DialogResult.OK
                                                           inputForm.Close()
                                                       End If
                                                   End Sub
            If listBoxOptions.Items.Count > 0 Then listBoxOptions.SelectedIndex = 0

            ' Keyboard shortcuts.
            inputForm.AcceptButton = buttonOK
            inputForm.CancelButton = buttonCancel
            AddHandler inputForm.KeyDown, Sub(sender As Object, e As System.Windows.Forms.KeyEventArgs)
                                              If e.KeyCode = System.Windows.Forms.Keys.Escape Then
                                                  selectedOption = "ESC"
                                                  inputForm.Close()
                                                  e.Handled = True
                                              End If
                                          End Sub

            ' Resize handler: keep label wrapping sensible.
            AddHandler inputForm.Resize,
                Sub()
                    Dim available As Integer = Math.Max(200, inputForm.ClientSize.Width - 40)
                    labelPrompt.MaximumSize = New System.Drawing.Size(available, 0)
                End Sub

            ' Show dialog.
            inputForm.TopMost = True

            SharedMethods.AttachForeignForegroundWatchdog(inputForm)

            Dim ownerWnd As System.Windows.Forms.IWin32Window = SharedMethods.ResolveSameThreadDialogOwner()
            If ownerWnd IsNot Nothing Then
                inputForm.ShowDialog(ownerWnd)
            Else
                inputForm.ShowDialog()
            End If

            Return selectedOption
        End Function


        ''' <summary>
        ''' Shows a modal input dialog supporting single-line or multi-line text entry.
        ''' </summary>
        ''' <param name="prompt">Prompt text shown above the input field.</param>
        ''' <param name="title">Window title.</param>
        ''' <param name="SimpleInput">
        ''' If <c>True</c>, uses a single-line TextBox; otherwise uses a multi-line TextBox with vertical scrolling.
        ''' </param>
        ''' <param name="DefaultValue">Initial text in the input field.</param>
        ''' <param name="CtrlP">Text inserted at caret position when Ctrl+P is pressed (if non-empty).</param>
        ''' <param name="OptionalButtons">
        ''' Optional extra buttons (up to 5). Each tuple is (ButtonLabel, TooltipText, PrefixToPrepend).
        ''' When clicked, the dialog returns OK and the prefix may be prepended to the final text.
        ''' </param>
        ''' <param name="InsertButtons">
        ''' Optional insert buttons shown only in multi-line mode. Each tuple is (ButtonLabel, TooltipText, TextToInsert).
        ''' When clicked, the specified text is inserted at the current caret position in the input field.
        ''' Examples: ("📄", "Insert document trigger", "{doc}"), ("📑", "Insert additional document trigger", "(adddoc)"), ("📊", "Insert worksheet trigger", "(addws)")
        ''' </param>
        ''' <returns>
        ''' On OK: the entered (and possibly prefixed) text.
        ''' On Cancel: returns <c>"ESC"</c> for multi-line mode and <c>""</c> for single-line mode.
        ''' </returns>
        Public Shared Function ShowCustomInputBox(
                                                    prompt As String,
                                                    title As String,
                                                    SimpleInput As Boolean,
                                                    Optional DefaultValue As String = "",
                                                    Optional CtrlP As String = "",
                                                    Optional OptionalButtons As System.Tuple(Of System.String, System.String, System.String)() = Nothing,
                                                    Optional InsertButtons As System.Tuple(Of System.String, System.String, System.String)() = Nothing,
                                                    Optional Context As ISharedContext = Nothing,
                                                    Optional InsertButtonMaxOccurrences As System.Collections.Generic.IDictionary(Of System.String, System.Int32) = Nothing
                                                ) As String

            ' Screen working area (accounts for taskbar, etc.).
            Dim wa As System.Drawing.Rectangle = Screen.FromPoint(Cursor.Position).WorkingArea

            ' Multi-line sizing rule: height = 1/6 of screen; width based on height.
            Dim desiredInputHeight As Integer = 0
            Dim desiredInputWidth As Integer = 0
            If Not SimpleInput Then
                desiredInputHeight = Math.Max(150, CInt(wa.Height / 6.0))
                desiredInputWidth = CInt(desiredInputHeight * 3)
                desiredInputWidth = Math.Min(desiredInputWidth, wa.Width - 60) ' Margin to fit in screen.
            End If

            ' Create and configure the form (resizable in both modes).
            Dim inputForm As New Form() With {
                .Opacity = 0,
                .Text = title,
                .FormBorderStyle = FormBorderStyle.Sizable,
                .StartPosition = FormStartPosition.Manual, ' Center within working area after layout.
                .MaximizeBox = False,
                .MinimizeBox = False,
                .ShowInTaskbar = False,
                .TopMost = True,
                .AutoScaleMode = AutoScaleMode.Font,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink
            }

            ' Set the icon.
            Dim bmp As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
            inputForm.Icon = Icon.FromHandle(bmp.GetHicon())

            ' Standard font.
            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            inputForm.Font = standardFont

            ' Main layout for dynamic resizing.
            Dim mainLayout As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 3,
                .Padding = New Padding(20),
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink
            }
            mainLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            If SimpleInput Then
                mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))         ' Label.
                mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))         ' Single-line TextBox.
                mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))         ' Buttons.
            Else
                mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))         ' Label.
                mainLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))  ' Multi-line TextBox grows/shrinks.
                mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))         ' Buttons.
            End If

            ' Prompt label (wrap to initial target width; updated on resize).
            Dim initialLabelWrap As Integer = If(SimpleInput,
                                                 Math.Min(wa.Width - 120, 700),
                                                 Math.Max(400, desiredInputWidth))
            Dim promptLabel As New System.Windows.Forms.Label() With {
                .Text = prompt,
                .Font = standardFont,
                .AutoSize = True,
                .MaximumSize = New Size(initialLabelWrap, 0)
            }
            promptLabel.Dock = DockStyle.Top
            mainLayout.Controls.Add(promptLabel, 0, 0)

            ' Input TextBox.
            Dim inputTextBox As New TextBox() With {
                .Font = standardFont,
                .Multiline = Not SimpleInput,
                .WordWrap = True,
                .ScrollBars = If(SimpleInput, ScrollBars.None, ScrollBars.Vertical),
                .Text = DefaultValue
            }
            If SimpleInput Then
                ' Single-line: compute height, stretch horizontally with the form.
                inputTextBox.Height = TextRenderer.MeasureText("Wy", standardFont).Height + 6
                inputTextBox.Anchor = AnchorStyles.Left Or AnchorStyles.Right
                inputTextBox.Width = initialLabelWrap
            Else
                ' Multi-line: initial size by rule; allow growing with the form.
                inputTextBox.MinimumSize = New Size(desiredInputWidth, desiredInputHeight)
                inputTextBox.Dock = DockStyle.Fill
            End If
            mainLayout.Controls.Add(inputTextBox, 0, 1)

            ' OK and Cancel buttons.
            Dim okButton As New Button() With {.Text = "OK", .AutoSize = True, .Font = standardFont}
            Dim cancelButton As New Button() With {.Text = "Cancel", .AutoSize = True, .Font = standardFont}

            AddHandler okButton.Click,
                Sub()
                    Dim violatingToken As System.String = System.String.Empty
                    Dim maximumOccurrences As System.Int32 = 0
                    If Not SharedMethods.ValidateCustomInputInsertOccurrenceLimits(
                        inputTextBox.Text,
                        InsertButtonMaxOccurrences,
                        violatingToken,
                        maximumOccurrences) Then

                        SharedMethods.ShowCustomMessageBox(
                            "'" & violatingToken & "' can be included at most " &
                            maximumOccurrences.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                            " time(s) in one request.")
                        inputTextBox.Focus()
                        Return
                    End If

                    inputForm.DialogResult = DialogResult.OK
                    inputForm.Close()
                End Sub
            AddHandler cancelButton.Click, Sub()
                                               inputForm.DialogResult = DialogResult.Cancel
                                               inputForm.Close()
                                           End Sub

            ' Bottom flow with wrapping so all buttons remain visible if space narrows.
            Dim bottomFlow As New FlowLayoutPanel() With {
                .FlowDirection = FlowDirection.LeftToRight,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Margin = New Padding(0, 20, 0, 0),
                .Dock = DockStyle.Top,
                .WrapContents = True
            }
            bottomFlow.Controls.Add(okButton)
            bottomFlow.Controls.Add(cancelButton)

            ' Optional extra buttons (max 5): label, tooltip, and prefix.
            Dim selectedPrefix As String = Nothing
            If OptionalButtons IsNot Nothing AndAlso OptionalButtons.Length > 0 Then
                Dim tip As New System.Windows.Forms.ToolTip()
                Dim count As Integer = Math.Min(5, OptionalButtons.Length)
                For i As Integer = 0 To count - 1
                    Dim item = OptionalButtons(i)
                    Dim extraBtn As New System.Windows.Forms.Button() With {
                        .Text = item.Item1,
                        .AutoSize = True,
                        .Font = standardFont
                    }
                    tip.SetToolTip(extraBtn, item.Item2)
                    If i = 0 Then
                        extraBtn.Margin = New Padding(cancelButton.Margin.Left * 2, cancelButton.Margin.Top, cancelButton.Margin.Right, cancelButton.Margin.Bottom)
                    End If
                    AddHandler extraBtn.Click,
                        Sub()
                            selectedPrefix = item.Item3
                            inputForm.DialogResult = DialogResult.OK
                            inputForm.Close()
                        End Sub
                    bottomFlow.Controls.Add(extraBtn)
                Next
            End If

            ' Insert buttons for multi-line mode: insert text at caret position.
            If Not SimpleInput AndAlso InsertButtons IsNot Nothing AndAlso InsertButtons.Length > 0 Then
                Dim insertTip As New System.Windows.Forms.ToolTip()
                Dim emojiFont As New System.Drawing.Font("Segoe UI Emoji", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
                For i As Integer = 0 To InsertButtons.Length - 1
                    Dim insertItem = InsertButtons(i)
                    Dim insertBtn As New System.Windows.Forms.Button() With {
                        .Text = insertItem.Item1,
                        .AutoSize = True,
                        .Font = emojiFont
                    }
                    insertTip.SetToolTip(insertBtn, insertItem.Item2)
                    If i = 0 Then
                        ' Add extra left margin to visually separate insert buttons from action buttons.
                        insertBtn.Margin = New Padding(cancelButton.Margin.Left * 3, cancelButton.Margin.Top, cancelButton.Margin.Right, cancelButton.Margin.Bottom)
                    End If
                    Dim textToInsert As String = insertItem.Item3
                    AddHandler insertBtn.Click,
                        Sub()
                            Dim maxOccurrences As System.Int32 = SharedMethods.GetCustomInputInsertMaximum(
                                InsertButtonMaxOccurrences,
                                textToInsert)

                            If maxOccurrences > 0 AndAlso
                               SharedMethods.CountFreestyleTokenOccurrences(inputTextBox.Text, textToInsert) >= maxOccurrences Then

                                SharedMethods.ShowCustomMessageBox(
                                    "This item can be included at most " &
                                    maxOccurrences.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                                    " time(s) in one request.")
                                inputTextBox.Focus()
                                Return
                            End If

                            Dim selPos = inputTextBox.SelectionStart
                            inputTextBox.Text = inputTextBox.Text.Insert(selPos, textToInsert)
                            inputTextBox.SelectionStart = selPos + textToInsert.Length
                            inputTextBox.Focus()
                        End Sub
                    bottomFlow.Controls.Add(insertBtn)
                Next
            End If

            mainLayout.Controls.Add(bottomFlow, 0, 2)
            inputForm.Controls.Add(mainLayout)

            ' Resize handler to keep label wrapping sensible when the user resizes the form.
            AddHandler inputForm.Resize, Sub()
                                             ' Available width for content inside padding.
                                             Dim available As Integer = Math.Max(300, mainLayout.ClientSize.Width)
                                             promptLabel.MaximumSize = New Size(available, 0)
                                             promptLabel.Invalidate()
                                         End Sub

            ' KeyDown handlers for Enter/Escape.
            If SimpleInput Then
                AddHandler inputTextBox.KeyDown, Sub(sender, e)
                                                     If e.KeyCode = Keys.Enter Then
                                                         inputForm.DialogResult = DialogResult.OK
                                                         inputForm.Close()
                                                         e.SuppressKeyPress = True
                                                     End If
                                                 End Sub
            Else
                AddHandler inputTextBox.KeyDown, Sub(sender, e)
                                                     If e.KeyCode = Keys.Enter AndAlso e.Modifiers = Keys.Control Then
                                                         inputForm.DialogResult = DialogResult.OK
                                                         inputForm.Close()
                                                         e.SuppressKeyPress = True
                                                     ElseIf e.KeyCode = Keys.Escape Then
                                                         inputForm.DialogResult = DialogResult.Cancel
                                                         inputForm.Close()
                                                         e.SuppressKeyPress = True
                                                     End If
                                                 End Sub
            End If

            ' Ctrl+P insertion, if provided.
            If Not String.IsNullOrEmpty(CtrlP) Then
                AddHandler inputTextBox.KeyDown, Sub(sender, e)
                                                     If e.KeyCode = Keys.P AndAlso e.Modifiers = Keys.Control Then
                                                         Dim selPos = inputTextBox.SelectionStart
                                                         inputTextBox.Text = inputTextBox.Text.Insert(selPos, CtrlP)
                                                         inputTextBox.SelectionStart = selPos + CtrlP.Length
                                                         e.SuppressKeyPress = True
                                                     End If
                                                 End Sub
            End If

            ' Slash-triggered prompt library insertion for multi-line mode.
            If Not SimpleInput AndAlso Context IsNot Nothing AndAlso Context.INI_PromptLib Then
                Dim promptLibraryPath As String = Context.INI_PromptLibPath
                Dim promptLibraryPathLocal As String = Context.INI_PromptLibPathLocal
                Dim promptLibraryContext As ISharedContext = Context

                AddHandler inputTextBox.KeyPress,
                    Sub(sender, e)
                        If e.KeyChar <> "/"c Then Return

                        Dim slashAction As SharedMethods.PromptLibrarySlashAction =
                            SharedMethods.HandlePromptLibrarySlash(
                                inputTextBox,
                                promptLibraryPath,
                                promptLibraryPathLocal,
                                promptLibraryContext,
                                CtrlP,
                                True   ' NoWarning
                            )

                        If slashAction <> SharedMethods.PromptLibrarySlashAction.NotTriggered Then
                            e.Handled = True
                        End If
                    End Sub
            End If

            ' After AutoSize computed, clamp to screen, set MinimumSize (so buttons stay visible),
            ' disable AutoSize to allow user resizing, and center within the working area.
            AddHandler inputForm.Shown, Sub()
                                            ' Let AutoSize produce the preferred size first.
                                            inputForm.PerformLayout()

                                            Dim maxW As Integer = wa.Width - 40
                                            Dim maxH As Integer = wa.Height - 40

                                            ' Ensure the form is wide enough to show all buttons in one row.
                                            bottomFlow.PerformLayout()
                                            Dim requiredButtonWidth As Integer = bottomFlow.PreferredSize.Width + mainLayout.Padding.Horizontal
                                            Dim chromeW As Integer = inputForm.Width - inputForm.ClientSize.Width
                                            Dim minClientW As Integer = Math.Max(inputForm.ClientSize.Width, requiredButtonWidth)
                                            minClientW = Math.Min(minClientW, maxW - chromeW)

                                            ' Compute space used by non-textbox rows and window chrome.
                                            Dim chromeH As Integer = inputForm.Height - inputForm.ClientSize.Height
                                            Dim labelH As Integer = promptLabel.PreferredSize.Height
                                            Dim buttonsH As Integer = bottomFlow.PreferredSize.Height
                                            Dim paddingV As Integer = mainLayout.Padding.Vertical
                                            Dim gaps As Integer = bottomFlow.Margin.Top ' Vertical gap above buttons.

                                            Dim fixedRowsH As Integer = paddingV + labelH + gaps + buttonsH
                                            Dim maxClientH As Integer = maxH - chromeH

                                            If Not SimpleInput Then
                                                ' Allocate remaining height to the textbox, but stay within working area.
                                                Dim textH As Integer = Math.Max(100, Math.Min(desiredInputHeight, maxClientH - fixedRowsH))

                                                ' Set client size so all rows are visible.
                                                Dim newClientH As Integer = Math.Min(fixedRowsH + textH, maxClientH)

                                                ' Use the wider of autosized width or required button width, clamped to screen.
                                                Dim newClientW As Integer = Math.Min(minClientW, maxW)

                                                inputForm.ClientSize = New Size(newClientW, newClientH)
                                            Else
                                                ' SimpleInput: ensure button width is accommodated, then clamp to screen.
                                                inputForm.ClientSize = New Size(minClientW, inputForm.ClientSize.Height)
                                                If inputForm.Width > maxW Then inputForm.Width = maxW
                                                If inputForm.Height > maxH Then inputForm.Height = maxH
                                            End If

                                            ' Minimum cannot be smaller than the current fully-visible content.
                                            inputForm.MinimumSize = inputForm.Size

                                            ' Now allow resizing (keep MinimumSize so content/buttons never get clipped).
                                            inputForm.AutoSize = False

                                            ' Center within working area.
                                            inputForm.Location = New System.Drawing.Point(
                                                wa.X + (wa.Width - inputForm.Width) \ 2,
                                                wa.Y + (wa.Height - inputForm.Height) \ 2
                                            )
                                        End Sub

            ' Ensure focus/topmost.
            inputForm.TopMost = True
            inputForm.BringToFront()
            inputForm.Focus()

            SharedMethods.AttachForeignForegroundWatchdog(inputForm)

            Dim Result As DialogResult


            ' Show the dialog, must be owned by Outlook (only then the title may contains "Browser").

            If title.Contains("Browser") Then
                ' Activate Outlook window via Win32 (no COM object needed since we're already in-process).
                Dim outlookHwnd As IntPtr = FindWindow("rctrl_renwnd32", Nothing)
                If outlookHwnd <> IntPtr.Zero Then
                    Const SW_RESTORE As Integer = 9
                    Const WM_SYSCOMMAND As Integer = &H112
                    Const SC_RESTORE As Integer = &HF120
                    SendMessage(outlookHwnd, WM_SYSCOMMAND, New IntPtr(SC_RESTORE), IntPtr.Zero)
                    SetForegroundWindow(outlookHwnd)
                End If

                inputForm.Opacity = 1

                Dim __browserOwner As System.Windows.Forms.IWin32Window = Nothing
                If outlookHwnd <> IntPtr.Zero Then
                    __browserOwner = SharedMethods.IfOwnerOnCurrentThread(New WindowWrapper(outlookHwnd))
                End If
                If __browserOwner IsNot Nothing Then
                    Result = inputForm.ShowDialog(__browserOwner)
                Else
                    Dim __owner As System.Windows.Forms.IWin32Window = SharedMethods.ResolveSameThreadDialogOwner()
                    If __owner IsNot Nothing Then
                        Result = inputForm.ShowDialog(__owner)
                    Else
                        Result = inputForm.ShowDialog()
                    End If
                End If
            Else
                inputForm.Opacity = 1
                Dim __owner As System.Windows.Forms.IWin32Window = SharedMethods.ResolveSameThreadDialogOwner()
                If __owner IsNot Nothing Then
                    Result = inputForm.ShowDialog(__owner)
                Else
                    Result = inputForm.ShowDialog()
                End If
            End If

            ' Return the entered text or appropriate default.
            If Result = DialogResult.OK Then
                Dim finalText As String = inputTextBox.Text
                If Not String.IsNullOrEmpty(selectedPrefix) AndAlso Not finalText.StartsWith(selectedPrefix, StringComparison.OrdinalIgnoreCase) Then
                    finalText = selectedPrefix & " " & finalText
                End If
                Debug.WriteLine("Final text: " & finalText)
                Return finalText
            Else
                Return If(Not SimpleInput, "ESC", "")
            End If
        End Function


        <DllImport("user32.dll")>
        Private Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
        End Function


        ''' <summary>
        ''' Shows a modal Yes/No-style dialog with two custom button labels and an optional auto-close timer.
        ''' When nonModal is True, the dialog stays topmost but allows interaction with other windows.
        ''' </summary>
        ''' <param name="bodyText">Dialog body text (truncated to 10000 characters as implemented).</param>
        ''' <param name="button1Text">Text for the first button (result 1).</param>
        ''' <param name="button2Text">Text for the second button (result 2).</param>
        ''' <param name="header">Dialog title. Defaults to <c>AN</c>.</param>
        ''' <param name="autoCloseSeconds">If set, the dialog closes after this many seconds and returns 3.</param>
        ''' <param name="Defaulttext">Suffix appended to the countdown label text.</param>
        ''' <param name="extraButtonText">Optional extra button text (only when no auto-close is active).</param>
        ''' <param name="extraButtonAction">Action invoked when the extra button is clicked.</param>
        ''' <param name="CloseAfterExtra">If <c>True</c>, closes the dialog after invoking the extra action.</param>
        ''' <param name="nonModal">If <c>True</c>, shows the dialog non-modally with topmost behavior, allowing interaction with other windows.</param>
        ''' <returns>1 for button1, 2 for button2, 3 for auto-close; otherwise 0 (initial value/cancelled).</returns>
        Public Shared Function ShowCustomYesNoBox(
                        ByVal bodyText As String,
                        ByVal button1Text As String,
                        ByVal button2Text As String,
                        Optional header As String = AN,
                        Optional autoCloseSeconds As Integer? = Nothing,
                        Optional Defaulttext As String = "",
                        Optional extraButtonText As String = Nothing,
                        Optional extraButtonAction As System.Action = Nothing,
                        Optional CloseAfterExtra As Boolean = False,
                        Optional nonModal As Boolean = False
                    ) As Integer

            ' Screen working area.
            Dim wa As Rectangle = Screen.FromPoint(Cursor.Position).WorkingArea
            Dim maxScreenHeight As Integer = CInt(wa.Height * 0.5)
            Dim maxScreenWidth As Integer = CInt(wa.Width * 0.9)

            ' Constants.
            Const MIN_WIDTH As Integer = 450
            Const PADDING As Integer = 20
            Const BUTTON_GAP As Integer = 10
            Const ASPECT_RATIO As Double = 16.0 / 9.0

            ' Create and configure form (resizable).
            Dim messageForm As New Form() With {
                .Opacity = 0,
                .Text = header,
                .FormBorderStyle = FormBorderStyle.Sizable,
                .StartPosition = FormStartPosition.CenterScreen,
                .MaximizeBox = False,
                .MinimizeBox = False,
                .ShowInTaskbar = If(nonModal, True, False),
                .TopMost = True,
                .AutoScaleMode = AutoScaleMode.Font
            }

            ' Icon.
            Dim bmpIcon As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
            messageForm.Icon = Icon.FromHandle(bmpIcon.GetHicon())

            ' Font.
            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            messageForm.Font = standardFont

            ' Create buttons first to measure their size.
            Dim button1 As New Button() With {
                .Text = button1Text,
                .AutoSize = True,
                .Font = standardFont
            }
            Dim button2 As New Button() With {
                .Text = button2Text,
                .AutoSize = True,
                .Font = standardFont
            }
            Dim countdownLabel As New System.Windows.Forms.Label() With {
                .Font = standardFont,
                .AutoSize = True
            }

            ' Result variable.
            Dim result As Integer = 0

            ' For non-modal, we need a signal to know when dialog is closed
            Dim dialogClosed As Threading.ManualResetEvent = Nothing
            If nonModal Then
                dialogClosed = New Threading.ManualResetEvent(False)
            End If

            AddHandler button1.Click, Sub()
                                          result = 1
                                          messageForm.Close()
                                      End Sub
            AddHandler button2.Click, Sub()
                                          result = 2
                                          messageForm.Close()
                                      End Sub

            ' Bottom flow for buttons.
            Dim bottomFlow As New FlowLayoutPanel() With {
                .FlowDirection = FlowDirection.LeftToRight,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Dock = DockStyle.Bottom,
                .Padding = New Padding(PADDING, BUTTON_GAP, PADDING, PADDING),
                .WrapContents = True
            }
            bottomFlow.Controls.Add(button1)
            bottomFlow.Controls.Add(button2)

            ' Optional extra button.
            Dim extraButton As Button = Nothing
            If (Not autoCloseSeconds.HasValue) AndAlso
               (Not String.IsNullOrEmpty(extraButtonText)) AndAlso
               (extraButtonAction IsNot Nothing) Then

                extraButton = New Button() With {
                    .Text = extraButtonText,
                    .AutoSize = True,
                    .Font = standardFont,
                    .Margin = New Padding(BUTTON_GAP, button1.Margin.Top, 0, button1.Margin.Bottom)
                }

                AddHandler extraButton.Click,
                    Sub()
                        Try
                            extraButtonAction.Invoke()
                        Catch ex As System.Exception
                            ' Swallow to keep dialog functional.
                        End Try
                        If CloseAfterExtra Then messageForm.Close()
                    End Sub

                bottomFlow.Controls.Add(extraButton)
            End If

            If autoCloseSeconds.HasValue Then
                bottomFlow.Controls.Add(countdownLabel)
            End If

            ' Measure button panel height.
            bottomFlow.PerformLayout()
            Dim buttonPanelHeight As Integer = bottomFlow.PreferredSize.Height

            ' Body label for text measurement.
            Dim bodyLabel As New System.Windows.Forms.Label() With {
                .Text = bodyText,
                .Font = standardFont,
                .AutoSize = True
            }

            ' Calculate optimal dimensions with 16:9 aspect ratio preference.
            Dim chromeWidth As Integer = messageForm.Width - messageForm.ClientSize.Width + 20
            Dim chromeHeight As Integer = messageForm.Height - messageForm.ClientSize.Height

            ' Start with minimum width and calculate text height.
            Dim contentWidth As Integer = MIN_WIDTH - 2 * PADDING
            bodyLabel.MaximumSize = New Size(contentWidth, 0)
            Dim textSize As Size = bodyLabel.GetPreferredSize(New Size(contentWidth, 0))

            ' Try to achieve 16:9 ratio by widening if text is tall.
            Dim targetHeight As Integer = textSize.Height + buttonPanelHeight + PADDING
            Dim targetWidth As Integer = MIN_WIDTH

            ' Iteratively widen to approach 16:9 ratio while text is taller than optimal.
            Dim iterations As Integer = 0
            While iterations < 20 AndAlso targetHeight > 0
                Dim optimalHeight As Integer = CInt(targetWidth / ASPECT_RATIO)
                If targetHeight <= optimalHeight OrElse targetWidth >= maxScreenWidth Then
                    Exit While
                End If

                ' Increase width.
                targetWidth = Math.Min(targetWidth + 50, maxScreenWidth)
                contentWidth = targetWidth - 2 * PADDING - chromeWidth

                bodyLabel.MaximumSize = New Size(contentWidth, 0)
                textSize = bodyLabel.GetPreferredSize(New Size(contentWidth, 0))
                targetHeight = textSize.Height + buttonPanelHeight + PADDING

                iterations += 1
            End While

            ' Determine if scrolling is needed.
            Dim needsScroll As Boolean = textSize.Height > (maxScreenHeight - buttonPanelHeight - PADDING)
            Dim bodyPanelHeight As Integer

            If needsScroll Then
                bodyPanelHeight = maxScreenHeight - buttonPanelHeight - PADDING
                ' Account for scrollbar width.
                contentWidth = contentWidth - SystemInformation.VerticalScrollBarWidth
                bodyLabel.MaximumSize = New Size(contentWidth, 0)
            Else
                bodyPanelHeight = textSize.Height
            End If

            ' Create scrollable body container.
            Dim bodyScrollPanel As New Panel() With {
                .Dock = DockStyle.Fill,
                .AutoScroll = needsScroll,
                .Padding = New Padding(PADDING, PADDING, PADDING, BUTTON_GAP)
            }

            bodyLabel.MaximumSize = New Size(contentWidth, 0)
            bodyLabel.Location = New System.Drawing.Point(PADDING, PADDING)  ' Respect the left and top padding            
            bodyScrollPanel.Controls.Add(bodyLabel)

            If needsScroll Then
                bodyScrollPanel.AutoScrollMinSize = New Size(contentWidth, textSize.Height + PADDING)  ' Add padding to scroll size
            End If

            ' Main layout using TableLayoutPanel for proper resizing.
            Dim mainLayout As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 2,
                .Padding = New Padding(0),
                .Margin = New Padding(0)
            }
            mainLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            mainLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F)) ' Body expands.
            mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))        ' Buttons fixed.

            mainLayout.Controls.Add(bodyScrollPanel, 0, 0)
            mainLayout.Controls.Add(bottomFlow, 0, 1)

            messageForm.Controls.Add(mainLayout)

            ' Calculate final form size.
            ' Account for the bodyScrollPanel internal padding (top PADDING + bottom BUTTON_GAP)
            ' plus the bodyLabel offset within the panel, to ensure the last wrapped line is visible.
            Dim bodyScrollPaddingV As Integer = bodyScrollPanel.Padding.Vertical  ' PADDING + BUTTON_GAP = 30
            Dim clientWidth As Integer = Math.Max(MIN_WIDTH, targetWidth) - chromeWidth
            Dim clientHeight As Integer = bodyPanelHeight + bodyScrollPaddingV + buttonPanelHeight + PADDING + BUTTON_GAP

            messageForm.ClientSize = New Size(clientWidth, clientHeight)

            ' Set minimum size to ensure buttons always visible.
            Dim minButtonWidth As Integer = bottomFlow.PreferredSize.Width + 2 * PADDING
            messageForm.MinimumSize = New Size(
                Math.Max(MIN_WIDTH, minButtonWidth + chromeWidth),
                buttonPanelHeight + 100 + chromeHeight
            )

            ' Resize handler 
            Dim ApplyLayout As System.Action =
                        Sub()
                            Dim availableWidth As Integer = bodyScrollPanel.ClientSize.Width - 2 * PADDING
                            If bodyScrollPanel.AutoScroll Then
                                availableWidth -= SystemInformation.VerticalScrollBarWidth
                            End If

                            bodyLabel.MaximumSize = New Size(Math.Max(100, availableWidth), 0)
                            bodyLabel.PerformLayout()

                            If bodyScrollPanel.AutoScroll Then
                                bodyScrollPanel.AutoScrollMinSize = New Size(availableWidth, bodyLabel.PreferredHeight + PADDING)
                            End If
                        End Sub

            ' Run once at the start
            messageForm.PerformLayout()
            ApplyLayout.Invoke()

            ' Run on every resize
            AddHandler messageForm.Resize, Sub() ApplyLayout.Invoke()

            ' Auto-close timer.
            If autoCloseSeconds.HasValue Then
                Dim remaining = autoCloseSeconds.Value
                countdownLabel.Text = $"(closes in {remaining} seconds{Defaulttext})"
                Dim t As New System.Windows.Forms.Timer() With {.Interval = 1000}
                AddHandler t.Tick, Sub()
                                       remaining -= 1
                                       If remaining > 0 Then
                                           countdownLabel.Text = $"(closes in {remaining} seconds{Defaulttext})"
                                       Else
                                           t.Stop()
                                           result = 3
                                           messageForm.Close()
                                       End If
                                   End Sub
                t.Start()
            End If

            ' For non-modal: keep topmost on deactivate and signal when closed
            If nonModal Then
                AddHandler messageForm.Deactivate, Sub(sender, e)
                                                       Try
                                                           messageForm.TopMost = True
                                                       Catch
                                                       End Try
                                                   End Sub

                AddHandler messageForm.FormClosed, Sub(sender, e)
                                                       dialogClosed.Set()
                                                   End Sub
            End If

            ' Show and return.
            messageForm.TopMost = True
            messageForm.Opacity = 1
            messageForm.BringToFront()
            messageForm.Focus()
            messageForm.Activate()

            AddHandler messageForm.Shown,
                Sub(sender, e)
                    SharedMethods.ForceDialogToForeground(messageForm)
                End Sub

            If nonModal Then
                ' Show non-modal and pump messages until closed
                messageForm.Show()
                messageForm.BringToFront()
                messageForm.Activate()

                ' Pump messages until dialog is closed
                While Not dialogClosed.WaitOne(50)
                    System.Windows.Forms.Application.DoEvents()
                End While

                messageForm.Dispose()
            Else
                ' Show modal.
                SharedMethods.AttachForeignForegroundWatchdog(messageForm)

                Dim __owner As System.Windows.Forms.IWin32Window = SharedMethods.ResolveSameThreadDialogOwner()
                If __owner IsNot Nothing Then
                    messageForm.ShowDialog(__owner)
                Else
                    messageForm.ShowDialog()
                End If
            End If

            Return result
        End Function


        ''' <summary>
        ''' Shows a modal message dialog with OK button and optional auto-close behavior.
        ''' </summary>
        ''' <param name="bodyText">Text content (truncated to 10000 characters as implemented).</param>
        ''' <param name="header">Dialog title. Defaults to <c>AN</c> if empty/whitespace.</param>
        ''' <param name="autoCloseSeconds">If set, counts down and closes the dialog automatically.</param>
        ''' <param name="Defaulttext">Suffix appended to the countdown label text.</param>
        ''' <param name="SeparateThread">
        ''' If <c>True</c> and auto-close is enabled, shows the dialog using <c>ShowDialog()</c>;
        ''' otherwise uses <c>Show()</c> with <c>Application.DoEvents()</c>.
        ''' </param>
        ''' <param name="extraButtonText">Optional extra button text (only when no auto-close is active).</param>
        ''' <param name="extraButtonAction">Action invoked when the extra button is clicked.</param>
        ''' <param name="CloseAfterExtra">If <c>True</c>, closes the dialog after invoking the extra action.</param>
        Public Shared Sub ShowCustomMessageBox(
    ByVal bodyText As String,
    Optional header As String = AN,
    Optional autoCloseSeconds As System.Nullable(Of Integer) = Nothing,
    Optional Defaulttext As String = " - execution continues meanwhile",
    Optional SeparateThread As Boolean = False,
    Optional extraButtonText As String = Nothing,
    Optional extraButtonAction As System.Action = Nothing,
    Optional CloseAfterExtra As Boolean = False
)
            If System.String.IsNullOrWhiteSpace(header) Then header = AN
            Dim isTruncated As System.Boolean = False
            If bodyText IsNot Nothing AndAlso bodyText.Length > 10000 Then
                bodyText = bodyText.Substring(0, 10000) & "(...)"
                isTruncated = True
            End If

            Dim messageForm As New System.Windows.Forms.Form() With {
        .Opacity = 0,
        .Text = header,
        .FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog,
        .StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen,
        .MaximizeBox = False,
        .MinimizeBox = False,
        .ShowInTaskbar = False,
        .TopMost = True,
        .AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font,
        .AutoSize = False
    }

            Dim bmpIcon As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
            messageForm.Icon = System.Drawing.Icon.FromHandle(bmpIcon.GetHicon())

            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
            messageForm.Font = standardFont

            Dim wa As System.Drawing.Rectangle = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea
            Dim paddingAll As System.Int32 = 20
            Dim gapAboveButtons As System.Int32 = 10 ' Keep existing gap logic.
            Dim spacerExtra As System.Int32 = 20    ' Extra space between text and buttons.
            Dim minContentWidth As System.Int32 = 360
            Dim startContentWidth As System.Int32 = 500
            Dim maxWindowWidth As System.Int32 = CInt(System.Math.Floor(wa.Width * 0.5))
            Dim maxWindowHeight As System.Int32 = CInt(System.Math.Floor(wa.Height * 0.9))

            Dim okButton As New System.Windows.Forms.Button() With {.Text = "OK", .AutoSize = True, .Font = standardFont, .Margin = New System.Windows.Forms.Padding(0)}
            Dim countdownLabel As New System.Windows.Forms.Label() With {.Font = standardFont, .AutoSize = True, .Margin = New System.Windows.Forms.Padding(8, 0, 0, 0)}
            Dim userClicked As System.Boolean = False
            AddHandler okButton.Click, Sub()
                                           userClicked = True
                                           messageForm.Close()
                                       End Sub

            Dim bottomFlow As New System.Windows.Forms.FlowLayoutPanel() With {
        .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
        .AutoSize = True,
        .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
        .Margin = New System.Windows.Forms.Padding(0)
    }
            bottomFlow.Controls.Add(okButton)

            ' Optional extra button.
            If (Not autoCloseSeconds.HasValue) AndAlso
       (Not System.String.IsNullOrEmpty(extraButtonText)) AndAlso
       (extraButtonAction IsNot Nothing) Then


                Dim extraButton As New System.Windows.Forms.Button() With {
                            .Text = extraButtonText,
                            .AutoSize = True,
                            .Font = standardFont,
                            .Margin = New System.Windows.Forms.Padding(8, okButton.Margin.Top, 0, okButton.Margin.Bottom)
                        }


                AddHandler extraButton.Click,
            Sub()
                Try
                    extraButtonAction.Invoke()
                Catch ex As System.Exception
                    ' Swallow to keep dialog functional.
                End Try
                If CloseAfterExtra Then messageForm.Close()
            End Sub
                bottomFlow.Controls.Add(extraButton)
            End If
            If autoCloseSeconds.HasValue Then bottomFlow.Controls.Add(countdownLabel)

            bottomFlow.PerformLayout()
            Dim bottomSize As System.Drawing.Size = bottomFlow.PreferredSize
            Dim reservedBottomHeight As System.Int32 = bottomSize.Height + gapAboveButtons

            Dim bodyLabel As New System.Windows.Forms.Label() With {
        .Text = If(bodyText, System.String.Empty),
        .Font = standardFont,
        .AutoSize = True,
        .Margin = New System.Windows.Forms.Padding(0)
    }

            Dim GetLabelPreferred As System.Func(Of System.Int32, System.Drawing.Size) =
        Function(w As System.Int32) As System.Drawing.Size
            bodyLabel.MaximumSize = New System.Drawing.Size(System.Math.Max(1, w), 0)
            Return bodyLabel.GetPreferredSize(New System.Drawing.Size(System.Math.Max(1, w), 0))
        End Function

            Dim contentWidth As System.Int32 = System.Math.Max(minContentWidth, System.Math.Min(startContentWidth, maxWindowWidth - 2 * paddingAll))
            Dim pref As System.Drawing.Size = GetLabelPreferred(contentWidth)
            Dim maxBodyHeightNoScroll As System.Int32 = System.Math.Max(100, maxWindowHeight - reservedBottomHeight - spacerExtra - 2 * paddingAll) ' Include spacer in budget.

            While (pref.Height > maxBodyHeightNoScroll) AndAlso ((contentWidth + 2 * paddingAll) < maxWindowWidth)
                Dim stepW As System.Int32 = System.Math.Max(24, (maxWindowWidth - 2 * paddingAll - contentWidth) \ 3)
                contentWidth = System.Math.Min(maxWindowWidth - 2 * paddingAll, contentWidth + stepW)
                pref = GetLabelPreferred(contentWidth)
            End While

            Dim needScroll As System.Boolean = pref.Height > maxBodyHeightNoScroll
            Dim usableTextWidth As System.Int32 = contentWidth
            If needScroll Then
                usableTextWidth = System.Math.Max(100, contentWidth - System.Windows.Forms.SystemInformation.VerticalScrollBarWidth)
                pref = GetLabelPreferred(usableTextWidth)
            End If

            Dim bodyPanelHeight As System.Int32 = If(needScroll, maxBodyHeightNoScroll, pref.Height)

            Dim bodyScrollPanel As New System.Windows.Forms.Panel() With {
        .AutoScroll = False,
        .AutoSize = False,
        .Size = New System.Drawing.Size(contentWidth, bodyPanelHeight),
        .Margin = New System.Windows.Forms.Padding(0),
        .Padding = New System.Windows.Forms.Padding(0)
    }
            bodyScrollPanel.HorizontalScroll.Enabled = False
            bodyScrollPanel.HorizontalScroll.Visible = False

            bodyLabel.MaximumSize = New System.Drawing.Size(usableTextWidth, 0)
            bodyScrollPanel.Controls.Add(bodyLabel)
            bodyLabel.Location = New System.Drawing.Point(0, 0)

            If needScroll Then
                bodyScrollPanel.AutoScroll = True
                bodyScrollPanel.AutoScrollMinSize = New System.Drawing.Size(usableTextWidth, pref.Height)
            End If

            ' Main table: [text][spacer][buttons].
            Dim table As New System.Windows.Forms.TableLayoutPanel() With {
        .Dock = System.Windows.Forms.DockStyle.Fill,
        .ColumnCount = 1,
        .RowCount = 3,
        .Padding = New System.Windows.Forms.Padding(paddingAll),
        .AutoSize = False,
        .Margin = New System.Windows.Forms.Padding(0)
    }
            table.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
            table.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, bodyPanelHeight))  ' Text.
            table.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, spacerExtra))       ' Spacer.
            table.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))                  ' Buttons.

            table.Controls.Add(bodyScrollPanel, 0, 0)

            ' Spacer: exact spacerExtra above the buttons.
            Dim spacer As New System.Windows.Forms.Panel() With {.Height = spacerExtra, .Width = 1, .Margin = New System.Windows.Forms.Padding(0)}
            table.Controls.Add(spacer, 0, 1)

            Dim bottomHost As New System.Windows.Forms.Panel() With {.AutoSize = True, .Margin = New System.Windows.Forms.Padding(0)}
            bottomHost.Padding = New System.Windows.Forms.Padding(0, gapAboveButtons, 0, 0)
            bottomHost.Controls.Add(bottomFlow)
            table.Controls.Add(bottomHost, 0, 2)

            messageForm.Controls.Clear()
            messageForm.Controls.Add(table)

            ' Final size: include spacerExtra.
            Dim clientW As System.Int32 = contentWidth + 2 * paddingAll
            Dim clientH As System.Int32 = bodyPanelHeight + spacerExtra + reservedBottomHeight + 2 * paddingAll
            clientW = System.Math.Min(clientW, maxWindowWidth)
            clientH = System.Math.Min(clientH, maxWindowHeight)
            messageForm.ClientSize = New System.Drawing.Size(clientW, clientH)

            SharedMethods.AttachForeignForegroundWatchdog(messageForm)

            If autoCloseSeconds.HasValue Then
                Dim remaining As System.Int32 = autoCloseSeconds.Value
                countdownLabel.Text = $"(closes in {remaining} seconds{Defaulttext})"
                Dim t As New System.Windows.Forms.Timer() With {.Interval = 1000}
                AddHandler t.Tick,
            Sub()
                remaining -= 1
                If remaining > 0 Then
                    countdownLabel.Text = $"(closes in {remaining} seconds{Defaulttext})"
                Else
                    t.Stop()
                    If Not userClicked Then
                        messageForm.Close()
                    End If
                End If
            End Sub
                t.Start()

                messageForm.Opacity = 1
                If SeparateThread Then
                    messageForm.BringToFront()
                    messageForm.Focus()
                    messageForm.Activate()

                    AddHandler messageForm.Shown,
                            Sub(sender, e)
                                SharedMethods.ForceDialogToForeground(messageForm)
                            End Sub
                    Dim __owner As System.Windows.Forms.IWin32Window = SharedMethods.ResolveSameThreadDialogOwner()
                    If __owner IsNot Nothing Then
                        messageForm.ShowDialog(__owner)
                    Else
                        messageForm.ShowDialog()
                    End If
                Else
                    messageForm.Show()
                    System.Windows.Forms.Application.DoEvents()
                End If
            Else

                messageForm.BringToFront()
                messageForm.Focus()
                messageForm.Activate()

                AddHandler messageForm.Shown,
                        Sub(sender, e)
                            SharedMethods.ForceDialogToForeground(messageForm)
                        End Sub

                messageForm.Opacity = 1
                Dim __owner As System.Windows.Forms.IWin32Window = SharedMethods.ResolveSameThreadDialogOwner()
                If __owner IsNot Nothing Then
                    messageForm.ShowDialog(__owner)
                Else
                    messageForm.ShowDialog()
                End If

            End If
        End Sub


        ''' <summary>
        ''' Shows a modal RichTextBox-based message dialog (RTF content) with optional auto-close.
        ''' Returns 1 for the primary button, 2 for the secondary button, 3 for auto-close, and 0 for close/Esc.
        ''' </summary>
        ''' <param name="bodyText">RTF content assigned to <see cref="RichTextBox.Rtf"/>.</param>
        ''' <param name="header">Dialog title. Defaults to <c>AN</c> if empty/whitespace.</param>
        ''' <param name="autoCloseSeconds">If set, closes the dialog after this many seconds.</param>
        ''' <param name="Defaulttext">Suffix appended to the countdown label text.</param>
        ''' <param name="RestoreWindow">If True, saves and restores window position and size from settings.</param>
        ''' <param name="okButtonText">Primary button text.</param>
        ''' <param name="secondaryButtonText">Optional secondary button text.</param>
        ''' <param name="okButtonAction">Optional callback invoked when the primary button is clicked.</param>
        ''' <param name="secondaryButtonAction">Optional callback invoked when the secondary button is clicked.</param>
        ''' <param name="CloseAfterOk">If True, closes after the primary button is clicked.</param>
        ''' <param name="CloseAfterSecondary">If True, closes after the secondary button is clicked.</param>
        Public Shared Function ShowRTFCustomMessageBox(
            ByVal bodyText As String,
            Optional header As String = AN,
            Optional autoCloseSeconds As Integer? = Nothing,
            Optional Defaulttext As String = " - execution continues meanwhile",
            Optional RestoreWindow As Boolean = False,
            Optional okButtonText As String = "OK",
            Optional secondaryButtonText As String = Nothing,
            Optional okButtonAction As System.Action = Nothing,
            Optional secondaryButtonAction As System.Action = Nothing,
            Optional CloseAfterOk As Boolean = True,
            Optional CloseAfterSecondary As Boolean = True
        ) As Integer

            Dim RTFMessageForm As New System.Windows.Forms.Form()
            Dim bodyLabel As New System.Windows.Forms.RichTextBox()
            Dim okButton As New System.Windows.Forms.Button()
            Dim secondaryButton As System.Windows.Forms.Button = Nothing
            Dim countdownLabel As New System.Windows.Forms.Label()
            Dim result As Integer = 0
            Dim userClicked As Boolean = False

            If String.IsNullOrWhiteSpace(header) Then header = AN

            RTFMessageForm.Opacity = 0
            RTFMessageForm.Text = header
            RTFMessageForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
            RTFMessageForm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            RTFMessageForm.MaximizeBox = True
            RTFMessageForm.MinimizeBox = True
            RTFMessageForm.ShowInTaskbar = False
            RTFMessageForm.TopMost = True
            RTFMessageForm.KeyPreview = True
            RTFMessageForm.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
            RTFMessageForm.AutoScaleDimensions = New System.Drawing.SizeF(96.0F, 96.0F)
            RTFMessageForm.MinimumSize = New System.Drawing.Size(650, 335)

            Dim bmp As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
            RTFMessageForm.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)

            bodyLabel.Font = standardFont
            bodyLabel.ReadOnly = True
            bodyLabel.BorderStyle = System.Windows.Forms.BorderStyle.None
            bodyLabel.BackColor = RTFMessageForm.BackColor
            bodyLabel.TabStop = True
            bodyLabel.HideSelection = False
            bodyLabel.ShortcutsEnabled = True
            bodyLabel.Rtf = bodyText
            bodyLabel.Location = New System.Drawing.Point(20, 20)
            bodyLabel.Width = 600
            bodyLabel.Height = 200
            bodyLabel.Anchor = System.Windows.Forms.AnchorStyles.Top _
                     Or System.Windows.Forms.AnchorStyles.Left _
                     Or System.Windows.Forms.AnchorStyles.Right _
                     Or System.Windows.Forms.AnchorStyles.Bottom
            RTFMessageForm.Controls.Add(bodyLabel)

            okButton.Font = standardFont
            okButton.Text = okButtonText
            okButton.AutoSize = True

            If Not String.IsNullOrWhiteSpace(secondaryButtonText) Then
                secondaryButton = New System.Windows.Forms.Button() With {
                    .Font = standardFont,
                    .Text = secondaryButtonText,
                    .AutoSize = True
                }
            End If

            countdownLabel.Font = standardFont
            countdownLabel.AutoSize = True

            Dim bottomPanel As New System.Windows.Forms.Panel()
            bottomPanel.Dock = System.Windows.Forms.DockStyle.Bottom
            bottomPanel.Padding = New System.Windows.Forms.Padding(20)
            bottomPanel.Height = okButton.PreferredSize.Height + bottomPanel.Padding.Top + bottomPanel.Padding.Bottom
            RTFMessageForm.Controls.Add(bottomPanel)

            bottomPanel.Controls.Add(okButton)

            If secondaryButton IsNot Nothing Then
                bottomPanel.Controls.Add(secondaryButton)
            End If

            bottomPanel.Controls.Add(countdownLabel)

            Dim nextLeft As Integer = bottomPanel.Padding.Left
            okButton.Location = New System.Drawing.Point(nextLeft, bottomPanel.Padding.Top)
            nextLeft = okButton.Right + 10

            If secondaryButton IsNot Nothing Then
                secondaryButton.Location = New System.Drawing.Point(nextLeft, bottomPanel.Padding.Top)
                nextLeft = secondaryButton.Right + 10
            End If

            countdownLabel.Location = New System.Drawing.Point(nextLeft, bottomPanel.Padding.Top + 4)

            AddHandler RTFMessageForm.Resize,
                Sub(sender As Object, e As EventArgs)
                    Dim availableWidth As Integer = RTFMessageForm.ClientSize.Width - bodyLabel.Left - 20
                    Dim availableHeight As Integer = RTFMessageForm.ClientSize.Height - bottomPanel.Height - bodyLabel.Top - 20
                    bodyLabel.Size = New System.Drawing.Size(availableWidth, availableHeight)
                End Sub

            AddHandler okButton.Click,
                Sub(sender As Object, e As EventArgs)
                    result = 1
                    userClicked = True
                    Try
                        If okButtonAction IsNot Nothing Then okButtonAction()
                    Catch
                    End Try
                    If CloseAfterOk Then
                        RTFMessageForm.Close()
                    End If
                End Sub

            If secondaryButton IsNot Nothing Then
                AddHandler secondaryButton.Click,
                    Sub(sender As Object, e As EventArgs)
                        result = 2
                        userClicked = True
                        Try
                            If secondaryButtonAction IsNot Nothing Then secondaryButtonAction()
                        Catch
                        End Try
                        If CloseAfterSecondary Then
                            RTFMessageForm.Close()
                        End If
                    End Sub
            End If

            AddHandler RTFMessageForm.KeyDown,
                Sub(sender As Object, e As System.Windows.Forms.KeyEventArgs)
                    If e.KeyCode = System.Windows.Forms.Keys.Escape Then
                        result = 0
                        userClicked = True
                        RTFMessageForm.Close()
                        e.SuppressKeyPress = True
                    End If
                End Sub

            AddHandler RTFMessageForm.Shown,
                Sub(sender As Object, e As EventArgs)
                    RTFMessageForm.PerformLayout()
                    RTFMessageForm.Activate()
                    bodyLabel.Focus()
                End Sub

            Dim formWidth As Integer = Math.Max(RTFMessageForm.MinimumSize.Width, bodyLabel.Width + 40)
            Dim formHeight As Integer = Math.Max(RTFMessageForm.MinimumSize.Height, bodyLabel.Bottom + 20 + bottomPanel.Height)
            RTFMessageForm.ClientSize = New System.Drawing.Size(formWidth, formHeight)

            If RestoreWindow Then
                Try
                    Dim savedBounds As Rectangle = My.Settings.RTFMessageBoxBounds
                    If savedBounds <> Rectangle.Empty AndAlso
                       savedBounds.Width >= RTFMessageForm.MinimumSize.Width AndAlso
                       savedBounds.Height >= RTFMessageForm.MinimumSize.Height Then

                        Dim isOnScreen As Boolean = False
                        For Each scr As Screen In Screen.AllScreens
                            If scr.WorkingArea.IntersectsWith(savedBounds) Then
                                isOnScreen = True
                                Exit For
                            End If
                        Next

                        If isOnScreen Then
                            RTFMessageForm.StartPosition = FormStartPosition.Manual
                            RTFMessageForm.Bounds = savedBounds
                        End If
                    End If
                Catch
                End Try

                AddHandler RTFMessageForm.FormClosing,
                    Sub(sender As Object, e As FormClosingEventArgs)
                        Try
                            If RTFMessageForm.WindowState = FormWindowState.Normal Then
                                My.Settings.RTFMessageBoxBounds = RTFMessageForm.Bounds
                                My.Settings.Save()
                            End If
                        Catch
                        End Try
                    End Sub
            End If

            SharedMethods.AttachForeignForegroundWatchdog(RTFMessageForm)

            If autoCloseSeconds.HasValue AndAlso autoCloseSeconds > 0 Then
                Dim remainingTime As Integer = autoCloseSeconds.Value
                countdownLabel.Text = $"(closes in {remainingTime} seconds{Defaulttext})"

                Dim timer As New System.Windows.Forms.Timer()
                timer.Interval = 1000

                AddHandler timer.Tick,
                    Sub(sender As Object, e As EventArgs)
                        remainingTime -= 1
                        If remainingTime > 0 Then
                            countdownLabel.Text = $"(closes in {remainingTime} seconds{Defaulttext})"
                        Else
                            timer.Stop()
                            If Not userClicked Then
                                result = 3
                                RTFMessageForm.Close()
                            End If
                        End If
                    End Sub

                timer.Start()

                RTFMessageForm.BringToFront()
                RTFMessageForm.Focus()
                RTFMessageForm.Activate()

                AddHandler RTFMessageForm.Shown,
                    Sub(sender, e)
                        SharedMethods.ForceDialogToForeground(RTFMessageForm)
                    End Sub

                RTFMessageForm.Opacity = 1
                RTFMessageForm.Show()
                RTFMessageForm.BringToFront()
                RTFMessageForm.Activate()
                System.Windows.Forms.Application.DoEvents()

                Return result
            Else
                RTFMessageForm.BringToFront()
                RTFMessageForm.Focus()
                RTFMessageForm.Activate()

                AddHandler RTFMessageForm.Shown,
                    Sub(sender, e)
                        SharedMethods.ForceDialogToForeground(RTFMessageForm)
                    End Sub

                RTFMessageForm.Opacity = 1
                Dim __owner As System.Windows.Forms.IWin32Window = SharedMethods.ResolveSameThreadDialogOwner()
                If __owner IsNot Nothing Then
                    RTFMessageForm.ShowDialog(__owner)
                Else
                    RTFMessageForm.ShowDialog()
                End If
                Return result
            End If

        End Function


        ''' <summary>
        ''' Shows an HTML message dialog using a WinForms WebBrowser control on an STA thread.
        ''' </summary>
        ''' <param name="bodyText">HTML assigned to DocumentText.</param>
        ''' <param name="header">Dialog title.</param>
        ''' <param name="Defaulttext">Unused parameter (kept for signature compatibility).</param>
        ''' <param name="extraButtonText">Optional extra button text.</param>
        ''' <param name="extraButtonAction">Action invoked when the extra button is clicked.</param>
        ''' <param name="CloseAfterExtra">If True, closes the dialog after invoking the extra action.</param>
        ''' <param name="additionalButtons">Optional array of additional buttons (Text, Action, CloseAfter).</param>
        ''' <param name="nonModal">If True, shows non-modally with topmost behavior.</param>
        ''' <param name="onClose">Optional action invoked when the dialog is closed.</param>
        Public Shared Sub ShowHTMLCustomMessageBox(
            ByVal bodyText As String,
            Optional header As String = AN,
            Optional Defaulttext As String = " - execution continues meanwhile",
            Optional extraButtonText As String = Nothing,
            Optional extraButtonAction As System.Action = Nothing,
            Optional CloseAfterExtra As Boolean = False,
            Optional additionalButtons As System.Tuple(Of System.String, System.Action, System.Boolean)() = Nothing,
            Optional nonModal As Boolean = False,
            Optional onClose As System.Action = Nothing
        )
            ' For non-modal on the current thread
            If nonModal Then
                ShowHTMLCustomMessageBoxNonModal(bodyText, header, extraButtonText, extraButtonAction, CloseAfterExtra, additionalButtons, onClose)
                Return
            End If

            Dim t As New Thread(Sub()
                                    ShowHTMLCustomMessageBoxInternal(bodyText, header, extraButtonText, extraButtonAction, CloseAfterExtra, additionalButtons, onClose)
                                End Sub)
            t.SetApartmentState(System.Threading.ApartmentState.STA)
            t.Start()
        End Sub

        Private Shared Sub ShowHTMLCustomMessageBoxInternal(
            ByVal bodyText As String,
            header As String,
            extraButtonText As String,
            extraButtonAction As System.Action,
            CloseAfterExtra As Boolean,
            additionalButtons As System.Tuple(Of System.String, System.Action, System.Boolean)(),
            onClose As System.Action
        )
            ' Create and configure form
            Dim HTMLMessageForm As New System.Windows.Forms.Form() With {
                .Opacity = 0,
                .Text = If(String.IsNullOrWhiteSpace(header), AN, header),
                .FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable,
                .StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen,
                .MaximizeBox = True,
                .MinimizeBox = True,
                .ShowInTaskbar = True,
                .TopMost = False,
                .KeyPreview = True,
                .AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            }

            ' Set the icon
            Try
                Dim bmp As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                HTMLMessageForm.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())
            Catch
            End Try

            ' Standard font
            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
            HTMLMessageForm.Font = standardFont

            ' WebBrowser
            Dim htmlBrowser As New System.Windows.Forms.WebBrowser() With {
                .AllowNavigation = False,
                .WebBrowserShortcutsEnabled = True,
                .IsWebBrowserContextMenuEnabled = True,
                .ScrollBarsEnabled = True,
                .ScriptErrorsSuppressed = True,
                .DocumentText = bodyText,
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .BackColor = HTMLMessageForm.BackColor,
                .Margin = New System.Windows.Forms.Padding(20)
            }
            AddHandler htmlBrowser.DocumentCompleted, Sub(sender2, e2)
                                                          If htmlBrowser.Document?.Body IsNot Nothing Then
                                                              htmlBrowser.Document.Body.Style =
                                                                  $"background-color: rgb({HTMLMessageForm.BackColor.R}, {HTMLMessageForm.BackColor.G}, {HTMLMessageForm.BackColor.B}); " &
                                                                  "font-family: 'Segoe UI'; font-size: 9pt; margin: 20px;"
                                                          End If
                                                      End Sub

            ' OK button
            Dim okButton As New System.Windows.Forms.Button() With {
                .Text = "OK",
                .AutoSize = True,
                .Font = standardFont,
                .Margin = New System.Windows.Forms.Padding(0)
            }
            AddHandler okButton.Click, Sub()
                                           HTMLMessageForm.Close()
                                       End Sub

            ' Form-level Escape
            AddHandler HTMLMessageForm.KeyDown, Sub(sender2, e2)
                                                    If e2.KeyCode = System.Windows.Forms.Keys.Escape Then
                                                        HTMLMessageForm.Close()
                                                        e2.SuppressKeyPress = True
                                                    End If
                                                End Sub

            ' Bottom flow panel
            Dim bottomFlow As New System.Windows.Forms.FlowLayoutPanel() With {
                .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
                .Dock = System.Windows.Forms.DockStyle.Bottom,
                .AutoSize = True,
                .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                .Padding = New System.Windows.Forms.Padding(20),
                .WrapContents = False
            }
            bottomFlow.Controls.Add(okButton)

            ' First extra button (legacy parameter)
            If (Not System.String.IsNullOrEmpty(extraButtonText)) AndAlso (extraButtonAction IsNot Nothing) Then
                Dim extraButton As New System.Windows.Forms.Button() With {
                    .Text = extraButtonText,
                    .AutoSize = True,
                    .Font = standardFont,
                    .Margin = New System.Windows.Forms.Padding(10, okButton.Margin.Top, 0, okButton.Margin.Bottom)
                }
                AddHandler extraButton.Click,
                    Sub()
                        Try
                            extraButtonAction.Invoke()
                        Catch ex As System.Exception
                        End Try
                        If CloseAfterExtra Then HTMLMessageForm.Close()
                    End Sub
                bottomFlow.Controls.Add(extraButton)
            End If

            ' Additional buttons array
            If additionalButtons IsNot Nothing AndAlso additionalButtons.Length > 0 Then
                For Each btnDef In additionalButtons
                    If System.String.IsNullOrEmpty(btnDef.Item1) OrElse btnDef.Item2 Is Nothing Then Continue For
                    Dim addBtn As New System.Windows.Forms.Button() With {
                        .Text = btnDef.Item1,
                        .AutoSize = True,
                        .Font = standardFont,
                        .Margin = New System.Windows.Forms.Padding(10, okButton.Margin.Top, 0, okButton.Margin.Bottom)
                    }
                    Dim closeAfter As Boolean = btnDef.Item3
                    Dim action As System.Action = btnDef.Item2
                    AddHandler addBtn.Click,
                        Sub()
                            Try
                                action.Invoke()
                            Catch ex As System.Exception
                            End Try
                            If closeAfter Then HTMLMessageForm.Close()
                        End Sub
                    bottomFlow.Controls.Add(addBtn)
                Next
            End If

            ' Calculate minimum width to fit all buttons in one row
            bottomFlow.PerformLayout()
            Dim totalButtonWidth As Integer = 0
            For Each ctrl As Control In bottomFlow.Controls
                totalButtonWidth += ctrl.PreferredSize.Width + ctrl.Margin.Left + ctrl.Margin.Right
            Next
            totalButtonWidth += bottomFlow.Padding.Left + bottomFlow.Padding.Right + 40

            Dim minFormWidth As Integer = Math.Max(600, totalButtonWidth)
            HTMLMessageForm.MinimumSize = New System.Drawing.Size(minFormWidth, 400)
            HTMLMessageForm.Size = New System.Drawing.Size(Math.Max(minFormWidth, 1000), 700)

            ' Compose form
            HTMLMessageForm.Controls.Add(htmlBrowser)
            HTMLMessageForm.Controls.Add(bottomFlow)

            ' onClose callback
            If onClose IsNot Nothing Then
                AddHandler HTMLMessageForm.FormClosed, Sub(sender, e)
                                                           Try
                                                               onClose.Invoke()
                                                           Catch
                                                           End Try
                                                       End Sub
            End If

            AddHandler HTMLMessageForm.Shown,
                Sub(sender, e)
                    SharedMethods.ForceDialogToForeground(HTMLMessageForm)
                End Sub

            HTMLMessageForm.Opacity = 1
            Dim __owner As System.Windows.Forms.IWin32Window = SharedMethods.ResolveSameThreadDialogOwner()
            If __owner IsNot Nothing Then
                HTMLMessageForm.ShowDialog(__owner)
            Else
                HTMLMessageForm.ShowDialog()
            End If

        End Sub

        Private Shared Sub ShowHTMLCustomMessageBoxNonModal(
            ByVal bodyText As String,
            header As String,
            extraButtonText As String,
            extraButtonAction As System.Action,
            CloseAfterExtra As Boolean,
            additionalButtons As System.Tuple(Of System.String, System.Action, System.Boolean)(),
            onClose As System.Action
        )
            Dim dialogClosed As New Threading.ManualResetEvent(False)

            ' Create and configure form
            Dim HTMLMessageForm As New System.Windows.Forms.Form() With {
                .Opacity = 0,
                .Text = If(String.IsNullOrWhiteSpace(header), AN, header),
                .FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable,
                .StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen,
                .MaximizeBox = True,
                .MinimizeBox = True,
                .ShowInTaskbar = True,
                .TopMost = True,
                .KeyPreview = True,
                .AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            }

            ' Set the icon
            Try
                Dim bmp As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                HTMLMessageForm.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())
            Catch
            End Try

            ' Standard font
            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
            HTMLMessageForm.Font = standardFont

            ' WebBrowser
            Dim htmlBrowser As New System.Windows.Forms.WebBrowser() With {
                .AllowNavigation = False,
                .WebBrowserShortcutsEnabled = True,
                .IsWebBrowserContextMenuEnabled = True,
                .ScrollBarsEnabled = True,
                .ScriptErrorsSuppressed = True,
                .DocumentText = bodyText,
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .BackColor = HTMLMessageForm.BackColor,
                .Margin = New System.Windows.Forms.Padding(20)
            }
            AddHandler htmlBrowser.DocumentCompleted, Sub(sender2, e2)
                                                          If htmlBrowser.Document?.Body IsNot Nothing Then
                                                              htmlBrowser.Document.Body.Style =
                                                                  $"background-color: rgb({HTMLMessageForm.BackColor.R}, {HTMLMessageForm.BackColor.G}, {HTMLMessageForm.BackColor.B}); " &
                                                                  "font-family: 'Segoe UI'; font-size: 9pt; margin: 20px;"
                                                          End If
                                                      End Sub

            ' OK button
            Dim okButton As New System.Windows.Forms.Button() With {
                .Text = "OK",
                .AutoSize = True,
                .Font = standardFont,
                .Margin = New System.Windows.Forms.Padding(0)
            }
            AddHandler okButton.Click, Sub()
                                           HTMLMessageForm.Close()
                                       End Sub

            ' Form-level Escape
            AddHandler HTMLMessageForm.KeyDown, Sub(sender2, e2)
                                                    If e2.KeyCode = System.Windows.Forms.Keys.Escape Then
                                                        HTMLMessageForm.Close()
                                                        e2.SuppressKeyPress = True
                                                    End If
                                                End Sub

            ' Bottom flow panel - no wrapping
            Dim bottomFlow As New System.Windows.Forms.FlowLayoutPanel() With {
                .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
                .Dock = System.Windows.Forms.DockStyle.Bottom,
                .AutoSize = True,
                .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                .Padding = New System.Windows.Forms.Padding(20),
                .WrapContents = False
            }
            bottomFlow.Controls.Add(okButton)

            ' First extra button
            If (Not String.IsNullOrEmpty(extraButtonText)) AndAlso (extraButtonAction IsNot Nothing) Then
                Dim extraButton As New System.Windows.Forms.Button() With {
                    .Text = extraButtonText,
                    .AutoSize = True,
                    .Font = standardFont,
                    .Margin = New System.Windows.Forms.Padding(10, okButton.Margin.Top, 0, okButton.Margin.Bottom)
                }
                AddHandler extraButton.Click,
                    Sub()
                        Try
                            extraButtonAction.Invoke()
                        Catch
                        End Try
                        If CloseAfterExtra Then HTMLMessageForm.Close()
                    End Sub
                bottomFlow.Controls.Add(extraButton)
            End If

            ' Additional buttons
            If additionalButtons IsNot Nothing AndAlso additionalButtons.Length > 0 Then
                For Each btnDef In additionalButtons
                    If String.IsNullOrEmpty(btnDef.Item1) OrElse btnDef.Item2 Is Nothing Then Continue For
                    Dim addBtn As New System.Windows.Forms.Button() With {
                        .Text = btnDef.Item1,
                        .AutoSize = True,
                        .Font = standardFont,
                        .Margin = New System.Windows.Forms.Padding(10, okButton.Margin.Top, 0, okButton.Margin.Bottom)
                    }
                    Dim closeAfter As Boolean = btnDef.Item3
                    Dim action As System.Action = btnDef.Item2
                    AddHandler addBtn.Click,
                        Sub()
                            Try
                                action.Invoke()
                            Catch
                            End Try
                            If closeAfter Then HTMLMessageForm.Close()
                        End Sub
                    bottomFlow.Controls.Add(addBtn)
                Next
            End If

            ' Calculate minimum width to fit all buttons
            bottomFlow.PerformLayout()
            Dim totalButtonWidth As Integer = 0
            For Each ctrl As Control In bottomFlow.Controls
                totalButtonWidth += ctrl.PreferredSize.Width + ctrl.Margin.Left + ctrl.Margin.Right
            Next
            totalButtonWidth += bottomFlow.Padding.Left + bottomFlow.Padding.Right + 40

            Dim minFormWidth As Integer = Math.Max(600, totalButtonWidth)
            HTMLMessageForm.MinimumSize = New System.Drawing.Size(minFormWidth, 400)
            HTMLMessageForm.Size = New System.Drawing.Size(Math.Max(minFormWidth, 1000), 700)

            ' Compose form
            HTMLMessageForm.Controls.Add(htmlBrowser)
            HTMLMessageForm.Controls.Add(bottomFlow)

            ' Keep topmost on deactivate.
            AddHandler HTMLMessageForm.Deactivate, Sub(sender, e)
                                                       Try
                                                           HTMLMessageForm.TopMost = True
                                                       Catch
                                                       End Try
                                                   End Sub

            SharedMethods.AttachForeignForegroundWatchdog(HTMLMessageForm)

            ' Signal when closed and invoke onClose
            AddHandler HTMLMessageForm.FormClosed, Sub(sender, e)
                                                       If onClose IsNot Nothing Then
                                                           Try
                                                               onClose.Invoke()
                                                           Catch
                                                           End Try
                                                       End If
                                                       dialogClosed.Set()
                                                   End Sub

            AddHandler HTMLMessageForm.Shown,
                Sub(sender, e)
                    SharedMethods.ForceDialogToForeground(HTMLMessageForm)
                End Sub

            HTMLMessageForm.Opacity = 1
            HTMLMessageForm.Show()
            HTMLMessageForm.BringToFront()
            HTMLMessageForm.Activate()

            ' Pump messages until dialog is closed
            While Not dialogClosed.WaitOne(50)
                System.Windows.Forms.Application.DoEvents()
            End While

            HTMLMessageForm.Dispose()
        End Sub
        ''' <summary>
        ''' Shows an HTML message dialog non-modally with topmost behavior, allowing interaction with other windows.
        ''' Blocks the calling thread until the dialog is closed using a message pump.
        ''' </summary>
        Private Shared Sub ShowHTMLCustomMessageBoxNonModal(
            ByVal bodyText As String,
            header As String,
            extraButtonText As String,
            extraButtonAction As System.Action,
            CloseAfterExtra As Boolean,
            additionalButtons As System.Tuple(Of System.String, System.Action, System.Boolean)()
        )
            Dim dialogClosed As New Threading.ManualResetEvent(False)

            ' Create and configure form.
            Dim HTMLMessageForm As New System.Windows.Forms.Form() With {
                .Opacity = 0,
                .Text = header,
                .FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable,
                .StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen,
                .MaximizeBox = True,
                .MinimizeBox = True,
                .ShowInTaskbar = True,
                .TopMost = True,
                .KeyPreview = True,
                .MinimumSize = New System.Drawing.Size(800, 500),
                .AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            }

            ' Header fallback.
            If String.IsNullOrWhiteSpace(header) Then
                HTMLMessageForm.Text = AN
            End If

            ' Set the icon.
            Try
                Dim bmp As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                HTMLMessageForm.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())
            Catch
            End Try

            ' Standard font.
            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
            HTMLMessageForm.Font = standardFont

            ' WebBrowser with margin - enable shortcuts for copy/paste.
            Dim htmlBrowser As New System.Windows.Forms.WebBrowser() With {
                .AllowNavigation = False,
                .WebBrowserShortcutsEnabled = True,
                .IsWebBrowserContextMenuEnabled = True,
                .ScrollBarsEnabled = True,
                .ScriptErrorsSuppressed = True,
                .DocumentText = bodyText,
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .BackColor = HTMLMessageForm.BackColor,
                .Margin = New System.Windows.Forms.Padding(20)
            }
            AddHandler htmlBrowser.DocumentCompleted, Sub(sender2, e2)
                                                          If htmlBrowser.Document?.Body IsNot Nothing Then
                                                              htmlBrowser.Document.Body.Style =
                                                                  $"background-color: rgb({HTMLMessageForm.BackColor.R}, {HTMLMessageForm.BackColor.G}, {HTMLMessageForm.BackColor.B}); " &
                                                                  "font-family: 'Segoe UI'; font-size: 9pt; margin: 20px;"
                                                          End If
                                                      End Sub

            ' OK button.
            Dim okButton As New System.Windows.Forms.Button() With {
                .Text = "OK",
                .AutoSize = True,
                .Font = standardFont,
                .Margin = New System.Windows.Forms.Padding(0)
            }
            AddHandler okButton.Click, Sub()
                                           HTMLMessageForm.Close()
                                       End Sub

            ' Form-level Escape.
            AddHandler HTMLMessageForm.KeyDown, Sub(sender2, e2)
                                                    If e2.KeyCode = System.Windows.Forms.Keys.Escape Then
                                                        HTMLMessageForm.Close()
                                                        e2.SuppressKeyPress = True
                                                    End If
                                                End Sub

            ' Bottom flow panel.
            Dim bottomFlow As New System.Windows.Forms.FlowLayoutPanel() With {
                .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
                .Dock = System.Windows.Forms.DockStyle.Bottom,
                .AutoSize = True,
                .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                .Padding = New System.Windows.Forms.Padding(20)
            }
            bottomFlow.Controls.Add(okButton)

            ' First extra button (legacy parameter).
            If (Not System.String.IsNullOrEmpty(extraButtonText)) AndAlso (extraButtonAction IsNot Nothing) Then
                Dim extraButton As New System.Windows.Forms.Button() With {
                    .Text = extraButtonText,
                    .AutoSize = True,
                    .Font = standardFont,
                    .Margin = New System.Windows.Forms.Padding(10, okButton.Margin.Top, 0, okButton.Margin.Bottom)
                }

                AddHandler extraButton.Click,
                    Sub()
                        Try
                            extraButtonAction.Invoke()
                        Catch ex As System.Exception
                        End Try
                        If CloseAfterExtra Then HTMLMessageForm.Close()
                    End Sub

                bottomFlow.Controls.Add(extraButton)
            End If

            ' Additional buttons array.
            If additionalButtons IsNot Nothing AndAlso additionalButtons.Length > 0 Then
                For Each btnDef In additionalButtons
                    If System.String.IsNullOrEmpty(btnDef.Item1) OrElse btnDef.Item2 Is Nothing Then Continue For
                    Dim addBtn As New System.Windows.Forms.Button() With {
                        .Text = btnDef.Item1,
                        .AutoSize = True,
                        .Font = standardFont,
                        .Margin = New System.Windows.Forms.Padding(10, okButton.Margin.Top, 0, okButton.Margin.Bottom)
                    }
                    Dim closeAfter As Boolean = btnDef.Item3
                    Dim action As System.Action = btnDef.Item2
                    AddHandler addBtn.Click,
                        Sub()
                            Try
                                action.Invoke()
                            Catch ex As System.Exception
                            End Try
                            If closeAfter Then HTMLMessageForm.Close()
                        End Sub
                    bottomFlow.Controls.Add(addBtn)
                Next
            End If

            ' Compose form.
            HTMLMessageForm.Controls.Add(htmlBrowser)
            HTMLMessageForm.Controls.Add(bottomFlow)

            ' Keep topmost on deactivate.
            AddHandler HTMLMessageForm.Deactivate, Sub(sender, e)
                                                       Try
                                                           HTMLMessageForm.TopMost = True
                                                       Catch
                                                       End Try
                                                   End Sub

            SharedMethods.AttachForeignForegroundWatchdog(HTMLMessageForm)

            ' Signal when closed
            AddHandler HTMLMessageForm.FormClosed, Sub(sender, e)
                                                       dialogClosed.Set()
                                                   End Sub

            AddHandler HTMLMessageForm.Shown,
                Sub(sender, e)
                    SharedMethods.ForceDialogToForeground(HTMLMessageForm)
                End Sub

            HTMLMessageForm.Opacity = 1

            ' Show non-modal
            HTMLMessageForm.Show()
            HTMLMessageForm.BringToFront()
            HTMLMessageForm.Activate()

            ' Pump messages until dialog is closed
            While Not dialogClosed.WaitOne(50)
                System.Windows.Forms.Application.DoEvents()
            End While

            HTMLMessageForm.Dispose()
        End Sub



        ''' <summary>
        ''' Shows a modal form that renders an array of input parameters as appropriate WinForms controls.
        ''' </summary>
        ''' <param name="prompt">Prompt text shown above the parameter list.</param>
        ''' <param name="header">Dialog title (empty when null/whitespace).</param>
        ''' <param name="params">Parameter array; each item is updated in-place when OK is pressed.</param>
        ''' <param name="extraButtonText">Optional extra button text.</param>
        ''' <param name="extraButtonAction">Action invoked when the extra button is clicked.</param>
        ''' <param name="CloseAfterExtra">
        ''' If <c>True</c>, closes the dialog after invoking the extra action and sets <see cref="DialogResult.Cancel"/>.
        ''' </param>
        ''' <returns><c>True</c> when OK is pressed; otherwise <c>False</c>.</returns>
        Public Shared Function ShowCustomVariableInputForm(
                                            ByVal prompt As String,
                                            ByVal header As String,
                                            ByRef params() As InputParameter,
                                            Optional extraButtonText As System.String = Nothing,
                                            Optional extraButtonAction As System.Action = Nothing,
                                            Optional CloseAfterExtra As System.Boolean = False
                                        ) As Boolean
            If String.IsNullOrWhiteSpace(header) Then header = String.Empty

            Dim inputForm As New Form() With {
                .Text = header,
                .FormBorderStyle = FormBorderStyle.FixedDialog,
                .StartPosition = FormStartPosition.CenterScreen,
                .MaximizeBox = False,
                .MinimizeBox = False,
                .Font = New System.Drawing.Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point),
                .AutoScaleMode = AutoScaleMode.Font,
                .AutoScaleDimensions = New SizeF(6.0F, 13.0F),
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .KeyPreview = True ' Allow form to see Ctrl+Enter before controls.
            }

            ' Set icon.
            Dim bmpIcon As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
            inputForm.Icon = Icon.FromHandle(bmpIcon.GetHicon())

            ' Layout.
            Dim mainLayout As New TableLayoutPanel() With {
                .ColumnCount = 2,
                .Dock = DockStyle.Fill,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Padding = New Padding(12),
                .GrowStyle = TableLayoutPanelGrowStyle.AddRows
            }
            mainLayout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            mainLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))

            ' Prompt label.
            Dim promptLabel As New System.Windows.Forms.Label() With {
                .Text = prompt,
                .AutoSize = True,
                .MaximumSize = New Size(600, 0),
                .Margin = New Padding(0, 0, 0, 12)
            }
            mainLayout.Controls.Add(promptLabel, 0, 0)
            mainLayout.SetColumnSpan(promptLabel, 2)

            ' Component container + tooltip.
            Dim components As New System.ComponentModel.Container()
            Dim toolTip As New System.Windows.Forms.ToolTip(components) With {
                .ShowAlways = True
            }

            For i As Integer = 0 To params.Length - 1
                Dim param = params(i)
                Dim rawValue As Object = param.Value

                Dim showLabel As Boolean = Not String.IsNullOrWhiteSpace(param.Name)
                Dim lbl As System.Windows.Forms.Label = Nothing

                If showLabel Then
                    lbl = New System.Windows.Forms.Label() With {
                        .Text = param.Name & ":",
                        .AutoSize = True,
                        .Anchor = AnchorStyles.Left,
                        .Margin = New Padding(0, 0, 8, 8)
                    }
                    mainLayout.Controls.Add(lbl, 0, i + 1)
                End If

                Dim ctrl As Control

                ' Rules:
                ' 1. If value Is Nothing -> show disabled CheckBox (unchecked).
                ' 2. If value Is Boolean -> show enabled CheckBox with that state.
                ' 3. Else if options exist -> ComboBox.
                ' 4. Else -> TextBox.
                Dim isNothing As Boolean = (rawValue Is Nothing)
                Dim isBool As Boolean = TypeOf rawValue Is Boolean

                Dim sentinelDisabled As String = "<<disabled>>"
                Dim disableForSentinel As Boolean =
                    (TypeOf rawValue Is String AndAlso
                     String.Equals(CStr(rawValue), sentinelDisabled, System.StringComparison.Ordinal))

                If disableForSentinel Then rawValue = ""

                If isNothing OrElse isBool Then
                    Dim initial As Boolean = If(isBool, CBool(rawValue), False)
                    Dim chk As New System.Windows.Forms.CheckBox() With {
                        .Checked = initial,
                        .AutoSize = True,
                        .Anchor = AnchorStyles.Left,
                        .Margin = New Padding(0, 0, 0, 8),
                        .Enabled = Not isNothing
                    }
                    If isNothing Then
                        chk.BackColor = SystemColors.Control
                        toolTip.SetToolTip(chk, "Not available")
                    End If
                    ctrl = chk

                ElseIf param.Options IsNot Nothing AndAlso param.Options.Count > 0 AndAlso TypeOf rawValue Is String Then
                    Dim cb As New System.Windows.Forms.ComboBox() With {
                        .DropDownStyle = ComboBoxStyle.DropDownList,
                        .MaxDropDownItems = 5,
                        .IntegralHeight = False,
                        .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
                        .Margin = New Padding(0, 0, 0, 12),
                        .MinimumSize = New Size(400, 0)
                    }
                    cb.Items.AddRange(param.Options.ToArray())
                    If param.Options.Contains(CStr(rawValue)) Then cb.SelectedItem = rawValue

                    ' Adjust dropdown width.
                    Dim maxItemWidth As Integer = 0
                    For Each it In cb.Items
                        Dim w = TextRenderer.MeasureText(CStr(it), cb.Font).Width
                        If w > maxItemWidth Then maxItemWidth = w
                    Next
                    Dim needsScroll = cb.Items.Count > cb.MaxDropDownItems
                    Dim scrollW = If(needsScroll, SystemInformation.VerticalScrollBarWidth, 0)
                    cb.DropDownWidth = Math.Max(cb.DropDownWidth, maxItemWidth + scrollW + 16)

                    ' Tooltip if truncated.
                    Dim updateToolTip As EventHandler =
                        Sub(sender As Object, eArgs As EventArgs)
                            Dim combo = DirectCast(sender, ComboBox)
                            Dim t = combo.Text
                            Dim tw = TextRenderer.MeasureText(t, combo.Font).Width
                            Dim usable = Math.Max(0, combo.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - 6)
                            If tw > usable Then
                                toolTip.SetToolTip(combo, t)
                            Else
                                toolTip.SetToolTip(combo, Nothing)
                            End If
                        End Sub
                    AddHandler cb.SelectedIndexChanged, updateToolTip
                    AddHandler cb.TextChanged, updateToolTip
                    AddHandler cb.Resize, updateToolTip
                    AddHandler cb.MouseEnter, updateToolTip
                    updateToolTip(cb, EventArgs.Empty)

                    ctrl = cb

                Else
                    Dim stringValue As String = rawValue.ToString()
                    Dim useMultiline As Boolean =
                        TypeOf rawValue Is String AndAlso
                        (param.Multiline OrElse
                         stringValue.Contains(vbCr) OrElse
                         stringValue.Contains(vbLf))

                    Dim txt As New TextBox() With {
                        .Text = stringValue,
                        .Anchor = If(useMultiline,
                                     AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right,
                                     AnchorStyles.Left Or AnchorStyles.Right),
                        .Margin = New Padding(0, 0, 0, 8)
                    }

                    If TypeOf rawValue Is String Then
                        If useMultiline Then
                            Dim multilineHeight As Integer = If(param.MultilineHeight > 0, param.MultilineHeight, 96)
                            txt.Multiline = True
                            txt.AcceptsReturn = True
                            txt.ScrollBars = ScrollBars.Vertical
                            txt.WordWrap = True
                            txt.MinimumSize = New Size(400, multilineHeight)
                            txt.Size = New Size(400, multilineHeight)
                            If lbl IsNot Nothing Then
                                lbl.Anchor = AnchorStyles.Left Or AnchorStyles.Top
                            End If
                        Else
                            txt.MinimumSize = New Size(400, 0)
                        End If
                    Else
                        txt.MinimumSize = New Size(50, 0)
                    End If

                    ctrl = txt
                End If

                If disableForSentinel Then
                    ctrl.Enabled = False
                    toolTip.SetToolTip(ctrl, "Not available")
                End If

                param.InputControl = ctrl
                If showLabel Then
                    mainLayout.Controls.Add(ctrl, 1, i + 1)
                Else
                    mainLayout.Controls.Add(ctrl, 0, i + 1)
                    mainLayout.SetColumnSpan(ctrl, 2)
                End If
            Next

            ' Buttons.
            Dim buttonFlow As New FlowLayoutPanel() With {
                .FlowDirection = FlowDirection.RightToLeft,
                .Dock = DockStyle.Bottom,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Padding = New Padding(12, 8, 12, 12)
            }
            Dim btnOK As New Button() With {.Text = "OK", .AutoSize = True, .DialogResult = DialogResult.OK}
            Dim btnCancel As New Button() With {.Text = "Cancel", .AutoSize = True, .DialogResult = DialogResult.Cancel}

            ' Add in this order so visual order is [OK][Cancel] with RightToLeft.
            buttonFlow.Controls.Add(btnCancel)
            buttonFlow.Controls.Add(btnOK)

            ' Ensure Tab order prefers OK when tabbing out of the last field.
            btnOK.TabIndex = 0
            btnCancel.TabIndex = 2 ' Will move to 1 if no extra button is added.

            ' Optional extra button: same behavior as ShowCustomMessageBox.
            Dim extraButton As System.Windows.Forms.Button = Nothing
            If (Not System.String.IsNullOrEmpty(extraButtonText)) AndAlso (extraButtonAction IsNot Nothing) Then
                extraButton = New System.Windows.Forms.Button() With {
                    .Text = extraButtonText,
                    .AutoSize = True,
                    .Margin = New System.Windows.Forms.Padding(8, btnOK.Margin.Top, 0, btnOK.Margin.Bottom)
                }
                AddHandler extraButton.Click,
                    Sub()
                        Try
                            extraButtonAction.Invoke()
                        Catch ex As System.Exception
                            ' Swallow to keep dialog functional; mirror ShowCustomMessageBox behavior.
                        End Try
                        If CloseAfterExtra Then
                            inputForm.DialogResult = DialogResult.Cancel ' Do not commit changes implicitly.
                            inputForm.Close()
                        End If
                    End Sub

                ' Place the extra button to the left of OK (RightToLeft flow).
                buttonFlow.Controls.Add(extraButton)

                ' Tab order: OK first, then extra, then Cancel.
                extraButton.TabIndex = 1
            Else
                ' No extra button: let Cancel be second.
                btnCancel.TabIndex = 1
            End If

            inputForm.Controls.Add(mainLayout)
            inputForm.Controls.Add(buttonFlow)

            ' Ctrl+Enter should trigger OK anywhere on the form.
            AddHandler inputForm.KeyDown,
                Sub(sender As Object, e As KeyEventArgs)
                    If e.KeyCode = Keys.Enter AndAlso e.Control Then
                        btnOK.PerformClick()
                        e.SuppressKeyPress = True
                        e.Handled = True
                    End If
                End Sub

            Dim result As DialogResult
            Dim __owner As System.Windows.Forms.IWin32Window = SharedMethods.ResolveSameThreadDialogOwner()
            If __owner IsNot Nothing Then
                result = inputForm.ShowDialog(__owner)
            Else
                result = inputForm.ShowDialog()
            End If

            If result = DialogResult.OK Then
                For Each param In params
                    ' Skip disabled controls: keep existing Value.
                    If param.InputControl IsNot Nothing AndAlso Not param.InputControl.Enabled Then
                        Continue For
                    End If
                    Try
                        If TypeOf param.InputControl Is System.Windows.Forms.ComboBox Then
                            Dim cb = DirectCast(param.InputControl, System.Windows.Forms.ComboBox)
                            param.Value = If(cb.SelectedItem IsNot Nothing, cb.SelectedItem.ToString(), cb.Text)
                        ElseIf TypeOf param.Value Is Boolean Then
                            param.Value = CType(param.InputControl, System.Windows.Forms.CheckBox).Checked
                        ElseIf TypeOf param.Value Is Integer Then
                            Dim valI As Integer
                            If Integer.TryParse(CType(param.InputControl, TextBox).Text, valI) Then
                                param.Value = valI
                            Else
                                Throw New Exception($"Invalid value for {param.Name}.")
                            End If
                        ElseIf TypeOf param.Value Is Double Then
                            Dim valD As Double
                            Dim inputText As String = CType(param.InputControl, TextBox).Text.Trim()

                            ' Normalize: replace comma with dot, then parse with invariant culture.
                            Dim normalizedInput As String = inputText.Replace(","c, "."c)

                            If Double.TryParse(normalizedInput, NumberStyles.Float, CultureInfo.InvariantCulture, valD) Then
                                param.Value = valD
                            Else
                                Throw New Exception($"Invalid value for {param.Name}.")
                            End If
                        Else
                            ' Generic / string.
                            If TypeOf param.InputControl Is TextBox Then
                                param.Value = CType(param.InputControl, TextBox).Text
                            End If
                        End If
                    Catch ex As Exception
                        ShowCustomMessageBox($"{ex.Message} Using original ('{If(param.Value Is Nothing, "Nothing", param.Value)}').")
                    End Try
                Next
            End If

            inputForm.Dispose()
            Return (result = DialogResult.OK)
        End Function

        ''' <summary>
        ''' Shows an editable dialog window with intro text, an editable RichTextBox (or plain text),
        ''' and multiple completion options (use edited/original text; optional special return modes).
        ''' </summary>
        ''' <param name="introLine">Intro line displayed above the editor.</param>
        ''' <param name="bodyText">Initial text content; converted to RTF unless <paramref name="NoRTF"/> is True.</param>
        ''' <param name="finalRemark">Optional remark text shown below the editor.</param>
        ''' <param name="header">Dialog title.</param>
        ''' <param name="NoRTF">If True, uses plain text; otherwise assigns RTF into the editor.</param>
        ''' <param name="Getfocus">If True and no parent handle is passed, attempts to parent to a detected Office window.</param>
        ''' <param name="InsertMarkdown">If True, adds a button that returns the sentinel value "Markdown".</param>
        ''' <param name="TransferToPane">If True, adds a button that returns the sentinel value "Pane".</param>
        ''' <param name="parentWindowHwnd">Optional explicit parent window handle for dialog ownership.</param>
        ''' <param name="PreserveLiterals">Passed through to Markdown-to-RTF conversion.</param>
        ''' <param name="ReturnPlainText">If True, returns plain text even when RTF editing is enabled; otherwise returns RTF when edited.</param>
        ''' <returns>
        ''' On OK buttons: returns edited text (RTF or plain) or original text (RTF or original input) as implemented.
        ''' On Cancel: returns <see cref="String.Empty"/>.
        ''' On special buttons: returns the sentinel strings "Markdown" or "Pane" as implemented.
        ''' </returns>
        Public Shared Function ShowCustomWindow(
    introLine As String,
    ByVal bodyText As String,
    finalRemark As String,
    header As String,
    Optional NoRTF As Boolean = False,
    Optional Getfocus As Boolean = False,
    Optional InsertMarkdown As Boolean = False,
    Optional TransferToPane As Boolean = False,
    Optional parentWindowHwnd As IntPtr = Nothing,
    Optional PreserveLiterals As Boolean = False,
    Optional ReturnPlainText As Boolean = False
                ) As String

            ' Store original body text.
            Dim OriginalText As String = bodyText

            ' Spacing & constants.
            Const leftMargin As Integer = 10
            Const rightPadding As Integer = 10
            Const spacing As Integer = 10
            Const gapButtons As Integer = 10
            Const remarkToButtonSpacing As Integer = 20
            Const bottomPadding As Integer = 20

            ' Create controls.
            Dim styledForm As New System.Windows.Forms.Form()
            Dim introLabel As New System.Windows.Forms.Label()
            Dim bodyTextBox As New RichTextBox()
            Dim finalRemarkLabel As New System.Windows.Forms.Label()
            Dim btnEdited As New System.Windows.Forms.Button()
            Dim btnOriginal As New System.Windows.Forms.Button()
            Dim btnMark As New System.Windows.Forms.Button()
            Dim btnPane As New System.Windows.Forms.Button()
            Dim btnCancel As New System.Windows.Forms.Button()
            Dim toolStrip As New System.Windows.Forms.ToolStrip()
            Dim lblHint As New System.Windows.Forms.Label() With {
        .AutoSize = False,
        .TextAlign = ContentAlignment.MiddleRight
    }

            ' Screen / max size calculation.
            Dim scrW = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea.Width
            Dim scrH = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea.Height
            Dim maxW = scrW \ 2
            Dim maxH = Math.Min(scrH \ 2, (maxW * 9) \ 16)
            maxW = Math.Min(maxW, (maxH * 16) \ 9)

            ' Fallback minima.
            Const minFormWStatic As Integer = 400
            Const minFormHStatic As Integer = 300

            ' Form properties.
            styledForm.Text = header
            styledForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
            styledForm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            styledForm.MaximizeBox = True
            styledForm.MinimizeBox = False
            styledForm.ShowInTaskbar = False
            styledForm.TopMost = True
            styledForm.CancelButton = btnCancel
            styledForm.MinimumSize = New System.Drawing.Size(minFormWStatic, minFormHStatic)

            ' Icon.
            Dim bmp As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
            styledForm.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

            ' Standard font.
            Dim stdFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
            styledForm.Font = stdFont

            ' Intro label.
            introLabel.Text = introLine
            introLabel.Font = stdFont
            introLabel.AutoSize = False
            introLabel.Location = New System.Drawing.Point(leftMargin, spacing)
            introLabel.Width = maxW - leftMargin - rightPadding
            introLabel.Height = introLabel.PreferredHeight
            introLabel.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            styledForm.Controls.Add(introLabel)

            ' Buttons.
            btnEdited.Text = "OK, use edited text"
            Dim szE = TextRenderer.MeasureText(btnEdited.Text, stdFont)
            btnEdited.Size = New Size(szE.Width + 20, szE.Height + 10)

            btnOriginal.Text = "OK, use original text"
            Dim szO = TextRenderer.MeasureText(btnOriginal.Text, stdFont)
            btnOriginal.Size = New Size(szO.Width + 20, szE.Height + 10)

            If TransferToPane Then
                btnPane.Text = "Transfer to pane"
                Dim szP = TextRenderer.MeasureText(btnPane.Text, stdFont)
                btnPane.Size = New Size(szP.Width + 20, szE.Height + 10)
                styledForm.Controls.Add(btnPane)
            End If

            If InsertMarkdown Then
                btnMark.Text = "Insert original text with formatting"
                Dim szM = TextRenderer.MeasureText(btnMark.Text, stdFont)
                btnMark.Size = New Size(szM.Width + 20, szE.Height + 10)
                styledForm.Controls.Add(btnMark)
            End If

            btnCancel.Text = "Cancel"
            Dim szC = TextRenderer.MeasureText(btnCancel.Text, stdFont)
            btnCancel.Size = New Size(szC.Width + 20, szE.Height + 10)

            styledForm.Controls.Add(btnEdited)
            styledForm.Controls.Add(btnOriginal)
            styledForm.Controls.Add(btnCancel)

            ' BodyTextBox (align with CustomPaneControl).
            bodyTextBox.Font = New System.Drawing.Font("Segoe UI", 10.0F, FontStyle.Regular, GraphicsUnit.Point)
            bodyTextBox.Multiline = True
            bodyTextBox.ScrollBars = RichTextBoxScrollBars.Vertical
            bodyTextBox.WordWrap = True
            bodyTextBox.HideSelection = False
            bodyTextBox.DetectUrls = True
            bodyTextBox.Location = New System.Drawing.Point(leftMargin, introLabel.Bottom + spacing)
            bodyTextBox.Width = maxW - leftMargin - rightPadding
            bodyTextBox.Height = maxH - introLabel.Bottom - spacing
            bodyTextBox.MinimumSize = New Size(bodyTextBox.Width, bodyTextBox.Height)
            bodyTextBox.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            styledForm.Controls.Add(bodyTextBox)

            ' LinkClicked: open directly (no Ctrl modifier), like CustomPaneControl.
            AddHandler bodyTextBox.LinkClicked,
        Sub(senderObj As Object, e As LinkClickedEventArgs)
            Try
                System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(e.LinkText) With {.UseShellExecute = True})
            Catch
                ' Ignore.
            End Try
        End Sub

            ' Copy handler: match CustomPaneControl behavior.
            AddHandler bodyTextBox.KeyDown,
        Sub(sender As Object, e As System.Windows.Forms.KeyEventArgs)
            If (e.Control AndAlso (e.KeyCode = Keys.C OrElse e.KeyCode = Keys.Insert)) Then
                Try
                    If Not NoRTF Then
                        SharedMethods.CopySelectionExcludingTrailingNbsp(bodyTextBox)
                    Else
                        If bodyTextBox.SelectionLength > 0 Then
                            SharedMethods.PutInClipboard(bodyTextBox.SelectedText)
                        Else
                            SharedMethods.PutInClipboard(bodyTextBox.Text)
                        End If
                    End If
                    e.Handled = True
                Catch
                    ' Fallback to default if anything goes wrong.
                End Try
            End If
            ' Do not intercept Ctrl+A (same as CustomPaneControl).
        End Sub

            ' Optional final remark label.
            Dim hasRemark = Not String.IsNullOrEmpty(finalRemark)
            If hasRemark Then
                finalRemarkLabel.Text = finalRemark
                finalRemarkLabel.Font = stdFont
                finalRemarkLabel.AutoSize = False
                finalRemarkLabel.Width = bodyTextBox.MinimumSize.Width
                finalRemarkLabel.Height = finalRemarkLabel.GetPreferredSize(New Size(finalRemarkLabel.Width, 0)).Height
                finalRemarkLabel.Anchor = AnchorStyles.Left Or AnchorStyles.Right
                styledForm.Controls.Add(finalRemarkLabel)
            End If

            ' ToolStrip.
            toolStrip.Dock = DockStyle.None
            For Each sym In New String() {"B", "I", "U", "•"}
                Dim tsb As New ToolStripButton(sym) With {
            .Font = New System.Drawing.Font(stdFont, If(sym = "B",
                                                FontStyle.Bold,
                                                If(sym = "I",
                                                   FontStyle.Italic,
                                                   If(sym = "U",
                                                      FontStyle.Underline,
                                                      FontStyle.Regular)))),
            .Name = "tsb" & sym
        }
                AddHandler tsb.Click,
            Sub(s, e)
                If bodyTextBox.SelectionLength > 0 Then
                    Select Case DirectCast(s, ToolStripButton).Name
                        Case "tsbB"
                            bodyTextBox.SelectionFont = New System.Drawing.Font(bodyTextBox.SelectionFont, bodyTextBox.SelectionFont.Style Xor FontStyle.Bold)
                        Case "tsbI"
                            bodyTextBox.SelectionFont = New System.Drawing.Font(bodyTextBox.SelectionFont, bodyTextBox.SelectionFont.Style Xor FontStyle.Italic)
                        Case "tsbU"
                            bodyTextBox.SelectionFont = New System.Drawing.Font(bodyTextBox.SelectionFont, bodyTextBox.SelectionFont.Style Xor FontStyle.Underline)
                        Case "tsb•"
                            bodyTextBox.SelectionIndent = If(bodyTextBox.SelectionIndent = 20, 0, 20)
                            bodyTextBox.SelectionBullet = Not bodyTextBox.SelectionBullet
                            bodyTextBox.BulletIndent = If(bodyTextBox.BulletIndent = 15, 0, 15)
                    End Select
                End If
            End Sub
                toolStrip.Items.Add(tsb)
            Next
            styledForm.Controls.Add(toolStrip)

            ' Hint label.
            lblHint.Text = "Click a link to open"
            lblHint.Font = New System.Drawing.Font(stdFont, FontStyle.Italic)
            lblHint.ForeColor = Color.DimGray
            lblHint.Height = szE.Height + 6
            styledForm.Controls.Add(lblHint)

            ' Dynamic MinimumSize.
            Dim bodyTop = bodyTextBox.Top
            Dim bodyMinH = bodyTextBox.MinimumSize.Height
            Dim remHeight = If(hasRemark,
               finalRemarkLabel.GetPreferredSize(New Size(bodyTextBox.MinimumSize.Width, 0)).Height,
               0)
            Dim btnH = btnEdited.Height

            Dim dynamicMinH = bodyTop +
              bodyMinH +
              If(hasRemark,
                 spacing + remHeight + remarkToButtonSpacing,
                 remarkToButtonSpacing) +
              btnH +
              bottomPadding

            Dim w1 = leftMargin + bodyTextBox.MinimumSize.Width + rightPadding
            Dim introMinW = leftMargin + introLabel.PreferredWidth + rightPadding
            Dim totalBtnW = btnEdited.Width + gapButtons + btnOriginal.Width +
            If(InsertMarkdown, gapButtons + btnMark.Width, 0) +
            If(TransferToPane, gapButtons + btnPane.Width, 0) +
            gapButtons + btnCancel.Width
            Dim w3 = leftMargin + totalBtnW + rightPadding
            Dim dynamicMinW = Math.Max(Math.Max(w1, introMinW), w3)

            styledForm.MinimumSize = New Size(
        Math.Max(minFormWStatic, dynamicMinW),
        Math.Max(minFormHStatic, dynamicMinH)
    )

            ' Resize handler.
            AddHandler styledForm.Resize,
        Sub(s, e)
            Dim fW = styledForm.ClientSize.Width
            Dim fH = styledForm.ClientSize.Height

            introLabel.Width = fW - leftMargin - rightPadding

            Dim newW = fW - leftMargin - rightPadding
            bodyTextBox.Width = Math.Max(bodyTextBox.MinimumSize.Width, newW)

            Dim usedBelow = If(hasRemark,
                               spacing + finalRemarkLabel.Height + remarkToButtonSpacing,
                               remarkToButtonSpacing) +
                            btnH + bottomPadding
            Dim availH = fH - bodyTop - usedBelow
            bodyTextBox.Height = Math.Max(bodyTextBox.MinimumSize.Height, availH)

            If hasRemark Then
                finalRemarkLabel.Width = bodyTextBox.Width
                finalRemarkLabel.Height = finalRemarkLabel.GetPreferredSize(New Size(finalRemarkLabel.Width, 0)).Height
                finalRemarkLabel.Location = New System.Drawing.Point(leftMargin, bodyTextBox.Bottom + spacing)
            End If

            Dim btnY = fH - btnH - bottomPadding
            btnEdited.Location = New System.Drawing.Point(leftMargin, btnY)
            btnOriginal.Location = New System.Drawing.Point(btnEdited.Right + gapButtons, btnY)

            Dim nextX = btnOriginal.Right
            If InsertMarkdown Then
                btnMark.Location = New System.Drawing.Point(btnOriginal.Right + gapButtons, btnY)
                nextX = btnMark.Right
            End If
            If TransferToPane Then
                btnPane.Location = New System.Drawing.Point(nextX + gapButtons, btnY)
                nextX = btnPane.Right
            End If
            btnCancel.Location = New System.Drawing.Point(nextX + gapButtons, btnY)

            ' Toolstrip above textbox right aligned.
            toolStrip.Location = New System.Drawing.Point(
                leftMargin + bodyTextBox.Width - toolStrip.Width,
                bodyTextBox.Top - toolStrip.Height - spacing
            )
            toolStrip.BringToFront()

            ' Hint label aligns with right edge above buttons.
            lblHint.Width = 180
            lblHint.Location = New System.Drawing.Point(fW - lblHint.Width - rightPadding, introLabel.Top)
        End Sub

            ' Initial size.
            Dim initW = Math.Max(maxW, styledForm.MinimumSize.Width)
            Dim initH = Math.Max(maxH, styledForm.MinimumSize.Height)
            styledForm.ClientSize = New Size(initW, initH)
            styledForm.PerformLayout()
            styledForm.MinimumSize = styledForm.Size

            ' Content assignment (match CustomPaneControl).
            Dim rtf As String = Nothing
            If Not NoRTF Then
                rtf = MarkdownToRtfConverter.Convert(bodyText, PreserveLiterals)
                Debug.WriteLine("Converted RTF: " & rtf)
            End If

            Try
                If NoRTF Then
                    bodyTextBox.Text = bodyText
                Else
                    bodyTextBox.Rtf = rtf
                    ' Append NBSPs for hyperlinks (same as CustomPaneControl).
                    SharedMethods.AppendNbspForHyperlinks(bodyTextBox, rtf)
                End If
            Catch ex As System.ComponentModel.Win32Exception
                bodyTextBox.Text = bodyText
            Catch
                bodyTextBox.Text = bodyText
            End Try

            ' Ensure URL detection is enabled (same as CustomPaneControl).
            bodyTextBox.DetectUrls = True
            bodyTextBox.Select(0, 0)

            Dim OriginalTextBox As String = bodyTextBox.Text

            ' Button handlers.
            Dim returnValue As String = String.Empty

            AddHandler btnEdited.Click,
                    Sub()
                        If ReturnPlainText Then
                            returnValue = bodyTextBox.Text
                        Else
                            returnValue = If(NoRTF, bodyTextBox.Text, bodyTextBox.Rtf)
                        End If
                        styledForm.DialogResult = DialogResult.OK
                        styledForm.Close()
                    End Sub

            AddHandler btnOriginal.Click,
                    Sub()
                        If ReturnPlainText Then
                            returnValue = OriginalText
                        Else
                            returnValue = If(NoRTF, OriginalText, If(rtf, bodyText))
                        End If
                        styledForm.DialogResult = DialogResult.OK
                        styledForm.Close()
                    End Sub

            If InsertMarkdown Then
                AddHandler btnMark.Click,
            Sub()
                returnValue = "Markdown"
                styledForm.DialogResult = DialogResult.OK
                styledForm.Close()
            End Sub
            End If

            If TransferToPane Then
                AddHandler btnPane.Click,
            Sub()
                If bodyTextBox.Text.Trim() = OriginalTextBox.Trim() OrElse
                   ShowCustomYesNoBox($"Your changes will be lost and the pane will again show the original text (unless you put it in the clipboard manually). Continue?", "Yes", "No") = 1 Then
                    returnValue = "Pane"
                    styledForm.DialogResult = DialogResult.OK
                    styledForm.Close()
                End If
            End Sub
            End If

            AddHandler btnCancel.Click,
        Sub()
            returnValue = String.Empty
            styledForm.DialogResult = DialogResult.Cancel
            styledForm.Close()
        End Sub

            ' Show dialog.
            SharedMethods.AttachForeignForegroundWatchdog(styledForm)

            styledForm.BringToFront()
            styledForm.Focus()
            styledForm.Activate()

            AddHandler styledForm.Shown,
                    Sub(sender, e)
                        SharedMethods.ForceDialogToForeground(styledForm)
                    End Sub

            If parentWindowHwnd <> IntPtr.Zero Then
                Dim __ownerHwnd As System.Windows.Forms.IWin32Window =
                    SharedMethods.IfOwnerOnCurrentThread(New WindowWrapper(parentWindowHwnd))
                If __ownerHwnd IsNot Nothing Then
                    styledForm.ShowDialog(__ownerHwnd)
                Else
                    styledForm.ShowDialog()
                End If
            ElseIf Getfocus Then
                Dim officeHwnd As IntPtr = GetOfficeApplicationHwnd()
                Dim __ownerOffice As System.Windows.Forms.IWin32Window = Nothing
                If officeHwnd <> IntPtr.Zero Then
                    __ownerOffice = SharedMethods.IfOwnerOnCurrentThread(New WindowWrapper(officeHwnd))
                End If
                If __ownerOffice IsNot Nothing Then
                    styledForm.ShowDialog(__ownerOffice)
                Else
                    Dim __owner As System.Windows.Forms.IWin32Window = SharedMethods.ResolveSameThreadDialogOwner()
                    If __owner IsNot Nothing Then
                        styledForm.ShowDialog(__owner)
                    Else
                        styledForm.ShowDialog()
                    End If
                End If
            Else
                Dim __owner As System.Windows.Forms.IWin32Window = SharedMethods.ResolveSameThreadDialogOwner()
                If __owner IsNot Nothing Then
                    styledForm.ShowDialog(__owner)
                Else
                    styledForm.ShowDialog()
                End If
            End If

            Return returnValue
        End Function


        ''' <summary>
        ''' Represents a single input parameter for <see cref="ShowCustomVariableInputForm"/>,
        ''' including the UI control created to edit the parameter.
        ''' </summary>
        Public Class InputParameter
            ''' <summary>
            ''' Display name used for the label.
            ''' </summary>
            Public Property Name As System.String

            ''' <summary>
            ''' Current value. Its runtime type determines which control is created and how values are parsed back.
            ''' </summary>
            Public Property Value As System.Object

            ''' <summary>
            ''' Optional list of allowed values (used for a ComboBox when <see cref="Value"/> is a string).
            ''' </summary>
            Public Property Options As System.Collections.Generic.List(Of System.String)

            ''' <summary>
            ''' The WinForms control created for this parameter during dialog generation.
            ''' </summary>
            Public Property InputControl As System.Windows.Forms.Control

            ''' <summary>
            ''' When True, string values are rendered in a multiline TextBox.
            ''' </summary>
            Public Property Multiline As System.Boolean

            ''' <summary>
            ''' Preferred height for multiline text boxes.
            ''' </summary>
            Public Property MultilineHeight As System.Int32

            ' Important: parameterless constructor (required for "New InputParameter() With {...}").
            Public Sub New()
                Me.Options = New System.Collections.Generic.List(Of System.String)()
                Me.Multiline = False
                Me.MultilineHeight = 96
            End Sub

            ''' <summary>
            ''' Creates an <see cref="InputParameter"/> with a name and initial value.
            ''' </summary>
            ''' <param name="name">Display name.</param>
            ''' <param name="value">Initial value.</param>
            Public Sub New(ByVal name As System.String, ByVal value As System.Object)
                Me.New()
                Me.Name = name
                Me.Value = value
            End Sub

            ''' <summary>
            ''' Creates an <see cref="InputParameter"/> with a name, initial value, and a list of selectable options.
            ''' </summary>
            ''' <param name="name">Display name.</param>
            ''' <param name="value">Initial value.</param>
            ''' <param name="options">Selectable options.</param>
            Public Sub New(ByVal name As System.String,
                   ByVal value As System.Object,
                   ByVal options As System.Collections.Generic.IEnumerable(Of System.String))
                Me.New()
                Me.Name = name
                Me.Value = value
                If options IsNot Nothing Then
                    Me.Options = New System.Collections.Generic.List(Of System.String)(options)
                End If
            End Sub
        End Class

#Region "Freestyle Prompt Form"

        Public Class FreestylePromptMode

            Public Property Id As System.String
            Public Property Text As System.String
            Public Property Description As System.String
            Public Property Prefix As System.String
            Public Property Prefixes As System.Collections.Generic.List(Of System.String)
            Public Property ManualSyntax As System.String
            Public Property IsDefault As System.Boolean
            Public Property IsAvailable As System.Boolean
            Public Property UnavailableReason As System.String

            Public Sub New()
                Me.Id = System.String.Empty
                Me.Text = System.String.Empty
                Me.Description = System.String.Empty
                Me.Prefix = System.String.Empty
                Me.Prefixes = New System.Collections.Generic.List(Of System.String)()
                Me.ManualSyntax = System.String.Empty
                Me.IsDefault = False
                Me.IsAvailable = True
                Me.UnavailableReason = System.String.Empty
            End Sub

            Public Overrides Function ToString() As System.String

                Dim result As System.String = If(Me.Text, System.String.Empty)

                If Not System.String.IsNullOrWhiteSpace(Me.ManualSyntax) Then
                    result &= "   [" & Me.ManualSyntax & "]"
                End If

                If Not Me.IsAvailable Then
                    result &= "   (not available now)"
                End If

                Return result

            End Function

        End Class


        Public Class FreestylePromptInsertOption

            Public Property Id As System.String
            Public Property Text As System.String
            Public Property Description As System.String
            Public Property InsertText As System.String
            Public Property MaxOccurrences As System.Int32

            Public Property RequiresValue As System.Boolean
            Public Property ValuePrompt As System.String
            Public Property ValueTitle As System.String
            Public Property ValueTemplate As System.String
            Public Property ValuePlaceholder As System.String

            Public Sub New()
                Me.Id = System.String.Empty
                Me.Text = System.String.Empty
                Me.Description = System.String.Empty
                Me.InsertText = System.String.Empty
                Me.MaxOccurrences = 0
                Me.RequiresValue = False
                Me.ValuePrompt = System.String.Empty
                Me.ValueTitle = System.String.Empty
                Me.ValueTemplate = System.String.Empty
                Me.ValuePlaceholder = "[value]"
            End Sub

            Public Function BuildInsertText(ByVal value As System.String) As System.String

                If Not Me.RequiresValue Then
                    Return If(Me.InsertText, System.String.Empty)
                End If

                Dim cleanValue As System.String = If(value, System.String.Empty).Trim()

                If cleanValue.Length = 0 Then
                    Return System.String.Empty
                End If

                If System.String.IsNullOrWhiteSpace(Me.ValueTemplate) Then
                    Return cleanValue
                End If

                Dim placeholder As System.String = If(System.String.IsNullOrWhiteSpace(Me.ValuePlaceholder), "[value]", Me.ValuePlaceholder)

                Return Me.ValueTemplate.Replace(placeholder, cleanValue)

            End Function

        End Class


        Public Class FreestylePromptQuickButton

            Public Property Id As System.String
            Public Property Text As System.String
            Public Property Description As System.String
            Public Property Prefix As System.String

            Public Sub New()
                Me.Id = System.String.Empty
                Me.Text = System.String.Empty
                Me.Description = System.String.Empty
                Me.Prefix = System.String.Empty
            End Sub

        End Class


        Public Class FreestylePromptToggleOption

            Public Property Id As System.String
            Public Property Text As System.String
            Public Property Description As System.String
            Public Property Trigger As System.String
            Public Property ManualSyntax As System.String
            Public Property IsChecked As System.Boolean

            Public Property ArgumentPrefix As System.String
            Public Property ArgumentSuffix As System.String
            Public Property ArgumentTemplate As System.String
            Public Property ArgumentPlaceholder As System.String
            Public Property ArgumentHint As System.String
            Public Property ArgumentRequired As System.Boolean

            Public Sub New()
                Me.Id = System.String.Empty
                Me.Text = System.String.Empty
                Me.Description = System.String.Empty
                Me.Trigger = System.String.Empty
                Me.ManualSyntax = System.String.Empty
                Me.IsChecked = False
                Me.ArgumentPrefix = System.String.Empty
                Me.ArgumentSuffix = System.String.Empty
                Me.ArgumentTemplate = System.String.Empty
                Me.ArgumentPlaceholder = "[value]"
                Me.ArgumentHint = System.String.Empty
                Me.ArgumentRequired = False
            End Sub

            Public ReadOnly Property HasArgument As System.Boolean
                Get
                    Return Not System.String.IsNullOrWhiteSpace(Me.ArgumentTemplate) OrElse Not System.String.IsNullOrWhiteSpace(Me.ArgumentPrefix)
                End Get
            End Property

            Public Function BuildTrigger(ByVal argument As System.String) As System.String

                Dim cleanArgument As System.String = If(argument, System.String.Empty).Trim()

                If cleanArgument.Length = 0 Then
                    Return If(Me.Trigger, System.String.Empty)
                End If

                If Not System.String.IsNullOrWhiteSpace(Me.ArgumentTemplate) Then

                    Dim placeholder As System.String = If(System.String.IsNullOrWhiteSpace(Me.ArgumentPlaceholder), "[value]", Me.ArgumentPlaceholder)

                    Return Me.ArgumentTemplate.Replace(placeholder, cleanArgument)

                End If

                Return If(Me.ArgumentPrefix, System.String.Empty) & cleanArgument & If(Me.ArgumentSuffix, System.String.Empty)

            End Function

        End Class


        Public Class FreestylePromptSection

            Public Property Id As System.String
            Public Property Caption As System.String
            Public Property Options As System.Collections.Generic.List(Of FreestylePromptToggleOption)

            Public Sub New()
                Me.Id = System.String.Empty
                Me.Caption = System.String.Empty
                Me.Options = New System.Collections.Generic.List(Of FreestylePromptToggleOption)()
            End Sub

        End Class


        Public Class FreestylePromptOptions

            Public Property Title As System.String
            Public Property Heading As System.String
            Public Property ModeCaption As System.String
            Public Property ModelText As System.String
            Public Property ContextStatusText As System.String

            Public Property InitialPrompt As System.String
            Public Property LastPrompt As System.String

            Public Property PromptLibraryEnabled As System.Boolean
            Public Property ShowShortCommandsHint As System.Boolean
            Public Property Context As ISharedContext

            Public Property CallerId As System.String
            Public Property PersistedState As System.String
            Public Property RestorePersistedState As System.Boolean

            Public Property Modes As System.Collections.Generic.List(Of FreestylePromptMode)
            Public Property QuickButtons As System.Collections.Generic.List(Of FreestylePromptQuickButton)
            Public Property InsertOptions As System.Collections.Generic.List(Of FreestylePromptInsertOption)
            Public Property Sections As System.Collections.Generic.List(Of FreestylePromptSection)

            Public Sub New()
                Me.Title = "Freestyle"
                Me.Heading = "What would you like Red Ink to do?"
                Me.ModeCaption = "Output"
                Me.ModelText = System.String.Empty
                Me.ContextStatusText = System.String.Empty
                Me.InitialPrompt = System.String.Empty
                Me.LastPrompt = System.String.Empty
                Me.PromptLibraryEnabled = False
                Me.ShowShortCommandsHint = False
                Me.Context = Nothing
                Me.CallerId = System.String.Empty
                Me.PersistedState = System.String.Empty
                Me.RestorePersistedState = True
                Me.Modes = New System.Collections.Generic.List(Of FreestylePromptMode)()
                Me.QuickButtons = New System.Collections.Generic.List(Of FreestylePromptQuickButton)()
                Me.InsertOptions = New System.Collections.Generic.List(Of FreestylePromptInsertOption)()
                Me.Sections = New System.Collections.Generic.List(Of FreestylePromptSection)()
            End Sub

        End Class


        Public Class FreestylePromptResult

            Public Property Accepted As System.Boolean
            Public Property Prompt As System.String
            Public Property SelectedModeId As System.String
            Public Property SelectedPrefix As System.String
            Public Property SelectedOptionIds As System.Collections.Generic.List(Of System.String)
            Public Property SelectedTriggers As System.Collections.Generic.List(Of System.String)
            Public Property KnownPrefixes As System.Collections.Generic.List(Of System.String)
            Public Property PersistedState As System.String

            Public Sub New()
                Me.Accepted = False
                Me.Prompt = System.String.Empty
                Me.SelectedModeId = System.String.Empty
                Me.SelectedPrefix = System.String.Empty
                Me.SelectedOptionIds = New System.Collections.Generic.List(Of System.String)()
                Me.SelectedTriggers = New System.Collections.Generic.List(Of System.String)()
                Me.KnownPrefixes = New System.Collections.Generic.List(Of System.String)()
                Me.PersistedState = System.String.Empty
            End Sub

        End Class


        Private Shared Function FindFreestylePromptMode(ByVal text As System.String, ByVal modes As System.Collections.Generic.IEnumerable(Of FreestylePromptMode)) As System.Tuple(Of FreestylePromptMode, System.String)

            Dim source As System.String = If(text, System.String.Empty).TrimStart()
            Dim bestMode As FreestylePromptMode = Nothing
            Dim bestPrefix As System.String = System.String.Empty

            For Each mode As FreestylePromptMode In modes

                If mode Is Nothing OrElse mode.Prefixes Is Nothing Then
                    Continue For
                End If

                For Each prefix As System.String In mode.Prefixes

                    If System.String.IsNullOrWhiteSpace(prefix) Then
                        Continue For
                    End If

                    If source.StartsWith(prefix, System.StringComparison.OrdinalIgnoreCase) AndAlso prefix.Length > bestPrefix.Length Then
                        bestMode = mode
                        bestPrefix = prefix
                    End If

                Next

            Next

            If bestMode Is Nothing Then
                Return Nothing
            End If

            Return System.Tuple.Create(bestMode, bestPrefix)

        End Function


        Private Shared Function GetFreestyleLeadingColonToken(ByVal text As System.String) As System.String

            Dim source As System.String = If(text, System.String.Empty).TrimStart()

            If source.Length = 0 Then
                Return System.String.Empty
            End If

            Dim endIndex As System.Int32 = source.Length

            Dim firstSpace As System.Int32 = source.IndexOf(" "c)
            If firstSpace >= 0 Then endIndex = System.Math.Min(endIndex, firstSpace)

            Dim firstTab As System.Int32 = source.IndexOf(Microsoft.VisualBasic.ControlChars.Tab)
            If firstTab >= 0 Then endIndex = System.Math.Min(endIndex, firstTab)

            Dim firstCr As System.Int32 = source.IndexOf(Microsoft.VisualBasic.ControlChars.Cr)
            If firstCr >= 0 Then endIndex = System.Math.Min(endIndex, firstCr)

            Dim firstLf As System.Int32 = source.IndexOf(Microsoft.VisualBasic.ControlChars.Lf)
            If firstLf >= 0 Then endIndex = System.Math.Min(endIndex, firstLf)

            If endIndex <= 0 Then
                Return System.String.Empty
            End If

            Dim token As System.String = source.Substring(0, endIndex)

            If token.EndsWith(":", System.StringComparison.Ordinal) Then
                Return token
            End If

            Return System.String.Empty

        End Function


        Private Shared Function GetCustomInputInsertMaximum(
            ByVal limits As System.Collections.Generic.IDictionary(Of System.String, System.Int32),
            ByVal token As System.String) As System.Int32

            If limits Is Nothing OrElse System.String.IsNullOrWhiteSpace(token) Then
                Return 0
            End If

            For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, System.Int32) In limits
                If System.String.Equals(pair.Key, token, System.StringComparison.OrdinalIgnoreCase) Then
                    Return System.Math.Max(0, pair.Value)
                End If
            Next

            Return 0

        End Function


        Private Shared Function ValidateCustomInputInsertOccurrenceLimits(
            ByVal prompt As System.String,
            ByVal limits As System.Collections.Generic.IDictionary(Of System.String, System.Int32),
            ByRef violatingToken As System.String,
            ByRef maximumOccurrences As System.Int32) As System.Boolean

            violatingToken = System.String.Empty
            maximumOccurrences = 0

            If limits Is Nothing Then
                Return True
            End If

            For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, System.Int32) In limits
                If pair.Value <= 0 OrElse System.String.IsNullOrWhiteSpace(pair.Key) Then
                    Continue For
                End If

                If SharedMethods.CountFreestyleTokenOccurrences(If(prompt, System.String.Empty), pair.Key) > pair.Value Then
                    violatingToken = pair.Key
                    maximumOccurrences = pair.Value
                    Return False
                End If
            Next

            Return True

        End Function


        Private Shared Function CountFreestyleTokenOccurrences(ByVal source As System.String, ByVal token As System.String) As System.Int32

            If System.String.IsNullOrEmpty(source) OrElse System.String.IsNullOrEmpty(token) Then
                Return 0
            End If

            Dim count As System.Int32 = 0
            Dim searchIndex As System.Int32 = 0

            Do
                Dim foundIndex As System.Int32 = source.IndexOf(token, searchIndex, System.StringComparison.OrdinalIgnoreCase)
                If foundIndex < 0 Then
                    Exit Do
                End If

                count += 1
                searchIndex = foundIndex + token.Length
            Loop While searchIndex < source.Length

            Return count

        End Function


        Private Shared Function ValidateFreestyleInsertOccurrenceLimits(ByVal prompt As System.String, ByVal insertOptions As System.Collections.Generic.IEnumerable(Of FreestylePromptInsertOption), ByRef violatingOption As FreestylePromptInsertOption) As System.Boolean

            violatingOption = Nothing

            If insertOptions Is Nothing Then
                Return True
            End If

            For Each definition As FreestylePromptInsertOption In insertOptions
                If definition Is Nothing OrElse definition.MaxOccurrences <= 0 OrElse System.String.IsNullOrWhiteSpace(definition.InsertText) Then
                    Continue For
                End If

                If CountFreestyleTokenOccurrences(If(prompt, System.String.Empty), definition.InsertText) > definition.MaxOccurrences Then
                    violatingOption = definition
                    Return False
                End If
            Next

            Return True

        End Function


        Private Shared Sub InsertFreestyleTextAtCaret(ByVal textBox As System.Windows.Forms.TextBox, ByVal textToInsert As System.String)

            If textBox Is Nothing OrElse System.String.IsNullOrEmpty(textToInsert) Then
                Return
            End If

            Dim selectionStart As System.Int32 = textBox.SelectionStart

            textBox.Text = textBox.Text.Insert(selectionStart, textToInsert)
            textBox.SelectionStart = selectionStart + textToInsert.Length
            textBox.Focus()

        End Sub


        Private Shared Function FindFreestylePromptModeById(ByVal modeId As System.String, ByVal modes As System.Collections.Generic.IEnumerable(Of FreestylePromptMode)) As FreestylePromptMode

            If System.String.IsNullOrWhiteSpace(modeId) OrElse modes Is Nothing Then
                Return Nothing
            End If

            For Each mode As FreestylePromptMode In modes

                If mode Is Nothing Then
                    Continue For
                End If

                If System.String.Equals(mode.Id, modeId, System.StringComparison.OrdinalIgnoreCase) Then
                    Return mode
                End If

            Next

            Return Nothing

        End Function


        Private Shared Function FindFreestylePromptModeByPrefix(ByVal prefix As System.String, ByVal modes As System.Collections.Generic.IEnumerable(Of FreestylePromptMode)) As FreestylePromptMode

            If System.String.IsNullOrWhiteSpace(prefix) OrElse modes Is Nothing Then
                Return Nothing
            End If

            Dim targetPrefix As System.String = prefix.Trim()

            For Each mode As FreestylePromptMode In modes

                If mode Is Nothing Then
                    Continue For
                End If

                If Not System.String.IsNullOrWhiteSpace(mode.Prefix) AndAlso
                   System.String.Equals(mode.Prefix.Trim(), targetPrefix, System.StringComparison.OrdinalIgnoreCase) Then
                    Return mode
                End If

                If mode.Prefixes Is Nothing Then
                    Continue For
                End If

                For Each modePrefix As System.String In mode.Prefixes

                    If System.String.IsNullOrWhiteSpace(modePrefix) Then
                        Continue For
                    End If

                    If System.String.Equals(modePrefix.Trim(), targetPrefix, System.StringComparison.OrdinalIgnoreCase) Then
                        Return mode
                    End If

                Next

            Next

            Return Nothing

        End Function


        Private Shared Function GetFreestylePromptKnownPrefixes(ByVal modes As System.Collections.Generic.IEnumerable(Of FreestylePromptMode)) As System.Collections.Generic.List(Of System.String)

            Dim prefixes As New System.Collections.Generic.List(Of System.String)()

            If modes Is Nothing Then
                Return prefixes
            End If

            For Each mode As FreestylePromptMode In modes

                If mode Is Nothing Then
                    Continue For
                End If

                If Not System.String.IsNullOrWhiteSpace(mode.Prefix) Then
                    Dim singlePrefix As System.String = mode.Prefix.Trim()
                    If Not prefixes.Exists(Function(item As System.String) System.String.Equals(item, singlePrefix, System.StringComparison.OrdinalIgnoreCase)) Then
                        prefixes.Add(singlePrefix)
                    End If
                End If

                If mode.Prefixes Is Nothing Then
                    Continue For
                End If

                For Each prefix As System.String In mode.Prefixes

                    If System.String.IsNullOrWhiteSpace(prefix) Then
                        Continue For
                    End If

                    Dim cleanedPrefix As System.String = prefix.Trim()

                    If Not prefixes.Exists(Function(item As System.String) System.String.Equals(item, cleanedPrefix, System.StringComparison.OrdinalIgnoreCase)) Then
                        prefixes.Add(cleanedPrefix)
                    End If

                Next

            Next

            prefixes.Sort(Function(left As System.String, right As System.String) right.Length.CompareTo(left.Length))

            Return prefixes

        End Function


        Private Shared Function GetFreestylePromptStateDocument(ByVal persistedState As System.String) As System.Xml.Linq.XDocument

            Try

                If Not System.String.IsNullOrWhiteSpace(persistedState) Then

                    Dim document As System.Xml.Linq.XDocument = System.Xml.Linq.XDocument.Parse(persistedState)

                    If document.Root IsNot Nothing AndAlso
                       System.String.Equals(document.Root.Name.LocalName, "freestylePromptState", System.StringComparison.OrdinalIgnoreCase) Then
                        Return document
                    End If

                End If

            Catch
            End Try

            Return New System.Xml.Linq.XDocument(
                New System.Xml.Linq.XElement(
                    "freestylePromptState",
                    New System.Xml.Linq.XAttribute("version", "1")))

        End Function


        Private Shared Function GetFreestylePromptCallerStateElement(ByVal document As System.Xml.Linq.XDocument, ByVal callerId As System.String, ByVal createIfMissing As System.Boolean) As System.Xml.Linq.XElement

            If document Is Nothing OrElse document.Root Is Nothing OrElse System.String.IsNullOrWhiteSpace(callerId) Then
                Return Nothing
            End If

            For Each callerElement As System.Xml.Linq.XElement In document.Root.Elements("caller")

                Dim idAttribute As System.Xml.Linq.XAttribute = callerElement.Attribute("id")

                If idAttribute IsNot Nothing AndAlso
                   System.String.Equals(idAttribute.Value, callerId, System.StringComparison.OrdinalIgnoreCase) Then
                    Return callerElement
                End If

            Next

            If Not createIfMissing Then
                Return Nothing
            End If

            Dim newCallerElement As New System.Xml.Linq.XElement(
                "caller",
                New System.Xml.Linq.XAttribute("id", callerId))

            document.Root.Add(newCallerElement)

            Return newCallerElement

        End Function


        Private Shared Function GetFreestylePromptStateAttribute(ByVal element As System.Xml.Linq.XElement, ByVal attributeName As System.String) As System.String

            If element Is Nothing OrElse System.String.IsNullOrWhiteSpace(attributeName) Then
                Return System.String.Empty
            End If

            Dim attribute As System.Xml.Linq.XAttribute = element.Attribute(attributeName)

            If attribute Is Nothing Then
                Return System.String.Empty
            End If

            Return If(attribute.Value, System.String.Empty)

        End Function


        Private Shared Function BuildFreestylePromptPersistedState(
            ByVal callerId As System.String,
            ByVal existingState As System.String,
            ByVal selectedModeId As System.String,
            ByVal selectedPrefix As System.String,
            ByVal expanded As System.Boolean,
            ByVal toggleCheckBoxes As System.Collections.Generic.IEnumerable(Of System.Windows.Forms.CheckBox),
            ByVal argumentTextBoxes As System.Collections.Generic.IDictionary(Of System.Windows.Forms.CheckBox, System.Windows.Forms.TextBox)) As System.String

            If System.String.IsNullOrWhiteSpace(callerId) Then
                Return If(existingState, System.String.Empty)
            End If

            Dim document As System.Xml.Linq.XDocument = SharedMethods.GetFreestylePromptStateDocument(existingState)
            Dim callerElement As System.Xml.Linq.XElement = SharedMethods.GetFreestylePromptCallerStateElement(document, callerId, True)

            If callerElement Is Nothing Then
                Return If(existingState, System.String.Empty)
            End If

            callerElement.RemoveNodes()
            callerElement.RemoveAttributes()

            callerElement.SetAttributeValue("id", callerId)
            callerElement.SetAttributeValue("modeId", If(selectedModeId, System.String.Empty))
            callerElement.SetAttributeValue("prefix", If(selectedPrefix, System.String.Empty))
            callerElement.SetAttributeValue("expanded", If(expanded, "1", "0"))

            If toggleCheckBoxes IsNot Nothing Then

                For Each optionCheckBox As System.Windows.Forms.CheckBox In toggleCheckBoxes

                    If optionCheckBox Is Nothing Then
                        Continue For
                    End If

                    Dim definition As FreestylePromptToggleOption = TryCast(optionCheckBox.Tag, FreestylePromptToggleOption)

                    If definition Is Nothing OrElse System.String.IsNullOrWhiteSpace(definition.Id) Then
                        Continue For
                    End If

                    Dim optionElement As New System.Xml.Linq.XElement("option")

                    optionElement.SetAttributeValue("id", definition.Id)
                    optionElement.SetAttributeValue("checked", If(optionCheckBox.Checked, "1", "0"))

                    If argumentTextBoxes IsNot Nothing AndAlso argumentTextBoxes.ContainsKey(optionCheckBox) Then
                        Dim argumentValue As System.String = argumentTextBoxes(optionCheckBox).Text.Trim()
                        If argumentValue.Length > 0 Then
                            optionElement.SetAttributeValue("argument", argumentValue)
                        End If
                    End If

                    callerElement.Add(optionElement)

                Next

            End If

            Return document.ToString(System.Xml.Linq.SaveOptions.DisableFormatting)

        End Function


        Private Shared Function ReplaceFreestyleLeadingPrefix(ByVal text As System.String, ByVal knownPrefixes As System.Collections.Generic.IEnumerable(Of System.String), ByVal replacementPrefix As System.String) As System.String

            Dim prompt As System.String = If(text, System.String.Empty).Trim()
            Dim prefixToApply As System.String = If(replacementPrefix, System.String.Empty).Trim()
            Dim prefixToRemove As System.String = System.String.Empty

            If knownPrefixes IsNot Nothing Then

                For Each knownPrefix As System.String In knownPrefixes

                    If System.String.IsNullOrWhiteSpace(knownPrefix) Then
                        Continue For
                    End If

                    Dim cleanedPrefix As System.String = knownPrefix.Trim()

                    If prompt.StartsWith(cleanedPrefix, System.StringComparison.OrdinalIgnoreCase) Then
                        prefixToRemove = cleanedPrefix
                        Exit For
                    End If

                Next

            End If

            If prefixToRemove.Length = 0 Then
                prefixToRemove = SharedMethods.GetFreestyleLeadingColonToken(prompt)
            End If

            If prefixToRemove.Length > 0 AndAlso prompt.StartsWith(prefixToRemove, System.StringComparison.OrdinalIgnoreCase) Then
                prompt = prompt.Substring(prefixToRemove.Length).TrimStart()
            End If

            If prefixToApply.Length = 0 Then
                Return prompt.Trim()
            End If

            Return prefixToApply & If(prompt.Length > 0, " " & prompt.Trim(), System.String.Empty)

        End Function


        Private Shared Function TryGetFreestyleToggleTriggerMatch(
            ByVal text As System.String,
            ByVal definition As FreestylePromptToggleOption,
            ByRef matchedTrigger As System.String,
            ByRef argument As System.String) As System.Boolean

            matchedTrigger = System.String.Empty
            argument = System.String.Empty

            If definition Is Nothing Then
                Return False
            End If

            Dim source As System.String = If(text, System.String.Empty)

            If definition.HasArgument Then

                Dim argumentPrefix As System.String = System.String.Empty
                Dim argumentSuffix As System.String = System.String.Empty

                If Not System.String.IsNullOrWhiteSpace(definition.ArgumentTemplate) Then

                    Dim template As System.String = definition.ArgumentTemplate
                    Dim placeholder As System.String = If(System.String.IsNullOrWhiteSpace(definition.ArgumentPlaceholder), "[value]", definition.ArgumentPlaceholder)
                    Dim placeholderIndex As System.Int32 = template.IndexOf(placeholder, System.StringComparison.Ordinal)

                    If placeholderIndex >= 0 Then
                        argumentPrefix = template.Substring(0, placeholderIndex)
                        argumentSuffix = template.Substring(placeholderIndex + placeholder.Length)
                    Else
                        argumentPrefix = template
                    End If

                Else

                    argumentPrefix = If(definition.ArgumentPrefix, System.String.Empty)
                    argumentSuffix = If(definition.ArgumentSuffix, System.String.Empty)

                End If

                If argumentPrefix.Length > 0 Then

                    Dim searchIndex As System.Int32 = 0

                    Do

                        Dim prefixIndex As System.Int32 = source.IndexOf(argumentPrefix, searchIndex, System.StringComparison.OrdinalIgnoreCase)

                        If prefixIndex < 0 Then
                            Exit Do
                        End If

                        Dim valueStart As System.Int32 = prefixIndex + argumentPrefix.Length
                        Dim valueEnd As System.Int32 = -1
                        Dim triggerEnd As System.Int32 = -1

                        If argumentSuffix.Length > 0 Then

                            valueEnd = source.IndexOf(argumentSuffix, valueStart, System.StringComparison.OrdinalIgnoreCase)

                            If valueEnd >= 0 Then
                                triggerEnd = valueEnd + argumentSuffix.Length
                            End If

                        Else

                            valueEnd = valueStart

                            While valueEnd < source.Length AndAlso Not System.Char.IsWhiteSpace(source(valueEnd))
                                valueEnd += 1
                            End While

                            triggerEnd = valueEnd

                        End If

                        If valueEnd >= valueStart AndAlso triggerEnd >= valueStart Then

                            Dim extractedArgument As System.String = source.Substring(valueStart, valueEnd - valueStart)

                            If extractedArgument.Length > 0 OrElse Not definition.ArgumentRequired Then
                                matchedTrigger = source.Substring(prefixIndex, triggerEnd - prefixIndex)
                                argument = extractedArgument
                                Return True
                            End If

                        End If

                        searchIndex = prefixIndex + System.Math.Max(1, argumentPrefix.Length)

                    Loop While searchIndex < source.Length

                End If

            End If

            Dim fixedTrigger As System.String = If(definition.Trigger, System.String.Empty).Trim()

            If fixedTrigger.Length > 0 Then

                Dim fixedIndex As System.Int32 = source.IndexOf(fixedTrigger, System.StringComparison.OrdinalIgnoreCase)

                If fixedIndex >= 0 Then
                    matchedTrigger = source.Substring(fixedIndex, fixedTrigger.Length)
                    Return True
                End If

            End If

            Return False

        End Function


        Private Shared Function RemoveFreestyleToggleTriggers(ByVal text As System.String, ByVal definition As FreestylePromptToggleOption) As System.String

            Dim result As System.String = If(text, System.String.Empty)

            If definition Is Nothing Then
                Return result
            End If

            Do

                Dim matchedTrigger As System.String = System.String.Empty
                Dim argument As System.String = System.String.Empty

                If Not SharedMethods.TryGetFreestyleToggleTriggerMatch(result, definition, matchedTrigger, argument) Then
                    Exit Do
                End If

                If matchedTrigger.Length = 0 Then
                    Exit Do
                End If

                Dim matchIndex As System.Int32 = result.IndexOf(matchedTrigger, System.StringComparison.OrdinalIgnoreCase)

                If matchIndex < 0 Then
                    Exit Do
                End If

                Dim removeStart As System.Int32 = matchIndex
                Dim removeLength As System.Int32 = matchedTrigger.Length
                Dim characterBefore As System.Int32 = removeStart - 1
                Dim characterAfter As System.Int32 = removeStart + removeLength

                If characterBefore >= 0 AndAlso characterAfter < result.Length AndAlso
                   result(characterBefore) = " "c AndAlso result(characterAfter) = " "c Then

                    removeLength += 1

                ElseIf removeStart = 0 AndAlso characterAfter < result.Length AndAlso result(characterAfter) = " "c Then

                    removeLength += 1

                ElseIf characterAfter = result.Length AndAlso characterBefore >= 0 AndAlso result(characterBefore) = " "c Then

                    removeStart -= 1
                    removeLength += 1

                End If

                result = result.Remove(removeStart, removeLength)

            Loop

            Return result

        End Function


        Private Shared Function AppendFreestyleToggleTrigger(ByVal text As System.String, ByVal trigger As System.String) As System.String

            Dim result As System.String = If(text, System.String.Empty)
            Dim cleanedTrigger As System.String = If(trigger, System.String.Empty).Trim()

            If cleanedTrigger.Length = 0 Then
                Return result
            End If

            If result.Length = 0 Then
                Return cleanedTrigger
            End If

            If System.Char.IsWhiteSpace(result(result.Length - 1)) Then
                Return result & cleanedTrigger
            End If

            Return result & " " & cleanedTrigger

        End Function


        Public Shared Function ShowFreestylePromptForm(ByVal options As FreestylePromptOptions) As FreestylePromptResult

            Dim returnValue As New FreestylePromptResult()

            If options Is Nothing Then
                Return returnValue
            End If

            Try

                Dim workingArea As System.Drawing.Rectangle = System.Windows.Forms.Screen.FromPoint(System.Windows.Forms.Cursor.Position).WorkingArea

                Using standardFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)

                    Using headingFont As New System.Drawing.Font(standardFont.FontFamily, standardFont.Size + 2.0F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)

                        Using sectionFont As New System.Drawing.Font(standardFont.FontFamily, standardFont.Size, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)

                            Using freestyleForm As New System.Windows.Forms.Form()

                                Dim lineHeight As System.Int32 = System.Windows.Forms.TextRenderer.MeasureText("Ag", standardFont).Height
                                Dim scaleFactor As System.Double = System.Math.Max(0.8R, lineHeight / 15.0R)

                                Dim outerPadding As System.Int32 = System.Math.Max(1, CInt(System.Math.Round(16.0R * scaleFactor)))
                                Dim smallGap As System.Int32 = System.Math.Max(1, CInt(System.Math.Round(5.0R * scaleFactor)))
                                Dim normalGap As System.Int32 = System.Math.Max(1, CInt(System.Math.Round(8.0R * scaleFactor)))
                                Dim largeGap As System.Int32 = System.Math.Max(1, CInt(System.Math.Round(13.0R * scaleFactor)))
                                Dim buttonPadX As System.Int32 = System.Math.Max(1, CInt(System.Math.Round(5.0R * scaleFactor)))
                                Dim buttonPadY As System.Int32 = System.Math.Max(1, CInt(System.Math.Round(2.0R * scaleFactor)))

                                Dim maximumClientWidth As System.Int32 = System.Math.Max(1, workingArea.Width - 60)
                                Dim maximumClientHeight As System.Int32 = System.Math.Max(1, workingArea.Height - 60)

                                Dim preferredWidth As System.Int32 = CInt(System.Math.Round(workingArea.Width * 0.312R))
                                preferredWidth = System.Math.Max(420, preferredWidth)
                                preferredWidth = System.Math.Min(preferredWidth, maximumClientWidth)

                                Dim preferredHeight As System.Int32 = CInt(System.Math.Round(workingArea.Height * 0.56R))
                                preferredHeight = System.Math.Max(lineHeight * 27, preferredHeight)
                                preferredHeight = System.Math.Min(preferredHeight, maximumClientHeight)

                                Dim minimumPromptHeight As System.Int32 = lineHeight * 7

                                Dim layoutBusy As System.Boolean = False
                                Dim expanded As System.Boolean = False
                                Dim collapsedClientHeight As System.Int32 = preferredHeight
                                Dim quickSelectedPrefix As System.String = System.String.Empty
                                Dim quickSelectedModeId As System.String = System.String.Empty

                                freestyleForm.Opacity = 0
                                freestyleForm.Text = If(System.String.IsNullOrWhiteSpace(options.Title), "Freestyle", options.Title)
                                freestyleForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
                                freestyleForm.StartPosition = System.Windows.Forms.FormStartPosition.Manual
                                freestyleForm.MaximizeBox = False
                                freestyleForm.MinimizeBox = False
                                freestyleForm.ShowInTaskbar = False
                                freestyleForm.TopMost = True
                                freestyleForm.KeyPreview = True
                                freestyleForm.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
                                freestyleForm.Font = standardFont
                                freestyleForm.ClientSize = New System.Drawing.Size(preferredWidth, preferredHeight)

                                Dim bmp As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                                freestyleForm.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

                                ' =========================================================
                                ' ROOT
                                ' =========================================================

                                Dim root As New System.Windows.Forms.TableLayoutPanel() With {
                                    .Dock = System.Windows.Forms.DockStyle.Fill,
                                    .ColumnCount = 1,
                                    .RowCount = 4,
                                    .Padding = New System.Windows.Forms.Padding(outerPadding),
                                    .Margin = New System.Windows.Forms.Padding(0)
                                }

                                root.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
                                root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
                                root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
                                root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, lineHeight * 10))
                                root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))

                                freestyleForm.Controls.Add(root)

                                ' =========================================================
                                ' HEADER
                                ' =========================================================

                                Dim header As New System.Windows.Forms.TableLayoutPanel() With {
                                    .Dock = System.Windows.Forms.DockStyle.Fill,
                                    .ColumnCount = 2,
                                    .RowCount = 1,
                                    .AutoSize = True,
                                    .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                                    .Margin = New System.Windows.Forms.Padding(0, 0, 0, normalGap)
                                }

                                header.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
                                header.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize))
                                header.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))

                                Dim headingLabel As New System.Windows.Forms.Label() With {
                                    .Text = If(options.Heading, System.String.Empty),
                                    .Font = headingFont,
                                    .AutoSize = True,
                                    .Margin = New System.Windows.Forms.Padding(0)
                                }

                                header.Controls.Add(headingLabel, 0, 0)

                                If Not System.String.IsNullOrWhiteSpace(options.ModelText) Then

                                    Dim modelLabel As New System.Windows.Forms.Label() With {
                                        .Text = options.ModelText,
                                        .Font = standardFont,
                                        .AutoSize = True,
                                        .Anchor = System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right,
                                        .TextAlign = System.Drawing.ContentAlignment.MiddleRight,
                                        .Margin = New System.Windows.Forms.Padding(largeGap, 3, 0, 0)
                                    }

                                    header.Controls.Add(modelLabel, 1, 0)

                                End If

                                root.Controls.Add(header, 0, 0)

                                ' =========================================================
                                ' PROMPT
                                ' =========================================================

                                Dim promptTextBox As New System.Windows.Forms.TextBox() With {
                                    .Text = If(options.InitialPrompt, System.String.Empty),
                                    .Font = standardFont,
                                    .Multiline = True,
                                    .AcceptsReturn = True,
                                    .AcceptsTab = True,
                                    .WordWrap = True,
                                    .ScrollBars = System.Windows.Forms.ScrollBars.Vertical,
                                    .Dock = System.Windows.Forms.DockStyle.Fill,
                                    .MinimumSize = New System.Drawing.Size(0, minimumPromptHeight),
                                    .Margin = New System.Windows.Forms.Padding(0)
                                }

                                root.Controls.Add(promptTextBox, 0, 1)

                                ' =========================================================
                                ' SETTINGS HOST
                                '
                                ' IMPORTANT:
                                ' Margin belongs to the TableLayout cell.
                                ' The root-row height therefore always includes
                                ' settingsHost.Margin.Vertical.
                                ' =========================================================

                                Dim settingsHost As New System.Windows.Forms.Panel() With {
                                    .Dock = System.Windows.Forms.DockStyle.Fill,
                                    .AutoScroll = True,
                                    .Margin = New System.Windows.Forms.Padding(0, normalGap, 0, normalGap)
                                }

                                Dim settingsCanvas As New System.Windows.Forms.Panel() With {
                                    .AutoSize = False,
                                    .Margin = New System.Windows.Forms.Padding(0),
                                    .Padding = New System.Windows.Forms.Padding(0),
                                    .Location = New System.Drawing.Point(0, 0)
                                }

                                settingsHost.Controls.Add(settingsCanvas)
                                root.Controls.Add(settingsHost, 0, 2)

                                ' =========================================================
                                ' SHORTCUTS
                                ' =========================================================

                                Dim shortcutText As System.String = "Ctrl+Enter Run"

                                If options.PromptLibraryEnabled Then
                                    shortcutText &= "   •   / Prompt library   •   Empty + Run opens Prompt library"
                                End If

                                If Not System.String.IsNullOrWhiteSpace(options.LastPrompt) Then
                                    shortcutText &= "   •   Ctrl+P Previous prompt"
                                End If

                                If options.ShowShortCommandsHint Then shortcutText &= "   •   ? Short commands"

                                Dim shortcutLabel As New System.Windows.Forms.Label() With {
                                    .Text = shortcutText,
                                    .Font = standardFont,
                                    .AutoSize = False,
                                    .Margin = New System.Windows.Forms.Padding(0)
                                }

                                settingsCanvas.Controls.Add(shortcutLabel)

                                ' =========================================================
                                ' OUTPUT
                                ' =========================================================

                                Dim modeLabel As New System.Windows.Forms.Label() With {
                                    .Text = If(options.ModeCaption, "Output"),
                                    .Font = sectionFont,
                                    .AutoSize = True,
                                    .Margin = New System.Windows.Forms.Padding(0)
                                }

                                settingsCanvas.Controls.Add(modeLabel)

                                Dim modeCombo As New System.Windows.Forms.ComboBox() With {
                                    .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList,
                                    .Font = standardFont,
                                    .Margin = New System.Windows.Forms.Padding(0)
                                }

                                settingsCanvas.Controls.Add(modeCombo)

                                Dim defaultModeIndex As System.Int32 = -1

                                For modeIndex As System.Int32 = 0 To options.Modes.Count - 1

                                    Dim modeItem As FreestylePromptMode = options.Modes(modeIndex)

                                    If modeItem.Prefixes Is Nothing Then
                                        modeItem.Prefixes = New System.Collections.Generic.List(Of System.String)()
                                    End If

                                    If modeItem.Prefixes.Count = 0 AndAlso Not System.String.IsNullOrWhiteSpace(modeItem.Prefix) Then
                                        modeItem.Prefixes.Add(modeItem.Prefix)
                                    End If

                                    If System.String.IsNullOrWhiteSpace(modeItem.ManualSyntax) AndAlso modeItem.Prefixes.Count > 0 Then
                                        modeItem.ManualSyntax = System.String.Join(" / ", modeItem.Prefixes)
                                    End If

                                    modeCombo.Items.Add(modeItem)

                                    If modeItem.IsDefault Then
                                        defaultModeIndex = modeIndex
                                    End If

                                Next

                                If modeCombo.Items.Count > 0 Then

                                    If defaultModeIndex < 0 Then
                                        defaultModeIndex = 0
                                    End If

                                    modeCombo.SelectedIndex = defaultModeIndex

                                End If

                                ' =========================================================
                                ' PREFIX SYNCHRONISATION
                                ' =========================================================

                                Dim synchronizingMode As System.Boolean = False
                                Dim synchronizingToggleOptions As System.Boolean = False
                                Dim synchronizeToggleOptionsFromPrompt As System.Action = Nothing
                                Dim synchronizePromptFromCheckedToggleOptions As System.Action = Nothing
                                Dim updateLayout As System.Action = Nothing
                                Dim defaultMode As FreestylePromptMode = Nothing
                                Dim lastAvailableMode As FreestylePromptMode = Nothing

                                If defaultModeIndex >= 0 AndAlso defaultModeIndex < modeCombo.Items.Count Then

                                    defaultMode = TryCast(modeCombo.Items(defaultModeIndex), FreestylePromptMode)

                                    If defaultMode IsNot Nothing AndAlso defaultMode.IsAvailable Then
                                        lastAvailableMode = defaultMode
                                    End If

                                End If

                                AddHandler promptTextBox.TextChanged,
                                    Sub(sender As System.Object, e As System.EventArgs)

                                        If Not synchronizingMode Then

                                            Dim match As System.Tuple(Of FreestylePromptMode, System.String) = SharedMethods.FindFreestylePromptMode(promptTextBox.Text, options.Modes)

                                            synchronizingMode = True

                                            Try

                                                If match IsNot Nothing Then

                                                    modeCombo.SelectedItem = match.Item1

                                                Else

                                                    Dim manualPrefix As System.String = SharedMethods.GetFreestyleLeadingColonToken(promptTextBox.Text)

                                                    If manualPrefix.Length > 0 AndAlso defaultMode IsNot Nothing Then
                                                        modeCombo.SelectedItem = defaultMode
                                                    End If

                                                End If

                                            Finally
                                                synchronizingMode = False
                                            End Try

                                        End If

                                        If synchronizeToggleOptionsFromPrompt IsNot Nothing Then
                                            synchronizeToggleOptionsFromPrompt.Invoke()
                                        End If

                                    End Sub

                                AddHandler modeCombo.SelectedIndexChanged,
                                    Sub(sender As System.Object, e As System.EventArgs)

                                        If synchronizingMode Then
                                            Return
                                        End If

                                        Dim selectedMode As FreestylePromptMode = TryCast(modeCombo.SelectedItem, FreestylePromptMode)

                                        If selectedMode Is Nothing Then
                                            Return
                                        End If

                                        If Not selectedMode.IsAvailable Then

                                            synchronizingMode = True

                                            Try

                                                If lastAvailableMode IsNot Nothing Then
                                                    modeCombo.SelectedItem = lastAvailableMode
                                                ElseIf defaultMode IsNot Nothing Then
                                                    modeCombo.SelectedItem = defaultMode
                                                End If

                                            Finally
                                                synchronizingMode = False
                                            End Try

                                            If Not System.String.IsNullOrWhiteSpace(selectedMode.UnavailableReason) Then
                                                SharedMethods.ShowCustomMessageBox(selectedMode.UnavailableReason, "Freestyle")
                                            End If

                                            Return

                                        End If

                                        lastAvailableMode = selectedMode

                                        Dim source As System.String = promptTextBox.Text
                                        Dim match As System.Tuple(Of FreestylePromptMode, System.String) = SharedMethods.FindFreestylePromptMode(source, options.Modes)
                                        Dim existingPrefix As System.String = System.String.Empty

                                        If match IsNot Nothing Then
                                            existingPrefix = match.Item2
                                        Else
                                            existingPrefix = SharedMethods.GetFreestyleLeadingColonToken(source)
                                        End If

                                        If existingPrefix.Length = 0 Then
                                            Return
                                        End If

                                        Dim leadingCount As System.Int32 = source.Length - source.TrimStart().Length
                                        Dim leadingWhitespace As System.String = source.Substring(0, leadingCount)
                                        Dim trimmedSource As System.String = source.Substring(leadingCount)
                                        Dim bodyStart As System.Int32 = System.Math.Min(existingPrefix.Length, trimmedSource.Length)
                                        Dim body As System.String = trimmedSource.Substring(bodyStart).TrimStart()
                                        Dim replacement As System.String = If(selectedMode.Prefix, System.String.Empty).Trim()

                                        synchronizingMode = True

                                        Try

                                            If replacement.Length = 0 Then
                                                promptTextBox.Text = leadingWhitespace & body
                                            Else
                                                promptTextBox.Text = leadingWhitespace & replacement & If(body.Length > 0, " " & body, System.String.Empty)
                                            End If

                                            promptTextBox.SelectionStart = promptTextBox.TextLength

                                        Finally
                                            synchronizingMode = False
                                        End Try

                                    End Sub

                                ' =========================================================
                                ' CONTEXT
                                ' =========================================================

                                Dim contextLabel As New System.Windows.Forms.Label() With {
                                    .Text = "Add context",
                                    .Font = sectionFont,
                                    .AutoSize = True,
                                    .Margin = New System.Windows.Forms.Padding(0)
                                }

                                settingsCanvas.Controls.Add(contextLabel)

                                Dim contextFlow As New System.Windows.Forms.FlowLayoutPanel() With {
                                    .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
                                    .WrapContents = True,
                                    .AutoSize = False,
                                    .Margin = New System.Windows.Forms.Padding(0),
                                    .Padding = New System.Windows.Forms.Padding(0)
                                }

                                settingsCanvas.Controls.Add(contextFlow)

                                Dim promptToolTip As New System.Windows.Forms.ToolTip()

                                For Each insertDefinition As FreestylePromptInsertOption In options.InsertOptions

                                    Dim insertButton As New System.Windows.Forms.Button() With {
                                        .Text = insertDefinition.Text,
                                        .AutoSize = True,
                                        .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                                        .Font = standardFont,
                                        .Padding = New System.Windows.Forms.Padding(buttonPadX, buttonPadY, buttonPadX, buttonPadY),
                                        .Tag = insertDefinition,
                                        .Margin = New System.Windows.Forms.Padding(0, 0, normalGap, smallGap)
                                    }

                                    promptToolTip.SetToolTip(insertButton, insertDefinition.Description)

                                    AddHandler insertButton.Click,
                                        Sub(sender As System.Object, e As System.EventArgs)

                                            Dim clickedButton As System.Windows.Forms.Button = TryCast(sender, System.Windows.Forms.Button)

                                            If clickedButton Is Nothing Then
                                                Return
                                            End If

                                            Dim clickedDefinition As FreestylePromptInsertOption = TryCast(clickedButton.Tag, FreestylePromptInsertOption)

                                            If clickedDefinition Is Nothing Then
                                                Return
                                            End If

                                            Dim textToInsert As System.String = clickedDefinition.InsertText

                                            If clickedDefinition.RequiresValue Then

                                                Dim value As System.String = SharedMethods.ShowCustomInputBox(clickedDefinition.ValuePrompt, clickedDefinition.ValueTitle, True, System.String.Empty)

                                                If System.String.IsNullOrWhiteSpace(value) OrElse value.Equals("ESC", System.StringComparison.OrdinalIgnoreCase) Then
                                                    promptTextBox.Focus()
                                                    Return
                                                End If

                                                textToInsert = clickedDefinition.BuildInsertText(value)

                                            End If

                                            If System.String.IsNullOrWhiteSpace(textToInsert) Then
                                                promptTextBox.Focus()
                                                Return
                                            End If

                                            If clickedDefinition.MaxOccurrences > 0 AndAlso
                                               Not System.String.IsNullOrWhiteSpace(clickedDefinition.InsertText) AndAlso
                                               SharedMethods.CountFreestyleTokenOccurrences(promptTextBox.Text, clickedDefinition.InsertText) >= clickedDefinition.MaxOccurrences Then

                                                SharedMethods.ShowCustomMessageBox(
                                                    "'" & clickedDefinition.Text & "' can be included at most " & clickedDefinition.MaxOccurrences.ToString(System.Globalization.CultureInfo.InvariantCulture) & " time(s) in one Freestyle request.",
                                                    "Freestyle")
                                                promptTextBox.Focus()
                                                Return
                                            End If

                                            SharedMethods.InsertFreestyleTextAtCaret(promptTextBox, textToInsert)

                                        End Sub

                                    contextFlow.Controls.Add(insertButton)

                                Next

                                contextLabel.Visible = options.InsertOptions.Count > 0
                                contextFlow.Visible = options.InsertOptions.Count > 0

                                ' =========================================================
                                ' MORE OPTIONS
                                ' =========================================================

                                Dim moreButton As New System.Windows.Forms.Button() With {
                                    .Text = "More options ▸",
                                    .AutoSize = True,
                                    .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                                    .Font = standardFont,
                                    .Padding = New System.Windows.Forms.Padding(buttonPadX, buttonPadY, buttonPadX, buttonPadY),
                                    .Margin = New System.Windows.Forms.Padding(0)
                                }

                                settingsCanvas.Controls.Add(moreButton)

                                ' =========================================================
                                ' ADVANCED
                                ' =========================================================

                                Dim advancedGrid As New System.Windows.Forms.TableLayoutPanel() With {
                                    .ColumnCount = 2,
                                    .RowCount = 0,
                                    .AutoSize = True,
                                    .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                                    .Visible = False,
                                    .Padding = New System.Windows.Forms.Padding(0),
                                    .Margin = New System.Windows.Forms.Padding(0)
                                }

                                advancedGrid.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0F))
                                advancedGrid.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0F))

                                settingsCanvas.Controls.Add(advancedGrid)

                                Dim toggleCheckBoxes As New System.Collections.Generic.List(Of System.Windows.Forms.CheckBox)()
                                Dim argumentTextBoxes As New System.Collections.Generic.Dictionary(Of System.Windows.Forms.CheckBox, System.Windows.Forms.TextBox)()
                                Dim toggleById As New System.Collections.Generic.Dictionary(Of System.String, System.Windows.Forms.CheckBox)(System.StringComparer.OrdinalIgnoreCase)

                                Dim visibleSectionIndex As System.Int32 = 0

                                For Each section As FreestylePromptSection In options.Sections

                                    If section Is Nothing OrElse section.Options Is Nothing OrElse section.Options.Count = 0 Then
                                        Continue For
                                    End If

                                    Dim sectionRow As System.Int32 = visibleSectionIndex \ 2
                                    Dim sectionColumn As System.Int32 = visibleSectionIndex Mod 2

                                    If advancedGrid.RowCount <= sectionRow Then
                                        advancedGrid.RowCount = sectionRow + 1
                                        advancedGrid.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
                                    End If

                                    Dim sectionPanel As New System.Windows.Forms.TableLayoutPanel() With {
                                        .Dock = System.Windows.Forms.DockStyle.Fill,
                                        .ColumnCount = 1,
                                        .RowCount = 2,
                                        .AutoSize = True,
                                        .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                                        .Padding = New System.Windows.Forms.Padding(0),
                                        .Margin = New System.Windows.Forms.Padding(If(sectionColumn = 0, 0, normalGap), 0, If(sectionColumn = 0, normalGap, 0), normalGap)
                                    }

                                    sectionPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
                                    sectionPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
                                    sectionPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))

                                    Dim sectionLabel As New System.Windows.Forms.Label() With {
                                        .Text = section.Caption,
                                        .Font = sectionFont,
                                        .AutoSize = True,
                                        .Margin = New System.Windows.Forms.Padding(0, 0, 0, smallGap)
                                    }

                                    sectionPanel.Controls.Add(sectionLabel, 0, 0)

                                    Dim optionsTable As New System.Windows.Forms.TableLayoutPanel() With {
                                        .Dock = System.Windows.Forms.DockStyle.Top,
                                        .ColumnCount = 1,
                                        .RowCount = 0,
                                        .AutoSize = True,
                                        .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                                        .Padding = New System.Windows.Forms.Padding(0),
                                        .Margin = New System.Windows.Forms.Padding(0)
                                    }

                                    optionsTable.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))

                                    Dim optionRow As System.Int32 = 0

                                    For Each definition As FreestylePromptToggleOption In section.Options

                                        optionsTable.RowCount = optionRow + 1
                                        optionsTable.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))

                                        Dim optionHost As New System.Windows.Forms.TableLayoutPanel() With {
                                            .Dock = System.Windows.Forms.DockStyle.Top,
                                            .ColumnCount = 1,
                                            .RowCount = If(definition.HasArgument, 2, 1),
                                            .AutoSize = True,
                                            .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                                            .Padding = New System.Windows.Forms.Padding(0),
                                            .Margin = New System.Windows.Forms.Padding(0, 1, 0, 3)
                                        }

                                        optionHost.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
                                        optionHost.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))

                                        If definition.HasArgument Then
                                            optionHost.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
                                        End If

                                        Dim displayText As System.String = definition.Text

                                        If Not System.String.IsNullOrWhiteSpace(definition.ManualSyntax) Then
                                            displayText &= "   [" & definition.ManualSyntax & "]"
                                        End If

                                        Dim optionCheckBox As New System.Windows.Forms.CheckBox() With {
                                            .Text = displayText,
                                            .Checked = definition.IsChecked,
                                            .AutoSize = False,
                                            .AutoEllipsis = True,
                                            .Font = standardFont,
                                            .Tag = definition,
                                            .Dock = System.Windows.Forms.DockStyle.Top,
                                            .Height = lineHeight + smallGap + 4,
                                            .Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
                                        }

                                        promptToolTip.SetToolTip(optionCheckBox, displayText & If(System.String.IsNullOrWhiteSpace(definition.Description), System.String.Empty, System.Environment.NewLine & definition.Description))

                                        optionHost.Controls.Add(optionCheckBox, 0, 0)
                                        toggleCheckBoxes.Add(optionCheckBox)

                                        If Not System.String.IsNullOrWhiteSpace(definition.Id) AndAlso Not toggleById.ContainsKey(definition.Id) Then
                                            toggleById.Add(definition.Id, optionCheckBox)
                                        End If

                                        If definition.HasArgument Then

                                            Dim argumentTextBox As New System.Windows.Forms.TextBox() With {
                                                .Font = standardFont,
                                                .Dock = System.Windows.Forms.DockStyle.Top,
                                                .Enabled = definition.IsChecked,
                                                .Margin = New System.Windows.Forms.Padding(22, 2, 0, 2)
                                            }

                                            promptToolTip.SetToolTip(argumentTextBox, definition.ArgumentHint)

                                            optionHost.Controls.Add(argumentTextBox, 0, 1)
                                            argumentTextBoxes.Add(optionCheckBox, argumentTextBox)

                                        End If

                                        AddHandler optionCheckBox.CheckedChanged,
                                            Sub(sender As System.Object, e As System.EventArgs)

                                                Dim changedCheckBox As System.Windows.Forms.CheckBox = TryCast(sender, System.Windows.Forms.CheckBox)

                                                If changedCheckBox Is Nothing Then
                                                    Return
                                                End If

                                                Dim changedDefinition As FreestylePromptToggleOption = TryCast(changedCheckBox.Tag, FreestylePromptToggleOption)

                                                If changedDefinition Is Nothing Then
                                                    Return
                                                End If

                                                Dim linkedTextBox As System.Windows.Forms.TextBox = Nothing

                                                If argumentTextBoxes.ContainsKey(changedCheckBox) Then
                                                    linkedTextBox = argumentTextBoxes(changedCheckBox)
                                                    linkedTextBox.Enabled = changedCheckBox.Checked
                                                End If

                                                If synchronizingToggleOptions Then
                                                    Return
                                                End If

                                                synchronizingToggleOptions = True

                                                Try

                                                    Dim updatedPrompt As System.String = promptTextBox.Text

                                                    If changedCheckBox.Checked Then

                                                        Dim matchedTrigger As System.String = System.String.Empty
                                                        Dim matchedArgument As System.String = System.String.Empty

                                                        If SharedMethods.TryGetFreestyleToggleTriggerMatch(updatedPrompt, changedDefinition, matchedTrigger, matchedArgument) Then

                                                            If linkedTextBox IsNot Nothing AndAlso matchedArgument.Length > 0 Then
                                                                linkedTextBox.Text = matchedArgument
                                                            End If

                                                        Else

                                                            Dim argumentValue As System.String = System.String.Empty

                                                            If linkedTextBox IsNot Nothing Then
                                                                argumentValue = linkedTextBox.Text.Trim()
                                                            End If

                                                            Dim triggerToInsert As System.String = changedDefinition.BuildTrigger(argumentValue)

                                                            If Not System.String.IsNullOrWhiteSpace(triggerToInsert) Then
                                                                updatedPrompt = SharedMethods.AppendFreestyleToggleTrigger(updatedPrompt, triggerToInsert)
                                                            End If

                                                        End If

                                                    Else

                                                        updatedPrompt = SharedMethods.RemoveFreestyleToggleTriggers(updatedPrompt, changedDefinition)

                                                    End If

                                                    If Not System.String.Equals(updatedPrompt, promptTextBox.Text, System.StringComparison.Ordinal) Then
                                                        Dim oldSelectionStart As System.Int32 = promptTextBox.SelectionStart
                                                        promptTextBox.Text = updatedPrompt
                                                        promptTextBox.SelectionStart = System.Math.Min(oldSelectionStart, promptTextBox.TextLength)
                                                    End If

                                                Finally

                                                    synchronizingToggleOptions = False

                                                End Try

                                                If changedCheckBox.Checked AndAlso linkedTextBox IsNot Nothing Then
                                                    linkedTextBox.Focus()
                                                End If

                                            End Sub

                                        If definition.HasArgument Then

                                            Dim argumentTextBoxForHandler As System.Windows.Forms.TextBox = argumentTextBoxes(optionCheckBox)
                                            argumentTextBoxForHandler.Tag = optionCheckBox

                                            AddHandler argumentTextBoxForHandler.TextChanged,
                                                Sub(sender As System.Object, e As System.EventArgs)

                                                    If synchronizingToggleOptions Then
                                                        Return
                                                    End If

                                                    Dim changedTextBox As System.Windows.Forms.TextBox = TryCast(sender, System.Windows.Forms.TextBox)

                                                    If changedTextBox Is Nothing Then
                                                        Return
                                                    End If

                                                    Dim linkedCheckBox As System.Windows.Forms.CheckBox = TryCast(changedTextBox.Tag, System.Windows.Forms.CheckBox)

                                                    If linkedCheckBox Is Nothing OrElse Not linkedCheckBox.Checked Then
                                                        Return
                                                    End If

                                                    Dim changedDefinition As FreestylePromptToggleOption = TryCast(linkedCheckBox.Tag, FreestylePromptToggleOption)

                                                    If changedDefinition Is Nothing Then
                                                        Return
                                                    End If

                                                    synchronizingToggleOptions = True

                                                    Try

                                                        Dim updatedPrompt As System.String = SharedMethods.RemoveFreestyleToggleTriggers(promptTextBox.Text, changedDefinition)
                                                        Dim triggerToInsert As System.String = changedDefinition.BuildTrigger(changedTextBox.Text.Trim())

                                                        If Not System.String.IsNullOrWhiteSpace(triggerToInsert) Then
                                                            updatedPrompt = SharedMethods.AppendFreestyleToggleTrigger(updatedPrompt, triggerToInsert)
                                                        End If

                                                        If Not System.String.Equals(updatedPrompt, promptTextBox.Text, System.StringComparison.Ordinal) Then
                                                            Dim oldSelectionStart As System.Int32 = promptTextBox.SelectionStart
                                                            promptTextBox.Text = updatedPrompt
                                                            promptTextBox.SelectionStart = System.Math.Min(oldSelectionStart, promptTextBox.TextLength)
                                                        End If

                                                    Finally

                                                        synchronizingToggleOptions = False

                                                    End Try

                                                End Sub

                                        End If

                                        optionsTable.Controls.Add(optionHost, 0, optionRow)

                                        optionRow += 1

                                    Next

                                    sectionPanel.Controls.Add(optionsTable, 0, 1)
                                    advancedGrid.Controls.Add(sectionPanel, sectionColumn, sectionRow)

                                    visibleSectionIndex += 1

                                Next

                                synchronizeToggleOptionsFromPrompt =
                                    Sub()

                                        If synchronizingToggleOptions Then
                                            Return
                                        End If

                                        synchronizingToggleOptions = True

                                        Try

                                            Dim anyDetected As System.Boolean = False
                                            Dim promptText As System.String = If(promptTextBox.Text, System.String.Empty)

                                            For Each optionCheckBox As System.Windows.Forms.CheckBox In toggleCheckBoxes

                                                If optionCheckBox Is Nothing Then
                                                    Continue For
                                                End If

                                                Dim definition As FreestylePromptToggleOption = TryCast(optionCheckBox.Tag, FreestylePromptToggleOption)

                                                If definition Is Nothing Then
                                                    Continue For
                                                End If

                                                Dim canDetect As System.Boolean = definition.HasArgument OrElse Not System.String.IsNullOrWhiteSpace(definition.Trigger)

                                                If Not canDetect Then
                                                    Continue For
                                                End If

                                                Dim matchedTrigger As System.String = System.String.Empty
                                                Dim matchedArgument As System.String = System.String.Empty
                                                Dim detected As System.Boolean = SharedMethods.TryGetFreestyleToggleTriggerMatch(promptText, definition, matchedTrigger, matchedArgument)

                                                optionCheckBox.Checked = detected

                                                If argumentTextBoxes.ContainsKey(optionCheckBox) Then

                                                    Dim argumentTextBox As System.Windows.Forms.TextBox = argumentTextBoxes(optionCheckBox)
                                                    argumentTextBox.Enabled = detected

                                                    If detected Then
                                                        argumentTextBox.Text = matchedArgument
                                                    End If

                                                End If

                                                If detected Then
                                                    anyDetected = True
                                                End If

                                            Next

                                            If anyDetected AndAlso Not expanded Then

                                                If freestyleForm.Visible Then
                                                    collapsedClientHeight = freestyleForm.ClientSize.Height
                                                End If

                                                expanded = True
                                                advancedGrid.Visible = True
                                                moreButton.Text = "More options ▾"

                                                If updateLayout IsNot Nothing Then
                                                    updateLayout.Invoke()
                                                End If

                                            End If

                                        Finally

                                            synchronizingToggleOptions = False

                                        End Try

                                    End Sub

                                synchronizePromptFromCheckedToggleOptions =
                                    Sub()

                                        If synchronizingToggleOptions Then
                                            Return
                                        End If

                                        synchronizingToggleOptions = True

                                        Try

                                            Dim updatedPrompt As System.String = promptTextBox.Text

                                            For Each optionCheckBox As System.Windows.Forms.CheckBox In toggleCheckBoxes

                                                If optionCheckBox Is Nothing OrElse Not optionCheckBox.Checked Then
                                                    Continue For
                                                End If

                                                Dim definition As FreestylePromptToggleOption = TryCast(optionCheckBox.Tag, FreestylePromptToggleOption)

                                                If definition Is Nothing Then
                                                    Continue For
                                                End If

                                                Dim matchedTrigger As System.String = System.String.Empty
                                                Dim matchedArgument As System.String = System.String.Empty

                                                If SharedMethods.TryGetFreestyleToggleTriggerMatch(updatedPrompt, definition, matchedTrigger, matchedArgument) Then

                                                    If argumentTextBoxes.ContainsKey(optionCheckBox) AndAlso matchedArgument.Length > 0 Then
                                                        argumentTextBoxes(optionCheckBox).Text = matchedArgument
                                                    End If

                                                    Continue For

                                                End If

                                                Dim argumentValue As System.String = System.String.Empty

                                                If argumentTextBoxes.ContainsKey(optionCheckBox) Then
                                                    argumentValue = argumentTextBoxes(optionCheckBox).Text.Trim()
                                                End If

                                                Dim triggerToInsert As System.String = definition.BuildTrigger(argumentValue)

                                                If Not System.String.IsNullOrWhiteSpace(triggerToInsert) Then
                                                    updatedPrompt = SharedMethods.AppendFreestyleToggleTrigger(updatedPrompt, triggerToInsert)
                                                End If

                                            Next

                                            If Not System.String.Equals(updatedPrompt, promptTextBox.Text, System.StringComparison.Ordinal) Then
                                                Dim oldSelectionStart As System.Int32 = promptTextBox.SelectionStart
                                                promptTextBox.Text = updatedPrompt
                                                promptTextBox.SelectionStart = System.Math.Min(oldSelectionStart, promptTextBox.TextLength)
                                            End If

                                        Finally

                                            synchronizingToggleOptions = False

                                        End Try

                                    End Sub

                                moreButton.Visible = visibleSectionIndex > 0

                                If options.RestorePersistedState AndAlso
                                   Not System.String.IsNullOrWhiteSpace(options.CallerId) AndAlso
                                   Not System.String.IsNullOrWhiteSpace(options.PersistedState) Then

                                    Dim persistedStateDocument As System.Xml.Linq.XDocument =
                                        SharedMethods.GetFreestylePromptStateDocument(options.PersistedState)

                                    Dim persistedCallerElement As System.Xml.Linq.XElement =
                                        SharedMethods.GetFreestylePromptCallerStateElement(
                                            persistedStateDocument,
                                            options.CallerId,
                                            False)

                                    If persistedCallerElement IsNot Nothing Then

                                        Dim savedPrefix As System.String =
                                            SharedMethods.GetFreestylePromptStateAttribute(persistedCallerElement, "prefix")

                                        Dim savedModeId As System.String =
                                            SharedMethods.GetFreestylePromptStateAttribute(persistedCallerElement, "modeId")

                                        Dim restoredMode As FreestylePromptMode = Nothing

                                        If savedPrefix.Length > 0 Then
                                            restoredMode = SharedMethods.FindFreestylePromptModeByPrefix(savedPrefix, options.Modes)
                                        End If

                                        If restoredMode Is Nothing AndAlso savedModeId.Length > 0 Then
                                            restoredMode = SharedMethods.FindFreestylePromptModeById(savedModeId, options.Modes)
                                        End If

                                        If restoredMode IsNot Nothing AndAlso restoredMode.IsAvailable Then
                                            synchronizingMode = True

                                            Try
                                                modeCombo.SelectedItem = restoredMode
                                                lastAvailableMode = restoredMode
                                            Finally
                                                synchronizingMode = False
                                            End Try
                                        End If

                                        synchronizingToggleOptions = True

                                        Try

                                            For Each optionElement As System.Xml.Linq.XElement In persistedCallerElement.Elements("option")

                                                Dim optionId As System.String =
                                                SharedMethods.GetFreestylePromptStateAttribute(optionElement, "id")

                                                Dim optionCheckBox As System.Windows.Forms.CheckBox = Nothing

                                                If optionId.Length = 0 OrElse Not toggleById.TryGetValue(optionId, optionCheckBox) Then
                                                    Continue For
                                                End If

                                                optionCheckBox.Checked =
                                                SharedMethods.GetFreestylePromptStateAttribute(optionElement, "checked") = "1"

                                                If argumentTextBoxes.ContainsKey(optionCheckBox) Then
                                                    argumentTextBoxes(optionCheckBox).Text =
                                                    SharedMethods.GetFreestylePromptStateAttribute(optionElement, "argument")
                                                End If

                                            Next

                                        Finally

                                            synchronizingToggleOptions = False

                                        End Try

                                        expanded =
                                            SharedMethods.GetFreestylePromptStateAttribute(persistedCallerElement, "expanded") = "1"

                                        advancedGrid.Visible = expanded
                                        moreButton.Text = If(expanded, "More options ▾", "More options ▸")

                                    End If

                                End If

                                If synchronizePromptFromCheckedToggleOptions IsNot Nothing Then
                                    synchronizePromptFromCheckedToggleOptions.Invoke()
                                End If

                                If synchronizeToggleOptionsFromPrompt IsNot Nothing Then
                                    synchronizeToggleOptionsFromPrompt.Invoke()
                                End If

                                ' =========================================================
                                ' FOOTER
                                ' =========================================================

                                Dim footer As New System.Windows.Forms.TableLayoutPanel() With {
                                    .Dock = System.Windows.Forms.DockStyle.Fill,
                                    .ColumnCount = 2,
                                    .RowCount = 1,
                                    .AutoSize = True,
                                    .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                                    .Margin = New System.Windows.Forms.Padding(0)
                                }

                                footer.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
                                footer.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize))
                                footer.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))

                                Dim statusLabel As New System.Windows.Forms.Label() With {
                                    .Text = If(options.ContextStatusText, System.String.Empty),
                                    .Font = standardFont,
                                    .AutoSize = True,
                                    .Anchor = System.Windows.Forms.AnchorStyles.Left,
                                    .Margin = New System.Windows.Forms.Padding(0, 7, largeGap, 0)
                                }

                                Dim actionFlow As New System.Windows.Forms.FlowLayoutPanel() With {
                                    .AutoSize = True,
                                    .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                                    .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
                                    .WrapContents = False,
                                    .Anchor = System.Windows.Forms.AnchorStyles.Right,
                                    .Margin = New System.Windows.Forms.Padding(0)
                                }

                                Dim cancelButton As New System.Windows.Forms.Button() With {
                                    .Text = "Cancel",
                                    .DialogResult = System.Windows.Forms.DialogResult.Cancel,
                                    .AutoSize = True,
                                    .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                                    .Font = standardFont,
                                    .Padding = New System.Windows.Forms.Padding(buttonPadX + 2, buttonPadY + 2, buttonPadX + 2, buttonPadY + 2),
                                    .Margin = New System.Windows.Forms.Padding(0, 0, normalGap, 0)
                                }

                                Dim runButton As New System.Windows.Forms.Button() With {
                                    .Text = "Run",
                                    .DialogResult = System.Windows.Forms.DialogResult.OK,
                                    .AutoSize = True,
                                    .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                                    .Font = standardFont,
                                    .Padding = New System.Windows.Forms.Padding(buttonPadX + 2, buttonPadY + 2, buttonPadX + 2, buttonPadY + 2),
                                    .Margin = New System.Windows.Forms.Padding(0)
                                }

                                If options.QuickButtons IsNot Nothing Then

                                    For Each quickButtonDefinition As FreestylePromptQuickButton In options.QuickButtons

                                        If quickButtonDefinition Is Nothing OrElse
                                           System.String.IsNullOrWhiteSpace(quickButtonDefinition.Text) OrElse
                                           System.String.IsNullOrWhiteSpace(quickButtonDefinition.Prefix) Then
                                            Continue For
                                        End If

                                        Dim quickRunButton As New System.Windows.Forms.Button() With {
                                            .Text = quickButtonDefinition.Text,
                                            .AutoSize = True,
                                            .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                                            .Font = standardFont,
                                            .Padding = New System.Windows.Forms.Padding(buttonPadX + 2, buttonPadY + 2, buttonPadX + 2, buttonPadY + 2),
                                            .Margin = New System.Windows.Forms.Padding(0, 0, normalGap, 0),
                                            .Tag = quickButtonDefinition
                                        }

                                        promptToolTip.SetToolTip(quickRunButton, quickButtonDefinition.Description)

                                        AddHandler quickRunButton.Click,
                                            Sub(sender As System.Object, e As System.EventArgs)

                                                Dim clickedButton As System.Windows.Forms.Button = TryCast(sender, System.Windows.Forms.Button)

                                                If clickedButton Is Nothing Then
                                                    Return
                                                End If

                                                Dim clickedDefinition As FreestylePromptQuickButton =
                                                    TryCast(clickedButton.Tag, FreestylePromptQuickButton)

                                                If clickedDefinition Is Nothing OrElse
                                                   System.String.IsNullOrWhiteSpace(clickedDefinition.Prefix) Then
                                                    Return
                                                End If

                                                Dim quickMode As FreestylePromptMode =
                                                    SharedMethods.FindFreestylePromptModeByPrefix(clickedDefinition.Prefix, options.Modes)

                                                If quickMode IsNot Nothing AndAlso Not quickMode.IsAvailable Then
                                                    If Not System.String.IsNullOrWhiteSpace(quickMode.UnavailableReason) Then
                                                        SharedMethods.ShowCustomMessageBox(quickMode.UnavailableReason, "Freestyle")
                                                    End If
                                                    Return
                                                End If

                                                quickSelectedPrefix = clickedDefinition.Prefix.Trim()
                                                quickSelectedModeId = If(quickMode Is Nothing, System.String.Empty, quickMode.Id)

                                                freestyleForm.DialogResult = System.Windows.Forms.DialogResult.OK
                                                freestyleForm.Close()

                                            End Sub

                                        actionFlow.Controls.Add(quickRunButton)

                                    Next

                                End If

                                actionFlow.Controls.Add(cancelButton)
                                actionFlow.Controls.Add(runButton)

                                footer.Controls.Add(statusLabel, 0, 0)
                                footer.Controls.Add(actionFlow, 1, 0)

                                root.Controls.Add(footer, 0, 3)

                                freestyleForm.AcceptButton = runButton
                                freestyleForm.CancelButton = cancelButton

                                ' =========================================================
                                ' SETTINGS CANVAS
                                ' =========================================================

                                Dim layoutSettingsCanvas As System.Func(Of System.Int32, System.Int32) =
                                    Function(contentWidth As System.Int32) As System.Int32

                                        contentWidth = System.Math.Max(250, contentWidth)

                                        Dim y As System.Int32 = 0

                                        shortcutLabel.Location = New System.Drawing.Point(0, y)
                                        shortcutLabel.Width = contentWidth

                                        Dim shortcutPreferred As System.Drawing.Size = shortcutLabel.GetPreferredSize(New System.Drawing.Size(contentWidth, 0))

                                        shortcutLabel.Height = shortcutPreferred.Height

                                        y = shortcutLabel.Bottom + normalGap

                                        modeLabel.Location = New System.Drawing.Point(0, y)

                                        y = modeLabel.Bottom + smallGap

                                        modeCombo.Location = New System.Drawing.Point(0, y)
                                        modeCombo.Width = contentWidth

                                        y = modeCombo.Bottom + normalGap

                                        If contextFlow.Visible Then

                                            contextLabel.Location = New System.Drawing.Point(0, y)

                                            y = contextLabel.Bottom + smallGap

                                            contextFlow.Location = New System.Drawing.Point(0, y)
                                            contextFlow.Width = contentWidth

                                            Dim contextPreferred As System.Drawing.Size = contextFlow.GetPreferredSize(New System.Drawing.Size(contentWidth, 0))

                                            contextFlow.Height = System.Math.Max(contextPreferred.Height, lineHeight + buttonPadY * 2 + 12)

                                            y = contextFlow.Bottom + smallGap

                                        End If

                                        moreButton.Location = New System.Drawing.Point(0, y)

                                        ' Explicit bottom breathing room prevents a button
                                        ' from visually touching the clipping boundary.
                                        y = moreButton.Bottom + smallGap

                                        If advancedGrid.Visible Then

                                            y += normalGap

                                            advancedGrid.Location = New System.Drawing.Point(0, y)
                                            advancedGrid.Width = contentWidth
                                            advancedGrid.MinimumSize = New System.Drawing.Size(contentWidth, 0)
                                            advancedGrid.MaximumSize = New System.Drawing.Size(contentWidth, 0)

                                            advancedGrid.PerformLayout()

                                            Dim advancedPreferred As System.Drawing.Size = advancedGrid.GetPreferredSize(New System.Drawing.Size(contentWidth, 0))

                                            advancedGrid.Height = advancedPreferred.Height

                                            y = advancedGrid.Bottom + smallGap

                                        Else

                                            advancedGrid.Location = New System.Drawing.Point(0, y)
                                            advancedGrid.Width = contentWidth
                                            advancedGrid.Height = 0

                                        End If

                                        settingsCanvas.Location = New System.Drawing.Point(0, 0)
                                        settingsCanvas.Size = New System.Drawing.Size(contentWidth, y)

                                        Return y

                                    End Function

                                ' =========================================================
                                ' LAYOUT ENGINE
                                ' =========================================================

                                updateLayout =
                                    Sub()

                                        If layoutBusy OrElse freestyleForm.IsDisposed Then
                                            Return
                                        End If

                                        layoutBusy = True

                                        Try

                                            header.PerformLayout()
                                            footer.PerformLayout()
                                            root.PerformLayout()

                                            Dim scrollbarWidth As System.Int32 = System.Windows.Forms.SystemInformation.VerticalScrollBarWidth

                                            ' First pass assumes no vertical scrollbar.
                                            Dim contentWidth As System.Int32 = settingsHost.ClientSize.Width

                                            contentWidth = System.Math.Max(250, contentWidth)

                                            Dim requiredSettingsHeight As System.Int32 = layoutSettingsCanvas.Invoke(contentWidth)

                                            Dim headerRowHeight As System.Int32 = root.GetRowHeights()(0)
                                            Dim footerRowHeight As System.Int32 = root.GetRowHeights()(3)

                                            ' This is the conceptual correction:
                                            ' the row contains the viewport PLUS its margins.
                                            Dim settingsMargins As System.Int32 = settingsHost.Margin.Vertical

                                            Dim requiredSettingsRowHeight As System.Int32 = requiredSettingsHeight + settingsMargins

                                            Dim requiredClientHeight As System.Int32 = root.Padding.Vertical + headerRowHeight + minimumPromptHeight + requiredSettingsRowHeight + footerRowHeight

                                            If requiredClientHeight > freestyleForm.ClientSize.Height AndAlso freestyleForm.ClientSize.Height < maximumClientHeight Then

                                                Dim targetClientHeight As System.Int32 = System.Math.Min(requiredClientHeight, maximumClientHeight)

                                                freestyleForm.ClientSize = New System.Drawing.Size(freestyleForm.ClientSize.Width, targetClientHeight)

                                                root.PerformLayout()

                                            End If

                                            Dim availableRowHeight As System.Int32 = freestyleForm.ClientSize.Height - root.Padding.Vertical - headerRowHeight - footerRowHeight - minimumPromptHeight

                                            availableRowHeight = System.Math.Max(lineHeight * 6 + settingsMargins, availableRowHeight)

                                            Dim actualSettingsRowHeight As System.Int32 = System.Math.Min(requiredSettingsRowHeight, availableRowHeight)

                                            root.RowStyles(2).Height = actualSettingsRowHeight

                                            Dim actualViewportHeight As System.Int32 = actualSettingsRowHeight - settingsMargins

                                            actualViewportHeight = System.Math.Max(1, actualViewportHeight)

                                            Dim needsVerticalScroll As System.Boolean = requiredSettingsHeight > actualViewportHeight

                                            If needsVerticalScroll Then
                                                contentWidth = settingsHost.ClientSize.Width - scrollbarWidth
                                            Else
                                                contentWidth = settingsHost.ClientSize.Width
                                            End If

                                            contentWidth = System.Math.Max(250, contentWidth)

                                            requiredSettingsHeight = layoutSettingsCanvas.Invoke(contentWidth)

                                            settingsHost.AutoScrollMinSize = New System.Drawing.Size(0, requiredSettingsHeight)

                                            ' Prevent horizontal scrolling by keeping the
                                            ' canvas strictly narrower than the viewport.
                                            settingsCanvas.Width = contentWidth

                                            root.PerformLayout()
                                            settingsHost.PerformLayout()

                                        Finally

                                            layoutBusy = False

                                        End Try

                                    End Sub

                                ' =========================================================
                                ' MORE OPTIONS
                                ' =========================================================

                                AddHandler moreButton.Click,
                                    Sub(sender As System.Object, e As System.EventArgs)

                                        If layoutBusy Then
                                            Return
                                        End If

                                        If Not expanded Then

                                            collapsedClientHeight = freestyleForm.ClientSize.Height

                                            expanded = True
                                            advancedGrid.Visible = True
                                            moreButton.Text = "More options ▾"

                                        Else

                                            expanded = False
                                            advancedGrid.Visible = False
                                            advancedGrid.Height = 0
                                            moreButton.Text = "More options ▸"

                                            Dim restoreHeight As System.Int32 = System.Math.Min(collapsedClientHeight, maximumClientHeight)

                                            freestyleForm.ClientSize = New System.Drawing.Size(freestyleForm.ClientSize.Width, restoreHeight)

                                        End If

                                        updateLayout.Invoke()

                                        Dim centredY As System.Int32 = workingArea.Y + (workingArea.Height - freestyleForm.Height) \ 2

                                        centredY = System.Math.Max(workingArea.Top, centredY)
                                        centredY = System.Math.Min(centredY, workingArea.Bottom - freestyleForm.Height)

                                        freestyleForm.Location = New System.Drawing.Point(freestyleForm.Left, centredY)

                                    End Sub

                                AddHandler freestyleForm.Resize,
                                    Sub(sender As System.Object, e As System.EventArgs)
                                        updateLayout.Invoke()
                                    End Sub

                                ' =========================================================
                                ' ARGUMENT VALIDATION
                                ' =========================================================

                                AddHandler freestyleForm.FormClosing,
                                    Sub(sender As System.Object, e As System.Windows.Forms.FormClosingEventArgs)

                                        If freestyleForm.DialogResult <> System.Windows.Forms.DialogResult.OK Then
                                            Return
                                        End If

                                        Dim violatingInsertOption As FreestylePromptInsertOption = Nothing
                                        If Not SharedMethods.ValidateFreestyleInsertOccurrenceLimits(promptTextBox.Text, options.InsertOptions, violatingInsertOption) Then
                                            e.Cancel = True
                                            Dim optionLabel As System.String = If(violatingInsertOption Is Nothing OrElse System.String.IsNullOrWhiteSpace(violatingInsertOption.Text), "This context option", violatingInsertOption.Text)
                                            Dim maximum As System.Int32 = If(violatingInsertOption Is Nothing, 1, violatingInsertOption.MaxOccurrences)
                                            SharedMethods.ShowCustomMessageBox(
                                                "'" & optionLabel & "' can be included at most " & maximum.ToString(System.Globalization.CultureInfo.InvariantCulture) & " time(s) in one Freestyle request.",
                                                "Freestyle")
                                            promptTextBox.Focus()
                                            Return
                                        End If

                                        For Each pair As System.Collections.Generic.KeyValuePair(Of System.Windows.Forms.CheckBox, System.Windows.Forms.TextBox) In argumentTextBoxes

                                            If Not pair.Key.Checked Then
                                                Continue For
                                            End If

                                            Dim definition As FreestylePromptToggleOption = TryCast(pair.Key.Tag, FreestylePromptToggleOption)

                                            If definition Is Nothing Then
                                                Continue For
                                            End If

                                            If definition.ArgumentRequired AndAlso System.String.IsNullOrWhiteSpace(pair.Value.Text) Then

                                                e.Cancel = True

                                                SharedMethods.ShowCustomMessageBox("'" & definition.Text & "' requires a value.", "Freestyle")

                                                pair.Value.Focus()

                                                Return

                                            End If

                                        Next

                                    End Sub

                                ' =========================================================
                                ' KEYBOARD
                                ' =========================================================

                                AddHandler promptTextBox.KeyDown,
                                    Sub(sender As System.Object, e As System.Windows.Forms.KeyEventArgs)

                                        If e.KeyCode = System.Windows.Forms.Keys.Enter AndAlso e.Modifiers = System.Windows.Forms.Keys.Control Then
                                            freestyleForm.DialogResult = System.Windows.Forms.DialogResult.OK
                                            freestyleForm.Close()
                                            e.SuppressKeyPress = True
                                            Return
                                        End If

                                        If e.KeyCode = System.Windows.Forms.Keys.Escape Then
                                            freestyleForm.DialogResult = System.Windows.Forms.DialogResult.Cancel
                                            freestyleForm.Close()
                                            e.SuppressKeyPress = True
                                            Return
                                        End If

                                        If e.KeyCode = System.Windows.Forms.Keys.P AndAlso e.Modifiers = System.Windows.Forms.Keys.Control AndAlso Not System.String.IsNullOrEmpty(options.LastPrompt) Then
                                            SharedMethods.InsertFreestyleTextAtCaret(promptTextBox, options.LastPrompt)
                                            e.SuppressKeyPress = True
                                        End If

                                    End Sub

                                ' =========================================================
                                ' PROMPT LIBRARY
                                ' =========================================================

                                If options.PromptLibraryEnabled AndAlso options.Context IsNot Nothing AndAlso options.Context.INI_PromptLib Then

                                    Dim promptLibraryPath As System.String = options.Context.INI_PromptLibPath
                                    Dim promptLibraryPathLocal As System.String = options.Context.INI_PromptLibPathLocal
                                    Dim promptLibraryContext As ISharedContext = options.Context

                                    AddHandler promptTextBox.KeyPress,
                                        Sub(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs)

                                            If e.KeyChar <> "/"c Then
                                                Return
                                            End If

                                            Dim slashAction As SharedMethods.PromptLibrarySlashAction = SharedMethods.HandlePromptLibrarySlash(promptTextBox, promptLibraryPath, promptLibraryPathLocal, promptLibraryContext, options.LastPrompt, True)

                                            If slashAction <> SharedMethods.PromptLibrarySlashAction.NotTriggered Then
                                                e.Handled = True
                                            End If

                                        End Sub

                                End If

                                ' =========================================================
                                ' INITIAL PREFIX
                                ' =========================================================

                                If promptTextBox.TextLength > 0 Then

                                    Dim initialMatch As System.Tuple(Of FreestylePromptMode, System.String) = SharedMethods.FindFreestylePromptMode(promptTextBox.Text, options.Modes)

                                    If initialMatch IsNot Nothing Then

                                        synchronizingMode = True

                                        Try
                                            modeCombo.SelectedItem = initialMatch.Item1
                                        Finally
                                            synchronizingMode = False
                                        End Try

                                    End If

                                End If

                                ' =========================================================
                                ' INITIAL SIZE
                                ' =========================================================

                                freestyleForm.PerformLayout()
                                updateLayout.Invoke()

                                collapsedClientHeight = freestyleForm.ClientSize.Height

                                Dim chromeWidth As System.Int32 = freestyleForm.Width - freestyleForm.ClientSize.Width
                                Dim chromeHeight As System.Int32 = freestyleForm.Height - freestyleForm.ClientSize.Height

                                freestyleForm.MinimumSize = New System.Drawing.Size(System.Math.Min(390 + chromeWidth, workingArea.Width), System.Math.Min(lineHeight * 21 + chromeHeight, workingArea.Height))

                                freestyleForm.Location = New System.Drawing.Point(workingArea.X + (workingArea.Width - freestyleForm.Width) \ 2, workingArea.Y + (workingArea.Height - freestyleForm.Height) \ 2)

                                ' =========================================================
                                ' WINDOW HANDLING
                                ' =========================================================

                                SharedMethods.AttachForeignForegroundWatchdog(freestyleForm)

                                AddHandler freestyleForm.Shown,
                                    Sub(sender As System.Object, e As System.EventArgs)

                                        updateLayout.Invoke()
                                        SharedMethods.ForceDialogToForeground(freestyleForm)
                                        promptTextBox.Focus()

                                        If promptTextBox.TextLength > 0 Then
                                            promptTextBox.SelectionStart = promptTextBox.TextLength
                                        End If

                                    End Sub

                                freestyleForm.TopMost = True
                                freestyleForm.Opacity = 1
                                freestyleForm.BringToFront()
                                freestyleForm.Focus()
                                freestyleForm.Activate()

                                Dim dialogResult As System.Windows.Forms.DialogResult
                                Dim owner As System.Windows.Forms.IWin32Window = SharedMethods.ResolveSameThreadDialogOwner()

                                If owner IsNot Nothing Then
                                    dialogResult = freestyleForm.ShowDialog(owner)
                                Else
                                    dialogResult = freestyleForm.ShowDialog()
                                End If

                                If dialogResult <> System.Windows.Forms.DialogResult.OK Then
                                    Return returnValue
                                End If

                                ' =========================================================
                                ' RESULT
                                ' =========================================================

                                returnValue.Accepted = True
                                returnValue.Prompt = promptTextBox.Text
                                returnValue.KnownPrefixes.AddRange(SharedMethods.GetFreestylePromptKnownPrefixes(options.Modes))

                                Dim finalSelectedMode As FreestylePromptMode = TryCast(modeCombo.SelectedItem, FreestylePromptMode)
                                Dim effectiveSelectedPrefix As System.String = quickSelectedPrefix.Trim()

                                If quickSelectedModeId.Length > 0 Then
                                    finalSelectedMode = SharedMethods.FindFreestylePromptModeById(quickSelectedModeId, options.Modes)
                                End If

                                If finalSelectedMode IsNot Nothing Then

                                    returnValue.SelectedModeId = finalSelectedMode.Id

                                    If effectiveSelectedPrefix.Length = 0 Then

                                        Dim finalMatch As System.Tuple(Of FreestylePromptMode, System.String) =
                                            SharedMethods.FindFreestylePromptMode(promptTextBox.Text, options.Modes)

                                        If finalMatch IsNot Nothing AndAlso finalMatch.Item1 Is finalSelectedMode Then
                                            effectiveSelectedPrefix = finalMatch.Item2
                                        Else
                                            effectiveSelectedPrefix = finalSelectedMode.Prefix
                                        End If

                                    End If

                                ElseIf quickSelectedModeId.Length > 0 Then

                                    returnValue.SelectedModeId = quickSelectedModeId

                                End If

                                returnValue.SelectedPrefix = effectiveSelectedPrefix

                                returnValue.PersistedState =
                                    SharedMethods.BuildFreestylePromptPersistedState(
                                        options.CallerId,
                                        options.PersistedState,
                                        returnValue.SelectedModeId,
                                        returnValue.SelectedPrefix,
                                        expanded,
                                        toggleCheckBoxes,
                                        argumentTextBoxes)

                                For Each optionCheckBox As System.Windows.Forms.CheckBox In toggleCheckBoxes

                                    If Not optionCheckBox.Checked Then
                                        Continue For
                                    End If

                                    Dim definition As FreestylePromptToggleOption = TryCast(optionCheckBox.Tag, FreestylePromptToggleOption)

                                    If definition Is Nothing Then
                                        Continue For
                                    End If

                                    If Not System.String.IsNullOrWhiteSpace(definition.Id) Then
                                        returnValue.SelectedOptionIds.Add(definition.Id)
                                    End If

                                    Dim argumentValue As System.String = System.String.Empty

                                    If argumentTextBoxes.ContainsKey(optionCheckBox) Then
                                        argumentValue = argumentTextBoxes(optionCheckBox).Text.Trim()
                                    End If

                                    Dim finalTrigger As System.String = definition.BuildTrigger(argumentValue)

                                    If Not System.String.IsNullOrWhiteSpace(finalTrigger) AndAlso Not returnValue.SelectedTriggers.Contains(finalTrigger) Then
                                        returnValue.SelectedTriggers.Add(finalTrigger)
                                    End If

                                Next

                                Return returnValue

                            End Using
                        End Using
                    End Using
                End Using

            Catch ex As System.Exception

                SharedMethods.ShowCustomMessageBox("The Freestyle dialog could not be displayed." & System.Environment.NewLine & System.Environment.NewLine & ex.Message, "Freestyle")

                Return returnValue

            End Try

        End Function


        Public Shared Function ComposeFreestylePrompt(ByVal result As FreestylePromptResult) As System.String

            If result Is Nothing OrElse Not result.Accepted Then
                Return System.String.Empty
            End If

            Dim prompt As System.String = If(result.Prompt, System.String.Empty).Trim()

            If Not System.String.IsNullOrWhiteSpace(result.SelectedPrefix) Then
                prompt = SharedMethods.ReplaceFreestyleLeadingPrefix(prompt, result.KnownPrefixes, result.SelectedPrefix)
            End If

            For Each trigger As System.String In result.SelectedTriggers

                If System.String.IsNullOrWhiteSpace(trigger) Then
                    Continue For
                End If

                If prompt.IndexOf(trigger, System.StringComparison.OrdinalIgnoreCase) < 0 Then

                    If prompt.Length > 0 Then
                        prompt &= " "
                    End If

                    prompt &= trigger.Trim()

                End If

            Next

            Return prompt.Trim()

        End Function

#End Region


    End Class
End Namespace
