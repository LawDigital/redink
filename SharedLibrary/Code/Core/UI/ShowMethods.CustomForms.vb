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
            ' Try Word first.
            Dim hwnd As IntPtr = FindWindow("OpusApp", Nothing)
            If hwnd <> IntPtr.Zero Then Return hwnd

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

            ' Configure the form and DPI support.
            Dim inputForm As New System.Windows.Forms.Form() With {
        .Text = title,
        .FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog,
        .StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen,
        .MinimizeBox = False,
        .MaximizeBox = False,
        .ShowInTaskbar = False,
        .KeyPreview = True,
        .AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font,
        .ClientSize = New System.Drawing.Size(450, 320),
        .MinimumSize = New System.Drawing.Size(450, 240)
    }
            inputForm.Font = New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)

            ' Use logo as icon.
            Dim bmp As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
            inputForm.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

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
            listBoxOptions.Items.AddRange(options.ToArray())
            listPanel.Controls.Add(listBoxOptions)

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

            ' Show dialog.
            inputForm.TopMost = True

            Dim hwndOwner As IntPtr = GetOfficeApplicationHwnd() ' Uses OpusApp/XLMAIN/rctrl_renwnd32
            If hwndOwner <> IntPtr.Zero Then
                inputForm.ShowDialog(New WindowWrapper(hwndOwner))
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
                                                    Optional OptionalButtons As System.Tuple(Of System.String, System.String, System.String)() = Nothing
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

            AddHandler okButton.Click, Sub()
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

            ' After AutoSize computed, clamp to screen, set MinimumSize (so buttons stay visible),
            ' disable AutoSize to allow user resizing, and center within the working area.
            AddHandler inputForm.Shown, Sub()
                                            ' Let AutoSize produce the preferred size first.
                                            inputForm.PerformLayout()

                                            Dim maxW As Integer = wa.Width - 40
                                            Dim maxH As Integer = wa.Height - 40

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

                                                ' Keep current width (already autosized) but clamp to screen.
                                                Dim newClientW As Integer = Math.Min(inputForm.ClientSize.Width, maxW)

                                                inputForm.ClientSize = New Size(newClientW, newClientH)
                                            Else
                                                ' SimpleInput: just clamp to screen.
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

            ' Show the dialog, optionally owned by Outlook.
            'Dim Result As DialogResult
            'If title.Contains("Browser") Then
            'Dim outlookApp As Object = CreateObject("Outlook.Application")
            'If outlookApp IsNot Nothing Then
            'Dim explorer As Object = outlookApp.GetType().InvokeMember("ActiveExplorer", BindingFlags.GetProperty, Nothing, outlookApp, Nothing)
            'If explorer IsNot Nothing Then
            'explorer.GetType().InvokeMember("WindowState", BindingFlags.SetProperty, Nothing, explorer, New Object() {1})
            'explorer.GetType().InvokeMember("Activate", BindingFlags.InvokeMethod, Nothing, explorer, Nothing)
            'End If
            'End If
            'inputForm.Opacity = 1
            'Dim outlookHwnd As IntPtr = FindWindow("rctrl_renwnd32", Nothing)
            'Result = inputForm.ShowDialog(New WindowWrapper(outlookHwnd))
            'Else
            'inputForm.Opacity = 1
            'Result = inputForm.ShowDialog()
            'End If

            ' Show the dialog, must be owned by Outlook (only then the title may contains "Browser").
            Dim Result As DialogResult
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
                Result = inputForm.ShowDialog(New WindowWrapper(outlookHwnd))
            Else
                inputForm.Opacity = 1
                Result = inputForm.ShowDialog()
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
        ''' <returns>1 for button1, 2 for button2, 3 for auto-close; otherwise 0 (initial value).</returns>
        Public Shared Function ShowCustomYesNoBox(
                        ByVal bodyText As String,
                        ByVal button1Text As String,
                        ByVal button2Text As String,
                        Optional header As String = AN,
                        Optional autoCloseSeconds As Integer? = Nothing,
                        Optional Defaulttext As String = "",
                        Optional extraButtonText As String = Nothing,
                        Optional extraButtonAction As System.Action = Nothing,
                        Optional CloseAfterExtra As Boolean = False
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
                .ShowInTaskbar = False,
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
            Dim clientWidth As Integer = Math.Max(MIN_WIDTH, targetWidth) - chromeWidth
            Dim clientHeight As Integer = bodyPanelHeight + buttonPanelHeight + PADDING + BUTTON_GAP

            messageForm.ClientSize = New Size(clientWidth, clientHeight)

            ' Set minimum size to ensure buttons always visible.
            Dim minButtonWidth As Integer = bottomFlow.PreferredSize.Width + 2 * PADDING
            messageForm.MinimumSize = New Size(
                Math.Max(MIN_WIDTH, minButtonWidth + chromeWidth),
                buttonPanelHeight + 100 + chromeHeight
            )

            ' Handle resize to update label wrapping.
            AddHandler messageForm.Resize,
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

            ' Show and return.
            messageForm.TopMost = True
            messageForm.Opacity = 1
            messageForm.BringToFront()
            messageForm.Focus()
            messageForm.Activate()

            AddHandler messageForm.Shown,
                Sub(sender, e)
                    messageForm.TopMost = False
                    messageForm.TopMost = True
                    messageForm.Activate()
                    messageForm.BringToFront()
                End Sub

            messageForm.ShowDialog()
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
                                messageForm.TopMost = False  ' Reset first.
                                messageForm.TopMost = True   ' Then set again.
                                messageForm.Activate()
                                messageForm.BringToFront()
                            End Sub
                    messageForm.ShowDialog()
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
                            messageForm.TopMost = False  ' Reset first.
                            messageForm.TopMost = True   ' Then set again.
                            messageForm.Activate()
                            messageForm.BringToFront()
                        End Sub

                messageForm.Opacity = 1
                messageForm.ShowDialog()
            End If
        End Sub






        ''' <summary>
        ''' Shows a modal RichTextBox-based message dialog (RTF content) with optional auto-close.
        ''' </summary>
        ''' <param name="bodyText">RTF content assigned to <see cref="RichTextBox.Rtf"/>.</param>
        ''' <param name="header">Dialog title. Defaults to <c>AN</c> if empty/whitespace.</param>
        ''' <param name="autoCloseSeconds">If set, closes the dialog after this many seconds.</param>
        ''' <param name="Defaulttext">Suffix appended to the countdown label text.</param>
        Public Shared Sub ShowRTFCustomMessageBox(ByVal bodyText As String, Optional header As String = AN, Optional autoCloseSeconds As Integer? = Nothing, Optional Defaulttext As String = " - execution continues meanwhile")

            Dim RTFMessageForm As New System.Windows.Forms.Form()
            Dim bodyLabel As New System.Windows.Forms.RichTextBox()
            Dim okButton As New System.Windows.Forms.Button()
            Dim countdownLabel As New System.Windows.Forms.Label()
            Dim Truncated As Boolean = False

            If String.IsNullOrWhiteSpace(header) Then header = AN

            ' Form attributes.
            RTFMessageForm.Opacity = 0
            RTFMessageForm.Text = header
            RTFMessageForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
            RTFMessageForm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            RTFMessageForm.MaximizeBox = True
            RTFMessageForm.MinimizeBox = True
            RTFMessageForm.ShowInTaskbar = False
            RTFMessageForm.TopMost = True
            RTFMessageForm.KeyPreview = True

            ' Autoscale for fonts & DPI.
            RTFMessageForm.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
            RTFMessageForm.AutoScaleDimensions = New System.Drawing.SizeF(96.0F, 96.0F)

            RTFMessageForm.MinimumSize = New System.Drawing.Size(650, 335)

            ' Icon.
            Dim bmp As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
            RTFMessageForm.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

            ' Standard font.
            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)

            ' Body RTF box.
            bodyLabel.Font = standardFont
            bodyLabel.ReadOnly = True
            bodyLabel.BorderStyle = System.Windows.Forms.BorderStyle.None
            bodyLabel.BackColor = RTFMessageForm.BackColor
            bodyLabel.TabStop = False
            bodyLabel.Rtf = bodyText
            bodyLabel.Location = New System.Drawing.Point(20, 20)
            bodyLabel.Width = 600
            bodyLabel.Height = 200
            ' Anchor to all sides so it resizes with the form.
            bodyLabel.Anchor = System.Windows.Forms.AnchorStyles.Top _
                     Or System.Windows.Forms.AnchorStyles.Left _
                     Or System.Windows.Forms.AnchorStyles.Right _
                     Or System.Windows.Forms.AnchorStyles.Bottom
            RTFMessageForm.Controls.Add(bodyLabel)


            ' OK button & countdown label setup.
            okButton.Font = standardFont
            okButton.Text = "OK"
            okButton.AutoSize = True

            countdownLabel.Font = standardFont
            countdownLabel.AutoSize = True

            ' Bottom panel to hold button + countdown, docked so it moves with resizing.
            Dim bottomPanel As New System.Windows.Forms.Panel()
            bottomPanel.Dock = System.Windows.Forms.DockStyle.Bottom
            bottomPanel.Padding = New System.Windows.Forms.Padding(20)  ' 20px padding on all sides.
            bottomPanel.Height = okButton.PreferredSize.Height + bottomPanel.Padding.Top + bottomPanel.Padding.Bottom
            RTFMessageForm.Controls.Add(bottomPanel)

            ' Add controls into panel.
            bottomPanel.Controls.Add(okButton)
            bottomPanel.Controls.Add(countdownLabel)
            okButton.Location = New System.Drawing.Point(bottomPanel.Padding.Left, bottomPanel.Padding.Top)
            countdownLabel.Location = New System.Drawing.Point(okButton.Right + 10, bottomPanel.Padding.Top)

            ' Ensure bodyLabel resizes when form is resized.
            AddHandler RTFMessageForm.Resize, Sub(sender As Object, e As EventArgs)
                                                  Dim availableWidth As Integer = RTFMessageForm.ClientSize.Width - bodyLabel.Left - 20
                                                  Dim availableHeight As Integer = RTFMessageForm.ClientSize.Height - bottomPanel.Height - bodyLabel.Top - 20
                                                  bodyLabel.Size = New System.Drawing.Size(availableWidth, availableHeight)
                                              End Sub

            ' Handlers.
            Dim userClicked As Boolean = False
            AddHandler okButton.Click, Sub(sender As Object, e As EventArgs)
                                           userClicked = True
                                           RTFMessageForm.Close()
                                           RTFMessageForm = Nothing
                                       End Sub
            AddHandler RTFMessageForm.KeyDown, Sub(sender As Object, e As System.Windows.Forms.KeyEventArgs)
                                                   If e.KeyCode = System.Windows.Forms.Keys.Escape Then
                                                       userClicked = True
                                                       RTFMessageForm.Close()
                                                       RTFMessageForm = Nothing
                                                       e.SuppressKeyPress = True
                                                   End If
                                               End Sub
            AddHandler RTFMessageForm.Shown, Sub(sender As Object, e As EventArgs)
                                                 ' Trigger initial resize layout.
                                                 RTFMessageForm.PerformLayout()
                                                 RTFMessageForm.Activate()
                                             End Sub

            ' Initial form sizing: ensure 20px padding around button and RTF label sizing.
            Dim formWidth As Integer = Math.Max(RTFMessageForm.MinimumSize.Width, bodyLabel.Width + 40)
            Dim formHeight As Integer = Math.Max(RTFMessageForm.MinimumSize.Height,
                                         bodyLabel.Bottom + 20 + bottomPanel.Height)
            RTFMessageForm.ClientSize = New System.Drawing.Size(formWidth, formHeight)

            ' Auto-close timer.
            If autoCloseSeconds.HasValue AndAlso autoCloseSeconds > 0 Then
                Dim remainingTime As Integer = autoCloseSeconds.Value
                countdownLabel.Text = $"(closes in {remainingTime} seconds{Defaulttext})"

                Dim timer As New System.Windows.Forms.Timer()
                timer.Interval = 1000
                AddHandler timer.Tick, Sub(sender As Object, e As EventArgs)
                                           remainingTime -= 1
                                           If remainingTime > 0 Then
                                               countdownLabel.Text = $"(closes in {remainingTime} seconds{Defaulttext})"
                                           Else
                                               timer.Stop()
                                               If Not userClicked Then
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
                                            RTFMessageForm.TopMost = False  ' Reset first.
                                            RTFMessageForm.TopMost = True   ' Then set again.
                                            RTFMessageForm.Activate()
                                            RTFMessageForm.BringToFront()
                                        End Sub

                RTFMessageForm.Opacity = 1
                RTFMessageForm.Show()
                RTFMessageForm.BringToFront()
                RTFMessageForm.Activate()
                System.Windows.Forms.Application.DoEvents()
            Else


                RTFMessageForm.BringToFront()
                RTFMessageForm.Focus()
                RTFMessageForm.Activate()

                AddHandler RTFMessageForm.Shown,
                                        Sub(sender, e)
                                            RTFMessageForm.TopMost = False  ' Reset first.
                                            RTFMessageForm.TopMost = True   ' Then set again.
                                            RTFMessageForm.Activate()
                                            RTFMessageForm.BringToFront()
                                        End Sub

                RTFMessageForm.Opacity = 1
                RTFMessageForm.ShowDialog()
            End If

        End Sub


        ''' <summary>
        ''' Shows an HTML message dialog using a WinForms WebBrowser control on an STA thread.
        ''' </summary>
        ''' <param name="bodyText">HTML assigned to <see cref="WebBrowser.DocumentText"/>.</param>
        ''' <param name="header">Dialog title. Defaults to <c>AN</c>.</param>
        ''' <param name="Defaulttext">Unused parameter (kept for signature compatibility).</param>
        ''' <param name="extraButtonText">Optional extra button text.</param>
        ''' <param name="extraButtonAction">Action invoked when the extra button is clicked.</param>
        ''' <param name="CloseAfterExtra">If <c>True</c>, closes the dialog after invoking the extra action.</param>
        Public Shared Sub ShowHTMLCustomMessageBox(
    ByVal bodyText As String,
    Optional header As String = AN,
    Optional Defaulttext As String = " - execution continues meanwhile",
    Optional extraButtonText As String = Nothing,
    Optional extraButtonAction As System.Action = Nothing,
    Optional CloseAfterExtra As Boolean = False
)
            Dim t As New Thread(Sub()
                                    ' Create and configure form.
                                    Dim HTMLMessageForm As New System.Windows.Forms.Form() With {
                                .Opacity = 0,
                                .Text = header,
                                .FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable,
                                .StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen,
                                .MaximizeBox = True,
                                .MinimizeBox = True,
                                .ShowInTaskbar = True,
                                .TopMost = False,
                                .KeyPreview = True,
                                .MinimumSize = New System.Drawing.Size(800, 500),
                                .AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
                            }

                                    ' Header fallback.
                                    If String.IsNullOrWhiteSpace(header) Then
                                        HTMLMessageForm.Text = AN
                                    End If

                                    ' Set the icon.
                                    Dim bmp As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                                    HTMLMessageForm.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

                                    ' Standard font.
                                    Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
                                    HTMLMessageForm.Font = standardFont

                                    ' WebBrowser with margin.
                                    Dim htmlBrowser As New System.Windows.Forms.WebBrowser() With {
                                .AllowNavigation = False,
                                .WebBrowserShortcutsEnabled = False,
                                .ScrollBarsEnabled = True,
                                .ScriptErrorsSuppressed = True,
                                .DocumentText = bodyText,
                                .Dock = System.Windows.Forms.DockStyle.Fill,
                                .BackColor = HTMLMessageForm.BackColor,
                                .Margin = New System.Windows.Forms.Padding(20)
                            }
                                    AddHandler htmlBrowser.DocumentCompleted, Sub(sender2, e2)
                                                                                  If htmlBrowser.Document?.Body IsNot Nothing Then
                                                                                      ' Body style with margin matching the form.
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

                                    ' Activate on shown.
                                    AddHandler HTMLMessageForm.Shown, Sub(sender2, e2)
                                                                          HTMLMessageForm.Activate()
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

                                    ' Optional extra button.
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
                                                    ' Execute the action. Recursive calls will spawn their own STA threads.
                                                    extraButtonAction.Invoke()
                                                Catch ex As System.Exception
                                                    ' Swallow to keep dialog functional.
                                                End Try
                                                If CloseAfterExtra Then HTMLMessageForm.Close()
                                            End Sub

                                        bottomFlow.Controls.Add(extraButton)
                                    End If

                                    ' Compose form.
                                    HTMLMessageForm.Controls.Add(htmlBrowser)
                                    HTMLMessageForm.Controls.Add(bottomFlow)

                                    HTMLMessageForm.BringToFront()
                                    HTMLMessageForm.Focus()
                                    HTMLMessageForm.Activate()

                                    AddHandler HTMLMessageForm.Shown,
                                        Sub(sender, e)
                                            HTMLMessageForm.TopMost = False  ' Reset first.
                                            HTMLMessageForm.TopMost = True   ' Then set again.
                                            HTMLMessageForm.Activate()
                                            HTMLMessageForm.BringToFront()
                                        End Sub

                                    HTMLMessageForm.Opacity = 1

                                    ' Show dialog.
                                    HTMLMessageForm.ShowDialog()
                                End Sub)
            t.SetApartmentState(System.Threading.ApartmentState.STA)
            t.Start()
        End Sub


        ''' <summary>
        ''' Legacy HTML message dialog variant without extra button support; kept for compatibility.
        ''' </summary>
        ''' <param name="bodyText">HTML assigned to <see cref="WebBrowser.DocumentText"/>.</param>
        ''' <param name="header">Dialog title. Defaults to <c>AN</c>.</param>
        ''' <param name="Defaulttext">Unused parameter (kept for signature compatibility).</param>
        Public Shared Sub oldShowHTMLCustomMessageBox(
    ByVal bodyText As String,
    Optional header As String = AN,
    Optional Defaulttext As String = " - execution continues meanwhile"
)
            Dim t As New Thread(Sub()
                                    ' Create and configure form.
                                    Dim HTMLMessageForm As New System.Windows.Forms.Form() With {
                                .Opacity = 0,
                                .Text = header,
                                .FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable,
                                .StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen,
                                .MaximizeBox = True,
                                .MinimizeBox = True,
                                .ShowInTaskbar = True,
                                .TopMost = False,
                                .KeyPreview = True,
                                .MinimumSize = New System.Drawing.Size(800, 500),
                                .AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
                            }

                                    ' Header fallback.
                                    If String.IsNullOrWhiteSpace(header) Then
                                        HTMLMessageForm.Text = AN
                                    End If

                                    ' Set the icon.
                                    Dim bmp As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                                    HTMLMessageForm.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

                                    ' Standard font.
                                    Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
                                    HTMLMessageForm.Font = standardFont

                                    ' WebBrowser with margin.
                                    Dim htmlBrowser As New System.Windows.Forms.WebBrowser() With {
                                .AllowNavigation = False,
                                .WebBrowserShortcutsEnabled = False,
                                .ScrollBarsEnabled = True,
                                .ScriptErrorsSuppressed = True,
                                .DocumentText = bodyText,
                                .Dock = System.Windows.Forms.DockStyle.Fill,
                                .BackColor = HTMLMessageForm.BackColor,
                                .Margin = New System.Windows.Forms.Padding(20)
                            }
                                    AddHandler htmlBrowser.DocumentCompleted, Sub(sender2, e2)
                                                                                  If htmlBrowser.Document?.Body IsNot Nothing Then
                                                                                      ' Body style with margin matching the form.
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
                                .Margin = New System.Windows.Forms.Padding(0) ' No extra spacing here.
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

                                    ' Activate on shown.
                                    AddHandler HTMLMessageForm.Shown, Sub(sender2, e2)
                                                                          HTMLMessageForm.Activate()
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

                                    ' Compose form.
                                    HTMLMessageForm.Controls.Add(htmlBrowser)
                                    HTMLMessageForm.Controls.Add(bottomFlow)

                                    HTMLMessageForm.BringToFront()
                                    HTMLMessageForm.Focus()
                                    HTMLMessageForm.Activate()

                                    AddHandler HTMLMessageForm.Shown,
                                        Sub(sender, e)
                                            HTMLMessageForm.TopMost = False  ' Reset first.
                                            HTMLMessageForm.TopMost = True   ' Then set again.
                                            HTMLMessageForm.Activate()
                                            HTMLMessageForm.BringToFront()
                                        End Sub

                                    HTMLMessageForm.Opacity = 1

                                    ' Show dialog.
                                    HTMLMessageForm.ShowDialog()
                                End Sub)
            t.SetApartmentState(System.Threading.ApartmentState.STA)
            t.Start()
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

                Dim lbl As New System.Windows.Forms.Label() With {
                    .Text = param.Name & ":",
                    .AutoSize = True,
                    .Anchor = AnchorStyles.Left,
                    .Margin = New Padding(0, 0, 8, 8)
                }
                mainLayout.Controls.Add(lbl, 0, i + 1)

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
                    Dim txt As New TextBox() With {
                        .Text = rawValue.ToString(),
                        .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
                        .Margin = New Padding(0, 0, 0, 8)
                    }
                    If TypeOf rawValue Is String Then
                        txt.MinimumSize = New Size(400, 0)
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
                mainLayout.Controls.Add(ctrl, 1, i + 1)
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

            Dim result = inputForm.ShowDialog()

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
            Optional PreserveLiterals As Boolean = False
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
            returnValue = If(NoRTF, bodyTextBox.Text, bodyTextBox.Rtf)
            styledForm.DialogResult = DialogResult.OK
            styledForm.Close()
        End Sub

            AddHandler btnOriginal.Click,
        Sub()
            returnValue = If(NoRTF, OriginalText, If(rtf, bodyText))
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
            styledForm.BringToFront()
            styledForm.Focus()
            styledForm.Activate()

            AddHandler styledForm.Shown,
                    Sub(sender, e)
                        styledForm.TopMost = False  ' Reset first.
                        styledForm.TopMost = True   ' Then set again.
                        styledForm.Activate()
                        styledForm.BringToFront()
                    End Sub

            If parentWindowHwnd <> IntPtr.Zero Then
                styledForm.ShowDialog(New WindowWrapper(parentWindowHwnd))
            ElseIf Getfocus Then
                Dim officeHwnd As IntPtr = GetOfficeApplicationHwnd()
                If officeHwnd <> IntPtr.Zero Then
                    styledForm.ShowDialog(New WindowWrapper(officeHwnd))
                Else
                    styledForm.ShowDialog()
                End If
            Else
                styledForm.ShowDialog()
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

            ' Important: parameterless constructor (required for "New InputParameter() With {...}").
            Public Sub New()
                Me.Options = New System.Collections.Generic.List(Of System.String)()
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



    End Class
End Namespace