' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.PromptLibrary.vb
' Purpose: Loads prompt library entries from one or two text files and provides
'          a WinForms UI to select a prompt with preview and output-format options.
'
' Prompt Library File Format (line-based):
'  - Empty lines are ignored.
'  - Comment lines starting with ";" are ignored.
'  - Each entry is "Title|Prompt". The prompt may contain additional "|" characters;
'    everything after the first "|" is treated as part of the prompt.
'
' Behavior:
'  - Paths are resolved via ExpandEnvironmentVariables.
'  - Dual-source mode: loads local and central files independently, combines them
'    with local entries first, and optionally appends a title suffix (e.g. " (local)").
'  - Edit workflow: the user can open an editor for the local file (if configured),
'    otherwise the central file, and the lists are reloaded afterwards.
'  - Selection UI: shows titles and a prompt preview; Enter confirms the current selection.
'  - Output options: mutually exclusive checkboxes for markup / bubbles / window output.
'  - Context sync: best-effort copy of the combined titles and prompts into ISharedContext.
'
' Return Values:
'  - ShowPromptSelector / oldShowPromptSelector return a tuple:
'      (SelectedPrompt, MarkupSelected, BubblesSelected, WindowSelected)
'    or ("", False, False, False) if canceled or no valid selection exists.
' =============================================================================

Option Strict On
Option Explicit On
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary

    Partial Public Class SharedMethods

        ''' <summary>
        ''' Shows a prompt selection dialog that loads prompts from an optional local file and a central file,
        ''' displays a title list with prompt preview, and returns the selected prompt and output options.
        ''' </summary>
        ''' <param name="filePath">Central prompt library file path (environment variables are expanded).</param>
        ''' <param name="filepathlocal">Optional local prompt library file path (environment variables are expanded).</param>
        ''' <param name="enableMarkup">If <c>True</c>, the markup checkbox is enabled; if passed as <c>Nothing</c>, it is hidden and forced <c>False</c>.</param>
        ''' <param name="enableBubbles">If <c>True</c>, the bubbles checkbox is enabled; if passed as <c>Nothing</c>, it is hidden and forced <c>False</c>.</param>
        ''' <param name="Context">Optional context to receive the combined prompt titles and prompt texts (best-effort).</param>
        ''' <returns>
        ''' (SelectedPrompt, MarkupSelected, BubblesSelected, WindowSelected), or ("", False, False, False) if canceled or no prompts exist.
        ''' </returns>
        Public Shared Function ShowPromptSelector(filePath As String, filepathlocal As String, enableMarkup As Boolean, enableBubbles As Boolean, Context As ISharedContext) As (String, Boolean, Boolean, Boolean)

            Dim centralPath As String = ExpandEnvironmentVariables(filePath)
            Dim localPath As String = ExpandEnvironmentVariables(filepathlocal)
            Dim hasLocal As Boolean = Not String.IsNullOrWhiteSpace(localPath)

            ' Load prompts from both files independently (local first), ignore missing/non-existing silently
            Dim localTitles As New List(Of String)()
            Dim localPrompts As New List(Of String)()
            Dim centralTitles As New List(Of String)()
            Dim centralPrompts As New List(Of String)()

            LoadPromptsIntoLists(localPath, localTitles, localPrompts, " (local)")
            LoadPromptsIntoLists(centralPath, centralTitles, centralPrompts, Nothing)

            Dim combinedTitles As New List(Of String)()
            Dim combinedPrompts As New List(Of String)()

            ' Local first
            combinedTitles.AddRange(localTitles)
            combinedPrompts.AddRange(localPrompts)

            ' Then central
            combinedTitles.AddRange(centralTitles)
            combinedPrompts.AddRange(centralPrompts)

            ' Optionally keep Context in sync with what the user sees
            Try
                If Context IsNot Nothing Then
                    If Context.PromptTitles Is Nothing Then Context.PromptTitles = New List(Of String)()
                    If Context.PromptLibrary Is Nothing Then Context.PromptLibrary = New List(Of String)()
                    Context.PromptTitles.Clear()
                    Context.PromptLibrary.Clear()
                    Context.PromptTitles.AddRange(combinedTitles)
                    Context.PromptLibrary.AddRange(combinedPrompts)
                End If
            Catch
                ' Best-effort only
            End Try

            Dim NoBubbles As Boolean = False
            Dim NoMarkup As Boolean = False

            ' If enableMarkup is not used; comparing to Nothing treats Nothing as False.
            If enableMarkup = Nothing Then
                NoMarkup = True
                enableMarkup = False
            End If

            ' If enableBubbles is not used; comparing to Nothing treats Nothing as False.
            If enableBubbles = Nothing Then
                NoBubbles = True
                enableBubbles = False
            End If

            If combinedPrompts.Count = 0 Then
                ShowCustomMessageBox("No prompts have been found in the configured prompt library files.")
                Return ("", False, False, False)
            End If

            ' --- Form -----------------------------------------------------------------
            Dim settingsForm As New System.Windows.Forms.Form With {
                    .Text = "Select Prompt",
                    .AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi,
                    .AutoScaleDimensions = New System.Drawing.SizeF(96.0F, 96.0F),
                    .AutoSize = False,
                    .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowOnly,
                    .StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen,
                    .Padding = New System.Windows.Forms.Padding(10),
                    .MinimizeBox = True,
                    .MaximizeBox = True
                }
            settingsForm.MinimumSize = New System.Drawing.Size(900, 650)

            Dim bmp As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
            settingsForm.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
            settingsForm.Font = standardFont

            ' --- Layout grid ----------------------------------------------------------
            Dim layout As New System.Windows.Forms.TableLayoutPanel With {
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 3,
                .Padding = New System.Windows.Forms.Padding(10)
            }
            layout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0F))
            layout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0F))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 70.0F))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            settingsForm.Controls.Add(layout)

            ' --- Selector --------------------------------------------------------------
            Dim titleListBox As New System.Windows.Forms.ListBox With {
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .Margin = New System.Windows.Forms.Padding(10)
            }
            titleListBox.Items.AddRange(combinedTitles.ToArray())
            layout.Controls.Add(titleListBox, 0, 0)

            ' --- Preview ---------------------------------------------------------------
            Dim promptTextBox As New System.Windows.Forms.TextBox With {
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .Multiline = True,
                .ReadOnly = True,
                .ScrollBars = System.Windows.Forms.ScrollBars.Vertical,
                .Margin = New System.Windows.Forms.Padding(10)
            }
            layout.Controls.Add(promptTextBox, 1, 0)

            If combinedTitles.Count > 0 Then
                titleListBox.SelectedIndex = 0
                promptTextBox.Text = combinedPrompts(0).Replace("\n", vbCrLf)
            End If

            ' Updates the preview on selection changes.
            AddHandler titleListBox.SelectedIndexChanged,
                Sub()
                    Dim selectedIndex = titleListBox.SelectedIndex
                    If selectedIndex >= 0 AndAlso selectedIndex < combinedPrompts.Count Then
                        Dim selectedPrompt = combinedPrompts(selectedIndex).Replace("\n", vbCrLf)
                        promptTextBox.Text = selectedPrompt
                    End If
                End Sub

            ' Confirms the dialog on Enter.
            AddHandler titleListBox.KeyDown,
                Sub(sender As Object, e As System.Windows.Forms.KeyEventArgs)
                    If e.KeyCode = System.Windows.Forms.Keys.Enter Then
                        settingsForm.DialogResult = System.Windows.Forms.DialogResult.OK
                        settingsForm.Close()
                    End If
                End Sub

            ' --- Checkboxes (wrapping) ------------------------------------------------
            Dim checkboxPanel As New System.Windows.Forms.FlowLayoutPanel With {
                .FlowDirection = System.Windows.Forms.FlowDirection.TopDown,
                .WrapContents = False,
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .Margin = New System.Windows.Forms.Padding(10),
                .AutoSize = True,
                .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
            }
            layout.Controls.Add(checkboxPanel, 0, 1)

            Dim markupCheckbox As New System.Windows.Forms.CheckBox With {
                .Text = "The output shall be provided as a markup",
                .AutoSize = True,
                .Enabled = enableMarkup,
                .Visible = Not NoMarkup,
                .Margin = New System.Windows.Forms.Padding(3, 3, 3, 6)
            }

            Dim clipboardCheckbox As New System.Windows.Forms.CheckBox With {
                .Text = "The output shall be shown in a window",
                .AutoSize = True,
                .Checked = True,
                .Margin = New System.Windows.Forms.Padding(3, 3, 3, 6)
            }

            Dim bubblesCheckbox As New System.Windows.Forms.CheckBox With {
                .Text = "The output shall be in the form of bubbles",
                .AutoSize = True,
                .Enabled = enableBubbles,
                .Visible = Not NoBubbles,
                .Margin = New System.Windows.Forms.Padding(3, 3, 3, 6)
            }

            checkboxPanel.Controls.Add(markupCheckbox)
            checkboxPanel.Controls.Add(clipboardCheckbox)
            checkboxPanel.Controls.Add(bubblesCheckbox)

            ' Applies MaximumSize to trigger line wrapping based on the left grid cell width.
            Dim ApplyCheckboxWrap As System.Action =
                Sub()
                    Dim cellWidthLeft As Integer = CInt((layout.ClientSize.Width - layout.Padding.Horizontal) * layout.ColumnStyles(0).Width / 100.0F) - 20
                    If cellWidthLeft < 100 Then cellWidthLeft = 100
                    markupCheckbox.MaximumSize = New System.Drawing.Size(cellWidthLeft, 0)
                    clipboardCheckbox.MaximumSize = New System.Drawing.Size(cellWidthLeft, 0)
                    bubblesCheckbox.MaximumSize = New System.Drawing.Size(cellWidthLeft, 0)
                End Sub
            AddHandler layout.SizeChanged, Sub() ApplyCheckboxWrap()

            ' Mutual exclusivity
            AddHandler markupCheckbox.CheckedChanged, Sub() If markupCheckbox.Checked Then bubblesCheckbox.Checked = False : clipboardCheckbox.Checked = False
            AddHandler bubblesCheckbox.CheckedChanged, Sub() If bubblesCheckbox.Checked Then markupCheckbox.Checked = False : clipboardCheckbox.Checked = False
            AddHandler clipboardCheckbox.CheckedChanged, Sub() If clipboardCheckbox.Checked Then markupCheckbox.Checked = False : bubblesCheckbox.Checked = False

            ' --- Source label (wrapping) ----------------------------------------------
            Dim sourceText As String
            If hasLocal Then
                sourceText = $"Source: {localPath} (local, editable) | {centralPath} (central)"
            Else
                sourceText = $"Source: {centralPath} (central, editable)"
            End If

            Dim filePathLabel As New System.Windows.Forms.Label With {
                .Text = sourceText,
                .AutoSize = True,
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .Margin = New System.Windows.Forms.Padding(10),
                .AutoEllipsis = False
            }
            layout.Controls.Add(filePathLabel, 1, 1)

            ' Applies MaximumSize to trigger line wrapping based on the right grid cell width.
            Dim ApplyFilePathWrap As System.Action =
                Sub()
                    Dim cellWidthRight As Integer = CInt((layout.ClientSize.Width - layout.Padding.Horizontal) * layout.ColumnStyles(1).Width / 100.0F) - 20
                    If cellWidthRight < 100 Then cellWidthRight = 100
                    filePathLabel.MaximumSize = New System.Drawing.Size(cellWidthRight, 0)
                End Sub
            AddHandler layout.SizeChanged, Sub() ApplyFilePathWrap()

            ' --- Buttons (LEFT aligned, OK | Cancel | Edit) ---------------------------
            Dim buttonPanel As New System.Windows.Forms.FlowLayoutPanel With {
                .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
                .WrapContents = False,
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .AutoSize = True,
                .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                .Margin = New System.Windows.Forms.Padding(4),
                .Padding = New System.Windows.Forms.Padding(4)
            }
            layout.Controls.Add(buttonPanel, 0, 2)
            layout.SetColumnSpan(buttonPanel, 2)

            Dim okButton As New System.Windows.Forms.Button With {
                .Text = "OK",
                .AutoSize = True,
                .DialogResult = System.Windows.Forms.DialogResult.OK,
                .Margin = New System.Windows.Forms.Padding(3),
                .Padding = New System.Windows.Forms.Padding(8, 4, 8, 4)
            }
            Dim cancelButton As New System.Windows.Forms.Button With {
                .Text = "Cancel",
                .AutoSize = True,
                .DialogResult = System.Windows.Forms.DialogResult.Cancel,
                .Margin = New System.Windows.Forms.Padding(3),
                .Padding = New System.Windows.Forms.Padding(8, 4, 8, 4)
            }
            Dim editButton As New System.Windows.Forms.Button With {
                .Text = "Edit",
                .AutoSize = True,
                .Margin = New System.Windows.Forms.Padding(3),
                .Padding = New System.Windows.Forms.Padding(8, 4, 8, 4)
            }

            buttonPanel.Controls.Add(okButton)
            buttonPanel.Controls.Add(cancelButton)
            buttonPanel.Controls.Add(editButton)

            ' --- Edit button: edit ONLY local if defined, else central; then reload both
            AddHandler editButton.Click,
                Sub()
                    Dim target As String = If(hasLocal, localPath, centralPath)
                    Dim targetKind As String = If(hasLocal, "local", "central")
                    ShowTextFileEditor(target, $"You can now edit your {targetKind} prompts (stored at {target}). Make sure that on each line, the description and the prompt is separated by a '|'; you can use ';' for indicating comments.")

                    ' Reload both sources after editing
                    localTitles.Clear() : localPrompts.Clear()
                    centralTitles.Clear() : centralPrompts.Clear()
                    LoadPromptsIntoLists(localPath, localTitles, localPrompts, " (local)")
                    LoadPromptsIntoLists(centralPath, centralTitles, centralPrompts, Nothing)

                    combinedTitles.Clear() : combinedPrompts.Clear()
                    combinedTitles.AddRange(localTitles) : combinedPrompts.AddRange(localPrompts)
                    combinedTitles.AddRange(centralTitles) : combinedPrompts.AddRange(centralPrompts)

                    ' Keep Context synced with the combined view
                    Try
                        If Context IsNot Nothing Then
                            Context.PromptTitles.Clear()
                            Context.PromptLibrary.Clear()
                            Context.PromptTitles.AddRange(combinedTitles)
                            Context.PromptLibrary.AddRange(combinedPrompts)
                        End If
                    Catch
                        ' Best-effort only
                    End Try

                    titleListBox.Items.Clear()
                    titleListBox.Items.AddRange(combinedTitles.ToArray())

                    If combinedTitles.Count > 0 Then
                        titleListBox.SelectedIndex = 0
                        promptTextBox.Text = combinedPrompts(0).Replace("\n", vbCrLf)
                    Else
                        promptTextBox.Clear()
                    End If

                    titleListBox.Focus()
                End Sub

            ApplyCheckboxWrap()
            ApplyFilePathWrap()

            Dim result As System.Windows.Forms.DialogResult = settingsForm.ShowDialog()

            If result = System.Windows.Forms.DialogResult.OK Then
                Dim selectedIndex = titleListBox.SelectedIndex
                If selectedIndex >= 0 AndAlso selectedIndex < combinedPrompts.Count Then
                    Return (
                        combinedPrompts(selectedIndex),
                        markupCheckbox.Checked,
                        bubblesCheckbox.Checked,
                        clipboardCheckbox.Checked
                    )
                End If
            End If

            Return ("", False, False, False)
        End Function

        ' Helper: read prompts from a single file into provided lists; ignore missing files silently.
        ' If titleSuffix is provided (e.g., " (local)"), it is appended to every title from this file.
        ''' <summary>
        ''' Loads prompts from a single file into the provided title and prompt lists.
        ''' Missing/non-existing files are ignored and errors are swallowed.
        ''' </summary>
        ''' <param name="filePath">Prompt library file path (environment variables are expanded).</param>
        ''' <param name="titles">Destination list for prompt titles.</param>
        ''' <param name="prompts">Destination list for prompt texts.</param>
        ''' <param name="titleSuffix">Optional suffix appended to each title loaded from this file.</param>
        Private Shared Sub LoadPromptsIntoLists(filePath As String,
                                               titles As List(Of String),
                                               prompts As List(Of String),
                                               Optional titleSuffix As String = Nothing)
            Try
                If String.IsNullOrWhiteSpace(filePath) Then Return
                filePath = ExpandEnvironmentVariables(filePath)
                If Not System.IO.File.Exists(filePath) Then Return

                Dim lines = System.IO.File.ReadAllLines(filePath)
                For Each line As String In lines
                    Dim trimmedLine = line.Trim()
                    If trimmedLine.Length = 0 OrElse trimmedLine.StartsWith(";") Then Continue For

                    Dim parts = trimmedLine.Split("|"c)
                    If parts.Length >= 2 Then
                        Dim title = parts(0).Trim()
                        Dim prompt As String
                        If parts.Length = 2 Then
                            prompt = parts(1).Trim()
                        Else
                            ' Avoid LINQ; keep everything after the first '|' intact
                            prompt = String.Join("|", parts, 1, parts.Length - 1).Trim()
                        End If
                        If Not String.IsNullOrEmpty(titleSuffix) Then title &= titleSuffix

                        titles.Add(title)
                        prompts.Add(prompt)
                    End If
                Next
            Catch
                ' Swallow errors to avoid noisy UX in dual-source mode
            End Try
        End Sub


        ''' <summary>
        ''' Legacy prompt selector that loads prompts from a single file via LoadPrompts and uses Context for storage.
        ''' </summary>
        ''' <param name="filePath">Prompt library file path (environment variables are expanded).</param>
        ''' <param name="filepathlocal">Unused in this legacy implementation.</param>
        ''' <param name="enableMarkup">If <c>True</c>, the markup checkbox is enabled; if passed as <c>Nothing</c>, it is hidden and forced <c>False</c>.</param>
        ''' <param name="enableBubbles">If <c>True</c>, the bubbles checkbox is enabled; if passed as <c>Nothing</c>, it is hidden and forced <c>False</c>.</param>
        ''' <param name="Context">Context receiving prompt titles and prompts loaded from the file.</param>
        ''' <returns>
        ''' (SelectedPrompt, MarkupSelected, BubblesSelected, WindowSelected), or ("", False, False, False) if canceled or no valid selection exists.
        ''' </returns>
        Public Shared Function oldShowPromptSelector(filePath As String, filepathlocal As String, enableMarkup As Boolean, enableBubbles As Boolean, Context As ISharedContext) As (String, Boolean, Boolean, Boolean)

            filePath = ExpandEnvironmentVariables(filePath)

            Dim LoadResult = LoadPrompts(filePath, Context)
            Dim NoBubbles As Boolean = False
            Dim NoMarkup As Boolean = False

            ' NOTE: enableMarkup / enableBubbles are Boolean. Comparing to Nothing treats Nothing as False.
            If enableMarkup = Nothing Then
                NoMarkup = True
                enableMarkup = False
            End If

            ' NOTE: enableMarkup / enableBubbles are Boolean. Comparing to Nothing treats Nothing as False.
            If enableBubbles = Nothing Then
                NoBubbles = True
                enableBubbles = False
            End If

            If LoadResult <> 0 Then Return ("", False, False, False)

            ' --- Form -----------------------------------------------------------------
            Dim settingsForm As New System.Windows.Forms.Form With {
                    .Text = "Select Prompt",
                    .AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi,
                    .AutoScaleDimensions = New System.Drawing.SizeF(96.0F, 96.0F),
                    .AutoSize = False,
                    .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowOnly,
                    .StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen,
                    .Padding = New System.Windows.Forms.Padding(10),
                    .MinimizeBox = True,
                    .MaximizeBox = True
                }
            settingsForm.MinimumSize = New System.Drawing.Size(900, 650)

            Dim bmp As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
            settingsForm.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
            settingsForm.Font = standardFont

            ' --- Layout grid ----------------------------------------------------------
            Dim layout As New System.Windows.Forms.TableLayoutPanel With {
        .Dock = System.Windows.Forms.DockStyle.Fill,
        .ColumnCount = 2,
        .RowCount = 3,
        .Padding = New System.Windows.Forms.Padding(10)
    }
            layout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0F))
            layout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0F))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 70.0F))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            settingsForm.Controls.Add(layout)

            ' --- Selector --------------------------------------------------------------
            Dim titleListBox As New System.Windows.Forms.ListBox With {
        .Dock = System.Windows.Forms.DockStyle.Fill,
        .Margin = New System.Windows.Forms.Padding(10)
    }
            titleListBox.Items.AddRange(Context.PromptTitles.ToArray())
            layout.Controls.Add(titleListBox, 0, 0)

            ' --- Preview ---------------------------------------------------------------
            Dim promptTextBox As New System.Windows.Forms.TextBox With {
        .Dock = System.Windows.Forms.DockStyle.Fill,
        .Multiline = True,
        .ReadOnly = True,
        .ScrollBars = System.Windows.Forms.ScrollBars.Vertical,
        .Margin = New System.Windows.Forms.Padding(10)
    }
            layout.Controls.Add(promptTextBox, 1, 0)

            If Context.PromptTitles.Count > 0 Then
                titleListBox.SelectedIndex = 0
                promptTextBox.Text = Context.PromptLibrary(0).Replace("\n", vbCrLf)
            End If

            ' Updates the preview on selection changes.
            AddHandler titleListBox.SelectedIndexChanged,
        Sub()
            Dim selectedIndex = titleListBox.SelectedIndex
            If selectedIndex >= 0 Then
                Dim selectedPrompt = Context.PromptLibrary(selectedIndex).Replace("\n", vbCrLf)
                promptTextBox.Text = selectedPrompt
            End If
        End Sub

            ' Confirms the dialog on Enter.
            AddHandler titleListBox.KeyDown,
        Sub(sender As Object, e As System.Windows.Forms.KeyEventArgs)
            If e.KeyCode = System.Windows.Forms.Keys.Enter Then
                settingsForm.DialogResult = System.Windows.Forms.DialogResult.OK
                settingsForm.Close()
            End If
        End Sub

            ' --- Checkboxes (wrapping) ------------------------------------------------
            Dim checkboxPanel As New System.Windows.Forms.FlowLayoutPanel With {
        .FlowDirection = System.Windows.Forms.FlowDirection.TopDown,
        .WrapContents = False,
        .Dock = System.Windows.Forms.DockStyle.Fill,
        .Margin = New System.Windows.Forms.Padding(10),
        .AutoSize = True,
        .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
    }
            layout.Controls.Add(checkboxPanel, 0, 1)

            Dim markupCheckbox As New System.Windows.Forms.CheckBox With {
        .Text = "The output shall be provided as a markup",
        .AutoSize = True,
        .Enabled = enableMarkup,
        .Visible = Not NoMarkup,
        .Margin = New System.Windows.Forms.Padding(3, 3, 3, 6)
    }

            Dim clipboardCheckbox As New System.Windows.Forms.CheckBox With {
        .Text = "The output shall be shown in a window",
        .AutoSize = True,
        .Checked = True,
        .Margin = New System.Windows.Forms.Padding(3, 3, 3, 6)
    }

            Dim bubblesCheckbox As New System.Windows.Forms.CheckBox With {
        .Text = "The output shall be in the form of bubbles",
        .AutoSize = True,
        .Enabled = enableBubbles,
        .Visible = Not NoBubbles,
        .Margin = New System.Windows.Forms.Padding(3, 3, 3, 6)
    }

            checkboxPanel.Controls.Add(markupCheckbox)
            checkboxPanel.Controls.Add(clipboardCheckbox)
            checkboxPanel.Controls.Add(bubblesCheckbox)

            ' Applies MaximumSize to trigger line wrapping based on the left grid cell width.
            Dim ApplyCheckboxWrap As System.Action =
        Sub()
            Dim cellWidthLeft As Integer = CInt((layout.ClientSize.Width - layout.Padding.Horizontal) * layout.ColumnStyles(0).Width / 100.0F) - 20
            If cellWidthLeft < 100 Then cellWidthLeft = 100
            markupCheckbox.MaximumSize = New System.Drawing.Size(cellWidthLeft, 0)
            clipboardCheckbox.MaximumSize = New System.Drawing.Size(cellWidthLeft, 0)
            bubblesCheckbox.MaximumSize = New System.Drawing.Size(cellWidthLeft, 0)
        End Sub
            AddHandler layout.SizeChanged, Sub() ApplyCheckboxWrap()

            ' Mutual exclusivity
            AddHandler markupCheckbox.CheckedChanged, Sub() If markupCheckbox.Checked Then bubblesCheckbox.Checked = False : clipboardCheckbox.Checked = False
            AddHandler bubblesCheckbox.CheckedChanged, Sub() If bubblesCheckbox.Checked Then markupCheckbox.Checked = False : clipboardCheckbox.Checked = False
            AddHandler clipboardCheckbox.CheckedChanged, Sub() If clipboardCheckbox.Checked Then markupCheckbox.Checked = False : bubblesCheckbox.Checked = False

            ' --- Source label (wrapping) ----------------------------------------------
            Dim filePathLabel As New System.Windows.Forms.Label With {
        .Text = $"Source: {filePath}",
        .AutoSize = True,
        .Dock = System.Windows.Forms.DockStyle.Fill,
        .Margin = New System.Windows.Forms.Padding(10),
        .AutoEllipsis = False
    }
            layout.Controls.Add(filePathLabel, 1, 1)

            ' Applies MaximumSize to trigger line wrapping based on the right grid cell width.
            Dim ApplyFilePathWrap As System.Action =
        Sub()
            Dim cellWidthRight As Integer = CInt((layout.ClientSize.Width - layout.Padding.Horizontal) * layout.ColumnStyles(1).Width / 100.0F) - 20
            If cellWidthRight < 100 Then cellWidthRight = 100
            filePathLabel.MaximumSize = New System.Drawing.Size(cellWidthRight, 0)
        End Sub
            AddHandler layout.SizeChanged, Sub() ApplyFilePathWrap()

            ' --- Buttons (LEFT aligned, OK | Cancel | Edit) ---------------------------
            Dim buttonPanel As New System.Windows.Forms.FlowLayoutPanel With {
    .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
    .WrapContents = False,
    .Dock = System.Windows.Forms.DockStyle.Fill,
    .AutoSize = True,
    .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
    .Margin = New System.Windows.Forms.Padding(4),
    .Padding = New System.Windows.Forms.Padding(4) ' Less outer padding
}
            layout.Controls.Add(buttonPanel, 0, 2)
            layout.SetColumnSpan(buttonPanel, 2)

            Dim okButton As New System.Windows.Forms.Button With {
    .Text = "OK",
    .AutoSize = True,
    .DialogResult = System.Windows.Forms.DialogResult.OK,
    .Margin = New System.Windows.Forms.Padding(3), ' Less gap between buttons
    .Padding = New System.Windows.Forms.Padding(8, 4, 8, 4) ' Slimmer buttons
}
            Dim cancelButton As New System.Windows.Forms.Button With {
    .Text = "Cancel",
    .AutoSize = True,
    .DialogResult = System.Windows.Forms.DialogResult.Cancel,
    .Margin = New System.Windows.Forms.Padding(3),
    .Padding = New System.Windows.Forms.Padding(8, 4, 8, 4)
}
            Dim editButton As New System.Windows.Forms.Button With {
    .Text = "Edit",
    .AutoSize = True,
    .Margin = New System.Windows.Forms.Padding(3),
    .Padding = New System.Windows.Forms.Padding(8, 4, 8, 4)
}

            buttonPanel.Controls.Add(okButton)
            buttonPanel.Controls.Add(cancelButton)
            buttonPanel.Controls.Add(editButton)


            ' --- Edit button: show editor + reload list and preview afterwards --------
            AddHandler editButton.Click,
        Sub()
            ShowTextFileEditor(filePath, $"You can now edit your prompts (stored at {filePath}). Make sure that on each line, the description and the prompt is separated by a '|'; you can use ';' for indicating comments.")

            ' Reload prompts after editing
            LoadPrompts(filePath, Context)
            titleListBox.Items.Clear()
            titleListBox.Items.AddRange(Context.PromptTitles.ToArray())

            ' Select first prompt again if available
            If Context.PromptTitles.Count > 0 Then
                titleListBox.SelectedIndex = 0
                promptTextBox.Text = Context.PromptLibrary(0).Replace("\n", vbCrLf)
            Else
                promptTextBox.Clear()
            End If

            titleListBox.Focus()
        End Sub

            ApplyCheckboxWrap()
            ApplyFilePathWrap()

            Dim result As System.Windows.Forms.DialogResult = settingsForm.ShowDialog()

            If result = System.Windows.Forms.DialogResult.OK Then
                Dim selectedIndex = titleListBox.SelectedIndex
                If selectedIndex >= 0 Then
                    Return (
                Context.PromptLibrary(selectedIndex),
                markupCheckbox.Checked,
                bubblesCheckbox.Checked,
                clipboardCheckbox.Checked
            )
                End If
            End If

            Return ("", False, False, False)
        End Function


        ''' <summary>
        ''' Loads prompts from a prompt library file into the provided context.
        ''' </summary>
        ''' <param name="filePath">Prompt library file path (environment variables are expanded).</param>
        ''' <param name="context">Destination context holding prompt titles and prompt texts.</param>
        ''' <returns>
        ''' 0 on success; 1 if missing file; 2 on format-related exception; 3 if no prompts were found; 99 on unexpected errors.
        ''' </returns>
        Public Shared Function LoadPrompts(filePath As String, context As ISharedContext) As Integer

            ' Initialize the return code to 0 (no error)
            Dim returnCode As Integer = 0

            filePath = ExpandEnvironmentVariables(filePath)

            Try
                ' Verify the file exists
                If Not System.IO.File.Exists(filePath) Then
                    ShowCustomMessageBox("The prompt library file was not found.")
                    Return 1
                End If

                context.PromptTitles.Clear()
                context.PromptLibrary.Clear()

                ' Read all lines from the file
                Dim lines = System.IO.File.ReadAllLines(filePath)

                For Each line As String In lines
                    ' Trim leading and trailing spaces
                    Dim trimmedLine = line.Trim()

                    ' Ignore empty lines and lines starting with ';'
                    If Not String.IsNullOrEmpty(trimmedLine) AndAlso Not trimmedLine.StartsWith(";") Then
                        ' Split the line by the delimiter '|'
                        Dim promptData = trimmedLine.Split("|"c)

                        ' Ensure there are at least two parts (title and prompt)
                        If promptData.Length >= 2 Then
                            Dim title = promptData(0).Trim()
                            Dim prompt = String.Join("|", promptData.Skip(1)).Trim()

                            ' Add title and prompt to the respective lists
                            context.PromptTitles.Add(title)
                            context.PromptLibrary.Add(prompt)
                        End If
                    End If
                Next

                ' Check if no prompts were found
                If context.PromptLibrary.Count = 0 Then
                    returnCode = 3
                    ShowCustomMessageBox("No prompts have been found in the configured prompt library file.")
                End If

            Catch ex As System.IO.FileNotFoundException
                returnCode = 1
                ShowCustomMessageBox("The prompt library file was not found: " & ex.Message)

            Catch ex As IndexOutOfRangeException
                returnCode = 2
                ShowCustomMessageBox("The format of the prompt library file is not correct (is a '|' or text thereafter missing?): " & ex.Message)

            Catch ex As Exception
                returnCode = 99
                ShowCustomMessageBox("An unexpected error occurred while loading prompts: " & ex.Message)
            End Try

            Return returnCode
        End Function


        ' Call example from your existing Sub:
        ' ExtractAndStorePromptFromAnalysis(analysis, INI_MyStylePath)

        ''' <summary>
        ''' Extracts [Title = ...] and [Prompt = ...] markers from the provided text, sanitizes them,
        ''' and appends a new "Prefix|Title|Prompt" line to the specified MyStyle prompt file.
        ''' </summary>
        ''' <param name="analysis">Source text containing markers.</param>
        ''' <param name="MyStylePath">Target MyStyle prompt file path.</param>
        ''' <param name="Prefix">Prefix to prepend (defaults to "All" if blank).</param>
        Public Shared Sub ExtractAndStorePromptFromAnalysis(ByVal analysis As System.String, ByVal MyStylePath As System.String, ByVal Prefix As String)
            Try
                ' Basic input validation
                If analysis Is Nothing OrElse analysis.Trim().Length = 0 Then
                    ShowCustomMessageBox("No analysis text was provided.")
                    Return
                End If
                If MyStylePath Is Nothing OrElse MyStylePath.Trim().Length = 0 Then
                    ShowCustomMessageBox("No MyStyle file path ('INI_MyStylePath') is set in the configuration file.")
                    Return
                End If

                ' Try to extract [Title = ...] and [Prompt = ...] near the end of the text (case-insensitive)
                Dim title As System.String = TryGetMarkerValue(analysis, "Title")
                Dim prompt As System.String = TryGetMarkerValue(analysis, "Prompt")

                If title Is Nothing OrElse prompt Is Nothing Then
                    ShowCustomMessageBox("Could not find both [Title = ...] and [Prompt = ...] markers in the analysis text (the text is in the clipboard, so you can manually add it to the file).")
                    Return
                End If

                ' Sanitize to ensure single-line Title|Prompt format (no newlines; safe delimiter)
                title = SanitizeForSingleLine(title)
                prompt = SanitizeForSingleLine(prompt)

                ' Ensure directory exists
                Dim dir As System.String = System.IO.Path.GetDirectoryName(MyStylePath)
                If dir IsNot Nothing AndAlso dir.Trim().Length > 0 AndAlso System.IO.Directory.Exists(dir) = False Then
                    System.IO.Directory.CreateDirectory(dir)
                End If

                ' If file does not exist, create with header and an empty line
                If System.IO.File.Exists(MyStylePath) = False Then
                    Dim header As System.String = "; MyStyle prompt file" & System.Environment.NewLine & System.Environment.NewLine & "; Format: [All|Word|Outlook]|Title of style prompt|style prompt" & System.Environment.NewLine
                    Dim enc As System.Text.Encoding = New System.Text.UTF8Encoding(False) ' UTF-8 without BOM
                    System.IO.File.WriteAllText(MyStylePath, header, enc)
                End If

                If String.IsNullOrWhiteSpace(Prefix) Then Prefix = "All"

                ' Append the new entry: Prefix|Title|Prompt
                Dim line As System.String = System.Environment.NewLine & Prefix & "|" & title & "|" & prompt & System.Environment.NewLine
                System.IO.File.AppendAllText(MyStylePath, line, New System.Text.UTF8Encoding(False))

                ShowCustomMessageBox($"Prompt saved to the MyStyle prompt file ({MyStylePath}).")

            Catch ex As System.Exception
                ShowCustomMessageBox("An error occurred while saving the MyStyle prompt: " & ex.Message)
            End Try
        End Sub

        ' --- Helpers ---

        ''' <summary>
        ''' Returns the value for [Title = ...] or [Prompt = ...] allowing nested brackets in the value.
        ''' Falls back to unbracketed "Title = ..." / "Prompt = ..." (end of line).
        ''' </summary>
        ''' <param name="analysis">Source text to search.</param>
        ''' <param name="markerName">Marker name (e.g., "Title" or "Prompt").</param>
        ''' <returns>The extracted value, or <c>Nothing</c> if not found.</returns>
        Private Shared Function TryGetMarkerValue(ByVal analysis As System.String, ByVal markerName As System.String) As System.String
            ' 1) Prefer bracketed form with balanced square brackets: [Marker = value-with-[nested]-brackets]
            Dim bracketed As System.String = TryGetBracketedMarkerValue(analysis, markerName)
            If bracketed IsNot Nothing Then
                bracketed = bracketed.Trim()
                If bracketed.Length > 0 Then
                    Return bracketed
                End If
            End If

            ' 2) Fallback: unbracketed "Marker = value" up to end of line
            Dim patternLoose As System.String =
        "(?im)^\s*" & System.Text.RegularExpressions.Regex.Escape(markerName) & "\s*=\s*(.+?)\s*$"
            Dim options As System.Text.RegularExpressions.RegexOptions =
        System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.Singleline

            Dim mCol2 As System.Text.RegularExpressions.MatchCollection =
        System.Text.RegularExpressions.Regex.Matches(analysis, patternLoose, options)
            If mCol2 IsNot Nothing AndAlso mCol2.Count > 0 Then
                Dim value As System.String = mCol2(mCol2.Count - 1).Groups(1).Value
                value = value.Trim()
                If value.Length > 0 Then
                    Return value
                End If
            End If

            Return Nothing
        End Function

        ''' <summary>
        ''' Finds the last occurrence of a bracketed marker like:
        ''' [Marker = value]
        ''' and returns the value portion while balancing outer square brackets.
        ''' Matching of the marker name is case-insensitive.
        ''' </summary>
        ''' <param name="analysis">Source text to search.</param>
        ''' <param name="markerName">Marker name (e.g., "Title" or "Prompt").</param>
        ''' <returns>The extracted value, or <c>Nothing</c> if not found or malformed.</returns>
        Private Shared Function TryGetBracketedMarkerValue(ByVal analysis As System.String, ByVal markerName As System.String) As System.String
            If analysis Is Nothing OrElse analysis.Length = 0 Then
                Return Nothing
            End If

            ' Find all occurrences of the opening token "[ marker ="
            Dim openPattern As System.String = "\[\s*" & System.Text.RegularExpressions.Regex.Escape(markerName) & "\s*="
            Dim options As System.Text.RegularExpressions.RegexOptions =
        System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.Singleline

            Dim matches As System.Text.RegularExpressions.MatchCollection =
        System.Text.RegularExpressions.Regex.Matches(analysis, openPattern, options)

            If matches Is Nothing OrElse matches.Count = 0 Then
                Return Nothing
            End If

            ' Use the LAST occurrence to prefer the final summary at the end of the LLM output
            Dim m As System.Text.RegularExpressions.Match = matches(matches.Count - 1)

            ' pos points just after the '='; allow optional spaces before the value
            Dim pos As System.Int32 = m.Index + m.Length
            While pos < analysis.Length AndAlso System.Char.IsWhiteSpace(analysis(pos))
                pos += 1
            End While

            ' Balance square brackets starting from the initial '[' at m.Index
            Dim depth As System.Int32 = 1 ' We are inside the first '['
            Dim i As System.Int32 = pos

            While i < analysis.Length
                Dim ch As System.Char = analysis(i)

                If ch = "["c Then
                    depth += 1
                ElseIf ch = "]"c Then
                    depth -= 1
                    If depth = 0 Then
                        ' The value is everything from pos up to i (excluded)
                        Dim raw As System.String = analysis.Substring(pos, i - pos)
                        Return raw
                    End If
                End If

                i += 1
            End While

            ' If we got here, we never closed the outer '['; treat as not found / malformed
            Return Nothing
        End Function


        ''' <summary>
        ''' Makes a value safe for a single-line "Title|Prompt" config:
        ''' - Replaces CR/LF with spaces
        ''' - Collapses consecutive whitespace
        ''' - Replaces "|" with "¦" (broken bar) to avoid delimiter collision
        ''' - Trims surrounding whitespace
        ''' </summary>
        ''' <param name="input">Input to sanitize.</param>
        ''' <returns>Sanitized single-line string.</returns>
        Private Shared Function SanitizeForSingleLine(ByVal input As System.String) As System.String
            If input Is Nothing Then
                Return System.String.Empty
            End If

            Dim s As System.String = input.Replace(vbCr, " ").Replace(vbLf, " ")
            s = System.Text.RegularExpressions.Regex.Replace(s, "\s+", " ")
            s = s.Replace("|", "¦")
            Return s.Trim()
        End Function



    End Class

End Namespace