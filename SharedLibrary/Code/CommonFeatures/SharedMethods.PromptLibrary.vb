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
Imports System.Runtime.InteropServices
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary

    Partial Public Class SharedMethods

        Private Const PromptLibraryAllCategoriesLabel As String = "(All categories)"
        Private Const PromptLibraryUncategorizedLabel As String = "(Uncategorized)"

        Private NotInheritable Class PromptLibraryEntry

            Public ReadOnly Title As String
            Public ReadOnly Prompt As String
            Public ReadOnly Category As String

            Public Sub New(title As String, prompt As String, category As String)
                Me.Title = title
                Me.Prompt = prompt
                Me.Category = category
            End Sub

        End Class

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

            Dim allEntries As New List(Of PromptLibraryEntry)()
            Dim combinedTitles As New List(Of String)()
            Dim combinedPrompts As New List(Of String)()

            Dim SyncContextFromEntries As System.Action =
                Sub()
                    combinedTitles.Clear()
                    combinedPrompts.Clear()

                    For Each entry As PromptLibraryEntry In allEntries
                        combinedTitles.Add(entry.Title)
                        combinedPrompts.Add(entry.Prompt)
                    Next

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
                End Sub

            Dim ReloadPromptEntries As System.Action =
                Sub()
                    Dim localEntries As New List(Of PromptLibraryEntry)()
                    Dim centralEntries As New List(Of PromptLibraryEntry)()

                    allEntries.Clear()

                    LoadPromptEntriesIntoList(localPath, localEntries, " (local)")
                    LoadPromptEntriesIntoList(centralPath, centralEntries, Nothing)

                    allEntries.AddRange(localEntries)
                    allEntries.AddRange(centralEntries)

                    SyncContextFromEntries()
                End Sub

            ReloadPromptEntries()

            Dim NoBubbles As Boolean = False
            Dim NoMarkup As Boolean = False

            If enableMarkup = Nothing Then
                NoMarkup = True
                enableMarkup = False
            End If

            If enableBubbles = Nothing Then
                NoBubbles = True
                enableBubbles = False
            End If

            If allEntries.Count = 0 Then
                ShowCustomMessageBox("No prompts have been found in the configured prompt library files.")
                Return ("", False, False, False)
            End If

            Dim settingsForm As New System.Windows.Forms.Form With {
                    .Text = "Select Prompt",
                    .AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi,
                    .AutoScaleDimensions = New System.Drawing.SizeF(96.0F, 96.0F),
                    .AutoSize = False,
                    .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowOnly,
                    .StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen,
                    .Padding = New System.Windows.Forms.Padding(10),
                    .MinimizeBox = True,
                    .MaximizeBox = True,
                    .TopMost = True
                }
            settingsForm.MinimumSize = New System.Drawing.Size(900, 650)

            Dim bmp As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
            settingsForm.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
            settingsForm.Font = standardFont

            Dim layout As New System.Windows.Forms.TableLayoutPanel With {
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 4,
                .Padding = New System.Windows.Forms.Padding(10)
            }
            layout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0F))
            layout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0F))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 70.0F))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            settingsForm.Controls.Add(layout)

            Dim categoryPanel As New System.Windows.Forms.FlowLayoutPanel With {
                .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
                .WrapContents = False,
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .Margin = New System.Windows.Forms.Padding(10, 10, 10, 0),
                .AutoSize = True,
                .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
            }

            Dim categoryLabel As New System.Windows.Forms.Label With {
                .Text = "Category:",
                .AutoSize = True,
                .Margin = New System.Windows.Forms.Padding(3, 8, 8, 3)
            }

            Dim categoryComboBox As New System.Windows.Forms.ComboBox With {
                .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList,
                .Width = 320,
                .Margin = New System.Windows.Forms.Padding(3)
            }

            categoryPanel.Controls.Add(categoryLabel)
            categoryPanel.Controls.Add(categoryComboBox)
            layout.Controls.Add(categoryPanel, 0, 0)
            layout.SetColumnSpan(categoryPanel, 2)

            Dim titleListBox As New System.Windows.Forms.ListBox With {
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .Margin = New System.Windows.Forms.Padding(10)
            }
            layout.Controls.Add(titleListBox, 0, 1)

            Dim promptTextBox As New System.Windows.Forms.TextBox With {
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .Multiline = True,
                .ReadOnly = True,
                .ScrollBars = System.Windows.Forms.ScrollBars.Vertical,
                .Margin = New System.Windows.Forms.Padding(10)
            }
            layout.Controls.Add(promptTextBox, 1, 1)

            Dim checkboxPanel As New System.Windows.Forms.FlowLayoutPanel With {
                .FlowDirection = System.Windows.Forms.FlowDirection.TopDown,
                .WrapContents = False,
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .Margin = New System.Windows.Forms.Padding(10),
                .AutoSize = True,
                .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
            }
            layout.Controls.Add(checkboxPanel, 0, 2)

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

            Dim ApplyCheckboxWrap As System.Action =
                Sub()
                    Dim cellWidthLeft As Integer = CInt((layout.ClientSize.Width - layout.Padding.Horizontal) * layout.ColumnStyles(0).Width / 100.0F) - 20
                    If cellWidthLeft < 100 Then cellWidthLeft = 100
                    markupCheckbox.MaximumSize = New System.Drawing.Size(cellWidthLeft, 0)
                    clipboardCheckbox.MaximumSize = New System.Drawing.Size(cellWidthLeft, 0)
                    bubblesCheckbox.MaximumSize = New System.Drawing.Size(cellWidthLeft, 0)
                End Sub
            AddHandler layout.SizeChanged, Sub() ApplyCheckboxWrap()

            AddHandler markupCheckbox.CheckedChanged, Sub() If markupCheckbox.Checked Then bubblesCheckbox.Checked = False : clipboardCheckbox.Checked = False
            AddHandler bubblesCheckbox.CheckedChanged, Sub() If bubblesCheckbox.Checked Then markupCheckbox.Checked = False : clipboardCheckbox.Checked = False
            AddHandler clipboardCheckbox.CheckedChanged, Sub() If clipboardCheckbox.Checked Then markupCheckbox.Checked = False : bubblesCheckbox.Checked = False

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
            layout.Controls.Add(filePathLabel, 1, 2)

            Dim ApplyFilePathWrap As System.Action =
                Sub()
                    Dim cellWidthRight As Integer = CInt((layout.ClientSize.Width - layout.Padding.Horizontal) * layout.ColumnStyles(1).Width / 100.0F) - 20
                    If cellWidthRight < 100 Then cellWidthRight = 100
                    filePathLabel.MaximumSize = New System.Drawing.Size(cellWidthRight, 0)
                End Sub
            AddHandler layout.SizeChanged, Sub() ApplyFilePathWrap()

            Dim buttonPanel As New System.Windows.Forms.FlowLayoutPanel With {
                .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
                .WrapContents = False,
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .AutoSize = True,
                .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                .Margin = New System.Windows.Forms.Padding(4),
                .Padding = New System.Windows.Forms.Padding(4)
            }
            layout.Controls.Add(buttonPanel, 0, 3)
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

            Dim visibleEntries As New List(Of PromptLibraryEntry)()
            Dim suppressCategoryRefresh As Boolean = False

            Dim RefreshVisiblePromptList As System.Action(Of String) =
                Sub(preferredTitle As String)
                    Dim selectedCategory As String = PromptLibraryAllCategoriesLabel
                    If categoryComboBox.SelectedItem IsNot Nothing Then
                        selectedCategory = CStr(categoryComboBox.SelectedItem)
                    End If

                    visibleEntries.Clear()

                    For Each entry As PromptLibraryEntry In allEntries
                        If String.Equals(selectedCategory, PromptLibraryAllCategoriesLabel, StringComparison.OrdinalIgnoreCase) OrElse
                           String.Equals(entry.Category, selectedCategory, StringComparison.OrdinalIgnoreCase) Then
                            visibleEntries.Add(entry)
                        End If
                    Next

                    titleListBox.BeginUpdate()
                    titleListBox.Items.Clear()
                    For Each entry As PromptLibraryEntry In visibleEntries
                        titleListBox.Items.Add(entry.Title)
                    Next
                    titleListBox.EndUpdate()

                    Dim selectedIndex As Integer = -1

                    If Not String.IsNullOrWhiteSpace(preferredTitle) Then
                        For i As Integer = 0 To visibleEntries.Count - 1
                            If String.Equals(visibleEntries(i).Title, preferredTitle, StringComparison.OrdinalIgnoreCase) Then
                                selectedIndex = i
                                Exit For
                            End If
                        Next
                    End If

                    If selectedIndex = -1 AndAlso visibleEntries.Count > 0 Then
                        selectedIndex = 0
                    End If

                    If selectedIndex >= 0 Then
                        titleListBox.SelectedIndex = selectedIndex
                        promptTextBox.Text = visibleEntries(selectedIndex).Prompt.Replace("\n", vbCrLf)
                    Else
                        promptTextBox.Clear()
                    End If
                End Sub

            Dim RefreshCategoryFilter As System.Action(Of String) =
                Sub(preferredCategory As String)
                    Dim categories As New List(Of String)()
                    Dim seenCategories As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                    Dim hasDefinedCategories As Boolean = False

                    categories.Add(PromptLibraryAllCategoriesLabel)

                    For Each entry As PromptLibraryEntry In allEntries
                        If String.Equals(entry.Category, PromptLibraryUncategorizedLabel, StringComparison.OrdinalIgnoreCase) Then
                            Continue For
                        End If

                        If seenCategories.Add(entry.Category) Then
                            categories.Add(entry.Category)
                            hasDefinedCategories = True
                        End If
                    Next

                    If hasDefinedCategories Then
                        For Each entry As PromptLibraryEntry In allEntries
                            If String.Equals(entry.Category, PromptLibraryUncategorizedLabel, StringComparison.OrdinalIgnoreCase) Then
                                If seenCategories.Add(entry.Category) Then
                                    categories.Add(entry.Category)
                                End If
                            End If
                        Next
                    End If

                    suppressCategoryRefresh = True

                    categoryComboBox.BeginUpdate()
                    categoryComboBox.Items.Clear()
                    categoryComboBox.Items.AddRange(categories.ToArray())
                    categoryComboBox.EndUpdate()

                    categoryComboBox.Enabled = hasDefinedCategories
                    categoryLabel.Enabled = hasDefinedCategories

                    Dim selectedIndex As Integer = 0

                    If hasDefinedCategories AndAlso Not String.IsNullOrWhiteSpace(preferredCategory) Then
                        For i As Integer = 0 To categories.Count - 1
                            If String.Equals(categories(i), preferredCategory, StringComparison.OrdinalIgnoreCase) Then
                                selectedIndex = i
                                Exit For
                            End If
                        Next
                    End If

                    If categoryComboBox.Items.Count > 0 Then
                        categoryComboBox.SelectedIndex = selectedIndex
                    End If

                    suppressCategoryRefresh = False
                End Sub

            RefreshCategoryFilter(PromptLibraryAllCategoriesLabel)
            RefreshVisiblePromptList(Nothing)

            AddHandler categoryComboBox.SelectedIndexChanged,
                Sub()
                    If suppressCategoryRefresh Then Return
                    RefreshVisiblePromptList(Nothing)
                End Sub

            AddHandler titleListBox.SelectedIndexChanged,
                Sub()
                    Dim selectedIndex = titleListBox.SelectedIndex
                    If selectedIndex >= 0 AndAlso selectedIndex < visibleEntries.Count Then
                        promptTextBox.Text = visibleEntries(selectedIndex).Prompt.Replace("\n", vbCrLf)
                    Else
                        promptTextBox.Clear()
                    End If
                End Sub

            AddHandler titleListBox.KeyDown,
                Sub(sender As Object, e As System.Windows.Forms.KeyEventArgs)
                    If e.KeyCode = System.Windows.Forms.Keys.Enter Then
                        settingsForm.DialogResult = System.Windows.Forms.DialogResult.OK
                        settingsForm.Close()
                    End If
                End Sub

            AddHandler editButton.Click,
                Sub()
                    Dim previousCategory As String = PromptLibraryAllCategoriesLabel
                    If categoryComboBox.SelectedItem IsNot Nothing Then
                        previousCategory = CStr(categoryComboBox.SelectedItem)
                    End If

                    Dim previousTitle As String = Nothing
                    If titleListBox.SelectedIndex >= 0 AndAlso titleListBox.SelectedIndex < visibleEntries.Count Then
                        previousTitle = visibleEntries(titleListBox.SelectedIndex).Title
                    End If

                    Dim target As String = If(hasLocal, localPath, centralPath)
                    Dim targetKind As String = If(hasLocal, "local", "central")

                    EnsurePromptLibraryDirectoryExists(target)
                    ShowTextFileEditor(
                        target,
                        $"You can now edit your {targetKind} prompts (stored at {target}). Make sure that on each prompt line, the description and the prompt are separated by a '|'; you can use ';' for comments; optional category blocks use <Category Name> and </Category Name>."
                    )

                    ReloadPromptEntries()
                    RefreshCategoryFilter(previousCategory)
                    RefreshVisiblePromptList(previousTitle)

                    titleListBox.Focus()
                End Sub

            ApplyCheckboxWrap()
            ApplyFilePathWrap()

            AddHandler settingsForm.Shown,
                Sub()
                    settingsForm.TopMost = True
                    Global.SharedLibrary.SharedLibrary.SharedMethods.ForceDialogToForeground(settingsForm)
                    Global.SharedLibrary.SharedLibrary.SharedMethods.AttachForeignForegroundWatchdog(settingsForm)
                End Sub
            Dim __safeDialogOwner490 As System.Windows.Forms.IWin32Window = Global.SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
            Dim result As System.Windows.Forms.DialogResult = If(__safeDialogOwner490 IsNot Nothing, settingsForm.ShowDialog(__safeDialogOwner490), settingsForm.ShowDialog())

            If result = System.Windows.Forms.DialogResult.OK Then
                Dim selectedIndex = titleListBox.SelectedIndex
                If selectedIndex >= 0 AndAlso selectedIndex < visibleEntries.Count Then
                    Return (
                        visibleEntries(selectedIndex).Prompt,
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
        Private Shared Sub LoadPromptEntriesIntoList(filePath As String,
                                                     entries As List(Of PromptLibraryEntry),
                                                     Optional titleSuffix As String = Nothing)
            Try
                If String.IsNullOrWhiteSpace(filePath) Then Return
                filePath = ExpandEnvironmentVariables(filePath)
                If Not System.IO.File.Exists(filePath) Then Return

                Dim currentCategory As String = Nothing
                Dim lines = System.IO.File.ReadAllLines(filePath)

                For Each line As String In lines
                    Dim trimmedLine As String = line.Trim()
                    If trimmedLine.Length = 0 OrElse trimmedLine.StartsWith(";") Then Continue For

                    If trimmedLine.StartsWith("</", StringComparison.Ordinal) AndAlso trimmedLine.EndsWith(">", StringComparison.Ordinal) Then
                        currentCategory = Nothing
                        Continue For
                    End If

                    If trimmedLine.StartsWith("<", StringComparison.Ordinal) AndAlso
                       trimmedLine.EndsWith(">", StringComparison.Ordinal) AndAlso
                       Not trimmedLine.Contains("|") Then

                        Dim openingTagName As String = trimmedLine.Substring(1, trimmedLine.Length - 2).Trim()
                        If openingTagName.Length > 0 AndAlso Not openingTagName.StartsWith("/", StringComparison.Ordinal) Then
                            currentCategory = openingTagName
                            Continue For
                        End If
                    End If

                    Dim parts = trimmedLine.Split("|"c)
                    If parts.Length >= 2 Then
                        Dim title As String = parts(0).Trim()
                        Dim prompt As String

                        If parts.Length = 2 Then
                            prompt = parts(1).Trim()
                        Else
                            prompt = String.Join("|", parts, 1, parts.Length - 1).Trim()
                        End If

                        If Not String.IsNullOrEmpty(titleSuffix) Then title &= titleSuffix

                        Dim category As String = If(
                            String.IsNullOrWhiteSpace(currentCategory),
                            PromptLibraryUncategorizedLabel,
                            currentCategory.Trim()
                        )

                        entries.Add(New PromptLibraryEntry(title, prompt, category))
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

            AddHandler settingsForm.Shown,
                Sub()
                    settingsForm.TopMost = True
                    Global.SharedLibrary.SharedLibrary.SharedMethods.ForceDialogToForeground(settingsForm)
                    Global.SharedLibrary.SharedLibrary.SharedMethods.AttachForeignForegroundWatchdog(settingsForm)
                End Sub
            Dim __safeDialogOwner821 As System.Windows.Forms.IWin32Window = Global.SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
            Dim result As System.Windows.Forms.DialogResult = If(__safeDialogOwner821 IsNot Nothing, settingsForm.ShowDialog(__safeDialogOwner821), settingsForm.ShowDialog())

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


        Public Enum PromptLibrarySlashAction
            NotTriggered = 0
            Canceled = 1
            Inserted = 2
        End Enum

        ''' <summary>
        ''' Returns <c>True</c> when a typed slash should open the prompt library instead of being inserted literally.
        ''' A slash is treated as a command trigger only at the start of the text or after whitespace.
        ''' </summary>
        Public Shared Function IsPromptLibrarySlashTrigger(targetTextBox As System.Windows.Forms.TextBoxBase) As Boolean
            If targetTextBox Is Nothing Then Return False

            Dim insertionIndex As Integer = targetTextBox.SelectionStart
            If insertionIndex <= 0 Then Return True
            If insertionIndex > targetTextBox.TextLength Then insertionIndex = targetTextBox.TextLength

            Dim previousChar As Char = targetTextBox.Text.Chars(insertionIndex - 1)
            Return Char.IsWhiteSpace(previousChar)
        End Function


        Private Shared Function GetEditablePromptLibraryPath(localPath As String, centralPath As String) As String
            If Not String.IsNullOrWhiteSpace(localPath) Then
                Return localPath
            End If

            Return centralPath
        End Function

        Private Shared Sub EnsurePromptLibraryDirectoryExists(filePath As String)
            If String.IsNullOrWhiteSpace(filePath) Then
                Throw New InvalidOperationException("No prompt library file path is configured.")
            End If

            Dim dir As String = System.IO.Path.GetDirectoryName(filePath)
            If String.IsNullOrWhiteSpace(dir) Then
                Throw New InvalidOperationException("The prompt library path does not contain a valid directory.")
            End If

            If Not System.IO.Directory.Exists(dir) Then
                System.IO.Directory.CreateDirectory(dir)
            End If
        End Sub

        Private Shared Sub EnsurePromptLibraryFileExists(filePath As String)
            EnsurePromptLibraryDirectoryExists(filePath)

            If System.IO.File.Exists(filePath) Then
                Return
            End If

            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine("; Local prompt library")
            sb.AppendLine("; Format: Title|Prompt")
            sb.AppendLine("; Optional categories: <Category Name> ... </Category Name>")
            sb.AppendLine()

            System.IO.File.WriteAllText(filePath, sb.ToString(), New System.Text.UTF8Encoding(False))
        End Sub

        Private Shared Function NormalizePromptLibraryPromptText(value As String) As String
            If value Is Nothing Then
                Return String.Empty
            End If

            Return value.Replace(vbCrLf, "\n").
                         Replace(vbCr, "\n").
                         Replace(vbLf, "\n").
                         Replace("|", "¦").
                         Trim()
        End Function

        Private Shared Sub AppendPromptLibraryEntry(filePath As String,
                                                    title As String,
                                                    prompt As String,
                                                    Optional category As String = Nothing)
            title = SanitizeForSingleLine(title)
            category = SanitizeForSingleLine(category)
            prompt = NormalizePromptLibraryPromptText(prompt)

            If String.IsNullOrWhiteSpace(title) Then
                Throw New InvalidOperationException("Please provide a title.")
            End If

            If String.IsNullOrWhiteSpace(prompt) Then
                Throw New InvalidOperationException("Please provide a prompt.")
            End If

            EnsurePromptLibraryFileExists(filePath)

            Dim sb As New System.Text.StringBuilder()

            Try
                If System.IO.File.Exists(filePath) AndAlso New System.IO.FileInfo(filePath).Length > 0 Then
                    sb.AppendLine()
                End If
            Catch
            End Try

            If Not String.IsNullOrWhiteSpace(category) Then
                sb.AppendLine("<" & category & ">")
            End If

            sb.AppendLine(title & "|" & prompt)

            If Not String.IsNullOrWhiteSpace(category) Then
                sb.AppendLine("</" & category & ">")
            End If

            System.IO.File.AppendAllText(filePath, sb.ToString(), New System.Text.UTF8Encoding(False))
        End Sub

        Private Shared Function ShowAddPromptLibraryEntryDialog(targetPath As String,
                                                        defaultCategory As String,
                                                        lastPromptForCtrlP As String,
                                                        ByRef newTitle As String,
                                                        ByRef newPrompt As String,
                                                        ByRef newCategory As String,
                                                        Optional ownerHandle As IntPtr = Nothing) As Boolean


            newTitle = Nothing
            newPrompt = Nothing
            newCategory = Nothing

            Dim pendingTitle As String = Nothing
            Dim pendingPrompt As String = Nothing
            Dim pendingCategory As String = Nothing

            Dim form As New System.Windows.Forms.Form With {
                .Text = "Add Prompt",
                .StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen,
                .AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi,
                .AutoScaleDimensions = New System.Drawing.SizeF(96.0F, 96.0F),
                .Padding = New System.Windows.Forms.Padding(10),
                .MinimizeBox = False,
                .MaximizeBox = False,
                .TopMost = True
            }
            form.MinimumSize = New System.Drawing.Size(820, 560)

            Dim bmp As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
            form.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())
            form.Font = New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)

            Dim layout As New System.Windows.Forms.TableLayoutPanel With {
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 5,
                .Padding = New System.Windows.Forms.Padding(10)
            }
            layout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize))
            layout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            form.Controls.Add(layout)

            Dim infoLabel As New System.Windows.Forms.Label With {
                .Text = $"The new prompt will be stored in: {targetPath}",
                .AutoSize = True,
                .Margin = New System.Windows.Forms.Padding(3, 3, 3, 10)
            }
            layout.Controls.Add(infoLabel, 0, 0)
            layout.SetColumnSpan(infoLabel, 2)

            Dim titleLabel As New System.Windows.Forms.Label With {
                .Text = "Title:",
                .AutoSize = True,
                .Margin = New System.Windows.Forms.Padding(3, 8, 8, 3)
            }
            Dim titleTextBox As New System.Windows.Forms.TextBox With {
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .Margin = New System.Windows.Forms.Padding(3)
            }
            layout.Controls.Add(titleLabel, 0, 1)
            layout.Controls.Add(titleTextBox, 1, 1)

            Dim categoryLabel As New System.Windows.Forms.Label With {
                .Text = "Category:",
                .AutoSize = True,
                .Margin = New System.Windows.Forms.Padding(3, 8, 8, 3)
            }
            Dim categoryTextBox As New System.Windows.Forms.TextBox With {
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .Margin = New System.Windows.Forms.Padding(3),
                .Text = If(defaultCategory, "")
            }
            layout.Controls.Add(categoryLabel, 0, 2)
            layout.Controls.Add(categoryTextBox, 1, 2)

            Dim promptLabel As New System.Windows.Forms.Label With {
                .Text = "Prompt:",
                .AutoSize = True,
                .Margin = New System.Windows.Forms.Padding(3, 8, 8, 3)
            }
            Dim promptTextBox As New System.Windows.Forms.TextBox With {
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .Multiline = True,
                .AcceptsReturn = True,
                .ScrollBars = System.Windows.Forms.ScrollBars.Vertical,
                .Margin = New System.Windows.Forms.Padding(3)
            }
            layout.Controls.Add(promptLabel, 0, 3)
            layout.Controls.Add(promptTextBox, 1, 3)

            Dim hintLabel As New System.Windows.Forms.Label With {
                .Text = "Ctrl+P inserts the caller's last prompt. Ctrl+Enter adds the new prompt.",
                .AutoSize = True,
                .Margin = New System.Windows.Forms.Padding(3, 8, 3, 3)
            }
            layout.Controls.Add(hintLabel, 0, 4)
            layout.SetColumnSpan(hintLabel, 2)

            Dim buttonPanel As New System.Windows.Forms.FlowLayoutPanel With {
                .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
                .WrapContents = False,
                .Dock = System.Windows.Forms.DockStyle.Bottom,
                .AutoSize = True,
                .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                .Margin = New System.Windows.Forms.Padding(4),
                .Padding = New System.Windows.Forms.Padding(4)
            }
            form.Controls.Add(buttonPanel)

            Dim addButton As New System.Windows.Forms.Button With {
                .Text = "Add",
                .AutoSize = True,
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

            buttonPanel.Controls.Add(addButton)
            buttonPanel.Controls.Add(cancelButton)

            form.AcceptButton = addButton
            form.CancelButton = cancelButton

            AddHandler addButton.Click,
                Sub()
                    Dim titleValue As String = titleTextBox.Text.Trim()
                    Dim promptValue As String = promptTextBox.Text.Trim()

                    If titleValue.Length = 0 Then
                        ShowCustomMessageBox("Please provide a title.")
                        titleTextBox.Focus()
                        Return
                    End If

                    If promptValue.Length = 0 Then
                        ShowCustomMessageBox("Please provide a prompt.")
                        promptTextBox.Focus()
                        Return
                    End If

                    pendingTitle = titleValue
                    pendingPrompt = promptValue
                    pendingCategory = categoryTextBox.Text.Trim()

                    form.DialogResult = System.Windows.Forms.DialogResult.OK
                    form.Close()
                End Sub

            AddHandler promptTextBox.KeyDown,
                Sub(sender As Object, e As System.Windows.Forms.KeyEventArgs)
                    If e.Control AndAlso e.KeyCode = System.Windows.Forms.Keys.P Then
                        If Not String.IsNullOrWhiteSpace(lastPromptForCtrlP) Then
                            Dim startIndex As Integer = promptTextBox.SelectionStart
                            Dim selectionLength As Integer = promptTextBox.SelectionLength

                            promptTextBox.Text =
                                promptTextBox.Text.Remove(startIndex, selectionLength).Insert(startIndex, lastPromptForCtrlP)

                            promptTextBox.SelectionStart = startIndex + lastPromptForCtrlP.Length
                            promptTextBox.SelectionLength = 0
                        End If

                        e.SuppressKeyPress = True
                        e.Handled = True
                        Return
                    End If

                    If e.Control AndAlso e.KeyCode = System.Windows.Forms.Keys.Enter Then
                        addButton.PerformClick()
                        e.SuppressKeyPress = True
                        e.Handled = True
                    End If
                End Sub

            AddHandler form.Shown,
    Sub()
        titleTextBox.Focus()
        Try
            Global.SharedLibrary.SharedLibrary.SharedMethods.ForceDialogToForeground(form)
            Global.SharedLibrary.SharedLibrary.SharedMethods.AttachForeignForegroundWatchdog(form)
        Catch ex As System.Exception
        End Try
    End Sub

            Dim result As System.Windows.Forms.DialogResult
            Dim owner As System.Windows.Forms.IWin32Window = Nothing

            If ownerHandle <> System.IntPtr.Zero Then
                owner = New WindowWrapper(ownerHandle)
            Else
                owner = System.Windows.Forms.Form.ActiveForm
            End If

            ' Inspect and same-thread filter the resolved owner before using it as a modal
            ' owner. A foreign owner (e.g. another Office process's main window, whether passed
            ' in via ownerHandle or picked up as ActiveForm) would be disabled by ShowDialog and
            ' never re-enabled, deadlocking that host. InspectDialogOwner logs cross-process
            ' attempts; IfOwnerOnCurrentThread rejects a cross-thread owner so we fall back to
            ' an ownerless dialog instead of deadlocking.
            If owner IsNot Nothing Then
                OfficeWindowWatchdog.InspectDialogOwner(owner, "ShowAddPromptLibraryEntryDialog", "ShowAddPromptLibraryEntryDialog")
                owner = SharedMethods.IfOwnerOnCurrentThread(owner)
            End If

            If owner IsNot Nothing Then
                result = form.ShowDialog(owner)
            Else
                result = form.ShowDialog()
            End If

            If result = System.Windows.Forms.DialogResult.OK Then
                newTitle = pendingTitle
                newPrompt = pendingPrompt
                newCategory = pendingCategory
                Return True
            End If

            Return False
        End Function

        Private Shared Sub WarnIfPromptContainsUnsupportedChatPlaceholders(prompt As String)
            If String.IsNullOrEmpty(prompt) Then
                Return
            End If

            If prompt.IndexOf("{doc}", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               prompt.IndexOf("{dir}", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               prompt.IndexOf("{[path]}", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               prompt.IndexOf("(file)", StringComparison.OrdinalIgnoreCase) >= 0 Then

                ShowCustomMessageBox(
                    "This prompt contains placeholder text such as {doc}, {dir}, {[path]} or (file). " &
                    "These placeholders are not supported when the prompt library is used within a chat, so they will be inserted and processed literally."
                )
            End If
        End Sub

        ''' <summary>
        ''' Opens a prompt picker that can be filtered by prompt name and category and returns the selected prompt text.
        ''' Returns an empty string if the user cancels or no prompt is available.
        ''' </summary>
        Public Shared Function ShowPromptInsertionSelector(filePath As String,
                                                   filepathlocal As String,
                                                   Context As ISharedContext,
                                                   Optional initialFilter As String = Nothing,
                                                   Optional lastPromptForCtrlP As String = Nothing,
                                                   Optional ownerHandle As IntPtr = Nothing,
                                                   Optional NoWarning As Boolean = False) As String

            Dim centralPath As String = ExpandEnvironmentVariables(filePath)
            Dim localPath As String = ExpandEnvironmentVariables(filepathlocal)
            Dim hasLocal As Boolean = Not String.IsNullOrWhiteSpace(localPath)

            Dim allEntries As New List(Of PromptLibraryEntry)()
            Dim combinedTitles As New List(Of String)()
            Dim combinedPrompts As New List(Of String)()

            Dim SyncContextFromEntries As System.Action =
                Sub()
                    combinedTitles.Clear()
                    combinedPrompts.Clear()

                    For Each entry As PromptLibraryEntry In allEntries
                        combinedTitles.Add(entry.Title)
                        combinedPrompts.Add(entry.Prompt)
                    Next

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
                End Sub

            Dim ReloadPromptEntries As System.Action =
                Sub()
                    Dim localEntries As New List(Of PromptLibraryEntry)()
                    Dim centralEntries As New List(Of PromptLibraryEntry)()

                    allEntries.Clear()

                    LoadPromptEntriesIntoList(localPath, localEntries, " (local)")
                    LoadPromptEntriesIntoList(centralPath, centralEntries, Nothing)

                    allEntries.AddRange(localEntries)
                    allEntries.AddRange(centralEntries)

                    SyncContextFromEntries()
                End Sub

            ReloadPromptEntries()

            Dim pickerForm As New System.Windows.Forms.Form With {
                .Text = "Insert Prompt",
                .AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi,
                .AutoScaleDimensions = New System.Drawing.SizeF(96.0F, 96.0F),
                .StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen,
                .Padding = New System.Windows.Forms.Padding(10),
                .MinimizeBox = True,
                .MaximizeBox = True,
                .TopMost = True
            }
            pickerForm.MinimumSize = New System.Drawing.Size(950, 680)

            Dim bmp As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
            pickerForm.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())
            pickerForm.Font = New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)

            Dim layout As New System.Windows.Forms.TableLayoutPanel With {
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 5,
                .Padding = New System.Windows.Forms.Padding(10)
            }
            layout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0F))
            layout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0F))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            pickerForm.Controls.Add(layout)

            Dim searchPanel As New System.Windows.Forms.FlowLayoutPanel With {
                .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
                .WrapContents = False,
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .Margin = New System.Windows.Forms.Padding(10, 10, 10, 0),
                .AutoSize = True,
                .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
            }

            Dim searchLabel As New System.Windows.Forms.Label With {
                .Text = "Search:",
                .AutoSize = True,
                .Margin = New System.Windows.Forms.Padding(3, 8, 8, 3)
            }

            Dim searchTextBox As New System.Windows.Forms.TextBox With {
                .Width = 420,
                .Margin = New System.Windows.Forms.Padding(3),
                .Text = If(initialFilter, "")
            }

            searchPanel.Controls.Add(searchLabel)
            searchPanel.Controls.Add(searchTextBox)
            layout.Controls.Add(searchPanel, 0, 0)
            layout.SetColumnSpan(searchPanel, 2)

            Dim categoryPanel As New System.Windows.Forms.FlowLayoutPanel With {
                .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
                .WrapContents = False,
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .Margin = New System.Windows.Forms.Padding(10, 4, 10, 0),
                .AutoSize = True,
                .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
            }

            Dim categoryLabel As New System.Windows.Forms.Label With {
                .Text = "Category:",
                .AutoSize = True,
                .Margin = New System.Windows.Forms.Padding(3, 8, 8, 3)
            }

            Dim categoryComboBox As New System.Windows.Forms.ComboBox With {
                .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList,
                .Width = 320,
                .Margin = New System.Windows.Forms.Padding(3)
            }

            categoryPanel.Controls.Add(categoryLabel)
            categoryPanel.Controls.Add(categoryComboBox)
            layout.Controls.Add(categoryPanel, 0, 1)
            layout.SetColumnSpan(categoryPanel, 2)

            Dim titleListBox As New System.Windows.Forms.ListBox With {
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .Margin = New System.Windows.Forms.Padding(10)
            }
            layout.Controls.Add(titleListBox, 0, 2)

            Dim promptTextBox As New System.Windows.Forms.TextBox With {
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .Multiline = True,
                .ReadOnly = True,
                .ScrollBars = System.Windows.Forms.ScrollBars.Vertical,
                .Margin = New System.Windows.Forms.Padding(10)
            }
            layout.Controls.Add(promptTextBox, 1, 2)

            Dim sourceText As String
            If hasLocal Then
                sourceText = $"Source: {localPath} (local, editable) | {centralPath} (central)"
            Else
                sourceText = $"Source: {centralPath} (central, editable)"
            End If

            Dim sourceLabel As New System.Windows.Forms.Label With {
                .Text = sourceText,
                .AutoSize = True,
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .Margin = New System.Windows.Forms.Padding(10),
                .AutoEllipsis = False
            }
            layout.Controls.Add(sourceLabel, 0, 3)
            layout.SetColumnSpan(sourceLabel, 2)

            Dim buttonPanel As New System.Windows.Forms.FlowLayoutPanel With {
                .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
                .WrapContents = False,
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .AutoSize = True,
                .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                .Margin = New System.Windows.Forms.Padding(4),
                .Padding = New System.Windows.Forms.Padding(4)
            }
            layout.Controls.Add(buttonPanel, 0, 4)
            layout.SetColumnSpan(buttonPanel, 2)

            Dim insertButton As New System.Windows.Forms.Button With {
                .Text = "Insert",
                .AutoSize = True,
                .DialogResult = System.Windows.Forms.DialogResult.OK,
                .Margin = New System.Windows.Forms.Padding(3),
                .Padding = New System.Windows.Forms.Padding(8, 4, 8, 4),
                .Enabled = False
            }

            Dim cancelButton As New System.Windows.Forms.Button With {
                .Text = "Cancel",
                .AutoSize = True,
                .DialogResult = System.Windows.Forms.DialogResult.Cancel,
                .Margin = New System.Windows.Forms.Padding(3),
                .Padding = New System.Windows.Forms.Padding(8, 4, 8, 4)
            }

            Dim addButton As New System.Windows.Forms.Button With {
                .Text = "Add",
                .AutoSize = True,
                .Margin = New System.Windows.Forms.Padding(3),
                .Padding = New System.Windows.Forms.Padding(8, 4, 8, 4)
            }

            Dim editButton As New System.Windows.Forms.Button With {
                .Text = If(hasLocal, "Edit Local", "Edit"),
                .AutoSize = True,
                .Margin = New System.Windows.Forms.Padding(3),
                .Padding = New System.Windows.Forms.Padding(8, 4, 8, 4)
            }

            buttonPanel.Controls.Add(insertButton)
            buttonPanel.Controls.Add(cancelButton)
            buttonPanel.Controls.Add(addButton)
            buttonPanel.Controls.Add(editButton)

            pickerForm.AcceptButton = insertButton
            pickerForm.CancelButton = cancelButton

            Dim visibleEntries As New List(Of PromptLibraryEntry)()
            Dim suppressCategoryRefresh As Boolean = False

            Dim GetDisplayTitle As System.Func(Of PromptLibraryEntry, String) =
                Function(entry As PromptLibraryEntry) As String
                    If entry Is Nothing Then Return ""

                    If String.Equals(entry.Category, PromptLibraryUncategorizedLabel, StringComparison.OrdinalIgnoreCase) Then
                        Return entry.Title
                    End If

                    Return $"{entry.Title} [{entry.Category}]"
                End Function

            Dim RefreshVisiblePromptList As System.Action(Of String) = Nothing
            Dim RefreshCategoryFilter As System.Action(Of String) = Nothing

            RefreshVisiblePromptList =
                Sub(preferredTitle As String)
                    Dim selectedCategory As String = PromptLibraryAllCategoriesLabel
                    If categoryComboBox.SelectedItem IsNot Nothing Then
                        selectedCategory = CStr(categoryComboBox.SelectedItem)
                    End If

                    Dim searchText As String = searchTextBox.Text.Trim()

                    visibleEntries.Clear()

                    For Each entry As PromptLibraryEntry In allEntries
                        Dim categoryMatches As Boolean =
                            String.Equals(selectedCategory, PromptLibraryAllCategoriesLabel, StringComparison.OrdinalIgnoreCase) OrElse
                            String.Equals(entry.Category, selectedCategory, StringComparison.OrdinalIgnoreCase)

                        If Not categoryMatches Then Continue For

                        Dim searchMatches As Boolean =
                            searchText.Length = 0 OrElse
                            entry.Title.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                            entry.Category.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0

                        If searchMatches Then
                            visibleEntries.Add(entry)
                        End If
                    Next

                    titleListBox.BeginUpdate()
                    titleListBox.Items.Clear()
                    For Each entry As PromptLibraryEntry In visibleEntries
                        titleListBox.Items.Add(GetDisplayTitle(entry))
                    Next
                    titleListBox.EndUpdate()

                    Dim selectedIndex As Integer = -1

                    If Not String.IsNullOrWhiteSpace(preferredTitle) Then
                        For i As Integer = 0 To visibleEntries.Count - 1
                            If String.Equals(visibleEntries(i).Title, preferredTitle, StringComparison.OrdinalIgnoreCase) Then
                                selectedIndex = i
                                Exit For
                            End If
                        Next
                    End If

                    If selectedIndex = -1 AndAlso visibleEntries.Count > 0 Then
                        selectedIndex = 0
                    End If

                    If selectedIndex >= 0 Then
                        titleListBox.SelectedIndex = selectedIndex
                        promptTextBox.Text = visibleEntries(selectedIndex).Prompt.Replace("\n", vbCrLf)
                    Else
                        titleListBox.ClearSelected()
                        promptTextBox.Clear()
                    End If

                    insertButton.Enabled = selectedIndex >= 0 AndAlso selectedIndex < visibleEntries.Count
                End Sub

            RefreshCategoryFilter =
                Sub(preferredCategory As String)
                    Dim categories As New List(Of String)()
                    Dim seenCategories As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                    Dim hasDefinedCategories As Boolean = False

                    categories.Add(PromptLibraryAllCategoriesLabel)

                    For Each entry As PromptLibraryEntry In allEntries
                        If String.Equals(entry.Category, PromptLibraryUncategorizedLabel, StringComparison.OrdinalIgnoreCase) Then
                            Continue For
                        End If

                        If seenCategories.Add(entry.Category) Then
                            categories.Add(entry.Category)
                            hasDefinedCategories = True
                        End If
                    Next

                    If hasDefinedCategories Then
                        For Each entry As PromptLibraryEntry In allEntries
                            If String.Equals(entry.Category, PromptLibraryUncategorizedLabel, StringComparison.OrdinalIgnoreCase) Then
                                If seenCategories.Add(entry.Category) Then
                                    categories.Add(entry.Category)
                                End If
                            End If
                        Next
                    End If

                    suppressCategoryRefresh = True

                    categoryComboBox.BeginUpdate()
                    categoryComboBox.Items.Clear()
                    categoryComboBox.Items.AddRange(categories.ToArray())
                    categoryComboBox.EndUpdate()

                    categoryComboBox.Enabled = hasDefinedCategories
                    categoryLabel.Enabled = hasDefinedCategories

                    Dim selectedIndex As Integer = 0

                    If hasDefinedCategories AndAlso Not String.IsNullOrWhiteSpace(preferredCategory) Then
                        For i As Integer = 0 To categories.Count - 1
                            If String.Equals(categories(i), preferredCategory, StringComparison.OrdinalIgnoreCase) Then
                                selectedIndex = i
                                Exit For
                            End If
                        Next
                    End If

                    If categoryComboBox.Items.Count > 0 Then
                        categoryComboBox.SelectedIndex = selectedIndex
                    End If

                    suppressCategoryRefresh = False
                End Sub

            AddHandler searchTextBox.TextChanged,
                Sub()
                    Dim currentTitle As String = Nothing

                    If titleListBox.SelectedIndex >= 0 AndAlso titleListBox.SelectedIndex < visibleEntries.Count Then
                        currentTitle = visibleEntries(titleListBox.SelectedIndex).Title
                    End If

                    RefreshVisiblePromptList(currentTitle)
                End Sub

            AddHandler categoryComboBox.SelectedIndexChanged,
                Sub()
                    If suppressCategoryRefresh Then Return
                    RefreshVisiblePromptList(Nothing)
                End Sub

            AddHandler titleListBox.SelectedIndexChanged,
                Sub()
                    Dim selectedIndex As Integer = titleListBox.SelectedIndex
                    If selectedIndex >= 0 AndAlso selectedIndex < visibleEntries.Count Then
                        promptTextBox.Text = visibleEntries(selectedIndex).Prompt.Replace("\n", vbCrLf)
                        insertButton.Enabled = True
                    Else
                        promptTextBox.Clear()
                        insertButton.Enabled = False
                    End If
                End Sub

            AddHandler titleListBox.DoubleClick,
                Sub()
                    If titleListBox.SelectedIndex >= 0 Then
                        pickerForm.DialogResult = System.Windows.Forms.DialogResult.OK
                        pickerForm.Close()
                    End If
                End Sub

            AddHandler titleListBox.KeyDown,
                Sub(sender As Object, e As System.Windows.Forms.KeyEventArgs)
                    If e.KeyCode = System.Windows.Forms.Keys.Enter Then
                        e.SuppressKeyPress = True
                        If titleListBox.SelectedIndex >= 0 Then
                            pickerForm.DialogResult = System.Windows.Forms.DialogResult.OK
                            pickerForm.Close()
                        End If
                    End If
                End Sub

            AddHandler addButton.Click,
                Sub()
                    Try
                        Dim target As String = GetEditablePromptLibraryPath(localPath, centralPath)
                        If String.IsNullOrWhiteSpace(target) Then
                            ShowCustomMessageBox("No editable prompt library file is configured.")
                            Return
                        End If

                        Dim selectedCategory As String = Nothing
                        If categoryComboBox.SelectedItem IsNot Nothing Then
                            selectedCategory = CStr(categoryComboBox.SelectedItem)
                        End If

                        If String.Equals(selectedCategory, PromptLibraryAllCategoriesLabel, StringComparison.OrdinalIgnoreCase) OrElse
                           String.Equals(selectedCategory, PromptLibraryUncategorizedLabel, StringComparison.OrdinalIgnoreCase) Then
                            selectedCategory = Nothing
                        End If

                        Dim addedTitle As String = Nothing
                        Dim addedPrompt As String = Nothing
                        Dim addedCategory As String = Nothing

                        If Not ShowAddPromptLibraryEntryDialog(target, selectedCategory, lastPromptForCtrlP, addedTitle, addedPrompt, addedCategory, ownerHandle) Then
                            Return
                        End If

                        AppendPromptLibraryEntry(target, addedTitle, addedPrompt, addedCategory)

                        ReloadPromptEntries()

                        Dim preferredCategory As String =
                            If(String.IsNullOrWhiteSpace(addedCategory), PromptLibraryAllCategoriesLabel, addedCategory)

                        RefreshCategoryFilter(preferredCategory)

                        Dim preferredTitle As String = addedTitle
                        If hasLocal AndAlso String.Equals(target, localPath, StringComparison.OrdinalIgnoreCase) Then
                            preferredTitle &= " (local)"
                        End If

                        RefreshVisiblePromptList(preferredTitle)
                        titleListBox.Focus()
                    Catch ex As Exception
                        ShowCustomMessageBox("Failed to add prompt: " & ex.Message)
                    End Try
                End Sub

            AddHandler editButton.Click,
                Sub()
                    Try
                        Dim previousCategory As String = PromptLibraryAllCategoriesLabel
                        If categoryComboBox.SelectedItem IsNot Nothing Then
                            previousCategory = CStr(categoryComboBox.SelectedItem)
                        End If

                        Dim previousTitle As String = Nothing
                        If titleListBox.SelectedIndex >= 0 AndAlso titleListBox.SelectedIndex < visibleEntries.Count Then
                            previousTitle = visibleEntries(titleListBox.SelectedIndex).Title
                        End If

                        Dim target As String = GetEditablePromptLibraryPath(localPath, centralPath)
                        If String.IsNullOrWhiteSpace(target) Then
                            ShowCustomMessageBox("No editable prompt library file is configured.")
                            Return
                        End If

                        Dim targetKind As String = If(hasLocal, "local", "central")

                        EnsurePromptLibraryDirectoryExists(target)

                        ShowTextFileEditor(
                                target,
                                $"You can now edit your {targetKind} prompts (stored at {target}). Make sure that on each prompt line, the description and the prompt are separated by a '|'; you can use ';' for comments; optional category blocks use <Category Name> and </Category Name>.",
                                ownerHandle:=ownerHandle
                            )

                        ReloadPromptEntries()
                        RefreshCategoryFilter(previousCategory)
                        RefreshVisiblePromptList(previousTitle)

                        titleListBox.Focus()
                    Catch ex As Exception
                        ShowCustomMessageBox("Failed to open the prompt library editor: " & ex.Message)
                    End Try
                End Sub

            RefreshCategoryFilter(PromptLibraryAllCategoriesLabel)
            RefreshVisiblePromptList(Nothing)

            AddHandler pickerForm.Shown,
                    Sub()
                        searchTextBox.Focus()
                        searchTextBox.SelectionStart = 0
                        searchTextBox.SelectionLength = searchTextBox.TextLength
                        Try
                            Global.SharedLibrary.SharedLibrary.SharedMethods.ForceDialogToForeground(pickerForm)
                            Global.SharedLibrary.SharedLibrary.SharedMethods.AttachForeignForegroundWatchdog(pickerForm)
                        Catch ex As System.Exception
                        End Try
                    End Sub

            Dim result As System.Windows.Forms.DialogResult
            Dim owner As System.Windows.Forms.IWin32Window = Nothing

            If ownerHandle <> System.IntPtr.Zero Then
                owner = New WindowWrapper(ownerHandle)
            Else
                owner = System.Windows.Forms.Form.ActiveForm
            End If

            ' Inspect and same-thread filter the resolved owner before using it as a modal
            ' owner. A foreign owner (e.g. another Office process's main window, whether passed
            ' in via ownerHandle or picked up as ActiveForm) would be disabled by ShowDialog and
            ' never re-enabled, deadlocking that host. InspectDialogOwner logs cross-process
            ' attempts; IfOwnerOnCurrentThread rejects a cross-thread owner so we fall back to
            ' an ownerless dialog instead of deadlocking.
            If owner IsNot Nothing Then
                OfficeWindowWatchdog.InspectDialogOwner(owner, "ShowPromptInsertionSelector", "ShowPromptInsertionSelector")
                owner = SharedMethods.IfOwnerOnCurrentThread(owner)
            End If

            If owner IsNot Nothing Then
                result = pickerForm.ShowDialog(owner)
            Else
                result = pickerForm.ShowDialog()
            End If

            If result = System.Windows.Forms.DialogResult.OK Then
                Dim selectedIndex As Integer = titleListBox.SelectedIndex
                If selectedIndex >= 0 AndAlso selectedIndex < visibleEntries.Count Then
                    Dim selectedPrompt As String = visibleEntries(selectedIndex).Prompt
                    If Not NoWarning Then WarnIfPromptContainsUnsupportedChatPlaceholders(selectedPrompt)
                    Return selectedPrompt
                End If
            End If

            Return ""
        End Function


        ''' <summary>
        ''' Handles slash-triggered prompt insertion for chat input text boxes.
        ''' </summary>
        Public Shared Function HandlePromptLibrarySlash(targetTextBox As System.Windows.Forms.TextBoxBase,
                                                        filePath As String,
                                                        filepathlocal As String,
                                                        Context As ISharedContext,
                                                        Optional lastPromptForCtrlP As String = Nothing,
                                                        Optional NoWarning As Boolean = False) As PromptLibrarySlashAction

            If targetTextBox Is Nothing Then
                Return PromptLibrarySlashAction.NotTriggered
            End If

            If Not IsPromptLibrarySlashTrigger(targetTextBox) Then
                Return PromptLibrarySlashAction.NotTriggered
            End If

            Dim selectedPrompt As String =
                ShowPromptInsertionSelector(filePath, filepathlocal, Context, Nothing, lastPromptForCtrlP, NoWarning:=NoWarning)

            If String.IsNullOrEmpty(selectedPrompt) Then
                Return PromptLibrarySlashAction.Canceled
            End If

            Dim insertText As String = selectedPrompt.Replace("\n", Environment.NewLine)
            Dim insertionIndex As Integer = targetTextBox.SelectionStart
            Dim selectionLength As Integer = targetTextBox.SelectionLength

            Dim newText As String = targetTextBox.Text.Remove(insertionIndex, selectionLength).Insert(insertionIndex, insertText)
            targetTextBox.Text = newText
            targetTextBox.SelectionStart = insertionIndex + insertText.Length
            targetTextBox.SelectionLength = 0
            targetTextBox.Focus()

            Return PromptLibrarySlashAction.Inserted
        End Function

    End Class

End Namespace
