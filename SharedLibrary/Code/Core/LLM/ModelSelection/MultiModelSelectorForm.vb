' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: MultiModelSelectorForm.vb
' Purpose: Implements a modal Windows Forms dialog that lets the user select multiple
'          alternative model configurations from a provided list, with filtering,
'          persisted check state across filter changes, and a Select All / Unselect All toggle.
'
' Architecture:
'  - UI Layout: Uses a single-column `TableLayoutPanel` (`outer`):
'      Row 1: Title label (`lblTitle`)
'      Row 2: Filter textbox (`txtFilter`) with a Win32 cue-banner placeholder
'      Row 3: Checked list of models (`chkList`) supporting multi-selection
'      Row 4: Optional reset checkbox (`chkReset`) (currently hidden via `.Visible = False`)
'      Row 5: OK/Cancel/toggle buttons in a right-aligned `FlowLayoutPanel` (`pnlButtons`)
'  - Model Source: Receives a list of `ModelConfig` (`altModels`) from the caller and maps each
'    display label to its `ModelConfig` via `displayToModel`.
'  - Display Labels: Uses `ModelDescription` when available; otherwise uses `Model`. Ensures a
'    unique display label via `MakeUniqueDisplay` to prevent collisions in `displayToModel`.
'  - Filtering: `txtFilter` filters the visible items in the checked list without losing selections
'    by persisting selected labels in `selectedLabels`.
'  - Selection Output: `SelectedModels` returns all selected (`checked`) `ModelConfig` items based on
'    the persisted label set (not only the currently visible items).
'  - Preselection: Supports single-key preselection (`preselectKey`) and multi-key preselection
'    (`preselectKeys`) that is resolved during `PopulateList` and `ApplyPreselection`.
'  - Bulk Selection: `btnToggleAll` toggles the check state of all currently visible items and adapts
'    its text between "Select All" and "Unselect All".
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Windows.Forms

Namespace SharedLibrary

    ''' <summary>
    ''' Modal dialog that supports selecting multiple alternate models, with filtering and
    ''' persisted selection state across filter changes.
    ''' </summary>
    Public Class MultiModelSelectorForm
        Inherits System.Windows.Forms.Form

        ''' <summary>Title label displayed above the filter input.</summary>
        Private lblTitle As System.Windows.Forms.Label

        ''' <summary>Filter textbox used to restrict visible model entries.</summary>
        Private txtFilter As System.Windows.Forms.TextBox

        ''' <summary>Checked list box holding model display labels with multi-check selection.</summary>
        Private chkList As System.Windows.Forms.CheckedListBox

        ''' <summary>Optional checkbox to control whether the selection should reset to the default model after use.</summary>
        Private chkReset As System.Windows.Forms.CheckBox

        ''' <summary>OK button closing the dialog with <see cref="DialogResult.OK"/>.</summary>
        Private btnOK As Button

        ''' <summary>Cancel button closing the dialog with <see cref="DialogResult.Cancel"/>.</summary>
        Private btnCancel As Button

        ''' <summary>Toggle button to select or unselect all currently visible items.</summary>
        Private btnToggleAll As Button

        ''' <summary>Panel containing the dialog buttons.</summary>
        Private pnlButtons As System.Windows.Forms.FlowLayoutPanel

        ''' <summary>Root layout container.</summary>
        Private outer As System.Windows.Forms.TableLayoutPanel

        ''' <summary>Maps unique display labels to their corresponding <see cref="ModelConfig"/> instance.</summary>
        Private displayToModel As New System.Collections.Generic.Dictionary(Of String, ModelConfig)(System.StringComparer.OrdinalIgnoreCase)

        ''' <summary>Tracks already-used display labels to avoid collisions.</summary>
        Private seenDisplays As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.OrdinalIgnoreCase)

        ''' <summary>All model display labels (unfiltered master list).</summary>
        Private allDisplayItems As New System.Collections.Generic.List(Of String)

        ''' <summary>Optional label or model key to preselect when the dialog opens.</summary>
        Private preselectKey As String = Nothing

        ''' <summary>Original list of alternative models provided by the caller.</summary>
        Private ReadOnly altModels As System.Collections.Generic.List(Of ModelConfig)

        ''' <summary>
        ''' Persists checked selections across filtering by storing selected display labels.
        ''' </summary>
        Private ReadOnly selectedLabels As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.OrdinalIgnoreCase)

        ''' <summary>
        ''' Guards against re-entrancy and event-driven updates while programmatically rebuilding list items.
        ''' </summary>
        Private isUpdating As Boolean = False

        ''' <summary>
        ''' Win32 message ID for setting a cue banner (placeholder) on an Edit control.
        ''' </summary>
        Private Const EM_SETCUEBANNER As Integer = &H1501

        ''' <summary>Optional set of labels/keys to preselect when the dialog opens (multi-select).</summary>
        Private ReadOnly preselectKeys As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.OrdinalIgnoreCase)

        ''' <summary>
        ''' Sends a Win32 message to a window handle (used for setting the cue banner on <see cref="txtFilter"/>).
        ''' </summary>
        <System.Runtime.InteropServices.DllImport("user32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Unicode)>
        Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As String) As IntPtr
        End Function

        ''' <summary>
        ''' Gets the selected model configurations based on the persisted set of checked display labels.
        ''' </summary>
        Public ReadOnly Property SelectedModels As System.Collections.Generic.List(Of ModelConfig)
            Get
                ' Return models based on the persisted selectedLabels (not only the currently visible items)
                Dim result As New System.Collections.Generic.List(Of ModelConfig)
                For Each key In selectedLabels
                    If displayToModel.ContainsKey(key) Then
                        result.Add(displayToModel(key))
                    End If
                Next
                Return result
            End Get
        End Property

        ''' <summary>
        ''' Gets whether the default model should be restored after use (based on the reset checkbox state).
        ''' </summary>
        Public ReadOnly Property UseDefault As Boolean
            Get
                Return chkReset.Checked
            End Get
        End Property

        Protected Overrides Sub OnShown(e As System.EventArgs)
            MyBase.OnShown(e)
            SharedMethods.AttachForeignForegroundWatchdog(Me)
        End Sub

        ''' <summary>
        ''' Initializes a new instance of the multi-model selector dialog.
        ''' </summary>
        ''' <param name="models">Alternative model configurations to display.</param>
        ''' <param name="preselect">Optional preselection key (label, <see cref="ModelConfig.ModelDescription"/>, or <see cref="ModelConfig.Model"/>).</param>
        ''' <param name="title">Optional window title.</param>
        ''' <param name="resetChecked">Initial value for the reset checkbox.</param>
        ''' <param name="instruction">Optional instruction text displayed above the filter input.</param>
        Public Sub New(models As System.Collections.Generic.List(Of ModelConfig),
                   preselect As System.String,
                   Optional title As System.String = Nothing,
                   Optional resetChecked As System.Boolean = True,
                   Optional instruction As System.String = "")
            Me.altModels = If(models, New System.Collections.Generic.List(Of ModelConfig))
            Me.preselectKey = preselect
            InitializeComponent(title, resetChecked, instruction)
            PopulateList()
            ApplyPreselection()
        End Sub

        ''' <summary>
        ''' Initializes a new instance of the multi-model selector dialog with multi-selection precheck support.
        ''' </summary>
        ''' <param name="models">Alternative model configurations to display.</param>
        ''' <param name="preselect">Optional preselection key (label, ModelDescription, or Model).</param>
        ''' <param name="title">Optional window title.</param>
        ''' <param name="resetChecked">Initial value for the reset checkbox.</param>
        ''' <param name="preselectMany">Optional set of labels/keys to preselect (multi-select).</param>
        ''' <param name="instruction">Optional instruction text displayed above the filter input.</param>
        Public Sub New(models As System.Collections.Generic.List(Of ModelConfig),
                   preselect As System.String,
                   Optional title As System.String = Nothing,
                   Optional resetChecked As System.Boolean = True,
                   Optional preselectMany As System.Collections.Generic.IEnumerable(Of System.String) = Nothing,
                   Optional instruction As String = "")

            Me.altModels = If(models, New System.Collections.Generic.List(Of ModelConfig))
            Me.preselectKey = preselect

            If preselectMany IsNot Nothing Then
                For Each k In preselectMany
                    If Not String.IsNullOrWhiteSpace(k) Then preselectKeys.Add(k.Trim())
                Next
            End If

            InitializeComponent(title, resetChecked, instruction)

            ' Seed selectedLabels from preselectMany by resolving against unique display labels.
            If preselectKeys.Count > 0 Then
                For Each m In Me.altModels
                    Dim display As String = If(Not String.IsNullOrWhiteSpace(m.ModelDescription), m.ModelDescription, m.Model)
                    If preselectKeys.Contains(display) OrElse preselectKeys.Contains(m.Model) OrElse preselectKeys.Contains(m.ModelDescription) Then
                        ' We don't know the final unique display label yet; we’ll map it during PopulateList
                        ' by doing the check there.
                    End If
                Next
            End If

            PopulateList()
            ApplyPreselection()
        End Sub

        ''' <summary>
        ''' Populates the checked list with all available model display labels and restores any persisted selection.
        ''' Also builds the internal label-to-model map used by <see cref="SelectedModels"/>.
        ''' </summary>
        ''' <remarks>
        ''' This method rebuilds <see cref="displayToModel"/>, <see cref="seenDisplays"/>, and <see cref="allDisplayItems"/>.
        ''' If <see cref="preselectKeys"/> contains entries, matching labels/models are pre-checked by seeding
        ''' <see cref="selectedLabels"/> before the UI list is rendered.
        ''' </remarks>
        Private Sub PopulateList()
            displayToModel.Clear()
            seenDisplays.Clear()
            allDisplayItems.Clear()

            For Each m In altModels
                Dim display As System.String = GetDisplayTextForList(m)
                Dim unique As System.String = MakeUniqueDisplay(display)
                displayToModel(unique) = m
                allDisplayItems.Add(unique)

                If preselectKeys.Count > 0 Then
                    If preselectKeys.Contains(unique) OrElse
               preselectKeys.Contains(display) OrElse
               preselectKeys.Contains(m.Model) OrElse
               preselectKeys.Contains(m.ModelDescription) OrElse
               preselectKeys.Contains(m.ToolName) Then
                        selectedLabels.Add(unique)
                    End If
                End If
            Next

            isUpdating = True
            Try
                Me.chkList.Items.Clear()
                For Each label In allDisplayItems
                    Me.chkList.Items.Add(label, selectedLabels.Contains(label))
                Next
            Finally
                isUpdating = False
            End Try

            UpdateToggleAllButtonText()
            UpdatePreferredDialogWidth()
        End Sub

        Private Shared Function StripTrailingDisplaySuffix(value As String, suffix As String) As String
            Dim result As String = If(value, "")

            If suffix = "" Then
                Return result
            End If

            Do While result.EndsWith(suffix, System.StringComparison.OrdinalIgnoreCase)
                result = result.Substring(0, result.Length - suffix.Length).TrimEnd()
            Loop

            Return result
        End Function

        Private Shared Function RemoveLegacyInternalDisplaySuffixes(value As String) As String
            Dim result As String = If(value, "").Trim()

            result = StripTrailingDisplaySuffix(result, " (Outlook only)")
            result = StripTrailingDisplaySuffix(result, " (Word only)")
            result = StripTrailingDisplaySuffix(result, " (local only)")
            result = StripTrailingDisplaySuffix(result, " (built-in)")
            result = StripTrailingDisplaySuffix(result, " (internal)")

            Return result
        End Function

        Private Shared Function GetDisplayTextForList(m As ModelConfig) As System.String
            If m Is Nothing Then
                Return "(Unnamed model)"
            End If

            Dim display As System.String =
                If(Not String.IsNullOrWhiteSpace(m.ModelDescription), m.ModelDescription, m.Model)

            If String.IsNullOrWhiteSpace(m.ToolName) Then
                Return display
            End If

            Dim suffix As String =
                Global.SharedLibrary.Agents.HostToolRegistration.GetSelectorDisplaySuffix(m.ToolName)

            If suffix = "" Then
                Return display
            End If

            Return RemoveLegacyInternalDisplaySuffixes(display) & suffix
        End Function

        Private Sub UpdateInstructionLabelLayout()
            If Me.lblTitle Is Nothing OrElse Me.outer Is Nothing Then
                Return
            End If

            Dim width As Integer = Math.Max(320, Me.ClientSize.Width - Me.outer.Padding.Left - Me.outer.Padding.Right - 6)
            Me.lblTitle.MaximumSize = New System.Drawing.Size(width, 0)
        End Sub

        ''' <summary>
        ''' Creates and configures all UI controls and event handlers for the dialog.
        ''' </summary>
        ''' <summary>
        ''' Creates and configures all UI controls and event handlers for the dialog.
        ''' </summary>
        Private Sub InitializeComponent(Optional title As System.String = Nothing, Optional resetChecked As System.Boolean = True, Optional instruction As System.String = "Select one or more alternate models:")
            Me.Text = If(String.IsNullOrWhiteSpace(title), SharedMethods.AN & " - Select Alternate Models", title)
            Me.Icon = Icon.FromHandle((New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))).GetHicon())
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
            Me.MinimizeBox = True
            Me.MaximizeBox = True
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
            Me.Width = 640
            Me.Height = 600
            Me.MinimumSize = New System.Drawing.Size(640, 600)
            Me.TopMost = True

            Me.outer = New System.Windows.Forms.TableLayoutPanel() With {
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 5,
                .Padding = New System.Windows.Forms.Padding(16, 12, 16, 12)
            }
            Me.outer.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            Me.outer.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            Me.outer.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
            Me.outer.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            Me.outer.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))

            Me.lblTitle = New System.Windows.Forms.Label() With {
                .Text = instruction,
                .Dock = System.Windows.Forms.DockStyle.Top,
                .AutoSize = True,
                .MaximumSize = New System.Drawing.Size(580, 0),
                .TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                .Margin = New System.Windows.Forms.Padding(0, 0, 0, 8)
            }

            Me.txtFilter = New System.Windows.Forms.TextBox() With {
                .Dock = System.Windows.Forms.DockStyle.Top
            }
            AddHandler Me.txtFilter.HandleCreated,
                Sub()
                    Try
                        Dim showEvenIfFocused As IntPtr = CType(1, IntPtr)
                        SendMessage(Me.txtFilter.Handle, EM_SETCUEBANNER, showEvenIfFocused, "Filter…")
                    Catch
                    End Try
                End Sub
            AddHandler Me.txtFilter.TextChanged, AddressOf OnFilterChanged

            Me.chkList = New System.Windows.Forms.CheckedListBox() With {
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .CheckOnClick = True
            }
            AddHandler Me.chkList.DoubleClick, AddressOf OnListDoubleClick
            AddHandler Me.chkList.ItemCheck, AddressOf OnItemCheck

            Me.chkReset = New System.Windows.Forms.CheckBox() With {
                .Text = "Reset to default model after use",
                .Dock = System.Windows.Forms.DockStyle.Top,
                .Checked = resetChecked,
                .Visible = False
            }

            Me.pnlButtons = New System.Windows.Forms.FlowLayoutPanel() With {
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft,
                .Padding = New System.Windows.Forms.Padding(0, 8, 0, 0),
                .AutoSize = True,
                .WrapContents = False
            }

            Me.btnOK = New System.Windows.Forms.Button() With {
                .Text = "OK",
                .DialogResult = System.Windows.Forms.DialogResult.OK
            }
            ConfigureStandardButton(Me.btnOK)

            Me.btnCancel = New System.Windows.Forms.Button() With {
                .Text = "Cancel",
                .DialogResult = System.Windows.Forms.DialogResult.Cancel
            }
            ConfigureStandardButton(Me.btnCancel)

            Me.btnToggleAll = New System.Windows.Forms.Button() With {
                .Text = "Select All"
            }
            ConfigureStandardButton(Me.btnToggleAll)

            AddHandler Me.btnToggleAll.Click, AddressOf OnToggleAllClicked

            Me.pnlButtons.Controls.Add(Me.btnCancel)
            Me.pnlButtons.Controls.Add(Me.btnOK)
            Me.pnlButtons.Controls.Add(Me.btnToggleAll)

            Me.outer.Controls.Add(Me.lblTitle, 0, 0)
            Me.outer.Controls.Add(Me.txtFilter, 0, 1)
            Me.outer.Controls.Add(Me.chkList, 0, 2)
            Me.outer.Controls.Add(Me.chkReset, 0, 3)
            Me.outer.Controls.Add(Me.pnlButtons, 0, 4)
            Me.Controls.Add(Me.outer)

            AddHandler Me.Shown,
                    Sub()
                        UpdateInstructionLabelLayout()
                    End Sub

            AddHandler Me.Resize,
                Sub()
                    UpdateInstructionLabelLayout()
                End Sub

            Me.AcceptButton = Me.btnOK
            Me.CancelButton = Me.btnCancel

            AddHandler Me.Shown,
                Sub()
                    UpdatePreferredDialogWidth()
                End Sub
        End Sub

        Private Sub UpdatePreferredDialogWidth()
            If Me.chkList Is Nothing OrElse Me.pnlButtons Is Nothing Then Return

            Dim buttonsWidth As Integer = Me.pnlButtons.PreferredSize.Width

            Dim widestItemWidth As Integer = 0
            For Each itemText In allDisplayItems
                Dim itemWidth = TextRenderer.MeasureText(itemText, Me.chkList.Font).Width
                If itemWidth > widestItemWidth Then
                    widestItemWidth = itemWidth
                End If
            Next

            Dim listChromeWidth As Integer = 40 ' checkbox + list padding + scrollbar margin
            Dim contentWidth As Integer = Math.Max(buttonsWidth, widestItemWidth + listChromeWidth)

            Dim outerPaddingWidth As Integer = Me.outer.Padding.Left + Me.outer.Padding.Right
            Dim formNonClientWidth As Integer = Me.Width - Me.ClientSize.Width
            Dim smallMarginWidth As Integer = 24

            Dim preferredWidth As Integer = contentWidth + outerPaddingWidth + formNonClientWidth + smallMarginWidth
            Dim maxWidth As Integer = Screen.FromControl(Me).WorkingArea.Width - 80

            preferredWidth = Math.Max(640, Math.Min(preferredWidth, maxWidth))

            Me.Width = preferredWidth
            Me.MinimumSize = New Size(preferredWidth, 600)
        End Sub

        Private Shared Sub ConfigureStandardButton(button As Button)
            button.AutoSize = True
            button.AutoSizeMode = AutoSizeMode.GrowAndShrink
            button.UseVisualStyleBackColor = True
            button.Padding = New Padding(10, 4, 10, 4)
            button.MinimumSize = New Size(0, 0)
            button.Margin = New Padding(6, 0, 0, 0)
        End Sub

        ''' <summary>
        ''' Returns a display label that is unique within this dialog instance.
        ''' </summary>
        ''' <param name="baseText">Proposed display text.</param>
        ''' <returns>A unique display label (with numeric suffix if needed).</returns>
        Private Function MakeUniqueDisplay(baseText As System.String) As System.String
            Dim s As System.String = If(String.IsNullOrWhiteSpace(baseText), "(Unnamed model)", baseText.Trim())
            Dim unique As System.String = s
            Dim suffix As Integer = 2
            While seenDisplays.Contains(unique)
                unique = s & " (" & suffix.ToString() & ")"
                suffix += 1
            End While
            seenDisplays.Add(unique)
            Return unique
        End Function

        ''' <summary>
        ''' Returns True when all currently visible items are checked.
        ''' </summary>
        Private Function AreAllVisibleItemsChecked() As Boolean
            If Me.chkList.Items.Count = 0 Then
                Return False
            End If

            For i As Integer = 0 To Me.chkList.Items.Count - 1
                If Not Me.chkList.GetItemChecked(i) Then
                    Return False
                End If
            Next

            Return True
        End Function

        ''' <summary>
        ''' Updates the toggle button caption based on the current visible check state.
        ''' </summary>
        Private Sub UpdateToggleAllButtonText()
            If Me.btnToggleAll Is Nothing Then
                Return
            End If

            Me.btnToggleAll.Text = If(AreAllVisibleItemsChecked(), "Unselect All", "Select All")
        End Sub

        ''' <summary>
        ''' Toggles all currently visible items between checked and unchecked.
        ''' </summary>
        Private Sub OnToggleAllClicked(sender As Object, e As EventArgs)
            Dim shouldCheckAll As Boolean = Not AreAllVisibleItemsChecked()

            isUpdating = True
            Me.chkList.BeginUpdate()

            Try
                For i As Integer = 0 To Me.chkList.Items.Count - 1
                    Dim label As String = Me.chkList.Items(i).ToString()
                    Me.chkList.SetItemChecked(i, shouldCheckAll)

                    If shouldCheckAll Then
                        selectedLabels.Add(label)
                    Else
                        selectedLabels.Remove(label)
                    End If
                Next
            Finally
                Me.chkList.EndUpdate()
                isUpdating = False
            End Try

            UpdateToggleAllButtonText()
        End Sub

        ''' <summary>
        ''' Rebuilds the visible checked list based on the current filter text.
        ''' Check-state is restored from <see cref="selectedLabels"/>.
        ''' </summary>
        Private Sub OnFilterChanged(sender As System.Object, e As System.EventArgs)
            Dim filter As System.String = If(Me.txtFilter.Text, String.Empty).Trim().ToLowerInvariant()

            isUpdating = True
            Me.chkList.BeginUpdate()
            Try
                Me.chkList.Items.Clear()
                For Each itemText In allDisplayItems
                    If filter.Length = 0 OrElse itemText.ToLowerInvariant().Contains(filter) Then
                        Dim isChecked = selectedLabels.Contains(itemText)
                        Me.chkList.Items.Add(itemText, isChecked)
                    End If
                Next
            Finally
                Me.chkList.EndUpdate()
                isUpdating = False
            End Try

            UpdateToggleAllButtonText()
            UpdatePreferredDialogWidth()
        End Sub

        ''' <summary>
        ''' Updates <see cref="selectedLabels"/> when a list item check state changes.
        ''' </summary>
        Private Sub OnItemCheck(sender As Object, e As System.Windows.Forms.ItemCheckEventArgs)
            If isUpdating Then Return

            Dim label As String = Me.chkList.Items(e.Index).ToString()
            If e.NewValue = CheckState.Checked Then
                selectedLabels.Add(label)
            Else
                selectedLabels.Remove(label)
            End If

            Me.BeginInvoke(New MethodInvoker(AddressOf UpdateToggleAllButtonText))
        End Sub

        ''' <summary>
        ''' Toggles the checked state of the currently selected list item on double-click.
        ''' </summary>
        Private Sub OnListDoubleClick(sender As System.Object, e As System.EventArgs)
            Dim idx As Integer = Me.chkList.SelectedIndex
            If idx >= 0 Then
                Dim state As Boolean = Not Me.chkList.GetItemChecked(idx)
                Me.chkList.SetItemChecked(idx, state)
            End If
        End Sub

        ''' <summary>
        ''' Applies an initial preselection by matching the provided preselect key against
        ''' the display label or underlying model values.
        ''' </summary>
        Private Sub ApplyPreselection()
            If String.IsNullOrWhiteSpace(preselectKey) Then
                UpdateToggleAllButtonText()
                Return
            End If

            ' Try by label first
            For i = 0 To Me.chkList.Items.Count - 1
                Dim label As System.String = Me.chkList.Items(i).ToString()
                If String.Equals(label, preselectKey, System.StringComparison.OrdinalIgnoreCase) Then
                    Me.chkList.SetItemChecked(i, True)
                    UpdateToggleAllButtonText()
                    Return
                End If
            Next

            ' Fallback: try to match underlying ModelDescription/Model
            Dim idxToCheck As Integer = -1
            For j = 0 To Me.chkList.Items.Count - 1
                Dim label As System.String = Me.chkList.Items(j).ToString()
                If displayToModel.ContainsKey(label) Then
                    Dim mc As ModelConfig = displayToModel(label)
                    If String.Equals(mc.ModelDescription, preselectKey, System.StringComparison.OrdinalIgnoreCase) _
                   OrElse String.Equals(mc.Model, preselectKey, System.StringComparison.OrdinalIgnoreCase) Then
                        idxToCheck = j
                        Exit For
                    End If
                End If
            Next

            If idxToCheck >= 0 Then
                Me.chkList.SetItemChecked(idxToCheck, True)
            End If

            UpdateToggleAllButtonText()
        End Sub

        ''' <summary>
        ''' Adds an extra button (left-aligned next to Select-All/OK/Cancel) that invokes
        ''' the supplied handler when clicked. Use this to surface secondary actions such
        ''' as "Manage Skills &amp; Agents…" or "Memory…" from the source-selection window
        ''' without subclassing the form.
        ''' </summary>
        Public Sub AddExtraButton(text As String, handler As EventHandler)
            If String.IsNullOrWhiteSpace(text) OrElse handler Is Nothing Then Return

            Dim btn As New Button() With {
                .Text = text
            }
            ConfigureStandardButton(btn)

            AddHandler btn.Click, handler

            Me.pnlButtons.Controls.Add(btn)
            UpdatePreferredDialogWidth()
        End Sub

    End Class

End Namespace
