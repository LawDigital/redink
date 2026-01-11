' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: MultiModelSelectorForm.vb
' Purpose: Implements a modal Windows Forms dialog that lets the user select multiple
'          alternative model configurations from a provided list, with filtering and
'          a persisted check state across filter changes.
'
' Architecture:
'  - UI Layout: Uses a single-column `TableLayoutPanel` (`outer`):
'      Row 1: Title label (`lblTitle`)
'      Row 2: Filter textbox (`txtFilter`) with a Win32 cue-banner placeholder
'      Row 3: Checked list of models (`chkList`) supporting multi-selection
'      Row 4: Optional reset checkbox (`chkReset`) (currently hidden via `.Visible = False`)
'      Row 5: OK/Cancel buttons in a right-aligned `FlowLayoutPanel` (`pnlButtons`)
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

        ''' <summary>Panel containing the OK/Cancel buttons.</summary>
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
            InitializeComponent(title, resetChecked)
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
                        ' by doing the check there (see PopulateList changes below).
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
                Dim display As System.String = If(Not String.IsNullOrWhiteSpace(m.ModelDescription), m.ModelDescription, m.Model)
                Dim unique As System.String = MakeUniqueDisplay(display)
                displayToModel(unique) = m
                allDisplayItems.Add(unique)

                ' NEW: seed multi-preselect into selectedLabels (checked state)
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
        End Sub

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
            Me.Width = 520
            Me.Height = 460
            Me.MinimumSize = New System.Drawing.Size(520, 460)

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
                .Height = 28,
                .TextAlign = System.Drawing.ContentAlignment.MiddleLeft
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
                .Visible = False            ' Hide it for the time being
            }

            Me.pnlButtons = New System.Windows.Forms.FlowLayoutPanel() With {
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft,
                .Padding = New System.Windows.Forms.Padding(0, 8, 0, 0),
                .AutoSize = True
            }

            ' Match ModelSelectorForm: autosize, GrowAndShrink, and padding (10,5,10,5)
            Me.btnOK = New System.Windows.Forms.Button() With {
                .Text = "OK",
                .DialogResult = System.Windows.Forms.DialogResult.OK,
                .AutoSize = True,
                .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                .Padding = New System.Windows.Forms.Padding(10, 5, 10, 5)
            }
            Me.btnCancel = New System.Windows.Forms.Button() With {
                .Text = "Cancel",
                .DialogResult = System.Windows.Forms.DialogResult.Cancel,
                .AutoSize = True,
                .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                .Padding = New System.Windows.Forms.Padding(10, 5, 10, 5)
            }
            Me.pnlButtons.Controls.Add(Me.btnOK)
            Me.pnlButtons.Controls.Add(Me.btnCancel)

            Me.outer.Controls.Add(Me.lblTitle, 0, 0)
            Me.outer.Controls.Add(Me.txtFilter, 0, 1)
            Me.outer.Controls.Add(Me.chkList, 0, 2)
            Me.outer.Controls.Add(Me.chkReset, 0, 3)
            Me.outer.Controls.Add(Me.pnlButtons, 0, 4)
            Me.Controls.Add(Me.outer)

            Me.AcceptButton = Me.btnOK
            Me.CancelButton = Me.btnCancel
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
                        ' Restore check state from the persisted selection set
                        Dim isChecked = selectedLabels.Contains(itemText)
                        Me.chkList.Items.Add(itemText, isChecked)
                    End If
                Next
            Finally
                Me.chkList.EndUpdate()
                isUpdating = False
            End Try
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
            If String.IsNullOrWhiteSpace(preselectKey) Then Return

            ' Try by label first
            For i = 0 To Me.chkList.Items.Count - 1
                Dim label As System.String = Me.chkList.Items(i).ToString()
                If String.Equals(label, preselectKey, System.StringComparison.OrdinalIgnoreCase) Then
                    Me.chkList.SetItemChecked(i, True)
                    ' selectedLabels will be updated by ItemCheck handler
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
                ' selectedLabels will be updated by ItemCheck handler
            End If
        End Sub
    End Class


End Namespace