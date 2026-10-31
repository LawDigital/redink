' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: MultiModelSelectorForm.Grouping.vb
' PURPOSE
'   Adds opt-in group presentation to the existing multi-model selector without changing selection semantics.
'
' KEY ARCHITECTURE
'   • Adapter: Opt-in grouped selector behavior wraps the existing flat form.
'   • Compatibility: Legacy callers and ungrouped selections preserve existing semantics.
' =============================================================================

' Opt-in adapter. The legacy flat selector and all callers that do not configure
' grouping retain their existing behavior. Stable ToolName IDs are used only here.
Option Strict On
Option Explicit On

Namespace SharedLibrary
    Partial Public Class MultiModelSelectorForm
        Private _groupingStore As SelectorGroupingStore
        Private _groupItems As System.Collections.Generic.List(Of SelectorGroupItem)
        Private _groupedView As SelectorGroupedList
        Private _groupHost As System.Windows.Forms.Panel
        Private _groupedActive As System.Boolean
        Private _groupingRenderFailed As System.Boolean
        Private _groupsButton As System.Windows.Forms.Button
        Private _groupsToolTip As System.Windows.Forms.ToolTip
        Private _updatingGroupingWidth As System.Boolean

        Public Sub ConfigureGrouping(store As SelectorGroupingStore)
            If store Is Nothing OrElse _groupingStore IsNot Nothing Then Return
            Try
                Dim items As New System.Collections.Generic.List(Of SelectorGroupItem)()
                For Each display As System.String In allDisplayItems
                    Dim model As ModelConfig = displayToModel(display)
                    items.Add(New SelectorGroupItem() With {.Id = model.ToolName, .Label = display})
                Next
                ' Validate IDs before changing any UI; do not guess IDs from descriptions.
                Dim validation As New SelectorGroupSelection(items, Nothing)
                store.Reload()
                _groupItems = items
                _groupingStore = store
                AddGroupEditorButton()
                ApplyGroupingDocument()
            Catch ex As System.Exception
                SelectorGroupingStore.Log("selector-flat-fallback", store.SourceFile, ex)
                _groupingRenderFailed = True
                RestoreFlatView()
                UpdateGroupStatus()
            End Try
        End Sub

        Private Sub AddGroupEditorButton()
            ' Only the opt-in selectors use this layout. An implicit AutoSize column
            ' can outgrow the client area when the tree/button captions replace the
            ' flat list. Keep every row inside the actual available width instead.
            outer.ColumnStyles.Clear()
            outer.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
            Dim bar As New System.Windows.Forms.TableLayoutPanel() With {
                .AutoSize = True,
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 1,
                .Margin = New System.Windows.Forms.Padding(0)
            }
            bar.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
            bar.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize))
            outer.Controls.Remove(txtFilter)
            txtFilter.Dock = System.Windows.Forms.DockStyle.Fill
            bar.Controls.Add(txtFilter, 0, 0)
            _groupsButton = New System.Windows.Forms.Button() With {
                .Text = "Groups…",
                .AutoSize = True,
                .Margin = New System.Windows.Forms.Padding(8, 0, 0, 0),
                .Padding = New System.Windows.Forms.Padding(6, 0, 6, 0)
            }
            AddHandler _groupsButton.Click, AddressOf EditGroups
            bar.Controls.Add(_groupsButton, 1, 0)
            outer.Controls.Add(bar, 0, 1)
            _groupsToolTip = New System.Windows.Forms.ToolTip()
            AddHandler Me.Disposed,
                Sub(sender As System.Object, args As System.EventArgs)
                    _groupsToolTip.Dispose()
                End Sub
            UpdateGroupStatus()
        End Sub

        Private Sub UpdateGroupStatus()
            If _groupsButton Is Nothing OrElse _groupingStore Is Nothing Then Return
            _groupsButton.Text = If(_groupingStore.Warning.Length > 0 OrElse _groupingRenderFailed, "Groups (!)…", "Groups…")
            Dim hint As System.String = "Edit presentation groups. Tool selections are unaffected."
            If _groupingStore.SourceFile.Length > 0 Then hint &= System.Environment.NewLine & _groupingStore.SourceFile
            If _groupingStore.Warning.Length > 0 Then hint &= System.Environment.NewLine & _groupingStore.Warning
            If _groupingRenderFailed Then hint &= System.Environment.NewLine & "The grouped view could not be displayed. The flat list is being used without changing your selections."
            _groupsToolTip.SetToolTip(_groupsButton, hint)
            ' ApplyGroupingDocument calls this AFTER switching view and captions;
            ' the initial Shown event cannot size a later flat-to-tree transition.
            UpdatePreferredDialogWidth()
        End Sub


        Private Sub UpdateGroupingDialogWidth()
            If _updatingGroupingWidth OrElse Me.IsDisposed OrElse Me.Disposing Then Return
            If Me.outer Is Nothing OrElse Me.pnlButtons Is Nothing OrElse _groupsButton Is Nothing Then Return
            _updatingGroupingWidth = True
            Me.SuspendLayout()
            Try
                Dim scale As System.Double = Me.DeviceDpi / 96.0R
                ' Measure the controls, not the already constrained FlowLayoutPanel.
                ' Reserve the longer toggle caption so checking/filtering cannot clip
                ' buttons or repeatedly grow/shrink the window.
                Dim buttonRowWidth As System.Int32 = Me.pnlButtons.Padding.Horizontal + Me.pnlButtons.Margin.Horizontal
                For Each control As System.Windows.Forms.Control In Me.pnlButtons.Controls
                    Dim width As System.Int32 = control.GetPreferredSize(System.Drawing.Size.Empty).Width
                    If control Is Me.btnToggleAll Then
                        Dim currentTextWidth As System.Int32 = System.Windows.Forms.TextRenderer.MeasureText(control.Text, control.Font).Width
                        Dim longestTextWidth As System.Int32 = System.Windows.Forms.TextRenderer.MeasureText("Unselect all matches", control.Font).Width
                        width += System.Math.Max(0, longestTextWidth - currentTextWidth)
                    End If
                    buttonRowWidth += width + control.Margin.Horizontal
                Next

                ' Keep a useful filter field next to the complete Groups button.
                Dim filterRowWidth As System.Int32 = CInt(System.Math.Ceiling(160.0R * scale)) + Me.txtFilter.Margin.Horizontal
                filterRowWidth += _groupsButton.GetPreferredSize(System.Drawing.Size.Empty).Width + _groupsButton.Margin.Horizontal
                If _groupsButton.Parent IsNot Nothing Then
                    filterRowWidth += _groupsButton.Parent.Padding.Horizontal + _groupsButton.Parent.Margin.Horizontal
                End If
                Dim widestItemWidth As System.Int32 = 0
                For Each itemText As System.String In allDisplayItems
                    widestItemWidth = System.Math.Max(widestItemWidth, System.Windows.Forms.TextRenderer.MeasureText(itemText, Me.chkList.Font).Width)
                Next
                Dim listChromeWidth As System.Int32 = CInt(System.Math.Ceiling(If(_groupedActive, 64.0R, 40.0R) * scale))
                Dim contentWidth As System.Int32 = System.Math.Max(System.Math.Max(buttonRowWidth, filterRowWidth), widestItemWidth + listChromeWidth)
                Dim nonClientWidth As System.Int32 = Me.Width - Me.ClientSize.Width
                Dim marginWidth As System.Int32 = CInt(System.Math.Ceiling(24.0R * scale))
                Dim workingArea As System.Drawing.Rectangle = System.Windows.Forms.Screen.FromControl(Me).WorkingArea
                Dim maxWidth As System.Int32 = System.Math.Max(1, workingArea.Width - CInt(System.Math.Ceiling(80.0R * scale)))
                Dim preferredWidth As System.Int32 = System.Math.Min(maxWidth, System.Math.Max(640, contentWidth + Me.outer.Padding.Horizontal + nonClientWidth + marginWidth))

                ' Set the new minimum first: the previous mode's minimum must not
                ' prevent a screen-constrained width. Keep a manual enlargement.
                Me.MinimumSize = New System.Drawing.Size(preferredWidth, Me.MinimumSize.Height)
                If Me.WindowState = System.Windows.Forms.FormWindowState.Normal Then
                    Me.Width = System.Math.Min(maxWidth, System.Math.Max(Me.Width, preferredWidth))
                    If Me.Visible Then
                        Me.Left = System.Math.Max(workingArea.Left, System.Math.Min(Me.Left, workingArea.Right - Me.Width))
                    End If
                End If
                ' If a single button row cannot fit this screen, let the existing
                ' auto-height row wrap instead of placing actions outside the form.
                Dim availableWidth As System.Int32 = System.Math.Max(1, Me.ClientSize.Width - Me.outer.Padding.Horizontal)
                Me.pnlButtons.WrapContents = buttonRowWidth > availableWidth
                UpdateInstructionLabelLayout()
            Finally
                Me.ResumeLayout(True)
                _updatingGroupingWidth = False
            End Try
            Me.outer.PerformLayout()
        End Sub

        Private Function CurrentSelectedIds() As System.Collections.Generic.List(Of System.String)
            If _groupedActive Then Return _groupedView.State.SelectedIds()
            Dim ids As New System.Collections.Generic.List(Of System.String)()
            For Each label As System.String In allDisplayItems
                If selectedLabels.Contains(label) Then ids.Add(displayToModel(label).ToolName)
            Next
            Return ids
        End Function

        Private Sub SyncGroupedSelectionToLabels()
            If Not _groupedActive OrElse _groupedView Is Nothing Then Return
            ' Preserve the existing HashSet lifecycle/return order and original model
            ' references, while the tree itself uses stable IDs. Do not clear/reseed it.
            For Each label As System.String In allDisplayItems
                Dim model As ModelConfig = displayToModel(label)
                If _groupedView.State.IsSelected(model.ToolName) Then
                    selectedLabels.Add(label)
                Else
                    selectedLabels.Remove(label)
                End If
            Next
        End Sub

        Private Sub ApplyGroupingDocument()
            Dim selected As System.Collections.Generic.List(Of System.String) = CurrentSelectedIds()
            _groupingRenderFailed = False
            If _groupingStore.Document.Groups.Count = 0 OrElse Not HasApplicableGroups() Then
                RestoreFlatView()
                UpdateGroupStatus()
                Return
            End If
            If _groupHost Is Nothing Then
                _groupHost = New System.Windows.Forms.Panel() With {.Dock = System.Windows.Forms.DockStyle.Fill, .Margin = chkList.Margin}
                outer.Controls.Remove(chkList)
                _groupHost.Controls.Add(chkList)
                outer.Controls.Add(_groupHost, 0, 2)
            End If
            If _groupedView Is Nothing Then
                Dim state As New SelectorGroupSelection(_groupItems, selected)
                state.Layout = _groupingStore.Document.Copy()
                state.Filter = txtFilter.Text
                _groupedView = New SelectorGroupedList(state)
                AddHandler _groupedView.SelectionChanged,
                    Sub(sender As System.Object, args As System.EventArgs)
                        SyncGroupedSelectionToLabels()
                        UpdateToggleAllButtonText()
                    End Sub
                AddHandler _groupedView.PresentationFailed,
                    Sub(sender As System.Object, args As System.EventArgs)
                        _groupingRenderFailed = True
                        RestoreFlatView()
                        UpdateGroupStatus()
                    End Sub
                _groupHost.Controls.Add(_groupedView)
            Else
                Dim checkedIds As New System.Collections.Generic.HashSet(Of System.String)(selected, System.StringComparer.OrdinalIgnoreCase)
                For Each item As SelectorGroupItem In _groupItems
                    _groupedView.State.SetSelected(item.Id, checkedIds.Contains(item.Id))
                Next
                _groupedView.State.Filter = txtFilter.Text
                _groupedView.ApplyLayout(_groupingStore.Document)
            End If
            If _groupingRenderFailed Then Return
            _groupedActive = True
            chkList.Visible = False
            _groupedView.Visible = True
            _groupedView.BringToFront()
            UpdateToggleAllButtonText()
            UpdateGroupStatus()
        End Sub

        Private Function HasApplicableGroups() As System.Boolean
            For Each item As SelectorGroupItem In _groupItems
                If _groupingStore.Document.GroupFor(item.Id).Length > 0 Then Return True
            Next
            Return False
        End Function

        Private Sub RestoreFlatView()
            SyncGroupedSelectionToLabels()
            _groupedActive = False
            If _groupedView IsNot Nothing Then _groupedView.Visible = False
            chkList.Visible = True
            chkList.BringToFront()
            OnFilterChanged(Me, System.EventArgs.Empty)
        End Sub

        Private Sub EditGroups(sender As System.Object, args As System.EventArgs)
            Try
                ' Opening the editor is an explicit refresh point, not a per-keystroke read.
                _groupingStore.Reload()
                Using editor As New SelectorGroupsEditorForm(_groupingStore, _groupItems)
                    Using SharedMethods.PushDialogOwner(Me)
                        Dim owner As System.Windows.Forms.IWin32Window = SharedMethods.IfOwnerOnCurrentThread(Me)
                        If If(owner IsNot Nothing, editor.ShowDialog(owner), editor.ShowDialog()) = System.Windows.Forms.DialogResult.OK Then
                            ApplyGroupingDocument()
                        End If
                    End Using
                End Using
                UpdateGroupStatus()
            Catch ex As System.Exception
                SelectorGroupingStore.Log("editor-open-failed", _groupingStore.SourceFile, ex)
                Using SharedMethods.PushDialogOwner(Me)
                    SharedMethods.ShowCustomMessageBox("The group settings could not be opened. Your tool selections have not been changed.", header:=SharedMethods.AN & " - Selector groups")
                End Using
            End Try
        End Sub
    End Class
End Namespace
