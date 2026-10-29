' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' One-level presentation of a host-supplied catalog. Headers never change selection.
Option Strict On
Option Explicit On

Namespace SharedLibrary
    Public NotInheritable Class SelectorGroupedList
        Inherits System.Windows.Forms.UserControl

        Private ReadOnly _tree As New GroupTree()
        Private ReadOnly _expanded As New System.Collections.Generic.Dictionary(Of System.String, System.Boolean)(System.StringComparer.OrdinalIgnoreCase)
        Private _building As System.Boolean
        Private _state As SelectorGroupSelection
        Public Event SelectionChanged As System.EventHandler
        Public Event PresentationFailed As System.EventHandler

        Private NotInheritable Class GroupTree
            Inherits System.Windows.Forms.TreeView
            Private _ignoreDoubleClickUp As System.Boolean

            Protected Overrides Sub WndProc(ByRef message As System.Windows.Forms.Message)
                ' A native checkbox double-click can bypass the normal AfterCheck path.
                ' Treat it as one click, including its matching mouse-up; never let the
                ' displayed check state drift away from the stable-ID state model.
                If message.Msg = &H203 Then
                    Dim hit As System.Windows.Forms.TreeViewHitTestInfo = Me.HitTest(Me.PointToClient(System.Windows.Forms.Cursor.Position))
                    If hit.Node IsNot Nothing Then
                        _ignoreDoubleClickUp = True
                        message.Result = System.IntPtr.Zero
                        Return
                    End If
                ElseIf message.Msg = &H202 AndAlso _ignoreDoubleClickUp Then
                    _ignoreDoubleClickUp = False
                    message.Result = System.IntPtr.Zero
                    Return
                ElseIf message.Msg = &H201 Then
                    _ignoreDoubleClickUp = False
                End If
                MyBase.WndProc(message)
            End Sub
        End Class

        Private NotInheritable Class Header
            Public Property Name As System.String
            Public Property Label As System.String
        End Class

        ' WinForms exposes CheckBoxes for the whole tree, whereas native TreeView permits
        ' an empty state image for headers. BeforeCheck also rejects keyboard/programmatic
        ' checks of headers; removing the glyph alone would not be a sufficient guard.
        <System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential, CharSet:=System.Runtime.InteropServices.CharSet.Unicode)>
        Private Structure TreeItem
            Public Mask As System.UInt32
            Public ItemHandle As System.IntPtr
            Public State As System.UInt32
            Public StateMask As System.UInt32
            Public TextPointer As System.IntPtr
            Public TextLength As System.Int32
            Public Image As System.Int32
            Public SelectedImage As System.Int32
            Public Children As System.Int32
            Public Parameter As System.IntPtr
        End Structure

        <System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint:="SendMessageW", CharSet:=System.Runtime.InteropServices.CharSet.Unicode)>
        Private Shared Function SendTreeMessage(window As System.IntPtr, message As System.UInt32, parameter As System.IntPtr, ByRef item As TreeItem) As System.IntPtr
        End Function

        Public Sub New(state As SelectorGroupSelection)
            _state = state
            Me.Dock = System.Windows.Forms.DockStyle.Fill
            _tree.Dock = System.Windows.Forms.DockStyle.Fill
            _tree.CheckBoxes = True
            _tree.HideSelection = False
            _tree.FullRowSelect = False
            _tree.ShowNodeToolTips = True
            _tree.AccessibleName = "Grouped entries"
            Me.Controls.Add(_tree)
            AddHandler _tree.BeforeCheck, AddressOf BeforeCheck
            AddHandler _tree.AfterCheck, AddressOf AfterCheck
            AddHandler _tree.NodeMouseClick, AddressOf NodeClicked
            AddHandler _tree.AfterExpand, AddressOf ExpansionChanged
            AddHandler _tree.AfterCollapse, AddressOf ExpansionChanged
            AddHandler _tree.KeyDown, AddressOf TreeKeyDown
            AddHandler _tree.HandleCreated,
                Sub(sender As System.Object, args As System.EventArgs)
                    If Not _tree.IsDisposed Then
                        _tree.BeginInvoke(New System.Action(AddressOf HideHeaderChecks))
                    End If
                End Sub
            Rebuild()
        End Sub

        Public ReadOnly Property State As SelectorGroupSelection
            Get
                Return _state
            End Get
        End Property

        Public Sub ApplyLayout(document As SelectorGroupDocument)
            _state.Layout = document.Copy()
            Rebuild()
        End Sub

        Public Sub ApplyFilter(filter As System.String)
            _state.Filter = If(filter, "")
            Rebuild()
        End Sub

        Public Sub ToggleMatching()
            _state.ToggleMatching()
            Rebuild()
            RaiseEvent SelectionChanged(Me, System.EventArgs.Empty)
        End Sub

        Private Sub Rebuild()
            Dim selectedItem As SelectorGroupItem = Nothing
            If _tree.SelectedNode IsNot Nothing Then selectedItem = TryCast(_tree.SelectedNode.Tag, SelectorGroupItem)
            _building = True
            _tree.BeginUpdate()
            Try
                _tree.Nodes.Clear()
                Dim nodes As New System.Collections.Generic.Dictionary(Of System.String, System.Windows.Forms.TreeNode)(System.StringComparer.OrdinalIgnoreCase)
                For Each definition As SelectorGroupDefinition In _state.Layout.Groups
                    If Not _expanded.ContainsKey(definition.Name) Then _expanded(definition.Name) = definition.Expanded
                    nodes(definition.Name) = New System.Windows.Forms.TreeNode() With {
                        .Tag = New Header() With {.Name = definition.Name, .Label = definition.Name}
                    }
                Next
                nodes("") = New System.Windows.Forms.TreeNode() With {
                    .Tag = New Header() With {.Name = "", .Label = "Other (unassigned)"}
                }
                If Not _expanded.ContainsKey("") Then _expanded("") = True
                Dim restoreSelection As System.Windows.Forms.TreeNode = Nothing
                For Each item As SelectorGroupItem In _state.Items
                    If Not _state.Matches(item) Then Continue For
                    Dim leaf As New System.Windows.Forms.TreeNode(item.Label) With {
                        .Tag = item,
                        .Checked = _state.IsSelected(item.Id),
                        .ToolTipText = item.Label & System.Environment.NewLine & item.Id
                    }
                    nodes(_state.Layout.GroupFor(item.Id)).Nodes.Add(leaf)
                    If selectedItem IsNot Nothing AndAlso System.String.Equals(item.Id, selectedItem.Id, System.StringComparison.OrdinalIgnoreCase) Then restoreSelection = leaf
                Next
                For Each definition As SelectorGroupDefinition In _state.Layout.Groups
                    AddHeader(nodes(definition.Name))
                Next
                AddHeader(nodes(""))
                ' Do not EnsureVisible here: that would expand a deliberately collapsed group.
                If restoreSelection IsNot Nothing AndAlso restoreSelection.Parent.IsExpanded Then _tree.SelectedNode = restoreSelection
                RefreshCounts()
            Finally
                _tree.EndUpdate()
                _building = False
            End Try
            HideHeaderChecks()
        End Sub

        Private Sub AddHeader(node As System.Windows.Forms.TreeNode)
            If node.Nodes.Count = 0 Then Return
            Dim headerValue As Header = DirectCast(node.Tag, Header)
            _tree.Nodes.Add(node)
            ' A search temporarily exposes matches; it never overwrites the saved expansion.
            If Not System.String.IsNullOrWhiteSpace(_state.Filter) OrElse _expanded(headerValue.Name) Then node.Expand()
        End Sub

        Private Sub RefreshCounts()
            For Each node As System.Windows.Forms.TreeNode In _tree.Nodes
                Dim headerValue As Header = DirectCast(node.Tag, Header)
                Dim total As System.Int32 = 0
                Dim selected As System.Int32 = 0
                For Each item As SelectorGroupItem In _state.Items
                    If System.String.Equals(_state.Layout.GroupFor(item.Id), headerValue.Name, System.StringComparison.OrdinalIgnoreCase) Then
                        total += 1
                        If _state.IsSelected(item.Id) Then selected += 1
                    End If
                Next
                node.Text = headerValue.Label & " (" & selected.ToString() & " / " & total.ToString() & " selected)"
                node.ToolTipText = "Expand or collapse only. This does not select or deselect entries."
            Next
        End Sub

        Private Sub HideHeaderChecks()
            If _tree.IsDisposed OrElse Not _tree.IsHandleCreated Then Return
            Try
                For Each node As System.Windows.Forms.TreeNode In _tree.Nodes
                    Dim item As New TreeItem() With {
                        .Mask = &H8UI,
                        .ItemHandle = node.Handle,
                        .State = 0UI,
                        .StateMask = &HF000UI
                    }
                    If SendTreeMessage(_tree.Handle, &H113FUI, System.IntPtr.Zero, item) = System.IntPtr.Zero Then
                        Throw New System.InvalidOperationException("The tree header could not be rendered without a selection checkbox.")
                    End If
                Next
            Catch ex As System.Exception
                SelectorGroupingStore.Log("group-rendering-flat-fallback", "", ex)
                RaiseEvent PresentationFailed(Me, System.EventArgs.Empty)
            End Try
        End Sub

        Private Sub BeforeCheck(sender As System.Object, args As System.Windows.Forms.TreeViewCancelEventArgs)
            If Not TypeOf args.Node.Tag Is SelectorGroupItem Then args.Cancel = True
        End Sub

        Private Sub AfterCheck(sender As System.Object, args As System.Windows.Forms.TreeViewEventArgs)
            If _building Then Return
            Dim item As SelectorGroupItem = TryCast(args.Node.Tag, SelectorGroupItem)
            If item Is Nothing Then Return
            _state.SetSelected(item.Id, args.Node.Checked)
            RefreshCounts()
            RaiseEvent SelectionChanged(Me, System.EventArgs.Empty)
        End Sub

        Private Sub NodeClicked(sender As System.Object, args As System.Windows.Forms.TreeNodeMouseClickEventArgs)
            If args.Button <> System.Windows.Forms.MouseButtons.Left OrElse args.Clicks <> 1 Then Return
            Dim hit As System.Windows.Forms.TreeViewHitTestInfo = _tree.HitTest(args.Location)
            ' None is the native NOWHERE flag (1), NOT a zero bit mask. Compare
            ' with the requested Label flag so native +/- and checkbox clicks are
            ' not handled a second time here (which would undo the native toggle).
            If hit.Node IsNot args.Node OrElse
               (hit.Location And System.Windows.Forms.TreeViewHitTestLocations.Label) <> System.Windows.Forms.TreeViewHitTestLocations.Label Then Return
            If TypeOf args.Node.Tag Is Header Then
                args.Node.Toggle()
            Else
                ' Native state-image clicks already toggle. Only label clicks need this.
                args.Node.Checked = Not args.Node.Checked
            End If
        End Sub

        Private Sub ExpansionChanged(sender As System.Object, args As System.Windows.Forms.TreeViewEventArgs)
            If _building OrElse Not System.String.IsNullOrWhiteSpace(_state.Filter) Then Return
            Dim headerValue As Header = TryCast(args.Node.Tag, Header)
            If headerValue IsNot Nothing Then _expanded(headerValue.Name) = args.Node.IsExpanded
        End Sub

        Private Sub TreeKeyDown(sender As System.Object, args As System.Windows.Forms.KeyEventArgs)
            If args.KeyCode = System.Windows.Forms.Keys.Space AndAlso _tree.SelectedNode IsNot Nothing AndAlso TypeOf _tree.SelectedNode.Tag Is Header Then
                _tree.SelectedNode.Toggle()
                args.Handled = True
                args.SuppressKeyPress = True
            End If
        End Sub
    End Class
End Namespace
