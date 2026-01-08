' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.SelectionFormSmall.vb
' Purpose:
'   Provides a small modal WinForms selection dialog and a helper API
'   (`SharedMethods.SelectValue`) to let the caller choose an integer value from
'   a list of labeled options.
'
' How it works:
'   - `SelectionItem` is a value/display pair; `ToString()` returns `DisplayText`
'     so a `ListBox` can render the item directly.
'   - `SelectionFormSmall` renders a prompt label and a ListBox showing up to five
'     rows initially, then sizes the form based on the computed list height.
'   - Keyboard/mouse handling:
'       - Enter or double-click accepts the current selection (DialogResult.OK).
'       - Escape cancels and resets the result to 0 (DialogResult.Cancel).
'   - `SelectValue` validates input, shows the dialog, and returns `Result`.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary
    Partial Public Class SharedMethods

        ''' <summary>
        ''' Represents an item shown in the selection list, pairing display text with an integer value.
        ''' </summary>
        Public Structure SelectionItem
            ''' <summary>Text shown in the UI list.</summary>
            Public ReadOnly DisplayText As String

            ''' <summary>Integer value returned to the caller when selected.</summary>
            Public ReadOnly Value As Integer

            ''' <summary>
            ''' Initializes a new selection item.
            ''' </summary>
            ''' <param name="text">Display text shown to the user.</param>
            ''' <param name="value">Integer value associated with the item.</param>
            Public Sub New(text As String, value As Integer)
                Me.DisplayText = text
                Me.Value = value
            End Sub

            ''' <summary>
            ''' Returns the display text so the ListBox renders the item as intended.
            ''' </summary>
            Public Overrides Function ToString() As String
                Return DisplayText
            End Function
        End Structure

        ''' <summary>
        ''' Modal selection dialog used by <see cref="SelectValue"/> to capture an integer choice from a list.
        ''' </summary>
        Friend NotInheritable Class SelectionFormSmall
            Inherits System.Windows.Forms.Form

            ''' <summary>List box containing <see cref="SelectionItem"/> entries.</summary>
            Private ReadOnly _lst As System.Windows.Forms.ListBox

            ''' <summary>Prompt label displayed above the list.</summary>
            Private ReadOnly _lbl As System.Windows.Forms.Label

            ''' <summary>Selected result value. 0 indicates cancellation/no selection in this implementation.</summary>
            Private _result As Integer = 0

            ''' <summary>
            ''' Initializes a new instance of the selection form.
            ''' </summary>
            ''' <param name="items">Items to display in the ListBox.</param>
            ''' <param name="defaultValue">Default selected value (matched against <see cref="SelectionItem.Value"/>).</param>
            ''' <param name="promptText">Prompt text shown above the list.</param>
            ''' <param name="headerText">Optional window title. If missing/blank, <c>AN</c> is used.</param>
            Friend Sub New(items As IReadOnlyList(Of SelectionItem),
                           defaultValue As Integer,
                           promptText As String,
                           Optional headerText As String = Nothing)

                Const minWidth As Integer = 400
                Const maxWidth As Integer = 900
                Const sidePadding As Integer = 10
                Const bottomPadding As Integer = 24

                Const minVisibleRows As Integer = 5
                Const maxVisibleRows As Integer = 18

                Me.SuspendLayout()

                Dim stdFont As New System.Drawing.Font(
                    "Segoe UI",
                    9.0F,
                    System.Drawing.FontStyle.Regular,
                    System.Drawing.GraphicsUnit.Point
                )
                Me.Font = stdFont
                Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font

                If String.IsNullOrWhiteSpace(headerText) Then headerText = AN
                Me.Text = headerText

                Dim bmp As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                Me.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

                ' Important: don't reference baseWidth (it no longer exists).
                ' We compute ClientSize after measuring content.
                Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
                Me.KeyPreview = True

                _lbl = New System.Windows.Forms.Label With {
                    .AutoSize = True,
                    .Text = promptText,
                    .Location = New System.Drawing.Point(sidePadding, sidePadding),
                    .Anchor = System.Windows.Forms.AnchorStyles.Top Or
                              System.Windows.Forms.AnchorStyles.Left Or
                              System.Windows.Forms.AnchorStyles.Right
                }
                Controls.Add(_lbl)

                ' Important: do NOT anchor Bottom yet (avoid early stretch)
                _lst = New System.Windows.Forms.ListBox With {
                    .IntegralHeight = False,
                    .SelectionMode = System.Windows.Forms.SelectionMode.One,
                    .Anchor = System.Windows.Forms.AnchorStyles.Top Or
                              System.Windows.Forms.AnchorStyles.Left Or
                              System.Windows.Forms.AnchorStyles.Right
                }

                ' Fill items first (so width can be measured)
                For Each it In items
                    _lst.Items.Add(it)
                Next

                ' --- Dynamic width based on widest item / prompt ---
                Dim measuredContentWidth As Integer = minWidth

                Using g As System.Drawing.Graphics = Me.CreateGraphics()
                    Dim maxItemTextWidth As Integer = 0

                    For Each it As SelectionItem In items
                        Dim w As Integer = System.Windows.Forms.TextRenderer.MeasureText(
                            g,
                            it.DisplayText,
                            stdFont,
                            New System.Drawing.Size(Integer.MaxValue, Integer.MaxValue),
                            System.Windows.Forms.TextFormatFlags.SingleLine
                        ).Width
                        If w > maxItemTextWidth Then maxItemTextWidth = w
                    Next

                    Dim promptWidth As Integer = System.Windows.Forms.TextRenderer.MeasureText(
                        g,
                        promptText,
                        stdFont,
                        New System.Drawing.Size(Integer.MaxValue, Integer.MaxValue),
                        System.Windows.Forms.TextFormatFlags.SingleLine
                    ).Width

                    ' Add ListBox borders/scrollbar + padding
                    Dim listChrome As Integer = 35 ' scrollbar + borders (approx)
                    Dim neededClientWidth As Integer =
                        Math.Max(promptWidth, maxItemTextWidth + listChrome) + (2 * sidePadding)

                    measuredContentWidth = Math.Max(minWidth, Math.Min(maxWidth, neededClientWidth))
                End Using

                ' --- Dynamic height based on item count (more than 5 rows if long list) ---
                Dim visibleRows As Integer = Math.Min(maxVisibleRows, Math.Max(minVisibleRows, items.Count))
                _lst.ItemHeight = CInt(stdFont.GetHeight())
                Dim desiredListHeight As Integer = _lst.ItemHeight * visibleRows + 9

                ' Place controls now that width/height is known
                _lbl.MaximumSize = New System.Drawing.Size(measuredContentWidth - (2 * sidePadding), 0) ' allow wrapping if needed

                _lst.Location = New System.Drawing.Point(sidePadding, _lbl.Bottom + 10)
                _lst.Width = measuredContentWidth - (2 * sidePadding)
                _lst.Height = desiredListHeight
                Controls.Add(_lst)

                Dim requiredHeight As Integer = _lst.Top + desiredListHeight + bottomPadding
                Me.ClientSize = New System.Drawing.Size(measuredContentWidth, requiredHeight)

                ' After ClientSize is finalized, enable Bottom anchoring
                _lst.Anchor = System.Windows.Forms.AnchorStyles.Top Or
                              System.Windows.Forms.AnchorStyles.Left Or
                              System.Windows.Forms.AnchorStyles.Right Or
                              System.Windows.Forms.AnchorStyles.Bottom

                ' Ensure width matches final client width
                _lst.Width = Me.ClientSize.Width - (2 * sidePadding)

                ' Keep minimum width at least 400, but allow resizing smaller than the initial computed width if desired.
                Me.MinimumSize = New System.Drawing.Size(minWidth, Me.Size.Height)

                AddHandler _lst.KeyDown,
                    Sub(s, e)
                        If e.KeyCode = System.Windows.Forms.Keys.Enter Then AcceptCurrentSelection()
                    End Sub

                AddHandler _lst.DoubleClick, Sub() AcceptCurrentSelection()

                AddHandler Me.KeyDown,
                    Sub(sender, e)
                        If e.KeyCode = System.Windows.Forms.Keys.Escape Then
                            _result = 0
                            Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
                            Close()
                        End If
                    End Sub

                AddHandler Me.FormClosing,
                    Sub(s, e)
                        If Me.DialogResult <> System.Windows.Forms.DialogResult.OK Then _result = 0
                    End Sub

                ' Keep padding and width on resize
                AddHandler Me.Resize,
                    Sub()
                        _lbl.Width = Me.ClientSize.Width - (2 * sidePadding)
                        _lst.Width = Me.ClientSize.Width - (2 * sidePadding)
                        Dim newHeight = Me.ClientSize.Height - _lst.Top - bottomPadding
                        If newHeight > 40 Then _lst.Height = newHeight
                    End Sub

                ' Preselect default
                Dim defIdx As Integer = items.ToList().FindIndex(Function(it) it.Value = defaultValue)
                If defIdx >= 0 Then
                    _lst.SelectedIndex = defIdx
                    _result = items(defIdx).Value
                End If

                Me.ResumeLayout(False)
                Me.PerformLayout()
                _lst.Focus()
            End Sub

            ''' <summary>
            ''' Accepts the current ListBox selection and closes the form with <see cref="DialogResult.OK"/>.
            ''' </summary>
            Private Sub AcceptCurrentSelection()
                If _lst.SelectedIndex >= 0 Then
                    Dim item As SelectionItem = DirectCast(_lst.SelectedItem, SelectionItem)
                    _result = item.Value
                    Me.DialogResult = System.Windows.Forms.DialogResult.OK
                    Close()
                End If
            End Sub

            ''' <summary>
            ''' Gets the selected integer result (0 when canceled/no selection in this implementation).
            ''' </summary>
            Friend ReadOnly Property Result As Integer
                Get
                    Return _result
                End Get
            End Property
        End Class

        ''' <summary>
        ''' Shows a modal selection dialog and returns the chosen integer value.
        ''' </summary>
        ''' <param name="items">Items to offer in the selection list.</param>
        ''' <param name="defaultValue">Default selected value.</param>
        ''' <param name="prompt">Prompt shown above the selection list.</param>
        ''' <param name="header">Optional window title. If missing/blank, <c>AN</c> is used.</param>
        ''' <returns>The selected value, or 0 when canceled or when <paramref name="items"/> is Nothing.</returns>
        Public Shared Function SelectValue(items As IEnumerable(Of SelectionItem),
                                           defaultValue As Integer,
                                           Optional prompt As String = "Please choose …",
                                           Optional header As String = Nothing) As Integer

            If items Is Nothing Then
                System.Windows.Forms.MessageBox.Show("SelectValue Error: Items collection must not be null.")
                Return 0
            End If

            Using frm As New SelectionFormSmall(items.ToList(), defaultValue, prompt, header)
                frm.ShowDialog()
                Return frm.Result
            End Using
        End Function

    End Class
End Namespace