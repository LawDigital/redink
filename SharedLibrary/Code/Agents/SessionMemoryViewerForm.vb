' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SessionMemoryViewerForm.vb
' Purpose: Lets the user inspect and manage the persistent session memory used by
'          the agent layer. Displays memory entries with summaries, allows opening
'          values in a text editor, deleting entries, and clearing all memory.
'
' UI Features:
'  - Lists "[KEY] SUMMARY" entries with optional filter.
'  - "Open" exports the selected value to a temp text/JSON file and opens it.
'  - "Delete" removes the selected entry.
'  - "Clear All" wipes the entire memory (with confirmation).
'  - "Refresh" re-reads from storage.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agents

    Public Class SessionMemoryViewerForm
        Inherits Form

        Protected Overrides Sub OnShown(e As System.EventArgs)
            MyBase.OnShown(e)
            Me.TopMost = True
            Global.SharedLibrary.SharedLibrary.SharedMethods.ForceDialogToForeground(Me)
            Global.SharedLibrary.SharedLibrary.SharedMethods.AttachForeignForegroundWatchdog(Me)
        End Sub


        Private Const EM_SETCUEBANNER As Integer = &H1501

        <System.Runtime.InteropServices.DllImport("user32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Unicode)>
        Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As String) As IntPtr
        End Function

        Private lblTitle As Label
        Private txtFilter As TextBox
        Private lstItems As ListBox
        Private btnOpen As Button
        Private btnDelete As Button
        Private btnClearAll As Button
        Private btnRefresh As Button
        Private btnClose As Button
        Private pnlButtons As FlowLayoutPanel
        Private outer As TableLayoutPanel
        Private _dialogOwnerScope As System.IDisposable = Nothing

        Private entries As New List(Of SessionMemoryEntry)()

        Public Sub New(Optional title As String = Nothing)
            InitializeComponent(title)
            ReloadEntries()
        End Sub

        Private Sub InitializeComponent(title As String)
            Me.SuspendLayout()

            Me.Text = If(String.IsNullOrWhiteSpace(title), SharedLibrary.SharedMethods.AN & " - Session Memory", title)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.MinimizeBox = False
            Me.MaximizeBox = False
            Me.ShowInTaskbar = False
            Me.FormBorderStyle = FormBorderStyle.Sizable
            Me.AutoScaleMode = AutoScaleMode.Dpi
            Me.AutoScaleDimensions = New SizeF(96.0F, 96.0F)
            Me.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            Me.Width = 860
            Me.Height = 620
            Me.MinimumSize = New Size(680, 440)

            Try
                Dim bmp As New Bitmap(SharedLibrary.SharedMethods.GetLogoBitmap(SharedLibrary.SharedMethods.LogoType.Standard))
                Me.Icon = Icon.FromHandle(bmp.GetHicon())
            Catch
            End Try

            Me.outer = New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 4,
                .Padding = New Padding(12)
            }
            Me.outer.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
            Me.outer.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            Me.outer.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            Me.outer.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
            Me.outer.RowStyles.Add(New RowStyle(SizeType.AutoSize))

            Me.lblTitle = New Label() With {
                .AutoSize = True,
                .Text = "Entries stored in session memory:",
                .Anchor = AnchorStyles.Left,
                .Margin = New Padding(0, 0, 0, 8)
            }

            Me.txtFilter = New TextBox() With {
                .Dock = DockStyle.Fill,
                .Margin = New Padding(0, 0, 0, 8)
            }
            AddHandler Me.txtFilter.TextChanged, AddressOf OnFilterChanged

            Me.lstItems = New ListBox() With {
                .Dock = DockStyle.Fill,
                .IntegralHeight = False,
                .SelectionMode = SelectionMode.One,
                .HorizontalScrollbar = True
            }
            Me.lstItems.DisplayMember = NameOf(ItemWrap.Display)
            AddHandler Me.lstItems.DoubleClick, AddressOf OnOpenClicked
            AddHandler Me.lstItems.SelectedIndexChanged, AddressOf OnSelectionChanged

            Me.btnOpen = New Button() With {.Text = "Open"}
            ConfigureStandardButton(Me.btnOpen)
            AddHandler Me.btnOpen.Click, AddressOf OnOpenClicked

            Me.btnDelete = New Button() With {.Text = "Delete"}
            ConfigureStandardButton(Me.btnDelete)
            AddHandler Me.btnDelete.Click, AddressOf OnDeleteClicked

            Me.btnClearAll = New Button() With {.Text = "Clear All"}
            ConfigureStandardButton(Me.btnClearAll)
            AddHandler Me.btnClearAll.Click, AddressOf OnClearAllClicked

            Me.btnRefresh = New Button() With {.Text = "Refresh"}
            ConfigureStandardButton(Me.btnRefresh)
            AddHandler Me.btnRefresh.Click, Sub(s, e) ReloadEntries()

            Me.btnClose = New Button() With {.Text = "Close", .DialogResult = DialogResult.OK}
            ConfigureStandardButton(Me.btnClose)

            Me.pnlButtons = New FlowLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.RightToLeft,
                .AutoSize = True,
                .WrapContents = False,
                .Margin = New Padding(0, 8, 0, 0)
            }
            Me.pnlButtons.Controls.Add(Me.btnClose)
            Me.pnlButtons.Controls.Add(Me.btnOpen)
            Me.pnlButtons.Controls.Add(Me.btnDelete)
            Me.pnlButtons.Controls.Add(Me.btnClearAll)
            Me.pnlButtons.Controls.Add(Me.btnRefresh)

            Me.outer.Controls.Add(Me.lblTitle, 0, 0)
            Me.outer.Controls.Add(Me.txtFilter, 0, 1)
            Me.outer.Controls.Add(Me.lstItems, 0, 2)
            Me.outer.Controls.Add(Me.pnlButtons, 0, 3)
            Me.Controls.Add(Me.outer)

            Me.AcceptButton = Me.btnClose
            Me.CancelButton = Me.btnClose

            AddHandler Me.Shown,
                Sub(s, e)
                    SendMessage(Me.txtFilter.Handle, EM_SETCUEBANNER, CType(1, IntPtr), "Filter…")
                    UpdateButtons()
                End Sub

            Me.ResumeLayout(False)
        End Sub

        Protected Overrides Sub OnHandleCreated(e As System.EventArgs)
            MyBase.OnHandleCreated(e)

            If _dialogOwnerScope Is Nothing Then
                _dialogOwnerScope = SharedLibrary.SharedMethods.PushDialogOwner(Me)
            End If
        End Sub

        Protected Overrides Sub OnHandleDestroyed(e As System.EventArgs)
            Dim scope As System.IDisposable = _dialogOwnerScope
            _dialogOwnerScope = Nothing

            If scope IsNot Nothing Then
                Try
                    scope.Dispose()
                Catch
                End Try
            End If

            MyBase.OnHandleDestroyed(e)
        End Sub

        Private Shared Sub ConfigureStandardButton(button As Button)
            button.AutoSize = True
            button.AutoSizeMode = AutoSizeMode.GrowAndShrink
            button.UseVisualStyleBackColor = True
            button.Padding = New Padding(10, 4, 10, 4)
            button.MinimumSize = New Size(0, 0)
            button.Margin = New Padding(6, 0, 0, 0)
        End Sub

        Private Shared Function BuildDisplay(entry As SessionMemoryEntry) As String
            If entry Is Nothing Then Return ""

            Dim summary As String = If(entry.Summary, "").Trim()
            If summary.Length = 0 Then
                summary = "(no summary)"
            End If

            Return "[" & If(entry.Key, "") & "] " & summary
        End Function

        Private Sub ReloadEntries()
            Try
                entries = If(SessionMemory.List(), New List(Of SessionMemoryEntry)())
            Catch ex As Exception
                Debug.WriteLine("[SessionMemoryViewerForm] ReloadEntries ERROR: " & ex.Message)
                entries = New List(Of SessionMemoryEntry)()
            End Try

            Me.lblTitle.Text = "Entries stored in session memory (persisted until you delete them): " & entries.Count.ToString()
            ApplyFilter()
        End Sub

        Private Sub ApplyFilter()
            Dim filter As String = If(Me.txtFilter.Text, "").Trim().ToLowerInvariant()

            Me.lstItems.BeginUpdate()
            Try
                Me.lstItems.Items.Clear()

                For Each entry As SessionMemoryEntry In entries
                    Dim display As String = BuildDisplay(entry)
                    Dim searchText As String = (If(entry.Key, "") & " " & If(entry.Summary, "") & " " & display).ToLowerInvariant()

                    If filter.Length = 0 OrElse searchText.Contains(filter) Then
                        Me.lstItems.Items.Add(New ItemWrap With {
                            .Entry = entry,
                            .Display = display
                        })
                    End If
                Next

                If Me.lstItems.Items.Count > 0 Then
                    Me.lstItems.SelectedIndex = 0
                Else
                    Me.lstItems.ClearSelected()
                End If
            Finally
                Me.lstItems.EndUpdate()
            End Try

            UpdateButtons()
        End Sub

        Private Sub UpdateButtons()
            Dim hasSelection As Boolean = TryCast(Me.lstItems.SelectedItem, ItemWrap) IsNot Nothing
            Dim hasEntries As Boolean = Me.entries.Count > 0

            Me.btnOpen.Enabled = hasSelection
            Me.btnDelete.Enabled = hasSelection
            Me.btnClearAll.Enabled = hasEntries
        End Sub

        Private Sub OnFilterChanged(sender As Object, e As EventArgs)
            ApplyFilter()
        End Sub

        Private Sub OnSelectionChanged(sender As Object, e As EventArgs)
            UpdateButtons()
        End Sub

        Private Sub OnOpenClicked(sender As Object, e As EventArgs)
            Dim sel = TryCast(Me.lstItems.SelectedItem, ItemWrap)
            If sel Is Nothing OrElse sel.Entry Is Nothing Then Return

            Try
                Dim forceJson As Boolean = False
                Dim editorPath As String = ExportEntryToEditorFile(sel.Entry, forceJson)
                Dim headerText As String =
                    "Session memory entry: " & sel.Entry.Key & Environment.NewLine &
                    "Summary: " & If(String.IsNullOrWhiteSpace(sel.Entry.Summary), "(none)", sel.Entry.Summary.Trim())

                SharedLibrary.SharedMethods.ShowTextFileEditor(
                    editorPath,
                    headerText,
                    ForceJson:=forceJson,
                    ownerHandle:=Me.Handle)
            Catch ex As Exception
                SharedLibrary.SharedMethods.ShowCustomMessageBox("Failed to open memory entry: " & ex.Message)
            End Try
        End Sub

        Private Shared Function ExportEntryToEditorFile(entry As SessionMemoryEntry, ByRef forceJson As Boolean) As String
            Dim folderPath As String = Path.Combine(Path.GetTempPath(), SharedLibrary.SharedMethods.AN, "SessionMemory")
            Directory.CreateDirectory(folderPath)

            Dim safeKey As String = MakeSafeFileNamePart(If(entry.Key, "memory-entry"))
            Dim fileExtension As String = ".txt"
            Dim content As String = ""

            If entry IsNot Nothing AndAlso entry.Value IsNot Nothing Then
                If entry.Value.Type = JTokenType.Object OrElse entry.Value.Type = JTokenType.Array Then
                    fileExtension = ".json"
                    forceJson = True
                    content = entry.Value.ToString(Formatting.Indented)
                Else
                    content = entry.Value.ToString()
                End If
            End If

            Dim filePath As String = Path.Combine(folderPath, safeKey & fileExtension)
            File.WriteAllText(filePath, content, System.Text.Encoding.UTF8)
            Return filePath
        End Function

        Private Shared Function MakeSafeFileNamePart(value As String) As String
            Dim result As String = If(value, "").Trim()
            If result.Length = 0 Then
                Return "memory-entry"
            End If

            For Each invalidChar As Char In Path.GetInvalidFileNameChars()
                result = result.Replace(invalidChar, "_"c)
            Next

            If result.Length > 80 Then
                result = result.Substring(0, 80).Trim()
            End If

            If result.Length = 0 Then
                result = "memory-entry"
            End If

            Return result
        End Function

        Private Sub OnDeleteClicked(sender As Object, e As EventArgs)
            Dim sel = TryCast(Me.lstItems.SelectedItem, ItemWrap)
            If sel Is Nothing OrElse sel.Entry Is Nothing Then Return

            If SharedLibrary.SharedMethods.ShowCustomYesNoBox(
                "Delete memory entry '" & sel.Entry.Key & "'?",
                "Yes",
                "No",
                "Confirm") = 1 Then
                SessionMemory.Delete(sel.Entry.Key)
                ReloadEntries()
            End If
        End Sub

        Private Sub OnClearAllClicked(sender As Object, e As EventArgs)
            If Me.entries.Count = 0 Then Return

            If SharedLibrary.SharedMethods.ShowCustomYesNoBox(
                "Delete ALL session-memory entries? This cannot be undone.",
                "Yes",
                "No",
                "Confirm") = 1 Then
                SessionMemory.Clear()
                ReloadEntries()
            End If
        End Sub

        Private Class ItemWrap
            Public Property Entry As SessionMemoryEntry
            Public Property Display As String

            Public Overrides Function ToString() As String
                Return If(Display, MyBase.ToString())
            End Function
        End Class

    End Class

End Namespace
