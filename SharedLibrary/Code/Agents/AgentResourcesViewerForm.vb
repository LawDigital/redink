' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: AgentResourcesViewerForm.vb
' Purpose: Read-only inspector form for discovered Skills and Agents (central +
'          local resources). Allows filtering, viewing details, opening folders,
'          and toggling skill author mode for resource management.
'
' UI Features:
'  - Lists discovered skills and agents with descriptions.
'  - Filter textbox for searching by name.
'  - "View" button to inspect resource details.
'  - "Open Folder" button to browse the resource directory.
'  - "Refresh" to rescan the discovery paths.
'  - "Author Mode" checkbox to enable/disable skill authoring.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms

Namespace Agents

    Public Class AgentResourcesViewerForm
        Inherits Form

        Private Const EM_SETCUEBANNER As Integer = &H1501

        <System.Runtime.InteropServices.DllImport("user32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Unicode)>
        Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As String) As IntPtr
        End Function

        Private lblTitle As Label
        Private txtFilter As TextBox
        Private lstItems As ListBox
        Private btnView As Button
        Private btnRefresh As Button
        Private btnClose As Button
        Private btnOpenFolder As Button
        Private pnlButtons As FlowLayoutPanel
        Private outer As TableLayoutPanel
        Private chkAuthorMode As CheckBox
        Private chkAuthorCentral As CheckBox
        Private _dialogOwnerScope As System.IDisposable = Nothing

        Private Class Entry
            Public Property Display As String
            Public Property Path As String
            Public Property Description As String

            Public Overrides Function ToString() As String
                Return If(Display, MyBase.ToString())
            End Function
        End Class

        Private ReadOnly entries As New List(Of Entry)

        Public Sub New(Optional title As String = Nothing)
            InitializeComponent(title)
            ReloadEntries()
        End Sub

        Private Sub InitializeComponent(title As String)
            Me.Text = If(String.IsNullOrWhiteSpace(title), SharedLibrary.SharedMethods.AN & " - Manage Skills & Agents", title)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.MinimizeBox = False
            Me.MaximizeBox = False
            Me.ShowInTaskbar = False
            Me.FormBorderStyle = FormBorderStyle.Sizable
            Me.AutoScaleMode = AutoScaleMode.Dpi
            Me.AutoScaleDimensions = New SizeF(96.0F, 96.0F)
            Me.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            Me.Width = 980
            Me.Height = 620
            Me.MinimumSize = New Size(820, 440)

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
                .Text = "Discovered skills and agents:",
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
            Me.lstItems.DisplayMember = NameOf(Entry.Display)
            AddHandler Me.lstItems.DoubleClick, AddressOf OnViewClicked
            AddHandler Me.lstItems.SelectedIndexChanged, AddressOf OnSelectionChanged

            Me.btnView = New Button() With {.Text = "Edit"}
            ConfigureStandardButton(Me.btnView)
            AddHandler Me.btnView.Click, AddressOf OnViewClicked

            Me.btnRefresh = New Button() With {.Text = "Refresh"}
            ConfigureStandardButton(Me.btnRefresh)
            AddHandler Me.btnRefresh.Click, Sub(s, e) ReloadEntries()

            Me.btnOpenFolder = New Button() With {.Text = "Open Folder"}
            ConfigureStandardButton(Me.btnOpenFolder)
            AddHandler Me.btnOpenFolder.Click, AddressOf OnOpenFolderClicked

            Me.chkAuthorMode = New CheckBox() With {
                .Text = "Skill-author mode (agent may write inside skills)",
                .AutoSize = True,
                .Checked = SkillAuthorMode.IsActive,
                .Margin = New Padding(0, 8, 12, 0)
            }
            AddHandler Me.chkAuthorMode.CheckedChanged,
                Sub(s, e)
                    If chkAuthorMode.Checked AndAlso Not SkillAuthorMode.IsActive Then
                        SkillAuthorMode.Enable()
                    ElseIf Not chkAuthorMode.Checked AndAlso SkillAuthorMode.IsActive Then
                        SkillAuthorMode.Disable()
                    End If
                    Me.chkAuthorCentral.Enabled = chkAuthorMode.Checked
                End Sub

            Me.chkAuthorCentral = New CheckBox() With {
                .Text = "Also allow writing to the shared/central resource folder",
                .AutoSize = True,
                .Checked = SkillAuthorMode.AllowCentralWrites,
                .Enabled = SkillAuthorMode.IsActive,
                .Margin = New Padding(0, 8, 12, 0)
            }
            AddHandler Me.chkAuthorCentral.CheckedChanged,
                Sub(s, e)
                    SkillAuthorMode.AllowCentralWrites = chkAuthorCentral.Checked
                End Sub

            Me.btnClose = New Button() With {.Text = "Close", .DialogResult = DialogResult.OK}
            ConfigureStandardButton(Me.btnClose)

            Me.pnlButtons = New FlowLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.RightToLeft,
                .AutoSize = True,
                .WrapContents = True,
                .Margin = New Padding(0, 8, 0, 0)
            }
            Me.pnlButtons.Controls.Add(Me.btnClose)
            Me.pnlButtons.Controls.Add(Me.btnView)
            Me.pnlButtons.Controls.Add(Me.btnOpenFolder)
            Me.pnlButtons.Controls.Add(Me.btnRefresh)
            Me.pnlButtons.Controls.Add(Me.chkAuthorMode)
            Me.pnlButtons.Controls.Add(Me.chkAuthorCentral)

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

        Private Sub ReloadEntries()
            entries.Clear()

            Try
                AgentResources.Refresh()

                For Each sk In AgentResources.Skills
                    entries.Add(New Entry With {
                        .Display = "[Skill]  " & sk.Name & "  (" & If(sk.IsLocal, "local", "central") & ")" &
                                   If(String.IsNullOrWhiteSpace(sk.Description), "", " — " & sk.Description.Trim()),
                        .Path = sk.FilePath,
                        .Description = If(sk.Description, "")
                    })
                Next

                For Each ag In AgentResources.Agents
                    entries.Add(New Entry With {
                        .Display = "[Agent] " & ag.Name & "  (" & If(ag.IsLocal, "local", "central") & ")" &
                                   If(String.IsNullOrWhiteSpace(ag.Description), "", " — " & ag.Description.Trim()),
                        .Path = ag.FilePath,
                        .Description = If(ag.Description, "")
                    })
                Next

                Me.lblTitle.Text = $"Discovered skills and agents: {AgentResources.Skills.Count} skills, {AgentResources.Agents.Count} agents (local entries override central)"
            Catch ex As Exception
                Debug.WriteLine("[AgentResourcesViewerForm] ReloadEntries ERROR: " & ex.Message)
                Me.lblTitle.Text = "Discovered skills and agents:"
            End Try

            ApplyFilter()
        End Sub

        Private Sub ApplyFilter()
            Dim filter As String = If(Me.txtFilter.Text, "").Trim().ToLowerInvariant()

            Me.lstItems.BeginUpdate()
            Try
                Me.lstItems.Items.Clear()

                For Each e In entries
                    If filter.Length = 0 OrElse
                       e.Display.ToLowerInvariant().Contains(filter) OrElse
                       e.Description.ToLowerInvariant().Contains(filter) Then
                        Me.lstItems.Items.Add(e)
                    End If
                Next

                If Me.lstItems.Items.Count > 0 Then
                    Me.lstItems.SelectedIndex = 0
                End If
            Finally
                Me.lstItems.EndUpdate()
            End Try

            UpdateButtons()
        End Sub

        Private Sub UpdateButtons()
            Dim hasSelection As Boolean = Me.lstItems.SelectedItem IsNot Nothing
            Me.btnView.Enabled = hasSelection
            Me.btnOpenFolder.Enabled = Not String.IsNullOrWhiteSpace(AgentResources.ConfiguredLocalPath)
        End Sub

        Private Sub OnSelectionChanged(sender As Object, e As EventArgs)
            UpdateButtons()
        End Sub

        Private Sub OnOpenFolderClicked(sender As Object, e As EventArgs)
            Dim localPath As String = AgentResources.ConfiguredLocalPath

            If String.IsNullOrWhiteSpace(localPath) Then
                SharedLibrary.SharedMethods.ShowCustomMessageBox("No local agent resources path is configured.")
                Return
            End If

            Try
                If Not Directory.Exists(localPath) Then
                    Directory.CreateDirectory(localPath)
                End If

                Dim psi As New ProcessStartInfo("explorer.exe", """" & localPath & """") With {
                    .UseShellExecute = True
                }
                Process.Start(psi)
            Catch ex As Exception
                SharedLibrary.SharedMethods.ShowCustomMessageBox("Failed to open folder: " & ex.Message)
            End Try
        End Sub

        Private Sub OnFilterChanged(sender As Object, e As EventArgs)
            ApplyFilter()
        End Sub

        Private Sub OnViewClicked(sender As Object, e As EventArgs)
            Dim sel = TryCast(Me.lstItems.SelectedItem, Entry)
            If sel Is Nothing Then Return

            If String.IsNullOrWhiteSpace(sel.Path) OrElse Not File.Exists(sel.Path) Then
                SharedLibrary.SharedMethods.ShowCustomMessageBox("File not found: " & If(sel.Path, ""))
                Return
            End If

            Try
                Dim headerText As String =
                    "Edit agent resource file:" & Environment.NewLine &
                    sel.Path

                SharedLibrary.SharedMethods.ShowTextFileEditor(
                    sel.Path,
                    headerText,
                    ownerHandle:=Me.Handle)
            Catch ex As Exception
                SharedLibrary.SharedMethods.ShowCustomMessageBox("Failed to open file: " & ex.Message)
            End Try
        End Sub

    End Class

End Namespace
