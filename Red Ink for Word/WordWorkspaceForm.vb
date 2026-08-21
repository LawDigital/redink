' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: WordWorkspaceForm.vb
' PURPOSE
'   Manage the Word agent workspace folder and permissions. The agent writes files there.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports SharedLibrary.Agents
Imports SharedLibrary.SharedLibrary

''' <summary>
''' Modal dialog for designating and managing the Word agent workspace.
''' </summary>
Public Class WordWorkspaceForm
    Inherits Form

    Protected Overrides Sub OnShown(e As System.EventArgs)
        MyBase.OnShown(e)
        Me.TopMost = True
        SharedLibrary.SharedLibrary.SharedMethods.ForceDialogToForeground(Me)
        SharedLibrary.SharedLibrary.SharedMethods.AttachForeignForegroundWatchdog(Me)
    End Sub


    Private Const HostKey As String = "word"
    Private Const EM_SETCUEBANNER As Integer = &H1501

    <System.Runtime.InteropServices.DllImport("user32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Unicode)>
    Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As String) As IntPtr
    End Function

    Private lblTitle As Label
    Private txtRoot As TextBox
    Private btnBrowse As Button
    Private chkPersist As CheckBox
    Private chkRead As CheckBox
    Private chkWrite As CheckBox
    Private chkMoveCopyRename As CheckBox
    Private chkDelete As CheckBox
    Private chkIncludeHidden As CheckBox
    Private chkAllowActiveDocWrites As CheckBox
    Private lblFallback As Label
    Private btnSave As Button
    Private btnClear As Button
    Private btnCancel As Button
    Private pnlButtons As FlowLayoutPanel
    Private outer As TableLayoutPanel
    Private _dialogOwnerScope As System.IDisposable = Nothing

    Protected Overrides Sub OnHandleCreated(e As System.EventArgs)
        MyBase.OnHandleCreated(e)

        If _dialogOwnerScope Is Nothing Then
            _dialogOwnerScope = SharedLibrary.SharedLibrary.SharedMethods.PushDialogOwner(Me)
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

    Public Sub New()
        InitializeComponent()
        LoadCurrent()
    End Sub

    Private Sub InitializeComponent()
        Me.SuspendLayout()

        Me.Text = SharedLibrary.SharedLibrary.SharedMethods.AN & " - Workspace"
        Me.StartPosition = FormStartPosition.CenterParent
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        Me.ShowInTaskbar = False
        Me.FormBorderStyle = FormBorderStyle.Sizable
        Me.AutoScaleMode = AutoScaleMode.Dpi
        Me.AutoScaleDimensions = New SizeF(96.0F, 96.0F)
        Me.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
        Me.MinimumSize = New Size(700, 320)
        Me.Size = Me.MinimumSize

        Try
            Dim bmp As New Bitmap(SharedLibrary.SharedLibrary.SharedMethods.GetLogoBitmap(SharedLibrary.SharedLibrary.SharedMethods.LogoType.Standard))
            Me.Icon = Icon.FromHandle(bmp.GetHicon())
        Catch
        End Try

        outer = New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 2,
            .RowCount = 10,
            .Padding = New Padding(12, 12, 12, 0),
            .Margin = New Padding(0)
        }
        outer.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        outer.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))

        For i As Integer = 1 To 10
            outer.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        Next

        lblTitle = New Label() With {
            .AutoSize = True,
            .Anchor = AnchorStyles.Left,
            .Text = "Designate a workspace folder to make available files to the agent. Agent file output goes there, too; otherwise it goes to your Desktop.",
            .Margin = New Padding(0, 0, 0, 4)
        }

        txtRoot = New TextBox() With {
            .Dock = DockStyle.Fill,
            .Margin = New Padding(0, 0, 6, 0)
        }

        btnBrowse = New Button() With {.Text = "Browse…"}
        ConfigureBrowseButton(btnBrowse, txtRoot)
        AddHandler btnBrowse.Click, AddressOf OnBrowseClicked

        chkPersist = New CheckBox() With {
            .Text = "Remember this workspace until I revoke it",
            .AutoSize = True,
            .Checked = True,
            .Margin = New Padding(0, 0, 0, 2)
        }

        chkRead = New CheckBox() With {
            .Text = "Allow the agent to read files",
            .AutoSize = True,
            .Checked = True,
            .Margin = New Padding(0, 0, 0, 2)
        }

        chkWrite = New CheckBox() With {
            .Text = "Allow the agent to write/create files",
            .AutoSize = True,
            .Checked = True,
            .Margin = New Padding(0, 0, 0, 2)
        }

        chkMoveCopyRename = New CheckBox() With {
            .Text = "Allow the agent to move/copy/rename",
            .AutoSize = True,
            .Checked = True,
            .Margin = New Padding(0, 0, 0, 2)
        }

        chkDelete = New CheckBox() With {
            .Text = "Allow the agent to delete (to Recycle Bin)",
            .AutoSize = True,
            .Checked = SharedMethods.WorkspaceDeleteByDefaultOn,
            .Margin = New Padding(0, 0, 0, 2)
        }

        chkIncludeHidden = New CheckBox() With {
            .Text = "Include hidden/system files in listings",
            .AutoSize = True,
            .Checked = False,
            .Margin = New Padding(0, 0, 0, 2)
        }

        chkAllowActiveDocWrites = New CheckBox() With {
            .Text = "Allow the agent to modify the currently open Word document (worddoc_* writes)",
            .AutoSize = True,
            .Checked = Not SharedLibrary.Agents.WordHostPolicy.ActiveDocReadOnly,
            .Margin = New Padding(0, 0, 0, 2)
        }

        lblFallback = New Label() With {
            .AutoSize = True,
            .ForeColor = SystemColors.GrayText,
            .Text = "If no workspace is configured, the agent writes to your Desktop. Skill scripts/references remain read-only.",
            .Margin = New Padding(0, 2, 0, 0)
        }

        btnSave = New Button() With {.Text = "Save", .DialogResult = DialogResult.OK}
        ConfigureStandardButton(btnSave)
        AddHandler btnSave.Click, AddressOf OnSaveClicked

        btnClear = New Button() With {.Text = "Clear"}
        ConfigureStandardButton(btnClear)
        AddHandler btnClear.Click, AddressOf OnClearClicked

        btnCancel = New Button() With {.Text = "Cancel", .DialogResult = DialogResult.Cancel}
        ConfigureStandardButton(btnCancel)

        pnlButtons = New FlowLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .FlowDirection = FlowDirection.RightToLeft,
            .AutoSize = True,
            .WrapContents = False,
            .Margin = New Padding(0, 8, 0, 0)
        }
        pnlButtons.Controls.Add(btnCancel)
        pnlButtons.Controls.Add(btnSave)
        pnlButtons.Controls.Add(btnClear)

        outer.Controls.Add(lblTitle, 0, 0)
        outer.SetColumnSpan(lblTitle, 2)

        outer.Controls.Add(txtRoot, 0, 1)
        outer.Controls.Add(btnBrowse, 1, 1)

        outer.Controls.Add(chkPersist, 0, 2)
        outer.SetColumnSpan(chkPersist, 2)

        outer.Controls.Add(chkRead, 0, 3)
        outer.SetColumnSpan(chkRead, 2)

        outer.Controls.Add(chkWrite, 0, 4)
        outer.SetColumnSpan(chkWrite, 2)

        outer.Controls.Add(chkMoveCopyRename, 0, 5)
        outer.SetColumnSpan(chkMoveCopyRename, 2)

        outer.Controls.Add(chkDelete, 0, 6)
        outer.SetColumnSpan(chkDelete, 2)

        outer.Controls.Add(chkIncludeHidden, 0, 7)
        outer.SetColumnSpan(chkIncludeHidden, 2)

        outer.Controls.Add(chkAllowActiveDocWrites, 0, 8)
        outer.SetColumnSpan(chkAllowActiveDocWrites, 2)

        outer.Controls.Add(lblFallback, 0, 9)
        outer.SetColumnSpan(lblFallback, 2)

        outer.Controls.Add(pnlButtons, 0, 10)
        outer.SetColumnSpan(pnlButtons, 2)

        Me.Controls.Add(outer)
        Me.AcceptButton = btnSave
        Me.CancelButton = btnCancel

        AddHandler Me.Shown,
            Sub(s, e)
                UpdateWrappingWidths()
                SendMessage(txtRoot.Handle, EM_SETCUEBANNER, CType(1, IntPtr), "Path to workspace folder…")
                Me.BeginInvoke(New MethodInvoker(AddressOf AdjustHeightToContent))
            End Sub

        AddHandler Me.Resize,
            Sub(s, e)
                UpdateWrappingWidths()
            End Sub

        Me.ResumeLayout(False)
    End Sub

    Private Sub UpdateWrappingWidths()
        If outer Is Nothing Then Return

        Dim availableWidth As Integer = Math.Max(420, Me.ClientSize.Width - outer.Padding.Left - outer.Padding.Right)
        Dim wrappedWidth As Integer = Math.Max(320, availableWidth - 6)

        lblTitle.MaximumSize = New Size(wrappedWidth, 0)
        chkAllowActiveDocWrites.MaximumSize = New Size(wrappedWidth, 0)
        lblFallback.MaximumSize = New Size(wrappedWidth, 0)
    End Sub

    Private Shared Sub ConfigureBrowseButton(button As Button, relatedTextBox As TextBox)
        button.AutoSize = False
        button.UseVisualStyleBackColor = True
        button.Padding = New Padding(10, 0, 10, 0)
        button.Height = relatedTextBox.PreferredHeight + 2
        button.Width = Math.Max(90, TextRenderer.MeasureText(button.Text, button.Font).Width + 24)
        button.Margin = New Padding(0)
    End Sub

    Private Shared Sub ConfigureStandardButton(button As Button)
        button.AutoSize = True
        button.AutoSizeMode = AutoSizeMode.GrowAndShrink
        button.UseVisualStyleBackColor = True
        button.Padding = New Padding(10, 4, 10, 4)
        button.MinimumSize = New Size(0, 0)
        button.Margin = New Padding(6, 0, 0, 0)
    End Sub

    Private Sub AdjustHeightToContent()
        Me.SuspendLayout()
        UpdateWrappingWidths()
        outer.PerformLayout()
        pnlButtons.PerformLayout()

        Dim preferred As Size = outer.GetPreferredSize(New Size(Me.ClientSize.Width, Integer.MaxValue))
        Dim targetClientHeight As Integer = preferred.Height + 10
        Dim nonClientHeight As Integer = Me.Height - Me.ClientSize.Height
        Dim targetFormHeight As Integer = targetClientHeight + nonClientHeight

        If targetFormHeight < 320 Then
            targetFormHeight = 320
        End If

        Me.MinimumSize = New Size(Me.MinimumSize.Width, targetFormHeight)
        Me.Height = targetFormHeight
        Me.ResumeLayout(True)
    End Sub

    Private Sub LoadCurrent()
        Dim st = WorkspaceStore.Load(HostKey)
        txtRoot.Text = If(st.RootPath, "")
        chkPersist.Checked = st.PersistUntilRevoked
        chkRead.Checked = st.AllowRead
        chkWrite.Checked = st.AllowWrite
        chkMoveCopyRename.Checked = st.AllowMoveCopyRename
        chkDelete.Checked = st.AllowDelete
        chkIncludeHidden.Checked = st.IncludeHiddenSystem
        chkAllowActiveDocWrites.Checked = Not SharedLibrary.Agents.WordHostPolicy.ActiveDocReadOnly
    End Sub

    Private Sub OnBrowseClicked(sender As Object, e As EventArgs)
        Using fbd As New FolderBrowserDialog()
            fbd.Description = "Choose a folder to use as the agent workspace."
            If Not String.IsNullOrWhiteSpace(txtRoot.Text) AndAlso Directory.Exists(txtRoot.Text) Then
                fbd.SelectedPath = txtRoot.Text
            End If

            If fbd.ShowDialog(Me) = DialogResult.OK Then
                txtRoot.Text = fbd.SelectedPath
            End If
        End Using
    End Sub

    Private Sub OnSaveClicked(sender As Object, e As EventArgs)
        Dim root As String = If(txtRoot.Text, "").Trim()

        If String.IsNullOrWhiteSpace(root) Then
            SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("Please specify a folder, or press Clear to unset the workspace.")
            Me.DialogResult = DialogResult.None
            Return
        End If

        If Not Directory.Exists(root) Then
            SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("Folder does not exist: " & root)
            Me.DialogResult = DialogResult.None
            Return
        End If

        Dim st As New WorkspaceState() With {
            .RootPath = Path.GetFullPath(root),
            .PersistUntilRevoked = chkPersist.Checked,
            .AllowRead = chkRead.Checked,
            .AllowWrite = chkWrite.Checked,
            .AllowMoveCopyRename = chkMoveCopyRename.Checked,
            .AllowDelete = chkDelete.Checked,
            .IncludeHiddenSystem = chkIncludeHidden.Checked
        }

        WorkspaceStore.Save(HostKey, st)
        WorkspaceTools.SetActive(st)
        SharedLibrary.Agents.WordHostPolicy.ActiveDocReadOnly = Not chkAllowActiveDocWrites.Checked
    End Sub

    Private Sub OnClearClicked(sender As Object, e As EventArgs)
        If MessageBox.Show("Clear the configured workspace? The agent will then write to your Desktop.", "Workspace",
                           MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            WorkspaceStore.Clear(HostKey)
            WorkspaceTools.SetActive(New WorkspaceState())
            SharedLibrary.Agents.WordHostPolicy.ActiveDocReadOnly = True
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End If
    End Sub

End Class
