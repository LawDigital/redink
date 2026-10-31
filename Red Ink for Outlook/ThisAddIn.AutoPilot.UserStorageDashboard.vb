' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.UserStorageDashboard.vb
' Purpose:
'   Admin dashboard for managing per-user AutoPilot storage (memory + files).
'   Provides a DataGridView of all user directories with actions to inspect,
'   edit, and delete memory and files.
'
' Actions:
'   - View Memory:    Opens the selected user's memory.txt in the text editor.
'   - Edit Memory:    Same as View Memory (file is editable in the editor).
'   - Delete Memory:  Deletes the selected user's memory file.
'   - View Files:     Shows a file listing dialog for the selected user.
'   - Delete File:    Deletes a single file from the selected user's home.
'   - Delete All Files: Deletes all files from the selected user's home.
'   - Delete User:    Deletes the entire user directory (memory + all files).
'   - Refresh:        Reloads the user list.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.IO
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    Private _apUserStorageDashboard As Form = Nothing

    ''' <summary>Shows the user storage admin dashboard. Creates it if not yet open.</summary>
    Friend Sub ShowUserStorageDashboard()
        If _apUserStorageDashboard IsNot Nothing AndAlso Not _apUserStorageDashboard.IsDisposed Then
            RefreshUserStorageDashboard()
            ShowHostModelessForm(_apUserStorageDashboard)
            _apUserStorageDashboard.BringToFront()
            Return
        End If

        _apUserStorageDashboard = CreateUserStorageDashboardForm()
        RefreshUserStorageDashboard()
        ShowHostModelessForm(_apUserStorageDashboard)
    End Sub

    ''' <summary>Closes the user storage dashboard if it is open.</summary>
    Friend Sub CloseUserStorageDashboard()
        Try
            If _apUserStorageDashboard IsNot Nothing AndAlso Not _apUserStorageDashboard.IsDisposed Then
                _apUserStorageDashboard.Close()
                _apUserStorageDashboard.Dispose()
            End If
        Catch
        End Try
        _apUserStorageDashboard = Nothing
    End Sub

    ''' <summary>Creates the user storage dashboard form with a DataGridView and action buttons.</summary>
    Private Function CreateUserStorageDashboardForm() As Form
        Dim frm As New Form() With {
            .Text = $"{AN6} AutoPilot — User Storage Manager",
            .Width = 1150,
            .Height = 520,
            .StartPosition = FormStartPosition.CenterScreen,
            .FormBorderStyle = FormBorderStyle.Sizable,
            .MinimumSize = New Drawing.Size(750, 350),
            .TopMost = False,
            .MaximizeBox = True,
            .MinimizeBox = True,
            .ShowInTaskbar = True,
            .AutoScaleMode = AutoScaleMode.Font,
            .Font = New Drawing.Font("Segoe UI", 9.0F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point)
        }

        Try
            Global.SharedLibrary.SharedLibrary.SharedMethods.ApplyStandardDialogIcon(frm)
        Catch
        End Try

        Dim mainPanel As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill, .ColumnCount = 1, .RowCount = 2, .Padding = New Padding(8)
        }
        mainPanel.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        mainPanel.RowStyles.Add(New RowStyle(SizeType.AutoSize))

        Dim dgv As New DataGridView() With {
            .Dock = DockStyle.Fill,
            .Name = "dgvUsers",
            .ReadOnly = True,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .AllowUserToResizeRows = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .MultiSelect = False,
            .RowHeadersVisible = False,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            .BackgroundColor = Drawing.SystemColors.Window,
            .BorderStyle = BorderStyle.None,
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
            .DefaultCellStyle = New DataGridViewCellStyle() With {
                .Font = New Drawing.Font("Segoe UI", 8.5F),
                .WrapMode = DataGridViewTriState.False
            },
            .AlternatingRowsDefaultCellStyle = New DataGridViewCellStyle() With {
                .BackColor = Drawing.Color.FromArgb(245, 245, 250)
            }
        }

        dgv.Columns.Add("colEmail", "Verified e-mail")
        dgv.Columns.Add("colUser", "User (folder)")
        dgv.Columns.Add("colMemory", "Memory")
        dgv.Columns.Add("colFiles", "Files")
        dgv.Columns.Add("colSize", "Size")
        dgv.Columns.Add("colCoaching", "Inky recommendations")

        dgv.Columns("colEmail").FillWeight = 27
        dgv.Columns("colUser").FillWeight = 20
        dgv.Columns("colMemory").FillWeight = 10
        dgv.Columns("colFiles").FillWeight = 8
        dgv.Columns("colSize").FillWeight = 10
        dgv.Columns("colCoaching").FillWeight = 25

        mainPanel.Controls.Add(dgv, 0, 0)

        ' Button panel
        Dim buttonPanel As New FlowLayoutPanel() With {
            .Dock = DockStyle.Fill, .FlowDirection = FlowDirection.LeftToRight,
            .AutoSize = True, .Padding = New Padding(0, 4, 0, 0), .WrapContents = True
        }

        Dim btnRefresh As New Button() With {
            .Text = "Refresh", .AutoSize = True, .Padding = New Padding(8, 4, 8, 4)
        }
        AddHandler btnRefresh.Click, Sub(s, e) RefreshUserStorageDashboard()

        Dim btnViewMemory As New Button() With {
            .Text = "View/Edit Memory", .AutoSize = True, .Padding = New Padding(8, 4, 8, 4)
        }
        AddHandler btnViewMemory.Click, Sub(s, e) UserStorageAction_ViewMemory(dgv)

        Dim btnDeleteMemory As New Button() With {
            .Text = "Delete Memory", .AutoSize = True, .Padding = New Padding(8, 4, 8, 4)
        }
        AddHandler btnDeleteMemory.Click, Sub(s, e) UserStorageAction_DeleteMemory(dgv)

        Dim btnViewFiles As New Button() With {
            .Text = "View Files", .AutoSize = True, .Padding = New Padding(8, 4, 8, 4)
        }
        AddHandler btnViewFiles.Click, Sub(s, e) UserStorageAction_ViewFiles(dgv)

        Dim btnDeleteFile As New Button() With {
            .Text = "Delete File…", .AutoSize = True, .Padding = New Padding(8, 4, 8, 4)
        }
        AddHandler btnDeleteFile.Click, Sub(s, e) UserStorageAction_DeleteFile(dgv)

        Dim btnDeleteAllFiles As New Button() With {
            .Text = "Delete All Files", .AutoSize = True, .Padding = New Padding(8, 4, 8, 4)
        }
        AddHandler btnDeleteAllFiles.Click, Sub(s, e) UserStorageAction_DeleteAllFiles(dgv)

        Dim btnDeleteUser As New Button() With {
            .Text = "Delete User", .AutoSize = True, .Padding = New Padding(8, 4, 8, 4),
            .ForeColor = Drawing.Color.DarkRed
        }
        AddHandler btnDeleteUser.Click, Sub(s, e) UserStorageAction_DeleteUser(dgv)

        Dim btnPromptCoaching As New System.Windows.Forms.Button() With {
            .Text = "Inky recommendations…", .AutoSize = True, .Padding = New System.Windows.Forms.Padding(8, 4, 8, 4)
        }
        AddHandler btnPromptCoaching.Click, Sub(s, e) UserStorageAction_PromptCoaching(dgv, False)

        Dim btnAddPromptCoachingUser As New System.Windows.Forms.Button() With {
            .Text = "Add recommendation user…", .AutoSize = True, .Padding = New System.Windows.Forms.Padding(8, 4, 8, 4)
        }
        AddHandler btnAddPromptCoachingUser.Click, Sub(s, e) UserStorageAction_PromptCoaching(dgv, True)

        Dim btnClose As New Button() With {
            .Text = "Close", .AutoSize = True, .Padding = New Padding(8, 4, 8, 4)
        }
        AddHandler btnClose.Click, Sub(s, e) frm.Hide()

        buttonPanel.Controls.Add(btnRefresh)
        buttonPanel.Controls.Add(btnViewMemory)
        buttonPanel.Controls.Add(btnDeleteMemory)
        buttonPanel.Controls.Add(btnViewFiles)
        buttonPanel.Controls.Add(btnDeleteFile)
        buttonPanel.Controls.Add(btnDeleteAllFiles)
        buttonPanel.Controls.Add(btnDeleteUser)
        buttonPanel.Controls.Add(btnPromptCoaching)
        buttonPanel.Controls.Add(btnAddPromptCoachingUser)
        buttonPanel.Controls.Add(btnClose)
        mainPanel.Controls.Add(buttonPanel, 0, 1)

        frm.Controls.Add(mainPanel)

        ' Hide instead of closing
        AddHandler frm.FormClosing, Sub(s As Object, e As FormClosingEventArgs)
                                        If e.CloseReason = CloseReason.UserClosing AndAlso Not IsHostModelessClosing Then
                                            e.Cancel = True
                                            frm.Hide()
                                        End If
                                    End Sub

        Return frm
    End Function

    ''' <summary>Refreshes the user storage dashboard DataGridView.</summary>
    Friend Sub RefreshUserStorageDashboard()
        If _apUserStorageDashboard Is Nothing OrElse _apUserStorageDashboard.IsDisposed Then Return

        Dim dgv = _apUserStorageDashboard.Controls.Find("dgvUsers", True).
                    OfType(Of DataGridView)().FirstOrDefault()
        If dgv Is Nothing Then Return

        If dgv.InvokeRequired Then
            dgv.BeginInvoke(New MethodInvoker(Sub() RefreshUserStorageDashboardCore(dgv)))
        Else
            RefreshUserStorageDashboardCore(dgv)
        End If
    End Sub

    ''' <summary>Core logic for populating the user storage DataGridView.</summary>
    Private Sub RefreshUserStorageDashboardCore(dgv As DataGridView)
        dgv.Rows.Clear()
        Dim users = ListAllUserStorageDirs()
        Dim representedEmails As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)

        For Each u In users.OrderBy(Function(x) x.FolderName)
            Dim memoryLabel = If(u.HasMemory, "✓ Active", "—")
            Dim filesLabel = If(u.HomeFileCount > 0, u.HomeFileCount.ToString() & " file(s)", "—")
            Dim sizeLabel = If(u.HomeSizeBytes > 0, $"{u.HomeSizeBytes / 1024.0 / 1024.0:F1} MB", "—")

            Dim email As System.String = GetAutoPilotPromptCoachingDashboardEmail(u.FolderName)
            If email.Length > 0 Then representedEmails.Add(email)
            Dim rowIdx = dgv.Rows.Add(If(email.Length > 0, email, "(unverified legacy identity)"), u.FolderName, memoryLabel, filesLabel, sizeLabel, "")

            ' Color-code
            Dim row = dgv.Rows(rowIdx)
            If u.HasMemory Then
                row.Cells("colMemory").Style.ForeColor = Drawing.Color.DarkGreen
            Else
                row.Cells("colMemory").Style.ForeColor = Drawing.Color.Gray
            End If

            ' Store folder name in Tag for action lookups
            row.Tag = u.FolderName
            row.Cells("colEmail").Tag = email
            UpdateUserStoragePromptCoachingCell(row, email)
        Next

        For Each email As System.String In ListAutoPilotPromptCoachingKnownUsers()
            If representedEmails.Contains(email) Then Continue For
            Dim rowIndex As System.Int32 = dgv.Rows.Add(email, "(recommendations only)", "—", "—", "—", "")
            Dim row As System.Windows.Forms.DataGridViewRow = dgv.Rows(rowIndex)
            row.Tag = Nothing
            row.Cells("colEmail").Tag = email
            UpdateUserStoragePromptCoachingCell(row, email)
        Next
    End Sub

    Private Sub UpdateUserStoragePromptCoachingCell(row As System.Windows.Forms.DataGridViewRow, email As System.String)
        Dim cell As System.Windows.Forms.DataGridViewCell = row.Cells("colCoaching")
        If System.String.IsNullOrWhiteSpace(email) Then
            cell.Value = "Unavailable (unverified identity)"
            cell.ToolTipText = "A legacy folder name does not prove its owner's e-mail address. Use Add recommendation user with the exact address."
            Return
        End If
        Dim status As Newtonsoft.Json.Linq.JObject = GetAutoPilotPromptCoachingPreferenceStatus(email)
        Dim preference As System.String = status.Value(Of System.String)("preference")
        Dim enabled As System.Boolean = status.Value(Of System.Boolean)("effectiveEnabled")
        If preference = "unavailable" Then
            cell.Value = "Withheld (preference unreadable)"
            cell.Style.ForeColor = System.Drawing.Color.DarkRed
        ElseIf Not status.Value(Of System.Boolean)("globalEnabled") Then
            cell.Value = "System off; stored: " & preference
            cell.Style.ForeColor = System.Drawing.Color.Gray
        Else
            cell.Value = If(enabled, "On", "Off") & " — " &
                If(preference = "default", status.Value(Of System.String)("reason").Replace("_", " "), status.Value(Of System.String)("source") & " choice")
            cell.Style.ForeColor = If(enabled, System.Drawing.Color.DarkGreen, System.Drawing.Color.Gray)
        End If
        Dim timing As System.String = If(status.Value(Of System.String)("deliveryMode") = "immediate", "after analysis", "daily from 20:00")
        If preference <> "unavailable" Then cell.Value = System.Convert.ToString(cell.Value, System.Globalization.CultureInfo.InvariantCulture) & " — " & timing
        cell.ToolTipText = "Preference: " & preference & "; effective delivery: " & timing &
            "; personal timing: " & status.Value(Of System.String)("deliveryPreference") &
            ". Personal timing overrides the startup schedule, never a global disable or opt-out. Preferences survive memory/file deletion."
    End Sub

    Private Sub UserStorageAction_PromptCoaching(dgv As System.Windows.Forms.DataGridView, addUser As System.Boolean)
        Using SharedLibrary.SharedLibrary.SharedMethods.PushDialogOwner(_apUserStorageDashboard)
            Dim email As System.String = System.String.Empty
            If addUser Then
                Dim input As System.String = ShowCustomInputBox(
                    "Enter the exact e-mail address whose Inky recommendations you want to manage:",
                    "Inky recommendations — User", True, System.String.Empty)
                If input Is Nothing Then Return
                email = NormalizeAutoPilotPromptCoachingEmail(input)
            Else
                If dgv.SelectedRows.Count = 0 Then
                    ShowCustomMessageBox("Please select a user first.")
                    Return
                End If
                email = NormalizeAutoPilotPromptCoachingEmail(TryCast(dgv.SelectedRows(0).Cells("colEmail").Tag, System.String))
            End If
            If email.Length = 0 Then
                ShowCustomMessageBox("A verified, exact e-mail address is required. Legacy folder names cannot identify a recommendation recipient. Use Add recommendation user if needed.")
                Return
            End If
            Dim current As Newtonsoft.Json.Linq.JObject = GetAutoPilotPromptCoachingPreferenceStatus(email)
            Dim choices As New System.Collections.Generic.List(Of SharedLibrary.SharedLibrary.SharedMethods.SelectionItem) From {
                New SharedLibrary.SharedLibrary.SharedMethods.SelectionItem("Enable recommendations for this user", 1),
                New SharedLibrary.SharedLibrary.SharedMethods.SelectionItem("Disable recommendations for this user", 2),
                New Global.SharedLibrary.SharedLibrary.SharedMethods.SelectionItem("Reset both personal choices to system defaults / inclusion list", 3),
                New Global.SharedLibrary.SharedLibrary.SharedMethods.SelectionItem("Enable recommendations, once daily from 20:00", 4),
                New Global.SharedLibrary.SharedLibrary.SharedMethods.SelectionItem("Enable recommendations, after each qualifying request analysis", 5),
                New Global.SharedLibrary.SharedLibrary.SharedMethods.SelectionItem("Use system delivery schedule (keep current on/off choice)", 6)
            }
            Dim preference As System.String = current.Value(Of System.String)("preference")
            Dim selected As System.Int32 = SelectValue(choices, If(preference = "enabled", 1, If(preference = "disabled", 2, 3)),
                "Inky recommendations for " & email & System.Environment.NewLine &
                "Explicit disable overrides the inclusion list and default-for-all setting. Personal timing overrides the startup schedule, not a global disable." &
                System.Environment.NewLine & "Effective delivery: " & current.Value(Of System.String)("deliveryMode") & "; personal timing: " & current.Value(Of System.String)("deliveryPreference") & "." &
                System.Environment.NewLine & "Users can ask Inky or send INKY_TIPS ON/OFF/DAILY/IMMEDIATE/SYSTEM/STATUS/DEFAULT. Preferences survive memory/file deletion.",
                "Inky recommendations — Administrator")
            If selected < 1 OrElse selected > 6 Then Return
            Try
                Select Case selected
                    Case 1 : SetAutoPilotPromptCoachingPreference(email, "enabled", "admin")
                    Case 2 : SetAutoPilotPromptCoachingPreference(email, "disabled", "admin")
                    Case 3 : SetAutoPilotPromptCoachingPreference(email, "default", "admin", "default")
                    Case 4 : SetAutoPilotPromptCoachingPreference(email, "enabled", "admin", "daily")
                    Case 5 : SetAutoPilotPromptCoachingPreference(email, "enabled", "admin", "immediate")
                    Case 6 : SetAutoPilotPromptCoachingPreference(email, Nothing, "admin", "default")
                End Select
                RefreshUserStorageDashboard()
                If _apConfig Is Nothing OrElse Not _apConfig.PromptCoachingEnabled Then
                    ShowCustomMessageBox("The preference was saved. Inky recommendations are globally disabled, so no recommendation e-mails will be sent until they are enabled in the startup configuration.")
                End If
            Catch ex As System.Exception
                ApDashboardLog("Inky recommendation preference could not be saved: " & ex.GetType().Name & ": " & ex.Message, "warn")
                ShowCustomMessageBox("The Inky recommendation preference could not be saved safely. No successful change is confirmed. Please try again or check the diagnostic log.")
            End Try
        End Using
    End Sub

    ''' <summary>Gets the selected user's folder name from the DataGridView, or shows an error.</summary>
    Private Function GetSelectedUserFolder(dgv As DataGridView) As String
        If dgv.SelectedRows.Count = 0 Then
            ShowCustomMessageBox("Please select a user first.")
            Return Nothing
        End If
        If dgv.SelectedRows(0).Tag Is Nothing Then
            ShowCustomMessageBox("This user currently has a recommendation preference only, with no memory or home-file directory. Use Inky recommendations to change the mailing preference.")
            Return Nothing
        End If
        Return CStr(dgv.SelectedRows(0).Tag)
    End Function

    ' ── ACTION: View/Edit Memory ──

    Private Sub UserStorageAction_ViewMemory(dgv As DataGridView)
        Dim folder = GetSelectedUserFolder(dgv)
        If folder Is Nothing Then Return
        Try
            Dim memoryPath = Path.Combine(GetUserStorageRootDir(), folder, AP_UserMemoryFileName)
            If Not File.Exists(memoryPath) Then
                ' Create with default content so the admin can add items
                Dim dir = Path.GetDirectoryName(memoryPath)
                If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
                WriteFileWithRetry(memoryPath, GetDefaultMemoryFileContent())
            End If
            ShowTextFileEditor(memoryPath, $"Inky Memory for user: {folder}")
            RefreshUserStorageDashboard()
        Catch ex As Exception
            ShowCustomMessageBox($"Error opening memory file: {ex.Message}")
        End Try
    End Sub

    ' ── ACTION: Delete Memory ──

    Private Sub UserStorageAction_DeleteMemory(dgv As DataGridView)
        Dim folder = GetSelectedUserFolder(dgv)
        If folder Is Nothing Then Return
        Dim confirm = ShowCustomYesNoBox(
            $"Delete memory file for user '{folder}'?" & vbCrLf &
            "This will remove all stored preferences. Home files will not be affected.",
            "Delete", "Cancel",
            header:=$"{AN6} — Delete User Memory")
        If confirm <> 1 Then Return
        Try
            Dim memoryPath = Path.Combine(GetUserStorageRootDir(), folder, AP_UserMemoryFileName)
            If File.Exists(memoryPath) Then File.Delete(memoryPath)
            ApDashboardLog($"🧠 Admin deleted memory for: {folder}", "info")
            RefreshUserStorageDashboard()
        Catch ex As Exception
            ShowCustomMessageBox($"Error deleting memory: {ex.Message}")
        End Try
    End Sub

    ' ── ACTION: View Files ──

    Private Sub UserStorageAction_ViewFiles(dgv As DataGridView)
        Dim folder = GetSelectedUserFolder(dgv)
        If folder Is Nothing Then Return
        Try
            Dim homeDir = Path.Combine(GetUserStorageRootDir(), folder, AP_UserHomeSubdir)
            If Not Directory.Exists(homeDir) OrElse Directory.GetFiles(homeDir).Length = 0 Then
                ShowCustomMessageBox($"No files stored for user '{folder}'.")
                Return
            End If

            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine($"Files stored for user: {folder}")
            sb.AppendLine($"Directory: {homeDir}")
            sb.AppendLine()
            Dim totalSize As Long = 0
            For Each filePath In Directory.GetFiles(homeDir)
                Dim fi As New FileInfo(filePath)
                sb.AppendLine($"  • {fi.Name}  ({fi.Length / 1024.0:F0} KB, modified: {fi.LastWriteTime:yyyy-MM-dd HH:mm})")
                totalSize += fi.Length
            Next
            sb.AppendLine()
            sb.AppendLine($"Total: {totalSize / 1024.0 / 1024.0:F1} MB / {AP_UserHomeMaxBytes / 1024 / 1024:F0} MB limit")

            ShowCustomMessageBox(sb.ToString(), $"{AN6} — User Files: {folder}")
        Catch ex As Exception
            ShowCustomMessageBox($"Error listing files: {ex.Message}")
        End Try
    End Sub

    ' ── ACTION: Delete Single File ──

    Private Sub UserStorageAction_DeleteFile(dgv As DataGridView)
        Dim folder = GetSelectedUserFolder(dgv)
        If folder Is Nothing Then Return
        Try
            Dim homeDir = Path.Combine(GetUserStorageRootDir(), folder, AP_UserHomeSubdir)
            If Not Directory.Exists(homeDir) OrElse Directory.GetFiles(homeDir).Length = 0 Then
                ShowCustomMessageBox($"No files stored for user '{folder}'.")
                Return
            End If

            ' Build a selection list using MultiModelSelectorForm
            Dim files = Directory.GetFiles(homeDir)
            Dim items As New List(Of ModelConfig)()
            For Each filePath In files
                Dim fi As New FileInfo(filePath)
                items.Add(New ModelConfig() With {
                    .ModelDescription = $"{fi.Name}  ({fi.Length / 1024.0:F0} KB)",
                    .Model = fi.Name
                })
            Next

            Using dlg As New SharedLibrary.SharedLibrary.MultiModelSelectorForm(
                items, Nothing,
                title:=$"{AN6} — Select File(s) to Delete from {folder}",
                resetChecked:=True,
                instruction:="Select file(s) to delete:")

                dlg.StartPosition = FormStartPosition.CenterScreen
                Dim __safeDialogOwner345 As System.Windows.Forms.IWin32Window = SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
                If If(__safeDialogOwner345 IsNot Nothing, dlg.ShowDialog(__safeDialogOwner345), dlg.ShowDialog()) = DialogResult.OK Then
                    Dim selected = dlg.SelectedModels
                    If selected IsNot Nothing AndAlso selected.Count > 0 Then
                        Dim deleted = 0
                        For Each sel In selected
                            Dim filePath = Path.Combine(homeDir, sel.Model)
                            ' Security: validate containment
                            If Path.GetFullPath(filePath).StartsWith(Path.GetFullPath(homeDir), StringComparison.OrdinalIgnoreCase) Then
                                Try
                                    If File.Exists(filePath) Then
                                        File.Delete(filePath)
                                        deleted += 1
                                    End If
                                Catch
                                End Try
                            End If
                        Next
                        ApDashboardLog($"📁 Admin deleted {deleted} file(s) for: {folder}", "info")
                        RefreshUserStorageDashboard()
                    End If
                End If
            End Using
        Catch ex As Exception
            ShowCustomMessageBox($"Error deleting file: {ex.Message}")
        End Try
    End Sub

    ' ── ACTION: Delete All Files ──

    Private Sub UserStorageAction_DeleteAllFiles(dgv As DataGridView)
        Dim folder = GetSelectedUserFolder(dgv)
        If folder Is Nothing Then Return
        Dim confirm = ShowCustomYesNoBox(
            $"Delete ALL files for user '{folder}'?" & vbCrLf &
            "Memory will not be affected.",
            "Delete All", "Cancel",
            header:=$"{AN6} — Delete All User Files")
        If confirm <> 1 Then Return
        Try
            Dim homeDir = Path.Combine(GetUserStorageRootDir(), folder, AP_UserHomeSubdir)
            If Directory.Exists(homeDir) Then
                For Each filePath In Directory.GetFiles(homeDir)
                    Try : File.Delete(filePath) : Catch : End Try
                Next
            End If
            ApDashboardLog($"📁 Admin deleted all files for: {folder}", "info")
            RefreshUserStorageDashboard()
        Catch ex As Exception
            ShowCustomMessageBox($"Error deleting files: {ex.Message}")
        End Try
    End Sub

    ' ── ACTION: Delete Entire User ──

    Private Sub UserStorageAction_DeleteUser(dgv As DataGridView)
        Dim folder = GetSelectedUserFolder(dgv)
        If folder Is Nothing Then Return
        Dim confirm = ShowCustomYesNoBox(
            $"DELETE entire storage for user '{folder}'?" & vbCrLf & vbCrLf &
            "This will permanently remove:" & vbCrLf &
            "  • Memory file (all preferences)" & vbCrLf &
            "  • All stored files (home directory)" & vbCrLf &
            "  • The user directory itself" & vbCrLf & vbCrLf &
            "The separate Inky recommendation preference is retained, so an opt-out will not be lost." & vbCrLf &
            "This action cannot be undone.",
            "Delete Everything", "Cancel",
            header:=$"{AN6} — Delete User Storage")
        If confirm <> 1 Then Return
        Try
            Dim userDir = Path.Combine(GetUserStorageRootDir(), folder)
            If Directory.Exists(userDir) Then
                Directory.Delete(userDir, recursive:=True)
            End If
            ApDashboardLog($"🗑 Admin deleted entire user storage for: {folder}", "info")
            RefreshUserStorageDashboard()
        Catch ex As Exception
            ShowCustomMessageBox($"Error deleting user storage: {ex.Message}")
        End Try
    End Sub

End Class
