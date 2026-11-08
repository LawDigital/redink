' Part of Red Ink for Outlook. Copyright (c) LawDigital Ltd., Switzerland.
' UI only: dispatches file jobs through the existing host tooling harness.
Option Strict On
Option Explicit On

Friend NotInheritable Class BatchAgentJobForm
    Inherits System.Windows.Forms.Form

    Private ReadOnly _host As ThisAddIn
    Private ReadOnly _models As System.Collections.Generic.List(Of ThisAddIn.BatchModelChoice)
    Private ReadOnly _tools As System.Collections.Generic.List(Of Global.SharedLibrary.SharedLibrary.ModelConfig)
    Private _toolNames As System.Collections.Generic.List(Of System.String)
    Private ReadOnly _configuration As New System.Windows.Forms.TableLayoutPanel()
    Private ReadOnly _source As New System.Windows.Forms.TextBox()
    Private ReadOnly _instruction As New System.Windows.Forms.TextBox()
    Private ReadOnly _model As New System.Windows.Forms.ComboBox()
    Private ReadOnly _output As New System.Windows.Forms.ComboBox()
    Private ReadOnly _extensions As New System.Windows.Forms.TextBox()
    Private ReadOnly _recursive As New System.Windows.Forms.CheckBox With {.Text = "Include subfolders", .AutoSize = True}
    Private ReadOnly _carry As New System.Windows.Forms.CheckBox With {.Text = "Carry previous successful result forward", .AutoSize = True}
    Private ReadOnly _stop As New System.Windows.Forms.CheckBox With {.Text = "Stop after an exhausted file error", .AutoSize = True}
    Private ReadOnly _require As New System.Windows.Forms.CheckBox With {.Text = "Require a result file (not just an answer)", .AutoSize = True}
    Private ReadOnly _showDashboard As New System.Windows.Forms.CheckBox With {.Text = "Show dashboard with tooling log", .AutoSize = True}
    Private ReadOnly _retries As System.Windows.Forms.NumericUpDown = NumberBox(0, 10, 2)
    Private ReadOnly _delay As System.Windows.Forms.NumericUpDown = NumberBox(0, 300, 3)
    Private ReadOnly _minutes As System.Windows.Forms.NumericUpDown = NumberBox(1, 1440, 30)
    Private ReadOnly _iterations As System.Windows.Forms.NumericUpDown = NumberBox(1, 1000, 100)
    Private ReadOnly _maxSize As System.Windows.Forms.NumericUpDown = NumberBox(1, 2048, 100)
    Private ReadOnly _selectTools As New System.Windows.Forms.Button With {.Text = "Tools...", .AutoSize = True}
    Private ReadOnly _start As New System.Windows.Forms.Button With {.Text = "Start new job", .AutoSize = True}
    Private ReadOnly _resume As New System.Windows.Forms.Button With {.Text = "Resume last run", .AutoSize = True}
    Private ReadOnly _load As New System.Windows.Forms.Button With {.Text = "Resume saved run...", .AutoSize = True}
    Private ReadOnly _cancel As New System.Windows.Forms.Button With {.Text = "Cancel", .AutoSize = True, .Enabled = False}
    Private ReadOnly _open As New System.Windows.Forms.Button With {.Text = "Open results", .AutoSize = True, .Enabled = False}
    Private ReadOnly _errors As New System.Windows.Forms.Button With {.Text = "Error report", .AutoSize = True, .Enabled = False}
    Private ReadOnly _bar As New System.Windows.Forms.ProgressBar With {.Dock = System.Windows.Forms.DockStyle.Fill, .Minimum = 0, .Maximum = 1}
    Private ReadOnly _status As New System.Windows.Forms.Label With {.AutoSize = True, .Text = "Ready. Each file uses the existing agent harness."}
    Private ReadOnly _log As New System.Windows.Forms.TextBox With {
        .Multiline = True, .ReadOnly = True, .ScrollBars = System.Windows.Forms.ScrollBars.Vertical,
        .Dock = System.Windows.Forms.DockStyle.Fill, .WordWrap = True}
    Private _cts As System.Threading.CancellationTokenSource
    Private _busy As System.Boolean
    Private _closeWhenDone As System.Boolean
    Private _manifest As System.String = ""

    Public Sub New(host As ThisAddIn,
                   models As System.Collections.Generic.List(Of ThisAddIn.BatchModelChoice),
                   tools As System.Collections.Generic.List(Of Global.SharedLibrary.SharedLibrary.ModelConfig))
        _host = host
        _models = models
        _tools = tools
        _toolNames = host.GetBatchDefaultToolNames(tools)
        Me.Text = "Red Ink - Batch Agent Job"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.MinimumSize = New System.Drawing.Size(760, 640)
        Me.Size = New System.Drawing.Size(960, 850)
        Me.ShowInTaskbar = False
        _instruction.Multiline = True
        _instruction.AcceptsReturn = True
        _instruction.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        _instruction.Height = 108
        _instruction.MaxLength = 1000000
        _model.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        For Each choice As ThisAddIn.BatchModelChoice In models
            _model.Items.Add(choice)
        Next
        If _model.Items.Count > 0 Then _model.SelectedIndex = 0
        _output.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        _output.Items.AddRange(New System.Object() {"Desktop", "Source directory"})
        _output.SelectedIndex = 0
        _extensions.Text = New Global.SharedLibrary.Agents.BatchJobOptions().Extensions

        _configuration.ColumnCount = 2
        _configuration.AutoSize = True
        _configuration.Dock = System.Windows.Forms.DockStyle.Top
        _configuration.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 135.0F))
        _configuration.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
        Dim sourcePanel As New System.Windows.Forms.TableLayoutPanel With {.ColumnCount = 2, .AutoSize = True}
        sourcePanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
        sourcePanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize))
        _source.Dock = System.Windows.Forms.DockStyle.Fill
        Dim browse As New System.Windows.Forms.Button With {.Text = "Browse...", .AutoSize = True}
        AddHandler browse.Click, AddressOf BrowseSource
        sourcePanel.Controls.Add(_source, 0, 0)
        sourcePanel.Controls.Add(browse, 1, 0)
        AddSetting("Source directory", sourcePanel)
        AddSetting("Task / problem", _instruction)
        AddSetting("Tooling model", _model)
        AddSetting("Output location", _output)
        AddSetting("File extensions", _extensions)
        Dim switches As New System.Windows.Forms.FlowLayoutPanel With {.AutoSize = True, .WrapContents = True}
        switches.Controls.AddRange(New System.Windows.Forms.Control() {_recursive, _carry, _stop, _require})
        AddSetting("Processing", switches)
        Dim limits As New System.Windows.Forms.FlowLayoutPanel With {.AutoSize = True, .WrapContents = True}
        AddNumber(limits, "Retries / file", _retries)
        AddNumber(limits, "Delay (s)", _delay)
        AddNumber(limits, "Timeout (min)", _minutes)
        AddNumber(limits, "Tool turns", _iterations)
        AddNumber(limits, "Max input (MB)", _maxSize)
        AddSetting("Limits", limits)
        AddSetting("Tool scope", _selectTools)
        AddSetting("Display", _showDashboard)
        AddSetting("Notes", New System.Windows.Forms.Label With {.AutoSize = True, .MaximumSize = New System.Drawing.Size(720, 0),
            .Text = "Inputs are processed in path order using working copies. Carry-forward jobs stop at a failed item to preserve the chain. " &
                    "Resume keeps successful items. Selected tools keep their normal permissions; retries may repeat external side effects. Prefer file-only tasks."})

        Dim root As New System.Windows.Forms.TableLayoutPanel With {.Dock = System.Windows.Forms.DockStyle.Fill, .ColumnCount = 1, .RowCount = 5,
            .Padding = New System.Windows.Forms.Padding(12)}
        root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
        root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
        root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 25.0F))
        root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
        root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
        Dim configurationScroll As New System.Windows.Forms.Panel With {
            .AutoScroll = True, .Dock = System.Windows.Forms.DockStyle.Fill,
            .MinimumSize = New System.Drawing.Size(0, 250), .Height = 480}
        configurationScroll.Controls.Add(_configuration)
        root.RowStyles(0).SizeType = System.Windows.Forms.SizeType.Percent
        root.RowStyles(0).Height = 72.0F
        root.RowStyles(3).Height = 28.0F
        root.Controls.Add(configurationScroll, 0, 0)
        root.Controls.Add(_status, 0, 1)
        root.Controls.Add(_bar, 0, 2)
        root.Controls.Add(_log, 0, 3)
        Dim actions As New System.Windows.Forms.FlowLayoutPanel With {.AutoSize = True, .Dock = System.Windows.Forms.DockStyle.Fill, .WrapContents = True}
        actions.Controls.AddRange(New System.Windows.Forms.Control() {_start, _resume, _load, _cancel, _open, _errors})
        root.Controls.Add(actions, 0, 4)
        Me.Controls.Add(root)
        AddHandler _start.Click, AddressOf StartJob
        AddHandler _resume.Click, AddressOf ResumeLast
        AddHandler _load.Click, AddressOf ResumeSaved
        AddHandler _cancel.Click, Sub(sender, e) CancelJob()
        AddHandler _open.Click, Sub(sender, e) OpenResult(False)
        AddHandler _errors.Click, Sub(sender, e) OpenResult(True)
        AddHandler _selectTools.Click, AddressOf SelectTools
        AddHandler _carry.CheckedChanged, Sub(sender, e)
                                              _stop.Enabled = Not _carry.Checked
                                          End Sub
        UpdateToolCaption()
    End Sub

    Protected Overrides Sub OnShown(e As System.EventArgs)
        MyBase.OnShown(e)
        Global.SharedLibrary.SharedLibrary.SharedMethods.ForceDialogToForeground(Me)
        Global.SharedLibrary.SharedLibrary.SharedMethods.AttachForeignForegroundWatchdog(Me)
    End Sub

    Private Shared Function NumberBox(minimum As System.Decimal, maximum As System.Decimal, value As System.Decimal) As System.Windows.Forms.NumericUpDown
        Return New System.Windows.Forms.NumericUpDown With {.Minimum = minimum, .Maximum = maximum, .Value = value, .Width = 65}
    End Function

    Private Shared Sub AddNumber(panel As System.Windows.Forms.FlowLayoutPanel, caption As System.String, control As System.Windows.Forms.NumericUpDown)
        panel.Controls.Add(New System.Windows.Forms.Label With {.Text = caption, .AutoSize = True, .Margin = New System.Windows.Forms.Padding(3, 7, 3, 0)})
        panel.Controls.Add(control)
    End Sub

    Private Sub AddSetting(caption As System.String, control As System.Windows.Forms.Control)
        Dim row As System.Int32 = _configuration.RowCount
        _configuration.RowCount += 1
        _configuration.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
        _configuration.Controls.Add(New System.Windows.Forms.Label With {.Text = caption, .AutoSize = True, .Margin = New System.Windows.Forms.Padding(3, 7, 3, 3)}, 0, row)
        control.Dock = System.Windows.Forms.DockStyle.Fill
        _configuration.Controls.Add(control, 1, row)
    End Sub

    Private Sub BrowseSource(sender As System.Object, e As System.EventArgs)
        Using dialog As New System.Windows.Forms.FolderBrowserDialog With {.Description = "Select the source directory", .SelectedPath = _source.Text}
            Using Global.SharedLibrary.SharedLibrary.SharedMethods.PushDialogOwner(Me)
                If dialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then _source.Text = dialog.SelectedPath
            End Using
        End Using
    End Sub

    Private Sub SelectTools(sender As System.Object, e As System.EventArgs)
        ' Reuse the existing selector UI, but do not call ShowToolSelectionDialog: that
        ' convenience method persists selections shared with the user's Local Agent.
        Dim labels As New System.Collections.Generic.List(Of System.String)()
        For Each tool As Global.SharedLibrary.SharedLibrary.ModelConfig In _tools
            If _toolNames.Contains(tool.ToolName) Then labels.Add(If(System.String.IsNullOrWhiteSpace(tool.ModelDescription), tool.Model, tool.ModelDescription))
        Next
        Using selector As New Global.SharedLibrary.SharedLibrary.MultiModelSelectorForm(_tools, "", "Batch Agent Job - Tools", False,
            preselectMany:=labels, instruction:="Select the existing tools available to this batch. Other agent tool selections are not changed.")
            Using Global.SharedLibrary.SharedLibrary.SharedMethods.PushDialogOwner(Me)
                If selector.ShowDialog(Me) <> System.Windows.Forms.DialogResult.OK Then Return
                _toolNames = New System.Collections.Generic.List(Of System.String)()
                For Each tool As Global.SharedLibrary.SharedLibrary.ModelConfig In selector.SelectedModels
                    _toolNames.Add(tool.ToolName)
                Next
            End Using
        End Using
        UpdateToolCaption()
    End Sub

    Private Sub UpdateToolCaption()
        _selectTools.Text = "Tools... (" & _toolNames.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) & " selected)"
    End Sub

    Private Function ReadOptions() As Global.SharedLibrary.Agents.BatchJobOptions
        Dim selected As ThisAddIn.BatchModelChoice = TryCast(_model.SelectedItem, ThisAddIn.BatchModelChoice)
        If selected Is Nothing Then Throw New System.InvalidOperationException("Select a tooling-capable model.")
        Return New Global.SharedLibrary.Agents.BatchJobOptions With {
            .SourceDirectory = _source.Text.Trim(),
            .OutputBaseDirectory = If(_output.SelectedIndex = 1, _source.Text.Trim(), System.Environment.GetFolderPath(System.Environment.SpecialFolder.DesktopDirectory)),
            .Instruction = _instruction.Text, .ModelKey = selected.Key, .ModelDisplayName = selected.DisplayName,
            .ToolNames = New System.Collections.Generic.List(Of System.String)(_toolNames), .Extensions = _extensions.Text,
            .Recursive = _recursive.Checked, .CarryForward = _carry.Checked, .StopOnError = _stop.Checked,
            .RequireArtifact = _require.Checked, .ShowDashboard = _showDashboard.Checked, .MaxRetries = CInt(_retries.Value), .RetryDelaySeconds = CInt(_delay.Value),
            .AttemptTimeoutMinutes = CInt(_minutes.Value), .MaxToolIterations = CInt(_iterations.Value), .MaxInputMegabytes = CInt(_maxSize.Value)}
    End Function

    Private Async Sub StartJob(sender As System.Object, e As System.EventArgs)
        If _busy Then Return
        Try
            Await ExecuteAsync(ReadOptions(), Nothing)
        Catch ex As System.Exception
            ShowFailure(ex)
        End Try
    End Sub

    Private Async Sub ResumeLast(sender As System.Object, e As System.EventArgs)
        If _busy Then Return
        Try
            Dim pointer As System.String = ThisAddIn.BatchLastRunPointer()
            If Not System.IO.File.Exists(pointer) Then Throw New System.IO.FileNotFoundException("No previous batch is recorded. Use 'Resume saved run...' to select its batch-job.json.")
            Await ExecuteAsync(Nothing, System.IO.File.ReadAllText(pointer, System.Text.Encoding.UTF8).Trim())
        Catch ex As System.Exception
            ShowFailure(ex)
        End Try
    End Sub

    Private Async Sub ResumeSaved(sender As System.Object, e As System.EventArgs)
        If _busy Then Return
        Try
            Dim selectedPath As System.String
            Using dialog As New System.Windows.Forms.OpenFileDialog With {.Title = "Resume a batch manifest", .Filter = "Batch job (batch-job.json)|batch-job.json", .CheckFileExists = True}
                If dialog.ShowDialog(Me) <> System.Windows.Forms.DialogResult.OK Then Return
                selectedPath = dialog.FileName
            End Using
            Await ExecuteAsync(Nothing, selectedPath)
        Catch ex As System.Exception
            ShowFailure(ex)
        End Try
    End Sub

    Private Async Function ExecuteAsync(options As Global.SharedLibrary.Agents.BatchJobOptions, savedPath As System.String) As System.Threading.Tasks.Task
        Dim showDashboard As System.Boolean = _showDashboard.Checked
        SetBusy(True)
        _cts = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(_host.HostWindowLifetimeToken)
        Dim runCancellation As System.Threading.CancellationTokenSource = _cts
        Dim token As System.Threading.CancellationToken = runCancellation.Token
        Dim progress As System.IProgress(Of Global.SharedLibrary.Agents.BatchJobProgress) =
            New System.Progress(Of Global.SharedLibrary.Agents.BatchJobProgress)(AddressOf UpdateProgress)
        Try
            _log.Clear()
            _bar.Value = 0
            If options IsNot Nothing Then
                _manifest = Await System.Threading.Tasks.Task.Run(Function() Global.SharedLibrary.Agents.BatchJobQueue.Create(options, progress, token), token)
            Else
                _manifest = System.IO.Path.GetFullPath(savedPath)
                Dim saved = Await System.Threading.Tasks.Task.Run(Function() Global.SharedLibrary.Agents.BatchJobQueue.Load(_manifest), token)
                ApplyOptions(saved.Options)
                AppendLog("Resuming only failed, cancelled, interrupted and not-yet-run items with the saved prompt/settings.")
            End If
            AppendLog("Manifest: " & _manifest)
            Dim completed = Await System.Threading.Tasks.Task.Run(Function() _host.RunBatchJobAsync(_manifest, _models, _tools, progress, token, showDashboard, Sub() runCancellation.Cancel()), token)
            _status.Text = "Batch: " & completed.Status & " - " & System.IO.Path.GetDirectoryName(_manifest)
            AppendLog(_status.Text)
        Catch ex As System.OperationCanceledException
            _status.Text = "Cancelled. Completed results remain committed; resume continues unfinished items."
            AppendLog(_status.Text)
        Catch ex As System.Exception
            _status.Text = "Batch interrupted: " & ex.Message
            AppendLog(ex.ToString())
            Throw
        Finally
            _cts.Dispose()
            _cts = Nothing
            SetBusy(False)
            If _closeWhenDone Then Me.Close()
        End Try
    End Function

    Private Sub ApplyOptions(options As Global.SharedLibrary.Agents.BatchJobOptions)
        _source.Text = options.SourceDirectory
        _instruction.Text = options.Instruction
        _extensions.Text = options.Extensions
        _recursive.Checked = options.Recursive
        _carry.Checked = options.CarryForward
        _stop.Checked = options.StopOnError
        _require.Checked = options.RequireArtifact
        ' Display is a current UI choice and may change on resume; substantive
        ' task/model/tool options always come from the saved manifest.
        _retries.Value = options.MaxRetries
        _delay.Value = options.RetryDelaySeconds
        _minutes.Value = options.AttemptTimeoutMinutes
        _iterations.Value = options.MaxToolIterations
        _maxSize.Value = options.MaxInputMegabytes
        _output.SelectedIndex = If(System.String.Equals(options.SourceDirectory, options.OutputBaseDirectory, System.StringComparison.OrdinalIgnoreCase), 1, 0)
        _model.SelectedIndex = -1
        For index As System.Int32 = 0 To _models.Count - 1
            If _models(index).Key = options.ModelKey Then _model.SelectedIndex = index
        Next
        _toolNames = New System.Collections.Generic.List(Of System.String)(options.ToolNames)
        UpdateToolCaption()
    End Sub

    Private Sub UpdateProgress(value As Global.SharedLibrary.Agents.BatchJobProgress)
        If Me.IsDisposed OrElse value Is Nothing Then Return
        _bar.Maximum = System.Math.Max(1, value.Total)
        _bar.Value = System.Math.Min(_bar.Maximum, System.Math.Max(0, value.Completed))
        Dim prefix As System.String = If(value.ItemIndex > 0, value.ItemIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) & "/" & value.Total.ToString(System.Globalization.CultureInfo.InvariantCulture) &
            " | " & value.FileName & " | attempt " & value.Attempt.ToString(System.Globalization.CultureInfo.InvariantCulture) & " | ", "")
        _status.Text = prefix & value.Message
        AppendLog(_status.Text)
    End Sub

    Private Sub AppendLog(message As System.String)
        If Me.IsDisposed Then Return
        If _log.TextLength > 160000 Then _log.Text = _log.Text.Substring(_log.TextLength - 80000)
        _log.AppendText(System.DateTime.Now.ToString("HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture) & " " & message & System.Environment.NewLine)
    End Sub

    Private Sub SetBusy(value As System.Boolean)
        _busy = value
        _configuration.Enabled = Not value
        _start.Enabled = Not value
        _resume.Enabled = Not value
        _load.Enabled = Not value
        _cancel.Enabled = value
        _open.Enabled = Not value AndAlso System.IO.File.Exists(_manifest)
        _errors.Enabled = _open.Enabled
    End Sub

    Private Sub CancelJob()
        If _cts Is Nothing Then Return
        _cancel.Enabled = False
        _status.Text = "Cancelling the current harness run; waiting for its cleanup."
        _cts.Cancel()
    End Sub

    Private Sub ShowFailure(ex As System.Exception)
        If Me.IsDisposed Then Return
        Using Global.SharedLibrary.SharedLibrary.SharedMethods.PushDialogOwner(Me)
            Global.SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("Batch Agent Job: " & ex.Message)
        End Using
    End Sub

    Private Sub OpenResult(errors As System.Boolean)
        Try
            Dim path As System.String = System.IO.Path.GetDirectoryName(_manifest)
            If errors Then path = System.IO.Path.Combine(path, "errors.csv")
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(path) With {.UseShellExecute = True})
        Catch ex As System.Exception
            ShowFailure(ex)
        End Try
    End Sub

    Protected Overrides Sub OnFormClosing(e As System.Windows.Forms.FormClosingEventArgs)
        If _busy Then
            e.Cancel = True
            _closeWhenDone = True
            CancelJob()
        End If
        MyBase.OnFormClosing(e)
    End Sub
End Class
