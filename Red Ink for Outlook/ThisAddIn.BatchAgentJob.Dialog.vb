' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
Option Explicit On
Option Strict On
Option Infer On
Imports System.Linq

Namespace BatchJobs
    Friend NotInheritable Class JobDialog
        Inherits System.Windows.Forms.Form

        Private ReadOnly _host As ThisAddIn
        Private ReadOnly _options As Settings
        Private ReadOnly _choices As System.Collections.Generic.List(Of ModelChoice)
        Private ReadOnly _source As New System.Windows.Forms.TextBox()
        Private ReadOnly _prompt As New System.Windows.Forms.TextBox()
        Private ReadOnly _model As New System.Windows.Forms.ComboBox()
        Private ReadOnly _selectTools As System.Windows.Forms.Button = MakeButton("Tools / skills…")
        Private ReadOnly _useLocalChatTools As System.Windows.Forms.Button = MakeButton("Use Local Chat selection")
        Private ReadOnly _toolSelectionSummary As New System.Windows.Forms.Label()
        Private ReadOnly _destination As New System.Windows.Forms.ComboBox()
        Private ReadOnly _retries As New System.Windows.Forms.NumericUpDown()
        Private ReadOnly _iterations As New System.Windows.Forms.NumericUpDown()
        Private ReadOnly _extensions As New System.Windows.Forms.TextBox()
        Private ReadOnly _carry As New System.Windows.Forms.CheckBox()
        Private ReadOnly _consistent As New System.Windows.Forms.CheckBox()
        Private ReadOnly _recursive As New System.Windows.Forms.CheckBox()
        Private ReadOnly _stop As New System.Windows.Forms.CheckBox()
        Private ReadOnly _open As New System.Windows.Forms.CheckBox()
        Private ReadOnly _progress As New System.Windows.Forms.ProgressBar()
        Private ReadOnly _status As New System.Windows.Forms.Label()
        Private ReadOnly _log As New System.Windows.Forms.TextBox()
        Private ReadOnly _settingsPanel As New System.Windows.Forms.Panel()
        Private ReadOnly _start As System.Windows.Forms.Button = MakeButton("Start batch")
        Private ReadOnly _retry As System.Windows.Forms.Button = MakeButton("Retry / resume last job")
        Private ReadOnly _cancel As System.Windows.Forms.Button = MakeButton("Cancel job")
        Private ReadOnly _results As System.Windows.Forms.Button = MakeButton("Open Results")
        Private ReadOnly _report As System.Windows.Forms.Button = MakeButton("Open report")
        Private ReadOnly _close As System.Windows.Forms.Button = MakeButton("Close")
        Private _cts As System.Threading.CancellationTokenSource
        Private _busy As System.Boolean
        Private _resultsDirectory As System.String = ""

        Public Sub New(host As ThisAddIn, options As Settings, choices As System.Collections.Generic.List(Of ModelChoice))
            _host = host
            _options = options
            _choices = choices
            Text = "Red Ink — Batch Agent Job"
            Font = New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular)
            AutoScaleDimensions = New System.Drawing.SizeF(7.0F, 15.0F)
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
            MinimizeBox = False
            MaximizeBox = True
            ShowInTaskbar = False
            KeyPreview = True
            ClientSize = New System.Drawing.Size(1000, 960)
            MinimumSize = New System.Drawing.Size(800, 600)
            Global.SharedLibrary.SharedLibrary.SharedMethods.ApplyStandardDialogIcon(Me)
            Dim layout As New System.Windows.Forms.TableLayoutPanel With {.Dock = System.Windows.Forms.DockStyle.Fill, .Padding = New System.Windows.Forms.Padding(16), .ColumnCount = 1, .RowCount = 3}
            layout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 74.0F))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 26.0F))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            Controls.Add(layout)
            _settingsPanel.Dock = System.Windows.Forms.DockStyle.Fill
            _settingsPanel.AutoScroll = True
            layout.Controls.Add(_settingsPanel, 0, 0)
            Dim settingsTable As New System.Windows.Forms.TableLayoutPanel With {.Dock = System.Windows.Forms.DockStyle.Top, .AutoSize = True, .ColumnCount = 2, .Padding = New System.Windows.Forms.Padding(0, 0, 10, 10)}
            settingsTable.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 175.0F))
            settingsTable.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
            _settingsPanel.Controls.Add(settingsTable)
            Dim sourcePanel As New System.Windows.Forms.TableLayoutPanel With {.Dock = System.Windows.Forms.DockStyle.Fill, .AutoSize = True, .ColumnCount = 2, .Margin = New System.Windows.Forms.Padding(0)}
            sourcePanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
            sourcePanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize))
            _source.Dock = System.Windows.Forms.DockStyle.Fill
            _source.Anchor = System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right
            Dim browse = MakeButton("Browse…")
            AddHandler browse.Click, AddressOf BrowseSource
            sourcePanel.Controls.Add(_source, 0, 0)
            sourcePanel.Controls.Add(browse, 1, 0)
            AddRow(settingsTable, "Source directory", sourcePanel)
            _prompt.Multiline = True
            _prompt.AcceptsReturn = True
            _prompt.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
            _prompt.Height = 155
            _prompt.MaxLength = 100000
            _prompt.Dock = System.Windows.Forms.DockStyle.Fill
            AddRow(settingsTable, "Instruction for each file" & System.Environment.NewLine & "Ctrl+P: last used prompt", _prompt, 165.0F)
            _model.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            _model.Dock = System.Windows.Forms.DockStyle.Fill
            For Each choice In choices
                _model.Items.Add(choice)
            Next
            AddRow(settingsTable, "Tooling model", _model)
            Dim toolSelectionPanel As New System.Windows.Forms.TableLayoutPanel With {.Dock = System.Windows.Forms.DockStyle.Fill, .AutoSize = True, .ColumnCount = 1, .RowCount = 2, .Margin = New System.Windows.Forms.Padding(0)}
            toolSelectionPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
            toolSelectionPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            toolSelectionPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            Dim toolSelectionButtons As New System.Windows.Forms.FlowLayoutPanel With {.Dock = System.Windows.Forms.DockStyle.Fill, .AutoSize = True, .WrapContents = True, .Margin = New System.Windows.Forms.Padding(0)}
            toolSelectionButtons.Controls.AddRange(New System.Windows.Forms.Control() {_selectTools, _useLocalChatTools})
            toolSelectionPanel.Controls.Add(toolSelectionButtons, 0, 0)
            _toolSelectionSummary.AutoSize = True
            _toolSelectionSummary.Dock = System.Windows.Forms.DockStyle.Fill
            _toolSelectionSummary.Margin = New System.Windows.Forms.Padding(3, 5, 3, 3)
            toolSelectionPanel.Controls.Add(_toolSelectionSummary, 0, 1)
            AddRow(settingsTable, "Tools / skills", toolSelectionPanel)
            AddHandler _selectTools.Click, AddressOf SelectTools
            AddHandler _useLocalChatTools.Click, AddressOf UseLocalChatTools
            _destination.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            _destination.Items.AddRange(New System.Object() {"Desktop\Inky_Batch_Jobs", "Source directory\Inky_Batch_Jobs"})
            _destination.Dock = System.Windows.Forms.DockStyle.Fill
            AddRow(settingsTable, "Output location", _destination)
            Dim limits As New System.Windows.Forms.FlowLayoutPanel With {.Dock = System.Windows.Forms.DockStyle.Fill, .AutoSize = True, .WrapContents = True, .Margin = New System.Windows.Forms.Padding(0)}
            _retries.Minimum = 0D
            _retries.Maximum = 10D
            _retries.Width = 65
            _iterations.Minimum = 1D
            _iterations.Maximum = 1000D
            _iterations.Width = 80
            limits.Controls.Add(_retries)
            limits.Controls.Add(New System.Windows.Forms.Label With {.Text = "additional attempts per failed file; maximum tool iterations:", .AutoSize = True, .Margin = New System.Windows.Forms.Padding(8, 6, 5, 0)})
            limits.Controls.Add(_iterations)
            AddRow(settingsTable, "Retries / limits", limits)
            Dim switches As New System.Windows.Forms.TableLayoutPanel With {.Dock = System.Windows.Forms.DockStyle.Fill, .AutoSize = True, .ColumnCount = 2, .Margin = New System.Windows.Forms.Padding(0)}
            switches.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0F))
            switches.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0F))
            Dim boxes = New System.Windows.Forms.CheckBox() {_carry, _consistent, _recursive, _stop, _open}
            Dim captions = New System.String() {"Carry results into the next run", "Validate cumulative layout (XLSX / MD)", "Include subdirectories", "Stop after an unrecovered error", "Open Results when finished"}
            For index As System.Int32 = 0 To boxes.Length - 1
                boxes(index).Text = captions(index)
                boxes(index).AutoSize = True
                boxes(index).Margin = New System.Windows.Forms.Padding(3, 6, 8, 6)
                switches.Controls.Add(boxes(index), index Mod 2, index \ 2)
            Next
            AddRow(settingsTable, "Options", switches)
            _extensions.Dock = System.Windows.Forms.DockStyle.Fill
            AddRow(settingsTable, "Extensions (optional)", _extensions)
            Dim note As New System.Windows.Forms.Label With {.AutoSize = True, .Dock = System.Windows.Forms.DockStyle.Fill, .Text =
                "Empty extensions = all files; example: .txt;.docx;.pdf;.xlsx;.msg. Inputs are processed in filename order using working copies. " &
                "Cumulative mode stops at a failed item and resumes from the last verified result. Strict layout is append-only: fixed Excel columns/native row styles or an unchanged Markdown prefix. " &
                "Other formats remain available with strict layout disabled. New batches use the active Local Chat's tools, skills and agents by default. " &
                "Choose Tools / skills… for a separate batch selection; Local Chat keeps its own selection. Enabled memory still follows Local Chat. " &
                "The batch reads resource/template sources from the same configured resource roots and readable connected workspace, and edits working copies. " &
                "Retry / resume uses the recorded job's tool and skill selection. Select the tooling model above. Selected tools may access external services; retries can repeat their external actions.", .Margin = New System.Windows.Forms.Padding(3, 6, 3, 6)}
            AddRow(settingsTable, "Behaviour", note)
            AddHandler settingsTable.SizeChanged, Sub(sender, e)
                                                      Dim fieldWidth = System.Math.Max(180, settingsTable.ClientSize.Width - 195)
                                                      note.MaximumSize = New System.Drawing.Size(fieldWidth, 0)
                                                      _toolSelectionSummary.MaximumSize = New System.Drawing.Size(fieldWidth, 0)
                                                      For Each box In boxes
                                                          box.MaximumSize = New System.Drawing.Size(System.Math.Max(160, fieldWidth \ 2 - 18), 0)
                                                      Next
                                                  End Sub
            Dim progressPanel As New System.Windows.Forms.TableLayoutPanel With {.Dock = System.Windows.Forms.DockStyle.Fill, .ColumnCount = 1, .RowCount = 3, .Padding = New System.Windows.Forms.Padding(0, 6, 0, 6)}
            progressPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
            AddHandler progressPanel.SizeChanged, Sub(sender, e) _status.MaximumSize = New System.Drawing.Size(System.Math.Max(200, progressPanel.ClientSize.Width - 12), 0)
            progressPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 24.0F))
            progressPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            progressPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
            _progress.Dock = System.Windows.Forms.DockStyle.Fill
            _status.AutoSize = True
            _status.Text = "Ready. Open report shows completion, answers and any unfinished inputs."
            _status.Margin = New System.Windows.Forms.Padding(3, 6, 3, 6)
            _log.Multiline = True
            _log.ReadOnly = True
            _log.ScrollBars = System.Windows.Forms.ScrollBars.Both
            _log.WordWrap = False
            _log.Dock = System.Windows.Forms.DockStyle.Fill
            progressPanel.Controls.Add(_progress, 0, 0)
            progressPanel.Controls.Add(_status, 0, 1)
            progressPanel.Controls.Add(_log, 0, 2)
            layout.Controls.Add(progressPanel, 0, 1)
            Dim buttons As New System.Windows.Forms.FlowLayoutPanel With {.Dock = System.Windows.Forms.DockStyle.Fill, .AutoSize = True, .WrapContents = True, .Padding = New System.Windows.Forms.Padding(0, 8, 0, 0)}
            buttons.Controls.AddRange(New System.Windows.Forms.Control() {_start, _retry, _cancel, _results, _report, _close})
            layout.Controls.Add(buttons, 0, 2)
            AddHandler _start.Click, Async Sub(sender, e) Await RunAsync(False)
            AddHandler _retry.Click, Async Sub(sender, e) Await RunAsync(True)
            AddHandler _cancel.Click, Sub(sender, e)
                                          If _cts IsNot Nothing Then _cts.Cancel()
                                          _cancel.Enabled = False
                                          _status.Text = "Cancellation requested; waiting for the current tool to stop safely."
                                      End Sub
            AddHandler _results.Click, Sub(sender, e) OpenResults()
            AddHandler _report.Click, Sub(sender, e) OpenReport()
            AddHandler _close.Click, Sub(sender, e) Close()
            AddHandler _carry.CheckedChanged, Sub(sender, e) UpdateCarryMode()
            AddHandler FormClosing, AddressOf ClosingDialog
            AddHandler Shown, Sub(sender, e)
                                  ' Clamp after font/DPI autoscaling. The body scrolls on small
                                  ' screens; the progress/footer controls remain reachable.
                                  Dim area = System.Windows.Forms.Screen.FromControl(Me).WorkingArea
                                  MinimumSize = New System.Drawing.Size(System.Math.Min(MinimumSize.Width, area.Width - 20), System.Math.Min(MinimumSize.Height, area.Height - 20))
                                  Size = New System.Drawing.Size(System.Math.Min(Width, area.Width - 24), System.Math.Min(Height, area.Height - 24))
                                  Location = New System.Drawing.Point(area.Left + (area.Width - Width) \ 2, area.Top + (area.Height - Height) \ 2)
                              End Sub
            LoadOptions()
            SetBusy(False)
        End Sub

        Private Shared Function MakeButton(caption As System.String) As System.Windows.Forms.Button
            Return New System.Windows.Forms.Button With {.Text = caption, .AutoSize = True, .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink, .Padding = New System.Windows.Forms.Padding(14, 7, 14, 7), .Margin = New System.Windows.Forms.Padding(3, 3, 9, 3), .UseVisualStyleBackColor = True, .MinimumSize = New System.Drawing.Size(90, 36)}
        End Function

        Private Shared Sub AddRow(table As System.Windows.Forms.TableLayoutPanel, caption As System.String, control As System.Windows.Forms.Control, Optional height As System.Single = 0.0F)
            Dim row = table.RowCount
            table.RowCount += 1
            table.RowStyles.Add(If(height > 0.0F, New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, height), New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize)))
            table.Controls.Add(New System.Windows.Forms.Label With {.Text = caption, .AutoSize = True, .Margin = New System.Windows.Forms.Padding(3, 9, 10, 8), .MaximumSize = New System.Drawing.Size(165, 0)}, 0, row)
            control.Margin = New System.Windows.Forms.Padding(3, 5, 3, 5)
            table.Controls.Add(control, 1, row)
        End Sub

        Private Sub LoadOptions()
            _source.Text = _options.SourceDirectory
            _prompt.Text = _options.Prompt
            _extensions.Text = _options.Extensions
            _retries.Value = System.Math.Min(_retries.Maximum, System.Math.Max(_retries.Minimum, CDec(_options.Retries)))
            _iterations.Value = System.Math.Min(_iterations.Maximum, System.Math.Max(_iterations.Minimum, CDec(_options.MaxIterations)))
            _destination.SelectedIndex = If(_options.OutputOnDesktop, 0, 1)
            _carry.Checked = _options.CarryForward
            _consistent.Checked = _options.ConsistentLayout
            _recursive.Checked = _options.IncludeSubdirectories
            _stop.Checked = _options.StopOnError
            _open.Checked = _options.OpenResults
            _model.SelectedItem = _choices.FirstOrDefault(Function(choice) choice.Key.Equals(_options.ModelKey, System.StringComparison.OrdinalIgnoreCase))
            If _model.SelectedItem Is Nothing AndAlso System.String.IsNullOrWhiteSpace(_options.ModelKey) AndAlso _model.Items.Count > 0 Then _model.SelectedIndex = 0
            If System.IO.File.Exists(_options.LastManifestPath) Then _resultsDirectory = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(_options.LastManifestPath), "Results")
            UpdateCarryMode()
            UpdateToolSelectionSummary()
        End Sub

        Private Sub ReadOptions()
            _options.SourceDirectory = _source.Text.Trim()
            _options.Prompt = _prompt.Text
            _options.Extensions = _extensions.Text.Trim()
            Dim model = TryCast(_model.SelectedItem, ModelChoice)
            If model IsNot Nothing Then _options.ModelKey = model.Key
            _options.OutputOnDesktop = _destination.SelectedIndex = 0
            _options.Retries = CInt(_retries.Value)
            _options.MaxIterations = CInt(_iterations.Value)
            _options.CarryForward = _carry.Checked
            _options.ConsistentLayout = _consistent.Checked
            _options.IncludeSubdirectories = _recursive.Checked
            _options.StopOnError = _stop.Checked
            _options.OpenResults = _open.Checked
            _options.IncludeSelectedTools = True
        End Sub

        Private Sub UpdateToolSelectionSummary()
            If _options.UseLocalChatSelection Then
                _toolSelectionSummary.Text = "Use the active Local Chat selection when the batch starts. Tools / skills… lets you choose a separate selection for this batch."
            Else
                Dim mainCount As System.Int32 = If(_options.SelectedMainToolNames, New System.Collections.Generic.List(Of System.String)()).
                    Where(Function(name) Not System.String.IsNullOrWhiteSpace(name)).Select(Function(name) name.Trim()).Distinct(System.StringComparer.OrdinalIgnoreCase).Count()
                Dim advancedCount As System.Int32 = If(_options.SelectedAdvancedToolNames, New System.Collections.Generic.List(Of System.String)()).
                    Where(Function(name) Not System.String.IsNullOrWhiteSpace(name)).Select(Function(name) name.Trim()).Distinct(System.StringComparer.OrdinalIgnoreCase).Count()
                _toolSelectionSummary.Text = "Batch selection: " & mainCount.ToString(System.Globalization.CultureInfo.InvariantCulture) & " main tools / skills; " &
                    advancedCount.ToString(System.Globalization.CultureInfo.InvariantCulture) & " advanced tools. Local Chat keeps its own selection."
            End If
            _useLocalChatTools.Enabled = Not _options.UseLocalChatSelection
        End Sub

        Private Sub SelectTools(sender As System.Object, e As System.EventArgs)
            If _busy Then Return
            Using ownerScope = Global.SharedLibrary.SharedLibrary.SharedMethods.PushDialogOwner(Me)
                Try
                    If Not _host.SelectBatchAgentTools(_options) Then Return
                    UpdateToolSelectionSummary()
                    ReadOptions()
                    Store.SaveSettings(_options)
                Catch ex As System.Exception
                    Global.SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("Could not update the batch tool selection: " & ex.Message, "Red Ink — Batch Agent Job")
                End Try
            End Using
        End Sub

        Private Sub UseLocalChatTools(sender As System.Object, e As System.EventArgs)
            If _busy Then Return
            Using ownerScope = Global.SharedLibrary.SharedLibrary.SharedMethods.PushDialogOwner(Me)
                Try
                    _options.UseLocalChatSelection = True
                    _options.SelectedMainToolNames = New System.Collections.Generic.List(Of System.String)()
                    _options.SelectedAdvancedToolNames = New System.Collections.Generic.List(Of System.String)()
                    UpdateToolSelectionSummary()
                    ReadOptions()
                    Store.SaveSettings(_options)
                Catch ex As System.Exception
                    Global.SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("Could not save the Local Chat selection for this batch: " & ex.Message, "Red Ink — Batch Agent Job")
                End Try
            End Using
        End Sub

        Private Sub UpdateCarryMode()
            _consistent.Enabled = _carry.Checked
            _stop.Enabled = Not _carry.Checked
        End Sub

        Protected Overrides Function ProcessCmdKey(ByRef msg As System.Windows.Forms.Message, keyData As System.Windows.Forms.Keys) As System.Boolean
            If keyData = (System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.P) AndAlso Not _busy Then
                If Not System.String.IsNullOrWhiteSpace(_options.LastPrompt) Then
                    _prompt.Text = _options.LastPrompt
                    _prompt.Focus()
                    _prompt.SelectionStart = _prompt.TextLength
                End If
                Return True
            End If
            Return MyBase.ProcessCmdKey(msg, keyData)
        End Function

        Private Sub BrowseSource(sender As System.Object, e As System.EventArgs)
            Using ownerScope = Global.SharedLibrary.SharedLibrary.SharedMethods.PushDialogOwner(Me)
                Using picker As New System.Windows.Forms.FolderBrowserDialog With {.Description = "Select the source files for this batch", .ShowNewFolderButton = False, .SelectedPath = _source.Text}
                    If picker.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then _source.Text = picker.SelectedPath
                End Using
            End Using
        End Sub

        Private Async Function RunAsync(resumeLastJob As System.Boolean) As System.Threading.Tasks.Task
            If _busy Then Return
            Try
                ReadOptions()
                If resumeLastJob Then
                    If Not System.IO.File.Exists(_options.LastManifestPath) Then Throw New System.IO.FileNotFoundException("No previous batch checkpoint is available.")
                Else
                    If _model.SelectedItem Is Nothing Then Throw New System.IO.InvalidDataException("Select an available tooling-capable model.")
                    Store.ValidateOptions(_options)
                End If
                Store.SaveSettings(_options)
                _cts = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(_host.BatchAgentLifetimeToken)
                SetBusy(True)
                _log.Clear()
                _progress.Value = 0
                _progress.Style = System.Windows.Forms.ProgressBarStyle.Marquee
                _status.Text = If(resumeLastJob, "Verifying the recorded job and its saved tool / skill selection…",
                    If(_options.UseLocalChatSelection, "Preparing the source inventory and capturing the active Local Chat selection…", "Preparing the source inventory and the batch tool / skill selection…"))
                Dim sink As System.IProgress(Of ProgressInfo) = New System.Progress(Of ProgressInfo)(AddressOf ShowProgress)
                Dim result = Await System.Threading.Tasks.Task.Run(Async Function() Await _host.RunBatchAgentJobAsync(_options, resumeLastJob, _choices, sink, _cts.Token).ConfigureAwait(False))
                If IsDisposed OrElse Disposing Then Return
                _resultsDirectory = System.IO.Path.Combine(result.RootDirectory, "Results")
                _status.Text = System.Linq.Enumerable.Count(result.Items, Function(item) item.Status = "succeeded").ToString(System.Globalization.CultureInfo.InvariantCulture) & " / " & result.Items.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) & " files succeeded. Choose Open report for details."
                If Not System.String.IsNullOrWhiteSpace(result.ReportWarning) Then
                    _status.Text &= " The host could not save every report; see the log below."
                    _log.AppendText(result.ReportWarning & System.Environment.NewLine)
                End If
                If result.Options.OpenResults Then OpenResults()
            Catch ex As System.OperationCanceledException
                If IsDisposed OrElse Disposing Then Return
                _status.Text = "You cancelled the job. The host keeps verified results; Retry / resume continues unfinished files."
            Catch ex As System.Exception
                If IsDisposed OrElse Disposing Then Return
                _status.Text = "Stopped: " & ex.Message
                _log.AppendText(ex.ToString() & System.Environment.NewLine)
                Using ownerScope = Global.SharedLibrary.SharedLibrary.SharedMethods.PushDialogOwner(Me)
                    Global.SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox(ex.Message, "Red Ink — Batch Agent Job")
                End Using
            Finally
                If _cts IsNot Nothing Then _cts.Dispose()
                _cts = Nothing
                If Not IsDisposed AndAlso Not Disposing Then SetBusy(False)
            End Try
        End Function

        Private Sub SetBusy(value As System.Boolean)
            _busy = value
            If Not value Then _progress.Style = System.Windows.Forms.ProgressBarStyle.Blocks
            _settingsPanel.Enabled = Not value
            _start.Enabled = Not value AndAlso _choices.Count > 0
            _retry.Enabled = Not value AndAlso System.IO.File.Exists(_options.LastManifestPath)
            _cancel.Enabled = value
            _close.Enabled = Not value
            _results.Enabled = System.IO.Directory.Exists(_resultsDirectory)
            _report.Enabled = System.IO.File.Exists(ReportPath())
            UseWaitCursor = False
        End Sub

        Private Sub ShowProgress(value As ProgressInfo)
            If IsDisposed OrElse Disposing Then Return
            _progress.Style = System.Windows.Forms.ProgressBarStyle.Blocks
            _progress.Maximum = System.Math.Max(1, value.Total)
            _progress.Value = System.Math.Min(_progress.Maximum, System.Math.Max(0, value.Completed))
            _status.Text = value.Completed.ToString(System.Globalization.CultureInfo.InvariantCulture) & " / " & value.Total.ToString(System.Globalization.CultureInfo.InvariantCulture) & " processed; " & value.Failed.ToString(System.Globalization.CultureInfo.InvariantCulture) & " failed. " & value.Current
            _resultsDirectory = System.IO.Path.Combine(value.RootDirectory, "Results")
            _results.Enabled = System.IO.Directory.Exists(_resultsDirectory)
            _report.Enabled = System.IO.File.Exists(ReportPath())
            If _log.TextLength > 60000 Then _log.Text = _log.Text.Substring(_log.TextLength - 30000)
            _log.AppendText(value.Message & System.Environment.NewLine)
        End Sub

        Private Function ReportPath() As System.String
            If System.String.IsNullOrWhiteSpace(_options.LastManifestPath) Then Return ""
            Return System.IO.Path.Combine(System.IO.Path.GetDirectoryName(_options.LastManifestPath), "Run-report.md")
        End Function

        Private Sub OpenReport()
            Try
                Dim path = ReportPath()
                If Not System.IO.File.Exists(path) Then
                    _log.AppendText("The host has no Markdown report yet. Read the progress log for the current failure." & System.Environment.NewLine)
                    Return
                End If
                System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo With {.FileName = path, .UseShellExecute = True})
            Catch ex As System.Exception
                _log.AppendText("Cannot open the report: " & ex.Message & System.Environment.NewLine)
            End Try
        End Sub

        Private Sub OpenResults()
            Try
                If Not System.IO.Directory.Exists(_resultsDirectory) Then Return
                System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo With {.FileName = _resultsDirectory, .UseShellExecute = True})
            Catch ex As System.Exception
                _log.AppendText("Cannot open Results: " & ex.Message & System.Environment.NewLine)
            End Try
        End Sub

        Private Sub ClosingDialog(sender As System.Object, e As System.Windows.Forms.FormClosingEventArgs)
            If _busy Then
                If _cts IsNot Nothing Then _cts.Cancel()
                ' Do not block Office/Windows shutdown. Ordinary user Close waits for
                ' the active operation; queued progress ignores a disposed dialog.
                e.Cancel = e.CloseReason = System.Windows.Forms.CloseReason.UserClosing
                If e.Cancel Then _status.Text = "Cancellation requested. Close after the active tool has stopped."
                Return
            End If
            Try
                ReadOptions()
                Store.SaveSettings(_options)
            Catch ex As System.Exception
                e.Cancel = True
                Using ownerScope = Global.SharedLibrary.SharedLibrary.SharedMethods.PushDialogOwner(Me)
                    Global.SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("Settings could not be saved: " & ex.Message, "Red Ink — Batch Agent Job")
                End Using
            End Try
        End Sub
    End Class
End Namespace
