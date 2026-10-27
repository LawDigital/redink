' Part of Red Ink for Excel. Independent Collection Analyzer UI adapter.
' Existing Data Extractor, worksheet output and Outlook export are untouched.
Option Strict On
Option Explicit On

Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks
Imports SharedLibrary.Agents
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports Newtonsoft.Json

Partial Public Class ThisAddIn
    Private _collectionAnalyzerOpen As Boolean

    Public Sub OpenCollectionAnalyzer()
        If _collectionAnalyzerOpen Then
            Global.SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("A Collection Analyzer window is already open.")
            Return
        End If
        _collectionAnalyzerOpen = True
        Try
            Using dialog As New CollectionAnalyzerDialog(_context)
                Dim owner As System.Windows.Forms.IWin32Window =
                    Global.SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
                If owner IsNot Nothing Then
                    dialog.ShowDialog(owner)
                Else
                    dialog.ShowDialog()
                End If
            End Using
        Finally
            _collectionAnalyzerOpen = False
        End Try
    End Sub

End Class

Friend NotInheritable Class CollectionAnalyzerDialog
    Inherits System.Windows.Forms.Form

    Private ReadOnly _context As ISharedContext
    Private ReadOnly _modelChoice As New System.Windows.Forms.ComboBox()
    Private ReadOnly _choices As New System.Collections.Generic.List(Of ModelChoice)()
    Private _missingSavedModel As String = String.Empty
    Private ReadOnly _source As New System.Windows.Forms.TextBox()
    Private ReadOnly _output As New System.Windows.Forms.TextBox()
    Private ReadOnly _objective As New System.Windows.Forms.TextBox()
    Private ReadOnly _duplicates As New System.Windows.Forms.CheckBox()
    Private ReadOnly _annex As New System.Windows.Forms.CheckBox()
    Private ReadOnly _progress As New System.Windows.Forms.TextBox()
    Private ReadOnly _run As New System.Windows.Forms.Button()
    Private ReadOnly _cancel As New System.Windows.Forms.Button()
    Private _cts As System.Threading.CancellationTokenSource
    Private _running As Boolean

    ' Stores preferences only, never API keys or model configurations.
    Private NotInheritable Class CollectionAnalyzerPreferences
        Public Sub New()
        End Sub

        Public Property SourceDirectory As String = String.Empty
        Public Property WorkspaceDirectory As String = String.Empty
        Public Property Objective As String = String.Empty
        Public Property ModelKey As String = String.Empty
        Public Property RemoveExactDuplicates As Boolean = False
        Public Property IncludeEvidenceAppendix As Boolean = True
        Public Property WindowWidth As Integer = 825
        Public Property WindowHeight As Integer = 645
    End Class

    Private NotInheritable Class ModelChoice
        Public Property Key As String
        Public Property Display As String
        Public Property PrimaryModelNumber As Integer
        Public Property UseSecondary As Boolean
        Public Property Alternate As ModelConfig
        Public Overrides Function ToString() As String
            Return Display
        End Function
    End Class

    Private Shared ReadOnly Property SettingsPath As String
        Get
            Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                                "RedInk", "CollectionAnalyzer", "settings.json")
        End Get
    End Property

    Private Shared ReadOnly Property LastPromptPath As String
        Get
            Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                                "RedInk", "CollectionAnalyzer", "last-objective.txt")
        End Get
    End Property

    Public Sub New(context As ISharedContext)
        If context Is Nothing Then Throw New System.ArgumentNullException(NameOf(context))
        _context = context
        Me.Text = "Red Ink – Collection Analyzer"
        Me.MinimumSize = New System.Drawing.Size(660, 540)
        Me.Size = New System.Drawing.Size(825, 645)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        ' Use the same Red Ink logo as the existing parameter dialogs.
        Dim logoBitmap As New System.Drawing.Bitmap(Global.SharedLibrary.SharedLibrary.SharedMethods.GetLogoBitmap(
            Global.SharedLibrary.SharedLibrary.SharedMethods.LogoType.Standard))
        Me.Icon = System.Drawing.Icon.FromHandle(logoBitmap.GetHicon())

        Dim panel As New System.Windows.Forms.TableLayoutPanel With {
            .Dock = System.Windows.Forms.DockStyle.Fill, .ColumnCount = 3, .RowCount = 9,
            .Padding = New System.Windows.Forms.Padding(12)
        }
        panel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 150))
        panel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100))
        panel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 100))
        panel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 38))
        panel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 38))
        panel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 38))
        panel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28))
        panel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 48))
        panel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 36))
        panel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35))
        panel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 52))
        panel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 48))
        Me.Controls.Add(panel)

        _source.Dock = System.Windows.Forms.DockStyle.Fill
        panel.Controls.Add(NewLabel("Source folder:"), 0, 0)
        panel.Controls.Add(_source, 1, 0)
        Dim browseSource As New System.Windows.Forms.Button With {.Text = "Browse...", .Dock = System.Windows.Forms.DockStyle.Fill}
        panel.Controls.Add(browseSource, 2, 0)
        AddHandler browseSource.Click, Sub(sender, args) ChooseFolder(_source)

        _output.Text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "Red Ink Collection Analysis")
        _output.Dock = System.Windows.Forms.DockStyle.Fill
        panel.Controls.Add(NewLabel("Output folder:"), 0, 1)
        panel.Controls.Add(_output, 1, 1)
        Dim browseOutput As New System.Windows.Forms.Button With {.Text = "Browse...", .Dock = System.Windows.Forms.DockStyle.Fill}
        panel.Controls.Add(browseOutput, 2, 1)
        AddHandler browseOutput.Click, Sub(sender, args) ChooseFolder(_output)

        panel.Controls.Add(NewLabel("Model:"), 0, 2)
        _modelChoice.Dock = System.Windows.Forms.DockStyle.Fill
        _modelChoice.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        panel.Controls.Add(_modelChoice, 1, 2)
        panel.SetColumnSpan(_modelChoice, 2)

        Dim hint As New System.Windows.Forms.Label With {
            .Text = "Analysis objective (Ctrl+P restores the last submitted objective):", .Dock = System.Windows.Forms.DockStyle.Fill,
            .TextAlign = System.Drawing.ContentAlignment.MiddleLeft, .AutoSize = False
        }
        panel.Controls.Add(hint, 0, 3)
        panel.SetColumnSpan(hint, 3)
        _objective.Multiline = True
        _objective.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        _objective.AcceptsReturn = True
        _objective.Dock = System.Windows.Forms.DockStyle.Fill
        panel.Controls.Add(_objective, 0, 4)
        panel.SetColumnSpan(_objective, 3)
        AddHandler _objective.KeyDown, AddressOf ObjectiveKeyDown

        _duplicates.Text = "Skip exact duplicate source contents (retain all filenames)"
        _duplicates.Checked = False
        _duplicates.AutoSize = True
        panel.Controls.Add(_duplicates, 0, 5)
        panel.SetColumnSpan(_duplicates, 3)
        _annex.Text = "Include full validated evidence annex in the Markdown report"
        _annex.Checked = True
        _annex.AutoSize = True
        panel.Controls.Add(_annex, 0, 6)
        panel.SetColumnSpan(_annex, 3)

        _progress.Multiline = True
        _progress.ReadOnly = True
        _progress.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        _progress.Dock = System.Windows.Forms.DockStyle.Fill
        _progress.Text = "The analyzer reads existing text exports. Original files are never modified." & Environment.NewLine &
                         "On Start, choose a NEW analysis or explicitly select an unfinished matching analysis." & Environment.NewLine &
                         "Completed validated batches are rechecked and reused when resuming, including after repeated interruptions." & Environment.NewLine &
                         "Failed batches and unavailable files are reported; partial results are clearly marked. Keep the recovery folder."
        panel.Controls.Add(_progress, 0, 7)
        panel.SetColumnSpan(_progress, 3)
        _run.Text = "Start analysis..."
        _run.Dock = System.Windows.Forms.DockStyle.Fill
        _cancel.Text = "Close"
        _cancel.Dock = System.Windows.Forms.DockStyle.Fill
        panel.Controls.Add(_run, 1, 8)
        panel.Controls.Add(_cancel, 2, 8)
        AddHandler _run.Click, AddressOf StartClicked
        AddHandler _cancel.Click, AddressOf CancelClicked
        AddHandler Me.FormClosing, AddressOf ClosingDialog
        PopulateModels()
        RestorePreferences()
        AddHandler _modelChoice.SelectedIndexChanged,
            Sub(sender, args)
                If _modelChoice.SelectedIndex >= 0 Then _missingSavedModel = String.Empty
            End Sub
    End Sub

    Private Sub PopulateModels()
        Dim primary As System.Collections.Generic.List(Of Integer) = PrimaryModelManager.GetAvailableModels()
        For Each number As Integer In primary
            _choices.Add(New ModelChoice With {
                .Key = "primary:" & number.ToString(System.Globalization.CultureInfo.InvariantCulture),
                .Display = "Excel model " & number.ToString() & " – " & PrimaryModelManager.GetModelDisplayName(number),
                .PrimaryModelNumber = number
            })
        Next
        ' The current primary model is always available even in a single-model configuration.
        If _choices.Count = 0 Then
            _choices.Add(New ModelChoice With {
                .Key = "current", .Display = "Current Excel model – " & If(_context.INI_Model, "")
            })
        End If
        If _context.INI_SecondAPI AndAlso Not System.String.IsNullOrWhiteSpace(_context.INI_Model_2) Then
            _choices.Add(New ModelChoice With {
                .Key = "secondary", .Display = "Secondary model – " & _context.INI_Model_2,
                .UseSecondary = True
            })
        End If
        If Not System.String.IsNullOrWhiteSpace(_context.INI_AlternateModelPath) Then
            Try
                Dim path As String = SharedMethods.ExpandEnvironmentVariables(_context.INI_AlternateModelPath)
                Dim loadContext As ISharedContext = SharedMethods.CreateIsolatedModelCallContext(_context)
                For Each config As ModelConfig In SharedMethods.LoadAlternativeModels(path, loadContext)
                    Dim identity As String = String.Join(vbLf, New String() {
                        If(config.ModelDescription, ""), If(config.Model, ""),
                        If(config.Endpoint, ""), If(config.APICall, ""),
                        If(config.Response, "")
                    })
                    Dim key As String = "alternate:" & Digest(identity)
                    If Not _choices.Exists(Function(c) c.Key = key) Then
                        _choices.Add(New ModelChoice With {
                            .Key = key,
                            .Display = "Alternate model – " & If(Not String.IsNullOrWhiteSpace(config.ModelDescription), config.ModelDescription, config.Model),
                            .UseSecondary = True, .Alternate = config
                        })
                    End If
                Next
            Catch ex As System.Exception
                AppendStatus("Alternate models could not be loaded: " & ex.Message)
            End Try
        End If
        For Each choice As ModelChoice In _choices
            _modelChoice.Items.Add(choice)
        Next
        Dim defaultKey As String = "primary:" & PrimaryModelManager.GetCurrentModelNumber().ToString(System.Globalization.CultureInfo.InvariantCulture)
        Dim defaultIndex As Integer = _choices.FindIndex(Function(c) c.Key = defaultKey)
        _modelChoice.SelectedIndex = If(defaultIndex >= 0, defaultIndex, 0)
    End Sub

    Private Shared Function NormalizeOutputFolder(selected As String) As String
        Dim value As String = System.IO.Path.GetFullPath(selected.Trim())
        Dim root As String = System.IO.Path.GetPathRoot(value)
        If Not System.String.Equals(value, root, System.StringComparison.OrdinalIgnoreCase) Then
            value = value.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar)
        End If
        Dim lastPart As String = System.IO.Path.GetFileName(value)
        If System.String.Equals(lastPart, "Red Ink Collection Analysis", System.StringComparison.OrdinalIgnoreCase) Then Return value
        If System.String.Equals(lastPart, "RedInk Collection Analysis", System.StringComparison.OrdinalIgnoreCase) Then
            If System.IO.Directory.Exists(value) Then
                Try
                    For Each previousJob As String In System.IO.Directory.EnumerateDirectories(value, "job_*", System.IO.SearchOption.TopDirectoryOnly)
                        If System.IO.File.Exists(System.IO.Path.Combine(previousJob, "manifest.json")) AndAlso
                           Not System.IO.File.Exists(System.IO.Path.Combine(previousJob, "Result.md")) Then
                            ' Older incomplete runs must remain reachable for recovery.
                            Return value
                        End If
                    Next
                Catch ex As System.Exception When TypeOf ex Is System.IO.IOException OrElse
                                                  TypeOf ex Is System.UnauthorizedAccessException
                    Return value
                End Try
            End If
            Return System.IO.Path.Combine(System.IO.Path.GetDirectoryName(value), "Red Ink Collection Analysis")
        End If
        Return System.IO.Path.Combine(value, "Red Ink Collection Analysis")
    End Function

    Private Sub RestorePreferences()
        If Not File.Exists(SettingsPath) Then Return
        Try
            Dim saved As CollectionAnalyzerPreferences = JsonConvert.DeserializeObject(Of CollectionAnalyzerPreferences)(
                File.ReadAllText(SettingsPath, System.Text.Encoding.UTF8))
            If saved Is Nothing Then Throw New System.IO.InvalidDataException("Empty settings file.")
            _source.Text = If(saved.SourceDirectory, String.Empty)
            If Not System.String.IsNullOrWhiteSpace(saved.WorkspaceDirectory) Then
                _output.Text = NormalizeOutputFolder(saved.WorkspaceDirectory)
            End If
            _objective.Text = If(saved.Objective, String.Empty)
            _duplicates.Checked = saved.RemoveExactDuplicates
            _annex.Checked = saved.IncludeEvidenceAppendix
            If saved.WindowWidth >= Me.MinimumSize.Width AndAlso saved.WindowHeight >= Me.MinimumSize.Height Then
                Me.Size = New System.Drawing.Size(saved.WindowWidth, saved.WindowHeight)
            End If
            If Not System.String.IsNullOrEmpty(saved.ModelKey) Then
                Dim index As Integer = _choices.FindIndex(Function(c) c.Key = saved.ModelKey)
                If index >= 0 Then
                    _modelChoice.SelectedIndex = index
                Else
                    _missingSavedModel = saved.ModelKey
                    _modelChoice.SelectedIndex = -1
                    AppendStatus("The previously selected model is no longer available. Select another model before starting.")
                End If
            End If
        Catch ex As System.Exception
            AppendStatus("Saved Collection Analyzer preferences could not be read: " & ex.Message)
        End Try
    End Sub

    Private Sub SavePreferences()
        Dim selected As ModelChoice = TryCast(_modelChoice.SelectedItem, ModelChoice)
        Dim prefs As New CollectionAnalyzerPreferences With {
            .SourceDirectory = _source.Text, .WorkspaceDirectory = _output.Text,
            .Objective = _objective.Text,
            .ModelKey = If(selected IsNot Nothing, selected.Key, _missingSavedModel),
            .RemoveExactDuplicates = _duplicates.Checked,
            .IncludeEvidenceAppendix = _annex.Checked,
            .WindowWidth = Me.Width, .WindowHeight = Me.Height
        }
        Dim directoryName As String = Path.GetDirectoryName(SettingsPath)
        Directory.CreateDirectory(directoryName)
        Dim temp As String = Path.Combine(directoryName, "settings." & System.Guid.NewGuid().ToString("N") & ".tmp")
        Try
            File.WriteAllText(temp, JsonConvert.SerializeObject(prefs, Formatting.Indented),
                              New System.Text.UTF8Encoding(False))
            If File.Exists(SettingsPath) Then
                File.Replace(temp, SettingsPath, Nothing)
            Else
                File.Move(temp, SettingsPath)
            End If
        Finally
            If File.Exists(temp) Then File.Delete(temp)
        End Try
    End Sub

    Private Shared Function Digest(value As String) As String
        Using algorithm As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
            Return System.BitConverter.ToString(algorithm.ComputeHash(System.Text.Encoding.UTF8.GetBytes(value))).Replace("-", "").ToLowerInvariant()
        End Using
    End Function

    Private Function CreateModelRun(choice As ModelChoice, cancellation As System.Threading.CancellationToken) As (InvokeModel As System.Func(Of String, String, System.Threading.Tasks.Task(Of String)), Signature As String)
        Dim isolated As ISharedContext
        If choice.PrimaryModelNumber > 0 Then
            isolated = PrimaryModelManager.CreateIsolatedPrimaryModelContext(_context, choice.PrimaryModelNumber)
        Else
            isolated = SharedMethods.CreateIsolatedModelCallContext(_context)
        End If
        If choice.Alternate IsNot Nothing Then SharedMethods.ApplyModelConfig(isolated, choice.Alternate)
        Dim secondary As Boolean = choice.UseSecondary
        Dim fingerprint As String = String.Join(vbLf, New String() {
            choice.Key,
            If(secondary, isolated.INI_Model_2, isolated.INI_Model),
            If(secondary, isolated.INI_Endpoint_2, isolated.INI_Endpoint),
            If(secondary, isolated.INI_HeaderA_2, isolated.INI_HeaderA),
            If(secondary, isolated.INI_HeaderB_2, isolated.INI_HeaderB),
            If(secondary, isolated.INI_APICall_2, isolated.INI_APICall),
            If(secondary, isolated.INI_APICall_Object_2, isolated.INI_APICall_Object),
            If(secondary, isolated.INI_Response_2, isolated.INI_Response),
            If(secondary, isolated.INI_Temperature_2, isolated.INI_Temperature),
            If(secondary, isolated.INI_MaxOutputToken_2, isolated.INI_MaxOutputToken).ToString(System.Globalization.CultureInfo.InvariantCulture),
            If(secondary, isolated.INI_Timeout_2, isolated.INI_Timeout).ToString(System.Globalization.CultureInfo.InvariantCulture),
            isolated.INI_Model_Parameter1, isolated.INI_Model_Parameter2,
            isolated.INI_Model_Parameter3, isolated.INI_Model_Parameter4
        })
        Dim signature As String = Digest(fingerprint)
        Dim modelInvoker As System.Func(Of String, String, System.Threading.Tasks.Task(Of String)) =
            Function(systemPrompt As String, userPrompt As String)
                Return SharedMethods.LLM(isolated, systemPrompt, userPrompt,
                                         UseSecondAPI:=secondary, Hidesplash:=True, cancellationToken:=cancellation)
            End Function
        Return (modelInvoker, signature)
    End Function

    ' Shared host UI: logo, fonts and same-thread owner handling stay consistent
    ' with Excel Data Extractor and other Red Ink parameter dialogs.
    Private Sub ShowRedInkMessage(message As String)
        Using Global.SharedLibrary.SharedLibrary.SharedMethods.PushDialogOwner(Me)
            Global.SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox(message, Me.Text)
        End Using
    End Sub

    Private Function ConfirmCloseWithoutSaving(message As String) As Boolean
        Using Global.SharedLibrary.SharedLibrary.SharedMethods.PushDialogOwner(Me)
            Return Global.SharedLibrary.SharedLibrary.SharedMethods.ShowCustomYesNoBox(
                message, "Close without saving", "Return to analyzer", Me.Text) = 1
        End Using
    End Function

    Private Shared Function NewLabel(label As String) As System.Windows.Forms.Label
        Return New System.Windows.Forms.Label With {
            .Text = label, .TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
            .Dock = System.Windows.Forms.DockStyle.Fill, .AutoSize = False
        }
    End Function

    Private Sub ObjectiveKeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs)
        If e.Control AndAlso e.KeyCode = System.Windows.Forms.Keys.P Then
            e.SuppressKeyPress = True
            e.Handled = True
            Try
                If File.Exists(LastPromptPath) Then
                    _objective.Text = File.ReadAllText(LastPromptPath, System.Text.Encoding.UTF8)
                    _objective.SelectionStart = _objective.TextLength
                End If
            Catch ex As System.Exception
                AppendStatus("Unable to restore the previous objective: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub ChooseFolder(target As System.Windows.Forms.TextBox)
        Using chooser As New System.Windows.Forms.FolderBrowserDialog With {.Description = "Choose a folder"}
            If Directory.Exists(target.Text) Then chooser.SelectedPath = target.Text
            If chooser.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
                target.Text = If(System.Object.ReferenceEquals(target, _output),
                                 NormalizeOutputFolder(chooser.SelectedPath), chooser.SelectedPath)
            End If
        End Using
    End Sub

    Private Function ChooseAnalysisRun(runs As System.Collections.Generic.List(Of CollectionAnalysisRunInfo)) As String
        If runs.Count = 0 Then Return String.Empty
        Dim items As New System.Collections.Generic.List(Of Global.SharedLibrary.SharedLibrary.SharedMethods.SelectionItem) From {
            New Global.SharedLibrary.SharedLibrary.SharedMethods.SelectionItem("Start a NEW analysis (do not reuse earlier results)", 1)
        }
        For index As Integer = 0 To runs.Count - 1
            items.Add(New Global.SharedLibrary.SharedLibrary.SharedMethods.SelectionItem(
                "Resume " & runs(index).DisplayLabel, index + 2))
        Next
        Dim selection As Integer = Global.SharedLibrary.SharedLibrary.SharedMethods.SelectValue(
            items, 1,
            "Select an earlier analysis or start a new one. Changing models does not hide prior runs: " &
            "verified findings will be reused, and any new model calls will use the currently selected model. " &
            "Cancel leaves all files unchanged.",
            "Red Ink – Collection Analyzer", Me)
        If selection = 0 Then Return Nothing
        If selection = 1 Then Return String.Empty
        Return runs(selection - 2).DirectoryPath
    End Function

    Private Async Sub StartClicked(sender As Object, e As EventArgs)
        If _running Then Return
        If Not Directory.Exists(_source.Text.Trim()) Then
            ShowRedInkMessage("Select an existing source directory.")
            Return
        End If
        If String.IsNullOrWhiteSpace(_objective.Text) Then
            ShowRedInkMessage("Enter an analysis objective.")
            Return
        End If
        Dim selectedModel As ModelChoice = TryCast(_modelChoice.SelectedItem, ModelChoice)
        If selectedModel Is Nothing Then
            ShowRedInkMessage("Select an available model. The previously selected model may no longer be configured.")
            Return
        End If
        Dim root As String = _output.Text.Trim()
        If String.IsNullOrWhiteSpace(root) Then
            ShowRedInkMessage("Select a work directory.")
            Return
        End If
        Try
            root = NormalizeOutputFolder(root)
            _output.Text = root
            SavePreferences()
        Catch ex As System.Exception
            AppendStatus("Unable to save preferences: " & ex.Message)
            ShowRedInkMessage("The analysis preferences could not be saved. Check access to your profile folder and try again.")
            Return
        End Try
        Try
            Directory.CreateDirectory(Path.GetDirectoryName(LastPromptPath))
            File.WriteAllText(LastPromptPath, _objective.Text, New System.Text.UTF8Encoding(False))
        Catch ex As System.Exception
            AppendStatus("Unable to save objective: " & ex.Message)
            ShowRedInkMessage("The last analysis objective could not be saved for Ctrl+P. Check access to your profile folder and try again.")
            Return
        End Try
        _cts = New System.Threading.CancellationTokenSource()
        _running = True
        _run.Enabled = False
        _cancel.Text = "Cancel"
        _source.Enabled = False
        _output.Enabled = False
        _objective.Enabled = False
        _duplicates.Enabled = False
        _annex.Enabled = False
        _modelChoice.Enabled = False
        Try
            Dim modelRun = CreateModelRun(selectedModel, _cts.Token)
            Dim options As New CollectionAnalysisOptions With {
                .SourceDirectory = _source.Text.Trim(), .WorkspaceDirectory = root,
                .Objective = _objective.Text.Trim(), .ModelSignature = modelRun.Signature,
                .RemoveExactDuplicates = _duplicates.Checked,
                .IncludeEvidenceAppendix = _annex.Checked
            }
            AppendStatus("Checking for earlier compatible analyses...")
            Dim previousRuns As System.Collections.Generic.List(Of CollectionAnalysisRunInfo) =
                Await System.Threading.Tasks.Task.Run(Function() CollectionAnalysisService.ListResumableRuns(options, _cts.Token), _cts.Token)
            Dim selectedRun As String = ChooseAnalysisRun(previousRuns)
            If selectedRun Is Nothing Then
                AppendStatus("No analysis started. Selection cancelled.")
                Return
            End If
            options.ResumeDirectory = selectedRun
            If String.IsNullOrEmpty(selectedRun) Then
                AppendStatus("New analysis selected. Existing checkpoints will not be reused.")
            Else
                AppendStatus("Resuming selected analysis: " & System.IO.Path.GetFileName(selectedRun))
                Dim chosenInfo As CollectionAnalysisRunInfo = previousRuns.Find(
                    Function(run) String.Equals(run.DirectoryPath, selectedRun, System.StringComparison.OrdinalIgnoreCase))
                If chosenInfo IsNot Nothing AndAlso chosenInfo.ModelChanged Then
                    AppendStatus("Model changed: verified checkpoints from the earlier model will be reused. " &
                                 "Remaining processing will use the currently selected model.")
                End If
            End If
            Dim reporter As New Progress(Of String)(Sub(message) AppendStatus(message))
            Dim result As CollectionAnalysisResult = Await CollectionAnalysisService.RunAsync(options, modelRun.InvokeModel, _cts.Token, reporter)
            AppendStatus("Report: " & result.ReportPath)
            If result.IsComplete Then
                _output.Text = NormalizeOutputFolder(_output.Text)
                ShowRedInkMessage("The collection analysis is complete. The report was saved to:" & Environment.NewLine & result.ReportPath)
            Else
                ShowRedInkMessage("A partial report was saved. " & result.UnreadableSourceCount.ToString() &
                    " source(s) were unreadable and " & result.FailedPackageCount.ToString() &
                    " analysis package(s) were not accepted. Open the report for details. " &
                    "Fix the source issues if needed, then start again and select this same unfinished run to retry the missing work." &
                    Environment.NewLine & result.ReportPath)
            End If
        Catch ex As System.OperationCanceledException
            AppendStatus("Analysis cancelled. Previously saved validated checkpoints remain in the unfinished run. Select it on the next start to resume.")
        Catch ex As System.Exception
            AppendStatus("Analysis stopped: " & ex.Message)
            Dim explanation As String = "The collection analysis could not be completed. The existing source files were not changed." &
                Environment.NewLine & "Any saved validated progress remains in its original run folder. Select that unfinished run on the next start to retry."
            If TypeOf ex Is System.IO.InvalidDataException AndAlso
               ex.Message.StartsWith("The selected model could not provide verifiable source references", System.StringComparison.Ordinal) Then
                explanation &= Environment.NewLine & Environment.NewLine &
                    "The selected AI model did not provide sufficiently verifiable source references. " &
                    "Please try again. If the problem continues, select another model. " &
                    "Diagnostic responses, if available, are saved in the job workspace."
            Else
                explanation &= Environment.NewLine & Environment.NewLine &
                    "Please check the progress details, source folder and model configuration before retrying."
            End If
            ShowRedInkMessage(explanation)
        Finally
            _cts.Dispose()
            _cts = Nothing
            _running = False
            _run.Enabled = True
            _cancel.Text = "Close"
            _source.Enabled = True
            _output.Enabled = True
            _objective.Enabled = True
            _duplicates.Enabled = True
            _annex.Enabled = True
            _modelChoice.Enabled = True
        End Try
    End Sub

    Private Sub CancelClicked(sender As Object, e As EventArgs)
        If _running Then
            _cts.Cancel()
            AppendStatus("Cancelling after the current operation...")
        Else
            Me.Close()
        End If
    End Sub

    Private Sub ClosingDialog(sender As Object, e As System.Windows.Forms.FormClosingEventArgs)
        If _running Then
            _cts.Cancel()
            e.Cancel = True
            AppendStatus("Cancellation requested. The window will remain open until the current call returns.")
            Return
        End If
        Try
            SavePreferences()
        Catch ex As System.Exception
            AppendStatus("Unable to save preferences on close: " & ex.Message)
            If Not ConfirmCloseWithoutSaving("The Collection Analyzer preferences could not be saved." & Environment.NewLine &
                                             "Close without saving these preferences?") Then e.Cancel = True
        End Try
    End Sub

    Private Sub AppendStatus(message As String)
        If Me.IsDisposed Then Return
        _progress.AppendText(Environment.NewLine & DateTime.Now.ToString("HH:mm:ss") & "  " & message)
    End Sub
End Class
