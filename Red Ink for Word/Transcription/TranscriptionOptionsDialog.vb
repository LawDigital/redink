' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: TranscriptionOptionsDialog.vb
' Purpose: Provides a user interface for configuring transcription settings.
'          This includes selecting the transcription engine, audio input
'          device, and other engine-specific options.
'
' Architecture:
'  - UI Design: A Windows Form that presents transcription options to the user
'    in an organized manner.
'  - Configuration Loading/Saving: Loads the current settings upon opening and
'    saves any changes made by the user.
'  - Control Binding: Binds UI controls (e.g., dropdowns, checkboxes) to the
'    underlying configuration properties.
'  - User Interaction: Handles user input to update settings and provides
'    mechanisms to confirm or cancel changes.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Drawing
Imports System.Linq
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.Transcription

Namespace Transcription

    Public Class TranscriptionOptionsDialog
        Inherits Form

        Private Class NamedValueItem
            Public Property Name As String
            Public Property Value As String

            Public Sub New(name As String, value As String)
                Me.Name = name
                Me.Value = value
            End Sub

            Public Overrides Function ToString() As String
                Return Name
            End Function
        End Class

        Public Property Options As TranscriptionOptions
        Public Property SelectedSourceMode As String
        Public Property SelectedOutputDeviceId As String

        Private ReadOnly _kind As EngineKind
        Private ReadOnly _toolTip As New ToolTip()
        Private ReadOnly _root As TableLayoutPanel
        Private ReadOnly _contentPanel As Panel
        Private ReadOnly _contentLayout As TableLayoutPanel
        Private ReadOnly _buttonsTable As TableLayoutPanel

        Private cboLang As ComboBox
        Private cboSourceMode As ComboBox
        Private cboOutputDevice As ComboBox
        Private chkDiar As CheckBox
        Private nudMin As NumericUpDown
        Private nudMax As NumericUpDown
        Private chkMulti As CheckBox
        Private chkTrans As CheckBox
        Private nudVad As NumericUpDown
        Private nudVosk As NumericUpDown
        Private txtModel As TextBox
        Private cboTurn As ComboBox
        Private txtPrompt As TextBox
        Private chkDebug As CheckBox

        Public Sub New(kind As EngineKind,
                       recognizerDisplayName As String,
                       currentOpts As TranscriptionOptions,
                       langChoices As String(),
                       currentSourceMode As String,
                       currentOutputDeviceId As String,
                       outputDevices As IEnumerable(Of KeyValuePair(Of String, String)))

            _kind = kind
            Options = CloneOptions(currentOpts)
            SelectedSourceMode = currentSourceMode
            SelectedOutputDeviceId = currentOutputDeviceId

            Me.Text = $"{Globals.ThisAddIn.AN} Transcriptor Options"
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.Sizable
            Me.MinimizeBox = False
            Me.MaximizeBox = True
            Me.ShowInTaskbar = False
            Me.AutoScaleMode = AutoScaleMode.Dpi
            Me.Font = New System.Drawing.Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            Me.Padding = New Padding(10)
            Me.MinimumSize = New Size(900, 660)
            Me.ClientSize = New Size(1000, 720)

            Try
                Dim bmp As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                Me.Icon = Icon.FromHandle(bmp.GetHicon())
            Catch
            End Try

            _root = New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 2
            }
            _root.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
            _root.RowStyles.Add(New RowStyle(SizeType.AutoSize))

            _contentPanel = New Panel() With {
                .Dock = DockStyle.Fill,
                .AutoScroll = True,
                .Padding = New Padding(0)
            }

            _contentLayout = New TableLayoutPanel() With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .ColumnCount = 1,
                .Padding = New Padding(0),
                .Margin = New Padding(0)
            }

            BuildControls(langChoices, outputDevices)
            BuildSections(kind, recognizerDisplayName)
            UpdateDiarizationUi()

            _contentPanel.Controls.Add(_contentLayout)
            _root.Controls.Add(_contentPanel, 0, 0)

            _buttonsTable = New TableLayoutPanel() With {
                .Dock = DockStyle.Bottom,
                .ColumnCount = 3,
                .RowCount = 1,
                .AutoSize = True,
                .Padding = New Padding(0, 12, 0, 0)
            }
            _buttonsTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            _buttonsTable.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            _buttonsTable.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))

            Dim btnOk As New Button() With {
                .Text = "OK",
                .DialogResult = DialogResult.OK,
                .Width = 96,
                .Height = 30,
                .Anchor = AnchorStyles.Right
            }

            Dim btnCancel As New Button() With {
                .Text = "Cancel",
                .DialogResult = DialogResult.Cancel,
                .Width = 96,
                .Height = 30,
                .Anchor = AnchorStyles.Right,
                .Margin = New Padding(8, 0, 0, 0)
            }

            AddHandler btnOk.Click, AddressOf Ok_Click

            _buttonsTable.Controls.Add(New Panel() With {.Dock = DockStyle.Fill}, 0, 0)
            _buttonsTable.Controls.Add(btnOk, 1, 0)
            _buttonsTable.Controls.Add(btnCancel, 2, 0)

            Me.AcceptButton = btnOk
            Me.CancelButton = btnCancel

            _root.Controls.Add(_buttonsTable, 0, 1)
            Me.Controls.Add(_root)
        End Sub

        Private Shared Function CloneOptions(source As TranscriptionOptions) As TranscriptionOptions
            If source Is Nothing Then
                Return New TranscriptionOptions()
            End If

            Dim clone As New TranscriptionOptions With {
                .LanguageCode = source.LanguageCode,
                .EnableDiarization = source.EnableDiarization,
                .MinSpeakers = source.MinSpeakers,
                .MaxSpeakers = source.MaxSpeakers,
                .Translate = source.Translate,
                .Model = source.Model,
                .VadThreshold = source.VadThreshold,
                .VoskSimilarityThreshold = source.VoskSimilarityThreshold,
                .MultiChannelDiarization = source.MultiChannelDiarization,
                .AudioDebugDump = source.AudioDebugDump,
                .TurnDetection = source.TurnDetection,
                .Prompt = source.Prompt
            }

            SetOptionDiarization(clone, GetOptionDiarization(source))
            SetOptionDiarizationMaxSpeakers(clone, GetOptionDiarizationMaxSpeakers(source))

            Return clone
        End Function

        Private Shared Function GetOptionDiarization(opts As TranscriptionOptions) As Boolean
            If opts Is Nothing Then
                Return False
            End If

            Dim value As Boolean = opts.EnableDiarization

            Try
                Dim prop = opts.GetType().GetProperty("Diarization")
                If prop IsNot Nothing Then
                    Dim raw As Object = prop.GetValue(opts, Nothing)
                    If raw IsNot Nothing Then
                        value = System.Convert.ToBoolean(raw, System.Globalization.CultureInfo.InvariantCulture)
                    End If
                End If
            Catch
            End Try

            Return value
        End Function

        Private Shared Sub SetOptionDiarization(opts As TranscriptionOptions, value As Boolean)
            If opts Is Nothing Then
                Return
            End If

            opts.EnableDiarization = value

            Try
                Dim prop = opts.GetType().GetProperty("Diarization")
                If prop IsNot Nothing AndAlso prop.CanWrite Then
                    prop.SetValue(opts, value, Nothing)
                End If
            Catch
            End Try
        End Sub

        Private Shared Function GetOptionDiarizationMaxSpeakers(opts As TranscriptionOptions) As Integer
            If opts Is Nothing Then
                Return 2
            End If

            Dim value As Integer = opts.MaxSpeakers

            Try
                Dim prop = opts.GetType().GetProperty("DiarizationMaxSpeakers")
                If prop IsNot Nothing Then
                    Dim raw As Object = prop.GetValue(opts, Nothing)
                    If raw IsNot Nothing Then
                        value = System.Convert.ToInt32(raw, System.Globalization.CultureInfo.InvariantCulture)
                    End If
                End If
            Catch
            End Try

            If value < 2 Then
                value = 2
            End If

            If value > 35 Then
                value = 35
            End If

            Return value
        End Function

        Private Shared Sub SetOptionDiarizationMaxSpeakers(opts As TranscriptionOptions, value As Integer)
            If opts Is Nothing Then
                Return
            End If

            If value < 2 Then
                value = 2
            End If

            If value > 35 Then
                value = 35
            End If

            opts.MaxSpeakers = value

            Try
                Dim prop = opts.GetType().GetProperty("DiarizationMaxSpeakers")
                If prop IsNot Nothing AndAlso prop.CanWrite Then
                    prop.SetValue(opts, value, Nothing)
                End If
            Catch
            End Try
        End Sub

        Private Sub BuildControls(langChoices As IEnumerable(Of String), outputDevices As IEnumerable(Of KeyValuePair(Of String, String)))
            cboLang = New ComboBox() With {
                .Dock = DockStyle.Top,
                .DropDownStyle = ComboBoxStyle.DropDown,
                .Margin = New Padding(0, 2, 0, 6)
            }
            cboLang.Items.AddRange(langChoices.Select(Function(x) CObj(x)).ToArray())
            cboLang.Text = Options.LanguageCode

            cboSourceMode = New ComboBox() With {
                .Dock = DockStyle.Top,
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Margin = New Padding(0, 2, 0, 6)
            }
            cboSourceMode.Items.AddRange(New Object() {"MicrophoneOnly", "SystemAudioOnly", "MicrophoneAndSystem"})
            If cboSourceMode.Items.Contains(SelectedSourceMode) Then
                cboSourceMode.SelectedItem = SelectedSourceMode
            Else
                cboSourceMode.SelectedItem = "MicrophoneOnly"
            End If

            cboOutputDevice = New ComboBox() With {
                .Dock = DockStyle.Top,
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Margin = New Padding(0, 2, 0, 6)
            }

            For Each kvp In outputDevices
                cboOutputDevice.Items.Add(New NamedValueItem(kvp.Key, kvp.Value))
            Next

            Dim matchedOutput As Boolean = False
            For i As Integer = 0 To cboOutputDevice.Items.Count - 1
                Dim item = TryCast(cboOutputDevice.Items(i), NamedValueItem)
                If item IsNot Nothing AndAlso String.Equals(item.Value, SelectedOutputDeviceId, StringComparison.Ordinal) Then
                    cboOutputDevice.SelectedIndex = i
                    matchedOutput = True
                    Exit For
                End If
            Next
            If Not matchedOutput AndAlso cboOutputDevice.Items.Count > 0 Then
                cboOutputDevice.SelectedIndex = 0
            End If

            chkDiar = New CheckBox() With {
                .AutoSize = True,
                .Text = "Enable speaker diarization",
                .Checked = GetOptionDiarization(Options),
                .Margin = New Padding(0, 4, 0, 6)
            }

            nudMin = New NumericUpDown() With {
                .Dock = DockStyle.Left,
                .Minimum = 2,
                .Maximum = 8,
                .Value = Math.Max(2, Math.Min(8, Options.MinSpeakers)),
                .Width = 110,
                .Margin = New Padding(0, 2, 0, 6)
            }

            nudMax = New NumericUpDown() With {
                .Dock = DockStyle.Left,
                .Minimum = 2,
                .Maximum = 35,
                .Value = Math.Max(2, Math.Min(35, GetOptionDiarizationMaxSpeakers(Options))),
                .Width = 110,
                .Margin = New Padding(0, 2, 0, 6)
            }

            chkMulti = New CheckBox() With {
                .AutoSize = True,
                .Text = "Use separate mic/system channels",
                .Checked = Options.MultiChannelDiarization,
                .Margin = New Padding(0, 4, 0, 6)
            }

            chkTrans = New CheckBox() With {
                .AutoSize = True,
                .Text = "Translate output to English",
                .Checked = Options.Translate,
                .Margin = New Padding(0, 4, 0, 6)
            }

            nudVad = New NumericUpDown() With {
                .Dock = DockStyle.Left,
                .DecimalPlaces = 2,
                .Increment = 0.05D,
                .Minimum = 0.05D,
                .Maximum = 0.95D,
                .Value = CDec(Math.Max(0.05F, Math.Min(0.95F, Options.VadThreshold))),
                .Width = 110,
                .Margin = New Padding(0, 2, 0, 6)
            }

            nudVosk = New NumericUpDown() With {
                .Dock = DockStyle.Left,
                .DecimalPlaces = 2,
                .Increment = 0.1D,
                .Minimum = 0.2D,
                .Maximum = 2.5D,
                .Value = CDec(Math.Max(0.2, Math.Min(2.5, Options.VoskSimilarityThreshold))),
                .Width = 110,
                .Margin = New Padding(0, 2, 0, 6)
            }

            txtModel = New TextBox() With {
                .Dock = DockStyle.Top,
                .Text = Options.Model,
                .Margin = New Padding(0, 2, 0, 6)
            }

            cboTurn = New ComboBox() With {
                .Dock = DockStyle.Top,
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Margin = New Padding(0, 2, 0, 6)
            }
            cboTurn.Items.AddRange(New Object() {"server_vad", "none"})
            If String.IsNullOrWhiteSpace(Options.TurnDetection) Then
                cboTurn.SelectedItem = "server_vad"
            ElseIf cboTurn.Items.Contains(Options.TurnDetection) Then
                cboTurn.SelectedItem = Options.TurnDetection
            Else
                cboTurn.SelectedItem = "server_vad"
            End If

            txtPrompt = New TextBox() With {
                .Dock = DockStyle.Top,
                .Text = Options.Prompt,
                .Margin = New Padding(0, 2, 0, 6)
            }

            chkDebug = New CheckBox() With {
                .AutoSize = True,
                .Text = "Write debug WAV to %TEMP%",
                .Checked = Options.AudioDebugDump,
                .Margin = New Padding(0, 4, 0, 6)
            }

            _toolTip.SetToolTip(cboOutputDevice, "Select the system output device used for loopback capture.")
            _toolTip.SetToolTip(cboSourceMode, "Choose microphone-only, system-audio-only, or mixed capture.")
            _toolTip.SetToolTip(chkDiar, "Enable speaker labeling when supported by the selected recognizer.")
            _toolTip.SetToolTip(nudMin, "Minimum speaker hint.")
            _toolTip.SetToolTip(nudMax, "Maximum speaker hint. Azure Fast REST supports 2 to 35 speakers.")
            _toolTip.SetToolTip(chkMulti, "Routes microphone to the left channel and system audio to the right channel.")
            _toolTip.SetToolTip(nudVad, "Whisper no-speech threshold. Default is 0.60.")
            _toolTip.SetToolTip(nudVosk, "Vosk speaker similarity threshold.")
            _toolTip.SetToolTip(chkDebug, "Only use for diagnostics. Writes RedInk_AudioDebug.wav to %TEMP%.")

            AddHandler chkDiar.CheckedChanged, AddressOf OnDiarizationChanged
            AddHandler cboLang.SelectedIndexChanged, Sub() UpdateComboToolTip(cboLang)
            AddHandler cboSourceMode.SelectedIndexChanged, Sub() UpdateComboToolTip(cboSourceMode)
            AddHandler cboOutputDevice.SelectedIndexChanged, Sub() UpdateComboToolTip(cboOutputDevice)
            AddHandler cboTurn.SelectedIndexChanged, Sub() UpdateComboToolTip(cboTurn)

            AddHandler cboLang.MouseMove, Sub() UpdateComboToolTip(cboLang)
            AddHandler cboSourceMode.MouseMove, Sub() UpdateComboToolTip(cboSourceMode)
            AddHandler cboOutputDevice.MouseMove, Sub() UpdateComboToolTip(cboOutputDevice)
            AddHandler cboTurn.MouseMove, Sub() UpdateComboToolTip(cboTurn)
        End Sub

        Private Sub BuildSections(kind As EngineKind, recognizerDisplayName As String)
            _contentLayout.Controls.Add(BuildGlobalGroup(), 0, _contentLayout.RowCount)
            _contentLayout.RowCount += 1

            _contentLayout.Controls.Add(BuildRecognizerGroup(kind, recognizerDisplayName), 0, _contentLayout.RowCount)
            _contentLayout.RowCount += 1
        End Sub

        Private Function BuildGlobalGroup() As Control
            Dim grp As New GroupBox() With {
                .Text = "Global capture settings",
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Padding = New Padding(12)
            }

            Dim grid As New TableLayoutPanel() With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .ColumnCount = 2
            }
            grid.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 240))
            grid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))

            AddLabeledRow(grid, "Capture source", cboSourceMode)
            AddLabeledRow(grid, "System output device", cboOutputDevice)
            AddFullWidthRow(grid, chkDebug)

            grp.Controls.Add(grid)
            Return grp
        End Function

        Private Function BuildRecognizerGroup(kind As EngineKind, recognizerDisplayName As String) As Control
            Dim grp As New GroupBox() With {
                .Text = "Current recognizer settings: " & recognizerDisplayName,
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Padding = New Padding(12),
                .Margin = New Padding(0, 12, 0, 0)
            }

            Dim grid As New TableLayoutPanel() With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .ColumnCount = 2
            }
            grid.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 240))
            grid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))

            AddLabeledRow(grid, "Language / code", cboLang)

            Select Case kind
                Case EngineKind.Vosk
                    AddFullWidthRow(grid, chkDiar)
                    AddLabeledRow(grid, "Similarity threshold", nudVosk)

                Case EngineKind.WhisperLocal
                    AddFullWidthRow(grid, chkTrans)
                    AddLabeledRow(grid, "Whisper VAD", nudVad)
                    AddLabeledRow(grid, "Model override", txtModel)

                Case EngineKind.GoogleV1
                    AddFullWidthRow(grid, chkDiar)
                    AddLabeledRow(grid, "Minimum speakers", nudMin)
                    AddLabeledRow(grid, "Maximum speakers", nudMax)
                    AddLabeledRow(grid, "Model", txtModel)

                Case EngineKind.GoogleV2
                    AddFullWidthRow(grid, chkMulti)
                    AddLabeledRow(grid, "Model", txtModel)

                Case EngineKind.OpenAiRest
                    AddLabeledRow(grid, "Model", txtModel)
                    AddLabeledRow(grid, "Prompt bias", txtPrompt)

                Case EngineKind.OpenAiRealtime
                    AddFullWidthRow(grid, chkTrans)
                    AddLabeledRow(grid, "Model", txtModel)
                    AddLabeledRow(grid, "Turn detection", cboTurn)

                Case EngineKind.AzureSpeechRealtime
                    AddFullWidthRow(grid, BuildInfoLabel("Realtime Azure diarization is not available for the current non-SDK WebSocket engine."))

                Case EngineKind.AzureSpeechFastRest
                    AddFullWidthRow(grid, chkDiar)
                    AddLabeledRow(grid, "Maximum speakers", nudMax)

                Case Else
                    AddLabeledRow(grid, "Model", txtModel)
            End Select

            grp.Controls.Add(grid)
            Return grp
        End Function

        Private Shared Function BuildInfoLabel(text As String) As Control
            Return New Label() With {
                .AutoSize = True,
                .MaximumSize = New Size(700, 0),
                .Text = text,
                .ForeColor = SystemColors.GrayText,
                .Margin = New Padding(0, 2, 0, 6)
            }
        End Function

        Private Sub AddLabeledRow(grid As TableLayoutPanel, labelText As String, editor As Control)
            Dim rowIndex As Integer = grid.RowCount
            grid.RowCount += 1
            grid.RowStyles.Add(New RowStyle(SizeType.AutoSize))

            Dim lbl As New Label() With {
                .Text = labelText,
                .AutoSize = True,
                .Anchor = AnchorStyles.Left Or AnchorStyles.Top,
                .Margin = New Padding(0, 6, 12, 6)
            }

            editor.Anchor = AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top

            grid.Controls.Add(lbl, 0, rowIndex)
            grid.Controls.Add(editor, 1, rowIndex)
        End Sub

        Private Sub AddFullWidthRow(grid As TableLayoutPanel, editor As Control)
            Dim rowIndex As Integer = grid.RowCount
            grid.RowCount += 1
            grid.RowStyles.Add(New RowStyle(SizeType.AutoSize))

            editor.Anchor = AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top

            grid.Controls.Add(editor, 0, rowIndex)
            grid.SetColumnSpan(editor, 2)
        End Sub

        Private Sub UpdateComboToolTip(cbo As ComboBox)
            If cbo Is Nothing Then
                Return
            End If

            Dim text As String = If(cbo.SelectedItem IsNot Nothing, cbo.SelectedItem.ToString(), cbo.Text)
            _toolTip.SetToolTip(cbo, text)
        End Sub

        Private Sub OnDiarizationChanged(sender As Object, e As EventArgs)
            UpdateDiarizationUi()
        End Sub

        Private Sub UpdateDiarizationUi()
            If nudMax IsNot Nothing Then
                Dim maxLimit As Integer = If(_kind = EngineKind.AzureSpeechFastRest, 35, 16)
                nudMax.Maximum = maxLimit

                If nudMax.Value > nudMax.Maximum Then
                    nudMax.Value = nudMax.Maximum
                End If
            End If

            Select Case _kind
                Case EngineKind.GoogleV1
                    nudMin.Enabled = chkDiar.Checked
                    nudMax.Enabled = chkDiar.Checked

                Case EngineKind.AzureSpeechFastRest
                    nudMin.Enabled = False
                    nudMax.Enabled = chkDiar.Checked

                Case Else
                    nudMin.Enabled = False
                    nudMax.Enabled = False
            End Select
        End Sub

        Private Sub Ok_Click(sender As Object, e As EventArgs)
            Commit()
        End Sub

        Private Sub Commit()
            Options.LanguageCode = cboLang.Text.Trim()
            SetOptionDiarization(Options, chkDiar.Checked)
            Options.MinSpeakers = CInt(nudMin.Value)
            Options.MaxSpeakers = CInt(nudMax.Value)
            SetOptionDiarizationMaxSpeakers(Options, CInt(nudMax.Value))
            Options.MultiChannelDiarization = chkMulti.Checked
            Options.Translate = chkTrans.Checked
            Options.VadThreshold = CSng(nudVad.Value)
            Options.VoskSimilarityThreshold = CDbl(nudVosk.Value)
            Options.Model = txtModel.Text.Trim()
            Options.Prompt = txtPrompt.Text
            Options.AudioDebugDump = chkDebug.Checked

            If cboTurn.SelectedItem IsNot Nothing Then
                Options.TurnDetection = cboTurn.SelectedItem.ToString()
            Else
                Options.TurnDetection = "server_vad"
            End If

            If cboSourceMode.SelectedItem IsNot Nothing Then
                SelectedSourceMode = cboSourceMode.SelectedItem.ToString()
            Else
                SelectedSourceMode = "MicrophoneOnly"
            End If

            Dim item = TryCast(cboOutputDevice.SelectedItem, NamedValueItem)
            If item IsNot Nothing Then
                SelectedOutputDeviceId = item.Value
            Else
                SelectedOutputDeviceId = ""
            End If
        End Sub

    End Class

End Namespace