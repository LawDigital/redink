' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: TalkToMeWidget.vb
' Purpose: Provides the user interface for the "Talk to Me" feature. It's a
'          floating, non-modal window that displays recording status and handles
'          user interaction to start/stop transcription.
'
' Architecture:
'  - UI Component: A `System.Windows.Forms.Form` designed as a lightweight,
'    always-on-top widget.
'  - State Display: Visually indicates the current state of transcription
'    (e.g., listening, processing) to the user.
'  - User Interaction: Captures user clicks to toggle the recording state via
'    the `ITalkToMeHost` interface.
'  - Host Communication: Interacts with a host application (e.g., Word Add-in)
'    through the `ITalkToMeHost` interface to control the transcription process.
' =============================================================================


Option Explicit On
Option Strict On

Imports System.Drawing
Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Newtonsoft.Json


Namespace SharedLibrary
    Public Class TalkToMeWidget
        Inherits Form

        Private NotInheritable Class WidgetSettings
            Public Property X As Integer = Integer.MinValue
            Public Property Y As Integer = Integer.MinValue
            Public Property Width As Integer = 420
            Public Property Height As Integer = 72
            Public Property IncludeFullDocument As Boolean = False
        End Class

        Private Const DefaultPromptText As String = "Ready. Click ▶ to start."
        Private Const ListeningPromptText As String = "Listening…"
        Private Const DispatchPauseMilliseconds As Integer = 250

        Private ReadOnly _speechAdapter As ITalkToMeSpeechAdapter
        Private ReadOnly _coordinator As TalkToMeCoordinator

        Private _settings As WidgetSettings
        Private _isDisplaySettingsHooked As Boolean = False
        Private _isClosing As Boolean = False
        Private _dispatchCts As CancellationTokenSource = Nothing
        Private _isBusy As Boolean = False
        Private _returnFocusAfterStart As Action = Nothing
        Private _speechOutputRefreshTimer As System.Windows.Forms.Timer

        Private _rootLayout As TableLayoutPanel
        Private WithEvents btnStartStop As Button
        Private WithEvents btnConfigure As Button
        Private WithEvents btnSpeechOutput As Button
        Private lblTranscript As Label

        Public Sub New(speechAdapter As ITalkToMeSpeechAdapter,
                       coordinator As TalkToMeCoordinator)
            _speechAdapter = speechAdapter
            _coordinator = coordinator

            InitializeComponent()
            LoadSettings()
            RestoreBoundsSafe()

            _speechOutputRefreshTimer = New System.Windows.Forms.Timer() With {
                .Interval = 250,
                .Enabled = True
            }
            AddHandler _speechOutputRefreshTimer.Tick, Sub(sender As Object, e As EventArgs) UpdateSpeechOutputUi()

            AddHandler _speechAdapter.PartialTranscriptReceived, AddressOf OnPartialTranscriptReceived
            AddHandler _speechAdapter.FinalTranscriptReceived, AddressOf OnFinalTranscriptReceived
        End Sub

        Public Function GetIncludeFullDocumentSetting() As Boolean
            Return _settings IsNot Nothing AndAlso _settings.IncludeFullDocument
        End Function

        Public Sub SetReturnFocusAfterStart(action As Action)
            _returnFocusAfterStart = action
        End Sub

        Private Sub InitializeComponent()
            Me.Text = SharedMethods.AN & " - Talk to me!"
            Me.AutoScaleDimensions = New SizeF(96.0F, 96.0F)
            Me.AutoScaleMode = AutoScaleMode.Dpi
            Me.FormBorderStyle = FormBorderStyle.Sizable
            Me.TopMost = True
            Me.ShowInTaskbar = False
            Me.StartPosition = FormStartPosition.Manual
            Me.MinimumSize = New Size(360, 72)
            Me.Size = New Size(520, 72)
            Me.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            Me.Padding = New Padding(0)

            Try
                Dim bmp As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                Me.Icon = Icon.FromHandle(bmp.GetHicon())
            Catch
            End Try

            _rootLayout = New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 4,
                .RowCount = 1,
                .Padding = New Padding(8, 4, 8, 4),
                .Margin = New Padding(0)
            }
            _rootLayout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            _rootLayout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            _rootLayout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            _rootLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            _rootLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))

            btnStartStop = New Button() With {
                .Text = ChrW(&H25B6),
                .Font = New Font("Segoe UI Symbol", 9.0F, FontStyle.Regular, GraphicsUnit.Point),
                .Size = New Size(22, 22),
                .MinimumSize = New Size(22, 22),
                .MaximumSize = New Size(22, 22),
                .Margin = New Padding(0, 0, 4, 0),
                .Padding = New Padding(0),
                .UseVisualStyleBackColor = True
            }

            btnSpeechOutput = New Button() With {
                .Text = Char.ConvertFromUtf32(&H1F50A),
                .Font = New Font("Segoe UI Emoji", 9.0F, FontStyle.Regular, GraphicsUnit.Point),
                .Size = New Size(22, 22),
                .MinimumSize = New Size(22, 22),
                .MaximumSize = New Size(22, 22),
                .Margin = New Padding(0, 0, 4, 0),
                .Padding = New Padding(0),
                .UseVisualStyleBackColor = True,
                .Visible = False,
                .TabStop = False
            }

            btnConfigure = New Button() With {
                .Text = ChrW(&H2699),
                .Font = New Font("Segoe UI Symbol", 9.0F, FontStyle.Regular, GraphicsUnit.Point),
                .Size = New Size(22, 22),
                .MinimumSize = New Size(22, 22),
                .MaximumSize = New Size(22, 22),
                .Margin = New Padding(0, 0, 6, 0),
                .Padding = New Padding(0),
                .UseVisualStyleBackColor = True
            }

            lblTranscript = New Label() With {
                .Dock = DockStyle.Fill,
                .Text = DefaultPromptText,
                .AutoSize = False,
                .AutoEllipsis = True,
                .TextAlign = ContentAlignment.MiddleLeft,
                .Margin = New Padding(0)
            }

            _rootLayout.Controls.Add(btnStartStop, 0, 0)
            _rootLayout.Controls.Add(btnSpeechOutput, 1, 0)
            _rootLayout.Controls.Add(btnConfigure, 2, 0)
            _rootLayout.Controls.Add(lblTranscript, 3, 0)

            Me.Controls.Add(_rootLayout)

            ApplyCalculatedMinimumSize()
            UpdateUiState()
        End Sub

        Protected Overrides Sub OnHandleCreated(e As EventArgs)
            MyBase.OnHandleCreated(e)

            If Not _isDisplaySettingsHooked Then
                AddHandler Microsoft.Win32.SystemEvents.DisplaySettingsChanged, AddressOf OnDisplaySettingsChanged
                _isDisplaySettingsHooked = True
            End If
        End Sub

        Private _speechStopTask As Task = Nothing

        Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
            BeginShutdown()
            MyBase.OnFormClosing(e)
        End Sub

        Public Sub PrepareForHostShutdown()
            BeginShutdown()

            Try
                If Not Me.IsDisposed Then
                    Me.Hide()
                End If
            Catch
            End Try
        End Sub

        Private Sub BeginShutdown()
            If _isClosing Then
                Return
            End If

            _isClosing = True

            If _isDisplaySettingsHooked Then
                RemoveHandler Microsoft.Win32.SystemEvents.DisplaySettingsChanged, AddressOf OnDisplaySettingsChanged
                _isDisplaySettingsHooked = False
            End If

            SaveSettings()

            Try
                _dispatchCts?.Cancel()
                _dispatchCts?.Dispose()
                _dispatchCts = Nothing
            Catch
            End Try

            Try
                If _speechOutputRefreshTimer IsNot Nothing Then
                    _speechOutputRefreshTimer.Stop()
                    _speechOutputRefreshTimer.Dispose()
                    _speechOutputRefreshTimer = Nothing
                End If
            Catch
            End Try

            RemoveHandler _speechAdapter.PartialTranscriptReceived, AddressOf OnPartialTranscriptReceived
            RemoveHandler _speechAdapter.FinalTranscriptReceived, AddressOf OnFinalTranscriptReceived

            BeginStopSpeechAdapter()
        End Sub

        Private Sub BeginStopSpeechAdapter()
            If _speechAdapter Is Nothing Then
                Return
            End If

            If _speechStopTask IsNot Nothing Then
                Return
            End If

            Try
                _speechStopTask =
                    Task.Run(
                        Async Function()
                            Try
                                Await _speechAdapter.StopListeningAsync().ConfigureAwait(False)
                            Catch
                            End Try
                        End Function)
            Catch
            End Try
        End Sub

        Public Sub ShowWidget()
            ApplyCalculatedMinimumSize()
            EnsureWidgetVisible()

            If Me.Visible Then
                Me.BringToFront()
                Me.Activate()
            Else
                Me.Show()
            End If
        End Sub

        Public Async Function SubmitExternalSpeechAsync(speakerName As String,
                                                        text As String) As Task(Of Boolean)
            If _isClosing OrElse Me.IsDisposed Then
                Return False
            End If

            If Not _speechAdapter.CanAcceptExternalSpeech Then
                UpdateSpeechOutputUi()
                Return False
            End If

            Return Await _speechAdapter.SubmitExternalSpeechAsync(
                speakerName,
                text,
                CancellationToken.None).ConfigureAwait(False)
        End Function

        Private Async Sub btnStartStop_Click(sender As Object, e As EventArgs) Handles btnStartStop.Click
            If _isBusy Then
                Return
            End If

            If _speechAdapter.IsListening Then
                Await StopListeningAsync().ConfigureAwait(True)
            Else
                Await StartListeningAsync().ConfigureAwait(True)
            End If
        End Sub

        Private Sub btnConfigure_Click(sender As Object, e As EventArgs) Handles btnConfigure.Click
            ConfigureSpeech()
        End Sub

        Private Sub btnSpeechOutput_Click(sender As Object, e As EventArgs) Handles btnSpeechOutput.Click
            If _isBusy Then
                Return
            End If

            ConfigureSpeech()
            UpdateSpeechOutputUi()
        End Sub

        Private Sub ConfigureSpeech()
            Dim result As TalkToMeSpeechConfigurationResult =
                _speechAdapter.Configure(Me, GetIncludeFullDocumentSetting())

            If result IsNot Nothing AndAlso result.Applied Then
                If _settings Is Nothing Then
                    _settings = New WidgetSettings()
                End If

                _settings.IncludeFullDocument = result.IncludeFullDocument
                SaveSettings()

                If String.IsNullOrWhiteSpace(result.Summary) Then
                    SetDisplayText("Configuration updated.")
                Else
                    SetDisplayText(result.Summary)
                End If
            End If

            UpdateUiState()
            UpdateSpeechOutputUi()
        End Sub

        Private Async Function StartListeningAsync() As Task
            If Not _speechAdapter.IsConfigured Then
                ConfigureSpeech()

                If Not _speechAdapter.IsConfigured Then
                    SetDisplayText("Configuration required.")
                    UpdateUiState()
                    Return
                End If
            End If

            _isBusy = True
            UpdateUiState()
            SetDisplayText("Starting…")

            Try
                Await _speechAdapter.StartListeningAsync(CancellationToken.None).ConfigureAwait(True)
                SetDisplayText(ListeningPromptText)

                If _returnFocusAfterStart IsNot Nothing Then
                    BeginInvoke(
                        New MethodInvoker(
                            Sub()
                                Try
                                    _returnFocusAfterStart.Invoke()
                                Catch
                                End Try
                            End Sub))
                End If
            Catch ex As System.Exception
                SetDisplayText("Start failed: " & ex.Message)
            Finally
                _isBusy = False
                UpdateUiState()
            End Try
        End Function

        Private Async Function StopListeningAsync() As Task
            _isBusy = True
            UpdateUiState()
            SetDisplayText("Stopping…")

            Try
                Await _speechAdapter.StopListeningAsync().ConfigureAwait(True)
                SetDisplayText(DefaultPromptText)
            Catch ex As System.Exception
                SetDisplayText("Stop failed: " & ex.Message)
            Finally
                _isBusy = False
                UpdateUiState()
            End Try
        End Function

        Private Sub OnPartialTranscriptReceived(sender As Object, e As TalkToMeTranscriptEventArgs)
            If _isClosing OrElse Me.IsDisposed Then
                Return
            End If

            If _speechAdapter.IsSpeechOutputActive Then
                Return
            End If

            If Me.InvokeRequired Then
                Me.BeginInvoke(New Action(Of Object, TalkToMeTranscriptEventArgs)(AddressOf OnPartialTranscriptReceived), sender, e)
                Return
            End If

            SetDisplayText(e.Text)
        End Sub

        Private Async Sub OnFinalTranscriptReceived(sender As Object, e As TalkToMeTranscriptEventArgs)
            If _isClosing OrElse Me.IsDisposed Then
                Return
            End If

            If _speechAdapter.IsSpeechOutputActive Then
                Return
            End If

            If Me.InvokeRequired Then
                Me.BeginInvoke(New Action(Of Object, TalkToMeTranscriptEventArgs)(AddressOf OnFinalTranscriptReceived), sender, e)
                Return
            End If

            Dim finalText As String = If(e.Text, "").Trim()

            If String.IsNullOrWhiteSpace(finalText) Then
                Return
            End If

            SetDisplayText(finalText)

            If finalText.StartsWith("Error:", StringComparison.OrdinalIgnoreCase) Then
                Return
            End If

            Try
                _dispatchCts?.Cancel()
            Catch
            End Try

            Dim localCts As New CancellationTokenSource()
            _dispatchCts = localCts

            Try
                Await Task.Delay(DispatchPauseMilliseconds, localCts.Token).ConfigureAwait(True)

                Dim result As TalkToMeDispatchResult =
                    Await _coordinator.ProcessTranscriptAsync(finalText, localCts.Token).ConfigureAwait(True)

                If localCts.IsCancellationRequested Then
                    Return
                End If

                If result Is Nothing Then
                    SetDisplayText("No result.")
                ElseIf Not String.IsNullOrWhiteSpace(result.TranscriptToDisplay) Then
                    SetDisplayText(result.TranscriptToDisplay)
                ElseIf Not String.IsNullOrWhiteSpace(result.StatusText) Then
                    SetDisplayText(result.StatusText)
                ElseIf _speechAdapter.IsListening Then
                    SetDisplayText(ListeningPromptText)
                Else
                    SetDisplayText(DefaultPromptText)
                End If
            Catch ex As OperationCanceledException
            Catch ex As System.Exception
                SetDisplayText("Error: " & ex.Message)
            Finally
                If ReferenceEquals(_dispatchCts, localCts) Then
                    Try
                        localCts.Dispose()
                    Catch
                    End Try

                    _dispatchCts = Nothing
                Else
                    Try
                        localCts.Dispose()
                    Catch
                    End Try
                End If
            End Try
        End Sub

        Private Sub SetDisplayText(text As String)
            If Me.IsDisposed OrElse lblTranscript Is Nothing Then
                Return
            End If

            If Me.InvokeRequired Then
                Try
                    Me.BeginInvoke(New Action(Of String)(AddressOf SetDisplayText), text)
                Catch
                End Try
                Return
            End If

            Dim canShowMultipleLines As Boolean = CanShowMultipleTranscriptLines()
            Dim value As String

            If canShowMultipleLines Then
                value = FitTranscriptMultiline(If(text, "").Trim())
                lblTranscript.AutoEllipsis = False
                lblTranscript.TextAlign = ContentAlignment.TopLeft
            Else
                value = FitTranscript(If(text, "").Trim())
                lblTranscript.AutoEllipsis = True
                lblTranscript.TextAlign = ContentAlignment.MiddleLeft
            End If

            If String.IsNullOrWhiteSpace(value) Then
                value = DefaultPromptText
            End If

            lblTranscript.Text = value
        End Sub

        Private Sub UpdateUiState()
            If Me.IsDisposed OrElse btnStartStop Is Nothing OrElse btnConfigure Is Nothing Then
                Return
            End If

            If Me.InvokeRequired Then
                Try
                    Me.BeginInvoke(New MethodInvoker(AddressOf UpdateUiState))
                Catch
                End Try
                Return
            End If

            btnStartStop.Enabled = Not _isBusy
            btnConfigure.Enabled = Not _isBusy AndAlso Not _speechAdapter.IsListening

            btnStartStop.Text = If(_speechAdapter.IsListening, ChrW(&H25A0), ChrW(&H25B6))

            If _speechAdapter.IsListening Then
                btnStartStop.BackColor = Color.FromArgb(220, 240, 220)
            Else
                btnStartStop.BackColor = SystemColors.Control
            End If

            UpdateSpeechOutputUi()
        End Sub

        Private Shared Function FitTranscript(text As String) As String
            Return System.Text.RegularExpressions.Regex.Replace(If(text, ""), "\s+", " ").Trim()
        End Function

        Private Shared Function FitTranscriptMultiline(text As String) As String
            Dim normalized As String = If(text, "").Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            Dim lines As String() =
                normalized.Split(New String() {vbLf}, StringSplitOptions.None).
                    Select(Function(line) System.Text.RegularExpressions.Regex.Replace(line, "\s+", " ").Trim()).
                    ToArray()

            Return String.Join(Environment.NewLine, lines).Trim()
        End Function

        Private Function CanShowMultipleTranscriptLines() As Boolean
            If lblTranscript Is Nothing Then
                Return False
            End If

            Dim lineHeight As Integer = CInt(Math.Ceiling(lblTranscript.Font.GetHeight()))
            Return lblTranscript.Height >= (lineHeight * 2) + 4
        End Function

        Private Sub UpdateSpeechOutputUi()
            If Me.IsDisposed OrElse btnSpeechOutput Is Nothing Then
                Return
            End If

            If Me.InvokeRequired Then
                Try
                    Me.BeginInvoke(New MethodInvoker(AddressOf UpdateSpeechOutputUi))
                Catch
                End Try
                Return
            End If

            btnSpeechOutput.Visible = _speechAdapter.IsSpeechOutputActive
            btnSpeechOutput.Enabled = False
            btnSpeechOutput.BackColor = Color.FromArgb(220, 240, 220)
        End Sub

        Private Sub ApplyCalculatedMinimumSize()
            If Me.IsDisposed OrElse _rootLayout Is Nothing Then
                Return
            End If

            _rootLayout.PerformLayout()

            Dim preferredClientSize As Size =
                _rootLayout.GetPreferredSize(New Size(Integer.MaxValue, Integer.MaxValue))

            Dim nonClientWidth As Integer = Me.Width - Me.ClientSize.Width
            Dim nonClientHeight As Integer = Me.Height - Me.ClientSize.Height

            Dim minWidth As Integer = Math.Max(360, preferredClientSize.Width + nonClientWidth + 8)
            Dim minHeight As Integer = Math.Max(72, preferredClientSize.Height + nonClientHeight + 8)

            Me.MinimumSize = New Size(minWidth, minHeight)

            If Me.Width < minWidth OrElse Me.Height < minHeight Then
                Me.Size = New Size(Math.Max(Me.Width, minWidth), Math.Max(Me.Height, minHeight))
            End If
        End Sub

        Private Sub LoadSettings()
            _settings = New WidgetSettings()

            Try
                _settings.X = My.Settings.TalkToMeWidgetX
                _settings.Y = My.Settings.TalkToMeWidgetY
                _settings.Width = If(My.Settings.TalkToMeWidgetWidth > 0, My.Settings.TalkToMeWidgetWidth, 420)
                _settings.Height = Math.Max(Me.MinimumSize.Height, If(My.Settings.TalkToMeWidgetHeight > 0, My.Settings.TalkToMeWidgetHeight, Me.MinimumSize.Height))
                _settings.IncludeFullDocument = My.Settings.TalkToMeIncludeFullDocument
            Catch
                _settings = New WidgetSettings()
            End Try
        End Sub

        Private Sub SaveSettings()
            Try
                If _settings Is Nothing Then
                    _settings = New WidgetSettings()
                End If

                _settings.X = Me.Left
                _settings.Y = Me.Top
                _settings.Width = Me.Width
                _settings.Height = Math.Max(Me.MinimumSize.Height, Me.Height)

                My.Settings.TalkToMeWidgetX = _settings.X
                My.Settings.TalkToMeWidgetY = _settings.Y
                My.Settings.TalkToMeWidgetWidth = _settings.Width
                My.Settings.TalkToMeWidgetHeight = _settings.Height
                My.Settings.TalkToMeIncludeFullDocument = _settings.IncludeFullDocument
                My.Settings.Save()
            Catch
            End Try
        End Sub

        Private Sub RestoreBoundsSafe()
            If _settings Is Nothing Then
                PositionOnScreen()
                EnsureWidgetVisible()
                Return
            End If

            If _settings.Width > 0 AndAlso _settings.Height > 0 AndAlso
               _settings.X <> Integer.MinValue AndAlso _settings.Y <> Integer.MinValue Then

                Dim safeWidth As Integer = Math.Max(Me.MinimumSize.Width, _settings.Width)
                Dim safeHeight As Integer = Math.Max(Me.MinimumSize.Height, _settings.Height)
                Me.SetBounds(_settings.X, _settings.Y, safeWidth, safeHeight)
            Else
                PositionOnScreen()
            End If

            EnsureWidgetVisible()
        End Sub

        Private Sub PositionOnScreen()
            Dim wa As Rectangle = Screen.FromPoint(Cursor.Position).WorkingArea
            Const margin As Integer = 30
            Me.Location = New Point(wa.Right - Me.Width - margin, wa.Top + margin)
        End Sub

        Private Sub EnsureWidgetVisible()
            SharedMethods.EnsureVisibleOnScreen(Me)
        End Sub

        Private Sub OnDisplaySettingsChanged(sender As Object, e As EventArgs)
            If _isClosing OrElse Me.IsDisposed Then
                Return
            End If

            Try
                If Me.InvokeRequired Then
                    Me.BeginInvoke(
                        New MethodInvoker(
                            Sub()
                                ApplyCalculatedMinimumSize()
                                EnsureWidgetVisible()
                            End Sub))
                Else
                    ApplyCalculatedMinimumSize()
                    EnsureWidgetVisible()
                End If
            Catch
            End Try
        End Sub
    End Class
End Namespace
