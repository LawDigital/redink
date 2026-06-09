' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Transcriptor.vb
' Purpose: Manages the real-time transcription workflow within Microsoft Word.
'          It integrates audio capture, transcription engine management, and
'          the insertion of transcribed text into the active document.
'
' Architecture:
'  - Engine Management: Dynamically loads and initializes the selected
'    transcription engine (e.g., OpenAI, Google, Vosk, Whisper).
'  - Audio Handling: Coordinates with the AudioCaptureService to start and
'    stop audio recording and stream data to the transcription engine.
'  - UI Integration: Manages the transcription lifecycle, including starting,
'    stopping, and displaying status updates to the user.
'  - Text Insertion: Handles the insertion of real-time and final transcription
'    results into the Word document at the current cursor position.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Word = Microsoft.Office.Interop.Word
Imports NAudio.CoreAudioApi
Imports NAudio.Wave
Imports Newtonsoft.Json
Imports Red_Ink_for_Word.Transcription
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods
Imports SharedLibrary.Transcription

Partial Public Class ThisAddIn

    Public Class TranscriptionForm
        Inherits Form

        <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
        Private Shared Function SetThreadExecutionState(esFlags As UInteger) As UInteger
        End Function

        Private Const ES_CONTINUOUS As UInteger = &H80000000UI
        Private Const ES_SYSTEM_REQUIRED As UInteger = &H1UI

        Private _setSleepLock As Boolean = False
        Private _isStopping As Boolean = False
        Private _closeAfterStop As Boolean = False

        Private rtb As RichTextBox
        Private cboEngine As ComboBox
        Private cboLang As ComboBox
        Private cboDevice As ComboBox
        Private cboProcess As ComboBox
        Private btnStart As Button
        Private btnStop As Button
        Private btnLoad As Button
        Private btnClear As Button
        Private btnOptions As Button
        Private btnProcess As Button
        Private btnQuit As Button
        Private lblLiveState As Label
        Private tt As New ToolTip()

        Private _engine As ITranscriptionEngine
        Private _capture As AudioCaptureService
        Private _opts As New TranscriptionOptions()
        Private _capturing As Boolean
        Private _cts As CancellationTokenSource
        Private _currentEngineDisplayName As String = ""

        Private _lastLiveStateText As String = ""
        Private _lastLiveStateUtc As DateTime = DateTime.MinValue
        Private _lastPartialText As String = ""

        Private _suspendSettingsPersistence As Boolean = False
        Private _alternateOpenAiConfig As ModelConfig = Nothing
        Private _alternateGoogleConfig As ModelConfig = Nothing

        Private _fileTranscribing As Boolean = False

        Private _dialogOwnerScope As IDisposable

        Public Const ACS_Bridge_Address As String = ""

        Private Class EngineDescriptor
            Public DisplayName As String
            Public Kind As EngineKind
            Public ModelOrTag As String
        End Class

        Private Class AudioInputDeviceChoice
            Public Property DeviceId As String = ""
            Public Property WaveDeviceIndex As Integer = 0
            Public Property DisplayText As String = ""
            Public Property ToolTipText As String = ""

            Public Overrides Function ToString() As String
                Return DisplayText
            End Function
        End Class

        Private _engines As New List(Of EngineDescriptor)()
        Private _audioInputDevices As New List(Of AudioInputDeviceChoice)()
        Private _promptTitles As New List(Of String)()
        Private _promptBodies As New List(Of String)()

        Public Sub New()
            _suspendSettingsPersistence = True

            Try
                InitializeComponents()
                WireHandlers()
                LoadEngines()
                LoadAudioInputDevices()
                LoadAndPopulatePrompts(Globals.ThisAddIn.INI_PromptLibPath_Transcript, cboProcess)
                RestoreSettings()
                RefreshEngineUi()
            Finally
                _suspendSettingsPersistence = False
            End Try
        End Sub

        Protected Overrides Sub OnHandleCreated(e As EventArgs)
            MyBase.OnHandleCreated(e)

            If _dialogOwnerScope Is Nothing Then
                _dialogOwnerScope = SLib.PushDialogOwner(Me)
            End If
        End Sub

        Protected Overrides Sub OnHandleDestroyed(e As EventArgs)
            Dim scope As IDisposable = _dialogOwnerScope
            _dialogOwnerScope = Nothing

            If scope IsNot Nothing Then
                Try
                    scope.Dispose()
                Catch
                End Try
            End If

            MyBase.OnHandleDestroyed(e)
        End Sub

        Private Sub InitializeComponents()
            Me.Text = $"{AN} Transcriptor (audio is not persisted)"
            Me.MinimumSize = New Size(1280, 560)
            Me.AutoScaleMode = AutoScaleMode.Dpi
            Me.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)

            rtb = New RichTextBox() With {
                .Dock = DockStyle.Fill,
                .Font = New Font("Segoe UI", 10.0F, FontStyle.Regular, GraphicsUnit.Point)
            }

            lblLiveState = New Label() With {
                .Dock = DockStyle.Top,
                .AutoSize = False,
                .Height = 56,
                .Text = "Ready.",
                .Padding = New Padding(0, 4, 0, 4)
            }

            cboEngine = NewCombo(360)
            cboLang = NewCombo(220)
            cboDevice = NewCombo(360)
            cboProcess = NewCombo(320)

            btnStart = NewBtn("Start")
            btnStop = NewBtn("Stop")
            btnStop.Enabled = False
            btnLoad = NewBtn("Load")
            btnClear = NewBtn("Clear")
            btnOptions = NewBtn("Options…")
            btnQuit = NewBtn("Quit")
            btnProcess = NewBtn("Process:")

            Dim root As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 3,
                .Padding = New Padding(10)
            }
            root.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
            root.RowStyles.Add(New RowStyle(SizeType.AutoSize))

            Dim top As New TableLayoutPanel() With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .ColumnCount = 7,
                .RowCount = 1
            }
            top.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            top.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 370))
            top.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            top.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 230))
            top.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            top.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 370))
            top.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))

            top.Controls.Add(New Label With {.Text = "Engine:", .AutoSize = True, .Anchor = AnchorStyles.Left, .Margin = New Padding(0, 8, 6, 0)}, 0, 0)
            top.Controls.Add(cboEngine, 1, 0)
            top.Controls.Add(New Label With {.Text = "Lang/Model:", .AutoSize = True, .Anchor = AnchorStyles.Left, .Margin = New Padding(0, 8, 6, 0)}, 2, 0)
            top.Controls.Add(cboLang, 3, 0)
            top.Controls.Add(New Label With {.Text = "Input:", .AutoSize = True, .Anchor = AnchorStyles.Left, .Margin = New Padding(0, 8, 6, 0)}, 4, 0)
            top.Controls.Add(cboDevice, 5, 0)
            top.Controls.Add(btnOptions, 6, 0)

            Dim mid As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 2
            }
            mid.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            mid.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
            mid.Controls.Add(lblLiveState, 0, 0)
            mid.Controls.Add(rtb, 0, 1)

            Dim bottom As New TableLayoutPanel() With {
                .Dock = DockStyle.Bottom,
                .ColumnCount = 8,
                .RowCount = 1,
                .AutoSize = True,
                .Padding = New Padding(0, 10, 0, 0)
            }
            bottom.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            bottom.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            bottom.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            bottom.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            bottom.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            bottom.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            bottom.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            bottom.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))

            cboProcess.Dock = DockStyle.Fill

            bottom.Controls.Add(btnStart, 0, 0)
            bottom.Controls.Add(btnStop, 1, 0)
            bottom.Controls.Add(btnLoad, 2, 0)
            bottom.Controls.Add(btnClear, 3, 0)
            bottom.Controls.Add(btnQuit, 4, 0)
            bottom.Controls.Add(btnOptions, 5, 0)
            bottom.Controls.Add(btnProcess, 6, 0)
            bottom.Controls.Add(cboProcess, 7, 0)

            root.Controls.Add(top, 0, 0)
            root.Controls.Add(mid, 0, 1)
            root.Controls.Add(bottom, 0, 2)

            Me.Controls.Add(root)

            Try
                Dim bmp As New Bitmap(SLib.GetLogoBitmap(SLib.LogoType.Standard))
                Me.Icon = Icon.FromHandle(bmp.GetHicon())
            Catch
            End Try
        End Sub

        Private Function NewCombo(width As Integer) As ComboBox
            Return New ComboBox With {
                .Width = width,
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Margin = New Padding(0, 3, 10, 3)
            }
        End Function

        Private Function NewBtn(text As String) As Button
            Return New Button With {
                .Text = text,
                .AutoSize = True,
                .Margin = New Padding(0, 3, 10, 3)
            }
        End Function

        Private Sub WireHandlers()
            AddHandler btnStart.Click, AddressOf OnStart
            AddHandler btnStop.Click, AddressOf OnStop
            AddHandler btnLoad.Click, AddressOf OnLoadFile
            AddHandler btnClear.Click, Sub() rtb.Clear()
            AddHandler btnOptions.Click, AddressOf OnOptions
            AddHandler btnQuit.Click, AddressOf OnQuit
            AddHandler btnProcess.Click, AddressOf OnProcess
            AddHandler Me.FormClosing, AddressOf OnClosing
            AddHandler cboEngine.SelectedIndexChanged, AddressOf OnEngineChanged

            AddHandler cboEngine.SelectedIndexChanged, Sub() UpdateComboToolTip(cboEngine)
            AddHandler cboLang.SelectedIndexChanged, Sub() UpdateComboToolTip(cboLang)
            AddHandler cboDevice.SelectedIndexChanged, Sub() UpdateComboToolTip(cboDevice)
            AddHandler cboProcess.SelectedIndexChanged, Sub() UpdateComboToolTip(cboProcess)

            AddHandler cboEngine.MouseMove, Sub() UpdateComboToolTip(cboEngine)
            AddHandler cboLang.MouseMove, Sub() UpdateComboToolTip(cboLang)
            AddHandler cboDevice.MouseMove, Sub() UpdateComboToolTip(cboDevice)
            AddHandler cboProcess.MouseMove, Sub() UpdateComboToolTip(cboProcess)

            AddHandler cboLang.SelectedIndexChanged, Sub() PersistSettings()

            AddHandler cboEngine.SelectedIndexChanged, Sub() PersistSettings()
            AddHandler cboDevice.SelectedIndexChanged, Sub() PersistSettings()

        End Sub

        Private Sub UpdateComboToolTip(cbo As ComboBox)
            If cbo Is Nothing Then
                Return
            End If

            Dim text As String = ""

            If cbo Is cboDevice Then
                Dim choice As AudioInputDeviceChoice = TryCast(cbo.SelectedItem, AudioInputDeviceChoice)
                If choice IsNot Nothing Then
                    text = choice.ToolTipText
                End If
            End If

            If String.IsNullOrWhiteSpace(text) Then
                text = If(cbo.SelectedItem IsNot Nothing, cbo.SelectedItem.ToString(), cbo.Text)
            End If

            tt.SetToolTip(cbo, text)
        End Sub


        Private Sub LoadEngines()
            _engines.Clear()
            cboEngine.Items.Clear()

            Dim modelRoot As String = ExpandEnvironmentVariables(Globals.ThisAddIn.INI_SpeechModelPath)

            If Directory.Exists(modelRoot) Then
                For Each d In Directory.GetDirectories(modelRoot).
                    OrderBy(Function(p) Path.GetFileName(p), StringComparer.OrdinalIgnoreCase)

                    Dim name As String = Path.GetFileName(d)
                    If name.StartsWith("vosk-model", StringComparison.OrdinalIgnoreCase) Then
                        _engines.Add(New EngineDescriptor With {
                            .DisplayName = "Vosk: " & name,
                            .Kind = EngineKind.Vosk,
                            .ModelOrTag = name
                        })
                    End If
                Next

                For Each f In Directory.GetFiles(modelRoot, "ggml*").
                    OrderBy(Function(p) Path.GetFileName(p), StringComparer.OrdinalIgnoreCase)

                    Dim name As String = Path.GetFileName(f)
                    _engines.Add(New EngineDescriptor With {
                        .DisplayName = "Whisper: " & name,
                        .Kind = EngineKind.WhisperLocal,
                        .ModelOrTag = name
                    })
                Next
            End If

            LoadAlternateProviderFallbacks()

            If HasConfiguredGoogleV1Provider() Then
                _engines.Add(New EngineDescriptor With {
                    .DisplayName = GoogleV1Engine.DisplayName,
                    .Kind = EngineKind.GoogleV1,
                    .ModelOrTag = "google-v1"
                })
            End If

            If HasConfiguredGoogleV2Provider() Then
                _engines.Add(New EngineDescriptor With {
                    .DisplayName = GoogleV2Engine.DisplayName,
                    .Kind = EngineKind.GoogleV2,
                    .ModelOrTag = "google-v2"
                })
            End If

            If HasConfiguredOpenAiProvider() Then
                _engines.Add(New EngineDescriptor With {
                    .DisplayName = "OpenAI gpt-4o-transcribe (REST)",
                    .Kind = EngineKind.OpenAiRest,
                    .ModelOrTag = "gpt-4o-transcribe"
                })

                _engines.Add(New EngineDescriptor With {
                    .DisplayName = "OpenAI gpt-4o-mini-transcribe (REST)",
                    .Kind = EngineKind.OpenAiRest,
                    .ModelOrTag = "gpt-4o-mini-transcribe"
                })

                _engines.Add(New EngineDescriptor With {
                    .DisplayName = "OpenAI gpt-4o-mini-transcribe-2025-12-15 (REST)",
                    .Kind = EngineKind.OpenAiRest,
                    .ModelOrTag = "gpt-4o-mini-transcribe-2025-12-15"
                })

                _engines.Add(New EngineDescriptor With {
                    .DisplayName = "OpenAI gpt-4o-transcribe-diarize (REST)",
                    .Kind = EngineKind.OpenAiRest,
                    .ModelOrTag = "gpt-4o-transcribe-diarize"
                })

                _engines.Add(New EngineDescriptor With {
                    .DisplayName = "OpenAI whisper-1 (REST legacy)",
                    .Kind = EngineKind.OpenAiRest,
                    .ModelOrTag = "whisper-1"
                })

                _engines.Add(New EngineDescriptor With {
                    .DisplayName = "OpenAI Realtime Whisper (streaming)",
                    .Kind = EngineKind.OpenAiRealtime,
                    .ModelOrTag = "gpt-realtime-whisper"
                })
            End If

            If HasConfiguredAzureProvider() Then
                _engines.Add(New EngineDescriptor With {
                    .DisplayName = AzureSpeechRealtimeEngine.DisplayNameValue,
                    .Kind = EngineKind.AzureSpeechRealtime,
                    .ModelOrTag = "azure-speech-realtime"
                })

                _engines.Add(New EngineDescriptor With {
                    .DisplayName = AzureSpeechFastRestEngine.DisplayNameValue,
                    .Kind = EngineKind.AzureSpeechFastRest,
                    .ModelOrTag = "azure-speech-fast-rest"
                })
            End If

            If HasConfiguredTeamsAcsProvider() Then
                _engines.Add(New EngineDescriptor With {
                    .DisplayName = TeamsAcsRealtimeEngine.DisplayNameValue,
                    .Kind = EngineKind.TeamsAcsRealtime,
                    .ModelOrTag = "teams-acs-realtime"
                })
            End If

            For Each e In _engines
                cboEngine.Items.Add(e.DisplayName)
            Next

            If cboEngine.Items.Count = 0 Then
                ShowCustomMessageBox("No transcription engines available. Install Vosk/Whisper models or configure Google/OpenAI/Azure transcription.")
                Me.BeginInvoke(Sub() Me.Close())
                Return
            End If

            RestoreLastEngineSelection()
        End Sub

        Private Shared Function EndpointMatchesProvider(endpoint As String, providerIdentifier As String) As Boolean
            If String.IsNullOrWhiteSpace(endpoint) OrElse String.IsNullOrWhiteSpace(providerIdentifier) Then
                Return False
            End If

            Return endpoint.IndexOf(providerIdentifier, StringComparison.OrdinalIgnoreCase) >= 0
        End Function

        Private Sub LoadAlternateProviderFallbacks()
            _alternateOpenAiConfig = Nothing
            _alternateGoogleConfig = Nothing

            Try
                Dim altPath As String = ExpandEnvironmentVariables(Globals.ThisAddIn.INI_AlternateModelPath)

                If String.IsNullOrWhiteSpace(altPath) OrElse Not File.Exists(altPath) Then
                    Return
                End If

                Dim models As List(Of ModelConfig) = SLib.LoadAlternativeModels(
                    altPath,
                    ThisAddIn._context,
                    "Transcriptor Fallback",
                    includeToolOnly:=True,
                    toolsOnly:=False)

                If models Is Nothing OrElse models.Count = 0 Then
                    Return
                End If

                _alternateOpenAiConfig = models.FirstOrDefault(Function(m) IsUsableOpenAiConfig(m))
                _alternateGoogleConfig = models.FirstOrDefault(Function(m) IsUsableGoogleConfig(m))
            Catch
            End Try
        End Sub

        Private Function HasConfiguredGoogleV1Provider() As Boolean
            Return (EndpointMatchesProvider(INI_Endpoint, GoogleIdentifier) AndAlso INI_OAuth2) OrElse
                   (EndpointMatchesProvider(INI_Endpoint_2, GoogleIdentifier) AndAlso INI_OAuth2_2) OrElse
                   IsUsableGoogleConfig(_alternateGoogleConfig)
        End Function

        Private Function HasConfiguredGoogleV2Provider() As Boolean
            Return HasConfiguredGoogleV1Provider() AndAlso
                   Not String.IsNullOrWhiteSpace(ResolveGoogleProjectId())
        End Function

        Private Function HasConfiguredAzureProvider() As Boolean
            Return Not String.IsNullOrWhiteSpace(ResolveAzureSpeechKey())
        End Function

        Private Function HasConfiguredTeamsAcsProvider() As Boolean
            Return Not String.IsNullOrWhiteSpace(ResolveTeamsAcsBridgeWebSocketUri())
        End Function

        Private Function ResolveTeamsAcsBridgeWebSocketUri() As String
            Return NormalizeIniValue(If(ACS_Bridge_Address, ""))
        End Function

        Private Function ResolveTeamsMeetingJoinUrl() As String
            Return NormalizeIniValue(If(ACS_Bridge_Address, ""))
        End Function

        Private Function ResolveTeamsAcsBridgeBearerToken() As String
            Return DecodeWrappedEncryptedValue(
                NormalizeIniValue(INI_Model_Parameter3),
                "Teams ACS bridge bearer token")
        End Function

        Private Shared Function EngineNeedsLocalAudioCapture(kind As EngineKind) As Boolean
            Select Case kind
                Case EngineKind.TeamsAcsRealtime
                    Return False
                Case Else
                    Return True
            End Select
        End Function

        Private Function HasConfiguredOpenAiProvider() As Boolean
            Return EndpointMatchesProvider(INI_Endpoint, OpenAIIdentifier) OrElse
                   EndpointMatchesProvider(INI_Endpoint_2, OpenAIIdentifier) OrElse
                   IsUsableOpenAiConfig(_alternateOpenAiConfig)
        End Function

        Private Function ResolveAzureRegionForHeader(modelOrTag As String) As String
            Return ResolveAzureSttSetting(modelOrTag, "region", "")
        End Function

        Private Function ResolveAzureRealtimeLocation(modelOrTag As String) As String
            Dim region As String = ResolveAzureRegionForHeader(modelOrTag)
            If Not String.IsNullOrWhiteSpace(region) Then
                Return region
            End If

            Return ResolveAzureSttSetting(modelOrTag, "endpoint", "")
        End Function

        Private Function ResolveAzureFastRestLocation(modelOrTag As String) As String
            Dim endpoint As String = ResolveAzureSttSetting(modelOrTag, "endpoint", "")
            If Not String.IsNullOrWhiteSpace(endpoint) Then
                Return endpoint
            End If

            Return ResolveAzureSttSetting(modelOrTag, "region", "")
        End Function

        Private Function BuildConfiguredGoogleModelConfig(useSecond As Boolean) As ModelConfig
            Dim endpoint As String = If(useSecond, INI_Endpoint_2, INI_Endpoint)
            Dim oauthEnabled As Boolean = If(useSecond, INI_OAuth2_2, INI_OAuth2)

            If Not EndpointMatchesProvider(endpoint, GoogleIdentifier) OrElse Not oauthEnabled Then
                Return Nothing
            End If

            Return New ModelConfig With {
                .Endpoint = endpoint,
                .OAuth2 = oauthEnabled,
                .OAuth2ClientMail = If(useSecond, INI_OAuth2ClientMail_2, INI_OAuth2ClientMail),
                .OAuth2Scopes = If(useSecond, INI_OAuth2Scopes_2, INI_OAuth2Scopes),
                .OAuth2Endpoint = If(useSecond, INI_OAuth2Endpoint_2, INI_OAuth2Endpoint),
                .OAuth2ATExpiry = If(useSecond, INI_OAuth2ATExpiry_2, INI_OAuth2ATExpiry),
                .APIKey = If(useSecond, INI_APIKey_2, INI_APIKey)
            }
        End Function

        Private Function BuildConfiguredOpenAiModelConfig(useSecond As Boolean) As ModelConfig
            Dim endpoint As String = If(useSecond, INI_Endpoint_2, INI_Endpoint)

            If Not EndpointMatchesProvider(endpoint, OpenAIIdentifier) Then
                Return Nothing
            End If

            Return New ModelConfig With {
                .Endpoint = endpoint,
                .APIKey = If(useSecond, INI_APIKey_2, INI_APIKey),
                .DecodedAPI = If(useSecond, DecodedAPI_2, DecodedAPI)
            }
        End Function

        Private Function GetApiKeyFromModelConfig(config As ModelConfig) As String
            If config Is Nothing Then
                Return ""
            End If

            If Not String.IsNullOrWhiteSpace(config.DecodedAPI) Then
                Return config.DecodedAPI.Trim()
            End If

            Return If(config.APIKey, "").Trim()
        End Function

        Private Function IsUsableOpenAiConfig(config As ModelConfig) As Boolean
            If config Is Nothing Then
                Return False
            End If

            If Not EndpointMatchesProvider(config.Endpoint, OpenAIIdentifier) Then
                Return False
            End If

            Return Not String.IsNullOrWhiteSpace(GetApiKeyFromModelConfig(config))
        End Function

        Private Function IsUsableGoogleConfig(config As ModelConfig) As Boolean
            If config Is Nothing Then
                Return False
            End If

            If Not EndpointMatchesProvider(config.Endpoint, GoogleIdentifier) Then
                Return False
            End If

            If Not config.OAuth2 Then
                Return False
            End If

            If String.IsNullOrWhiteSpace(config.OAuth2ClientMail) Then
                Return False
            End If

            If String.IsNullOrWhiteSpace(config.OAuth2Endpoint) Then
                Return False
            End If

            Return Not String.IsNullOrWhiteSpace(config.APIKey)
        End Function

        Private Function ResolveGoogleTranscriptionConfig(ByRef cacheSlot As String) As ModelConfig
            Dim primaryConfig As ModelConfig = BuildConfiguredGoogleModelConfig(False)
            If IsUsableGoogleConfig(primaryConfig) Then
                cacheSlot = "primary"
                Return primaryConfig
            End If

            Dim secondaryConfig As ModelConfig = BuildConfiguredGoogleModelConfig(True)
            If IsUsableGoogleConfig(secondaryConfig) Then
                cacheSlot = "secondary"
                Return secondaryConfig
            End If

            If IsUsableGoogleConfig(_alternateGoogleConfig) Then
                cacheSlot = "alternate"
                Return _alternateGoogleConfig
            End If

            cacheSlot = ""
            Return Nothing
        End Function

        Private Function ResolveOpenAiConfig() As ModelConfig
            Dim primaryConfig As ModelConfig = BuildConfiguredOpenAiModelConfig(False)
            If IsUsableOpenAiConfig(primaryConfig) Then
                Return primaryConfig
            End If

            Dim secondaryConfig As ModelConfig = BuildConfiguredOpenAiModelConfig(True)
            If IsUsableOpenAiConfig(secondaryConfig) Then
                Return secondaryConfig
            End If

            If IsUsableOpenAiConfig(_alternateOpenAiConfig) Then
                Return _alternateOpenAiConfig
            End If

            Return Nothing
        End Function

        Private Sub RefreshEngineUi()
            OnEngineChanged(Me, EventArgs.Empty)
        End Sub

        Private Sub OnEngineChanged(sender As Object, e As EventArgs)
            _suspendSettingsPersistence = True

            Try
                cboLang.Items.Clear()

                Dim d As EngineDescriptor = CurrentDescriptor()
                If d Is Nothing Then
                    Return
                End If

                _currentEngineDisplayName = d.DisplayName

                Select Case d.Kind
                    Case EngineKind.Vosk
                        cboLang.Items.Add("(language comes from selected Vosk model)")
                        cboLang.SelectedIndex = 0

                    Case EngineKind.WhisperLocal
                        cboLang.Items.AddRange(
                            WhisperEngine.SupportedLanguages.
                                OrderBy(Function(x) If(String.Equals(x, "auto", StringComparison.OrdinalIgnoreCase), "", x), StringComparer.OrdinalIgnoreCase).
                                Select(Function(x) CObj(x)).
                                ToArray())
                        If cboLang.Items.Contains("auto") Then
                            cboLang.SelectedItem = "auto"
                        ElseIf cboLang.Items.Count > 0 Then
                            cboLang.SelectedIndex = 0
                        End If

                    Case EngineKind.GoogleV1
                        cboLang.Items.AddRange(
                            GoogleV1Engine.SupportedLanguages.
                                OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase).
                                Select(Function(x) CObj(x)).
                                ToArray())
                        If cboLang.Items.Count > 0 Then
                            cboLang.SelectedIndex = 0
                        End If

                    Case EngineKind.GoogleV2
                        cboLang.Items.AddRange(
                            GoogleV2Engine.SupportedLanguages.
                                OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase).
                                Select(Function(x) CObj(x)).
                                ToArray())
                        If cboLang.Items.Count > 0 Then
                            cboLang.SelectedIndex = 0
                        End If

                    Case EngineKind.OpenAiRest
                        cboLang.Items.AddRange(
                            OpenAiRestEngine.SupportedLanguages.
                                OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase).
                                Select(Function(x) CObj(x)).
                                ToArray())
                        If cboLang.Items.Count > 0 Then
                            cboLang.SelectedIndex = 0
                        End If

                    Case EngineKind.OpenAiRealtime
                        cboLang.Items.AddRange(
                            OpenAiRealtimeEngine.SupportedLanguages.
                                OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase).
                                Select(Function(x) CObj(x)).
                                ToArray())
                        If cboLang.Items.Count > 0 Then
                            cboLang.SelectedIndex = 0
                        End If

                    Case EngineKind.AzureSpeechRealtime
                        cboLang.Items.AddRange(
                            AzureSpeechRealtimeEngine.SupportedLanguages.
                                OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase).
                                Select(Function(x) CObj(x)).
                                ToArray())
                        If cboLang.Items.Count > 0 Then
                            cboLang.SelectedIndex = 0
                        End If

                    Case EngineKind.AzureSpeechFastRest
                        cboLang.Items.AddRange(
                            AzureSpeechFastRestEngine.SupportedLanguages.
                                OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase).
                                Select(Function(x) CObj(x)).
                                ToArray())
                        If cboLang.Items.Count > 0 Then
                            cboLang.SelectedIndex = 0
                        End If

                    Case EngineKind.TeamsAcsRealtime
                        cboLang.Items.AddRange(
                            TeamsAcsRealtimeEngine.SupportedLanguages.
                                OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase).
                                Select(Function(x) CObj(x)).
                                ToArray())
                        If cboLang.Items.Count > 0 Then
                            cboLang.SelectedIndex = 0
                        End If

                End Select

                Dim savedLanguage As String = GetSavedLanguageForDescriptor(d)
                If Not String.IsNullOrWhiteSpace(savedLanguage) AndAlso cboLang.Items.Contains(savedLanguage) Then
                    cboLang.SelectedItem = savedLanguage
                End If
            Finally
                _suspendSettingsPersistence = False
            End Try

            UpdateComboToolTip(cboEngine)
            UpdateComboToolTip(cboLang)

            If Not _capturing AndAlso Not _fileTranscribing Then
                SetLiveState(GetIdleLiveState())
            End If
        End Sub

        Private Shared Function IsAzureSpeechLocation(value As String) As Boolean
            Dim normalized As String = If(value, "").Trim()

            If String.IsNullOrWhiteSpace(normalized) Then
                Return False
            End If

            If normalized.IndexOf(".cognitiveservices.azure.com", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return True
            End If

            If normalized.IndexOf(".api.cognitive.microsoft.com", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return True
            End If

            If normalized.StartsWith("http://", StringComparison.OrdinalIgnoreCase) OrElse
               normalized.StartsWith("https://", StringComparison.OrdinalIgnoreCase) Then
                Return False
            End If

            For Each ch As Char In normalized
                If Not Char.IsLetterOrDigit(ch) AndAlso ch <> "-"c Then
                    Return False
                End If
            Next

            Return True
        End Function

        Private Function IsUsableAzureConfig(config As ModelConfig) As Boolean
            If config Is Nothing Then
                Return False
            End If

            If Not IsAzureSpeechLocation(config.Endpoint) Then
                Return False
            End If

            Return Not String.IsNullOrWhiteSpace(GetApiKeyFromModelConfig(config))
        End Function

        Private Function BuildConfiguredAzureModelConfig(useSecond As Boolean) As ModelConfig
            Dim endpointOrRegion As String = If(useSecond, INI_Endpoint_2, INI_Endpoint)

            If Not IsAzureSpeechLocation(endpointOrRegion) Then
                Return Nothing
            End If

            Return New ModelConfig With {
                .Endpoint = endpointOrRegion,
                .APIKey = If(useSecond, INI_APIKey_2, INI_APIKey),
                .DecodedAPI = If(useSecond, DecodedAPI_2, DecodedAPI)
            }
        End Function

        Private Function ResolveAzureConfig() As ModelConfig
            Dim primaryConfig As ModelConfig = BuildConfiguredAzureModelConfig(False)
            If IsUsableAzureConfig(primaryConfig) Then
                Return primaryConfig
            End If

            Dim secondaryConfig As ModelConfig = BuildConfiguredAzureModelConfig(True)
            If IsUsableAzureConfig(secondaryConfig) Then
                Return secondaryConfig
            End If

            Return Nothing
        End Function


        Private Function FindEngineIndexByDisplayName(displayName As String) As Integer
            If String.IsNullOrWhiteSpace(displayName) Then
                Return -1
            End If

            Dim normalized As String = displayName.Trim()

            For i As Integer = 0 To _engines.Count - 1
                If String.Equals(_engines(i).DisplayName, normalized, StringComparison.OrdinalIgnoreCase) Then
                    Return i
                End If
            Next

            Return -1
        End Function

        Private Sub RestoreLastEngineSelection()
            Dim engineIndex As Integer = FindEngineIndexByDisplayName(My.Settings.LastEngineName)

            If engineIndex >= 0 Then
                cboEngine.SelectedIndex = engineIndex
                Return
            End If

            If cboEngine.Items.Count > 0 AndAlso cboEngine.SelectedIndex < 0 Then
                cboEngine.SelectedIndex = 0
            End If
        End Sub

        Private Sub SaveCurrentEngineSelection()
            Dim d As EngineDescriptor = CurrentDescriptor()
            If d Is Nothing Then
                Return
            End If

            My.Settings.LastEngineName = d.DisplayName
        End Sub

        Private Function CurrentDescriptor() As EngineDescriptor
            If cboEngine.SelectedIndex < 0 Then
                Return Nothing
            End If
            Return _engines(cboEngine.SelectedIndex)
        End Function

        Private Sub LoadAudioInputDevices()
            cboDevice.Items.Clear()
            _audioInputDevices.Clear()

            Try
                Dim enumr As New MMDeviceEnumerator()
                Dim devs = enumr.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active)

                For Each d In devs
                    Dim friendlyName As String = If(d.FriendlyName, "").Trim()
                    If friendlyName.Length = 0 Then
                        friendlyName = d.ID
                    End If

                    _audioInputDevices.Add(New AudioInputDeviceChoice With {
                        .DeviceId = d.ID,
                        .WaveDeviceIndex = FindLegacyWaveInputDeviceIndex(friendlyName),
                        .DisplayText = friendlyName,
                        .ToolTipText = friendlyName
                    })
                Next
            Catch
            End Try

            If _audioInputDevices.Count = 0 Then
                For i As Integer = 0 To WaveInEvent.DeviceCount - 1
                    Dim productName As String = WaveInEvent.GetCapabilities(i).ProductName

                    _audioInputDevices.Add(New AudioInputDeviceChoice With {
                        .DeviceId = "",
                        .WaveDeviceIndex = i,
                        .DisplayText = $"{i}: {productName}",
                        .ToolTipText = productName
                    })
                Next
            End If

            For Each choice As AudioInputDeviceChoice In _audioInputDevices
                cboDevice.Items.Add(choice)
            Next

            If cboDevice.Items.Count > 0 Then
                cboDevice.SelectedIndex = 0
            End If
        End Sub

        Private Shared Function FindLegacyWaveInputDeviceIndex(friendlyName As String) As Integer
            Dim normalizedFriendlyName As String = If(friendlyName, "").Trim()

            For i As Integer = 0 To WaveInEvent.DeviceCount - 1
                Dim waveName As String = If(WaveInEvent.GetCapabilities(i).ProductName, "").Trim()

                If String.Equals(waveName, normalizedFriendlyName, StringComparison.OrdinalIgnoreCase) Then
                    Return i
                End If

                If normalizedFriendlyName.IndexOf(waveName, StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                   waveName.IndexOf(normalizedFriendlyName, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Return i
                End If
            Next

            Return 0
        End Function

        Private Function GetSelectedAudioInputDeviceChoice() As AudioInputDeviceChoice
            Return TryCast(cboDevice.SelectedItem, AudioInputDeviceChoice)
        End Function

        Private Sub RestoreAudioInputSelection()
            Dim preferredDeviceId As String = ""

            If _opts IsNot Nothing Then
                preferredDeviceId = If(_opts.PreferredMicrophoneDeviceId, "").Trim()
            End If

            If preferredDeviceId.Length > 0 Then
                For i As Integer = 0 To _audioInputDevices.Count - 1
                    If String.Equals(_audioInputDevices(i).DeviceId, preferredDeviceId, StringComparison.OrdinalIgnoreCase) Then
                        cboDevice.SelectedIndex = i
                        Return
                    End If
                Next
            End If

            If My.Settings.LastAudioInputDeviceIndex >= 0 AndAlso My.Settings.LastAudioInputDeviceIndex < cboDevice.Items.Count Then
                cboDevice.SelectedIndex = My.Settings.LastAudioInputDeviceIndex
                Return
            End If

            If cboDevice.Items.Count > 0 Then
                cboDevice.SelectedIndex = 0
            End If
        End Sub

        Private Sub UpdateSelectedMicrophonePreference()
            If _opts Is Nothing Then
                _opts = New TranscriptionOptions()
            End If

            Dim choice As AudioInputDeviceChoice = GetSelectedAudioInputDeviceChoice()

            If choice Is Nothing Then
                _opts.PreferredMicrophoneDeviceId = ""
                _opts.PreferredMicrophoneDisplayName = ""
                Return
            End If

            _opts.PreferredMicrophoneDeviceId = choice.DeviceId
            _opts.PreferredMicrophoneDisplayName = choice.ToolTipText
        End Sub

        Private Function GetAudioOutputDeviceChoices() As List(Of KeyValuePair(Of String, String))
            Dim result As New List(Of KeyValuePair(Of String, String)) From {
                New KeyValuePair(Of String, String)("Default Audio Output Device", "")
            }

            Dim enumr As New MMDeviceEnumerator()
            Dim devs = enumr.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active)

            For Each d In devs
                result.Add(New KeyValuePair(Of String, String)(d.FriendlyName, d.ID))
            Next

            Return result
        End Function

        Private Function GetConfiguredSourceMode() As AudioSourceMode
            Dim raw As String = If(String.IsNullOrWhiteSpace(My.Settings.LastAudioSourceMode), "MicrophoneOnly", My.Settings.LastAudioSourceMode)

            Try
                Return CType([Enum].Parse(GetType(AudioSourceMode), raw), AudioSourceMode)
            Catch
                Return AudioSourceMode.MicrophoneOnly
            End Try
        End Function

        Private Function GetConfiguredOutputDeviceId() As String
            Return If(My.Settings.LastAudioOutputDeviceId, "")
        End Function

        Private Sub RestoreSettings()
            Try
                RestoreLastEngineSelection()

                RefreshEngineUi()

                If cboLang.Items.Count > 0 AndAlso cboLang.SelectedIndex < 0 Then
                    cboLang.SelectedIndex = 0
                End If

                If Not String.IsNullOrEmpty(My.Settings.LastEngineOptionsJson) Then
                    Try
                        _opts = JsonConvert.DeserializeObject(Of TranscriptionOptions)(My.Settings.LastEngineOptionsJson)
                    Catch
                    End Try
                End If

                RestoreAudioInputSelection()
            Catch
            End Try

            SetLiveState(GetIdleLiveState())
        End Sub

        Private Sub PersistSettings()
            If Me.IsDisposed OrElse _suspendSettingsPersistence Then
                Return
            End If

            If Me.InvokeRequired Then
                Try
                    Me.BeginInvoke(New System.Action(AddressOf PersistSettings))
                Catch
                End Try
                Return
            End If

            Try
                SaveCurrentEngineSelection()
                UpdateSelectedMicrophonePreference()
                My.Settings.LastAudioInputDeviceIndex = cboDevice.SelectedIndex
                My.Settings.LastEngineOptionsJson = JsonConvert.SerializeObject(_opts)
                SaveCurrentLanguageForCurrentEngine()
                My.Settings.Save()
            Catch
            End Try
        End Sub


        Private Sub OnOptions(sender As Object, e As EventArgs)
            Dim d As EngineDescriptor = CurrentDescriptor()
            If d Is Nothing Then
                Return
            End If

            Dim langs As String() = (From o In cboLang.Items Select CStr(o)).ToArray()

            Using dlg As New TranscriptionOptionsDialog(
                d.Kind,
                d.DisplayName,
                _opts,
                langs,
                If(String.IsNullOrWhiteSpace(My.Settings.LastAudioSourceMode), "MicrophoneOnly", My.Settings.LastAudioSourceMode),
                GetConfiguredOutputDeviceId(),
                GetAudioOutputDeviceChoices())

                If dlg.ShowDialog(Me) = DialogResult.OK Then
                    _opts = dlg.Options

                    If Not String.IsNullOrWhiteSpace(_opts.LanguageCode) AndAlso cboLang.Items.Contains(_opts.LanguageCode) Then
                        cboLang.SelectedItem = _opts.LanguageCode
                    End If

                    My.Settings.LastAudioSourceMode = dlg.SelectedSourceMode
                    My.Settings.LastAudioOutputDeviceId = dlg.SelectedOutputDeviceId
                    PersistSettings()
                    SetLiveState(GetIdleLiveState())
                End If
            End Using
        End Sub

        Private Async Function CreateEngineAsync(d As EngineDescriptor) As Task(Of ITranscriptionEngine)
            Dim modelRoot As String = ExpandEnvironmentVariables(Globals.ThisAddIn.INI_SpeechModelPath)

            Select Case d.Kind
                Case EngineKind.Vosk
                    Return New VoskEngine(modelRoot, d.ModelOrTag)

                Case EngineKind.WhisperLocal
                    Return New WhisperEngine(modelRoot, d.ModelOrTag)

                Case EngineKind.GoogleV1
                    Dim googleCacheSlot As String = ""
                    Dim googleConfig As ModelConfig = ResolveGoogleTranscriptionConfig(googleCacheSlot)

                    If googleConfig Is Nothing Then
                        Throw New InvalidOperationException("No Google transcription credentials are available.")
                    End If

                    _opts.Model = ResolveGoogleSttSetting(d.ModelOrTag, "model", "")

                    Dim tf As Func(Of System.Threading.Tasks.Task(Of String)) =
                        Function() GetFreshGoogleTokenAsync(googleConfig, googleCacheSlot)

                    Return New GoogleV1Engine("", tf)

                Case EngineKind.GoogleV2
                    Dim googleCacheSlot As String = ""
                    Dim googleConfig As ModelConfig = ResolveGoogleTranscriptionConfig(googleCacheSlot)

                    If googleConfig Is Nothing Then
                        Throw New InvalidOperationException("No Google transcription credentials are available.")
                    End If

                    If String.IsNullOrWhiteSpace(ResolveGoogleProjectId()) Then
                        Throw New InvalidOperationException("INI_STT_Google_ProjectID is missing.")
                    End If

                    Dim dbgProjectIdRaw As String = If(INI_STT_Google_ProjectID, "")
                    Dim dbgProjectIdResolved As String = ResolveGoogleProjectId()
                    Dim dbgIniSttGoogleRaw As String =
                        If(INI_STT_Google, "").
                            Replace(vbCrLf, "\n").
                            Replace(vbCr, "\n").
                            Replace(vbLf, "\n")
                    Dim dbgResolvedEndpoint As String = ResolveGoogleSttSetting(d.ModelOrTag, "endpoint", "")
                    Dim dbgResolvedLocation As String = ResolveGoogleSttSetting(d.ModelOrTag, "location", "")
                    Dim dbgResolvedRecognizer As String = ResolveGoogleSttSetting(d.ModelOrTag, "recognizer", "")
                    Dim dbgResolvedModel As String = ResolveGoogleSttSetting(d.ModelOrTag, "model", "")
                    Dim dbgResolvedLanguage As String = ResolveGoogleSttSetting(d.ModelOrTag, "language", "")

                    System.Diagnostics.Debug.WriteLine("[Transcriptor.GoogleV2] _context Is Nothing=" & (ThisAddIn._context Is Nothing).ToString())
                    System.Diagnostics.Debug.WriteLine("[Transcriptor.GoogleV2] Codebasis present=" & (ThisAddIn._context IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(ThisAddIn._context.Codebasis)).ToString())
                    System.Diagnostics.Debug.WriteLine("[Transcriptor.GoogleV2] CacheSlot=" & googleCacheSlot)
                    System.Diagnostics.Debug.WriteLine("[Transcriptor.GoogleV2] OAuthClientMail=" & If(googleConfig.OAuth2ClientMail, ""))
                    System.Diagnostics.Debug.WriteLine("[Transcriptor.GoogleV2] OAuthTokenEndpoint=" & If(googleConfig.OAuth2Endpoint, ""))
                    System.Diagnostics.Debug.WriteLine("[Transcriptor.GoogleV2] APIKeyLength=" & If(googleConfig.APIKey, "").Length.ToString())
                    System.Diagnostics.Debug.WriteLine("[Transcriptor.GoogleV2] INI_STT_Google_ProjectID raw=" & dbgProjectIdRaw)
                    System.Diagnostics.Debug.WriteLine("[Transcriptor.GoogleV2] ResolveGoogleProjectId()=" & dbgProjectIdResolved)
                    System.Diagnostics.Debug.WriteLine("[Transcriptor.GoogleV2] INI_STT_Google raw=" & dbgIniSttGoogleRaw)
                    System.Diagnostics.Debug.WriteLine(
                        "[Transcriptor.GoogleV2] Resolved STT settings " &
                        "endpoint=" & dbgResolvedEndpoint &
                        "; location=" & dbgResolvedLocation &
                        "; recognizer=" & dbgResolvedRecognizer &
                        "; model=" & dbgResolvedModel &
                        "; language=" & dbgResolvedLanguage)



                    Return New GoogleV2Engine(
                        googleConfig.OAuth2ClientMail,
                        googleConfig.APIKey,
                        googleConfig.OAuth2Endpoint,
                        ResolveGoogleProjectId(),
                        ResolveGoogleSttSetting(d.ModelOrTag, "endpoint", ""),
                        ResolveGoogleSttSetting(d.ModelOrTag, "location", ""),
                        ResolveGoogleSttSetting(d.ModelOrTag, "recognizer", ""),
                        ResolveGoogleSttSetting(d.ModelOrTag, "model", ""),
                        ResolveGoogleSttSetting(d.ModelOrTag, "language", ""),
                        googleConfig.OAuth2Scopes,
                        "Transcriptor")

                Case EngineKind.OpenAiRest
                    Dim key As String = ResolveOpenAiKey()
                    If String.IsNullOrWhiteSpace(key) Then
                        Throw New InvalidOperationException("No OpenAI API key is available.")
                    End If

                    _opts.Model = ResolveOpenAiSttSetting(d.ModelOrTag, "model", d.ModelOrTag)
                    Return New OpenAiRestEngine(key)

                Case EngineKind.OpenAiRealtime
                    Dim key As String = ResolveOpenAiKey()
                    If String.IsNullOrWhiteSpace(key) Then
                        Throw New InvalidOperationException("No OpenAI API key is available.")
                    End If

                    _opts.Model = ResolveOpenAiSttSetting(d.ModelOrTag, "model", "gpt-realtime-whisper")
                    Return New OpenAiRealtimeEngine(key)

                Case EngineKind.AzureSpeechRealtime
                    Dim azureSpeechKey As String = ResolveAzureSpeechKey()
                    If String.IsNullOrWhiteSpace(azureSpeechKey) Then
                        Throw New InvalidOperationException("INI_STT_Azure_SpeechKey is missing.")
                    End If

                    Return New AzureSpeechRealtimeEngine(
                        azureSpeechKey,
                        ResolveAzureRealtimeLocation(d.ModelOrTag),
                        ResolveAzureRegionForHeader(d.ModelOrTag))

                Case EngineKind.AzureSpeechFastRest
                    Dim azureSpeechKey As String = ResolveAzureSpeechKey()
                    If String.IsNullOrWhiteSpace(azureSpeechKey) Then
                        Throw New InvalidOperationException("INI_STT_Azure_SpeechKey is missing.")
                    End If

                    Return New AzureSpeechFastRestEngine(
                        azureSpeechKey,
                        ResolveAzureFastRestLocation(d.ModelOrTag),
                        ResolveAzureSttSetting(d.ModelOrTag, "api-version", ""),
                        ResolveAzureRegionForHeader(d.ModelOrTag))

                Case EngineKind.TeamsAcsRealtime
                    Return New TeamsAcsRealtimeEngine(
                        ResolveTeamsAcsBridgeWebSocketUri(),
                        ResolveTeamsMeetingJoinUrl(),
                        ResolveTeamsAcsBridgeBearerToken())

            End Select

            Throw New NotSupportedException(d.Kind.ToString())
        End Function

        Private Function ResolveOpenAiKey() As String
            Dim config As ModelConfig = ResolveOpenAiConfig()
            Return GetApiKeyFromModelConfig(config)
        End Function

        Private _g1Token As String = ""
        Private _g2Token As String = ""
        Private _gAltToken As String = ""
        Private _g1Exp As DateTime = DateTime.MinValue
        Private _g2Exp As DateTime = DateTime.MinValue
        Private _gAltExp As DateTime = DateTime.MinValue

        Private Async Function GetFreshGoogleTokenAsync(config As ModelConfig, cacheSlot As String) As Task(Of String)
            If config Is Nothing Then
                Return ""
            End If

            Dim token As String = ""
            Dim exp As DateTime = DateTime.MinValue

            Select Case cacheSlot
                Case "secondary"
                    token = _g2Token
                    exp = _g2Exp

                Case "alternate"
                    token = _gAltToken
                    exp = _gAltExp

                Case Else
                    token = _g1Token
                    exp = _g1Exp
            End Select

            If String.IsNullOrEmpty(token) OrElse DateTime.UtcNow >= exp Then
                Dim life As Long = If(config.OAuth2ATExpiry > 0, config.OAuth2ATExpiry, 3600)

                GoogleOAuthHelper.client_email = config.OAuth2ClientMail
                GoogleOAuthHelper.private_key = FormatPrivateKey(config.APIKey)
                GoogleOAuthHelper.scopes = config.OAuth2Scopes
                GoogleOAuthHelper.token_uri = config.OAuth2Endpoint
                GoogleOAuthHelper.token_life = life

                token = Await GoogleOAuthHelper.GetAccessToken()
                Dim newExp As DateTime = DateTime.UtcNow.AddSeconds(Math.Max(60L, life - 300L))

                Select Case cacheSlot
                    Case "secondary"
                        _g2Token = token
                        _g2Exp = newExp

                    Case "alternate"
                        _gAltToken = token
                        _gAltExp = newExp

                    Case Else
                        _g1Token = token
                        _g1Exp = newExp
                End Select
            End If

            Return token
        End Function

        Public Shared Function FormatPrivateKey(rawKey As String) As String
            Dim noEsc As String = rawKey.Replace("\n", "")
            Dim sb As New StringBuilder()

            For i As Integer = 0 To noEsc.Length - 1 Step 64
                sb.AppendLine(If(i + 64 <= noEsc.Length, noEsc.Substring(i, 64), noEsc.Substring(i)))
            Next

            Return "-----BEGIN PRIVATE KEY-----" & vbLf & sb.ToString() & "-----END PRIVATE KEY-----" & vbLf
        End Function


        Private Function GetFileTranscribingState() As String
            If String.IsNullOrWhiteSpace(_currentEngineDisplayName) Then
                Return "Transcribing file…"
            End If

            Return _currentEngineDisplayName & ": Transcribing file…"
        End Function

        Private Sub CancelCurrentFileTranscription()
            If Not _fileTranscribing Then
                Return
            End If

            SetLiveState(_currentEngineDisplayName & ": Canceling file transcription…")

            If _cts IsNot Nothing Then
                Try
                    _cts.Cancel()
                Catch
                End Try
            End If
        End Sub


        Private Function GetIdleLiveState() As String
            If String.IsNullOrWhiteSpace(_currentEngineDisplayName) Then
                Return "Ready."
            End If

            Return _currentEngineDisplayName & ": Ready."
        End Function

        Private Function GetListeningLiveState() As String
            If String.IsNullOrWhiteSpace(_currentEngineDisplayName) Then
                Return "Listening…"
            End If

            Return _currentEngineDisplayName & ": Listening…"
        End Function

        Private Sub SetLiveState(text As String)
            If Me.IsDisposed OrElse lblLiveState Is Nothing Then
                Return
            End If

            Dim normalized As String = If(text, "").Trim()

            If String.Equals(_lastLiveStateText, normalized, StringComparison.Ordinal) Then
                If (DateTime.UtcNow - _lastLiveStateUtc).TotalMilliseconds < 500 Then
                    Return
                End If
            End If

            If Me.InvokeRequired Then
                Try
                    Me.BeginInvoke(New System.Action(Of String)(AddressOf SetLiveState), normalized)
                Catch
                End Try
                Return
            End If

            _lastLiveStateText = normalized
            _lastLiveStateUtc = DateTime.UtcNow
            lblLiveState.Text = normalized
            Debug.WriteLine("[Transcriptor] " & normalized)
        End Sub

        Private Async Sub OnStart(sender As Object, e As EventArgs)
            If _capturing OrElse _fileTranscribing OrElse _isStopping Then
                Return
            End If

            Dim d As EngineDescriptor = CurrentDescriptor()
            If d Is Nothing Then
                Return
            End If

            If IsFileOnlyEngine(d.Kind) Then
                ShowCustomMessageBox("Selected engine is file/request-response transcription only. Use Load for an audio file, or select a live engine for microphone transcription.")
                Return
            End If

            Dim selectedLanguage As String = ""
            If cboLang.SelectedItem IsNot Nothing Then
                selectedLanguage = CStr(cboLang.SelectedItem)
            End If

            Dim selectedDevice As AudioInputDeviceChoice = GetSelectedAudioInputDeviceChoice()
            Dim selectedDeviceText As String = If(selectedDevice IsNot Nothing, selectedDevice.DisplayText, If(cboDevice.Text, ""))
            Dim micDeviceIndex As Integer = If(selectedDevice IsNot Nothing, selectedDevice.WaveDeviceIndex, ParseDeviceIndexFromText(selectedDeviceText))
            Dim micDeviceId As String = If(selectedDevice IsNot Nothing, selectedDevice.DeviceId, "")
            Dim sourceMode As AudioSourceMode = GetConfiguredSourceMode()

            UpdateSelectedMicrophonePreference()

            If d.Kind <> EngineKind.Vosk AndAlso Not String.IsNullOrWhiteSpace(selectedLanguage) Then
                _opts.LanguageCode = selectedLanguage
            End If

            _currentEngineDisplayName = d.DisplayName
            SetLiveState(_currentEngineDisplayName & ": Starting…")

            Dim startException As Exception = Nothing
            Dim engineToDisposeAfterStartFailure As ITranscriptionEngine = Nothing
            Dim ctsToDisposeAfterStartFailure As CancellationTokenSource = Nothing

            Try
                _engine = Await CreateEngineAsync(d)
                AttachEngineEvents(_engine)
                _cts = New CancellationTokenSource()

                System.Diagnostics.Debug.WriteLine(
                    "[Transcriptor.Live] About to call StartLiveAsync " &
                    "Engine=" & d.DisplayName &
                    "; LanguageCode=" & If(_opts Is Nothing OrElse String.IsNullOrWhiteSpace(_opts.LanguageCode), "(empty)", _opts.LanguageCode) &
                    "; Model=" & If(_opts Is Nothing OrElse String.IsNullOrWhiteSpace(_opts.Model), "(empty)", _opts.Model) &
                    "; MultiChannelDiarization=" & If(_opts IsNot Nothing AndAlso _opts.MultiChannelDiarization, "True", "False"))

                Await _engine.StartLiveAsync(_opts, _cts.Token)
            Catch ex As Exception
                startException = ex
                engineToDisposeAfterStartFailure = _engine
                ctsToDisposeAfterStartFailure = _cts

                _engine = Nothing
                _cts = Nothing
            End Try

            If ctsToDisposeAfterStartFailure IsNot Nothing Then
                Try
                    ctsToDisposeAfterStartFailure.Dispose()
                Catch
                End Try
            End If

            If engineToDisposeAfterStartFailure IsNot Nothing Then
                Try
                    Await engineToDisposeAfterStartFailure.DisposeAsync()
                Catch
                End Try
            End If

            If startException IsNot Nothing Then
                SetLiveState(_currentEngineDisplayName & ": Error.")
                ShowTranscriptorMessageBox("Failed to start engine: " & startException.Message)
                Return
            End If

            If EngineNeedsLocalAudioCapture(d.Kind) Then
                _capture = New AudioCaptureService With {
                    .MicDeviceIndex = micDeviceIndex,
                    .MicDeviceId = micDeviceId,
                    .SourceMode = sourceMode,
                    .SystemAudioRenderDeviceId = GetConfiguredOutputDeviceId(),
                    .MultiChannelStereo = _opts.MultiChannelDiarization AndAlso _engine.SupportsMultiChannelDiarization,
                    .AudioDebugDump = _opts.AudioDebugDump OrElse INI_APIDebug
                }

                AddHandler _capture.Frame, AddressOf OnCaptureFrame
                AddHandler _capture.CaptureError,
                    Sub(s, ev)
                        SetLiveState(_currentEngineDisplayName & ": Capture error: " & ev.Message)
                    End Sub

                _capture.Start()
            End If

            AcquireSleepLock()

            _capturing = True
            ToggleCaptureUi(True)
            PersistSettings()
            SetLiveState(GetListeningLiveState())
        End Sub

        Private Function ParseDeviceIndexFromText(deviceText As String) As Integer
            If String.IsNullOrWhiteSpace(deviceText) Then
                Return 0
            End If

            Dim c As Integer = deviceText.IndexOf(":"c)
            If c <= 0 Then
                Return 0
            End If

            Dim idx As Integer = 0
            Integer.TryParse(deviceText.Substring(0, c), idx)
            Return idx
        End Function

        Private Async Sub OnCaptureFrame(sender As Object, e As AudioCaptureService.FrameEventArgs)
            Dim eng As ITranscriptionEngine = _engine
            Dim ctsLocal As CancellationTokenSource = _cts

            If eng Is Nothing OrElse Not _capturing OrElse ctsLocal Is Nothing Then
                Return
            End If

            Try
                Await eng.PushAudioAsync(e.Pcm, e.BytesValid, ctsLocal.Token)
            Catch ex As OperationCanceledException
            Catch ex As ObjectDisposedException
            Catch ex As Exception
                If Not _isStopping Then
                    SetLiveState(_currentEngineDisplayName & ": Audio push failed: " & ex.Message)
                End If
            End Try
        End Sub

        Private Async Function StopCurrentSessionAsync() As Task
            If _isStopping Then
                Return
            End If

            _isStopping = True

            Dim captureToStop As AudioCaptureService = _capture
            Dim engineToStop As ITranscriptionEngine = _engine
            Dim ctsToDispose As CancellationTokenSource = _cts

            _capture = Nothing
            _engine = Nothing
            _cts = Nothing
            _capturing = False

            SetLiveState(_currentEngineDisplayName & ": Stopping…")
            ToggleCaptureUi(False)

            Try
                If captureToStop IsNot Nothing Then
                    Try
                        RemoveHandler captureToStop.Frame, AddressOf OnCaptureFrame
                    Catch
                    End Try

                    Try
                        captureToStop.Stop()
                    Catch
                    End Try

                    Try
                        captureToStop.Dispose()
                    Catch
                    End Try
                End If

                If engineToStop IsNot Nothing Then
                    Await engineToStop.StopLiveAsync()
                    Await engineToStop.DisposeAsync()
                End If
            Catch ex As OperationCanceledException
            Catch ex As Exception
                SetLiveState(_currentEngineDisplayName & ": Stop error: " & ex.Message)
            Finally
                If ctsToDispose IsNot Nothing Then
                    Try
                        ctsToDispose.Cancel()
                    Catch
                    End Try

                    Try
                        ctsToDispose.Dispose()
                    Catch
                    End Try
                End If

                ReleaseSleepLock()
                _isStopping = False
                ToggleCaptureUi(False)
                SetLiveState(GetIdleLiveState())

                If _closeAfterStop Then
                    _closeAfterStop = False
                    If Not Me.IsDisposed Then
                        Try
                            Me.BeginInvoke(New System.Action(Sub() Me.Close()))
                        Catch
                        End Try
                    End If
                End If
            End Try
        End Function



        Private Async Sub OnStop(sender As Object, e As EventArgs)
            If _fileTranscribing Then
                CancelCurrentFileTranscription()
                Return
            End If

            Await StopCurrentSessionAsync()
        End Sub


        Private Async Sub OnQuit(sender As Object, e As EventArgs)
            If _capturing Then
                _closeAfterStop = True
                Await StopCurrentSessionAsync()
            ElseIf _fileTranscribing Then
                _closeAfterStop = True
                CancelCurrentFileTranscription()
            Else
                Me.Close()
            End If
        End Sub

        Private Async Sub OnLoadFile(sender As Object, e As EventArgs)
            If _capturing OrElse _fileTranscribing OrElse _isStopping Then
                Return
            End If

            Dim d As EngineDescriptor = CurrentDescriptor()
            If d Is Nothing Then
                Return
            End If

            If IsLiveOnlyEngine(d.Kind) Then
                ShowTranscriptorMessageBox("Realtime engine is live-only; pick another engine for file mode.")
                Return
            End If

            DragDropFormLabel = "Audio file (*.wav, *.mp3, *.aac, *.m4a, *.mp4, *.wma)"
            DragDropFormFilter = "Audio|*.wav;*.mp3;*.aac;*.m4a;*.mp4;*.wma|All|*.*"

            Dim filePath As String = ""

            Try
                Using f As New DragDropForm()
                    If f.ShowDialog() = DialogResult.OK Then
                        filePath = f.SelectedFilePath
                    End If
                End Using
            Finally
                DragDropFormLabel = ""
                DragDropFormFilter = ""
            End Try

            If String.IsNullOrEmpty(filePath) OrElse Not File.Exists(filePath) Then
                Return
            End If

            If cboLang.SelectedItem IsNot Nothing AndAlso d.Kind <> EngineKind.Vosk Then
                _opts.LanguageCode = CStr(cboLang.SelectedItem)
            End If

            _currentEngineDisplayName = d.DisplayName
            _lastPartialText = ""
            _fileTranscribing = True
            ToggleCaptureUi(False)
            PersistSettings()
            SetLiveState(GetFileTranscribingState())

            Dim engineToDispose As ITranscriptionEngine = Nothing
            Dim ctsToDispose As CancellationTokenSource = Nothing
            Dim failed As Boolean = False
            Dim canceled As Boolean = False
            Dim fileTranscriptionException As Exception = Nothing

            Try
                _engine = Await CreateEngineAsync(d)
                AttachEngineEvents(_engine)
                _cts = New CancellationTokenSource()
                Await _engine.TranscribeFileAsync(filePath, _opts, _cts.Token)
                canceled = (_cts IsNot Nothing AndAlso _cts.IsCancellationRequested)
            Catch ex As OperationCanceledException
                canceled = True
            Catch ex As Exception
                failed = True
                fileTranscriptionException = ex
            Finally
                engineToDispose = _engine
                ctsToDispose = _cts

                _engine = Nothing
                _cts = Nothing
                _fileTranscribing = False
                ToggleCaptureUi(False)
            End Try

            If ctsToDispose IsNot Nothing Then
                Try
                    ctsToDispose.Dispose()
                Catch
                End Try
            End If

            If engineToDispose IsNot Nothing Then
                Try
                    Await engineToDispose.DisposeAsync()
                Catch
                End Try
            End If

            If fileTranscriptionException IsNot Nothing Then
                SetLiveState(_currentEngineDisplayName & ": File transcription failed.")
                ShowTranscriptorMessageBox("File transcription failed: " & fileTranscriptionException.Message)
            End If

            If canceled Then
                SetLiveState(_currentEngineDisplayName & ": File transcription canceled.")
            ElseIf Not failed Then
                SetLiveState(_currentEngineDisplayName & ": File transcription complete.")
                ShowTranscriptorMessageBox("File transcription complete.")
            End If

            If _closeAfterStop Then
                _closeAfterStop = False
                If Not Me.IsDisposed Then
                    Try
                        Me.BeginInvoke(New System.Action(Sub() Me.Close()))
                    Catch
                    End Try
                End If
            End If
        End Sub

        Private Sub SafeAppendTranscript(text As String)
            If String.IsNullOrWhiteSpace(text) Then
                Return
            End If

            If Me.IsDisposed Then
                Return
            End If

            If Me.InvokeRequired Then
                Try
                    Me.BeginInvoke(New Action(Of String)(AddressOf SafeAppendTranscript), text)
                Catch
                End Try
                Return
            End If

            rtb.AppendText(text.Trim() & vbCrLf)
            rtb.SelectionStart = rtb.TextLength
            rtb.ScrollToCaret()
        End Sub

        Private Sub AttachEngineEvents(eng As ITranscriptionEngine)
            AddHandler eng.PartialResult,
        Sub(s, ev)
            Dim msg As String = If(String.IsNullOrEmpty(ev.Speaker), ev.Text, ev.Speaker & ": " & ev.Text)
            If Not String.IsNullOrWhiteSpace(msg) Then
                _lastPartialText = msg.Trim()
                SetLiveState(_currentEngineDisplayName & ": " & _lastPartialText)
            End If
        End Sub

            AddHandler eng.FinalResult,
        Sub(s, ev)
            Dim line As String = If(String.IsNullOrEmpty(ev.Speaker), ev.Text, ev.Speaker & ": " & ev.Text)
            SafeAppendTranscript(line)
            _lastPartialText = ""

            If _capturing Then
                SetLiveState(GetListeningLiveState())
            ElseIf _fileTranscribing Then
                SetLiveState(GetFileTranscribingState())
            Else
                SetLiveState(GetIdleLiveState())
            End If
        End Sub

            AddHandler eng.EngineError,
        Sub(s, ev)
            LogSttError(_currentEngineDisplayName, ev.Message, ev.Exception)
            SetLiveState(GetFriendlySttErrorText(_currentEngineDisplayName, ev.Message))
        End Sub

            AddHandler eng.Status,
        Sub(s, ev)
            If String.IsNullOrWhiteSpace(_lastPartialText) Then
                Dim friendly As String = GetFriendlyStatusText(_currentEngineDisplayName, ev.Message)
                If Not String.IsNullOrWhiteSpace(friendly) Then
                    SetLiveState(friendly)
                End If
            End If
        End Sub
        End Sub

        Private Sub ToggleCaptureUi(capturing As Boolean)
            If Me.IsDisposed Then
                Return
            End If

            If Me.InvokeRequired Then
                Try
                    Me.BeginInvoke(New Action(Of Boolean)(AddressOf ToggleCaptureUi), capturing)
                Catch
                End Try
                Return
            End If

            Dim busy As Boolean = capturing OrElse _capturing OrElse _fileTranscribing OrElse _isStopping

            btnStart.Enabled = Not busy
            btnStop.Enabled = _capturing OrElse _fileTranscribing
            btnStop.Text = If(_fileTranscribing, "Cancel", "Stop")
            btnLoad.Enabled = Not busy
            btnOptions.Enabled = Not busy
            cboEngine.Enabled = Not busy
            cboLang.Enabled = Not busy
            cboDevice.Enabled = Not busy
        End Sub

        Private Sub AcquireSleepLock()
            Dim prev As UInteger = SetThreadExecutionState(ES_CONTINUOUS Or ES_SYSTEM_REQUIRED)
            _setSleepLock = ((prev And ES_SYSTEM_REQUIRED) = 0)
        End Sub

        Private Sub ReleaseSleepLock()
            If _setSleepLock Then
                SetThreadExecutionState(ES_CONTINUOUS)
                _setSleepLock = False
            End If
        End Sub

        Private Sub LoadAndPopulatePrompts(filePath As String, combo As ComboBox)
            _promptTitles.Clear()
            _promptBodies.Clear()
            combo.Items.Clear()

            Try
                Dim path As String = ExpandEnvironmentVariables(filePath)
                If Not File.Exists(path) Then
                    Return
                End If

                For Each line As String In File.ReadAllLines(path)
                    Dim t As String = line.Trim()
                    If String.IsNullOrEmpty(t) OrElse t.StartsWith(";") Then
                        Continue For
                    End If

                    Dim parts = t.Split("|"c)
                    If parts.Length >= 2 Then
                        _promptTitles.Add(parts(0).Trim())
                        _promptBodies.Add(String.Join("|", parts.Skip(1)).Trim())
                    End If
                Next

                For Each title As String In _promptTitles
                    combo.Items.Add(title)
                Next
            Catch
            End Try
        End Sub

        Private Async Sub OnProcess(sender As Object, e As EventArgs)
            If cboProcess.SelectedIndex < 0 OrElse cboProcess.SelectedIndex >= _promptBodies.Count Then
                Return
            End If

            Dim prompt As String = _promptBodies(cboProcess.SelectedIndex)
            Dim payload As String = If(String.IsNullOrWhiteSpace(rtb.SelectedText), rtb.Text, rtb.SelectedText)
            Dim contextDocumentText As String = ""
            Dim contextDocumentPath As String = ""
            Dim contextDocumentIsDirectory As Boolean = False

            Dim askContext As Integer = ShowCustomYesNoBox(
                "Do you want to add a document or folder as additional context for processing the transcript?",
                "Yes",
                "No",
                Me.Text,
                extraButtonText:="Cancel",
                extraButtonAction:=Sub()
                                   End Sub,
                CloseAfterExtra:=True)

            If askContext = 0 Then
                Return
            End If

            If askContext = 1 Then
                DragDropFormLabel = "Context document or folder"
                DragDropFormFilter = "Supported|*.txt;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm;*.md;*.yaml;*.yml;*.vb;*.cs;*.js;*.ts;*.py;*.java;*.cpp;*.c;*.h;*.sql;*.rtf;*.doc;*.docx;*.xlsx;*.pptx;*.pdf;*.eml;*.msg|All|*.*"

                Try
                    Using f As New DragDropForm(DragDropMode.FileOrDirectory)
                        If f.ShowDialog() = DialogResult.OK Then
                            contextDocumentPath = f.SelectedFilePath
                            contextDocumentIsDirectory = f.IsDirectory
                        End If
                    End Using
                Finally
                    DragDropFormLabel = ""
                    DragDropFormFilter = ""
                End Try

                If Not String.IsNullOrWhiteSpace(contextDocumentPath) Then
                    If contextDocumentIsDirectory Then
                        Dim ctx As New FileLoadingContext()
                        Dim directoryResult As String = ""

                        Try
                            directoryResult = Await Globals.ThisAddIn.LoadDirectoryFilesAsync(
                                contextDocumentPath,
                                False,
                                ctx,
                                ensureProgressBar:=True)
                        Finally
                            ProgressBarModule.CancelOperation = True
                        End Try

                        If String.Equals(directoryResult, "ABORT", StringComparison.Ordinal) Then
                            Return
                        End If

                        contextDocumentText = directoryResult.Trim()

                        If String.IsNullOrWhiteSpace(contextDocumentText) Then
                            ShowTranscriptorMessageBox("The selected context folder could not be read or returned no usable text.")
                        End If
                    ElseIf File.Exists(contextDocumentPath) Then
                        Dim fileResult = Await Globals.ThisAddIn.GetFileContentEx(
                            contextDocumentPath,
                            Silent:=True,
                            DoOCR:=True,
                            AskUser:=True,
                            AskWorksheetSelection:=True)

                        contextDocumentText = If(fileResult.Content, "").Trim()

                        If String.IsNullOrWhiteSpace(contextDocumentText) Then
                            ShowTranscriptorMessageBox("The selected context document could not be read or returned no usable text.")
                        End If
                    Else
                        ShowTranscriptorMessageBox("The selected context path could not be found.")
                    End If
                End If
            End If

            Dim combinedPayload As String = payload.Trim()

            If Not String.IsNullOrWhiteSpace(contextDocumentText) Then
                Dim contextLabel As String = "Additional Context Document"
                Dim contextName As String = Path.GetFileName(contextDocumentPath)

                If contextDocumentIsDirectory Then
                    contextLabel = "Additional Context Folder"

                    Dim normalizedContextDirectoryPath As String =
                        contextDocumentPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)

                    contextName = Path.GetFileName(normalizedContextDirectoryPath)

                    If String.IsNullOrWhiteSpace(contextName) Then
                        contextName = normalizedContextDirectoryPath
                    End If
                ElseIf String.IsNullOrWhiteSpace(contextName) Then
                    contextName = contextDocumentPath
                End If

                combinedPayload &= vbCrLf & vbCrLf &
                    "=== " & contextLabel & ": " & contextName & " ===" & vbCrLf &
                    contextDocumentText
            End If

            Dim suffix As String = " (Current Date: " & DateTime.Now.ToString("dd MMM yyyy", CultureInfo.GetCultureInfo("en-US")) & ")"
            Dim result As String = Await LLM(prompt & suffix, combinedPayload, "", "", 0, False)

            Dim wordApp = Globals.ThisAddIn.Application
            If wordApp.Documents.Count > 0 Then
                Dim sel = wordApp.Selection
                sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                InsertTextWithMarkdown(sel, result, True)
            End If
        End Sub
        Private Async Sub OnClosing(sender As Object, e As FormClosingEventArgs)
            If _isStopping Then
                e.Cancel = True
                Return
            End If

            If _capturing Then
                e.Cancel = True
                _closeAfterStop = True
                Await StopCurrentSessionAsync()
            ElseIf _fileTranscribing Then
                e.Cancel = True
                _closeAfterStop = True
                CancelCurrentFileTranscription()
            End If
        End Sub

        Private Function GetLanguagePersistenceKey(d As EngineDescriptor) As String
            If d Is Nothing Then
                Return ""
            End If

            Return d.DisplayName
        End Function

        Private Function LoadLastLanguageMap() As Newtonsoft.Json.Linq.JObject
            Try
                Dim raw As String = My.Settings.LastLanguageByEngineJson
                If String.IsNullOrWhiteSpace(raw) Then
                    Return New Newtonsoft.Json.Linq.JObject()
                End If

                Return Newtonsoft.Json.Linq.JObject.Parse(raw)
            Catch
                Return New Newtonsoft.Json.Linq.JObject()
            End Try
        End Function

        Private Sub SaveLastLanguageMap(map As Newtonsoft.Json.Linq.JObject)
            Try
                My.Settings.LastLanguageByEngineJson = map.ToString(Newtonsoft.Json.Formatting.None)
            Catch
            End Try
        End Sub

        Private Function GetSavedLanguageForDescriptor(d As EngineDescriptor) As String
            Dim key As String = GetLanguagePersistenceKey(d)
            If String.IsNullOrWhiteSpace(key) Then
                Return ""
            End If

            Dim map = LoadLastLanguageMap()
            Dim token = map(key)
            If token Is Nothing Then
                Return ""
            End If

            Return token.ToString()
        End Function

        Private Sub SaveCurrentLanguageForCurrentEngine()
            Dim d As EngineDescriptor = CurrentDescriptor()
            If d Is Nothing Then
                Return
            End If

            If cboLang Is Nothing OrElse cboLang.SelectedItem Is Nothing Then
                Return
            End If

            Dim value As String = cboLang.SelectedItem.ToString()
            If String.IsNullOrWhiteSpace(value) Then
                Return
            End If

            Dim key As String = GetLanguagePersistenceKey(d)
            If String.IsNullOrWhiteSpace(key) Then
                Return
            End If

            Dim map = LoadLastLanguageMap()
            map(key) = value
            SaveLastLanguageMap(map)
        End Sub

        Private Function GetFriendlySttErrorText(engineDisplayName As String, rawMessage As String) As String
            Dim original As String = If(rawMessage, "").Trim()
            Dim m As String = original.ToLowerInvariant()

            If m.StartsWith("openai realtime error:") OrElse
       m.StartsWith("openai realtime ws send failed:") OrElse
       m.StartsWith("openai realtime ws read failed:") OrElse
       m.StartsWith("openai rest request failed:") Then
                Return engineDisplayName & ": " & original
            End If

            If m.Contains("permissiondenied") OrElse m.Contains("permission denied") Then
                Return engineDisplayName & ": Access was denied. Please check the configured account and permissions."
            End If

            If m.Contains("unavailable") OrElse
               m.Contains("winhttp") OrElse
               m.Contains("connection") OrElse
               m.Contains("timeout") OrElse
               m.Contains("interrupted") Then
                Return engineDisplayName & ": Connection to the speech service was interrupted. Please try again."
            End If

            If m.Contains("audio write failed") Then
                Return engineDisplayName & ": The live audio stream was interrupted. Please try again."
            End If

            If m.Contains("reader") OrElse m.Contains("writer") OrElse m.Contains("stream") Then
                Return engineDisplayName & ": The live transcription session ended unexpectedly."
            End If

            Return engineDisplayName & ": Transcription encountered an error."
        End Function

        Private Function GetFriendlyStatusText(engineDisplayName As String, rawMessage As String) As String
            Dim m As String = If(rawMessage, "").Trim()
            If String.IsNullOrWhiteSpace(m) Then
                Return ""
            End If

            Dim ml As String = m.ToLowerInvariant()

            If ml.Contains("preparing") OrElse
       ml.Contains("uploading") OrElse
       ml.Contains("streaming file") OrElse
       ml.Contains("transcribing file") OrElse
       ml.Contains("processing") OrElse
       ml.Contains("parsing") OrElse
       ml.Contains("finalizing") Then
                Return engineDisplayName & ": " & m
            End If

            If ml.Contains("cancel") Then
                Return engineDisplayName & ": Canceling…"
            End If

            If ml.Contains("stream configured") OrElse
       ml.Contains("streaming session opened") OrElse
       ml.Contains("session configured") Then
                If _fileTranscribing Then
                    Return GetFileTranscribingState()
                End If

                Return engineDisplayName & ": Listening…"
            End If

            If ml.Contains("starting") OrElse
       ml.Contains("opening") OrElse
       ml.Contains("connecting") Then
                Return engineDisplayName & ": Starting…"
            End If

            If ml.Contains("stopped") Then
                If _fileTranscribing Then
                    Return engineDisplayName & ": Finalizing file transcription…"
                End If

                Return engineDisplayName & ": Ready."
            End If

            Return engineDisplayName & ": " & m
        End Function

        Private Sub LogSttError(engineDisplayName As String, detail As String, Optional ex As Exception = Nothing)
            If Not INI_APIDebug Then
                Return
            End If

            Try
                Dim dir As String = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "RedInk")

                System.IO.Directory.CreateDirectory(dir)

                Dim logPath As String = System.IO.Path.Combine(dir, "RI_STT_Log.txt")
                Dim sb As New StringBuilder()

                sb.AppendLine("============================================================")
                sb.AppendLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                sb.AppendLine("Engine: " & engineDisplayName)
                sb.AppendLine("Detail: " & detail)

                If ex IsNot Nothing Then
                    sb.AppendLine("Exception: " & ex.ToString())
                End If

                System.IO.File.AppendAllText(logPath, sb.ToString() & Environment.NewLine)
            Catch
            End Try
        End Sub

        Private Shared Function NormalizeIniValue(value As String) As String
            Dim result As String = If(value, "").Trim()

            If result.Length >= 2 AndAlso result.StartsWith("""", StringComparison.Ordinal) AndAlso result.EndsWith("""", StringComparison.Ordinal) Then
                result = result.Substring(1, result.Length - 2).Trim()
            End If

            Return result
        End Function

        Private Shared Function ParseSttSettings(raw As String) As Dictionary(Of String, String)
            Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            For Each part As String In If(raw, "").Split(";"c)
                Dim trimmed As String = part.Trim()
                If trimmed.Length = 0 Then
                    Continue For
                End If

                Dim pos As Integer = trimmed.IndexOf("="c)
                If pos <= 0 Then
                    Continue For
                End If

                Dim key As String = trimmed.Substring(0, pos).Trim()
                Dim value As String = trimmed.Substring(pos + 1).Trim()

                If key.Length > 0 Then
                    result(key) = value
                End If
            Next

            Return result
        End Function

        Private Shared Function ResolveSttSetting(raw As String, modelOrTag As String, settingName As String, defaultValue As String) As String
            Dim settings As Dictionary(Of String, String) = ParseSttSettings(raw)
            Dim value As String = ""

            If Not String.IsNullOrWhiteSpace(modelOrTag) AndAlso
               settings.TryGetValue(modelOrTag & "." & settingName, value) Then
                Return value
            End If

            If settings.TryGetValue("default." & settingName, value) Then
                Return value
            End If

            If settings.TryGetValue(settingName, value) Then
                Return value
            End If

            Return defaultValue
        End Function

        Private Function ResolveGoogleProjectId() As String
            Return If(INI_STT_Google_ProjectID, "").Trim()
        End Function

        Private Function ResolveGoogleSttSetting(modelOrTag As String, settingName As String, defaultValue As String) As String
            Return ResolveSttSetting(INI_STT_Google, modelOrTag, settingName, defaultValue)
        End Function

        Private Function ResolveOpenAiSttSetting(modelOrTag As String, settingName As String, defaultValue As String) As String
            Return ResolveSttSetting(INI_STT_OpenAI, modelOrTag, settingName, defaultValue)
        End Function

        Private Function ResolveAzureSpeechKey() As String
            Return DecodeWrappedEncryptedValue(NormalizeIniValue(INI_STT_Azure_SpeechKey), "Azure Speech key")
        End Function

        Private Shared Function DecodeWrappedEncryptedValue(value As String, valueName As String) As String
            Dim normalized As String = NormalizeIniValue(value)

            If normalized.StartsWith("encrypted(", StringComparison.OrdinalIgnoreCase) AndAlso
               normalized.EndsWith(")", StringComparison.Ordinal) Then

                Dim innerValue As String = normalized.Substring(
                    "encrypted(".Length,
                    normalized.Length - "encrypted(".Length - 1).Trim()

                If String.IsNullOrWhiteSpace(innerValue) Then
                    Return ""
                End If

                Dim codeBasis As String = ResolveCodeBasis()
                If String.IsNullOrWhiteSpace(codeBasis) Then
                    Throw New InvalidOperationException("Missing CodeBasis for encrypted " & valueName & ".")
                End If

                Dim decoded As String = DecodeString(innerValue, codeBasis)
                If decoded.StartsWith("Error:", StringComparison.OrdinalIgnoreCase) Then
                    Throw New InvalidOperationException("Failed to decrypt " & valueName & ": " & decoded)
                End If

                Return NormalizeIniValue(decoded)
            End If

            Return normalized
        End Function

        Private Shared Function ResolveCodeBasis() As String
            Dim codeBasis As String = ""

            Try
                If ThisAddIn._context IsNot Nothing Then
                    codeBasis = If(ThisAddIn._context.Codebasis, "").Trim()
                End If
            Catch
            End Try

            If String.IsNullOrWhiteSpace(codeBasis) Then
                Try
                    If IsEmptyOrBlank(Int_CodeBasis) Then
                        codeBasis = GetFromRegistry(RegPath_Base, RegPath_CodeBasis, False)
                    Else
                        codeBasis = Int_CodeBasis
                    End If
                Catch
                End Try
            End If

            Try
                If ThisAddIn._context IsNot Nothing AndAlso
                   String.IsNullOrWhiteSpace(ThisAddIn._context.Codebasis) AndAlso
                   Not String.IsNullOrWhiteSpace(codeBasis) Then

                    ThisAddIn._context.Codebasis = codeBasis
                End If
            Catch
            End Try

            Return If(codeBasis, "").Trim()
        End Function

        Private Function ResolveAzureSttSetting(modelOrTag As String, settingName As String, defaultValue As String) As String
            Return NormalizeIniValue(ResolveSttSetting(INI_STT_Azure, modelOrTag, settingName, defaultValue))
        End Function

        Private Function ResolveAzureSpeechLocation(modelOrTag As String) As String
            Dim endpoint As String = ResolveAzureSttSetting(modelOrTag, "endpoint", "")
            If Not String.IsNullOrWhiteSpace(endpoint) Then
                Return endpoint
            End If

            Return ResolveAzureSttSetting(modelOrTag, "region", "")
        End Function

        Private Shared Function IsLiveOnlyEngine(kind As EngineKind) As Boolean
            Select Case kind
                Case EngineKind.OpenAiRealtime, EngineKind.AzureSpeechRealtime, EngineKind.TeamsAcsRealtime
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Function IsFileOnlyEngine(kind As EngineKind) As Boolean
            Select Case kind
                Case EngineKind.OpenAiRest, EngineKind.AzureSpeechFastRest
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Sub ShowTranscriptorMessageBox(
    bodyText As String,
    Optional header As String = Nothing,
    Optional autoCloseSeconds As Integer? = Nothing,
    Optional defaultText As String = " - execution continues meanwhile",
    Optional separateThread As Boolean = False,
    Optional extraButtonText As String = Nothing,
    Optional extraButtonAction As System.Action = Nothing,
    Optional closeAfterExtra As Boolean = False)

            If Me.IsDisposed Then
                Return
            End If

            If Me.InvokeRequired Then
                Try
                    Me.Invoke(
                        New Action(Of String, String, Integer?, String, Boolean, String, System.Action, Boolean)(
                            AddressOf ShowTranscriptorMessageBox),
                        bodyText,
                        header,
                        autoCloseSeconds,
                        defaultText,
                        separateThread,
                        extraButtonText,
                        extraButtonAction,
                        closeAfterExtra)
                Catch
                End Try
                Return
            End If

            Using SLib.PushDialogOwner(Me)
                ShowCustomMessageBox(
                    bodyText,
                    If(String.IsNullOrWhiteSpace(header), Me.Text, header),
                    autoCloseSeconds,
                    defaultText,
                    separateThread,
                    extraButtonText,
                    extraButtonAction,
                    closeAfterExtra)
            End Using

            Try
                Me.Activate()
                Me.BringToFront()
            Catch
            End Try
        End Sub


    End Class



End Class
