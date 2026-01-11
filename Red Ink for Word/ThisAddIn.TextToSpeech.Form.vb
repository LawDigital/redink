' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.TextToSpeech.Form.vb
' Purpose: Hosts the Text-to-Speech selection dialog that lets users pick a TTS
'          provider (Google/OpenAI), configure languages/voices, preview speech,
'          and choose the output mp3 destination used by Red Ink for Word.
'
' Architecture:
'  - UI Initialization: Builds the WinForms surface (engine selector, two voice
'    sets, sample text, output path controls) and persists user choices via My.Settings.
'  - Voice Loading: Populates combos with either cached Google Cloud voices
'    (fetched via REST + OAuth token) or locally defined OpenAI voices/descriptions.
'  - Playback: Generates preview audio through GenerateAndPlayAudio for the
'    currently highlighted language/voice combination.
'  - Output Selection: Normalizes output paths, honors “Temp only”, and exposes
'    SelectedVoices/SelectedOutputPath back to the caller.
'  - Dependencies: Relies on SharedLibrary.SharedMethods for TTS helpers,
'    OAuth token retrieval, and UI utilities; uses Newtonsoft.Json for parsing
'    Google voice responses.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Data
Imports System.Net.Http
Imports System.Windows
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SharedLibrary.SharedLibrary

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Provides a dialog for selecting Text-to-Speech provider, languages, voices, and output options.
    ''' Supports both Google Cloud TTS and OpenAI TTS engines with voice preview and persistence.
    ''' </summary>
    Public Class TTSSelectionForm
        Inherits Form

        ' -- Controls --
        Private lblIntro As Label

        ' engine selector combo:
        Private cmbEngine As Forms.ComboBox


        ' --- Set 1 Controls ---
        Private lblSet1 As Label
        Private cmbLanguage1 As Forms.ComboBox
        Private cmbVoice1A As Forms.ComboBox
        Private btnPlay1A As Forms.Button
        Private cmbVoice1B As Forms.ComboBox
        Private btnPlay1B As Forms.Button

        ' --- Set 2 Controls ---
        Private lblSet2 As Label
        Private cmbLanguage2 As Forms.ComboBox
        Private cmbVoice2A As Forms.ComboBox
        Private btnPlay2A As Forms.Button
        Private cmbVoice2B As Forms.ComboBox
        Private btnPlay2B As Forms.Button

        ' --- Sample text to play ---
        Private lblSampleText As Label
        Private txtSampleText As Forms.TextBox

        ' --- Bottom buttons ---
        Private btnOK As Forms.Button
        Private btnCancel As Forms.Button
        Private btnDesktop As Forms.Button

        ' --- For output path ---
        Private lblOutputPath As Label
        Private txtOutputPath As Forms.TextBox
        Private chkTemporary As Forms.CheckBox

        ' --- For storing voices from Google TTS ---
        ' This class helps us parse the JSON response from the voices API
        Private Class GoogleVoicesList
            <JsonProperty("voices")>
            Public Property Voices As List(Of GoogleVoice)
        End Class

        ''' <summary>
        ''' Represents a single voice returned by Google Cloud TTS with name, language codes, and gender attributes.
        ''' </summary>
        Private Class GoogleVoice
            <JsonProperty("name")>
            Public Property Name As String

            <JsonProperty("languageCodes")>
            Public Property LanguageCodes As List(Of String)

            <JsonProperty("ssmlGender")>
            Public Property SsmlGender As String
        End Class

        ' We can cache voices once retrieved for each language
        Private voiceCache As New Dictionary(Of String, List(Of GoogleVoice))()

        ' --- New parameters/fields for the amended form ---
        Private _twoVoicesRequired As Boolean
        Private _topLabelText As String
        Private _formTitle As String

        ' Radio buttons for voice selection.
        ' In one‐voice mode all four are in one group.
        ' In two‐voice mode we group each voice set separately (using Panels).
        Private rdoVoice1A As RadioButton, rdoVoice1B As RadioButton
        Private rdoVoice2A As RadioButton, rdoVoice2B As RadioButton
        'Private pnlVoiceSet1 As Panel, pnlVoiceSet2 As Panel

        ' --- Public properties to return results ---
        ' In one‑voice mode SelectedVoices will contain one item;
        ' in two‑voice mode it will contain two items.
        Public Property SelectedVoices As List(Of String) = New List(Of String)()
        Public Property SelectedOutputPath As String = ""
        Public Property SelectedLanguage As String = ""

        Public NoOutputFileRequired As Boolean = False

        ''' <summary>
        ''' Initializes the TTS selection form with specified display text, title, and voice mode.
        ''' </summary>
        ''' <param name="topLabelText">Introductory text displayed at the top of the form.</param>
        ''' <param name="formTitle">Window title for the dialog.</param>
        ''' <param name="twoVoicesRequired">True if caller needs two distinct voices; False for single-voice selection.</param>
        ''' <param name="NoOutputFile">When True, disables output path controls (for preview-only scenarios).</param>
        Public Sub New(topLabelText As String,
               formTitle As String,
               twoVoicesRequired As Boolean,
               Optional NoOutputFile As Boolean = False)

            NoOutputFileRequired = NoOutputFile

            'context As ISharedContext,
            'clientMail As String,
            'scopes As String,
            'apiKey As String,
            'oauth2Endpoint As String,
            'oauth2Expiry As Long,

            ' Assign external parameters
            '_context = context
            'INI_OAuth2ClientMail = clientMail
            'INI_OAuth2Scopes = scopes
            'INI_APIKey = apiKey
            'INI_OAuth2Endpoint = oauth2Endpoint
            'INI_OAuth2ATExpiry = oauth2Expiry

            ' Store our extra parameters
            _topLabelText = topLabelText
            _formTitle = formTitle
            _twoVoicesRequired = twoVoicesRequired

            ' --- FORM PROPERTIES ---
            Dim bmp As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
            Me.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())
            Me.Text = _formTitle
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.Font = New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            Me.AutoSize = False
            Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
            Me.MinimumSize = New System.Drawing.Size(810, 480)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
            Me.MaximizeBox = True

            Me.SuspendLayout()
            CreateControls()
            LayoutControls()
            Me.ResumeLayout()

            AddHandlers()

            Dim saved = My.Settings.TTSProvider
            If saved = "OpenAI" AndAlso TTS_openAIAvailable Then
                cmbEngine.SelectedItem = "OpenAI"
            ElseIf saved = "Google" AndAlso TTS_googleAvailable Then
                cmbEngine.SelectedItem = "Google"
            Else
                ' fall back to whichever is first in the list
                cmbEngine.SelectedIndex = 0
            End If

            PopulateLanguageComboBoxes()
            LoadSettingsAndVoices()

            txtSampleText.Text = If(
        String.IsNullOrEmpty(My.Settings.TTSSampleText),
        $"Hello, I am talking using {_formTitle}!",
        My.Settings.TTSSampleText
    )
        End Sub

        ''' <summary>
        ''' Instantiates all UI controls (labels, combos, buttons, checkboxes) and configures their initial properties.
        ''' </summary>
        Private Sub CreateControls()
            ' --- Intro ---
            lblIntro = New System.Windows.Forms.Label() With {
        .Font = Me.Font,
        .Text = _topLabelText,
        .AutoSize = True,
        .MaximumSize = New System.Drawing.Size(700, 0)
    }

            ' --- Engine selector ---
            cmbEngine = New System.Windows.Forms.ComboBox() With {
    .Font = Me.Font,
    .DropDownStyle = ComboBoxStyle.DropDownList,
    .Width = 150,
    .Margin = New System.Windows.Forms.Padding(0, -4, 0, 10)
}
            cmbEngine.Items.Clear()
            If TTS_googleAvailable Then cmbEngine.Items.Add("Google")
            If TTS_openAIAvailable Then cmbEngine.Items.Add("OpenAI")
            ' default to first available
            If cmbEngine.Items.Count > 0 Then cmbEngine.SelectedIndex = 0
            AddHandler cmbEngine.SelectedIndexChanged, AddressOf EngineChanged


            ' --- Voice Set 1 ---
            lblSet1 = New System.Windows.Forms.Label() With {.Font = Me.Font, .Text = "Your default voice set 1:", .AutoSize = True}
            cmbLanguage1 = New System.Windows.Forms.ComboBox() With {
        .Font = Me.Font,
        .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList,
        .Width = 300,
        .MaxDropDownItems = 10
    }
            cmbVoice1A = New System.Windows.Forms.ComboBox() With {
        .Font = Me.Font,
        .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList,
        .Width = 300,
        .MaxDropDownItems = 10
    }
            btnPlay1A = New System.Windows.Forms.Button() With {
        .Font = New System.Drawing.Font("Segoe UI Symbol", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point),
        .Text = "▶",
        .Size = New System.Drawing.Size(24, cmbVoice1A.PreferredHeight),
        .AutoSize = False
    }
            cmbVoice1B = New System.Windows.Forms.ComboBox() With {
        .Font = Me.Font,
        .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList,
        .Width = 300,
        .MaxDropDownItems = 10
    }
            btnPlay1B = New System.Windows.Forms.Button() With {
        .Font = New System.Drawing.Font("Segoe UI Symbol", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point),
        .Text = "▶",
        .Size = New System.Drawing.Size(24, cmbVoice1B.PreferredHeight),
        .AutoSize = False
    }

            ' --- Voice Set 2 ---
            lblSet2 = New System.Windows.Forms.Label() With {.Font = Me.Font, .Text = "Your default voice set 2:", .AutoSize = True}
            cmbLanguage2 = New System.Windows.Forms.ComboBox() With {
        .Font = Me.Font,
        .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList,
        .Width = 300,
        .MaxDropDownItems = 10
    }
            cmbVoice2A = New System.Windows.Forms.ComboBox() With {
        .Font = Me.Font,
        .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList,
        .Width = 300,
        .MaxDropDownItems = 10
    }
            btnPlay2A = New System.Windows.Forms.Button() With {
        .Font = New System.Drawing.Font("Segoe UI Symbol", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point),
        .Text = "▶",
        .Size = New System.Drawing.Size(24, cmbVoice2A.PreferredHeight),
        .AutoSize = False
    }
            cmbVoice2B = New System.Windows.Forms.ComboBox() With {
        .Font = Me.Font,
        .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList,
        .Width = 300,
        .MaxDropDownItems = 10
    }
            btnPlay2B = New System.Windows.Forms.Button() With {
        .Font = New System.Drawing.Font("Segoe UI Symbol", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point),
        .Text = "▶",
        .Size = New System.Drawing.Size(24, cmbVoice2B.PreferredHeight),
        .AutoSize = False
    }

            ' --- Radio Buttons ---
            If Not _twoVoicesRequired Then
                rdoVoice1A = New System.Windows.Forms.RadioButton() With {.Font = Me.Font, .AutoSize = True}
                rdoVoice1B = New System.Windows.Forms.RadioButton() With {.Font = Me.Font, .AutoSize = True}
                rdoVoice2A = New System.Windows.Forms.RadioButton() With {.Font = Me.Font, .AutoSize = True}
                rdoVoice2B = New System.Windows.Forms.RadioButton() With {.Font = Me.Font, .AutoSize = True}
                Select Case My.Settings.TTSLastRdoOneVoice
                    Case "Voice1A"
                        rdoVoice1A.Checked = True
                    Case "Voice1B"
                        rdoVoice1B.Checked = True
                    Case "Voice2A"
                        rdoVoice2A.Checked = True
                    Case "Voice2B"
                        rdoVoice2B.Checked = True
                    Case Else
                        rdoVoice1A.Checked = True ' Default if no previous selection
                End Select
            Else
                rdoVoice1A = New System.Windows.Forms.RadioButton() With {.Font = Me.Font, .AutoSize = True}
                rdoVoice2A = New System.Windows.Forms.RadioButton() With {.Font = Me.Font, .AutoSize = True}
                Select Case My.Settings.TTSLastRdoTwoVoices
                    Case "Voice1"
                        rdoVoice1A.Checked = True
                    Case "Voice2"
                        rdoVoice2A.Checked = True
                    Case Else
                        rdoVoice1A.Checked = True ' Default if no previous selection
                End Select
            End If

            ' --- Sample & Output rows ---
            lblSampleText = New System.Windows.Forms.Label() With {.Font = Me.Font, .Text = "Sample text:", .AutoSize = True}
            txtSampleText = New System.Windows.Forms.TextBox() With {.Font = Me.Font, .Width = 467}
            lblOutputPath = New System.Windows.Forms.Label() With {.Font = Me.Font, .Text = "Output (.mp3):", .AutoSize = True, .Enabled = Not NoOutputFileRequired}
            txtOutputPath = New System.Windows.Forms.TextBox() With {.Font = Me.Font, .Width = 330, .Enabled = Not NoOutputFileRequired}
            chkTemporary = New System.Windows.Forms.CheckBox() With {.Font = Me.Font, .Text = "Temp only", .AutoSize = True, .Enabled = Not NoOutputFileRequired}

            ' --- Bottom Buttons ---
            btnOK = New System.Windows.Forms.Button() With {.Font = Me.Font, .Text = "OK", .AutoSize = True}
            btnCancel = New System.Windows.Forms.Button() With {.Font = Me.Font, .Text = "Cancel", .AutoSize = True}
            btnDesktop = New System.Windows.Forms.Button() With {.Font = Me.Font, .Text = "Save on Desktop", .AutoSize = True, .Enabled = Not NoOutputFileRequired}

            ' --- Wire up mutual‑exclusion for all radios ---
            Dim radios As New List(Of System.Windows.Forms.RadioButton)
            For Each rb In New RadioButton() {rdoVoice1A, rdoVoice1B, rdoVoice2A, rdoVoice2B}
                If rb IsNot Nothing Then radios.Add(rb)
            Next
            For Each rb In radios
                AddHandler rb.CheckedChanged, Sub(s, e)
                                                  Dim meRb = DirectCast(s, RadioButton)
                                                  If meRb.Checked Then
                                                      For Each other In radios
                                                          If other IsNot meRb Then other.Checked = False
                                                      Next
                                                  End If
                                              End Sub
            Next
        End Sub

        ''' <summary>
        ''' Arranges controls into a two-column TableLayoutPanel layout with proper spacing and alignment.
        ''' </summary>
        Private Sub LayoutControls()

            Me.Controls.Clear()

            ' Root: 2 cols, 9 rows, bottom padding = 20px
            Dim root As New System.Windows.Forms.TableLayoutPanel() With {
                      .Dock = System.Windows.Forms.DockStyle.Fill,
                      .AutoSize = True,
                      .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                      .ColumnCount = 2,
                      .RowCount = 9,
                      .Padding = New System.Windows.Forms.Padding(10, 10, 10, 20)
                    }
            root.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0F))
            root.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0F))
            For i = 0 To 8
                root.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            Next

            ' Row0: Intro
            root.Controls.Add(lblIntro, 0, 0)
            root.SetColumnSpan(lblIntro, 2)

            ' Row1: Provider
            root.RowStyles.Insert(1, New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            root.Controls.Add(New Label() With {
                    .Font = Me.Font,
                    .Text = "Text-to-Speech Provider:",
                    .AutoSize = True
                }, 0, 1)
            root.Controls.Add(cmbEngine, 1, 1)


            ' Row1: Headings
            root.Controls.Add(lblSet1, 0, 2)
            root.Controls.Add(lblSet2, 1, 2)

            ' Row2: Language (+ two‑voice radio)
            Dim fl2a As New System.Windows.Forms.FlowLayoutPanel() With {
      .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
      .AutoSize = True
    }
            If _twoVoicesRequired Then fl2a.Controls.Add(rdoVoice1A)
            fl2a.Controls.Add(cmbLanguage1)
            root.Controls.Add(fl2a, 0, 3)

            Dim fl2b As New System.Windows.Forms.FlowLayoutPanel() With {
      .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
      .AutoSize = True
    }
            If _twoVoicesRequired Then fl2b.Controls.Add(rdoVoice2A)
            fl2b.Controls.Add(cmbLanguage2)
            root.Controls.Add(fl2b, 1, 3)

            ' Row3: Voice1A + play + single‑voice radio or indent if two‑voice mode
            Dim fl3a As New System.Windows.Forms.FlowLayoutPanel() With {
      .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
      .AutoSize = True
    }
            If _twoVoicesRequired Then
                ' indent by the radio’s width so it lines up with language
                fl3a.Padding = New System.Windows.Forms.Padding(rdoVoice1A.PreferredSize.Width, 0, 0, 0)
            Else
                rdoVoice1A.Margin = New System.Windows.Forms.Padding(0,
        (cmbVoice1A.PreferredHeight - rdoVoice1A.PreferredSize.Height) \ 2, 0, 0)
                fl3a.Controls.Add(rdoVoice1A)
            End If
            fl3a.Controls.Add(cmbVoice1A)
            fl3a.Controls.Add(btnPlay1A)
            root.Controls.Add(fl3a, 0, 4)

            Dim fl3b As New System.Windows.Forms.FlowLayoutPanel() With {
      .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
      .AutoSize = True
    }
            If _twoVoicesRequired Then
                fl3b.Padding = New System.Windows.Forms.Padding(rdoVoice2A.PreferredSize.Width, 0, 0, 0)
            Else
                rdoVoice2A.Margin = New System.Windows.Forms.Padding(0,
        (cmbVoice2A.PreferredHeight - rdoVoice2A.PreferredSize.Height) \ 2, 0, 0)
                fl3b.Controls.Add(rdoVoice2A)
            End If
            fl3b.Controls.Add(cmbVoice2A)
            fl3b.Controls.Add(btnPlay2A)
            root.Controls.Add(fl3b, 1, 4)

            ' Row4: Voice1B + play (same indent logic)
            Dim fl4a As New System.Windows.Forms.FlowLayoutPanel() With {
      .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
      .AutoSize = True
    }
            If _twoVoicesRequired Then
                fl4a.Padding = New System.Windows.Forms.Padding(rdoVoice1A.PreferredSize.Width, 0, 0, 0)
            Else
                rdoVoice1B.Margin = New System.Windows.Forms.Padding(0,
        (cmbVoice1B.PreferredHeight - rdoVoice1B.PreferredSize.Height) \ 2, 0, 0)
                fl4a.Controls.Add(rdoVoice1B)
            End If
            fl4a.Controls.Add(cmbVoice1B)
            fl4a.Controls.Add(btnPlay1B)
            root.Controls.Add(fl4a, 0, 5)

            Dim fl4b As New System.Windows.Forms.FlowLayoutPanel() With {
      .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
      .AutoSize = True
    }
            If _twoVoicesRequired Then
                fl4b.Padding = New System.Windows.Forms.Padding(rdoVoice2A.PreferredSize.Width, 0, 0, 0)
            Else
                rdoVoice2B.Margin = New System.Windows.Forms.Padding(0,
        (cmbVoice2B.PreferredHeight - rdoVoice2B.PreferredSize.Height) \ 2, 0, 0)
                fl4b.Controls.Add(rdoVoice2B)
            End If
            fl4b.Controls.Add(cmbVoice2B)
            fl4b.Controls.Add(btnPlay2B)
            root.Controls.Add(fl4b, 1, 5)

            ' Row5: Sample text (2‑col table for vertical centering)
            Dim tbl5 As New System.Windows.Forms.TableLayoutPanel() With {
      .ColumnCount = 2,
      .RowCount = 1,
      .AutoSize = True,
      .Dock = System.Windows.Forms.DockStyle.Top
    }
            tbl5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize))
            tbl5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
            lblSampleText.Dock = System.Windows.Forms.DockStyle.Fill
            lblSampleText.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            txtSampleText.Dock = System.Windows.Forms.DockStyle.Fill
            tbl5.Controls.Add(lblSampleText, 0, 0)
            tbl5.Controls.Add(txtSampleText, 1, 0)
            root.Controls.Add(tbl5, 0, 6)
            root.SetColumnSpan(tbl5, 2)

            ' Row6: Output path (3‑col table)
            Dim tbl6 As New System.Windows.Forms.TableLayoutPanel() With {
      .ColumnCount = 3,
      .RowCount = 1,
      .AutoSize = True,
      .Dock = System.Windows.Forms.DockStyle.Top
    }
            tbl6.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize))
            tbl6.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
            tbl6.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize))
            lblOutputPath.Dock = System.Windows.Forms.DockStyle.Fill
            lblOutputPath.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            txtOutputPath.Dock = System.Windows.Forms.DockStyle.Fill
            chkTemporary.Anchor = System.Windows.Forms.AnchorStyles.Left
            tbl6.Controls.Add(lblOutputPath, 0, 0)
            tbl6.Controls.Add(txtOutputPath, 1, 0)
            tbl6.Controls.Add(chkTemporary, 2, 0)
            root.Controls.Add(tbl6, 0, 7)
            root.SetColumnSpan(tbl6, 2)

            Dim pnlButtons As New System.Windows.Forms.FlowLayoutPanel() With {
                  .Dock = System.Windows.Forms.DockStyle.Bottom,
                  .AutoSize = True,
                  .Padding = New Padding(10),
                  .FlowDirection = FlowDirection.LeftToRight
                }
            pnlButtons.Controls.Add(btnOK)
            pnlButtons.Controls.Add(btnCancel)
            pnlButtons.Controls.Add(btnDesktop)

            Me.Controls.Clear()
            Me.Controls.Add(pnlButtons)
            Me.Controls.Add(root)

        End Sub

        ''' <summary>
        ''' Handles engine selection changes by updating the global TTS_SelectedEngine, persisting the choice,
        ''' and reloading language/voice lists for the newly selected provider.
        ''' </summary>
        Private Sub EngineChanged(sender As Object, e As EventArgs)
            ' set our global
            TTS_SelectedEngine = If(cmbEngine.SelectedItem.ToString() = "OpenAI",
                             TTSEngine.OpenAI,
                             TTSEngine.Google)

            My.Settings.TTSProvider = cmbEngine.SelectedItem.ToString()
            My.Settings.Save()

            ' rebuild the combos
            PopulateLanguageComboBoxes()
            LoadSettingsAndVoices()
        End Sub

        ''' <summary>
        ''' Wires up all event handlers for language changes, play buttons, and action buttons.
        ''' </summary>
        Private Sub AddHandlers()
            AddHandler cmbLanguage1.SelectedIndexChanged, AddressOf cmbLanguage1_SelectedIndexChanged
            AddHandler cmbLanguage2.SelectedIndexChanged, AddressOf cmbLanguage2_SelectedIndexChanged

            AddHandler btnPlay1A.Click, AddressOf btnPlay1A_Click
            AddHandler btnPlay1B.Click, AddressOf btnPlay1B_Click
            AddHandler btnPlay2A.Click, AddressOf btnPlay2A_Click
            AddHandler btnPlay2B.Click, AddressOf btnPlay2B_Click

            AddHandler btnOK.Click, AddressOf btnOK_Click
            AddHandler btnCancel.Click, AddressOf btnCancel_Click
            AddHandler btnDesktop.Click, AddressOf btnDesktop_Click
            AddHandler chkTemporary.CheckedChanged, AddressOf chkTemporary_CheckedChanged
        End Sub

        ''' <summary>
        ''' Populates both language combo boxes with either OpenAI or Google language codes based on the selected engine.
        ''' </summary>
        Private Sub PopulateLanguageComboBoxes()
            cmbLanguage1.Items.Clear()
            cmbLanguage2.Items.Clear()

            If TTS_SelectedEngine = TTSEngine.OpenAI Then
                cmbLanguage1.Items.AddRange(OpenAILanguages)
                cmbLanguage2.Items.AddRange(OpenAILanguages)
            Else
                For Each lang In GoogleTTSsupportedLanguages
                    cmbLanguage1.Items.Add(lang)
                    cmbLanguage2.Items.Add(lang)
                Next
            End If

            If cmbLanguage1.Items.Count > 0 Then cmbLanguage1.SelectedIndex = 0
            If cmbLanguage2.Items.Count > 0 Then cmbLanguage2.SelectedIndex = 0
        End Sub

        ''' <summary>
        ''' Restores previously selected languages and voices from My.Settings, then asynchronously loads voice lists
        ''' for Google or synchronously for OpenAI.
        ''' </summary>
        Private Async Sub LoadSettingsAndVoices()
            RemoveHandler cmbLanguage1.SelectedIndexChanged, AddressOf cmbLanguage1_SelectedIndexChanged
            RemoveHandler cmbLanguage2.SelectedIndexChanged, AddressOf cmbLanguage2_SelectedIndexChanged

            ' restore last‐used languages
            cmbLanguage1.SelectedItem = My.Settings.TTS1languagecode
            cmbLanguage2.SelectedItem = My.Settings.TTS2languagecode

            Dim tasks As New List(Of System.Threading.Tasks.Task)

            If TTS_SelectedEngine = TTSEngine.OpenAI Then
                ' immediate, sync fill of voices
                PopulateOpenAIVoices(cmbLanguage1.Text, cmbVoice1A, cmbVoice1B)
                PopulateOpenAIVoices(cmbLanguage2.Text, cmbVoice2A, cmbVoice2B)
            Else
                ' Google: async fetch
                If Not String.IsNullOrEmpty(cmbLanguage1.Text) Then
                    tasks.Add(LoadVoicesIntoComboBoxesAsync(cmbLanguage1.Text, cmbVoice1A, cmbVoice1B))
                End If
                If Not String.IsNullOrEmpty(cmbLanguage2.Text) Then
                    tasks.Add(LoadVoicesIntoComboBoxesAsync(cmbLanguage2.Text, cmbVoice2A, cmbVoice2B))
                End If
            End If

            If tasks.Count > 0 Then Await System.Threading.Tasks.Task.WhenAll(tasks)

            ' restore last‐used voice selections
            cmbVoice1A.SelectedItem = My.Settings.TTS1voiceA
            cmbVoice1B.SelectedItem = My.Settings.TTS1voiceB
            cmbVoice2A.SelectedItem = My.Settings.TTS2voiceA
            cmbVoice2B.SelectedItem = My.Settings.TTS2voiceB

            AddHandler cmbLanguage1.SelectedIndexChanged, AddressOf cmbLanguage1_SelectedIndexChanged
            AddHandler cmbLanguage2.SelectedIndexChanged, AddressOf cmbLanguage2_SelectedIndexChanged
        End Sub


        ''' <summary>
        ''' Fills the specified voice combo boxes with OpenAI voice names and descriptions.
        ''' </summary>
        ''' <param name="lang">Language code (not used by OpenAI but kept for consistency).</param>
        ''' <param name="comboA">First voice combo box to populate.</param>
        ''' <param name="comboB">Second voice combo box to populate.</param>
        Private Sub PopulateOpenAIVoices(lang As String,
                                 comboA As Forms.ComboBox,
                                 comboB As Forms.ComboBox)
            comboA.Items.Clear() : comboB.Items.Clear()
            For Each v In OpenAIVoices
                Dim disp = $"{v} — {OpenAIDescriptions(v)}"
                comboA.Items.Add(disp)
                comboB.Items.Add(disp)
            Next
            If comboA.Items.Count > 0 Then comboA.SelectedIndex = 0
            If comboB.Items.Count > 0 Then comboB.SelectedIndex = 0
        End Sub

        ''' <summary>
        ''' Handles language selection changes for voice set 1 by reloading the corresponding voice combo boxes.
        ''' </summary>
        Private Async Sub cmbLanguage1_SelectedIndexChanged(sender As Object, e As EventArgs)
            Dim lang = TryCast(cmbLanguage1.SelectedItem, String)
            If String.IsNullOrEmpty(lang) Then Return

            If TTS_SelectedEngine = TTSEngine.OpenAI Then
                PopulateOpenAIVoices(lang, cmbVoice1A, cmbVoice1B)
            Else
                Await LoadVoicesIntoComboBoxesAsync(lang, cmbVoice1A, cmbVoice1B)
            End If
        End Sub

        ''' <summary>
        ''' Handles language selection changes for voice set 2 by reloading the corresponding voice combo boxes.
        ''' </summary>
        Private Async Sub cmbLanguage2_SelectedIndexChanged(sender As Object, e As EventArgs)
            Dim lang = TryCast(cmbLanguage2.SelectedItem, String)
            If String.IsNullOrEmpty(lang) Then Return

            If TTS_SelectedEngine = TTSEngine.OpenAI Then
                PopulateOpenAIVoices(lang, cmbVoice2A, cmbVoice2B)
            Else
                Await LoadVoicesIntoComboBoxesAsync(lang, cmbVoice2A, cmbVoice2B)
            End If
        End Sub

        ''' <summary>
        ''' Fetches Google Cloud TTS voices for the specified language code, filters by prefix match,
        ''' and populates the given combo boxes with display names including gender.
        ''' </summary>
        ''' <param name="languageCode">BCP-47 language code (e.g., "de-DE").</param>
        ''' <param name="comboA">First voice combo box to populate.</param>
        ''' <param name="comboB">Second voice combo box to populate.</param>

        Private Async Function LoadVoicesIntoComboBoxesAsync(languageCode As String,
                                                           comboA As Forms.ComboBox,
                                                           comboB As Forms.ComboBox) As System.Threading.Tasks.Task
            Try
                Dim allVoices As List(Of GoogleVoice) = Await GetVoicesByLanguageAsync(languageCode)
                comboA.Items.Clear()
                comboB.Items.Clear()
                If allVoices Is Nothing Then Exit Function

                ' STRICT FILTER:
                ' Only keep voices whose Name starts with "<langCode>-"
                ' (e.g. "de-DE-Wavenet-A"). This removes generic / multi‑locale
                ' voices like "Schedar" etc.
                Dim filtered = allVoices.Where(
                    Function(v) Not String.IsNullOrEmpty(v.Name) AndAlso
                                v.Name.StartsWith(languageCode & "-", StringComparison.OrdinalIgnoreCase)
                ).ToList()

                For Each v In filtered
                    Dim displayName As String = $"{v.Name} ({v.SsmlGender.ToLower()})"
                    comboA.Items.Add(displayName)
                    comboB.Items.Add(displayName)
                Next

                If comboA.Items.Count > 0 Then comboA.SelectedIndex = 0
                If comboB.Items.Count > 0 Then comboB.SelectedIndex = 0

                ' If previously saved selections no longer exist, they will be ignored automatically.
            Catch ex As System.Exception
                ShowCustomMessageBox("When trying to load the voices from the Google server, an error occurred: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Retrieves the list of Google Cloud TTS voices for the specified language, using cached data when available.
        ''' Authenticates via OAuth token and parses the JSON response.
        ''' </summary>
        ''' <param name="languageCode">BCP-47 language code to query.</param>
        ''' <returns>List of GoogleVoice objects, or Nothing on error.</returns>

        Private Async Function GetVoicesByLanguageAsync(languageCode As String) As System.Threading.Tasks.Task(Of List(Of GoogleVoice))
            If voiceCache.ContainsKey(languageCode) Then
                Return voiceCache(languageCode)
            End If

            Dim AccessToken As String = Await GetFreshTTSToken(UseSecondaryFor(TTSEngine.Google))
            If String.IsNullOrEmpty(AccessToken) Then
                ShowCustomMessageBox("Error accessing Google API - authentication failed (no token).")
                Return Nothing
            End If

            ' Build request
            Dim url As String = TTS_GoogleEndpoint & "voices?languageCode=" & languageCode
            Using httpClient As New HttpClient()
                httpClient.DefaultRequestHeaders.Authorization = New Net.Http.Headers.AuthenticationHeaderValue("Bearer", AccessToken)

                Dim response As HttpResponseMessage = Await httpClient.GetAsync(url)
                If response.IsSuccessStatusCode Then
                    Dim responseContent As String = Await response.Content.ReadAsStringAsync()
                    Dim voicesList As GoogleVoicesList = JsonConvert.DeserializeObject(Of GoogleVoicesList)(responseContent)

                    If voicesList IsNot Nothing AndAlso voicesList.Voices IsNot Nothing Then
                        voiceCache(languageCode) = voicesList.Voices
                        Return voicesList.Voices
                    Else
                        Return New List(Of GoogleVoice)()
                    End If
                Else
                    ShowCustomMessageBox("Failed to retrieve voices: " & response.StatusCode.ToString())
                    Return Nothing
                End If
            End Using
        End Function

        ' --- Play button event handlers ---
        Private Async Sub btnPlay1A_Click(sender As Object, e As EventArgs)
            Await PlaySelectedVoiceAsync(cmbLanguage1, cmbVoice1A)
        End Sub

        Private Async Sub btnPlay1B_Click(sender As Object, e As EventArgs)
            Await PlaySelectedVoiceAsync(cmbLanguage1, cmbVoice1B)
        End Sub

        Private Async Sub btnPlay2A_Click(sender As Object, e As EventArgs)
            Await PlaySelectedVoiceAsync(cmbLanguage2, cmbVoice2A)
        End Sub

        Private Async Sub btnPlay2B_Click(sender As Object, e As EventArgs)
            Await PlaySelectedVoiceAsync(cmbLanguage2, cmbVoice2B)
        End Sub

        ''' <summary>
        ''' Generates and plays audio for the selected language and voice using the sample text.
        ''' Strips gender/description suffixes before invoking GenerateAndPlayAudio.
        ''' </summary>
        ''' <param name="cmbLang">Language combo box to read from.</param>
        ''' <param name="cmbVoice">Voice combo box to read from.</param>
        Private Async Function PlaySelectedVoiceAsync(cmbLang As Forms.ComboBox, cmbVoice As Forms.ComboBox) As System.Threading.Tasks.Task
            Try
                Dim lang As String = TryCast(cmbLang.SelectedItem, String)
                Dim voiceName As String = TryCast(cmbVoice.SelectedItem, String)
                Dim sampleText As String = txtSampleText.Text

                If String.IsNullOrEmpty(lang) OrElse String.IsNullOrEmpty(voiceName) Then
                    ShowCustomMessageBox("Please select both language and voice before playing.")
                    Return
                End If
                voiceName = voiceName.Replace(" (male)", "").Replace(" (female)", "")
                If TTS_SelectedEngine = TTSEngine.OpenAI Then
                    ' remove OpenAI voice description suffix
                    voiceName = voiceName.Split(" "c)(0)
                End If
                Await System.Threading.Tasks.Task.Run(Sub()
                                                          GenerateAndPlayAudio(sampleText, "", lang, voiceName)
                                                      End Sub)
            Catch ex As System.Exception
                ShowCustomMessageBox("When trying to play the voice, an error occurred: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Validates selections, normalizes output path, saves settings, and closes the dialog with DialogResult.OK.
        ''' Strips OpenAI description suffixes and handles both one-voice and two-voice modes.
        ''' </summary>
        Private Sub btnOK_Click(sender As Object, e As EventArgs)

            TTS_SelectedEngine = If(cmbEngine.SelectedItem.ToString() = "OpenAI",
                         TTSEngine.OpenAI,
                         TTSEngine.Google)

            My.Settings.TTSProvider = cmbEngine.SelectedItem.ToString()
            My.Settings.Save()

            Dim NotAllSelected As Boolean = False

            ' Determine which voice(s) were selected based on radio buttons
            SelectedVoices.Clear()
            If Not _twoVoicesRequired Then
                ' ONE VOICE mode: the four radio buttons are one group.

                If rdoVoice1A.Checked Then
                    If cmbVoice1A.SelectedItem IsNot Nothing AndAlso cmbVoice1A.SelectedItem.ToString() <> "" Then
                        Dim sel As String = cmbVoice1A.SelectedItem.ToString()
                        If TTS_SelectedEngine = TTSEngine.OpenAI Then
                            ' drop the “ — Beschreibung” part
                            sel = sel.Split(" "c)(0)
                        End If
                        SelectedVoices.Add(sel)
                        SelectedLanguage = cmbLanguage1.SelectedItem.ToString()
                    Else
                        NotAllSelected = True
                    End If
                ElseIf rdoVoice1B.Checked Then
                    If cmbVoice1B.SelectedItem IsNot Nothing AndAlso cmbVoice1B.SelectedItem.ToString() <> "" Then

                        Dim sel As String = cmbVoice1B.SelectedItem.ToString()
                        If TTS_SelectedEngine = TTSEngine.OpenAI Then
                            ' drop the “ — Beschreibung” part
                            sel = sel.Split(" "c)(0)
                        End If
                        SelectedVoices.Add(sel)

                        SelectedLanguage = cmbLanguage1.SelectedItem.ToString()
                    Else
                        NotAllSelected = True
                    End If
                ElseIf rdoVoice2A.Checked Then
                    If cmbVoice2A.SelectedItem IsNot Nothing AndAlso cmbVoice2A.SelectedItem.ToString() <> "" Then

                        Dim sel As String = cmbVoice2A.SelectedItem.ToString()
                        If TTS_SelectedEngine = TTSEngine.OpenAI Then
                            ' drop the “ — Beschreibung” part
                            sel = sel.Split(" "c)(0)
                        End If
                        SelectedVoices.Add(sel)

                        SelectedLanguage = cmbLanguage2.SelectedItem.ToString()
                    Else
                        NotAllSelected = True
                    End If
                ElseIf rdoVoice2B.Checked Then
                    If cmbVoice2B.SelectedItem IsNot Nothing AndAlso cmbVoice2B.SelectedItem.ToString() <> "" Then
                        Dim sel As String = cmbVoice2B.SelectedItem.ToString()
                        If TTS_SelectedEngine = TTSEngine.OpenAI Then
                            ' drop the “ — Beschreibung” part
                            sel = sel.Split(" "c)(0)
                        End If
                        SelectedVoices.Add(sel)

                        SelectedLanguage = cmbLanguage2.SelectedItem.ToString()
                    Else
                        NotAllSelected = True
                    End If
                End If
            Else
                ' TWO VOICES mode: one voice from each set.
                If rdoVoice1A.Checked Then
                    If cmbVoice1A.SelectedItem IsNot Nothing AndAlso cmbVoice1A.SelectedItem.ToString() <> "" Then
                        SelectedVoices.Add(cmbVoice1A.SelectedItem.ToString())
                        SelectedLanguage = cmbLanguage1.SelectedItem.ToString()
                    Else
                        NotAllSelected = True
                    End If
                    If cmbVoice1B.SelectedItem IsNot Nothing AndAlso cmbVoice1B.SelectedItem.ToString() <> "" Then
                        SelectedVoices.Add(cmbVoice1B.SelectedItem.ToString())
                        SelectedLanguage = cmbLanguage1.SelectedItem.ToString()
                    Else
                        NotAllSelected = True
                    End If
                ElseIf rdoVoice2A.Checked Then
                    If cmbVoice2A.SelectedItem IsNot Nothing AndAlso cmbVoice2A.SelectedItem.ToString() <> "" Then
                        SelectedVoices.Add(cmbVoice2A.SelectedItem.ToString())
                        SelectedLanguage = cmbLanguage2.SelectedItem.ToString()
                    Else
                        NotAllSelected = True
                    End If
                    If cmbVoice2B.SelectedItem IsNot Nothing AndAlso cmbVoice2B.SelectedItem.ToString() <> "" Then
                        SelectedVoices.Add(cmbVoice2B.SelectedItem.ToString())
                        SelectedLanguage = cmbLanguage2.SelectedItem.ToString()
                    Else
                        NotAllSelected = True
                    End If
                End If
            End If

            If NotAllSelected Then
                ShowCustomMessageBox("Please complete your voice selection (Or 'Cancel').")
                Return
            End If

            ' Save selected radio button (for one-voice mode)
            If Not _twoVoicesRequired Then
                If rdoVoice1A.Checked Then
                    My.Settings.TTSLastRdoOneVoice = "Voice1A"
                ElseIf rdoVoice1B.Checked Then
                    My.Settings.TTSLastRdoOneVoice = "Voice1B"
                ElseIf rdoVoice2A.Checked Then
                    My.Settings.TTSLastRdoOneVoice = "Voice2A"
                ElseIf rdoVoice2B.Checked Then
                    My.Settings.TTSLastRdoOneVoice = "Voice2B"
                End If
            Else
                ' Save selected radio button (for two-voices mode)
                If rdoVoice1A.Checked Then
                    My.Settings.TTSLastRdoTwoVoices = "Voice1"
                ElseIf rdoVoice2A.Checked Then
                    My.Settings.TTSLastRdoTwoVoices = "Voice2"
                End If
            End If
            ' Save settings as before
            My.Settings.Save()

            ' Determine output path: if Temporary is checked, return blank.

            SelectedOutputPath = txtOutputPath.Text

            If String.IsNullOrWhiteSpace(SelectedOutputPath) Then
                ' Use default path (Desktop) with default filename
                SelectedOutputPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), TTSDefaultFile)
            ElseIf SelectedOutputPath.EndsWith("\") OrElse SelectedOutputPath.EndsWith("/") Then
                ' If only a folder is given, append default filename
                SelectedOutputPath = System.IO.Path.Combine(SelectedOutputPath, TTSDefaultFile)
            Else
                Dim dir As String = System.IO.Path.GetDirectoryName(SelectedOutputPath)
                Dim fileName As String = System.IO.Path.GetFileName(SelectedOutputPath)

                ' If no directory is found, assume Desktop as the base
                If String.IsNullOrWhiteSpace(dir) Then
                    SelectedOutputPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), fileName)
                    dir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                End If

                ' If no filename is given, use the default filename
                If String.IsNullOrWhiteSpace(fileName) Then
                    SelectedOutputPath = System.IO.Path.Combine(dir, TTSDefaultFile)
                End If

                ' Ensure the filename has ".mp3" extension
                If Not fileName.EndsWith(".mp3", StringComparison.OrdinalIgnoreCase) Then
                    SelectedOutputPath = System.IO.Path.Combine(dir, fileName & ".mp3")
                End If
            End If

            ' Update the TextBox with the corrected path
            txtOutputPath.Text = SelectedOutputPath

            SaveSettings()

            If chkTemporary.Checked Then
                SelectedOutputPath = ""
            End If

            Me.DialogResult = DialogResult.OK
            Me.Close()

        End Sub

        ''' <summary>
        ''' Clears selected voices and closes the dialog with DialogResult.Cancel.
        ''' </summary>
        Private Sub btnCancel_Click(sender As Object, e As EventArgs)
            ' If cancelled, clear any voice selection.
            SelectedVoices.Clear()
            Me.DialogResult = DialogResult.Cancel
            Me.Close()
        End Sub

        ''' <summary>
        ''' Persists the current language, voice, sample text, and output path selections to My.Settings.
        ''' </summary>
        Private Sub SaveSettings()
            My.Settings.TTS1languagecode = If(cmbLanguage1.SelectedItem?.ToString(), "")
            My.Settings.TTS1voiceA = If(cmbVoice1A.SelectedItem?.ToString(), "")
            My.Settings.TTS1voiceB = If(cmbVoice1B.SelectedItem?.ToString(), "")
            My.Settings.TTS2languagecode = If(cmbLanguage2.SelectedItem?.ToString(), "")
            My.Settings.TTS2voiceA = If(cmbVoice2A.SelectedItem?.ToString(), "")
            My.Settings.TTS2voiceB = If(cmbVoice2B.SelectedItem?.ToString(), "")
            My.Settings.TTSSampleText = If(txtSampleText.Text, "")
            My.Settings.TTSOutputPath = txtOutputPath.Text
            My.Settings.Save()
        End Sub

        ' --- chkTemporary CheckedChanged handler ---
        Private Sub chkTemporary_CheckedChanged(sender As Object, e As EventArgs)
            txtOutputPath.Enabled = Not chkTemporary.Checked
        End Sub

        ''' <summary>
        ''' Sets the output path to the user's Desktop folder while preserving the current filename.
        ''' </summary>
        Private Sub btnDesktop_Click(sender As Object, e As EventArgs)
            ' Get the filename
            Dim fileName As String = System.IO.Path.GetFileName(txtOutputPath.Text)

            ' Get the user's Desktop path
            Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)

            ' Construct new file path
            txtOutputPath.Text = System.IO.Path.Combine(desktopPath, fileName)

        End Sub

    End Class
End Class
