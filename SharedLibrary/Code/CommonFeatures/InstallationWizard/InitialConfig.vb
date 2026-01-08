' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: InitialConfig.vb
' Purpose: First-run installation wizard form for configuring Red Ink LLM API access.
'          Provides provider-specific templates (OpenAI, Azure, Google Vertex, etc.),
'          validates user input, supports remote configuration updates, and generates
'          INI files for Word/Excel/Outlook add-ins.
'
' Architecture Overview:
' ----------------------
' InitialConfig is a modal WinForms wizard dialog typically shown when no Red Ink INI
' configuration file exists for the host application.
'
' Configuration Templates:
' ------------------------
' Provider defaults are built into the form in PrepareConfigData() as a dictionary:
'   providerConfigs(ProviderName) -> List(Of AppConfigurationVariable)
' plus optional providerNotes(ProviderName) displayed below the fields.
'
' Remote Template Override (optional, user-confirmed):
' ---------------------------------------------------
' TryOverrideDefaultsFromRemote() can download an updated template list from
' RemoteDefaultsUrl (custom INI with pipe-delimited field definitions), parse it via
' TryParseRemoteDefaults(), and compare it against built-in defaults via AreDifferent().
' If differences are found, the user can load the remote defaults.
'
' Data Flow:
' ----------
' 1. InitializeComponent() builds the static wizard UI.
' 2. PrepareConfigData() creates provider templates and calls TryOverrideDefaultsFromRemote().
' 3. Provider selection changes rebuild the dynamic UI via LoadConfigForSelectedProvider().
'    Values are preserved across provider switching via SaveCurrentInputToSpecificConfig().
' 4. OK click: SaveCurrentInputToConfig() -> ValidateAllConfigs() -> map config variables to
'    ISharedContext -> CreateAppConfig() writes INI files -> closes wizard.
'
' Validation:
' -----------
' ValidateAllConfigs() applies AppConfigurationVariable.ValidationRule using simple checks:
' NotEmpty, E-Mail ("@" heuristic), Hyperlink (http/https prefix), >0 integer, 0.0-2.0 range,
' and max-value rules expressed as a decimal literal like "2.0" (accepts "." or "," separator).
'
' OAuth2 (Google Vertex):
' -----------------------
' Google Vertex includes additional OAuth2 fields. INI_OAuth2ATExpiry is an integer value in
' seconds (not milliseconds).
'
' File Output:
' ------------
' CreateAppConfig() writes a minimal INI using SharedMethods.GetDefaultINIPath(appName).
'
' Known Comment/Code Mismatches to keep in mind:
' ----------------------------------------------
' - The layout uses a bottom "invisibleLabel" to force height computation; it is currently
'   Visible=True (not invisible).
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Net
Imports System.Net.Http
Imports System.Text
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

Namespace SharedLibrary

    ''' <summary>
    ''' First-run installation wizard for Red Ink LLM configuration.
    ''' Guides users through provider selection, credential entry, and INI file generation.
    ''' </summary>
    ''' <remarks>
    ''' Supports provider-specific templates stored as <see cref="AppConfigurationVariable"/> lists.
    ''' State is preserved when switching providers by copying current TextBox content into each
    ''' variable's <c>CurrentValue</c>.
    ''' </remarks>
    Public Class InitialConfig
        Inherits Form

        ''' <summary>Shared configuration context passed ByRef from host add-in (Word/Excel/Outlook).</summary>
        Private _context As ISharedContext

        ''' <summary>Provider selection dropdown. Items populated from <c>providerConfigs</c> keys.</summary>
        Private WithEvents cmbProvider As ComboBox

        ''' <summary>Checkbox: Apply configuration to Word add-in (uses <c>GetDefaultINIPath("Word")</c>).</summary>
        Private chkWord As System.Windows.Forms.CheckBox

        ''' <summary>Checkbox: Apply configuration to Outlook add-in (uses <c>GetDefaultINIPath("Outlook")</c>).</summary>
        Private chkOutlook As System.Windows.Forms.CheckBox

        ''' <summary>Checkbox: Apply configuration to Excel add-in (uses <c>GetDefaultINIPath("Excel")</c>).</summary>
        Private chkExcel As System.Windows.Forms.CheckBox

        ''' <summary>Scrollable panel containing dynamically generated Label+TextBox pairs for selected provider.</summary>
        Private panelConfig As Panel

        ''' <summary>Label displaying "Configuration For {ProviderName}:" above configuration fields.</summary>
        Private lblCurrentProvider As System.Windows.Forms.Label

        ''' <summary>
        ''' Provider-to-configuration mapping. Key: provider name, Value: list of provider variables.
        ''' </summary>
        Private providerConfigs As New Dictionary(Of String, List(Of AppConfigurationVariable))(StringComparer.OrdinalIgnoreCase)

        ''' <summary>Optional per-provider note displayed at the bottom of the provider field list.</summary>
        Private providerNotes As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        ''' <summary>
        ''' Controls currently displayed in <c>panelConfig</c>.
        ''' Expected layout order: Label, TextBox, Label, TextBox, ... plus optional note Label at the end.
        ''' </summary>
        Private currentConfigControls As New List(Of Control)

        ''' <summary>OK button: Validates input, maps to <see cref="ISharedContext"/>, writes INI files, closes wizard.</summary>
        Private btnOK As Button

        ''' <summary>Cancel button: Sets DialogResult.Cancel, <c>InitialConfigFailed=True</c>, closes wizard.</summary>
        Private btnCancel As Button

        ''' <summary>Target width for form content: <c>Min(OverallWidth + 150, 80% of screen width)</c>.</summary>
        Private ReadOnly _targetWidth As Integer

        ''' <summary>Flag to prevent layout/event handlers from acting during initialization.</summary>
        Private isInitializing As Boolean = False

        ''' <summary>
        ''' Bottom spacer label used to force height recalculation.
        ''' Note: This label is currently configured with <c>Visible=True</c>.
        ''' </summary>
        Private invisibleLabel As New System.Windows.Forms.Label() With {
            .Size = New System.Drawing.Size(1, 10),
            .Visible = True
        }

        ''' <summary>Base width constant used for responsive layout calculations.</summary>
        Private Const OverallWidth As Integer = 900

        ''' <summary>Label for the "Use this config ..." checkbox row (positioned dynamically below <c>panelConfig</c>).</summary>
        Private lblUseThisConfig As System.Windows.Forms.Label

        ''' <summary>
        ''' Tracks the provider whose fields are currently displayed in <c>panelConfig</c>.
        ''' Used to save/restore values during provider switching.
        ''' </summary>
        Private _activeProvider As String = "OpenAI"

        ''' <summary>
        ''' Initializes the wizard with the provided shared configuration context.
        ''' </summary>
        ''' <param name="context">Shared configuration interface passed ByRef from host add-in.</param>
        Public Sub New(ByRef context As ISharedContext)
            _context = context
            _targetWidth = Math.Min(OverallWidth + 150, CInt(Screen.PrimaryScreen.WorkingArea.Width * 0.8))
            Me.Size = New System.Drawing.Size(_targetWidth + 20, 800)
            Me.AutoScroll = False
            Me.AutoSize = True
            Me.InitializeComponent()
            Me.FormBorderStyle = FormBorderStyle.Fixed3D
        End Sub

        ''' <summary>
        ''' Builds the wizard UI by creating and positioning all form controls.
        ''' </summary>
        ''' <remarks>
        ''' This method also builds provider templates via <see cref="PrepareConfigData"/> and then renders the
        ''' first provider's fields via <see cref="LoadConfigForSelectedProvider"/>.
        ''' </remarks>
        Private Sub InitializeComponent()
            isInitializing = True

            ' Configure form properties
            Me.Text = $"{SharedMethods.AN} Initial Configuration Wizard"
            Me.FormBorderStyle = FormBorderStyle.None
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.BackColor = ColorTranslator.FromWin32(&H8000000F)
            Me.ControlBox = False  ' No min/max/close buttons
            Me.AutoScroll = True

            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            Me.Font = standardFont

            ' PictureBox (Logo)
            Dim bmp As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Large))
            Dim pictureBox As New PictureBox() With {
            .Image = bmp,
            .SizeMode = PictureBoxSizeMode.Zoom
        }
            pictureBox.SetBounds(10, 10, 50, 50)
            Me.Controls.Add(pictureBox)

            ' Label "Welcome to {AN}" next to the logo
            Dim lblWelcome As New System.Windows.Forms.Label() With {
            .Text = $"Welcome to {SharedMethods.AN}",
            .AutoSize = True,
            .Font = New System.Drawing.Font("Segoe UI", 12.0F, FontStyle.Bold, GraphicsUnit.Point)
        }
            lblWelcome.Location = New System.Drawing.Point(pictureBox.Right + 10, pictureBox.Top + (pictureBox.Height \ 2) - (lblWelcome.Height \ 2))
            Me.Controls.Add(lblWelcome)

            ' Resolve DefaultINIPath for Word (expanded for display in instructions)
            Dim defaultWordPath As String = ""
            Try
                If SharedMethods.DefaultINIPaths IsNot Nothing AndAlso SharedMethods.DefaultINIPaths.ContainsKey("Word") Then
                    defaultWordPath = SharedMethods.DefaultINIPaths("Word")
                    defaultWordPath = SharedMethods.ExpandEnvironmentVariables(defaultWordPath)
                End If
            Catch
                ' Ignore errors resolving Word path; use empty string
            End Try

            ' LinkLabel with instructions (wraps to target width for responsive layout)
            Dim lblInfo As New LinkLabel() With {
            .AutoSize = True,
            .MaximumSize = New Size(_targetWidth, 0),
            .Text =
                $"No configuration file '{SharedMethods.AN2}.ini' was found, in which all settings " &
                "can be made locally or centrally. Therefore, you can make the basic settings here, " &
                "which will then be saved to such a file. You can then expand it manually (e.g., to add more models); go to 'Settings', then 'Expert Config'. " &
                $"How all this works is explained in the manual, which you can find at {SharedMethods.AN4}." &
                If(String.IsNullOrWhiteSpace(defaultWordPath),
                   "",
                   $" {SharedMethods.AN2} will be stored at {defaultWordPath} for Word, which will also be used by Excel and Outlook unless they have their own {SharedMethods.AN2}.ini.")
        }
            lblInfo.Location = New System.Drawing.Point(10, pictureBox.Bottom + 15)
            AddHandler lblInfo.LinkClicked, AddressOf LinkLabel_LinkClicked
            lblInfo.Links.Add(New LinkLabel.Link() With {
            .LinkData = $"{SharedMethods.AN4}",
            .Start = lblInfo.Text.IndexOf($"{SharedMethods.AN4}", StringComparison.Ordinal),
            .Length = $"{SharedMethods.AN4}".Length
        })
            Me.Controls.Add(lblInfo)

            ' Label + ComboBox "Select API provider:"
            Dim lblWhichAI As New System.Windows.Forms.Label() With {
            .Text = "Select API provider:",
            .AutoSize = True,
            .Font = New System.Drawing.Font(standardFont, FontStyle.Bold)
        }
            lblWhichAI.Location = New System.Drawing.Point(10, lblInfo.Bottom + 20)
            Me.Controls.Add(lblWhichAI)

            ' ComboBox with responsive width (expanded from original 520px, max _targetWidth - label - margins)
            cmbProvider = New ComboBox() With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Width = Math.Min(520 + 150, Math.Max(300, _targetWidth - lblWhichAI.Right - 30))
        }
            cmbProvider.Location = New System.Drawing.Point(lblWhichAI.Right + 10, lblWhichAI.Top - 2)
            Me.Controls.Add(cmbProvider)

            ' Second LinkLabel row (below combo, indented)
            Dim lblMoreInfo As New LinkLabel() With {
            .AutoSize = True,
            .MaximumSize = New Size(_targetWidth - 20, 0),
            .Text = $"Note: More on how to obtain access to one of these providers is described on {SharedMethods.AN4}. Getting an API access is not expensive. You can use the below form also for other providers. If this does not work or you need to configure more, abort and do it manually before restarting your application."
        }
            lblMoreInfo.Location = New System.Drawing.Point(30, cmbProvider.Bottom + 5)
            AddHandler lblMoreInfo.LinkClicked, AddressOf LinkLabel_LinkClicked
            lblMoreInfo.Links.Add(New LinkLabel.Link() With {
            .LinkData = $"{SharedMethods.AN4}",
            .Start = lblMoreInfo.Text.IndexOf($"{SharedMethods.AN4}", StringComparison.Ordinal),
            .Length = $"{SharedMethods.AN4}".Length
        })
            Me.Controls.Add(lblMoreInfo)

            ' Label for "Configuration for <AI Provider>:"
            lblCurrentProvider = New System.Windows.Forms.Label() With {
            .AutoSize = True,
            .Font = New System.Drawing.Font(standardFont, FontStyle.Bold),
            .Location = New System.Drawing.Point(10, lblMoreInfo.Bottom + 20)
        }
            Me.Controls.Add(lblCurrentProvider)

            ' Panel for dynamic input fields (scrollable, responsive width)
            panelConfig = New Panel() With {
            .AutoScroll = True,
            .Location = New System.Drawing.Point(10, lblCurrentProvider.Bottom + 5),
            .Width = _targetWidth
        }
            AddHandler panelConfig.SizeChanged, AddressOf PanelConfig_SizeChanged
            Me.Controls.Add(panelConfig)

            ' Build provider configuration data before populating combo
            PrepareConfigData()

            ' Populate combo in preferred order (OpenAI, Azure, Google Gemini, Google Vertex), then alphabetical
            Dim defaultOrder As New List(Of String) From {
            "OpenAI",
            "Microsoft Azure OpenAI Services",
            "Google Gemini",
            "Google Vertex"
        }
            For Each providerName As String In defaultOrder
                If providerConfigs.ContainsKey(providerName) Then cmbProvider.Items.Add(providerName)
            Next
            For Each providerName As String In providerConfigs.Keys
                If cmbProvider.Items.IndexOf(providerName) = -1 Then cmbProvider.Items.Add(providerName)
            Next
            If cmbProvider.Items.Count > 0 Then
                cmbProvider.SelectedIndex = 0
                _activeProvider = cmbProvider.SelectedItem.ToString()
            End If

            lblUseThisConfig = New System.Windows.Forms.Label() With {
            .Text = $"Use this config for {SharedMethods.AN}:",
            .Font = New System.Drawing.Font(Me.Font, FontStyle.Bold),
            .AutoSize = True
        }
            lblUseThisConfig.Location = New System.Drawing.Point(10, panelConfig.Bottom + 10)
            Me.Controls.Add(lblUseThisConfig)

            chkWord = New System.Windows.Forms.CheckBox() With {
            .Text = "for Word",
            .AutoSize = True,
            .Checked = _context.RDV.StartsWith("Word")
        }
            chkWord.Location = New System.Drawing.Point(lblUseThisConfig.Right + 10, lblUseThisConfig.Top)
            Me.Controls.Add(chkWord)

            chkOutlook = New System.Windows.Forms.CheckBox() With {
            .Text = "for Outlook (as separate config)",
            .AutoSize = True,
            .Checked = _context.RDV.StartsWith("Outlook")
        }
            chkOutlook.Location = New System.Drawing.Point(chkWord.Right + 17, lblUseThisConfig.Top)
            Me.Controls.Add(chkOutlook)

            chkExcel = New System.Windows.Forms.CheckBox() With {
            .Text = "for Excel (as separate config)",
            .AutoSize = True,
            .Checked = _context.RDV.StartsWith("Excel")
        }
            chkExcel.Location = New System.Drawing.Point(chkOutlook.Right + 17, lblUseThisConfig.Top)
            Me.Controls.Add(chkExcel)

            btnOK = New Button() With {
            .Text = "OK, save this configuration and continue",
            .AutoSize = True
        }
            btnOK.Location = New System.Drawing.Point(10, lblUseThisConfig.Bottom + 20)
            AddHandler btnOK.Click, AddressOf btnOK_Click
            Me.Controls.Add(btnOK)

            btnCancel = New Button() With {
            .Text = "Cancel",
            .AutoSize = True
        }
            btnCancel.Location = New System.Drawing.Point(btnOK.Right + 10, btnOK.Top)
            AddHandler btnCancel.Click, AddressOf btnCancel_Click
            Me.Controls.Add(btnCancel)

            invisibleLabel.Location = New System.Drawing.Point(10, btnCancel.Bottom + 10)
            Me.Controls.Add(invisibleLabel)

            LoadConfigForSelectedProvider()
            isInitializing = False
        End Sub

        ''' <summary>
        ''' Handles panel size changes by repositioning controls below <c>panelConfig</c> and adjusting the form height.
        ''' </summary>
        ''' <param name="sender">The <see cref="Panel"/> whose size changed.</param>
        ''' <param name="e">Event arguments.</param>
        Private Sub PanelConfig_SizeChanged(sender As Object, e As EventArgs)
            If isInitializing OrElse lblUseThisConfig Is Nothing Then Exit Sub
            Dim panel = DirectCast(sender, Panel)
            lblUseThisConfig.Location = New System.Drawing.Point(10, panel.Bottom + 20)
            chkWord.Location = New System.Drawing.Point(lblUseThisConfig.Right + 10, lblUseThisConfig.Top)
            chkOutlook.Location = New System.Drawing.Point(chkWord.Right + 20, lblUseThisConfig.Top)
            chkExcel.Location = New System.Drawing.Point(chkOutlook.Right + 20, lblUseThisConfig.Top)
            btnOK.Location = New System.Drawing.Point(10, lblUseThisConfig.Bottom + 20)
            btnCancel.Location = New System.Drawing.Point(btnOK.Right + 10, btnOK.Top)
            invisibleLabel.Location = New System.Drawing.Point(10, btnCancel.Bottom + 10)
            Me.Height = invisibleLabel.Bottom + 20
        End Sub

        ''' <summary>
        ''' Builds provider configuration templates with default values and optional notes.
        ''' </summary>
        ''' <remarks>
        ''' Creates templates for 6 providers (OpenAI, Azure, Google Gemini, Google Vertex, MTF, SafeSwissCloud).
        ''' Each has 9-13 fields. Google Vertex includes 4 OAuth2 fields.
        ''' Temperature validation is implemented either via explicit range rule "0.0-2.0" or via max-value literals (e.g., "2.0").
        ''' Calls <see cref="TryOverrideDefaultsFromRemote"/> to optionally replace built-in defaults with remote defaults.
        ''' </remarks>
        Private Sub PrepareConfigData()
            providerConfigs.Clear()
            providerNotes.Clear()

            ' Helper lambda to add provider with cloned variable list (prevents reference sharing)
            Dim SubAdd As Action(Of String, List(Of AppConfigurationVariable)) =
                Sub(name As String, vars As List(Of AppConfigurationVariable))
                    Dim clone As New List(Of AppConfigurationVariable)
                    For Each v In vars
                        clone.Add(New AppConfigurationVariable With {
                            .DisplayName = v.DisplayName,
                            .VarName = v.VarName,
                            .VarType = v.VarType,
                            .ValidationRule = v.ValidationRule,
                            .DefaultValue = v.DefaultValue,
                            .CurrentValue = v.DefaultValue
                        })
                    Next
                    providerConfigs(name) = clone
                End Sub

            ' OPENAI provider (10 fields + payment note)
            SubAdd("OpenAI",
                New List(Of AppConfigurationVariable) From {
                    New AppConfigurationVariable With {.DisplayName = "API Key:", .VarName = "INI_APIKey", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "[[Your OpenAI API Key]]"},
                    New AppConfigurationVariable With {.DisplayName = "Timeout (ms):", .VarName = "INI_Timeout", .VarType = "Integer", .ValidationRule = ">0", .DefaultValue = "200000"},
                    New AppConfigurationVariable With {.DisplayName = "Model:", .VarName = "INI_Model", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "gpt-5.2"},
                    New AppConfigurationVariable With {.DisplayName = "Endpoint:", .VarName = "INI_Endpoint", .VarType = "String", .ValidationRule = "Hyperlink", .DefaultValue = "https://api.openai.com/v1/responses"},
                    New AppConfigurationVariable With {.DisplayName = "HeaderA:", .VarName = "INI_HeaderA", .VarType = "String", .ValidationRule = "", .DefaultValue = "Authorization"},
                    New AppConfigurationVariable With {.DisplayName = "HeaderB:", .VarName = "INI_HeaderB", .VarType = "String", .ValidationRule = "", .DefaultValue = "Bearer {apikey}"},
                    New AppConfigurationVariable With {.DisplayName = "APICall:", .VarName = "INI_APICall", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "{""model"": ""{model}"", ""input"": [{""role"": ""developer"", ""content"": [{""type"": ""input_text"",""text"": ""{promptsystem}""}]},{""role"": ""user"",""content"": [{""type"": ""input_text"",""text"": ""{promptuser}""}{objectcall}]}],""reasoning"": {""effort"": ""low""}}"},
                    New AppConfigurationVariable With {.DisplayName = "APICall_Object:", .VarName = "INI_APICall_Object", .VarType = "String", .ValidationRule = "", .DefaultValue = "[application/pdf],{""type"": ""input_file"",""filename"": ""userfile.pdf"", ""file_data"": ""data:{mimetype};base64,{encodeddata}""}¦[image/png,image/jpeg,image/webp, image/gif],{""type"": ""input_image"",""image_url"": ""data:{mimetype};base64,{encodeddata}""}"},
                    New AppConfigurationVariable With {.DisplayName = "Response tag:", .VarName = "INI_Response", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "text (rkmode_first)"}
                })
            providerNotes("OpenAI") = "Note: When generating the API key with OpenAI, make sure you have added a valid payment method (e.g., credit card), even if you use ChatGPT for free or with an already paid subscription. You still need the payment method and a budget to pay for the actual consumption (costs are in our experience low)."

            ' MICROSOFT AZURE OPENAI SERVICES provider (9 fields, no note)
            SubAdd("Microsoft Azure OpenAI Services",
                New List(Of AppConfigurationVariable) From {
                    New AppConfigurationVariable With {.DisplayName = "API Key:", .VarName = "INI_APIKey", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = ""},
                    New AppConfigurationVariable With {.DisplayName = "Temperature:", .VarName = "INI_Temperature", .VarType = "String", .ValidationRule = "0.0-2.0", .DefaultValue = "1.0"},
                    New AppConfigurationVariable With {.DisplayName = "Timeout (ms):", .VarName = "INI_Timeout", .VarType = "Integer", .ValidationRule = ">0", .DefaultValue = "200000"},
                    New AppConfigurationVariable With {.DisplayName = "Model:", .VarName = "INI_Model", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "gpt-5.2"},
                    New AppConfigurationVariable With {.DisplayName = "Endpoint:", .VarName = "INI_Endpoint", .VarType = "String", .ValidationRule = "Hyperlink", .DefaultValue = "https://[[Your Azure Endpoint]]/openai/deployments/[your deployment-id]/chat/completions?api-version=2024-06-01"},
                    New AppConfigurationVariable With {.DisplayName = "HeaderA:", .VarName = "INI_HeaderA", .VarType = "String", .ValidationRule = "", .DefaultValue = "api-key"},
                    New AppConfigurationVariable With {.DisplayName = "HeaderB:", .VarName = "INI_HeaderB", .VarType = "String", .ValidationRule = "", .DefaultValue = "{apikey}"},
                    New AppConfigurationVariable With {.DisplayName = "APICall:", .VarName = "INI_APICall", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "{""messages"": [{""role"": ""system"",""content"": ""{promptsystem}""},{""role"": ""user"", ""content"": ""{promptuser}""}],""temperature"": {temperature}}"},
                    New AppConfigurationVariable With {.DisplayName = "Response tag:", .VarName = "INI_Response", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "content (rkmode_first)"}
                })
            providerNotes("Microsoft Azure OpenAI Services") = "These may not be the latest possible settings available in your environment. Check your documentation to update this."

            ' GOOGLE GEMINI provider (9 fields, no note)
            SubAdd("Google Gemini",
                New List(Of AppConfigurationVariable) From {
                    New AppConfigurationVariable With {.DisplayName = "API Key:", .VarName = "INI_APIKey", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "[[Your Google API Key]]"},
                    New AppConfigurationVariable With {.DisplayName = "Temperature:", .VarName = "INI_Temperature", .VarType = "String", .ValidationRule = "0.0-2.0", .DefaultValue = "0.2"},
                    New AppConfigurationVariable With {.DisplayName = "Timeout (ms):", .VarName = "INI_Timeout", .VarType = "Integer", .ValidationRule = ">0", .DefaultValue = "200000"},
                    New AppConfigurationVariable With {.DisplayName = "Model:", .VarName = "INI_Model", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "gemini-2.5-pro"},
                    New AppConfigurationVariable With {.DisplayName = "Endpoint:", .VarName = "INI_Endpoint", .VarType = "String", .ValidationRule = "Hyperlink", .DefaultValue = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={apikey}"},
                    New AppConfigurationVariable With {.DisplayName = "HeaderA:", .VarName = "INI_HeaderA", .VarType = "String", .ValidationRule = "", .DefaultValue = "X-Goog-Api-Key"},
                    New AppConfigurationVariable With {.DisplayName = "HeaderB:", .VarName = "INI_HeaderB", .VarType = "String", .ValidationRule = "", .DefaultValue = "{apikey}"},
                    New AppConfigurationVariable With {.DisplayName = "APICall:", .VarName = "INI_APICall", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "{""contents"": [{""role"": ""user"",""parts"": [{ ""text"": ""{promptsystem} {promptuser}"" }]}], ""generationConfig"": {""temperature"": {temperature}}}"},
                    New AppConfigurationVariable With {.DisplayName = "Response tag:", .VarName = "INI_Response", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "text"}
                })
            providerNotes("Google Gemini") = ""

            ' GOOGLE VERTEX provider (14 fields: 10 common + APICall_Object + 4 OAuth2, includes service account note)
            SubAdd("Google Vertex",
                New List(Of AppConfigurationVariable) From {
                    New AppConfigurationVariable With {.DisplayName = "Private Key (barebones, not PEM):", .VarName = "INI_APIKey", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "[[Your Google private_key]]"},
                    New AppConfigurationVariable With {.DisplayName = "Temperature:", .VarName = "INI_Temperature", .VarType = "String", .ValidationRule = "2.0", .DefaultValue = "0.2"},
                    New AppConfigurationVariable With {.DisplayName = "Timeout (ms):", .VarName = "INI_Timeout", .VarType = "Integer", .ValidationRule = ">0", .DefaultValue = "200000"},
                    New AppConfigurationVariable With {.DisplayName = "Model:", .VarName = "INI_Model", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "gemini-2.5-pro"},
                    New AppConfigurationVariable With {.DisplayName = "Endpoint:", .VarName = "INI_Endpoint", .VarType = "String", .ValidationRule = "Hyperlink", .DefaultValue = "https://europe-west4-aiplatform.googleapis.com/v1/projects/[[Your Google project_id]]/locations/europe-west4/publishers/google/models/{model}:generateContent"},
                    New AppConfigurationVariable With {.DisplayName = "HeaderA:", .VarName = "INI_HeaderA", .VarType = "String", .ValidationRule = "", .DefaultValue = "Authorization"},
                    New AppConfigurationVariable With {.DisplayName = "HeaderB:", .VarName = "INI_HeaderB", .VarType = "String", .ValidationRule = "", .DefaultValue = "Bearer {apikey}"},
                    New AppConfigurationVariable With {.DisplayName = "APICall:", .VarName = "INI_APICall", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "{""contents"": [{""role"": ""user"", ""parts"":[{""text"": ""{promptsystem} {promptuser}""}{objectcall}]}], ""generationConfig"": {""temperature"": {temperature},  ""thinking_config"": {""thinking_budget"": 128}}}"},
                    New AppConfigurationVariable With {.DisplayName = "APICall_Object:", .VarName = "INI_APICall_Object", .VarType = "String", .ValidationRule = "", .DefaultValue = ", {""inlineData"": {""mimeType"": ""{mimetype}"",""data"": ""{encodeddata}""}}"},
                    New AppConfigurationVariable With {.DisplayName = "Response tag:", .VarName = "INI_Response", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "text"},
                    New AppConfigurationVariable With {.DisplayName = "OAuth2 'client_mail':", .VarName = "INI_OAuth2ClientMail", .VarType = "String", .ValidationRule = "E-Mail", .DefaultValue = "[[Your Google client_email]]"},
                    New AppConfigurationVariable With {.DisplayName = "OAuth2 'scopes':", .VarName = "INI_OAuth2Scopes", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "https://www.googleapis.com/auth/cloud-platform"},
                    New AppConfigurationVariable With {.DisplayName = "OAuth2 Endpoint:", .VarName = "INI_OAuth2Endpoint", .VarType = "String", .ValidationRule = "Hyperlink", .DefaultValue = "https://oauth2.googleapis.com/token"},
                    New AppConfigurationVariable With {.DisplayName = "OAuth2 Access Token Expiry (seconds):", .VarName = "INI_OAuth2ATExpiry", .VarType = "Integer", .ValidationRule = ">0", .DefaultValue = "3600"}
                })
            providerNotes("Google Vertex") = "Note: Requires OAuth2 service account to be configured via the GCP console. Private Key must be the raw key (not PEM)."
            ' MTF provider (9 fields, no note)
            SubAdd("MTF",
                New List(Of AppConfigurationVariable) From {
                    New AppConfigurationVariable With {.DisplayName = "API Key:", .VarName = "INI_APIKey", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = ""},
                    New AppConfigurationVariable With {.DisplayName = "Temperature:", .VarName = "INI_Temperature", .VarType = "String", .ValidationRule = "2.0", .DefaultValue = "0.2"},
                    New AppConfigurationVariable With {.DisplayName = "Timeout (ms):", .VarName = "INI_Timeout", .VarType = "Integer", .ValidationRule = ">0", .DefaultValue = "200000"},
                    New AppConfigurationVariable With {.DisplayName = "Model:", .VarName = "INI_Model", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "meta-llama-ai"},
                    New AppConfigurationVariable With {.DisplayName = "Endpoint:", .VarName = "INI_Endpoint", .VarType = "String", .ValidationRule = "Hyperlink", .DefaultValue = "https://api.ai.mtf.cloud/chatbot/ask"},
                    New AppConfigurationVariable With {.DisplayName = "HeaderA:", .VarName = "INI_HeaderA", .VarType = "String", .ValidationRule = "", .DefaultValue = "Authorization"},
                    New AppConfigurationVariable With {.DisplayName = "HeaderB:", .VarName = "INI_HeaderB", .VarType = "String", .ValidationRule = "", .DefaultValue = "Bearer {apikey}"},
                    New AppConfigurationVariable With {.DisplayName = "APICall:", .VarName = "INI_APICall", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "{""model"":   ""{model}"",  ""messages"": [{""role"": ""system"",""content"": ""{promptsystem}""},{""role"": ""user"",""content"": ""{promptuser}""}],""temperature"": {temperature}}"},
                    New AppConfigurationVariable With {.DisplayName = "Response tag:", .VarName = "INI_Response", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "content"}
                })
            providerNotes("MTF") = ""

            ' SAFESWISSCLOUD provider (9 fields, no note)
            SubAdd("SafeSwissCloud",
                New List(Of AppConfigurationVariable) From {
                    New AppConfigurationVariable With {.DisplayName = "API Key:", .VarName = "INI_APIKey", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = ""},
                    New AppConfigurationVariable With {.DisplayName = "Temperature:", .VarName = "INI_Temperature", .VarType = "String", .ValidationRule = "0.0-2.0", .DefaultValue = "0.2"},
                    New AppConfigurationVariable With {.DisplayName = "Timeout (ms):", .VarName = "INI_Timeout", .VarType = "Integer", .ValidationRule = ">0", .DefaultValue = "200000"},
                    New AppConfigurationVariable With {.DisplayName = "Model:", .VarName = "INI_Model", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "gpt-oss-120b"},
                    New AppConfigurationVariable With {.DisplayName = "Endpoint:", .VarName = "INI_Endpoint", .VarType = "String", .ValidationRule = "Hyperlink", .DefaultValue = "https://llm01.safeswisscloud.ch/engines/{model}/chat/completions"},
                    New AppConfigurationVariable With {.DisplayName = "HeaderA:", .VarName = "INI_HeaderA", .VarType = "String", .ValidationRule = "", .DefaultValue = "Authorization"},
                    New AppConfigurationVariable With {.DisplayName = "HeaderB:", .VarName = "INI_HeaderB", .VarType = "String", .ValidationRule = "", .DefaultValue = "Bearer {apikey}"},
                    New AppConfigurationVariable With {.DisplayName = "APICall:", .VarName = "INI_APICall", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "{""model"":   ""{model}"",  ""messages"": [{""role"": ""system"",""content"": ""{promptsystem}""},{""role"": ""user"",""content"": ""{promptuser}""}],""temperature"": {temperature}}"},
                    New AppConfigurationVariable With {.DisplayName = "Response tag:", .VarName = "INI_Response", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "content"}
                })
            providerNotes("SafeSwissCloud") = ""

            ' Attempt to override with remote configuration if available and different
            TryOverrideDefaultsFromRemote()

        End Sub

        ''' <summary>
        ''' Downloads text content from a URL with a timeout.
        ''' </summary>
        ''' <param name="url">URL to download.</param>
        ''' <param name="timeoutMs">Timeout in milliseconds. A minimum timeout of 10 seconds is enforced.</param>
        ''' <param name="content">On success, receives the downloaded text.</param>
        ''' <returns>True when non-empty content was downloaded; otherwise False.</returns>
        Private Function TryDownloadString(url As String, timeoutMs As Integer, ByRef content As String) As Boolean
            content = Nothing
            Try
                ' Ensure TLS 1.2 for HTTPS endpoints (many servers reject TLS 1.0/1.1)
                ServicePointManager.SecurityProtocol = ServicePointManager.SecurityProtocol Or SecurityProtocolType.Tls12

                Dim handler As New HttpClientHandler() With {
                .AutomaticDecompression = DecompressionMethods.GZip Or DecompressionMethods.Deflate
            }

                Using client As New System.Net.Http.HttpClient(handler)
                    client.Timeout = TimeSpan.FromMilliseconds(Math.Max(10000, timeoutMs)) ' 10s minimum

                    ' Use GetStringAsync for simplicity (auto-detects encoding from Content-Type header)
                    Dim readTask = client.GetStringAsync(url)
                    readTask.Wait()

                    If readTask.Status = TaskStatus.RanToCompletion Then
                        Dim s = readTask.Result
                        If Not String.IsNullOrWhiteSpace(s) Then
                            content = s
                            Return True
                        End If
                    End If
                End Using
            Catch ex As Exception
                ' Log error for diagnostics (visible in Output window during debugging)
                System.Diagnostics.Debug.WriteLine($"TryDownloadString error for {url}: {ex}")
            End Try
            Return False
        End Function

        ''' <summary>
        ''' Parses remote INI-format configuration text into provider dictionaries.
        ''' </summary>
        ''' <param name="ini">Remote INI text.</param>
        ''' <param name="outConfigs">On success, receives provider configuration templates.</param>
        ''' <param name="outNotes">On success, receives optional provider notes.</param>
        ''' <returns>True on successful parse with at least one provider; otherwise False.</returns>
        ''' <remarks>
        ''' Format per section:
        ''' <c>[ProviderName]</c>
        ''' <c>FieldN = DisplayName|VarName|VarType|ValidationRule|DefaultValue</c>
        ''' <c>Note = ...</c>
        ''' Lines starting with <c>;</c> or <c>#</c> are ignored.
        ''' </remarks>
        Private Function TryParseRemoteDefaults(ini As String,
                                            ByRef outConfigs As Dictionary(Of String, List(Of AppConfigurationVariable)),
                                            ByRef outNotes As Dictionary(Of String, String)) As Boolean
            outConfigs = Nothing
            outNotes = Nothing
            If String.IsNullOrWhiteSpace(ini) Then Return False

            Try
                Dim cfg As New Dictionary(Of String, List(Of AppConfigurationVariable))(StringComparer.OrdinalIgnoreCase)
                Dim notes As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

                Dim lines = ini.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split(New Char() {ChrW(10)}, StringSplitOptions.None)
                Dim section As String = Nothing
                Dim sectionFields As New List(Of KeyValuePair(Of String, String))()

                Dim flushSection As Action =
                Sub()
                    If String.IsNullOrWhiteSpace(section) Then Return

                    Dim vars As New List(Of AppConfigurationVariable)()

                    ' Gather FieldN entries in numeric order (not alphabetical)
                    Dim ordered = sectionFields.
                        Where(Function(kv) kv.Key.StartsWith("Field", StringComparison.OrdinalIgnoreCase)).
                        Select(Function(kv)
                                   Dim numStr = New String(kv.Key.SkipWhile(Function(c) Not Char.IsDigit(c)).ToArray())
                                   Dim n As Integer = 0
                                   Integer.TryParse(numStr, n)
                                   Return New With {.Num = n, .Val = kv.Value}
                               End Function).
                        OrderBy(Function(x) x.Num).
                        ToList()

                    For Each f In ordered
                        Dim parts = (If(f.Val, "")).Split("|"c)
                        ' Expected format: DisplayName|VarName|VarType|ValidationRule|DefaultValue
                        If parts.Length >= 4 Then
                            Dim v As New AppConfigurationVariable With {
                                .DisplayName = parts(0).Trim(),
                                .VarName = parts(1).Trim(),
                                .VarType = parts(2).Trim(),
                                .ValidationRule = parts(3).Trim(),
                                .DefaultValue = If(parts.Length >= 5, parts(4), "")
                            }
                            v.CurrentValue = v.DefaultValue
                            vars.Add(v)
                        End If
                    Next

                    ' Extract optional Note entry
                    Dim noteValue = sectionFields.
                        FirstOrDefault(Function(kv) kv.Key.Equals("Note", StringComparison.OrdinalIgnoreCase)).Value
                    If Not String.IsNullOrWhiteSpace(noteValue) Then
                        notes(section) = noteValue.Trim()
                    End If

                    If vars.Count > 0 Then
                        cfg(section) = vars
                    End If

                    sectionFields.Clear()
                End Sub

                For Each raw In lines
                    Dim line = raw.Trim()
                    If line.Length = 0 Then Continue For
                    If line.StartsWith(";", StringComparison.Ordinal) OrElse line.StartsWith("#", StringComparison.Ordinal) Then Continue For

                    If line.StartsWith("[") AndAlso line.EndsWith("]") Then
                        flushSection()
                        section = line.Substring(1, line.Length - 2).Trim()
                        Continue For
                    End If

                    Dim eqIdx = line.IndexOf("="c)
                    If eqIdx > 0 AndAlso section IsNot Nothing Then
                        Dim key = line.Substring(0, eqIdx).Trim()
                        Dim val = line.Substring(eqIdx + 1).Trim()
                        sectionFields.Add(New KeyValuePair(Of String, String)(key, val))
                    End If
                Next

                flushSection()

                If cfg.Count > 0 Then
                    outConfigs = cfg
                    outNotes = notes
                    Return True
                End If
            Catch
                ' Parsing failure -> silently ignore, return False
            End Try

            Return False
        End Function

        ''' <summary>
        ''' Compares two strings for equality using normalization suitable for JSON-like payloads.
        ''' </summary>
        ''' <param name="a">First string.</param>
        ''' <param name="b">Second string.</param>
        ''' <returns>True if values are considered equal; otherwise False.</returns>
        Private Function StringsEqual(a As String, b As String) As Boolean
            If Object.ReferenceEquals(a, b) Then Return True
            If a Is Nothing OrElse b Is Nothing Then
                Return String.IsNullOrEmpty(a) AndAlso String.IsNullOrEmpty(b)
            End If

            Dim sa = a.Replace(vbCrLf, vbLf).Trim()
            Dim sb = b.Replace(vbCrLf, vbLf).Trim()

            If String.Equals(sa, sb, StringComparison.Ordinal) Then Return True

            ' Heuristic: if both strings look JSON-ish, compare ignoring whitespace outside quotes
            Dim looksJsonA = (sa.IndexOf(":"c) >= 0) AndAlso (sa.Contains("{") OrElse sa.Contains("["))
            Dim looksJsonB = (sb.IndexOf(":"c) >= 0) AndAlso (sb.Contains("{") OrElse sb.Contains("["))

            If looksJsonA AndAlso looksJsonB Then
                Return String.Equals(StripWsOutsideQuotes(sa), StripWsOutsideQuotes(sb), StringComparison.Ordinal)
            End If

            Return False
        End Function

        ''' <summary>
        ''' Removes whitespace characters occurring outside of quoted strings.
        ''' </summary>
        ''' <param name="s">Input string.</param>
        ''' <returns>Normalized string without whitespace outside quotes.</returns>
        ''' <remarks>
        ''' Used by <see cref="StringsEqual"/> to compare JSON-like text while ignoring formatting differences.
        ''' </remarks>
        Private Function StripWsOutsideQuotes(s As String) As String
            Dim sb As New StringBuilder(s.Length)
            Dim inStr As Boolean = False
            Dim esc As Boolean = False

            For i As Integer = 0 To s.Length - 1
                Dim ch = s(i)
                If inStr Then
                    sb.Append(ch)
                    If esc Then
                        esc = False
                    ElseIf ch = "\"c Then
                        esc = True
                    ElseIf ch = """"c Then
                        inStr = False
                    End If
                Else
                    If ch = """"c Then
                        inStr = True
                        sb.Append(ch)
                    ElseIf Not Char.IsWhiteSpace(ch) Then
                        sb.Append(ch)
                    End If
                End If
            Next

            Return sb.ToString()
        End Function

        ''' <summary>
        ''' Performs deep comparison of local vs. remote provider configurations.
        ''' </summary>
        ''' <param name="localCfg">Local provider templates.</param>
        ''' <param name="localNotes">Local provider notes.</param>
        ''' <param name="remoteCfg">Remote provider templates.</param>
        ''' <param name="remoteNotes">Remote provider notes.</param>
        ''' <returns>True if differences are detected; otherwise False.</returns>
        Private Function AreDifferent(localCfg As Dictionary(Of String, List(Of AppConfigurationVariable)),
                                  localNotes As Dictionary(Of String, String),
                                  remoteCfg As Dictionary(Of String, List(Of AppConfigurationVariable)),
                                  remoteNotes As Dictionary(Of String, String)) As Boolean
            If remoteCfg Is Nothing OrElse remoteNotes Is Nothing Then
                System.Diagnostics.Debug.WriteLine("No remote config/notes available; treating as 'no differences'.")
                Return False
            End If

            Dim foundDiff As Boolean = False
            Dim S As Func(Of String, String) = Function(x) If(x, "<null>")

            ' Check for providers added/removed
            For Each p In remoteCfg.Keys
                If Not localCfg.ContainsKey(p) Then
                    System.Diagnostics.Debug.WriteLine($"Difference: provider present in remote but missing locally: '{p}'")
                    foundDiff = True
                End If
            Next
            For Each p In localCfg.Keys
                If Not remoteCfg.ContainsKey(p) Then
                    System.Diagnostics.Debug.WriteLine($"Difference: provider present locally but missing in remote: '{p}'")
                    foundDiff = True
                End If
            Next

            ' Per-provider comparison: check variables (added/removed/changed)
            For Each p In remoteCfg.Keys
                Dim rList = remoteCfg(p)
                Dim lList As List(Of AppConfigurationVariable) = Nothing
                If Not localCfg.TryGetValue(p, lList) Then
                    System.Diagnostics.Debug.WriteLine($"Difference: provider '{p}' exists in remote but not found in local config map.")
                    foundDiff = True
                    Continue For
                End If

                Dim rByName = rList.ToDictionary(Function(v) v.VarName, StringComparer.OrdinalIgnoreCase)
                Dim lByName = lList.ToDictionary(Function(v) v.VarName, StringComparer.OrdinalIgnoreCase)

                ' Check variables added/removed
                For Each k In rByName.Keys
                    If Not lByName.ContainsKey(k) Then
                        System.Diagnostics.Debug.WriteLine($"Difference[{p}]: variable added in remote: '{k}'")
                        foundDiff = True
                    End If
                Next
                For Each k In lByName.Keys
                    If Not rByName.ContainsKey(k) Then
                        System.Diagnostics.Debug.WriteLine($"Difference[{p}]: variable removed in remote: '{k}'")
                        foundDiff = True
                    End If
                Next

                ' Field-by-field comparison (including CurrentValue)
                For Each k In rByName.Keys
                    If Not lByName.ContainsKey(k) Then
                        Continue For
                    End If

                    Dim r = rByName(k)
                    Dim l = lByName(k)

                    If Not StringsEqual(l.DisplayName, r.DisplayName) Then
                        System.Diagnostics.Debug.WriteLine($"Difference[{p}.{k}]: DisplayName local='{S(l.DisplayName)}' remote='{S(r.DisplayName)}'")
                        foundDiff = True
                    End If
                    If Not StringsEqual(l.VarType, r.VarType) Then
                        System.Diagnostics.Debug.WriteLine($"Difference[{p}.{k}]: VarType local='{S(l.VarType)}' remote='{S(r.VarType)}'")
                        foundDiff = True
                    End If
                    If Not StringsEqual(l.ValidationRule, r.ValidationRule) Then
                        System.Diagnostics.Debug.WriteLine($"Difference[{p}.{k}]: ValidationRule local='{S(l.ValidationRule)}' remote='{S(r.ValidationRule)}'")
                        foundDiff = True
                    End If
                    If Not StringsEqual(l.DefaultValue, r.DefaultValue) Then
                        System.Diagnostics.Debug.WriteLine($"Difference[{p}.{k}]: DefaultValue local='{S(l.DefaultValue)}' remote='{S(r.DefaultValue)}'")
                        foundDiff = True
                    End If
                    If Not StringsEqual(l.CurrentValue, r.CurrentValue) Then
                        System.Diagnostics.Debug.WriteLine($"Difference[{p}.{k}]: CurrentValue local='{S(l.CurrentValue)}' remote='{S(r.CurrentValue)}'")
                        foundDiff = True
                    End If
                Next
            Next

            ' Compare notes: check union of providers to catch added/removed notes
            Dim allProviders = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each k In localCfg.Keys : allProviders.Add(k) : Next
            For Each k In remoteCfg.Keys : allProviders.Add(k) : Next

            For Each p In allProviders
                Dim ln As String = Nothing
                Dim rn As String = Nothing
                localNotes.TryGetValue(p, ln)
                remoteNotes.TryGetValue(p, rn)
                If Not StringsEqual(ln, rn) Then
                    System.Diagnostics.Debug.WriteLine($"Difference[{p}]: Provider note changed. localLen={If(ln, "").Length} remoteLen={If(rn, "").Length}")
                    foundDiff = True
                End If
            Next

            If Not foundDiff Then
                System.Diagnostics.Debug.WriteLine("No differences detected between local and remote config/notes.")
            End If

            Return foundDiff
        End Function

        ''' <summary>
        ''' Attempts to download and apply remote provider configuration defaults.
        ''' </summary>
        ''' <remarks>
        ''' Uses <c>RemoteDefaultsUrl</c> as the remote source and asks the user before downloading.
        ''' If remote defaults differ from built-in defaults, the user can choose to replace the built-in maps.
        ''' Network and parsing failures are caught and ignored.
        ''' </remarks>
        Private Sub TryOverrideDefaultsFromRemote()

            Dim answer = ShowCustomYesNoBox($"You are about to run the {SharedMethods.AN} Installation Wizard. Do you want to check on {RemoteDefaultsUrl} for updated default configuration information?", "Yes", "No, keep built-in")

            If answer <> 1 Then Return

            Try
                Dim remoteText As String = Nothing
                If Not TryDownloadString(RemoteDefaultsUrl, 10000, remoteText) Then Exit Sub

                Dim rCfg As Dictionary(Of String, List(Of AppConfigurationVariable)) = Nothing
                Dim rNotes As Dictionary(Of String, String) = Nothing
                If Not TryParseRemoteDefaults(remoteText, rCfg, rNotes) Then Exit Sub

                If rCfg Is Nothing OrElse rCfg.Count = 0 Then Exit Sub

                If AreDifferent(providerConfigs, providerNotes, rCfg, rNotes) Then
                    Dim choice = SharedMethods.ShowCustomYesNoBox(
                    "Updated Default provider configurations are available online. Do you want To load and use those instead Of the built-In defaults now?",
                    "Use online defaults",
                    "Keep built-In")

                    ' Convention: return 1 for first button ("Use online defaults")
                    If choice = 1 Then
                        providerConfigs = New Dictionary(Of String, List(Of AppConfigurationVariable))(rCfg, StringComparer.OrdinalIgnoreCase)
                        providerNotes = New Dictionary(Of String, String)(rNotes, StringComparer.OrdinalIgnoreCase)
                    End If
                Else
                    ShowCustomMessageBox("No updates found. Keeping built-in defaults.")

                End If
            Catch
                ' Never fail the wizard on remote errors (network, parsing, comparison exceptions all ignored)
            End Try
        End Sub

        ''' <summary>
        ''' Handles provider selection changes by saving current inputs, switching the active provider, and rebuilding the provider UI.
        ''' </summary>
        Private Sub cmbProvider_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbProvider.SelectedIndexChanged
            Try
                ' Step 1: Save values into the previously displayed provider
                Dim prevList As List(Of AppConfigurationVariable) = GetConfigListByName(_activeProvider)
                SaveCurrentInputToSpecificConfig(prevList)

                ' Step 2: Switch active provider to the newly selected one
                If cmbProvider.SelectedItem IsNot Nothing Then
                    _activeProvider = cmbProvider.SelectedItem.ToString()
                End If

                ' Step 3: Load UI for new provider
                LoadConfigForSelectedProvider()
            Catch
                ' Ignore minor UI timing issues (silent failure acceptable)
            End Try
        End Sub

        ''' <summary>
        ''' Copies the current panel input into <see cref="AppConfigurationVariable.CurrentValue"/> for the active provider.
        ''' </summary>
        ''' <remarks>
        ''' Relies on <c>currentConfigControls</c> being ordered as Label/TextBox pairs and matches variables by <c>DisplayName</c>.
        ''' </remarks>
        Private Sub SaveCurrentInputToConfig()
            Dim selectedList = GetSelectedConfigList()
            If selectedList Is Nothing OrElse currentConfigControls.Count = 0 Then Return

            For i As Integer = 0 To currentConfigControls.Count - 1
                Dim ctrl = currentConfigControls(i)
                If TypeOf ctrl Is System.Windows.Forms.Label Then
                    Dim labelText = CType(ctrl, System.Windows.Forms.Label).Text
                    Dim configVar = selectedList.FirstOrDefault(Function(x) x.DisplayName = labelText)
                    If configVar IsNot Nothing Then
                        If i + 1 < currentConfigControls.Count Then
                            Dim inputControl = currentConfigControls(i + 1)
                            If TypeOf inputControl Is TextBox Then
                                configVar.CurrentValue = CType(inputControl, TextBox).Text
                            End If
                        End If
                    End If
                End If
            Next
        End Sub

        ''' <summary>
        ''' Saves current UI inputs into a specified provider config list during provider switching.
        ''' </summary>
        ''' <param name="targetConfig">Provider variable list to update.</param>
        ''' <remarks>
        ''' Relies on <c>currentConfigControls</c> being ordered as Label/TextBox pairs and matches variables by <c>DisplayName</c>.
        ''' </remarks>
        Private Sub SaveCurrentInputToSpecificConfig(targetConfig As List(Of AppConfigurationVariable))
            If targetConfig Is Nothing OrElse currentConfigControls.Count = 0 Then Return

            For i As Integer = 0 To currentConfigControls.Count - 1
                Dim ctrl As System.Windows.Forms.Control = currentConfigControls(i)
                If TypeOf ctrl Is System.Windows.Forms.Label Then
                    Dim labelText As String = CType(ctrl, System.Windows.Forms.Label).Text
                    Dim configVar As AppConfigurationVariable = targetConfig.FirstOrDefault(Function(x) x.DisplayName = labelText)
                    If configVar IsNot Nothing AndAlso i + 1 < currentConfigControls.Count Then
                        Dim inputControl As System.Windows.Forms.Control = currentConfigControls(i + 1)
                        If TypeOf inputControl Is System.Windows.Forms.TextBox Then
                            configVar.CurrentValue = CType(inputControl, System.Windows.Forms.TextBox).Text
                        End If
                    End If
                End If
            Next
        End Sub

        ''' <summary>
        ''' Dynamically generates Label+TextBox pairs for the active provider's configuration fields.
        ''' </summary>
        ''' <remarks>
        ''' Creates a label and a textbox per configuration variable, then optionally appends a note label.
        ''' The note label is informational only; it does not map to any <see cref="AppConfigurationVariable"/>.
        ''' </remarks>
        Private Sub LoadConfigForSelectedProvider()
            Dim selectedList As List(Of AppConfigurationVariable) = GetConfigListByName(_activeProvider)
            If selectedList Is Nothing Then Return

            ' Clear panel contents
            panelConfig.Controls.Clear()
            currentConfigControls.Clear()

            ' Update header label
            lblCurrentProvider.Text = "Configuration For " & _activeProvider & ":"

            ' Pass 1: Calculate maximum label width for alignment
            Dim yPos As Integer = 0
            Dim maxLabelWidth As Integer = 0
            For Each configVar In selectedList
                Dim lbl As New System.Windows.Forms.Label() With {
                .Text = configVar.DisplayName,
                .AutoSize = True,
                .Font = New System.Drawing.Font(Me.Font, FontStyle.Regular)
            }
                maxLabelWidth = Math.Max(maxLabelWidth, lbl.PreferredWidth)
            Next

            ' Pass 2: Create and position Label+TextBox pairs
            For Each configVar In selectedList
                ' Create label with DisplayName
                Dim lbl As New System.Windows.Forms.Label() With {
                .Text = configVar.DisplayName,
                .AutoSize = True,
                .Font = New System.Drawing.Font(Me.Font, FontStyle.Regular)
            }
                lbl.Location = New System.Drawing.Point(0, yPos)
                panelConfig.Controls.Add(lbl)
                currentConfigControls.Add(lbl)

                ' Create TextBox with CurrentValue
                Dim txt As New TextBox() With {
                .Width = panelConfig.Width - maxLabelWidth - 30,
                .Text = configVar.CurrentValue
            }
                txt.Location = New System.Drawing.Point(maxLabelWidth + 10, yPos - 2)
                panelConfig.Controls.Add(txt)
                currentConfigControls.Add(txt)

                yPos += lbl.Height + 8
            Next

            ' Append provider note (if exists) at end of field list
            Dim endNote As String = Nothing
            providerNotes.TryGetValue(_activeProvider, endNote)
            If Not String.IsNullOrWhiteSpace(endNote) Then
                Dim noteLabel As New System.Windows.Forms.Label() With {
                .AutoSize = True,
                .MaximumSize = New Size(panelConfig.Width - maxLabelWidth - 30, 0),
                .Text = endNote,
                .ForeColor = SystemColors.GrayText
            }
                noteLabel.Location = New System.Drawing.Point(maxLabelWidth + 10, yPos)
                panelConfig.Controls.Add(noteLabel)
                currentConfigControls.Add(noteLabel)
                yPos += noteLabel.Height + 8
            End If

            panelConfig.Height = yPos + 2
        End Sub

        ''' <summary>
        ''' Retrieves a provider configuration list by provider name.
        ''' </summary>
        ''' <param name="name">Provider name (e.g., "OpenAI", "Google Vertex").</param>
        ''' <returns>Provider variable list, or Nothing if provider is not found.</returns>
        Private Function GetConfigListByName(name As String) As List(Of AppConfigurationVariable)
            If String.IsNullOrEmpty(name) Then Return Nothing
            Dim list As List(Of AppConfigurationVariable) = Nothing
            If providerConfigs.TryGetValue(name, list) Then
                Return list
            End If
            Return Nothing
        End Function

        ''' <summary>
        ''' Returns the variable list for the currently active provider.
        ''' </summary>
        ''' <returns>Provider variable list, or Nothing if provider is not found.</returns>
        Private Function GetSelectedConfigList() As List(Of AppConfigurationVariable)
            Return GetConfigListByName(_activeProvider)
        End Function

        ''' <summary>
        ''' OK button click handler.
        ''' </summary>
        ''' <remarks>
        ''' Flow: SaveCurrentInputToConfig -> ValidateAllConfigs -> copy values to <see cref="ISharedContext"/> ->
        ''' write INI files via <see cref="CreateAppConfig"/> -> close dialog.
        ''' </remarks>
        Private Sub btnOK_Click(sender As Object, e As EventArgs)
            Try
                ' Save inputs from current panel to CurrentValue
                SaveCurrentInputToConfig()

                ' Validate all fields
                If Not ValidateAllConfigs() Then
                    Return
                End If

                ' If validation passed: Get selected provider's variable list
                Dim finalList = GetSelectedConfigList()
                If finalList Is Nothing Then
                    SharedMethods.ShowCustomMessageBox("No AI provider selected.")
                    Return
                End If

                ' Map VarName -> CurrentValue to _context properties
                For Each cv In finalList
                    Select Case cv.VarName
                        Case "INI_APIKey" : _context.INI_APIKey = cv.CurrentValue
                        Case "INI_Temperature" : _context.INI_Temperature = cv.CurrentValue
                        Case "INI_Timeout" : _context.INI_Timeout = CInt(cv.CurrentValue)
                        Case "INI_Model" : _context.INI_Model = cv.CurrentValue
                        Case "INI_Endpoint" : _context.INI_Endpoint = cv.CurrentValue
                        Case "INI_HeaderA" : _context.INI_HeaderA = cv.CurrentValue
                        Case "INI_HeaderB" : _context.INI_HeaderB = cv.CurrentValue
                        Case "INI_APICall" : _context.INI_APICall = cv.CurrentValue
                        Case "INI_APICall_Object" : _context.INI_APICall_Object = cv.CurrentValue
                        Case "INI_Response" : _context.INI_Response = cv.CurrentValue
                        Case "INI_OAuth2ClientMail" : _context.INI_OAuth2ClientMail = cv.CurrentValue
                        Case "INI_OAuth2Scopes" : _context.INI_OAuth2Scopes = cv.CurrentValue
                        Case "INI_OAuth2Endpoint" : _context.INI_OAuth2Endpoint = cv.CurrentValue
                        Case "INI_OAuth2ATExpiry" : _context.INI_OAuth2ATExpiry = CInt(cv.CurrentValue)
                    End Select
                Next

                ' Only Google Vertex requires OAuth2 by default
                If String.Equals(_activeProvider, "Google Vertex", StringComparison.OrdinalIgnoreCase) Then
                    _context.INI_OAuth2 = True
                End If

                _context.INIloaded = False

                Dim providerName As String = _activeProvider

                If chkWord.Checked Then CreateAppConfig("Word", providerName)
                If chkExcel.Checked Then CreateAppConfig("Excel", providerName)
                If chkOutlook.Checked Then CreateAppConfig("Outlook", providerName)

                ShowCustomMessageBox(
                            "You have completed a basic installation." & vbCrLf & vbCrLf &
                            "To add more model configurations, Special Service configurations, get sample files " &
                            "(including a prompt library) and change your settings, go to 'Settings', where you " &
                            "will find in particular a button with 'Get More'. Use it." & vbCrLf & vbCrLf &
                            $"Have fun using {SharedMethods.AN}!"
                        )

                ' Close wizard
                Me.DialogResult = DialogResult.OK
                _context.InitialConfigFailed = False
                Me.Close()

            Catch ex As System.Exception
                MessageBox.Show("Error finalizing configuration: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        ''' <summary>
        ''' Cancel button click handler.
        ''' </summary>
        ''' <remarks>
        ''' Closes the wizard without writing configuration files and sets <c>InitialConfigFailed=True</c>.
        ''' </remarks>
        Private Sub btnCancel_Click(sender As Object, e As EventArgs)
            Me.DialogResult = DialogResult.Cancel
            _context.InitialConfigFailed = True
            Me.Close()
        End Sub

        ''' <summary>
        ''' Validates all configuration fields for the active provider against validation rules.
        ''' </summary>
        ''' <returns>True if validation succeeded; otherwise False.</returns>
        ''' <remarks>
        ''' Validation rules are applied based on substring checks on <see cref="AppConfigurationVariable.ValidationRule"/>.
        ''' Accepts "." or "," as decimal separator for the max-value validation rule.
        '''
        ''' Obvious issue in current implementation:
        ''' The "0.0-2.0" rule returns True immediately after validating a single field and therefore skips the
        ''' validation of remaining fields.
        ''' </remarks>
        Private Function ValidateAllConfigs() As Boolean
            Dim selectedList = GetSelectedConfigList()

            ' Check if at least one relevant checkbox is checked for current host
            If _context.RDV.StartsWith("Word") AndAlso Not chkWord.Checked Then
                SharedMethods.ShowCustomMessageBox("At least the 'for Word' checkbox needs to be checked.")
                Return False
            ElseIf _context.RDV.StartsWith("Outlook") AndAlso Not chkOutlook.Checked Then
                SharedMethods.ShowCustomMessageBox("At least the 'for Outlook' checkbox needs to be checked.")
                Return False
            ElseIf _context.RDV.StartsWith("Excel") AndAlso Not chkExcel.Checked Then
                SharedMethods.ShowCustomMessageBox("At least the 'for Excel' checkbox needs to be checked.")
                Return False
            End If

            For Each cv In selectedList
                Dim valRule = cv.ValidationRule
                Dim valValue = cv.CurrentValue

                Debug.WriteLine("Validating: valrule=" & valRule & ", valValue='" & valValue & "'")

                ' NotEmpty validation
                If valRule.Contains("NotEmpty") Then
                    If String.IsNullOrWhiteSpace(valValue) Then
                        SharedMethods.ShowCustomMessageBox("Value For '" & cv.DisplayName & "' cannot be empty.")
                        Return False
                    End If
                End If

                ' E-Mail validation (simple @ check)
                If valRule.Contains("E-Mail") Then
                    If Not valValue.Contains("@") Then
                        SharedMethods.ShowCustomMessageBox("Value for '" & cv.DisplayName & "' must be a valid e-mail address.")
                        Return False
                    End If
                End If

                ' Hyperlink validation (http/https protocol check)
                If valRule.Contains("Hyperlink") Then
                    If Not (valValue.StartsWith("http://") OrElse valValue.StartsWith("https://")) Then
                        SharedMethods.ShowCustomMessageBox("Value for '" & cv.DisplayName & "' must be a valid URL (http/https).")
                        Return False
                    End If
                End If

                ' Positive integer validation (>0)
                If valRule.Contains(">0") Then
                    Dim intVal As Integer
                    If Not Integer.TryParse(valValue, intVal) OrElse intVal <= 0 Then
                        SharedMethods.ShowCustomMessageBox("Value for '" & cv.DisplayName & "' must be an integer larger than 0.")
                        Return False
                    End If
                End If

                ' Explicit range validation (0.0-2.0) [backwards compatibility with old field validation rule]
                If valRule.Contains("0.0-2.0") Then
                    Dim dblVal As Double
                    If Not Double.TryParse(valValue, dblVal) Then
                        SharedMethods.ShowCustomMessageBox("Value for '" & cv.DisplayName & "' must be a floating number between 0.0 and 2.0.")
                        Return False
                    End If
                    If dblVal < 0.0 OrElse dblVal > 2.0 Then
                        SharedMethods.ShowCustomMessageBox("Value for '" & cv.DisplayName & "' must be in [0.0 .. 2.0].", "Validation Error")
                        Return False
                    End If
                    Return True  ' Do not continue with further validation to avoid conflicting with next rule
                End If

                ' Max value validation (regex pattern \d+\.\d+, e.g., "2.0")
                If System.Text.RegularExpressions.Regex.IsMatch(valRule.Trim(), "^\d+\.\d+$") Then
                    Dim maxVal As Double
                    If Not Double.TryParse(valRule.Trim(),
                                               System.Globalization.NumberStyles.Float,
                                               System.Globalization.CultureInfo.InvariantCulture,
                                               maxVal) Then
                        SharedMethods.ShowCustomMessageBox("Internal validation error: cannot parse max value rule '" & valRule & "'.")
                        Return False
                    End If

                    Dim rawValue As String = valValue.Trim()

                    ' Normalize decimal separator: allow either "," or "."
                    ' Replace comma with dot; reject if more than one dot afterwards (invalid format like thousand separators)
                    Dim normalized As String = rawValue.Replace(",", ".")
                    If normalized.Count(Function(c) c = "."c) > 1 Then
                        SharedMethods.ShowCustomMessageBox("Value for '" & cv.DisplayName & "' is not a valid decimal number. Use one decimal point ('.' or ','). " &
                                                               "(Example: 1.25 or 1,25)")
                        Return False
                    End If

                    Dim dblVal As Double
                    If Not Double.TryParse(normalized,
                                               System.Globalization.NumberStyles.Float,
                                               System.Globalization.CultureInfo.InvariantCulture,
                                               dblVal) Then
                        SharedMethods.ShowCustomMessageBox("Value for '" & cv.DisplayName & "' must be a decimal number between 0 and " &
                                                               maxVal.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) &
                                                               " (accepts '.' or ',' as decimal separator).")
                        Return False
                    End If

                    If dblVal < 0.0 OrElse dblVal > maxVal Then
                        SharedMethods.ShowCustomMessageBox("Value for '" & cv.DisplayName & "' must be between 0 and " &
                                                               maxVal.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) &
                                                               ". Entered: " & rawValue)
                        Return False
                    End If
                End If

            Next

            Return True
        End Function

        ''' <summary>
        ''' Handles a <see cref="LinkLabel.LinkClicked"/> event by opening the link target with the default application.
        ''' </summary>
        Private Sub LinkLabel_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
            Try
                Dim link = e.Link.LinkData.ToString()
                System.Diagnostics.Process.Start(link)
            Catch ex As System.Exception
                MessageBox.Show("Could not open link. Error: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Writes a Red Ink configuration INI file for the specified Office application.
        ''' </summary>
        ''' <param name="App">Office application name ("Word", "Excel", or "Outlook").</param>
        ''' <param name="provider">Provider display name for the INI comment.</param>
        ''' <remarks>
        ''' The output path is obtained via <see cref="SharedMethods.GetDefaultINIPath"/>.
        ''' Temperature is normalized to a dot decimal separator using invariant culture formatting.
        ''' </remarks>
        Private Sub CreateAppConfig(App As String, provider As String)
            Try
                ' Define the file path
                Dim filepath = SharedMethods.GetDefaultINIPath(App)

                Debug.WriteLine($"Creating {SharedMethods.AN} configuration file: " & filepath)

                ' Open a StreamWriter to create the file
                Using writer As New System.IO.StreamWriter(filepath)
                    ' Write the header
                    writer.WriteLine($"; {SharedMethods.AN} configuration file (automatically generated)")
                    writer.WriteLine(";")
                    writer.WriteLine($"; Go to {SharedMethods.AN4} on how to find the instructions to manually add or change the configuration settings")

                    ' Write an empty line
                    writer.WriteLine()

                    ' Write provider information
                    writer.WriteLine($"; Minimum configuration for {provider}")

                    ' Write another empty line
                    writer.WriteLine()

                    ' Normalize Temperature to use dot as decimal separator
                    Dim normalizedTemp As String = _context.INI_Temperature
                    If Not String.IsNullOrWhiteSpace(normalizedTemp) Then
                        Dim tempValue As Double
                        ' Parse with current culture, then format with invariant culture (dot separator)
                        If Double.TryParse(normalizedTemp.Replace(","c, "."c),
                                       System.Globalization.NumberStyles.Float,
                                       System.Globalization.CultureInfo.InvariantCulture,
                                       tempValue) Then
                            normalizedTemp = tempValue.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)
                        End If
                    End If

                    ' Loop through the dictionary and write each configuration value
                    Dim MinimumConfigValues As New Dictionary(Of String, String) From {
                        {"APIKey", _context.INI_APIKey},
                        {"Endpoint", _context.INI_Endpoint},
                        {"HeaderA", _context.INI_HeaderA},
                        {"HeaderB", _context.INI_HeaderB},
                        {"Response", _context.INI_Response},
                        {"APICall", _context.INI_APICall},
                        {"APICall_Object", _context.INI_APICall_Object},
                        {"Timeout", _context.INI_Timeout.ToString()},
                        {"Temperature", normalizedTemp},
                        {"Model", _context.INI_Model},
                        {"OAuth2", _context.INI_OAuth2.ToString()},
                        {"OAuth2ClientMail", _context.INI_OAuth2ClientMail},
                        {"OAuth2Scopes", _context.INI_OAuth2Scopes},
                        {"OAuth2Endpoint", _context.INI_OAuth2Endpoint},
                        {"OAuth2ATExpiry", _context.INI_OAuth2ATExpiry.ToString()}
                    }

                    For Each kvp In MinimumConfigValues
                        writer.WriteLine($"{kvp.Key} = {kvp.Value}")
                    Next
                End Using

            Catch ex As System.Exception
                ' Handle errors by showing a custom message box
                SharedMethods.ShowCustomMessageBox($"Error creating configuration file: {ex.Message}")
            End Try
        End Sub
    End Class
End Namespace