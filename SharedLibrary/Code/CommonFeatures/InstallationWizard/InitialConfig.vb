' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: InitialConfig.vb
' Purpose: First-run installation wizard form for configuring Red Ink LLM API access.
'          Provides provider-specific templates (OpenAI, Azure, Google Enterprise Agent Platform, etc.),
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
' Provider defaults are loaded from the embedded redink-defaultconfig.ini in PrepareConfigData():
'   providerConfigs(ProviderName) -> List(Of AppConfigurationVariable)
' plus optional providerNotes(ProviderName) displayed below the fields.
'
' Remote Template Override (optional, user-confirmed):
' ---------------------------------------------------
' ProviderTemplateCatalog.OfferOnlineUpdate() is the shared opt-in workflow used by
' both configuration wizards. It downloads, validates, compares, and offers remote
' provider values without changing a saved user INI automatically.
'
' Data Flow:
' ----------
' 1. InitializeComponent() builds the static wizard UI.
' 2. PrepareConfigData() loads the shared provider catalog and offers the same opt-in update as ConfigWizard.
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
' OAuth2 (Google Enterprise Agent Platform):
' -----------------------
' Google Enterprise Agent Platform includes additional OAuth2 fields. INI_OAuth2ATExpiry is an integer value in
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

        Protected Overrides Sub OnShown(e As System.EventArgs)
            MyBase.OnShown(e)
            Me.TopMost = True
            Global.SharedLibrary.SharedLibrary.SharedMethods.ForceDialogToForeground(Me)
            Global.SharedLibrary.SharedLibrary.SharedMethods.AttachForeignForegroundWatchdog(Me)
        End Sub


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
        Private providerOrder As New System.Collections.Generic.List(Of System.String)()

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

            ' Populate from the same ordered INI catalog as the Configuration Wizard.
            ' Do not maintain a second hard-coded provider-name list here.
            For Each providerName As System.String In providerOrder
                cmbProvider.Items.Add(providerName)
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
            .Text = "for all/Word",
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
        ''' Provider field definitions and order come entirely from the embedded INI catalog.
        ''' Both wizards use the same opt-in online update workflow and field validation.
        ''' </remarks>
        Private Sub PrepareConfigData()
            Dim embedded As ProviderTemplateCatalogSnapshot = ProviderTemplateCatalog.LoadEmbedded()
            Dim selected As ProviderTemplateCatalogSnapshot
            ' Same downloader, parser and comparison logic as ConfigWizard.
            Using SharedMethods.PushDialogOwner(Me)
                selected = ProviderTemplateCatalog.OfferOnlineUpdate(embedded, "Installation Wizard")
            End Using
            providerConfigs = selected.Configurations
            providerNotes = selected.Notes
            providerOrder = selected.ProviderOrder.ToList()
        End Sub

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
        ''' <param name="name">Provider name (e.g., "OpenAI", "Google Enterprise Agent Platform").</param>
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

                _context.INI_APIKey = ""
                _context.INI_APIEncrypted = False
                _context.INI_APIKeyPrefix = ""
                _context.INI_Temperature = ""
                _context.INI_Timeout = 0
                _context.INI_Model = ""
                _context.INI_Endpoint = ""
                _context.INI_HeaderA = ""
                _context.INI_HeaderB = ""
                _context.INI_APICall = ""
                _context.INI_APICall_Object = ""
                _context.INI_Response = ""
                _context.INI_OAuth2 = False
                _context.INI_OAuth2ClientMail = ""
                _context.INI_OAuth2Scopes = ""
                _context.INI_OAuth2Endpoint = ""
                _context.INI_OAuth2ATExpiry = 0

                For Each cv In finalList
                    Select Case cv.VarName
                        Case "INI_APIKey" : _context.INI_APIKey = cv.CurrentValue
                        Case "INI_APIEncrypted"
                            Dim isEncrypted As Boolean
                            _context.INI_APIEncrypted = Boolean.TryParse(cv.CurrentValue, isEncrypted) AndAlso isEncrypted
                        Case "INI_APIKeyPrefix" : _context.INI_APIKeyPrefix = cv.CurrentValue
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

                ' Only Google Enterprise Agent Platform requires OAuth2 by default
                If String.Equals(_activeProvider, "Google Enterprise Agent Platform", StringComparison.OrdinalIgnoreCase) Then
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
                Global.SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("Error finalizing configuration: " & ex.Message, "Error")
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
                Global.SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("Could not open link. Error: " & ex.Message)
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
                            {"APIKeyEncrypted", _context.INI_APIEncrypted.ToString()},
                            {"APIKeyPrefix", _context.INI_APIKeyPrefix},
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
