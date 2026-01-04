' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SharedMethods.Settings.vb
' Purpose:
'   Provides Windows Forms UI and helper methods to view and modify runtime configuration values stored
'   on `ISharedContext` and to persist configuration changes to disk and/or `My.Settings`.
'
' Architecture / Responsibilities:
'   - Settings UI (curated subset):
'       `ShowSettingsWindow` builds a modal settings dialog dynamically from two dictionaries:
'         - `Settings`:     settingKey -> label template (may contain "{model}" and "{model2}")
'         - `SettingsTips`: settingKey -> tooltip text
'       The UI creates a control per key (TextBox or CheckBox) based on `IsBooleanSetting`.
'
'   - Expert configuration UI (arbitrary variables):
'       `ShowExpertConfiguration` materializes a variable name/value dictionary from the current `ISharedContext`,
'       shows it via `ShowVariableConfigurationWindow`, then maps edited values back into `ISharedContext`.
'       `ShowVariableConfigurationWindow` displays an editable two-column grid (variable/value) and can open
'       selected `.ini` files via `ShowTextFileEditor`.
'
'   - In-memory configuration access:
'       `GetSettingValue` and `SetSettingValue` provide string-based mapping between UI setting keys and
'       concrete properties on `ISharedContext`. `SetSettingValue` performs parsing for numeric/Boolean keys
'       and updates derived context flags (`INI_PromptLib`, `Ignore`).
'
'   - Model switching:
'       `SwitchModels` swaps primary and secondary ("_2") configuration values directly on `ISharedContext`.
'
'   - Persisting configuration:
'       `UpdateAppConfig` rewrites configuration content using a read/transform/write approach:
'         1) Resolves the active `.ini` input file path using:
'              - Registry setting (`RegPath_IniPath`) and precedence (`RegPath_IniPrio`)
'              - Per-application default path (`GetDefaultINIPath(context.RDV)`)
'              - Word default path fallback (`GetDefaultINIPath("Word")`)
'         2) Reads the resolved `IniFilePath` and updates known keys from an `expectedKeys` dictionary.
'         3) Skips selected keys when their values equal defaults (`KeysToSkipWhenDefault`).
'         4) Persists a small, explicit subset of keys to `My.Settings` instead of the `.ini`
'            (`SaveToMySettings` / `pendingMySettings`).
'         5) Writes the updated configuration to a temporary file and replaces the local default `.ini`
'            (`DefaultPath`) by moving the temporary file into place.
'
'       `ResetLocalAppConfig` rewrites the per-application default `.ini` file to contain only a defined
'       set of keys, preserving comment/empty lines where encountered.
'
'       `GetActiveConfigFilePath` exposes the resolved active configuration file path using the same path
'       precedence rules used by `UpdateAppConfig`.
'
' External dependencies (within `SharedMethods` / SharedLibrary):
'   - Configuration bootstrap and persistence helpers: `InitializeConfig`, `GetDefaultINIPath`, `RemoveCR`,
'     registry helpers (`GetFromRegistry`, `RegPath_Base`, `RegPath_IniPath`, `RegPath_IniPrio`).
'   - UI helpers: `ShowCustomMessageBox`, `ShowCustomYesNoBox`, `ShowSelectionForm`, `ShowTextFileEditor`,
'     and related forms/modules.
'   - Update workflow integration: `UpdateHandler` (invoked from the settings UI when applicable).
' =============================================================================


Option Strict On
Option Explicit On

Imports System.Deployment.Application
Imports System.Drawing
Imports System.IO
Imports System.Net.Http
Imports System.Reflection.Emit
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms
Imports Markdig.Extensions
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Tools
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary

    Partial Public Class SharedMethods

        ''' <summary>
        ''' Shows a modal settings dialog that allows temporarily editing a subset of configuration values.
        ''' Values are read from and written to <paramref name="context"/> via <see cref="GetSettingValue"/> and
        ''' <see cref="SetSettingValue"/>.
        ''' </summary>
        ''' <param name="Settings">
        ''' Map of setting key to label text template. Templates may contain "{model}" and "{model2}" placeholders.
        ''' </param>
        ''' <param name="SettingsTips">Map of setting key to tooltip text.</param>
        ''' <param name="context">Shared context containing the current in-memory configuration values.</param>
        Public Shared Sub ShowSettingsWindow(Settings As Dictionary(Of String, String), SettingsTips As Dictionary(Of String, String), ByRef context As ISharedContext)

            InitializeConfig(context, False, False)
            If context.INIloaded = False Then Return

            Dim settingsForm As New System.Windows.Forms.Form()
            settingsForm.Text = $"{AN} Settings"
            settingsForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            settingsForm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            settingsForm.MaximizeBox = False
            settingsForm.MinimizeBox = False
            settingsForm.ShowInTaskbar = False
            settingsForm.TopMost = True

            Dim bmp As New System.Drawing.Bitmap(My.Resources.Red_Ink_Logo)
            settingsForm.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
            settingsForm.Font = standardFont

            Dim descriptionLabel As New System.Windows.Forms.Label()
            descriptionLabel.Text = "You can temporarily change the following values (save to keep them):"
            descriptionLabel.AutoSize = True
            descriptionLabel.Location = New System.Drawing.Point(10, 20)
            settingsForm.Controls.Add(descriptionLabel)

            Dim labelControls As New Dictionary(Of String, System.Windows.Forms.Label)
            Dim settingControls As New Dictionary(Of String, System.Windows.Forms.Control)

            Dim maxLabelWidth As Integer = 0
            For Each setting In Settings
                Dim textSize As System.Drawing.Size
                If context.INI_SecondAPI Then
                    textSize = TextRenderer.MeasureText(setting.Value.Replace("{model}", context.INI_Model).Replace("{model2}", context.INI_Model_2) & ":", standardFont)
                Else
                    textSize = TextRenderer.MeasureText(setting.Value.Replace("{model}", context.INI_Model).Replace("{model2}", "2nd model (none)") & ":", standardFont)
                End If
                maxLabelWidth = Math.Max(maxLabelWidth, textSize.Width)
            Next

            ' --- sizes / layout core ---
            Dim controlXOffset As Integer = maxLabelWidth + 20

            ' (1) Widen input fields a bit more
            Dim defaultControlWidth As Integer = 400

            Dim lineSpacing As Integer = CInt(TextRenderer.MeasureText("Sample", standardFont).Height * 1.5)

            ' (2) Scrollable panel with extra width padding to prevent horizontal scrollbar
            Dim scrollPanel As New Panel() With {
                .AutoScroll = True,
                .Location = New System.Drawing.Point(10, descriptionLabel.Bottom + 20),
                .Width = controlXOffset + defaultControlWidth + 10 + SystemInformation.VerticalScrollBarWidth + 8
            }
            settingsForm.Controls.Add(scrollPanel)

            Dim yPos As Integer = 0

            For Each setting In Settings
                Dim label As New System.Windows.Forms.Label()
                Dim ToolTip As New System.Windows.Forms.ToolTip()
                If context.INI_SecondAPI Then
                    label.Text = setting.Value.Replace("{model}", context.INI_Model).Replace("{model2}", context.INI_Model_2) & ":"
                Else
                    label.Text = setting.Value.Replace("{model}", context.INI_Model).Replace("{model2}", "2nd model (none)") & ":"
                End If
                label.AutoSize = True
                label.Font = standardFont
                label.Location = New System.Drawing.Point(0, yPos)
                scrollPanel.Controls.Add(label)
                labelControls.Add(setting.Key, label)
                Dim ToolTipText As String = SettingsTips(setting.Key)
                ToolTip.SetToolTip(label, ToolTipText)

                If IsBooleanSetting(setting.Key) Then
                    Dim checkBox As New System.Windows.Forms.CheckBox()
                    checkBox.Checked = Boolean.Parse(GetSettingValue(setting.Key, context))
                    checkBox.Location = New System.Drawing.Point(controlXOffset, yPos - 2)
                    checkBox.Enabled = Not (setting.Key.Contains("_2") AndAlso Not context.INI_SecondAPI)
                    scrollPanel.Controls.Add(checkBox)
                    settingControls.Add(setting.Key, checkBox)
                    ToolTip.SetToolTip(checkBox, ToolTipText)
                Else
                    Dim textBox As New System.Windows.Forms.TextBox()
                    textBox.Text = GetSettingValue(setting.Key, context)
                    textBox.Size = New System.Drawing.Size(defaultControlWidth, 20)
                    textBox.Location = New System.Drawing.Point(controlXOffset, yPos)
                    textBox.Enabled = Not (setting.Key.Contains("_2") AndAlso Not context.INI_SecondAPI)
                    scrollPanel.Controls.Add(textBox)
                    settingControls.Add(setting.Key, textBox)
                    ToolTip.SetToolTip(textBox, ToolTipText)
                End If

                yPos += lineSpacing
            Next

            ' After populating:
            Dim contentHeight As Integer = yPos

            ' (3) Dynamic max height (up to 70% of working area minus space for buttons), fallback to previous if smaller)
            Dim workArea = Screen.FromPoint(Cursor.Position).WorkingArea
            Dim reservedBelow As Integer = 180   ' space for buttons + margins
            Dim dynamicCap As Integer = Math.Max(450, CInt(workArea.Height * 0.7) - reservedBelow) ' ensure at least a bit taller than old 400
            Dim maxPanelHeight As Integer = dynamicCap

            scrollPanel.Height = If(contentHeight > maxPanelHeight, maxPanelHeight, contentHeight)

            ' (4) Buttons below panel 

            ' Top of button row
            Dim topButtonYPos As Integer = scrollPanel.Bottom + 20
            ' Height of one button row (derived once, reused)
            Dim buttonRowHeight As Integer = TextRenderer.MeasureText("Sample", standardFont).Height + 10
            ' Dim buttonYPos As Integer = scrollPanel.Bottom + 20
            Dim buttonYPos As Integer = topButtonYPos + buttonRowHeight + 10

            Dim buttonSpacing As Integer = 10

            ' --- INI Importer Buttons -------------------------------------------------

            Dim getMoreStuffButton As New System.Windows.Forms.Button()
            getMoreStuffButton.Text = "Get More"
            Dim getMoreStuffSize As System.Drawing.Size = TextRenderer.MeasureText(getMoreStuffButton.Text, standardFont)
            getMoreStuffButton.Size = New System.Drawing.Size(getMoreStuffSize.Width + 20, getMoreStuffSize.Height + 10)
            getMoreStuffButton.Location = New System.Drawing.Point(10, topButtonYPos)
            settingsForm.Controls.Add(getMoreStuffButton)

            Dim getMoreStuffButtonToolTip As New System.Windows.Forms.ToolTip()
            getMoreStuffButtonToolTip.SetToolTip(
                    getMoreStuffButton,
                    $"Will open {GetMoreStuffURL} to show you additional AI models, Special Services and other settings you can load into your configuration."
                )

            Dim loadProviderSettingsButton As New System.Windows.Forms.Button()
            loadProviderSettingsButton.Text = "Get Model/Special Service"
            Dim loadProviderSettingsSize As System.Drawing.Size = TextRenderer.MeasureText(loadProviderSettingsButton.Text, standardFont)
            loadProviderSettingsButton.Size = New System.Drawing.Size(loadProviderSettingsSize.Width + 20, loadProviderSettingsSize.Height + 10)
            loadProviderSettingsButton.Location = New System.Drawing.Point(getMoreStuffButton.Right + buttonSpacing, topButtonYPos)
            settingsForm.Controls.Add(loadProviderSettingsButton)

            Dim loadProviderSettingsToolTip As New System.Windows.Forms.ToolTip()
            loadProviderSettingsToolTip.SetToolTip(
                    loadProviderSettingsButton,
                    $"Allows you to configure AI models and Special Services based on an URL (or file) you provide. See '{getMoreStuffButton.Text}' for URLs."
                )

            Dim loadOtherSettingsButton As New System.Windows.Forms.Button()
            loadOtherSettingsButton.Text = "Get Settings"
            Dim loadOtherSettingsSize As System.Drawing.Size = TextRenderer.MeasureText(loadOtherSettingsButton.Text, standardFont)
            loadOtherSettingsButton.Size = New System.Drawing.Size(loadOtherSettingsSize.Width + 20, loadOtherSettingsSize.Height + 10)
            loadOtherSettingsButton.Location = New System.Drawing.Point(loadProviderSettingsButton.Right + buttonSpacing, topButtonYPos)
            settingsForm.Controls.Add(loadOtherSettingsButton)

            Dim loadOtherSettingsToolTip As New System.Windows.Forms.ToolTip()
            loadOtherSettingsToolTip.SetToolTip(
                        loadOtherSettingsButton,
                        $"Allows you to add configuration settings for {AN} based on an URL (or file) you provide. See '{getMoreStuffButton.Text}' for URLs."
                    )

            Dim downloadSampleFilesButton As New System.Windows.Forms.Button()
            downloadSampleFilesButton.Text = "Get Sample Files"
            Dim downloadSampleFilesSize As System.Drawing.Size = TextRenderer.MeasureText(downloadSampleFilesButton.Text, standardFont)
            downloadSampleFilesButton.Size = New System.Drawing.Size(downloadSampleFilesSize.Width + 20, downloadSampleFilesSize.Height + 10)
            downloadSampleFilesButton.Location = New System.Drawing.Point(loadOtherSettingsButton.Right + buttonSpacing, topButtonYPos)
            settingsForm.Controls.Add(downloadSampleFilesButton)

            Dim downloadSampleFilesToolTip As New System.Windows.Forms.ToolTip()
            downloadSampleFilesToolTip.SetToolTip(
                            downloadSampleFilesButton,
                            $"Downloads from {AppsUrl} sample files you can use with {AN} and update your configuration, if necessary."
                        )


            Dim activeIniPath As String = GetActiveConfigFilePath(context)
            If Not IniImportManager.CanUseImportFeature(context, activeIniPath, "") Then
                getMoreStuffButton.Enabled = False
                loadProviderSettingsButton.Enabled = False
                loadOtherSettingsButton.Enabled = False
                downloadSampleFilesButton.Enabled = False
            End If

            ' ------------------------------------------------------------------------

            Dim switchButton As New System.Windows.Forms.Button()
            switchButton.Text = "Switch Model"
            Dim switchButtonSize As System.Drawing.Size = TextRenderer.MeasureText(switchButton.Text, standardFont)
            switchButton.Size = New System.Drawing.Size(switchButtonSize.Width + 20, switchButtonSize.Height + 10)
            switchButton.Location = New System.Drawing.Point(10, buttonYPos)
            switchButton.Enabled = context.INI_SecondAPI
            settingsForm.Controls.Add(switchButton)

            Dim SwitchButtonToolTip As New System.Windows.Forms.ToolTip()
            SwitchButtonToolTip.SetToolTip(switchButton, "Will accept the current settings and switch the primary model with the secondary model.")

            Dim expertConfigButton As New System.Windows.Forms.Button()
            expertConfigButton.Text = "Expert Config"
            Dim expertButtonSize As System.Drawing.Size = TextRenderer.MeasureText(expertConfigButton.Text, standardFont)
            expertConfigButton.Size = New System.Drawing.Size(expertButtonSize.Width + 20, expertButtonSize.Height + 10)
            expertConfigButton.Location = New System.Drawing.Point(switchButton.Right + buttonSpacing, buttonYPos)
            settingsForm.Controls.Add(expertConfigButton)

            Dim expertConfigButtonToolTip As New System.Windows.Forms.ToolTip()
            expertConfigButtonToolTip.SetToolTip(expertConfigButton, $"Will accept the current settings and in a separate window let you amend all configuration variables from '{AN2}.ini'.")

            Dim saveConfigButton As New System.Windows.Forms.Button()
            saveConfigButton.Text = "Save Config"
            Dim saveButtonSize As System.Drawing.Size = TextRenderer.MeasureText(saveConfigButton.Text, standardFont)
            saveConfigButton.Size = New System.Drawing.Size(saveButtonSize.Width + 20, saveButtonSize.Height + 10)
            saveConfigButton.Location = New System.Drawing.Point(expertConfigButton.Right + buttonSpacing, buttonYPos)
            settingsForm.Controls.Add(saveConfigButton)

            Dim saveConfigToolTip As New System.Windows.Forms.ToolTip()
            saveConfigToolTip.SetToolTip(saveConfigButton, $"Will save the current configuration to a local copy of '{AN2}.ini' (overwriting any existing such file).")

            Dim CentralConfigAvailable As Boolean = System.IO.File.Exists(System.IO.Path.Combine(ExpandEnvironmentVariables(GetFromRegistry(RegPath_Base, RegPath_IniPath, True)), $"{AN2}.ini"))
            Dim delLocalConfigButton As New System.Windows.Forms.Button()
            Dim LocalConfigAvailable As Boolean = System.IO.File.Exists(GetDefaultINIPath(context.RDV))
            If CentralConfigAvailable Then
                delLocalConfigButton.Text = "Give Up Local Config"
            Else
                delLocalConfigButton.Text = "Reset Optional Values"
            End If
            Dim delLocalButtonSize As System.Drawing.Size = TextRenderer.MeasureText(delLocalConfigButton.Text, standardFont)
            delLocalConfigButton.Size = New System.Drawing.Size(delLocalButtonSize.Width + 20, delLocalButtonSize.Height + 10)
            delLocalConfigButton.Location = New System.Drawing.Point(saveConfigButton.Right + buttonSpacing, buttonYPos)
            settingsForm.Controls.Add(delLocalConfigButton)
            If Not LocalConfigAvailable Then
                delLocalConfigButton.Enabled = False
            End If

            Dim delLocalConfigToolTip As New System.Windows.Forms.ToolTip()
            If CentralConfigAvailable Then
                If Left(context.RDV, 4) = "Word" Then
                    delLocalConfigToolTip.SetToolTip(delLocalConfigButton, $"This will deactivate the local configuration in '{AN2}.ini' (by renaming it to '.bak', overwriting any existing such file) and have the central configuration file applied going forward.")
                Else
                    delLocalConfigToolTip.SetToolTip(delLocalConfigButton, $"This will deactivate the local configuration in '{AN2}.ini' (by renaming it to '.bak', overwriting any existing such file), and have the configuration file of your 'Word' add-in (if available) and otherwise the central one applied going forward.")
                End If
            Else
                delLocalConfigToolTip.SetToolTip(delLocalConfigButton, $"This will reset all parameters that are not mandatory by removing them from your local configuration file '{AN2}.ini'. A copy will be saved beforehand to '.bak', overwriting any existing such file.")
            End If

            Dim okButton As New System.Windows.Forms.Button()
            okButton.Text = "OK"
            Dim okButtonSize As System.Drawing.Size = TextRenderer.MeasureText(okButton.Text, standardFont)
            okButton.Size = New System.Drawing.Size(okButtonSize.Width + 20, okButtonSize.Height + 10)
            okButton.Location = New System.Drawing.Point(10, buttonYPos + 50)
            settingsForm.Controls.Add(okButton)

            Dim cancelButton As New System.Windows.Forms.Button()
            cancelButton.Text = "Cancel"
            Dim cancelButtonSize As System.Drawing.Size = TextRenderer.MeasureText(cancelButton.Text, standardFont)
            cancelButton.Size = New System.Drawing.Size(cancelButtonSize.Width + 20, cancelButtonSize.Height + 10)
            cancelButton.Location = New System.Drawing.Point(okButton.Right + buttonSpacing, buttonYPos + 50)
            settingsForm.Controls.Add(cancelButton)

            Dim aboutButton As New System.Windows.Forms.Button()
            aboutButton.Text = $"About {AN}"
            Dim aboutButtonSize As System.Drawing.Size = TextRenderer.MeasureText(aboutButton.Text, standardFont)
            aboutButton.Size = New System.Drawing.Size(aboutButtonSize.Width + 20, aboutButtonSize.Height + 10)
            aboutButton.Location = New System.Drawing.Point(cancelButton.Right + buttonSpacing, cancelButton.Top)
            settingsForm.Controls.Add(aboutButton)

            Dim RightSide As Integer = aboutButton.Right

            Dim updateButton As New System.Windows.Forms.Button()
            updateButton.Text = "Check for Updates"
            If Not String.IsNullOrWhiteSpace(context.INI_UpdatePath) Then
                updateButton.Text = "Do Local Update"
            End If
            Dim updateButtonSize As System.Drawing.Size = TextRenderer.MeasureText(updateButton.Text, standardFont)
            updateButton.Size = New System.Drawing.Size(updateButtonSize.Width + 20, updateButtonSize.Height + 10)
            updateButton.Location = New System.Drawing.Point(aboutButton.Right + buttonSpacing, cancelButton.Top)
            If ApplicationDeployment.IsNetworkDeployed OrElse Not String.IsNullOrWhiteSpace(context.INI_UpdatePath) OrElse (context.INI_UpdateIniSilentMode = 0 AndAlso context.INI_UpdateIni) Then
                settingsForm.Controls.Add(updateButton)
                RightSide = updateButton.Right
            End If

            Dim FilePath As String = ""
            Dim IsExcel As Boolean = True
            If context.RDV.Contains("Word") Then
                FilePath = ExpandEnvironmentVariables(HelperPaths("Word"))
                IsExcel = False
            ElseIf context.RDV.Contains("Excel") Then
                FilePath = ExpandEnvironmentVariables(HelperPaths("Excel"))
            End If
            Debug.WriteLine("Filepath=" & FilePath)

            Dim helperButton As New System.Windows.Forms.Button()
            If Not String.IsNullOrEmpty(FilePath) Then
                If File.Exists(FilePath) Then
                    helperButton.Text = "Remove Helper"
                Else
                    helperButton.Text = "Install Helper"
                    If context.INI_NoHelperDownload Then
                        helperButton.Enabled = False
                    End If
                End If
                Dim HelperButtonSize As System.Drawing.Size = TextRenderer.MeasureText(helperButton.Text, standardFont)
                helperButton.Size = New System.Drawing.Size(HelperButtonSize.Width + 20, HelperButtonSize.Height + 10)
                helperButton.Location = New System.Drawing.Point(RightSide + buttonSpacing, cancelButton.Top)
                settingsForm.Controls.Add(helperButton)
            End If
            Dim CapturedContext As ISharedContext = context

            AddHandler getMoreStuffButton.Click, Sub(sender, e)
                                                     Try
                                                         System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo With {
                                                                .FileName = GetMoreStuffURL,
                                                                .UseShellExecute = True
                                                            })
                                                     Catch ex As System.Exception
                                                         ShowCustomMessageBox($"Unable to open browser. Try opening {GetMoreStuffURL} on your own.")
                                                     End Try
                                                 End Sub

            AddHandler loadProviderSettingsButton.Click, Sub(sender, e)
                                                             If IniImportManager.RunInteractiveImportProvidersOnly(CapturedContext, settingsForm) Then
                                                                 Dim answer = ShowCustomYesNoBox("Your main configuration settings have changed. You need to reload them for them to become active. Proceed?", "Yes, reload", "No, load later")
                                                                 If answer = 1 Then
                                                                     ' Mark config as not loaded so InitializeConfig will re-read from disk
                                                                     CapturedContext.INIloaded = False
                                                                     ' Reload configuration from disk into memory
                                                                     InitializeConfig(CapturedContext, False, True)
                                                                     ' Refresh the UI with the newly loaded values
                                                                     RefreshFormValues(settingControls, labelControls, CapturedContext, Settings)
                                                                     switchButton.Enabled = CapturedContext.INI_SecondAPI
                                                                     CapturedContext.MenusAdded = False
                                                                 End If
                                                             End If
                                                         End Sub

            AddHandler loadOtherSettingsButton.Click, Sub(sender, e)
                                                          If IniImportManager.RunInteractiveImportOtherParameters(CapturedContext, settingsForm) Then
                                                              Dim answer = ShowCustomYesNoBox("Your main configuration settings have changed. You need to reload them for them to become active. Proceed?", "Yes, reload", "No, load later")
                                                              If answer = 1 Then
                                                                  ' Mark config as not loaded so InitializeConfig will re-read from disk
                                                                  CapturedContext.INIloaded = False
                                                                  ' Reload configuration from disk into memory
                                                                  InitializeConfig(CapturedContext, False, True)
                                                                  ' Refresh the UI with the newly loaded values
                                                                  RefreshFormValues(settingControls, labelControls, CapturedContext, Settings)
                                                                  switchButton.Enabled = CapturedContext.INI_SecondAPI
                                                                  CapturedContext.MenusAdded = False
                                                              End If
                                                          End If
                                                      End Sub

            AddHandler downloadSampleFilesButton.Click, Sub(sender, e)

                                                            If IniImportManager.RunDownloadSampleFiles(CapturedContext, settingsForm) Then
                                                                Dim answer = ShowCustomYesNoBox("Your main configuration settings have changed. You need to reload them for them to become active. Proceed?", "Yes, reload", "No, load later")
                                                                If answer = 1 Then
                                                                    ' Mark config as not loaded so InitializeConfig will re-read from disk
                                                                    CapturedContext.INIloaded = False
                                                                    ' Reload configuration from disk into memory
                                                                    InitializeConfig(CapturedContext, False, True)
                                                                    ' Refresh the UI with the newly loaded values
                                                                    RefreshFormValues(settingControls, labelControls, CapturedContext, Settings)
                                                                    switchButton.Enabled = CapturedContext.INI_SecondAPI
                                                                    CapturedContext.MenusAdded = False
                                                                End If
                                                            End If
                                                        End Sub

            AddHandler switchButton.Click, Sub(sender, e)
                                               If CapturedContext.INI_SecondAPI Then
                                                   For Each settingKey In settingControls.Keys
                                                       Dim control = settingControls(settingKey)
                                                       If TypeOf control Is System.Windows.Forms.TextBox Then
                                                           Dim textValue As String = DirectCast(control, System.Windows.Forms.TextBox).Text
                                                           SetSettingValue(settingKey, textValue, CapturedContext)
                                                       ElseIf TypeOf control Is System.Windows.Forms.CheckBox Then
                                                           Dim boolValue As Boolean = DirectCast(control, System.Windows.Forms.CheckBox).Checked
                                                           SetSettingValue(settingKey, boolValue.ToString(), CapturedContext)
                                                       Else
                                                           MessageBox.Show($"Error in ShowSettingsWindow - unsupported control type for setting '{settingKey}' in ShowSettingsWindow (Switch).", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                       End If
                                                   Next
                                                   SwitchModels(CapturedContext)
                                                   RefreshFormValues(settingControls, labelControls, CapturedContext, Settings)
                                                   switchButton.Enabled = CapturedContext.INI_SecondAPI
                                               End If
                                               CapturedContext.MenusAdded = False
                                           End Sub

            AddHandler expertConfigButton.Click, Sub(sender, e)
                                                     For Each settingKey In settingControls.Keys
                                                         Dim control = settingControls(settingKey)
                                                         If TypeOf control Is System.Windows.Forms.TextBox Then
                                                             Dim textValue As String = DirectCast(control, System.Windows.Forms.TextBox).Text
                                                             SetSettingValue(settingKey, textValue, CapturedContext)
                                                         ElseIf TypeOf control Is System.Windows.Forms.CheckBox Then
                                                             Dim boolValue As Boolean = DirectCast(control, System.Windows.Forms.CheckBox).Checked
                                                             SetSettingValue(settingKey, boolValue.ToString(), CapturedContext)
                                                         Else
                                                             MessageBox.Show($"Error in ShowSettingsWindow - unsupported control type for setting '{settingKey}' in ShowSettingsWindow (ExpertConfig).", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                         End If
                                                     Next
                                                     ShowExpertConfiguration(CapturedContext, settingsForm)
                                                     RefreshFormValues(settingControls, labelControls, CapturedContext, Settings)
                                                     switchButton.Enabled = CapturedContext.INI_SecondAPI
                                                     CapturedContext.MenusAdded = False
                                                 End Sub

            AddHandler saveConfigButton.Click, Sub(sender, e)
                                                   For Each settingKey In settingControls.Keys
                                                       Dim control = settingControls(settingKey)
                                                       If TypeOf control Is System.Windows.Forms.TextBox Then
                                                           Dim textValue As String = DirectCast(control, System.Windows.Forms.TextBox).Text
                                                           SetSettingValue(settingKey, textValue, CapturedContext)
                                                       ElseIf TypeOf control Is System.Windows.Forms.CheckBox Then
                                                           Dim boolValue As Boolean = DirectCast(control, System.Windows.Forms.CheckBox).Checked
                                                           SetSettingValue(settingKey, boolValue.ToString(), CapturedContext)
                                                       Else
                                                           MessageBox.Show($"Error in ShowSettingsWindow - unsupported control type for setting '{settingKey}' in ShowSettingsWindow (Save).", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                       End If
                                                   Next
                                                   UpdateAppConfig(CapturedContext)
                                                   CapturedContext.MenusAdded = False
                                               End Sub

            AddHandler delLocalConfigButton.Click, Sub(sender, e)
                                                       If CentralConfigAvailable Then
                                                           If ShowCustomYesNoBox($"Do you really want to deactivate your local configuration file? The file '{AN2}.ini' will be renamed to '.bak' overwriting any existing such file", "Yes", "No") = 1 Then
                                                               If RenameFileToBak(GetDefaultINIPath(CapturedContext.RDV)) Then
                                                                   ShowCustomMessageBox("Local configuration deactivated. The central configuration will be applied going forward.", "OK")
                                                                   InitializeConfig(CapturedContext, False, True)
                                                               End If
                                                           End If
                                                       Else
                                                           If ShowCustomYesNoBox($"Do you really want to reset your local configuration file by removing non-mandatory entries? The current configuration file '{AN2}.ini' will beforehand be saved to a '.bak' file overwriting any existing such file.", "Yes", "No") = 1 Then
                                                               If RenameFileToBak(GetDefaultINIPath(CapturedContext.RDV)) Then
                                                                   ResetLocalAppConfig(CapturedContext)
                                                               End If
                                                           End If
                                                       End If
                                                       RefreshFormValues(settingControls, labelControls, CapturedContext, Settings)
                                                       switchButton.Enabled = CapturedContext.INI_SecondAPI
                                                       CapturedContext.MenusAdded = False
                                                   End Sub

            AddHandler helperButton.Click, Async Sub(sender, e)
                                               If helperButton.Text = "Remove Helper" Then
                                                   If ShowCustomYesNoBox($"Do you really want to remove the helper file '{FilePath}' from your system? It will be unloaded and deleted. You can re-install it later.", "Yes", "No") = 1 Then
                                                       If IsExcel Then UnloadExcelAddin(ExcelHelper) Else UnloadWordAddin(WordHelper)
                                                       Try
                                                           System.IO.File.Delete(FilePath)
                                                       Catch ex As System.Exception
                                                       End Try
                                                       If System.IO.File.Exists(FilePath) Then
                                                           ShowCustomMessageBox($"The helper file could not be deleted. Try to manually delete the file '{FilePath}' after having closed the application.")
                                                       Else
                                                           ShowCustomMessageBox("The helper file was successfully deleted.")
                                                           helperButton.Text = "Install Helper"
                                                           CapturedContext.MenusAdded = False
                                                           RemoveMenu = True
                                                       End If
                                                   End If
                                               Else
                                                   If ShowCustomYesNoBox($"Do you really want to download the helper file from {AppsUrl} and have it installed to '{FilePath}'? Next time you start the application, it will be automatically loaded.", "Yes", "No") = 1 Then
                                                       Dim DownloadUrl As String = ""
                                                       If IsExcel Then DownloadUrl = ExcelHelperUrl Else DownloadUrl = WordHelperUrl
                                                       Try
                                                           Using client As New HttpClient()
                                                               client.Timeout = TimeSpan.FromMinutes(10)
                                                               client.DefaultRequestHeaders.AcceptEncoding.Clear()
                                                               Using response As HttpResponseMessage = Await client.GetAsync(DownloadUrl, HttpCompletionOption.ResponseHeadersRead)
                                                                   response.EnsureSuccessStatusCode()
                                                                   Using fileStream As FileStream = New FileStream(FilePath, FileMode.Create, FileAccess.Write, FileShare.None)
                                                                       Using httpStream As Stream = Await response.Content.ReadAsStreamAsync()
                                                                           Dim buffer(8192) As Byte
                                                                           Dim bytesRead As Integer
                                                                           Do
                                                                               bytesRead = Await httpStream.ReadAsync(buffer, 0, buffer.Length)
                                                                               If bytesRead = 0 Then Exit Do
                                                                               Await fileStream.WriteAsync(buffer, 0, bytesRead)
                                                                           Loop
                                                                       End Using
                                                                   End Using
                                                               End Using
                                                           End Using
                                                           ShowCustomMessageBox($"Download to '{FilePath}' completed. You must restart the application for it to be loaded.")
                                                           helperButton.Text = "Remove Helper"
                                                       Catch ex As System.Exception
                                                           ShowCustomMessageBox($"Error when downloading from '{DownloadUrl}' to '{FilePath}'. You may have to download and install the helper file manually.")
                                                       End Try
                                                   End If
                                               End If
                                               RefreshFormValues(settingControls, labelControls, CapturedContext, Settings)
                                               switchButton.Enabled = CapturedContext.INI_SecondAPI
                                               CapturedContext.MenusAdded = False
                                           End Sub

            AddHandler aboutButton.Click, Sub(sender, e)
                                              ShowAboutWindow(settingsForm, CapturedContext)
                                          End Sub

            If ApplicationDeployment.IsNetworkDeployed OrElse Not String.IsNullOrWhiteSpace(CapturedContext.INI_UpdatePath) OrElse (context.INI_UpdateIniSilentMode = 0 AndAlso context.INI_UpdateIni) Then
                AddHandler updateButton.Click, Sub(sender, e)
                                                   Dim updater As New UpdateHandler()
                                                   updater.CheckAndInstallUpdates(CapturedContext.RDV, CapturedContext.INI_UpdatePath, CapturedContext)
                                               End Sub
            End If

            AddHandler okButton.Click, Sub(sender, e)

                                           Dim SaveToMySettings As New Dictionary(Of String, String) From {
                                                    {"DefaultPrefix", "DefaultPrefix"},
                                                    {"ReplaceText2Override", "ReplaceText2Override"},
                                                    {"MarkupMethodWordOverride", "MarkupMethodWordOverride"},
                                                    {"MarkupMethodOutlookOverride", "MarkupMethodOutlookOverride"}
                                                }

                                           For Each settingKey In settingControls.Keys
                                               Dim control = settingControls(settingKey)
                                               If TypeOf control Is System.Windows.Forms.TextBox Then
                                                   Dim textValue As String = DirectCast(control, System.Windows.Forms.TextBox).Text
                                                   SetSettingValue(settingKey, textValue, CapturedContext)
                                               ElseIf TypeOf control Is System.Windows.Forms.CheckBox Then
                                                   Dim boolValue As Boolean = DirectCast(control, System.Windows.Forms.CheckBox).Checked
                                                   SetSettingValue(settingKey, boolValue.ToString(), CapturedContext)
                                               Else
                                                   MessageBox.Show($"Error in ShowSettingsWindow - unsupported control type for setting '{settingKey}' in ShowSettingsWindow (OK).", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                               End If
                                           Next

                                           ' Persist the current values of keys mapped to My.Settings (do not alter file persistence logic)
                                           If SaveToMySettings.Count > 0 Then
                                               For Each kvp In SaveToMySettings
                                                   Dim iniKey = kvp.Key
                                                   Dim mySettingsKey = kvp.Value
                                                   Dim currentValue As String = ""

                                                   If settingControls.ContainsKey(iniKey) Then
                                                       Dim ctrl = settingControls(iniKey)
                                                       If TypeOf ctrl Is System.Windows.Forms.TextBox Then
                                                           currentValue = DirectCast(ctrl, System.Windows.Forms.TextBox).Text
                                                       ElseIf TypeOf ctrl Is System.Windows.Forms.CheckBox Then
                                                           currentValue = DirectCast(ctrl, System.Windows.Forms.CheckBox).Checked.ToString()
                                                       Else
                                                           currentValue = GetSettingValue(iniKey, CapturedContext)
                                                       End If
                                                   Else
                                                       currentValue = GetSettingValue(iniKey, CapturedContext)
                                                   End If

                                                   Try
                                                       My.Settings.Item(mySettingsKey) = currentValue
                                                   Catch
                                                       ' Ignore if the My.Settings entry does not exist
                                                   End Try
                                               Next
                                               Try
                                                   My.Settings.Save()
                                               Catch
                                                   ' Ignore save errors silently
                                               End Try
                                           End If

                                           CapturedContext.MenusAdded = False
                                           settingsForm.Close()
                                       End Sub

            AddHandler cancelButton.Click, Sub(sender, e)
                                               settingsForm.Close()
                                           End Sub


            ' (5) Recalculate final form size AFTER buttons are placed
            settingsForm.ClientSize = New System.Drawing.Size(
                scrollPanel.Left + scrollPanel.Width + 20,
                cancelButton.Bottom + 20
            )

            settingsForm.ShowDialog()
        End Sub

        ''' <summary>
        ''' Unloads an Excel COM add-in from the currently running Excel instance (if available).
        ''' If no running Excel instance exists, a new hidden instance is created and used.
        ''' </summary>
        ''' <param name="addinName">
        ''' Substring matched (case-insensitively) against <see cref="Excel.AddIn.FullName"/> to select the add-in.
        ''' </param>
        Public Shared Sub UnloadExcelAddin(addinName As String)
            Dim excelApp As Excel.Application = Nothing
            Dim foundAddin As Boolean = False
            Try
                ' Start or get running instance of Excel
                excelApp = TryCast(System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"), Excel.Application)
                If excelApp Is Nothing Then
                    excelApp = New Excel.Application()
                    excelApp.Visible = False
                End If

                For Each addin As Excel.AddIn In excelApp.AddIns2
                    Try
                        If addin.FullName.ToLower().Contains(addinName.ToLower()) Then
                            Debug.WriteLine("Unloading add-in: " & addin.FullName)
                            addin.Installed = False  ' Unload the add-in
                            foundAddin = True
                            Exit For
                        End If
                    Finally
                        ' Release each AddIn COM object
                        Marshal.ReleaseComObject(addin)
                    End Try
                Next

                If foundAddin Then
                    GC.Collect()
                    GC.WaitForPendingFinalizers()
                    Debug.WriteLine("Waiting for Excel to release file lock...")
                    Thread.Sleep(1000)
                End If

            Catch ex As Exception
                Debug.WriteLine("Error unloading Excel add-In: " & ex.Message)
            Finally
                If excelApp IsNot Nothing Then
                    Marshal.ReleaseComObject(excelApp)
                    excelApp = Nothing
                End If
            End Try
        End Sub

        ''' <summary>
        ''' Unloads a Word COM add-in from the currently running Word instance (if available).
        ''' </summary>
        ''' <param name="addInName">
        ''' Add-in name matched (case-insensitively) against <see cref="Microsoft.Office.Interop.Word.AddIn.Name"/>.
        ''' </param>
        Public Shared Sub UnloadWordAddin(addInName As String)
            Try
                ' Attempt to get the active (running) Word Application instance.
                Dim wordApp As Microsoft.Office.Interop.Word.Application = CType(Marshal.GetActiveObject("Word.Application"), Microsoft.Office.Interop.Word.Application)

                ' Iterate through all loaded AddIns in Word.
                For Each addIn As Microsoft.Office.Interop.Word.AddIn In wordApp.AddIns
                    ' Compare names in a case-insensitive manner (if desired).
                    Debug.WriteLine("Addin: " & addIn.Name)
                    If addIn.Name.Equals(addInName, StringComparison.OrdinalIgnoreCase) Then
                        ' Unload the add-in from the current Word session.
                        addIn.Installed = False
                        addIn.Delete()
                        Debug.WriteLine("Deleted!")
                        Exit For
                    End If
                Next

            Catch ex As System.Exception
                Debug.WriteLine("Error unloading Word add-in: " & ex.Message)
            End Try
        End Sub



        ''' <summary>
        ''' Refreshes the settings form UI by updating labels (including "{model}" / "{model2}" placeholders)
        ''' and reloading current values from <paramref name="context"/>.
        ''' </summary>
        ''' <param name="settingControls">Map of setting key to input control (<see cref="TextBox"/> or <see cref="CheckBox"/>).</param>
        ''' <param name="labelControls">Map of setting key to its label control.</param>
        ''' <param name="context">Shared context providing the in-memory configuration values.</param>
        ''' <param name="Settings">Map of setting key to label text template.</param>
        Public Shared Sub RefreshFormValues(settingControls As Dictionary(Of String, System.Windows.Forms.Control),
                              labelControls As Dictionary(Of String, System.Windows.Forms.Label), ByRef context As ISharedContext, Settings As Dictionary(Of String, String))
            ' Update the labels and input controls dynamically
            For Each setting In Settings
                ' Update label text
                If labelControls.ContainsKey(setting.Key) Then
                    If context.INI_SecondAPI Then
                        labelControls(setting.Key).Text = setting.Value.Replace("{model}", context.INI_Model).Replace("{model2}", context.INI_Model_2) & ":"
                    Else
                        labelControls(setting.Key).Text = setting.Value.Replace("{model}", context.INI_Model).Replace("{model2}", "of 2nd model (none)") & ":"
                    End If

                    ' Update input controls
                    If TypeOf settingControls(setting.Key) Is System.Windows.Forms.TextBox Then
                        settingControls(setting.Key).Text = GetSettingValue(setting.Key, context)
                    ElseIf TypeOf settingControls(setting.Key) Is System.Windows.Forms.CheckBox Then
                        DirectCast(settingControls(setting.Key), System.Windows.Forms.CheckBox).Checked = Boolean.Parse(GetSettingValue(setting.Key, context))
                    End If
                End If
            Next
        End Sub

        ''' <summary>
        ''' Determines whether a setting key is represented as a Boolean input in the settings UI.
        ''' </summary>
        ''' <param name="settingKey">Setting key to check.</param>
        ''' <returns><c>True</c> if the key is listed as Boolean; otherwise, <c>False</c>.</returns>
        Public Shared Function IsBooleanSetting(settingKey As String) As Boolean
            ' Determine if a setting is a Boolean based on its key
            Dim booleanSettings As New List(Of String) From {
        "DoubleS", "NoEmDash", "Clean", "MarkdownBubbles", "KeepFormat1", "MarkdownConvert", "ReplaceText1",
        "KeepFormat2", "KeepParaFormatInline", "ReplaceText2", "DoMarkupOutlook", "DoMarkupWord",
        "APIDebug", "ISearch_Approve", "ISearch", "Lib", "ContextMenu", "SecondAPI", "APIEncrypted", "APIEncrypted_2",
        "OAuth2", "OAuth2_2", "PromptLib", "Ignore", "ToolingLogWindow", "ToolingDryRun",
        "UpdateIni", "UpdateIniAllowRemote", "UpdateIniNoSignature", "UpdateIniSilentLog", "NoHelperDownload"
            }
            Return booleanSettings.Contains(settingKey)
        End Function


        ''' <summary>
        ''' Returns the current string representation of a named setting as stored on <paramref name="context"/>.
        ''' </summary>
        ''' <param name="settingName">Setting key to read.</param>
        ''' <param name="context">Shared context containing the in-memory configuration values.</param>
        ''' <returns>The setting value as a string, or an empty string if the key is not handled.</returns>
        Public Shared Function GetSettingValue(settingName As String, ByRef context As ISharedContext) As String
            ' Return the value of the setting based on its name
            Select Case settingName
                Case "APIKey"
                    Return context.INI_APIKeyBack
                Case "Temperature"
                    Return context.INI_Temperature
                Case "Timeout"
                    Return context.INI_Timeout.ToString() ' Convert Long to String
                Case "Model"
                    Return context.INI_Model
                Case "Endpoint"
                    Return context.INI_Endpoint
                Case "HeaderA"
                    Return context.INI_HeaderA
                Case "HeaderB"
                    Return context.INI_HeaderB
                Case "APICall"
                    Return context.INI_APICall
                Case "APICall_Object"
                    Return context.INI_APICall_Object
                Case "Response"
                    Return context.INI_Response
                Case "Anon"
                    Return context.INI_Anon
                Case "Anon_2"
                    Return context.INI_Anon_2
                Case "MaxOutputToken"
                    Return context.INI_MaxOutputToken.ToString()
                Case "MaxOutputToken_2"
                    Return context.INI_MaxOutputToken_2.ToString()
                Case "TokenCount"
                    Return context.INI_TokenCount
                Case "TokenCount_2"
                    Return context.INI_TokenCount_2
                Case "APIKey_2"
                    Return context.INI_APIKeyBack_2
                Case "Temperature_2"
                    Return context.INI_Temperature_2
                Case "Timeout_2"
                    Return context.INI_Timeout_2.ToString() ' Convert Long to String
                Case "Model_2"
                    Return context.INI_Model_2
                Case "Endpoint_2"
                    Return context.INI_Endpoint_2
                Case "HeaderA_2"
                    Return context.INI_HeaderA_2
                Case "HeaderB_2"
                    Return context.INI_HeaderB_2
                Case "APICall_2"
                    Return context.INI_APICall_2
                Case "APICall_Object_2"
                    Return context.INI_APICall_Object_2
                Case "Response_2"
                    Return context.INI_Response_2
                Case "OAuth2ClientMail"
                    Return context.INI_OAuth2ClientMail
                Case "OAuth2Scopes"
                    Return context.INI_OAuth2Scopes
                Case "OAuth2Endpoint"
                    Return context.INI_OAuth2Endpoint
                Case "OAuth2ATExpiry"
                    Return context.INI_OAuth2ATExpiry.ToString() ' Convert to String
                Case "OAuth2ClientMail_2"
                    Return context.INI_OAuth2ClientMail_2
                Case "OAuth2Scopes_2"
                    Return context.INI_OAuth2Scopes_2
                Case "OAuth2Endpoint_2"
                    Return context.INI_OAuth2Endpoint_2
                Case "OAuth2ATExpiry_2"
                    Return context.INI_OAuth2ATExpiry_2.ToString() ' Convert to String
                Case "APIKeyPrefix"
                    Return context.INI_APIKeyPrefix
                Case "APIKeyPrefix_2"
                    Return context.INI_APIKeyPrefix_2
                Case "Codebasis"
                    Return context.Codebasis
                Case "DoubleS"
                    Return context.INI_DoubleS.ToString()
                Case "Clean"
                    Return context.INI_Clean.ToString()
                Case "Ignore"
                    Return context.INI_Ignore.ToString()
                Case "Location"
                    Return context.INI_Location
                Case "NoEmDash"
                    Return context.INI_NoDash.ToString()
                Case "MarkdownBubbles"
                    Return context.INI_MarkdownBubbles.ToString()
                Case "DefaultPrefix"
                    Return context.INI_DefaultPrefix
                Case "KeepFormat1"
                    Return context.INI_KeepFormat1.ToString()
                Case "MarkdownConvert"
                    Return context.INI_MarkdownConvert.ToString()
                Case "ReplaceText1"
                    Return context.INI_ReplaceText1.ToString()
                Case "KeepFormat2"
                    Return context.INI_KeepFormat2.ToString()
                Case "KeepFormatCap"
                    Return context.INI_KeepFormatCap.ToString()
                Case "KeepParaFormatInline"
                    Return context.INI_KeepParaFormatInline.ToString()
                Case "ReplaceText2"
                    Return context.INI_ReplaceText2.ToString()
                Case "ReplaceText2Override"
                    Return context.INI_ReplaceText2Override
                Case "DoMarkupOutlook"
                    Return context.INI_DoMarkupOutlook.ToString()
                Case "DoMarkupWord"
                    Return context.INI_DoMarkupWord.ToString()
                Case "MarkupMethodHelper"
                    Return context.INI_MarkupMethodHelper.ToString()
                Case "MarkupMethodWord"
                    Return context.INI_MarkupMethodWord.ToString()
                Case "MarkupMethodWordOverride"
                    Return context.INI_MarkupMethodWordOverride
                Case "MarkupMethodOutlookOverride"
                    Return context.INI_MarkupMethodOutlookOverride
                Case "MarkupMethodOutlook"
                    Return context.INI_MarkupMethodOutlook.ToString()
                Case "MarkupDiffCap"
                    Return context.INI_MarkupDiffCap.ToString()
                Case "MarkupRegexCap"
                    Return context.INI_MarkupRegexCap.ToString()
                Case "ChatCap"
                    Return context.INI_ChatCap.ToString()
                Case "PreCorrection"
                    Return context.INI_PreCorrection.ToString()
                Case "PostCorrection"
                    Return context.INI_PostCorrection.ToString()
                Case "Language1"
                    Return context.INI_Language1.ToString()
                Case "Language2"
                    Return context.INI_Language2.ToString()
                Case "ShortcutsWordExcel"
                    Return context.INI_ShortcutsWordExcel
                Case "PromptLibPath"
                    Return context.INI_PromptLibPath
                Case "PromptLibPathLocal"
                    Return context.INI_PromptLibPathLocal
                Case "MyStylePath"
                    Return context.INI_MyStylePath
                Case "AlternateModelPath"
                    Return context.INI_AlternateModelPath
                Case "SpecialServicePath"
                    Return context.INI_SpecialServicePath
                Case "FindClausePath"
                    Return context.INI_FindClausePath
                Case "FindClausePathLocal"
                    Return context.INI_FindClausePathLocal
                Case "WebAgentPath"
                    Return context.INI_WebAgentPath
                Case "WebAgentPathLocal"
                    Return context.INI_WebAgentPathLocal
                Case "DocCheckPath"
                    Return context.INI_DocCheckPath
                Case "DocCheckPathLocal"
                    Return context.INI_DocCheckPathLocal
                Case "DocStylePath"
                    Return context.INI_DocStylePath
                Case "DocStylePathLocal"
                    Return context.INI_DocStylePathLocal
                Case "PromptLibPath_Transcript"
                    Return context.INI_PromptLibPath_Transcript
                Case "SpeechModelPath"
                    Return context.INI_SpeechModelPath
                Case "LocalModelPath"
                    Return context.INI_LocalModelPath
                Case "APIDebug"
                    Return context.INI_APIDebug.ToString()
                Case "ISearch"
                    Return context.INI_ISearch.ToString()
                Case "ISearch_Approve"
                    Return context.INI_ISearch_Approve.ToString()
                Case "ISearch_URL"
                    Return context.INI_ISearch_URL
                Case "ISearch_ResponseURLStart"
                    Return context.INI_ISearch_ResponseURLStart
                Case "ISearch_ResponseMask1"
                    Return context.INI_ISearch_ResponseMask1
                Case "ISearch_ResponseMask2"
                    Return context.INI_ISearch_ResponseMask2
                Case "ISearch_Name"
                    Return context.INI_ISearch_Name
                Case "ISearch_Tries"
                    Return context.INI_ISearch_Tries.ToString() ' Convert Integer to String
                Case "ISearch_Results"
                    Return context.INI_ISearch_Results.ToString() ' Convert Integer to String
                Case "ISearch_MaxDepth"
                    Return context.INI_ISearch_MaxDepth.ToString() ' Convert Integer to String
                Case "ISearch_Timeout"
                    Return context.INI_ISearch_Timeout.ToString() ' Convert Long to String
                Case "ISearch_SearchTerm_SP"
                    Return context.INI_ISearch_SearchTerm_SP
                Case "ISearch_Apply_SP"
                    Return context.INI_ISearch_Apply_SP
                Case "ISearch_Apply_SP_Markup"
                    Return context.INI_ISearch_Apply_SP_Markup
                Case "Lib"
                    Return context.INI_Lib.ToString()
                Case "Lib_File"
                    Return context.INI_Lib_File
                Case "Lib_Timeout"
                    Return context.INI_Lib_Timeout.ToString() ' Convert Long to String
                Case "Lib_Find_SP"
                    Return context.INI_Lib_Find_SP
                Case "Lib_Apply_SP"
                    Return context.INI_Lib_Apply_SP
                Case "Lib_Apply_SP_Markup"
                    Return context.INI_Lib_Apply_SP_Markup
                Case "SecondAPI"
                    Return context.INI_SecondAPI.ToString()
                Case "APIEncrypted"
                    Return context.INI_APIEncrypted.ToString()
                Case "APIEncrypted_2"
                    Return context.INI_APIEncrypted_2.ToString()
                Case "UsageRestrictions"
                    Return context.INI_UsageRestrictions
                Case "LogPath"
                    Return context.INI_LogPath
                Case "LogPath"
                    Return context.INI_LogPath
                Case "ContextMenu"
                    Return context.INI_ContextMenu.ToString()
                Case "UpdateCheckInterval"
                    Return context.INI_UpdateCheckInterval.ToString()
                Case "UpdatePath"
                    Return context.INI_UpdatePath
                Case "HelpMeInkyPath"
                    Return context.INI_HelpMeInkyPath
                Case "DiscussInkyPath"
                    Return context.INI_DiscussInkyPath
                Case "DiscussInkyPathLocal"
                    Return context.INI_DiscussInkyPathLocal
                Case "RedactionInstructionsPath"
                    Return context.INI_RedactionInstructionsPath
                Case "RedactionInstructionsPathLocal"
                    Return context.INI_RedactionInstructionsPathLocal
                Case "ExtractorPath"
                    Return context.INI_ExtractorPath
                Case "ExtractorPathLocal"
                    Return context.INI_ExtractorPathLocal
                Case "RenameLibPath"
                    Return context.INI_RenameLibPath
                Case "RenameLibPathLocal"
                    Return context.INI_RenameLibPathLocal
                Case "TTSEndpoint"
                    Return context.INI_TTSEndpoint
                Case "OAuth2"
                    Return context.INI_OAuth2.ToString()
                Case "OAuth2_2"
                    Return context.INI_OAuth2_2.ToString()
                Case "NoHelperDownload"
                    Return context.INI_NoHelperDownload.ToString()
                Case "ToolingLogWindow"
                    Return context.INI_ToolingLogWindow.ToString()
                Case "ToolingDryRun"
                    Return context.INI_ToolingDryRun.ToString()
                Case "ToolingMaximumIterations"
                    Return context.INI_ToolingMaximumIterations.ToString()
                Case "UpdateIni"
                    Return context.INI_UpdateIni.ToString()
                Case "UpdateIniAllowRemote"
                    Return context.INI_UpdateIniAllowRemote.ToString()
                Case "UpdateIniNoSignature"
                    Return context.INI_UpdateIniNoSignature.ToString()
                Case "UpdateSource"
                    Return context.INI_UpdateSource
                Case "UpdateClients"
                    Return context.INI_UpdateClients
                Case "UpdateIniIgnoreOverride"
                    Return context.INI_UpdateIniIgnoreOverride
                Case "UpdateIniSilentMode"
                    Return context.INI_UpdateIniSilentMode.ToString()
                Case "UpdateIniSilentLog"
                    Return context.INI_UpdateIniSilentLog.ToString()
                Case Else
                    Return ""
            End Select
        End Function


        ''' <summary>
        ''' Sets a named setting value on <paramref name="context"/> by mapping a string key to a specific context field.
        ''' Some settings are parsed into numeric or Boolean types based on the key.
        ''' After assignment, derived context flags are refreshed (<c>INI_PromptLib</c> and <c>Ignore</c>).
        ''' </summary>
        ''' <param name="settingName">Setting key to write.</param>
        ''' <param name="value">String representation of the value to assign (parsed for numeric/Boolean keys).</param>
        ''' <param name="context">Shared context that receives the updated in-memory configuration values.</param>
        Public Shared Sub SetSettingValue(settingName As String, value As String, ByRef context As ISharedContext)
            ' Set the value of the setting based on its name

            Select Case Trim(settingName)
                Case "APIKey"
                    context.INI_APIKeyBack = value
                Case "APIKeyPrefix"
                    context.INI_APIKeyPrefix = value
                Case "Temperature"
                    context.INI_Temperature = value
                Case "Timeout"
                    context.INI_Timeout = Long.Parse(value) ' Parse String to Long
                Case "Model"
                    context.INI_Model = value
                Case "Endpoint"
                    context.INI_Endpoint = value
                Case "HeaderA"
                    context.INI_HeaderA = value
                Case "HeaderB"
                    context.INI_HeaderB = value
                Case "APICall"
                    context.INI_APICall = value
                Case "APICall_Object"
                    context.INI_APICall_Object = value
                Case "Response"
                    context.INI_Response = value
                Case "Anon"
                    context.INI_Anon = value
                Case "TokenCount"
                    context.INI_TokenCount = value
                Case "APIKey_2"
                    context.INI_APIKeyBack_2 = value
                Case "APIKeyPrefix_2"
                    context.INI_APIKeyPrefix_2 = value
                Case "Temperature_2"
                    context.INI_Temperature_2 = value
                Case "Timeout_2"
                    context.INI_Timeout_2 = Long.Parse(value) ' Parse String to Long
                Case "Model_2"
                    context.INI_Model_2 = value
                Case "Endpoint_2"
                    context.INI_Endpoint_2 = value
                Case "HeaderA_2"
                    context.INI_HeaderA_2 = value
                Case "HeaderB_2"
                    context.INI_HeaderB_2 = value
                Case "APICall_2"
                    context.INI_APICall_2 = value
                Case "APICall_Object_2"
                    context.INI_APICall_Object_2 = value
                Case "Response_2"
                    context.INI_Response_2 = value
                Case "TokenCount_2"
                    context.INI_TokenCount_2 = value
                Case "OAuth2ClientMail"
                    context.INI_OAuth2ClientMail = value
                Case "OAuth2Scopes"
                    context.INI_OAuth2Scopes = value
                Case "OAuth2Endpoint"
                    context.INI_OAuth2Endpoint = value
                Case "OAuth2ATExpiry"
                    context.INI_OAuth2ATExpiry = Long.Parse(value) ' Parse String to Long
                Case "OAuth2ClientMail_2"
                    context.INI_OAuth2ClientMail_2 = value
                Case "OAuth2Scopes_2"
                    context.INI_OAuth2Scopes_2 = value
                Case "OAuth2Endpoint_2"
                    context.INI_OAuth2Endpoint_2 = value
                Case "OAuth2ATExpiry_2"
                    context.INI_OAuth2ATExpiry_2 = Long.Parse(value)
                Case "Codebasis"
                    context.Codebasis = value
                Case "DoubleS"
                    context.INI_DoubleS = Boolean.Parse(value)
                Case "Clean"
                    context.INI_Clean = Boolean.Parse(value)
                Case "Ignore"
                    context.INI_Ignore = Boolean.Parse(value)
                Case "Location"
                    context.INI_Location = value
                Case "NoEmDash"
                    context.INI_NoDash = Boolean.Parse(value)
                Case "MarkdownBubbles"
                    context.INI_MarkdownBubbles = Boolean.Parse(value)
                Case "DefaultPrefix"
                    context.INI_DefaultPrefix = value
                Case "KeepFormat1"
                    context.INI_KeepFormat1 = Boolean.Parse(value)
                Case "MarkdownConvert"
                    context.INI_MarkdownConvert = Boolean.Parse(value)
                Case "ReplaceText1"
                    context.INI_ReplaceText1 = Boolean.Parse(value)
                Case "KeepFormat2"
                    context.INI_KeepFormat2 = Boolean.Parse(value)
                Case "KeepFormatCap"
                    context.INI_KeepFormatCap = Integer.Parse(value)
                Case "KeepParaFormatInline"
                    context.INI_KeepParaFormatInline = Boolean.Parse(value)
                Case "ReplaceText2"
                    context.INI_ReplaceText2 = Boolean.Parse(value)
                Case "ReplaceText2Override"
                    context.INI_ReplaceText2Override = value
                Case "DoMarkupOutlook"
                    context.INI_DoMarkupOutlook = Boolean.Parse(value)
                Case "DoMarkupWord"
                    context.INI_DoMarkupWord = Boolean.Parse(value)
                Case "MarkupMethodHelper"
                    context.INI_MarkupMethodHelper = Integer.Parse(value)
                Case "MarkupMethodWord"
                    context.INI_MarkupMethodWord = Integer.Parse(value)
                Case "MarkupMethodWordOverride"
                    context.INI_MarkupMethodWordOverride = value
                Case "MarkupMethodOutlookOverride"
                    context.INI_MarkupMethodOutlookOverride = value
                Case "MarkupMethodOutlook"
                    context.INI_MarkupMethodOutlook = Integer.Parse(value)
                Case "MarkupDiffCap"
                    context.INI_MarkupDiffCap = Integer.Parse(value)
                Case "MarkupRegexCap"
                    context.INI_MarkupRegexCap = Integer.Parse(value)
                Case "ChatCap"
                    context.INI_ChatCap = Integer.Parse(value)
                Case "PreCorrection"
                    context.INI_PreCorrection = value
                Case "PostCorrection"
                    context.INI_PostCorrection = value
                Case "Language1"
                    context.INI_Language1 = value
                Case "Language2"
                    context.INI_Language2 = value
                Case "ShortcutsWordExcel"
                    context.INI_ShortcutsWordExcel = value
                Case "PromptLibPath"
                    context.INI_PromptLibPath = value
                Case "PromptLibPathLocal"
                    context.INI_PromptLibPathLocal = value
                Case "MyStylePath"
                    context.INI_MyStylePath = value
                Case "PromptLibPath_Transcript"
                    context.INI_PromptLibPath_Transcript = value
                Case "AlternateModelPath"
                    context.INI_AlternateModelPath = value
                Case "SpecialServicePath"
                    context.INI_SpecialServicePath = value
                Case "WebAgentPath"
                    context.INI_WebAgentPath = value
                Case "WebAgentPathLocal"
                    context.INI_WebAgentPathLocal = value
                Case "FindClausePath"
                    context.INI_FindClausePath = value
                Case "FindClausePathLocal"
                    context.INI_FindClausePathLocal = value
                Case "DocCheckPath"
                    context.INI_DocCheckPath = value
                Case "DocCheckPathLocal"
                    context.INI_DocCheckPathLocal = value
                Case "DocStylePath"
                    context.INI_DocStylePath = value
                Case "DocStylePathLocal"
                    context.INI_DocStylePathLocal = value
                Case "SpeechModelPath"
                    context.INI_SpeechModelPath = value
                Case "LocalModelPath"
                    context.INI_LocalModelPath = value
                Case "APIDebug"
                    context.INI_APIDebug = Boolean.Parse(value)
                Case "ISearch"
                    context.INI_ISearch = Boolean.Parse(value)
                Case "ISearch_Approve"
                    context.INI_ISearch_Approve = Boolean.Parse(value)
                Case "ISearch_URL"
                    context.INI_ISearch_URL = value
                Case "ISearch_ResponseURLStart"
                    context.INI_ISearch_ResponseURLStart = value
                Case "ISearch_ResponseMask1"
                    context.INI_ISearch_ResponseMask1 = value
                Case "ISearch_ResponseMask2"
                    context.INI_ISearch_ResponseMask2 = value
                Case "ISearch_Name"
                    context.INI_ISearch_Name = value
                Case "ISearch_Tries"
                    context.INI_ISearch_Tries = Integer.Parse(value) ' Parse String to Integer
                Case "ISearch_Results"
                    context.INI_ISearch_Results = Integer.Parse(value) ' Parse String to Integer
                Case "ISearch_MaxDepth"
                    context.INI_ISearch_MaxDepth = Integer.Parse(value) ' Parse String to Integer
                Case "ISearch_Timeout"
                    context.INI_ISearch_Timeout = Long.Parse(value) ' Parse String to Long
                Case "ISearch_SearchTerm_SP"
                    context.INI_ISearch_SearchTerm_SP = value
                Case "ISearch_Apply_SP"
                    context.INI_ISearch_Apply_SP = value
                Case "ISearch_Apply_SP_Markup"
                    context.INI_ISearch_Apply_SP_Markup = value
                Case "Lib"
                    context.INI_Lib = Boolean.Parse(value)
                Case "Lib_File"
                    context.INI_Lib_File = value
                Case "Lib_Timeout"
                    context.INI_Lib_Timeout = Long.Parse(value) ' Parse String to Long
                Case "Lib_Find_SP"
                    context.INI_Lib_Find_SP = value
                Case "Lib_Apply_SP"
                    context.INI_Lib_Apply_SP = value
                Case "Lib_Apply_SP_Markup"
                    context.INI_Lib_Apply_SP_Markup = value
                Case "Anon_2"
                    context.INI_Anon_2 = value
                Case "MaxOutputToken"
                    context.INI_MaxOutputToken = Integer.Parse(value)
                Case "MaxOutputToken_2"
                    context.INI_MaxOutputToken_2 = Integer.Parse(value)
                Case "ContextMenu"
                    context.INI_ContextMenu = Boolean.Parse(value)
                Case "UpdateCheckInterval"
                    context.INI_UpdateCheckInterval = Integer.Parse(value)
                Case "UpdatePath"
                    context.INI_UpdatePath = value
                Case "HelpMeInkyPath"
                    context.INI_HelpMeInkyPath = value
                Case "DiscussInkyPath"
                    context.INI_DiscussInkyPath = value
                Case "DiscussInkyPathLocal"
                    context.INI_DiscussInkyPathLocal = value
                Case "RedactionInstructionsPath"
                    context.INI_RedactionInstructionsPath = value
                Case "RedactionInstructionsPathLocal"
                    context.INI_RedactionInstructionsPathLocal = value
                Case "ExtractorPath"
                    context.INI_ExtractorPath = value
                Case "ExtractorPathLocal"
                    context.INI_ExtractorPathLocal = value
                Case "RenameLibPath"
                    context.INI_RenameLibPath = value
                Case "RenameLibPathLocal"
                    context.INI_RenameLibPathLocal = value
                Case "NoHelperDownload"
                    context.INI_NoHelperDownload = Boolean.Parse(value)
                Case "ToolingLogWindow"
                    context.INI_ToolingLogWindow = Boolean.Parse(value)
                Case "ToolingDryRun"
                    context.INI_ToolingDryRun = Boolean.Parse(value)
                Case "ToolingMaximumIterations"
                    context.INI_ToolingMaximumIterations = Integer.Parse(value)
                Case "UpdateIni"
                    context.INI_UpdateIni = Boolean.Parse(value)
                Case "UpdateIniAllowRemote"
                    context.INI_UpdateIniAllowRemote = Boolean.Parse(value)
                Case "UpdateIniNoSignature"
                    context.INI_UpdateIniNoSignature = Boolean.Parse(value)
                Case "UpdateSource"
                    context.INI_UpdateSource = value
                Case "UpdateClients"
                    context.INI_UpdateClients = value
                Case "UpdateIniIgnoreOverride"
                    context.INI_UpdateIniIgnoreOverride = value
                Case "UpdateIniSilentMode"
                    context.INI_UpdateIniSilentMode = Integer.Parse(value)
                Case "UpdateIniSilentLog"
                    context.INI_UpdateIniSilentLog = Boolean.Parse(value)

                Case Else
                    MessageBox.Show($"Error in SetSettingValue - could not save the value for '{settingName}'.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Select

            If context.INI_PromptLibPath.Trim() = "" And context.INI_PromptLibPathLocal.Trim() = "" Then context.INI_PromptLib = False Else context.INI_PromptLib = True
            If context.INI_Ignore Then context.Ignore = context.SP_Ignore Else context.Ignore = ""
            context.Location = context.INI_Location.Trim()

        End Sub


        ''' <summary>
        ''' Switches primary and secondary ("_2") configuration values on <paramref name="context"/> by swapping
        ''' each supported pair of fields.
        ''' </summary>
        ''' <param name="context">Shared context whose primary/secondary settings are swapped in-place.</param>
        Public Shared Sub SwitchModels(ByRef context As ISharedContext)
            ' Switch the content of variables with a _2 suffix with their corresponding variables without the _2 suffix
            Dim temp As String
            Dim tempb As Boolean
            Dim templ As Long
            Dim tempi As Integer
            Dim tempt As DateTime

            temp = context.INI_Model
            context.INI_Model = context.INI_Model_2
            context.INI_Model_2 = temp

            temp = context.INI_APIKeyPrefix
            context.INI_APIKeyPrefix = context.INI_APIKeyPrefix_2
            context.INI_APIKeyPrefix_2 = temp

            temp = context.INI_APIKey
            context.INI_APIKey = context.INI_APIKey_2
            context.INI_APIKey_2 = temp

            tempb = context.INI_APIEncrypted
            context.INI_APIEncrypted = context.INI_APIEncrypted_2
            context.INI_APIEncrypted_2 = tempb

            temp = context.INI_Temperature
            context.INI_Temperature = context.INI_Temperature_2
            context.INI_Temperature_2 = temp

            templ = context.INI_Timeout
            context.INI_Timeout = context.INI_Timeout_2
            context.INI_Timeout_2 = templ

            tempi = context.INI_MaxOutputToken
            context.INI_MaxOutputToken = context.INI_MaxOutputToken_2
            context.INI_MaxOutputToken_2 = tempi

            temp = context.INI_Endpoint
            context.INI_Endpoint = context.INI_Endpoint_2
            context.INI_Endpoint_2 = temp

            temp = context.INI_HeaderA
            context.INI_HeaderA = context.INI_HeaderA_2
            context.INI_HeaderA_2 = temp

            temp = context.INI_HeaderB
            context.INI_HeaderB = context.INI_HeaderB_2
            context.INI_HeaderB_2 = temp

            temp = context.INI_Response
            context.INI_Response = context.INI_Response_2
            context.INI_Response_2 = temp

            temp = context.INI_Anon
            context.INI_Anon = context.INI_Anon_2
            context.INI_Anon_2 = temp

            temp = context.INI_TokenCount
            context.INI_TokenCount = context.INI_TokenCount_2
            context.INI_TokenCount_2 = temp

            temp = context.INI_APICall
            context.INI_APICall = context.INI_APICall_2
            context.INI_APICall_2 = temp

            temp = context.INI_APICall_Object
            context.INI_APICall_Object = context.INI_APICall_Object_2
            context.INI_APICall_Object_2 = temp

            temp = context.INI_OAuth2ClientMail
            context.INI_OAuth2ClientMail = context.INI_OAuth2ClientMail_2
            context.INI_OAuth2ClientMail_2 = temp

            temp = context.INI_OAuth2Scopes
            context.INI_OAuth2Scopes = context.INI_OAuth2Scopes_2
            context.INI_OAuth2Scopes_2 = temp

            temp = context.INI_OAuth2Endpoint
            context.INI_OAuth2Endpoint = context.INI_OAuth2Endpoint_2
            context.INI_OAuth2Endpoint_2 = temp

            templ = context.INI_OAuth2ATExpiry
            context.INI_OAuth2ATExpiry = context.INI_OAuth2ATExpiry_2
            context.INI_OAuth2ATExpiry_2 = templ

            temp = context.DecodedAPI
            context.DecodedAPI = context.DecodedAPI_2
            context.DecodedAPI_2 = temp

            temp = context.INI_APIKeyBack
            context.INI_APIKeyBack = context.INI_APIKeyBack_2
            context.INI_APIKeyBack_2 = temp

            tempt = context.TokenExpiry
            context.TokenExpiry = context.TokenExpiry_2
            context.TokenExpiry_2 = tempt

            tempb = context.INI_OAuth2
            context.INI_OAuth2 = context.INI_OAuth2_2
            context.INI_OAuth2_2 = tempb



        End Sub

        ''' <summary>
        ''' Writes the current in-memory configuration from <paramref name="context"/> back to an `.ini` file and
        ''' persists selected keys into <c>My.Settings</c>.
        ''' The `.ini` content is rewritten by updating existing keys and appending missing keys.
        ''' </summary>
        ''' <param name="context">Shared context providing the in-memory configuration values to persist.</param>
        Public Shared Sub UpdateAppConfig(ByRef context As ISharedContext)
            Try

                Dim IniFilePath As String = ""
                Dim RegFilePath As String = ""
                Dim DefaultPath As String = ""
                Dim DefaultPath2 As String = ""
                Dim TempIniFilePath As String = ""

                ' Determine the configuration file path

                RegFilePath = GetFromRegistry(RegPath_Base, RegPath_IniPath, True)
                DefaultPath = GetDefaultINIPath(context.RDV)
                DefaultPath2 = GetDefaultINIPath("Word")

                If Not String.IsNullOrWhiteSpace(RegFilePath) AndAlso RegPath_IniPrio Then
                    IniFilePath = System.IO.Path.Combine(ExpandEnvironmentVariables(RegFilePath), $"{AN2}.ini")
                ElseIf System.IO.File.Exists(DefaultPath) Then
                    IniFilePath = DefaultPath
                ElseIf System.IO.File.Exists(DefaultPath2) Then
                    IniFilePath = DefaultPath2
                ElseIf Not String.IsNullOrWhiteSpace(RegFilePath) Then
                    IniFilePath = System.IO.Path.Combine(ExpandEnvironmentVariables(RegFilePath), $"{AN2}.ini")
                Else
                    IniFilePath = DefaultPath
                End If

                IniFilePath = RemoveCR(IniFilePath)

                ' Validate IniFilePath
                If Not System.IO.File.Exists(IniFilePath) Then
                    ShowCustomMessageBox($"The configuration file '{IniFilePath}' was not found.")
                    Return
                End If

                ' Create a temporary file for the updated configuration
                TempIniFilePath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(IniFilePath), $"{AN2}_temp.ini")

                ' Define all expected keys and their default or in-memory values
                Dim expectedKeys As New Dictionary(Of String, String) From {
                    {"APIKey", context.INI_APIKeyBack},
                    {"Endpoint", context.INI_Endpoint},
                    {"HeaderA", context.INI_HeaderA},
                    {"HeaderB", context.INI_HeaderB},
                    {"Response", context.INI_Response},
                    {"Anon", context.INI_Anon},
                    {"TokenCount", context.INI_TokenCount},
                    {"APICall", context.INI_APICall},
                    {"APICall_Object", context.INI_APICall_Object},
                    {"Timeout", context.INI_Timeout.ToString()},
                    {"MaxOutputToken", context.INI_MaxOutputToken.ToString()},
                    {"Temperature", context.INI_Temperature},
                    {"Model", context.INI_Model},
                    {"PreCorrection", context.INI_PreCorrection},
                    {"PostCorrection", context.INI_PostCorrection},
                    {"APIKeyPrefix", context.INI_APIKeyPrefix},
                    {"UsageRestrictions", context.INI_UsageRestrictions},
                    {"LogPath", context.INI_LogPath},
                    {"Language1", context.INI_Language1},
                    {"Language2", context.INI_Language2},
                    {"DoubleS", context.INI_DoubleS.ToString()},
                    {"Clean", context.INI_Clean.ToString()},
                    {"Ignore", context.INI_Ignore.ToString()},
                    {"Location", context.INI_Location},
                    {"NoEmDash", context.INI_NoDash.ToString()},
                    {"DefaultPrefix", context.INI_DefaultPrefix},
                    {"MarkdownBubbles", context.INI_MarkdownBubbles.ToString()},
                    {"KeepFormat1", context.INI_KeepFormat1.ToString()},
                    {"MarkdownConvert", context.INI_MarkdownConvert.ToString()},
                    {"ReplaceText1", context.INI_ReplaceText1.ToString()},
                    {"KeepFormat2", context.INI_KeepFormat2.ToString()},
                    {"KeepFormatCap", context.INI_KeepFormatCap.ToString()},
                    {"KeepParaFormatInline", context.INI_KeepParaFormatInline.ToString()},
                    {"ReplaceText2", context.INI_ReplaceText2.ToString()},
                    {"ReplaceText2Override", context.INI_ReplaceText2Override},
                    {"DoMarkupOutlook", context.INI_DoMarkupOutlook.ToString()},
                    {"DoMarkupWord", context.INI_DoMarkupWord.ToString()},
                    {"MarkupMethodOutlook", context.INI_MarkupMethodOutlook.ToString()},
                    {"MarkupDiffCap", context.INI_MarkupDiffCap.ToString()},
                    {"MarkupRegexCap", context.INI_MarkupRegexCap.ToString()},
                    {"ChatCap", context.INI_ChatCap.ToString()},
                    {"APIDebug", context.INI_APIDebug.ToString()},
                    {"APIKeyEncrypted", context.INI_APIEncrypted.ToString()},
                    {"SecondAPI", context.INI_SecondAPI.ToString()},
                    {"APIKey_2", context.INI_APIKeyBack_2},
                    {"Endpoint_2", context.INI_Endpoint_2},
                    {"HeaderA_2", context.INI_HeaderA_2},
                    {"HeaderB_2", context.INI_HeaderB_2},
                    {"Response_2", context.INI_Response_2},
                    {"Anon_2", context.INI_Anon_2},
                    {"TokenCount_2", context.INI_TokenCount_2},
                    {"APICall_2", context.INI_APICall_2},
                    {"APICall_Object_2", context.INI_APICall_Object_2},
                    {"Timeout_2", context.INI_Timeout_2.ToString()},
                    {"MaxOutputToken_2", context.INI_MaxOutputToken_2.ToString()},
                    {"Temperature_2", context.INI_Temperature_2},
                    {"Model_2", context.INI_Model_2},
                    {"APIKeyEncrypted_2", context.INI_APIEncrypted_2.ToString()},
                    {"APIKeyPrefix_2", context.INI_APIKeyPrefix_2},
                    {"OAuth2", context.INI_OAuth2.ToString()},
                    {"OAuth2ClientMail", context.INI_OAuth2ClientMail},
                    {"OAuth2Scopes", context.INI_OAuth2Scopes},
                    {"OAuth2Endpoint", context.INI_OAuth2Endpoint},
                    {"OAuth2ATExpiry", context.INI_OAuth2ATExpiry.ToString()},
                    {"OAuth2_2", context.INI_OAuth2_2.ToString()},
                    {"OAuth2ClientMail_2", context.INI_OAuth2ClientMail_2},
                    {"OAuth2Scopes_2", context.INI_OAuth2Scopes_2},
                    {"OAuth2Endpoint_2", context.INI_OAuth2Endpoint_2},
                    {"OAuth2ATExpiry_2", context.INI_OAuth2ATExpiry_2.ToString()},
                    {"ISearch", context.INI_ISearch.ToString()},
                    {"ISearch_Approve", context.INI_ISearch_Approve.ToString()},
                    {"ISearch_URL", context.INI_ISearch_URL},
                    {"ISearch_ResponseMask1", context.INI_ISearch_ResponseMask1},
                    {"ISearch_ResponseMask2", context.INI_ISearch_ResponseMask2},
                    {"ISearch_Name", context.INI_ISearch_Name},
                    {"ISearch_Tries", context.INI_ISearch_Tries.ToString()},
                    {"ISearch_Results", context.INI_ISearch_Results.ToString()},
                    {"ISearch_MaxDepth", context.INI_ISearch_MaxDepth.ToString()},
                    {"ISearch_Timeout", context.INI_ISearch_Timeout.ToString()},
                    {"ISearch_SearchTerm_SP", context.INI_ISearch_SearchTerm_SP},
                    {"ISearch_Apply_SP", context.INI_ISearch_Apply_SP},
                    {"ISearch_Apply_SP_Markup", context.INI_ISearch_Apply_SP_Markup},
                    {"Lib", context.INI_Lib.ToString()},
                    {"Lib_File", context.INI_Lib_File},
                    {"Lib_Timeout", context.INI_Lib_Timeout.ToString()},
                    {"Lib_Find_SP", context.INI_Lib_Find_SP},
                    {"Lib_Apply_SP", context.INI_Lib_Apply_SP},
                    {"Lib_Apply_SP_Markup", context.INI_Lib_Apply_SP_Markup},
                    {"MarkupMethodHelper", context.INI_MarkupMethodHelper.ToString()},
                    {"MarkupMethodWord", context.INI_MarkupMethodWord.ToString()},
                    {"MarkupMethodWordOverride", context.INI_MarkupMethodWordOverride},
                    {"MarkupMethodOutlookOverride", context.INI_MarkupMethodOutlookOverride},
                    {"ShortcutsWordExcel", context.INI_ShortcutsWordExcel},
                    {"ContextMenu", context.INI_ContextMenu.ToString()},
                    {"UpdateCheckInterval", context.INI_UpdateCheckInterval.ToString()},
                    {"UpdatePath", context.INI_UpdatePath},
                    {"HelpMeInkyPath", context.INI_HelpMeInkyPath},
                    {"DiscussInkyPath", context.INI_DiscussInkyPath},
                    {"DiscussInkyPathLocal", context.INI_DiscussInkyPathLocal},
                    {"RedactionInstructionsPath", context.INI_RedactionInstructionsPath},
                    {"RedactionInstructionsPathLocal", context.INI_RedactionInstructionsPathLocal},
                    {"ExtractorPath", context.INI_ExtractorPath},
                    {"ExtractorPathLocal", context.INI_ExtractorPathLocal},
                    {"RenameLibPath", context.INI_RenameLibPath},
                    {"RenameLibPathLocal", context.INI_RenameLibPathLocal},
                    {"SpeechModelPath", context.INI_SpeechModelPath},
                    {"LocalModelPath", context.INI_LocalModelPath},
                    {"TTSEndpoint", context.INI_TTSEndpoint},
                    {"PromptLib", context.INI_PromptLibPath},
                    {"PromptLibLocal", context.INI_PromptLibPathLocal},
                    {"MyStylePath", context.INI_MyStylePath},
                    {"AlternateModelPath", context.INI_AlternateModelPath},
                    {"SpecialServicePath", context.INI_SpecialServicePath},
                    {"FindClausePath", context.INI_FindClausePath},
                    {"FindClausePathLocal", context.INI_FindClausePathLocal},
                    {"WebAgentPath", context.INI_WebAgentPath},
                    {"WebAgentPathLocal", context.INI_WebAgentPathLocal},
                    {"DocCheckPath", context.INI_DocCheckPath},
                    {"DocCheckPathLocal", context.INI_DocCheckPathLocal},
                    {"DocStylePath", context.INI_DocStylePath},
                    {"DocStylePathLocal", context.INI_DocStylePathLocal},
                    {"PromptLib_Transcript", context.INI_PromptLibPath_Transcript},
                    {"SP_Translate", context.SP_Translate},
                    {"SP_Correct", context.SP_Correct},
                    {"SP_Improve", context.SP_Improve},
                    {"SP_Explain", context.SP_Explain},
                    {"SP_FindClause", context.SP_FindClause},
                    {"SP_FindClause_Clean", context.SP_FindClause_Clean},
                    {"SP_ApplyDocStyle", context.SP_ApplyDocStyle},
                    {"SP_ApplyDocStyle_NumberingHint", context.SP_ApplyDocStyle_NumberingHint},
                    {"SP_DocCheck_Clause", context.SP_DocCheck_Clause},
                    {"SP_DocCheck_MultiClause", context.SP_DocCheck_MultiClause},
                    {"SP_DocCheck_MultiClauseSum", context.SP_DocCheck_MultiClauseSum},
                    {"SP_DocCheck_MultiClauseSum_Bubbles", context.SP_DocCheck_MultiClauseSum_Bubbles},
                    {"SP_SuggestTitles", context.SP_SuggestTitles},
                    {"SP_Friendly", context.SP_Friendly},
                    {"SP_Convincing", context.SP_Convincing},
                    {"SP_NoFillers", context.SP_NoFillers},
                    {"SP_Podcast", context.SP_Podcast},
                    {"SP_MyStyle_Word", context.SP_MyStyle_Word},
                    {"SP_MyStyle_Outlook", context.SP_MyStyle_Outlook},
                    {"SP_MyStyle_Apply", context.SP_MyStyle_Apply},
                    {"SP_Shorten", context.SP_Shorten},
                    {"SP_Filibuster", context.SP_Filibuster},
                    {"SP_ArgueAgainst", context.SP_ArgueAgainst},
                    {"SP_InsertClipboard", context.SP_InsertClipboard},
                    {"SP_Summarize", context.SP_Summarize},
                    {"SP_Markup", context.SP_Markup},
                    {"SP_MailReply", context.SP_MailReply},
                    {"SP_MailSumup", context.SP_MailSumup},
                    {"SP_MailSumup2", context.SP_MailSumup2},
                    {"SP_FreestyleText", context.SP_FreestyleText},
                    {"SP_FreestyleNoText", context.SP_FreestyleNoText},
                    {"SP_SwitchParty", context.SP_SwitchParty},
                    {"SP_Anonymize", context.SP_Anonymize},
                    {"SP_Rename", context.SP_Rename},
                    {"SP_Redact", context.SP_Redact},
                    {"SP_CheckforII", context.SP_CheckforII},
                    {"SP_Extract", context.SP_Extract},
                    {"SP_ExtractSchema", context.SP_ExtractSchema},
                    {"SP_MergeDateRows", context.SP_MergeDateRows},
                    {"SP_ContextSearch", context.SP_ContextSearch},
                    {"SP_ContextSearchMulti", context.SP_ContextSearchMulti},
                    {"SP_RangeOfCells", context.SP_RangeOfCells},
                    {"SP_ParseFile", context.SP_ParseFile},
                    {"SP_Ignore", context.SP_Ignore},
                    {"SP_WriteNeatly", context.SP_WriteNeatly},
                    {"SP_Add_KeepFormulasIntact", context.SP_Add_KeepFormulasIntact},
                    {"SP_Add_KeepHTMLIntact", context.SP_Add_KeepHTMLIntact},
                    {"SP_Add_KeepInlineIntact", context.SP_Add_KeepInlineIntact},
                    {"SP_Add_Bubbles", context.SP_Add_Bubbles},
                    {"SP_Add_BubblesReply", context.SP_Add_BubblesReply},
                    {"SP_Add_BubblesExtract", context.SP_Add_BubblesExtract},
                    {"SP_Add_Bubbles_Format", context.SP_Add_Bubbles_Format},
                    {"SP_Add_Batch", context.SP_Add_Batch},
                    {"SP_Add_Tooling", context.SP_Add_Tooling},
                    {"SP_Add_Slides", context.SP_Add_Slides},
                    {"SP_BubblesExcel", context.SP_BubblesExcel},
                    {"SP_Add_Revisions", context.SP_Add_Revisions},
                    {"SP_MarkupRegex", context.SP_MarkupRegex},
                    {"SP_ChatWord", context.SP_ChatWord},
                    {"SP_HelpMe", context.SP_HelpMe},
                    {"SP_Chat", context.SP_Chat},
                    {"SP_Add_ChatWord_Commands", context.SP_Add_ChatWord_Commands},
                    {"SP_Add_Chat_NoCommands", context.SP_Add_Chat_NoCommands},
                    {"SP_ChatExcel", context.SP_ChatExcel},
                    {"SP_Add_ChatExcel_Commands", context.SP_Add_ChatExcel_Commands},
                    {"SP_FindPrompts", context.SP_FindPrompts},
                    {"SP_MergePrompt", context.SP_MergePrompt},
                    {"SP_MergePrompt2", context.SP_MergePrompt2},
                    {"SP_Add_MergePrompt", context.SP_Add_MergePrompt},
                    {"NoHelperDownload", context.INI_NoHelperDownload.ToString()},
                    {"ToolingLogWindow", context.INI_ToolingLogWindow.ToString()},
                    {"ToolingDryRun", context.INI_ToolingDryRun.ToString()},
                    {"ToolingMaximumIterations", context.INI_ToolingMaximumIterations.ToString()},
                    {"UpdateIni", context.INI_UpdateIni.ToString()},
                    {"UpdateIniAllowRemote", context.INI_UpdateIniAllowRemote.ToString()},
                    {"UpdateIniNoSignature", context.INI_UpdateIniNoSignature.ToString()},
                    {"UpdateSource", context.INI_UpdateSource},
                    {"UpdateClients", context.INI_UpdateClients},
                    {"UpdateIniIgnoreOverride", context.INI_UpdateIniIgnoreOverride},
                    {"UpdateIniSilentMode", context.INI_UpdateIniSilentMode.ToString()},
                    {"UpdateIniSilentLog", context.INI_UpdateIniSilentLog.ToString()}
                }

                Dim KeysToSkipWhenDefault As New Dictionary(Of String, Object) From {
                    {"ISearch_SearchTerm_SP", Default_INI_ISearch_SearchTerm_SP},
                    {"ISearch_Apply_SP", Default_INI_ISearch_Apply_SP},
                    {"ISearch_Apply_SP_Markup", Default_INI_ISearch_Apply_SP_Markup},
                    {"SP_Translate", Default_SP_Translate},
                    {"SP_Correct", Default_SP_Correct},
                    {"SP_Improve", Default_SP_Improve},
                    {"SP_Explain", Default_SP_Explain},
                    {"SP_FindClause", Default_SP_FindClause},
                    {"SP_FindClause_Clean", Default_SP_FindClause_Clean},
                    {"SP_ApplyDocStyle", Default_SP_ApplyDocStyle},
                    {"SP_ApplyDocStyle_NumberingHint", Default_SP_ApplyDocStyle_NumberingHint},
                    {"SP_DocCheck_Clause", Default_SP_DocCheck_Clause},
                    {"SP_DocCheck_MultiClause", Default_SP_DocCheck_MultiClause},
                    {"SP_DocCheck_MultiClauseSum", Default_SP_DocCheck_MulticlauseSum},
                    {"SP_DocCheck_MultiClauseSum_Bubbles", Default_SP_DocCheck_MultiClauseSum_Bubbles},
                    {"SP_SuggestTitles", Default_SP_SuggestTitles},
                    {"SP_Friendly", Default_SP_Friendly},
                    {"SP_Convincing", Default_SP_Convincing},
                    {"SP_NoFillers", Default_SP_NoFillers},
                    {"SP_Podcast", Default_SP_Podcast},
                    {"SP_MyStyle_Word", Default_SP_MyStyle_Word},
                    {"SP_MyStyle_Outlook", Default_SP_MyStyle_Outlook},
                    {"SP_MyStyle_Apply", Default_SP_MyStyle_Apply},
                    {"SP_Shorten", Default_SP_Shorten},
                    {"SP_Filibuster", Default_SP_Filibuster},
                    {"SP_ArgueAgainst", Default_SP_ArgueAgainst},
                    {"SP_InsertClipboard", Default_SP_InsertClipboard},
                    {"SP_Summarize", Default_SP_Summarize},
                    {"SP_Markup", Default_SP_Markup},
                    {"SP_MailReply", Default_SP_MailReply},
                    {"SP_MailSumup", Default_SP_MailSumup},
                    {"SP_MailSumup2", Default_SP_MailSumup2},
                    {"SP_FreestyleText", Default_SP_FreestyleText},
                    {"SP_FreestyleNoText", Default_SP_FreestyleNoText},
                    {"SP_SwitchParty", Default_SP_SwitchParty},
                    {"SP_Anonymize", Default_SP_Anonymize},
                    {"SP_Rename", Default_SP_Rename},
                    {"SP_Redact", Default_SP_Redact},
                    {"SP_CheckforII", Default_SP_CheckforII},
                    {"SP_Extract", Default_SP_Extract},
                    {"SP_ExtractSchema", Default_SP_ExtractSchema},
                    {"SP_MergeDateRows", Default_SP_MergeDateRows},
                    {"SP_ContextSearch", Default_SP_ContextSearch},
                    {"SP_ContextSearchMulti", Default_SP_ContextSearchMulti},
                    {"SP_RangeOfCells", Default_SP_RangeOfCells},
                    {"SP_ParseFile", Default_SP_ParseFile},
                    {"SP_Ignore", Default_SP_Ignore},
                    {"SP_WriteNeatly", Default_SP_WriteNeatly},
                    {"SP_Add_KeepFormulasIntact", Default_SP_Add_KeepFormulasIntact},
                    {"SP_Add_KeepHTMLIntact", Default_SP_Add_KeepHTMLIntact},
                    {"SP_Add_KeepInlineIntact", Default_SP_Add_KeepInlineIntact},
                    {"SP_Add_Bubbles", Default_SP_Add_Bubbles},
                    {"SP_Add_BubblesReply", Default_SP_Add_BubblesReply},
                    {"SP_Add_BubblesExtract", Default_SP_Add_BubblesExtract},
                    {"SP_Add_Bubbles_Format", Default_SP_Add_Bubbles_Format},
                    {"SP_Add_Batch", Default_SP_Add_Batch},
                    {"SP_Add_Tooling", Default_SP_Add_Tooling},
                    {"SP_Add_Slides", Default_SP_Add_Slides},
                    {"SP_BubblesExcel", Default_SP_BubblesExcel},
                    {"SP_Add_Revisions", Default_SP_Add_Revisions},
                    {"SP_MarkupRegex", Default_SP_MarkupRegex},
                    {"SP_ChatWord", Default_SP_ChatWord},
                    {"SP_HelpMe", Default_SP_HelpMe},
                    {"SP_Chat", Default_SP_Chat},
                    {"SP_Add_ChatWord_Commands", Default_SP_Add_ChatWord_Commands},
                    {"SP_Add_Chat_NoCommands", Default_SP_Add_Chat_NoCommands},
                    {"SP_ChatExcel", Default_SP_ChatExcel},
                    {"SP_Add_ChatExcel_Commands", Default_SP_Add_ChatExcel_Commands},
                    {"SP_FindPrompts", Default_SP_FindPrompts},
                    {"SP_Add_MergePrompt", Default_SP_Add_MergePrompt},
                    {"SP_MergePrompt", Default_SP_MergePrompt},
                    {"SP_MergePrompt2", Default_SP_MergePrompt2},
                    {"Temperature", DEFAULT_TEMPERATURE},
                    {"Timeout", DEFAULT_TIMEOUT_LONG},
                    {"Temperature_2", DEFAULT_TEMPERATURE},
                    {"Timeout_2", DEFAULT_TIMEOUT_LONG},
                    {"Language1", DEFAULT_LANGUAGE_1},
                    {"Language2", DEFAULT_LANGUAGE_2},
                    {"KeepFormatCap", DEFAULT_KEEPFORMAT_CAP},
                    {"MarkupMethodHelper", DEFAULT_MARKUP_METHOD_HELPER},
                    {"MarkupMethodWord", DEFAULT_MARKUP_METHOD_WORD},
                    {"MarkupMethodOutlook", DEFAULT_MARKUP_METHOD_OUTLOOK},
                    {"MarkupDiffCap", DEFAULT_MARKUP_DIFF_CAP},
                    {"MarkupRegexCap", DEFAULT_MARKUP_REGEX_CAP},
                    {"ChatCap", DEFAULT_CHAT_CAP},
                    {"Lib_Timeout", DEFAULT_TIMEOUT_LIB},
                    {"UpdateIniSilentMode", DEFAULT_UPDATE_INI_SILENT_MODE},
                    {"ReplaceText1", DEFAULT_BOOL_REPLACETEXT1},
                    {"MarkdownConvert", DEFAULT_BOOL_MARKDOWNCONVERT},
                    {"ReplaceText2", DEFAULT_BOOL_REPLACETEXT2},
                    {"DoMarkupOutlook", DEFAULT_BOOL_DOMARKUPOUTLOOK},
                    {"DoMarkupWord", DEFAULT_BOOL_DOMARKUPWORD},
                    {"ContextMenu", DEFAULT_BOOL_CONTEXTMENU},
                    {"ISearch", DEFAULT_BOOL_ISEARCH_ENABLED},
                    {"ToolingLogWindow", DEFAULT_BOOL_TOOLINGLOGWINDOW},
                    {"ToolingMaximumIterations", DEFAULT_TOOLING_MAXIMUMITERATIONS},
                    {"UpdateIni", DEFAULT_BOOL_UPDATEINI},
                    {"UpdateIniAllowRemote", DEFAULT_BOOL_UPDATEINI_ALLOWREMOTE},
                    {"UpdateIniSilentLog", DEFAULT_BOOL_UPDATEINISILENTLOG},
                    {"ISearch_URL", DEFAULT_ISEARCH_URL},
                    {"ISearch_ResponseMask1", DEFAULT_ISEARCH_RESPONSE_MASK_1},
                    {"ISearch_ResponseMask2", DEFAULT_ISEARCH_RESPONSE_MASK_2},
                    {"ISearch_Name", DEFAULT_ISEARCH_NAME},
                    {"ISearch_Tries", ISearch_DefTries},
                    {"ISearch_Results", ISearch_DefResults},
                    {"ISearch_MaxDepth", ISearch_DefMaxDepth},
                    {"ISearch_Timeout", ISearch_DefSearchTimeout},
                    {"UpdateCheckInterval", DefaultUpdateIntervalDays}
                }

                Dim SaveToMySettings As New Dictionary(Of String, String) From {
                    {"DefaultPrefix", "DefaultPrefix"},
                    {"ReplaceText2Override", "ReplaceText2Override"},
                    {"MarkupMethodWordOverride", "MarkupMethodWordOverride"},
                    {"MarkupMethodOutlookOverride", "MarkupMethodOutlookOverride"}
                }

                ' Accumulate settings to persist to My.Settings at the end
                Dim pendingMySettings As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

                ' Read the original ini file content
                Dim originalContent As String = System.IO.File.ReadAllText(IniFilePath)
                Dim updatedContent As New StringBuilder()
                Dim foundKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                ' Split into lines and process each line
                Dim iniLines As String() = originalContent.Split({vbCrLf}, StringSplitOptions.None)
                For Each line As String In iniLines
                    Dim trimmedLine As String = line.Trim()

                    ' Preserve comments and empty lines
                    If String.IsNullOrEmpty(trimmedLine) OrElse trimmedLine.StartsWith(";") Then
                        updatedContent.AppendLine(line)
                        Continue For
                    End If

                    ' Process key-value pairs
                    Dim keyValue As String() = trimmedLine.Split(New Char() {"="c}, 2)
                    If keyValue.Length = 2 Then
                        Dim key As String = keyValue(0).Trim()
                        Dim value As String = keyValue(1).Trim()

                        ' Update values for known keys
                        If expectedKeys.ContainsKey(key) Then
                            value = expectedKeys(key)
                        End If

                        foundKeys.Add(key)

                        ' If this key should be stored in My.Settings, queue it and do NOT write to the file
                        If SaveToMySettings.ContainsKey(key) Then
                            Dim settingsKey As String = SaveToMySettings(key)
                            pendingMySettings(settingsKey) = value
                            ' Skip writing this key to the file
                            Continue For
                        End If

                        ' Write the updated key-value pair
                        updatedContent.AppendLine($"{key} = {value}")
                    Else
                        ' Preserve lines that are not key-value pairs
                        updatedContent.AppendLine(line)
                    End If
                Next

                ' Add missing keys to the updated content
                For Each key In expectedKeys.Keys.Except(foundKeys, StringComparer.OrdinalIgnoreCase)
                    Dim value As String = expectedKeys(key)

                    ' If this key is mapped to My.Settings, store there only (respecting default-skip behavior)
                    If SaveToMySettings.ContainsKey(key) Then
                        If IsDefaultValue(key, value, KeysToSkipWhenDefault) Then
                            ' Skip adding default to settings to mirror previous "skip When Default" behavior
                        Else
                            Dim settingsKey As String = SaveToMySettings(key)
                            pendingMySettings(settingsKey) = value
                        End If
                        ' Never write mapped keys to the file
                        Continue For
                    End If


                    ' For normal keys: skip adding to file if default matches the skip rule
                    If IsDefaultValue(key, value, KeysToSkipWhenDefault) Then
                        Continue For
                    End If

                    ' Write the key-value pair to the updated content, but not for empty keys                    
                    If ShouldWriteKey(key, value, KeysToSkipWhenDefault) Then
                        updatedContent.AppendLine($"{key} = {value}")
                    End If
                Next

                ' Write the updated content to the temporary ini file
                System.IO.File.WriteAllText(TempIniFilePath, updatedContent.ToString())

                ' Replace the original file with the updated file
                ' Ensure the target directory exists
                Dim targetDir As String = System.IO.Path.GetDirectoryName(DefaultPath)
                If Not System.IO.Directory.Exists(targetDir) Then
                    System.IO.Directory.CreateDirectory(targetDir)
                End If

                ' Delete existing file only if it exists
                If System.IO.File.Exists(DefaultPath) Then
                    System.IO.File.Delete(DefaultPath)
                End If
                System.IO.File.Move(TempIniFilePath, DefaultPath)

                ' Persist any keys mapped to My.Settings at the end
                If pendingMySettings.Count > 0 Then
                    For Each kvp In pendingMySettings
                        ' Use late-bound access to avoid requiring strongly-typed settings properties
                        My.Settings.Item(kvp.Key) = kvp.Value
                    Next
                    My.Settings.Save()
                End If

                context.INIloaded = False

                If IniFilePath = DefaultPath Then
                    ShowCustomMessageBox("Your configuration file has been updated.")
                Else
                    ShowCustomMessageBox("Your configuration has been saved To a local configuration file (which will be used going forward until deleted).")
                End If

                InitializeConfig(context, False, True)

            Catch ex As System.Exception
                ShowCustomMessageBox($"Error updating configuration file: {ex.Message}")
            End Try
        End Sub


        ''' <summary>
        ''' Determines whether the specified INI key/value pair represents its defined default value
        ''' and therefore should be skipped when writing the configuration file.
        ''' </summary>
        ''' <param name="key">
        ''' The INI key name to evaluate.
        ''' </param>
        ''' <param name="currentValue">
        ''' The current string value that would be written to the INI file.
        ''' </param>
        ''' <param name="defaults">
        ''' A dictionary mapping INI keys to their strongly-typed default values.
        ''' </param>
        ''' <returns>
        ''' <c>True</c> if the current value matches the default value for the specified key;
        ''' otherwise, <c>False</c>.
        ''' </returns>
        Private Shared Function IsDefaultValue(
    ByVal key As String,
    ByVal currentValue As String,
    ByVal defaults As Dictionary(Of String, Object)
) As Boolean

            If Not defaults.ContainsKey(key) Then
                Return False
            End If

            Dim defaultValue As Object = defaults(key)

            Dim normalizedCurrent As String = NormalizeIniValue(currentValue)
            Dim normalizedDefault As String = NormalizeIniValue(defaultValue)

            Return String.Equals(normalizedCurrent, normalizedDefault, StringComparison.OrdinalIgnoreCase)
        End Function


        ''' <summary>
        ''' Normalizes an INI value or default value into a canonical string representation
        ''' suitable for reliable, culture-invariant comparison.
        ''' </summary>
        ''' <param name="value">
        ''' The value to normalize. Supported types include <see cref="String"/>,
        ''' <see cref="Boolean"/>, and numeric primitives.
        ''' </param>
        ''' <returns>
        ''' A normalized string representation of the value.
        ''' </returns>
        Private Shared Function NormalizeIniValue(ByVal value As Object) As String
            If value Is Nothing Then
                Return String.Empty
            End If

            If TypeOf value Is Boolean Then
                Return CBool(value).ToString().ToLowerInvariant()
            End If

            If TypeOf value Is Integer OrElse
       TypeOf value Is Long OrElse
       TypeOf value Is Double OrElse
       TypeOf value Is Decimal Then
                Return System.Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture)
            End If

            Return value.ToString().Trim()
        End Function



        ''' <summary>
        ''' Determines whether an INI key should be written to disk based on its value.
        ''' </summary>
        ''' <param name="key">
        ''' The INI key name being evaluated.
        ''' </param>
        ''' <param name="value">
        ''' The value associated with the INI key. Empty or whitespace-only strings are considered invalid.
        ''' Boolean "False" values for keys not in the defaults skiplist are also skipped.
        ''' </param>
        ''' <param name="defaults">
        ''' Dictionary of keys with explicit default values. Boolean keys in this dictionary default to True;
        ''' boolean keys NOT in this dictionary default to False and should be skipped when False.
        ''' </param>
        ''' <returns>
        ''' <c>True</c> if the value should be written; otherwise, <c>False</c>.
        ''' </returns>
        Private Shared Function ShouldWriteKey(ByVal key As String, ByVal value As String, ByVal defaults As Dictionary(Of String, Object)) As Boolean
            ' Skip empty or whitespace-only values
            If String.IsNullOrWhiteSpace(value) Then
                Return False
            End If

            ' For boolean "False" values not in the skiplist, skip writing (they default to False implicitly)
            If String.Equals(value, "False", StringComparison.OrdinalIgnoreCase) Then
                If Not defaults.ContainsKey(key) Then
                    Return False
                End If
            End If

            Return True
        End Function



        ''' <summary>
        ''' Resets the local `.ini` configuration file (at the default INI path for the current RDV) by rewriting it to
        ''' contain only keys listed in <c>expectedKeys</c>, using current values from <paramref name="context"/>.
        ''' Comment and empty lines are preserved from the original file where encountered.
        ''' </summary>
        ''' <param name="context">Shared context providing the in-memory configuration values to keep.</param>
        Public Shared Sub ResetLocalAppConfig(ByRef context As ISharedContext)
            Try
                ' Determine the path to the existing .ini file
                Dim IniFilePath As String = System.IO.Path.Combine(GetDefaultINIPath(context.RDV))
                Dim TempIniFilePath As String = ""

                ' Validate IniFilePath
                If Not System.IO.File.Exists(IniFilePath) Then
                    ShowCustomMessageBox($"The configuration file '{IniFilePath}' was not found.")
                    Return
                End If

                ' Create a temporary file for the updated configuration
                TempIniFilePath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(IniFilePath), $"{AN2}_temp.ini")

                ' Define all expected keys and their default or in-memory values --> they will remain in case of a reset of the configuration
                Dim expectedKeys As New Dictionary(Of String, String) From {
                    {"APIKey", context.INI_APIKeyBack},
                    {"Endpoint", context.INI_Endpoint},
                    {"HeaderA", context.INI_HeaderA},
                    {"HeaderB", context.INI_HeaderB},
                    {"Response", context.INI_Response},
                    {"Anon", context.INI_Anon},
                    {"TokenCount", context.INI_TokenCount},
                    {"APICall", context.INI_APICall},
                    {"APICall_Object", context.INI_APICall_Object},
                    {"Timeout", context.INI_Timeout.ToString()},
                    {"MaxOutputToken", context.INI_MaxOutputToken.ToString()},
                    {"Temperature", context.INI_Temperature},
                    {"Model", context.INI_Model},
                    {"APIKeyPrefix", context.INI_APIKeyPrefix},
                    {"APIKeyEncrypted", context.INI_APIEncrypted.ToString()},
                    {"SecondAPI", context.INI_SecondAPI.ToString()},
                    {"APIKey_2", context.INI_APIKeyBack_2},
                    {"Endpoint_2", context.INI_Endpoint_2},
                    {"HeaderA_2", context.INI_HeaderA_2},
                    {"HeaderB_2", context.INI_HeaderB_2},
                    {"Response_2", context.INI_Response_2},
                    {"Anon_2", context.INI_Anon_2},
                    {"TokenCount_2", context.INI_TokenCount_2},
                    {"APICall_2", context.INI_APICall_2},
                    {"APICall_Object_2", context.INI_APICall_Object_2},
                    {"Timeout_2", context.INI_Timeout_2.ToString()},
                    {"MaxOutputToken_2", context.INI_MaxOutputToken_2.ToString()},
                    {"Temperature_2", context.INI_Temperature_2},
                    {"Model_2", context.INI_Model_2},
                    {"APIKeyEncrypted_2", context.INI_APIEncrypted_2.ToString()},
                    {"APIKeyPrefix_2", context.INI_APIKeyPrefix_2},
                    {"OAuth2", context.INI_OAuth2.ToString()},
                    {"OAuth2ClientMail", context.INI_OAuth2ClientMail},
                    {"OAuth2Scopes", context.INI_OAuth2Scopes},
                    {"OAuth2Endpoint", context.INI_OAuth2Endpoint},
                    {"OAuth2ATExpiry", context.INI_OAuth2ATExpiry.ToString()},
                    {"OAuth2_2", context.INI_OAuth2_2.ToString()},
                    {"OAuth2ClientMail_2", context.INI_OAuth2ClientMail_2},
                    {"OAuth2Scopes_2", context.INI_OAuth2Scopes_2},
                    {"OAuth2Endpoint_2", context.INI_OAuth2Endpoint_2},
                    {"OAuth2ATExpiry_2", context.INI_OAuth2ATExpiry_2.ToString()},
                    {"SpeechModelPath", context.INI_SpeechModelPath},
                    {"LocalModelPath", context.INI_LocalModelPath},
                    {"TTSEndpoint", context.INI_TTSEndpoint},
                    {"PromptLib", context.INI_PromptLibPath},
                    {"PromptLibLocal", context.INI_PromptLibPathLocal},
                    {"MyStylePath", context.INI_MyStylePath},
                    {"AlternateModelPath", context.INI_AlternateModelPath},
                    {"SpecialServicePath", context.INI_SpecialServicePath},
                    {"FindClausePath", context.INI_FindClausePath},
                    {"FindClausePathLocal", context.INI_FindClausePathLocal},
                    {"WebAgentPath", context.INI_WebAgentPath},
                    {"WebAgentPathLocal", context.INI_WebAgentPathLocal},
                    {"DocCheckPath", context.INI_DocCheckPath},
                    {"DocCheckPathLocal", context.INI_DocCheckPathLocal},
                    {"DocStylePath", context.INI_DocStylePath},
                    {"DocStylePathLocal", context.INI_DocStylePathLocal},
                    {"PromptLib_Transcript", context.INI_PromptLibPath_Transcript},
                    {"RedactionInstructionsPath", context.INI_RedactionInstructionsPath},
                    {"RedactionInstructionsPathLocal", context.INI_RedactionInstructionsPathLocal},
                    {"ExtractorPath", context.INI_ExtractorPath},
                    {"ExtractorPathLocal", context.INI_ExtractorPathLocal},
                    {"RenameLibPath", context.INI_RenameLibPath},
                    {"RenameLibPathLocal", context.INI_RenameLibPathLocal},
                    {"HelpMeInkyPath", context.INI_HelpMeInkyPath},
                    {"DiscussInkyPath", context.INI_DiscussInkyPath},
                    {"DiscussInkyPathLocal", context.INI_DiscussInkyPathLocal},
                    {"UpdateCheckInterval", context.INI_UpdateCheckInterval.ToString()},
                    {"UpdatePath", context.INI_UpdatePath},
                    {"NoHelperDownload", context.INI_NoHelperDownload.ToString()},
                    {"ToolingLogWindow", context.INI_ToolingLogWindow.ToString()},
                    {"ToolingDryRun", context.INI_ToolingDryRun.ToString()},
                    {"ToolingMaximumIterations", context.INI_ToolingMaximumIterations.ToString()},
                    {"UpdateIni", context.INI_UpdateIni.ToString()},
                    {"UpdateIniAllowRemote", context.INI_UpdateIniAllowRemote.ToString()},
                    {"UpdateIniNoSignature", context.INI_UpdateIniNoSignature.ToString()},
                    {"UpdateSource", context.INI_UpdateSource},
                    {"UpdateClients", context.INI_UpdateClients},
                    {"UpdateIniIgnoreOverride", context.INI_UpdateIniIgnoreOverride},
                    {"UpdateIniSilentMode", context.INI_UpdateIniSilentMode.ToString()},
                    {"UpdateIniSilentLog", context.INI_UpdateIniSilentLog.ToString()}
                }

                ' Read the original ini file content
                Dim originalContent As String = System.IO.File.ReadAllText(IniFilePath)
                Dim updatedContent As New StringBuilder()
                Dim foundKeys As New HashSet(Of String)()

                ' Split into lines and process each line
                Dim iniLines As String() = originalContent.Split({vbCrLf}, StringSplitOptions.None)
                For Each line As String In iniLines
                    Dim trimmedLine As String = line.Trim()

                    ' Preserve comments and empty lines
                    If String.IsNullOrEmpty(trimmedLine) OrElse trimmedLine.StartsWith(";") Then
                        updatedContent.AppendLine(line)
                        Continue For
                    End If

                    ' Process key-value pairs
                    Dim keyValue As String() = trimmedLine.Split(New Char() {"="c}, 2)
                    If keyValue.Length = 2 Then
                        Dim key As String = keyValue(0).Trim()
                        Dim value As String = keyValue(1).Trim()

                        ' Retain keys that are in the expectedKeys dictionary
                        If expectedKeys.ContainsKey(key) Then
                            value = expectedKeys(key)
                            foundKeys.Add(key)
                            updatedContent.AppendLine($"{key} = {value}")
                        End If
                    End If
                Next

                ' Add missing keys to the updated content
                For Each key In expectedKeys.Keys.Except(foundKeys)
                    updatedContent.AppendLine($"{key} = {expectedKeys(key)}")
                Next

                ' Write the updated content to the temporary ini file
                System.IO.File.WriteAllText(TempIniFilePath, updatedContent.ToString())

                ' Replace the original file with the updated file
                System.IO.File.Delete(IniFilePath)
                System.IO.File.Move(TempIniFilePath, IniFilePath)

                context.INIloaded = False

                ShowCustomMessageBox("Configuration file has been updated.")

                InitializeConfig(context, False, True)

            Catch ex As System.Exception
                ShowCustomMessageBox($"Error resetting configuration file: {ex.Message}")

            End Try
        End Sub


        ''' <summary>
        ''' Returns the resolved path of the active configuration file (<c>{AN2}.ini</c>) based on registry and default
        ''' location precedence.
        ''' </summary>
        ''' <param name="context">Shared context providing the current application identifier (<c>RDV</c>).</param>
        ''' <returns>The active configuration file path with CR characters removed.</returns>
        Public Shared Function GetActiveConfigFilePath(context As ISharedContext) As String
            Dim regPath As String = GetFromRegistry(RegPath_Base, RegPath_IniPath, True)
            Dim defaultPathApp As String = GetDefaultINIPath(context.RDV)
            Dim defaultPathWord As String = GetDefaultINIPath("Word")
            Dim candidate As String

            If Not String.IsNullOrWhiteSpace(regPath) AndAlso RegPath_IniPrio Then
                candidate = System.IO.Path.Combine(ExpandEnvironmentVariables(regPath), $"{AN2}.ini")
            ElseIf System.IO.File.Exists(defaultPathApp) Then
                candidate = defaultPathApp
            ElseIf System.IO.File.Exists(defaultPathWord) Then
                candidate = defaultPathWord
            ElseIf Not String.IsNullOrWhiteSpace(regPath) Then
                candidate = System.IO.Path.Combine(ExpandEnvironmentVariables(regPath), $"{AN2}.ini")
            Else
                candidate = defaultPathApp
            End If
            Return RemoveCR(candidate)
        End Function


        ''' <summary>
        ''' Shows a modal "Expert Configuration" window containing a grid of variable names and editable values.
        ''' </summary>
        ''' <param name="variableNames">List of variable names shown as read-only in the first column.</param>
        ''' <param name="variableValues">
        ''' Dictionary of variable name to current value. Values are written back as strings when saved.
        ''' If <c>Nothing</c>, a new case-insensitive dictionary is created.
        ''' </param>
        ''' <param name="context">Shared context used when resolving and editing related `.ini` files.</param>
        ''' <param name="ownerForm">Optional dialog owner.</param>
        ''' <returns>
        ''' The edited dictionary when the dialog is saved; otherwise <c>Nothing</c>.
        ''' Returns <c>Nothing</c> as well when the user chooses the "reload from disk" flow.
        ''' </returns>
        Public Shared Function ShowVariableConfigurationWindow(
        variableNames As List(Of String),
        variableValues As Dictionary(Of String, Object),
        context As ISharedContext,
        Optional ownerForm As Form = Nothing
    ) As Dictionary(Of String, Object)

            If variableValues Is Nothing Then
                variableValues = New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
            End If

            Dim abortAndReload As Boolean = False
            Dim gridTouched As Boolean = False

            Dim form As New Form() With {
                        .Text = "Expert Configuration",
                        .StartPosition = FormStartPosition.CenterScreen,
                        .ClientSize = New Size(900, 520),
                        .Font = New System.Drawing.Font("Segoe UI", 9.0F)
                    }

            ' Icon / logo
            Try
                Dim bmp As New System.Drawing.Bitmap(My.Resources.Red_Ink_Logo)
                form.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())
            Catch
            End Try

            Dim dgv As New DataGridView() With {
                        .Dock = DockStyle.Fill,
                        .AllowUserToAddRows = False,
                        .AllowUserToDeleteRows = False,
                        .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                        .RowHeadersVisible = False,
                        .SelectionMode = DataGridViewSelectionMode.CellSelect,
                        .MultiSelect = False
                    }

            Dim colVar As New System.Windows.Forms.DataGridViewTextBoxColumn() With {
                    .HeaderText = "Variable",
                    .ReadOnly = True,
                    .Name = "colVar",
                    .FillWeight = 30
                }

            Dim colVal As New System.Windows.Forms.DataGridViewTextBoxColumn() With {
                    .HeaderText = "Value",
                    .Name = "colVal",
                    .FillWeight = 70
                }


            dgv.Columns.AddRange(colVar, colVal)

            dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            dgv.ColumnHeadersHeight = CInt(form.Font.Height * 2.4)

            For Each vName In variableNames
                Dim val As String = ""
                If variableValues.ContainsKey(vName) AndAlso variableValues(vName) IsNot Nothing Then
                    val = variableValues(vName).ToString()
                End If
                dgv.Rows.Add(vName, val)
            Next

            AddHandler dgv.CellValueChanged, Sub() gridTouched = True
            AddHandler dgv.CurrentCellDirtyStateChanged, Sub()
                                                             If dgv.IsCurrentCellDirty Then dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
                                                         End Sub

            Dim btnSaveClose As New System.Windows.Forms.Button() With {.Text = "Save && Close", .AutoSize = True, .Margin = New Padding(10)}
            Dim btnEditIni As New System.Windows.Forms.Button() With {.Text = "Edit .ini Files", .AutoSize = True, .Margin = New Padding(10)}
            Dim btnCancel As New System.Windows.Forms.Button() With {.Text = "Cancel", .AutoSize = True, .Margin = New Padding(10)}
            Dim btnImportIni As New System.Windows.Forms.Button() With {.Text = "Load Settings From A Source", .AutoSize = True, .Margin = New Padding(10)}

            Dim pnlGridHost As New System.Windows.Forms.Panel() With {
            .Dock = System.Windows.Forms.DockStyle.Fill,
            .Padding = New System.Windows.Forms.Padding(15, 15, 15, 15)
        }


            Dim pnlButtons As New FlowLayoutPanel() With {
                    .Dock = DockStyle.Bottom,
                    .FlowDirection = FlowDirection.RightToLeft,
                    .AutoSize = True,
                    .Padding = New Padding(15, 15, 15, 15),
                    .WrapContents = False
                }
            pnlButtons.Controls.AddRange({btnCancel, btnSaveClose, btnEditIni, btnImportIni})

            pnlGridHost.Controls.Add(dgv)
            form.Controls.Add(pnlGridHost)
            form.Controls.Add(pnlButtons)

            ' Enable / disable Import Settings button using same eligibility rules
            Try
                Dim activeIniPath As String = Nothing
                Try
                    activeIniPath = GetActiveConfigFilePath(context)
                Catch
                End Try

                Dim disableReason As String = Nothing
                Dim canImport As Boolean = True

                ' Registry-controlled INI => disable
                Try
                    If RegPath_IniPrio Then
                        canImport = False
                    End If
                Catch
                End Try

                ' Excel / Outlook using Word INI => disable
                If canImport Then
                    Dim rdv As String = Nothing
                    Try
                        rdv = context.RDV
                    Catch
                    End Try

                    If String.Equals(rdv, "Excel", StringComparison.OrdinalIgnoreCase) OrElse
           String.Equals(rdv, "Outlook", StringComparison.OrdinalIgnoreCase) Then

                        Dim wordPath As String = Nothing
                        Try
                            wordPath = GetDefaultINIPath("Word")
                        Catch
                        End Try

                        If Not String.IsNullOrWhiteSpace(wordPath) AndAlso
               String.Equals(activeIniPath, wordPath, StringComparison.OrdinalIgnoreCase) Then
                            canImport = False
                        End If
                    End If
                End If

                btnImportIni.Enabled = canImport
            Catch
                btnImportIni.Enabled = False
            End Try


            ''' <summary>
            ''' Copies the current grid values into <c>variableValues</c> as strings by variable name.
            ''' </summary>
            Dim syncGridToDictionary As Action =
                            Sub()
                                For Each row As DataGridViewRow In dgv.Rows
                                    If row.IsNewRow Then Continue For
                                    Dim name = CStr(row.Cells("colVar").Value)
                                    Dim valueStr = CStr(If(row.Cells("colVal").Value, ""))
                                    If Not String.IsNullOrWhiteSpace(name) Then
                                        variableValues(name) = valueStr
                                    End If
                                Next
                            End Sub

            AddHandler btnEditIni.Click,
        Sub()
            ' Build selectable INI file list
            Dim fileMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim mainPath As String = Nothing
            Try
                mainPath = GetActiveConfigFilePath(context)
            Catch
            End Try

            If Not String.IsNullOrWhiteSpace(mainPath) AndAlso System.IO.File.Exists(mainPath) Then
                fileMap.Add($"Main ({AN2}.ini) - " & mainPath, mainPath)
            End If

            ' Default per-application INIs
            For Each kvp In DefaultINIPaths
                Dim p = ExpandEnvironmentVariables(kvp.Value)
                If System.IO.File.Exists(p) AndAlso Not fileMap.ContainsValue(p) Then
                    fileMap.Add($"{kvp.Key} - {p}", p)
                End If
            Next

            ' Alternate model path
            If Not String.IsNullOrWhiteSpace(context.INI_AlternateModelPath) Then
                Dim alt = ExpandEnvironmentVariables(context.INI_AlternateModelPath)
                If System.IO.File.Exists(alt) AndAlso Not fileMap.ContainsValue(alt) Then
                    fileMap.Add("Alternate Model - " & alt, alt)
                End If
            End If

            ' Special service path
            If Not String.IsNullOrWhiteSpace(context.INI_SpecialServicePath) Then
                Dim sp = ExpandEnvironmentVariables(context.INI_SpecialServicePath)
                If System.IO.File.Exists(sp) AndAlso Not fileMap.ContainsValue(sp) Then
                    fileMap.Add("Special Service - " & sp, sp)
                End If
            End If

            If fileMap.Count = 0 Then
                ShowCustomMessageBox("No .ini files found.")
                Exit Sub
            End If

            ' Z-order fix: temporarily disable form so selection UI is not blocked
            Dim wasTopMost = form.TopMost
            form.TopMost = False
            form.Enabled = False
            System.Windows.Forms.Application.DoEvents()

            Dim choice As String = Nothing
            Try
                choice = ShowSelectionForm("Select a .ini file to open (ESC to cancel):", "INI Files", fileMap.Keys)
            Finally
                form.Enabled = True
                form.TopMost = wasTopMost
                form.Activate()
            End Try

            If String.IsNullOrWhiteSpace(choice) OrElse Not fileMap.ContainsKey(choice) Then Exit Sub
            Dim selectedPath = fileMap(choice)

            ' Only warn about in-memory vs disk version IF the main config file was chosen
            If Not String.IsNullOrWhiteSpace(mainPath) AndAlso
               String.Equals(selectedPath, mainPath, StringComparison.OrdinalIgnoreCase) Then

                Dim proceed = ShowCustomYesNoBox(
                    "You selected the main '" & AN2 & ".ini' configuration file. " &
                    "The editor will open the on-disk version. Unsaved grid changes are NOT included unless you click 'Save && Close' first. " &
                    "Proceed opening the file?",
                    "Yes, proceed", "No, cancel")
                If proceed = 2 Then Exit Sub
            End If

            Try
                ShowTextFileEditor(selectedPath, "Editing " & selectedPath)
            Catch ex As Exception
                ShowCustomMessageBox("Could not open editor: " & ex.Message)
                Exit Sub
            End Try

            ' Post-edit logic ONLY for main config file
            If Not String.IsNullOrWhiteSpace(mainPath) AndAlso
               String.Equals(selectedPath, mainPath, StringComparison.OrdinalIgnoreCase) Then

                Dim answer = ShowCustomYesNoBox(
                    "Do you want to reload the updated '" & AN2 & ".ini' from the disk NOW (and discard any changes on the grid or in the memory)?",
                    "Yes, close and reload", "No, stay here")
                If answer = 1 Then
                    abortAndReload = True
                    form.DialogResult = DialogResult.OK
                    form.Close()
                End If
            End If
        End Sub

            AddHandler btnImportIni.Click,
                Sub()
                    ' Same Z-order handling as Edit .ini Files
                    Dim wasTopMost As Boolean = form.TopMost
                    form.TopMost = False
                    form.Enabled = False
                    System.Windows.Forms.Application.DoEvents()

                    Dim ChangesToMainConfig = False

                    Try
                        ChangesToMainConfig = IniImportManager.RunImportFromVariableConfigurationWindow(context, form)
                    Catch ex As System.Exception
                        ShowCustomMessageBox("Import failed: " & ex.Message)
                    Finally
                        form.Enabled = True
                        form.TopMost = wasTopMost
                        form.Activate()
                    End Try

                    If ChangesToMainConfig Then
                        Dim NextStep = ShowCustomYesNoBox("Your import of settings changed the main configuration file. They are not yet reflected in this table. You should close this window and reload the configuration to avoid conflicts (if you cancel, this Expert Configuration window closes without reloading). Proceed?", "Yes, close and reload", "No, stay here")
                        Select Case NextStep
                            Case 1
                                abortAndReload = True
                                form.DialogResult = DialogResult.OK
                                form.Close()
                            Case 0
                                abortAndReload = False
                                form.DialogResult = DialogResult.Cancel
                                form.Close()
                        End Select
                    End If

                End Sub


            AddHandler btnSaveClose.Click,
        Sub()
            syncGridToDictionary()
            form.DialogResult = DialogResult.OK
            form.Close()
        End Sub

            AddHandler btnCancel.Click,
        Sub()
            variableValues = Nothing
            form.DialogResult = DialogResult.Cancel
            form.Close()
        End Sub

            If ownerForm IsNot Nothing Then
                form.ShowDialog(ownerForm)
            Else
                form.ShowDialog()
            End If

            If abortAndReload Then

                context.INIloaded = False
                InitializeConfig(context, False, True)

                Return Nothing

            End If

            If form.DialogResult = DialogResult.OK Then
                Return variableValues
            End If

            Return Nothing
        End Function



        ''' <summary>
        ''' Shows the expert configuration UI for the current <paramref name="context"/> by building a variable name/value
        ''' dictionary from many context fields, displaying it in <see cref="ShowVariableConfigurationWindow"/>, and then
        ''' copying returned values back into <paramref name="context"/>.
        ''' </summary>
        ''' <param name="context">Shared context whose configuration values are displayed and (optionally) updated.</param>
        ''' <param name="ownerform">Owner form for the modal dialog.</param>
        Public Shared Sub ShowExpertConfiguration(ByRef context As ISharedContext, ownerform As Form)
            ' Dictionary to store variable names and their current values
            Dim variableValues As New Dictionary(Of String, Object)

            ' Populate the dictionary with all the required variables
            variableValues.Add("APIKey", context.INI_APIKeyBack) ' Use Context.INI_APIKeyBack, display as Context.INI_APIKey
            variableValues.Add("Temperature", context.INI_Temperature)
            variableValues.Add("Timeout", context.INI_Timeout)
            variableValues.Add("MaxOutputToken", context.INI_MaxOutputToken)
            variableValues.Add("Model", context.INI_Model)
            variableValues.Add("Endpoint", context.INI_Endpoint)
            variableValues.Add("HeaderA", context.INI_HeaderA)
            variableValues.Add("HeaderB", context.INI_HeaderB)
            variableValues.Add("APICall", context.INI_APICall)
            variableValues.Add("APICall_Object", context.INI_APICall_Object)
            variableValues.Add("Response", context.INI_Response)
            variableValues.Add("Anon", context.INI_Anon)
            variableValues.Add("TokenCount", context.INI_TokenCount)
            variableValues.Add("DoubleS", context.INI_DoubleS)
            variableValues.Add("Clean", context.INI_Clean)
            variableValues.Add("Ignore", context.INI_Ignore)
            variableValues.Add("Location", context.INI_Location)
            variableValues.Add("NoEmDash", context.INI_NoDash)
            variableValues.Add("DefaultPrefix", context.INI_DefaultPrefix)
            variableValues.Add("MarkdownBubbles", context.INI_MarkdownBubbles)
            variableValues.Add("PreCorrection", context.INI_PreCorrection)
            variableValues.Add("PostCorrection", context.INI_PostCorrection)
            variableValues.Add("APIEncrypted", context.INI_APIEncrypted)
            variableValues.Add("APIKeyPrefix", context.INI_APIKeyPrefix)
            variableValues.Add("MarkupMethodOutlook", context.INI_MarkupMethodOutlook)
            variableValues.Add("MarkupDiffCap", context.INI_MarkupDiffCap)
            variableValues.Add("MarkupRegexCap", context.INI_MarkupRegexCap)
            variableValues.Add("ChatCap", context.INI_ChatCap)
            variableValues.Add("OAuth2", context.INI_OAuth2)
            variableValues.Add("OAuth2ClientMail", context.INI_OAuth2ClientMail)
            variableValues.Add("OAuth2Scopes", context.INI_OAuth2Scopes)
            variableValues.Add("OAuth2Endpoint", context.INI_OAuth2Endpoint)
            variableValues.Add("OAuth2ATExpiry", context.INI_OAuth2ATExpiry)
            variableValues.Add("SecondAPI", context.INI_SecondAPI)
            variableValues.Add("APIKey_2", context.INI_APIKeyBack_2) ' Use Context.INI_APIKeyBack_2, display as Context.INI_APIKey_2
            variableValues.Add("Temperature_2", context.INI_Temperature_2)
            variableValues.Add("Timeout_2", context.INI_Timeout_2)
            variableValues.Add("MaxOutputToken_2", context.INI_MaxOutputToken_2)
            variableValues.Add("Model_2", context.INI_Model_2)
            variableValues.Add("Endpoint_2", context.INI_Endpoint_2)
            variableValues.Add("HeaderA_2", context.INI_HeaderA_2)
            variableValues.Add("HeaderB_2", context.INI_HeaderB_2)
            variableValues.Add("APICall_2", context.INI_APICall_2)
            variableValues.Add("APICall_Object_2", context.INI_APICall_Object_2)
            variableValues.Add("Response_2", context.INI_Response_2)
            variableValues.Add("Anon_2", context.INI_Anon_2)
            variableValues.Add("TokenCount_2", context.INI_TokenCount_2)
            variableValues.Add("APIEncrypted_2", context.INI_APIEncrypted_2)
            variableValues.Add("APIKeyPrefix_2", context.INI_APIKeyPrefix_2)
            variableValues.Add("OAuth2_2", context.INI_OAuth2_2)
            variableValues.Add("OAuth2ClientMail_2", context.INI_OAuth2ClientMail_2)
            variableValues.Add("OAuth2Scopes_2", context.INI_OAuth2Scopes_2)
            variableValues.Add("OAuth2Endpoint_2", context.INI_OAuth2Endpoint_2)
            variableValues.Add("OAuth2ATExpiry_2", context.INI_OAuth2ATExpiry_2)
            variableValues.Add("APIDebug", context.INI_APIDebug)
            variableValues.Add("UsageRestrictions", context.INI_UsageRestrictions)
            variableValues.Add("LogPath", context.INI_LogPath)
            variableValues.Add("Language1", context.INI_Language1)
            variableValues.Add("Language2", context.INI_Language2)
            variableValues.Add("KeepFormat1", context.INI_KeepFormat1)
            variableValues.Add("MarkdownConvert", context.INI_MarkdownConvert)
            variableValues.Add("KeepFormat2", context.INI_KeepFormat2)
            variableValues.Add("KeepFormatCap", context.INI_KeepFormatCap)
            variableValues.Add("KeepParaFormatInline", context.INI_KeepParaFormatInline)
            variableValues.Add("ReplaceText1", context.INI_ReplaceText1)
            variableValues.Add("ReplaceText2", context.INI_ReplaceText2)
            variableValues.Add("ReplaceText2Override", context.INI_ReplaceText2Override)
            variableValues.Add("DoMarkupOutlook", context.INI_DoMarkupOutlook)
            variableValues.Add("DoMarkupWord", context.INI_DoMarkupWord)
            variableValues.Add("ISearch", context.INI_ISearch)
            variableValues.Add("ISearch_Approve", context.INI_ISearch_Approve)
            variableValues.Add("ISearch_URL", context.INI_ISearch_URL)
            variableValues.Add("ISearch_ResponseMask1", context.INI_ISearch_ResponseMask1)
            variableValues.Add("ISearch_ResponseMask2", context.INI_ISearch_ResponseMask2)
            variableValues.Add("ISearch_Name", context.INI_ISearch_Name)
            variableValues.Add("ISearch_Tries", context.INI_ISearch_Tries)
            variableValues.Add("ISearch_Results", context.INI_ISearch_Results)
            variableValues.Add("ISearch_MaxDepth", context.INI_ISearch_MaxDepth)
            variableValues.Add("ISearch_Timeout", context.INI_ISearch_Timeout)
            variableValues.Add("ISearch_SearchTerm_SP", context.INI_ISearch_SearchTerm_SP)
            variableValues.Add("ISearch_Apply_SP", context.INI_ISearch_Apply_SP)
            variableValues.Add("ISearch_Apply_SP_Markup", context.INI_ISearch_Apply_SP_Markup)
            variableValues.Add("Lib", context.INI_Lib)
            variableValues.Add("Lib_File", context.INI_Lib_File)
            variableValues.Add("Lib_Timeout", context.INI_Lib_Timeout)
            variableValues.Add("Lib_Find_SP", context.INI_Lib_Find_SP)
            variableValues.Add("Lib_Apply_SP", context.INI_Lib_Apply_SP)
            variableValues.Add("Lib_Apply_SP_Markup", context.INI_Lib_Apply_SP_Markup)
            variableValues.Add("MarkupMethodHelper", context.INI_MarkupMethodHelper)
            variableValues.Add("MarkupMethodWord", context.INI_MarkupMethodWord)
            variableValues.Add("MarkupMethodWordOverride", context.INI_MarkupMethodWordOverride)
            variableValues.Add("MarkupMethodOutlookOverride", context.INI_MarkupMethodOutlookOverride)
            variableValues.Add("ContextMenu", context.INI_ContextMenu)
            variableValues.Add("UpdateCheckInterval", context.INI_UpdateCheckInterval)
            variableValues.Add("UpdatePath", context.INI_UpdatePath)
            variableValues.Add("HelpMeInkyPath", context.INI_HelpMeInkyPath)
            variableValues.Add("DiscussInkyPath", context.INI_DiscussInkyPath)
            variableValues.Add("DiscussInkyPathLocal", context.INI_DiscussInkyPathLocal)
            variableValues.Add("RedactionInstructionsPath", context.INI_RedactionInstructionsPath)
            variableValues.Add("RedactionInstructionsPathLocal", context.INI_RedactionInstructionsPathLocal)
            variableValues.Add("ExtractorPath", context.INI_ExtractorPath)
            variableValues.Add("ExtractorPathLocal", context.INI_ExtractorPathLocal)
            variableValues.Add("RenameLibPath", context.INI_RenameLibPath)
            variableValues.Add("RenameLibPathLocal", context.INI_RenameLibPathLocal)
            variableValues.Add("SpeechModelPath", context.INI_SpeechModelPath)
            variableValues.Add("LocalModelPath", context.INI_LocalModelPath)
            variableValues.Add("TTSEndpoint", context.INI_TTSEndpoint)
            variableValues.Add("ShortcutsWordExcel", context.INI_ShortcutsWordExcel)
            variableValues.Add("PromptLib", context.INI_PromptLibPath)
            variableValues.Add("PromptLibLocal", context.INI_PromptLibPathLocal)
            variableValues.Add("MyStylePath", context.INI_MyStylePath)
            variableValues.Add("AlternateModelPath", context.INI_AlternateModelPath)
            variableValues.Add("SpecialServicePath", context.INI_SpecialServicePath)
            variableValues.Add("FindClausePath", context.INI_FindClausePath)
            variableValues.Add("FindClausePathLocal", context.INI_FindClausePathLocal)
            variableValues.Add("WebAgentPath", context.INI_WebAgentPath)
            variableValues.Add("WebAgentPathLocal", context.INI_WebAgentPathLocal)
            variableValues.Add("DocCheckPath", context.INI_DocCheckPath)
            variableValues.Add("DocCheckPathLocal", context.INI_DocCheckPathLocal)
            variableValues.Add("DocStylePath", context.INI_DocStylePath)
            variableValues.Add("DocStylePathLocal", context.INI_DocStylePathLocal)
            variableValues.Add("PromptLib_Transcript", context.INI_PromptLibPath_Transcript)
            variableValues.Add("SP_Translate", context.SP_Translate)
            variableValues.Add("SP_Correct", context.SP_Correct)
            variableValues.Add("SP_Improve", context.SP_Improve)
            variableValues.Add("SP_Explain", context.SP_Explain)
            variableValues.Add("SP_FindClause", context.SP_FindClause)
            variableValues.Add("SP_FindClause_Clean", context.SP_FindClause_Clean)
            variableValues.Add("SP_ApplyDocStyle", context.SP_ApplyDocStyle)
            variableValues.Add("SP_ApplyDocStyle_NumberingHint", context.SP_ApplyDocStyle_NumberingHint)
            variableValues.Add("SP_DocCheck_Clause", context.SP_DocCheck_Clause)
            variableValues.Add("SP_DocCheck_MultiClause", context.SP_DocCheck_MultiClause)
            variableValues.Add("SP_DocCheck_MultiClauseSum", context.SP_DocCheck_MultiClauseSum)
            variableValues.Add("SP_DocCheck_MultiClauseSum_Bubbles", context.SP_DocCheck_MultiClauseSum_Bubbles)
            variableValues.Add("SP_SuggestTitles", context.SP_SuggestTitles)
            variableValues.Add("SP_Friendly", context.SP_Friendly)
            variableValues.Add("SP_Convincing", context.SP_Convincing)
            variableValues.Add("SP_NoFillers", context.SP_NoFillers)
            variableValues.Add("SP_Podcast", context.SP_Podcast)
            variableValues.Add("SP_MyStyle_Word", context.SP_MyStyle_Word)
            variableValues.Add("SP_MyStyle_Outlook", context.SP_MyStyle_Outlook)
            variableValues.Add("SP_MyStyle_Apply", context.SP_MyStyle_Apply)
            variableValues.Add("SP_Shorten", context.SP_Shorten)
            variableValues.Add("SP_Filibuster", context.SP_Filibuster)
            variableValues.Add("SP_ArgueAgainst", context.SP_ArgueAgainst)
            variableValues.Add("SP_InsertClipboard", context.SP_InsertClipboard)
            variableValues.Add("SP_Summarize", context.SP_Summarize)
            variableValues.Add("SP_Markup", context.SP_Markup)
            variableValues.Add("SP_MailReply", context.SP_MailReply)
            variableValues.Add("SP_MailSumup", context.SP_MailSumup)
            variableValues.Add("SP_MailSumup2", context.SP_MailSumup2)
            variableValues.Add("SP_FreestyleText", context.SP_FreestyleText)
            variableValues.Add("SP_FreestyleNoText", context.SP_FreestyleNoText)
            variableValues.Add("SP_SwitchParty", context.SP_SwitchParty)
            variableValues.Add("SP_Anonymize", context.SP_Anonymize)
            variableValues.Add("SP_Rename", context.SP_Rename)
            variableValues.Add("SP_Redact", context.SP_Redact)
            variableValues.Add("SP_CheckforII", context.SP_CheckforII)
            variableValues.Add("SP_Extract", context.SP_Extract)
            variableValues.Add("SP_ExtractSchema", context.SP_ExtractSchema)
            variableValues.Add("SP_MergeDateRows", context.SP_MergeDateRows)
            variableValues.Add("SP_ContextSearch", context.SP_ContextSearch)
            variableValues.Add("SP_ContextSearchMulti", context.SP_ContextSearchMulti)
            variableValues.Add("SP_RangeOfCells", context.SP_RangeOfCells)
            variableValues.Add("SP_ParseFile", context.SP_ParseFile)
            variableValues.Add("SP_Ignore", context.SP_Ignore)
            variableValues.Add("SP_WriteNeatly", context.SP_WriteNeatly)
            variableValues.Add("SP_Add_KeepFormulasIntact", context.SP_Add_KeepFormulasIntact)
            variableValues.Add("SP_Add_KeepHTMLIntact", context.SP_Add_KeepHTMLIntact)
            variableValues.Add("SP_Add_KeepInlineIntact", context.SP_Add_KeepInlineIntact)
            variableValues.Add("SP_Add_Bubbles", context.SP_Add_Bubbles)
            variableValues.Add("SP_Add_BubblesReply", context.SP_Add_BubblesReply)
            variableValues.Add("SP_Add_BubblesExtract", context.SP_Add_BubblesExtract)
            variableValues.Add("SP_Add_Bubbles_Format", context.SP_Add_Bubbles_Format)
            variableValues.Add("SP_Add_Batch", context.SP_Add_Batch)
            variableValues.Add("SP_Add_Tooling", context.SP_Add_Tooling)
            variableValues.Add("SP_Add_Slides", context.SP_Add_Slides)
            variableValues.Add("SP_BubblesExcel", context.SP_BubblesExcel)
            variableValues.Add("SP_Add_Revisions", context.SP_Add_Revisions)
            variableValues.Add("SP_MarkupRegex", context.SP_MarkupRegex)
            variableValues.Add("SP_ChatWord", context.SP_ChatWord)
            variableValues.Add("SP_HelpMe", context.SP_HelpMe)
            variableValues.Add("SP_Chat", context.SP_Chat)
            variableValues.Add("SP_Add_ChatWord_Commands", context.SP_Add_ChatWord_Commands)
            variableValues.Add("SP_Add_Chat_NoCommands", context.SP_Add_Chat_NoCommands)
            variableValues.Add("SP_ChatExcel", context.SP_ChatExcel)
            variableValues.Add("SP_Add_ChatExcel_Commands", context.SP_Add_ChatExcel_Commands)
            variableValues.Add("SP_Add_MergePrompt", context.SP_Add_MergePrompt)
            variableValues.Add("SP_FindPrompts", context.SP_FindPrompts)
            variableValues.Add("SP_MergePrompt", context.SP_MergePrompt)
            variableValues.Add("SP_MergePrompt2", context.SP_MergePrompt2)
            variableValues.Add("NoHelperDownload", context.INI_NoHelperDownload)
            variableValues.Add("ToolingLogWindow", context.INI_ToolingLogWindow)
            variableValues.Add("ToolingDryRun", context.INI_ToolingDryRun)
            variableValues.Add("ToolingMaximumIterations", context.INI_ToolingMaximumIterations)
            variableValues.Add("UpdateIni", context.INI_UpdateIni)
            variableValues.Add("UpdateIniAllowRemote", context.INI_UpdateIniAllowRemote)
            variableValues.Add("UpdateIniNoSignature", context.INI_UpdateIniNoSignature)
            variableValues.Add("UpdateSource", context.INI_UpdateSource)
            variableValues.Add("UpdateClients", context.INI_UpdateClients)
            variableValues.Add("UpdateIniIgnoreOverride", context.INI_UpdateIniIgnoreOverride)
            variableValues.Add("UpdateIniSilentMode", context.INI_UpdateIniSilentMode)
            variableValues.Add("UpdateIniSilentLog", context.INI_UpdateIniSilentLog)

            ' Extract variable names from the dictionary
            Dim variableNames As New List(Of String)(variableValues.Keys)

            ' Call the ShowVariableConfigurationWindow function and get the updated values            
            Dim updatedValues = ShowVariableConfigurationWindow(variableNames, variableValues, context, ownerform)

            If Not IsNothing(updatedValues) Then

                ' Update the original variables with the returned values
                If updatedValues.ContainsKey("APIKey") Then context.INI_APIKeyBack = CStr(updatedValues("APIKey"))
                If updatedValues.ContainsKey("Temperature") Then context.INI_Temperature = CStr(updatedValues("Temperature"))
                If updatedValues.ContainsKey("Timeout") Then context.INI_Timeout = CLng(updatedValues("Timeout"))
                If updatedValues.ContainsKey("MaxOutputToken") Then context.INI_MaxOutputToken = CInt(updatedValues("MaxOutputToken"))
                If updatedValues.ContainsKey("Model") Then context.INI_Model = CStr(updatedValues("Model"))
                If updatedValues.ContainsKey("Endpoint") Then context.INI_Endpoint = CStr(updatedValues("Endpoint"))
                If updatedValues.ContainsKey("HeaderA") Then context.INI_HeaderA = CStr(updatedValues("HeaderA"))
                If updatedValues.ContainsKey("HeaderB") Then context.INI_HeaderB = CStr(updatedValues("HeaderB"))
                If updatedValues.ContainsKey("APICall") Then context.INI_APICall = CStr(updatedValues("APICall"))
                If updatedValues.ContainsKey("APICall_Object") Then context.INI_APICall_Object = CStr(updatedValues("APICall_Object"))
                If updatedValues.ContainsKey("Response") Then context.INI_Response = CStr(updatedValues("Response"))
                If updatedValues.ContainsKey("Anon") Then context.INI_Anon = CStr(updatedValues("Anon"))
                If updatedValues.ContainsKey("TokenCount") Then context.INI_TokenCount = CStr(updatedValues("TokenCount"))
                If updatedValues.ContainsKey("DoubleS") Then context.INI_DoubleS = CBool(updatedValues("DoubleS"))
                If updatedValues.ContainsKey("Clean") Then context.INI_Clean = CBool(updatedValues("Clean"))
                If updatedValues.ContainsKey("Ignore") Then context.INI_Ignore = CBool(updatedValues("Ignore"))
                If updatedValues.ContainsKey("Location") Then context.INI_Location = CStr(updatedValues("Location"))
                If updatedValues.ContainsKey("NoEmDash") Then context.INI_NoDash = CBool(updatedValues("NoEmDash"))
                If updatedValues.ContainsKey("DefaultPrefix") Then context.INI_DefaultPrefix = CStr(updatedValues("DefaultPrefix"))
                If updatedValues.ContainsKey("MarkdownBubbles") Then context.INI_MarkdownBubbles = CBool(updatedValues("MarkdownBubbles"))
                If updatedValues.ContainsKey("PreCorrection") Then context.INI_PreCorrection = CStr(updatedValues("PreCorrection"))
                If updatedValues.ContainsKey("PostCorrection") Then context.INI_PostCorrection = CStr(updatedValues("PostCorrection"))
                If updatedValues.ContainsKey("APIEncrypted") Then context.INI_APIEncrypted = CBool(updatedValues("APIEncrypted"))
                If updatedValues.ContainsKey("APIKeyPrefix") Then context.INI_APIKeyPrefix = CStr(updatedValues("APIKeyPrefix"))
                If updatedValues.ContainsKey("MarkupMethodOutlook") Then context.INI_MarkupMethodOutlook = CInt(updatedValues("MarkupMethodOutlook"))
                If updatedValues.ContainsKey("MarkupDiffCap") Then context.INI_MarkupDiffCap = CInt(updatedValues("MarkupDiffCap"))
                If updatedValues.ContainsKey("MarkupRegexCap") Then context.INI_MarkupRegexCap = CInt(updatedValues("MarkupRegexCap"))
                If updatedValues.ContainsKey("ChatCap") Then context.INI_ChatCap = CInt(updatedValues("ChatCap"))
                If updatedValues.ContainsKey("OAuth2") Then context.INI_OAuth2 = CBool(updatedValues("OAuth2"))
                If updatedValues.ContainsKey("OAuth2ClientMail") Then context.INI_OAuth2ClientMail = CStr(updatedValues("OAuth2ClientMail"))
                If updatedValues.ContainsKey("OAuth2Scopes") Then context.INI_OAuth2Scopes = CStr(updatedValues("OAuth2Scopes"))
                If updatedValues.ContainsKey("OAuth2Endpoint") Then context.INI_OAuth2Endpoint = CStr(updatedValues("OAuth2Endpoint"))
                If updatedValues.ContainsKey("OAuth2ATExpiry") Then context.INI_OAuth2ATExpiry = CLng(updatedValues("OAuth2ATExpiry"))
                If updatedValues.ContainsKey("SecondAPI") Then context.INI_SecondAPI = CBool(updatedValues("SecondAPI"))
                If updatedValues.ContainsKey("APIKey_2") Then context.INI_APIKeyBack_2 = CStr(updatedValues("APIKey_2"))
                If updatedValues.ContainsKey("Temperature_2") Then context.INI_Temperature_2 = CStr(updatedValues("Temperature_2"))
                If updatedValues.ContainsKey("Timeout_2") Then context.INI_Timeout_2 = CLng(updatedValues("Timeout_2"))
                If updatedValues.ContainsKey("MaxOutputToken_2") Then context.INI_MaxOutputToken_2 = CInt(updatedValues("MaxOutputToken_2"))
                If updatedValues.ContainsKey("Model_2") Then context.INI_Model_2 = CStr(updatedValues("Model_2"))
                If updatedValues.ContainsKey("Endpoint_2") Then context.INI_Endpoint_2 = CStr(updatedValues("Endpoint_2"))
                If updatedValues.ContainsKey("HeaderA_2") Then context.INI_HeaderA_2 = CStr(updatedValues("HeaderA_2"))
                If updatedValues.ContainsKey("HeaderB_2") Then context.INI_HeaderB_2 = CStr(updatedValues("HeaderB_2"))
                If updatedValues.ContainsKey("APICall_2") Then context.INI_APICall_2 = CStr(updatedValues("APICall_2"))
                If updatedValues.ContainsKey("APICall_Object_2") Then context.INI_APICall_Object_2 = CStr(updatedValues("APICall_Object_2"))
                If updatedValues.ContainsKey("Response_2") Then context.INI_Response_2 = CStr(updatedValues("Response_2"))
                If updatedValues.ContainsKey("Anon_2") Then context.INI_Anon_2 = CStr(updatedValues("Anon_2"))
                If updatedValues.ContainsKey("TokenCount_2") Then context.INI_TokenCount_2 = CStr(updatedValues("TokenCount_2"))
                If updatedValues.ContainsKey("APIEncrypted_2") Then context.INI_APIEncrypted_2 = CBool(updatedValues("APIEncrypted_2"))
                If updatedValues.ContainsKey("APIKeyPrefix_2") Then context.INI_APIKeyPrefix_2 = CStr(updatedValues("APIKeyPrefix_2"))
                If updatedValues.ContainsKey("OAuth2_2") Then context.INI_OAuth2_2 = CBool(updatedValues("OAuth2_2"))
                If updatedValues.ContainsKey("OAuth2ClientMail_2") Then context.INI_OAuth2ClientMail_2 = CStr(updatedValues("OAuth2ClientMail_2"))
                If updatedValues.ContainsKey("OAuth2Scopes_2") Then context.INI_OAuth2Scopes_2 = CStr(updatedValues("OAuth2Scopes_2"))
                If updatedValues.ContainsKey("OAuth2Endpoint_2") Then context.INI_OAuth2Endpoint_2 = CStr(updatedValues("OAuth2Endpoint_2"))
                If updatedValues.ContainsKey("OAuth2ATExpiry_2") Then context.INI_OAuth2ATExpiry_2 = CLng(updatedValues("OAuth2ATExpiry_2"))
                If updatedValues.ContainsKey("APIDebug") Then context.INI_APIDebug = CBool(updatedValues("APIDebug"))
                If updatedValues.ContainsKey("UsageRestrictions") Then context.INI_UsageRestrictions = CStr(updatedValues("UsageRestrictions"))
                If updatedValues.ContainsKey("LogPath") Then context.INI_LogPath = CStr(updatedValues("LogPath"))
                If updatedValues.ContainsKey("Language1") Then context.INI_Language1 = CStr(updatedValues("Language1"))
                If updatedValues.ContainsKey("Language2") Then context.INI_Language2 = CStr(updatedValues("Language2"))
                If updatedValues.ContainsKey("KeepFormat1") Then context.INI_KeepFormat1 = CBool(updatedValues("KeepFormat1"))
                If updatedValues.ContainsKey("MarkdownConvert") Then context.INI_MarkdownConvert = CBool(updatedValues("MarkdownConvert"))
                If updatedValues.ContainsKey("KeepFormat2") Then context.INI_KeepFormat2 = CBool(updatedValues("KeepFormat2"))
                If updatedValues.ContainsKey("KeepFormatCap") Then context.INI_KeepFormatCap = CInt(updatedValues("KeepFormatCap"))
                If updatedValues.ContainsKey("KeepParaFormatInline") Then context.INI_KeepParaFormatInline = CBool(updatedValues("KeepParaFormatInline"))
                If updatedValues.ContainsKey("ReplaceText1") Then context.INI_ReplaceText1 = CBool(updatedValues("ReplaceText1"))
                If updatedValues.ContainsKey("ReplaceText2") Then context.INI_ReplaceText2 = CBool(updatedValues("ReplaceText2"))
                If updatedValues.ContainsKey("ReplaceText2Override") Then context.INI_ReplaceText2Override = CStr(updatedValues("ReplaceText2Override"))
                If updatedValues.ContainsKey("DoMarkupOutlook") Then context.INI_DoMarkupOutlook = CBool(updatedValues("DoMarkupOutlook"))
                If updatedValues.ContainsKey("DoMarkupWord") Then context.INI_DoMarkupWord = CBool(updatedValues("DoMarkupWord"))
                If updatedValues.ContainsKey("SP_Translate") Then context.SP_Translate = CStr(updatedValues("SP_Translate"))
                If updatedValues.ContainsKey("SP_Correct") Then context.SP_Correct = CStr(updatedValues("SP_Correct"))
                If updatedValues.ContainsKey("SP_Improve") Then context.SP_Improve = CStr(updatedValues("SP_Improve"))
                If updatedValues.ContainsKey("SP_Explain") Then context.SP_Explain = CStr(updatedValues("SP_Explain"))
                If updatedValues.ContainsKey("SP_FindClause") Then context.SP_FindClause = CStr(updatedValues("SP_FindClause"))
                If updatedValues.ContainsKey("SP_FindClause_Clean") Then context.SP_FindClause_Clean = CStr(updatedValues("SP_FindClause_Clean"))
                If updatedValues.ContainsKey("SP_ApplyDocStyle") Then context.SP_ApplyDocStyle = CStr(updatedValues("SP_ApplyDocStyle"))
                If updatedValues.ContainsKey("SP_ApplyDocStyle_NumberingHint") Then context.SP_ApplyDocStyle_NumberingHint = CStr(updatedValues("SP_ApplyDocStyle_NumberingHint"))
                If updatedValues.ContainsKey("SP_DocCheck_Clause") Then context.SP_DocCheck_Clause = CStr(updatedValues("SP_DocCheck_Clause"))
                If updatedValues.ContainsKey("SP_DocCheck_MultiClause") Then context.SP_DocCheck_MultiClause = CStr(updatedValues("SP_DocCheck_MultiClause"))
                If updatedValues.ContainsKey("SP_DocCheck_MultiClauseSum") Then context.SP_DocCheck_MultiClauseSum = CStr(updatedValues("SP_DocCheck_MultiClauseSum"))
                If updatedValues.ContainsKey("SP_DocCheck_MultiClauseSum_Bubbles") Then context.SP_DocCheck_MultiClauseSum_Bubbles = CStr(updatedValues("SP_DocCheck_MultiClauseSum_Bubbles"))
                If updatedValues.ContainsKey("SP_SuggestTitles") Then context.SP_SuggestTitles = CStr(updatedValues("SP_SuggestTitles"))
                If updatedValues.ContainsKey("SP_Friendly") Then context.SP_Friendly = CStr(updatedValues("SP_Friendly"))
                If updatedValues.ContainsKey("SP_Convincing") Then context.SP_Convincing = CStr(updatedValues("SP_Convincing"))
                If updatedValues.ContainsKey("SP_NoFillers") Then context.SP_NoFillers = CStr(updatedValues("SP_NoFillers"))
                If updatedValues.ContainsKey("SP_Podcast") Then context.SP_Podcast = CStr(updatedValues("SP_Podcast"))
                If updatedValues.ContainsKey("SP_MyStyle_Word") Then context.SP_MyStyle_Word = CStr(updatedValues("SP_MyStyle_Word"))
                If updatedValues.ContainsKey("SP_MyStyle_Outlook") Then context.SP_MyStyle_Outlook = CStr(updatedValues("SP_MyStyle_Outlook"))
                If updatedValues.ContainsKey("SP_MyStyle_Apply") Then context.SP_MyStyle_Apply = CStr(updatedValues("SP_MyStyle_Apply"))
                If updatedValues.ContainsKey("SP_Shorten") Then context.SP_Shorten = CStr(updatedValues("SP_Shorten"))
                If updatedValues.ContainsKey("SP_Filibuster") Then context.SP_Filibuster = CStr(updatedValues("SP_Filibuster"))
                If updatedValues.ContainsKey("SP_ArgueAgainst") Then context.SP_ArgueAgainst = CStr(updatedValues("SP_ArgueAgainst"))
                If updatedValues.ContainsKey("SP_InsertClipboard") Then context.SP_InsertClipboard = CStr(updatedValues("SP_InsertClipboard"))
                If updatedValues.ContainsKey("SP_Summarize") Then context.SP_Summarize = CStr(updatedValues("SP_Summarize"))
                If updatedValues.ContainsKey("SP_Markup") Then context.SP_Markup = CStr(updatedValues("SP_Markup"))
                If updatedValues.ContainsKey("SP_MailReply") Then context.SP_MailReply = CStr(updatedValues("SP_MailReply"))
                If updatedValues.ContainsKey("SP_MailSumup") Then context.SP_MailSumup = CStr(updatedValues("SP_MailSumup"))
                If updatedValues.ContainsKey("SP_MailSumup2") Then context.SP_MailSumup2 = CStr(updatedValues("SP_MailSumup2"))
                If updatedValues.ContainsKey("SP_FreestyleText") Then context.SP_FreestyleText = CStr(updatedValues("SP_FreestyleText"))
                If updatedValues.ContainsKey("SP_FreestyleNoText") Then context.SP_FreestyleNoText = CStr(updatedValues("SP_FreestyleNoText"))
                If updatedValues.ContainsKey("SP_SwitchParty") Then context.SP_SwitchParty = CStr(updatedValues("SP_SwitchParty"))
                If updatedValues.ContainsKey("SP_Anonymize") Then context.SP_Anonymize = CStr(updatedValues("SP_Anonymize"))
                If updatedValues.ContainsKey("SP_Rename") Then context.SP_Rename = CStr(updatedValues("SP_Rename"))
                If updatedValues.ContainsKey("SP_Redact") Then context.SP_Redact = CStr(updatedValues("SP_Redact"))
                If updatedValues.ContainsKey("SP_CheckforII") Then context.SP_CheckforII = CStr(updatedValues("SP_CheckforII"))
                If updatedValues.ContainsKey("SP_Extract") Then context.SP_Extract = CStr(updatedValues("SP_Extract"))
                If updatedValues.ContainsKey("SP_ExtractSchema") Then context.SP_ExtractSchema = CStr(updatedValues("SP_ExtractSchema"))
                If updatedValues.ContainsKey("SP_MergeDateRows") Then context.SP_MergeDateRows = CStr(updatedValues("SP_MergeDateRows"))
                If updatedValues.ContainsKey("SP_ContextSearch") Then context.SP_ContextSearch = CStr(updatedValues("SP_ContextSearch"))
                If updatedValues.ContainsKey("SP_ContextSearchMulti") Then context.SP_ContextSearchMulti = CStr(updatedValues("SP_ContextSearchMulti"))
                If updatedValues.ContainsKey("SP_RangeOfCells") Then context.SP_RangeOfCells = CStr(updatedValues("SP_RangeOfCells"))
                If updatedValues.ContainsKey("SP_ParseFile") Then context.SP_ParseFile = CStr(updatedValues("SP_ParseFile"))
                If updatedValues.ContainsKey("SP_Ignore") Then context.SP_Ignore = CStr(updatedValues("SP_Ignore"))
                If updatedValues.ContainsKey("SP_WriteNeatly") Then context.SP_WriteNeatly = CStr(updatedValues("SP_WriteNeatly"))
                If updatedValues.ContainsKey("SP_Add_KeepFormulasIntact") Then context.SP_Add_KeepFormulasIntact = CStr(updatedValues("SP_Add_KeepFormulasIntact"))
                If updatedValues.ContainsKey("SP_Add_KeepHTMLIntact") Then context.SP_Add_KeepHTMLIntact = CStr(updatedValues("SP_Add_KeepHTMLIntact"))
                If updatedValues.ContainsKey("SP_Add_KeepInlineIntact") Then context.SP_Add_KeepInlineIntact = CStr(updatedValues("SP_Add_KeepInlineIntact"))
                If updatedValues.ContainsKey("SP_Add_Bubbles") Then context.SP_Add_Bubbles = CStr(updatedValues("SP_Add_Bubbles"))
                If updatedValues.ContainsKey("SP_Add_BubblesReply") Then context.SP_Add_BubblesReply = CStr(updatedValues("SP_Add_BubblesReply"))
                If updatedValues.ContainsKey("SP_Add_BubblesExtract") Then context.SP_Add_BubblesExtract = CStr(updatedValues("SP_Add_BubblesExtract"))
                If updatedValues.ContainsKey("SP_Add_Bubbles_Format") Then context.SP_Add_Bubbles_Format = CStr(updatedValues("SP_Add_Bubbles_Format"))
                If updatedValues.ContainsKey("SP_Add_Batch") Then context.SP_Add_Batch = CStr(updatedValues("SP_Add_Batch"))
                If updatedValues.ContainsKey("SP_Add_Tooling") Then context.SP_Add_Tooling = CStr(updatedValues("SP_Add_Tooling"))
                If updatedValues.ContainsKey("SP_Add_Slides") Then context.SP_Add_Slides = CStr(updatedValues("SP_Add_Slides"))
                If updatedValues.ContainsKey("SP_BubblesExcel") Then context.SP_BubblesExcel = CStr(updatedValues("SP_BubblesExcel"))
                If updatedValues.ContainsKey("SP_Add_Revisions") Then context.SP_Add_Revisions = CStr(updatedValues("SP_Add_Revisions"))
                If updatedValues.ContainsKey("SP_MarkupRegex") Then context.SP_MarkupRegex = CStr(updatedValues("SP_MarkupRegex"))
                If updatedValues.ContainsKey("SP_ChatWord") Then context.SP_ChatWord = CStr(updatedValues("SP_ChatWord"))
                If updatedValues.ContainsKey("SP_HelpMe") Then context.SP_HelpMe = CStr(updatedValues("SP_HelpMe"))
                If updatedValues.ContainsKey("SP_Chat") Then context.SP_Chat = CStr(updatedValues("SP_Chat"))
                If updatedValues.ContainsKey("SP_Add_ChatWord_Commands") Then context.SP_Add_ChatWord_Commands = CStr(updatedValues("SP_Add_ChatWord_Commands"))
                If updatedValues.ContainsKey("SP_Add_Chat_NoCommands") Then context.SP_Add_Chat_NoCommands = CStr(updatedValues("SP_Add_Chat_NoCommands"))
                If updatedValues.ContainsKey("SP_ChatExcel") Then context.SP_ChatExcel = CStr(updatedValues("SP_ChatExcel"))
                If updatedValues.ContainsKey("SP_Add_ChatExcel_Commands") Then context.SP_Add_ChatExcel_Commands = CStr(updatedValues("SP_Add_ChatExcel_Commands"))
                If updatedValues.ContainsKey("SP_Add_MergePrompt") Then context.SP_Add_MergePrompt = CStr(updatedValues("SP_Add_MergePrompt"))
                If updatedValues.ContainsKey("SP_FindPrompts") Then context.SP_FindPrompts = CStr(updatedValues("SP_FindPrompts"))
                If updatedValues.ContainsKey("SP_MergePrompt") Then context.SP_MergePrompt = CStr(updatedValues("SP_MergePrompt"))
                If updatedValues.ContainsKey("SP_MergePrompt2") Then context.SP_MergePrompt2 = CStr(updatedValues("SP_MergePrompt2"))
                If updatedValues.ContainsKey("ISearch") Then context.INI_ISearch = CBool(updatedValues("ISearch"))
                If updatedValues.ContainsKey("ISearch_Approve") Then context.INI_ISearch_Approve = CBool(updatedValues("ISearch_Approve"))
                If updatedValues.ContainsKey("ISearch_URL") Then context.INI_ISearch_URL = CStr(updatedValues("ISearch_URL"))
                If updatedValues.ContainsKey("ISearch_ResponseMask1") Then context.INI_ISearch_ResponseMask1 = CStr(updatedValues("ISearch_ResponseMask1"))
                If updatedValues.ContainsKey("ISearch_ResponseMask2") Then context.INI_ISearch_ResponseMask2 = CStr(updatedValues("ISearch_ResponseMask2"))
                If updatedValues.ContainsKey("ISearch_Name") Then context.INI_ISearch_Name = CStr(updatedValues("ISearch_Name"))
                If updatedValues.ContainsKey("ISearch_Tries") Then context.INI_ISearch_Tries = CInt(updatedValues("ISearch_Tries"))
                If updatedValues.ContainsKey("ISearch_Results") Then context.INI_ISearch_Results = CInt(updatedValues("ISearch_Results"))
                If updatedValues.ContainsKey("ISearch_MaxDepth") Then context.INI_ISearch_MaxDepth = CInt(updatedValues("ISearch_MaxDepth"))
                If updatedValues.ContainsKey("ISearch_Timeout") Then context.INI_ISearch_Timeout = CLng(updatedValues("ISearch_Timeout"))
                If updatedValues.ContainsKey("ISearch_SearchTerm_SP") Then context.INI_ISearch_SearchTerm_SP = CStr(updatedValues("ISearch_SearchTerm_SP"))
                If updatedValues.ContainsKey("ISearch_Apply_SP") Then context.INI_ISearch_Apply_SP = CStr(updatedValues("ISearch_Apply_SP"))
                If updatedValues.ContainsKey("ISearch_Apply_SP_Markup") Then context.INI_ISearch_Apply_SP_Markup = CStr(updatedValues("ISearch_Apply_SP_Markup"))
                If updatedValues.ContainsKey("Lib") Then context.INI_Lib = CBool(updatedValues("Lib"))
                If updatedValues.ContainsKey("Lib_File") Then context.INI_Lib_File = CStr(updatedValues("Lib_File"))
                If updatedValues.ContainsKey("Lib_Timeout") Then context.INI_Lib_Timeout = CLng(updatedValues("Lib_Timeout"))
                If updatedValues.ContainsKey("Lib_Find_SP") Then context.INI_Lib_Find_SP = CStr(updatedValues("Lib_Find_SP"))
                If updatedValues.ContainsKey("Lib_Apply_SP") Then context.INI_Lib_Apply_SP = CStr(updatedValues("Lib_Apply_SP"))
                If updatedValues.ContainsKey("Lib_Apply_SP_Markup") Then context.INI_Lib_Apply_SP_Markup = CStr(updatedValues("Lib_Apply_SP_Markup"))
                If updatedValues.ContainsKey("MarkupMethodHelper") Then context.INI_MarkupMethodHelper = CInt(updatedValues("MarkupMethodHelper"))
                If updatedValues.ContainsKey("MarkupMethodWord") Then context.INI_MarkupMethodWord = CInt(updatedValues("MarkupMethodWord"))
                If updatedValues.ContainsKey("MarkupMethodWordOverride") Then context.INI_MarkupMethodWordOverride = CStr(updatedValues("MarkupMethodWordOverride"))
                If updatedValues.ContainsKey("MarkupMethodOutlookOverride") Then context.INI_MarkupMethodOutlookOverride = CStr(updatedValues("MarkupMethodOutlookOverride"))
                If updatedValues.ContainsKey("ShortcutsWordExcel") Then context.INI_ShortcutsWordExcel = CStr(updatedValues("ShortcutsWordExcel"))
                If updatedValues.ContainsKey("ContextMenu") Then context.INI_ContextMenu = CBool(updatedValues("ContextMenu"))
                If updatedValues.ContainsKey("UpdateCheckInterval") Then context.INI_UpdateCheckInterval = CInt(updatedValues("UpdateCheckInterval"))
                If updatedValues.ContainsKey("UpdatePath") Then context.INI_UpdatePath = CStr(updatedValues("UpdatePath"))
                If updatedValues.ContainsKey("HelpMeInkyPath") Then context.INI_HelpMeInkyPath = CStr(updatedValues("HelpMeInkyPath"))
                If updatedValues.ContainsKey("DiscussInkyPath") Then context.INI_DiscussInkyPath = CStr(updatedValues("DiscussInkyPath"))
                If updatedValues.ContainsKey("DiscussInkyPathLocal") Then context.INI_DiscussInkyPathLocal = CStr(updatedValues("DiscussInkyPathLocal"))
                If updatedValues.ContainsKey("RedactionInstructionsPath") Then context.INI_RedactionInstructionsPath = CStr(updatedValues("RedactionInstructionsPath"))
                If updatedValues.ContainsKey("RedactionInstructionsPathLocal") Then context.INI_RedactionInstructionsPathLocal = CStr(updatedValues("RedactionInstructionsPathLocal"))
                If updatedValues.ContainsKey("ExtractorPath") Then context.INI_ExtractorPath = CStr(updatedValues("ExtractorPath"))
                If updatedValues.ContainsKey("ExtractorPathLocal") Then context.INI_ExtractorPathLocal = CStr(updatedValues("ExtractorPathLocal"))
                If updatedValues.ContainsKey("RenameLibPath") Then context.INI_RenameLibPath = CStr(updatedValues("RenameLibPath"))
                If updatedValues.ContainsKey("RenameLibPathLocal") Then context.INI_RenameLibPathLocal = CStr(updatedValues("RenameLibPathLocal"))
                If updatedValues.ContainsKey("SpeechModelPath") Then context.INI_SpeechModelPath = CStr(updatedValues("SpeechModelPath"))
                If updatedValues.ContainsKey("LocalModelPath") Then context.INI_LocalModelPath = CStr(updatedValues("LocalModelPath"))
                If updatedValues.ContainsKey("TTSEndpoint") Then context.INI_TTSEndpoint = CStr(updatedValues("TTSEndpoint"))
                If updatedValues.ContainsKey("PromptLib") Then context.INI_PromptLibPath = CStr(updatedValues("PromptLib"))
                If updatedValues.ContainsKey("PromptLibLocal") Then context.INI_PromptLibPathLocal = CStr(updatedValues("PromptLibLocal"))
                If updatedValues.ContainsKey("MyStylePath") Then context.INI_MyStylePath = CStr(updatedValues("MyStylePath"))
                If updatedValues.ContainsKey("AlternateModelPath") Then context.INI_AlternateModelPath = CStr(updatedValues("AlternateModelPath"))
                If updatedValues.ContainsKey("SpecialServicePath") Then context.INI_SpecialServicePath = CStr(updatedValues("SpecialServicePath"))
                If updatedValues.ContainsKey("FindClausePath") Then context.INI_FindClausePath = CStr(updatedValues("FindClausePath"))
                If updatedValues.ContainsKey("FindClausePathLocal") Then context.INI_FindClausePathLocal = CStr(updatedValues("FindClausePathLocal"))
                If updatedValues.ContainsKey("WebAgentPath") Then context.INI_WebAgentPath = CStr(updatedValues("WebAgentPath"))
                If updatedValues.ContainsKey("WebAgentPathLocal") Then context.INI_WebAgentPathLocal = CStr(updatedValues("WebAgentPathLocal"))
                If updatedValues.ContainsKey("DocCheckPath") Then context.INI_DocCheckPath = CStr(updatedValues("DocCheckPath"))
                If updatedValues.ContainsKey("DocCheckPathLocal") Then context.INI_DocCheckPathLocal = CStr(updatedValues("DocCheckPathLocal"))
                If updatedValues.ContainsKey("DocStylePath") Then context.INI_DocStylePath = CStr(updatedValues("DocStylePath"))
                If updatedValues.ContainsKey("DocStylePathLocal") Then context.INI_DocStylePathLocal = CStr(updatedValues("DocStylePathLocal"))
                If updatedValues.ContainsKey("PromptLib_Transcript") Then context.INI_PromptLibPath_Transcript = CStr(updatedValues("PromptLib_Transcript"))
                If updatedValues.ContainsKey("NoHelperDownload") Then context.INI_NoHelperDownload = CBool(updatedValues("NoHelperDownload"))
                If updatedValues.ContainsKey("ToolingLogWindow") Then context.INI_ToolingLogWindow = CBool(updatedValues("ToolingLogWindow"))
                If updatedValues.ContainsKey("ToolingDryRun") Then context.INI_ToolingDryRun = CBool(updatedValues("ToolingDryRun"))
                If updatedValues.ContainsKey("ToolingMaximumIterations") Then context.INI_ToolingMaximumIterations = CInt(updatedValues("ToolingMaximumIterations"))
                If updatedValues.ContainsKey("UpdateIni") Then context.INI_UpdateIni = CBool(updatedValues("UpdateIni"))
                If updatedValues.ContainsKey("UpdateIniAllowRemote") Then context.INI_UpdateIniAllowRemote = CBool(updatedValues("UpdateIniAllowRemote"))
                If updatedValues.ContainsKey("UpdateIniNoSignature") Then context.INI_UpdateIniNoSignature = CBool(updatedValues("UpdateIniNoSignature"))
                If updatedValues.ContainsKey("UpdateSource") Then context.INI_UpdateSource = CStr(updatedValues("UpdateSource"))
                If updatedValues.ContainsKey("UpdateClients") Then context.INI_UpdateClients = CStr(updatedValues("UpdateClients"))
                If updatedValues.ContainsKey("UpdateIniIgnoreOverride") Then context.INI_UpdateIniIgnoreOverride = CStr(updatedValues("UpdateIniIgnoreOverride"))
                If updatedValues.ContainsKey("UpdateIniSilentMode") Then context.INI_UpdateIniSilentMode = CInt(updatedValues("UpdateIniSilentMode"))
                If updatedValues.ContainsKey("UpdateIniSilentLog") Then context.INI_UpdateIniSilentLog = CBool(updatedValues("UpdateIniSilentLog"))

                ' Call UpdateAppConfig after all updates
                UpdateAppConfig(context)
            End If
        End Sub

        ''' <summary>
        ''' Shows a modal "About" dialog for the current application instance.
        ''' The dialog displays application identifiers, license status information, and additional static text,
        ''' and provides buttons to view third-party license text and to reset stored license information.
        ''' </summary>
        ''' <param name="owner">Owner window used for theming (background color) and modal display.</param>
        ''' <param name="context">Shared context providing the current application identifier (<c>RDV</c>).</param>

        Public Shared Sub ShowAboutWindow(owner As System.Windows.Forms.Form, context As ISharedContext)
            ' Example of using the same font and appearance as ShowWindowsSettings
            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)

            Dim baseWidth As Integer = 450
            Dim formWidth As Integer = CInt(baseWidth * 1.3)

            ' Calculate height based on text content
            Dim ExpireText As String = $"{vbCrLf}{vbCrLf}(your {If(String.IsNullOrEmpty(LicenseStatus), "(undefined license type)", LicenseStatus)} for {LicenseUsers} user(s) expires on {LicensedTill.ToString("dd-MMM-yyyy")})"
            Dim testRichTextBox As New System.Windows.Forms.RichTextBox() With {
                            .Font = standardFont,
                            .Text = $"{AN}{vbCrLf}{context.RDV}{ExpireText}{vbCrLf}{vbCrLf}By David Rosenthal & Team{vbCrLf}{vbCrLf}{CopyrightNotice}{vbCrLf}{vbCrLf}All rights reserved.{vbCrLf}{vbCrLf}{AN4}{vbCrLf}{vbCrLf}Local Chat: {AN7}"
                        }
            Dim graphics As System.Drawing.Graphics = testRichTextBox.CreateGraphics()
            Dim textSize As System.Drawing.SizeF = graphics.MeasureString(testRichTextBox.Text, standardFont, formWidth - 40)
            graphics.Dispose()
            testRichTextBox.Dispose()

            ' Calculate required height: logo + text + buttons + margins
            Dim logoSize As Integer = 120
            Dim buttonHeight As Integer = 30
            Dim buttonSpacing As Integer = 10
            Dim margins As Integer = 80 ' top/bottom margins and spacing

            Dim requiredHeight As Integer = logoSize + CInt(textSize.Height) + (buttonHeight * 3) + (buttonSpacing * 4) + margins

            ' Clamp to screen working area if needed
            Dim workArea = Screen.FromPoint(Cursor.Position).WorkingArea
            Dim maxHeight As Integer = CInt(workArea.Height * 0.85)
            Dim formHeight As Integer = Math.Min(requiredHeight, maxHeight)

            ' Create the form
            Dim aboutForm As New System.Windows.Forms.Form() With {
                        .FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog,
                        .StartPosition = System.Windows.Forms.FormStartPosition.CenterParent,
                        .ClientSize = New System.Drawing.Size(formWidth, formHeight),
                        .BackColor = owner.BackColor,
                        .Font = standardFont,
                        .MaximizeBox = False,
                        .MinimizeBox = False,
                        .ControlBox = False,
                        .ShowInTaskbar = False
                    }

            ' Add a logo
            Dim logo As New System.Windows.Forms.PictureBox() With {
                        .Image = My.Resources.Red_Ink_Logo_Large,
                        .SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom,
                        .Size = New System.Drawing.Size(logoSize, logoSize),
                        .Location = New System.Drawing.Point((formWidth - logoSize) \ 2, 20)
                    }
            aboutForm.Controls.Add(logo)

            ' Add a RichTextBox for text
            Dim aboutTextBox As New System.Windows.Forms.RichTextBox() With {
                        .ReadOnly = True,
                        .BorderStyle = System.Windows.Forms.BorderStyle.None,
                        .BackColor = owner.BackColor,
                        .Font = standardFont,
                        .DetectUrls = True,
                        .ScrollBars = RichTextBoxScrollBars.None
                    }

            Dim topOffset As Integer = logo.Bottom + 10
            Dim bottomPadding As Integer = (buttonHeight * 3) + (buttonSpacing * 4) + 20
            Dim availableHeight As Integer = formHeight - topOffset - bottomPadding
            aboutTextBox.Size = New System.Drawing.Size(formWidth - 40, availableHeight)
            aboutTextBox.Location = New System.Drawing.Point(20, topOffset)
            aboutForm.Controls.Add(aboutTextBox)

            Dim aboutContent As String =
        $"{AN}<P>{context.RDV}{ExpireText}<P><P>By David Rosenthal & Team<P><P>{CopyrightNotice}<P><P>All rights reserved.<P><P>{AN4}<P><P>Local Chat: {AN7}"

            ' Replace <P> with vbCrLf
            Dim plainText As New System.Text.StringBuilder()

            While aboutContent.Contains("<P>")
                Dim index = aboutContent.IndexOf("<P>")
                plainText.Append(aboutContent.Substring(0, index))
                plainText.Append(vbCrLf)
                aboutContent = aboutContent.Substring(index + 3)
            End While
            plainText.Append(aboutContent)

            ' Set the text and apply formatting
            aboutTextBox.Text = plainText.ToString()

            ' Center the text
            aboutTextBox.SelectAll()
            aboutTextBox.SelectionAlignment = HorizontalAlignment.Center
            aboutTextBox.DeselectAll()

            ' Hide the blinking cursor
            aboutTextBox.SelectionStart = aboutTextBox.Text.Length
            aboutTextBox.SelectionLength = 0
            aboutTextBox.ScrollToCaret() ' Ensures the caret is out of visible range

            ' Add a handler for link clicks
            AddHandler aboutTextBox.LinkClicked,
        Sub(sender, e)
            Try
                Process.Start(New ProcessStartInfo(e.LinkText) With {.UseShellExecute = True})
            Catch ex As System.Exception
                MessageBox.Show("Error in ShowAboutWindow - unable to open the link.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

            ' Measure button text widths to size buttons appropriately
            Dim licenseButtonText As String = "3rd Party Software Used"
            Dim resetButtonText As String = "Update License Info"
            Dim licenseTextSize As Size = TextRenderer.MeasureText(licenseButtonText, standardFont)
            Dim resetTextSize As Size = TextRenderer.MeasureText(resetButtonText, standardFont)
            Dim buttonPadding As Integer = 20

            Dim licenseButtonWidth As Integer = licenseTextSize.Width + buttonPadding
            Dim resetButtonWidth As Integer = resetTextSize.Width + buttonPadding
            Dim stackedButtonWidth As Integer = Math.Max(licenseButtonWidth, resetButtonWidth)
            Dim buttonsLeft As Integer = (formWidth - stackedButtonWidth) \ 2

            ' Add a "License" button
            Dim licenseButton As New System.Windows.Forms.Button() With {
                        .Text = licenseButtonText,
                        .Size = New System.Drawing.Size(stackedButtonWidth, buttonHeight),
                        .Location = New System.Drawing.Point(buttonsLeft, aboutTextBox.Bottom + 10)
                    }
            AddHandler licenseButton.Click, Sub(sender, e) ShowRTFCustomMessageBox(ConvertMarkupToRTF(LicenseText), AN)
            aboutForm.Controls.Add(licenseButton)

            ' Add a "Reset License" button
            Dim resetLicenseButton As New System.Windows.Forms.Button() With {
                        .Text = resetButtonText,
                        .Size = New System.Drawing.Size(stackedButtonWidth, buttonHeight),
                        .Location = New System.Drawing.Point(buttonsLeft, licenseButton.Bottom + buttonSpacing),
                        .Enabled = Not LicenseFromConfig AndAlso Not IsBetaVersion() AndAlso Not LicenseStatus = "Beta Test License"
                    }
            AddHandler resetLicenseButton.Click, Sub(sender, e)
                                                     Try
                                                         ' Reset license information in My.Settings
                                                         My.Settings.LicenseStatus = ""
                                                         My.Settings.LicenseUsers = 1
                                                         My.Settings.LicensedTill = Date.MinValue
                                                         My.Settings.Save()

                                                         ' Reset global license variables
                                                         LicenseStatus = ""
                                                         LicenseUsers = 1
                                                         LicensedTill = Date.MinValue

                                                         ' Close the current About window
                                                         aboutForm.Close()

                                                         ' Show the license configuration form
                                                         ShowLicenseEntryForm(context)

                                                         ' Re-show the About window with updated info
                                                         ShowAboutWindow(owner, context)
                                                     Catch ex As Exception
                                                         ShowCustomMessageBox($"Error resetting license: {ex.Message}", AN)
                                                     End Try
                                                 End Sub
            aboutForm.Controls.Add(resetLicenseButton)

            ' Add an OK button
            Dim okButton As New System.Windows.Forms.Button() With {
                        .Text = "OK",
                        .Size = New System.Drawing.Size(80, buttonHeight),
                        .Location = New System.Drawing.Point((formWidth - 80) \ 2, resetLicenseButton.Bottom + buttonSpacing)
                    }
            AddHandler okButton.Click, Sub(sender, e) aboutForm.Close()
            aboutForm.Controls.Add(okButton)

            ' Adjust form height to fit OK button if needed
            Dim finalHeight As Integer = okButton.Bottom + 20
            If finalHeight > formHeight Then
                aboutForm.ClientSize = New System.Drawing.Size(formWidth, Math.Min(finalHeight, maxHeight))
            End If

            ' Show the form
            aboutForm.ShowDialog(owner)
        End Sub


    End Class

End Namespace
