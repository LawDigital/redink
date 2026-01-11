' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SharedMethods.AlternateModels.vb
' Purpose: Loads, filters, and applies alternative LLM/model configurations ("alternate models")
'          from an INI file, and transfers configuration values between `ISharedContext` and
'          `ModelConfig` for UI-driven selection and task-driven switching.
'
' Architecture:
'  - INI Parsing: `LoadAlternativeModels` reads an INI file, parses each section into a
'    `Dictionary(Of String, String)`, and materializes each section as a `ModelConfig`.
'  - ModelConfig Materialization: `CreateModelConfigFromDict` maps INI keys to `ModelConfig`
'    properties, initializes OAuth2-related fields, and resolves key material using `RealAPIKeyMC`.
'  - Context Snapshot/Restore: `GetCurrentConfig` snapshots the active configuration from
'    `ISharedContext`; `ApplyModelConfig` writes a `ModelConfig` back into `ISharedContext`.
'  - Filtering Rules: `ProcessModelSection` excludes deprecated entries and optionally excludes
'    `ToolOnly` entries; can also filter to only tool-capable definitions.
'  - UI Integration: `ShowModelSelection` uses `ModelSelectorForm` (single selection) and
'    `ShowMultipleModelSelection` uses `MultiModelSelectorForm` (multi-selection).
'  - Task-Based Selection: `GetSpecialTaskModel` scans INI sections for a given key (task flag)
'    set to a truthy value and applies the first matching section.
'
' Notes:
'  - The INI key `APIKeyEncrypted` is parsed into `ModelConfig.APIEncrypted`.
'  - Deprecated can be configured via both `Deprecated` and `NotAvailable`.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary.SharedContext
Imports System.IO

Namespace SharedLibrary

    Partial Public Class SharedMethods

        ''' <summary>
        ''' Creates a <see cref="ModelConfig"/> instance from an INI-section dictionary.
        ''' Populates properties from known keys, initializes OAuth2-related fields, and resolves
        ''' API key material using <c>RealAPIKeyMC</c>.
        ''' </summary>
        ''' <param name="configDict">Key/value pairs parsed from an INI section.</param>
        ''' <param name="context">Shared context used for defaults and API key resolution.</param>
        ''' <param name="Description">Model display description (typically taken from the INI section name).</param>
        ''' <returns>A populated <see cref="ModelConfig"/> instance.</returns>
        Public Shared Function CreateModelConfigFromDict(ByVal configDict As Dictionary(Of String, String), context As ISharedContext, Description As String) As ModelConfig
            Dim mc As New ModelConfig()
            Try
                mc.APIKey = If(configDict.ContainsKey("APIKey"), configDict("APIKey"), "")
                mc.Endpoint = If(configDict.ContainsKey("Endpoint"), configDict("Endpoint"), "")
                mc.HeaderA = If(configDict.ContainsKey("HeaderA"), configDict("HeaderA"), "")
                mc.HeaderB = If(configDict.ContainsKey("HeaderB"), configDict("HeaderB"), "")
                mc.Response = If(configDict.ContainsKey("Response"), configDict("Response"), "")
                mc.APICall = If(configDict.ContainsKey("APICall"), configDict("APICall"), "")
                mc.APICall_Object = If(configDict.ContainsKey("APICall_Object"), configDict("APICall_Object"), "")
                mc.Timeout = If(configDict.ContainsKey("Timeout"), CLng(configDict("Timeout")), 0)
                mc.MaxOutputToken = If(configDict.ContainsKey("MaxOutputToken"), CInt(configDict("MaxOutputToken")), 0)
                mc.Temperature = If(configDict.ContainsKey("Temperature"), configDict("Temperature"), "")
                mc.Model = If(configDict.ContainsKey("Model"), configDict("Model"), "")
                mc.APIEncrypted = ParseBoolean(configDict, "APIKeyEncrypted")
                mc.APIKeyPrefix = If(configDict.ContainsKey("APIKeyPrefix"), configDict("APIKeyPrefix"), "")
                mc.OAuth2 = ParseBoolean(configDict, "OAuth2")
                mc.OAuth2ClientMail = If(configDict.ContainsKey("OAuth2ClientMail"), configDict("OAuth2ClientMail"), "")
                mc.OAuth2Scopes = If(configDict.ContainsKey("OAuth2Scopes"), configDict("OAuth2Scopes"), "")
                mc.OAuth2Endpoint = If(configDict.ContainsKey("OAuth2Endpoint"), configDict("OAuth2Endpoint"), "")
                mc.OAuth2ATExpiry = If(configDict.ContainsKey("OAuth2ATExpiry"), CLng(configDict("OAuth2ATExpiry")), 3600)
                mc.Parameter1 = If(configDict.ContainsKey("Parameter1"), configDict("Parameter1"), "")
                mc.Parameter2 = If(configDict.ContainsKey("Parameter2"), configDict("Parameter2"), "")
                mc.Parameter3 = If(configDict.ContainsKey("Parameter3"), configDict("Parameter3"), "")
                mc.Parameter4 = If(configDict.ContainsKey("Parameter4"), configDict("Parameter4"), "")
                mc.MergePrompt = If(configDict.ContainsKey("MergePrompt"), configDict("MergePrompt"), context.SP_MergePrompt)
                mc.QueryPrompt = If(configDict.ContainsKey("QueryPrompt"), configDict("QueryPrompt"), "")
                mc.ModelDescription = Description

                mc.APIKeyBack = mc.APIKey

                ' OAuth2-related runtime fields (default values for later refresh/usage).
                mc.TokenExpiry = Microsoft.VisualBasic.DateAndTime.DateAdd(Microsoft.VisualBasic.DateInterval.Year, -1, DateTime.Now)
                mc.DecodedAPI = ""

                ' API key resolution:
                ' - OAuth2: use RealAPIKeyMC(...) as returned and strip literal "\n" sequences.
                ' - Non-OAuth2: store decoded key material in DecodedAPI; APIKey remains the original configured value.
                If mc.OAuth2 Then
                    mc.APIKey = Trim(Replace(RealAPIKeyMC(mc.APIKey, True, mc, context), "\n", ""))
                Else
                    mc.DecodedAPI = RealAPIKeyMC(mc.APIKey, False, mc, context)
                End If

                ' === TOOLING PROPERTIES ===
                ' Tooling configuration used by LLM/tool-call related features.
                mc.Tool = ParseBoolean(configDict, "Tool")
                mc.ToolOnly = ParseBoolean(configDict, "ToolOnly")
                mc.Deprecated = ParseBoolean(configDict, "Off") OrElse ParseBoolean(configDict, "Deprecated")

                mc.APICall_ToolInstructions = If(configDict.ContainsKey("APICall_ToolInstructions"), configDict("APICall_ToolInstructions"), "")
                mc.APICall_ToolInstructions_Template = If(configDict.ContainsKey("APICall_ToolInstructions_Template"), configDict("APICall_ToolInstructions_Template"), "")
                mc.APICall_ToolResponses = If(configDict.ContainsKey("APICall_ToolResponses"), configDict("APICall_ToolResponses"), "")
                mc.APICall_ToolResponses_Template = If(configDict.ContainsKey("APICall_ToolResponses_Template"), configDict("APICall_ToolResponses_Template"), "")
                mc.APICall_ToolCallPart_Template = If(configDict.ContainsKey("APICall_ToolCallPart_Template"), configDict("APICall_ToolCallPart_Template"), "")
                mc.ToolCallDetectionPattern = If(configDict.ContainsKey("ToolCallDetectionPattern"), configDict("ToolCallDetectionPattern"), "")
                mc.ToolCallExtractionMap = If(configDict.ContainsKey("ToolCallExtractionMap"), configDict("ToolCallExtractionMap"), "")

                mc.ToolName = If(configDict.ContainsKey("ToolName"), configDict("ToolName"), "")
                mc.ToolInstructionsPrompt = If(configDict.ContainsKey("ToolInstructionsPrompt"), configDict("ToolInstructionsPrompt"), "")
                mc.ToolDefinition = If(configDict.ContainsKey("ToolDefinition"), configDict("ToolDefinition"), "")
                mc.ToolAPICall = If(configDict.ContainsKey("ToolAPICall"), configDict("ToolAPICall"), "")
                mc.ToolErrorHandling = If(configDict.ContainsKey("ToolErrorHandling"), configDict("ToolErrorHandling"), "skip")
                mc.ToolParameterDefaults = If(configDict.ContainsKey("ToolParameterDefaults"), configDict("ToolParameterDefaults"), "")

                If configDict.ContainsKey("ToolPriority") Then
                    Dim priorityVal As Integer
                    If Integer.TryParse(configDict("ToolPriority"), priorityVal) Then
                        mc.ToolPriority = priorityVal
                    Else
                        mc.ToolPriority = 100
                    End If
                Else
                    mc.ToolPriority = 100
                End If

            Catch ex As System.Exception
                MessageBox.Show("Error in CreateModelConfigFromDict: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

            Return mc
        End Function

        ''' <summary>
        ''' Creates a <see cref="ModelConfig"/> snapshot from values currently stored in <see cref="ISharedContext"/>.
        ''' </summary>
        ''' <param name="context">Source context containing the active configuration values.</param>
        ''' <returns>A populated <see cref="ModelConfig"/> instance.</returns>
        Public Shared Function GetCurrentConfig(ByVal context As ISharedContext) As ModelConfig
            Dim mc As New ModelConfig()
            Try
                ' Copy active context values into a ModelConfig snapshot.
                mc.APIKey = If(String.IsNullOrEmpty(context.INI_APIKey_2), "", context.INI_APIKey_2)
                mc.APIKeyBack = If(String.IsNullOrEmpty(context.INI_APIKeyBack_2), "", context.INI_APIKeyBack_2)
                mc.Endpoint = If(String.IsNullOrEmpty(context.INI_Endpoint_2), "", context.INI_Endpoint_2)
                mc.HeaderA = If(String.IsNullOrEmpty(context.INI_HeaderA_2), "", context.INI_HeaderA_2)
                mc.HeaderB = If(String.IsNullOrEmpty(context.INI_HeaderB_2), "", context.INI_HeaderB_2)
                mc.Response = If(String.IsNullOrEmpty(context.INI_Response_2), "", context.INI_Response_2)
                mc.Anon = If(String.IsNullOrEmpty(context.INI_Anon_2), "", context.INI_Anon_2)
                mc.TokenCount = If(String.IsNullOrEmpty(context.INI_TokenCount_2), "", context.INI_TokenCount_2)
                mc.APICall = If(String.IsNullOrEmpty(context.INI_APICall_2), "", context.INI_APICall_2)
                mc.APICall_Object = If(String.IsNullOrEmpty(context.INI_APICall_Object_2), "", context.INI_APICall_Object_2)
                mc.Timeout = context.INI_Timeout_2
                mc.MaxOutputToken = context.INI_MaxOutputToken_2
                mc.Temperature = If(String.IsNullOrEmpty(context.INI_Temperature_2), "", context.INI_Temperature_2)
                mc.Model = If(String.IsNullOrEmpty(context.INI_Model_2), "", context.INI_Model_2)
                mc.APIEncrypted = context.INI_APIEncrypted_2
                mc.APIKeyPrefix = If(String.IsNullOrEmpty(context.INI_APIKeyPrefix_2), "", context.INI_APIKeyPrefix_2)
                mc.OAuth2 = context.INI_OAuth2_2
                mc.OAuth2ClientMail = If(String.IsNullOrEmpty(context.INI_OAuth2ClientMail_2), "", context.INI_OAuth2ClientMail_2)
                mc.OAuth2Scopes = If(String.IsNullOrEmpty(context.INI_OAuth2Scopes_2), "", context.INI_OAuth2Scopes_2)
                mc.OAuth2Endpoint = If(String.IsNullOrEmpty(context.INI_OAuth2Endpoint_2), "", context.INI_OAuth2Endpoint_2)
                mc.OAuth2ATExpiry = context.INI_OAuth2ATExpiry_2
                mc.MergePrompt = If(String.IsNullOrEmpty(context.SP_MergePrompt), "", context.SP_MergePrompt)
                mc.DecodedAPI = context.DecodedAPI_2
                mc.TokenExpiry = context.TokenExpiry_2

                ' === CAPTURE TOOLING PROPERTIES ===
                mc.APICall_ToolInstructions = context.INI_APICall_ToolInstructions_2
                mc.APICall_ToolInstructions_Template = context.INI_APICall_ToolInstructions_Template_2
                mc.APICall_ToolResponses = context.INI_APICall_ToolResponses_2
                mc.APICall_ToolResponses_Template = context.INI_APICall_ToolResponses_Template_2
                mc.APICall_ToolCallPart_Template = context.INI_APICall_ToolCallPart_Template_2
                mc.ToolCallDetectionPattern = context.INI_ToolCallDetectionPattern_2
                mc.ToolCallExtractionMap = context.INI_ToolCallExtractionMap_2

            Catch ex As System.Exception
                MessageBox.Show("Error in GetCurrentConfig: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            Return mc
        End Function

        ''' <summary>
        ''' Applies the specified <see cref="ModelConfig"/> to the provided <see cref="ISharedContext"/>.
        ''' </summary>
        ''' <param name="context">Target context to receive the configuration values.</param>
        ''' <param name="config">Source configuration values to apply.</param>
        ''' <param name="ErrorFlag">
        ''' Optional error-suppression flag. When <c>True</c>, repeated UI error reporting is suppressed by this method.
        ''' </param>
        Public Shared Sub ApplyModelConfig(ByVal context As ISharedContext, ByVal config As ModelConfig, Optional ByRef ErrorFlag As Boolean = False)
            Try
                context.INI_APIKey_2 = If(Not String.IsNullOrEmpty(config.APIKey), config.APIKey, "")
                context.INI_APIKeyBack_2 = If(Not String.IsNullOrEmpty(config.APIKeyBack), config.APIKeyBack, "")
                context.INI_Endpoint_2 = If(Not String.IsNullOrEmpty(config.Endpoint), config.Endpoint, "")
                context.INI_HeaderA_2 = If(Not String.IsNullOrEmpty(config.HeaderA), config.HeaderA, "")
                context.INI_HeaderB_2 = If(Not String.IsNullOrEmpty(config.HeaderB), config.HeaderB, "")
                context.INI_Response_2 = If(Not String.IsNullOrEmpty(config.Response), config.Response, "")
                context.INI_Anon_2 = If(Not String.IsNullOrEmpty(config.Anon), config.Anon, "")
                context.INI_TokenCount_2 = If(Not String.IsNullOrEmpty(config.TokenCount), config.TokenCount, "")
                context.INI_APICall_2 = If(Not String.IsNullOrEmpty(config.APICall), config.APICall, "")
                context.INI_APICall_Object_2 = If(Not String.IsNullOrEmpty(config.APICall_Object), config.APICall_Object, "")
                context.INI_Timeout_2 = If(config.Timeout <> 0, config.Timeout, 0)
                context.INI_MaxOutputToken_2 = If(config.MaxOutputToken <> 0, config.MaxOutputToken, 0)
                context.INI_Temperature_2 = If(Not String.IsNullOrEmpty(config.Temperature), config.Temperature, "")
                context.INI_Model_2 = If(Not String.IsNullOrEmpty(config.Model), config.Model, "")
                context.INI_APIEncrypted_2 = config.APIEncrypted
                context.INI_APIKeyPrefix_2 = If(Not String.IsNullOrEmpty(config.APIKeyPrefix), config.APIKeyPrefix, "")
                context.INI_OAuth2_2 = config.OAuth2
                context.INI_OAuth2ClientMail_2 = If(Not String.IsNullOrEmpty(config.OAuth2ClientMail), config.OAuth2ClientMail, "")
                context.INI_OAuth2Scopes_2 = If(Not String.IsNullOrEmpty(config.OAuth2Scopes), config.OAuth2Scopes, "")
                context.INI_OAuth2Endpoint_2 = If(Not String.IsNullOrEmpty(config.OAuth2Endpoint), config.OAuth2Endpoint, "")
                context.INI_OAuth2ATExpiry_2 = If(config.OAuth2ATExpiry <> 0, config.OAuth2ATExpiry, 3600)
                context.DecodedAPI_2 = config.DecodedAPI
                context.TokenExpiry_2 = config.TokenExpiry
                context.INI_Model_Parameter1 = If(Not String.IsNullOrEmpty(config.Parameter1), config.Parameter1, "")
                context.INI_Model_Parameter2 = If(Not String.IsNullOrEmpty(config.Parameter2), config.Parameter2, "")
                context.INI_Model_Parameter3 = If(Not String.IsNullOrEmpty(config.Parameter3), config.Parameter3, "")
                context.INI_Model_Parameter4 = If(Not String.IsNullOrEmpty(config.Parameter4), config.Parameter4, "")
                context.SP_MergePrompt = If(Not String.IsNullOrEmpty(config.MergePrompt), config.MergePrompt, "")
                SP_QueryPrompt = If(Not String.IsNullOrEmpty(config.QueryPrompt), config.QueryPrompt, "")

                ' === APPLY TOOLING PROPERTIES ===
                ' These are shared/module-level variables used by LLM() for tool-related behavior.
                context.INI_APICall_ToolInstructions_2 = If(Not String.IsNullOrEmpty(config.APICall_ToolInstructions), config.APICall_ToolInstructions, "")
                context.INI_APICall_ToolInstructions_Template_2 = If(Not String.IsNullOrEmpty(config.APICall_ToolInstructions_Template), config.APICall_ToolInstructions_Template, "")
                context.INI_APICall_ToolResponses_2 = If(Not String.IsNullOrEmpty(config.APICall_ToolResponses), config.APICall_ToolResponses, "")
                context.INI_APICall_ToolResponses_Template_2 = If(Not String.IsNullOrEmpty(config.APICall_ToolResponses_Template), config.APICall_ToolResponses_Template, "")
                context.INI_APICall_ToolCallPart_Template_2 = If(Not String.IsNullOrEmpty(config.APICall_ToolCallPart_Template), config.APICall_ToolCallPart_Template, "")
                context.INI_ToolCallDetectionPattern_2 = If(Not String.IsNullOrEmpty(config.ToolCallDetectionPattern), config.ToolCallDetectionPattern, "")
                context.INI_ToolCallExtractionMap_2 = If(Not String.IsNullOrEmpty(config.ToolCallExtractionMap), config.ToolCallExtractionMap, "")

                ErrorFlag = False

            Catch ex As System.Exception
                If Not ErrorFlag Then
                    MessageBox.Show("Error in ApplyModelConfig: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                ErrorFlag = True
            End Try
        End Sub

        ''' <summary>
        ''' Restores a previously captured default configuration by applying it to the provided context.
        ''' </summary>
        ''' <param name="context">Target context to receive the configuration values.</param>
        ''' <param name="originalConfig">Configuration snapshot to restore.</param>
        Public Shared Sub RestoreDefaults(ByVal context As ISharedContext, ByVal originalConfig As ModelConfig)
            ApplyModelConfig(context, originalConfig)
        End Sub

        ''' <summary>
        ''' Loads alternative model configurations from an INI file.
        ''' Each INI section is treated as one model configuration; section names are stored in
        ''' <see cref="ModelConfig.ModelDescription"/> (possibly extended with <c>ModelNote</c> and tooling suffix).
        ''' Deprecation and tool-only filters are applied in <c>ProcessModelSection</c>.
        ''' </summary>
        ''' <param name="iniFilePath">Path to the INI file.</param>
        ''' <param name="context">Shared context used when creating <see cref="ModelConfig"/> instances.</param>
        ''' <param name="Title">The purpose for which the models are used (used only for error message text).</param>
        ''' <param name="includeToolOnly">If <c>True</c>, includes <c>ToolOnly</c> entries; otherwise excludes them.</param>
        ''' <param name="toolsOnly">
        ''' If <c>True</c>, only returns entries that contain a <c>ToolDefinition</c> or <c>ToolInstructionsPrompt</c>.
        ''' </param>
        ''' <returns>List of parsed <see cref="ModelConfig"/> entries (empty if the file is missing or has no sections).</returns>
        Public Shared Function LoadAlternativeModels(ByVal iniFilePath As String,
                                                      context As ISharedContext,
                                                      Optional Title As String = "",
                                                      Optional includeToolOnly As Boolean = False,
                                                      Optional toolsOnly As Boolean = False) As List(Of ModelConfig)
            Dim models As New List(Of ModelConfig)()
            Try
                If Not File.Exists(iniFilePath) Then
                    ShowCustomMessageBox($"INI file for alternative models not found (update {AN2}.ini): " & iniFilePath)
                    Return models
                End If

                Dim currentDict As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                Dim Description As String = ""

                For Each XLine In File.ReadAllLines(iniFilePath)
                    Dim trimmedLine As String = XLine.Trim()
                    ' Skip empty lines and comments (';').
                    If String.IsNullOrEmpty(trimmedLine) OrElse trimmedLine.StartsWith(";") Then
                        Continue For
                    End If

                    ' Section header (e.g., [Model1]) starts a new model section.
                    If trimmedLine.StartsWith("[") AndAlso trimmedLine.EndsWith("]") Then
                        If currentDict.Count > 0 Then
                            ProcessModelSection(currentDict, Description, context, models, includeToolOnly, toolsOnly)
                            currentDict.Clear()
                        End If
                        Description = trimmedLine.Substring(1, trimmedLine.Length - 2).Trim()
                        Continue For
                    End If

                    ' Parse key=value lines.
                    Dim tokens() As String = trimmedLine.Split(New Char() {"="c}, 2)
                    If tokens.Length = 2 Then
                        Dim key As String = tokens(0).Trim()
                        Dim value As String = tokens(1).Trim()
                        If Not currentDict.ContainsKey(key) Then
                            currentDict.Add(key, value)
                        Else
                            currentDict(key) = value
                        End If
                    End If
                Next

                ' Materialize the final section (if any).
                If currentDict.Count > 0 Then
                    ProcessModelSection(currentDict, Description, context, models, includeToolOnly, toolsOnly)
                End If

            Catch ex As System.Exception
                ShowCustomMessageBox($"Error reading INI file for models {If(String.IsNullOrWhiteSpace(Title), " ", $"for '{Title}' ")}({iniFilePath}): " & ex.Message)
            End Try
            Return models
        End Function

        ''' <summary>
        ''' Processes a single INI section dictionary and adds a corresponding <see cref="ModelConfig"/> to <paramref name="models"/>
        ''' if the section passes deprecation and tool-related filters.
        ''' </summary>
        ''' <param name="currentDict">The parsed key/value pairs for the current INI section.</param>
        ''' <param name="description">The base description derived from the INI section header.</param>
        ''' <param name="context">Shared context passed to <see cref="CreateModelConfigFromDict"/>.</param>
        ''' <param name="models">The destination list for accepted model configurations.</param>
        ''' <param name="includeToolOnly">Whether to include <c>ToolOnly</c> entries.</param>
        ''' <param name="toolsOnly">Whether to only include tool definitions.</param>
        Private Shared Sub ProcessModelSection(currentDict As Dictionary(Of String, String),
                                               description As String,
                                               context As ISharedContext,
                                               models As List(Of ModelConfig),
                                               includeToolOnly As Boolean,
                                               toolsOnly As Boolean)
            ' Exclude deprecated sections.
            If ParseBoolean(currentDict, "Depreciated") OrElse ParseBoolean(currentDict, "Deprecated") Then
                Return
            End If

            ' Exclude ToolOnly sections unless explicitly included.
            Dim isToolOnly = ParseBoolean(currentDict, "ToolOnly")
            If isToolOnly AndAlso Not includeToolOnly Then
                Return
            End If

            ' If selecting tools only, require a tool definition or tool prompt.
            If toolsOnly Then
                Dim hasToolDef = currentDict.ContainsKey("ToolDefinition") AndAlso Not String.IsNullOrWhiteSpace(currentDict("ToolDefinition"))
                Dim hasToolPrompt = currentDict.ContainsKey("ToolInstructionsPrompt") AndAlso Not String.IsNullOrWhiteSpace(currentDict("ToolInstructionsPrompt"))
                If Not hasToolDef AndAlso Not hasToolPrompt Then
                    Return
                End If
            End If

            ' Build description with ModelNote
            Dim finalDescription As String = BuildModelDescription(description, currentDict)

            ' Only add ToolingSuffix for MODELS that can call tools (have APICall_ToolInstructions),
            ' NOT for tools/sources themselves (toolsOnly=True or Tool=True).
            If Not toolsOnly Then
                Dim hasToolInstructions = currentDict.ContainsKey("APICall_ToolInstructions") AndAlso
                                          Not String.IsNullOrWhiteSpace(currentDict("APICall_ToolInstructions"))
                Dim isTool = ParseBoolean(currentDict, "Tool")

                ' Add suffix only if: can call tools AND is not itself a tool
                If hasToolInstructions AndAlso Not isTool AndAlso Not finalDescription.EndsWith(ToolingSuffix) Then
                    finalDescription &= ToolingSuffix
                End If
            End If

            Dim mc = CreateModelConfigFromDict(currentDict, context, finalDescription)
            models.Add(mc)
        End Sub

        ''' <summary>
        ''' Builds the model display description by combining the section header with the optional <c>ModelNote</c> value.
        ''' </summary>
        ''' <param name="sectionHeader">The INI section header (e.g., "GPT-4").</param>
        ''' <param name="configDict">The configuration dictionary for the section.</param>
        ''' <returns>The combined description: "SectionHeader - ModelNote" or "SectionHeader" if no ModelNote is present.</returns>
        Private Shared Function BuildModelDescription(ByVal sectionHeader As String, ByVal configDict As Dictionary(Of String, String)) As String
            Dim modelNote As String = ""
            If configDict.ContainsKey("ModelNote") Then
                modelNote = configDict("ModelNote").Trim()
            End If

            If Not String.IsNullOrWhiteSpace(modelNote) Then
                Return sectionHeader & " - " & modelNote
            Else
                Return sectionHeader
            End If
        End Function

        ''' <summary>
        ''' Snapshot of the default configuration captured prior to applying an alternate model.
        ''' </summary>
        Public Shared originalConfig As ModelConfig

        ''' <summary>
        ''' Stores the current "reset to default after use" UI checkbox state as used by callers.
        ''' </summary>
        Public Shared OptionChecked As Boolean = False

        ''' <summary>
        ''' Indicates whether <see cref="originalConfig"/> is currently considered valid/available for restore logic.
        ''' </summary>
        Public Shared originalConfigLoaded As Boolean = False

        ''' <summary>
        ''' Stores the models selected by the multi-model selector dialog.
        ''' </summary>
        Public Shared SelectedAlternateModels As List(Of ModelConfig)

        ''' <summary>
        ''' Stores the last selected alternate model display name (used for preselection).
        ''' </summary>
        Public Shared LastAlternateModel As String = ""

        ''' <summary>
        ''' Displays a single-selection model picker dialog and applies the chosen configuration to the provided context.
        ''' Also captures the current configuration into <see cref="originalConfig"/> for optional restoration.
        ''' </summary>
        ''' <param name="context">Shared context to be updated.</param>
        ''' <param name="iniFilePath">Path to the INI file containing alternative models.</param>
        ''' <param name="Title">Dialog title.</param>
        ''' <param name="Listtype">Dialog label above the listbox.</param>
        ''' <param name="OptionText">Checkbox label controlling reset-to-default behavior.</param>
        ''' <param name="UseCase">Selection mode (passed to <c>ModelSelectorForm</c>).</param>
        ''' <returns><c>True</c> if the dialog completed with OK; otherwise <c>False</c>.</returns>
        Public Shared Function ShowModelSelection(ByVal context As ISharedContext, iniFilePath As String, Optional Title As String = "Freestyle", Optional Listtype As String = "Select the model you want to use:", Optional OptionText As String = "Reset to default model after use", Optional UseCase As Integer = 1) As Boolean
            Try
                ' Back up the current configuration into `originalConfig` for restore behavior.
                originalConfig = GetCurrentConfig(context)
                originalConfigLoaded = True

                Dim selector As New ModelSelectorForm(iniFilePath, context, Title, Listtype, OptionText, UseCase)
                If selector.ShowDialog() = DialogResult.OK Then
                    If selector.UseDefault AndAlso UseCase = 1 Then
                        RestoreDefaults(context, originalConfig)
                    ElseIf selector.SelectedModel IsNot Nothing Then
                        ApplyModelConfig(context, selector.SelectedModel)
                    End If

                    If selector.SelectedModel IsNot Nothing Then
                        Dim m As ModelConfig = selector.SelectedModel
                        LastAlternateModel = If(Not String.IsNullOrWhiteSpace(m.ModelDescription), m.ModelDescription, m.Model)
                    End If

                    Return True
                Else
                    Return False
                End If
            Catch ex As System.Exception
                MessageBox.Show("Error in ShowModelSelection: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Displays a multi-selection model picker dialog and stores the chosen models in <see cref="SelectedAlternateModels"/>.
        ''' </summary>
        ''' <param name="context">Shared context used for INI parsing defaults and key resolution.</param>
        ''' <param name="modelPath">Path to the INI file; environment variables are expanded.</param>
        ''' <returns><c>True</c> if the dialog completed with OK and at least one model was selected; otherwise <c>False</c>.</returns>
        Public Shared Function ShowMultipleModelSelection(context As ISharedContext,
                                                  modelPath As String) As Boolean
            Try
                Dim iniPath As String = ExpandEnvironmentVariables(modelPath)
                If String.IsNullOrWhiteSpace(iniPath) OrElse Not System.IO.File.Exists(iniPath) Then
                    System.Windows.Forms.MessageBox.Show("The configured alternate model path does not exist.", AN, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning)
                    Return False
                End If

                Dim alternativeModels As System.Collections.Generic.List(Of ModelConfig) = LoadAlternativeModels(iniPath, context)
                If alternativeModels Is Nothing OrElse alternativeModels.Count = 0 Then
                    System.Windows.Forms.MessageBox.Show("No alternate model configurations found in the specified file.", AN, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
                    Return False
                End If

                Using form As New MultiModelSelectorForm(alternativeModels, LastAlternateModel, AN & " - Select Alternate Models", True, "")
                    If form.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
                        Return False
                    End If

                    SelectedAlternateModels = form.SelectedModels
                    If SelectedAlternateModels Is Nothing OrElse SelectedAlternateModels.Count = 0 Then
                        Return False
                    End If

                    Return True
                End Using
            Catch ex As System.Exception
                System.Windows.Forms.MessageBox.Show("Error during multi-model selection: " & ex.Message, AN, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Retrieves and applies the first model whose INI section contains a key matching <paramref name="Task"/>
        ''' with a truthy value (True/Yes/Wahr/Ja/On/1).
        ''' </summary>
        ''' <param name="context">Shared context to receive the selected model configuration.</param>
        ''' <param name="iniFilePath">Path to the INI file.</param>
        ''' <param name="Task">Key name to locate in each INI section.</param>
        ''' <param name="UseCase">Selection mode (currently unused by this method).</param>
        ''' <returns><c>True</c> if a matching model was found and applied; otherwise <c>False</c>.</returns>
        Public Shared Function GetSpecialTaskModel(ByVal context As ISharedContext,
                                               ByVal iniFilePath As String,
                                               ByVal Task As String,
                                               Optional ByVal UseCase As Integer = 1) As Boolean
            If String.IsNullOrWhiteSpace(Task) Then Return False
            Try
                If Not File.Exists(iniFilePath) Then
                    ShowCustomMessageBox($"INI file for alternative models not found (update {AN2}.ini): " & iniFilePath)
                    Return False
                End If

                ' Back up the current configuration into `originalConfig` (same pattern as ShowModelSelection).
                originalConfigLoaded = False
                originalConfig = GetCurrentConfig(context)
                originalConfigLoaded = True

                Dim normalizedTask As String = Task.Trim()
                Dim truthy = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
                    "true", "yes", "wahr", "ja", "on"
                }

                Dim currentDict As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                Dim description As String = ""

                ' Applies the current section if it contains the task key set to a truthy value.
                Dim applyIfMatch As Func(Of Boolean) =
                    Function()
                        If currentDict.Count = 0 Then Return False
                        If currentDict.ContainsKey(normalizedTask) Then
                            Dim raw As String = currentDict(normalizedTask)
                            If raw Is Nothing Then raw = ""

                            ' Strip inline comments ';' or '#', then trim and remove surrounding quotes.
                            Dim scIdx = raw.IndexOf(";"c)
                            If scIdx >= 0 Then raw = raw.Substring(0, scIdx)
                            Dim hashIdx = raw.IndexOf("#"c)
                            If hashIdx >= 0 Then raw = raw.Substring(0, hashIdx)
                            raw = raw.Trim()

                            If raw.Length >= 2 AndAlso ((raw.StartsWith("""") AndAlso raw.EndsWith("""")) OrElse (raw.StartsWith("'") AndAlso raw.EndsWith("'"))) Then
                                raw = raw.Substring(1, raw.Length - 2).Trim()
                            End If

                            Dim lowered = raw.ToLowerInvariant()
                            If truthy.Contains(lowered) OrElse lowered = "1" Then
                                Dim mc = CreateModelConfigFromDict(currentDict, context, description)
                                ApplyModelConfig(context, mc)
                                Return True
                            End If
                        End If
                        Return False
                    End Function

                For Each rawLine In File.ReadAllLines(iniFilePath)
                    Dim line = rawLine.Trim()
                    If line.Length = 0 OrElse line.StartsWith(";") OrElse line.StartsWith("#") Then
                        Continue For
                    End If

                    ' Section header.
                    If line.StartsWith("[") AndAlso line.EndsWith("]") Then
                        If applyIfMatch() Then
                            Return True
                        End If
                        currentDict.Clear()
                        description = line.Substring(1, line.Length - 2).Trim()
                        Continue For
                    End If

                    ' Parse key=value.
                    Dim tokens = line.Split(New Char() {"="c}, 2)
                    If tokens.Length = 2 Then
                        Dim key = tokens(0).Trim()
                        Dim value = tokens(1).Trim()
                        currentDict(key) = value
                    End If
                Next

                ' Final section.
                If applyIfMatch() Then
                    Return True
                End If

                Return False

            Catch ex As Exception
                MessageBox.Show("Error in GetSpecialTaskModel: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try
        End Function

    End Class
End Namespace