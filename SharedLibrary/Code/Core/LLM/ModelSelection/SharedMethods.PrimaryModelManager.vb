' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' Note: The "PrimaryModelManager" has been contributed by Roman Kost, Switzerland
'
' =============================================================================
' File: SharedMethods.PrimaryModelManager.vb
' Purpose: Manages multiple primary LLM model configurations parsed from `redink.ini`,
'          supports selection of the active primary model, and applies the selection
'          to the shared runtime context (`ISharedContext`).
'
' Architecture:
'  - Discovery: Scans INI key/value pairs for `MultiModel_<Key>_<Number>` entries and
'    groups them by model number (e.g., MultiModel_Model_1, MultiModel_Model_2).
'  - Validation: Ensures each model definition includes required keys (Endpoint, Model,
'    APICall, Response); invalid entries trigger a warning and are skipped.
'  - Backward Compatibility: If no `MultiModel_*` keys exist, uses the base (un-suffixed)
'    INI keys as Model 1.
'  - Selection: Copies the selected model's fields into the primary `INI_*` properties
'    in `ISharedContext` and resolves the runtime API key via `SharedMethods.RealAPIKey`.
'  - Persistence: Stores the selected model number in `My.Settings.SelectedModelNumber`.
'
' Notes:
'  - This class only manipulates PRIMARY fields (`INI_APIKey`, `INI_Model`, ...).
'  - The secondary API slot (`INI_*_2`) is not modified.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary

    ''' <summary>
    ''' Manages multiple LLM model configurations loaded from redink.ini.
    ''' </summary>
    ''' <remarks>
    ''' Detects numbered model variants, stores them in memory, and applies the selected
    ''' model to the primary `INI_*` fields in the shared context.
    ''' </remarks>
    Public Class PrimaryModelManager

        ''' <summary>Prefix used for per-model INI keys (e.g., MultiModel_Model_1).</summary>
        Private Const MultiModelPrefix As String = "MultiModel_"

        ''' <summary>Maximum allowed model index when parsing numbered INI keys.</summary>
        Private Const MaxModelNumber As Integer = 100

        ''' <summary>
        ''' Required model keys used for validation; must match `SharedMethods.LoadConfig.INIValuesMissing`.
        ''' </summary>
        Private Shared ReadOnly RequiredModelKeys As String() = {"Endpoint", "Model", "APICall", "Response"}

        ''' <summary>Storage for detected model configurations by model number.</summary>
        Private Shared ReadOnly _availableModels As New Dictionary(Of Integer, Dictionary(Of String, String))()

        ''' <summary>Current selected model number (1-based).</summary>
        Private Shared _currentModelNumber As Integer = 1

        ''' <summary>Display names for each model (e.g., Model_N or ItemName_N).</summary>
        Private Shared ReadOnly _modelDisplayNames As New Dictionary(Of Integer, String)()

        ''' <summary>
        ''' Detects and stores all model configurations from the parsed INI dictionary.
        ''' </summary>
        ''' <param name="configDict">INI key/value pairs parsed from `redink.ini`.</param>
        Public Shared Sub DetectAndStoreModels(configDict As Dictionary(Of String, String))
            _availableModels.Clear()
            _modelDisplayNames.Clear()

            ' Preferred format: MultiModel_<Key>_<Number>
            Dim perModel = ExtractPrefixedModels(configDict)

            ' Validate each detected model
            For Each n In perModel.Keys.OrderBy(Function(x) x)
                Dim mc = perModel(n)
                If ValidateModel(mc, $"#{n}") Then
                    _availableModels(n) = mc
                    _modelDisplayNames(n) = GetModelDisplayNameForConfig(mc, n)
                End If
            Next

            ' Backward compatibility: if no numbered models, extract a "primary" model as Model 1
            If _availableModels.Count = 0 Then
                Dim primary = ExtractPrimaryModelConfig(configDict)
                If ValidateModel(primary, "primary (no suffix)") Then
                    _availableModels(1) = primary
                    _modelDisplayNames(1) = GetModelDisplayNameForConfig(primary, 1)
                End If
            End If

            ' If previous current model is no longer available, fall back to 1
            If _availableModels.Count > 0 Then
                If Not _availableModels.ContainsKey(_currentModelNumber) Then
                    _currentModelNumber = 1
                End If
            Else
                _currentModelNumber = 1
            End If
        End Sub

        ''' <summary>
        ''' Returns the list of available model numbers.
        ''' </summary>
        ''' <returns>Sorted list of model numbers.</returns>
        Public Shared Function GetAvailableModels() As List(Of Integer)
            Return _availableModels.Keys.OrderBy(Function(n) n).ToList()
        End Function

        ''' <summary>
        ''' Returns a user-friendly display name for a model.
        ''' </summary>
        ''' <param name="modelNumber">Model number (1-based).</param>
        ''' <returns>Display name or a fallback label.</returns>
        Public Shared Function GetModelDisplayName(modelNumber As Integer) As String
            If _modelDisplayNames.ContainsKey(modelNumber) Then
                Return _modelDisplayNames(modelNumber)
            End If
            Return $"Model {modelNumber}"
        End Function

        ''' <summary>
        ''' Returns the currently selected model number (1-based).
        ''' </summary>
        ''' <returns>The active model number.</returns>
        Public Shared Function GetCurrentModelNumber() As Integer
            Return _currentModelNumber
        End Function

        ''' <summary>
        ''' Selects a model by copying its configuration into the primary INI_* fields.
        ''' </summary>
        ''' <param name="context">Shared context to update.</param>
        ''' <param name="modelNumber">Model number to select.</param>
        ''' <returns>True on success; otherwise False.</returns>
        Public Shared Function SelectModel(context As ISharedContext, modelNumber As Integer) As Boolean
            If context Is Nothing Then Return False
            If Not _availableModels.ContainsKey(modelNumber) Then Return False

            Dim mc = _availableModels(modelNumber)

            Try
                CopyPropertiesToContext(context, mc)
                UpdateDecodedAPI(context)

                _currentModelNumber = modelNumber

                ' Persist selection to settings; ignore failures
                Try
                    My.Settings.SelectedModelNumber = modelNumber
                    My.Settings.Save()
                Catch
                End Try

                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Loads the saved model selection from settings.
        ''' </summary>
        ''' <returns>Saved model number, or 1 if missing/invalid.</returns>
        Public Shared Function LoadSavedModelNumber() As Integer
            Try
                Dim n = My.Settings.SelectedModelNumber
                If n > 0 Then Return n
            Catch
            End Try
            Return 1
        End Function

        ' -----------------------
        '  Internal helpers
        ' -----------------------

        ''' <summary>
        ''' Extracts numbered model configurations using the `MultiModel_*_N` convention.
        ''' </summary>
        ''' <param name="configDict">INI key/value pairs.</param>
        ''' <returns>Dictionary of model number to key/value pairs.</returns>
        Private Shared Function ExtractPrefixedModels(configDict As Dictionary(Of String, String)) As Dictionary(Of Integer, Dictionary(Of String, String))
            Dim perModel As New Dictionary(Of Integer, Dictionary(Of String, String))()

            For Each kvp In configDict
                Dim key = kvp.Key
                If Not key.StartsWith(MultiModelPrefix, StringComparison.OrdinalIgnoreCase) Then Continue For

                Dim remainder = key.Substring(MultiModelPrefix.Length)
                Dim idx = remainder.LastIndexOf("_"c)
                If idx <= 0 OrElse idx = remainder.Length - 1 Then Continue For

                Dim baseKey = remainder.Substring(0, idx)
                Dim suffix = remainder.Substring(idx + 1)

                Dim n As Integer
                If Integer.TryParse(suffix, n) AndAlso n > 0 AndAlso n <= MaxModelNumber Then
                    If Not perModel.ContainsKey(n) Then
                        perModel(n) = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                    End If

                    perModel(n)(baseKey) = kvp.Value
                End If
            Next

            Return perModel
        End Function

        ''' <summary>
        ''' Extracts the base (non-numbered) model configuration for backward compatibility.
        ''' </summary>
        ''' <param name="configDict">INI key/value pairs.</param>
        ''' <returns>Dictionary containing the primary model configuration.</returns>
        Private Shared Function ExtractPrimaryModelConfig(configDict As Dictionary(Of String, String)) As Dictionary(Of String, String)
            Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            For Each kvp In configDict
                Dim key = kvp.Key
                ' Skip keys that clearly belong to numbered or secondary models
                Dim idx = key.LastIndexOf("_"c)
                Dim numericSuffix As Integer
                If idx > 0 AndAlso Integer.TryParse(key.Substring(idx + 1), numericSuffix) AndAlso numericSuffix > 0 AndAlso numericSuffix <= MaxModelNumber Then
                    ' Belongs to model N, handled elsewhere
                    Continue For
                End If
                result(key) = kvp.Value
            Next

            Return result
        End Function

        ''' <summary>
        ''' Validates a model configuration against required keys.
        ''' </summary>
        ''' <param name="config">Model configuration dictionary.</param>
        ''' <param name="modelLabel">Display label used for diagnostics.</param>
        ''' <returns>True if valid; otherwise False.</returns>
        Private Shared Function ValidateModel(config As Dictionary(Of String, String), modelLabel As String) As Boolean
            If config Is Nothing OrElse config.Count = 0 Then
                Dim emptyMessage = BuildMissingKeyMessage(modelLabel, Nothing, True, config)
                Debug.WriteLine($"[ModelConfigManager] {emptyMessage}")
                MessageBox.Show(emptyMessage,
                                "Model configuration invalid",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning)
                Return False
            End If

            Dim missing As New List(Of String)()

            For Each key In RequiredModelKeys
                If Not config.ContainsKey(key) OrElse String.IsNullOrWhiteSpace(config(key)) Then
                    missing.Add(key)
                End If
            Next

            If missing.Count > 0 Then
                Dim detailedMsg = BuildMissingKeyMessage(modelLabel, missing, False, config)
                Debug.WriteLine($"[ModelConfigManager] {detailedMsg}")
                MessageBox.Show(detailedMsg,
                                "Model configuration invalid",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning)
                Return False
            End If

            Return True
        End Function

        ''' <summary>
        ''' Builds an error message describing missing model keys and how to fix them.
        ''' </summary>
        ''' <param name="modelLabel">Model label used in the message.</param>
        ''' <param name="missingKeys">Missing keys; Nothing when no entries exist.</param>
        ''' <param name="noEntries">True when the configuration is empty.</param>
        ''' <param name="modelConfig">The model configuration used for label hints.</param>
        ''' <returns>Formatted error message.</returns>
        Private Shared Function BuildMissingKeyMessage(modelLabel As String,
                                                       missingKeys As List(Of String),
                                                       noEntries As Boolean,
                                                       modelConfig As Dictionary(Of String, String)) As String
            Dim prefix As String
            Dim fixHints As New List(Of String)()

            Dim labelNumber As Integer
            Dim hasNumericLabel As Boolean = False

            If Not String.IsNullOrWhiteSpace(modelLabel) AndAlso modelLabel.StartsWith("#", StringComparison.Ordinal) Then
                Dim suffix = modelLabel.TrimStart("#"c)
                If Integer.TryParse(suffix, labelNumber) Then
                    hasNumericLabel = True
                End If
            End If

            If hasNumericLabel Then
                Dim identifierSuffix As String = ""
                prefix = $"Model #{labelNumber}"
                If missingKeys IsNot Nothing Then
                    For Each key In missingKeys
                        fixHints.Add($"MultiModel_{key}_{labelNumber}")
                    Next
                Else
                    fixHints.Add($"MultiModel_<Key>_{labelNumber}")
                End If

                Dim itemNameExists As Boolean = modelConfig IsNot Nothing AndAlso modelConfig.ContainsKey("ItemName") AndAlso Not String.IsNullOrWhiteSpace(modelConfig("ItemName"))
                Dim modelNameExists As Boolean = modelConfig IsNot Nothing AndAlso modelConfig.ContainsKey("Model") AndAlso Not String.IsNullOrWhiteSpace(modelConfig("Model"))

                If itemNameExists Then
                    identifierSuffix = $"MultiModel_ItemName_{labelNumber} = '{modelConfig("ItemName")}'"
                ElseIf modelNameExists Then
                    identifierSuffix = $"MultiModel_Model_{labelNumber} = '{modelConfig("Model")}'"
                Else
                    identifierSuffix = $"MultiModel suffix #{labelNumber}"
                End If

                prefix = $"{prefix} ({identifierSuffix})"
            Else
                prefix = $"Model {modelLabel}"
                If missingKeys IsNot Nothing Then
                    fixHints.AddRange(missingKeys)
                Else
                    fixHints.Add("base INI keys (e.g., Endpoint, Model, APICall, Response)")
                End If

                If modelConfig IsNot Nothing Then
                    Dim primaryName As String = ""
                    If modelConfig.ContainsKey("ItemName") AndAlso Not String.IsNullOrWhiteSpace(modelConfig("ItemName")) Then
                        primaryName = $"ItemName = '{modelConfig("ItemName")}'"
                    ElseIf modelConfig.ContainsKey("Model") AndAlso Not String.IsNullOrWhiteSpace(modelConfig("Model")) Then
                        primaryName = $"Model = '{modelConfig("Model")}'"
                    End If

                    If Not String.IsNullOrWhiteSpace(primaryName) Then
                        prefix = $"{prefix} ({primaryName})"
                    End If
                End If
            End If

            Dim baseMessage As String
            If noEntries Then
                baseMessage = $"{prefix} did not contain any configuration entries."
            Else
                baseMessage = $"{prefix} is missing required fields: {String.Join(", ", missingKeys)}."
            End If

            Dim hintText = $"Please set the following INI keys: {String.Join(", ", fixHints.Distinct())}"
            Return $"{baseMessage}{Environment.NewLine}{hintText}"
        End Function

        ''' <summary>
        ''' Computes a display name for a model from its config.
        ''' </summary>
        ''' <param name="config">Model configuration dictionary.</param>
        ''' <param name="n">Model number.</param>
        ''' <returns>Display name to show in UI.</returns>
        Private Shared Function GetModelDisplayNameForConfig(config As Dictionary(Of String, String), n As Integer) As String
            If config.ContainsKey("ItemName") AndAlso Not String.IsNullOrWhiteSpace(config("ItemName")) Then
                Return config("ItemName")
            End If
            If config.ContainsKey("Model") AndAlso Not String.IsNullOrWhiteSpace(config("Model")) Then
                Return config("Model")
            End If
            Return $"Model {n}"
        End Function

        ''' <summary>
        ''' Copies model configuration values into the primary fields of the shared context.
        ''' </summary>
        ''' <param name="context">Shared context to update.</param>
        ''' <param name="model">Model configuration dictionary.</param>
        Private Shared Sub CopyPropertiesToContext(context As ISharedContext,
                                                   model As Dictionary(Of String, String))
            ' Strings
            context.INI_APIKey = GetValue(model, "APIKey", "")
            context.INI_APIKeyBack = context.INI_APIKey
            context.INI_APIKeyPrefix = GetValue(model, "APIKeyPrefix", "")
            context.INI_Endpoint = GetValue(model, "Endpoint", "")
            context.INI_Model = GetValue(model, "Model", "")
            context.INI_Temperature = GetValue(model, "Temperature", "")
            context.INI_HeaderA = GetValue(model, "HeaderA", "")
            context.INI_HeaderB = GetValue(model, "HeaderB", "")
            context.INI_Response = GetValue(model, "Response", "")
            context.INI_APICall = GetValue(model, "APICall", "")
            context.INI_APICall_Object = GetValue(model, "APICall_Object", "")
            context.INI_Anon = GetValue(model, "Anon", "")
            context.INI_TokenCount = GetValue(model, "TokenCount", "")
            context.INI_PreCorrection = GetValue(model, "PreCorrection", "")
            context.INI_PostCorrection = GetValue(model, "PostCorrection", "")
            context.INI_OAuth2ClientMail = GetValue(model, "OAuth2ClientMail", "")
            context.INI_OAuth2Scopes = GetValue(model, "OAuth2Scopes", "")
            context.INI_OAuth2Endpoint = GetValue(model, "OAuth2Endpoint", "")

            ' Numerics
            context.INI_Timeout = GetLongValue(model, "Timeout", 0)
            context.INI_MaxOutputToken = GetIntValue(model, "MaxOutputToken", 0)
            context.INI_OAuth2ATExpiry = GetLongValue(model, "OAuth2ATExpiry", 3600)

            ' Booleans
            context.INI_APIEncrypted = GetBoolValue(model, "APIKeyEncrypted", False)
            context.INI_DoubleS = GetBoolValue(model, "DoubleS", context.INI_DoubleS)
            context.INI_Clean = GetBoolValue(model, "Clean", context.INI_Clean)
            context.INI_NoDash = GetBoolValue(model, "NoEmDash", context.INI_NoDash)
            context.INI_OAuth2 = GetBoolValue(model, "OAuth2", False)
        End Sub

        ''' <summary>
        ''' Resolves and applies the runtime API key for the primary model.
        ''' </summary>
        ''' <param name="context">Shared context to update.</param>
        Private Shared Sub UpdateDecodedAPI(context As ISharedContext)
            ' Mirrors logic from InitializeConfig for primary API
            If context.INI_OAuth2 Then
                context.INI_APIKey = Trim(SharedMethods.RealAPIKey(context.INI_APIKey, False, True, context)).Replace(vbLf, "").Replace(vbCr, "")
                ' If something goes wrong, RealAPIKey will already have shown an error via SharedMethods
            Else
                context.DecodedAPI = SharedMethods.RealAPIKey(context.INI_APIKey, False, False, context)
            End If
        End Sub

        ''' <summary>
        ''' Returns a string value from a config dictionary or a default value.
        ''' </summary>
        ''' <param name="config">Model configuration dictionary.</param>
        ''' <param name="key">Key to read.</param>
        ''' <param name="defaultValue">Fallback value.</param>
        ''' <returns>Resolved value.</returns>
        Private Shared Function GetValue(config As Dictionary(Of String, String),
                                         key As String,
                                         defaultValue As String) As String
            If config.ContainsKey(key) Then
                Return config(key)
            End If
            Return defaultValue
        End Function

        ''' <summary>
        ''' Returns a long value from a config dictionary or a default value.
        ''' </summary>
        ''' <param name="config">Model configuration dictionary.</param>
        ''' <param name="key">Key to read.</param>
        ''' <param name="defaultValue">Fallback value.</param>
        ''' <returns>Resolved value.</returns>
        Private Shared Function GetLongValue(config As Dictionary(Of String, String),
                                             key As String,
                                             defaultValue As Long) As Long
            If config.ContainsKey(key) Then
                Dim v As Long
                If Long.TryParse(config(key), v) Then
                    Return v
                End If
            End If
            Return defaultValue
        End Function

        ''' <summary>
        ''' Returns an integer value from a config dictionary or a default value.
        ''' </summary>
        ''' <param name="config">Model configuration dictionary.</param>
        ''' <param name="key">Key to read.</param>
        ''' <param name="defaultValue">Fallback value.</param>
        ''' <returns>Resolved value.</returns>
        Private Shared Function GetIntValue(config As Dictionary(Of String, String),
                                            key As String,
                                            defaultValue As Integer) As Integer
            If config.ContainsKey(key) Then
                Dim v As Integer
                If Integer.TryParse(config(key), v) Then
                    Return v
                End If
            End If
            Return defaultValue
        End Function

        ''' <summary>
        ''' Returns a boolean value from a config dictionary or a default value.
        ''' </summary>
        ''' <param name="config">Model configuration dictionary.</param>
        ''' <param name="key">Key to read.</param>
        ''' <param name="defaultValue">Fallback value.</param>
        ''' <returns>Resolved value.</returns>
        Private Shared Function GetBoolValue(config As Dictionary(Of String, String),
                                             key As String,
                                             defaultValue As Boolean) As Boolean
            If config.ContainsKey(key) Then
                Dim raw = config(key).Trim().ToLowerInvariant()
                Return raw = "yes" OrElse raw = "true" OrElse raw = "ja" OrElse raw = "wahr" OrElse raw = "1"
            End If
            Return defaultValue
        End Function

    End Class

End Namespace
