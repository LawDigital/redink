' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: AppConfigurationVariable.vb
' Purpose: Data model for installation wizard configuration variables. Represents
'          user-configurable settings (API keys, endpoints, timeouts, etc.) with
'          metadata for dynamic UI generation, validation, and INI persistence.
'
' Architecture:
'   Simple POCO/DTO pattern with six string properties. No validation logic or
'   behavior—pure data container consumed by InitialConfig.vb wizard form.
'
'   Workflow (InitialConfig.vb orchestrates):
'   1. PrepareConfigData() creates provider dictionaries (e.g., "OpenAI" -> list of vars)
'   2. LoadConfigForSelectedProvider() generates Label+TextBox pairs from DisplayName/CurrentValue
'   3. User input updates TextBoxes; SaveCurrentInputToConfig() copies Text -> CurrentValue
'   4. ValidateAllConfigs() checks CurrentValue against ValidationRule
'   5. btnOK_Click() maps VarName -> CurrentValue to ISharedContext properties
'   6. CreateAppConfig() writes context to INI file (e.g., %APPDATA%\redink.ini)
'
'   Each instance represents one configuration field for one LLM provider.
'   Typical usage: 8-13 vars/provider × 4-6 providers = 40-80 instances total.
'
' Key Properties:
'   DisplayName: UI label (e.g., "API Key:", "Temperature:")
'   VarName: INI key and context mapping (e.g., "INI_APIKey" -> context.INI_APIKey)
'   VarType: Type hint ("String", "Integer"); informational only, not enforced
'   ValidationRule: Constraint expression ("NotEmpty", ">0", "Hyperlink", "2.0")
'   DefaultValue: Initial/recommended value (e.g., "0.2", "gpt-4.1")
'   CurrentValue: Mutable user input; preserves state during provider switching
'
' Dependencies:
'   Consumed by: InitialConfig.vb (wizard form)
'   Persisted to: ISharedContext properties -> INI files
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary

    ''' <summary>
    ''' Represents a single configurable setting for the installation wizard.
    ''' Contains metadata for UI generation, validation rules, default value,
    ''' and current user-entered value. Used by InitialConfig to dynamically
    ''' create provider-specific configuration forms.
    ''' </summary>
    ''' <remarks>
    ''' Lightweight DTO with no behavior. Designed for wizard state management
    ''' and INI persistence. All validation logic lives in InitialConfig.vb.
    ''' 
    ''' Typical lifecycle:
    ''' 1. Created in PrepareConfigData() with DefaultValue populated
    ''' 2. CurrentValue initialized = DefaultValue
    ''' 3. UI controls generated from DisplayName/CurrentValue
    ''' 4. User edits TextBox; copied to CurrentValue on provider switch or OK
    ''' 5. Validated against ValidationRule
    ''' 6. Mapped via VarName to ISharedContext properties
    ''' 7. Written to INI file
    ''' 
    ''' Supports LLM provider configurations (OpenAI, Azure, Google Vertex, etc.)
    ''' with fields like API keys, endpoints, model names, temperature, timeout,
    ''' HTTP headers, JSON templates, and OAuth2 settings.
    ''' </remarks>
    Public Class AppConfigurationVariable

        ''' <summary>
        ''' Gets or sets the user-friendly label displayed in the wizard UI.
        ''' </summary>
        ''' <value>Display text for Label control (e.g., "API Key:", "Temperature:").</value>
        ''' <remarks>
        ''' Used for Label.Text in LoadConfigForSelectedProvider(), validation error
        ''' messages, and mapping labels back to variables. Convention: end with colon,
        ''' include units if applicable (e.g., "Timeout (ms):").
        ''' </remarks>
        Public Property DisplayName As String

        ''' <summary>
        ''' Gets or sets the internal variable name used for INI persistence and
        ''' ISharedContext property mapping.
        ''' </summary>
        ''' <value>Unique identifier (e.g., "INI_APIKey", "INI_Temperature").</value>
        ''' <remarks>
        ''' All variables use "INI_*" prefix. Maps to context properties in btnOK_Click
        ''' Select Case. Stripped to "APIKey", "Temperature", etc. when written to INI.
        ''' Complete list includes: INI_APIKey, INI_Temperature, INI_Timeout, INI_Model,
        ''' INI_Endpoint, INI_HeaderA/B, INI_APICall, INI_Response, INI_OAuth2*.
        ''' </remarks>
        Public Property VarName As String

        ''' <summary>
        ''' Gets or sets the data type hint for this variable.
        ''' </summary>
        ''' <value>Type identifier (e.g., "String", "Integer").</value>
        ''' <remarks>
        ''' Currently informational only; all variables render as TextBox regardless
        ''' of VarType. Type conversion happens during validation/persistence (e.g.,
        ''' CInt for timeouts). Future enhancement: use to generate NumericUpDown,
        ''' ComboBox, CheckBox, etc.
        ''' </remarks>
        Public Property VarType As String

        ''' <summary>
        ''' Gets or sets the validation rule applied to CurrentValue before persistence.
        ''' </summary>
        ''' <value>Constraint expression (e.g., "NotEmpty", ">0", "Hyperlink", "2.0").</value>
        ''' <remarks>
        ''' Interpreted by ValidateAllConfigs() using Contains() and regex:
        ''' - "NotEmpty": Required field
        ''' - "E-Mail": Must contain "@"
        ''' - "Hyperlink": Must start with "http://" or "https://"
        ''' - ">0": Positive integer
        ''' - "0.0-2.0": Explicit range
        ''' - "2.0": Max value (validates 0 <= CurrentValue <= 2.0)
        ''' Supports "," and "." as decimal separators for European locales.
        ''' </remarks>
        Public Property ValidationRule As String

        ''' <summary>
        ''' Gets or sets the default/recommended value shown when variable is first presented.
        ''' </summary>
        ''' <value>Initial value string (e.g., "0.2", "gpt-4.1", "https://...").</value>
        ''' <remarks>
        ''' Used to initialize CurrentValue and populate TextBox controls. Can contain
        ''' placeholders like {apikey}, {model}, {promptsystem} for runtime replacement.
        ''' Can be overridden by remote defaults downloaded from RemoteDefaultsUrl.
        ''' Never modified; CurrentValue is the mutable copy.
        ''' </remarks>
        Public Property DefaultValue As String

        ''' <summary>
        ''' Gets or sets the current user-entered value for this variable.
        ''' </summary>
        ''' <value>User-modified value synchronized from TextBox.Text during save operations.</value>
        ''' <remarks>
        ''' Preserves state during wizard navigation (provider switching). NOT directly
        ''' data-bound; manually synchronized via SaveCurrentInputToConfig() and
        ''' LoadConfigForSelectedProvider(). Validated against ValidationRule before
        ''' mapping to ISharedContext. Stored as string; type conversion happens during
        ''' validation (Double.TryParse) and persistence (CInt, direct assignment).
        ''' </remarks>
        Public Property CurrentValue As String

    End Class

End Namespace