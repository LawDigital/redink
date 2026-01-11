' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ModelConfig.vb
' Purpose: Defines a mutable data container for the currently selected LLM/model configuration and
'          related runtime state used by model selection and invocation code.
'
' Architecture:
'  - Configuration Carrier: Instances are populated from INI/config dictionaries (e.g., via
'    `SharedMethods.CreateModelConfigFromDict`) and can be applied back to `ISharedContext` (e.g., via
'    `SharedMethods.ApplyModelConfig`).
'  - Credentials: Stores API key material and related flags (encryption/prefix handling) and OAuth2
'    configuration values used by callers.
'  - Runtime/Diagnostics State: Stores request/response strings, token counting, and token expiry information.
'  - Tooling Support: Stores tool definitions, instructions, and response handling configuration for
'    models/services that support tool use.
'  - Cloning: `Clone()` returns a shallow copy via `MemberwiseClone` for snapshotting.
' =============================================================================


Option Strict On
Option Explicit On

Namespace SharedLibrary
    Public Class ModelConfig

        Public Property APIKey As String
        Public Property APIKeyBack As String
        Public Property Temperature As String
        Public Property Timeout As Long
        Public Property MaxOutputToken As Integer
        Public Property Model As String
        Public Property Endpoint As String
        Public Property HeaderA As String
        Public Property HeaderB As String
        Public Property APICall As String
        Public Property APICall_Object As String
        Public Property Response As String
        Public Property Anon As String
        Public Property TokenCount As String
        Public Property APIEncrypted As Boolean
        Public Property APIKeyPrefix As String
        Public Property OAuth2 As Boolean
        Public Property OAuth2ClientMail As String
        Public Property OAuth2Scopes As String
        Public Property OAuth2Endpoint As String
        Public Property OAuth2ATExpiry As Long
        Public Property ModelDescription As String
        Public Property DecodedAPI As String
        Public Property TokenExpiry As DateTime
        Public Property Parameter1 As String
        Public Property Parameter2 As String
        Public Property Parameter3 As String
        Public Property Parameter4 As String
        Public Property MergePrompt As String
        Public Property QueryPrompt As String
        Public Property Tool As Boolean
        Public Property ToolOnly As Boolean
        Public Property Deprecated As Boolean
        Public Property APICall_ToolInstructions As String
        Public Property APICall_ToolInstructions_Template As String
        Public Property APICall_ToolResponses As String
        Public Property ToolParameterDefaults As String = ""
        Public Property APICall_ToolResponses_Template As String
        Public Property ToolCallDetectionPattern As String
        Public Property ToolCallExtractionMap As String
        Public Property ToolName As String
        Public Property ToolInstructionsPrompt As String
        Public Property ToolDefinition As String
        Public Property ToolAPICall As String
        Public Property ToolPriority As Integer = 100
        Public Property ToolErrorHandling As String = "skip"
        Public Property APICall_ToolCallPart_Template As String = ""

        Public Function Clone() As ModelConfig
            Return DirectCast(Me.MemberwiseClone(), ModelConfig)
        End Function
    End Class

End Namespace