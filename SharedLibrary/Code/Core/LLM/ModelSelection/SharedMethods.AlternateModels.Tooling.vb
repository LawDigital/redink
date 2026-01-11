' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SharedMethods.AlternateModels.Tooling.vb
' Purpose: Provides tooling-related helpers for alternate model configurations.
'
' Architecture:
'  - Tool Capability Detection: `ModelSupportsTooling` returns True when the model
'    has a non-empty `APICall_ToolInstructions` (meaning the model can CALL tools).
'    NOTE: `ModelConfig.Tool = True` means the entry IS a tool, not that it supports tooling.
'  - UI Display Normalization: `GetModelDisplayWithToolingSuffix` appends `ToolingSuffix`
'    to models that can call tools (for display in model selection dialogs).
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary

    Partial Public Class SharedMethods

        ''' <summary>
        ''' Determines if a MODEL supports calling tools based on its configuration.
        ''' Returns True when the model has APICall_ToolInstructions configured.
        ''' </summary>
        ''' <param name="config">ModelConfig to check.</param>
        ''' <returns>True if the model can call tools/sources.</returns>
        ''' <remarks>
        ''' NOTE: This checks if a MODEL can CALL tools, not if the entry IS a tool.
        ''' - ModelConfig.Tool = True means the entry IS a tool/source
        ''' - APICall_ToolInstructions being set means a model CAN CALL tools
        ''' </remarks>
        Public Shared Function ModelSupportsTooling(config As ModelConfig) As Boolean
            If config Is Nothing Then Return False
            Return Not String.IsNullOrWhiteSpace(config.APICall_ToolInstructions)
        End Function

        ''' <summary>
        ''' Gets the display description for a model, including tooling suffix if applicable.
        ''' Only adds suffix for MODELS that can call tools (not for tools themselves).
        ''' </summary>
        ''' <param name="config">ModelConfig to get description for.</param>
        ''' <returns>Display description with appropriate suffix.</returns>
        Public Shared Function GetModelDisplayWithToolingSuffix(config As ModelConfig) As String
            Dim baseDesc = If(Not String.IsNullOrWhiteSpace(config.ModelDescription),
                              config.ModelDescription, config.Model)

            ' Only add suffix for models that can CALL tools, not for tools themselves
            If ModelSupportsTooling(config) AndAlso Not config.Tool Then
                If Not baseDesc.EndsWith(ToolingSuffix) Then
                    baseDesc &= ToolingSuffix
                End If
            End If

            Return baseDesc
        End Function

    End Class

End Namespace