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

        ''' <summary>
        ''' Tries to load, without applying, the first alternate model whose specified task key is truthy.
        ''' </summary>
        ''' <param name="context">Shared context used to materialize the model configuration.</param>
        ''' <param name="iniFilePath">Path to the alternate models INI file.</param>
        ''' <param name="taskKey">Task flag to search for, such as "ToolDefaultModel".</param>
        ''' <param name="modelConfig">Receives the matching model configuration when found.</param>
        ''' <returns>True if a matching model was found and materialized; otherwise False.</returns>
        Public Shared Function TryGetSpecialTaskModelConfig(ByVal context As ISharedContext,
                                                            ByVal iniFilePath As String,
                                                            ByVal taskKey As String,
                                                            ByRef modelConfig As ModelConfig) As Boolean
            modelConfig = Nothing

            If context Is Nothing Then Return False
            If String.IsNullOrWhiteSpace(iniFilePath) Then Return False
            If String.IsNullOrWhiteSpace(taskKey) Then Return False

            Try
                Dim cacheHit As Boolean = False
                Dim sections As List(Of AlternativeModelIniSection) = GetAlternativeModelIniSections(iniFilePath, cacheHit)
                Dim normalizedTaskKey As String = taskKey.Trim()

                For Each section As AlternativeModelIniSection In sections
                    If section Is Nothing OrElse section.Values Is Nothing Then Continue For

                    Dim raw As String = Nothing
                    If Not section.Values.TryGetValue(normalizedTaskKey, raw) Then Continue For
                    If Not IsTruthyIniValue(raw) Then Continue For
                    If Not IsModelAccessibleForCurrentUser(section.Values, context) Then Continue For

                    ' Materialize the special-task model with the same display identity used by
                    ' LoadAlternativeModels. Local Chat persists that display key, so divergent
                    ' ModelNote/tooling-suffix handling would otherwise make a valid AgentDefaultModel
                    ' look like a different or stale alternate-model selection.
                    Dim displayDescription As System.String = BuildModelDescription(section.Description, section.Values)
                    Dim matched As ModelConfig = CreateModelConfigFromDict(section.Values, context, displayDescription)
                    If matched Is Nothing Then Continue For
                    matched.ModelDescription = GetModelDisplayWithToolingSuffix(matched)

                    modelConfig = matched
                    Return True
                Next

                Return False
            Catch
                modelConfig = Nothing
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Returns True when the alternate models INI contains a tooling-capable model
        ''' for the specified task key.
        ''' </summary>
        ''' <param name="context">Shared context used to materialize the model configuration.</param>
        ''' <param name="iniFilePath">Path to the alternate models INI file.</param>
        ''' <param name="taskKey">Task flag to search for, such as "ToolDefaultModel".</param>
        ''' <returns>True if a matching model exists and supports tooling; otherwise False.</returns>
        Public Shared Function HasToolingCapableSpecialTaskModel(ByVal context As ISharedContext,
                                                                 ByVal iniFilePath As String,
                                                                 ByVal taskKey As String) As Boolean
            Dim specialTaskModel As ModelConfig = Nothing

            Return TryGetSpecialTaskModelConfig(context, iniFilePath, taskKey, specialTaskModel) AndAlso
                   ModelSupportsTooling(specialTaskModel)
        End Function


    End Class

End Namespace