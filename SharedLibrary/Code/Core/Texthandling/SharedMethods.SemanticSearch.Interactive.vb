' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.SemanticSearch.Interactive.vb
' Purpose: Provides optional WinForms profile selection and interactive semantic-
'          index generation entry points. Agentic, service and batch builds may
'          omit this file and use the silent APIs from CoreAndGenerator instead.
' =============================================================================

Option Strict On
Option Explicit On
Option Infer On

Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary

    Partial Public Class SharedMethods

        ''' <summary>
        ''' Shows the small shared selection dialog and returns the chosen metadata profile.
        ''' A return value of Nothing means that the user canceled the dialog.
        ''' </summary>
        Public Shared Function SelectSemanticSearchMetadataProfile(
            Optional defaultProfile As SemanticSearchMetadataProfile = SemanticSearchMetadataProfile.Generic,
            Optional prompt As String = SemanticSearchDefaultProfileSelectionPrompt,
            Optional header As String = SemanticSearchDefaultProfileSelectionHeader,
            Optional owner As System.Windows.Forms.IWin32Window = Nothing
        ) As System.Nullable(Of SemanticSearchMetadataProfile)

            If Not System.Enum.IsDefined(GetType(SemanticSearchMetadataProfile), defaultProfile) Then
                defaultProfile = SemanticSearchMetadataProfile.Generic
            End If

            Dim profiles As System.Collections.Generic.List(Of SemanticSearchMetadataProfile) =
                GetSemanticSearchMetadataProfiles()

            Dim items As New System.Collections.Generic.List(Of SelectionItem)()
            For Each profile As SemanticSearchMetadataProfile In profiles
                ' SelectValue reserves zero for cancellation, so UI values are enum values plus one.
                items.Add(New SelectionItem(
                    GetSemanticSearchMetadataProfileDisplayName(profile),
                    CInt(profile) + 1))
            Next

            Dim selectedValue As Integer = SelectValue(
                items,
                CInt(defaultProfile) + 1,
                prompt,
                header,
                owner)

            If selectedValue = 0 Then
                Return Nothing
            End If

            Dim selectedProfile As SemanticSearchMetadataProfile =
                CType(selectedValue - 1, SemanticSearchMetadataProfile)

            If Not System.Enum.IsDefined(GetType(SemanticSearchMetadataProfile), selectedProfile) Then
                Return Nothing
            End If

            Return selectedProfile
        End Function

        ''' <summary>
        ''' Interactive indexing entry point intended for the Help Me UI. It asks the user to select
        ''' the source domain unless <paramref name="selectedProfile"/> is supplied. It does not ask
        ''' the user to write a narrative; Narrative is only one selectable source profile.
        ''' Returns Nothing when the profile-selection dialog is canceled.
        ''' </summary>
        Public Shared Async Function GenerateSemanticSearchIndexAsync(
            inputPath As String,
            outputPath As String,
            context As ISharedContext,
            Optional options As SemanticSearchIndexGeneratorOptions = Nothing,
            Optional selectedProfile As System.Nullable(Of SemanticSearchMetadataProfile) = Nothing,
            Optional owner As System.Windows.Forms.IWin32Window = Nothing,
            Optional progress As System.IProgress(Of SemanticSearchIndexGenerationProgress) = Nothing,
            Optional cancellationToken As System.Threading.CancellationToken = Nothing
        ) As System.Threading.Tasks.Task(Of SemanticSearchIndexGenerationResult)

            Dim effectiveOptions As SemanticSearchIndexGeneratorOptions =
                CloneSemanticSearchIndexGeneratorOptions(options)

            Dim effectiveProfile As System.Nullable(Of SemanticSearchMetadataProfile) = selectedProfile
            If Not effectiveProfile.HasValue Then
                effectiveProfile = SelectSemanticSearchMetadataProfile(
                    effectiveOptions.MetadataProfile,
                    SemanticSearchDefaultProfileSelectionPrompt,
                    SemanticSearchDefaultProfileSelectionHeader,
                    owner)
            End If

            If Not effectiveProfile.HasValue Then
                Return Nothing
            End If

            effectiveOptions.MetadataProfile = effectiveProfile.Value

            Return Await CreateSemanticSearchIndexedTextFileAsync(
                inputPath,
                outputPath,
                context,
                effectiveOptions,
                progress,
                cancellationToken).ConfigureAwait(False)
        End Function

        ''' <summary>
        ''' General interactive alias for callers outside the Help Me UI.
        ''' Returns Nothing when the profile-selection dialog is canceled.
        ''' </summary>
        Public Shared Async Function GenerateSemanticSearchIndexInteractiveAsync(
            inputPath As String,
            outputPath As String,
            context As ISharedContext,
            Optional options As SemanticSearchIndexGeneratorOptions = Nothing,
            Optional selectedProfile As System.Nullable(Of SemanticSearchMetadataProfile) = Nothing,
            Optional owner As System.Windows.Forms.IWin32Window = Nothing,
            Optional progress As System.IProgress(Of SemanticSearchIndexGenerationProgress) = Nothing,
            Optional cancellationToken As System.Threading.CancellationToken = Nothing
        ) As System.Threading.Tasks.Task(Of SemanticSearchIndexGenerationResult)

            Return Await GenerateSemanticSearchIndexAsync(
                inputPath,
                outputPath,
                context,
                options,
                selectedProfile,
                owner,
                progress,
                cancellationToken).ConfigureAwait(False)
        End Function

    End Class
End Namespace
