' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ProgressBarModule.vb
' Purpose: Provides shared state and a helper method for showing a modal progress UI
'          on a dedicated STA thread. The progress UI reads the shared state values
'          to update its progress bar, label, and cancellation status.
'
' Architecture:
'  - Shared State: Exposes module-level fields for progress value, maximum, status label,
'    and a cancellation flag.
'  - UI Threading: `ShowProgressBarInSeparateThread` starts a new STA thread and shows a
'    modal progress form on that thread.
'  - UI Consumption: A progress form (for example, `ProgressForm` / `DPIProgressForm`)
'    is expected to read these fields periodically to refresh the UI.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Threading

Namespace SharedLibrary

    ''' <summary>
    ''' Provides shared progress state and a helper for showing a progress dialog on a dedicated UI thread.
    ''' </summary>
    Public Module ProgressBarModule
        ' Global variables to control the progress form.
        Public GlobalProgressValue As Integer = 0
        Public GlobalProgressMax As Integer = 100
        Public GlobalProgressLabel As String = "Initializing..."
        Public CancelOperation As Boolean = False

        ''' <summary>
        ''' Starts a new STA thread and shows a modal progress form on that thread.
        ''' </summary>
        ''' <param name="headerText">The caption text provided to the progress form.</param>
        ''' <param name="initialLabel">The initial status label text provided to the progress form.</param>

        Public Sub ShowProgressBarInSeparateThread(headerText As String, initialLabel As String)
            Dim t As New Thread(Sub()
                                    ' Create and show the progress form modally.
                                    Dim progressForm As New ProgressForm(headerText, initialLabel)
                                    progressForm.ShowDialog()
                                End Sub)
            t.SetApartmentState(ApartmentState.STA)
            t.Start()
        End Sub
    End Module

End Namespace