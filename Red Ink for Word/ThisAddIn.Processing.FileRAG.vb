' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Processing.FileRAG.vb
' Purpose: Loads the configured instruction library, queries an LLM for relevant
'          guidance, and prepares the system prompt that governs subsequent Word
'          processing operations.
'
' Architecture:
'  - Library Acquisition: Resolves the library file path (INI_Lib_File), reads the
'    content into LibraryText, and validates availability.
'  - Retrieval Prompting: Builds the library-search system prompt via
'    InterpolateAtRuntime(INI_Lib_Find_SP) and invokes LLM with the current selection.
'  - Result Application: Validates LibResult and selects the apply prompt (markup or
'    plain) via InterpolateAtRuntime before downstream processing.
'  - User Feedback: Utilizes InfoBox for progress updates and ShowCustomMessageBox for
'    blocking error notifications; exception paths surface MessageBox alerts.
'  - External Dependencies: Relies on SharedLibrary.SharedLibrary.SharedMethods for
'    interpolation, file IO, InfoBox, message boxes, and LLM invocation helpers.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Consults the configured instruction library, queries the LLM for matching content,
    ''' and prepares the appropriate apply system prompt for later processing.
    ''' </summary>
    ''' <param name="DoMarkup">True to use the markup-specific apply prompt; otherwise, False.</param>
    ''' <returns>True when the library content is loaded and the LLM returns a result; otherwise, False.</returns>
    Public Async Function ConsultLibrary(DoMarkup As Boolean) As Task(Of Boolean)

        Try

            Dim SysPromptTemp As String

            ' Load the library text

            Dim LibFilePath As String = ExpandEnvironmentVariables(INI_Lib_File)

            InfoBox.ShowInfoBox("Loading the library from " & LibFilePath & " ...")

            LibraryText = ReadTextFile(LibFilePath)

            If String.IsNullOrWhiteSpace(LibraryText) Then
                InfoBox.ShowInfoBox("")
                ShowCustomMessageBox("The library file '" & LibFilePath & "' is empty or could not be read.")
                Return False
            End If

            InfoBox.ShowInfoBox("Asking the LLM to search the library based on the instruction ....")

            SysPromptTemp = InterpolateAtRuntime(INI_Lib_Find_SP)

            LibResult = Await LLM(SysPromptTemp, SelectedText, "", "", INI_Lib_Timeout)

            If String.IsNullOrWhiteSpace(LibResult) Then
                InfoBox.ShowInfoBox("")
                ShowCustomMessageBox("The LLM failed to retrieve relevant content from the library. Will abort.")
                Return False
            End If

            InfoBox.ShowInfoBox("Having the LLM apply the result from the library search: " & LibResult, 6)

            If DoMarkup And Not String.IsNullOrWhiteSpace(SelectedText) Then
                SysPrompt = InterpolateAtRuntime(INI_Lib_Apply_SP_Markup)
            Else
                SysPrompt = InterpolateAtRuntime(INI_Lib_Apply_SP)
            End If

            Return True

        Catch ex As System.Exception
            Global.SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("Error in ConsultLibrary: " & ex.Message, "Error")
            Return False
        End Try

    End Function

End Class
