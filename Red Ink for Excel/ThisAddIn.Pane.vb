' Part of "Red Ink for Excel"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Pane.vb
' Purpose: Provides asynchronous task pane display and an instruction merge workflow.
'          Shows a custom pane, supplies a callback for post-pane text handling,
'          parses returned instruction text, applies instructions, and updates UI state.
'
' Architecture:
' - Partial class ThisAddIn: Excel add-in logic split across files.
' - Thread marshalling: Each public/Private procedure checks mainThreadControl.InvokeRequired;
'   if True, execution is marshalled to the UI thread via Invoke.
' - Pane display: ShowPaneAsync calls PaneManager.ShowMyPane with parameters and an
'   IntelligentMergeCallback delegate referencing HandleIntelligentMerge.
' - Callback chain: PaneManager triggers HandleIntelligentMerge, which ensures UI thread
'   and calls IntelligentMerge.
' - Instruction processing: IntelligentMerge parses text (ParseLLMResponse), applies
'   instructions (ApplyLLMInstructions with DoBubbles=True), notifies user
'   (ShowCustomMessageBox), and updates ribbon state (UpdateUndoButton).
' - External dependencies: PaneManager, ParseLLMResponse, ApplyLLMInstructions, ShowCustomMessageBox,
'   Globals.Ribbons.Ribbon1, IntelligentMergeCallback are assumed supplied elsewhere.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Asynchronously shows the custom pane via PaneManager.ShowMyPane. Marshals to UI thread if required.
    ''' </summary>
    ''' <param name="introLine">Introductory line text.</param>
    ''' <param name="bodyText">Main body text displayed in pane.</param>
    ''' <param name="finalRemark">Final remark text.</param>
    ''' <param name="header">Pane header text.</param>
    ''' <param name="NoRtf">If True, disables RTF formatting.</param>
    ''' <param name="insertMarkdown">If True, interprets bodyText as Markdown.</param>
    ''' <param name="PreserveLiterals">If True, preserves literal content.</param>
    Private Async Sub ShowPaneAsync(
                          introLine As String,
                          bodyText As String,
                          finalRemark As String,
                          header As String,
                          Optional NoRtf As Boolean = False,
                          Optional insertMarkdown As Boolean = False,
                          Optional PreserveLiterals As Boolean = False
                        )
        Try

            ' Ensure we're on the UI thread for the pane operation
            Dim result As String = ""
            If mainThreadControl.InvokeRequired Then
                result = Await CType(mainThreadControl.Invoke(
                    New Func(Of Task(Of String))(
                        Function() As Task(Of String)
                            Return PaneManager.ShowMyPane(introLine, bodyText, finalRemark, header, NoRtf, insertMarkdown, New IntelligentMergeCallback(AddressOf HandleIntelligentMerge), PreserveLiterals)
                        End Function
                    )
                ), Task(Of String))
            Else
                result = Await PaneManager.ShowMyPane(introLine, bodyText, finalRemark, header, NoRtf, insertMarkdown, New IntelligentMergeCallback(AddressOf HandleIntelligentMerge), PreserveLiterals)
            End If

        Catch ex As Exception
            Global.SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("Error in ShowPaneAsync: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Callback wrapper for intelligent merge; ensures execution on UI thread then calls IntelligentMerge.
    ''' </summary>
    ''' <param name="selectedText">Text selected or produced for instruction parsing.</param>
    Private Sub HandleIntelligentMerge(selectedText As String)
        ' Ensure UI operations happen on the main thread
        If mainThreadControl.InvokeRequired Then
            mainThreadControl.Invoke(Sub() IntelligentMerge(selectedText))
        Else
            IntelligentMerge(selectedText)
        End If
    End Sub

    ''' <summary>
    ''' Parses instruction text, applies instructions, shows completion message, updates ribbon undo button. Marshals to UI thread if required.
    ''' </summary>
    ''' <param name="newtext">Raw instruction text returned from pane interaction.</param>
    Public Async Sub IntelligentMerge(newtext As String)
        ' Ensure we're on UI thread for Excel COM operations
        If mainThreadControl.InvokeRequired Then
            mainThreadControl.Invoke(
            Sub()
                Dim instructions As New List(Of String)
                instructions = ParseLLMResponse(newtext)
                ApplyLLMInstructions(instructions, True)  ' Always DoBubbles
                ShowCustomMessageBox("Implementation of the instructions completed (to the extent possible). They are also in the clipboard.")
                Dim result = Globals.Ribbons.Ribbon1.UpdateUndoButton()
            End Sub
        )
        Else
            Dim instructions As New List(Of String)
            instructions = ParseLLMResponse(newtext)
            ApplyLLMInstructions(instructions, True)  ' Always DoBubbles
            ShowCustomMessageBox("Implementation of the instructions completed (to the extent possible). They are also in the clipboard.")
            Dim result = Globals.Ribbons.Ribbon1.UpdateUndoButton()
        End If
    End Sub

End Class