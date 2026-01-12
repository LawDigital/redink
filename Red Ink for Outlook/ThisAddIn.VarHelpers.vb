' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland.
' All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.VarHelpers.vb
' Purpose: Variable and interaction helpers for the Outlook add-in.
'          Current scope: user-driven selection capture and comparison of two
'          text ranges inside the Outlook compose editor (WordEditor).
'
' Architecture:
' - CompareSelectedTextRangesOutlook:
'     - Obtains the active Outlook `Inspector` (compose window) and its Word-based
'       editor via `Inspector.WordEditor`.
'     - Captures the first selection (if present) or prompts the user to select it
'       using SharedLibrary non-modal dialogs (so Outlook remains interactive).
'     - Prompts the user to select a second text range and captures it.
'     - Performs a comparison using existing processing helpers:
'         - Primary: Word compare-docs pipeline via `CompareAndInsertTextCompareDocs`
'           (Word.CompareDocuments-based markup), when enabled by configuration.
'         - Fallback: DiffPlex inline diff via `CompareAndInsertText` (typically shown
'           in a viewer window rather than inserted).
'     - Provides user feedback for missing selections and no-diff cases.
'     - Uses best-effort COM cleanup for local references to `Inspector` / `Document`.
'
' Dependencies:
' - Microsoft Office Interop:
'     - Outlook: `Microsoft.Office.Interop.Outlook.Inspector`
'     - Word: `Microsoft.Office.Interop.Word.Document`, `Selection`
' - SharedLibrary:
'     - `SharedLibrary.SharedLibrary.SharedMethods` for UI prompts and messages.
' - Processing pipeline (same add-in):
'     - `CompareAndInsertTextCompareDocs`, `CompareAndInsertText` (see `ThisAddIn.Processing.vb`).
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Outlook
Imports Microsoft.Office.Interop.Word
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Compare two text selections in an Outlook compose inspector.
    ''' Flow:
    '''  - If user already selected text, use as first selection; otherwise prompt to select it.
    '''  - Prompt to select second text.
    '''  - Compare selections using Word Compare (CompareAndInsertTextCompareDocs) or DiffPlex fallback (CompareAndInsertText).
    ''' Result:
    '''  - Either inserts formatted comparison into the compose window (CompareDocs path)
    '''  - or shows the formatted result in a window (DiffPlex path).
    ''' </summary>
    Public Sub CompareSelectedTextRangesOutlook()
        Dim inspector As Inspector = Nothing
        Dim wordDoc As Document = Nothing
        Dim wordApp As Microsoft.Office.Interop.Word.Application = Nothing

        Try
            inspector = TryCast(Globals.ThisAddIn.Application.ActiveInspector(), Inspector)
            If inspector Is Nothing Then
                SLib.ShowCustomMessageBox("No active compose window found. Please open an email in compose mode and try again.", AN)
                Return
            End If

            wordDoc = TryCast(inspector.WordEditor, Document)
            If wordDoc Is Nothing Then
                SLib.ShowCustomMessageBox("Unable to access the Word editor for this item. Please ensure you are in an HTML/RTF compose window.", AN)
                Return
            End If

            wordApp = wordDoc.Application

            ' --- Step 1: capture first selection (if already selected) ---
            Dim firstText As String = Nothing
            Try
                Dim sel1 As Microsoft.Office.Interop.Word.Selection = wordApp.Selection
                If sel1 IsNot Nothing AndAlso sel1.Range IsNot Nothing AndAlso sel1.Start <> sel1.End Then
                    firstText = sel1.Range.Text
                End If
            Catch
            End Try

            If String.IsNullOrWhiteSpace(firstText) Then
                Dim step1 As Integer = SLib.ShowCustomYesNoBox(
                    "Please select the FIRST text range to compare in the compose window, then click 'Selection Ready'.",
                    "Selection Ready",
                    "Cancel",
                    $"{AN} Compare Selected - Step 1",
                    nonModal:=True)

                If step1 <> 1 Then Return

                Try
                    Dim sel1 As Microsoft.Office.Interop.Word.Selection = wordApp.Selection
                    If sel1 IsNot Nothing AndAlso sel1.Range IsNot Nothing AndAlso sel1.Start <> sel1.End Then
                        firstText = sel1.Range.Text
                    End If
                Catch
                End Try

                If String.IsNullOrWhiteSpace(firstText) Then
                    SLib.ShowCustomMessageBox("No text was selected for the first range. Operation cancelled.", AN)
                    Return
                End If
            End If

            ' --- Step 2: prompt and capture second selection ---
            Dim step2 As Integer = SLib.ShowCustomYesNoBox(
                $"First selection captured ({firstText.Length} characters).{vbCrLf}{vbCrLf}Now please select the SECOND text range to compare, then click 'Selection Ready'.",
                "Selection Ready",
                "Cancel",
                $"{AN} Compare Selected - Step 2",
                nonModal:=True)

            If step2 <> 1 Then Return

            Dim secondText As String = Nothing
            Try
                Dim sel2 As Microsoft.Office.Interop.Word.Selection = wordApp.Selection
                If sel2 IsNot Nothing AndAlso sel2.Range IsNot Nothing AndAlso sel2.Start <> sel2.End Then
                    secondText = sel2.Range.Text
                End If
            Catch
            End Try

            If String.IsNullOrWhiteSpace(secondText) Then
                SLib.ShowCustomMessageBox("No text was selected for the second range. Operation cancelled.", AN)
                Return
            End If

            ' --- identical check ---
            If String.Equals(firstText, secondText, StringComparison.Ordinal) Then
                SLib.ShowCustomMessageBox("The two selected text ranges are identical. No differences to show.", AN)
                Return
            End If

            ' --- compare (Word CompareDocs or fallback diff) ---
            ' Use your existing configuration switch (same pattern as Word CompareSelectionHalves).
            ' INI_MarkupMethodHelper = 1 => CompareDoc path (Word CompareDocuments)
            'If INI_MarkupMethodHelper = 1 Then
            ' Inserts formatted comparison at current cursor in compose window.
            'CompareAndInsertTextCompareDocs(firstText, secondText)
            'Else
            ' Shows the diff in a window (does not insert).
            CompareAndInsertText(
                    firstText,
                    secondText,
                    ShowInWindow:=True,
                    TextforWindow:="Comparison result (not inserted):",
                    DoNotWait:=True)
            'End If

        Catch ex As System.Exception
            SLib.ShowCustomMessageBox($"Failed to compare selected text: {ex.Message}", AN)
        Finally
            ' Best-effort COM cleanup (avoid over-releasing objects owned by Outlook/Word)
            If wordDoc IsNot Nothing Then
                Try : System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDoc) : Catch : End Try
                wordDoc = Nothing
            End If
            If inspector IsNot Nothing Then
                Try : System.Runtime.InteropServices.Marshal.ReleaseComObject(inspector) : Catch : End Try
                inspector = Nothing
            End If
        End Try
    End Sub


    Private _quickTranslateWidget As SharedLibrary.SharedLibrary.QuickTranslateWidget = Nothing

    Public Sub ShowQuickTranslate()
        If _quickTranslateWidget Is Nothing OrElse _quickTranslateWidget.IsDisposed Then
            _quickTranslateWidget = New SharedLibrary.SharedLibrary.QuickTranslateWidget(
                Async Function(text, lang, token)
                    TranslateLanguage = lang
                    Return Await LLM(InterpolateAtRuntime(SP_Translate),
                                    "<TEXTTOPROCESS>" & text & "</TEXTTOPROCESS>",
                                    "", "", 0,
                                    UseSecondAPI:=False,
                                    HideSplash:=True,
                                    cancellationToken:=token,
                                    EnsureUI:=False)
                End Function,
                INI_Language1)
        End If
        _quickTranslateWidget.ShowWidget()
    End Sub


End Class
