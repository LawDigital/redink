' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Helpers.vb
' Purpose: Helper components for the Outlook add-in: provides a Word undo scope
'          wrapper (WordUndoScope), robust clipboard setting with retry (SafeSetClipboard),
'          and runtime prompt/template interpolation (InterpolateAtRuntime).
'
' Architecture:
' - WordUndoScope: wraps Word.Application.UndoRecord for custom undo grouping (active only on Word >= 15).
' - SafeSetClipboard: retries clipboard persistence to mitigate temporary lock contention; reports non‑transient failures.
' - InterpolateAtRuntime: replaces {Placeholders} in a template string with field or property values on ThisAddIn; removes specific sensitive placeholders.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Provides a disposable undo scope for Word custom undo records (Word 2013 / version >= 15).
    ''' Starts a custom record only if none is currently active.
    ''' </summary>
    Friend NotInheritable Class WordUndoScope
        Implements System.IDisposable

        Private ReadOnly _app As Microsoft.Office.Interop.Word.Application
        Private ReadOnly _undo As Microsoft.Office.Interop.Word.UndoRecord
        Private ReadOnly _iStarted As System.Boolean

        ''' <summary>
        ''' Initializes the scope; if Word major version < 15 no custom record is started.
        ''' </summary>
        ''' <param name="app">Word application instance.</param>
        ''' <param name="name">Optional custom record name.</param>
        Public Sub New(app As Microsoft.Office.Interop.Word.Application, Optional name As System.String = Nothing)
            _app = app
            _undo = _app.UndoRecord

            ' Word versions earlier than 2013 (version < 15.0) have no UndoRecord.
            Dim ver As System.Version = New System.Version(_app.Version)
            If ver.Major < 15 Then
                Return
            End If

            ' Start only if no other custom record is currently running.
            If Not _undo.IsRecordingCustomRecord Then
                If name IsNot Nothing AndAlso name.Length > 0 Then
                    _undo.StartCustomRecord(name)
                Else
                    _undo.StartCustomRecord("VSTO-Aktion")
                End If
                _iStarted = True
            End If
        End Sub

        ''' <summary>
        ''' Ends the custom undo record if this instance started it.
        ''' </summary>
        Public Sub Dispose() Implements System.IDisposable.Dispose
            Try
                If _iStarted AndAlso _undo.IsRecordingCustomRecord Then
                    _undo.EndCustomRecord()
                End If
            Catch ex As System.Exception
                ' Do not throw – inside Dispose.
            End Try
        End Sub
    End Class

    ' Helper: robustly set clipboard with retries to avoid "clipboard locked" errors.
    ''' <summary>
    ''' Attempts to set the clipboard persistently with retry backoff; reports failure after max attempts.
    ''' </summary>
    ''' <param name="dataObj">DataObject to place on clipboard.</param>
    Private Sub SafeSetClipboard(dataObj As System.Windows.Forms.DataObject)
        Const maxAttempts As Integer = 8
        For attempt As Integer = 1 To maxAttempts
            Try
                System.Windows.Forms.Clipboard.SetDataObject(dataObj, True)
                Return
            Catch ex As System.Runtime.InteropServices.ExternalException
                ' Clipboard likely locked by another process, retry.
                System.Threading.Thread.Sleep(40 * attempt)
            Catch ex As System.Exception
                ' Non‑transient; bail out.
                SLib.ShowCustomMessageBox($"Clipboard copy failed: {ex.Message}")
                Return
            End Try
        Next
        SLib.ShowCustomMessageBox("Could not access the clipboard after several retries (another application may be holding it).")
    End Sub

    ''' <summary>
    ''' Replaces placeholders of form {Name} in a template with field or property values from ThisAddIn; removes selected sensitive placeholders.
    ''' </summary>
    ''' <param name="template">Input template string containing placeholders.</param>
    ''' <returns>Interpolated string or empty string on error.</returns>
    Public Function InterpolateAtRuntime(ByVal template As String) As String
        If template Is Nothing Then
            MessageBox.Show("Error InterpolateAtRuntime: Template is Nothing.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return ""
        End If

        template = Regex.Replace(template, "{Codebasis}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_DecodedAPI}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_DecodedAPI_2}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_APIKey}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_APIKeyBack}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_APIKey_2}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_APIKeyBack_2}", "", RegexOptions.IgnoreCase)

        Dim result As String = template

        If Regex.IsMatch(template, "{Dictionary}", RegexOptions.IgnoreCase) Then
            Dictionary = Global.SharedLibrary.SharedLibrary.SharedMethods.GetTranslationDictionaryText(_context, TranslateLanguage)
        End If

        Dim placeholderPattern As String = "\{([^}]+)\}"
        Dim matches As MatchCollection = Regex.Matches(template, placeholderPattern)

        For Each m As Match In matches
            Dim placeholder As String = m.Value          ' Example: "{Name}"
            Dim varName As String = m.Groups(1).Value    ' Example: "Name"

            ' Search for Field
            Dim fieldInfo = Me.GetType().GetField(varName)
            If fieldInfo IsNot Nothing Then
                Dim fieldValue = fieldInfo.GetValue(Me)
                If fieldValue IsNot Nothing Then
                    result = result.Replace(placeholder, fieldValue.ToString())
                End If
                Continue For
            End If

            ' Search for Property
            Dim propInfo = Me.GetType().GetProperty(varName)
            If propInfo IsNot Nothing Then
                Dim propValue = propInfo.GetValue(Me)
                If propValue IsNot Nothing Then
                    result = result.Replace(placeholder, propValue.ToString())
                End If
            End If
        Next

        Return result
    End Function


End Class
