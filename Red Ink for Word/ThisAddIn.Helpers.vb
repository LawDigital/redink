' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Helpers.vb
' Purpose: Provides utility helper methods for the Red Ink for Word add-in,
'          including undo scope management, escape sequence handling, document
'          gathering, text interpolation, progress tracking, API key encoding,
'          and VBA module validation.
'
' Architecture:
'  - WordUndoScope: Manages Word custom undo records for VSTO operations (Word 2013+).
'  - Escape Handling: HideEscape/UnHideEscape convert special characters to/from Unicode sequences.
'  - Document Gathering: GatherSelectedDocuments collects open Word documents with UI selection.
'  - Text Search: FindLongTextInChunks delegates to WordSearchHelper for anchored text searches.
'  - Word Counting: GetSelectedTextLength uses regex to count real words (letters with internal punctuation).
'  - Runtime Interpolation: InterpolateAtRuntime replaces placeholders with instance field/property values.
'  - ProgressScope: IDisposable wrapper for modeless progress windows with cancellation support.
'  - Language Detection: GetWordDefaultInterfaceLanguage retrieves Word UI language via CultureInfo.
'  - API Key Encoding: CodeAPIKey/DeCodeAPIKey encrypt/decrypt API keys with optional prefix handling.
'  - VBA Validation: VBAModuleWorking checks if required VBA helper module meets minimum version.
'  - Win32 Interop: GetAsyncKeyState for keyboard state detection (VK_ESCAPE).
'
' Dependencies: SharedLibrary.SharedMethods, WordSearchHelper, ProgressBarModule, 
'               DPIProgressForm/ProgressForm.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Diagnostics
Imports System.Globalization
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Interop.Word
Imports NetOffice.PowerPointApi
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods


Partial Public Class ThisAddIn

    ''' <summary>
    ''' Opens or activates the existing Translator widget.
    ''' This is the same widget used by the Freestyle command "translator".
    ''' </summary>
    Public Sub ShowTranslatorWidget()
        ShowQuickTranslate()
    End Sub


    ''' <summary>
    ''' Checks if INI configuration loading has failed and attempts initialization if needed.
    ''' Returns True if INI loading ultimately failed, False if successful.
    ''' </summary>
    Public Function INILoadFail() As Boolean
        If Not INIloaded Then
            If Not StartupInitialized Then
                DelayedStartupTasks()
                RemoveStartupHandlers()
                If Not INIloaded Then Return True
                Return False
            Else
                InitializeConfig(False, False)
                If Not INIloaded Then
                    Return True
                End If
                Return False
            End If
        Else
            Return False
        End If
    End Function


    ''' <summary>
    ''' Checks whether the active document is in an editable state.
    ''' Returns True if the document can be edited, False otherwise.
    ''' If not editable, displays a message informing the user that editing must be enabled for Red Ink.
    ''' </summary>
    ''' <param name="silent">If True, suppresses the user-facing message box.</param>
    ''' <returns>True if the active document is editable, False otherwise.</returns>
    Public Function IsDocumentEditable(Optional silent As Boolean = False) As Boolean
        Try
            Dim app As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
            Dim doc As Microsoft.Office.Interop.Word.Document = Nothing

            Try
                doc = app.ActiveDocument
            Catch
                If Not silent Then
                    ShowCustomMessageBox($"No document is open or editing is not enabled. Please open a document or enable editing on your document to use {AN}.")
                End If
                Return False
            End Try

            If doc Is Nothing Then
                If Not silent Then
                    ShowCustomMessageBox($"No document is open or editing is not enabled. Please open a document or enable editing on your document to use {AN}.")
                End If
                Return False
            End If

            ' Check if the document is read-only
            'If doc.ReadOnly Then
            'If Not silent Then
            'ShowCustomMessageBox($"The document '{doc.Name}' is opened as read-only. To use {AN}, editing must be enabled." & vbCrLf & vbCrLf &
            '                             "Please enable editing (e.g., click 'Enable Editing' in the message bar, uncheck read-only in File > Info, or save a writable copy).")
            'End If
            'Return False
            'End If

            ' Check if the document is protected (restricted editing)
            'If doc.ProtectionType <> Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection Then
            'If Not silent Then
            'ShowCustomMessageBox($"The document '{doc.Name}' has restricted editing enabled. To use {AN}, editing restrictions must be removed." & vbCrLf & vbCrLf &
            '                             "Please go to Review > Restrict Editing and stop protection.")
            'End If
            'Return False
            'End If

            ' Check if the document is marked as final
            'Try
            'If CBool(doc.Final) Then
            'If Not silent Then
            'ShowCustomMessageBox($"The document '{doc.Name}' is marked as final. To use {AN}, editing must be enabled." & vbCrLf & vbCrLf &
            '                                 "Please go to File > Info and click 'Edit Anyway'.")
            'End If
            'Return False
            'End If
            'Catch
            ' Final property may not be available in older versions — ignore
            'End Try

            Return True

        Catch ex As System.Exception
            If Not silent Then
                ShowCustomMessageBox($"Could not determine document edit status: {ex.Message}")
            End If
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Manages a custom Word undo record scope for VSTO operations.
    ''' Automatically starts a custom undo record on creation and ends it on disposal.
    ''' Only supported in Word 2013 (version 15.0) and later.
    ''' </summary>
    Friend NotInheritable Class WordUndoScope
        Implements System.IDisposable

        Private ReadOnly _app As Microsoft.Office.Interop.Word.Application
        Private ReadOnly _undo As Microsoft.Office.Interop.Word.UndoRecord
        Private ReadOnly _iStarted As System.Boolean

        ''' <summary>
        ''' Initializes a new undo scope. Starts a custom undo record if Word version supports it
        ''' and no other custom record is currently active.
        ''' </summary>
        ''' <param name="app">The Word Application instance.</param>
        ''' <param name="name">Optional custom name for the undo record. Defaults to "VSTO-Action".</param>
        Public Sub New(app As Microsoft.Office.Interop.Word.Application, Optional name As System.String = Nothing)
            _app = app
            _undo = _app.UndoRecord

            ' Word < 2013 (Version < 15.0) does not have UndoRecord.
            Dim ver As System.Version = New System.Version(_app.Version)
            If ver.Major < 15 Then
                Return
            End If

            ' Only start if no other custom record is currently running.
            If Not _undo.IsRecordingCustomRecord Then
                If name IsNot Nothing AndAlso name.Length > 0 Then
                    _undo.StartCustomRecord(name)
                Else
                    _undo.StartCustomRecord("VSTO-Action")
                End If
                _iStarted = True
            End If
        End Sub

        ''' <summary>
        ''' Ends the custom undo record if it was started by this scope.
        ''' </summary>
        Public Sub Dispose() Implements System.IDisposable.Dispose
            Try
                If _iStarted AndAlso _undo.IsRecordingCustomRecord Then
                    _undo.EndCustomRecord()
                End If
            Catch ex As System.Exception
                ' Do not throw exceptions in Dispose.
            End Try
        End Sub
    End Class

    ''' <summary>
    ''' Replaces each sequence \\X with \\uXXXX (double backslash).
    ''' Example: \\; becomes \\u003B, \\&lt; becomes \\u003C, etc.
    ''' </summary>
    Public Function HideEscape(ByVal input As String) As String
        Return System.Text.RegularExpressions.Regex.Replace(input, "\\\\(.)",
            Function(m As System.Text.RegularExpressions.Match) As String
                Dim c As Char = m.Groups(1).Value(0)
                Dim hex As String = System.Convert.ToInt32(c).ToString("X4")
                Return "\\u" & hex
            End Function)
    End Function

    ''' <summary>
    ''' Replaces each sequence \\uXXXX (double backslash) back to the corresponding character.
    ''' Example: \\u003B becomes ;, \\u003C becomes &lt;, etc.
    ''' </summary>
    Public Function UnHideEscape(ByVal input As String) As String
        Return System.Text.RegularExpressions.Regex.Replace(input, "\\\\u([0-9A-Fa-f]{4})",
            Function(m As System.Text.RegularExpressions.Match) As String
                Dim code As Integer = Integer.Parse(m.Groups(1).Value, System.Globalization.NumberStyles.HexNumber)
                Return System.Convert.ToChar(code).ToString()
            End Function)
    End Function

    ''' <summary>
    ''' Gathers selected Word documents from currently open documents.
    ''' Optionally prompts the user to select one, all, or none.
    ''' Returns formatted document content with XML-style tags, or special values "NONE", empty string, or "ERROR ...".
    ''' </summary>
    ''' <param name="IncludeName">Include document name in output header.</param>
    ''' <param name="IncludeNone">Offer "Do not add any document" option in selection.</param>
    ''' <param name="ExceptCurrent">Exclude the currently active document from the list.</param>
    ''' <param name="SilentAndGetAll">Skip user prompt and return all documents (after optional exclusion).</param>
    Public Function GatherSelectedDocuments(Optional IncludeName As Boolean = True,
                                            Optional IncludeNone As System.Boolean = False,
                                            Optional ExceptCurrent As Boolean = False,
                                            Optional SilentAndGetAll As Boolean = False) As System.String
        Try
            Dim app As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application

            ' Collect all open documents (unique by FullName/Name to avoid duplicates from multiple windows)
            Dim docList As New System.Collections.Generic.List(Of Microsoft.Office.Interop.Word.Document)()
            Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)

            For Each doc As Microsoft.Office.Interop.Word.Document In app.Documents
                Dim key As System.String = If(Not System.String.IsNullOrEmpty(doc.FullName), doc.FullName, doc.Name)
                If Not seen.Contains(key) Then
                    seen.Add(key)
                    docList.Add(doc)
                End If
            Next

            ' Optionally exclude the currently active document
            If ExceptCurrent Then
                Dim activeDoc As Microsoft.Office.Interop.Word.Document = Nothing
                Try
                    activeDoc = app.ActiveDocument
                Catch
                    activeDoc = Nothing
                End Try
                If activeDoc IsNot Nothing Then
                    For i As System.Int32 = docList.Count - 1 To 0 Step -1
                        If System.Object.ReferenceEquals(docList(i), activeDoc) Then
                            docList.RemoveAt(i)
                        End If
                    Next
                End If
            End If

            If docList.Count = 0 Then
                Return "NONE"
            End If

            ' If silent mode requested: return all (after optional exclusion) without prompting
            If SilentAndGetAll Then
                Return BuildDocumentsResult(docList, IncludeName)
            End If

            ' Build selection items for each open document
            Dim selItems As New System.Collections.Generic.List(Of SelectionItem)()
            For i As System.Int32 = 0 To docList.Count - 1
                Dim d As Microsoft.Office.Interop.Word.Document = docList(i)
                selItems.Add(New SelectionItem($"{d.Name} ({d.FullName})", i + 1))
            Next

            ' Add "All open documents" and optional "None"
            Dim indexAll As System.Int32 = selItems.Count + 1
            selItems.Add(New SelectionItem("Add all open documents", indexAll))

            Dim indexNone As System.Int32 = -1
            If IncludeNone Then
                indexNone = indexAll + 1
                selItems.Add(New SelectionItem("Do not add any document", indexNone))
            End If

            ' Prompt user (default/highlight on "All")
            Dim itemsArray As SelectionItem() = selItems.ToArray()
            Dim picked As System.Int32 = SelectValue(itemsArray, indexAll, "Choose document to add …")

            ' User cancelled or invalid choice
            If picked < 1 Then
                Return System.String.Empty
            End If

            ' User explicitly chose "None"
            If IncludeNone AndAlso picked = indexNone Then
                Return System.String.Empty
            End If

            ' Determine targets based on selection
            Dim targets As New System.Collections.Generic.List(Of Microsoft.Office.Interop.Word.Document)()
            If picked = indexAll Then
                targets.AddRange(docList)
            Else
                If picked - 1 >= 0 AndAlso picked - 1 < docList.Count Then
                    targets.Add(docList(picked - 1))
                Else
                    Return System.String.Empty
                End If
            End If

            Return BuildDocumentsResult(targets, IncludeName)

        Catch ex As System.Exception
            Return "ERROR " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' Builds the concatenated document content string with XML-style tags.
    ''' Each document is wrapped in &lt;DOCUMENTn&gt; tags with optional name header.
    ''' </summary>
    Private Function BuildDocumentsResult(docs As System.Collections.Generic.List(Of Microsoft.Office.Interop.Word.Document),
                                          includeName As System.Boolean) As System.String
        Dim insertedDocuments As System.String = System.String.Empty
        Dim tagIndex As System.Int32 = 1

        For Each d As Microsoft.Office.Interop.Word.Document In docs
            If includeName Then insertedDocuments &= $"Here follows document no. {tagIndex} with the name '" & d.Name & "': " & vbCrLf
            insertedDocuments &= $"<DOCUMENT{tagIndex}>" & vbCrLf
            insertedDocuments &= d.Content.Text & vbCrLf
            insertedDocuments &= $"</DOCUMENT{tagIndex}>" & vbCrLf
            tagIndex += 1
        Next

        If System.String.IsNullOrEmpty(insertedDocuments) Then
            ShowCustomMessageBox("No content could be retrieved from the selected document(s).")
            Return System.String.Empty
        End If

        Return insertedDocuments
    End Function

    ''' <summary>
    ''' Finds long text in the current selection using anchored fast search.
    ''' Delegates to WordSearchHelper.FindLongTextAnchoredFast with optional skip-deleted-text support.
    ''' </summary>
    ''' <param name="findText">The text to find.</param>
    ''' <param name="selection">The Word selection to search within (modified if found).</param>
    ''' <param name="Skipdeleted">If True, skips text marked as deleted revisions.</param>
    ''' <returns>True if text was found, False otherwise.</returns>
    Public Function FindLongTextInChunks(ByVal findText As String, ByRef selection As Word.Selection, Optional Skipdeleted As Boolean = True) As Boolean

        Debug.WriteLine("Entering into FindLongTextAnchoredFast")

        Dim answer As Boolean = WordSearchHelper.FindLongTextAnchoredFast(selection, findText, Skipdeleted)

        Debug.WriteLine("Text found: " & "'" & selection.Text & "'")

        Return answer

    End Function


    ''' <summary>
    ''' Counts real words in the current selection: sequences of letters (Unicode) optionally joined by internal apostrophes or hyphens; numeric/mixed tokens are ignored.
    ''' </summary>
    Private Function GetSelectedTextLength() As Integer
        Try
            Dim wordApp As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
            Dim selection As Microsoft.Office.Interop.Word.Selection = wordApp.Selection

            Dim selectedText As String = selection.Text
            If String.IsNullOrWhiteSpace(selectedText) Then
                Return 0
            End If

            ' Pattern:
            ' \b                Word boundary
            ' [\p{L}]+          One or more Unicode letters
            ' (?:[''\-‑–][\p{L}]+)*  Optional internal apostrophe/hyphen/dash + letters (e.g. don't, mother-in-law, rock'n'roll)
            ' \b                Word boundary
            ' Excludes tokens containing digits or starting with punctuation.
            Dim pattern As String = "\b[\p{L}]+(?:[''\-‑–][\p{L}]+)*\b"

            Return Regex.Matches(selectedText, pattern).Count
        Catch ex As System.Exception
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' Replaces placeholders in the template string with values from instance fields or properties.
    ''' Placeholders use {FieldName} or {PropertyName} syntax. Sensitive placeholders like {Codebasis}
    ''' and {INI_*API*} are cleared for security reasons.
    ''' </summary>
    ''' <param name="template">The template string containing placeholders.</param>
    ''' <returns>The interpolated string with placeholders replaced by actual values.</returns>
    Public Function InterpolateAtRuntime(ByVal template As String) As String
        If template Is Nothing Then
            MessageBox.Show("Error InterpolateAtRuntime: Template is Nothing.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return ""
        End If

        ' Clear sensitive placeholders
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
            Dim placeholder As String = m.Value          ' e.g. "{Name}"
            Dim varName As String = m.Groups(1).Value    ' e.g. "Name"

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


    ''' <summary>
    ''' Lightweight IDisposable scope to show a modeless progress window with cancellation support
    ''' for long-running operations. Integrates with ProgressBarModule and DPIProgressForm/ProgressForm.
    ''' The progress form runs on a dedicated STA UI thread with its own message loop.
    ''' </summary>
    Public NotInheritable Class ProgressScope
        Implements IDisposable

        Private ReadOnly _cts As System.Threading.CancellationTokenSource = New System.Threading.CancellationTokenSource()
        Private _uiThread As System.Threading.Thread
        Private _form As System.Windows.Forms.Form
        Private ReadOnly _useDpiForm As Boolean
        Private _closed As Integer = 0

        ''' <summary>
        ''' Initializes the progress scope and displays a modeless progress window.
        ''' The form updates its display by polling ProgressBarModule.* every 250ms.
        ''' </summary>
        ''' <param name="headerText">The header text displayed on the progress form.</param>
        ''' <param name="initialLabel">The initial status label text.</param>
        ''' <param name="max">Maximum progress value (default 100).</param>
        ''' <param name="useDpiForm">If True, uses DPIProgressForm; otherwise uses ProgressForm.</param>
        Public Sub New(headerText As String,
                   initialLabel As String,
                   Optional max As Integer = 100,
                   Optional useDpiForm As Boolean = True)
            _useDpiForm = useDpiForm

            ' Initialize global progress state
            ProgressBarModule.CancelOperation = False
            ProgressBarModule.GlobalProgressMax = System.Math.Max(1, max)
            ProgressBarModule.GlobalProgressValue = 0
            ProgressBarModule.GlobalProgressLabel = If(initialLabel, "")

            ' Spin up a dedicated STA UI thread with its own message loop for the progress form
            _uiThread = New System.Threading.Thread(
            Sub()
                Try
                    System.Windows.Forms.Application.EnableVisualStyles()
                    _form = If(_useDpiForm,
                               CType(New DPIProgressForm(headerText, initialLabel), System.Windows.Forms.Form),
                               CType(New ProgressForm(headerText, initialLabel), System.Windows.Forms.Form))

                    ' Run form (timer inside form pulls ProgressBarModule.* and closes itself when CancelOperation=True)
                    System.Windows.Forms.Application.Run(_form)
                Catch
                    ' Swallow exceptions — cleanup is always attempted in Dispose.
                End Try
            End Sub
        )
            _uiThread.IsBackground = True
            _uiThread.SetApartmentState(System.Threading.ApartmentState.STA)
            _uiThread.Start()
        End Sub

        ''' <summary>
        ''' Reports progress in a thread-safe manner via the global ProgressBarModule.
        ''' </summary>
        ''' <param name="current">Current progress value.</param>
        ''' <param name="max">Optional new maximum value (if >= 1).</param>
        ''' <param name="label">Optional new status label text.</param>
        Public Shared Sub Report(current As Integer,
                             Optional max As Integer = -1,
                             Optional label As String = Nothing)
            If max >= 1 Then ProgressBarModule.GlobalProgressMax = max
            If label IsNot Nothing Then ProgressBarModule.GlobalProgressLabel = label
            ProgressBarModule.GlobalProgressValue = System.Math.Max(0, System.Math.Min(current, ProgressBarModule.GlobalProgressMax))
        End Sub

        ''' <summary>
        ''' Requests cancellation of the operation.
        ''' Also triggered by the Cancel button in the progress UI.
        ''' </summary>
        Public Sub RequestCancel()
            _cts.Cancel()
            ProgressBarModule.CancelOperation = True
        End Sub

        ''' <summary>
        ''' Gets a value indicating whether cancellation has been requested.
        ''' Check this frequently at safe points (between steps/chunks) to bail out early.
        ''' </summary>
        Public ReadOnly Property CancelRequested As Boolean
            Get
                Return ProgressBarModule.CancelOperation OrElse _cts.IsCancellationRequested
            End Get
        End Property

        ''' <summary>
        ''' Gets the CancellationToken associated with this scope.
        ''' </summary>
        Public ReadOnly Property Token As System.Threading.CancellationToken
            Get
                Return _cts.Token
            End Get
        End Property

        ''' <summary>
        ''' Disposes the progress scope, signaling cancellation and closing the progress form.
        ''' Waits up to 1 second for the UI thread to exit cleanly.
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            If System.Threading.Interlocked.Exchange(_closed, 1) <> 0 Then Return
            Try
                ' Signal cancel so the form's timer path also closes itself
                ProgressBarModule.CancelOperation = True

                Dim f = _form
                If f IsNot Nothing Then
                    Try
                        If f.IsHandleCreated AndAlso Not f.IsDisposed Then
                            ' Request close on the form's thread
                            f.BeginInvoke(New System.Action(Sub()
                                                                Try
                                                                    If Not f.IsDisposed Then f.Close()
                                                                Catch
                                                                End Try
                                                            End Sub))
                        End If
                    Catch
                        ' Ignore cross-thread or shutdown races
                    End Try
                End If
            Finally
                ' Give the UI thread a moment to exit cleanly
                Try
                    If _uiThread IsNot Nothing AndAlso _uiThread.IsAlive Then
                        If Not _uiThread.Join(1000) Then
                            Try : _uiThread.Interrupt() : Catch : End Try
                        End If
                    End If
                Catch
                End Try
            End Try
        End Sub
    End Class

    ''' <summary>
    ''' Retrieves the display name of the Word user interface language.
    ''' Falls back to "English" if an error occurs.
    ''' </summary>
    Public Function GetWordDefaultInterfaceLanguage() As String
        Try
            ' Get the language ID of the Word user interface
            Dim uiLanguageID As Integer = Globals.ThisAddIn.Application.LanguageSettings.LanguageID(MsoAppLanguageID.msoLanguageIDUI)

            ' Convert the language ID to a human-readable name
            Dim cultureInfo As Globalization.CultureInfo = New Globalization.CultureInfo(uiLanguageID)

            ' Return the language display name
            Return cultureInfo.DisplayName
        Catch ex As System.Exception
            Return "English"
        End Try
    End Function

    ''' <summary>
    ''' Encodes (encrypts) an API key using a user-provided secret key.
    ''' Optionally handles a prefix (e.g., "sk-") that is preserved during encoding.
    ''' Prompts the user for the prefix and secret key via input dialogs.
    ''' </summary>
    ''' <param name="apiKey">The API key to encode.</param>
    ''' <returns>The encoded API key with prefix (if applicable), or "Error" on failure.</returns>
    Private Function CodeAPIKey(ByVal apiKey As String, Optional ByVal UseStrongAlgorithm As Boolean = True) As String
        Dim modifiedKey As String
        Dim resultKey As String
        Dim xcodebasis As String
        Dim HadPrefix As Boolean = False

        Dim PrefixValue As String = INI_APIKeyPrefix

        ' Check if an API key is provided
        apiKey = apiKey.Trim()
        If String.IsNullOrEmpty(apiKey) Then
            ShowCustomMessageBox("No text selected to encode. Select the API Key you wish to encode.")
            Return "Error"
        End If

        PrefixValue = SLib.ShowCustomInputBox("Please enter the API key prefix (as used in the configuration file, if any):", "API Key Encryptor", True, PrefixValue)

        xcodebasis = SLib.ShowCustomInputBox("Please enter the secret key:", "API Key Encryptor", True)
        If String.IsNullOrEmpty(xcodebasis) Then
            ShowCustomMessageBox("No secret key entered.")
            Return "Error"
        End If

        ' Check if the API key has the prefix
        If Not String.IsNullOrEmpty(PrefixValue) AndAlso apiKey.StartsWith(PrefixValue) Then
            HadPrefix = True
            ' Encrypt only the part after the prefix
            modifiedKey = apiKey.Substring(PrefixValue.Length)
        Else
            ' Encrypt the entire key if no prefix is present
            modifiedKey = apiKey
        End If

        ' Encrypt the modified key (without the prefix)
        resultKey = CodeString(modifiedKey, xcodebasis, UseStrongAlgorithm)

        ' Add the prefix back if it was present
        If HadPrefix Then
            resultKey = PrefixValue & resultKey
        End If

        Return resultKey
    End Function

    ''' <summary>
    ''' Decodes (decrypts) an API key using a user-provided secret key.
    ''' Optionally handles a prefix (e.g., "sk-") that is preserved during decoding.
    ''' Prompts the user for the prefix and secret key via input dialogs.
    ''' </summary>
    ''' <param name="apiKey">The encoded API key to decode.</param>
    ''' <returns>The decoded API key with prefix (if applicable), or "Error" on failure.</returns>
    Private Function DeCodeAPIKey(ByVal apiKey As String) As String
        Dim modifiedKey As String
        Dim resultKey As String
        Dim xcodebasis As String

        Dim PrefixValue As String = INI_APIKeyPrefix

        ' Check if an API key is provided
        apiKey = apiKey.Trim()
        If String.IsNullOrEmpty(apiKey) Then
            ShowCustomMessageBox("No text selected to decode. Select the API Key you wish to decode.")
            Return "Error"
        End If

        PrefixValue = SLib.ShowCustomInputBox("Please enter the API key prefix (as used in the configuration file, if any):", "API Key Decryptor", True, PrefixValue)

        xcodebasis = SLib.ShowCustomInputBox("Please enter the secret key:", "API Key Decryptor", True)
        If String.IsNullOrEmpty(xcodebasis) Then
            ShowCustomMessageBox("No secret key entered.")
            Return "Error"
        End If

        ' Check if the key starts with the prefix
        If Not String.IsNullOrEmpty(PrefixValue) AndAlso apiKey.StartsWith(PrefixValue) Then
            ' Decrypt only the part after the prefix
            modifiedKey = apiKey.Substring(PrefixValue.Length)
        Else
            ' Decrypt the entire key if no prefix is present
            modifiedKey = apiKey
        End If

        ' Decrypt the modified key (without the prefix)
        resultKey = DecodeString(modifiedKey, xcodebasis)

        ' Add the prefix back only if it was in the original key
        If Not String.IsNullOrEmpty(PrefixValue) AndAlso apiKey.StartsWith(PrefixValue) Then
            resultKey = PrefixValue & resultKey
        End If

        Return resultKey
    End Function

    ''' <summary>
    ''' Validates that the VBA helper module is working and meets the minimum required version.
    ''' Calls the VBA function "CheckAppHelper" and compares its version number.
    ''' </summary>
    ''' <returns>True if the VBA module version meets MinHelperVersion, False otherwise.</returns>
    Public Function VBAModuleWorking() As Boolean

        Dim xlApp As Microsoft.Office.Interop.Word.Application = Me.Application

        Try
            ' Call the VBA function
            Dim HelperVersion As Integer = CType(xlApp.Run("CheckAppHelper"), Integer)

            If HelperVersion >= MinHelperVersion Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            ' Return False if VBA call fails
            Return False
        End Try

    End Function

    ''' <summary>
    ''' Win32 API import: Retrieves the asynchronous state of a specified virtual key.
    ''' Used for detecting key presses (e.g., Escape key) outside normal event flow.
    ''' </summary>
    ''' <param name="vKey">The virtual-key code to check.</param>
    ''' <returns>Non-zero if the key is pressed, zero otherwise.</returns>
    <System.Runtime.InteropServices.DllImport("user32.dll",
    SetLastError:=True, CharSet:=System.Runtime.InteropServices.CharSet.Auto)>
    Public Shared Function GetAsyncKeyState(ByVal vKey As System.Int32) As System.Int16
    End Function

    ''' <summary>
    ''' Virtual-key code for the Escape key (0x1B).
    ''' Used with GetAsyncKeyState for cancellation detection.
    ''' </summary>
    Private Const VK_ESCAPE As System.Int32 = &H1B


    ' Serial Number Calculator

    ''' <summary>
    ''' Base date used for serial encoding/decoding. The serial stores the number of days since this date (UTC).
    ''' </summary>
    Private Shared ReadOnly BaseDate As DateTime = New DateTime(2000, 1, 1, 0, 0, 0, DateTimeKind.Utc)

    ''' <summary>
    ''' Number of bits allocated to the date component (days since <see cref="BaseDate"/>).
    ''' </summary>
    Private Const DateBits As Integer = 17

    ''' <summary>
    ''' Number of bits allocated to the numeric payload.
    ''' </summary>
    Private Const NumberBits As Integer = 10

    ''' <summary>
    ''' Number of bits allocated to the signature (integrity check).
    ''' </summary>
    Private Const SigBits As Integer = 13

    ''' <summary>
    ''' Total number of bits stored in the token.
    ''' </summary>
    Private Const TokenBits As Integer = DateBits + NumberBits + SigBits '40

    ''' <summary>
    ''' Fixed token length in Crockford Base32 characters.
    ''' 40 bits / 5 bits per char = 8 chars.
    ''' </summary>
    Private Const TokenChars As Integer = 8

    ''' <summary>
    ''' Maximum encodable day count (inclusive) based on <see cref="DateBits"/>.
    ''' </summary>
    Private Const MaxDays As Integer = (1 << DateBits) - 1           '131071

    ''' <summary>
    ''' Maximum encodable number value (inclusive) based on <see cref="NumberBits"/>.
    ''' </summary>
    Private Const MaxNumber As Integer = (1 << NumberBits) - 1       '1023

    ''' <summary>
    ''' Bitmask used to extract the signature field from the packed token.
    ''' </summary>
    Private Const SigMask As Long = (1L << SigBits) - 1L

    ''' <summary>
    ''' Prompts user for a date and a number, encodes them into a Base32 token, writes the token to the document,
    ''' demonstrates decoding, and copies the token to the clipboard.
    ''' </summary>
    ''' <param name="Selection">The Word selection used as an insertion point for output.</param>
    Private Shared Sub EncodeSerial(Selection As Word.Selection)

        Dim dateText As String = ShowCustomInputBox("Enter date (yyyy-MM-dd):", $"{AN} Serial Key Encoder", True)

        Dim inputDate As DateTime
        If Not DateTime.TryParseExact(dateText, "yyyy-MM-dd",
                                    System.Globalization.CultureInfo.InvariantCulture,
                                    DateTimeStyles.None,
                                    inputDate) Then
            ShowCustomMessageBox("Invalid date format.")
            Return
        End If

        Dim numText As String = ShowCustomInputBox("Enter number (0.." & MaxNumber.ToString(CultureInfo.InvariantCulture) & "):", $"{AN} Serial Key Encoder", True)

        Dim inputNumber As Integer
        If Not Integer.TryParse(numText, NumberStyles.Integer, CultureInfo.InvariantCulture, inputNumber) Then
            ShowCustomMessageBox("Invalid number.")
            Return
        End If

        If inputNumber < 0 OrElse inputNumber > MaxNumber Then
            ShowCustomMessageBox("Number out of range.")
            Return
        End If

        Dim token As String = EncodeToken(inputDate, inputNumber)

        ' Demonstrate decoding
        Dim decodedDate As DateTime = DecodeDate(token)
        Dim decodedNumber As Integer = DecodeNumber(token)

        Selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
        Selection.TypeText(vbCrLf & "Encoded key (also in clipboard):" & vbCrLf & token & vbCrLf)

        Try
            Selection.TypeText(vbCrLf & "Decoded date: " & decodedDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
            Selection.TypeText(vbCrLf & "Decoded number: " & decodedNumber.ToString(CultureInfo.InvariantCulture))
        Catch ex As System.Exception
            Selection.TypeText(vbCrLf & "Token invalid: " & ex.Message)
        End Try

        Selection.ParagraphFormat.Hyphenation = CInt(False)
        SLib.PutInClipboard(token)

    End Sub

    ''' <summary>
    ''' Prompts for a serial token, decodes it (including signature validation), and writes the decoded date and number
    ''' to the document using <see cref="Word.Selection.TypeText"/>.
    ''' </summary>
    ''' <param name="Selection">The Word selection used as an insertion point for output.</param>
    Private Shared Sub DecodeSerial(Selection As Word.Selection)
        Dim token As String = ShowCustomInputBox("Enter serial (8 chars, Crockford Base32):", $"{AN} Serial Key Decoder", True)

        If String.IsNullOrWhiteSpace(token) Then
            ShowCustomMessageBox("No serial entered.")
            Return
        End If

        token = token.Trim()

        Selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
        Selection.TypeText(vbCrLf & "Serial to decode:" & vbCrLf & token & vbCrLf)

        Try
            Dim decodedDate As DateTime = DecodeDate(token)
            Dim decodedNumber As Integer = DecodeNumber(token)

            Selection.TypeText(vbCrLf & "Decoded date: " & decodedDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
            Selection.TypeText(vbCrLf & "Decoded number: " & decodedNumber.ToString(CultureInfo.InvariantCulture) & vbCrLf)
        Catch ex As System.Exception
            Selection.TypeText(vbCrLf & "Token invalid: " & ex.Message & vbCrLf)
        End Try

        Selection.ParagraphFormat.Hyphenation = CInt(False)
    End Sub


    ''' <summary>
    ''' Encodes a date and number into an 8-character Crockford Base32 token with an embedded signature.
    ''' </summary>
    ''' <param name="inputDate">Date to encode (interpreted at UTC midnight).</param>
    ''' <param name="number">Number to encode (0..MaxNumber).</param>
    ''' <returns>8-character Crockford Base32 token.</returns>
    Private Shared Function EncodeToken(ByVal inputDate As DateTime, ByVal number As Integer) As String
        ' Normalize date to UTC midnight
        Dim dUtc As DateTime = New DateTime(inputDate.Year, inputDate.Month, inputDate.Day, 0, 0, 0, DateTimeKind.Utc)
        Dim days As Integer = CInt((dUtc - BaseDate).TotalDays)

        If days < 0 OrElse days > MaxDays Then
            Throw New System.Exception("Date out of encodable range.")
        End If
        If number < 0 OrElse number > MaxNumber Then
            Throw New System.Exception("Number out of encodable range.")
        End If

        Dim signature As Integer = ComputeSignature(days, number)

        ' Pack into 40 bits:
        ' packed = [days << (NumberBits+SigBits)] | [number << SigBits] | [sig]
        Dim packed As Long = (CLng(days) << (NumberBits + SigBits)) Or
                            (CLng(number) << SigBits) Or
                            CLng(signature)

        Return Base32CrockfordEncode(packed, TokenChars)
    End Function

    ''' <summary>
    ''' Decodes the date component from a token (validating the signature).
    ''' </summary>
    ''' <param name="token">8-character token.</param>
    ''' <returns>Date (UTC midnight) represented by the token.</returns>
    Private Shared Function DecodeDate(ByVal token As String) As DateTime
        Dim days As Integer = 0
        Dim number As Integer = 0
        UnpackAndValidate(token, days, number)

        Dim result As DateTime = BaseDate.AddDays(days)
        Return New DateTime(result.Year, result.Month, result.Day, 0, 0, 0, DateTimeKind.Utc)
    End Function

    ''' <summary>
    ''' Decodes the numeric payload from a token (validating the signature).
    ''' </summary>
    ''' <param name="token">8-character token.</param>
    ''' <returns>Number represented by the token.</returns>
    Private Shared Function DecodeNumber(ByVal token As String) As Integer
        Dim days As Integer = 0
        Dim number As Integer = 0
        UnpackAndValidate(token, days, number)
        Return number
    End Function

    ''' <summary>
    ''' Unpacks the token into its date-days and number fields and validates the embedded signature.
    ''' </summary>
    ''' <param name="token">8-character Crockford Base32 token.</param>
    ''' <param name="days">Outputs decoded days since <see cref="BaseDate"/>.</param>
    ''' <param name="number">Outputs decoded numeric payload.</param>
    Private Shared Sub UnpackAndValidate(ByVal token As String, ByRef days As Integer, ByRef number As Integer)
        If token Is Nothing Then
            Throw New System.Exception("Token is null.")
        End If

        token = token.Trim().ToUpperInvariant()
        If token.Length <> TokenChars Then
            Throw New System.Exception("Token must be exactly " & TokenChars.ToString(CultureInfo.InvariantCulture) & " characters.")
        End If

        Dim packed As Long = Base32CrockfordDecode(token)

        ' Extract fields:
        Dim sig As Integer = CInt(packed And SigMask)
        Dim numberVal As Integer = CInt((packed >> SigBits) And ((1L << NumberBits) - 1L))
        Dim daysVal As Integer = CInt((packed >> (SigBits + NumberBits)) And ((1L << DateBits) - 1L))

        Dim expectedSig As Integer = ComputeSignature(daysVal, numberVal)
        If sig <> expectedSig Then
            Throw New System.Exception("Signature mismatch (token manipulated or wrong key).")
        End If

        days = daysVal
        number = numberVal
    End Sub

    ''' <summary>
    ''' Computes the signature (integrity check) over the date-days and number fields using HMAC-SHA256,
    ''' truncated to <see cref="SigBits"/> bits.
    ''' </summary>
    ''' <param name="days">Days since <see cref="BaseDate"/>.</param>
    ''' <param name="number">Numeric payload.</param>
    ''' <returns>Signature value truncated to <see cref="SigBits"/> bits.</returns>
    Private Shared Function ComputeSignature(ByVal days As Integer, ByVal number As Integer) As Integer
        ' HMAC(key, days||number) truncated to SigBits
        Dim keyBytes As Byte() = System.Text.Encoding.UTF8.GetBytes(SK)

        ' 4 bytes days, 2 bytes number
        Dim msg As Byte() = New Byte(5) {} '6 bytes total
        Dim dBytes As Byte() = BitConverter.GetBytes(days)
        If BitConverter.IsLittleEndian Then Array.Reverse(dBytes)
        Array.Copy(dBytes, 0, msg, 0, 4)

        Dim nBytes As Byte() = BitConverter.GetBytes(CShort(number))
        If BitConverter.IsLittleEndian Then Array.Reverse(nBytes)
        Array.Copy(nBytes, 0, msg, 4, 2)

        Dim hash As Byte()
        Using h As System.Security.Cryptography.HMACSHA256 = New System.Security.Cryptography.HMACSHA256(keyBytes)
            hash = h.ComputeHash(msg)
        End Using

        ' Take first 4 bytes as unsigned int, then mask down
        Dim val As UInteger = CUInt(hash(0)) << 24 Or
                            CUInt(hash(1)) << 16 Or
                            CUInt(hash(2)) << 8 Or
                            CUInt(hash(3))

        Dim sig As Integer = CInt(val And CUInt((1 << SigBits) - 1))
        Return sig
    End Function

    ''' <summary>
    ''' Crockford Base32 alphabet used for token encoding.
    ''' </summary>
    Private Const CrockfordAlphabet As String = "0123456789ABCDEFGHJKMNPQRSTVWXYZ"

    ''' <summary>
    ''' Encodes a non-negative integer value to Crockford Base32 with a fixed character length.
    ''' </summary>
    ''' <param name="value">Value to encode (must be non-negative).</param>
    ''' <param name="length">Fixed output length in characters.</param>
    ''' <returns>Base32 encoded string of exactly <paramref name="length"/> characters.</returns>
    Private Shared Function Base32CrockfordEncode(ByVal value As Long, ByVal length As Integer) As String
        If value < 0 Then Throw New System.Exception("Value must be non-negative.")

        Dim chars As Char() = New Char(length - 1) {}
        Dim v As Long = value

        For i As Integer = length - 1 To 0 Step -1
            Dim idx As Integer = CInt(v And 31L) ' 5 bits
            chars(i) = CrockfordAlphabet(idx)
            v >>= 5
        Next

        If v <> 0 Then
            Throw New System.Exception("Value does not fit into " & length.ToString(CultureInfo.InvariantCulture) & " Base32 chars.")
        End If

        Return New String(chars)
    End Function

    ''' <summary>
    ''' Decodes a Crockford Base32 token to its numeric value.
    ''' </summary>
    ''' <param name="token">Token to decode.</param>
    ''' <returns>Decoded numeric value.</returns>
    Private Shared Function Base32CrockfordDecode(ByVal token As String) As Long
        Dim v As Long = 0

        For Each ch As Char In token
            v <<= 5
            v = v Or CrockfordValue(ch)
        Next

        Return v
    End Function

    ''' <summary>
    ''' Converts a single Crockford Base32 character into a 5-bit value.
    ''' Applies Crockford normalization rules (O->0, I/L->1; case-insensitive).
    ''' </summary>
    ''' <param name="ch">Character to decode.</param>
    ''' <returns>Value in range 0..31.</returns>
    Private Shared Function CrockfordValue(ByVal ch As Char) As Integer
        ' Crockford decoding rules: accept lowercase, treat O as 0, I/L as 1
        Dim c As Char = Char.ToUpperInvariant(ch)

        If c = "O"c Then c = "0"c
        If c = "I"c OrElse c = "L"c Then c = "1"c

        Dim idx As Integer = CrockfordAlphabet.IndexOf(c)
        If idx < 0 Then
            Throw New System.Exception("Invalid Base32 character: " & ch)
        End If
        Return idx
    End Function



End Class
