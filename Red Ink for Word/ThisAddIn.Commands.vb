' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Commands.vb
' Purpose: Implements command handlers for the Red Ink Word add-in, providing
'          AI-powered text processing operations including translation, correction,
'          improvement, anonymization, summarization, and style analysis.
'
' Architecture:
'  - Command Methods: Each public Sub corresponds to a ribbon button or context menu action
'  - INI Configuration: Commands check INILoadFail() and use INI_* settings for behavior
'  - Async Processing: Most commands use ProcessSelectedText() for asynchronous LLM calls
'  - Markup Support: Multiple markup methods (Word Track Changes, Diff, Regex) for showing changes
'  - Format Preservation: HTML/Markdown encoding to maintain text formatting through LLM processing
'  - Override System: Settings support personal overrides (e.g., INI_ReplaceText2Override)
'  - Model Selection: Support for alternate models via INI_AlternateModelPath
'  - MyStyle System: User writing style profiles stored in prompt files
'  - Anonymization: File-based or prompt-based entity replacement
'  - Dialog Integration: Custom input boxes, message boxes, and progress indicators
'
' Dependencies:
'  - SharedLibrary.SharedLibrary.SharedMethods: UI helpers, LLM communication, text processing
'  - Whisper.net.LibraryLoader: Speech-to-text functionality
'  - DocumentFormat.OpenXml: Document format handling
'  - Microsoft.Office.Interop.Word: Word automation
'
' Key Methods:
'  - Translation: InLanguage1(), InLanguage2(), InOther()
'  - Text Enhancement: Correct(), Improve(), Friendly(), Convincing(), NoFillers()
'  - Analysis: CheckDocumentII(), Explain(), SuggestTitles()
'  - Transformation: Anonymize(), Shorten(), Filibuster(), SwitchParty()
'  - Style: ApplyMyStyle(), DefineMyStyle()
'  - Utilities: ShowSettings(), Transcriptor(), ShowChatForm()
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.Windows.Forms
Imports DocumentFormat.OpenXml.Drawing
Imports DocumentFormat.OpenXml.Wordprocessing
Imports Google.Cloud.Speech.V1.LanguageCodes
Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Interop.Word
Imports NetOffice.PowerPointApi
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports Whisper.net.LibraryLoader
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn


    ''' <summary>
    ''' Translates selected text to the primary language configured in INI_Language1.
    ''' Uses ProcessSelectedText with SP_Translate prompt and INI_KeepFormat1 settings.
    ''' </summary>
    Public Async Sub InLanguage1()
        If INILoadFail() Then Return
        TranslateLanguage = INI_Language1
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Translate), True, INI_KeepFormat1, INI_KeepParaFormatInline, INI_ReplaceText1, False, 0, False, False, True, False, INI_KeepFormatCap, NoFormatAndFieldSaving:=Not INI_ReplaceText1)
    End Sub

    ''' <summary>
    ''' Translates selected text to the secondary language configured in INI_Language2.
    ''' Uses ProcessSelectedText with SP_Translate prompt and INI_KeepFormat1 settings.
    ''' </summary>
    Public Async Sub InLanguage2()

        If INILoadFail() Then Return
        TranslateLanguage = INI_Language2
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Translate), True, INI_KeepFormat1, INI_KeepParaFormatInline, INI_ReplaceText1, False, 0, False, False, True, False, INI_KeepFormatCap, NoFormatAndFieldSaving:=Not INI_ReplaceText1)
    End Sub

    ''' <summary>
    ''' Prompts user for target language and translates selected text.
    ''' Uses ProcessSelectedText with SP_Translate prompt and INI_KeepFormat1 settings.
    ''' </summary>
    Public Async Sub InOther()
        If INILoadFail() Then Return

        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim Selection As Microsoft.Office.Interop.Word.Selection = application.Selection

        If Selection.Type = WdSelectionType.wdSelectionIP Then
            TranslateWordDocuments()
            Return
        End If

        TranslateLanguage = SLib.ShowCustomInputBox("Enter your target language (e.g., English, German, French):", $"{AN} Translate", True)
        If Not String.IsNullOrEmpty(TranslateLanguage) Then
            Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Translate), True, INI_KeepFormat1, INI_KeepParaFormatInline, INI_ReplaceText1, False, 0, False, False, True, False, INI_KeepFormatCap, NoFormatAndFieldSaving:=Not INI_ReplaceText1)
        End If
    End Sub

    ''' <summary>
    ''' Corrects grammar and spelling in selected text using SP_Correct prompt.
    ''' Applies optional markup based on INI_DoMarkupWord and INI_MarkupMethodWord settings.
    ''' </summary>
    Public Async Sub Correct()
        If INILoadFail() Then Return

        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim Selection As Microsoft.Office.Interop.Word.Selection = application.Selection

        If Selection.Type = WdSelectionType.wdSelectionIP Then
            CorrectWordDocuments()
            Return
        End If


        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Correct), True, INI_KeepFormat2, INI_KeepParaFormatInline, Override(INI_ReplaceText2, INI_ReplaceText2Override), INI_DoMarkupWord, Override(INI_MarkupMethodWord, INI_MarkupMethodWordOverride), False, False, True, False, INI_KeepFormatCap, NoFormatAndFieldSaving:=Not Override(INI_ReplaceText2, INI_ReplaceText2Override))
    End Sub

    ''' <summary>
    ''' Enhances the selected text using the SP_Improve prompt while honoring the configured formatting and markup settings.
    ''' </summary>
    Public Async Sub Improve()
        If INILoadFail() Then Return
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Improve), True, INI_KeepFormat2, INI_KeepParaFormatInline, Override(INI_ReplaceText2, INI_ReplaceText2Override), INI_DoMarkupWord, Override(INI_MarkupMethodWord, INI_MarkupMethodWordOverride), False, False, True, False, INI_KeepFormatCap, NoFormatAndFieldSaving:=Not Override(INI_ReplaceText2, INI_ReplaceText2Override))
    End Sub

    ''' <summary>
    ''' Rewrites the selection with a friendlier tone through the SP_Friendly prompt, respecting the current formatting and markup options.
    ''' </summary>
    Public Async Sub Friendly()
        If INILoadFail() Then Return
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Friendly), True, INI_KeepFormat2, INI_KeepParaFormatInline, Override(INI_ReplaceText2, INI_ReplaceText2Override), INI_DoMarkupWord, Override(INI_MarkupMethodWord, INI_MarkupMethodWordOverride), False, False, True, False, INI_KeepFormatCap, NoFormatAndFieldSaving:=Not Override(INI_ReplaceText2, INI_ReplaceText2Override))
    End Sub

    ''' <summary>
    ''' Strengthens the persuasiveness of the selected text by invoking SP_Convincing with the active formatting and markup configuration.
    ''' </summary>
    Public Async Sub Convincing()
        If INILoadFail() Then Return
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Convincing), True, INI_KeepFormat2, INI_KeepParaFormatInline, Override(INI_ReplaceText2, INI_ReplaceText2Override), INI_DoMarkupWord, Override(INI_MarkupMethodWord, INI_MarkupMethodWordOverride), False, False, True, False, INI_KeepFormatCap, NoFormatAndFieldSaving:=Not Override(INI_ReplaceText2, INI_ReplaceText2Override))
    End Sub

    ''' <summary>
    ''' Removes filler phrases from the selection via SP_NoFillers while keeping the configured formatting logic intact.
    ''' </summary>
    Public Async Sub NoFillers()
        If INILoadFail() Then Return
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_NoFillers), True, INI_KeepFormat2, INI_KeepParaFormatInline, Override(INI_ReplaceText2, INI_ReplaceText2Override), INI_DoMarkupWord, Override(INI_MarkupMethodWord, INI_MarkupMethodWordOverride), False, False, True, False, INI_KeepFormatCap, NoFormatAndFieldSaving:=Not Override(INI_ReplaceText2, INI_ReplaceText2Override))
    End Sub

    ''' <summary>
    ''' Applies user-defined writing style from MyStyle prompt file to selected text.
    ''' Validates MyStyle configuration, prompts for style selection, then processes text.
    ''' </summary>
    Public Async Sub ApplyMyStyle()
        If INILoadFail() Then Return

        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim Selection As Microsoft.Office.Interop.Word.Selection = application.Selection

        Dim StylePath As String = ExpandEnvironmentVariables(INI_MyStylePath)

        ' Validate MyStyle configuration
        If String.IsNullOrWhiteSpace(StylePath) Then
            ShowCustomMessageBox("You have not defined a MyStyle prompt file. Please do so first in the configuration file or using 'Settings'.")
            Return
        End If
        If Not IO.File.Exists(StylePath) Then
            ShowCustomMessageBox("No MyStyle prompt file has been found. You may have to first create a MyStyle prompt. Go to 'Analyze' and use 'Define MyStyle' to do so - will abort.")
            Return
        End If
        If Selection.Type = WdSelectionType.wdSelectionIP Then
            ShowCustomMessageBox("Please select the text to be processed.")
            Return
        End If

        ' Prompt user to select style from MyStyle file
        MyStyleInsert = MyStyleHelpers.SelectPromptFromMyStyle(StylePath, "Word", 0, "Choose the style prompt to apply …", $"{AN} MyStyle", False)
        If MyStyleInsert = "ERROR" Then Return
        If MyStyleInsert = "NONE" OrElse String.IsNullOrWhiteSpace(MyStyleInsert) Then
            Return
        End If

        ' Apply selected style to text
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_MyStyle_Apply) & " " & MyStyleInsert, True, INI_KeepFormat2, INI_KeepParaFormatInline, Override(INI_ReplaceText2, INI_ReplaceText2Override), INI_DoMarkupWord, Override(INI_MarkupMethodWord, INI_MarkupMethodWordOverride), False, False, True, False, INI_KeepFormatCap, NoFormatAndFieldSaving:=Not Override(INI_ReplaceText2, INI_ReplaceText2Override))
    End Sub


    ''' <summary>
    ''' Checks document for identifiable information (PII/sensitive data).
    ''' Supports processing from selection, entire document, or external file.
    ''' Uses SP_CheckforII prompt with bubble extraction for results display.
    ''' </summary>
    Public Async Sub CheckDocumentII()
        If INILoadFail() Then Return

        Try
            Dim app As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
            Dim doc As Microsoft.Office.Interop.Word.Document = app.ActiveDocument
            If doc Is Nothing Then Return

            Dim sel As Microsoft.Office.Interop.Word.Selection = app.Selection
            If sel Is Nothing Then Return

            Dim rng As Microsoft.Office.Interop.Word.Range = sel.Range
            Dim FromFile As String = ""

            ' Handle no selection case - offer file import
            If sel.Type = WdSelectionType.wdSelectionIP Then
                Dim answer As Integer = ShowCustomYesNoBox("You have not selected any text. Do you instead want to analyze text from a document file or Powerpoint presentation?", "Yes", "No, proceed with this text", AN & " Check for Identifiable Information")
                If answer = 1 Then
                    ' Configure supported file types
                    DragDropFormLabel = "Document files (.txt, .docx, .pdf) or Powerpoint (.pptx)."
                    DragDropFormFilter = "Supported Files|*.txt;*.rtf;*.doc;*.docx;*.pdf;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm;*.pptx|" &
                             "Text Files (*.txt;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm)|*.txt;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm|" &
                             "Rich Text Files (*.rtf)|*.rtf|" &
                             "Word Documents (*.doc;*.docx)|*.doc;*.docx|" &
                             "PDF Files (*.pdf)|*.pdf|" &
                             "Powerpoint Files (*.pptx)|*.pptx"

                    Dim FilePath As String = GetFileName()
                    DragDropFormLabel = ""
                    DragDropFormFilter = ""
                    If String.IsNullOrWhiteSpace(FilePath) Then
                        ShowCustomMessageBox("No file has been selected - will abort.")
                        Return
                    End If

                    ' Extract content from selected file
                    Dim ext As String = IO.Path.GetExtension(FilePath).ToLowerInvariant()
                    FromFile = Await GetFileContent(FilePath, False, True, True)

                    If FromFile.StartsWith("Error:") Then
                        ShowCustomMessageBox(FromFile)
                        Return
                    End If
                    If String.IsNullOrWhiteSpace(FromFile) Then
                        ShowCustomMessageBox("The file you provided did not contain any text - will abort.")
                        Return
                    End If

                    ' Create new document with imported content
                    Dim newDoc As Word.Document = Globals.ThisAddIn.Application.Documents.Add()
                    newDoc.Activate()

                    rng = newDoc.Content
                    rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                    ' Sanitize: remove NULs and normalize line breaks
                    Dim safeText As String = If(FromFile, String.Empty)
                    safeText = safeText.Replace(ChrW(0), String.Empty)

                    rng.InsertAfter(safeText)

                    newDoc.Content.Select()
                    SelectedText = newDoc.Application.Selection.Text.Trim()

                    answer = ShowCustomYesNoBox("The content of your document has been inserted into a new document. Continue with the checking process?", "Yes", "No")
                    If answer <> 1 Then Return

                ElseIf answer = 0 Then
                    Return
                End If
            End If

            ' If no selection, select entire document
            If rng Is Nothing OrElse rng.Start = rng.End Then
                doc.Content.Select()
            End If

            ' Process document for identifiable information
            Dim result As System.String = Await ProcessSelectedText(
        InterpolateAtRuntime(SP_CheckforII),
        True,
        INI_KeepFormat2,
        INI_KeepParaFormatInline,
        Override(INI_ReplaceText2, INI_ReplaceText2Override),
        INI_DoMarkupWord,
        Override(INI_MarkupMethodWord, INI_MarkupMethodWordOverride),
        True,
        False,
        True,
        False,
        INI_KeepFormatCap,
        NoFormatAndFieldSaving:=Not Override(INI_ReplaceText2, INI_ReplaceText2Override),
        DoBubblesExtract:=True
    )

        Catch ex As System.Exception
            ' Silent catch
        End Try
    End Sub

    ''' <summary>
    ''' Opens text editor for local or central redaction instructions file.
    ''' Creates file with template if it doesn't exist.
    ''' </summary>
    Public Async Sub EditRedactionInstructions()
        If INILoadFail() Then Return
        Dim chosenPath As String = ""
        If Not String.IsNullOrEmpty(INI_RedactionInstructionsPathLocal) AndAlso Not String.IsNullOrEmpty(INI_RedactionInstructionsPath) Then
            Dim answer = ShowCustomYesNoBox("Do you want to edit the local or central redaction instructions set (if it does not yet exist, it will be created)?", "Local", "Central")
            If answer = 0 Then Return
            chosenPath = If(answer = 1, INI_RedactionInstructionsPathLocal, INI_RedactionInstructionsPath)
            chosenPath = ExpandEnvironmentVariables(chosenPath)
        ElseIf String.IsNullOrEmpty(INI_RedactionInstructionsPathLocal) Then
            chosenPath = ExpandEnvironmentVariables(INI_RedactionInstructionsPath)
        Else
            chosenPath = ExpandEnvironmentVariables(INI_RedactionInstructionsPathLocal)
        End If

        If String.IsNullOrWhiteSpace(chosenPath) Then
            ShowCustomMessageBox("No valid path found for redaction instructions.")
            Return
        End If

        ' Generate initial file if path is configured but file doesn't exist
        If Not String.IsNullOrWhiteSpace(chosenPath) AndAlso Not File.Exists(chosenPath) Then
            Try
                Dim directoryPath = System.IO.Path.GetDirectoryName(chosenPath)
                If Not String.IsNullOrWhiteSpace(directoryPath) AndAlso Not Directory.Exists(directoryPath) Then
                    Directory.CreateDirectory(directoryPath)
                End If

                Dim initialContent As String = "; Redaction Instructions (format: name|redaction instruction)" & vbCrLf & vbCrLf
                File.WriteAllText(chosenPath, initialContent, System.Text.Encoding.UTF8)
            Catch ex As Exception
                ShowCustomMessageBox($"Failed to create initial redaction instructions file ({chosenPath}): {ex.Message}")
                Return
            End Try
        End If

        SLib.ShowTextFileEditor(chosenPath, $"{AN} Redaction Instructions '{chosenPath}':")
    End Sub


    ''' <summary>
    ''' Anonymizes selected text by replacing identifiable entities.
    ''' Supports multiple markup methods; prompts user for method if configuration suggests alternatives.
    ''' Uses SP_Anonymize prompt with optional Regex markup for larger texts.
    ''' </summary>
    Public Async Sub Anonymize()
        If INILoadFail() Then Return

        Dim DoMarkup As Boolean = INI_DoMarkupWord
        Dim DoReplace As Boolean = Override(INI_ReplaceText2, INI_ReplaceText2Override)
        If Not DoMarkup Or Not DoReplace Then
            Dim result2 As Integer = ShowCustomYesNoBox($"As per your current settings no markup will be applied. For anonymizing a larger text, doing a markup may be a better choice. How Do you want To Continue?", "Continue As Is", "Continue With a markup")
            If result2 = 2 Then
                DoMarkup = True
            End If
        End If

        Dim MarkupMethod As Integer = Override(INI_MarkupMethodWord, INI_MarkupMethodWordOverride)
        If INI_DoMarkupWord And MarkupMethod <> 4 Then
            Dim MarkupNow As String = ""
            Select Case Override(INI_MarkupMethodWord, INI_MarkupMethodWordOverride)
                Case 1
                    MarkupNow = "Word markup method"
                Case 2
                    MarkupNow = "Diff markup method"
                Case 3
                    MarkupNow = "Diff markup method (With the output In a separate window)"
            End Select

            Dim result2 As Integer = ShowCustomYesNoBox($"You have chosen the {MarkupNow}. If you are anonymizing a larger text, the 'Regex' markup method may be a better choice. How do you want to continue?", "Continue as is", "Use Regex")
            If result2 = 2 Then
                MarkupMethod = 4
                DoReplace = True
            End If
        End If


        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Anonymize), True, INI_KeepFormat2, INI_KeepParaFormatInline, DoReplace, DoMarkup, MarkupMethod, False, False, True, False, INI_KeepFormatCap, NoFormatAndFieldSaving:=Not DoReplace)
    End Sub


    ''' <summary>
    ''' Provides explanations for the selected text using the SP_Explain prompt.
    ''' </summary>
    Public Async Sub Explain()
        If INILoadFail() Then Return

        Try
            Dim app As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
            Dim doc As Microsoft.Office.Interop.Word.Document = app.ActiveDocument
            If doc Is Nothing Then Return

            Dim sel As Microsoft.Office.Interop.Word.Selection = app.Selection
            If sel Is Nothing Then Return

            Dim rng As Microsoft.Office.Interop.Word.Range = sel.Range
            ' No selection -> select the entire document content
            If rng Is Nothing OrElse rng.Start = rng.End Then
                doc.Content.Select()
            End If

            If String.IsNullOrWhiteSpace(app.Selection.Text.Trim) Then
                ShowCustomMessageBox("There is no text I can explain. Open it or import it using 'Word Helpers'.")
                Return
            End If

            CurrentDate = "(Current Date: " & DateTime.Now.ToString("dd-MMM-yyyy", CultureInfo.GetCultureInfo("en-US")) & ")"

            Dim result As System.String = Await ProcessSelectedText(
            InterpolateAtRuntime(SP_Explain),
            True,
            INI_KeepFormat2,
            INI_KeepParaFormatInline,
            Override(INI_ReplaceText2, INI_ReplaceText2Override),
            INI_DoMarkupWord,
            Override(INI_MarkupMethodWord, INI_MarkupMethodWordOverride),
            True,
            False,
            True,
            False,
            INI_KeepFormatCap,
            NoFormatAndFieldSaving:=Not Override(INI_ReplaceText2, INI_ReplaceText2Override),
            DoBubblesExtract:=True
        )

        Catch ex As System.Exception
            ' 
        End Try
    End Sub

    ''' <summary>
    ''' Suggests titles for the selected text using the SP_SuggestTitles prompt.
    ''' Honors formatting and markup settings from configuration.
    ''' </summary>
    Public Async Sub SuggestTitles()
        If INILoadFail() Then Return
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_SuggestTitles), True, INI_KeepFormat2, INI_KeepParaFormatInline, Override(INI_ReplaceText2, INI_ReplaceText2Override), INI_DoMarkupWord, Override(INI_MarkupMethodWord, INI_MarkupMethodWordOverride), True, False, True, False, INI_KeepFormatCap, NoFormatAndFieldSaving:=Not Override(INI_ReplaceText2, INI_ReplaceText2Override))
    End Sub

    ''' <summary>
    ''' Sends clipboard content to the SP_InsertClipboard prompt (requires INI_APICall_Object) and inserts the LLM response beneath the caret.
    ''' </summary>
    Public Async Sub InsertClipboard()

        If String.IsNullOrWhiteSpace(INI_APICall_Object) Then
            ShowCustomMessageBox($"Your model ({INI_Model}) is not configured to process clipboard data (i.e. binary objects).")
            Return
        End If

        With Globals.ThisAddIn.Application.Selection
            If .Start <> .End Then
                .Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            End If
        End With

        If INILoadFail() Then Return

        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_InsertClipboard), False, False, False, False, False, 0, False, False, False, False, 0, False, "", False, "clipboard")

        If result <> "" Then
            Globals.ThisAddIn.Application.Selection.TypeParagraph()
            Globals.ThisAddIn.Application.Selection.TypeParagraph()
            InsertTextWithMarkdown(Globals.ThisAddIn.Application.Selection, vbCrLf & result & vbCrLf, False)

        End If

    End Sub

    ''' <summary>
    ''' Shortens selected text by specified percentage.
    ''' Prompts user for target reduction percentage, calculates word count, applies SP_Shorten prompt.
    ''' </summary>
    Public Async Sub Shorten()

        If INILoadFail() Then Return
        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim Selection As Microsoft.Office.Interop.Word.Selection = application.Selection

        If Selection.Type = WdSelectionType.wdSelectionIP Then
            ShowCustomMessageBox("Please select the text to be processed.")
            Return
        End If

        Dim Textlength As Integer = GetSelectedTextLength()
        Dim UserInput As String
        Dim ShortenPercentValue As Integer = 0
        Do
            UserInput = SLib.ShowCustomInputBox("Enter the percentage by which your text should be shortened (it has " & Textlength & " words; " & ShortenPercent & "% will cut approx. " & (Textlength * ShortenPercent / 100) & " words)", $"{AN} Shortener", True, CStr(ShortenPercent) & "%").Trim()
            If String.IsNullOrEmpty(UserInput) Then
                Return
            End If
            UserInput = UserInput.Replace("%", "").Trim()
            If Integer.TryParse(UserInput, ShortenPercentValue) AndAlso ShortenPercentValue >= 1 AndAlso ShortenPercentValue <= 99 Then
                Exit Do
            Else
                ShowCustomMessageBox("Please enter a valid percentage between 1 And 99.")
            End If
        Loop
        If ShortenPercentValue = 0 Then Return
        ShortenLength = Textlength * (100 - ShortenPercentValue) / 100
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Shorten), True, INI_KeepFormat2, INI_KeepParaFormatInline, Override(INI_ReplaceText2, INI_ReplaceText2Override), INI_DoMarkupWord, Override(INI_MarkupMethodWord, INI_MarkupMethodWordOverride), False, False, True, False, INI_KeepFormatCap, NoFormatAndFieldSaving:=Not Override(INI_ReplaceText2, INI_ReplaceText2Override))
    End Sub

    ''' <summary>
    ''' Expands the selected text to a user-defined word count using SP_Filibuster and the current formatting/markup preferences.
    ''' </summary>
    Public Async Sub Filibuster()

        If INILoadFail() Then Return
        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim Selection As Microsoft.Office.Interop.Word.Selection = application.Selection

        If Selection.Type = WdSelectionType.wdSelectionIP Then
            ShowCustomMessageBox("Please select the text to be processed.")
            Return
        End If

        Dim Textlength As Integer = GetSelectedTextLength()
        Dim UserInput As String
        FilibusterLength = Textlength * 10
        Do
            UserInput = SLib.ShowCustomInputBox("Enter the number of words your filibuster shall have (your base text has " & Textlength & " words)", $"{AN} Filibuster (Expand Text)", True, CStr(Filibusterlength)).Trim()
            If String.IsNullOrEmpty(UserInput) Then
                Return
            End If
            If Integer.TryParse(UserInput, FilibusterLength) AndAlso FilibusterLength > Textlength AndAlso FilibusterLength <= MaxFilibuster Then
                Exit Do
            Else
                ShowCustomMessageBox($"Please enter a range between {Textlength} and {MaxFilibuster} words.")
            End If
        Loop
        If FilibusterLength = 0 Then Return
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Filibuster), True, INI_KeepFormat2, INI_KeepParaFormatInline, Override(INI_ReplaceText2, INI_ReplaceText2Override), INI_DoMarkupWord, Override(INI_MarkupMethodWord, INI_MarkupMethodWordOverride), False, False, True, False, INI_KeepFormatCap, NoFormatAndFieldSaving:=Not Override(INI_ReplaceText2, INI_ReplaceText2Override))
    End Sub

    ''' <summary>
    ''' Generates an opposing argument of the requested length against the selection by calling SP_ArgueAgainst without replacement markup.
    ''' </summary>
    Public Async Sub ArgueAgainst()

        If INILoadFail() Then Return
        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim Selection As Microsoft.Office.Interop.Word.Selection = application.Selection

        If Selection.Type = WdSelectionType.wdSelectionIP Then
            ShowCustomMessageBox("Please select the text to be processed.")
            Return
        End If

        Dim UserInput As String
        FilibusterLength = ArgueAgainstDefault
        Do
            UserInput = SLib.ShowCustomInputBox("Enter the number of words your argument shall have:", $"{AN} Argue Against", True, CStr(FilibusterLength)).Trim()
            If String.IsNullOrEmpty(UserInput) Then
                Return
            End If
            If Integer.TryParse(UserInput, FilibusterLength) AndAlso FilibusterLength > 0 AndAlso FilibusterLength <= MaxFilibuster Then
                Exit Do
            Else
                ShowCustomMessageBox($"Please enter a range between 0 and {MaxFilibuster} words.")
            End If
        Loop
        If FilibusterLength = 0 Then Return
        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_ArgueAgainst), False, False, False, False, False, 0, True, False, True, False, 0)
    End Sub

    ''' <summary>
    ''' Swaps occurrences of two user-specified party names through SP_SwitchParty, optionally forcing Regex markup for larger texts.
    ''' </summary>
    Public Async Sub SwitchParty()
        If INILoadFail() Then Return
        Dim UserInput As String
        Do
            UserInput = SLib.ShowCustomInputBox("Please provide the original party name And the New party name, separated by a comma (example: Elvis Presley, Taylor Swift):", $"{AN} Switch Party", True).Trim()

            If String.IsNullOrEmpty(UserInput) Then
                Return
            End If

            Dim parts() As String = UserInput.Split(","c)
            If parts.Length = 2 Then
                OldParty = parts(0).Trim()
                NewParty = parts(1).Trim()
                Exit Do
            Else
                ShowCustomMessageBox("Please enter two names separated by a comma.")
            End If
        Loop

        Dim DoMarkup As Boolean = INI_DoMarkupWord
        Dim DoReplace As Boolean = Override(INI_ReplaceText2, INI_ReplaceText2Override)
        If Not DoMarkup Or Not DoReplace Then
            Dim result2 As Integer = ShowCustomYesNoBox($"As per your current settings no markup will be applied. For using 'Switch Party' on a larger texts, markup may be a better choice. How do you want to continue?", "Continue as is", "Continue with a markup")
            If result2 = 2 Then
                DoMarkup = True
                DoReplace = True
            End If
        End If

        Dim MarkupMethod As Integer = Override(INI_MarkupMethodWord, INI_MarkupMethodWordOverride)
        If INI_DoMarkupWord And MarkupMethod <> 4 Then
            Dim MarkupNow As String = ""
            Select Case Override(INI_MarkupMethodWord, INI_MarkupMethodWordOverride)
                Case 1
                    MarkupNow = "Word markup method"
                Case 2
                    MarkupNow = "Diff markup method"
                Case 3
                    MarkupNow = "Diff markup method (with the output in a separate window)"
            End Select

            Dim result2 As Integer = ShowCustomYesNoBox($"You have chosen the {MarkupNow}. If you are using 'Switch Party' with a larger text, the 'Regex' markup method may be a better choice. How do you want to continue?", "Continue as is", "Use Regex")
            If result2 = 2 Then
                MarkupMethod = 4
                DoReplace = True
            End If
        End If

        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_SwitchParty), True, INI_KeepFormat2, INI_KeepParaFormatInline, DoReplace, DoMarkup, MarkupMethod, False, False, True, False, INI_KeepFormatCap, NoFormatAndFieldSaving:=Not DoReplace)

    End Sub

    ''' <summary>
    ''' Summarizes the selected text to a user-defined word count using SP_Summarize without replacement markup.
    ''' </summary>
    Public Async Sub Summarize()
        If INILoadFail() Then Return
        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim Selection As Microsoft.Office.Interop.Word.Selection = application.Selection

        If Selection.Type = WdSelectionType.wdSelectionIP Then
            ShowCustomMessageBox("Please select the text to be processed.")
            Return
        End If

        Dim Textlength As Integer = GetSelectedTextLength()

        Dim UserInput As String
        SummaryLength = 0

        Do
            UserInput = SLib.ShowCustomInputBox("Enter the number of words your summary shall have (the selected text has " & Textlength & " words; the proposal " & SummaryPercent & "%):", $"{AN} Summarizer", True, CStr(System.Math.Round(SummaryPercent * Textlength / 100 / 5) * 5)).Trim()

            If String.IsNullOrEmpty(UserInput) Then
                Return
            End If

            If Integer.TryParse(UserInput, SummaryLength) AndAlso SummaryLength >= 1 AndAlso SummaryLength <= Textlength Then
                Exit Do
            Else
                ShowCustomMessageBox("Please enter a valid word count between 1 and " & Textlength & ".")
            End If
        Loop
        If SummaryLength = 0 Then Return

        Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SP_Summarize), False, False, False, False, False, 0, True, False, True, False, 0)
    End Sub


    ''' <summary>
    ''' Opens the speech transcription form for audio-to-text conversion.
    ''' Configures Whisper.net library path from INI_SpeechModelPath before launching.
    ''' </summary>
    Public Sub Transcriptor()
        If INILoadFail() Then Return
        If Not String.IsNullOrEmpty(INI_SpeechModelPath) Then
            Dim SpeechPath As String = ExpandEnvironmentVariables(INI_SpeechModelPath)
            If Not String.IsNullOrEmpty(SpeechPath) AndAlso Not SpeechPath.EndsWith("\") Then
                SpeechPath = SpeechPath & "\"
            End If
            Dim currentPath As String = Environment.GetEnvironmentVariable("PATH")

            If Not currentPath.Contains(SpeechPath) Then
                Environment.SetEnvironmentVariable("PATH", currentPath & ";" & SpeechPath)
            End If
            RuntimeOptions.LibraryPath = SpeechPath
            'RuntimeOptions.RuntimeLibraryOrder = New List(Of RuntimeLibrary) From {RuntimeLibrary.Cuda, RuntimeLibrary.Cpu}

        End If

        Dim TranscriptionForm = New TranscriptionForm()
        TranscriptionForm.Show()
    End Sub

    Private chatForm As frmAIChat

    ''' <summary>
    ''' Shows or activates the AI chat window.
    ''' Restores previous window position/size from My.Settings if available.
    ''' </summary>
    Public Sub ShowChatForm()
        If INILoadFail() Then Return
        If chatForm Is Nothing OrElse chatForm.IsDisposed Then
            chatForm = New frmAIChat(_context)

            ' Set the location and size before showing the form
            If My.Settings.FormLocation <> System.Drawing.Point.Empty AndAlso My.Settings.FormSize <> System.Drawing.Size.Empty Then
                chatForm.StartPosition = FormStartPosition.Manual
                chatForm.Location = My.Settings.FormLocation
                chatForm.Size = My.Settings.FormSize
            Else
                ' Default to center screen if no settings are available
                chatForm.StartPosition = FormStartPosition.Manual
                Dim screenBounds As System.Drawing.Rectangle = Screen.PrimaryScreen.WorkingArea
                chatForm.Location = New System.Drawing.Point((screenBounds.Width - chatForm.Width) \ 2, (screenBounds.Height - chatForm.Height) \ 2)
                chatForm.Size = New System.Drawing.Size(650, 500) ' Set default size if needed
            End If
        End If

        ' Show and bring the form to the front
        chatForm.Show()
        chatForm.BringToFront()
    End Sub


    ''' <summary>
    ''' Guides the user through sample selection, optional alternate model usage, and stores the SP_MyStyle_Word result in the MyStyle prompt file.
    ''' </summary>
    Public Async Sub DefineMyStyle()
        If INILoadFail() Then Return

        Dim StylePath As String = ExpandEnvironmentVariables(INI_MyStylePath)

        If String.IsNullOrWhiteSpace(StylePath) Then
            ShowCustomMessageBox("You have not configured a MyStyle prompt file path. Please do so in the configuration file or using 'Settings'.")
            Return
        End If

        Dim Label As String = $"You are about to have {AN} create a profile of your writing style. There are six steps:" & vbCrLf & vbCrLf &
                               "1. If you have selected text, this will be used as a sample." & vbCrLf &
                               "2. You select one, all or none of the open Word documents as further samples." & vbCrLf &
                               "3. You can provide further input as further instructions to the model (e.g., Internet links to check if the model is able to do so)." & vbCrLf &
                               "4. You select the model to perform the analysis (e.g., a reasoning model, Internet access if links are to be consulted)" & vbCrLf &
                               "5. You can review and amend the analysis, including at the end the prompt for the AI to implement your style." & vbCrLf &
                               $"6. The analysis will be saved to your personal MyStyle prompt file ({StylePath})."

        Dim Answer As Integer = ShowCustomYesNoBox(Label, "Continue", "Cancel", $"{AN} Define MyStyle",
                                    extraButtonText:="Edit MyStyle prompt file",
                                                            extraButtonAction:=Sub()
                                                                                   SLib.ShowTextFileEditor(StylePath, "Edit your MyStyle prompt file (use 'Define MyStyle' to create new prompts automatically):")
                                                                               End Sub)

        If Answer <> 1 Then Return

        ' Get Selected text

        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim selection As Word.Selection = application.Selection
        SelectedText = ""

        If selection.Type <> Word.WdSelectionType.wdSelectionIP Then
            SelectedText = selection.Text.Trim()
        End If

        ' Get other open documents

        InsertDocs = ""
        InsertDocs = GatherSelectedDocuments(False, True)
        Debug.WriteLine($"GatherSelectedDocs returned: {Left(InsertDocs, 3000)}")
        If InsertDocs.StartsWith("ERROR", StringComparison.OrdinalIgnoreCase) Then
            ShowCustomMessageBox($"An error occured gathering the additional document(s) ({InsertDocs.Substring(6).Trim()}) - will abort.")
            Return
        ElseIf InsertDocs.StartsWith("NONE", StringComparison.OrdinalIgnoreCase) Then
            ShowCustomMessageBox($"There are no other documents to add - will abort.")
            Return
        End If

        ' Get addition instructions

        OtherPrompt = ""
        OtherPrompt = SLib.ShowCustomInputBox("You can provide additional instructions for the analysis (e.g., Internet links to check [if your model will understand so], aspects to focus on etc.). This is optional.", $"{AN} Define MyStyle", False).Trim()

        If OtherPrompt = "ESC" Then
            Return
        End If

        ' Get the model to do it

        Dim UseSecondAPI As Boolean = False
        If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
            Answer = ShowCustomYesNoBox($"Do you want to use one of your alternate models?", "Yes, use alternate", "No, use primary", $"{AN} Define MyStyle")
            Debug.WriteLine("Answer=" & Answer)
            If Answer = 1 Then
                If Not ShowModelSelection(_context, INI_AlternateModelPath) Then
                    originalConfigLoaded = False
                    Return
                End If
                UseSecondAPI = True
            ElseIf Answer <> 2 Then
                Return
            End If
        End If

        ' SP_MyStyle will contain OtherPrompt and InsertDocs

        Dim LLMResponse As String = Await LLM(InterpolateAtRuntime(SP_MyStyle_Word), If(SelectedText <> "", "<DOCUMENT0>" & SelectedText & "</DOCUMENT0> ", "") & InsertDocs, "", "", 0, UseSecondAPI)

        If Not String.IsNullOrWhiteSpace(LLMResponse) Then

            Dim analysis As String = SLib.ShowCustomWindow("The AI provided the following style analysis and MyStyle prompt based on your text samples:", LLMResponse, "If you choose 'OK', the prompt and its title at the end of the analysis will be stored in your MyStyle prompt file for future usage (and the full report copied to the clipboard).", AN, False, False, False, False)

            If Not String.IsNullOrWhiteSpace(analysis) Then
                SLib.PutInClipboard(analysis)
                SLib.ExtractAndStorePromptFromAnalysis(analysis, StylePath, "Word")
            End If

        End If
        If UseSecondAPI And originalConfigLoaded Then
            RestoreDefaults(_context, originalConfig)
            originalConfigLoaded = False
        End If

    End Sub


    ''' <summary>
    ''' Easter egg feature: generates humorous wisdom based on Hasidic humor tradition.
    ''' Uses current selection, paragraph, or user context (name, date) as input.
    ''' Respects INI_RoastMe setting for tone (politically correct vs. provocative).
    ''' </summary>
    Public Async Sub EasterEgg()

        Dim splash As New SLib.SplashScreen($"{AN6} is preparing to tickle{If(INI_RoastMe, " (inofficial version)", "")}...")
        splash.Show()
        splash.Refresh()

        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim selection As Microsoft.Office.Interop.Word.Selection = application.Selection
        Dim inputtext As String = Nothing
        Dim result As String
        Dim UserLanguage As String = Globals.ThisAddIn.GetWordDefaultInterfaceLanguage()
        Dim basicinstruction As String = $"Write me in {UserLanguage} a piece of wisdom based on hasidic humour related To the following input. The output shall be politically correct, contain nothing religious and not be offensive, but within those bounds it can be provocative. Do tickle my brain!"
        If INI_RoastMe Then basicinstruction = $"Write me in {UserLanguage} a piece of wisdom based on hasidic humour related To the following input. You do not have to be politically correct. You can use black humour, if this fits the purpose. You don't have to take any sensitivities into consideration. But don't say anything that could be considered discrimination based on race, ethnicity, religion or sexual orientation. Be direct! Be provocative! Roast me! Tickle my brain!"

        If selection IsNot Nothing Then
            ' Case 1: A selection exists
            If Not String.IsNullOrWhiteSpace(selection.Text) Then
                inputtext = selection.Text
            End If
        End If

        If String.IsNullOrWhiteSpace(inputtext) Then
            ' Case 2: No selection or empty selection, use paragraph text at the cursor
            Dim currentParagraph As Word.Paragraph = Nothing

            If selection IsNot Nothing AndAlso selection.Range IsNot Nothing Then
                currentParagraph = selection.Range.Paragraphs.First
            End If

            If currentParagraph Is Nothing OrElse String.IsNullOrWhiteSpace(currentParagraph.Range.Text) Then
                ' Case 3: No cursor paragraph, fallback to the first paragraph with text
                For Each paragraph As Word.Paragraph In application.ActiveDocument.Paragraphs
                    If Not String.IsNullOrWhiteSpace(paragraph.Range.Text) Then
                        currentParagraph = paragraph
                        Exit For
                    End If
                Next
            End If

            If currentParagraph IsNot Nothing Then
                inputtext = currentParagraph.Range.Text.Trim()
            End If
        End If

        If String.IsNullOrWhiteSpace(inputtext) Then
            ' Case 4: No text in document, use fallback logic
            Dim userName As String = Globals.ThisAddIn.Application.UserName
            Dim currentMonth As String = DateTime.Now.ToString("MMMM") ' Full month name
            Dim currentDay As String = DateTime.Now.ToString("dddd")  ' Full day name
            result = Await LLM($"{basicinstruction} The only input you get is the name of the current user ({userName}) of this Word application (you may create friendly variations), the current month is {currentMonth} and the current day of week is {currentDay}.", "", "", "", 0, False, True)
        Else
            result = Await LLM($"{basicinstruction} This Is the text: {inputtext}", "", "", "", 0, False, True)
        End If
        splash.Close()

        result = result.TrimEnd()
        result = result.Replace($"{vbCrLf}* ", vbCrLf & ChrW(8226) & " ").Replace($"{vbCr}* ", vbCr & ChrW(8226) & " ").Replace($"{vbLf}* ", vbLf & ChrW(8226) & " ")
        result = result.Replace($"  *  ", "  " & ChrW(8226) & "  ")
        result = RemoveMarkdownFormatting(result)

        ShowCustomMessageBox(result, $"{AN6} tickles your brain ...")
    End Sub


    ''' <summary>
    ''' Runs the configured file/prompt-based anonymization on the current selection, displays entity mappings, and copies the chosen output to the clipboard.
    ''' </summary>

    Public Sub AnonymizeSelection()

        Dim sel As String = Globals.ThisAddIn.Application.Selection.Text

        If String.IsNullOrWhiteSpace(sel) Then
            SLib.ShowCustomMessageBox("Please select a text to anonymize.")
            Return
        End If

        Dim AnonSetting As String = INI_Anon
        Dim OverrideAnonSetting As String = LoadAnonSettingsForModel(INI_Model)

        If Not String.IsNullOrWhiteSpace(OverrideAnonSetting) Then AnonSetting = OverrideAnonSetting
        If Not String.IsNullOrWhiteSpace(AnonSetting) Then
            Dim AnonType As Integer = ShowCustomYesNoBox($"Which anonymization type do you want (using keys for '{INI_Model}')?", "3 - file based", "4 - prompt", $"{AN} Anonymization") + 2
            If AnonType > 2 Then
                Dim AnonMode As String = "silent"
                Dim AnonText As String = AnonymizeText(sel, INI_Model, AnonMode, AnonType)
                AnonText = AnonText & vbCrLf & vbCrLf & "**Entities:**  " & vbCrLf & vbCrLf & ExportEntitiesMappings()
                Dim result As String = ShowCustomWindow("The anonymization returned the following text:", AnonText, $"Beware that this anonymization depends entirely on the keys you provided in your file '{AnonFile}' (for your model '{INI_Model}') or your prompt. Check the result. Choose what to put into the clipboard.", $"{AN} Anonymization", False)

                If result <> "" Then
                    SLib.PutInClipboard(result)
                End If
            End If
        End If

        Return

        If String.IsNullOrEmpty(Globals.ThisAddIn.INI_LocalModelPath) Then
            SLib.ShowCustomMessageBox("No path set for the NER model ('LocalModelPath').")
            Return
        End If

        If Not EnsureInitialized() Then Return

        'Dim sel As String = Globals.ThisAddIn.Application.Selection.Text
        If String.IsNullOrWhiteSpace(sel) Then
            SLib.ShowCustomMessageBox("Please select a text to anonymize.")
            Return
        End If

        Dim anon As String = OnnxAnonymizer.Anonymize(sel)

        Dim sb As New StringBuilder()
        sb.AppendLine(anon)
        sb.AppendLine()
        sb.AppendLine("Entity-Mapping:")
        sb.AppendLine()
        For Each kvp In OnnxAnonymizer.mapping
            sb.AppendLine($"{kvp.Key} -> {kvp.Value}")
        Next

        Dim FinalText As String = ShowCustomWindow("The NER anonymization returned the following text:", sb.ToString(), "Beware that this anonymization method is fast, but not of very high precision. Check the result.", AN, False)

        If FinalText <> "" Then
            SLib.PutInClipboard(FinalText)
        End If

    End Sub


    Private _win As HelpMeInky = Nothing

    ''' <summary>
    ''' Displays (or raises) the HelpMeInky helper window bound to the current add-in context.
    ''' </summary>
    Public Sub HelpMeInky()
        If _win Is Nothing OrElse _win.IsDisposed Then
            _win = New HelpMeInky(_context, RDV)
        End If
        ' No owner needed
        _win.ShowRaised()
    End Sub

    Private _win2 As DiscussInky = Nothing

    ''' <summary>
    ''' Displays (or raises) the DiscussInky discussion window for the active add-in context.
    ''' </summary>
    Public Sub DiscussInky()
        If _win2 Is Nothing OrElse _win2.IsDisposed Then
            _win2 = New DiscussInky(_context)
        End If
        ' No owner needed
        _win2.ShowRaised()
    End Sub

    ''' <summary>
    ''' Displays the settings editor window with configuration options and tooltips.
    ''' Updates context menu after settings are changed.
    ''' </summary>
    Public Sub ShowSettings()

        If INILoadFail() Then Return
        Dim Settings As New Dictionary(Of String, String) From {
                {"Temperature", "Temperature of {model}"},
                {"Timeout", "Timeout of {model}"},
                {"Temperature_2", "Temperature of {model2}"},
                {"Timeout_2", "Timeout of {model2}"},
                {"DoubleS", "Convert '" & ChrW(223) & "' to 'ss'"},
                {"Clean", "Clean the LLM response"},
                {"NoEmDash", "Convert em to en dash"},
                {"Ignore", "Activate 'Ignore' prompt (for 'prompt injection' protection)"},
                {"MarkdownConvert", "Keep character formatting"},
                {"KeepFormat1", "Keep format (translations, additional coding)"},
                {"ReplaceText1", "Replace text (translations)"},
                {"KeepFormat2", "Keep format (other commands, additional coding)"},
                {"ReplaceText2", "Replace text (other commands)"},
                {"ReplaceText2Override", "Replace text (other commands) [override]"},
                {"KeepParaFormatInline", "Keep paragraph format (additional coding)"},
                {"KeepFormatCap", "Maximum text for keeping format (chars)"},
                {"DoMarkupWord", "Output as a markup (some functions)"},
                {"MarkupMethodHelper", "Markup method helpers (1 = Word, 2 = Diff, 3 = DiffW)"},
                {"MarkupMethodWord", "Markup method (1 = Word, 2 = Diff, 3 = DiffW, 4 = Regex)"},
                {"MarkupMethodWordOverride", "Markup method (1 = Word, 2 = Diff, 3 = DiffW, 4 = Regex) [override]"},
                {"MarkupDiffCap", "Maximum characters for Diff Markup"},
                {"MarkupRegexCap", "Maximum characters for Regex Markup"},
                {"MarkdownBubbles", "Use Markdown in Word bubbles"},
                {"PreCorrection", "Additional instruction for prompts"},
                {"PostCorrection", "Prompt to apply after queries"},
                {"Language1", "Default translation language 1"},
                {"Language2", "Default translation language 2"},
                {"PromptLibPath", "Prompt library file"},
                {"PromptLibPathLocal", "Prompt library file (local)"},
                {"PromptLibPath_Transcript", "Transcript prompt library file"},
                {"ShortcutsWordExcel", "Key shortcuts (for direct access)"},
                {"ChatCap", "Chat conversation memory (chars)"},
                {"MyStylePath", "Path to the MyStyle prompt file"},
                {"DefaultPrefix", "Default prefix to use in 'Freestyle'"},
                {"Location", "Location information to use, e.g., in 'Freestyle'"},
                {"ToolingLogWindow", "Tooling: Show log window"},
                {"ToolingDryRun", $"Tooling: Show {ToolFriendlyName.ToLower} overview before running"}
            }
        Dim SettingsTips As New Dictionary(Of String, String) From {
                {"Temperature", "The higher, the more creative the LLM will be (0.0-2.0)"},
                {"Timeout", "In milliseconds"},
                {"Temperature_2", "The higher, the more creative the LLM will be (0.0-2.0)"},
                {"Timeout_2", "In milliseconds"},
                {"DoubleS", "For Switzerland"},
                {"Clean", "To remove double-spaces and hidden markers that may have been inserted by the LLM"},
                {"NoEmDash", "This will convert long dashes typically generated by LLMs but that are not commonly used (thus suggesting that the text has been AI generated)"},
                {"Ignore", "Allow system prompts to use {Ignore} as a placeholder for text to ignore, such as malicious prompt injections; Freestyle and some other commands use {Ignore}; the chatbots have an independent protection"},
                {"MarkdownConvert", "If selected, bold, italic, underline and some more formatting will be preserved converting it to Markdown coding before passing it to the LLM (most LLM support it)"},
                {"KeepFormat1", "If selected, the original's text basic character and paragraph formatting of a translated text will be retained (by HTML encoding, takes time!)"},
                {"ReplaceText1", "If selected, the response of the LLM for translations will replace the original text"},
                {"KeepFormat2", "If selected, the original's text basic character formatting will be retained for commands other than translations (by HTML encoding, takes time!)"},
                {"ReplaceText2", "If selected, the response of the LLM for other commands (than translate) will replace the original text"},
                {"ReplaceText2Override", "Leave empty to not override the above value; use 0 or 'false' to disable and 1 or 'true' to enable 'Replace text' as a personal override"},
                {"KeepParaFormatInline", "If selected, the basic formatting of each paragraph will be retained by encoding it into the text (takes time, but less time encoding HTML), unless 'Keep Format' is selected"},
                {"KeepFormatCap", "If a text has more characters, then the format will not be retained (to prevent having to wait too long)"},
                {"DoMarkupWord", "Whether a markup should be done for functions that change only parts of a text"},
                {"MarkupMethodHelper", "Which markup method to use: 1 = Word compare, 2 = Simple Differ, 3 = Diff shown in a window"},
                {"MarkupMethodWord", "Which markup method to use: 1 = Word compare, 2 = Simple Differ, 3 = Diff shown in a window, 4 = LLM-based Regex Markup"},
                {"MarkupMethodWordOverride", "Leave empty to not override the above value; otherwise enter the personal override value for 'markup method'"},
                {"MarkupDiffCap", "The maximum size of the text that should be processed using the Diff method (to avoid you having to wait too long)"},
                {"MarkupRegexCap", "The maximum size of the text that should be processed using the Regex method (to avoid you having to wait too long)"},
                {"MarkdownBubbles", $"If selected, Word bubbles created by {AN} will support Markdown formatting (if provided by the LLM)"},
                {"PreCorrection", "Add prompting text that will be added to all basic requests (e.g., for special language tasks)"},
                {"PostCorrection", "Add a prompt that will be applied to each result before it is further processed (slow!)"},
                {"Language1", "The language (in English) that will be used for the first quick access button in the ribbon"},
                {"Language2", "The language (in English) that will be used for the second quick access button in the ribbon"},
                {"PromptLibPath", "The filename (including path, support environmental variables) for your prompt library (if any)"},
                {"PromptLibPathLocal", "The filename (including path, support environmental variables) for your local prompt library (if any)"},
                {"PromptLibPath_Transcript", "The filename (including path, support environmental variables) for your transcript prompt library (if any)"},
                {"ShortcutsWordExcel", "You can add key shortcuts by giving the name of the context menu, e.g., 'Correct=Ctrl-Shift-C', separated by ';' (only works if context menus are enabled and the Word helper is installed)"},
                {"ChatCap", "Use this to limit how many characters of your past chat discussion the chatbot will memorize (for saving costs and time)"},
                {"MyStylePath", "This is the path where the prompts are stored that convey your writing style (if defined, see 'Analyze')."},
                {"DefaultPrefix", "You can define here the default prefix to use within 'Freestyle' if no other prefix is used (will be added automatically)."},
                {"Location", "Provide location information (e.g., 'We are in Zurich, Switzerland') to be used in 'Freestyle', chatbot and some other prompts that contain {Location} to get more location specific results."},
                {"ToolingLogWindow", $"When an LLM is allowed to call {ToolFriendlyName.ToLower} within Red Ink (e.g., Special Services), a log window will automatically open and show the progress."},
                {"ToolingDryRun", $"When an LLM is allowed to call {ToolFriendlyName.ToLower} within Red Ink (e.g., Special Services), the {ToolFriendlyName.ToLower} made available to the LLM will be shown first, allowing the user to decide whether to proceed."}
            }

        ShowSettingsWindow(Settings, SettingsTips)

        Dim splash As New Slib.Splashscreen("Updating menu following your changes ...")
        splash.Show()
        splash.Refresh()

        AddContextMenu()

        splash.Close()

    End Sub



End Class
