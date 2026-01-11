' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Commands.vb
' Purpose: Dispatches ribbon and menu commands for the Outlook add‑in. Provides
'          LLM-driven transformations (translate, summarize, improve, style,
'          shorten, answer, freestyle) on selected email text with optional
'          markup, formatting preservation, model switching, clipboard/object
'          insertion, and personal style application.
'
' Architecture:
' - Reentrancy control: inMainMenu single-entry guard in MainMenu prevents concurrent execution.
' - Initialization: configuration and cooldown checks precede command dispatch.
' - Command routing: Select Case RI_Command invokes specific LLM-backed operations (Command_InsertAfter, FreeStyle_* etc.).
' - COM interop: Outlook (Application, Explorer, Inspector, Selection, MailItem) and Word (Document, Selection, Range) objects acquired via ComRetry for resilience.
' - Selection handling: inline response selection offsets captured and reapplied when promoting to Inspector window.
' - Formatting: plain text mail rejected; HTML/RTF processed; optional HTML preservation or Markdown conversion before LLM requests.
' - Markup: three methods (1 Word compare, 2 Diff inline, 3 Diff window) with caps and override logic; Diff size threshold triggers user choice.
' - Freestyle: prefix-driven behavior (markup, in-place, clipboard, new document, secondary model, MyStyle, object inclusion); last prompt persistence.
' - Clipboard: robust multi-attempt STA-setting with verification; fallback to editable window or temp file write.
' - MyStyle: aggregates open Inspector mail bodies into tagged blocks (<EMAILxxx>) for LLM style analysis and prompt extraction.
' - Summarization/translation: selected or aggregated mail text wrapped in tags and processed; Markdown converted to HTML for display.
' - Utility: latest non-quoted mail body extraction via marker/regex heuristics; selection range acquisition; temporary file persistence; settings window construction.
' =============================================================================

Option Strict Off
Option Explicit On

Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Markdig
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook
Imports Microsoft.Office.Interop.Word
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Main command dispatcher. Guards reentrancy, ensures configuration is loaded, validates active MailItem,
    ''' and executes the requested RI_Command (translate, summarize, improve, style, markup, freestyle, etc.).
    ''' </summary>
    ''' <param name="RI_Command">Internal command string selecting operation.</param>
    Public Sub MainMenu(RI_Command As String)

        ' Acquire single-entry guard; if already in MainMenu, bail out
        If System.Threading.Interlocked.CompareExchange(inMainMenu, 1, 0) <> 0 Then Return

        Try
            If IsInResumeCooldown() Then
                SLib.ShowCustomMessageBox("Outlook is resuming from sleep. Please try again in a few seconds.")
                Return
            End If

            If Not INIloaded Then
                If Not StartupInitialized Then
                    Try
                        DelayedStartupTasks()
                        RemoveHandler outlookExplorer.Activate, AddressOf Explorer_Activate
                    Catch ex As System.Exception
                    End Try
                    If Not INIloaded Then Exit Sub
                Else
                    InitializeConfig(False, False)
                    If Not INIloaded Then
                        Exit Sub
                    End If
                End If
            End If

            InitializeConfig(False, False)

            If GPTSetupError OrElse INIValuesMissing() Or Not INIloaded Then Return

            ' Use fully qualified names to avoid ambiguity
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application

            Dim inspector As Microsoft.Office.Interop.Outlook.Inspector = GetActiveInspector()

            Dim Textlength As Long

            If inspector Is Nothing Then

                InspectorOpened = False

                OpenInspectorAndReapplySelection(RI_Command)

                If Not InspectorOpened Then Exit Sub

                inspector = ComRetry(Function() outlookApp.ActiveInspector())
                If inspector Is Nothing Then
                    System.Windows.Forms.MessageBox.Show("Error in MainMenu: No active email item found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            End If

            Dim curr As Object = ComRetry(Function() inspector.CurrentItem)
            If curr Is Nothing OrElse Not TypeOf curr Is Microsoft.Office.Interop.Outlook.MailItem Then
                SLib.ShowCustomMessageBox($"Please open an email for editing for using {AN}.")
                Return
            End If

            Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = CType(curr, Microsoft.Office.Interop.Outlook.MailItem)
            Dim wordEditor As Microsoft.Office.Interop.Word.Document = ComRetry(Function() CType(inspector.WordEditor, Microsoft.Office.Interop.Word.Document))

            Select Case RI_Command

                Case "Translate"
                    TranslateLanguage = ""
                    TranslateLanguage = SLib.ShowCustomInputBox("Enter your target language:", $"{AN} Translate", True, INI_Language2)
                    If String.IsNullOrEmpty(TranslateLanguage) Then Return
                    Command_InsertAfter(InterpolateAtRuntime(SP_Translate), False, INI_KeepFormat1, INI_ReplaceText1)
                Case "PrimLang"
                    TranslateLanguage = INI_Language1
                    Command_InsertAfter(InterpolateAtRuntime(SP_Translate), False, INI_KeepFormat1, INI_ReplaceText1)
                Case "Correct"
                    Command_InsertAfter(InterpolateAtRuntime(SP_Correct), INI_DoMarkupOutlook, INI_KeepFormat2, Override(INI_ReplaceText2, INI_ReplaceText2Override), Override(INI_MarkupMethodOutlook, INI_MarkupMethodOutlookOverride))
                Case "Summarize"

                    Textlength = GetSelectedTextLength()

                    If Textlength = 0 Then
                        SLib.ShowCustomMessageBox("Please select the text to be processed.")
                        Exit Sub
                    End If

                    Dim UserInput As String
                    SummaryLength = 0

                    Do
                        UserInput = Trim(SLib.ShowCustomInputBox("Enter the number of words your summary shall have (the selected text has " & Textlength & " words; the proposal " & SummaryPercent & "%):", $"{AN} Summarizer", True, CStr(Math.Round(SummaryPercent * Textlength / 100 / 5) * 5)))

                        If String.IsNullOrEmpty(UserInput) Then
                            Exit Sub
                        End If

                        If Integer.TryParse(UserInput, SummaryLength) AndAlso SummaryLength >= 1 AndAlso SummaryLength <= Textlength Then
                            Exit Do
                        Else
                            SLib.ShowCustomMessageBox("Please enter a valid word count between 1 and " & Textlength & ".")
                        End If
                    Loop
                    If SummaryLength = 0 Then Exit Sub
                    'SummaryLength = (Textlength - (Textlength * SummaryPercent / 100))'

                    Command_InsertAfter(InterpolateAtRuntime(SP_Summarize), False)
                Case "Improve"
                    Command_InsertAfter(InterpolateAtRuntime(SP_Improve), INI_DoMarkupOutlook, INI_KeepFormat2, Override(INI_ReplaceText2, INI_ReplaceText2Override), Override(INI_MarkupMethodOutlook, INI_MarkupMethodOutlookOverride))
                Case "NoFillers"
                    Command_InsertAfter(InterpolateAtRuntime(SP_NoFillers), INI_DoMarkupOutlook, INI_KeepFormat2, Override(INI_ReplaceText2, INI_ReplaceText2Override), Override(INI_MarkupMethodOutlook, INI_MarkupMethodOutlookOverride))
                Case "ApplyMyStyle"
                    Dim StylePath As String = ExpandEnvironmentVariables(INI_MyStylePath)

                    If String.IsNullOrWhiteSpace(StylePath) Then
                        ShowCustomMessageBox("You have not defined a MyStyle prompt file. Please do so first in the configuration file or using 'Settings'.")
                        Return
                    End If
                    If Not IO.File.Exists(StylePath) Then
                        ShowCustomMessageBox("No MyStyle prompt file has been found. You may have to first create a MyStyle prompt. Go to 'Analyze' and use 'Define MyStyle' to do so - exiting.")
                        Return
                    End If

                    Textlength = GetSelectedTextLength()
                    If Textlength = 0 Then
                        SLib.ShowCustomMessageBox("Please select the text to be processed.")
                        Return
                    End If

                    MyStyleInsert = MyStyleHelpers.SelectPromptFromMyStyle(StylePath, "Outlook", 0, "Choose the style prompt to apply …", $"{AN} MyStyle", False)
                    If MyStyleInsert = "ERROR" Then Return
                    If MyStyleInsert = "NONE" OrElse String.IsNullOrWhiteSpace(MyStyleInsert) Then
                        Return
                    End If

                    Command_InsertAfter(InterpolateAtRuntime(SP_MyStyle_Apply) & " " & MyStyleInsert, INI_DoMarkupOutlook, INI_KeepFormat2, Override(INI_ReplaceText2, INI_ReplaceText2Override), Override(INI_MarkupMethodOutlook, INI_MarkupMethodOutlookOverride))

                Case "Friendly"
                    Command_InsertAfter(InterpolateAtRuntime(SP_Friendly), INI_DoMarkupOutlook, INI_KeepFormat2, Override(INI_ReplaceText2, INI_ReplaceText2Override), Override(INI_MarkupMethodOutlook, INI_MarkupMethodOutlookOverride))
                Case "Convincing"
                    Command_InsertAfter(InterpolateAtRuntime(SP_Convincing), INI_DoMarkupOutlook, INI_KeepFormat2, Override(INI_ReplaceText2, INI_ReplaceText2Override), Override(INI_MarkupMethodOutlook, INI_MarkupMethodOutlookOverride))
                Case "Shorten"
                    Textlength = GetSelectedTextLength()
                    If Textlength = 0 Then
                        SLib.ShowCustomMessageBox("Please select the text to be processed.")
                        Exit Sub
                    End If
                    Dim UserInput As String
                    Dim ShortenPercentValue As Integer = 0
                    Do
                        UserInput = Trim(SLib.ShowCustomInputBox("Enter the percentage by which your text should be shortened (it has " & Textlength & " words; " & ShortenPercent & "% will cut approx. " & (Textlength * ShortenPercent / 100) & " words)", $"{AN} Shortener", True, CStr(ShortenPercent) & "%"))
                        If String.IsNullOrEmpty(UserInput) Then
                            Exit Sub
                        End If
                        UserInput = UserInput.Replace("%", "").Trim()
                        If Integer.TryParse(UserInput, ShortenPercentValue) AndAlso ShortenPercentValue >= 1 AndAlso ShortenPercentValue <= 99 Then
                            Exit Do
                        Else
                            SLib.ShowCustomMessageBox("Please enter a valid percentage between 1 And 99.")
                        End If
                    Loop
                    ShortenLength = (Textlength * (100 - ShortenPercentValue) / 100)
                    Command_InsertAfter(InterpolateAtRuntime(SP_Shorten), INI_DoMarkupOutlook, INI_KeepFormat2, Override(INI_ReplaceText2, INI_ReplaceText2Override), Override(INI_MarkupMethodOutlook, INI_MarkupMethodOutlookOverride))
                Case "Sumup"

                    Dim selectedText As String = mailItem.Body
                    ShowSumup(selectedText)

                'FreeStyle_InsertBefore(SP_MailSumup, False)
                Case "Answers"
                    FreeStyle_InsertBefore(SP_MailReply, True)
                Case "Freestyle"
                    FreeStyle_InsertAfter()
                Case "InsertClipboard"
                    InsertClipboard()
                Case Else
                    System.Windows.Forms.MessageBox.Show("Error in MainMenu: Invalid internal command.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Select

            If inspector IsNot Nothing Then Marshal.ReleaseComObject(inspector) : inspector = Nothing
            'If outlookApp IsNot Nothing Then Marshal.ReleaseComObject(outlookApp) : outlookApp = Nothing

        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show("Error in MainMenu: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Always release the reentrancy guard so subsequent calls work
            System.Threading.Interlocked.Exchange(inMainMenu, 0)
        End Try
    End Sub


    ''' <summary>
    ''' Returns the active Inspector if the current window is an Inspector; otherwise releases window and returns Nothing.
    ''' </summary>
    Private Function GetActiveInspector() As Outlook.Inspector
        Try
            Dim activeWindow = Globals.ThisAddIn.Application.ActiveWindow()
            If activeWindow IsNot Nothing AndAlso TypeOf activeWindow Is Outlook.Inspector Then
                ' The active window is an inspector, return it.
                Return CType(activeWindow, Outlook.Inspector)
            End If

            ' If the active window is not an inspector (e.g., Explorer) or there is no active window, return Nothing.
            If activeWindow IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(activeWindow)
            End If
            Return Nothing
        Catch
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Ensures a MailItem editing context by promoting inline response to Inspector if needed,
    ''' reapplying selection/caret offsets, and handling special commands (Sumup, Translate, InsertClipboard).
    ''' </summary>
    ''' <param name="Command">Internal command name.</param>
    Public Sub OpenInspectorAndReapplySelection(Command As String)
        Try

            If Command = "InsertClipboard" Then InsertClipboard() : Return

            Dim Sumup As Boolean = (Command = "Sumup")
            Dim Translate As Boolean = (Command = "Translate" OrElse Command = "PrimLang")

            ' Grab Outlook instances
            Dim oApp As Outlook.Application = Globals.ThisAddIn.Application
            Dim oExplorer As Outlook.Explorer = ComRetry(Function() oApp.ActiveExplorer())

            ' Check for inline response
            Dim inlineResponse As Object = If(oExplorer Is Nothing, Nothing,
                                              ComRetry(Function() oExplorer.ActiveInlineResponse))

            If inlineResponse Is Nothing OrElse Sumup OrElse Translate Then
                ' Get the current selection in the explorer
                Dim selection As Outlook.Selection = If(oExplorer Is Nothing, Nothing,
                                                       ComRetry(Function() oExplorer.Selection))
                Dim selectionCount As Integer = If(selection Is Nothing, 0, ComRetry(Function() selection.Count))

                If selectionCount = 0 Then
                    ShowCustomMessageBox("No email is selected.")
                    Return
                End If


                If selection.Count > 1 Then
                    If Not Sumup Then
                        ShowCustomMessageBox("Multiple emails selected. Please select only one email when not using Sumup mode.")
                        Return
                    Else
                        ' Combine texts from all selected emails.
                        Dim mailItems As New List(Of Microsoft.Office.Interop.Outlook.MailItem)
                        For Each item As Object In selection
                            If TypeOf item Is Microsoft.Office.Interop.Outlook.MailItem Then
                                mailItems.Add(CType(item, Microsoft.Office.Interop.Outlook.MailItem))
                            End If
                        Next

                        If mailItems.Count = 0 Then
                            ShowCustomMessageBox("None of the selected items are emails.")
                            Return
                        End If

                        ' Order the emails: latest email first (descending order by ReceivedTime)
                        mailItems = mailItems.OrderByDescending(Function(m) m.ReceivedTime).ToList()

                        Const PR_LAST_VERB_EXECUTED As String = "http://schemas.microsoft.com/mapi/proptag/0x10810003"

                        Dim selectedText As String = String.Empty
                        Dim count As Integer = 1
                        For Each mail As Microsoft.Office.Interop.Outlook.MailItem In mailItems

                            Dim lastVerb As Integer = 0
                            Try
                                lastVerb = mail.PropertyAccessor.GetProperty(PR_LAST_VERB_EXECUTED)
                            Catch comEx As COMException
                                ' Property not set → treat as not answered yet
                                lastVerb = 0
                            Catch ex As System.Exception
                                lastVerb = 0
                            End Try

                            ' Include:
                            ' - Unanswered (not Reply/ReplyAll)
                            ' - Answered-but-unread (Reply/ReplyAll AND UnRead=True)
                            Dim include As Boolean = False

                            ' Unanswered mails
                            If lastVerb <> 102 AndAlso lastVerb <> 103 Then
                                include = True
                            End If

                            ' Answered (Reply/ReplyAll) but still unread
                            If Not include AndAlso mail.UnRead AndAlso (lastVerb = 102 OrElse lastVerb = 103) Then
                                include = True
                            End If

                            ' If you also want to treat Forward as "answered", add lastVerb = 104 to the check above.

                            If include Then
                                Dim tag As String = count.ToString("D4") ' 0001, 0002, ...
                                Dim latestBody As String = GetLatestMailBody(mail.Body)
                                selectedText &= "<EMAIL" & tag & ">" & latestBody & "</EMAIL" & tag & ">"
                                count += 1
                            End If
                        Next

                        If String.IsNullOrWhiteSpace(selectedText) Then
                            ShowCustomMessageBox("No unanswered or answered-but-unread emails found in the selection.")
                            Return
                        End If

                        ShowSumup2(selectedText)
                        Return
                    End If
                Else
                    ' Only one email is selected.
                    If Sumup Then
                        Dim selectedItem As Object = selection(1)
                        If TypeOf selectedItem Is Outlook.MailItem Then
                            Dim mail As Outlook.MailItem = CType(selectedItem, Outlook.MailItem)
                            Dim selectedText As String = GetMailBody(mail)
                            ShowSumup(selectedText)
                            Return
                        Else
                            ShowCustomMessageBox("The selected item is not an email.")
                            Return
                        End If
                    ElseIf Translate Then
                        Dim selectedItem As Object = selection(1)
                        If TypeOf selectedItem Is Outlook.MailItem Then

                            If Command = "Translate" Then
                                TranslateLanguage = ""
                                TranslateLanguage = SLib.ShowCustomInputBox("Enter your target language:", $"{AN} Translate", True, INI_Language2)
                                If String.IsNullOrEmpty(TranslateLanguage) Then Return
                            Else
                                TranslateLanguage = INI_Language1
                            End If

                            ' In OpenInspectorAndReapplySelection, replace the Translate single-item part:
                            Dim mail As Outlook.MailItem = CType(selectedItem, Outlook.MailItem)
                            Dim selectedText As String = GetMailBody(mail)
                            ShowTranslate(selectedText)
                            Return

                        Else
                            ShowCustomMessageBox("The selected item is not an email.")
                            Return
                        End If
                    Else
                        ShowCustomMessageBox("You can only use this function when you are editing one (single) e-mail.")
                        Return
                    End If
                End If

            End If

            ' Ensure it is a MailItem
            Dim mailItem As MailItem = TryCast(inlineResponse, MailItem)
            If mailItem Is Nothing Then
                ShowCustomMessageBox("You can only use this function when you are editing an e-mail (currently, there is no valid e-mail item).")
                Return
            End If

            ' Capture the user's current selection range (or caret) from the inline editor
            Dim oldSelStart As Integer = 0
            Dim oldSelEnd As Integer = 0
            If Not GetSelectionOrCaretRangeFromInlineEditor(oExplorer, oldSelStart, oldSelEnd) Then
                ' If this fails entirely (no Word editor, etc.), proceed without reapplying selection.
            End If

            ' Open the Inspector modelessly
            Dim inspector As Inspector = mailItem.GetInspector
            If inspector Is Nothing Then
                MessageBox.Show("Error in OpenInspectorAndReapplySelection: Failed to open the ActiveInspector.",
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            inspector.Display(False) ' modeless - do not block

            ' A short delay to let the new WordEditor initialize
            System.Threading.Thread.Sleep(500)

            ' Ensure it is still open and usable (guard COM with retries)
            If inspector Is Nothing Then
                inspector = ComRetry(Function() Globals.ThisAddIn.Application.ActiveInspector())
                If inspector Is Nothing Then
                    MessageBox.Show("Error in OpenInspectorAndReapplySelection: No active Inspector available.",
                                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            End If

            Dim curr As Object = ComRetry(Function() inspector.CurrentItem)
            If curr Is Nothing OrElse Not TypeOf curr Is Microsoft.Office.Interop.Outlook.MailItem Then
                MessageBox.Show("Error in OpenInspectorAndReapplySelection: The Inspector is not ready or no email item is active.",
                                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' Reapply the original selection (or caret position) to the new Inspector's WordEditor
            Try
                Dim wordDoc As Word.Document = ComRetry(Function() TryCast(inspector.WordEditor, Word.Document))
                If wordDoc IsNot Nothing Then
                    Dim wordSel As Word.Selection = wordDoc.Application.Selection

                    ' Only reapply if offsets are non-zero
                    If oldSelStart <> 0 OrElse oldSelEnd <> 0 Then
                        wordSel.SetRange(Start:=oldSelStart, End:=oldSelEnd)
                        wordSel.Select()
                    End If
                End If

            Catch ex As System.Exception
                MessageBox.Show("Error in OpenInspectorAndReapplySelection: Failed to restore the original selection: " & ex.Message,
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End Try

            ' Activate Inspector
            InspectorOpened = True
            inspector.Activate()

            ' Clean up COM references
            Marshal.ReleaseComObject(inspector)
            Marshal.ReleaseComObject(oExplorer)

            Return

        Catch ex As System.Exception
            MessageBox.Show("Error in OpenInspectorAndReapplySelection: " & ex.Message,
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Function GetMailBody(mi As Outlook.MailItem) As String
        If mi Is Nothing Then Return ""

        Const PR_BODY As String = "http://schemas.microsoft.com/mapi/proptag/0x1000001E" ' PidTagBody (string)

        Try
            ' Prefer MAPI property (works even if item is open read-only elsewhere)
            Dim pa = ComRetry(Function() mi.PropertyAccessor)
            If pa IsNot Nothing Then
                Dim bodyObj As Object = ComRetry(Function() pa.GetProperty(PR_BODY))
                Dim body As String = TryCast(bodyObj, String)
                If Not String.IsNullOrEmpty(body) Then Return body
            End If
        Catch
            ' ignore -> fall back
        End Try

        Try
            ' Fallback to Outlook Body (may fail if locked, so keep it last)
            Return ComRetry(Function() mi.Body)
        Catch
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Extracts latest (non-quoted) mail body portion by scanning for known reply/forward markers and header patterns.
    ''' Returns full body if no marker encountered.
    ''' </summary>
    Private Function GetLatestMailBody(ByVal fullBody As String) As String
        Try
            ' Define an array of candidate markers that are common indicators of quoted messages,
            ' including localized variants.
            Dim markers() As String = {
            "-----Original Message-----",
            "-----Ursprüngliche Nachricht-----",
            "-----Vorherige Nachricht-----",
            "-----Mensaje original-----",
            "-----Messaggio originale-----",
            "-----Courrier original-----",
            "On ",
            "wrote:"
        }

            ' Regular expression to detect header lines with a proper email address
            Dim emailPattern As String = "^(From:|Von:|De:|Da:)\s+[\w\.-]+@[\w\.-]+\.\w+"

            ' Split the email body into lines
            Dim lines() As String = fullBody.Split(New String() {Environment.NewLine}, StringSplitOptions.None)
            Dim sb As New StringBuilder()

            For i As Integer = 0 To lines.Length - 1
                Dim currentLine As String = lines(i)
                Dim trimmedLine As String = currentLine.TrimStart()
                Dim isChainMarker As Boolean = False

                ' First, check each line against list of known chain markers.
                For Each marker As String In markers
                    If trimmedLine.StartsWith(marker, StringComparison.InvariantCultureIgnoreCase) Then
                        ' Only consider short lines (heuristically less than 100 characters) as markers.
                        If trimmedLine.Length < 100 Then
                            isChainMarker = True
                            Exit For
                        End If
                    End If
                Next

                ' If no marker found, test for header indicators.
                If Not isChainMarker Then
                    If Regex.IsMatch(trimmedLine, emailPattern, RegexOptions.IgnoreCase) Then
                        isChainMarker = True
                    Else
                        ' Additional check: headers with name or parenthesized comment.
                        Dim headerMarkers() As String = {"From:", "Von:", "De:", "Da:"}
                        For Each header As String In headerMarkers
                            If trimmedLine.StartsWith(header, StringComparison.InvariantCultureIgnoreCase) Then
                                Dim remainingText As String = trimmedLine.Substring(header.Length).Trim()
                                If remainingText.Contains(",") OrElse (remainingText.Contains("(") AndAlso remainingText.Contains(")")) Then
                                    isChainMarker = True
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                End If

                ' If marker detected, return accumulated lines.
                If isChainMarker Then
                    Return sb.ToString().TrimEnd()
                End If

                sb.AppendLine(currentLine)
            Next

            ' No marker found; return full body.
            Return fullBody
        Catch ex As System.Exception
            ' On error, return full body.
            Return fullBody
        End Try
    End Function


    ''' <summary>
    ''' Retrieves selection start/end offsets from inline response Word editor (Explorer context). Returns True if successful.
    ''' </summary>
    Private Function GetSelectionOrCaretRangeFromInlineEditor(oExplorer As Outlook.Explorer, ByRef selStart As Integer, ByRef selEnd As Integer) As Boolean
        Try
            Dim inlineWordEditor As Object =
                If(oExplorer Is Nothing, Nothing,
                   ComRetry(Function() oExplorer.ActiveInlineResponseWordEditor))
            If inlineWordEditor Is Nothing Then
                Return False
            End If

            Dim wordSel As Word.Selection =
                ComRetry(Function() TryCast(inlineWordEditor.Application.Selection, Word.Selection))
            If wordSel Is Nothing Then
                Return False
            End If

            selStart = wordSel.Start
            selEnd = wordSel.End
            Return True
        Catch ex As System.Exception
            MessageBox.Show("Failed to retrieve the selection: " & ex.Message,
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function


    ''' <summary>
    ''' Generates summary HTML for a mail chain: calls LLM with SP_MailSumup, optional post-correction, Markdown to HTML conversion, displays result.
    ''' </summary>
    Private Async Sub ShowSumup(selectedtext As String)

        Dim LLMResult As String = ""

        LLMResult = Await LLM(InterpolateAtRuntime(SP_MailSumup), "<MAILCHAIN>" & selectedtext & "</MAILCHAIN>", "", "", 0, EnsureUI:=False)

        If INI_PostCorrection <> "" Then
            LLMResult = Await PostCorrection(LLMResult)
        End If

        'Dim markdownPipeline As MarkdownPipeline = New MarkdownPipelineBuilder().Build()

        Dim builder As New MarkdownPipelineBuilder()

        builder.UsePipeTables()
        builder.UseGridTables()
        builder.UseSoftlineBreakAsHardlineBreak()
        builder.UseListExtras()
        builder.UseFootnotes()
        builder.UseDefinitionLists()
        builder.UseAbbreviations()
        builder.UseAutoLinks()
        builder.UseTaskLists()
        builder.UseEmojiAndSmiley()
        builder.UseMathematics()
        builder.UseFigures()
        builder.UseAdvancedExtensions()
        builder.UseGenericAttributes()

        Dim markdownPipeline As MarkdownPipeline = builder.Build()

        Dim htmlText As String = Markdown.ToHtml(LLMResult, markdownPipeline)

        Dim fullHtml As String =
              "<!DOCTYPE html>" &
              "<html><head>" &
              "  <meta charset=""utf-8"" />" &
              "  <style>" &
              "    ul { margin-left: 0.5em; padding-left: 0; list-style-position: outside; }" &
              "    ul ul { margin-left: 1em; padding-left: 0; list-style-type: circle; }" &
              "    ul ul ul { margin-left: 1.5em; padding-left: 0; list-style-type: square; }" &
              "  </style>" &
              "</head><body>" &
                htmlText &
              "</body></html>"

        ShowHTMLCustomMessageBox(fullHtml, $"{AN} Sum-up")

    End Sub

    ''' <summary>
    ''' Performs translation of selected text using SP_Translate, optional post-correction, converts Markdown to HTML, displays translation.
    ''' </summary>
    Private Async Sub ShowTranslate(selectedtext As String)

        Dim LLMResult As String = ""

        LLMResult = Await LLM(InterpolateAtRuntime(SP_Translate), "<TEXTTOPROCESS>" & selectedtext & "</TEXTTOPROCESS>", "", "", 0, EnsureUI:=False)

        If INI_PostCorrection <> "" Then
            LLMResult = Await PostCorrection(LLMResult)
        End If

        'Dim markdownPipeline As MarkdownPipeline = New MarkdownPipelineBuilder().Build()

        Dim builder As New MarkdownPipelineBuilder()

        builder.UsePipeTables()
        builder.UseGridTables()
        builder.UseSoftlineBreakAsHardlineBreak()
        builder.UseListExtras()
        builder.UseFootnotes()
        builder.UseDefinitionLists()
        builder.UseAbbreviations()
        builder.UseAutoLinks()
        builder.UseTaskLists()
        builder.UseEmojiAndSmiley()
        builder.UseMathematics()
        builder.UseFigures()
        builder.UseAdvancedExtensions()
        builder.UseGenericAttributes()

        Dim markdownPipeline As MarkdownPipeline = builder.Build()

        Dim htmlText As String = Markdown.ToHtml(LLMResult, markdownPipeline)

        ShowHTMLCustomMessageBox(htmlText, $"{AN} Translation")

    End Sub

    ''' <summary>
    ''' Summarizes multiple unanswered mails (tagged blocks) using SP_MailSumup2 with timestamp, optional post-correction, renders HTML summary.
    ''' </summary>
    Private Async Sub ShowSumup2(selectedtext As String)

        Dim LLMResult As String = ""

        DateTimeNow = DateTime.Now.ToString("yyyy-MMM-dd HH:mm")

        LLMResult = Await LLM(InterpolateAtRuntime(SP_MailSumup2), selectedtext, "", "", 0, EnsureUI:=False)

        If INI_PostCorrection <> "" Then
            LLMResult = Await PostCorrection(LLMResult)
        End If

        ' Dim markdownPipeline As MarkdownPipeline = New MarkdownPipelineBuilder().Build()

        Dim builder As New MarkdownPipelineBuilder()

        builder.UsePipeTables()
        builder.UseGridTables()
        builder.UseSoftlineBreakAsHardlineBreak()
        builder.UseListExtras()
        builder.UseFootnotes()
        builder.UseDefinitionLists()
        builder.UseAbbreviations()
        builder.UseAutoLinks()
        builder.UseTaskLists()
        builder.UseEmojiAndSmiley()
        builder.UseMathematics()
        builder.UseFigures()
        builder.UseAdvancedExtensions()
        builder.UseGenericAttributes()

        Dim markdownPipeline As MarkdownPipeline = builder.Build()

        Dim htmlText As String = Markdown.ToHtml(LLMResult, markdownPipeline)

        Dim fullHtml As String =
              "<!DOCTYPE html>" &
              "<html><head>" &
              "  <meta charset=""utf-8"" />" &
              "  <style>" &
              "    ul { margin-left: 0.5em; padding-left: 0; list-style-position: outside; }" &
              "    ul ul { margin-left: 1em; padding-left: 0; list-style-type: circle; }" &
              "    ul ul ul { margin-left: 1.5em; padding-left: 0; list-style-type: square; }" &
              "  </style>" &
              "</head><body>" &
                htmlText &
              "</body></html>"

        ShowHTMLCustomMessageBox(fullHtml, $"{AN} Sum-up (of unanswered/unread mails)")

    End Sub


    ''' <summary>
    ''' Defines personal style (MyStyle) by aggregating open mail bodies, calling LLM for style analysis,
    ''' displaying and storing extracted prompt, with optional alternate model usage.
    ''' </summary>
    Public Async Sub DefineMyStyle()

        Try
            ' --- Check MyStyle path (like Word) ---
            Dim stylePath As System.String = System.Environment.ExpandEnvironmentVariables(INI_MyStylePath)

            If System.String.IsNullOrWhiteSpace(stylePath) Then
                ShowCustomMessageBox("You have not configured a MyStyle prompt file path. Please do so in the configuration file or using 'Settings'.")
                Return
            End If

            ' --- Intro info box (adapted to Outlook workflow) ---
            Dim introLabel As System.String =
                $"You are about to have {AN} create a profile of your writing style from selected emails. There are six steps:" & vbCrLf & vbCrLf &
                "1. You will enter your name (used by the prompt to detect your mails)." & vbCrLf &
                "2. All currently open emails (including those opened from .MSG files) will be gathered as samples." & vbCrLf &
                "3. You can provide additional instructions (e.g., links or aspects to focus on)." & vbCrLf &
                "4. You select the model to perform the analysis (e.g., a reasoning model, Internet access if links are to be consulted)." & vbCrLf &
                "5. You can review and amend the analysis, including the final prompt for the AI to implement your style." & vbCrLf &
                $"6. The analysis will be saved to your personal MyStyle prompt file ({stylePath})."

            Dim answer As System.Int32 = ShowCustomYesNoBox(introLabel, "Continue", "Cancel", $"{AN} Define MyStyle (Outlook)",
                                                            extraButtonText:="Edit MyStyle prompt file",
                                                            extraButtonAction:=Sub()
                                                                                   SLib.ShowTextFileEditor(stylePath, "Edit your MyStyle prompt file (use 'Define MyStyle' to create new prompts automatically):")
                                                                               End Sub)
            If answer <> 1 Then
                Return
            End If

            ' --- Ask for Username (default = OS user) ---
            Dim defaultUser As System.String = System.Environment.UserName
            Username = SLib.ShowCustomInputBox("Please enter your name (will be used to identify your mails within mailchains):", $"{AN} Define MyStyle (Outlook)", True, defaultUser)
            If Username Is Nothing OrElse Username.Trim().Length = 0 Then
                ShowCustomMessageBox("No username provided - exiting.")
                Return
            End If
            Username = Username.Trim()

            ' --- Collect all open emails from Outlook inspectors ---
            Dim app As Outlook.Application = Globals.ThisAddIn.Application
            Dim inspectors As Outlook.Inspectors = ComRetry(Function() app.Inspectors)

            '            Dim mailItems As New System.Collections.Generic.List(Of Outlook.MailItem)()

            'For i As System.Int32 = 1 To inspectors.Count
            'Dim insp As Outlook.Inspector = inspectors.Item(i)
            'If insp IsNot Nothing AndAlso insp.CurrentItem IsNot Nothing Then
            '   If TypeOf insp.CurrentItem Is Outlook.MailItem Then
            '       Dim mi As Outlook.MailItem = CType(insp.CurrentItem, Outlook.MailItem)
            '       If mi IsNot Nothing Then
            '            mailItems.Add(mi)
            '         End If
            '      End If
            '   End If
            'Next

            Dim mailItems As New System.Collections.Generic.List(Of Outlook.MailItem)()

            ' Get count safely
            Dim inspCount As Integer = 0
            Try
                inspCount = ComRetry(Function() inspectors.Count)
            Catch
                inspCount = 0
            End Try

            For i As System.Int32 = 1 To inspCount
                Dim insp As Outlook.Inspector = Nothing
                Try
                    insp = ComRetry(Function() inspectors.Item(i))
                    If insp Is Nothing Then Continue For

                    Dim curr As Object = ComRetry(Function() insp.CurrentItem)
                    Dim mi As Outlook.MailItem = TryCast(curr, Outlook.MailItem)
                    If mi IsNot Nothing Then
                        ' Intentionally keep the MailItem reference; used later in this method.
                        mailItems.Add(mi)
                    End If
                Catch
                    ' Ignore and continue scanning remaining inspectors
                Finally
                    If insp IsNot Nothing Then
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(insp)
                        insp = Nothing
                    End If
                End Try
            Next

            If mailItems.Count = 0 Then
                ShowCustomMessageBox("No open emails were found. Please open all emails you want to include and try again.")
                Return
            End If

            ' --- Show list of all emails that will be included (via MessageBox), then explicit proceed confirm ---
            Dim sbList As New System.Text.StringBuilder()
            sbList.AppendLine("The following emails will be included:").AppendLine()
            For idx As System.Int32 = 0 To mailItems.Count - 1
                Dim mi As Outlook.MailItem = mailItems(idx)
                Dim subj As System.String = If(mi.Subject, "(no subject)")
                Dim sender As System.String = If(mi.SenderName, "(unknown sender)")
                Dim sentOn As System.String
                Try
                    sentOn = If(mi.SentOn = Date.MinValue, "(no sent date)", mi.SentOn.ToString())
                Catch ex As System.Exception
                    sentOn = "(no sent date)"
                End Try
                sbList.AppendLine($"{(idx + 1).ToString("000")}. {subj} — {sender} — {sentOn}")
            Next

            ShowCustomMessageBox(sbList.ToString())
            Dim confirm As System.Int32 = ShowCustomYesNoBox($"Proceed with these emails (the AI will get the full text and be instructed to learn only from those that refer to '{Username}')?", "Continue", "Cancel", $"{AN} Define MyStyle (Outlook)")
            If confirm <> 1 Then
                Return
            End If

            ' --- Additional instructions (like Word: ESC cancels) ---
            OtherPrompt = ""
            OtherPrompt = SLib.ShowCustomInputBox("You can provide additional instructions for the analysis (e.g., Internet links to check [if your model will understand so], aspects to focus on etc.). This is optional.",
                                                    $"{AN} Define MyStyle (Outlook)", False).Trim()
            If OtherPrompt = "ESC" Then
                Return
            End If

            ' --- Optional: use alternate model (like Word) ---
            Dim useSecondAPI As System.Boolean = False
            If Not System.String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                answer = ShowCustomYesNoBox($"Do you want to use one of your alternate models?", "Yes, use alternate", "No, use primary", $"{AN} Define MyStyle (Outlook)")
                If answer = 1 Then
                    If Not ShowModelSelection(_context, INI_AlternateModelPath) Then
                        originalConfigLoaded = False
                        Return
                    End If
                    useSecondAPI = True
                ElseIf answer <> 2 Then
                    Return
                End If
            End If

            ' --- Build SelectedText from all open emails ---
            ' Format: <EMAILxxx> Mailtext </EMAILxxx>, xxx = 001, 002, ...
            Dim sbEmails As New System.Text.StringBuilder()

            For idx As System.Int32 = 0 To mailItems.Count - 1
                Dim mi As Outlook.MailItem = mailItems(idx)

                Dim bodyText As System.String = If(mi.Body, System.String.Empty)

                If System.String.IsNullOrWhiteSpace(bodyText) Then
                    Dim html As System.String = If(mi.HTMLBody, System.String.Empty)
                    If Not System.String.IsNullOrWhiteSpace(html) Then
                        ' simple HTML -> text (strip tags, decode)
                        Dim noTags As System.String = System.Text.RegularExpressions.Regex.Replace(html, "<[^>]+>", System.String.Empty)
                        bodyText = System.Net.WebUtility.HtmlDecode(noTags)
                    End If
                End If

                If bodyText Is Nothing Then
                    bodyText = System.String.Empty
                End If

                bodyText = bodyText.Trim()

                Dim tagId As System.String = (idx + 1).ToString("000")
                sbEmails.Append("<EMAIL").Append(tagId).Append(">").AppendLine()
                sbEmails.Append(bodyText).AppendLine()
                sbEmails.Append("</EMAIL").Append(tagId).Append(">").AppendLine().AppendLine()
            Next

            Dim SelectedText As String = sbEmails.ToString()

            ' --- Call LLM with SP_MyStyle_Outlook (like Word) ---
            ' Note: SP_MyStyle_Outlook should already include {OtherPrompt} via InterpolateAtRuntime.
            Dim llmResponse As System.String =
                Await LLM(InterpolateAtRuntime(SP_MyStyle_Outlook), SelectedText, "", "", 0, useSecondAPI)

            ' --- Show analysis and (on OK) save prompt + copy full report to clipboard (like Word) ---
            If Not System.String.IsNullOrWhiteSpace(llmResponse) Then
                Dim analysis As System.String = SLib.ShowCustomWindow($"The AI provided the following style analysis for {Username} and MyStyle prompt of your email samples:",
                                                                        llmResponse,
                                                                        "If you choose 'OK', the prompt and its title at the end of the analysis will be stored in your MyStyle prompt file for future usage (and the full report copied to the clipboard).",
                                                                        AN, False, False, False, False)

                If Not System.String.IsNullOrWhiteSpace(analysis) Then
                    SLib.PutInClipboard(analysis)
                    SLib.ExtractAndStorePromptFromAnalysis(analysis, stylePath, "Outlook")
                End If
            End If

            If useSecondAPI AndAlso originalConfigLoaded Then
                RestoreDefaults(_context, originalConfig)
                originalConfigLoaded = False
            End If


        Catch ex As System.Exception
            ShowCustomMessageBox($"An error occurred: {ex.Message}")
        End Try

    End Sub


    ''' <summary>
    ''' Inserts LLM result before existing content (prepend) in current MailItem; optionally prompts user for answer instructions and applies MyStyle.
    ''' </summary>
    Private Async Sub FreeStyle_InsertBefore(Command As String, Optional AskForPrompt As Boolean = False)
        Try
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim inspector As Inspector = ComRetry(Function() outlookApp.ActiveInspector())

            ' Ensure the inspector is open and the item is a MailItem
            'If inspector Is Nothing OrElse Not TypeOf inspector.CurrentItem Is MailItem Then
            'SLib.ShowCustomMessageBox($"Please create or open an email for editing to use {AN}.")
            'Return
            'End If

            'Dim mailItem As MailItem = DirectCast(inspector.CurrentItem, MailItem)


            ' Guard CurrentItem via ComRetry to avoid transient COM rejections
            Dim curr As Object = Nothing
            If inspector IsNot Nothing Then
                Try
                    curr = ComRetry(Function() inspector.CurrentItem)
                Catch
                    curr = Nothing
                End Try
            End If

            If inspector Is Nothing _
               OrElse curr Is Nothing _
               OrElse Not TypeOf curr Is Microsoft.Office.Interop.Outlook.MailItem Then
                SLib.ShowCustomMessageBox($"Please create or open an email for editing to use {AN}.")
                Return
            End If

            Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem =
                CType(curr, Microsoft.Office.Interop.Outlook.MailItem)


            ' Check if the email is in plain text format
            If mailItem.BodyFormat = OlBodyFormat.olFormatPlain Then
                SLib.ShowCustomMessageBox("This operation is not supported for plain text emails. Switch to HTML or RTF format.")
                Return
            End If

            ' Get the Word editor for the email
            Dim wordEditor As Microsoft.Office.Interop.Word.Document = ComRetry(Function() TryCast(inspector.WordEditor, Microsoft.Office.Interop.Word.Document))

            If wordEditor Is Nothing Then
                SLib.ShowCustomMessageBox("Unable to access the necessary email editor. Ensure the email is in HTML or RTF format.")
                Return
            End If

            ' Get the selected text
            Dim selectedText As String = wordEditor.Application.Selection.Text
            If String.IsNullOrWhiteSpace(selectedText) Then
                selectedText = wordEditor.Content.Text
            End If

            OtherPrompt = ""
            Dim LLMResult As String = ""

            If AskForPrompt Then

                MyStyleInsert = ""
                Dim DoMyStyle As Boolean = False
                Dim StylePath As String = ExpandEnvironmentVariables(INI_MyStylePath)
                If Not String.IsNullOrWhiteSpace(StylePath) And IO.File.Exists(StylePath) Then DoMyStyle = True

                ' Prompt for additional instructions
                OtherPrompt = SLib.ShowCustomInputBox("Please provide additional instructions for drafting an answer (or leave it empty for the most likely substantive response):", $"{AN} Answers", False)
                If OtherPrompt = "ESC" Then Return

                If DoMyStyle Then
                    MyStyleInsert = MyStyleHelpers.SelectPromptFromMyStyle(StylePath, "Outlook", 0, "Choose the style prompt to apply …", $"{AN} MyStyle", True)
                    If MyStyleInsert = "ERROR" Then Return
                    If MyStyleInsert = "NONE" OrElse String.IsNullOrWhiteSpace(MyStyleInsert) Then DoMyStyle = False
                End If

                ' Call LLM function with selected text
                LLMResult = Await LLM(InterpolateAtRuntime(SP_MailReply) & If(DoMyStyle, " " & MyStyleInsert, ""), "<MAILCHAIN>" & selectedText & "</MAILCHAIN>", "", "", 0)
            Else
                LLMResult = Await LLM(InterpolateAtRuntime(SP_MailSumup), "<MAILCHAIN>" & selectedText & "</MAILCHAIN>", "", "", 0)
            End If
            If INI_PostCorrection <> "" Then
                LLMResult = Await PostCorrection(LLMResult)
            End If

            'LLMResult = LLMResult.Replace("**", "")  ' Remove bold markers

            ' Convert Markdown to HTML using Markdig
            ' Dim markdownPipeline As MarkdownPipeline = New MarkdownPipelineBuilder().Build()

            Dim builder As New MarkdownPipelineBuilder()

            builder.UsePipeTables()
            builder.UseGridTables()
            builder.UseSoftlineBreakAsHardlineBreak()
            builder.UseListExtras()
            builder.UseFootnotes()
            builder.UseDefinitionLists()
            builder.UseAbbreviations()
            builder.UseAutoLinks()
            builder.UseTaskLists()
            builder.UseEmojiAndSmiley()
            builder.UseMathematics()
            builder.UseFigures()
            builder.UseAdvancedExtensions()
            builder.UseGenericAttributes()

            Dim markdownPipeline As MarkdownPipeline = builder.Build()

            Dim convertedHtml As String = Markdown.ToHtml(LLMResult, markdownPipeline)

            If mailItem.BodyFormat = OlBodyFormat.olFormatHTML Then
                ' Ensure consistent font and style for HTML emails
                Dim defaultStyle As String = "<div style='font-family:Arial, sans-serif; font-size:11pt;'>" ' Adjust as needed
                Dim formattedResult As String = defaultStyle & convertedHtml & "</div><br/><br/>"

                ' Append the formatted result to the HTML body
                mailItem.HTMLBody = formattedResult & mailItem.HTMLBody
            Else
                ' Convert HTML to plain text for non-HTML formats (optional)
                Dim doc As New HtmlAgilityPack.HtmlDocument()
                doc.LoadHtml(convertedHtml)
                Dim plainTextResult As String = doc.DocumentNode.InnerText

                ' Standard handling for Plain Text and Rich Text
                mailItem.Body = plainTextResult & vbCrLf & vbCrLf & mailItem.Body
            End If

            ' Refresh the inspector to show updated content
            inspector.Display()

            ' Release COM objects in reverse order of creation
            If wordEditor IsNot Nothing Then Marshal.ReleaseComObject(wordEditor) : wordEditor = Nothing
            If mailItem IsNot Nothing Then Marshal.ReleaseComObject(mailItem) : mailItem = Nothing
            If inspector IsNot Nothing Then Marshal.ReleaseComObject(inspector) : inspector = Nothing
            'If outlookApp IsNot Nothing Then Marshal.ReleaseComObject(outlookApp) : outlookApp = Nothing

        Catch ex As System.Exception
            MessageBox.Show("Error in Freestyle_InsertBefore: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Processes selected text (or markup) after selection using SysCommand prompt. Handles formatting preservation, in-place replacement,
    ''' optional markup generation via configured method, and insertion strategy.
    ''' </summary>
    Private Async Sub Command_InsertAfter(ByVal SysCommand As String, Optional ByVal DoMarkup As Boolean = False, Optional KeepFormat As Boolean = False, Optional Inplace As Boolean = False, Optional MarkupMethod As Integer = 3)
        Try
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim inspector As Microsoft.Office.Interop.Outlook.Inspector = ComRetry(Function() outlookApp.ActiveInspector())

            ' Ensure the inspector is open and the item is a MailItem
            'If inspector Is Nothing OrElse Not TypeOf inspector.CurrentItem Is Microsoft.Office.Interop.Outlook.MailItem Then
            '   ShowCustomMessageBox("Please open an email to use this function.")
            '   Return
            'End If

            'Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = DirectCast(inspector.CurrentItem, Microsoft.Office.Interop.Outlook.MailItem)

            ' Guard CurrentItem via ComRetry to avoid transient COM rejections
            Dim curr As Object = Nothing
            If inspector IsNot Nothing Then
                Try
                    curr = ComRetry(Function() inspector.CurrentItem)
                Catch
                    curr = Nothing
                End Try
            End If

            If inspector Is Nothing _
               OrElse curr Is Nothing _
               OrElse Not TypeOf curr Is Microsoft.Office.Interop.Outlook.MailItem Then
                ShowCustomMessageBox("Please open an email to use this function.")
                Return
            End If

            Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem =
                CType(curr, Microsoft.Office.Interop.Outlook.MailItem)

            ' Check if the email is in plain text format
            If mailItem.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatPlain Then
                ShowCustomMessageBox("This operation is not supported for plain text emails. Switch to HTML or RTF format.")
                Return
            End If

            ' Get the Word editor for the email
            Dim wordEditor As Microsoft.Office.Interop.Word.Document = ComRetry(Function() TryCast(inspector.WordEditor, Microsoft.Office.Interop.Word.Document))

            If wordEditor Is Nothing Then
                ShowCustomMessageBox("Unable to access the email editor. Ensure the email is in HTML or RTF format.")
                Return
            End If

            ' Get the selected text and range
            Dim selection As Microsoft.Office.Interop.Word.Selection = wordEditor.Application.Selection
            Dim range As Microsoft.Office.Interop.Word.Range = selection.Range.Duplicate ' Duplicate to preserve original
            Dim SelectedText As String

            'Try
            'Using New WordUndoScope(wordEditor, $"{AN} Changes")

            If INI_KeepFormatCap > 0 Then If Len(selection.Text) > INI_KeepFormatCap Then KeepFormat = False

            If KeepFormat Then
                SelectedText = SLib.GetRangeHtml(selection.Range)
            Else
                If INI_MarkdownConvert Then ConvertRangeToMarkdown(selection.Range)
                SelectedText = selection.Text
            End If

            If String.IsNullOrWhiteSpace(SelectedText) Then
                ShowCustomMessageBox($"Please select the text to be processed.")
                Return
            End If

            If DoMarkup And MarkupMethod = 2 And Len(SelectedText) > INI_MarkupDiffCap Then
                Dim MarkupChange As Integer = SLib.ShowCustomYesNoBox($"The selected text exceeds the defined cap for the Diff markup method at {INI_MarkupDiffCap} chars (your selection has {Len(SelectedText)} chars). {If(KeepFormat, "This may be because HTML codes have been inserted to keep the formatting (you can turn this off in the settings). ", "")}. How do you want to continue?", "Use Diff in Window compare instead", "Use Diff")
                Select Case MarkupChange
                    Case 1
                        MarkupMethod = 3
                    Case 2
                        MarkupMethod = 2
                    Case Else
                        Exit Sub
                End Select
            End If

            Dim trailingCR As Boolean = SelectedText.EndsWith(vbCrLf) Or SelectedText.EndsWith(vbCr) Or SelectedText.EndsWith(vbLf)

            ' Call LLM function with selected text
            Dim LLMResult As String = Await LLM(SysCommand & If(KeepFormat, " " & SP_Add_KeepHTMLIntact, SP_Add_KeepInlineIntact), "<TEXTTOPROCESS>" & SelectedText & "</TEXTTOPROCESS>", "", "", 0)

            LLMResult = LLMResult.Replace("<TEXTTOPROCESS>", "").Replace("</TEXTTOPROCESS>", "")

            If INI_PostCorrection <> "" Then
                LLMResult = Await PostCorrection(LLMResult)
            End If

            ' Remove horizontal whitespace (incl. NBSP) between real newline tokens (CRLF/CR/LF)
            LLMResult = System.Text.RegularExpressions.Regex.Replace(LLMResult, "(\r\n|\r|\n)[^\S\r\n]+(\r\n|\r|\n)", "$1$2")

            Debug.WriteLine("TrailingCR=" & trailingCR)
            Debug.WriteLine($"Selection='{selection.Text}'")

            ' Replace the selected text with the processed result
            If Not String.IsNullOrWhiteSpace(LLMResult) Then
                If KeepFormat Then

                    Dim Plaintext As String = ""

                    SelectedText = selection.Text
                    SLib.InsertTextWithFormat(LLMResult, range, Inplace, Not trailingCR)
                    If DoMarkup Then
                        LLMResult = SLib.RemoveHTML(LLMResult)
                        If MarkupMethod <> 3 Then
                            range.Text = vbCrLf & vbCrLf & "MARKUP:" & vbCrLf
                        End If
                        range.Collapse(WdCollapseDirection.wdCollapseEnd)
                        selection.SetRange(range.Start, selection.End)

                        CompareAndInsertText(SelectedText, LLMResult, MarkupMethod = 3, "This is the markup of the text inserted:", True)
                    End If

                Else

                    If Inplace Then
                        If Not trailingCR And LLMResult.EndsWith(ControlChars.Lf) Then LLMResult = LLMResult.TrimEnd(ControlChars.Lf)
                        If Not trailingCR And LLMResult.EndsWith(ControlChars.Cr) Then LLMResult = LLMResult.TrimEnd(ControlChars.Cr)
                        If DoMarkup And MarkupMethod <> 3 Then
                            SLib.InsertTextWithMarkdown(selection, LLMResult & "<p>MARKUP:<br></p>", trailingCR)
                        Else
                            SLib.InsertTextWithMarkdown(selection, LLMResult, trailingCR)
                        End If
                    Else
                        ' Insert two new line breaks and select final position while preserving formatting.
                        Dim selRange As Microsoft.Office.Interop.Word.Range = selection.Range.Duplicate
                        Dim originalFont As Microsoft.Office.Interop.Word.Font = selRange.Font.Duplicate

                        selRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                        selRange.Text = vbCrLf & vbCrLf

                        Dim newStart As Integer = selRange.End - 2
                        Dim newEnd As Integer = selRange.End
                        selection.SetRange(newStart, newEnd)

                        selection.Font = originalFont

                        If DoMarkup And MarkupMethod <> 3 Then
                            SLib.InsertTextWithMarkdown(selection, LLMResult & "<p>MARKUP:<br></p>" & vbCrLf, trailingCR)
                        Else
                            SLib.InsertTextWithMarkdown(selection, LLMResult, trailingCR)
                        End If
                    End If

                    ' Use Find to locate the nearest line break backward and adjust selection
                    range = selection.Range
                    With range.Find
                        .Text = vbCrLf
                        .Forward = False
                        .MatchWildcards = False
                        If .Execute() Then
                            selection.SetRange(range.Start, selection.End)
                        End If
                    End With

                    ' Perform markup comparison and insertion if necessary
                    If DoMarkup Then
                        If MarkupMethod = 2 Or MarkupMethod = 3 Then
                            CompareAndInsertText(SelectedText, LLMResult, MarkupMethod = 3, "This is the markup of the text inserted:", True)
                        Else
                            CompareAndInsertTextCompareDocs(SelectedText, LLMResult)
                        End If

                    End If

                End If

            Else
                ShowCustomMessageBox("The LLM did not return any content to insert.")

            End If

            ' End Using

            'Catch ex As System.Exception
            '   Debug.WriteLine("Error in Undo: " & ex.Message)
            'End Try

            ' Refresh the inspector to show updated content
            inspector.Display()

            ' Release COM objects in reverse order of creation
            If range IsNot Nothing Then Marshal.ReleaseComObject(range) : range = Nothing
            If selection IsNot Nothing Then Marshal.ReleaseComObject(selection) : selection = Nothing
            If wordEditor IsNot Nothing Then Marshal.ReleaseComObject(wordEditor) : wordEditor = Nothing
            If mailItem IsNot Nothing Then Marshal.ReleaseComObject(mailItem) : mailItem = Nothing
            If inspector IsNot Nothing Then Marshal.ReleaseComObject(inspector) : inspector = Nothing
            'If outlookApp IsNot Nothing Then Marshal.ReleaseComObject(outlookApp) : outlookApp = Nothing

        Catch ex As System.Exception
            MessageBox.Show("Error in Command_InsertAfter: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Executes freestyle prompt logic: parses prefix triggers (markup, in-place, clipboard/new doc, MyStyle, secondary model),
    ''' processes selected or empty context, inserts or displays result, and performs optional markup comparison.
    ''' </summary>
    Private Async Sub FreeStyle_InsertAfter()
        Try

            Dim DoMarkup As Boolean = False
            Dim DoInplace As Boolean = False
            Dim DoClipboard As Boolean = False
            Dim NoText As Boolean = False
            Dim MarkupMethod As Integer = Override(INI_MarkupMethodOutlook, INI_MarkupMethodOutlookOverride)
            Dim KeepFormatCap = INI_KeepFormatCap ' currently not used
            Dim DoKeepFormat As Boolean = INI_KeepFormat2 ' currently not used
            Dim DoKeepParaFormat As Boolean = INI_KeepParaFormatInline ' currently not used
            Dim DoFileObject As Boolean = False
            Dim FileObject As String = ""
            Dim DoNewDoc As Boolean = False
            Dim DoMyStyle As Boolean = False

            Dim UseSecondAPI As Boolean = False

            Dim MarkupInstruct As String = $"start with '{MarkupPrefixAll}' for markups"
            Dim InplaceInstruct As String = $"use '{InPlacePrefix}' for replacing your current selection"
            Dim ClipboardInstruct As String = $"with '{ClipboardPrefix}'/'{ClipboardPrefix2}' or '{NewDocPrefix}' to have the result in a window or new Word document"
            Dim PromptLibInstruct As String = If(INI_PromptLib, " or press 'OK' for the prompt library", "")
            Dim NoFormatInstruct As String = $"; add '{NoFormatTrigger2}'/'{KFTrigger2}'/'{KPFTrigger2}' for overriding formatting defaults"
            Dim MyStyleInstruct As String = $"; add '{MyStyleTrigger}' to apply your personal style"
            Dim SecondAPIInstruct As String = If(INI_SecondAPI, $"'{SecondAPICode}' to use {If(String.IsNullOrWhiteSpace(INI_AlternateModelPath), $"the secondary model ({INI_Model_2})", "one of the other models")}", "")
            Dim LastPromptInstruct As String = If(String.IsNullOrWhiteSpace(My.Settings.LastPrompt), "", "; Ctrl-P for your last prompt")
            Dim ObjectInstruct As String = $"; add '{ObjectTrigger2}' for including a clipboard object"

            Dim AddOnInstruct As String = "; add " & SecondAPIInstruct

            Dim lastCommaIndex As Integer = AddOnInstruct.LastIndexOf(","c)
            If lastCommaIndex <> -1 Then
                AddOnInstruct = AddOnInstruct.Substring(0, lastCommaIndex) & ", and" & AddOnInstruct.Substring(lastCommaIndex + 1)
            End If
            If Not String.IsNullOrWhiteSpace(INI_MyStylePath) Then
                AddOnInstruct += MyStyleInstruct.Replace("; add ", ", ")
            End If

            Dim DefaultPrefix As String = INI_DefaultPrefix
            Dim DefaultPrefixText As String = ""

            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim inspector As Microsoft.Office.Interop.Outlook.Inspector = ComRetry(Function() outlookApp.ActiveInspector())

            ' Ensure the inspector is open and the item is a MailItem
            'If inspector Is Nothing OrElse Not TypeOf inspector.CurrentItem Is Microsoft.Office.Interop.Outlook.MailItem Then
            '  SLib.ShowCustomMessageBox($"Please create or open an email for editing to use {AN}.")
            '   Return
            'End If

            'Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = DirectCast(inspector.CurrentItem, Microsoft.Office.Interop.Outlook.MailItem)

            ' Guard CurrentItem via ComRetry to avoid transient COM rejections
            Dim curr As Object = Nothing
            If inspector IsNot Nothing Then
                Try
                    curr = ComRetry(Function() inspector.CurrentItem)
                Catch
                    curr = Nothing
                End Try
            End If

            If inspector Is Nothing _
               OrElse curr Is Nothing _
               OrElse Not TypeOf curr Is Microsoft.Office.Interop.Outlook.MailItem Then
                SLib.ShowCustomMessageBox($"Please create or open an email for editing to use {AN}.")
                Return
            End If

            Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem =
                CType(curr, Microsoft.Office.Interop.Outlook.MailItem)

            ' Check if the email is in plain text format
            If mailItem.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatPlain Then
                SLib.ShowCustomMessageBox("This operation is not supported for plain text emails. Switch to HTML or RTF format.")
                Return
            End If

            ' Get the Word editor for the email
            Dim wordEditor As Microsoft.Office.Interop.Word.Document = ComRetry(Function() TryCast(inspector.WordEditor, Microsoft.Office.Interop.Word.Document))

            If wordEditor Is Nothing Then
                SLib.ShowCustomMessageBox("Unable to access the necessary email editor. Ensure the email is in HTML or RTF format.")
                Return
            End If

            ' Get the selected text
            Dim selection As Microsoft.Office.Interop.Word.Selection = wordEditor.Application.Selection
            Dim selectedText As String = selection.Text
            If String.IsNullOrWhiteSpace(selectedText) Then
                NoText = True
            End If

            If UseSecondAPI Then
                If Not String.IsNullOrWhiteSpace(INI_APICall_Object_2) Then
                    AddOnInstruct += ObjectInstruct.Replace("; add", ",")
                    DoFileObject = True
                End If
            Else
                If Not String.IsNullOrWhiteSpace(INI_APICall_Object) Then
                    AddOnInstruct += ObjectInstruct.Replace("; add", ",")
                    DoFileObject = True
                End If
            End If

            If DefaultPrefix.Trim() <> "" Then
                DefaultPrefixText = $" (default prefix: '{DefaultPrefix}')"
            End If

            ' Prompt for the text to process

            If Not NoText Then
                Dim OptionalButtons As System.Tuple(Of String, String, String)() = {
                            System.Tuple.Create("OK, use window", $"Use this to automatically insert '{ClipboardPrefix}' as a prefix.", ClipboardPrefix),
                            System.Tuple.Create("OK, do a new doc", $"Use this to automatically insert '{NewDocPrefix}' as a prefix.", NewDocPrefix),
                            System.Tuple.Create("OK, do a markup", $"Use this to automatically insert '{MarkupPrefixDiff}' as a prefix.", MarkupPrefixDiff)
                        }
                OtherPrompt = SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute on the selected text ({MarkupInstruct}, {InplaceInstruct}, {ClipboardInstruct}){PromptLibInstruct}{AddOnInstruct}{LastPromptInstruct}{DefaultPrefixText}:", $"{AN} Freestyle", False, "", My.Settings.LastPrompt, OptionalButtons)
            Else
                Dim OptionalButtons As System.Tuple(Of String, String, String)() = {
                            System.Tuple.Create("OK, use window", $"Use this to automatically insert '{ClipboardPrefix}' as a prefix.", ClipboardPrefix),
                            System.Tuple.Create("OK, do a new doc", $"Use this to automatically insert '{NewDocPrefix}' as a prefix.", NewDocPrefix)
                        }

                OtherPrompt = SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute ({ClipboardInstruct}){PromptLibInstruct}{AddOnInstruct}{LastPromptInstruct}{DefaultPrefixText}:", $"{AN} Freestyle", False, "", My.Settings.LastPrompt, OptionalButtons)
            End If

            If String.IsNullOrEmpty(OtherPrompt) AndAlso OtherPrompt <> "ESC" AndAlso INI_PromptLib Then

                Dim promptlibresult As (String, Boolean, Boolean, Boolean)

                promptlibresult = ShowPromptSelector(INI_PromptLibPath, INI_PromptLibPathLocal, Not NoText, Nothing)

                OtherPrompt = promptlibresult.Item1
                DoMarkup = promptlibresult.Item2
                DoClipboard = promptlibresult.Item4

                If OtherPrompt = "" Then
                    Exit Sub
                End If
            Else
                If String.IsNullOrEmpty(OtherPrompt) Or OtherPrompt = "ESC" Then Exit Sub
            End If

            ' Check if OtherPrompt starts with a word ending with a colon
            If Not String.IsNullOrWhiteSpace(OtherPrompt) Then
                Dim firstWord As String = OtherPrompt.Split({" "c}, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()
                If firstWord IsNot Nothing AndAlso Not firstWord.EndsWith(":"c) Then

                    Dim prefix As String = DefaultPrefix.Trim()

                    ' Ensure prefix ends with colon and space
                    If prefix <> "" AndAlso Not prefix.EndsWith(":"c) Then
                        prefix &= ":"
                    End If

                    OtherPrompt = prefix & " " & OtherPrompt.Trim()
                    OtherPrompt = OtherPrompt.Trim()
                End If
            End If


            My.Settings.LastPrompt = OtherPrompt
            My.Settings.Save()

            If Not SharedMethods.ProcessParameterPlaceholders(OtherPrompt) Then
                ShowCustomMessageBox("Freestyle canceled.", $"{AN} Freestyle")
                Exit Sub
            End If

            ' Check prefixes for behavior

            If OtherPrompt.StartsWith(ClipboardPrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(ClipboardPrefix.Length).Trim()
                DoClipboard = True
            ElseIf OtherPrompt.StartsWith(ClipboardPrefix2, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(ClipboardPrefix2.Length).Trim()
                DoClipboard = True
            ElseIf OtherPrompt.StartsWith(NewDocPrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(NewDocPrefix.Length).Trim()
                DoClipboard = True
                DoNewDoc = True

            ElseIf OtherPrompt.StartsWith(MarkupPrefix, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                OtherPrompt = OtherPrompt.Substring(MarkupPrefix.Length).Trim()
                DoMarkup = True
            ElseIf OtherPrompt.StartsWith(MarkupPrefixWord, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                OtherPrompt = OtherPrompt.Substring(MarkupPrefixWord.Length).Trim()
                DoMarkup = True
                MarkupMethod = 1
            ElseIf OtherPrompt.StartsWith(MarkupPrefixDiffW, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                OtherPrompt = OtherPrompt.Substring(MarkupPrefixDiffW.Length).Trim()
                DoMarkup = True
                MarkupMethod = 3
            ElseIf OtherPrompt.StartsWith(MarkupPrefixDiff, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                OtherPrompt = OtherPrompt.Substring(MarkupPrefixDiff.Length).Trim()
                DoMarkup = True
                MarkupMethod = 2
            ElseIf OtherPrompt.StartsWith(InPlacePrefix, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                OtherPrompt = OtherPrompt.Substring(InPlacePrefix.Length).Trim()
                DoMarkup = False
                MarkupMethod = 3
                DoInplace = True
            End If

            ' Formatting Trigger (currently not used)

            If OtherPrompt.IndexOf(NoFormatTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(NoFormatTrigger, "").Trim()
                KeepFormatCap = 1
            End If
            If OtherPrompt.IndexOf(NoFormatTrigger2, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(NoFormatTrigger2, "").Trim()
                KeepFormatCap = 1
            End If
            If OtherPrompt.IndexOf(KFTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(KFTrigger, "").Trim()
                DoKeepFormat = True
            End If
            If OtherPrompt.IndexOf(KFTrigger2, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(KFTrigger2, "").Trim()
                DoKeepFormat = True
            End If
            If OtherPrompt.IndexOf(KPFTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(KPFTrigger, "").Trim()
                DoKeepParaFormat = True
            End If
            If OtherPrompt.IndexOf(KPFTrigger2, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(KPFTrigger2, "").Trim()
                DoKeepParaFormat = True
            End If
            If DoFileObject AndAlso OtherPrompt.IndexOf(ObjectTrigger2, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(ObjectTrigger2, "(a file object follows)").Trim()
                FileObject = "clipboard"
            End If

            If Not String.IsNullOrWhiteSpace(INI_MyStylePath) And OtherPrompt.IndexOf(MyStyleTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                Dim StylePath As String = ExpandEnvironmentVariables(INI_MyStylePath)
                If Not IO.File.Exists(StylePath) Then
                    ShowCustomMessageBox("No MyStyle prompt file has been found. You may have to first create a MyStyle prompt. Go to 'Analyze' and use 'Define MyStyle' to do so - exiting.")
                    Return
                End If
                OtherPrompt = OtherPrompt.Replace(MyStyleTrigger, "").Trim()
                MyStyleInsert = MyStyleHelpers.SelectPromptFromMyStyle(StylePath, "Outlook", 0, "Choose the style prompt to apply …", $"{AN} MyStyle", True)
                If MyStyleInsert = "ERROR" Then Return
                If MyStyleInsert = "NONE" OrElse String.IsNullOrWhiteSpace(MyStyleInsert) Then Return
                DoMyStyle = True
            End If

            If INI_SecondAPI Then
                If OtherPrompt.Contains(SecondAPICode) Then
                    UseSecondAPI = True
                    OtherPrompt = OtherPrompt.Replace(SecondAPICode, "").Trim()

                    If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then

                        If Not ShowModelSelection(_context, INI_AlternateModelPath) Then
                            originalConfigLoaded = False
                            Return
                        End If

                    End If

                End If
            End If

            If DoMarkup And MarkupMethod = 2 And Len(selectedText) > INI_MarkupDiffCap Then
                Dim MarkupChange As Integer = SLib.ShowCustomYesNoBox($"The selected text exceeds the defined cap for the Diff markup method at {INI_MarkupDiffCap} chars (your selection has {Len(selectedText)} chars). How do you want to continue?", "Use Diff in Window compare instead", "Use Diff")
                Select Case MarkupChange
                    Case 1
                        MarkupMethod = 3
                    Case 2
                        MarkupMethod = 2
                    Case Else
                        Exit Sub
                End Select
            End If

            Dim trailingCR As Boolean = selectedText.EndsWith(vbCrLf)

            ' Call LLM function with selected text

            Dim LLMResult As String

            If Not NoText Then
                LLMResult = Await LLM(InterpolateAtRuntime(SP_FreestyleText) & If(DoMyStyle, " " & MyStyleInsert, ""), "<TEXTTOPROCESS>" & selectedText & "</TEXTTOPROCESS>", "", "", 0, UseSecondAPI, False, OtherPrompt, FileObject)

                LLMResult = LLMResult.Replace("<TEXTTOPROCESS>", "").Replace("</TEXTTOPROCESS>", "")
            Else
                LLMResult = Await LLM(InterpolateAtRuntime(SP_FreestyleNoText) & If(DoMyStyle, " " & MyStyleInsert, ""), "", "", "", 0, UseSecondAPI, False, OtherPrompt, FileObject)
            End If

            If INI_PostCorrection <> "" Then
                LLMResult = Await PostCorrection(LLMResult)
            End If

            OtherPrompt = ""

            If DoNewDoc Then
                Try
                    ' Create a new instance of Word
                    Dim wordApp As New Microsoft.Office.Interop.Word.Application
                    wordApp.Visible = True

                    ' Add a new document
                    Dim newDoc As Microsoft.Office.Interop.Word.Document = wordApp.Documents.Add()

                    ' Insert text at beginning
                    Dim docSelection As Microsoft.Office.Interop.Word.Selection = wordApp.Selection
                    InsertTextWithMarkdown(docSelection, LLMResult, True)

                Catch Ex As System.Exception
                    Dim FinalText As String = SLib.ShowCustomWindow("The Word document could not be created or the LLM output not inserted. Here is the result of the LLM (you can edit it):", LLMResult, "You can choose whether you want to have the original text put into the clipboard or your text with any changes you have made. If you select Cancel, nothing will be put into the clipboard (without formatting).", AN, False)

                    If FinalText <> "" Then
                        SLib.PutInClipboard(FinalText)
                    End If

                End Try

            ElseIf DoClipboard Then
                Dim FinalText As String = SLib.ShowCustomWindow("The LLM has provided the following result (you can edit it):", LLMResult, "You can choose whether you want to have the original text put into the clipboard or your text with any changes you have made. If you select Cancel, nothing will be put into the clipboard (without formatting).", AN, False)

                If FinalText <> "" Then
                    SLib.PutInClipboard(FinalText)
                End If
            Else
                ' Collapse selection to end if not in-place
                If Not DoInplace Then
                    selection.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                Else
                    If Not trailingCR And LLMResult.EndsWith(ControlChars.Lf) Then LLMResult = LLMResult.TrimEnd(ControlChars.Lf)
                    If Not trailingCR And LLMResult.EndsWith(ControlChars.Cr) Then LLMResult = LLMResult.TrimEnd(ControlChars.Cr)
                End If

                ' Insert result
                If DoMarkup And MarkupMethod <> 3 Then
                    SLib.InsertTextWithMarkdown(selection, vbCrLf & LLMResult & vbCrLf & "<p>MARKUP:<br></p>", trailingCR)
                Else
                    If DoInplace Then
                        SLib.InsertTextWithMarkdown(selection, LLMResult, trailingCR)
                    Else
                        SLib.InsertTextWithMarkdown(selection, vbCrLf & LLMResult & vbCrLf, trailingCR)
                    End If
                End If

                ' Adjust selection using backward line break find
                Dim range As Microsoft.Office.Interop.Word.Range = selection.Range
                With range.Find
                    .Text = vbCrLf
                    .Forward = False
                    .MatchWildcards = False
                    If .Execute() Then
                        selection.SetRange(range.Start, selection.End)
                    End If
                End With

                ' Markup comparison if requested
                If DoMarkup Then
                    If MarkupMethod = 2 Or MarkupMethod = 3 Then
                        CompareAndInsertText(selectedText, LLMResult, MarkupMethod = 3, "This is the markup of the text inserted:", True)
                    Else
                        CompareAndInsertTextCompareDocs(selectedText, LLMResult)
                    End If
                End If
            End If

            ' Refresh the inspector to show updated content
            inspector.Display()

            ' Release COM objects in reverse order of creation
            If selection IsNot Nothing Then Marshal.ReleaseComObject(selection) : selection = Nothing
            If wordEditor IsNot Nothing Then Marshal.ReleaseComObject(wordEditor) : wordEditor = Nothing
            If mailItem IsNot Nothing Then Marshal.ReleaseComObject(mailItem) : mailItem = Nothing
            If inspector IsNot Nothing Then Marshal.ReleaseComObject(inspector) : inspector = Nothing
            'If outlookApp IsNot Nothing Then Marshal.ReleaseComObject(outlookApp) : outlookApp = Nothing

            If UseSecondAPI And originalConfigLoaded Then
                RestoreDefaults(_context, originalConfig)
                originalConfigLoaded = False
            End If

        Catch ex As System.Exception
            MessageBox.Show("Error in Freestyle_InsertAfter: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    ''' <summary>
    ''' Inserts clipboard-derived object/content using SP_InsertClipboard; handles Inspector vs Explorer context,
    ''' robust clipboard setting, optional RTF conversion, and fallback paths with temp file persistence.
    ''' </summary>
    Private Async Function InsertClipboard() As System.Threading.Tasks.Task
        Try
            If System.String.IsNullOrWhiteSpace(INI_APICall_Object) Then
                SLib.ShowCustomMessageBox($"Your model ({INI_Model}) is not configured to process clipboard data (binary/object).")
                Return
            End If

            ' Acquire result off the UI thread
            Dim result As String = Await LLM(
                InterpolateAtRuntime(SP_InsertClipboard),
                "", "", "", 0,
                UseSecondAPI:=False,
                HideSplash:=False,
                AddUserPrompt:="",
                FileObject:="clipboard",
                cancellationToken:=Nothing,
                EnsureUI:=False
            ).ConfigureAwait(False)

            If String.IsNullOrWhiteSpace(result) Then Return

            ' Determine Outlook context
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim inspector As Microsoft.Office.Interop.Outlook.Inspector = ComRetry(Function() outlookApp.ActiveInspector())
            Dim curr As Object = Nothing
            If inspector IsNot Nothing Then
                Try : curr = ComRetry(Function() inspector.CurrentItem) : Catch : curr = Nothing : End Try
            End If

            Dim haveMail As Boolean =
                inspector IsNot Nothing AndAlso
                curr IsNot Nothing AndAlso
                TypeOf curr Is Microsoft.Office.Interop.Outlook.MailItem

            If Not haveMail Then
                ' Explorer context: copy to clipboard (with UI switch only for message box)
                Dim displayText As String = If(result.Length > 11000, result.Substring(0, 11000) & "…", result)

                ' Build DataObject outside UI (RTF conversion may be expensive)
                Dim dataObj As New System.Windows.Forms.DataObject()
                Dim includeRtf As Boolean = result.Length < 350000 ' Skip massive RTF
                If includeRtf Then
                    Try
                        Dim rtfText = MarkdownToRtfConverter.Convert(result)
                        If Not String.IsNullOrEmpty(rtfText) Then
                            dataObj.SetData(System.Windows.Forms.DataFormats.Rtf, rtfText)
                        End If
                    Catch
                        ' Ignore RTF failure
                    End Try
                End If
                dataObj.SetData(System.Windows.Forms.DataFormats.UnicodeText, result)
                dataObj.SetData(System.Windows.Forms.DataFormats.Text, result)

                Dim clipOk = Await SetClipboardRobustAsync(dataObj).ConfigureAwait(False)

                Await SwitchToUi(
                    Sub()
                        If clipOk Then
                            SLib.ShowCustomMessageBox($"The content has been copied to the clipboard:{Environment.NewLine}{Environment.NewLine}{displayText}")
                        Else
                            ' Fallback window
                            Dim edited As String = SLib.ShowCustomWindow(
                                "Clipboard is busy. You can copy the result below manually (Ctrl+A, Ctrl+C) or edit it and click OK:",
                                result,
                                "If copying still fails, the text will be saved to a temporary file.",
                                AN, False)

                            If Not String.IsNullOrWhiteSpace(edited) Then
                                Dim editedObj As New DataObject()
                                editedObj.SetData(DataFormats.UnicodeText, edited)
                                editedObj.SetData(DataFormats.Text, edited)
                                Dim editedOk = SetClipboardRobustAsync(editedObj).GetAwaiter().GetResult()
                                If Not editedOk Then
                                    Dim tmp = SaveTextToTempFile(edited)
                                    If Not String.IsNullOrWhiteSpace(tmp) Then
                                        SLib.ShowCustomMessageBox($"Clipboard is locked. The result was saved to: {tmp}")
                                    Else
                                        SLib.ShowCustomMessageBox("Clipboard is locked and saving failed.")
                                    End If
                                Else
                                    SLib.ShowCustomMessageBox("Your edited text has been copied to the clipboard.")
                                End If
                            End If
                        End If
                    End Sub).ConfigureAwait(False)

                Return
            End If

            ' Inspector context: insert at cursor
            Dim wordEditor As Microsoft.Office.Interop.Word.Document =
                ComRetry(Function() CType(inspector.WordEditor, Microsoft.Office.Interop.Word.Document))

            If wordEditor Is Nothing Then
                ' Fallback to clipboard path
                Dim txtObj As New DataObject()
                txtObj.SetData(DataFormats.UnicodeText, result)
                txtObj.SetData(DataFormats.Text, result)
                Dim clipOk = Await SetClipboardRobustAsync(txtObj).ConfigureAwait(False)

                Await SwitchToUi(
                    Sub()
                        If clipOk Then
                            SLib.ShowCustomMessageBox("Could not access the mail editor; result copied to clipboard instead.")
                        Else
                            Dim edited As String = SLib.ShowCustomWindow(
                                "Could not access the mail editor; clipboard is busy. Copy manually or edit and press OK:",
                                result,
                                "If copying still fails, the text will be saved to a temporary file.",
                                AN, False)

                            If Not String.IsNullOrWhiteSpace(edited) Then
                                Dim editedObj As New DataObject()
                                editedObj.SetData(DataFormats.UnicodeText, edited)
                                editedObj.SetData(DataFormats.Text, edited)
                                Dim editedOk = SetClipboardRobustAsync(editedObj).GetAwaiter().GetResult()
                                If Not editedOk Then
                                    Dim tmp = SaveTextToTempFile(edited)
                                    If Not String.IsNullOrWhiteSpace(tmp) Then
                                        SLib.ShowCustomMessageBox($"Clipboard is locked. The result was saved to: {tmp}")
                                    Else
                                        SLib.ShowCustomMessageBox("Clipboard is locked and saving failed.")
                                    End If
                                Else
                                    SLib.ShowCustomMessageBox("Your edited text has been copied to the clipboard.")
                                End If
                            End If
                        End If
                    End Sub).ConfigureAwait(False)
                Return
            End If

            Dim selection As Microsoft.Office.Interop.Word.Selection = wordEditor.Application.Selection
            If selection IsNot Nothing Then
                If selection.Start <> selection.End Then
                    selection.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                End If
                selection.TypeParagraph()
                InsertTextWithMarkdown(selection, result, True)
            End If

            inspector.Display()

            If selection IsNot Nothing Then Marshal.ReleaseComObject(selection) : selection = Nothing
            If wordEditor IsNot Nothing Then Marshal.ReleaseComObject(wordEditor) : wordEditor = Nothing
            If inspector IsNot Nothing Then Marshal.ReleaseComObject(inspector) : inspector = Nothing

        Catch ex As System.Runtime.InteropServices.ExternalException
            SLib.ShowCustomMessageBox($"InsertClipboard clipboard error (probably locked): {ex.Message}")
        Catch ex As System.Exception
            SLib.ShowCustomMessageBox($"InsertClipboard failed: {ex.GetType().FullName}: {ex.Message}")
        End Try
    End Function

    ''' <summary>
    ''' Attempts to set clipboard persistently with retries on UI and dedicated STA thread. Returns True if set and verified.
    ''' </summary>
    Private Async Function SetClipboardRobustAsync(
        dataObj As System.Windows.Forms.DataObject,
        Optional uiAttempts As Integer = 3,
        Optional staAttempts As Integer = 5,
        Optional verificationText As String = Nothing
    ) As System.Threading.Tasks.Task(Of Boolean)

        ' Try on current thread (if STA) with persistent copy.
        For i As Integer = 1 To uiAttempts
            If TrySetClipboardImmediate(dataObj, verificationText) Then
                Return True
            End If
            Await System.Threading.Tasks.Task.Delay(50 * i).ConfigureAwait(False)
        Next

        ' STA thread attempts.
        For i As Integer = 1 To staAttempts
            If TrySetClipboardSta(dataObj, verificationText) Then
                Return True
            End If
            Await System.Threading.Tasks.Task.Delay(80 * i).ConfigureAwait(False)
        Next

        Return False
    End Function

    ''' <summary>
    ''' Attempts immediate clipboard set on current STA thread; verifies content if expected text provided.
    ''' </summary>
    Private Function TrySetClipboardImmediate(dataObj As System.Windows.Forms.DataObject, verificationText As String) As Boolean
        Try
            ' Must be STA; if not, return False silently.
            If System.Threading.Thread.CurrentThread.GetApartmentState() <> System.Threading.ApartmentState.STA Then
                Return False
            End If
            System.Windows.Forms.Clipboard.SetDataObject(dataObj, True) ' True = persistent (critical)
            Return VerifyClipboard(verificationText)
        Catch ex As System.Threading.ThreadStateException
            Return False
        Catch ex As System.Runtime.InteropServices.ExternalException
            Return False
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Performs clipboard set on a dedicated STA background thread and verifies result.
    ''' </summary>
    Private Function TrySetClipboardSta(dataObj As System.Windows.Forms.DataObject, verificationText As String) As Boolean
        Dim success As Boolean = False
        Dim t As New System.Threading.Thread(
            Sub()
                Try
                    System.Windows.Forms.Clipboard.SetDataObject(dataObj, True) ' persistent copy
                    success = VerifyClipboard(verificationText)
                Catch
                    success = False
                End Try
            End Sub)
        t.SetApartmentState(System.Threading.ApartmentState.STA)
        t.IsBackground = True
        t.Start()
        t.Join()
        Return success
    End Function

    ''' <summary>
    ''' Verifies clipboard text content. If expected is Nothing, checks for presence of any text. If provided, compares prefix segment.
    ''' </summary>
    Private Function VerifyClipboard(expected As String) As Boolean
        If String.IsNullOrEmpty(expected) Then
            ' Basic sanity: at least one format present
            Return System.Windows.Forms.Clipboard.ContainsText()
        End If
        Try
            Dim got As String = System.Windows.Forms.Clipboard.GetText()
            If String.IsNullOrEmpty(got) Then Return False
            ' Allow minor truncation differences (some managers trim)
            If got.Length >= Math.Min(expected.Length, 32) Then
                ' Compare first 32 chars
                Return String.Compare(got.Substring(0, Math.Min(got.Length, 32)),
                                      expected.Substring(0, Math.Min(expected.Length, 32)),
                                      StringComparison.Ordinal) = 0
            End If
            Return False
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Writes text to a temporary file (Desktop or temp folder) in UTF-8; returns full path or Nothing on failure.
    ''' </summary>
    Private Function SaveTextToTempFile(text As String) As String
        Try
            Dim desktop As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.DesktopDirectory)
            If String.IsNullOrWhiteSpace(desktop) OrElse Not System.IO.Directory.Exists(desktop) Then
                ' Fallback to temp if desktop unavailable
                desktop = System.IO.Path.GetTempPath()
            End If

            Dim rawName As String = AN & "_" & DateTime.UtcNow.ToString("yyyyMMdd_HHmmss") &
                                    "_" & Guid.NewGuid().ToString("N").Substring(0, 6) & ".txt"

            ' Sanitize filename
            Dim invalid = System.IO.Path.GetInvalidFileNameChars()
            Dim safeName = New String(rawName.Select(Function(c) If(invalid.Contains(c), "_"c, c)).ToArray())

            Dim fullPath As String = System.IO.Path.Combine(desktop, safeName)

            System.IO.File.WriteAllText(fullPath, If(text, String.Empty), System.Text.Encoding.UTF8)
            Return fullPath
        Catch
            Return Nothing
        End Try
    End Function


    Private _win As HelpMeInky = Nothing

    ''' <summary>
    ''' Displays (or reuses) the HelpMeInky window for assistance features related to the add-in context.
    ''' </summary>
    Public Sub HelpMeInky()
        If _win Is Nothing OrElse _win.IsDisposed Then
            _win = New HelpMeInky(_context, RDV)
        End If
        ' No owner needed
        _win.ShowRaised()
    End Sub

    ''' <summary>
    ''' Constructs settings dictionaries (labels and tips), opens settings UI, then updates ribbons.
    ''' </summary>
    Public Sub ShowSettings()

        Dim Settings As New Dictionary(Of String, String) From {
                        {"Temperature", "Temperature of {model}"},
                        {"Timeout", "Timeout of {model}"},
                        {"Temperature_2", "Temperature of {model2}"},
                        {"Timeout_2", "Timeout of {model2}"},
                        {"DoubleS", "Convert '" & ChrW(223) & "' to 'ss'"},
                        {"Clean", "Clean the LLM response"},
                        {"NoEmDash", "Convert em to en dash"},
                        {"Ignore", "Activate 'Ignore' prompt (for 'prompt injection' protection)"},
                        {"MarkdownConvert", "Keep character formatting (try this first)"},
                        {"KeepFormat1", "Keep format (translations, additional coding)"},
                        {"ReplaceText1", "Replace text (translations)"},
                        {"KeepFormat2", "Keep format (other commands, additional coding)"},
                        {"ReplaceText2", "Replace text (other commands)"},
                        {"ReplaceText2Override", "Replace text (other commands) [override]"},
                        {"DoMarkupOutlook", "Also do a markup (other commands)"},
                        {"MarkupMethodOutlook", "Markup method (1 = Word, 2 = Diff, 3 = DiffW)"},
                        {"MarkupMethodOutlookOverride", "Markup method (1 = Word, 2 = Diff, 3 = DiffW) [override]"},
                        {"MarkupDiffCap", "Maximum characters for Diff Markup"},
                        {"PreCorrection", "Additional instruction for prompts"},
                        {"PostCorrection", "Prompt to apply after queries"},
                        {"Language1", "Default translation language"},
                        {"PromptLibPath", "Prompt library file"},
                        {"PromptLibPathLocal", "Prompt library file (local)"},
                        {"DefaultPrefix", "Default prefix to use in 'Freestyle'"},
                        {"Location", "Location information to use, e.g., in 'Freestyle'"}
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
                        {"KeepFormat1", "If selected, the original's text basic formatting of a translated text will be retained (by HTML encoding, takes time!)"},
                        {"ReplaceText1", "If selected, the response of the LLM for translations will replace the original text"},
                        {"KeepFormat2", "If selected, the original's text basic formatting of other text (other than translations) will be retained (by HTML encoding, takes time!)"},
                        {"ReplaceText2", "If selected, the response of the LLM for other commands (than translate) will replace the original text"},
                        {"ReplaceText2Override", "Leave empty to not override the above value; use 0 or 'false' to disable and 1 or 'true' to enable 'Replace text' as a personal override"},
                        {"DoMarkupOutlook", "Whether a markup should be done for functions that change only parts of a text"},
                        {"MarkupMethodOutlook", "Markup method to use: 1 = Compare using the Word compare function, 2 = Simple Differ, 3 = Simple Diff shown in a window"},
                        {"MarkupMethodOutlookOverride", "Leave empty to not override the above value; otherwise enter the personal override value for 'markup method'"},
                        {"MarkupDiffCap", "The maximum size of the text that should be processed using the Diff method (to avoid you having to wait too long)"},
                        {"PreCorrection", "Add prompting text that will be added to all basic requests (e.g., for special language tasks)"},
                        {"PostCorrection", "Add a prompt that will be applied to each result before it is further processed (slow!)"},
                        {"Language1", "The language (in English) that will be used for the quick access button in the ribbon"},
                        {"PromptLibPath", "The filename (including path, support environmental variables) for your prompt library (if any)"},
                        {"PromptLibPathLocal", "The filename (including path, support environmental variables) for your local prompt library (if any)"},
                        {"DefaultPrefix", "You can define here the default prefix to use within 'Freestyle' if no other prefix is used (will be added automatically)."},
                        {"Location", "Provide location information (e.g., 'We are in Zurich, Switzerland') to be used in 'Freestyle', chatbot and some other prompts that contain {Location} to get more location specific results."}
                    }

        ShowSettingsWindow(Settings, SettingsTips)

        Globals.Ribbons.Ribbon1.UpdateRibbon()
        Globals.Ribbons.Ribbon2.UpdateRibbon()

    End Sub


End Class