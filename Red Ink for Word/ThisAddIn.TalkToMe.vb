' Part of "Red Ink" (Red Ink for Word)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.TalkToMe.vb
' Purpose: Implements the "Talk to Me" feature for the Word Add-in, handling
'          user interactions and orchestrating transcription and text insertion.
'
' Architecture:
'  - Ribbon Interaction: Responds to the "Talk to Me" ribbon button click to
'    initiate a session.
'  - State Management: Manages the recording state and UI updates for the
'    TalkToMe widget.
'  - Transcription Service: Utilizes the transcription engine provided by the
'    SharedLibrary to convert speech to text.
'  - Text Insertion: Inserts the transcribed text into the active Word document
'    at the current selection.
'  - Error Handling: Provides user feedback in case of transcription errors.
' =============================================================================


Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports NAudio.CoreAudioApi
Imports NAudio.Wave
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.Transcription

Public NotInheritable Class WordTalkToMeHostAdapter
    Implements SharedLibrary.SharedLibrary.ITalkToMeHostAdapter

    Private ReadOnly _owner As ThisAddIn

    Public Sub New(owner As ThisAddIn)
        _owner = owner
    End Sub

    Public ReadOnly Property HostName As String Implements SharedLibrary.SharedLibrary.ITalkToMeHostAdapter.HostName
        Get
            Return "Word"
        End Get
    End Property

    Public Function GetSupportedCommands() As List(Of SharedLibrary.SharedLibrary.TalkToMeCommandDefinition) Implements SharedLibrary.SharedLibrary.ITalkToMeHostAdapter.GetSupportedCommands
        Dim ribbon As Ribbon1 = Globals.Ribbons.Ribbon1

        If ribbon Is Nothing Then
            Return New List(Of SharedLibrary.SharedLibrary.TalkToMeCommandDefinition)()
        End If

        Return ribbon.GetTalkToMeCommandDefinitions()
    End Function

    Public Function GetPromptContext(includeFullDocument As Boolean) As SharedLibrary.SharedLibrary.TalkToMeDocumentContext Implements SharedLibrary.SharedLibrary.ITalkToMeHostAdapter.GetPromptContext
        Dim app As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
        Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
        Dim sel As Microsoft.Office.Interop.Word.Selection = Nothing

        Try
            doc = app.ActiveDocument
        Catch
        End Try

        Try
            sel = app.Selection
        Catch
        End Try

        Dim canWriteToDocument As Boolean = CanWriteToDocumentNow(app, doc, sel)
        Dim canControlDocument As Boolean = CanControlWordDocumentNow(app, doc, sel)
        Dim hasSelection As Boolean =
            canControlDocument AndAlso
            sel IsNot Nothing AndAlso
            sel.Start <> sel.End

        Dim selectionText As String = If(hasSelection, sel.Text, "")
        Dim cursorContext As String = ""

        If canControlDocument AndAlso
           Not hasSelection AndAlso
           doc IsNot Nothing AndAlso
           sel IsNot Nothing Then

            Dim cursorPos As Integer = sel.Start
            Dim contextStart As Integer = Math.Max(doc.Content.Start, cursorPos - 120)
            Dim contextEnd As Integer = Math.Min(doc.Content.End, cursorPos + 120)
            Dim beforeRange As Microsoft.Office.Interop.Word.Range = doc.Range(contextStart, cursorPos)
            Dim afterRange As Microsoft.Office.Interop.Word.Range = doc.Range(cursorPos, contextEnd)
            cursorContext = beforeRange.Text & "[cursor is here]" & afterRange.Text
        End If

        Return New SharedLibrary.SharedLibrary.TalkToMeDocumentContext With {
            .HostName = "Word",
            .DocumentName = If(doc IsNot Nothing, doc.Name, ""),
            .DocumentText = If(includeFullDocument AndAlso doc IsNot Nothing, doc.Content.Text, ""),
            .SelectionText = selectionText,
            .CursorContext = cursorContext,
            .CaretPosition = If(sel IsNot Nothing AndAlso canControlDocument, $"{sel.Start}-{sel.End}", ""),
            .HasSelection = hasSelection,
            .ActiveSurface = GetActiveSurfaceName(app, doc, sel),
            .CanWriteToDocument = canWriteToDocument
        }
    End Function

    Public Async Function ResolveWithLlmAsync(spokenInstruction As String,
                                              context As SharedLibrary.SharedLibrary.TalkToMeDocumentContext,
                                              supportedCommands As List(Of SharedLibrary.SharedLibrary.TalkToMeCommandDefinition),
                                              cancellationToken As System.Threading.CancellationToken) As Task(Of SharedLibrary.SharedLibrary.TalkToMeStructuredResponse) Implements SharedLibrary.SharedLibrary.ITalkToMeHostAdapter.ResolveWithLlmAsync
        Dim commandList As New StringBuilder()

        For Each command As SharedLibrary.SharedLibrary.TalkToMeCommandDefinition In supportedCommands
            commandList.Append("- id=")
            commandList.Append(command.Name)

            If Not String.IsNullOrWhiteSpace(command.Label) Then
                commandList.Append("; label=""")
                commandList.Append(command.Label.Replace("""", "'"c))
                commandList.Append("""")
            End If

            If Not String.IsNullOrWhiteSpace(command.Category) Then
                commandList.Append("; category=""")
                commandList.Append(command.Category.Replace("""", "'"c))
                commandList.Append("""")
            End If

            If command.Aliases IsNot Nothing AndAlso command.Aliases.Count > 0 Then
                commandList.Append("; aliases=""")
                commandList.Append(String.Join(", ", command.Aliases).Replace("""", "'"c))
                commandList.Append("""")
            End If

            If Not String.IsNullOrWhiteSpace(command.Description) Then
                commandList.Append("; description=""")
                commandList.Append(command.Description.Replace("""", "'"c))
                commandList.Append("""")
            End If

            commandList.AppendLine()
        Next

        Dim systemPrompt As String =
            "You are the command resolver for Red Ink's 'Talk to me!' widget inside Microsoft Word. " &
            "The transcript may contain speech-recognition errors, partial speech, or wake-word variations. " &
            "You must decide whether the speaker was plausibly addressing this assistant. " &
            "Only treat the transcript as addressed to the assistant if either (1) a wake-word is likely present despite recognition errors, or (2) the context makes it clear that this is a command or request directed at an assistant. " &
            "If neither is true, return action='none'. " &
            "Return exactly one structured action. " &
            "Allowed action values are: host_command, type_text, insert_text, freestyle, goto_text, find_text, word_command, none. " &
            "The supported host commands are provided with stable command id, visible button label, menu category, button description, and optional aliases. Treat those as authoritative. " &
            "If ActiveSurface is anything other than 'word_document', then the user is typing into some other focused UI control and plain dictated text should usually be returned as type_text with cleaned dictated text in the user's language. " &
            "This includes chatbot surfaces, help surfaces, discuss surfaces, and any other focused add-in form or typing surface. " &
            "When ActiveSurface is not 'word_document', do NOT use freestyle for normal utterances. " &
            "Use host_command only for clear explicit commands to invoke Red Ink itself. " &
            "Use word_command for native Microsoft Word or keyboard actions such as pressing Enter, pressing Escape, saving, moving the caret, and selecting text. " &
            "When ActiveSurface is not 'word_document', still use word_command for clear explicit keyboard intent such as 'press enter', 'enter', 'new line', 'press escape', or similar commands. " &
            "Only return type_text for words like 'enter' or 'escape' when the user is clearly dictating literal text rather than asking for the key action. " &
            "If the user is dictating text and also explicitly wants inline key actions at exact positions, keep action='type_text' and encode those inline actions as markers inside text, for example [[press enter]] or [[press escape]], exactly where the key action should occur. " & "If action='word_command', place one canonical command in instruction and leave hostCommandName, text, and query empty unless truly needed. " &
            "Supported word_command instruction formats are: " &
            "press_enter; press_escape; save_document; read_selection; " &
            "move|up|line|N; move|down|line|N; " &
            "move|up|paragraph|N; move|down|paragraph|N; " &
            "move|up|page|N; move|down|page|N; " &
            "move|next|paragraph|N; move|previous|paragraph|N; " &
            "move|next|sentence|N; move|previous|sentence|N; " &
            "move|start|line|1; move|end|line|1; " &
            "move|start|paragraph|1; move|end|paragraph|1; " &
            "move|start|document|1; move|end|document|1; " &
            "goto|page|N; " &
            "select|current|line|1; select|current|sentence|1; select|current|paragraph|1; select|entire|document|1; " &
            "activate_document|next|1; activate_document|previous|1; activate_document|name|DOCUMENT_NAME. " &
            "Map synonyms such as 'enter', 'new paragraph', or 'press enter' to press_enter. " &
            "Map synonyms such as 'escape', 'press escape', 'press esc', or 'cancel' to press_escape. " &
            "Map 'save', 'save document', or 'word save' to save_document. " &
            "Map requests such as 'read the selected text', 'read selection aloud', 'vorlesen', 'lire la sélection', or equivalent wording in the user's language to read_selection. " &
            "Map 'go up three lines' to move|up|line|3. " &
            "Map 'go down two pages' or 'page down two pages' to move|down|page|2. " &
            "Map 'next paragraph' to move|next|paragraph|1. " &
            "Map 'paragraphs up two' or 'previous two paragraphs' to move|previous|paragraph|2. " &
            "Map 'next sentence' to move|next|sentence|1. " &
            "Map 'previous sentence' to move|previous|sentence|1. " &
            "Map 'start of line' to move|start|line|1. " &
            "Map 'end of document' to move|end|document|1. " &
            "Map 'go to page 5' to goto|page|5. " &
            "Map 'select line' to select|current|line|1. " &
            "Map 'select sentence' to select|current|sentence|1. " &
            "Map 'select paragraph' to select|current|paragraph|1. " &
            "Map 'select entire document' or 'select whole document' to select|entire|document|1. " &
            "Map 'next document' or 'switch to next document' to activate_document|next|1. " &
            "Map 'previous document' to activate_document|previous|1. " &
            "Use freestyle only when Red Ink itself should process document text, and only when CanWriteToDocument is true. " &
            "A text selection is context only; it does NOT by itself mean that the selected text should be changed, unless the user instructs so. " &
            "If the user wants you to redraft or otherwise change the selected text, always do so in the language of the text, unless directed otherwise. If the user wants you to provide a comment, an analysis or response, provide it in the language of the user's instruction." &
            "When the user wants the selected text or other document text to be corrected, revised, translated, shortened, expanded, reformulated, improved, or otherwise transformed in the document, use freestyle and make instruction start exactly with 'Markup: ' followed by the instruction in the user's language. " &
            "When the user wants new text to be drafted for insertion into the document, use freestyle with an instruction in the user's language that does NOT start with 'Clip: '. " &
            "When the user wants analysis, explanation, comments, feedback, an opinion, an answer, brainstorming, or any response that should NOT be inserted into the document, use freestyle and make instruction start exactly with 'Clip: ' followed by the instruction." &
            "Use 'Clip: ' even if text is selected and even if the request is about that selected text. " &
            "Use 'Markup: ' only when the selected or referenced document text should actually be changed in the document. " &
            "Use 'Replace: ' only when the selected or referenced document text should actually be replaced in the document (for example because markup makes no sense, e.g., for a translation). " &
            "Use 'Append: ' only when the selected or referenced document text should not be replaced but text added to it at the end (for example for a summary). " &
            "Examples: " &
            "'translate this selection into German' => action='freestyle', instruction='Replace: translate the selected text into German'; " &
            "'improve this wording' => action='freestyle', instruction='Markup: improve the wording of the selected text'; " &
            "'what does this paragraph mean' => action='freestyle', instruction='Clip: explain what the selected paragraph means'; " &
            "'is this argument convincing' => action='freestyle', instruction='Clip: assess whether the selected argument is convincing'; " &
            "'write a short conclusion about this topic' => action='freestyle', instruction='Append: write a short conclusion about this topic'. " &
            "If the user asks for chat, chatbot, open chat, open chatbot, or the chat window, choose the chat command. " &
            "Use help_me only when the user explicitly asks for help about Red Ink itself, such as help, help me Inky, Red Ink help, or how Red Ink works. " &
            "Return ONLY compact JSON with these exact properties: " &
            "{""action"":"""",""hostCommandName"":"""",""text"":"""",""query"":"""",""instruction"":"""",""reason"":""""}."


        Dim userPrompt As New StringBuilder()
        userPrompt.AppendLine("Host: Word")
        userPrompt.AppendLine("Document: " & context.DocumentName)
        userPrompt.AppendLine("Current position: " & context.CaretPosition)
        userPrompt.AppendLine("Active surface: " & context.ActiveSurface)
        userPrompt.AppendLine("Can write to document: " & context.CanWriteToDocument.ToString())
        userPrompt.AppendLine("HasSelection: " & context.HasSelection.ToString())
        userPrompt.AppendLine()

        userPrompt.AppendLine("Supported commands:")
        userPrompt.AppendLine(commandList.ToString())

        AppendOpenDocumentsToPrompt(userPrompt)

        If context.HasSelection Then
            userPrompt.AppendLine("Selected text:")
            userPrompt.AppendLine(context.SelectionText)
            userPrompt.AppendLine()
        End If

        If Not String.IsNullOrWhiteSpace(context.CursorContext) Then
            userPrompt.AppendLine("Cursor context:")
            userPrompt.AppendLine(context.CursorContext)
            userPrompt.AppendLine()
        End If

        If Not String.IsNullOrWhiteSpace(context.DocumentText) Then
            userPrompt.AppendLine("Full document:")
            userPrompt.AppendLine(context.DocumentText)
            userPrompt.AppendLine()
        End If

        userPrompt.AppendLine("Transcript:")
        userPrompt.AppendLine(spokenInstruction)

        Dim rawResult As String = Await RunTalkToMeLlmAsync(systemPrompt, userPrompt.ToString(), cancellationToken)

        Return ParseStructuredResponse(rawResult)
    End Function

    Private Async Function RunTalkToMeLlmAsync(systemPrompt As String,
                                               userPrompt As String,
                                               cancellationToken As System.Threading.CancellationToken) As Task(Of String)
        Dim originalConfig As ModelConfig = Nothing
        Dim talkToMeConfig As ModelConfig = Nothing
        Dim restoreRequired As Boolean = False

        Try
            If SharedMethods.TryGetSpecialTaskModelConfig(
                _owner._context,
                _owner.INI_AlternateModelPath,
                "TalkToMe",
                talkToMeConfig) Then

                originalConfig = SharedMethods.GetCurrentConfig(_owner._context)

                Dim errorFlag As Boolean = False
                SharedMethods.ApplyModelConfig(_owner._context, talkToMeConfig, errorFlag)

                If errorFlag Then
                    Throw New InvalidOperationException("Failed to apply the TalkToMe special task model.")
                End If

                restoreRequired = True

                Return Await _owner.LLM(
                    systemPrompt,
                    userPrompt,
                    "",
                    "",
                    0,
                    UseSecondAPI:=True,
                    Hidesplash:=True,
                    cancellationToken:=cancellationToken,
                    EnsureUI:=False)
            End If

            Return Await _owner.LLM(
                systemPrompt,
                userPrompt,
                "",
                "",
                0,
                UseSecondAPI:=False,
                Hidesplash:=True,
                cancellationToken:=cancellationToken,
                EnsureUI:=False)
        Finally
            If restoreRequired AndAlso originalConfig IsNot Nothing Then
                SharedMethods.RestoreDefaults(_owner._context, originalConfig)
            End If
        End Try
    End Function

    Public Async Function ExecuteAsync(response As SharedLibrary.SharedLibrary.TalkToMeStructuredResponse,
                                       cancellationToken As System.Threading.CancellationToken) As Task(Of SharedLibrary.SharedLibrary.TalkToMeDispatchResult) Implements SharedLibrary.SharedLibrary.ITalkToMeHostAdapter.ExecuteAsync
        Dim actionName As String = If(response.Action, "").Trim().ToLowerInvariant()

        Dim app As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
        Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
        Dim selection As Microsoft.Office.Interop.Word.Selection = Nothing

        Try
            doc = app.ActiveDocument
        Catch
        End Try

        Try
            selection = app.Selection
        Catch
        End Try

        Dim canWriteToDocument As Boolean = CanWriteToDocumentNow(app, doc, selection)

        If Not canWriteToDocument Then
            Select Case actionName
                Case "freestyle"
                    Dim fallbackText As String = GetLiteralTypingFallbackText(response)

                    If Not String.IsNullOrWhiteSpace(fallbackText) Then
                        Await RunOnUiThreadAsync(
                            Sub()
                                InsertLiteralText(fallbackText)
                            End Sub)

                        Return New SharedLibrary.SharedLibrary.TalkToMeDispatchResult With {
                            .Handled = True,
                            .StatusText = "Text inserted.",
                            .TranscriptToDisplay = fallbackText
                        }
                    End If

                    Return New SharedLibrary.SharedLibrary.TalkToMeDispatchResult With {
                        .Handled = False,
                        .StatusText = "Freestyle blocked outside writable document surface.",
                        .TranscriptToDisplay = response.Instruction
                    }

                Case "insert_text"
                    actionName = "type_text"
            End Select
        End If

        Select Case actionName
            Case "host_command"
                Await RunOnUiThreadAsync(
                    Sub()
                        ExecuteHostCommand(response.HostCommandName)
                    End Sub)

                Return New SharedLibrary.SharedLibrary.TalkToMeDispatchResult With {
                    .Handled = True,
                    .StatusText = "Command executed.",
                    .TranscriptToDisplay = response.Instruction
                }

            Case "type_text", "insert_text"
                Await RunOnUiThreadAsync(
                    Sub()
                        InsertLiteralText(response.Text)
                    End Sub)

                Return New SharedLibrary.SharedLibrary.TalkToMeDispatchResult With {
                    .Handled = True,
                    .StatusText = "Text inserted.",
                    .TranscriptToDisplay = response.Text
                }

            Case "goto_text", "find_text"
                Dim succeeded As Boolean =
                    Await RunOnUiThreadAsync(
                        Function()
                            Return FindAndSelect(response.Query)
                        End Function)

                Return New SharedLibrary.SharedLibrary.TalkToMeDispatchResult With {
                    .Handled = succeeded,
                    .StatusText = If(succeeded, "Text found.", "Text not found."),
                    .TranscriptToDisplay = response.Query
                }

            Case "word_command"
                Dim statusText As String =
                    Await RunOnUiThreadAsync(
                        Function()
                            Return ExecuteWordCommand(response.Instruction)
                        End Function)

                Return New SharedLibrary.SharedLibrary.TalkToMeDispatchResult With {
                    .Handled = True,
                    .StatusText = statusText,
                    .TranscriptToDisplay = response.Instruction
                }

            Case "freestyle"
                Await RunOnUiThreadAsync(
                    Sub()
                        ExecuteFreestyleInstruction(response.Instruction)
                    End Sub)

                Return New SharedLibrary.SharedLibrary.TalkToMeDispatchResult With {
                    .Handled = True,
                    .StatusText = "Freestyle executed.",
                    .TranscriptToDisplay = response.Instruction
                }

            Case Else
                Return New SharedLibrary.SharedLibrary.TalkToMeDispatchResult With {
                    .Handled = False,
                    .StatusText = "No actionable command.",
                    .TranscriptToDisplay = response.Instruction
                }
        End Select
    End Function

    Private Shared Function GetLiteralTypingFallbackText(response As SharedLibrary.SharedLibrary.TalkToMeStructuredResponse) As String
        If response Is Nothing Then
            Return ""
        End If

        Dim candidate As String = If(response.Text, "").Trim()

        If String.IsNullOrWhiteSpace(candidate) Then
            candidate = If(response.Instruction, "").Trim()
        End If

        If candidate.StartsWith("Clip:", StringComparison.OrdinalIgnoreCase) Then
            candidate = candidate.Substring(5).Trim()
        End If

        If candidate.StartsWith("Markup:", StringComparison.OrdinalIgnoreCase) Then
            candidate = candidate.Substring(7).Trim()
        End If

        Return candidate
    End Function


    Private Shared Function ParseStructuredResponse(rawResult As String) As SharedLibrary.SharedLibrary.TalkToMeStructuredResponse
        Dim raw As String = If(rawResult, "").Trim()
        Dim startIndex As Integer = raw.IndexOf("{"c)
        Dim endIndex As Integer = raw.LastIndexOf("}"c)

        If startIndex < 0 OrElse endIndex <= startIndex Then
            Return New SharedLibrary.SharedLibrary.TalkToMeStructuredResponse With {
                .Action = "none",
                .Reason = "No JSON found."
            }
        End If

        Dim jsonText As String = raw.Substring(startIndex, endIndex - startIndex + 1)

        Try
            Dim jobj As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(jsonText)

            Return New SharedLibrary.SharedLibrary.TalkToMeStructuredResponse With {
                .Action = If(CStr(jobj("action")), ""),
                .HostCommandName = If(CStr(jobj("hostCommandName")), ""),
                .Text = If(CStr(jobj("text")), ""),
                .Query = If(CStr(jobj("query")), ""),
                .Instruction = If(CStr(jobj("instruction")), ""),
                .Reason = If(CStr(jobj("reason")), "")
            }
        Catch
            Return New SharedLibrary.SharedLibrary.TalkToMeStructuredResponse With {
                .Action = "none",
                .Reason = "JSON parsing failed."
            }
        End Try
    End Function

    Private Sub ExecuteHostCommand(commandName As String)
        Dim ribbon As Ribbon1 = Globals.Ribbons.Ribbon1

        If ribbon IsNot Nothing AndAlso ribbon.TryExecuteTalkToMeCommand(commandName) Then
            Return
        End If

        Throw New InvalidOperationException("Unknown TalkToMe host command: " & If(commandName, ""))
    End Sub

    Private Sub ExecuteFreestyleInstruction(instruction As String)
        Dim app As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
        Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
        Dim selection As Microsoft.Office.Interop.Word.Selection = Nothing

        Try
            doc = app.ActiveDocument
        Catch
        End Try

        Try
            selection = app.Selection
        Catch
        End Try

        If Not CanWriteToDocumentNow(app, doc, selection) Then
            Throw New InvalidOperationException("Freestyle is not allowed outside a writable Word document surface.")
        End If

        Dim prompt As String = If(instruction, "").Trim()

        If String.IsNullOrWhiteSpace(prompt) Then
            Throw New InvalidOperationException("The freestyle instruction is empty.")
        End If

        Globals.ThisAddIn.FreeStyle(False, prompt)
    End Sub


    Private Sub InsertLiteralText(textToInsert As String)
        Dim app As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
        Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
        Dim selection As Microsoft.Office.Interop.Word.Selection = Nothing

        Try
            doc = app.ActiveDocument
        Catch
        End Try

        Try
            selection = app.Selection
        Catch
        End Try

        If CanWriteToDocumentNow(app, doc, selection) Then
            Dim originalTrackRevisions As Boolean = doc.TrackRevisions

            Try
                doc.TrackRevisions = True

                If selection.Start <> selection.End Then
                    selection.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                End If

                TypeTextIntoWordSelection(selection, textToInsert)
            Finally
                doc.TrackRevisions = originalTrackRevisions
            End Try

            Return
        End If

        TypeTextIntoActiveUiTarget(textToInsert)
    End Sub

    Private Function FindAndSelect(query As String) As Boolean
        Dim searchText As String = If(query, "").Trim()

        If String.IsNullOrWhiteSpace(searchText) Then
            Return False
        End If

        Dim app As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
        Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
        Dim selection As Microsoft.Office.Interop.Word.Selection = Nothing

        Try
            doc = app.ActiveDocument
        Catch
        End Try

        Try
            selection = app.Selection
        Catch
        End Try

        If Not CanControlWordDocumentNow(app, doc, selection) Then
            Return False
        End If

        Try
            app.Activate()
        Catch
        End Try

        Try
            doc.Activate()
        Catch
        End Try

        selection.SetRange(doc.Content.Start, doc.Content.End)

        If Globals.ThisAddIn.FindLongTextInChunks(searchText, selection, True) Then
            FocusWordRange(app, doc, selection.Range)
            Return True
        End If

        Return False
    End Function

    Private Function ExecuteWordCommand(commandText As String) As String
        Dim app As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
        Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
        Dim selection As Microsoft.Office.Interop.Word.Selection = Nothing

        Try
            doc = app.ActiveDocument
        Catch
        End Try

        Try
            selection = app.Selection
        Catch
        End Try

        Dim normalizedCommand As String = If(commandText, "").Trim()

        If String.IsNullOrWhiteSpace(normalizedCommand) Then
            Throw New InvalidOperationException("The Word command is empty.")
        End If

        Dim parts() As String = normalizedCommand.Split("|"c)
        Dim verb As String = GetWordCommandPart(parts, 0).ToLowerInvariant()

        Select Case verb
            Case "press_enter"
                Return ExecutePressEnterCommand(app, doc, selection)

            Case "press_escape"
                Return ExecutePressEscapeCommand()

            Case "save_document"
                Return ExecuteSaveDocumentCommand(doc)

            Case "read_selection"
                Return ExecuteReadSelectionCommand(selection)

            Case "move"
                If Not CanControlWordDocumentNow(app, doc, selection) Then
                    Throw New InvalidOperationException("Move commands require an active Word document.")
                End If

                Dim direction As String = NormalizeWordCommandDirection(GetWordCommandPart(parts, 1))
                Dim unitName As String = NormalizeWordCommandUnit(GetWordCommandPart(parts, 2))
                Dim count As Integer = ParsePositiveWordCommandCount(GetWordCommandPart(parts, 3))

                Return ExecuteMoveCommand(app, doc, selection, direction, unitName, count)

            Case "goto"
                If Not CanControlWordDocumentNow(app, doc, selection) Then
                    Throw New InvalidOperationException("Go-to commands require an active Word document.")
                End If

                Dim targetName As String = NormalizeWordCommandUnit(GetWordCommandPart(parts, 1))
                Dim count As Integer = ParsePositiveWordCommandCount(GetWordCommandPart(parts, 2))

                If targetName <> "page" Then
                    Throw New InvalidOperationException("Unsupported go-to target: " & targetName)
                End If

                Return ExecuteGotoPageCommand(app, doc, selection, count)

            Case "select"
                If Not CanControlWordDocumentNow(app, doc, selection) Then
                    Throw New InvalidOperationException("Selection commands require an active Word document.")
                End If

                Dim scopeName As String = NormalizeWordCommandScope(GetWordCommandPart(parts, 1))
                Dim unitName As String = NormalizeWordCommandUnit(GetWordCommandPart(parts, 2))

                Return ExecuteSelectCommand(app, doc, selection, scopeName, unitName)

            Case "activate_document"
                Dim mode As String = NormalizeDocumentActivationMode(GetWordCommandPart(parts, 1))
                Dim argument As String = GetWordCommandTail(parts, 2)

                Return ExecuteActivateDocumentCommand(app, mode, argument)

            Case Else
                Throw New InvalidOperationException("Unknown Word command: " & normalizedCommand)
        End Select
    End Function

    Private Shared Function GetWordCommandPart(parts() As String, index As Integer) As String
        If parts Is Nothing OrElse index < 0 OrElse index >= parts.Length Then
            Return ""
        End If

        Return If(parts(index), "").Trim()
    End Function

    Private Shared Function GetWordCommandTail(parts() As String, startIndex As Integer) As String
        If parts Is Nothing OrElse startIndex < 0 OrElse startIndex >= parts.Length Then
            Return ""
        End If

        Return String.Join("|", parts.Skip(startIndex)).Trim()
    End Function

    Private Shared Function NormalizeWordCommandDirection(value As String) As String
        Select Case If(value, "").Trim().ToLowerInvariant()
            Case "up"
                Return "up"

            Case "down"
                Return "down"

            Case "next", "forward"
                Return "next"

            Case "previous", "prev", "back", "backward"
                Return "previous"

            Case "start", "beginning", "home", "top"
                Return "start"

            Case "end", "finish", "bottom"
                Return "end"

            Case Else
                Throw New InvalidOperationException("Unsupported Word command direction: " & If(value, ""))
        End Select
    End Function

    Private Shared Function NormalizeWordCommandUnit(value As String) As String
        Select Case If(value, "").Trim().ToLowerInvariant()
            Case "line", "lines"
                Return "line"

            Case "sentence", "sentences"
                Return "sentence"

            Case "paragraph", "paragraphs", "para", "paras"
                Return "paragraph"

            Case "page", "pages"
                Return "page"

            Case "document", "documents", "doc"
                Return "document"

            Case Else
                Throw New InvalidOperationException("Unsupported Word command unit: " & If(value, ""))
        End Select
    End Function

    Private Shared Function NormalizeWordCommandScope(value As String) As String
        Select Case If(value, "").Trim().ToLowerInvariant()
            Case "current", "this"
                Return "current"

            Case "entire", "whole", "all"
                Return "entire"

            Case Else
                Throw New InvalidOperationException("Unsupported Word command scope: " & If(value, ""))
        End Select
    End Function

    Private Shared Function NormalizeDocumentActivationMode(value As String) As String
        Select Case If(value, "").Trim().ToLowerInvariant()
            Case "next"
                Return "next"

            Case "previous", "prev"
                Return "previous"

            Case "name", "named", "document"
                Return "name"

            Case Else
                Throw New InvalidOperationException("Unsupported document activation mode: " & If(value, ""))
        End Select
    End Function

    Private Shared Function ParsePositiveWordCommandCount(value As String) As Integer
        Dim parsed As Integer

        If Integer.TryParse(If(value, "").Trim(), parsed) AndAlso parsed > 0 Then
            Return parsed
        End If

        Return 1
    End Function

    Private Shared Function ExecutePressEnterCommand(app As Microsoft.Office.Interop.Word.Application,
                                                     doc As Microsoft.Office.Interop.Word.Document,
                                                     selection As Microsoft.Office.Interop.Word.Selection) As String
        If CanWriteToDocumentNow(app, doc, selection) Then
            If selection Is Nothing Then
                Throw New InvalidOperationException("No active Word selection is available.")
            End If

            If selection.Start <> selection.End Then
                selection.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
            End If

            selection.TypeParagraph()
            Return "Paragraph inserted."
        End If

        SendKeys.SendWait("{ENTER}")
        Return "Enter pressed."
    End Function

    Private Shared Function ExecutePressEscapeCommand() As String
        SendKeys.SendWait("{ESC}")
        Return "Escape pressed."
    End Function

    Private Shared Function ExecuteSaveDocumentCommand(doc As Microsoft.Office.Interop.Word.Document) As String
        If doc Is Nothing Then
            Throw New InvalidOperationException("No active Word document is available.")
        End If

        doc.Save()
        Return "Document saved."
    End Function

    Private Shared Function ExecuteReadSelectionCommand(selection As Microsoft.Office.Interop.Word.Selection) As String
        If selection Is Nothing OrElse selection.Start = selection.End Then
            Throw New InvalidOperationException("No text is selected.")
        End If

        Dim selectedText As String = If(selection.Text, "").Trim()
        If String.IsNullOrWhiteSpace(selectedText) Then
            Throw New InvalidOperationException("The selected text is empty.")
        End If

        Globals.ThisAddIn.SubmitTalkToMeExternalSpeech("", selectedText)
        Return "Selected text sent to speech output."
    End Function

    Private Shared Function ExecuteMoveCommand(app As Microsoft.Office.Interop.Word.Application,
                                               doc As Microsoft.Office.Interop.Word.Document,
                                               selection As Microsoft.Office.Interop.Word.Selection,
                                               direction As String,
                                               unitName As String,
                                               count As Integer) As String
        If selection Is Nothing Then
            Throw New InvalidOperationException("No active Word selection is available.")
        End If

        Select Case unitName
            Case "line"
                Select Case direction
                    Case "start"
                        selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdLine, Microsoft.Office.Interop.Word.WdMovementType.wdMove)
                        FocusCurrentSelection(app)
                        Return "Moved to the start of the line."

                    Case "end"
                        selection.EndKey(Microsoft.Office.Interop.Word.WdUnits.wdLine, Microsoft.Office.Interop.Word.WdMovementType.wdMove)
                        FocusCurrentSelection(app)
                        Return "Moved to the end of the line."

                    Case "up"
                        CollapseSelectionForMovement(selection, direction)
                        selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, count, Microsoft.Office.Interop.Word.WdMovementType.wdMove)
                        FocusCurrentSelection(app)
                        Return "Moved up " & count.ToString() & " line(s)."

                    Case "down"
                        CollapseSelectionForMovement(selection, direction)
                        selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, count, Microsoft.Office.Interop.Word.WdMovementType.wdMove)
                        FocusCurrentSelection(app)
                        Return "Moved down " & count.ToString() & " line(s)."
                End Select

            Case "sentence"
                Select Case direction
                    Case "next"
                        CollapseSelectionForMovement(selection, direction)
                        selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdSentence, count, Microsoft.Office.Interop.Word.WdMovementType.wdMove)
                        FocusCurrentSelection(app)
                        Return "Moved to the next sentence."

                    Case "previous"
                        CollapseSelectionForMovement(selection, direction)
                        selection.MoveLeft(Microsoft.Office.Interop.Word.WdUnits.wdSentence, count, Microsoft.Office.Interop.Word.WdMovementType.wdMove)
                        FocusCurrentSelection(app)
                        Return "Moved to the previous sentence."
                End Select

            Case "paragraph"
                Select Case direction
                    Case "start"
                        Dim startStatus As String = MoveToParagraphBoundary(doc, selection, True)
                        FocusCurrentSelection(app)
                        Return startStatus

                    Case "end"
                        Dim endStatus As String = MoveToParagraphBoundary(doc, selection, False)
                        FocusCurrentSelection(app)
                        Return endStatus

                    Case "up", "previous"
                        CollapseSelectionForMovement(selection, direction)
                        selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdParagraph, count, Microsoft.Office.Interop.Word.WdMovementType.wdMove)
                        FocusCurrentSelection(app)
                        If direction = "previous" Then
                            Return "Moved to the previous paragraph."
                        End If

                        Return "Moved up " & count.ToString() & " paragraph(s)."

                    Case "down", "next"
                        CollapseSelectionForMovement(selection, direction)
                        selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdParagraph, count, Microsoft.Office.Interop.Word.WdMovementType.wdMove)
                        FocusCurrentSelection(app)
                        If direction = "next" Then
                            Return "Moved to the next paragraph."
                        End If

                        Return "Moved down " & count.ToString() & " paragraph(s)."
                End Select

            Case "page"
                If direction <> "up" AndAlso direction <> "down" Then
                    Throw New InvalidOperationException("Pages support only up or down movement.")
                End If

                CollapseSelectionForMovement(selection, direction)
                Return ExecutePageMoveCommand(app, doc, selection, direction, count)

            Case "document"
                Select Case direction
                    Case "start"
                        Dim startStatus As String = MoveToDocumentBoundary(doc, selection, True)
                        FocusCurrentSelection(app)
                        Return startStatus

                    Case "end"
                        Dim endStatus As String = MoveToDocumentBoundary(doc, selection, False)
                        FocusCurrentSelection(app)
                        Return endStatus
                End Select
        End Select

        Throw New InvalidOperationException(
            "Unsupported move command: " &
            direction &
            " / " &
            unitName)
    End Function

    Private Shared Function ExecuteGotoPageCommand(app As Microsoft.Office.Interop.Word.Application,
                                                   doc As Microsoft.Office.Interop.Word.Document,
                                                   selection As Microsoft.Office.Interop.Word.Selection,
                                                   pageNumber As Integer) As String
        If app Is Nothing OrElse doc Is Nothing OrElse selection Is Nothing Then
            Throw New InvalidOperationException("Page navigation requires an active Word document.")
        End If

        Dim targetRange As Microsoft.Office.Interop.Word.Range =
            doc.GoTo(
                What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToPage,
                Which:=Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToAbsolute,
                Count:=pageNumber)

        If targetRange Is Nothing Then
            Throw New InvalidOperationException("Unable to navigate to the requested page.")
        End If

        selection.SetRange(targetRange.Start, targetRange.Start)
        FocusWordRange(app, doc, selection.Range)

        Return "Moved to page " & pageNumber.ToString() & "."
    End Function

    Private Shared Function ExecutePageMoveCommand(app As Microsoft.Office.Interop.Word.Application,
                                                   doc As Microsoft.Office.Interop.Word.Document,
                                                   selection As Microsoft.Office.Interop.Word.Selection,
                                                   direction As String,
                                                   count As Integer) As String
        If doc Is Nothing OrElse selection Is Nothing Then
            Throw New InvalidOperationException("Page navigation requires an active Word document.")
        End If

        Dim currentPageNumber As Integer =
            CInt(selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndAdjustedPageNumber))

        Dim targetPageNumber As Integer

        If direction = "up" Then
            targetPageNumber = Math.Max(1, currentPageNumber - count)
        Else
            targetPageNumber = Math.Max(1, currentPageNumber + count)
        End If

        Dim targetRange As Microsoft.Office.Interop.Word.Range =
            doc.GoTo(
                What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToPage,
                Which:=Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToAbsolute,
                Count:=targetPageNumber)

        If targetRange Is Nothing Then
            Throw New InvalidOperationException("Unable to navigate to the requested page.")
        End If

        selection.SetRange(targetRange.Start, targetRange.Start)
        FocusWordRange(app, doc, selection.Range)

        If direction = "up" Then
            Return "Moved up " & count.ToString() & " page(s)."
        End If

        Return "Moved down " & count.ToString() & " page(s)."
    End Function

    Private Shared Function MoveToDocumentBoundary(doc As Microsoft.Office.Interop.Word.Document,
                                                   selection As Microsoft.Office.Interop.Word.Selection,
                                                   moveToStart As Boolean) As String
        If doc Is Nothing OrElse selection Is Nothing Then
            Throw New InvalidOperationException("Document navigation requires an active Word document.")
        End If

        Dim targetPosition As Integer

        If moveToStart Then
            targetPosition = doc.Content.Start
            selection.SetRange(targetPosition, targetPosition)
            Return "Moved to the start of the document."
        End If

        targetPosition = Math.Max(doc.Content.Start, doc.Content.End - 1)
        selection.SetRange(targetPosition, targetPosition)
        Return "Moved to the end of the document."
    End Function

    Private Shared Function MoveToParagraphBoundary(doc As Microsoft.Office.Interop.Word.Document,
                                                    selection As Microsoft.Office.Interop.Word.Selection,
                                                    moveToStart As Boolean) As String
        If doc Is Nothing OrElse selection Is Nothing Then
            Throw New InvalidOperationException("Paragraph navigation requires an active Word document.")
        End If

        Dim caretPosition As Integer = selection.Start
        Dim caretRange As Microsoft.Office.Interop.Word.Range = doc.Range(caretPosition, caretPosition)

        If caretRange Is Nothing OrElse caretRange.Paragraphs Is Nothing OrElse caretRange.Paragraphs.Count = 0 Then
            Throw New InvalidOperationException("Unable to determine the current paragraph.")
        End If

        Dim paragraphRange As Microsoft.Office.Interop.Word.Range = caretRange.Paragraphs(1).Range
        Dim targetPosition As Integer

        If moveToStart Then
            targetPosition = paragraphRange.Start
            selection.SetRange(targetPosition, targetPosition)
            Return "Moved to the start of the paragraph."
        End If

        targetPosition = Math.Max(paragraphRange.Start, paragraphRange.End - 1)
        selection.SetRange(targetPosition, targetPosition)
        Return "Moved to the end of the paragraph."
    End Function

    Private Shared Function ExecuteSelectCommand(app As Microsoft.Office.Interop.Word.Application,
                                                 doc As Microsoft.Office.Interop.Word.Document,
                                                 selection As Microsoft.Office.Interop.Word.Selection,
                                                 scopeName As String,
                                                 unitName As String) As String
        If doc Is Nothing OrElse selection Is Nothing Then
            Throw New InvalidOperationException("Selection commands require an active Word document.")
        End If

        Select Case scopeName
            Case "entire"
                If unitName <> "document" Then
                    Throw New InvalidOperationException("Only the entire document can be selected with scope 'entire'.")
                End If

                doc.Content.Select()
                FocusCurrentSelection(app)
                Return "Entire document selected."

            Case "current"
                Select Case unitName
                    Case "line"
                        Dim anchorPosition As Integer = selection.Start
                        selection.SetRange(anchorPosition, anchorPosition)
                        selection.HomeKey(Microsoft.Office.Interop.Word.WdUnits.wdLine, Microsoft.Office.Interop.Word.WdMovementType.wdMove)
                        selection.EndKey(Microsoft.Office.Interop.Word.WdUnits.wdLine, Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        FocusCurrentSelection(app)
                        Return "Current line selected."

                    Case "sentence"
                        Dim sentenceRange As Microsoft.Office.Interop.Word.Range = doc.Range(selection.Start, selection.Start)

                        If sentenceRange Is Nothing OrElse sentenceRange.Sentences Is Nothing OrElse sentenceRange.Sentences.Count = 0 Then
                            Throw New InvalidOperationException("Unable to determine the current sentence.")
                        End If

                        sentenceRange.Sentences(1).Select()
                        FocusCurrentSelection(app)
                        Return "Current sentence selected."

                    Case "paragraph"
                        Dim paragraphRange As Microsoft.Office.Interop.Word.Range = doc.Range(selection.Start, selection.Start)

                        If paragraphRange Is Nothing OrElse paragraphRange.Paragraphs Is Nothing OrElse paragraphRange.Paragraphs.Count = 0 Then
                            Throw New InvalidOperationException("Unable to determine the current paragraph.")
                        End If

                        paragraphRange.Paragraphs(1).Range.Select()
                        FocusCurrentSelection(app)
                        Return "Current paragraph selected."

                    Case "document"
                        doc.Content.Select()
                        FocusCurrentSelection(app)
                        Return "Current document selected."
                End Select
        End Select

        Throw New InvalidOperationException(
            "Unsupported select command: " &
            scopeName &
            " / " &
            unitName)
    End Function

    Private Shared Function ExecuteActivateDocumentCommand(app As Microsoft.Office.Interop.Word.Application,
                                                           mode As String,
                                                           argument As String) As String
        If app Is Nothing OrElse app.Documents Is Nothing OrElse app.Documents.Count = 0 Then
            Throw New InvalidOperationException("No open Word documents are available.")
        End If

        Dim targetDoc As Microsoft.Office.Interop.Word.Document = Nothing

        Select Case mode
            Case "next"
                targetDoc = GetAdjacentOpenDocument(app, True)

            Case "previous"
                targetDoc = GetAdjacentOpenDocument(app, False)

            Case "name"
                targetDoc = FindBestMatchingOpenDocument(app, argument)

                If targetDoc Is Nothing Then
                    Throw New InvalidOperationException("No open document matches: " & If(argument, ""))
                End If

            Case Else
                Throw New InvalidOperationException("Unsupported document activation mode: " & mode)
        End Select

        targetDoc.Activate()

        Try
            app.Activate()
        Catch
        End Try

        FocusCurrentSelection(app)

        Return "Switched to document: " & targetDoc.Name
    End Function

    Private Shared Function GetAdjacentOpenDocument(app As Microsoft.Office.Interop.Word.Application,
                                                    moveNext As Boolean) As Microsoft.Office.Interop.Word.Document
        Dim activeDoc As Microsoft.Office.Interop.Word.Document = Nothing

        Try
            activeDoc = app.ActiveDocument
        Catch
        End Try

        If activeDoc Is Nothing Then
            Return app.Documents(1)
        End If

        Dim activeIndex As Integer = 1

        For i As Integer = 1 To app.Documents.Count
            If app.Documents(i) Is activeDoc Then
                activeIndex = i
                Exit For
            End If
        Next

        Dim targetIndex As Integer

        If moveNext Then
            targetIndex = activeIndex + 1
            If targetIndex > app.Documents.Count Then
                targetIndex = 1
            End If
        Else
            targetIndex = activeIndex - 1
            If targetIndex < 1 Then
                targetIndex = app.Documents.Count
            End If
        End If

        Return app.Documents(targetIndex)
    End Function

    Private Shared Function FindBestMatchingOpenDocument(app As Microsoft.Office.Interop.Word.Application,
                                                         selector As String) As Microsoft.Office.Interop.Word.Document
        Dim normalizedSelector As String = NormalizeDocumentSelector(selector)

        If String.IsNullOrWhiteSpace(normalizedSelector) Then
            Return Nothing
        End If

        Dim selectorTokens As List(Of String) = TokenizeDocumentSelector(normalizedSelector)
        Dim bestDoc As Microsoft.Office.Interop.Word.Document = Nothing
        Dim bestScore As Integer = Integer.MinValue

        For i As Integer = 1 To app.Documents.Count
            Dim currentDoc As Microsoft.Office.Interop.Word.Document = app.Documents(i)
            Dim currentScore As Integer = ComputeOpenDocumentMatchScore(currentDoc, normalizedSelector, selectorTokens)

            If currentScore > bestScore Then
                bestScore = currentScore
                bestDoc = currentDoc
            End If
        Next

        If bestScore <= 0 Then
            Return Nothing
        End If

        Return bestDoc
    End Function

    Private Shared Function ComputeOpenDocumentMatchScore(doc As Microsoft.Office.Interop.Word.Document,
                                                          normalizedSelector As String,
                                                          selectorTokens As List(Of String)) As Integer
        If doc Is Nothing Then
            Return Integer.MinValue
        End If

        Dim docName As String = If(doc.Name, "").Trim()
        Dim docNameWithoutExtension As String = Path.GetFileNameWithoutExtension(docName)
        Dim fullName As String = ""

        Try
            fullName = If(doc.FullName, "").Trim()
        Catch
        End Try

        Dim normalizedName As String = docName.ToLowerInvariant()
        Dim normalizedNameWithoutExtension As String = docNameWithoutExtension.ToLowerInvariant()
        Dim normalizedFullName As String = fullName.ToLowerInvariant()

        If normalizedName = normalizedSelector OrElse
           normalizedNameWithoutExtension = normalizedSelector OrElse
           normalizedFullName = normalizedSelector Then
            Return 1000
        End If

        Dim score As Integer = 0

        If normalizedNameWithoutExtension.Contains(normalizedSelector) Then
            score = Math.Max(score, 700)
        End If

        If normalizedName.Contains(normalizedSelector) Then
            score = Math.Max(score, 650)
        End If

        If normalizedFullName.Contains(normalizedSelector) Then
            score = Math.Max(score, 600)
        End If

        For Each token As String In selectorTokens
            If token.Length < 2 Then
                Continue For
            End If

            If normalizedNameWithoutExtension.Contains(token) Then
                score += 40
            ElseIf normalizedName.Contains(token) Then
                score += 30
            ElseIf normalizedFullName.Contains(token) Then
                score += 20
            End If
        Next

        Return score
    End Function

    Private Shared Function NormalizeDocumentSelector(value As String) As String
        Dim normalized As String = If(value, "").Trim().ToLowerInvariant()

        If normalized.StartsWith("""", StringComparison.Ordinal) AndAlso
           normalized.EndsWith("""", StringComparison.Ordinal) AndAlso
           normalized.Length >= 2 Then
            normalized = normalized.Substring(1, normalized.Length - 2).Trim()
        End If

        Return normalized
    End Function

    Private Shared Function TokenizeDocumentSelector(value As String) As List(Of String)
        Dim tokenBuilder As New StringBuilder()
        Dim result As New List(Of String)()

        For Each ch As Char In If(value, "")
            If Char.IsLetterOrDigit(ch) Then
                tokenBuilder.Append(ch)
            ElseIf tokenBuilder.Length > 0 Then
                result.Add(tokenBuilder.ToString())
                tokenBuilder.Clear()
            End If
        Next

        If tokenBuilder.Length > 0 Then
            result.Add(tokenBuilder.ToString())
        End If

        Return result.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
    End Function

    Private Shared Sub AppendOpenDocumentsToPrompt(userPrompt As StringBuilder)
        If userPrompt Is Nothing Then
            Return
        End If

        Dim app As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application

        If app Is Nothing OrElse app.Documents Is Nothing OrElse app.Documents.Count = 0 Then
            Return
        End If

        Dim activeDoc As Microsoft.Office.Interop.Word.Document = Nothing

        Try
            activeDoc = app.ActiveDocument
        Catch
        End Try

        userPrompt.AppendLine("Open documents:")

        For i As Integer = 1 To app.Documents.Count
            Dim currentDoc As Microsoft.Office.Interop.Word.Document = app.Documents(i)
            Dim prefix As String = If(currentDoc Is activeDoc, "* ", "- ")
            userPrompt.AppendLine(prefix & currentDoc.Name)
        Next

        userPrompt.AppendLine()
    End Sub

    Private Shared Sub FocusCurrentSelection(app As Microsoft.Office.Interop.Word.Application)
        If app Is Nothing Then
            Return
        End If

        Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
        Dim selection As Microsoft.Office.Interop.Word.Selection = Nothing

        Try
            doc = app.ActiveDocument
        Catch
        End Try

        Try
            selection = app.Selection
        Catch
        End Try

        If doc Is Nothing OrElse selection Is Nothing OrElse selection.Range Is Nothing Then
            Return
        End If

        FocusWordRange(app, doc, selection.Range)
    End Sub

    Private Shared Sub FocusWordRange(app As Microsoft.Office.Interop.Word.Application,
                                      doc As Microsoft.Office.Interop.Word.Document,
                                      targetRange As Microsoft.Office.Interop.Word.Range)
        If app Is Nothing OrElse doc Is Nothing OrElse targetRange Is Nothing Then
            Return
        End If

        Try
            app.Activate()
        Catch
        End Try

        Try
            doc.Activate()
        Catch
        End Try

        Try
            targetRange.Select()
        Catch
        End Try

        Try
            Dim scrollTarget As Object = targetRange
            app.ActiveWindow.ScrollIntoView(scrollTarget, True)
        Catch
        End Try
    End Sub

    Private Shared Sub CollapseSelectionForMovement(selection As Microsoft.Office.Interop.Word.Selection,
                                                    direction As String)
        If selection Is Nothing OrElse selection.Start = selection.End Then
            Return
        End If

        Select Case If(direction, "")
            Case "down", "next", "end"
                selection.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)

            Case Else
                selection.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseStart)
        End Select
    End Sub

    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=False)>
    Private Shared Function GetWindowThreadProcessId(hWnd As IntPtr,
                                                 ByRef lpdwProcessId As UInteger) As UInteger
    End Function

    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function GetForegroundWindow() As IntPtr
    End Function

    Private Shared Function TryGetWordMainWindowHandle(app As Microsoft.Office.Interop.Word.Application) As IntPtr
        Try
            If app Is Nothing Then
                Return IntPtr.Zero
            End If

            Dim rawHwnd As Object = app.Hwnd

            If rawHwnd Is Nothing Then
                Return IntPtr.Zero
            End If

            Return New IntPtr(CInt(rawHwnd))
        Catch
            Return IntPtr.Zero
        End Try
    End Function

    Private Shared Function IsWordMainWindowForeground(app As Microsoft.Office.Interop.Word.Application) As Boolean
        Try
            If app Is Nothing Then
                Return False
            End If

            Dim foregroundWindow As IntPtr = GetForegroundWindow()
            If foregroundWindow = IntPtr.Zero Then
                Return False
            End If

            Dim foregroundProcessId As UInteger = 0UI
            GetWindowThreadProcessId(foregroundWindow, foregroundProcessId)

            If foregroundProcessId = 0UI Then
                Return False
            End If

            Dim currentProcessId As UInteger = CUInt(System.Diagnostics.Process.GetCurrentProcess().Id)
            If foregroundProcessId <> currentProcessId Then
                Return False
            End If

            Dim activeWindow As Microsoft.Office.Interop.Word.Window = Nothing

            Try
                activeWindow = app.ActiveWindow
            Catch
            End Try

            Return activeWindow IsNot Nothing
        Catch
            Return False
        End Try
    End Function



    Private Shared Function GetActiveSurfaceName(app As Microsoft.Office.Interop.Word.Application,
                                                 doc As Microsoft.Office.Interop.Word.Document,
                                                 sel As Microsoft.Office.Interop.Word.Selection) As String
        Try
            Dim overlaySurface As String = GetFocusedOverlaySurfaceName()
            If Not String.IsNullOrWhiteSpace(overlaySurface) Then
                Return overlaySurface
            End If

            If app Is Nothing Then
                Return "unknown"
            End If

            If Not IsWordMainWindowForeground(app) Then
                Return "non_document_ui"
            End If

            Dim activeWindow As Microsoft.Office.Interop.Word.Window = app.ActiveWindow
            If activeWindow Is Nothing Then
                Return "unknown"
            End If

            If activeWindow.Type = Microsoft.Office.Interop.Word.WdWindowType.wdWindowDocument AndAlso
               doc IsNot Nothing AndAlso
               sel IsNot Nothing Then
                Return "word_document"
            End If

            Return "non_document_ui"
        Catch
            Return "unknown"
        End Try
    End Function

    Private Shared Function GetFocusedOverlaySurfaceName() As String
        Try
            Dim activeForm As Form = Form.ActiveForm

            If activeForm Is Nothing OrElse activeForm.IsDisposed Then
                Return ""
            End If

            Dim formTypeName As String = activeForm.GetType().Name

            Select Case formTypeName
                Case "TalkToMeWidget"
                    Return ""

                Case "frmAIChat"
                    Return "chatbot"

                Case "HelpMeInky"
                    Return "help_me_inky"

                Case "DiscussInky"
                    Return "discuss_inky"

                Case Else
                    If activeForm.ContainsFocus Then
                        Return "typing_surface"
                    End If
            End Select
        Catch
        End Try

        Return ""
    End Function

    Private Shared Function IsOverlayTypingSurfaceActive() As Boolean
        Return Not String.IsNullOrWhiteSpace(GetFocusedOverlaySurfaceName())
    End Function

    Private Shared Function CanControlWordDocumentNow(app As Microsoft.Office.Interop.Word.Application,
                                                      doc As Microsoft.Office.Interop.Word.Document,
                                                      sel As Microsoft.Office.Interop.Word.Selection) As Boolean
        Try
            If app Is Nothing OrElse doc Is Nothing OrElse sel Is Nothing Then
                Return False
            End If

            Dim activeWindow As Microsoft.Office.Interop.Word.Window = app.ActiveWindow
            If activeWindow Is Nothing Then
                Return False
            End If

            If activeWindow.Type <> Microsoft.Office.Interop.Word.WdWindowType.wdWindowDocument Then
                Return False
            End If

            Dim rng As Microsoft.Office.Interop.Word.Range = sel.Range
            Return rng IsNot Nothing
        Catch
            Return False
        End Try
    End Function

    Private Shared Function CanWriteToDocumentNow(app As Microsoft.Office.Interop.Word.Application, doc As Microsoft.Office.Interop.Word.Document, sel As Microsoft.Office.Interop.Word.Selection) As Boolean

        Try
            If app Is Nothing OrElse doc Is Nothing OrElse sel Is Nothing Then
                Return False
            End If

            Dim activeSurface As String = GetActiveSurfaceName(app, doc, sel)

            If Not String.Equals(activeSurface, "word_document", StringComparison.OrdinalIgnoreCase) Then
                Return False
            End If

            Dim activeWindow As Microsoft.Office.Interop.Word.Window = app.ActiveWindow
            If activeWindow Is Nothing Then
                Return False
            End If

            If activeWindow.Type <> Microsoft.Office.Interop.Word.WdWindowType.wdWindowDocument Then
                Return False
            End If

            Dim rng As Microsoft.Office.Interop.Word.Range = sel.Range
            Return rng IsNot Nothing
        Catch
            Return False
        End Try
    End Function

    Private Shared Function RunOnUiThreadAsync(action As System.Action) As Task
        If action Is Nothing Then
            Return Task.CompletedTask
        End If

        If ThisAddIn.UiSyncContext Is Nothing OrElse
           Thread.CurrentThread.ManagedThreadId = ThisAddIn.UiThreadId Then
            action.Invoke()
            Return Task.CompletedTask
        End If

        Dim tcs As New TaskCompletionSource(Of Object)(TaskCreationOptions.RunContinuationsAsynchronously)

        ThisAddIn.UiSyncContext.Post(
            Sub(state As Object)
                Try
                    action.Invoke()
                    tcs.TrySetResult(Nothing)
                Catch ex As Exception
                    tcs.TrySetException(ex)
                End Try
            End Sub,
            Nothing)

        Return tcs.Task
    End Function

    Private Shared Function RunOnUiThreadAsync(Of T)(func As System.Func(Of T)) As Task(Of T)
        If func Is Nothing Then
            Return Task.FromResult(Of T)(Nothing)
        End If

        If ThisAddIn.UiSyncContext Is Nothing OrElse
           Thread.CurrentThread.ManagedThreadId = ThisAddIn.UiThreadId Then
            Return Task.FromResult(Of T)(func.Invoke())
        End If

        Dim tcs As New TaskCompletionSource(Of T)(TaskCreationOptions.RunContinuationsAsynchronously)

        ThisAddIn.UiSyncContext.Post(
            Sub(state As Object)
                Try
                    tcs.TrySetResult(func.Invoke())
                Catch ex As Exception
                    tcs.TrySetException(ex)
                End Try
            End Sub,
            Nothing)

        Return tcs.Task
    End Function

    Private Shared Sub TypeTextIntoActiveUiTarget(textToInsert As String)
        Dim normalized As String = NormalizeInlineTypingCommands(textToInsert)
        If normalized.Length = 0 Then
            Return
        End If

        Dim currentText As New StringBuilder()

        For Each ch As Char In normalized
            Select Case ch
                Case ChrW(&HE000)
                    If currentText.Length > 0 Then
                        SendKeys.SendWait(EscapeSendKeysText(currentText.ToString()))
                        currentText.Clear()
                    End If

                    SendKeys.SendWait("{ENTER}")

                Case ChrW(&HE001)
                    If currentText.Length > 0 Then
                        SendKeys.SendWait(EscapeSendKeysText(currentText.ToString()))
                        currentText.Clear()
                    End If

                    SendKeys.SendWait("{ESC}")

                Case Else
                    currentText.Append(ch)
            End Select
        Next

        If currentText.Length > 0 Then
            SendKeys.SendWait(EscapeSendKeysText(currentText.ToString()))
        End If
    End Sub

    Private Shared Sub TypeTextIntoWordSelection(selection As Microsoft.Office.Interop.Word.Selection,
                                                 textToInsert As String)
        If selection Is Nothing Then
            Return
        End If

        Dim normalized As String = NormalizeInlineTypingCommands(textToInsert)
        Dim currentText As New StringBuilder()

        For Each ch As Char In normalized
            Select Case ch
                Case ChrW(&HE000)
                    If currentText.Length > 0 Then
                        selection.TypeText(currentText.ToString())
                        currentText.Clear()
                    End If

                    selection.TypeParagraph()

                Case Else
                    currentText.Append(ch)
            End Select
        Next

        If currentText.Length > 0 Then
            selection.TypeText(currentText.ToString())
        End If
    End Sub

    Private Shared Function NormalizeInlineTypingCommands(value As String) As String
        Dim normalized As String = If(value, "")

        normalized = Regex.Replace(
            normalized,
            "\[\[\s*(press\s+)?enter\s*\]\]",
            ChrW(&HE000).ToString(),
            RegexOptions.IgnoreCase)

        normalized = Regex.Replace(
            normalized,
            "\[\[\s*(press\s+)?new\s*line\s*\]\]",
            ChrW(&HE000).ToString(),
            RegexOptions.IgnoreCase)

        normalized = Regex.Replace(
            normalized,
            "\[\[\s*(press\s+)?newline\s*\]\]",
            ChrW(&HE000).ToString(),
            RegexOptions.IgnoreCase)

        normalized = Regex.Replace(
            normalized,
            "\[\[\s*(press\s+)?escape\s*\]\]",
            ChrW(&HE001).ToString(),
            RegexOptions.IgnoreCase)

        normalized = normalized.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
        Return normalized
    End Function

    Private Shared Function EscapeSendKeysText(value As String) As String
        Dim sb As New StringBuilder()

        For Each ch As Char In If(value, "")
            Select Case ch
                Case "+"c, "^"c, "%"c, "~"c, "("c, ")"c, "["c, "]"c, "{"c, "}"c
                    sb.Append("{")
                    sb.Append(ch)
                    sb.Append("}")
                Case Else
                    sb.Append(ch)
            End Select
        Next

        Return sb.ToString()
    End Function
End Class

Public NotInheritable Class WordTalkToMeSpeechAdapter
    Implements SharedLibrary.SharedLibrary.ITalkToMeSpeechAdapter

    Private NotInheritable Class SpeechSettings
        Public Property EngineDisplayName As String = ""
        Public Property LanguageCode As String = "auto"
        Public Property MicrophoneDeviceIndex As Integer = 0
        Public Property SpeechOutputEnabled As Boolean = False
        Public Property SpeechOutputMode As String = "Queue (progressive)"
        Public Property UseSpeakerSpecificVoices As Boolean = False
        Public Property SpeechOutputSpeed As Double = 1.0R
    End Class

    Private NotInheritable Class LiveEngineDescriptor
        Public Property DisplayName As String = ""
        Public Property Kind As EngineKind
        Public Property ModelOrTag As String = ""
        Public Property Languages As List(Of String)
    End Class

    Private ReadOnly _owner As ThisAddIn
    Private _settings As SpeechSettings
    Private _opts As TranscriptionOptions = New TranscriptionOptions()

    Private _alternateOpenAiConfig As ModelConfig = Nothing
    Private _alternateGoogleConfig As ModelConfig = Nothing

    Private _engine As ITranscriptionEngine = Nothing
    Private _capture As AudioCaptureService = Nothing
    Private _cts As CancellationTokenSource = Nothing
    Private _selectedDescriptor As LiveEngineDescriptor = Nothing
    Private _isListeningValue As Boolean = False
    Private _currentEngineDisplayName As String = ""
    Private ReadOnly _speechOutputSyncRoot As New Object()
    Private _speechOutputQueueTail As Task = Task.CompletedTask
    Private _speechOutputPlaybackCts As CancellationTokenSource = Nothing
    Private _speechOutputGeneration As Integer = 0
    Private _speechOutputBusy As Integer = 0
    Private _lastSpeechOutputSpeakerKey As String = ""
    Private _lastSpeechOutputVoiceSlot As Integer = 0

    Private _g1Token As String = ""
    Private _g2Token As String = ""
    Private _gAltToken As String = ""
    Private _g1Exp As DateTime = DateTime.MinValue
    Private _g2Exp As DateTime = DateTime.MinValue
    Private _gAltExp As DateTime = DateTime.MinValue

    Public Event PartialTranscriptReceived As EventHandler(Of SharedLibrary.SharedLibrary.TalkToMeTranscriptEventArgs) Implements SharedLibrary.SharedLibrary.ITalkToMeSpeechAdapter.PartialTranscriptReceived
    Public Event FinalTranscriptReceived As EventHandler(Of SharedLibrary.SharedLibrary.TalkToMeTranscriptEventArgs) Implements SharedLibrary.SharedLibrary.ITalkToMeSpeechAdapter.FinalTranscriptReceived
    Private _sessionRestartInProgress As Integer = 0

    Public Sub New(owner As ThisAddIn)
        _owner = owner
        LoadSettings()
    End Sub

    Public ReadOnly Property IsListening As Boolean Implements SharedLibrary.SharedLibrary.ITalkToMeSpeechAdapter.IsListening
        Get
            Return _isListeningValue
        End Get
    End Property

    Public ReadOnly Property IsConfigured As Boolean Implements SharedLibrary.SharedLibrary.ITalkToMeSpeechAdapter.IsConfigured
        Get
            Return _selectedDescriptor IsNot Nothing
        End Get
    End Property

    Public ReadOnly Property IsSpeechOutputAvailable As Boolean Implements SharedLibrary.SharedLibrary.ITalkToMeSpeechAdapter.IsSpeechOutputAvailable
        Get
            _owner.DetectTTSEngines()
            Return ThisAddIn.TTS_googleAvailable OrElse ThisAddIn.TTS_openAIAvailable
        End Get
    End Property

    Public ReadOnly Property IsSpeechOutputEnabled As Boolean Implements SharedLibrary.SharedLibrary.ITalkToMeSpeechAdapter.IsSpeechOutputEnabled
        Get
            Return _settings IsNot Nothing AndAlso _settings.SpeechOutputEnabled
        End Get
    End Property

    Public ReadOnly Property IsSpeechOutputActive As Boolean Implements SharedLibrary.SharedLibrary.ITalkToMeSpeechAdapter.IsSpeechOutputActive
        Get
            Return Threading.Interlocked.CompareExchange(_speechOutputBusy, 0, 0) <> 0
        End Get
    End Property

    Public ReadOnly Property CanAcceptExternalSpeech As Boolean Implements SharedLibrary.SharedLibrary.ITalkToMeSpeechAdapter.CanAcceptExternalSpeech
        Get
            Return IsListening AndAlso
                   IsSpeechOutputAvailable AndAlso
                   IsSpeechOutputEnabled AndAlso
                   HasConfiguredSpeechOutputSelection()
        End Get
    End Property

    Public Function Configure(ownerWindow As IWin32Window,
                              currentIncludeFullDocument As Boolean) As SharedLibrary.SharedLibrary.TalkToMeSpeechConfigurationResult Implements SharedLibrary.SharedLibrary.ITalkToMeSpeechAdapter.Configure
        Dim descriptors As List(Of LiveEngineDescriptor) = LoadLiveEngines()

        If descriptors.Count = 0 Then
            SharedMethods.ShowCustomMessageBox("No live transcription engines are available for Talk to me!")
            Return New SharedLibrary.SharedLibrary.TalkToMeSpeechConfigurationResult With {
                .Applied = False,
                .IncludeFullDocument = currentIncludeFullDocument,
                .Summary = ""
            }
        End If

        Dim currentEngineDisplayName As String = If(_selectedDescriptor Is Nothing, _settings.EngineDisplayName, _selectedDescriptor.DisplayName)
        Dim savedLanguages As Dictionary(Of String, String) = BuildSavedLanguageDictionary()

        Using dlg As New TalkToMeConfigForm(
            Me,
            descriptors,
            currentEngineDisplayName,
            _settings.LanguageCode,
            _settings.MicrophoneDeviceIndex,
            currentIncludeFullDocument,
            savedLanguages,
            _settings.SpeechOutputEnabled,
            NormalizeSpeechOutputMode(_settings.SpeechOutputMode),
            _settings.UseSpeakerSpecificVoices,
            _settings.SpeechOutputSpeed)

            ' Inspect and same-thread filter the caller-supplied owner before using it as a
            ' modal owner. A foreign (cross-thread/cross-process) owner would be disabled by
            ' ShowDialog and never re-enabled, deadlocking that host. InspectDialogOwner logs
            ' the attempt; IfOwnerOnCurrentThread rejects a cross-thread owner so we fall back
            ' to an ownerless dialog instead of deadlocking.
            Dim effectiveOwner As IWin32Window = ownerWindow
            If effectiveOwner IsNot Nothing Then
                SharedLibrary.SharedLibrary.OfficeWindowWatchdog.InspectDialogOwner(effectiveOwner, "TalkToMeConfigForm", "Configure")
                effectiveOwner = SharedLibrary.SharedLibrary.SharedMethods.IfOwnerOnCurrentThread(effectiveOwner)
            End If

            If dlg.ShowDialog(effectiveOwner) <> DialogResult.OK Then
                Return New SharedLibrary.SharedLibrary.TalkToMeSpeechConfigurationResult With {
                    .Applied = False,
                    .IncludeFullDocument = currentIncludeFullDocument,
                    .Summary = ""
                }
            End If

            _selectedDescriptor = dlg.SelectedDescriptor
            _currentEngineDisplayName = _selectedDescriptor.DisplayName

            _settings.EngineDisplayName = _selectedDescriptor.DisplayName
            _settings.LanguageCode = dlg.SelectedLanguage
            _settings.MicrophoneDeviceIndex = dlg.SelectedMicrophoneDeviceIndex
            _settings.SpeechOutputMode = NormalizeSpeechOutputMode(dlg.SelectedSpeechOutputMode)
            _settings.UseSpeakerSpecificVoices = dlg.SelectedUseSpeakerSpecificVoices
            _settings.SpeechOutputSpeed = NormalizeSpeechOutputSpeed(dlg.SelectedSpeechOutputSpeed)
            _settings.SpeechOutputEnabled = dlg.SelectedSpeechOutputEnabled AndAlso IsSpeechOutputAvailable AndAlso HasConfiguredSpeechOutputSelection()

            If _opts Is Nothing Then
                _opts = New TranscriptionOptions()
            End If

            If _selectedDescriptor.Kind <> EngineKind.Vosk Then
                _opts.LanguageCode = _settings.LanguageCode
            Else
                _opts.LanguageCode = ""
            End If

            PersistSettings()

            Return New SharedLibrary.SharedLibrary.TalkToMeSpeechConfigurationResult With {
                .Applied = True,
                .IncludeFullDocument = dlg.IncludeFullDocument,
                .Summary = GetConfigurationSummary() &
                           " / " &
                           GetSpeechOutputSummary() &
                           If(dlg.IncludeFullDocument, " / Full document enabled", " / Full document disabled")
            }
        End Using
    End Function



    Public Async Function StartListeningAsync(cancellationToken As CancellationToken) As Task Implements SharedLibrary.SharedLibrary.ITalkToMeSpeechAdapter.StartListeningAsync
        If _isListeningValue Then
            Return
        End If

        EnsureSelectedDescriptor()

        Dim d As LiveEngineDescriptor = _selectedDescriptor
        If d Is Nothing Then
            Throw New InvalidOperationException("No transcription engine has been configured.")
        End If

        Dim selectedLanguage As String = If(_settings Is Nothing, "", _settings.LanguageCode)
        Dim micDeviceIndex As Integer = If(_settings Is Nothing, 0, _settings.MicrophoneDeviceIndex)
        Dim sourceMode As AudioSourceMode = GetConfiguredSourceMode()

        If d.Kind <> EngineKind.Vosk AndAlso Not String.IsNullOrWhiteSpace(selectedLanguage) Then
            _opts.LanguageCode = selectedLanguage
        End If

        _currentEngineDisplayName = d.DisplayName

        Dim startException As Exception = Nothing
        Dim engineToDisposeAfterStartFailure As ITranscriptionEngine = Nothing
        Dim ctsToDisposeAfterStartFailure As CancellationTokenSource = Nothing

        Try
            _engine = Await CreateEngineAsync(d)
            AttachEngineEvents(_engine)
            _cts = New CancellationTokenSource()

            System.Diagnostics.Debug.WriteLine(
                "[TalkToMe.Live] About to call StartLiveAsync " &
                "Engine=" & d.DisplayName &
                "; LanguageCode=" & If(_opts Is Nothing OrElse String.IsNullOrWhiteSpace(_opts.LanguageCode), "(empty)", _opts.LanguageCode) &
                "; Model=" & If(_opts Is Nothing OrElse String.IsNullOrWhiteSpace(_opts.Model), "(empty)", _opts.Model) &
                "; MultiChannelDiarization=" & If(_opts IsNot Nothing AndAlso _opts.MultiChannelDiarization, "True", "False"))

            Await _engine.StartLiveAsync(_opts, _cts.Token)
        Catch ex As Exception
            startException = ex
            engineToDisposeAfterStartFailure = _engine
            ctsToDisposeAfterStartFailure = _cts

            _engine = Nothing
            _cts = Nothing
        End Try

        If ctsToDisposeAfterStartFailure IsNot Nothing Then
            Try
                ctsToDisposeAfterStartFailure.Dispose()
            Catch
            End Try
        End If

        If engineToDisposeAfterStartFailure IsNot Nothing Then
            Try
                Await engineToDisposeAfterStartFailure.DisposeAsync()
            Catch
            End Try
        End If

        If startException IsNot Nothing Then
            Throw startException
        End If

        If EngineNeedsLocalAudioCapture(d.Kind) Then
            _capture = New AudioCaptureService With {
                .MicDeviceIndex = micDeviceIndex,
                .MicDeviceId = "",
                .SourceMode = sourceMode,
                .SystemAudioRenderDeviceId = GetConfiguredOutputDeviceId(),
                .MultiChannelStereo = _opts.MultiChannelDiarization AndAlso _engine.SupportsMultiChannelDiarization,
                .AudioDebugDump = _opts.AudioDebugDump OrElse _owner.INI_APIDebug
            }

            AddHandler _capture.Frame, AddressOf OnCaptureFrame
            AddHandler _capture.CaptureError,
                Sub(sender As Object, ev As TranscriptionErrorEventArgs)
                    RaiseEvent FinalTranscriptReceived(
                        Me,
                        New SharedLibrary.SharedLibrary.TalkToMeTranscriptEventArgs("Error: " & ev.Message))

                    ScheduleListeningRestart("Capture error: " & ev.Message)
                End Sub

            _capture.Start()
        End If

        _isListeningValue = True
        PersistSettings()
    End Function

    Public Async Function StopListeningAsync() As Task Implements SharedLibrary.SharedLibrary.ITalkToMeSpeechAdapter.StopListeningAsync
        Dim captureToStop As AudioCaptureService = _capture
        Dim engineToStop As ITranscriptionEngine = _engine
        Dim ctsToDispose As CancellationTokenSource = _cts

        _capture = Nothing
        _engine = Nothing
        _cts = Nothing
        _isListeningValue = False

        CancelCurrentSpeechOutput(clearQueue:=True)

        If captureToStop IsNot Nothing Then
            Try
                RemoveHandler captureToStop.Frame, AddressOf OnCaptureFrame
            Catch
            End Try

            Try
                captureToStop.Stop()
            Catch
            End Try

            Try
                captureToStop.Dispose()
            Catch
            End Try
        End If

        If engineToStop IsNot Nothing Then
            Try
                Await engineToStop.StopLiveAsync()
            Catch
            End Try

            Try
                Await engineToStop.DisposeAsync()
            Catch
            End Try
        End If

        If ctsToDispose IsNot Nothing Then
            Try
                ctsToDispose.Cancel()
            Catch
            End Try

            Try
                ctsToDispose.Dispose()
            Catch
            End Try
        End If
    End Function

    Public Function GetConfigurationSummary() As String Implements SharedLibrary.SharedLibrary.ITalkToMeSpeechAdapter.GetConfigurationSummary
        If _selectedDescriptor Is Nothing Then
            Return "Not configured"
        End If

        Return _selectedDescriptor.DisplayName &
               " / " &
               If(String.IsNullOrWhiteSpace(_settings.LanguageCode), "auto", _settings.LanguageCode) &
               " / Mic " &
               _settings.MicrophoneDeviceIndex.ToString()
    End Function

    Public Function GetSpeechOutputSummary() As String Implements SharedLibrary.SharedLibrary.ITalkToMeSpeechAdapter.GetSpeechOutputSummary
        If Not IsSpeechOutputAvailable Then
            Return "No speech"
        End If

        If Not HasConfiguredSpeechOutputSelection() Then
            Return "Voice unset"
        End If

        Dim summary As String = If(IsSpeechOutputEnabled, "Speech on", "Speech off")

        If _settings IsNot Nothing AndAlso _settings.UseSpeakerSpecificVoices Then
            summary &= " / Two voices"
        End If

        Dim speed As Double = GetSafeSpeechOutputSpeed()
        If speed > 1.0R Then
            summary &= " / " & FormatSpeechOutputSpeed(speed)
        End If

        Return summary
    End Function

    Public Overloads Function ConfigureSpeechOutput(owner As IWin32Window) As String Implements SharedLibrary.SharedLibrary.ITalkToMeSpeechAdapter.ConfigureSpeechOutput
        Return ConfigureSpeechOutput(
            owner,
            _settings IsNot Nothing AndAlso _settings.UseSpeakerSpecificVoices)
    End Function

    Public Overloads Function ConfigureSpeechOutput(owner As IWin32Window,
                                                    useSpeakerSpecificVoices As Boolean) As String
        _owner.DetectTTSEngines()

        If Not IsSpeechOutputAvailable Then
            Return "Speech output unavailable."
        End If

        Using dlg As New ThisAddIn.TTSSelectionForm(
            "Select the speech provider and voice for Talk to me! output.",
            $"{SharedMethods.AN} - Talk to me! Speech Output",
            useSpeakerSpecificVoices,
            True)

            ' Inspect and same-thread filter the caller-supplied owner before using it as a
            ' modal owner. A foreign (cross-thread/cross-process) owner would be disabled by
            ' ShowDialog and never re-enabled, deadlocking that host. InspectDialogOwner logs
            ' the attempt; IfOwnerOnCurrentThread rejects a cross-thread owner so we fall back
            ' to an ownerless dialog instead of deadlocking.
            Dim effectiveOwner As IWin32Window = owner
            If effectiveOwner IsNot Nothing Then
                SharedLibrary.SharedLibrary.OfficeWindowWatchdog.InspectDialogOwner(effectiveOwner, "TTSSelectionForm", "ConfigureSpeechOutput")
                effectiveOwner = SharedLibrary.SharedLibrary.SharedMethods.IfOwnerOnCurrentThread(effectiveOwner)
            End If

            If dlg.ShowDialog(effectiveOwner) = DialogResult.OK Then
                PersistSettings()
            End If
        End Using

        Return GetSpeechOutputSummary()
    End Function

    Public Function ToggleSpeechOutputEnabled() As Boolean Implements SharedLibrary.SharedLibrary.ITalkToMeSpeechAdapter.ToggleSpeechOutputEnabled
        If Not IsSpeechOutputAvailable Then
            _settings.SpeechOutputEnabled = False
            PersistSettings()
            Return False
        End If

        If Not _settings.SpeechOutputEnabled AndAlso Not HasConfiguredSpeechOutputSelection() Then
            Return False
        End If

        _settings.SpeechOutputEnabled = Not _settings.SpeechOutputEnabled

        If Not _settings.SpeechOutputEnabled Then
            CancelCurrentSpeechOutput(clearQueue:=True)
        End If

        PersistSettings()
        Return _settings.SpeechOutputEnabled
    End Function

    Public Sub StopSpeechOutput() Implements SharedLibrary.SharedLibrary.ITalkToMeSpeechAdapter.StopSpeechOutput
        CancelCurrentSpeechOutput(clearQueue:=True)
    End Sub

    Public Function SubmitExternalSpeechAsync(speakerName As String,
                                              text As String,
                                              cancellationToken As CancellationToken) As Task(Of Boolean) Implements SharedLibrary.SharedLibrary.ITalkToMeSpeechAdapter.SubmitExternalSpeechAsync
        Dim speechText As String = PrepareSpeechOutputText(speakerName, text)

        If Not CanAcceptExternalSpeech OrElse String.IsNullOrWhiteSpace(speechText) Then
            Return Task.FromResult(False)
        End If

        Dim mode As String = NormalizeSpeechOutputMode(_settings.SpeechOutputMode)

        Select Case mode
            Case "Interrupt current speech"
                CancelCurrentSpeechOutput(clearQueue:=True)

                Dim immediateTask As Task =
                    Task.Run(
                        Async Function()
                            Await PlaySpeechOutputAsync(speakerName, speechText, cancellationToken).ConfigureAwait(False)
                        End Function)

                SyncLock _speechOutputSyncRoot
                    _speechOutputQueueTail = immediateTask
                End SyncLock

            Case "Skip new output while speaking"
                If Threading.Interlocked.CompareExchange(_speechOutputBusy, 0, 0) <> 0 Then
                    Return Task.FromResult(False)
                End If

                Dim skipTask As Task =
                    Task.Run(
                        Async Function()
                            Await PlaySpeechOutputAsync(speakerName, speechText, cancellationToken).ConfigureAwait(False)
                        End Function)

                SyncLock _speechOutputSyncRoot
                    _speechOutputQueueTail = skipTask
                End SyncLock

            Case "Queue (progressive)"
                Dim progressiveGeneration As Integer
                Dim progressivePreviousTask As Task
                Dim progressiveQueuedTask As Task

                SyncLock _speechOutputSyncRoot
                    progressiveGeneration = _speechOutputGeneration
                    progressivePreviousTask = _speechOutputQueueTail
                    progressiveQueuedTask =
                        QueueSpeechOutputAsync(
                            progressivePreviousTask,
                            progressiveGeneration,
                            speakerName,
                            speechText,
                            cancellationToken,
                            True)
                    _speechOutputQueueTail = progressiveQueuedTask
                End SyncLock

            Case Else
                Dim generation As Integer
                Dim previousTask As Task
                Dim queuedTask As Task

                SyncLock _speechOutputSyncRoot
                    generation = _speechOutputGeneration
                    previousTask = _speechOutputQueueTail
                    queuedTask =
                        QueueSpeechOutputAsync(
                            previousTask,
                            generation,
                            speakerName,
                            speechText,
                            cancellationToken,
                            False)
                    _speechOutputQueueTail = queuedTask
                End SyncLock
        End Select

        Return Task.FromResult(True)
    End Function

    Private Function QueueSpeechOutputAsync(previousTask As Task,
                                            generation As Integer,
                                            speakerName As String,
                                            speechText As String,
                                            cancellationToken As CancellationToken,
                                            useProgressivePlayback As Boolean) As Task
        Return Task.Run(
            Async Function()
                Try
                    If previousTask IsNot Nothing Then
                        Try
                            Await previousTask.ConfigureAwait(False)
                        Catch
                        End Try
                    End If

                    If cancellationToken.IsCancellationRequested Then
                        Return
                    End If

                    If generation <> Threading.Interlocked.CompareExchange(_speechOutputGeneration, 0, 0) Then
                        Return
                    End If

                    If Not CanAcceptExternalSpeech Then
                        Return
                    End If

                    If useProgressivePlayback Then
                        Await PlaySpeechOutputProgressivelyAsync(speakerName, speechText, cancellationToken).ConfigureAwait(False)
                    Else
                        Await PlaySpeechOutputAsync(speakerName, speechText, cancellationToken).ConfigureAwait(False)
                    End If
                Catch
                End Try
            End Function)
    End Function

    Private Async Function GenerateSpeechOutputAudioBytesAsync(speechText As String,
                                                               selection As Tuple(Of String, String, String),
                                                               Optional speakingRate As Double = 1.0R) As Task(Of Byte())
        ThisAddIn.TTS_SelectedEngine =
            If(String.Equals(selection.Item1, "OpenAI", StringComparison.OrdinalIgnoreCase),
               ThisAddIn.TTSEngine.OpenAI,
               ThisAddIn.TTSEngine.Google)

        Return Await ThisAddIn.GenerateAudioFromText(
            speechText,
            selection.Item2,
            selection.Item3,
            SpeakingRate:=NormalizeSpeechOutputSpeed(speakingRate)).ConfigureAwait(False)
    End Function

    Private Async Function PlaySpeechOutputAsync(speakerName As String,
                                                 speechText As String,
                                                 cancellationToken As CancellationToken) As Task
        If String.IsNullOrWhiteSpace(speechText) Then
            Return
        End If

        Dim selection As Tuple(Of String, String, String) = GetConfiguredSpeechOutputSelection(speakerName)
        Dim playbackRate As Double = GetSafeSpeechOutputSpeed()
        Dim tempFilePath As String = ""
        Dim playbackCts As CancellationTokenSource = Nothing

        Threading.Interlocked.Exchange(_speechOutputBusy, 1)

        Try
            playbackCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken)

            SyncLock _speechOutputSyncRoot
                _speechOutputPlaybackCts = playbackCts
            End SyncLock

            Dim audioBytes As Byte() =
                Await GenerateSpeechOutputAudioBytesAsync(
                    speechText,
                    selection).ConfigureAwait(False)

            If audioBytes Is Nothing OrElse audioBytes.Length = 0 OrElse playbackCts.Token.IsCancellationRequested Then
                Return
            End If

            tempFilePath = System.IO.Path.ChangeExtension(System.IO.Path.GetTempFileName(), ".mp3")
            ThisAddIn.SaveAudioToFile(audioBytes, tempFilePath)

            Await PlayAudioFileAsync(
                tempFilePath,
                playbackCts.Token,
                playbackRate,
                True,
                1200).ConfigureAwait(False)

        Catch ex As OperationCanceledException
        Catch
        Finally
            SyncLock _speechOutputSyncRoot
                If ReferenceEquals(_speechOutputPlaybackCts, playbackCts) Then
                    _speechOutputPlaybackCts = Nothing
                End If
            End SyncLock

            If playbackCts IsNot Nothing Then
                Try
                    playbackCts.Dispose()
                Catch
                End Try
            End If

            If Not String.IsNullOrWhiteSpace(tempFilePath) Then
                Try
                    If System.IO.File.Exists(tempFilePath) Then
                        System.IO.File.Delete(tempFilePath)
                    End If
                Catch
                End Try
            End If

            Threading.Interlocked.Exchange(_speechOutputBusy, 0)
        End Try
    End Function

    Private Async Function PlaySpeechOutputProgressivelyAsync(speakerName As String,
                                                              speechText As String,
                                                              cancellationToken As CancellationToken) As Task
        Dim chunks As List(Of String) = SplitSpeechOutputIntoChunks(speechText, 2)

        If chunks.Count <= 1 Then
            Await PlaySpeechOutputAsync(speakerName, speechText, cancellationToken).ConfigureAwait(False)
            Return
        End If

        Dim selection As Tuple(Of String, String, String) = GetConfiguredSpeechOutputSelection(speakerName)
        Dim playbackRate As Double = GetSafeSpeechOutputSpeed()
        Dim playbackCts As CancellationTokenSource = Nothing
        Dim tempFiles As New List(Of String)

        Threading.Interlocked.Exchange(_speechOutputBusy, 1)

        Try
            playbackCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken)

            SyncLock _speechOutputSyncRoot
                _speechOutputPlaybackCts = playbackCts
            End SyncLock

            Dim pendingAudioTask As Task(Of Byte()) =
                GenerateSpeechOutputAudioBytesAsync(chunks(0), selection)

            For chunkIndex As Integer = 0 To chunks.Count - 1
                Dim audioBytes As Byte() = Await pendingAudioTask.ConfigureAwait(False)

                If audioBytes Is Nothing OrElse audioBytes.Length = 0 OrElse playbackCts.Token.IsCancellationRequested Then
                    Return
                End If

                If chunkIndex < chunks.Count - 1 Then
                    pendingAudioTask = GenerateSpeechOutputAudioBytesAsync(chunks(chunkIndex + 1), selection)
                Else
                    pendingAudioTask = Nothing
                End If

                Dim tempFilePath As String = System.IO.Path.ChangeExtension(System.IO.Path.GetTempFileName(), ".mp3")
                tempFiles.Add(tempFilePath)
                ThisAddIn.SaveAudioToFile(audioBytes, tempFilePath)

                Await PlayAudioFileAsync(
                    tempFilePath,
                    playbackCts.Token,
                    playbackRate,
                    chunkIndex = 0,
                    1200).ConfigureAwait(False)
            Next

        Catch ex As OperationCanceledException
        Catch
        Finally
            SyncLock _speechOutputSyncRoot
                If ReferenceEquals(_speechOutputPlaybackCts, playbackCts) Then
                    _speechOutputPlaybackCts = Nothing
                End If
            End SyncLock

            If playbackCts IsNot Nothing Then
                Try
                    playbackCts.Dispose()
                Catch
                End Try
            End If

            For Each tempFilePath As String In tempFiles
                Try
                    If Not String.IsNullOrWhiteSpace(tempFilePath) AndAlso System.IO.File.Exists(tempFilePath) Then
                        System.IO.File.Delete(tempFilePath)
                    End If
                Catch
                End Try
            Next

            Threading.Interlocked.Exchange(_speechOutputBusy, 0)
        End Try
    End Function

    Private Shared Function SplitSpeechOutputIntoChunks(speechText As String,
                                                        sentencesPerChunk As Integer) As List(Of String)
        Dim result As New List(Of String)
        Dim normalized As String = Regex.Replace(If(speechText, ""), "\s+", " ").Trim()

        If String.IsNullOrWhiteSpace(normalized) Then
            Return result
        End If

        Dim sentences As New List(Of String)

        For Each m As Match In Regex.Matches(normalized, "[^.!?]+(?:[.!?]+|$)")
            Dim sentence As String = m.Value.Trim()
            If Not String.IsNullOrWhiteSpace(sentence) Then
                sentences.Add(sentence)
            End If
        Next

        If sentences.Count = 0 Then
            result.Add(normalized)
            Return result
        End If

        If sentences.Count <= sentencesPerChunk Then
            result.Add(String.Join(" ", sentences))
            Return result
        End If

        For i As Integer = 0 To sentences.Count - 1 Step sentencesPerChunk
            result.Add(String.Join(" ", sentences.Skip(i).Take(sentencesPerChunk)))
        Next

        Return result
    End Function

    Private NotInheritable Class LeadingSilenceSampleProvider
        Implements NAudio.Wave.ISampleProvider

        Private ReadOnly _waveFormat As NAudio.Wave.WaveFormat
        Private _remainingSamples As Integer

        Public Sub New(waveFormat As NAudio.Wave.WaveFormat, duration As TimeSpan)
            _waveFormat = waveFormat

            Dim totalSamples As Double =
                duration.TotalSeconds *
                CDbl(waveFormat.SampleRate) *
                CDbl(waveFormat.Channels)

            _remainingSamples = Math.Max(0, CInt(Math.Ceiling(totalSamples)))
        End Sub

        Public ReadOnly Property WaveFormat As NAudio.Wave.WaveFormat Implements NAudio.Wave.ISampleProvider.WaveFormat
            Get
                Return _waveFormat
            End Get
        End Property

        Public Function Read(buffer As Single(), offset As Integer, count As Integer) As Integer Implements NAudio.Wave.ISampleProvider.Read
            If _remainingSamples <= 0 Then
                Return 0
            End If

            Dim samplesToWrite As Integer = Math.Min(count, _remainingSamples)
            Array.Clear(buffer, offset, samplesToWrite)
            _remainingSamples -= samplesToWrite
            Return samplesToWrite
        End Function
    End Class

    Private NotInheritable Class PlaybackSpeedSampleProvider
        Implements NAudio.Wave.ISampleProvider

        Private ReadOnly _pipeline As NAudio.Wave.ISampleProvider

        Public Sub New(source As NAudio.Wave.ISampleProvider, speed As Double)
            If source Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(source))
            End If

            If speed <= 0.0R Then
                Throw New System.ArgumentOutOfRangeException(NameOf(speed), "Speed must be greater than zero.")
            End If

            If source.WaveFormat.Channels <> 1 Then
                Throw New System.Exception("This WSOLA implementation is mono-only. Convert to mono before using it.")
            End If

            If System.Math.Abs(speed - 1.0R) <= 0.01R Then
                _pipeline = source
            Else
                _pipeline = New MonoWsolaSpeedSampleProvider(source, speed)
            End If
        End Sub

        Public ReadOnly Property WaveFormat As NAudio.Wave.WaveFormat Implements NAudio.Wave.ISampleProvider.WaveFormat
            Get
                Return _pipeline.WaveFormat
            End Get
        End Property

        Public Function Read(buffer As Single(), offset As Integer, count As Integer) As Integer Implements NAudio.Wave.ISampleProvider.Read
            Return _pipeline.Read(buffer, offset, count)
        End Function

    End Class

    Private NotInheritable Class MonoWsolaSpeedSampleProvider
        Implements NAudio.Wave.ISampleProvider

        Private ReadOnly _waveFormat As NAudio.Wave.WaveFormat
        Private ReadOnly _samples As Single()
        Private _position As Integer

        Public Sub New(source As NAudio.Wave.ISampleProvider, speed As Double)
            If source Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(source))
            End If

            If source.WaveFormat.Channels <> 1 Then
                Throw New System.Exception("MonoWsolaSpeedSampleProvider requires mono audio.")
            End If

            If speed <= 0.0R Then
                Throw New System.ArgumentOutOfRangeException(NameOf(speed), "Speed must be greater than zero.")
            End If

            _waveFormat = source.WaveFormat

            Dim inputSamples As Single() = ReadAllSamples(source)

            If inputSamples.Length = 0 Then
                _samples = New Single() {}
                Return
            End If

            If System.Math.Abs(speed - 1.0R) <= 0.01R Then
                _samples = inputSamples
            Else
                _samples = TimeStretchMonoWsola(inputSamples, _waveFormat.SampleRate, speed)
            End If
        End Sub

        Public ReadOnly Property WaveFormat As NAudio.Wave.WaveFormat Implements NAudio.Wave.ISampleProvider.WaveFormat
            Get
                Return _waveFormat
            End Get
        End Property

        Public Function Read(buffer As Single(), offset As Integer, count As Integer) As Integer Implements NAudio.Wave.ISampleProvider.Read
            If buffer Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(buffer))
            End If

            If count <= 0 Then
                Return 0
            End If

            Dim available As Integer = _samples.Length - _position

            If available <= 0 Then
                Return 0
            End If

            Dim samplesToCopy As Integer = System.Math.Min(count, available)

            For i As Integer = 0 To samplesToCopy - 1
                buffer(offset + i) = _samples(_position + i)
            Next

            _position += samplesToCopy

            Return samplesToCopy
        End Function

        Private Shared Function ReadAllSamples(source As NAudio.Wave.ISampleProvider) As Single()
            Dim result As New System.Collections.Generic.List(Of Single)()
            Dim temp(32767) As Single

            Do
                Dim read As Integer = source.Read(temp, 0, temp.Length)

                If read <= 0 Then
                    Exit Do
                End If

                For i As Integer = 0 To read - 1
                    result.Add(temp(i))
                Next
            Loop

            Return result.ToArray()
        End Function

        Private Shared Function TimeStretchMonoWsola(input As Single(), sampleRate As Integer, speed As Double) As Single()
            Dim overlapLength As Integer = System.Math.Max(64, CInt(System.Math.Round(sampleRate * 0.01R)))
            Dim frameLength As Integer = System.Math.Max(overlapLength * 4, CInt(System.Math.Round(sampleRate * 0.04R)))
            Dim seekWindowLength As Integer = System.Math.Max(overlapLength, CInt(System.Math.Round(sampleRate * 0.02R)))

            If input.Length <= frameLength + seekWindowLength + overlapLength Then
                Return SimpleLinearResample(input, speed)
            End If

            Dim synthesisHop As Integer = frameLength - overlapLength
            Dim analysisHop As Double = CDbl(synthesisHop) * speed
            Dim targetLength As Integer = System.Math.Max(1, CInt(System.Math.Round(CDbl(input.Length) / speed)))

            Dim output As New System.Collections.Generic.List(Of Single)(targetLength + frameLength)

            For i As Integer = 0 To frameLength - 1
                output.Add(input(i))
            Next

            Dim inputPosition As Double = analysisHop

            Do While output.Count < targetLength
                Dim predictedPosition As Integer = CInt(System.Math.Round(inputPosition))

                If predictedPosition + frameLength >= input.Length Then
                    Exit Do
                End If

                Dim bestPosition As Integer =
                FindBestOverlapPosition(
                    input,
                    predictedPosition,
                    seekWindowLength,
                    overlapLength,
                    output)

                CrossfadeAppend(input, bestPosition, frameLength, overlapLength, output)

                inputPosition += analysisHop
            Loop

            If output.Count > targetLength Then
                output.RemoveRange(targetLength, output.Count - targetLength)
            End If

            Return output.ToArray()
        End Function

        Private Shared Function FindBestOverlapPosition(input As Single(),
                                                    predictedPosition As Integer,
                                                    seekWindowLength As Integer,
                                                    overlapLength As Integer,
                                                    output As System.Collections.Generic.List(Of Single)) As Integer

            Dim searchStart As Integer = System.Math.Max(0, predictedPosition - seekWindowLength)
            Dim searchEnd As Integer = System.Math.Min(input.Length - overlapLength - 1, predictedPosition + seekWindowLength)

            Dim outputOverlapStart As Integer = output.Count - overlapLength

            Dim bestPosition As Integer = predictedPosition
            Dim bestScore As Double = Double.NegativeInfinity

            For candidate As Integer = searchStart To searchEnd
                Dim score As Double = NormalizedCorrelation(input, candidate, output, outputOverlapStart, overlapLength)

                If score > bestScore Then
                    bestScore = score
                    bestPosition = candidate
                End If
            Next

            Return bestPosition
        End Function

        Private Shared Function NormalizedCorrelation(input As Single(),
                                                  inputStart As Integer,
                                                  output As System.Collections.Generic.List(Of Single),
                                                  outputStart As Integer,
                                                  length As Integer) As Double

            Dim sum As Double = 0.0R
            Dim inputEnergy As Double = 0.0R
            Dim outputEnergy As Double = 0.0R

            For i As Integer = 0 To length - 1
                Dim a As Double = CDbl(output(outputStart + i))
                Dim b As Double = CDbl(input(inputStart + i))

                sum += a * b
                outputEnergy += a * a
                inputEnergy += b * b
            Next

            If inputEnergy <= 0.000000000001R OrElse outputEnergy <= 0.000000000001R Then
                Return Double.NegativeInfinity
            End If

            Return sum / System.Math.Sqrt(inputEnergy * outputEnergy)
        End Function

        Private Shared Sub CrossfadeAppend(input As Single(),
                                       inputStart As Integer,
                                       frameLength As Integer,
                                       overlapLength As Integer,
                                       output As System.Collections.Generic.List(Of Single))

            Dim outputOverlapStart As Integer = output.Count - overlapLength

            For i As Integer = 0 To overlapLength - 1
                Dim fadeIn As Double = CDbl(i) / CDbl(overlapLength)
                Dim fadeOut As Double = 1.0R - fadeIn

                Dim oldSample As Double = CDbl(output(outputOverlapStart + i))
                Dim newSample As Double = CDbl(input(inputStart + i))

                output(outputOverlapStart + i) = CSng((oldSample * fadeOut) + (newSample * fadeIn))
            Next

            Dim appendStart As Integer = inputStart + overlapLength
            Dim appendEnd As Integer = System.Math.Min(input.Length, inputStart + frameLength)

            For i As Integer = appendStart To appendEnd - 1
                output.Add(input(i))
            Next
        End Sub

        Private Shared Function SimpleLinearResample(input As Single(), speed As Double) As Single()
            Dim targetLength As Integer = System.Math.Max(1, CInt(System.Math.Round(CDbl(input.Length) / speed)))
            Dim output(targetLength - 1) As Single

            For i As Integer = 0 To targetLength - 1
                Dim sourcePosition As Double = CDbl(i) * speed
                Dim index0 As Integer = CInt(System.Math.Floor(sourcePosition))
                Dim index1 As Integer = System.Math.Min(index0 + 1, input.Length - 1)
                Dim fraction As Double = sourcePosition - CDbl(index0)

                If index0 >= input.Length Then
                    output(i) = input(input.Length - 1)
                Else
                    output(i) = CSng((CDbl(input(index0)) * (1.0R - fraction)) + (CDbl(input(index1)) * fraction))
                End If
            Next

            Return output
        End Function

    End Class

    Private Shared Function CreateSpeechOutputWaveProvider(audioReader As AudioFileReader,
                                                           playbackRate As Double,
                                                           Optional includeLeadingSilence As Boolean = True,
                                                           Optional leadingSilenceMilliseconds As Integer = 1200) As IWaveProvider
        Dim originalSource As NAudio.Wave.ISampleProvider = audioReader
        Dim sampleProvider As NAudio.Wave.ISampleProvider = originalSource
        Dim normalizedRate As Double = NormalizeSpeechOutputSpeed(playbackRate)

        If normalizedRate > 1.0R Then
            Dim mono As NAudio.Wave.ISampleProvider = originalSource

            If originalSource.WaveFormat.Channels = 2 Then
                mono =
                    New NAudio.Wave.SampleProviders.StereoToMonoSampleProvider(originalSource) With {
                        .LeftVolume = 0.5F,
                        .RightVolume = 0.5F
                    }
            End If

            sampleProvider = New PlaybackSpeedSampleProvider(mono, normalizedRate)
        End If

        If includeLeadingSilence AndAlso leadingSilenceMilliseconds > 0 Then
            Dim leadingSilence As NAudio.Wave.ISampleProvider =
                New LeadingSilenceSampleProvider(
                    sampleProvider.WaveFormat,
                    TimeSpan.FromMilliseconds(leadingSilenceMilliseconds))

            sampleProvider =
                New NAudio.Wave.SampleProviders.ConcatenatingSampleProvider(
                    New List(Of NAudio.Wave.ISampleProvider) From {
                        leadingSilence,
                        sampleProvider
                    })
        End If

        Return New NAudio.Wave.SampleProviders.SampleToWaveProvider(sampleProvider)
    End Function

    Private Shared Async Function PlayAudioFileAsync(filePath As String,
                                                     cancellationToken As CancellationToken,
                                                     Optional playbackRate As Double = 1.0R,
                                                     Optional includeLeadingSilence As Boolean = True,
                                                     Optional leadingSilenceMilliseconds As Integer = 1200) As Task
        Using reader As New AudioFileReader(filePath)
            Using waveOut As New WaveOutEvent()
                ThisAddIn.ConfigurePlaybackOutput(waveOut)
                waveOut.Init(CreateSpeechOutputWaveProvider(reader, playbackRate, includeLeadingSilence, leadingSilenceMilliseconds))
                waveOut.Play()

                Try
                    While waveOut.PlaybackState <> PlaybackState.Stopped
                        cancellationToken.ThrowIfCancellationRequested()
                        Await Task.Delay(100, cancellationToken).ConfigureAwait(False)
                    End While
                Finally
                    Try
                        waveOut.Stop()
                    Catch
                    End Try
                End Try
            End Using
        End Using
    End Function

    Private Sub CancelCurrentSpeechOutput(clearQueue As Boolean)
        Dim ctsToCancel As CancellationTokenSource = Nothing

        SyncLock _speechOutputSyncRoot
            Threading.Interlocked.Increment(_speechOutputGeneration)
            ctsToCancel = _speechOutputPlaybackCts

            If clearQueue Then
                _speechOutputQueueTail = Task.CompletedTask
                _lastSpeechOutputSpeakerKey = ""
                _lastSpeechOutputVoiceSlot = 0
            End If
        End SyncLock

        If ctsToCancel IsNot Nothing Then
            Try
                ctsToCancel.Cancel()
            Catch
            End Try
        End If
    End Sub

    Private Function HasConfiguredSpeechOutputSelection() As Boolean
        If _settings IsNot Nothing AndAlso _settings.UseSpeakerSpecificVoices Then
            Dim firstSelection As Tuple(Of String, String, String) = Nothing
            Dim secondSelection As Tuple(Of String, String, String) = Nothing

            Return TryGetConfiguredDualSpeechOutputSelections(firstSelection, secondSelection)
        End If

        Return IsValidSpeechOutputSelection(GetConfiguredPrimarySpeechOutputSelection())
    End Function

    Private Shared Function IsValidSpeechOutputSelection(selection As Tuple(Of String, String, String)) As Boolean
        Return selection IsNot Nothing AndAlso
               Not String.IsNullOrWhiteSpace(selection.Item1) AndAlso
               Not String.IsNullOrWhiteSpace(selection.Item2) AndAlso
               Not String.IsNullOrWhiteSpace(selection.Item3)
    End Function

    Private Function GetConfiguredSpeechOutputSelection(Optional speakerName As String = "") As Tuple(Of String, String, String)
        If _settings IsNot Nothing AndAlso _settings.UseSpeakerSpecificVoices Then
            Dim firstSelection As Tuple(Of String, String, String) = Nothing
            Dim secondSelection As Tuple(Of String, String, String) = Nothing

            If TryGetConfiguredDualSpeechOutputSelections(firstSelection, secondSelection) Then
                Dim normalizedSpeakerName As String = NormalizeSpeakerNameForVoiceSelection(speakerName)

                If String.IsNullOrWhiteSpace(normalizedSpeakerName) Then
                    Return firstSelection
                End If

                SyncLock _speechOutputSyncRoot
                    If String.IsNullOrWhiteSpace(_lastSpeechOutputSpeakerKey) Then
                        _lastSpeechOutputSpeakerKey = normalizedSpeakerName
                        _lastSpeechOutputVoiceSlot = 0
                    ElseIf Not String.Equals(_lastSpeechOutputSpeakerKey, normalizedSpeakerName, StringComparison.OrdinalIgnoreCase) Then
                        _lastSpeechOutputSpeakerKey = normalizedSpeakerName
                        _lastSpeechOutputVoiceSlot = If(_lastSpeechOutputVoiceSlot = 0, 1, 0)
                    End If

                    Return If(_lastSpeechOutputVoiceSlot = 0, firstSelection, secondSelection)
                End SyncLock
            End If
        End If

        Return GetConfiguredPrimarySpeechOutputSelection()
    End Function

    Private Function GetConfiguredPrimarySpeechOutputSelection() As Tuple(Of String, String, String)
        Dim provider As String = If(My.Settings.TTSProvider, "").Trim()
        Dim languageCode As String = If(My.Settings.TTS1languagecode, "").Trim()
        Dim voiceName As String = NormalizeSpeechOutputVoice(If(My.Settings.TTS1voiceA, ""))

        Select Case If(My.Settings.TTSLastRdoOneVoice, "")
            Case "Voice1B"
                voiceName = NormalizeSpeechOutputVoice(If(My.Settings.TTS1voiceB, ""))

            Case "Voice2A"
                languageCode = If(My.Settings.TTS2languagecode, "").Trim()
                voiceName = NormalizeSpeechOutputVoice(If(My.Settings.TTS2voiceA, ""))

            Case "Voice2B"
                languageCode = If(My.Settings.TTS2languagecode, "").Trim()
                voiceName = NormalizeSpeechOutputVoice(If(My.Settings.TTS2voiceB, ""))
        End Select

        Return Tuple.Create(provider, languageCode, voiceName)
    End Function

    Private Function TryGetConfiguredDualSpeechOutputSelections(ByRef firstSelection As Tuple(Of String, String, String),
                                                                ByRef secondSelection As Tuple(Of String, String, String)) As Boolean
        Dim provider As String = If(My.Settings.TTSProvider, "").Trim()
        Dim setName As String = If(My.Settings.TTSLastRdoTwoVoices, "").Trim()
        Dim useSecondSet As Boolean = String.Equals(setName, "Voice2", StringComparison.OrdinalIgnoreCase)
        Dim languageCode As String =
            If(useSecondSet,
               If(My.Settings.TTS2languagecode, "").Trim(),
               If(My.Settings.TTS1languagecode, "").Trim())

        Dim firstVoiceName As String =
            NormalizeSpeechOutputVoice(
                If(useSecondSet, My.Settings.TTS2voiceA, My.Settings.TTS1voiceA))

        Dim secondVoiceName As String =
            NormalizeSpeechOutputVoice(
                If(useSecondSet, My.Settings.TTS2voiceB, My.Settings.TTS1voiceB))

        firstSelection = Tuple.Create(provider, languageCode, firstVoiceName)
        secondSelection = Tuple.Create(provider, languageCode, secondVoiceName)

        Return IsValidSpeechOutputSelection(firstSelection) AndAlso
               IsValidSpeechOutputSelection(secondSelection)
    End Function

    Private Shared Function NormalizeSpeakerNameForVoiceSelection(value As String) As String
        Return Regex.Replace(If(value, "").Trim(), "\s+", " ")
    End Function

    Private Shared Function NormalizeSpeechOutputVoice(value As String) As String
        Dim normalized As String = If(value, "").Trim()

        If normalized.EndsWith(" (male)", StringComparison.OrdinalIgnoreCase) Then
            normalized = normalized.Substring(0, normalized.Length - " (male)".Length).Trim()
        ElseIf normalized.EndsWith(" (female)", StringComparison.OrdinalIgnoreCase) Then
            normalized = normalized.Substring(0, normalized.Length - " (female)".Length).Trim()
        End If

        Dim separatorIndex As Integer = normalized.IndexOf(" — ", StringComparison.Ordinal)
        If separatorIndex > 0 Then
            normalized = normalized.Substring(0, separatorIndex).Trim()
        End If

        Return normalized
    End Function

    Private Shared Function NormalizeSpeechOutputMode(value As String) As String
        Dim normalizedValue As String = If(value, "").Trim().ToLowerInvariant()

        If String.IsNullOrWhiteSpace(normalizedValue) Then
            Return "Queue (progressive)"
        End If

        Select Case normalizedValue
            Case "queue (progressive)", "progressive", "chunked", "queue progressive"
                Return "Queue (progressive)"

            Case "interrupt current speech", "interrupt"
                Return "Interrupt current speech"

            Case "skip new output while speaking", "skip"
                Return "Skip new output while speaking"

            Case Else
                Return "Queue"
        End Select
    End Function

    Private Shared Function NormalizeSpeechOutputSpeed(value As Double) As Double
        If Double.IsNaN(value) OrElse Double.IsInfinity(value) Then
            Return 1.0R
        End If

        Return Math.Max(1.0R, Math.Min(2.0R, Math.Round(value, 2, MidpointRounding.AwayFromZero)))
    End Function

    Private Shared Function ParseSpeechOutputSpeed(value As String) As Double
        Dim normalized As String = If(value, "").Trim()
        normalized = normalized.Replace("(Normal)", "").Replace("(normal)", "")
        normalized = normalized.Replace("x", "").Replace("X", "").Trim()

        Dim parsed As Double

        If Double.TryParse(
            normalized,
            Globalization.NumberStyles.Float,
            Globalization.CultureInfo.InvariantCulture,
            parsed) Then

            Return NormalizeSpeechOutputSpeed(parsed)
        End If

        If Double.TryParse(normalized, parsed) Then
            Return NormalizeSpeechOutputSpeed(parsed)
        End If

        Return 1.0R
    End Function

    Private Shared Function FormatSpeechOutputSpeed(value As Double) As String
        Return NormalizeSpeechOutputSpeed(value).ToString("0.00", Globalization.CultureInfo.InvariantCulture) & "x"
    End Function

    Private Function GetSafeSpeechOutputSpeed() As Double
        If _settings Is Nothing Then
            Return 1.0R
        End If

        Return NormalizeSpeechOutputSpeed(_settings.SpeechOutputSpeed)
    End Function

    Private Shared Function PrepareSpeechOutputText(speakerName As String,
                                                    text As String) As String
        Dim normalized As String = If(text, "")

        normalized = Regex.Replace(normalized, "```[\s\S]*?```", " ")
        normalized = Regex.Replace(normalized, "`([^`]*)`", "$1")
        normalized = Regex.Replace(normalized, "!\[([^\]]*)\]\([^)]+\)", "$1")
        normalized = Regex.Replace(normalized, "\[([^\]]+)\]\([^)]+\)", "$1")
        normalized = normalized.Replace("#", " ")
        normalized = normalized.Replace("*", " ")
        normalized = normalized.Replace("_", " ")
        normalized = normalized.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
        normalized = Regex.Replace(normalized, "\n{3,}", vbLf & vbLf)
        normalized = Regex.Replace(normalized, "[ \t]{2,}", " ")
        normalized = normalized.Trim()

        If Not String.IsNullOrWhiteSpace(speakerName) Then
            Dim prefix As String = speakerName.Trim() & ":"
            If normalized.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then
                normalized = normalized.Substring(prefix.Length).Trim()
            End If
        End If

        If String.IsNullOrWhiteSpace(normalized) Then
            Return ""
        End If

        Return normalized
    End Function

    Private Shared Function EngineNeedsLocalAudioCapture(kind As EngineKind) As Boolean
        Select Case kind
            Case EngineKind.TeamsAcsRealtime
                Return False
            Case Else
                Return True
        End Select
    End Function

    Private Function GetConfiguredSourceMode() As AudioSourceMode
        Dim raw As String = If(String.IsNullOrWhiteSpace(My.Settings.LastAudioSourceMode), "MicrophoneOnly", My.Settings.LastAudioSourceMode)

        Try
            Return CType([Enum].Parse(GetType(AudioSourceMode), raw), AudioSourceMode)
        Catch
            Return AudioSourceMode.MicrophoneOnly
        End Try
    End Function

    Private Function GetConfiguredOutputDeviceId() As String
        Return If(My.Settings.LastAudioOutputDeviceId, "")
    End Function

    Private Sub LoadSettings()
        _settings = New SpeechSettings()
        _opts = New TranscriptionOptions()

        Try
            If Not String.IsNullOrWhiteSpace(My.Settings.TalkToMeSpeechEngineName) Then
                _settings.EngineDisplayName = My.Settings.TalkToMeSpeechEngineName
            End If

            If My.Settings.TalkToMeSpeechMicrophoneDeviceIndex >= 0 Then
                _settings.MicrophoneDeviceIndex = My.Settings.TalkToMeSpeechMicrophoneDeviceIndex
            End If

            If Not String.IsNullOrWhiteSpace(My.Settings.TalkToMeSpeechLanguageCode) Then
                _settings.LanguageCode = My.Settings.TalkToMeSpeechLanguageCode
            End If

            _settings.SpeechOutputEnabled = My.Settings.TalkToMeSpeechOutputEnabled

            If Not String.IsNullOrWhiteSpace(My.Settings.TalkToMeSpeechOutputMode) Then
                _settings.SpeechOutputMode = NormalizeSpeechOutputMode(My.Settings.TalkToMeSpeechOutputMode)
            End If

            _settings.UseSpeakerSpecificVoices = My.Settings.TalkToMeSpeechOutputUseSpeakerSpecificVoices
            _settings.SpeechOutputSpeed = NormalizeSpeechOutputSpeed(My.Settings.TalkToMeSpeechOutputSpeed)
        Catch
            _opts = New TranscriptionOptions()
        End Try

        EnsureSelectedDescriptor()

        If String.IsNullOrWhiteSpace(_settings.LanguageCode) Then
            _settings.LanguageCode = "auto"
        End If

        _settings.SpeechOutputMode = NormalizeSpeechOutputMode(_settings.SpeechOutputMode)
        _settings.SpeechOutputSpeed = NormalizeSpeechOutputSpeed(_settings.SpeechOutputSpeed)
        _opts.LanguageCode = _settings.LanguageCode
    End Sub

    Private Sub PersistSettings()
        Try
            SaveCurrentEngineSelection()
            My.Settings.TalkToMeSpeechMicrophoneDeviceIndex = _settings.MicrophoneDeviceIndex
            My.Settings.TalkToMeSpeechLanguageCode = _settings.LanguageCode
            My.Settings.TalkToMeSpeechOutputEnabled = _settings.SpeechOutputEnabled
            My.Settings.TalkToMeSpeechOutputMode = NormalizeSpeechOutputMode(_settings.SpeechOutputMode)
            My.Settings.TalkToMeSpeechOutputUseSpeakerSpecificVoices = _settings.UseSpeakerSpecificVoices
            My.Settings.TalkToMeSpeechOutputSpeed = NormalizeSpeechOutputSpeed(_settings.SpeechOutputSpeed)
            My.Settings.Save()
        Catch
        End Try
    End Sub

    Private Function FindEngineIndexByDisplayName(displayName As String) As Integer
        Dim descriptors As List(Of LiveEngineDescriptor) = LoadLiveEngines()

        If String.IsNullOrWhiteSpace(displayName) Then
            Return -1
        End If

        Dim normalized As String = displayName.Trim()

        For i As Integer = 0 To descriptors.Count - 1
            If String.Equals(descriptors(i).DisplayName, normalized, StringComparison.OrdinalIgnoreCase) Then
                Return i
            End If
        Next

        Return -1
    End Function

    Private Sub EnsureSelectedDescriptor()
        Dim descriptors As List(Of LiveEngineDescriptor) = LoadLiveEngines()
        Dim engineIndex As Integer = FindEngineIndexByDisplayName(_settings.EngineDisplayName)

        If engineIndex >= 0 AndAlso engineIndex < descriptors.Count Then
            _selectedDescriptor = descriptors(engineIndex)
            _currentEngineDisplayName = _selectedDescriptor.DisplayName
            Return
        End If

        _selectedDescriptor = descriptors.FirstOrDefault()
        If _selectedDescriptor IsNot Nothing Then
            _currentEngineDisplayName = _selectedDescriptor.DisplayName
        End If
    End Sub

    Private Sub SaveCurrentEngineSelection()
        If _selectedDescriptor Is Nothing Then
            Return
        End If

        My.Settings.TalkToMeSpeechEngineName = _selectedDescriptor.DisplayName
    End Sub

    Private Shared Function GetLanguagePersistenceKey(d As LiveEngineDescriptor) As String
        If d Is Nothing Then
            Return ""
        End If

        Return d.DisplayName
    End Function

    Private Function LoadLastLanguageMap() As JObject
        Return New JObject()
    End Function

    Private Sub SaveLastLanguageMap(map As JObject)
    End Sub

    Private Function BuildSavedLanguageDictionary() As Dictionary(Of String, String)
        Return New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
    End Function

    Private Function GetSavedLanguageForDescriptor(d As LiveEngineDescriptor) As String
        Return ""
    End Function

    Private Sub SaveCurrentLanguageForCurrentEngine()
    End Sub

    Private Function LoadLiveEngines() As List(Of LiveEngineDescriptor)
        Dim result As New List(Of LiveEngineDescriptor)()
        Dim modelRoot As String = SharedMethods.ExpandEnvironmentVariables(_owner.INI_SpeechModelPath)

        LoadAlternateProviderFallbacks()

        If Directory.Exists(modelRoot) Then
            For Each directoryPath As String In Directory.GetDirectories(modelRoot).OrderBy(Function(p) Path.GetFileName(p), StringComparer.OrdinalIgnoreCase)
                Dim name As String = Path.GetFileName(directoryPath)

                If name.StartsWith("vosk-model", StringComparison.OrdinalIgnoreCase) Then
                    result.Add(New LiveEngineDescriptor With {
                        .DisplayName = "Vosk: " & name,
                        .Kind = EngineKind.Vosk,
                        .ModelOrTag = name,
                        .Languages = New List(Of String) From {"(language comes from selected Vosk model)"}
                    })
                End If
            Next

            For Each filePath As String In Directory.GetFiles(modelRoot, "ggml*").OrderBy(Function(p) Path.GetFileName(p), StringComparer.OrdinalIgnoreCase)
                Dim name As String = Path.GetFileName(filePath)

                result.Add(New LiveEngineDescriptor With {
                    .DisplayName = "Whisper: " & name,
                    .Kind = EngineKind.WhisperLocal,
                    .ModelOrTag = name,
                    .Languages = WhisperEngine.SupportedLanguages.OrderBy(Function(x) If(String.Equals(x, "auto", StringComparison.OrdinalIgnoreCase), "", x), StringComparer.OrdinalIgnoreCase).ToList()
                })
            Next
        End If

        If HasConfiguredGoogleV1Provider() Then
            result.Add(New LiveEngineDescriptor With {
                .DisplayName = GoogleV1Engine.DisplayName,
                .Kind = EngineKind.GoogleV1,
                .ModelOrTag = "google-v1",
                .Languages = GoogleV1Engine.SupportedLanguages.OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase).ToList()
            })
        End If

        If HasConfiguredGoogleV2Provider() Then
            result.Add(New LiveEngineDescriptor With {
                .DisplayName = GoogleV2Engine.DisplayName,
                .Kind = EngineKind.GoogleV2,
                .ModelOrTag = "google-v2",
                .Languages = GoogleV2Engine.SupportedLanguages.OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase).ToList()
            })
        End If

        If HasConfiguredOpenAiProvider() Then
            result.Add(New LiveEngineDescriptor With {
                .DisplayName = OpenAiRealtimeEngine.DisplayNameValue,
                .Kind = EngineKind.OpenAiRealtime,
                .ModelOrTag = "gpt-realtime-whisper",
                .Languages = OpenAiRealtimeEngine.SupportedLanguages.OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase).ToList()
            })
        End If

        If HasConfiguredAzureProvider() Then
            result.Add(New LiveEngineDescriptor With {
                .DisplayName = AzureSpeechRealtimeEngine.DisplayNameValue,
                .Kind = EngineKind.AzureSpeechRealtime,
                .ModelOrTag = "azure-speech-realtime",
                .Languages = AzureSpeechRealtimeEngine.SupportedLanguages.OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase).ToList()
            })
        End If

        Return result
    End Function

    Private Sub LoadAlternateProviderFallbacks()
        _alternateOpenAiConfig = Nothing
        _alternateGoogleConfig = Nothing

        Try
            Dim altPath As String = SharedMethods.ExpandEnvironmentVariables(_owner.INI_AlternateModelPath)

            If String.IsNullOrWhiteSpace(altPath) OrElse Not File.Exists(altPath) Then
                Return
            End If

            Dim models As List(Of ModelConfig) = SharedMethods.LoadAlternativeModels(
                altPath,
                _owner._context,
                "TalkToMe Speech Fallback",
                includeToolOnly:=True,
                toolsOnly:=False)

            If models Is Nothing OrElse models.Count = 0 Then
                Return
            End If

            _alternateOpenAiConfig = models.FirstOrDefault(Function(m) IsUsableOpenAiConfig(m))
            _alternateGoogleConfig = models.FirstOrDefault(Function(m) IsUsableGoogleConfig(m))
        Catch
        End Try
    End Sub

    Private Shared Function EndpointMatchesProvider(endpoint As String, providerIdentifier As String) As Boolean
        If String.IsNullOrWhiteSpace(endpoint) OrElse String.IsNullOrWhiteSpace(providerIdentifier) Then
            Return False
        End If

        Return endpoint.IndexOf(providerIdentifier, StringComparison.OrdinalIgnoreCase) >= 0
    End Function

    Private Shared Function IsAzureSpeechLocation(value As String) As Boolean
        Dim normalized As String = If(value, "").Trim()

        If String.IsNullOrWhiteSpace(normalized) Then
            Return False
        End If

        If normalized.IndexOf(".cognitiveservices.azure.com", StringComparison.OrdinalIgnoreCase) >= 0 Then
            Return True
        End If

        If normalized.IndexOf(".api.cognitive.microsoft.com", StringComparison.OrdinalIgnoreCase) >= 0 Then
            Return True
        End If

        If normalized.StartsWith("http://", StringComparison.OrdinalIgnoreCase) OrElse
           normalized.StartsWith("https://", StringComparison.OrdinalIgnoreCase) Then
            Return False
        End If

        For Each ch As Char In normalized
            If Not Char.IsLetterOrDigit(ch) AndAlso ch <> "-"c Then
                Return False
            End If
        Next

        Return True
    End Function

    Private Function BuildConfiguredGoogleModelConfig(useSecond As Boolean) As ModelConfig
        Dim endpoint As String = If(useSecond, _owner.INI_Endpoint_2, _owner.INI_Endpoint)
        Dim oauthEnabled As Boolean = If(useSecond, _owner.INI_OAuth2_2, _owner.INI_OAuth2)

        If Not EndpointMatchesProvider(endpoint, ThisAddIn.GoogleIdentifier) OrElse Not oauthEnabled Then
            Return Nothing
        End If

        Return New ModelConfig With {
            .Endpoint = endpoint,
            .OAuth2 = oauthEnabled,
            .OAuth2ClientMail = If(useSecond, _owner.INI_OAuth2ClientMail_2, _owner.INI_OAuth2ClientMail),
            .OAuth2Scopes = If(useSecond, _owner.INI_OAuth2Scopes_2, _owner.INI_OAuth2Scopes),
            .OAuth2Endpoint = If(useSecond, _owner.INI_OAuth2Endpoint_2, _owner.INI_OAuth2Endpoint),
            .OAuth2ATExpiry = If(useSecond, _owner.INI_OAuth2ATExpiry_2, _owner.INI_OAuth2ATExpiry),
            .APIKey = If(useSecond, _owner.INI_APIKey_2, _owner.INI_APIKey)
        }
    End Function

    Private Function BuildConfiguredOpenAiModelConfig(useSecond As Boolean) As ModelConfig
        Dim endpoint As String = If(useSecond, _owner.INI_Endpoint_2, _owner.INI_Endpoint)

        If Not EndpointMatchesProvider(endpoint, ThisAddIn.OpenAIIdentifier) Then
            Return Nothing
        End If

        Return New ModelConfig With {
            .Endpoint = endpoint,
            .APIKey = If(useSecond, _owner.INI_APIKey_2, _owner.INI_APIKey),
            .DecodedAPI = If(useSecond, _owner.DecodedAPI_2, _owner.DecodedAPI)
        }
    End Function

    Private Function GetApiKeyFromModelConfig(config As ModelConfig) As String
        If config Is Nothing Then
            Return ""
        End If

        If Not String.IsNullOrWhiteSpace(config.DecodedAPI) Then
            Return config.DecodedAPI.Trim()
        End If

        Return If(config.APIKey, "").Trim()
    End Function

    Private Function IsUsableOpenAiConfig(config As ModelConfig) As Boolean
        If config Is Nothing Then
            Return False
        End If

        If Not EndpointMatchesProvider(config.Endpoint, ThisAddIn.OpenAIIdentifier) Then
            Return False
        End If

        Return Not String.IsNullOrWhiteSpace(GetApiKeyFromModelConfig(config))
    End Function

    Private Function IsUsableGoogleConfig(config As ModelConfig) As Boolean
        If config Is Nothing Then
            Return False
        End If

        If Not EndpointMatchesProvider(config.Endpoint, ThisAddIn.GoogleIdentifier) Then
            Return False
        End If

        If Not config.OAuth2 Then
            Return False
        End If

        If String.IsNullOrWhiteSpace(config.OAuth2ClientMail) Then
            Return False
        End If

        If String.IsNullOrWhiteSpace(config.OAuth2Endpoint) Then
            Return False
        End If

        Return Not String.IsNullOrWhiteSpace(config.APIKey)
    End Function

    Private Function ResolveGoogleTranscriptionConfig(ByRef cacheSlot As String) As ModelConfig
        Dim primaryConfig As ModelConfig = BuildConfiguredGoogleModelConfig(False)
        If IsUsableGoogleConfig(primaryConfig) Then
            cacheSlot = "primary"
            Return primaryConfig
        End If

        Dim secondaryConfig As ModelConfig = BuildConfiguredGoogleModelConfig(True)
        If IsUsableGoogleConfig(secondaryConfig) Then
            cacheSlot = "secondary"
            Return secondaryConfig
        End If

        If IsUsableGoogleConfig(_alternateGoogleConfig) Then
            cacheSlot = "alternate"
            Return _alternateGoogleConfig
        End If

        cacheSlot = ""
        Return Nothing
    End Function

    Private Function ResolveOpenAiConfig() As ModelConfig
        Dim primaryConfig As ModelConfig = BuildConfiguredOpenAiModelConfig(False)
        If IsUsableOpenAiConfig(primaryConfig) Then
            Return primaryConfig
        End If

        Dim secondaryConfig As ModelConfig = BuildConfiguredOpenAiModelConfig(True)
        If IsUsableOpenAiConfig(secondaryConfig) Then
            Return secondaryConfig
        End If

        If IsUsableOpenAiConfig(_alternateOpenAiConfig) Then
            Return _alternateOpenAiConfig
        End If

        Return Nothing
    End Function

    Private Function ResolveOpenAiKey() As String
        Dim config As ModelConfig = ResolveOpenAiConfig()
        Return GetApiKeyFromModelConfig(config)
    End Function

    Private Function HasConfiguredGoogleV1Provider() As Boolean
        Return (EndpointMatchesProvider(_owner.INI_Endpoint, ThisAddIn.GoogleIdentifier) AndAlso _owner.INI_OAuth2) OrElse
               (EndpointMatchesProvider(_owner.INI_Endpoint_2, ThisAddIn.GoogleIdentifier) AndAlso _owner.INI_OAuth2_2) OrElse
               IsUsableGoogleConfig(_alternateGoogleConfig)
    End Function

    Private Function HasConfiguredGoogleV2Provider() As Boolean
        Return HasConfiguredGoogleV1Provider() AndAlso
               Not String.IsNullOrWhiteSpace(ResolveGoogleProjectId())
    End Function

    Private Function HasConfiguredOpenAiProvider() As Boolean
        Return EndpointMatchesProvider(_owner.INI_Endpoint, ThisAddIn.OpenAIIdentifier) OrElse
               EndpointMatchesProvider(_owner.INI_Endpoint_2, ThisAddIn.OpenAIIdentifier) OrElse
               IsUsableOpenAiConfig(_alternateOpenAiConfig)
    End Function

    Private Function HasConfiguredAzureProvider() As Boolean
        Return Not String.IsNullOrWhiteSpace(ResolveAzureSpeechKey())
    End Function

    Private Function ResolveGoogleProjectId() As String
        Return NormalizeIniValue(If(_owner.INI_STT_Google_ProjectID, ""))
    End Function

    Private Function ResolveGoogleSttSetting(modelOrTag As String, settingName As String, defaultValue As String) As String
        Return ResolveSttSetting(_owner.INI_STT_Google, modelOrTag, settingName, defaultValue)
    End Function

    Private Function ResolveOpenAiSttSetting(modelOrTag As String, settingName As String, defaultValue As String) As String
        Return ResolveSttSetting(_owner.INI_STT_OpenAI, modelOrTag, settingName, defaultValue)
    End Function

    Private Function ResolveAzureSpeechKey() As String
        Return DecodeWrappedEncryptedValue(NormalizeIniValue(_owner.INI_STT_Azure_SpeechKey), "Azure Speech key")
    End Function

    Private Shared Function DecodeWrappedEncryptedValue(value As String, valueName As String) As String
        Dim normalized As String = NormalizeIniValue(value)

        If normalized.StartsWith("encrypted(", StringComparison.OrdinalIgnoreCase) AndAlso
           normalized.EndsWith(")", StringComparison.Ordinal) Then

            Dim innerValue As String = normalized.Substring(
                "encrypted(".Length,
                normalized.Length - "encrypted(".Length - 1).Trim()

            If String.IsNullOrWhiteSpace(innerValue) Then
                Return ""
            End If

            Dim codeBasis As String = ResolveCodeBasis()
            If String.IsNullOrWhiteSpace(codeBasis) Then
                Throw New InvalidOperationException("Missing CodeBasis for encrypted " & valueName & ".")
            End If

            Dim decoded As String = SharedMethods.DecodeString(innerValue, codeBasis)
            If decoded.StartsWith("Error:", StringComparison.OrdinalIgnoreCase) Then
                Throw New InvalidOperationException("Failed to decrypt " & valueName & ": " & decoded)
            End If

            Return NormalizeIniValue(decoded)
        End If

        Return normalized
    End Function

    Private Shared Function ResolveCodeBasis() As String
        Dim codeBasis As String = ""

        Try
            If ThisAddIn._context IsNot Nothing Then
                codeBasis = If(ThisAddIn._context.Codebasis, "").Trim()
            End If
        Catch
        End Try

        If String.IsNullOrWhiteSpace(codeBasis) Then
            Try
                codeBasis = SharedMethods.GetFromRegistry(SharedMethods.RegPath_Base, SharedMethods.RegPath_CodeBasis, True)
            Catch
            End Try
        End If

        Try
            If ThisAddIn._context IsNot Nothing AndAlso
               String.IsNullOrWhiteSpace(ThisAddIn._context.Codebasis) AndAlso
               Not String.IsNullOrWhiteSpace(codeBasis) Then

                ThisAddIn._context.Codebasis = codeBasis
            End If
        Catch
        End Try

        Return If(codeBasis, "").Trim()
    End Function

    Private Function ResolveAzureRegionForHeader(modelOrTag As String) As String
        Return ResolveAzureSttSetting(modelOrTag, "region", "")
    End Function

    Private Function ResolveAzureRealtimeLocation(modelOrTag As String) As String
        Dim region As String = ResolveAzureRegionForHeader(modelOrTag)
        If Not String.IsNullOrWhiteSpace(region) Then
            Return region
        End If

        Return ResolveAzureSttSetting(modelOrTag, "endpoint", "")
    End Function

    Private Function ResolveAzureSttSetting(modelOrTag As String, settingName As String, defaultValue As String) As String
        Return NormalizeIniValue(ResolveSttSetting(_owner.INI_STT_Azure, modelOrTag, settingName, defaultValue))
    End Function

    Private Function ResolveSttSetting(raw As String, modelOrTag As String, settingName As String, defaultValue As String) As String
        Dim settings As Dictionary(Of String, String) = ParseSttSettings(raw)
        Dim value As String = ""

        If Not String.IsNullOrWhiteSpace(modelOrTag) AndAlso
           settings.TryGetValue(modelOrTag & "." & settingName, value) Then
            Return value
        End If

        If settings.TryGetValue("default." & settingName, value) Then
            Return value
        End If

        If settings.TryGetValue(settingName, value) Then
            Return value
        End If

        Return defaultValue
    End Function

    Private Shared Function ParseSttSettings(raw As String) As Dictionary(Of String, String)
        Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        Dim normalized As String = If(raw, "")
        normalized = normalized.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
        normalized = normalized.Replace(";"c, vbLf(0))

        For Each line As String In normalized.Split(New Char() {vbLf(0)}, StringSplitOptions.RemoveEmptyEntries)
            Dim trimmedLine As String = line.Trim()

            If String.IsNullOrWhiteSpace(trimmedLine) Then
                Continue For
            End If

            If trimmedLine.StartsWith("#", StringComparison.OrdinalIgnoreCase) OrElse
               trimmedLine.StartsWith(";", StringComparison.OrdinalIgnoreCase) Then
                Continue For
            End If

            Dim separatorIndex As Integer = trimmedLine.IndexOf("="c)
            If separatorIndex <= 0 Then
                Continue For
            End If

            Dim key As String = trimmedLine.Substring(0, separatorIndex).Trim()
            Dim value As String = trimmedLine.Substring(separatorIndex + 1).Trim()

            Dim hashIndex As Integer = value.IndexOf("#"c)
            If hashIndex >= 0 Then
                value = value.Substring(0, hashIndex).Trim()
            End If

            If Not String.IsNullOrWhiteSpace(key) Then
                result(key) = value
            End If
        Next

        Return result
    End Function


    Private Shared Function NormalizeIniValue(value As String) As String
        Dim result As String = If(value, "").Trim()

        If result.Length >= 2 AndAlso
           result.StartsWith("""", StringComparison.Ordinal) AndAlso
           result.EndsWith("""", StringComparison.Ordinal) Then

            result = result.Substring(1, result.Length - 2).Trim()
        End If

        Return SharedMethods.ExpandEnvironmentVariables(result).Trim()
    End Function

    Private Async Function GetFreshGoogleTokenAsync(config As ModelConfig, cacheSlot As String) As Task(Of String)
        If config Is Nothing Then
            Return ""
        End If

        Dim token As String = ""
        Dim exp As DateTime = DateTime.MinValue

        Select Case cacheSlot
            Case "secondary"
                token = _g2Token
                exp = _g2Exp

            Case "alternate"
                token = _gAltToken
                exp = _gAltExp

            Case Else
                token = _g1Token
                exp = _g1Exp
        End Select

        If String.IsNullOrEmpty(token) OrElse DateTime.UtcNow >= exp Then
            Dim life As Long = If(config.OAuth2ATExpiry > 0, config.OAuth2ATExpiry, 3600)

            SharedMethods.GoogleOAuthHelper.client_email = config.OAuth2ClientMail
            SharedMethods.GoogleOAuthHelper.private_key = FormatPrivateKey(config.APIKey)
            SharedMethods.GoogleOAuthHelper.scopes = config.OAuth2Scopes
            SharedMethods.GoogleOAuthHelper.token_uri = config.OAuth2Endpoint
            SharedMethods.GoogleOAuthHelper.token_life = life

            token = Await SharedMethods.GoogleOAuthHelper.GetAccessToken()
            Dim newExp As DateTime = DateTime.UtcNow.AddSeconds(Math.Max(60L, life - 300L))

            Select Case cacheSlot
                Case "secondary"
                    _g2Token = token
                    _g2Exp = newExp

                Case "alternate"
                    _gAltToken = token
                    _gAltExp = newExp

                Case Else
                    _g1Token = token
                    _g1Exp = newExp
            End Select
        End If

        Return token
    End Function

    Public Shared Function FormatPrivateKey(rawKey As String) As String
        Dim noEsc As String = rawKey.Replace("\n", "")
        Dim sb As New StringBuilder()

        For i As Integer = 0 To noEsc.Length - 1 Step 64
            sb.AppendLine(If(i + 64 <= noEsc.Length, noEsc.Substring(i, 64), noEsc.Substring(i)))
        Next

        Return "-----BEGIN PRIVATE KEY-----" & vbLf & sb.ToString() & "-----END PRIVATE KEY-----" & vbLf
    End Function

    Private Async Function CreateEngineAsync(d As LiveEngineDescriptor) As Task(Of ITranscriptionEngine)
        Dim modelRoot As String = SharedMethods.ExpandEnvironmentVariables(_owner.INI_SpeechModelPath)

        Select Case d.Kind
            Case EngineKind.Vosk
                Return New VoskEngine(modelRoot, d.ModelOrTag)

            Case EngineKind.WhisperLocal
                Return New WhisperEngine(modelRoot, d.ModelOrTag)

            Case EngineKind.GoogleV1
                Dim googleCacheSlot As String = ""
                Dim googleConfig As ModelConfig = ResolveGoogleTranscriptionConfig(googleCacheSlot)

                If googleConfig Is Nothing Then
                    Throw New InvalidOperationException("No Google transcription credentials are available.")
                End If

                _opts.Model = ResolveGoogleSttSetting(d.ModelOrTag, "model", "")

                Dim tokenFactory As Func(Of Task(Of String)) =
                    Function() GetFreshGoogleTokenAsync(googleConfig, googleCacheSlot)

                Return New GoogleV1Engine("", tokenFactory)

            Case EngineKind.GoogleV2
                Dim googleCacheSlot As String = ""
                Dim googleConfig As ModelConfig = ResolveGoogleTranscriptionConfig(googleCacheSlot)

                If googleConfig Is Nothing Then
                    Throw New InvalidOperationException("No Google transcription credentials are available.")
                End If

                If String.IsNullOrWhiteSpace(ResolveGoogleProjectId()) Then
                    Throw New InvalidOperationException("INI_STT_Google_ProjectID is missing.")
                End If

                Dim dbgProjectIdRaw As String = If(_owner.INI_STT_Google_ProjectID, "")
                Dim dbgProjectIdResolved As String = ResolveGoogleProjectId()
                Dim dbgIniSttGoogleRaw As String =
                    If(_owner.INI_STT_Google, "").
                        Replace(vbCrLf, "\n").
                        Replace(vbCr, "\n").
                        Replace(vbLf, "\n")
                Dim dbgResolvedEndpoint As String = ResolveGoogleSttSetting(d.ModelOrTag, "endpoint", "")
                Dim dbgResolvedLocation As String = ResolveGoogleSttSetting(d.ModelOrTag, "location", "")
                Dim dbgResolvedRecognizer As String = ResolveGoogleSttSetting(d.ModelOrTag, "recognizer", "")
                Dim dbgResolvedModel As String = ResolveGoogleSttSetting(d.ModelOrTag, "model", "")
                Dim dbgResolvedLanguage As String = ResolveGoogleSttSetting(d.ModelOrTag, "language", "")

                System.Diagnostics.Debug.WriteLine("[TalkToMe.GoogleV2] _context Is Nothing=" & (ThisAddIn._context Is Nothing).ToString())
                System.Diagnostics.Debug.WriteLine("[TalkToMe.GoogleV2] Codebasis present=" & (ThisAddIn._context IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(ThisAddIn._context.Codebasis)).ToString())
                System.Diagnostics.Debug.WriteLine("[TalkToMe.GoogleV2] CacheSlot=" & googleCacheSlot)
                System.Diagnostics.Debug.WriteLine("[TalkToMe.GoogleV2] OAuthClientMail=" & If(googleConfig.OAuth2ClientMail, ""))
                System.Diagnostics.Debug.WriteLine("[TalkToMe.GoogleV2] OAuthTokenEndpoint=" & If(googleConfig.OAuth2Endpoint, ""))
                System.Diagnostics.Debug.WriteLine("[TalkToMe.GoogleV2] APIKeyLength=" & If(googleConfig.APIKey, "").Length.ToString())
                System.Diagnostics.Debug.WriteLine("[TalkToMe.GoogleV2] INI_STT_Google_ProjectID raw=" & dbgProjectIdRaw)
                System.Diagnostics.Debug.WriteLine("[TalkToMe.GoogleV2] ResolveGoogleProjectId()=" & dbgProjectIdResolved)
                System.Diagnostics.Debug.WriteLine("[TalkToMe.GoogleV2] INI_STT_Google raw=" & dbgIniSttGoogleRaw)
                System.Diagnostics.Debug.WriteLine(
                    "[TalkToMe.GoogleV2] Resolved STT settings " &
                    "endpoint=" & dbgResolvedEndpoint &
                    "; location=" & dbgResolvedLocation &
                    "; recognizer=" & dbgResolvedRecognizer &
                    "; model=" & dbgResolvedModel &
                    "; language=" & dbgResolvedLanguage)



                Return New GoogleV2Engine(
                        googleConfig.OAuth2ClientMail,
                        googleConfig.APIKey,
                        googleConfig.OAuth2Endpoint,
                        ResolveGoogleProjectId(),
                        ResolveGoogleSttSetting(d.ModelOrTag, "endpoint", ""),
                        ResolveGoogleSttSetting(d.ModelOrTag, "location", ""),
                        ResolveGoogleSttSetting(d.ModelOrTag, "recognizer", ""),
                        ResolveGoogleSttSetting(d.ModelOrTag, "model", ""),
                        ResolveGoogleSttSetting(d.ModelOrTag, "language", ""),
                        googleConfig.OAuth2Scopes,
                        "TalkToMe")

            Case EngineKind.OpenAiRealtime
                Dim key As String = ResolveOpenAiKey()
                If String.IsNullOrWhiteSpace(key) Then
                    Throw New InvalidOperationException("No OpenAI API key is available.")
                End If

                _opts.Model = ResolveOpenAiSttSetting(d.ModelOrTag, "model", "gpt-realtime-whisper")
                Return New OpenAiRealtimeEngine(key)

            Case EngineKind.AzureSpeechRealtime
                Dim azureSpeechKey As String = ResolveAzureSpeechKey()
                If String.IsNullOrWhiteSpace(azureSpeechKey) Then
                    Throw New InvalidOperationException("INI_STT_Azure_SpeechKey is missing.")
                End If

                Return New AzureSpeechRealtimeEngine(
                    azureSpeechKey,
                    ResolveAzureRealtimeLocation(d.ModelOrTag),
                    ResolveAzureRegionForHeader(d.ModelOrTag))

            Case Else
                Throw New NotSupportedException(d.Kind.ToString())
        End Select
    End Function

    Private Sub AttachEngineEvents(eng As ITranscriptionEngine)
        AddHandler eng.PartialResult,
            Sub(sender As Object, ev As TranscriptionEventArgs)
                Dim msg As String = If(String.IsNullOrEmpty(ev.Speaker), ev.Text, ev.Speaker & ": " & ev.Text)

                If Not String.IsNullOrWhiteSpace(msg) Then
                    RaiseEvent PartialTranscriptReceived(Me, New SharedLibrary.SharedLibrary.TalkToMeTranscriptEventArgs(msg.Trim()))
                End If
            End Sub

        AddHandler eng.FinalResult,
            Sub(sender As Object, ev As TranscriptionEventArgs)
                Dim msg As String = If(String.IsNullOrEmpty(ev.Speaker), ev.Text, ev.Speaker & ": " & ev.Text)

                RaiseEvent FinalTranscriptReceived(Me, New SharedLibrary.SharedLibrary.TalkToMeTranscriptEventArgs(If(msg, "").Trim()))
            End Sub

        AddHandler eng.EngineError,
            Sub(sender As Object, ev As TranscriptionErrorEventArgs)
                RaiseEvent FinalTranscriptReceived(
                    Me,
                    New SharedLibrary.SharedLibrary.TalkToMeTranscriptEventArgs("Error: " & ev.Message))

                ScheduleListeningRestart("Engine error: " & ev.Message)
            End Sub

        AddHandler eng.Status,
            Sub(sender As Object, ev As TranscriptionStatusEventArgs)
                Dim msg As String = If(ev.Message, "").Trim()
                Dim normalizedMsg As String = msg.ToLowerInvariant()

                If Not _isListeningValue Then
                    Return
                End If

                If normalizedMsg.IndexOf("turn ended", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                   normalizedMsg.IndexOf("timeout", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                   normalizedMsg.IndexOf("timed out", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                   normalizedMsg.IndexOf("session ended", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                   normalizedMsg.IndexOf("connection closed", StringComparison.OrdinalIgnoreCase) >= 0 Then

                    ScheduleListeningRestart(msg)
                End If
            End Sub
    End Sub

    Private Sub ScheduleListeningRestart(reason As String)
        If Not _isListeningValue Then
            Return
        End If

        Task.Run(
            Async Function()
                Await RestartListeningSessionAsync(reason).ConfigureAwait(False)
            End Function)
    End Sub

    Private Async Function RestartListeningSessionAsync(reason As String) As Task
        If Threading.Interlocked.Exchange(_sessionRestartInProgress, 1) <> 0 Then
            Return
        End If

        Try
            If Not _isListeningValue Then
                Return
            End If

            Dim d As LiveEngineDescriptor = _selectedDescriptor
            If d Is Nothing Then
                Return
            End If

            Dim oldEngine As ITranscriptionEngine = _engine
            Dim oldCts As CancellationTokenSource = _cts

            RaiseEvent FinalTranscriptReceived(
                Me,
                New SharedLibrary.SharedLibrary.TalkToMeTranscriptEventArgs("Restarting listening session…"))

            Dim newEngine As ITranscriptionEngine = Await CreateEngineAsync(d).ConfigureAwait(False)
            AttachEngineEvents(newEngine)

            Dim newCts As New CancellationTokenSource()
            Await newEngine.StartLiveAsync(_opts, newCts.Token).ConfigureAwait(False)

            _engine = newEngine
            _cts = newCts

            If oldEngine IsNot Nothing Then
                Try
                    Await oldEngine.StopLiveAsync().ConfigureAwait(False)
                Catch
                End Try

                Try
                    Await oldEngine.DisposeAsync().ConfigureAwait(False)
                Catch
                End Try
            End If

            If oldCts IsNot Nothing Then
                Try
                    oldCts.Cancel()
                Catch
                End Try

                Try
                    oldCts.Dispose()
                Catch
                End Try
            End If
        Catch ex As Exception
            RaiseEvent FinalTranscriptReceived(
                Me,
                New SharedLibrary.SharedLibrary.TalkToMeTranscriptEventArgs("Error: " & ex.Message))
        Finally
            Threading.Interlocked.Exchange(_sessionRestartInProgress, 0)
        End Try
    End Function

    Private Async Sub OnCaptureFrame(sender As Object, e As AudioCaptureService.FrameEventArgs)
        Dim eng As ITranscriptionEngine = _engine
        Dim ctsLocal As CancellationTokenSource = _cts

        If eng Is Nothing OrElse Not _isListeningValue OrElse ctsLocal Is Nothing OrElse IsSpeechOutputActive Then
            Return
        End If

        Try
            Await eng.PushAudioAsync(e.Pcm, e.BytesValid, ctsLocal.Token)
        Catch ex As OperationCanceledException
        Catch ex As ObjectDisposedException
        Catch ex As Exception
            RaiseEvent FinalTranscriptReceived(
                Me,
                New SharedLibrary.SharedLibrary.TalkToMeTranscriptEventArgs("Error: Audio push failed: " & ex.Message))
        End Try
    End Sub

    Private NotInheritable Class TalkToMeConfigForm
        Inherits Form

        Private ReadOnly _ownerAdapter As WordTalkToMeSpeechAdapter
        Private ReadOnly _descriptors As List(Of LiveEngineDescriptor)
        Private ReadOnly _savedLanguageByEngine As Dictionary(Of String, String)

        Private NotInheritable Class AudioInputDeviceChoice
            Public Property DeviceId As String = ""
            Public Property WaveDeviceIndex As Integer = 0
            Public Property DisplayText As String = ""

            Public Overrides Function ToString() As String
                Return DisplayText
            End Function
        End Class

        Private ReadOnly _audioInputDevices As New List(Of AudioInputDeviceChoice)()

        Private cboEngine As ComboBox
        Private cboLanguage As ComboBox
        Private cboMicrophone As ComboBox
        Private chkIncludeDocument As CheckBox
        Private chkSpeechOutputEnabled As CheckBox
        Private cboSpeechOutputMode As ComboBox
        Private chkUseSpeakerSpecificVoices As CheckBox
        Private cboSpeechOutputSpeed As ComboBox
        Private btnConfigureSpeechOutput As Button
        Private lblSpeechOutputSummary As Label
        Private btnOk As Button
        Private btnCancel As Button

        Public Property SelectedDescriptor As LiveEngineDescriptor = Nothing
        Public Property SelectedLanguage As String = "auto"
        Public Property SelectedMicrophoneDeviceIndex As Integer = 0
        Public Property IncludeFullDocument As Boolean
        Public Property SelectedSpeechOutputEnabled As Boolean
        Public Property SelectedSpeechOutputMode As String = "Queue (progressive)"
        Public Property SelectedUseSpeakerSpecificVoices As Boolean
        Public Property SelectedSpeechOutputSpeed As Double = 1.0R

        Public Sub New(ownerAdapter As WordTalkToMeSpeechAdapter,
                       descriptors As List(Of LiveEngineDescriptor),
                       currentEngineDisplayName As String,
                       currentLanguage As String,
                       currentMicrophoneDeviceIndex As Integer,
                       currentIncludeFullDocument As Boolean,
                       savedLanguageByEngine As Dictionary(Of String, String),
                       currentSpeechOutputEnabled As Boolean,
                       currentSpeechOutputMode As String,
                       currentUseSpeakerSpecificVoices As Boolean,
                       currentSpeechOutputSpeed As Double)
            _ownerAdapter = ownerAdapter
            _descriptors = descriptors
            _savedLanguageByEngine = If(savedLanguageByEngine, New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase))
            IncludeFullDocument = currentIncludeFullDocument
            SelectedLanguage = currentLanguage
            SelectedSpeechOutputEnabled = currentSpeechOutputEnabled
            SelectedSpeechOutputMode = NormalizeSpeechOutputMode(currentSpeechOutputMode)
            SelectedUseSpeakerSpecificVoices = currentUseSpeakerSpecificVoices
            SelectedSpeechOutputSpeed = NormalizeSpeechOutputSpeed(currentSpeechOutputSpeed)

            Me.Text = SharedMethods.AN & " - Talk to me!"
            Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0F, 96.0F)
            Me.AutoScaleMode = AutoScaleMode.Dpi
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.StartPosition = FormStartPosition.CenterParent
            Me.MinimizeBox = False
            Me.MaximizeBox = False
            Me.ShowInTaskbar = False
            Me.AutoScroll = True
            Me.ClientSize = New System.Drawing.Size(640, 360)
            Me.Font = New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)

            Dim root As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 10,
                .Padding = New Padding(10)
            }

            root.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))

            cboEngine = New ComboBox() With {
                .Dock = DockStyle.Fill,
                .DropDownStyle = ComboBoxStyle.DropDownList
            }

            cboLanguage = New ComboBox() With {
                .Dock = DockStyle.Fill,
                .DropDownStyle = ComboBoxStyle.DropDownList
            }

            cboMicrophone = New ComboBox() With {
                .Dock = DockStyle.Fill,
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .DropDownWidth = 700
            }

            chkIncludeDocument = New CheckBox() With {
                .Text = "Allow the LLM to see the current document",
                .Checked = currentIncludeFullDocument,
                .AutoSize = True,
                .Margin = New Padding(0, 5, 0, 0)
            }

            chkSpeechOutputEnabled = New CheckBox() With {
                .Text = "Enable speech output for incoming text while listening",
                .Checked = currentSpeechOutputEnabled,
                .AutoSize = True,
                .Margin = New Padding(0, 5, 0, 0)
            }

            cboSpeechOutputMode = New ComboBox() With {
                .Dock = DockStyle.Fill,
                .DropDownStyle = ComboBoxStyle.DropDownList
            }
            cboSpeechOutputMode.Items.Add("Queue")
            cboSpeechOutputMode.Items.Add("Queue (progressive)")
            cboSpeechOutputMode.Items.Add("Interrupt current speech")
            cboSpeechOutputMode.Items.Add("Skip new output while speaking")

            chkUseSpeakerSpecificVoices = New CheckBox() With {
                .Text = "Use alternating voices when different bots speak",
                .Checked = currentUseSpeakerSpecificVoices,
                .AutoSize = True,
                .Margin = New Padding(0, 5, 0, 0)
            }

            cboSpeechOutputSpeed = New ComboBox() With {
                .Dock = DockStyle.Fill,
                .DropDownStyle = ComboBoxStyle.DropDownList
            }
            cboSpeechOutputSpeed.Items.Add("1.00x (Normal)")
            cboSpeechOutputSpeed.Items.Add("1.10x")
            cboSpeechOutputSpeed.Items.Add("1.25x")
            cboSpeechOutputSpeed.Items.Add("1.50x")
            cboSpeechOutputSpeed.Items.Add("1.75x")
            cboSpeechOutputSpeed.Items.Add("2.00x")

            btnConfigureSpeechOutput = New Button() With {
                .Text = "Configure Speech Output...",
                .AutoSize = True
            }

            lblSpeechOutputSummary = New Label() With {
                .AutoSize = True,
                .MaximumSize = New System.Drawing.Size(420, 0),
                .Margin = New Padding(8, 6, 0, 0)
            }

            btnOk = New Button() With {
                .Text = "OK",
                .DialogResult = DialogResult.OK,
                .AutoSize = True
            }

            btnCancel = New Button() With {
                .Text = "Cancel",
                .DialogResult = DialogResult.Cancel,
                .AutoSize = True
            }

            Dim speechOutputPanel As New FlowLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.LeftToRight,
                .AutoSize = True,
                .WrapContents = True
            }
            speechOutputPanel.Controls.Add(btnConfigureSpeechOutput)
            speechOutputPanel.Controls.Add(lblSpeechOutputSummary)

            Dim buttons As New FlowLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.RightToLeft,
                .WrapContents = False,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink
            }
            buttons.Controls.Add(btnCancel)
            buttons.Controls.Add(btnOk)

            root.Controls.Add(New Label() With {.Text = "Engine:", .AutoSize = True, .Anchor = AnchorStyles.Left}, 0, 0)
            root.Controls.Add(cboEngine, 1, 0)

            root.Controls.Add(New Label() With {.Text = "Language:", .AutoSize = True, .Anchor = AnchorStyles.Left}, 0, 1)
            root.Controls.Add(cboLanguage, 1, 1)

            root.Controls.Add(New Label() With {.Text = "Microphone:", .AutoSize = True, .Anchor = AnchorStyles.Left}, 0, 2)
            root.Controls.Add(cboMicrophone, 1, 2)

            root.Controls.Add(New Label() With {.Text = "Document:", .AutoSize = True, .Anchor = AnchorStyles.Left}, 0, 3)
            root.Controls.Add(chkIncludeDocument, 1, 3)

            root.Controls.Add(New Label() With {.Text = "Speech output:", .AutoSize = True, .Anchor = AnchorStyles.Left}, 0, 4)
            root.Controls.Add(chkSpeechOutputEnabled, 1, 4)

            root.Controls.Add(New Label() With {.Text = "Speech mode:", .AutoSize = True, .Anchor = AnchorStyles.Left}, 0, 5)
            root.Controls.Add(cboSpeechOutputMode, 1, 5)

            root.Controls.Add(New Label() With {.Text = "Assistant voices:", .AutoSize = True, .Anchor = AnchorStyles.Left}, 0, 6)
            root.Controls.Add(chkUseSpeakerSpecificVoices, 1, 6)

            root.Controls.Add(New Label() With {.Text = "Speech speed:", .AutoSize = True, .Anchor = AnchorStyles.Left}, 0, 7)
            root.Controls.Add(cboSpeechOutputSpeed, 1, 7)

            root.Controls.Add(New Label() With {.Text = "Voice:", .AutoSize = True, .Anchor = AnchorStyles.Left}, 0, 8)
            root.Controls.Add(speechOutputPanel, 1, 8)

            root.Controls.Add(buttons, 1, 9)

            Me.Controls.Add(root)
            Me.AcceptButton = btnOk
            Me.CancelButton = btnCancel

            root.RowStyles.Clear()
            For i As Integer = 0 To root.RowCount - 1
                root.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            Next

            For Each descriptor As LiveEngineDescriptor In _descriptors
                cboEngine.Items.Add(descriptor.DisplayName)
            Next

            LoadAudioInputDevices(currentMicrophoneDeviceIndex)

            AddHandler cboEngine.SelectedIndexChanged, AddressOf OnEngineChanged
            AddHandler btnConfigureSpeechOutput.Click, AddressOf OnConfigureSpeechOutputClick
            AddHandler chkUseSpeakerSpecificVoices.CheckedChanged, Sub(sender As Object, e As EventArgs) RefreshSpeechOutputUi()
            AddHandler btnOk.Click, AddressOf OnOkClick

            Dim selectedIndex As Integer =
                _descriptors.FindIndex(Function(d) d.DisplayName.Equals(currentEngineDisplayName, StringComparison.OrdinalIgnoreCase))

            If selectedIndex < 0 Then
                selectedIndex = 0
            End If

            If cboEngine.Items.Count > 0 Then
                cboEngine.SelectedIndex = selectedIndex
            End If

            Dim modeIndex As Integer = cboSpeechOutputMode.FindStringExact(SelectedSpeechOutputMode)
            cboSpeechOutputMode.SelectedIndex = If(modeIndex >= 0, modeIndex, 0)

            Dim speedLabel As String =
                If(SelectedSpeechOutputSpeed <= 1.0R,
                   "1.00x (Normal)",
                   FormatSpeechOutputSpeed(SelectedSpeechOutputSpeed))

            Dim speedIndex As Integer = cboSpeechOutputSpeed.FindStringExact(speedLabel)
            cboSpeechOutputSpeed.SelectedIndex = If(speedIndex >= 0, speedIndex, 0)

            RefreshSpeechOutputUi()
            ApplyCalculatedMinimumSize(root)
        End Sub

        Private Sub ApplyCalculatedMinimumSize(root As TableLayoutPanel)
            If root Is Nothing Then
                Return
            End If

            root.PerformLayout()

            Dim preferredClientSize As System.Drawing.Size =
                root.GetPreferredSize(New System.Drawing.Size(Integer.MaxValue, Integer.MaxValue))

            Dim nonClientWidth As Integer = Me.Width - Me.ClientSize.Width
            Dim nonClientHeight As Integer = Me.Height - Me.ClientSize.Height

            Dim minWidth As Integer = Math.Max(640, preferredClientSize.Width + nonClientWidth + 10)
            Dim minHeight As Integer = Math.Max(360, preferredClientSize.Height + nonClientHeight + 10)

            Me.MinimumSize = New System.Drawing.Size(minWidth, minHeight)

            If Me.Width < minWidth OrElse Me.Height < minHeight Then
                Me.Size = New System.Drawing.Size(Math.Max(Me.Width, minWidth), Math.Max(Me.Height, minHeight))
            End If
        End Sub

        Private Sub LoadAudioInputDevices(selectedWaveDeviceIndex As Integer)
            cboMicrophone.Items.Clear()
            _audioInputDevices.Clear()

            Dim maxMicrophoneItemWidth As Integer = cboMicrophone.DropDownWidth

            Try
                Using enumr As New MMDeviceEnumerator()
                    Dim devs = enumr.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active)

                    For Each d As MMDevice In devs
                        Dim friendlyName As String = If(d.FriendlyName, "").Trim()
                        If friendlyName.Length = 0 Then
                            friendlyName = d.ID
                        End If

                        Dim choice As New AudioInputDeviceChoice With {
                            .DeviceId = d.ID,
                            .WaveDeviceIndex = FindLegacyWaveInputDeviceIndex(friendlyName),
                            .DisplayText = friendlyName
                        }

                        _audioInputDevices.Add(choice)
                        cboMicrophone.Items.Add(choice)
                        maxMicrophoneItemWidth = System.Math.Max(
                            maxMicrophoneItemWidth,
                            TextRenderer.MeasureText(choice.DisplayText, cboMicrophone.Font).Width + 30)
                    Next
                End Using
            Catch
            End Try

            If _audioInputDevices.Count = 0 Then
                For i As Integer = 0 To WaveInEvent.DeviceCount - 1
                    Dim microphoneText As String = $"{i}: {WaveInEvent.GetCapabilities(i).ProductName}"
                    Dim choice As New AudioInputDeviceChoice With {
                        .DeviceId = "",
                        .WaveDeviceIndex = i,
                        .DisplayText = microphoneText
                    }

                    _audioInputDevices.Add(choice)
                    cboMicrophone.Items.Add(choice)
                    maxMicrophoneItemWidth = System.Math.Max(
                        maxMicrophoneItemWidth,
                        TextRenderer.MeasureText(choice.DisplayText, cboMicrophone.Font).Width + 30)
                Next
            End If

            cboMicrophone.DropDownWidth = maxMicrophoneItemWidth

            If cboMicrophone.Items.Count = 0 Then
                Return
            End If

            Dim selectedIndex As Integer =
                _audioInputDevices.FindIndex(Function(choice) choice.WaveDeviceIndex = selectedWaveDeviceIndex)

            cboMicrophone.SelectedIndex = If(selectedIndex >= 0, selectedIndex, 0)
        End Sub

        Private Shared Function FindLegacyWaveInputDeviceIndex(friendlyName As String) As Integer
            Dim normalizedFriendlyName As String = If(friendlyName, "").Trim()

            For i As Integer = 0 To WaveInEvent.DeviceCount - 1
                Dim waveName As String = If(WaveInEvent.GetCapabilities(i).ProductName, "").Trim()

                If String.Equals(waveName, normalizedFriendlyName, StringComparison.OrdinalIgnoreCase) Then
                    Return i
                End If

                If normalizedFriendlyName.IndexOf(waveName, StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                   waveName.IndexOf(normalizedFriendlyName, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Return i
                End If
            Next

            Return 0
        End Function

        Private Function GetSelectedAudioInputDeviceChoice() As AudioInputDeviceChoice
            Return TryCast(cboMicrophone.SelectedItem, AudioInputDeviceChoice)
        End Function

        Private Sub RefreshSpeechOutputUi()
            Dim available As Boolean = _ownerAdapter IsNot Nothing AndAlso _ownerAdapter.IsSpeechOutputAvailable

            chkSpeechOutputEnabled.Enabled = available
            cboSpeechOutputMode.Enabled = available
            chkUseSpeakerSpecificVoices.Enabled = available
            cboSpeechOutputSpeed.Enabled = available
            btnConfigureSpeechOutput.Enabled = available

            If Not available Then
                chkSpeechOutputEnabled.Checked = False
            End If

            If _ownerAdapter Is Nothing Then
                lblSpeechOutputSummary.Text = "Speech output unavailable."
            Else
                lblSpeechOutputSummary.Text = _ownerAdapter.GetSpeechOutputSummary()
            End If
        End Sub

        Private Function GetPreferredLanguageForDescriptor(descriptor As LiveEngineDescriptor) As String
            If descriptor Is Nothing Then
                Return SelectedLanguage
            End If

            If _savedLanguageByEngine.ContainsKey(descriptor.DisplayName) Then
                Return _savedLanguageByEngine(descriptor.DisplayName)
            End If

            Return SelectedLanguage
        End Function

        Private Sub OnEngineChanged(sender As Object, e As EventArgs)
            cboLanguage.Items.Clear()

            If cboEngine.SelectedIndex < 0 OrElse cboEngine.SelectedIndex >= _descriptors.Count Then
                Return
            End If

            Dim descriptor As LiveEngineDescriptor = _descriptors(cboEngine.SelectedIndex)
            Dim preferredLanguage As String = GetPreferredLanguageForDescriptor(descriptor)

            If descriptor.Kind = EngineKind.Vosk Then
                cboLanguage.Items.Add("(language comes from selected Vosk model)")
                cboLanguage.SelectedIndex = 0
                cboLanguage.Enabled = False
            Else
                cboLanguage.Enabled = True

                If descriptor.Languages IsNot Nothing Then
                    For Each item As String In descriptor.Languages
                        cboLanguage.Items.Add(item)
                    Next
                End If

                If cboLanguage.Items.Count = 0 Then
                    cboLanguage.Items.Add("auto")
                End If

                Dim preferredIndex As Integer = cboLanguage.FindStringExact(preferredLanguage)
                cboLanguage.SelectedIndex = If(preferredIndex >= 0, preferredIndex, 0)
            End If
        End Sub

        Private Sub OnConfigureSpeechOutputClick(sender As Object, e As EventArgs)
            If _ownerAdapter Is Nothing Then
                Return
            End If

            _ownerAdapter.ConfigureSpeechOutput(Me, chkUseSpeakerSpecificVoices.Checked)
            RefreshSpeechOutputUi()
        End Sub

        Private Sub OnOkClick(sender As Object, e As EventArgs)
            If cboEngine.SelectedIndex < 0 Then
                DialogResult = DialogResult.None
                Return
            End If

            Dim selectedMicrophone As AudioInputDeviceChoice = GetSelectedAudioInputDeviceChoice()

            SelectedDescriptor = _descriptors(cboEngine.SelectedIndex)
            SelectedLanguage = If(cboLanguage.SelectedItem IsNot Nothing, cboLanguage.SelectedItem.ToString(), "auto")
            SelectedMicrophoneDeviceIndex = If(selectedMicrophone IsNot Nothing, selectedMicrophone.WaveDeviceIndex, Math.Max(0, cboMicrophone.SelectedIndex))
            IncludeFullDocument = chkIncludeDocument.Checked
            SelectedSpeechOutputEnabled = chkSpeechOutputEnabled.Checked
            SelectedSpeechOutputMode = If(cboSpeechOutputMode.SelectedItem IsNot Nothing, cboSpeechOutputMode.SelectedItem.ToString(), "Queue")
            SelectedUseSpeakerSpecificVoices = chkUseSpeakerSpecificVoices.Checked
            SelectedSpeechOutputSpeed =
                ParseSpeechOutputSpeed(
                    If(cboSpeechOutputSpeed.SelectedItem IsNot Nothing,
                       cboSpeechOutputSpeed.SelectedItem.ToString(),
                       "1.00x (Normal)"))
        End Sub
    End Class
End Class

Partial Public Class ThisAddIn

    Private _talkToMeWidget As SharedLibrary.SharedLibrary.TalkToMeWidget = Nothing

    Public Function IsTalkToMeAvailable() As Boolean
        Try
            If Globals.Ribbons Is Nothing OrElse Globals.Ribbons.Ribbon1 Is Nothing Then
                Return False
            End If

            If Globals.Ribbons.Ribbon1.GetTalkToMeCommandDefinitions().Count = 0 Then
                Return False
            End If

            Return Not INILoadFail()
        Catch
            Return False
        End Try
    End Function

    Public Sub ShowTalkToMeWidget(Optional returnFocusAfterStart As System.Action = Nothing)
        If INILoadFail() Then
            Return
        End If

        If _talkToMeWidget Is Nothing OrElse _talkToMeWidget.IsDisposed Then
            Dim hostAdapter As New WordTalkToMeHostAdapter(Me)
            Dim speechAdapter As New WordTalkToMeSpeechAdapter(Me)
            Dim coordinator As New SharedLibrary.SharedLibrary.TalkToMeCoordinator(
                hostAdapter,
                Function() If(_talkToMeWidget IsNot Nothing, _talkToMeWidget.GetIncludeFullDocumentSetting(), False))

            _talkToMeWidget = New SharedLibrary.SharedLibrary.TalkToMeWidget(
                speechAdapter,
                coordinator)
        End If

        _talkToMeWidget.SetReturnFocusAfterStart(returnFocusAfterStart)
        _talkToMeWidget.ShowWidget()
    End Sub

    Public Sub SubmitTalkToMeExternalSpeech(speakerName As String, text As String)
        Try
            If _talkToMeWidget Is Nothing OrElse _talkToMeWidget.IsDisposed Then
                Return
            End If

            Dim speechTask As Task(Of Boolean) = _talkToMeWidget.SubmitExternalSpeechAsync(speakerName, text)
        Catch
        End Try
    End Sub

    Friend Sub ShutdownTalkToMe()
        Try
            If _talkToMeWidget Is Nothing Then
                Return
            End If

            If Not _talkToMeWidget.IsDisposed Then
                _talkToMeWidget.PrepareForHostShutdown()
            End If
        Catch
        End Try

        Try
            If _talkToMeWidget IsNot Nothing AndAlso Not _talkToMeWidget.IsDisposed Then
                _talkToMeWidget.Dispose()
            End If
        Catch
        Finally
            _talkToMeWidget = Nothing
        End Try
    End Sub

End Class
