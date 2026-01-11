' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license terms see https://redink.ai.

' =============================================================================
' File: ThisAddIn.WebAgent.vb
' Purpose:
'   Word add-in entry points for running and authoring WebAgent JSON scripts.
'   This file contains UI-driven workflow orchestration; command semantics and
'   step execution are implemented by `WebAgentInterpreter`.
'
' Responsibilities (this file):
'   - Script discovery and selection from configured directories:
'       * `INI_WebAgentPathLocal` (preferred) and `INI_WebAgentPath`
'       * pattern: `{AN2}-ag-*.json` (top directory only)
'   - Pre-run substitutions:
'       * user parameters via `SharedMethods.ProcessParameterPlaceholders`
'       * secret placeholders in `env.secrets` via `ProcessWebAgentSecretPlaceholders`
'         (prompts for values matching `{{secret:Description; type}}`)
'   - Pre-flight validation and safety gates:
'       * JSON parse/step presence validation
'       * confirmation gate for `send_email_report` steps (`ConfirmEmailSendSteps`)
'       * warnings for relative/scheme-less URLs when `env.base_url` is missing
'   - Run control:
'       * cancellation via `_webAgentRunCts` and an ESC polling task (`PollEscForCancel`)
'   - Result handling:
'       * shows final Markdown report and supports clipboard / document insertion /
'         pane transfer (optionally via `InsertTextWithMarkdown`)
'   - Script authoring workflow:
'       * new or amend existing scripts (`CreateModifyWebAgentScript`)
'       * saves to configured directory with `.bak` backup on overwrite
'       * optional secondary model selection via `INI_SecondAPI` / `INI_AlternateModelPath`
'
' Threading notes:
'   - ESC monitoring runs on a background `System.Threading.Tasks.Task`.
'   - The interpreter run is cooperative-cancelled via `CancellationToken`.
'
' Key dependencies used in this file:
'   - `WebAgentInterpreter` (execution engine; see `WebAgentInterpreter.vb`)
'   - `SharedLibrary.SharedMethods` (UI dialogs, clipboard, editor, model selection)
'   - `Newtonsoft.Json` (`JObject`/`JArray` parsing and mutation)
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Data
Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports System.Threading
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Discovers and prompts user to select a WebAgent script file from configured local/global directories.
    ''' Local scripts (INI_WebAgentPathLocal) are listed first, followed by global (INI_WebAgentPath).
    ''' </summary>
    ''' <returns>Full path to selected script file, or Nothing if canceled or no scripts found.</returns>
    Private Function GetWebAgentScriptFile() As String
        Dim globalPath As String = ExpandEnvironmentVariables(INI_WebAgentPath)
        Dim localPath As String = ExpandEnvironmentVariables(INI_WebAgentPathLocal)

        Dim candidates As New List(Of (Display As String, FullPath As String, IsLocal As Boolean))

        If String.IsNullOrWhiteSpace(localPath) AndAlso String.IsNullOrWhiteSpace(globalPath) Then
            ShowCustomMessageBox("No WebAgent paths configured. Please set 'WebAgentPath' / 'WebAgentPathLocal' parameters in the configuration file.", $"{AN} WebAgent")
            Return Nothing
        End If

        Try
            If Not String.IsNullOrWhiteSpace(localPath) AndAlso IO.Directory.Exists(localPath) Then
                For Each f In IO.Directory.GetFiles(localPath, $"{AN2}-ag-*.json", IO.SearchOption.TopDirectoryOnly)
                    candidates.Add(($"{IO.Path.GetFileName(f)} (local)", f, True))
                Next
            End If
        Catch ex As Exception
            ShowCustomMessageBox("Error enumerating local WebAgent scripts: " & ex.Message, $"{AN} WebAgent")
        End Try

        Try
            If Not String.IsNullOrWhiteSpace(globalPath) AndAlso IO.Directory.Exists(globalPath) Then
                For Each f In IO.Directory.GetFiles(globalPath, $"{AN2}-ag-*.json", IO.SearchOption.TopDirectoryOnly)
                    candidates.Add(($"{IO.Path.GetFileName(f)}", f, False))
                Next
            End If
        Catch ex As Exception
            ShowCustomMessageBox("Error enumerating global WebAgent scripts: " & ex.Message, $"{AN} WebAgent")
        End Try

        If candidates.Count = 0 Then
            ShowCustomMessageBox($"No WebAgent scripts found. Please set 'WebAgentPath' / 'WebAgentPathLocal' parameters in the configuration file and place files named '{AN2}-ag-*.json'.", $"{AN} WebAgent")
            Return Nothing
        End If

        ' Local first, then global; alphabetical within group
        Dim ordered = candidates.
            OrderByDescending(Function(c) c.IsLocal).
            ThenBy(Function(c) c.Display, StringComparer.OrdinalIgnoreCase).
            ToList()

        Dim displayItems = ordered.Select(Function(c) c.Display).ToArray()
        Dim sel As String = ShowSelectionForm("Select the WebAgent script you want to run:", $"{AN} WebAgent", displayItems)
        If String.IsNullOrEmpty(sel) Then Return Nothing

        Dim chosen = ordered.FirstOrDefault(Function(o) o.Display = sel)
        Return chosen.FullPath
    End Function


    ''' <summary>
    ''' Detects secret placeholders in script's env.secrets section (pattern: {{secret:Description; type}}).
    ''' Prompts user for secret values via ShowCustomVariableInputForm (values not logged).
    ''' Substitutes raw values into JSON tree before interpreter sees the script.
    ''' </summary>
    ''' <param name="script">JSON script string passed by reference; mutated in place with secret values.</param>
    ''' <returns>True if successful or no secrets found; False if user cancels or processing fails.</returns>
    Private Function ProcessWebAgentSecretPlaceholders(ByRef script As String) As Boolean
        Try
            Dim root As JObject = Nothing
            Try
                root = JObject.Parse(script)
            Catch
                ' Let normal JSON error handling later handle issues
                Return True
            End Try
            Dim envObj = TryCast(root("env"), JObject)
            If envObj Is Nothing Then Return True
            Dim secretsObj = TryCast(envObj("secrets"), JObject)
            If secretsObj Is Nothing OrElse secretsObj.Count = 0 Then Return True

            Dim secretDefs As New List(Of (Name As String, Placeholder As String, Desc As String))()

            Dim rx As New Regex("^\{\{\s*secret\s*:(.+?)\}\}$", RegexOptions.IgnoreCase)
            For Each p In secretsObj.Properties()
                Dim v = p.Value.ToString().Trim()
                Dim m = rx.Match(v)
                If m.Success Then
                    Dim meta = m.Groups(1).Value.Trim()
                    ' Metadata format supports "Description; type" segments; only description is used for prompts
                    Dim segs = meta.Split(";"c).Select(Function(s) s.Trim()).Where(Function(s) s <> "").ToArray()
                    Dim desc = segs(0)
                    secretDefs.Add((p.Name, v, desc))
                End If
            Next
            If secretDefs.Count = 0 Then Return True

            Dim ipList As New List(Of SharedLibrary.SharedLibrary.SharedMethods.InputParameter)
            For Each sd In secretDefs
                ipList.Add(New SharedLibrary.SharedLibrary.SharedMethods.InputParameter($"Secret: {sd.Name} - {sd.Desc}", ""))
            Next

            Dim arr = ipList.ToArray()
            If Not ShowCustomVariableInputForm("Provide secret values (they are not logged):", "WebAgent Secrets", arr) Then
                Return False
            End If

            ' Substitute into JSON tree (not by raw string replace to avoid accidental collisions)
            For i = 0 To secretDefs.Count - 1
                Dim name = secretDefs(i).Name
                Dim userVal = CStr(arr(i).Value)
                secretsObj(name) = userVal
            Next

            script = root.ToString(Formatting.None)
            Return True
        Catch ex As Exception
            ShowCustomMessageBox("Secret placeholder processing failed: " & ex.Message, AN & " WebAgent")
            Return False
        End Try
    End Function


    ''' <summary>
    ''' Cancellation token source for the currently running WebAgent execution.
    ''' </summary>
    Private _webAgentRunCts As CancellationTokenSource

    ''' <summary>
    ''' Cancels the currently running WebAgent execution by triggering the cancellation token.
    ''' Safe to call even if no run is active.
    ''' </summary>
    Public Sub CancelWebAgentRun()
        Try : _webAgentRunCts?.Cancel() : Catch : End Try
    End Sub

    ''' <summary>
    ''' Background polling thread that monitors ESC key state via GetAsyncKeyState.
    ''' Cancels provided CancellationTokenSource when ESC pressed (debounced to require key release).
    ''' Runs until cancellation token signaled or ESC pressed.
    ''' </summary>
    ''' <param name="cts">CancellationTokenSource to cancel when ESC detected.</param>
    Private Sub PollEscForCancel(cts As CancellationTokenSource)
        Try
            Dim escLatched As Boolean = False
            While Not cts.IsCancellationRequested
                Thread.Sleep(150)
                Dim state = GetAsyncKeyState(VK_ESCAPE)
                If (state And &H8000) <> 0 Then
                    ' Debounce: ensure key released once before re-trigger
                    If Not escLatched Then
                        escLatched = True
                        cts.Cancel()
                        Exit While
                    End If
                Else
                    escLatched = False
                End If
            End While
        Catch
            ' Ignore polling errors
        End Try
    End Sub

    ''' <summary>
    ''' Scans script steps for 'send_email_report' commands and displays confirmation dialog
    ''' listing recipients (to, subject, smtp_host) before execution proceeds.
    ''' Security gate to prevent unintentional email sends.
    ''' </summary>
    ''' <param name="root">Parsed JSON root object of the script.</param>
    ''' <returns>1 if user confirms or no email steps found; 0 if user cancels or error occurs.</returns>
    Private Function ConfirmEmailSendSteps(root As JObject) As Integer
        Try
            Dim stepsArr = TryCast(root("steps"), JArray)
            If stepsArr Is Nothing OrElse stepsArr.Count = 0 Then Return 1

            Dim emailSteps As New List(Of JObject)
            For Each st As JObject In stepsArr
                Dim cmd = st.Value(Of String)("command")
                If Not String.IsNullOrWhiteSpace(cmd) AndAlso
                   String.Equals(cmd, "send_email_report", StringComparison.OrdinalIgnoreCase) Then
                    emailSteps.Add(st)
                End If
            Next
            ' No email steps found
            If emailSteps.Count = 0 Then Return 1

            ' Collect recipient info and optionally smtp_host to display
            Dim lines As New List(Of String)
            For Each st In emailSteps
                Dim p = TryCast(st("params"), JObject)
                Dim toRaw As String = p?("to")?.ToString()
                Dim smtpHost As String = p?("smtp_host")?.ToString()
                Dim subject As String = p?("subject")?.ToString()

                Dim recipients As String = ""
                If Not String.IsNullOrWhiteSpace(toRaw) Then
                    ' Normalize delimiters and trim each recipient
                    Dim parts = toRaw.Split(New String() {",", ";"}, StringSplitOptions.RemoveEmptyEntries).
                                      Select(Function(s) s.Trim()).
                                      Where(Function(s) s <> "")
                    recipients = String.Join(", ", parts)
                Else
                    recipients = "(missing)"
                End If

                Dim meta As New List(Of String)
                If Not String.IsNullOrWhiteSpace(subject) Then meta.Add($"Subject: {subject}")
                If Not String.IsNullOrWhiteSpace(smtpHost) Then meta.Add($"SMTP: {smtpHost}")

                If meta.Count > 0 Then
                    lines.Add($"- To: {recipients}  [{String.Join(" | ", meta)}]")
                Else
                    lines.Add($"- To: {recipients}")
                End If
            Next

            Dim msg As String =
                "Security check: This WebAgent script contains step(s) that will send an e-mail." & vbCrLf & vbCrLf &
                "Recipients detected:" & vbCrLf &
                String.Join(vbCrLf, lines) & vbCrLf & vbCrLf &
                "Do you wish to continue?"

            ' Ask for confirmation
            Dim ok As Integer = ShowCustomYesNoBox(msg, "Yes", "No, abort")
            Return ok
        Catch ex As Exception
            ' On any unexpected error, be safe and require explicit approval by failing the check
            ShowCustomMessageBox("Email send confirmation failed: " & ex.Message, $"{AN} WebAgent")
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' Main entry point for WebAgent script execution workflow.
    ''' Sequence: Script selection → Parameter/secret substitution → JSON validation → Pre-flight checks
    ''' (email confirmation, relative URL warnings) → Interpreter execution with ESC cancellation monitoring
    ''' → Output display with clipboard/insert/pane options.
    ''' Handles errors and provides user feedback at each stage.
    ''' </summary>
    Public Async Sub WebAgent()

        If INILoadFail() Then Return

        ' 1) Let user pick a script file (local first) 
        Dim selectedFile As String = GetWebAgentScriptFile()
        If String.IsNullOrWhiteSpace(selectedFile) Then Exit Sub

        Dim script As String = Nothing
        Try
            script = IO.File.ReadAllText(selectedFile, System.Text.Encoding.UTF8)
        Catch ex As Exception
            ShowCustomMessageBox("Failed to read script file:" & vbCrLf & selectedFile & vbCrLf & ex.Message, $"{AN} WebAgent")
            Exit Sub
        End Try

        If String.IsNullOrWhiteSpace(script) Then
            ShowCustomMessageBox("Script file is empty: " & selectedFile, $"{AN} WebAgent")
            Exit Sub
        End If

        ' 2) Detect and process parameter placeholders {{parameterN = ...}}
        If Not SLib.ProcessParameterPlaceholders(script) Then
            ShowCustomMessageBox("WebAgent canceled (no parameters applied).", $"{AN} WebAgent")
            Exit Sub
        End If

        If Not ProcessWebAgentSecretPlaceholders(script) Then
            ShowCustomMessageBox("WebAgent canceled (secrets).", $"{AN} WebAgent")
            Exit Sub
        End If

        ' 3) Parse (after substitution) to validate and for pre-flight checks
        Dim rootObj As JObject = Nothing
        Try
            rootObj = JObject.Parse(script)
        Catch ex As Exception
            SLib.PutInClipboard(ex.Message)
            ShowCustomMessageBox("Your script contains an incorrect JSON string (error copied to clipboard):" & vbCrLf & ex.Message, $"{AN} WebAgent")
            Exit Sub
        End Try

        Dim stepsArr = TryCast(rootObj("steps"), JArray)
        If stepsArr Is Nothing OrElse stepsArr.Count = 0 Then
            ShowCustomMessageBox("No steps found in script.", $"{AN} WebAgent")
            Exit Sub
        End If

        If ConfirmEmailSendSteps(rootObj) <> 1 Then
            ShowCustomMessageBox("WebAgent canceled (email send not approved).", $"{AN} WebAgent")
            Exit Sub
        End If

        ' 4) Pre-flight relative URL warnings
        ' Skip URLs containing template placeholders (they will be resolved at runtime)
        Dim baseUrl As String = rootObj.SelectToken("env.base_url")?.ToString()
        Dim relativeIssues As New List(Of String)
        For Each st As JObject In stepsArr
            Dim cmd = st.Value(Of String)("command")
            If cmd = "open_url" OrElse cmd = "http_request" Then
                Dim p = TryCast(st("params"), JObject)
                Dim rawUrl = p?("url")?.ToString()
                If Not String.IsNullOrWhiteSpace(rawUrl) Then
                    ' Skip validation if URL contains template placeholders - they resolve at runtime
                    If rawUrl.Contains("{{") Then
                        Continue For
                    End If
                    If Not rawUrl.StartsWith("http://", StringComparison.OrdinalIgnoreCase) AndAlso
                       Not rawUrl.StartsWith("https://", StringComparison.OrdinalIgnoreCase) Then
                        If String.IsNullOrWhiteSpace(baseUrl) Then
                            relativeIssues.Add($"Step id='{st.Value(Of String)("id")}' uses relative or scheme-less URL '{rawUrl}' but env.base_url is not set.")
                        End If
                    End If
                Else
                    ' Only warn about empty URLs if they don't appear to be template-based
                    Dim urlToken = p?("url")
                    If urlToken Is Nothing Then
                        relativeIssues.Add($"Step id='{st.Value(Of String)("id")}' has missing params.url.")
                    End If
                    ' If url exists but is empty string, skip warning - might be set dynamically
                End If
            End If
        Next
        If relativeIssues.Count > 0 Then
            Dim warnText As New StringBuilder()
            warnText.AppendLine("Pre-flight URL warnings:")
            For Each w In relativeIssues
                warnText.AppendLine(w)
            Next
            ShowCustomMessageBox(warnText.ToString().TrimEnd(), $"{AN} WebAgent", 20)
        End If

        ' 5) Run interpreter
        Dim finalMd As String = ""
        Dim webAgentCompleted As Boolean = False
        Dim abortedByEsc As Boolean = False

        _webAgentRunCts = New CancellationTokenSource()
        Dim ctsWeb = _webAgentRunCts

        ' Fully qualify Task to avoid ambiguity with Microsoft.Office.Interop.Word.Task
        Dim escMonitor As System.Threading.Tasks.Task =
            System.Threading.Tasks.Task.Run(Sub() PollEscForCancel(ctsWeb))

        Using interp As New WebAgentInterpreter()
            ' Set up cancellation callback so LogWindow close can abort the run
            interp.OnCancelRequested = Sub() CancelWebAgentRun()

            Dim sw As Stopwatch = Stopwatch.StartNew()
            Try
                finalMd = Await interp.RunAsync(
                    rootObj.ToString(Formatting.None),
                    _context,
                    True,   ' useSecondAPI
                    True,   ' autoselectModel
                    ctsWeb.Token)

                webAgentCompleted = True

            Catch ex As OperationCanceledException
                abortedByEsc = ctsWeb.IsCancellationRequested
                If abortedByEsc Then
                    finalMd = "# WebAgent aborted (ESC)" & vbCrLf & "User pressed ESC."
                Else
                    finalMd = "# WebAgent run canceled" & vbCrLf & ex.Message
                End If

            Catch ex As Exception
                Dim coreMsg = If(String.IsNullOrWhiteSpace(ex.Message), ex.GetType().Name, ex.Message)
                finalMd = "# WebAgent run failed" & vbCrLf & coreMsg
            End Try
            sw.Stop()
        End Using

        ' Stop ESC monitor task (fully qualified Task to suppress ambiguity)
        Try
            ctsWeb.Cancel()
            System.Threading.Tasks.Task.WaitAny(New System.Threading.Tasks.Task() {escMonitor}, 250)
        Catch
        End Try

        Dim isFailure As Boolean =
                    String.IsNullOrWhiteSpace(finalMd) OrElse
                    finalMd.StartsWith("# WebAgent run failed", StringComparison.OrdinalIgnoreCase) OrElse
                    finalMd.StartsWith("# WebAgent run canceled", StringComparison.OrdinalIgnoreCase) OrElse
                    finalMd.StartsWith("# WebAgent aborted", StringComparison.OrdinalIgnoreCase)

        If Not isFailure Then
            If finalMd.IndexOf("{{", StringComparison.Ordinal) >= 0 Then
                finalMd &= vbCrLf & vbCrLf & "_Warning: Unresolved template placeholders remain in output._"
            End If
        ElseIf String.IsNullOrWhiteSpace(finalMd) Then
            finalMd = "# WebAgent produced no output."
        End If

        InfoBox.ShowInfoBox("")

        Dim DialogResult As String = ShowCustomWindow("The WebAgent produced the following report:", finalMd, "You can edit the report. If you select OK, it will be copied to the clipboard in the original or edited form. You can also have the original inserted or transferred to the pane.", $"{AN} WebAgent", False, False, True, True)

        If DialogResult <> "" AndAlso DialogResult <> "Pane" Then
            If DialogResult = "Markdown" Then
                Dim NewDocChoice As Integer = ShowCustomYesNoBox("Do you want to insert the text into a new Word document (if you cancel, it will be in the clipboard with formatting)?", "Yes, new", "No, into my existing doc")

                If NewDocChoice = 1 Then
                    Dim newDoc As Word.Document = Globals.ThisAddIn.Application.Documents.Add()
                    Dim currentSelection As Word.Selection = newDoc.Application.Selection
                    currentSelection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    InsertTextWithMarkdown(currentSelection, finalMd, True, True)
                ElseIf NewDocChoice = 2 Then
                    Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    Globals.ThisAddIn.Application.Selection.TypeParagraph()
                    InsertTextWithMarkdown(Globals.ThisAddIn.Application.Selection, vbCrLf & finalMd, False)
                Else
                    ShowCustomMessageBox("No text was inserted (but included in the clipboard as RTF).")
                    SLib.PutInClipboard(finalMd)
                End If
            Else
                SLib.PutInClipboard(DialogResult)
            End If
        ElseIf DialogResult = "Pane" Then
            ShowPaneAsync(
                "The WebAgent produced the following report:",
                finalMd,
                "",
                AN,
                noRTF:=False,
                insertMarkdown:=True
            )
        End If
    End Sub


    ''' <summary>
    ''' LLM-assisted WebAgent script creation or amendment workflow.
    ''' Sequence: Directory validation → New/Amend mode selection → (if amend) Script editor for manual review
    ''' → User instruction prompt → Model selection (primary/secondary) → LLM generation with spec prompts
    ''' (WebAgentJSONInstruct + WebAgentParameterSpec) → JSON validation with optional auto-correction loop
    ''' → Save with .bak backup → Final editor display for manual parameter definition.
    ''' Supports overwrite confirmation and alternate model loading (INI_AlternateModelPath).
    ''' </summary>
    Public Async Sub CreateModifyWebAgentScript()

        If INILoadFail() Then Return

        ' 0) Preconditions: paths configured?
        Dim globalPath As String = ExpandEnvironmentVariables(INI_WebAgentPath)
        Dim localPath As String = ExpandEnvironmentVariables(INI_WebAgentPathLocal)

        If (String.IsNullOrWhiteSpace(globalPath) AndAlso String.IsNullOrWhiteSpace(localPath)) Then
            ShowCustomMessageBox("No WebAgent paths configured ('WebAgentPath' / 'WebAgentPathLocal' parameters in the configuration file).", $"{AN} WebAgent")
            Exit Sub
        End If

        ' Pick working directory (prefer local if available)
        Dim targetDir As String = If(Not String.IsNullOrWhiteSpace(localPath) AndAlso IO.Directory.Exists(localPath),
                                     localPath,
                                     If(Not String.IsNullOrWhiteSpace(globalPath) AndAlso IO.Directory.Exists(globalPath),
                                        globalPath,
                                        Nothing))
        If String.IsNullOrWhiteSpace(targetDir) Then
            ShowCustomMessageBox("Neither configured WebAgent directory exists.", $"{AN} WebAgent")
            Exit Sub
        End If

        ' 1) New vs Amend
        Dim modeChoice = ShowCustomYesNoBox(
            "Do you want to create a new WebAgent script or amend an existing one?",
            "New",
            "Amend existing",
            $"{AN} WebAgent")
        If modeChoice = 0 Then Exit Sub

        Dim existingScriptPath As String = Nothing
        Dim originalScriptText As String = Nothing
        Dim isAmend As Boolean = (modeChoice = 2)

        If isAmend Then
            existingScriptPath = GetWebAgentScriptFile()
            If String.IsNullOrWhiteSpace(existingScriptPath) Then Exit Sub
            Try
                ' Read current script (initial content before possible user edits)
                originalScriptText = IO.File.ReadAllText(existingScriptPath, System.Text.Encoding.UTF8)

                ' Open in embedded text editor so user can review and edit
                ' Editor saves directly to the same path when user clicks Save
                SharedLibrary.SharedLibrary.SharedMethods.ShowTextFileEditor(existingScriptPath, $"{AN} WebAgent Script '{existingScriptPath}' (you can have it amended by the LLM once you close the editor):")

                ' Ask whether to proceed with automatic LLM amendment after editor closes
                Dim amendChoice As Integer = SharedLibrary.SharedLibrary.SharedMethods.ShowCustomYesNoBox(
                    bodyText:="Do you want to automatically amend the script using the LLM now?",
                    button1Text:="Yes",
                    button2Text:="No",
                    header:=$"{AN} WebAgent"
                )

                ' ShowCustomYesNoBox returns 1 for first button ("Yes"), 2 for second ("No")
                If amendChoice <> 1 Then
                    Exit Sub
                End If

                ' Re-read file to capture any edits the user may have saved
                Try
                    originalScriptText = IO.File.ReadAllText(existingScriptPath, System.Text.Encoding.UTF8)
                Catch exReload As Exception
                    ShowCustomMessageBox("Failed to re-read edited script before amendment:" & vbCrLf & exReload.Message, $"{AN} WebAgent")
                    Exit Sub
                End Try

            Catch ex As Exception
                ShowCustomMessageBox("Unable to read existing script:" & vbCrLf & ex.Message, $"{AN} WebAgent")
                Exit Sub
            End Try
        End If

        ' 2) If new, gather filename with validation and overwrite confirmation
        Dim newScriptPath As String = existingScriptPath
        If Not isAmend Then
            Dim defaultName As String = $"{AN2}-ag-newscript.json"
            Dim rxName As New System.Text.RegularExpressions.Regex($"^{AN2}-ag-[a-z0-9._-]+\.json$", RegexOptions.IgnoreCase)
            Do
                Dim nameInput = ShowCustomInputBox($"Enter new script file name (pattern: {AN2}-ag-*.json)", $"{AN} WebAgent", True, defaultName)
                If String.IsNullOrWhiteSpace(nameInput) Then Exit Sub
                nameInput = nameInput.Trim()
                If Not rxName.IsMatch(nameInput) Then
                    ShowCustomMessageBox($"File name does not comply with pattern '{AN2}-ag-*.json'.", $"{AN} WebAgent")
                    Continue Do
                End If
                Dim fullCandidate = IO.Path.Combine(targetDir, nameInput)
                If IO.File.Exists(fullCandidate) Then
                    Dim ow = ShowCustomYesNoBox($"File '{nameInput}' already exists. Overwrite?", "Yes", "No", $"{AN} WebAgent")
                    If ow = 1 Then
                        newScriptPath = fullCandidate
                        Exit Do
                    ElseIf ow = 0 Then
                        Exit Sub
                    Else
                        Continue Do
                    End If
                Else
                    newScriptPath = fullCandidate
                    Exit Do
                End If
            Loop
        End If

        If String.IsNullOrWhiteSpace(newScriptPath) Then Exit Sub

        ' 3) User instructions
        Dim instruction = ShowCustomInputBox("You are about to have the LLM create or amend your script. For this purpose, describe what the WebAgent script shall do or how it shall be changed (note: user parameters need to be defined manually as per the user manual):", $"{AN} WebAgent", False, "")
        If String.IsNullOrWhiteSpace(instruction) Then Exit Sub
        OtherPrompt = instruction.Trim()

        If OtherPrompt = "" Then Return

        Dim useSecondAPI As Boolean = False

        Try

            ' 4) Primary vs Secondary model selection
            If INI_SecondAPI Then
                Dim modelChoice = ShowCustomYesNoBox("Select the model to generate or modify the script:", "Primary", "Secondary", $"{AN} WebAgent")
                If modelChoice = 0 Then Exit Sub
                If modelChoice = 2 Then
                    useSecondAPI = True

                    ' Attempt to autoload alternate model (best effort, non-fatal)
                    If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                        If Not ShowModelSelection(_context, INI_AlternateModelPath) Then
                            originalConfigLoaded = False
                            Return
                        End If
                    End If

                End If
            End If

            ' 5) Build user prompt (with optional existing script for amendment)
            Dim userPromptBuilder As New System.Text.StringBuilder()
            userPromptBuilder.Append("<USERPROMPT>").Append(OtherPrompt).Append("</USERPROMPT>")
            If isAmend AndAlso Not String.IsNullOrWhiteSpace(originalScriptText) Then
                userPromptBuilder.Append("<SCRIPT>").Append(originalScriptText).Append("</SCRIPT>")
            End If
            Dim userPromptFinal = userPromptBuilder.ToString()

            ' 6) Call LLM to generate or amend script
            Dim generated As String = Nothing
            Try
                generated = Await LLM(WebAgentJSONInstruct & " " & WebAgentParameterSpec, userPromptFinal, "", "", 0, useSecondAPI)
            Catch ex As Exception
                ShowCustomMessageBox("LLM call failed: " & ex.Message, $"{AN} WebAgent")
                Exit Sub
            End Try
            If String.IsNullOrWhiteSpace(generated) Then
                ShowCustomMessageBox("The LLM returned no content.", $"{AN} WebAgent")
                Exit Sub
            End If

            generated = WebAgentInterpreter.SanitizeLlmResult(generated)

            ' 7) Backup existing file before overwriting (only on first save, not during auto-correct loop)
            If IO.File.Exists(newScriptPath) AndAlso isAmend Then
                ' Amending existing file - make backup
                Try
                    SharedLibrary.SharedLibrary.SharedMethods.RenameFileToBak(newScriptPath)
                Catch
                End Try
            ElseIf IO.File.Exists(newScriptPath) AndAlso Not isAmend Then
                ' Overwriting a file in new mode - make backup
                Try
                    SharedLibrary.SharedLibrary.SharedMethods.RenameFileToBak(newScriptPath)
                Catch
                End Try
            End If

            ' 8) Save initial result
            Try
                IO.File.WriteAllText(newScriptPath, generated, System.Text.Encoding.UTF8)
            Catch ex As Exception
                ShowCustomMessageBox("Failed to save generated script (script is in the clipboard instead):" & vbCrLf & ex.Message, $"{AN} WebAgent")
                PutInClipboard(generated)
                Exit Sub
            End Try

            ' 9) Validate JSON, with optional auto-correction loop
            Dim keepLoop As Boolean = True
            Dim firstTry As Boolean = True
            While keepLoop
                Dim parseOk As Boolean = True
                Dim parseError As String = Nothing
                Try
                    Newtonsoft.Json.Linq.JObject.Parse(IO.File.ReadAllText(newScriptPath, System.Text.Encoding.UTF8))
                Catch ex As Exception
                    parseOk = False
                    parseError = ex.Message
                End Try

                If parseOk Then
                    ' Validation successful
                    keepLoop = False
                    If firstTry Then
                        ShowCustomMessageBox("The JSON script created by the LLM has been tested for validity and was saved to:" & vbCrLf & vbCrLf & newScriptPath, $"{AN} WebAgent")
                    Else
                        ShowCustomMessageBox("The JSON script created by the LLM is now valid and was saved to:" & vbCrLf & vbCrLf & newScriptPath, $"{AN} WebAgent")
                    End If
                Else
                    ' Validation failed - offer auto-correction
                    Dim choice = ShowCustomYesNoBox($"The script JSON is invalid (it was saved anyway). Error:{vbCrLf}{parseError}{vbCrLf}{vbCrLf}Attempt auto-correction?", "Yes", "No", $"{AN} WebAgent")
                    If choice = 1 Then
                        ' Attempt auto-correction via LLM
                        Dim brokenJson = IO.File.ReadAllText(newScriptPath, System.Text.Encoding.UTF8)
                        Dim correctionSystem =
                        "You are a strict JSON repair assistant. Fix structural JSON errors ONLY. Preserve semantic content & keys. Output EXACTLY one corrected JSON object, no comments, no prose."
                        Dim correctionUser =
                        "<BROKEN_JSON>" & brokenJson & "</BROKEN_JSON><ERROR_MESSAGE>" & parseError & "</ERROR_MESSAGE>"
                        Dim corrected As String = Nothing
                        Try
                            corrected = Await LLM(correctionSystem, correctionUser, "", "", 0, useSecondAPI)
                        Catch ex As Exception
                            ShowCustomMessageBox("Auto-correction LLM call failed: " & ex.Message, $"{AN} WebAgent")
                            keepLoop = False
                            Exit While
                        End Try
                        If String.IsNullOrWhiteSpace(corrected) Then
                            ShowCustomMessageBox("Auto-correction returned empty content.", $"{AN} WebAgent")
                            keepLoop = False
                            Exit While
                        End If
                        Try
                            IO.File.WriteAllText(newScriptPath, corrected, System.Text.Encoding.UTF8)
                        Catch ex As Exception
                            ShowCustomMessageBox("Failed to save corrected script (script is in the clipboard instead):" & ex.Message, $"{AN} WebAgent")
                            PutInClipboard(corrected)
                            keepLoop = False
                            Exit While
                        End Try
                        ' Loop will re-validate
                    Else
                        keepLoop = False
                    End If
                    firstTry = False
                End If
            End While

        Catch
        Finally
            If useSecondAPI And originalConfigLoaded Then
                RestoreDefaults(_context, originalConfig)
                originalConfigLoaded = False
            End If
        End Try

        ' 10) Show final file in editor for manual parameter definition
        Try
            ShowTextFileEditor(newScriptPath, $"{AN} WebAgent Script '{newScriptPath}'" & "(you can now add user parameters, e.g. '{{parameter1 = Date; String; 1.10.2025}}' - see user manual): ")
        Catch
        End Try

    End Sub


    ''' <summary>
    ''' Complete specification document for WebAgent JSON script format, provided to LLMs for reliable authoring.
    ''' </summary>
    Public Const WebAgentJSONInstruct As String =
        "WEBAGENT SCRIPT SPECIFICATION" & vbCrLf &
        "=============================" & vbCrLf &
        "" & vbCrLf &
        "OVERVIEW: WebAgent scripts are JSON documents that automate web interactions, data extraction, " &
        "LLM analysis, and report generation. The interpreter executes steps sequentially using an HTTP client " &
        "with cookie support, HTML parsing via HtmlAgilityPack, and template expansion for dynamic values." & vbCrLf &
        "" & vbCrLf &
        "EXECUTION FLOW:" & vbCrLf &
        "1. Script JSON is parsed and validated" & vbCrLf &
        "2. meta section configures timeouts and user-agent" & vbCrLf &
        "3. env section initializes base_url, headers, secrets, and variables" & vbCrLf &
        "4. steps array executes sequentially (jumps possible via on_error.goto or guard.else_goto)" & vbCrLf &
        "5. Each step: guard check -> wait_for delay -> command execution -> assign result -> post-checks" & vbCrLf &
        "6. Final output from render_report or execution log returned as Markdown" & vbCrLf &
        "" & vbCrLf &
        "=== TOP-LEVEL STRUCTURE ===" & vbCrLf &
        "{" & vbCrLf &
        "  ""meta"": { /* optional settings */ }," & vbCrLf &
        "  ""env"": { /* optional environment */ }," & vbCrLf &
        "  ""steps"": [ /* REQUIRED array of step objects */ ]" & vbCrLf &
        "}" & vbCrLf &
        "" & vbCrLf &
        "META SECTION (optional):" & vbCrLf &
        "- default_timeout_ms: Integer, default 30000. Applied to HTTP operations without explicit timeout." & vbCrLf &
        "- user_agent: String, default ""WebAgentInterpreter/1.0"". Sent with all HTTP requests." & vbCrLf &
        "" & vbCrLf &
        "ENV SECTION (optional):" & vbCrLf &
        "- base_url: String. Used to resolve relative URLs. Also available as {{base_url}} in templates." & vbCrLf &
        "- headers: Object {name:value,...}. Applied to all HTTP requests. Can be overridden per-request." & vbCrLf &
        "- secrets: Object {name:value,...}. Values starting with ""secret://"" trigger internal lookup. " &
        "  Secrets are masked in logs. User prompted for {{secret:Description}} placeholders before execution." & vbCrLf &
        "- variables: Object {name:value,...}. Initial values loaded into internal _vars dictionary. " &
        "  Any JSON type allowed. Accessible via {{varName}} templates." & vbCrLf &
        "" & vbCrLf &
        "=== STEP OBJECT STRUCTURE ===" & vbCrLf &
        "Each step in the steps array:" & vbCrLf &
        "{" & vbCrLf &
        "  ""id"": ""uniqueStepId"",        // RECOMMENDED: used for jumps, logging, debugging" & vbCrLf &
        "  ""command"": ""commandName"",    // REQUIRED: see Command Reference below" & vbCrLf &
        "  ""params"": { ... },            // Command-specific parameters" & vbCrLf &
        "  ""timeout_ms"": 15000,          // Override default timeout for this step" & vbCrLf &
        "  ""retry"": { ... },             // Retry configuration" & vbCrLf &
        "  ""on_error"": { ... },          // Error handling after retries exhausted" & vbCrLf &
        "  ""assign"": { ... },            // Store command result in variable" & vbCrLf &
        "  ""guard"": { ... },             // Conditional execution" & vbCrLf &
        "  ""wait_for"": { ... }           // Pre-step delay or post-step validation" & vbCrLf &
        "}" & vbCrLf &
        "" & vbCrLf &
        "RETRY OBJECT:" & vbCrLf &
        "{ ""max"": 3, ""delay_ms"": 1000, ""backoff"": 2.0 }" & vbCrLf &
        "- max: Maximum retry attempts (0 = no retries)" & vbCrLf &
        "- delay_ms: Initial delay before first retry" & vbCrLf &
        "- backoff: Multiplier for exponential backoff. Formula: delay = delay_ms * backoff^attemptIndex" & vbCrLf &
        "- Transient HTTP codes auto-trigger retry: 408, 425, 429, 500, 502, 503, 504" & vbCrLf &
        "" & vbCrLf &
        "ON_ERROR OBJECT:" & vbCrLf &
        "{ ""action"": ""continue|goto|abort"", ""goto"": ""stepId"" }" & vbCrLf &
        "- continue: Swallow error, proceed to next step" & vbCrLf &
        "- goto: Jump to specified step id (requires goto field)" & vbCrLf &
        "- abort: Re-throw exception, halt script execution" & vbCrLf &
        "- Applied only after ALL retries exhausted" & vbCrLf &
        "" & vbCrLf &
        "ASSIGN OBJECT:" & vbCrLf &
        "{ ""var"": ""variableName"", ""path"": ""json.path.to.value"" }" & vbCrLf &
        "- var: Name to store result under in _vars dictionary" & vbCrLf &
        "- path: Optional. Uses JToken.SelectToken to extract nested value (e.g., ""data.items[0].title"")" & vbCrLf &
        "- If path specified but not found, stores null" & vbCrLf &
        "- Result is deep-cloned before storage to prevent mutation" & vbCrLf &
        "" & vbCrLf &
        "GUARD OBJECT:" & vbCrLf &
        "{ ""if"": ""condition expression"", ""else_goto"": ""stepId"" }" & vbCrLf &
        "- if: Condition expression (see Condition Syntax below)" & vbCrLf &
        "- If condition FALSE: step is SKIPPED (treated as success, not error)" & vbCrLf &
        "- else_goto: Optional. If condition FALSE and else_goto set, jump to that step" & vbCrLf &
        "" & vbCrLf &
        "WAIT_FOR OBJECT:" & vbCrLf &
        "{ ""type"": ""time|url|selector"", ... }" & vbCrLf &
        "- type=""time"": { ""timeout_ms"": 2000 } - Delay BEFORE step execution" & vbCrLf &
        "- type=""url"": { ""value"": ""substring"" } - AFTER step, logs warning if URL doesn't contain substring" & vbCrLf &
        "- type=""selector"": { ""selector"": {...} } - AFTER step, logs warning if selector finds nothing" & vbCrLf &
        "" & vbCrLf &
        "=== TEMPLATE EXPANSION ===" & vbCrLf &
        "Placeholders {{varName}} are expanded in most string parameters BEFORE execution." & vbCrLf &
        "" & vbCrLf &
        "RESOLUTION ORDER:" & vbCrLf &
        "1. Check for special prefixes:" & vbCrLf &
        "   - {{env.VARNAME}} -> System.Environment.GetEnvironmentVariable" & vbCrLf &
        "   - {{env.DESKTOP}} -> Special case: user's Desktop folder path" & vbCrLf &
        "   - {{base_url}} -> env.base_url value" & vbCrLf &
        "2. Look up in _vars dictionary (case-insensitive)" & vbCrLf &
        "3. For nested paths like {{object.prop.sub}}, navigate through JToken or object properties" & vbCrLf &
        "4. If unresolved, placeholder remains as {{...}} in output (logged as warning)" & vbCrLf &
        "" & vbCrLf &
        "IMPORTANT: Unresolved placeholders are NOT errors - they remain literal. Check logs for warnings." & vbCrLf &
        "" & vbCrLf &
        "=== CONDITION SYNTAX ===" & vbCrLf &
        "Used in guard.if and if command's condition parameter." & vbCrLf &
        "" & vbCrLf &
        "EXISTENCE CHECK:" & vbCrLf &
        "  exists {{var}}   -> True if var exists and is non-empty" & vbCrLf &
        "" & vbCrLf &
        "EQUALITY:" & vbCrLf &
        "  {{var}} == ""literal""    -> Case-insensitive string comparison" & vbCrLf &
        "  {{var}} == true          -> Boolean comparison (handles string ""true""/""false"", ""1""/""0"")" & vbCrLf &
        "  {{var}} == false" & vbCrLf &
        "  {{var}} == []            -> True if null, empty string, or empty enumerable" & vbCrLf &
        "" & vbCrLf &
        "NUMERIC COMPARISONS:" & vbCrLf &
        "  {{var}} > 10             -> Greater than" & vbCrLf &
        "  {{var}} >= 10            -> Greater than or equal" & vbCrLf &
        "  {{var}} < 10             -> Less than" & vbCrLf &
        "  {{var}} <= 10            -> Less than or equal" & vbCrLf &
        "  {{var}} >= {{other}}     -> Compare two variables" & vbCrLf &
        "" & vbCrLf &
        "STRING OPERATIONS:" & vbCrLf &
        "  {{var}} contains ""text"" -> Case-insensitive substring search" & vbCrLf &
        "  {{var}} ~= ""regex""      -> Regex match (IgnoreCase, Singleline)" & vbCrLf &
        "" & vbCrLf &
        "LOGICAL OPERATORS:" & vbCrLf &
        "  expr1 || expr2           -> OR: returns True on first True (left-to-right)" & vbCrLf &
        "  expr1 && expr2           -> AND: returns False on first False (left-to-right)" & vbCrLf &
        "" & vbCrLf &
        "TRUTHINESS RULES:" & vbCrLf &
        "- False: null, empty string, ""false"", ""0"", ""null"", ""none"", ""nil"", empty enumerable" & vbCrLf &
        "- True: everything else (including non-empty strings, non-zero numbers, non-empty collections)" & vbCrLf &
        "" & vbCrLf &
        "=== SELECTOR OBJECT ===" & vbCrLf &
        "Used in extract_text, extract_html, wait_for selector." & vbCrLf &
        "" & vbCrLf &
        "{" & vbCrLf &
        "  ""strategy"": ""xpath|css|text|regex""," & vbCrLf &
        "  ""value"": ""selector expression""," & vbCrLf &
        "  ""within"": { /* nested SelectorObject for scoping */ }," & vbCrLf &
        "  ""relative"": { ""position"": ""first|last|nth"", ""nth"": 2 }" & vbCrLf &
        "}" & vbCrLf &
        "" & vbCrLf &
        "STRATEGIES:" & vbCrLf &
        "- xpath: Raw XPath expression (e.g., ""//div[@class='content']//a"")" & vbCrLf &
        "- css: Limited CSS selector support:" & vbCrLf &
        "  - Tag names: div, span, a" & vbCrLf &
        "  - Classes: .className" & vbCrLf &
        "  - IDs: #elementId" & vbCrLf &
        "  - Attributes: [attr=value]" & vbCrLf &
        "  - Pseudo: :nth-child(n)" & vbCrLf &
        "  - Combinators: space (descendant), > (direct child)" & vbCrLf &
        "- text: Substring match in element text (case-insensitive)" & vbCrLf &
        "  - ""exact:literal"" prefix for exact match" & vbCrLf &
        "- regex: Regex pattern matched against normalized inner text" & vbCrLf &
        "" & vbCrLf &
        "RELATIVE POSITIONING:" & vbCrLf &
        "- first: Return only first match" & vbCrLf &
        "- last: Return only last match" & vbCrLf &
        "- nth: Return nth match (1-based index)" & vbCrLf &
        "" & vbCrLf &
        "=== IMPLICIT VARIABLES ===" & vbCrLf &
        "These are automatically set by the interpreter:" & vbCrLf &
        "" & vbCrLf &
        "HTTP RESPONSE:" & vbCrLf &
        "- last_http_status: Integer status code from last open_url/http_request" & vbCrLf &
        "- last_http_elapsed_ms: Response time in milliseconds" & vbCrLf &
        "" & vbCrLf &
        "LLM RESULTS (after llm_analyze):" & vbCrLf &
        "- lastLlm: Parsed/sanitized JSON response with metadata (step_id, page_url)" & vbCrLf &
        "- lastLlm_raw: Raw LLM output text before sanitization" & vbCrLf &
        "- lastLlm_page_url: URL that was current when LLM was called" & vbCrLf &
        "- lastLlm_latency_ms: LLM call duration" & vbCrLf &
        "- last_step_id: ID of the most recently executed step" & vbCrLf &
        "" & vbCrLf &
        "LINK EXTRACTION (after open_url):" & vbCrLf &
        "- auto_links: List of extracted anchor URLs matching patterns" & vbCrLf &
        "- auto_link_enable: Boolean to enable/disable (default true)" & vbCrLf &
        "- auto_link_patterns: List of regex patterns for filtering links" & vbCrLf &
        "- auto_link_min: Minimum href length (default 15)" & vbCrLf &
        "" & vbCrLf &
        "=== COMMAND REFERENCE ===" & vbCrLf &
        "Command names are CASE-INSENSITIVE." & vbCrLf &
        "" & vbCrLf &
        "--- HTTP COMMANDS ---" & vbCrLf &
        "" & vbCrLf &
        "open_url: Load a page and parse HTML" & vbCrLf &
        "  params: {" & vbCrLf &
        "    ""url"": ""https://... or /relative"",  // REQUIRED. Template-expanded." & vbCrLf &
        "    ""method"": ""GET|POST|..."",          // Default: GET" & vbCrLf &
        "    ""headers"": { ""name"": ""value"" },    // Per-request headers" & vbCrLf &
        "    ""body"": ""content or object"",       // Request body" & vbCrLf &
        "    ""body_type"": ""json|form|raw"",      // How to encode body" & vbCrLf &
        "    ""return_body"": true,                // Include body in result" & vbCrLf &
        "    ""timeout_ms"": 15000,                // Override timeout" & vbCrLf &
        "    ""retry"": { ""max"":2, ""delay_ms"":1000, ""backoff"":2 }  // Inline retry" & vbCrLf &
        "  }" & vbCrLf &
        "  result: { status, url, elapsed_ms, body? }" & vbCrLf &
        "  SIDE EFFECTS: Sets DOM for selectors, auto-extracts links, sets lastResponseUrl/Body" & vbCrLf &
        "" & vbCrLf &
        "http_request: Generic HTTP request (no auto link extraction)" & vbCrLf &
        "  params: {" & vbCrLf &
        "    ""url"": ""..."",                      // REQUIRED" & vbCrLf &
        "    ""method"": ""GET|POST|PUT|DELETE"",  // Default: GET" & vbCrLf &
        "    ""headers"": { },                     // Per-request headers" & vbCrLf &
        "    ""query"": { ""key"": ""value"" },      // Query parameters (URL-encoded)" & vbCrLf &
        "    ""body"": ...,                        // Request body" & vbCrLf &
        "    ""body_type"": ""json|form|raw"",     // Encoding" & vbCrLf &
        "    ""timeout_ms"": 15000" & vbCrLf &
        "  }" & vbCrLf &
        "  result: { status, headers, body, url }" & vbCrLf &
        "  SIDE EFFECTS: Sets DOM for body content" & vbCrLf &
        "" & vbCrLf &
        "download_url: Download file to disk" & vbCrLf &
        "  params: {" & vbCrLf &
        "    ""url"": ""..."",                      // REQUIRED" & vbCrLf &
        "    ""target_dir"": ""C:\\path\\to\\dir"",  // REQUIRED. Created if missing." & vbCrLf &
        "    ""filename"": ""file.pdf"",            // Default: download.bin" & vbCrLf &
        "    ""method"": ""GET"",                   // HTTP method" & vbCrLf &
        "    ""headers"": { }," & vbCrLf &
        "    ""body"": ...," & vbCrLf &
        "    ""body_type"": ""json|form|raw""" & vbCrLf &
        "  }" & vbCrLf &
        "  result: { path, status }" & vbCrLf &
        "" & vbCrLf &
        "set_user_agent: Change HTTP User-Agent" & vbCrLf &
        "  params: { ""user_agent"": ""Custom Agent/1.0"" }" & vbCrLf &
        "  result: { user_agent }" & vbCrLf &
        "" & vbCrLf &
        "set_headers: Configure default headers" & vbCrLf &
        "  params: {" & vbCrLf &
        "    ""mode"": ""merge|replace"",          // merge=add/overwrite, replace=clear first" & vbCrLf &
        "    ""headers"": { ""Authorization"": ""Bearer ..."" }" & vbCrLf &
        "  }" & vbCrLf &
        "  result: { headers }" & vbCrLf &
        "" & vbCrLf &
        "set_cookies: Add cookies to jar" & vbCrLf &
        "  params: {" & vbCrLf &
        "    ""cookies"": [" & vbCrLf &
        "      { ""name"": ""session"", ""value"": ""abc123"", ""domain"": "".example.com""," & vbCrLf &
        "        ""path"": ""/"", ""secure"": true, ""httpOnly"": true }" & vbCrLf &
        "    ]" & vbCrLf &
        "  }" & vbCrLf &
        "  result: { count }" & vbCrLf &
        "" & vbCrLf &
        "--- EXTRACTION COMMANDS ---" & vbCrLf &
        "" & vbCrLf &
        "extract_text: Extract text from DOM elements" & vbCrLf &
        "  params: {" & vbCrLf &
        "    ""selector"": { /* SelectorObject */ },  // REQUIRED" & vbCrLf &
        "    ""all"": false,                         // false=first match, true=all matches as list" & vbCrLf &
        "    ""normalize_whitespace"": true,         // Collapse whitespace" & vbCrLf &
        "    ""regex"": ""pattern"",                  // Optional regex to extract portion" & vbCrLf &
        "    ""group"": 0                            // Regex capture group index" & vbCrLf &
        "  }" & vbCrLf &
        "  result: String (all=false) or List<String> (all=true)" & vbCrLf &
        "" & vbCrLf &
        "extract_html: Extract HTML from DOM element" & vbCrLf &
        "  params: {" & vbCrLf &
        "    ""selector"": { /* SelectorObject */ }," & vbCrLf &
        "    ""outer"": false                       // false=innerHTML, true=outerHTML" & vbCrLf &
        "  }" & vbCrLf &
        "  result: String (HTML content)" & vbCrLf &
        "" & vbCrLf &
        "extract_attribute: Extract attribute from serialized nodes" & vbCrLf &
        "  params: {" & vbCrLf &
        "    ""nodes_var"": ""varWithNodeList"",     // Variable holding serialized nodes" & vbCrLf &
        "    ""attribute"": ""href"",                // Attribute name to extract" & vbCrLf &
        "    ""var"": ""targetVariable""             // Where to store results" & vbCrLf &
        "  }" & vbCrLf &
        "  result: null (stores List<String> in target variable)" & vbCrLf &
        "  NOTE: nodes_var must contain objects with 'attributes' property" & vbCrLf &
        "" & vbCrLf &
        "find: Search for substring in variable" & vbCrLf &
        "  params: {" & vbCrLf &
        "    ""in"": ""variableName"",               // Variable to search" & vbCrLf &
        "    ""text"": ""needle"",                   // Substring to find (case-insensitive)" & vbCrLf &
        "    ""assign"": { ""var"": ""resultVar"" }   // Optional: store boolean result" & vbCrLf &
        "  }" & vbCrLf &
        "  result: { found: Boolean, index: Int }" & vbCrLf &
        "" & vbCrLf &
        "--- VARIABLE COMMANDS ---" & vbCrLf &
        "" & vbCrLf &
        "set_var: Set or update a variable" & vbCrLf &
        "  params: {" & vbCrLf &
        "    ""name"": ""variableName"",             // REQUIRED" & vbCrLf &
        "    ""value"": ""any JSON value""           // String values are template-expanded" & vbCrLf &
        "  }" & vbCrLf &
        "  result: { name, value }" & vbCrLf &
        "" & vbCrLf &
        "increment: Increment/decrement numeric variable" & vbCrLf &
        "  params: {" & vbCrLf &
        "    ""var"": ""counterName"",               // REQUIRED" & vbCrLf &
        "    ""by"": 1,                             // Amount to add (negative to subtract)" & vbCrLf &
        "    ""set_to"": 0                          // Optional: set absolute value instead" & vbCrLf &
        "  }" & vbCrLf &
        "  result: { var, old_value, new_value }" & vbCrLf &
        "" & vbCrLf &
        "array_push: Append item to array variable" & vbCrLf &
        "  params: {" & vbCrLf &
        "    ""array"": ""arrayVarName"",            // REQUIRED. Created if missing." & vbCrLf &
        "    ""item_var"": ""varToAppend"",          // Use existing variable value" & vbCrLf &
        "    ""item"": { /* inline value */ }      // OR specify inline value" & vbCrLf &
        "  }" & vbCrLf &
        "  result: { pushed, count, array }" & vbCrLf &
        "" & vbCrLf &
        "range: Generate integer array" & vbCrLf &
        "  params: {" & vbCrLf &
        "    ""var"": ""rangeVar"",                  // REQUIRED: target variable" & vbCrLf &
        "    ""from"": 0,                           // Start value (default 0)" & vbCrLf &
        "    ""to"": 10,                            // End value (exclusive)" & vbCrLf &
        "    ""step"": 1                            // Increment (default 1)" & vbCrLf &
        "  }" & vbCrLf &
        "  result: { var, count }" & vbCrLf &
        "" & vbCrLf &
        "--- CONTROL FLOW COMMANDS ---" & vbCrLf &
        "" & vbCrLf &
        "if: Conditional execution" & vbCrLf &
        "  params: {" & vbCrLf &
        "    ""condition"": ""condition expression"", // See Condition Syntax" & vbCrLf &
        "    ""steps"": [ /* steps if true */ ],    // REQUIRED" & vbCrLf &
        "    ""else_steps"": [ /* steps if false */ ] // Optional" & vbCrLf &
        "  }" & vbCrLf &
        "  result: null" & vbCrLf &
        "  NOTE: Sub-steps share same variable scope as parent" & vbCrLf &
        "" & vbCrLf &
        "foreach: Iterate over array" & vbCrLf &
        "  params: {" & vbCrLf &
        "    ""list"": ""arrayVarName"",             // REQUIRED: variable holding array" & vbCrLf &
        "    ""item_var"": ""itemName"",             // REQUIRED: loop variable name" & vbCrLf &
        "    ""steps"": [ /* steps per item */ ],  // REQUIRED" & vbCrLf &
        "    ""continue_on_error"": true,          // Default true: continue on item error" & vbCrLf &
        "    ""stop_on_error"": false,             // If true, overrides continue_on_error" & vbCrLf &
        "    ""max_items"": 100,                   // Limit iterations" & vbCrLf &
        "    ""break_if_var_true"": ""doneFlag""    // Exit if variable becomes truthy" & vbCrLf &
        "  }" & vbCrLf &
        "  result: { count, executed }" & vbCrLf &
        "  LOOP VARIABLES: item_var (current item), item_var_index (0-based index)" & vbCrLf &
        "  NOTE: If list variable missing, logs warning and returns {count:0, executed:0}" & vbCrLf &
        "" & vbCrLf &
        "while: Loop while condition true" & vbCrLf &
        "  params: {" & vbCrLf &
        "    ""condition"": ""condition expression"", // Re-evaluated each iteration" & vbCrLf &
        "    ""steps"": [ /* loop body */ ]," & vbCrLf &
        "    ""max_iterations"": 100,              // Safety limit (default 100)" & vbCrLf &
        "    ""break_if_var_true"": ""doneFlag""    // Early exit condition" & vbCrLf &
        "  }" & vbCrLf &
        "  result: { iterations, executed }" & vbCrLf &
        "" & vbCrLf &
        "wait: Pause execution" & vbCrLf &
        "  params: { ""ms"": 2000 }                // Milliseconds to wait" & vbCrLf &
        "  result: { slept }" & vbCrLf &
        "" & vbCrLf &
        "log: Write to execution log" & vbCrLf &
        "  params: {" & vbCrLf &
        "    ""level"": ""info|warn|error"",        // Log level" & vbCrLf &
        "    ""message"": ""Log message with {{vars}}""  // Template-expanded" & vbCrLf &
        "  }" & vbCrLf &
        "  result: null" & vbCrLf &
        "" & vbCrLf &
        "--- FILE COMMANDS ---" & vbCrLf &
        "" & vbCrLf &
        "save_file: Write content to file" & vbCrLf &
        "  params: {" & vbCrLf &
        "    ""path"": ""C:\\path\\file.txt"",        // REQUIRED. Directories created." & vbCrLf &
        "    ""content"": ""text or base64"",        // Content to write" & vbCrLf &
        "    ""encoding"": ""utf8|binary""           // binary requires Base64 content" & vbCrLf &
        "  }" & vbCrLf &
        "  result: { path }" & vbCrLf &
        "" & vbCrLf &
        "read_file: Read file content" & vbCrLf &
        "  params: {" & vbCrLf &
        "    ""path"": ""C:\\path\\file.txt"",        // REQUIRED" & vbCrLf &
        "    ""encoding"": ""utf8|binary""           // binary returns Base64" & vbCrLf &
        "  }" & vbCrLf &
        "  result: String (file content)" & vbCrLf &
        "" & vbCrLf &
        "delete_file: Delete file" & vbCrLf &
        "  params: { ""path"": ""C:\\path\\file.txt"" }" & vbCrLf &
        "  result: Boolean (true if deleted, false if not found)" & vbCrLf &
        "" & vbCrLf &
        "--- TEMPLATE & REPORT COMMANDS ---" & vbCrLf &
        "" & vbCrLf &
        "template: Render Mustache-like template" & vbCrLf &
        "  params: {" & vbCrLf &
        "    ""template"": ""Hello {{name}}!"",      // Template text" & vbCrLf &
        "    ""context"": { ""name"": ""World"" }     // Context for rendering" & vbCrLf &
        "  }" & vbCrLf &
        "  result: String (rendered output)" & vbCrLf &
        "  SYNTAX:" & vbCrLf &
        "  - {{var}} - Variable substitution" & vbCrLf &
        "  - {{{var}}} - Raw (unescaped) substitution" & vbCrLf &
        "  - {{#section}}...{{/section}} - Repeat for array, show if truthy" & vbCrLf &
        "  - {{^section}}...{{/section}} - Inverted: show if falsy/empty" & vbCrLf &
        "  NOTE: Second pass expands global _vars placeholders" & vbCrLf &
        "" & vbCrLf &
        "render_report: Generate final Markdown report" & vbCrLf &
        "  params: {" & vbCrLf &
        "    ""template"": ""# Report\\n\\n{{data}}"", // Mustache template" & vbCrLf &
        "    ""context"": { },                       // Optional context" & vbCrLf &
        "    ""output_path"": ""{{env.DESKTOP}}/report.md""  // Optional: save to file" & vbCrLf &
        "  }" & vbCrLf &
        "  result: { output }" & vbCrLf &
        "  SIDE EFFECT: Sets _finalMarkdown (returned as script output)" & vbCrLf &
        "  NOTE: If output_path contains unresolved {{...}}, file is NOT written" & vbCrLf &
        "" & vbCrLf &
        "send_email_report: Send email via SMTP" & vbCrLf &
        "  params: {" & vbCrLf &
        "    ""to"": ""user@example.com; other@example.com"",  // Semicolon/comma separated" & vbCrLf &
        "    ""subject"": ""Report Subject""," & vbCrLf &
        "    ""body_markdown"": ""# Email Body\\n\\nContent..."",  // Converted to HTML" & vbCrLf &
        "    ""smtp_host"": ""smtp.example.com"",    // REQUIRED" & vbCrLf &
        "    ""smtp_port"": 25," & vbCrLf &
        "    ""smtp_ssl"": ""true|false""," & vbCrLf &
        "    ""smtp_auth"": ""true|false""," & vbCrLf &
        "    ""smtp_user"": ""username""," & vbCrLf &
        "    ""smtp_pass"": ""password""," & vbCrLf &
        "    ""from_email"": ""sender@example.com""," & vbCrLf &
        "    ""from_name"": ""Sender Name""" & vbCrLf &
        "  }" & vbCrLf &
        "  result: Boolean (success)" & vbCrLf &
        "  NOTE: Sends multipart/alternative (text + HTML). Footer auto-added." & vbCrLf &
        "" & vbCrLf &
        "--- LLM COMMAND ---" & vbCrLf &
        "" & vbCrLf &
        "llm_analyze (aliases: llm, llmanalyze): Call LLM API" & vbCrLf &
        "  params: {" & vbCrLf &
        "    // PROMPTS (use one from each group):" & vbCrLf &
        "    ""system"": ""System prompt..."",       // OR ""systemPrompt""" & vbCrLf &
        "    ""user"": ""User prompt..."",           // OR ""prompt"", ""input"", ""arguments""" & vbCrLf &
        "    " & vbCrLf &
        "    // OPTIONS:" & vbCrLf &
        "    ""temperature"": ""0.7"",               // String or number" & vbCrLf &
        "    ""timeoutMs"": 60000,                 // Override timeout" & vbCrLf &
        "    ""status_var"": ""httpStatus"",         // Skip if var equals ""404"" (unless allow_llm_on_404)" & vbCrLf &
        "    " & vbCrLf &
        "    // INNER RETRY (before step-level retry):" & vbCrLf &
        "    ""inner_attempts"": 3,                // Retry if invalid JSON (default 1)" & vbCrLf &
        "    ""inner_delay_ms"": 800,              // Delay between inner attempts" & vbCrLf &
        "    " & vbCrLf &
        "    // VALIDATION FLAGS:" & vbCrLf &
        "    ""retry_on_invalid"": true,           // Throw if invalid (triggers step retry)" & vbCrLf &
        "    ""reject_if_empty"": true,            // Empty output = invalid" & vbCrLf &
        "    ""reject_if_plaintext"": true,        // Non-JSON = invalid (default true)" & vbCrLf &
        "    ""allow_non_json"": false,            // Override reject_if_plaintext" & vbCrLf &
        "    ""require_key"": ""key1,key2"",        // Required top-level JSON keys" & vbCrLf &
        "    ""require_key_all"": true,            // ALL keys required (default true)" & vbCrLf &
        "    ""require_array_key"": ""items"",      // Key must be JSON array" & vbCrLf &
        "    ""require_min_items"": 1,             // Minimum array length" & vbCrLf &
        "    " & vbCrLf &
        "    // DEBUG:" & vbCrLf &
        "    ""log_raw"": true,                    // Dump raw output to debug log" & vbCrLf &
        "    ""max_preview"": 250                  // UI preview length" & vbCrLf &
        "  }" & vbCrLf &
        "  result: JObject with parsed response + metadata (step_id, page_url)" & vbCrLf &
        "  " & vbCrLf &
        "  SANITIZATION PROCESS:" & vbCrLf &
        "  1. Extract content from ```json...``` code blocks if present" & vbCrLf &
        "  2. Attempt full JSON parse" & vbCrLf &
        "  3. Extract first balanced {...} or [...] substring" & vbCrLf &
        "  4. Strip stray backticks" & vbCrLf &
        "  5. Validate against require_key/require_array_key rules" & vbCrLf &
        "  " & vbCrLf &
        "  VARIABLES SET: lastLlm, lastLlm_raw, lastLlm_page_url, lastLlm_latency_ms, last_step_id" & vbCrLf &
        "" & vbCrLf &
        "--- SPECIAL COMMANDS ---" & vbCrLf &
        "" & vbCrLf &
        "enable_dynamic: Enable dynamic content expansion" & vbCrLf &
        "  params: { } (ignored)" & vbCrLf &
        "  result: { status, dynamic }" & vbCrLf &
        "  EFFECT: After page load, scans for AJAX endpoints and fetches up to 10 additional resources" & vbCrLf &
        "" & vbCrLf &
        "=== URL RESOLUTION ===" & vbCrLf &
        "Relative URLs are resolved using this precedence:" & vbCrLf &
        "1. lastResponseUrl (if previous request made)" & vbCrLf &
        "2. env.base_url (if configured)" & vbCrLf &
        "3. URL used as-is (will fail if not absolute)" & vbCrLf &
        "" & vbCrLf &
        "URL SANITIZATION:" & vbCrLf &
        "- Markdown [text](url) format: extracts url portion" & vbCrLf &
        "- Angle brackets <url>: strips brackets" & vbCrLf &
        "" & vbCrLf &
        "=== DEBUG FLAGS ===" & vbCrLf &
        "Set via env.variables or set_var command:" & vbCrLf &
        "" & vbCrLf &
        "- debug: Enable debug logging to file (Desktop/RI_Debug_Webagent.txt)" & vbCrLf &
        "- debug_allAttempts: Log all retry attempts, not just final" & vbCrLf &
        "- debug_substeps: Log sub-step execution in foreach/if" & vbCrLf &
        "- debug_var_changes: Log variable value changes" & vbCrLf &
        "- debug_include_script: Log masked script JSON at start" & vbCrLf &
        "- debug_summary: Log final execution summary" & vbCrLf &
        "- debug_to_logwindow: Mirror debug to UI log window" & vbCrLf &
        "- debug_clear_llm_state: Clear lastLlm between non-LLM steps" & vbCrLf &
        "- allow_llm_on_404: Don't skip LLM if status_var is ""404""" & vbCrLf &
        "- llm_rethrow_all: Re-throw all LLM exceptions" & vbCrLf &
        "" & vbCrLf &
        "=== BEST PRACTICES ===" & vbCrLf &
        "1. Always include unique 'id' for every step (for debugging and jumps)" & vbCrLf &
        "2. Use retry + on_error for network operations" & vbCrLf &
        "3. Initialize variables in env.variables before referencing" & vbCrLf &
        "4. Use require_key/require_array_key for LLM validation" & vbCrLf &
        "5. Set base_url or use absolute URLs to avoid resolution issues" & vbCrLf &
        "6. Prefer 'if' command over 'guard' for multi-step conditional blocks" & vbCrLf &
        "7. Use assign.path to extract specific values from complex results" & vbCrLf &
        "8. Check that list variables exist before foreach (missing = silent skip)" & vbCrLf &
        "" & vbCrLf &
        "=== COMMON MISTAKES ===" & vbCrLf &
        "- Unresolved {{placeholders}} are NOT errors - check logs for warnings" & vbCrLf &
        "- foreach with missing list silently returns {count:0, executed:0}" & vbCrLf &
        "- extract_attribute needs serialized node objects, not raw HTML" & vbCrLf &
        "- Binary file operations require Base64 encoding" & vbCrLf &
        "- Guard skip is success, not error - flow continues normally" & vbCrLf &
        "- render_report's output_path skips write if unresolved placeholders remain" & vbCrLf &
        "" & vbCrLf &
        "END OF SPECIFICATION"



    ''' <summary>
    ''' Parameter placeholder specification for user-defined runtime parameters in WebAgent scripts.
    ''' </summary>
    Public Const WebAgentParameterSpec As String =
        "WEBAGENT PARAMETER PLACEHOLDER SPECIFICATION" & vbCrLf &
        "=============================================" & vbCrLf &
        "" & vbCrLf &
        "PURPOSE: Enable runtime user input for WebAgent scripts. Parameters are defined inline " &
        "and replaced with user-provided values before script execution." & vbCrLf &
        "" & vbCrLf &
        "PROCESSING FLOW:" & vbCrLf &
        "1. Script loaded from file" & vbCrLf &
        "2. Regex scans for parameter DEFINITIONS: {parameterN=...}" & vbCrLf &
        "3. If definitions found, user prompted with dialog" & vbCrLf &
        "4. Values substituted into script (definitions AND references)" & vbCrLf &
        "5. Script proceeds to normal execution" & vbCrLf &
        "" & vbCrLf &
        "=== SYNTAX ===" & vbCrLf &
        "" & vbCrLf &
        "DEFINITION (where value should appear):" & vbCrLf &
        "  {parameterN=Description ; Type ; Default ; RangeOrOptions}" & vbCrLf &
        "" & vbCrLf &
        "REFERENCE (reuse same value elsewhere):" & vbCrLf &
        "  {parameterN}" & vbCrLf &
        "" & vbCrLf &
        "EXAMPLES:" & vbCrLf &
        "  ""base_url"": ""{parameter1=API URL ; string ; https://api.example.com}""" & vbCrLf &
        "  ""timeout"": {parameter2=Timeout seconds ; integer ; 30 ; 5-120}" & vbCrLf &
        "  ""mode"": ""{parameter3=Mode ; string ; prod ; prod<production>,dev<development>}""" & vbCrLf &
        "" & vbCrLf &
        "=== DEFINITION SEGMENTS ===" & vbCrLf &
        "Separated by semicolons, whitespace trimmed:" & vbCrLf &
        "" & vbCrLf &
        "[0] Description (REQUIRED)" & vbCrLf &
        "    - Shown in UI prompt" & vbCrLf &
        "    - Keep concise (<60 chars)" & vbCrLf &
        "" & vbCrLf &
        "[1] Type (optional, default: string)" & vbCrLf &
        "    - string: Any text value" & vbCrLf &
        "    - integer: Whole number" & vbCrLf &
        "    - long: Large whole number" & vbCrLf &
        "    - double: Decimal number" & vbCrLf &
        "    - boolean: true/false" & vbCrLf &
        "" & vbCrLf &
        "[2] Default (optional)" & vbCrLf &
        "    - Pre-filled value in UI" & vbCrLf &
        "    - For options: matches code value for pre-selection" & vbCrLf &
        "" & vbCrLf &
        "[3] Range OR Options (optional)" & vbCrLf &
        "    RANGE (numeric types only):" & vbCrLf &
        "      - Format: MIN-MAX (e.g., ""0-100"")" & vbCrLf &
        "      - Values clamped to range" & vbCrLf &
        "    " & vbCrLf &
        "    OPTIONS (any type):" & vbCrLf &
        "      - Comma-separated list" & vbCrLf &
        "      - Simple: ""opt1,opt2,opt3""" & vbCrLf &
        "      - With codes: ""Display Text<code>,Other<other>""" & vbCrLf &
        "      - UI shows display text, script receives code" & vbCrLf &
        "" & vbCrLf &
        "[4] Extra Options (optional)" & vbCrLf &
        "    - If [3] was a range, this adds option list" & vbCrLf &
        "    - Example: ""0-100 ; 25,50,75,100""" & vbCrLf &
        "" & vbCrLf &
        "=== DETAILED EXAMPLES ===" & vbCrLf &
        "" & vbCrLf &
        "STRING WITH OPTIONS (display<code> syntax):" & vbCrLf &
        "  {parameter1=Environment ; string ; prod ; " &
        "Production<https://api.prod.com>,Staging<https://api.staging.com>,Dev<http://localhost:8080>}" & vbCrLf &
        "  -> UI shows: Production, Staging, Dev" & vbCrLf &
        "  -> Script gets: https://api.prod.com (or selected URL)" & vbCrLf &
        "" & vbCrLf &
        "INTEGER WITH RANGE:" & vbCrLf &
        "  {parameter2=Max retries ; integer ; 3 ; 0-10}" & vbCrLf &
        "  -> Accepts 0-10, values outside clamped" & vbCrLf &
        "" & vbCrLf &
        "DOUBLE WITH PRESET OPTIONS:" & vbCrLf &
        "  {parameter3=Threshold ; double ; 0.75 ; 0.25,0.5,0.75,0.9,1.0}" & vbCrLf &
        "  -> Dropdown with common values" & vbCrLf &
        "" & vbCrLf &
        "BOOLEAN:" & vbCrLf &
        "  {parameter4=Enable debug ; boolean ; false}" & vbCrLf &
        "  -> Outputs: true or false (lowercase)" & vbCrLf &
        "" & vbCrLf &
        "SIMPLE STRING (no options):" & vbCrLf &
        "  {parameter5=Search term ; string ; default query}" & vbCrLf &
        "  -> Free text input" & vbCrLf &
        "" & vbCrLf &
        "=== REPLACEMENT RULES ===" & vbCrLf &
        "" & vbCrLf &
        "1. First definition for each N wins (duplicates ignored)" & vbCrLf &
        "2. UI prompts in ascending parameter number order" & vbCrLf &
        "3. After user confirms:" & vbCrLf &
        "   - Each {parameterN=...} replaced with final value" & vbCrLf &
        "   - Each {parameterN} replaced with same value" & vbCrLf &
        "4. Replacement proceeds from end of script backward (preserves positions)" & vbCrLf &
        "" & vbCrLf &
        "JSON ESCAPING:" & vbCrLf &
        "- Backslash: \\ -> \\\\" & vbCrLf &
        "- Quote: "" -> \\""" & vbCrLf &
        "- No other escaping applied" & vbCrLf &
        "" & vbCrLf &
        "EMPTY SELECTION:" & vbCrLf &
        "If user selects value starting with:" & vbCrLf &
        "  ""(no selection)"", ""(keine auswahl)"", or ""---""" & vbCrLf &
        "-> Empty string inserted" & vbCrLf &
        "" & vbCrLf &
        "=== PLACEMENT IN JSON ===" & vbCrLf &
        "" & vbCrLf &
        "CORRECT - Definition as complete string value:" & vbCrLf &
        "  ""base_url"": ""{parameter1=API URL ; string ; https://api.example.com}""" & vbCrLf &
        "  -> After: ""base_url"": ""https://api.example.com""" & vbCrLf &
        "" & vbCrLf &
        "CORRECT - Definition for numeric value (no quotes needed in result):" & vbCrLf &
        "  ""timeout"": {parameter2=Timeout ; integer ; 30}" & vbCrLf &
        "  -> After: ""timeout"": 30" & vbCrLf &
        "" & vbCrLf &
        "CORRECT - Reference after definition:" & vbCrLf &
        "  ""primary"": ""{parameter1=URL ; string ; https://example.com}""," & vbCrLf &
        "  ""secondary"": ""{parameter1}/backup""" & vbCrLf &
        "  -> After: ""primary"": ""https://example.com"", ""secondary"": ""https://example.com/backup""" & vbCrLf &
        "" & vbCrLf &
        "AVOID - Definition embedded in larger string:" & vbCrLf &
        "  ""url"": ""{parameter1=Base;string;https://api.com}/v1/items""" & vbCrLf &
        "  -> Works but less clear; prefer separate definition" & vbCrLf &
        "" & vbCrLf &
        "=== AUTHORING GUIDELINES ===" & vbCrLf &
        "" & vbCrLf &
        "1. Use sequential numbering: parameter1, parameter2, parameter3..." & vbCrLf &
        "2. Place definition where final value should appear" & vbCrLf &
        "3. Use references {parameterN} for reuse (not definitions)" & vbCrLf &
        "4. Provide sensible defaults for runnable scripts" & vbCrLf &
        "5. For string values inside JSON, wrap definition in quotes" & vbCrLf &
        "6. For numeric/boolean, omit quotes (bare definition)" & vbCrLf &
        "7. Keep descriptions concise but descriptive" & vbCrLf &
        "8. Use display<code> syntax when UI text differs from value" & vbCrLf &
        "" & vbCrLf &
        "=== COMMON MISTAKES ===" & vbCrLf &
        "" & vbCrLf &
        "INVALID - Missing description:" & vbCrLf &
        "  {parameter1=}" & vbCrLf &
        "" & vbCrLf &
        "INVALID - Non-numeric parameter number:" & vbCrLf &
        "  {parameterX=Description}" & vbCrLf &
        "" & vbCrLf &
        "PROBLEMATIC - Unquoted string in JSON context:" & vbCrLf &
        "  ""name"": {parameter1=Name ; string ; John}" & vbCrLf &
        "  -> Results in invalid JSON: ""name"": John" & vbCrLf &
        "  -> Correct: ""name"": ""{parameter1=Name ; string ; John}""" & vbCrLf &
        "" & vbCrLf &
        "PROBLEMATIC - Reference before definition:" & vbCrLf &
        "  References resolve only if definition exists somewhere in script" & vbCrLf &
        "" & vbCrLf &
        "=== CANCELLATION ===" & vbCrLf &
        "" & vbCrLf &
        "If user cancels the parameter dialog:" & vbCrLf &
        "- Script execution aborts" & vbCrLf &
        "- No changes applied to original script" & vbCrLf &
        "" & vbCrLf &
        "If no definitions found:" & vbCrLf &
        "- No dialog shown" & vbCrLf &
        "- Script executes immediately" & vbCrLf &
        "- Any {parameterN} references remain as literal text" & vbCrLf &
        "" & vbCrLf &
        "END OF PARAMETER SPECIFICATION"



End Class
