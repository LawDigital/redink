' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: HelpMeInky.vb
' Purpose: Provides a WinForms chat window ("Help me, AN8!") that lets the user ask questions and
'          forwards them to the configured LLM, enriched with a product manual and optional
'          configuration file content.
'
' Architecture:
'  - UI: A `WebBrowser` displays chat HTML; a multiline `TextBox` collects user input; buttons for
'    send/clear/close; checkboxes for "stay on top" and "include configuration files".
'  - Chat Rendering: A small HTML+CSS+JS document is injected into `WebBrowser`. Messages are appended
'    as HTML fragments; "Thinking..." is added/removed by DOM id.
'  - Conversation State: `_history` stores (role, content) tuples; transcript/HTML is persisted in
'    `My.Settings` and restored on next load.
'  - Welcome: On first open (or after clearing), a welcome message is generated via the LLM; a guard
'    (`_welcomeInProgress`) prevents concurrent welcomes.
'  - Manual Loading: Loads the manual text once per configured path/URL (`INI_HelpMeInkyPath`) with a
'    simple cache; supports local text formats and remote HTTP(S) with basic PDF detection.
'  - LLM Invocation: Uses `_context.SP_HelpMe` as a system prompt; optionally applies an alternate
'    model config via `GetSpecialTaskModel` and serializes requests via `_modelSemaphore`.
'  - Optional Config Enrichment: When enabled, reads known config files and redacts API keys before
'    appending them to the LLM prompt.
'
' Persistence:
'  - Window bounds, top-most preference, include-config preference, last transcript and last HTML are
'    stored in `My.Settings`.
'
' External Dependencies:
'  - Markdig: Markdown -> HTML rendering.
'  - HtmlAgilityPack: HTML -> plain text extraction for manual loading.
'  - SharedLibrary.SharedMethods: LLM invocation, PDF/RTF reading, config discovery, model switching.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Text
Imports System.Text.RegularExpressions
Imports SharedLibrary.SharedLibrary.SharedContext
Imports System.Net
Imports System.Threading
Imports Markdig
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports System.ComponentModel
Imports System.Drawing
Imports System.IO
Imports System.Net.Http
Imports System.Windows.Forms

Namespace SharedLibrary

    ''' <summary>
    ''' WinForms chat window that collects user questions, renders the conversation as HTML, and
    ''' invokes the configured LLM with manual/config context.
    ''' </summary>
    Public Class HelpMeInky
        Inherits System.Windows.Forms.Form

        ''' <summary>Window title template for this form.</summary>
        Private WindowTitle As String = $"Help me, {AN8}!"

        ''' <summary>Name shown in the chat UI for assistant messages.</summary>
        Private Const AssistantName As String = AN8

        ''' <summary>Shared configuration/context provider (prompts, INI paths, model settings, etc.).</summary>
        Private ReadOnly _context As ISharedContext

        ''' <summary>Host application name/version used for prompt enrichment (if provided).</summary>
        Private ReadOnly _hostAppName As String

        ''' <summary>Markdown pipeline used to render assistant responses into HTML.</summary>
        Private ReadOnly _mdPipeline As Markdig.MarkdownPipeline

        ''' <summary>
        ''' Browser control used to display the conversation. Hosts injected HTML with small JS helpers.
        ''' </summary>
        Private ReadOnly _chat As WebBrowser = New WebBrowser() With {
        .Dock = DockStyle.Fill,
        .AllowWebBrowserDrop = False,
        .IsWebBrowserContextMenuEnabled = True,
        .WebBrowserShortcutsEnabled = True,
        .ScriptErrorsSuppressed = True
    }

        ''' <summary>Text input for user messages.</summary>
        Private ReadOnly _txtInput As TextBox = New TextBox() With {
        .Dock = DockStyle.Fill,
        .Multiline = True,
        .AcceptsReturn = True,
        .WordWrap = True
    }

        ''' <summary>Clears conversation state and persisted chat.</summary>
        Private ReadOnly _btnClear As Button = New Button() With {.Text = "Clear", .AutoSize = True}

        ''' <summary>Closes the form.</summary>
        Private ReadOnly _btnClose As Button = New Button() With {.Text = "Close", .AutoSize = True}

        ''' <summary>Sends the current text input to the LLM.</summary>
        Private ReadOnly _btnSend As Button = New Button() With {.Text = $"Ask {AssistantName}", .AutoSize = True}

        ''' <summary>
        ''' When checked, disables TopMost behavior ("Do not stay on top").
        ''' </summary>
        Private ReadOnly _chkNoTopMost As System.Windows.Forms.CheckBox = New System.Windows.Forms.CheckBox() With {.Text = "Do not stay on top", .AutoSize = True}

        ''' <summary>
        ''' When checked, appends configuration files (with API key redaction) to the LLM prompt.
        ''' </summary>
        Private ReadOnly _chkIncludeConfig As System.Windows.Forms.CheckBox = New System.Windows.Forms.CheckBox() With {.Text = "Include configuration files", .AutoSize = True}

        ''' <summary>Indicates whether the chat HTML document finished loading and can be appended to.</summary>
        Private _htmlReady As Boolean = False

        ''' <summary>Queue for HTML fragments appended before the browser document is ready.</summary>
        Private ReadOnly _htmlQueue As New List(Of String)()

        ''' <summary>DOM id of the latest "Thinking..." placeholder message, if any.</summary>
        Private _lastThinkingId As String = Nothing

        ''' <summary>Conversation history used to build transcript and LLM context.</summary>
        Private ReadOnly _history As New List(Of (Role As String, Content As String))()

        ''' <summary>Cached manual text for the currently configured manual path/URL.</summary>
        Private _manualCache As String = Nothing

        ''' <summary>Path/URL that `_manualCache` was loaded from.</summary>
        Private _manualCachePath As String = Nothing

        ''' <summary>Welcome generation state flag (0 = none, 1 = running).</summary>
        Private _welcomeInProgress As Integer = 0 ' 0 = none, 1 = running

        ''' <summary>Indicates whether alternate HelpMe model availability has been evaluated.</summary>
        Private _helpMeAltResolved As Boolean = False

        ''' <summary>Indicates whether an alternate HelpMe model is available and will be used.</summary>
        Private _helpMeAltAvailable As Boolean = False

        ''' <summary>Serializes LLM calls to avoid concurrent model switching/use.</summary>
        Private ReadOnly _modelSemaphore As New Threading.SemaphoreSlim(1, 1)

        ''' <summary>
        ''' Creates the chat form and wires UI event handlers.
        ''' </summary>
        ''' <param name="context">Shared application context holding prompts and configuration.</param>
        ''' <param name="hostAppName">Host application name/version used for prompt enrichment.</param>
        Public Sub New(context As ISharedContext, Optional hostAppName As String = "")
            MyBase.New()
            _context = context
            _hostAppName = hostAppName

            Me.Text = WindowTitle
            Me.FormBorderStyle = FormBorderStyle.Sizable
            Me.StartPosition = FormStartPosition.Manual
            Me.MinimumSize = New System.Drawing.Size(720, 420)
            Me.Font = New System.Drawing.Font("Segoe UI", 9.0F)
            Try
                Me.Icon = Icon.FromHandle(New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard)).GetHicon())
            Catch
            End Try

            Dim table As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 3,
            .Padding = New Padding(10)
        }
            table.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            table.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            table.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            table.RowStyles.Add(New RowStyle(SizeType.AutoSize))

            _txtInput.Margin = New Padding(0, 10, 0, 6)
            Dim threeLines = (_txtInput.Font.Height * 3) + 10
            _txtInput.MinimumSize = New System.Drawing.Size(0, threeLines)
            _txtInput.Height = threeLines

            Dim pnlButtons As New FlowLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .FlowDirection = FlowDirection.LeftToRight,
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .Padding = New Padding(0, 0, 0, 4)
        }
            pnlButtons.Controls.Add(_btnSend)
            pnlButtons.Controls.Add(_btnClear)
            pnlButtons.Controls.Add(_btnClose)
            pnlButtons.Controls.Add(_chkNoTopMost)
            pnlButtons.Controls.Add(_chkIncludeConfig)

            table.Controls.Add(_chat, 0, 0)
            table.Controls.Add(_txtInput, 0, 1)
            table.Controls.Add(pnlButtons, 0, 2)
            Me.Controls.Add(table)

            _mdPipeline = New MarkdownPipelineBuilder().
            UseAdvancedExtensions().
            UseEmojiAndSmiley().
            UseSoftlineBreakAsHardlineBreak().
            Build()

            AddHandler Me.Load, AddressOf OnLoadForm
            AddHandler Me.FormClosing, AddressOf OnFormClosing
            AddHandler Me.Activated, AddressOf OnActivated ' ensure top-most behavior reapplied
            AddHandler _btnSend.Click, AddressOf OnSend
            AddHandler _btnClear.Click, AddressOf OnClear
            AddHandler _btnClose.Click, AddressOf OnClose
            AddHandler _txtInput.KeyDown, AddressOf OnInputKeyDown
            AddHandler _chat.DocumentCompleted, AddressOf Chat_DocumentCompleted
            AddHandler _chat.Navigating, AddressOf Chat_Navigating
            AddHandler _chat.NewWindow, AddressOf Chat_NewWindow
            AddHandler _chkNoTopMost.CheckedChanged, AddressOf OnTopMostChanged
            AddHandler _chkIncludeConfig.CheckedChanged, AddressOf OnIncludeConfigChanged
        End Sub

        ''' <summary>Writes a diagnostic message to the debug output.</summary>
        Private Sub Dbg(msg As String)
            Debug.WriteLine($"[HelpMeInky {DateTime.Now:HH:mm:ss.fff}] {msg}")
        End Sub

        ''' <summary>
        ''' Executes an action on the UI thread (no-op if the form is disposed).
        ''' </summary>
        Private Sub Ui(action As Action)
            If Me.IsDisposed Then Return
            If Me.InvokeRequired Then
                Try : Me.BeginInvoke(action) : Catch : End Try
            Else
                action()
            End If
        End Sub

        ''' <summary>
        ''' Shows or raises the window, restores from minimized state, and focuses the input box.
        ''' </summary>
        Public Sub ShowRaised(Optional owner As IWin32Window = Nothing)
            Dbg("ShowRaised")
            If Me.WindowState = FormWindowState.Minimized Then Me.WindowState = FormWindowState.Normal
            If Not Me.Visible Then
                If owner IsNot Nothing Then Me.Show(owner) Else Me.Show()
            End If
            Me.Activate()
            ApplyTopMostBehavior()
            _txtInput.Focus()
            _txtInput.SelectAll()
        End Sub

        ''' <summary>Re-applies TopMost behavior when the form becomes active.</summary>
        Private Sub OnActivated(sender As Object, e As EventArgs)
            ApplyTopMostBehavior()
        End Sub

        ''' <summary>Persists the TopMost preference and applies the behavior immediately.</summary>
        Private Sub OnTopMostChanged(sender As Object, e As EventArgs)
            Try
                My.Settings.HelpMeNoTopMost = _chkNoTopMost.Checked
                My.Settings.Save()
            Catch
            End Try
            ApplyTopMostBehavior()
        End Sub

        ''' <summary>Persists the include-config preference.</summary>
        Private Sub OnIncludeConfigChanged(sender As Object, e As EventArgs)
            Try
                My.Settings.HelpMeIncludeConfig = _chkIncludeConfig.Checked
                My.Settings.Save()
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' Applies TopMost based on `_chkNoTopMost`.
        ''' </summary>
        Private Sub ApplyTopMostBehavior()
            ' If unchecked => stay always on top
            If _chkNoTopMost IsNot Nothing AndAlso _chkNoTopMost.Checked Then
                Me.TopMost = False
            Else
                Me.TopMost = True
            End If
        End Sub

        ''' <summary>
        ''' Loads persisted window settings, initializes the chat HTML document, restores persisted chat
        ''' content, and generates a welcome message when needed.
        ''' </summary>
        Private Async Sub OnLoadForm(sender As Object, e As EventArgs)
            Dbg("OnLoadForm start")
            Try
                If My.Settings.HelpMeFormLocation <> System.Drawing.Point.Empty AndAlso My.Settings.HelpMeFormSize <> System.Drawing.Size.Empty Then
                    Me.Location = My.Settings.HelpMeFormLocation
                    Me.Size = My.Settings.HelpMeFormSize
                Else
                    Dim area = Screen.PrimaryScreen.WorkingArea
                    Dim w = Math.Max(Me.MinimumSize.Width, 820)
                    Dim h = Math.Max(Me.MinimumSize.Height, 500)
                    Me.Location = New System.Drawing.Point(area.Left + (area.Width - w) \ 2, area.Top + (area.Height - h) \ 2)
                    Me.Size = New System.Drawing.Size(w, h)
                End If
            Catch ex As Exception
                Dbg("Restore bounds error: " & ex.Message)
            End Try

            ' Load persisted TopMost choice (default False => window stays on top)
            Try
                _chkNoTopMost.Checked = My.Settings.HelpMeNoTopMost
            Catch
                _chkNoTopMost.Checked = False
            End Try

            ' Load persisted IncludeConfig choice
            Try
                _chkIncludeConfig.Checked = My.Settings.HelpMeIncludeConfig
            Catch
                _chkIncludeConfig.Checked = False
            End Try

            ApplyTopMostBehavior()

            InitializeChatHtml()

            If Not String.IsNullOrEmpty(My.Settings.LastHelpMeChatHtml) Then
                AppendHtml(My.Settings.LastHelpMeChatHtml)
            ElseIf Not String.IsNullOrEmpty(My.Settings.LastHelpMeChat) Then
                AppendTranscriptToHtml(My.Settings.LastHelpMeChat)
            Else
                Await SafeGenerateWelcomeAsync()
            End If
        End Sub

        ''' <summary>
        ''' Persists transcript/HTML and window settings when the form closes.
        ''' </summary>
        Private Sub OnFormClosing(sender As Object, e As FormClosingEventArgs)
            Dbg("OnFormClosing")
            Try
                PersistTranscriptLimited()
                PersistChatHtml()
                If Me.WindowState = FormWindowState.Normal Then
                    My.Settings.HelpMeFormLocation = Me.Location
                    My.Settings.HelpMeFormSize = Me.Size
                Else
                    My.Settings.HelpMeFormLocation = Me.RestoreBounds.Location
                    My.Settings.HelpMeFormSize = Me.RestoreBounds.Size
                End If
                My.Settings.HelpMeNoTopMost = _chkNoTopMost.Checked
                My.Settings.HelpMeIncludeConfig = _chkIncludeConfig.Checked
                My.Settings.Save()
            Catch ex As Exception
                Dbg("OnFormClosing error: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Sends the current input text: renders user text, appends it to history, shows "Thinking...",
        ''' and starts LLM invocation asynchronously.
        ''' </summary>
        Private Sub OnSend(sender As Object, e As EventArgs)
            Dim userText = _txtInput.Text.Trim()
            Dbg($"OnSend len={userText.Length}")
            If userText.Length = 0 Then Return

            AppendUserHtml(userText)
            _history.Add(("user", userText))
            _txtInput.Clear()
            ShowAssistantThinking()
            Dim __ = SendAsync(userText)
        End Sub

        ''' <summary>
        ''' Clears the conversation state, resets persisted chat, and then generates a new welcome.
        ''' </summary>
        Private Async Sub OnClear(sender As Object, e As EventArgs)
            Dbg("OnClear start")
            Try
                _history.Clear()
                InitializeChatHtml()
                My.Settings.LastHelpMeChat = ""
                My.Settings.LastHelpMeChatHtml = ""
                My.Settings.Save()
                Dbg("State cleared & saved")
                Await SafeGenerateWelcomeAsync().ConfigureAwait(False)
            Catch ex As Exception
                Dbg("OnClear error: " & ex.Message)
            Finally
                If _txtInput.InvokeRequired Then
                    _txtInput.BeginInvoke(Sub() _txtInput.Focus())
                Else
                    _txtInput.Focus()
                End If
            End Try
        End Sub

        ''' <summary>Closes the form.</summary>
        Private Sub OnClose(sender As Object, e As EventArgs)
            Dbg("OnClose")
            Me.Close()
        End Sub

        ''' <summary>
        ''' Handles keyboard shortcuts in the input box (Enter sends; Escape closes).
        ''' </summary>
        Private Sub OnInputKeyDown(sender As Object, e As KeyEventArgs)
            If e.KeyCode = Keys.Enter Then
                e.SuppressKeyPress = True
                OnSend(Me, EventArgs.Empty)
            ElseIf e.KeyCode = Keys.Escape Then
                Me.Close()
            End If
        End Sub

        ''' <summary>
        ''' Generates the welcome message, ensuring only one welcome is generated concurrently.
        ''' </summary>
        Private Async Function SafeGenerateWelcomeAsync() As Task
            If Interlocked.CompareExchange(_welcomeInProgress, 1, 0) <> 0 Then
                Dbg("SafeGenerateWelcomeAsync skipped: already running")
                Exit Function
            End If
            Try
                Await GenerateWelcomeAsync()
            Catch ex As Exception
                Dbg("SafeGenerateWelcomeAsync fatal: " & ex.Message)
                RemoveAssistantThinking()
                AppendAssistantMarkdown("*(Welcome failed: " & System.Security.SecurityElement.Escape(ex.Message) & ")*")
            Finally
                Interlocked.Exchange(_welcomeInProgress, 0)
            End Try
        End Function

        ''' <summary>
        ''' Builds the welcome prompt, invokes the LLM, appends the response, and persists chat HTML.
        ''' </summary>
        Private Async Function GenerateWelcomeAsync() As Task
            Dbg("GenerateWelcomeAsync start")
            Dim langName As String = System.Globalization.CultureInfo.CurrentUICulture.DisplayName
            Dim partOfDay As String = GetPartOfDay()
            Dim manualText As String = Await GetManualOnceAsync()
            Dim systemPrompt As String

            manualText = manualText.Trim()
            If manualText.StartsWith("Error", System.StringComparison.OrdinalIgnoreCase) OrElse manualText = "" Then
                systemPrompt = $"Generate a brief, friendly {langName} welcome that naturally references it is {partOfDay} now, but tell the user that you can't work because you have no access to the manual (which needs to be configured and is retrieved either via an URL or file path; most likely, the path/URL is wrong or not working). Advise that the configured source should be checked or configured as per the manual."
            Else
                systemPrompt = $"Generate a brief, friendly {langName} welcome that naturally references it is {partOfDay} now and asks what you can do. Do NOT state your name. One short short sentence, not talkative."
            End If
            Dim userPrompt As String = ""
            Dim answer As String = ""
            Try
                Dim sw = Stopwatch.StartNew()
                answer = Await CallHelpMeLlmAsync(systemPrompt, userPrompt)
                sw.Stop()
                Dbg($"Welcome LLM ms={sw.ElapsedMilliseconds} rawLen={If(answer, "").Length}")
            Catch ex As Exception
                Dbg("Welcome LLM error: " & ex.Message)
                answer = "Hello! How can I help?"
            End Try

            answer = If(answer, "").Trim()
            AppendAssistantMarkdown(answer)
            _history.Add(("assistant", answer))
            PersistChatHtml()
            Dbg("GenerateWelcomeAsync done")
        End Function

        ''' <summary>
        ''' Builds the user prompt (question + manual + conversation + optional config), invokes the LLM,
        ''' and appends the response to the chat.
        ''' </summary>
        Private Async Function SendAsync(userText As String) As Task
            Dbg("SendAsync start")
            Try
                Dim hostInfo As String = If(String.IsNullOrEmpty(_hostAppName), "", $" (Host application (and version of {AN} add-in): Microsoft {_hostAppName})")
                Dim systemPrompt As String = _context.SP_HelpMe & hostInfo
                Dim manualText As String = Await GetManualOnceAsync()
                Dim convo As String = BuildConversationForLlm()

                manualText = manualText.Trim()
                If manualText.StartsWith("Error", System.StringComparison.OrdinalIgnoreCase) Or manualText = "" Then
                    manualText = "No manual"
                End If
                Dim sb As New StringBuilder()
                sb.AppendLine("User question:")
                sb.AppendLine(userText)
                sb.AppendLine()
                sb.AppendLine("Manual:")
                sb.AppendLine(manualText)
                sb.AppendLine()
                sb.AppendLine("Conversation so far:")
                sb.AppendLine(convo)

                ' Include configuration files if checkbox is checked
                If _chkIncludeConfig.Checked Then
                    Dim configContent = GetConfigurationContent()
                    If Not String.IsNullOrEmpty(configContent) Then
                        sb.AppendLine()
                        sb.AppendLine(configContent)
                    End If
                End If

                Dim sw = Stopwatch.StartNew()
                Dim answer As String = Await CallHelpMeLlmAsync(systemPrompt, sb.ToString())
                sw.Stop()

                answer = If(answer, "").Trim()
                Dbg($"SendAsync ms={sw.ElapsedMilliseconds} ansLen={answer.Length}")

                RemoveAssistantThinking()
                AppendAssistantMarkdown(answer)
                _history.Add(("assistant", answer))
                PersistChatHtml()
            Catch ex As Exception
                Dbg("SendAsync error: " & ex.Message)
                RemoveAssistantThinking()
                AppendAssistantMarkdown("*(Error: " & System.Security.SecurityElement.Escape(ex.Message) & ")*")
            End Try
        End Function

        ''' <summary>
        ''' Reads known configuration files and returns a single string block to be appended to the LLM prompt.
        ''' API keys are redacted by `SanitizeConfigContent`.
        ''' </summary>
        Private Function GetConfigurationContent() As String
            Try
                Dim sb As New StringBuilder()
                sb.AppendLine("<Configuration Files>")

                ' Get main config file
                Dim mainPath As String = Nothing
                Try
                    mainPath = GetActiveConfigFilePath(_context)
                Catch
                End Try

                If Not String.IsNullOrWhiteSpace(mainPath) AndAlso File.Exists(mainPath) Then
                    sb.AppendLine($"<Main Configuration ({AN2}.ini)>")
                    sb.AppendLine($"Path: {mainPath}")
                    Try
                        Dim content = File.ReadAllText(mainPath)
                        sb.AppendLine(SanitizeConfigContent(content))
                    Catch ex As Exception
                        sb.AppendLine($"Error reading file: {ex.Message}")
                    End Try
                    sb.AppendLine($"</Main Configuration>")
                End If

                ' Get default INI paths if available (assuming DefaultINIPaths exists in SharedMethods)
                Try
                    Dim defaultPaths = SharedMethods.DefaultINIPaths
                    For Each kvp In defaultPaths
                        Dim p = Environment.ExpandEnvironmentVariables(kvp.Value)
                        If File.Exists(p) AndAlso Not String.Equals(p, mainPath, StringComparison.OrdinalIgnoreCase) Then
                            sb.AppendLine($"<{kvp.Key} Configuration>")
                            sb.AppendLine($"Path: {p}")
                            Try
                                Dim content = File.ReadAllText(p)
                                sb.AppendLine(SanitizeConfigContent(content))
                            Catch ex As Exception
                                sb.AppendLine($"Error reading file: {ex.Message}")
                            End Try
                            sb.AppendLine($"</{kvp.Key} Configuration>")
                        End If
                    Next
                Catch
                    ' DefaultINIPaths might not be accessible
                End Try

                ' Alternate model path
                If Not String.IsNullOrWhiteSpace(_context.INI_AlternateModelPath) Then
                    Dim alt = Environment.ExpandEnvironmentVariables(_context.INI_AlternateModelPath)
                    If File.Exists(alt) AndAlso Not String.Equals(alt, mainPath, StringComparison.OrdinalIgnoreCase) Then
                        sb.AppendLine("<Alternate Model Configuration>")
                        sb.AppendLine($"Path: {alt}")
                        Try
                            Dim content = File.ReadAllText(alt)
                            sb.AppendLine(SanitizeConfigContent(content))
                        Catch ex As Exception
                            sb.AppendLine($"Error reading file: {ex.Message}")
                        End Try
                        sb.AppendLine("</Alternate Model Configuration>")
                    End If
                End If

                ' Special service path
                If Not String.IsNullOrWhiteSpace(_context.INI_SpecialServicePath) Then
                    Dim sp = Environment.ExpandEnvironmentVariables(_context.INI_SpecialServicePath)
                    If File.Exists(sp) AndAlso Not String.Equals(sp, mainPath, StringComparison.OrdinalIgnoreCase) Then
                        sb.AppendLine("<Special Service Configuration>")
                        sb.AppendLine($"Path: {sp}")
                        Try
                            Dim content = File.ReadAllText(sp)
                            sb.AppendLine(SanitizeConfigContent(content))
                        Catch ex As Exception
                            sb.AppendLine($"Error reading file: {ex.Message}")
                        End Try
                        sb.AppendLine("</Special Service Configuration>")
                    End If
                End If

                sb.AppendLine("</Configuration Files>")
                Return sb.ToString()
            Catch ex As Exception
                Dbg($"GetConfigurationContent error: {ex.Message}")
                Return $"<Configuration Files>Error retrieving configuration: {ex.Message}</Configuration Files>"
            End Try
        End Function

        ''' <summary>
        ''' Redacts API keys and removes comment lines from INI-like configuration text before sending it to the LLM.
        ''' </summary>
        Private Function SanitizeConfigContent(content As String) As String
            If String.IsNullOrEmpty(content) Then Return content

            Dim lines = content.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
            Dim result As New StringBuilder()

            For Each line In lines
                Dim trimmedLine = line.TrimStart()

                ' Skip comment lines (starting with ;)
                If trimmedLine.StartsWith(";") Then
                    Continue For
                End If

                ' Check if this line contains an API key
                Dim apiKeyMatch = Regex.Match(line, "^(\s*APIKey(?:_2)?)\s*=\s*(.*)$", RegexOptions.IgnoreCase)
                If apiKeyMatch.Success Then
                    Dim key = apiKeyMatch.Groups(1).Value
                    Dim value = apiKeyMatch.Groups(2).Value.Trim()
                    If String.IsNullOrWhiteSpace(value) Then
                        ' Keep blank API keys as is
                        result.AppendLine(line)
                    Else
                        ' Replace non-blank API keys with placeholder
                        result.AppendLine($"{key}=[for security reasons, you are not provided with the real API key contained in this file]")
                    End If
                Else
                    ' Keep non-API key lines as is
                    result.AppendLine(line)
                End If
            Next

            Return result.ToString().TrimEnd()
        End Function

        ''' <summary>
        ''' Invokes the LLM for HelpMe. Optionally switches to a special model configuration for this task.
        ''' Serialized via `_modelSemaphore` to avoid concurrent model switching/restores.
        ''' </summary>
        Private Async Function CallHelpMeLlmAsync(systemPrompt As String, userPrompt As String) As Task(Of String)
            If _context Is Nothing Then Return ""
            If Not String.IsNullOrEmpty(_hostAppName) AndAlso systemPrompt.IndexOf(_hostAppName, StringComparison.OrdinalIgnoreCase) < 0 Then
                systemPrompt &= $" (This chat runs inside Microsoft {_hostAppName}.)"
            End If

            If Not _helpMeAltResolved Then
                _helpMeAltAvailable = False
                If Not String.IsNullOrWhiteSpace(_context.INI_AlternateModelPath) Then
                    If SharedMethods.GetSpecialTaskModel(_context, _context.INI_AlternateModelPath, "HelpMe") Then
                        _helpMeAltAvailable = True
                    End If
                End If
                _helpMeAltResolved = True
                Dbg($"Alternate HelpMe model available={_helpMeAltAvailable}")
            End If

            Await _modelSemaphore.WaitAsync().ConfigureAwait(False)
            Dim backupConfig As ModelConfig = Nothing
            Dim appliedAlternate As Boolean = False
            Dim useSecondApi As Boolean = False
            Dim timeout As Long = 0

            Try
                If _helpMeAltAvailable Then
                    backupConfig = SharedMethods.GetCurrentConfig(_context)
                    useSecondApi = True
                    appliedAlternate = True
                    timeout = If(_context.INI_Timeout_2 > 0, _context.INI_Timeout_2, _context.INI_Timeout)
                Else
                    timeout = _context.INI_Timeout
                End If

                Return Await SharedMethods.LLM(_context,
                                           systemPrompt,
                                           userPrompt,
                                           "",
                                           "",
                                           timeout,
                                           useSecondApi,
                                           True).ConfigureAwait(False)
            Finally
                If appliedAlternate AndAlso backupConfig IsNot Nothing Then
                    SharedMethods.RestoreDefaults(_context, backupConfig)
                End If
                _modelSemaphore.Release()
            End Try
        End Function

        ''' <summary>
        ''' Returns manual text from cache when possible; otherwise loads it from the configured path/URL.
        ''' </summary>
        Private Async Function GetManualOnceAsync() As Task(Of String)
            Dim path = If(_context IsNot Nothing, _context.INI_HelpMeInkyPath, "")
            If String.IsNullOrWhiteSpace(path) Then Return ""
            If _manualCache IsNot Nothing AndAlso String.Equals(_manualCachePath, path, StringComparison.OrdinalIgnoreCase) Then
                Return _manualCache
            End If
            Dbg("Loading manual fresh: " & path)
            Dim loaded = Await GetManualTextFreshAsync(path, _context)
            If Not String.IsNullOrEmpty(loaded) Then
                _manualCache = loaded
                _manualCachePath = path
            End If
            Return If(_manualCache, "")
        End Function

        ''' <summary>
        ''' Loads manual text from an HTTP(S) URL or a local file path.
        ''' Supports basic PDF detection for remote content and uses helper readers for PDF/RTF/DOCX.
        ''' </summary>
        Private Shared Async Function GetManualTextFreshAsync(pathOrUrl As String, Optional context As ISharedContext = Nothing) As Task(Of String)
            If String.IsNullOrWhiteSpace(pathOrUrl) Then Return ""
            Dim s As String = pathOrUrl.Trim()

            ' Ensure modern TLS (harmless if already enabled)
            Try
                '#If NETFRAMEWORK Then
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 Or CType(&HC00, SecurityProtocolType) ' include TLS 1.3 if supported
                '#End If
            Catch
            End Try

            ' Remote URL
            If s.StartsWith("http://", StringComparison.OrdinalIgnoreCase) OrElse s.StartsWith("https://", StringComparison.OrdinalIgnoreCase) Then
                Try
                    Dim handler As New HttpClientHandler()
                    handler.AllowAutoRedirect = True
                    handler.AutomaticDecompression = DecompressionMethods.GZip Or DecompressionMethods.Deflate

                    Using client As New HttpClient(handler)
                        client.Timeout = TimeSpan.FromSeconds(30)
                        ' Some servers reject requests without UA
                        Try
                            client.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", "RedInk/1.0 (+https://redink.ai)")
                            client.DefaultRequestHeaders.TryAddWithoutValidation("Accept", "application/pdf, text/*, */*")
                        Catch
                        End Try

                        Using resp As HttpResponseMessage = Await client.GetAsync(s, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(False)
                            If Not resp.IsSuccessStatusCode Then Return ""

                            Dim data As Byte() = Await resp.Content.ReadAsByteArrayAsync().ConfigureAwait(False)

                            ' Extract media type if provided
                            Dim mediaType As String = ""
                            If resp.Content IsNot Nothing AndAlso resp.Content.Headers IsNot Nothing AndAlso resp.Content.Headers.ContentType IsNot Nothing Then
                                If Not String.IsNullOrEmpty(resp.Content.Headers.ContentType.MediaType) Then
                                    mediaType = resp.Content.Headers.ContentType.MediaType.ToLowerInvariant()
                                End If
                            End If

                            ' PDF detection
                            Dim isPdf As Boolean = False

                            ' 1) Declared content-type
                            If Not String.IsNullOrEmpty(mediaType) AndAlso mediaType.IndexOf("pdf", StringComparison.OrdinalIgnoreCase) >= 0 Then
                                isPdf = True
                            End If

                            ' 2) URL contains ".pdf" anywhere (also handles querystring)
                            If Not isPdf Then
                                If s.IndexOf(".pdf", StringComparison.OrdinalIgnoreCase) >= 0 Then
                                    isPdf = True
                                End If
                            End If

                            ' 3) Magic header scan for "%PDF" within first KB (after possible BOM or garbage)
                            If Not isPdf AndAlso data IsNot Nothing AndAlso data.Length >= 4 Then
                                Dim scanMax As Integer = Math.Min(data.Length - 4, 1024)
                                Dim i As Integer = 0
                                While i <= scanMax
                                    If data(i) = AscW("%"c) AndAlso data(i + 1) = AscW("P"c) AndAlso data(i + 2) = AscW("D"c) AndAlso data(i + 3) = AscW("F"c) Then
                                        isPdf = True
                                        Exit While
                                    End If
                                    i += 1
                                End While
                            End If

                            If isPdf Then
                                Try
                                    Dim tmpPath As String = Path.Combine(Path.GetTempPath(), "manual_" & Guid.NewGuid().ToString("N") & ".pdf")
                                    File.WriteAllBytes(tmpPath, data)
                                    Return Await SharedMethods.ReadPdfAsText(tmpPath, True, False, False, context).ConfigureAwait(False)
                                Catch
                                    Return ""
                                End Try
                            End If

                            ' Fallback: decode as text
                            Dim enc As Encoding = Encoding.UTF8
                            Dim charset As String = ""
                            If resp.Content IsNot Nothing AndAlso resp.Content.Headers IsNot Nothing AndAlso resp.Content.Headers.ContentType IsNot Nothing Then
                                charset = resp.Content.Headers.ContentType.CharSet
                            End If
                            If Not String.IsNullOrEmpty(charset) Then
                                Try
                                    enc = Encoding.GetEncoding(charset)
                                Catch
                                    enc = Encoding.UTF8
                                End Try
                            End If

                            Dim text As String = enc.GetString(data)

                            ' HTML -> plain
                            If Not String.IsNullOrEmpty(mediaType) AndAlso mediaType.IndexOf("html", StringComparison.OrdinalIgnoreCase) >= 0 Then
                                If LooksLikeHtml(text) Then
                                    Return HtmlToPlain(text)
                                Else
                                    Return text
                                End If
                            End If

                            ' Generic octet-stream sometimes still is HTML
                            If LooksLikeHtml(text) Then
                                Return HtmlToPlain(text)
                            End If

                            Return text
                        End Using
                    End Using
                Catch
                    Return ""
                End Try
            End If

            ' Local file path
            Try
                If Not File.Exists(s) Then Return ""
                Select Case Path.GetExtension(s).ToLowerInvariant()
                    Case ".txt", ".md", ".log"
                        Return File.ReadAllText(s, Encoding.UTF8)
                    Case ".docx"
                        Return ReadDocxWithWordInterop(s)
                    Case ".rtf"
                        Try
                            Return SharedMethods.ReadRtfAsText(s)
                        Catch
                            Return ""
                        End Try
                    Case ".pdf"
                        Try
                            Return Await SharedMethods.ReadPdfAsText(s, True, False, False, context).ConfigureAwait(False)
                        Catch
                            Return ""
                        End Try
                    Case Else
                        Return File.ReadAllText(s, Encoding.UTF8)
                End Select
            Catch
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' Reads DOCX text using Microsoft Word interop (opens the document read-only and returns its text).
        ''' </summary>
        Private Shared Function ReadDocxWithWordInterop(path As String) As String
            Dim app As Microsoft.Office.Interop.Word.Application = Nothing
            Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
            Try
                Try
                    app = CType(Runtime.InteropServices.Marshal.GetActiveObject("Word.Application"), Microsoft.Office.Interop.Word.Application)
                Catch
                    app = New Microsoft.Office.Interop.Word.Application() With {.Visible = False}
                End Try
                Dim fileName As Object = path
                doc = app.Documents.Open(fileName, [ReadOnly]:=True, Visible:=False)
                Dim txt = doc.Content.Text
                doc.Close(SaveChanges:=False)
                Return If(txt, "")
            Catch
                Try
                    If doc IsNot Nothing Then doc.Close(SaveChanges:=False)
                Catch
                End Try
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' Heuristic to determine whether text looks like HTML.
        ''' </summary>
        Private Shared Function LooksLikeHtml(s As String) As Boolean
            If String.IsNullOrEmpty(s) Then Return False
            Dim t = s.TrimStart()
            Return t.StartsWith("<!DOCTYPE", StringComparison.OrdinalIgnoreCase) _
            OrElse t.StartsWith("<html", StringComparison.OrdinalIgnoreCase) _
            OrElse t.IndexOf("<body", StringComparison.OrdinalIgnoreCase) >= 0
        End Function

        ''' <summary>
        ''' Converts HTML to plain text by extracting the document's inner text.
        ''' </summary>
        Private Shared Function HtmlToPlain(html As String) As String
            Try
                Dim doc As New HtmlAgilityPack.HtmlDocument()
                doc.LoadHtml(html)
                Return doc.DocumentNode.InnerText
            Catch
                Return html
            End Try
        End Function

        ''' <summary>
        ''' Initializes the HTML document used by the chat browser and resets the append queue.
        ''' </summary>
        Private Sub InitializeChatHtml()
            Ui(Sub()
                   _htmlQueue.Clear()
                   _htmlReady = False
                   Dbg("InitializeChatHtml")
                   Dim baseSize = If(Me.Font IsNot Nothing, Me.Font.SizeInPoints, 9.0F)
                   Dim fontPt = Math.Max(CSng(baseSize + 1.0F), 10.0F)
                   Dim css =
                    $"html,body{{height:100%;margin:0;padding:0;background:#fff;color:#000;}}
                        body{{font-family:'Segoe UI',Tahoma,Arial,sans-serif;font-size:{fontPt}pt;line-height:1.45;}}
                        #chat{{padding:8px;}}
                        .msg{{margin:6px 0;word-wrap:break-word;}}
                        .msg .who{{font-weight:600;margin-right:4px;}}
                        .msg.user .who{{color:#333;}}
                        .msg.assistant .who{{color:#003366;}}
                        .msg.thinking .content{{opacity:.75;font-style:italic;}}
                        a{{color:#0068c9;text-decoration:underline;cursor:pointer;}}
                        pre{{white-space:pre-wrap;background:#f6f8fa;border:1px solid #e1e4e8;border-radius:4px;padding:6px;}}"
                   Dim html =
                    $"<!DOCTYPE html>
                        <html>
                        <head>
                        <meta http-equiv=""X-UA-Compatible"" content=""IE=edge"" />
                        <meta charset=""utf-8"">
                        <style>{css}</style>
                        <script>
                        function appendMessage(html) {{
                          var c=document.getElementById('chat'); if(!c) return;
                          var temp=document.createElement('div'); temp.innerHTML=html;
                          while(temp.firstChild){{c.appendChild(temp.firstChild);}}
                          window.scrollTo(0, document.body.scrollHeight);
                        }}
                        function removeById(id) {{
                          var el=document.getElementById(id); if(!el||!el.parentNode) return;
                          el.parentNode.removeChild(el);
                        }}
                        </script>
                        </head>
                        <body><div id=""chat""></div></body>
                        </html>"
                   _chat.DocumentText = html
               End Sub)
        End Sub

        ''' <summary>
        ''' Marks the chat HTML document as ready and flushes any queued HTML fragments.
        ''' </summary>
        Private Sub Chat_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs)
            _htmlReady = True
            Dbg("DocumentCompleted flushQueue=" & _htmlQueue.Count)
            If _htmlQueue.Count > 0 Then
                Try
                    For Each frag In _htmlQueue
                        _chat.Document.InvokeScript("appendMessage", New Object() {frag})
                    Next
                Catch ex As Exception
                    Dbg("Flush error: " & ex.Message)
                Finally
                    _htmlQueue.Clear()
                End Try
            End If
        End Sub

        ''' <summary>
        ''' Intercepts navigation to external links and opens them via the OS shell.
        ''' </summary>
        Private Sub Chat_Navigating(sender As Object, e As WebBrowserNavigatingEventArgs)
            Try
                Dim scheme = e.Url?.Scheme?.ToLowerInvariant()
                If scheme = "http" OrElse scheme = "https" OrElse scheme = "mailto" Then
                    e.Cancel = True
                    Process.Start(New ProcessStartInfo(e.Url.ToString()) With {.UseShellExecute = True})
                End If
            Catch ex As Exception
                Dbg("Navigating error: " & ex.Message)
            End Try
        End Sub

        ''' <summary>Prevents the WebBrowser control from opening new windows.</summary>
        Private Sub Chat_NewWindow(sender As Object, e As CancelEventArgs)
            e.Cancel = True
        End Sub

        ''' <summary>
        ''' Appends an HTML fragment to the chat document or queues it if the document is not ready.
        ''' </summary>
        Private Sub AppendHtml(fragment As String)
            If String.IsNullOrEmpty(fragment) Then Return
            Ui(Sub()
                   If Not _htmlReady OrElse _chat.Document Is Nothing Then
                       _htmlQueue.Add(fragment)
                       Return
                   End If
                   Try
                       _chat.Document.InvokeScript("appendMessage", New Object() {fragment})
                   Catch
                       _htmlQueue.Add(fragment)
                   End Try
               End Sub)
        End Sub

        ''' <summary>
        ''' HTML-encodes and appends a user message to the chat view, then persists chat HTML.
        ''' </summary>
        Private Sub AppendUserHtml(text As String)
            Dim encoded = WebUtility.HtmlEncode(text).Replace(vbCrLf, "<br>").Replace(vbLf, "<br>").Replace(vbCr, "<br>")
            AppendHtml($"<div class='msg user'><span class='who'>You:</span><span class='content'>{encoded}</span></div>")
            PersistChatHtml()
        End Sub

        ''' <summary>
        ''' Appends a "Thinking..." placeholder for the assistant and stores its DOM id for later removal.
        ''' </summary>
        Private Sub ShowAssistantThinking()
            _lastThinkingId = "thinking-" & Guid.NewGuid().ToString("N")
            AppendHtml($"<div id=""{_lastThinkingId}"" class='msg assistant thinking'><span class='who'>{WebUtility.HtmlEncode(AssistantName)}:</span><span class='content'>Thinking...</span></div>")
        End Sub

        ''' <summary>
        ''' Removes the last "Thinking..." placeholder from the chat DOM (if present).
        ''' </summary>
        Private Sub RemoveAssistantThinking()
            If String.IsNullOrEmpty(_lastThinkingId) Then Return
            Ui(Sub()
                   Try
                       If _chat.Document IsNot Nothing Then
                           _chat.Document.InvokeScript("removeById", New Object() {_lastThinkingId})
                       End If
                   Catch
                   Finally
                       _lastThinkingId = Nothing
                   End Try
               End Sub)
        End Sub

        ''' <summary>
        ''' Converts Markdown to HTML and appends an assistant message.
        ''' Single-paragraph replies are rendered inline; multi-block replies are split into inline first paragraph + block content.
        ''' </summary>
        Private Sub AppendAssistantMarkdown(md As String)
            md = If(md, "")
            Dim body = Markdig.Markdown.ToHtml(md, _mdPipeline)
            Dim t = body.Trim()
            Dim isSingle = Regex.IsMatch(t, "^\s*<p>[\s\S]*?</p>\s*$", RegexOptions.IgnoreCase) AndAlso
                       Not Regex.IsMatch(t, "<(ul|ol|pre|table|h[1-6]|blockquote|hr|div)\b", RegexOptions.IgnoreCase)

            If isSingle Then
                ' Single paragraph: keep fully inline
                Dim inlineHtml = Regex.Replace(t, "^\s*<p>|</p>\s*$", "", RegexOptions.IgnoreCase)
                AppendHtml($"<div class='msg assistant'><span class='who'>{WebUtility.HtmlEncode(AssistantName)}:</span><span class='content'>{inlineHtml}</span></div>")
            Else
                ' Multi-block: inline only the first paragraph; render the rest below
                Dim m = Regex.Match(t, "^\s*<p>([\s\S]*?)</p>\s*", RegexOptions.IgnoreCase)
                If m.Success Then
                    Dim firstInline = m.Groups(1).Value
                    Dim rest = t.Substring(m.Index + m.Length).Trim()
                    Dim sb As New StringBuilder()
                    sb.Append("<div class='msg assistant'>")
                    sb.Append("<span class='who'>").Append(WebUtility.HtmlEncode(AssistantName)).Append(":</span>")
                    sb.Append("<span class='content'>").Append(firstInline).Append("</span>")
                    If rest.Length > 0 Then
                        sb.Append("<div class='content'>").Append(rest).Append("</div>")
                    End If
                    sb.Append("</div>")
                    AppendHtml(sb.ToString())
                Else
                    ' Fallback: previous behavior
                    AppendHtml($"<div class='msg assistant'><span class='who'>{WebUtility.HtmlEncode(AssistantName)}:</span><div class='content'>{t}</div></div>")
                End If
            End If
        End Sub

        ''' <summary>
        ''' Persists the current chat DOM (`#chat` inner HTML) into `My.Settings.LastHelpMeChatHtml`.
        ''' </summary>
        Private Sub PersistChatHtml()
            Ui(Sub()
                   Try
                       If _chat.Document Is Nothing Then Return
                       Dim root = _chat.Document.GetElementById("chat")
                       If root Is Nothing Then Return
                       My.Settings.LastHelpMeChatHtml = root.InnerHtml
                       My.Settings.Save()
                   Catch ex As Exception
                       Dbg("PersistChatHtml error: " & ex.Message)
                   End Try
               End Sub)
        End Sub

        ''' <summary>
        ''' Restores a plain text transcript into the current chat HTML by parsing "You:" and "{AssistantName}:" prefixes.
        ''' </summary>
        Private Sub AppendTranscriptToHtml(transcript As String)
            If String.IsNullOrEmpty(transcript) Then Return
            Dim lines = transcript.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split({vbLf}, StringSplitOptions.None)
            Dim currentRole As String = Nothing
            Dim content As New StringBuilder()

            Dim flush =
            Sub()
                If content.Length = 0 OrElse String.IsNullOrEmpty(currentRole) Then
                    content.Clear() : currentRole = Nothing : Return
                End If
                If currentRole = "user" Then
                    Dim enc = WebUtility.HtmlEncode(content.ToString()).Replace(vbLf, "<br>")
                    AppendHtml($"<div class='msg user'><span class='who'>You:</span><span class='content'>{enc}</span></div>")
                Else
                    AppendAssistantMarkdown(content.ToString())
                End If
                content.Clear()
                currentRole = Nothing
            End Sub

            For Each ln In lines
                If ln.StartsWith("You:", StringComparison.OrdinalIgnoreCase) Then
                    flush() : currentRole = "user" : content.Append(ln.Substring(4).TrimStart())
                ElseIf ln.StartsWith(AssistantName & ":", StringComparison.OrdinalIgnoreCase) Then
                    flush() : currentRole = "assistant" : content.Append(ln.Substring((AssistantName & ":").Length).TrimStart())
                Else
                    If content.Length > 0 Then content.AppendLine()
                    content.Append(ln)
                End If
            Next
            flush()
            PersistChatHtml()
        End Sub

        ''' <summary>
        ''' Persists a bounded plain transcript into `My.Settings.LastHelpMeChat` using `INI_ChatCap` (minimum 5000).
        ''' </summary>
        Private Sub PersistTranscriptLimited()
            Dim transcript = BuildTranscriptPlain()
            Dim cap As Integer = Math.Max(5000, If(_context IsNot Nothing, _context.INI_ChatCap, 0))
            If transcript.Length > cap Then
                transcript = transcript.Substring(transcript.Length - cap)
            End If
            My.Settings.LastHelpMeChat = transcript
        End Sub

        ''' <summary>
        ''' Builds a plain text transcript from `_history` using "You:" and "{AssistantName}:" prefixes.
        ''' </summary>
        Private Function BuildTranscriptPlain() As String
            Dim sb As New StringBuilder()
            For Each m In _history
                If m.Role = "user" Then
                    sb.AppendLine("You: " & m.Content)
                Else
                    sb.AppendLine(AssistantName & ": " & m.Content)
                End If
            Next
            Return sb.ToString()
        End Function

        ''' <summary>
        ''' Builds conversation text for LLM context, capped by `INI_ChatCap` (minimum 5000), taking the newest messages first.
        ''' </summary>
        Private Function BuildConversationForLlm() As String
            Dim sb As New StringBuilder()
            Dim cap As Integer = Math.Max(5000, If(_context IsNot Nothing, _context.INI_ChatCap, 0))
            Dim acc = 0
            For i = _history.Count - 1 To 0 Step -1
                Dim line = If(_history(i).Role = "user", "User: ", AssistantName & ": ") & _history(i).Content & Environment.NewLine
                If acc + line.Length > cap Then
                    Dim remain = cap - acc
                    If remain > 0 Then sb.Insert(0, line.Substring(line.Length - remain))
                    Exit For
                Else
                    sb.Insert(0, line)
                    acc += line.Length
                End If
            Next
            Return sb.ToString()
        End Function

        ''' <summary>
        ''' Removes some Markdown formatting artifacts for plain transcript usage.
        ''' </summary>
        Private Shared Function StripMarkdownForTranscript(md As String) As String
            If String.IsNullOrEmpty(md) Then Return ""
            Dim s = Regex.Replace(md, "```.*?```", "", RegexOptions.Singleline)
            s = s.Replace("**", "").Replace("__", "").Replace("*", "").Replace("_", "").Replace("`", "")
            Return s
        End Function

        ''' <summary>
        ''' Returns a simple part-of-day bucket based on the current local hour.
        ''' </summary>
        Private Shared Function GetPartOfDay() As String
            Dim h = DateTime.Now.Hour
            If h < 12 Then Return "Morning"
            If h < 18 Then Return "Afternoon"
            Return "Evening"
        End Function
    End Class

End Namespace