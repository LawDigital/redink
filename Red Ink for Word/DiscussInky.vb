' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: DiscussInky.vb
' Purpose:
'   WinForms surface that hosts the "Discuss this" multi-persona chat inside Word.
'   Provides persona + mission selection, knowledge loading (file or directory),
'   optional inclusion of the active Word document, transcript rendering/persistence,
'   and LLM invocation with optional alternate-model / second-API support.
'   Includes two automation modes:
'     - Autorespond: simulates a second participant for multi-party dialogue
'     - Sort It Out: structured Advocate vs Challenger discussion using the same engine
'
' Architecture / How it works:
'  - UI Composition:
'     * WebBrowser transcript renderer (Markdown -> HTML via Markdig, custom CSS; external links open in browser)
'     * Multiline TextBox input with Enter-to-send (Shift+Enter for newline) and Esc-to-close
'     * Buttons: Send, Persona, Mission, Edit Local Persona Lib, Load Knowledge, Alternate Model,
'       Clear, Send to Doc, Close, Autorespond, Sort It Out, Tools
'     * Checkboxes: Include active document, Persist knowledge temporarily, Enable tooling, Tooling log
'
'  - Session State & Persistence (`My.Settings`):
'     * Persists window geometry, selected persona/mission, checkbox states, knowledge path, tooling settings
'     * Persists transcript in two forms:
'         - HTML fragment (`DiscussLastChatHtml`) for fast UI restoration
'         - Plain transcript (`DiscussLastChat`) to rebuild `_history` for LLM context
'     * Maintains an in-process runtime knowledge cache (`_cachedKnowledgeContent/_cachedKnowledgeFilePath`)
'       (survives closing/reopening the form while Word remains running)
'     * Optional temp-file persistence of knowledge when "Persist knowledge temporarily" is enabled
'       (`%TEMP%\redink-discussknowledge.txt`)
'
'  - Personas & Missions:
'     * Personas loaded from local and/or global persona libraries (Name|Prompt, `;` comments supported)
'       - Local personas are labeled "(local)" and display names are de-duplicated
'     * Missions loaded from a sibling file derived from the persona library filename:
'         [personaFileNameWithoutExtension]-missions.txt (Name|Prompt, `;` comments supported)
'     * Mission/Persona editing is available via the shared text editor; missions reload immediately
'
'  - Knowledge Loading (File/Directory):
'     * Loads a single file or all supported files from a directory (top-level only; capped to prevent overload)
'     * Supported formats include text/code/markup, RTF, DOC/DOCX, PPTX, PDF (optional OCR when available)
'     * When multiple files are loaded, wraps each into numbered tags:
'         <documentN name="file.ext"> ... </documentN>
'     * Produces a user-confirmed load summary (loaded/failed/ignored, PDF extraction warnings)
'
'  - LLM Pipeline & Context Unification:
'     * Builds prompts from shared context pieces:
'         - Persona system prompt (+ optional mission clause)
'         - Knowledge base (if loaded)
'         - Active Word document excerpt
'         - Conversation history (`_history`), capped/tail-truncated using `_context.INI_ChatCap`
'     * History roles:
'         - `user`
'         - `assistant` (main persona; Sort It Out stores prefixed display names like "X (Advocate): ...")
'         - `autoresponder` (second participant; stored already prefixed with its display name)
'       `BuildConversationForAutoResponder()` normalizes roles into a single transcript so both parties see the
'       full conversation regardless of who spoke.
'     * Generates a brief welcome message (persona/mission aware) and displays session info on startup when needed
'
'  - Automation Modes:
'     * Autorespond:
'         - Alternates between responder persona and main persona for up to N rounds
'         - Supports a stop phrase (`<AUTORESPOND_STOP>`) and optional progress UI
'         - Can optionally generate a discussion summary after completion
'     * Sort It Out:
'         - Runs a structured Advocate vs Challenger dialogue using the same alternation engine
'         - Missions can be auto-generated via LLM or selected manually; settings are persisted
'         - Can optionally generate a discussion summary after completion
'
'  - Model Selection / Tooling:
'     * Alternate model support:
'         - If an alternate model INI is configured, user can select an alternate model; the model is applied
'           only for the duration of each call and then restored (semaphore-protected)
'         - Legacy "second API" toggle is supported when configured
'     * Tooling support:
'         - Optional per-session tool selection and execution via `Globals.ThisAddIn.ExecuteToolingLoop`
'         - Tooling UI is enabled/disabled based on the currently selected model's tooling capability
'
'  - Export:
'     * "Send to Doc" exports the conversation to a new Word document using Markdown insertion,
'       formatting speakers explicitly for `user`, `assistant`, and `autoresponder`
'
' External Dependencies:
'  - Markdig: Markdown-to-HTML conversion for transcript rendering and summary display
'  - `SharedLibrary.SharedMethods`:
'     * LLM calls and model configuration switching
'     * OCR/PDF utilities and file extraction helpers
'     * Shared dialogs (selection, variable input, message boxes, text editor)
'  - `Globals.ThisAddIn`:
'     * Word interop access, tooling loop, presentation extraction, drag/drop picker, progress UI helpers
' =============================================================================

Option Strict Off
Option Explicit On

Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Net
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Markdig
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

''' <summary>
''' WinForms surface for persona-driven LLM discussions tied to knowledge files.
''' </summary>
Public Class DiscussInky
    Inherits System.Windows.Forms.Form

#Region "Constants and Fields"

    Private Const AssistantName As String = Globals.ThisAddIn.AN6
    Private Const PersistedKnowledgeFileName As String = "redink-discussknowledge.txt"
    Private _currentPersonaName As String = AssistantName
    Private _currentPersonaPrompt As String = ""

    ' Mission state
    Private _currentMissionName As String = ""
    Private _currentMissionPrompt As String = ""

    Private ReadOnly _context As ISharedContext
    Private ReadOnly _mdPipeline As Markdig.MarkdownPipeline

    ' Runtime knowledge cache (persists while Word is running, not in My.Settings)
    Private Shared _cachedKnowledgeContent As String = Nothing
    Private Shared _cachedKnowledgeFilePath As String = Nothing

    ' Supported file extensions for knowledge loading
    Private Shared ReadOnly SupportedKnowledgeExtensions As String() = {
        ".txt", ".rtf", ".doc", ".docx", ".pdf", ".pptx", ".ini", ".csv", ".log",
        ".json", ".xml", ".html", ".htm", ".md", ".vb", ".cs", ".js", ".ts",
        ".py", ".java", ".cpp", ".c", ".h", ".sql", ".yaml", ".yml"
    }

    ' Random words for response variety
    Private Shared ReadOnly _randomModifiers As String() = {
        "thoughtfully", "carefully", "precisely", "clearly", "concisely",
        "helpfully", "insightfully", "thoroughly", "directly", "naturally"
    }
    Private Shared ReadOnly _rng As New Random()

    ' Tooling support
    Private _selectedToolsForChat As List(Of ModelConfig) = Nothing

    ' Autorespond constants
    Private Const MaxAutoRespondRounds As Integer = 100
    Private Const DefaultRespondRounds As Integer = 5
    Private Const AutoRespondStopWord As String = "<AUTORESPOND_STOP>"
    Private Const ShowGeneratedMissionsConfirmation As Boolean = False
    Private DefaultAutoRespondBreakOff As String = $"If this chat is going in circles, if you have come to an agreement or solution, or if this chat is drifting away to a point that is no longer productive, stop the responses by including the exact text '{AutoRespondStopWord}' at the end of your message and explain why (if because a solution is found, explain the solution, common grounds, etc.)."

    ' Sort It Out feature state
    Private _sortOutInProgress As Boolean = False
    Private _sortOutMainMissionPrompt As String = ""
    Private _sortOutResponderMissionPrompt As String = ""
    Private _sortOutOriginalMissionName As String = ""
    Private _sortOutOriginalMissionPrompt As String = ""

    Private Const MinRoundsForAutoSummary As Integer = 2

    ' UI Controls
    Private ReadOnly _chat As WebBrowser = New WebBrowser() With {
        .Dock = DockStyle.Fill,
        .AllowWebBrowserDrop = False,
        .IsWebBrowserContextMenuEnabled = True,
        .WebBrowserShortcutsEnabled = True,
        .ScriptErrorsSuppressed = True
    }
    Private ReadOnly _txtInput As TextBox = New TextBox() With {
        .Dock = DockStyle.Fill,
        .Multiline = True,
        .AcceptsReturn = True,
        .WordWrap = True
    }

    Private ReadOnly _toolTip As ToolTip = New ToolTip() With {
    .AutoPopDelay = 10000,
    .InitialDelay = 500,
    .ReshowDelay = 200
}

    Private ReadOnly _btnClear As Button = New Button() With {.Text = "Clear", .AutoSize = True}
    Private ReadOnly _btnSendToDoc As Button = New Button() With {.Text = "Send to Doc", .AutoSize = True}
    Private ReadOnly _btnClose As Button = New Button() With {.Text = "Close", .AutoSize = True}
    Private ReadOnly _btnSend As Button = New Button() With {.Text = $"Send", .AutoSize = True}
    Private ReadOnly _btnPersona As Button = New Button() With {.Text = "Persona", .AutoSize = True}
    Private ReadOnly _btnMission As Button = New Button() With {.Text = "Mission", .AutoSize = True}
    Private ReadOnly _btnEditPersona As Button = New Button() With {.Text = "Edit Local Persona Lib", .AutoSize = True}
    Private ReadOnly _btnKnowledge As Button = New Button() With {.Text = "Load Knowledge (Docs)", .AutoSize = True}
    Private ReadOnly _btnAlternateModel As Button = New Button() With {.Text = "Alternate Model", .AutoSize = True}
    Private ReadOnly _chkIncludeActiveDoc As System.Windows.Forms.CheckBox = New System.Windows.Forms.CheckBox() With {.Text = "Include active document", .AutoSize = True}
    Private ReadOnly _chkPersistKnowledge As System.Windows.Forms.CheckBox = New System.Windows.Forms.CheckBox() With {.Text = "Persist knowledge temporarily", .AutoSize = True}
    Private ReadOnly _btnAutoRespond As Button = New Button() With {.Text = "Autorespond", .AutoSize = True}
    Private ReadOnly _btnSortOut As Button = New Button() With {.Text = "Sort It Out", .AutoSize = True}
    Private ReadOnly _btnTools As Button = New Button() With {.Text = Globals.ThisAddIn.ToolFriendlyName, .AutoSize = True}
    Private ReadOnly _chkEnableTooling As System.Windows.Forms.CheckBox = New System.Windows.Forms.CheckBox() With {.Text = $"Enable {Globals.ThisAddIn.ToolFriendlyName.ToLower}", .AutoSize = True}
    Private ReadOnly _chkShowToolingLog As System.Windows.Forms.CheckBox = New System.Windows.Forms.CheckBox() With {.Text = "Tooling log", .AutoSize = True}

    ' State
    Private _htmlReady As Boolean = False
    Private ReadOnly _htmlQueue As New List(Of String)()
    Private _lastThinkingId As String = Nothing
    Private ReadOnly _history As New List(Of (Role As String, Content As String))()
    Private _knowledgeContent As String = Nothing
    Private _knowledgeFilePath As String = Nothing
    Private _welcomeInProgress As Integer = 0
    Private _personaSelectedThisSession As Boolean = False
    Private _isUpdatingPersistCheckbox As Boolean = False ' Prevents recursive event handling

    ' Autorespond state
    Private _autoRespondInProgress As Boolean = False
    Private _autoRespondCancelled As Boolean = False
    Private _autoRespondPersonaName As String = ""
    Private _autoRespondPersonaPrompt As String = ""
    Private _autoRespondMissionName As String = ""
    Private _autoRespondMissionPrompt As String = ""
    Private _autoRespondMaxRounds As Integer = 5
    Private _autoRespondBreakOff As String = DefaultAutoRespondBreakOff

    ' Alternate model support (new implementation matching Form1.vb pattern)
    Private _alternateModelSelected As Boolean = False
    Private _alternateModelConfig As ModelConfig = Nothing
    Private _alternateModelDisplayName As String = Nothing
    Private ReadOnly _modelSemaphore As New Threading.SemaphoreSlim(1, 1)

    ''' <summary>
    ''' Holds a persona definition loaded from a file, including its prompt and display metadata.
    ''' </summary>
    Private Structure PersonaEntry
        Public Name As String
        Public Prompt As String
        Public IsLocal As Boolean
        Public DisplayName As String
    End Structure
    Private _personas As New List(Of PersonaEntry)()

    ''' <summary>
    ''' Holds a mission definition loaded from a file.
    ''' </summary>
    Private Structure MissionEntry
        Public Name As String
        Public Prompt As String
        Public DisplayName As String
    End Structure
    Private _missions As New List(Of MissionEntry)()

    ''' <summary>
    ''' Helper class to track file loading results for knowledge loading.
    ''' </summary>
    Private Class KnowledgeLoadingContext
        Public Property GlobalDocumentCounter As Integer = 0
        Public Property LoadedFiles As New List(Of Tuple(Of String, Integer))() ' (path, charCount)
        Public Property FailedFiles As New List(Of String)()
        Public Property IgnoredFilesPerDir As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        Public Property EnableOCR As Boolean = False
        Public Property HasPdfFiles As Boolean = False

        ''' <summary>PDFs that heuristics suggest may contain images/scanned content but OCR was not performed.</summary>
        Public Property PdfsWithPossibleImages As New List(Of String)()

        ''' <summary>Maximum files to load from a single directory.</summary>
        Public Const MaxFilesPerDirectory As Integer = 50

        ''' <summary>Ask user confirmation if directory has more than this many files.</summary>
        Public Const ConfirmDirectoryFileCount As Integer = 10
    End Class

#End Region

#Region "Constructor"

    ''' <summary>
    ''' Initializes UI, loads configuration references, and wires event handlers.
    ''' </summary>
    ''' <param name="context">Shared configuration context providing INI settings and model configuration.</param>
    Public Sub New(context As ISharedContext)
        MyBase.New()
        _context = context

        Me.Text = $"Discuss this, {AssistantName}"
        Me.FormBorderStyle = FormBorderStyle.Sizable
        Me.StartPosition = FormStartPosition.Manual
        Me.MinimumSize = New System.Drawing.Size(780, 480)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0F)
        Try
            Me.Icon = Icon.FromHandle(New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard)).GetHicon())
        Catch
        End Try

        ' Layout
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
        Dim fiveLines = (_txtInput.Font.Height * 5) + 10
        _txtInput.MinimumSize = New System.Drawing.Size(0, fiveLines)
        _txtInput.Height = fiveLines

        Dim pnlButtons As New FlowLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .FlowDirection = FlowDirection.LeftToRight,
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .Padding = New Padding(0, 0, 0, 4)
        }
        pnlButtons.Controls.Add(_btnSend)
        pnlButtons.Controls.Add(_btnPersona)
        pnlButtons.Controls.Add(_btnMission)
        pnlButtons.Controls.Add(_btnEditPersona)
        pnlButtons.Controls.Add(_btnKnowledge)

        ' Show alternate model button if either second API is configured or an alternate INI exists
        If _context.INI_SecondAPI OrElse Not String.IsNullOrWhiteSpace(_context.INI_AlternateModelPath) Then
            UpdateAlternateModelButtonText()
            pnlButtons.Controls.Add(_btnAlternateModel)
        End If

        pnlButtons.Controls.Add(_btnClear)
        pnlButtons.Controls.Add(_btnSendToDoc)
        pnlButtons.Controls.Add(_btnClose)
        pnlButtons.Controls.Add(_btnAutoRespond)
        pnlButtons.Controls.Add(_btnSortOut)
        pnlButtons.Controls.Add(_btnTools)
        pnlButtons.Controls.Add(_chkEnableTooling)
        pnlButtons.Controls.Add(_chkIncludeActiveDoc)
        pnlButtons.Controls.Add(_chkPersistKnowledge)
        pnlButtons.Controls.Add(_chkShowToolingLog)


        table.Controls.Add(_chat, 0, 0)
        table.Controls.Add(_txtInput, 0, 1)
        table.Controls.Add(pnlButtons, 0, 2)
        Me.Controls.Add(table)

        _mdPipeline = New MarkdownPipelineBuilder().
            UseAdvancedExtensions().
            UseEmojiAndSmiley().
            UseSoftlineBreakAsHardlineBreak().
            Build()

        ' Event handlers
        AddHandler Me.Load, AddressOf OnLoadForm
        AddHandler Me.FormClosing, AddressOf OnFormClosing
        AddHandler Me.Activated, AddressOf OnActivated
        AddHandler _btnSend.Click, AddressOf OnSend
        AddHandler _btnClear.Click, AddressOf OnClear
        AddHandler _btnSendToDoc.Click, AddressOf OnSendToDoc
        AddHandler _btnClose.Click, AddressOf OnClose
        AddHandler _btnPersona.Click, AddressOf OnSelectPersona
        AddHandler _btnMission.Click, AddressOf OnSelectMission
        AddHandler _btnEditPersona.Click, AddressOf OnEditLocalPersona
        AddHandler _btnKnowledge.Click, AddressOf OnLoadKnowledge
        AddHandler _btnAlternateModel.Click, AddressOf OnAlternateModelClick
        AddHandler _txtInput.KeyDown, AddressOf OnInputKeyDown
        AddHandler _chat.DocumentCompleted, AddressOf Chat_DocumentCompleted
        AddHandler _chat.Navigating, AddressOf Chat_Navigating
        AddHandler _chat.NewWindow, AddressOf Chat_NewWindow
        AddHandler _chkIncludeActiveDoc.CheckedChanged, AddressOf OnIncludeActiveDocChanged
        AddHandler _chkPersistKnowledge.CheckedChanged, AddressOf OnPersistKnowledgeChanged
        AddHandler _btnAutoRespond.Click, AddressOf OnAutoRespondClick
        AddHandler _btnSortOut.Click, AddressOf OnSortOutClick
        AddHandler _btnTools.Click, AddressOf OnToolsClick
        AddHandler _chkEnableTooling.CheckedChanged, AddressOf OnEnableToolingChanged
        AddHandler _chkShowToolingLog.CheckedChanged, AddressOf OnShowToolingLogChanged

    End Sub

#End Region

#Region "Utility Methods"

    ''' <summary>
    ''' Gets the location context string for inclusion in prompts.
    ''' </summary>
    ''' <returns>Location context string.</returns>
    Private Function GetLocationContext() As String
        Dim location = If(_context?.INI_Location, "")
        If String.IsNullOrWhiteSpace(location) Then
            Return ""
        End If
        Return $"Location of user: {location}."
    End Function

    ''' <summary>
    ''' Gets the language instruction for LLM responses.
    ''' </summary>
    ''' <returns>Language instruction string.</returns>
    Private Function GetLanguageInstruction() As String
        Return "Always respond in the same language the user uses in their messages, regardless of the language of these instructions or the knowledge base."
    End Function

    ''' <summary>
    ''' Executes an action on the UI thread, marshaling via BeginInvoke when required.
    ''' </summary>
    ''' <param name="action">Action to execute on the UI thread.</param>
    Private Sub Ui(action As System.Action)
        If Me.IsDisposed Then Return
        If Me.InvokeRequired Then
            Try : Me.BeginInvoke(action) : Catch : End Try
        Else
            action.Invoke()
        End If
    End Sub

    ''' <summary>
    ''' Builds the window caption to reflect persona, mission, knowledge file, and model state.
    ''' </summary>
    Private Sub UpdateWindowTitle()
        Dim title = $"Discuss this, {_currentPersonaName}"

        ' Add mission to title if active
        If Not String.IsNullOrEmpty(_currentMissionName) Then
            title &= $" [{_currentMissionName}]"
        End If

        If Not String.IsNullOrEmpty(_knowledgeFilePath) Then
            title &= $" - {Path.GetFileName(_knowledgeFilePath)}"
        End If

        ' Show current model in title if alternate is selected
        If _alternateModelSelected AndAlso Not String.IsNullOrWhiteSpace(_alternateModelDisplayName) Then
            title &= $" (using {_alternateModelDisplayName})"
        End If

        Ui(Sub() Me.Text = title)
    End Sub

    ''' <summary>
    ''' Refreshes the Send button label with the current persona name.
    ''' </summary>
    Private Sub UpdateSendButtonText()
        Ui(Sub() _btnSend.Text = $"Send to {_currentPersonaName}")
    End Sub

    ''' <summary>
    ''' Returns a random adverb used to vary assistant tone.
    ''' </summary>
    ''' <returns>Randomly selected adverb string.</returns>
    Private Function GetRandomModifier() As String
        Return _randomModifiers(_rng.Next(_randomModifiers.Length))
    End Function

    ''' <summary>
    ''' Formats the current date for inclusion in LLM prompts.
    ''' </summary>
    ''' <returns>Formatted date string.</returns>
    Private Function GetDateContext() As String
        Dim now = DateTime.Now
        Return $"Today is {now:dd-MMM-yyyy}."
    End Function

    ''' <summary>
    ''' Gets the full path to the persisted knowledge file in the temp folder.
    ''' </summary>
    ''' <returns>Full path to the persisted knowledge file.</returns>
    Private Function GetPersistedKnowledgeFilePath() As String
        Return Path.Combine(Path.GetTempPath(), PersistedKnowledgeFileName)
    End Function

    ''' <summary>
    ''' Checks if a trigger placeholder at a given index is wrapped in XML tags.
    ''' </summary>
    Private Function IsWrappedInXml(prompt As String, idx As Integer, trigger As String) As Boolean
        Dim wrappedPattern As String = "<(?<name>[A-Za-z][\w\-]*)\b[^>]*>\s*" & Regex.Escape(trigger) & "\s*</\k<name>>"
        Dim matches As MatchCollection = Regex.Matches(prompt, wrappedPattern, RegexOptions.IgnoreCase)
        For Each m As Match In matches
            If idx >= m.Index AndAlso idx < m.Index + m.Length Then
                Return True
            End If
        Next
        Return False
    End Function

#End Region

#Region "Form Events"

    ''' <summary>
    ''' Shows (or brings forward) the form and focuses the input box.
    ''' </summary>
    ''' <param name="owner">Optional owner window.</param>
    Public Sub ShowRaised(Optional owner As IWin32Window = Nothing)
        If Me.WindowState = FormWindowState.Minimized Then Me.WindowState = FormWindowState.Normal
        If Not Me.Visible Then
            If owner IsNot Nothing Then Me.Show(owner) Else Me.Show()
        End If
        Me.Activate()
        _txtInput.Focus()
        _txtInput.SelectAll()
    End Sub

    ''' <summary>
    ''' Handles form activation; TopMost behavior is disabled.
    ''' </summary>
    Private Sub OnActivated(sender As Object, e As EventArgs)
        ' No longer applying TopMost behavior
    End Sub

    ''' <summary>
    ''' Persists the 'include active document' checkbox state when changed.
    ''' </summary>
    Private Sub OnIncludeActiveDocChanged(sender As Object, e As EventArgs)
        Try
            My.Settings.DiscussIncludeActiveDoc = _chkIncludeActiveDoc.Checked
            My.Settings.Save()
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Handles the 'Persist knowledge temporarily' checkbox state changes.
    ''' When checked: persists current knowledge to temp file.
    ''' When unchecked: prompts user and deletes temp file if confirmed.
    ''' </summary>
    Private Sub OnPersistKnowledgeChanged(sender As Object, e As EventArgs)
        If _isUpdatingPersistCheckbox Then Return

        Try
            Dim persistPath = GetPersistedKnowledgeFilePath()

            If _chkPersistKnowledge.Checked Then
                ' User checked the box - persist current knowledge if available
                If Not String.IsNullOrWhiteSpace(_cachedKnowledgeContent) Then
                    Try
                        File.WriteAllText(persistPath, _cachedKnowledgeContent, Encoding.UTF8)
                        AppendSystemMessage($"Knowledge persisted to temporary storage ({_cachedKnowledgeContent.Length:N0} characters).")
                    Catch ex As Exception
                        AppendSystemMessage($"Failed to persist knowledge: {ex.Message}")
                        ' Revert checkbox state
                        _isUpdatingPersistCheckbox = True
                        _chkPersistKnowledge.Checked = False
                        _isUpdatingPersistCheckbox = False
                        Return
                    End Try
                Else
                    AppendSystemMessage("No knowledge loaded to persist. Load knowledge first, then check this box.")
                End If
            Else
                ' User unchecked the box - ask before deleting
                If File.Exists(persistPath) Then
                    Dim answer = ShowCustomYesNoBox(
                        "Do you want to delete the persisted knowledge file? This cannot be undone if you quit Word.",
                        "Yes, delete", "No, keep it")

                    If answer = 1 Then
                        Try
                            File.Delete(persistPath)
                            AppendSystemMessage("Persisted knowledge file deleted.")
                        Catch ex As Exception
                            AppendSystemMessage($"Failed to delete persisted knowledge: {ex.Message}")
                        End Try
                    Else
                        ' User chose not to delete - revert checkbox
                        _isUpdatingPersistCheckbox = True
                        _chkPersistKnowledge.Checked = True
                        _isUpdatingPersistCheckbox = False
                        Return
                    End If
                End If
            End If

            ' Save checkbox state
            My.Settings.DiscussPersistKnowledge = _chkPersistKnowledge.Checked
            My.Settings.Save()

            ' Update tooltip
            UpdatePersistKnowledgeTooltip()

        Catch ex As Exception
            AppendSystemMessage($"Error handling persist knowledge setting: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Updates the tooltip for the persist knowledge checkbox based on its state.
    ''' </summary>
    Private Sub UpdatePersistKnowledgeTooltip()
        If _chkPersistKnowledge.Checked Then
            Dim persistPath = GetPersistedKnowledgeFilePath()
            _toolTip.SetToolTip(_chkPersistKnowledge, $"Currently stored in: {persistPath}")
        Else
            _toolTip.SetToolTip(_chkPersistKnowledge, "")
        End If
    End Sub

    ''' <summary>
    ''' Restores persisted settings, persona, mission, knowledge cache, transcript, and optionally triggers a welcome.
    ''' </summary>
    Private Async Sub OnLoadForm(sender As Object, e As EventArgs)
        ' Restore window position/size
        Try
            If My.Settings.DiscussFormLocation <> System.Drawing.Point.Empty AndAlso My.Settings.DiscussFormSize <> System.Drawing.Size.Empty Then
                Me.Location = My.Settings.DiscussFormLocation
                Me.Size = My.Settings.DiscussFormSize
            Else
                Dim area = Screen.PrimaryScreen.WorkingArea
                Dim w = Math.Max(Me.MinimumSize.Width, 860)
                Dim h = Math.Max(Me.MinimumSize.Height, 540)
                Me.Location = New System.Drawing.Point(area.Left + (area.Width - w) \ 2, area.Top + (area.Height - h) \ 2)
                Me.Size = New System.Drawing.Size(w, h)
            End If
        Catch
        End Try

        ' Load persisted settings
        Try : _chkIncludeActiveDoc.Checked = My.Settings.DiscussIncludeActiveDoc : Catch : _chkIncludeActiveDoc.Checked = False : End Try

        ' Load persist knowledge checkbox state (set flag to prevent event firing during initialization)
        _isUpdatingPersistCheckbox = True
        Try : _chkPersistKnowledge.Checked = My.Settings.DiscussPersistKnowledge : Catch : _chkPersistKnowledge.Checked = False : End Try
        _isUpdatingPersistCheckbox = False

        ' Update tooltip for persist checkbox
        UpdatePersistKnowledgeTooltip()

        ' Clean up persisted knowledge file if checkbox is not checked
        If Not _chkPersistKnowledge.Checked Then
            Try
                Dim persistPath = GetPersistedKnowledgeFilePath()
                If File.Exists(persistPath) Then
                    File.Delete(persistPath)
                End If
            Catch
            End Try
        End If

        ' Restore tooling checkbox state
        Try : _chkEnableTooling.Checked = My.Settings.DiscussEnableTooling : Catch : _chkEnableTooling.Checked = False : End Try

        ' Load personas
        LoadPersonas()

        ' Load missions
        LoadMissions()

        ' Update tooling controls based on current model
        UpdateToolingControlsState()

        ' Check if persona was previously saved - if not, force selection
        Dim savedPersona = ""
        Try
            savedPersona = My.Settings.DiscussSelectedPersona
        Catch
        End Try

        Dim personaRestoredFromSettings = False
        If Not String.IsNullOrEmpty(savedPersona) Then
            Dim found = _personas.FirstOrDefault(Function(p) p.Name.Equals(savedPersona, StringComparison.OrdinalIgnoreCase))
            If Not String.IsNullOrEmpty(found.Name) Then
                _currentPersonaName = found.Name
                _currentPersonaPrompt = found.Prompt
                personaRestoredFromSettings = True
            End If
        End If

        ' Restore mission if previously saved
        Try
            Dim savedMission = My.Settings.DiscussSelectedMission
            If Not String.IsNullOrEmpty(savedMission) Then
                Dim found = _missions.FirstOrDefault(Function(m) m.Name.Equals(savedMission, StringComparison.OrdinalIgnoreCase))
                If Not String.IsNullOrEmpty(found.Name) Then
                    _currentMissionName = found.Name
                    _currentMissionPrompt = found.Prompt
                End If
            End If
        Catch
        End Try

        UpdateWindowTitle()
        UpdateSendButtonText()

        InitializeChatHtml()

        ' Restore chat or load knowledge
        Dim hasChat = False
        Dim restoredHtmlHadAlternateModel = False
        Try
            ' First, restore _history from plain transcript (this ensures LLM sees the conversation)
            Dim savedTranscript = My.Settings.DiscussLastChat
            If Not String.IsNullOrEmpty(savedTranscript) Then
                RestoreHistoryFromTranscript(savedTranscript)
            End If

            ' Then restore the HTML display
            Dim savedHtml = My.Settings.DiscussLastChatHtml
            If Not String.IsNullOrEmpty(savedHtml) Then
                ' Check if the restored HTML contains an alternate/secondary model switch message
                restoredHtmlHadAlternateModel = ChatHtmlIndicatesAlternateModel(savedHtml)
                AppendHtml(savedHtml)
                hasChat = True
            ElseIf Not String.IsNullOrEmpty(savedTranscript) Then
                AppendTranscriptToHtml(savedTranscript)
                hasChat = True
            End If
        Catch
        End Try

        ' If restored chat indicated an alternate model was active, notify user we're back on primary
        If hasChat AndAlso restoredHtmlHadAlternateModel Then
            ' Ensure alternate model state is reset (it should be by default, but be explicit)
            _alternateModelSelected = False
            _alternateModelConfig = Nothing
            _alternateModelDisplayName = Nothing
            UpdateAlternateModelButtonText()

            ' Notify user in chat that we're back on primary
            AppendSystemMessage($"Session restored. Now using primary model ({_context.INI_Model}).")
        End If

        ' Restore knowledge using the new loading flow
        Await RestoreKnowledgeAsync()

        ' Force persona selection on first run after Word starts (if not restored from settings)
        If Not personaRestoredFromSettings AndAlso _personas.Count > 0 AndAlso Not _personaSelectedThisSession Then
            OnSelectPersona(Nothing, EventArgs.Empty)
            _personaSelectedThisSession = True
        End If

        ' Prompt for knowledge if not available
        If String.IsNullOrEmpty(_knowledgeContent) AndAlso Not hasChat Then
            Await PromptForKnowledgeAsync()
        End If

        If Not hasChat Then
            Await SafeGenerateWelcomeAsync()
        End If
    End Sub

    ''' <summary>
    ''' Restores knowledge from various sources in priority order:
    ''' 1. Runtime cache (if Word hasn't been restarted)
    ''' 2. Persisted temp file (if checkbox is checked)
    ''' 3. Previously saved file path from settings
    ''' </summary>
    Private Async Function RestoreKnowledgeAsync() As Task
        ' 1. Check runtime cache first (survives form close but not Word restart)
        If Not String.IsNullOrEmpty(_cachedKnowledgeContent) AndAlso Not String.IsNullOrEmpty(_cachedKnowledgeFilePath) Then
            _knowledgeContent = _cachedKnowledgeContent
            _knowledgeFilePath = _cachedKnowledgeFilePath
            UpdateWindowTitle()
            Return
        End If

        ' 2. If persist checkbox is checked, try to load from temp file
        If _chkPersistKnowledge.Checked Then
            Dim persistPath = GetPersistedKnowledgeFilePath()
            If File.Exists(persistPath) Then
                Try
                    _knowledgeContent = File.ReadAllText(persistPath, Encoding.UTF8)
                    _knowledgeFilePath = "(Persisted Knowledge)"

                    ' Update runtime cache
                    _cachedKnowledgeContent = _knowledgeContent
                    _cachedKnowledgeFilePath = _knowledgeFilePath

                    UpdateWindowTitle()
                    AppendSystemMessage($"Knowledge restored from persisted storage ({_knowledgeContent.Length:N0} characters).")
                    Return
                Catch ex As Exception
                    AppendSystemMessage($"Failed to restore persisted knowledge: {ex.Message}")
                End Try
            End If
        End If

        ' 3. Try to reload from saved file path in settings
        Try
            Dim savedPath = My.Settings.DiscussKnowledgePath
            If Not String.IsNullOrEmpty(savedPath) AndAlso File.Exists(savedPath) Then
                ShowAssistantThinking()
                Dim result = Await LoadSingleKnowledgeFileAsync(savedPath, False, False)
                _knowledgeContent = result.Content
                RemoveAssistantThinking()

                If Not String.IsNullOrWhiteSpace(_knowledgeContent) Then
                    _knowledgeFilePath = savedPath

                    ' Update runtime cache
                    _cachedKnowledgeContent = _knowledgeContent
                    _cachedKnowledgeFilePath = _knowledgeFilePath

                    ' Persist if checkbox is checked
                    If _chkPersistKnowledge.Checked Then
                        PersistKnowledgeToTempFile()
                    End If

                    UpdateWindowTitle()
                End If
            End If
        Catch
        End Try
    End Function

    ''' <summary>
    ''' Persists the current knowledge content to the temp file.
    ''' </summary>
    Private Sub PersistKnowledgeToTempFile()
        If String.IsNullOrWhiteSpace(_cachedKnowledgeContent) Then Return

        Try
            Dim persistPath = GetPersistedKnowledgeFilePath()
            File.WriteAllText(persistPath, _cachedKnowledgeContent, Encoding.UTF8)
        Catch
            ' Silently fail - not critical
        End Try
    End Sub

    ''' <summary>
    ''' Persists geometry, transcript, persona, mission, knowledge path, and checkbox state on close.
    ''' </summary>
    Private Sub OnFormClosing(sender As Object, e As FormClosingEventArgs)
        Try
            PersistTranscriptLimited()
            PersistChatHtml()
            If Me.WindowState = FormWindowState.Normal Then
                My.Settings.DiscussFormLocation = Me.Location
                My.Settings.DiscussFormSize = Me.Size
            Else
                My.Settings.DiscussFormLocation = Me.RestoreBounds.Location
                My.Settings.DiscussFormSize = Me.RestoreBounds.Size
            End If
            My.Settings.DiscussIncludeActiveDoc = _chkIncludeActiveDoc.Checked
            My.Settings.DiscussPersistKnowledge = _chkPersistKnowledge.Checked
            My.Settings.DiscussSelectedPersona = _currentPersonaName
            My.Settings.DiscussSelectedMission = _currentMissionName
            My.Settings.DiscussKnowledgePath = If(_knowledgeFilePath, "")
            My.Settings.DiscussEnableTooling = _chkEnableTooling.Checked
            My.Settings.Save()
        Catch
        End Try
    End Sub

#End Region

#Region "Alternate Model Handling"

    ''' <summary>
    ''' Sets the alternate-model button caption according to availability and selection state.
    ''' </summary>
    Private Sub UpdateAlternateModelButtonText()
        If Not String.IsNullOrWhiteSpace(_context.INI_AlternateModelPath) Then
            _btnAlternateModel.Text = If(_alternateModelSelected, "Primary Model", "Alternate Model")
        Else
            _btnAlternateModel.Text = "Switch Model"
        End If
    End Sub

    ''' <summary>
    ''' Handles alternate model toggling or selection, mirroring Form1 pattern.
    ''' </summary>
    Private Sub OnAlternateModelClick(sender As Object, e As EventArgs)
        If Not String.IsNullOrWhiteSpace(_context.INI_AlternateModelPath) Then
            ' If an alternate is already active -> switch back to primary without dialog
            If _alternateModelSelected Then
                _alternateModelSelected = False
                _alternateModelConfig = Nothing
                _alternateModelDisplayName = Nothing
                UpdateAlternateModelButtonText()
                UpdateWindowTitle()
                AppendSystemMessage($"Switched back to primary model ({_context.INI_Model}).")
                Return
            End If

            ' Pre-check: verify the alternate model file exists and has content
            Dim altPath = ExpandEnvironmentVariables(_context.INI_AlternateModelPath)
            If String.IsNullOrWhiteSpace(altPath) OrElse Not File.Exists(altPath) Then
                AppendSystemMessage("Alternate model configuration file not found.")
                Return
            End If

            ' Selecting an alternate
            SharedMethods.LastAlternateModel = "" ' sentinel
            Dim ok As Boolean = SharedMethods.ShowModelSelection(
                _context,
                _context.INI_AlternateModelPath,
                "Alternate Model",
                "Select the alternate model you want to use:",
                "",
                2
            )
            If Not ok Then
                ' User cancelled
                Return
            End If

            ' The selector applies the chosen model to the context at this point.
            ' Snapshot it, then restore the original immediately so globals remain clean.
            Dim justApplied As ModelConfig = SharedMethods.GetCurrentConfig(_context)

            If SharedMethods.originalConfigLoaded Then
                SharedMethods.RestoreDefaults(_context, SharedMethods.originalConfig)
            End If
            SharedMethods.originalConfigLoaded = False

            Dim userChoseAlternate As Boolean = Not String.IsNullOrWhiteSpace(SharedMethods.LastAlternateModel)

            If userChoseAlternate Then
                _alternateModelSelected = True
                _alternateModelConfig = justApplied
                _alternateModelDisplayName = SharedMethods.LastAlternateModel
                AppendSystemMessage($"Switched to alternate model: {_alternateModelDisplayName}")
            Else
                _alternateModelSelected = False
                _alternateModelConfig = Nothing
                _alternateModelDisplayName = Nothing
            End If

            UpdateAlternateModelButtonText()
            UpdateWindowTitle()
            UpdateToolingControlsState()

        Else
            ' Legacy behavior: simple toggle to secondary model (if configured)
            If _context.INI_SecondAPI Then
                ' Toggle between primary and secondary
                If _alternateModelSelected Then
                    _alternateModelSelected = False
                    _alternateModelConfig = Nothing
                    _alternateModelDisplayName = Nothing
                    AppendSystemMessage($"Switched back to primary model ({_context.INI_Model}).")
                Else
                    _alternateModelSelected = True
                    _alternateModelDisplayName = _context.INI_Model_2
                    AppendSystemMessage($"Switched to secondary model: {_alternateModelDisplayName}")
                End If
                UpdateAlternateModelButtonText()
                UpdateWindowTitle()
                UpdateToolingControlsState()

            End If
        End If
    End Sub

    ''' <summary>
    ''' Runs an LLM request while temporarily applying any selected alternate model, restoring afterward.
    ''' </summary>
    ''' <summary>
    ''' Runs an LLM request while temporarily applying any selected alternate model, restoring afterward.
    ''' Supports tooling when enabled and model supports it.
    ''' </summary>
    Private Async Function CallLlmWithSelectedModelAsync(systemPrompt As String, userPrompt As String) As Task(Of String)
        Await _modelSemaphore.WaitAsync().ConfigureAwait(False)
        Dim backupConfig As ModelConfig = Nothing
        Dim appliedAlternate As Boolean = False
        Dim useSecondApi As Boolean = False

        Try
            ' If the user selected an alternate model, apply it to the context as the "second model" just for this call.
            If _alternateModelSelected AndAlso _alternateModelConfig IsNot Nothing Then
                ' Back up current config (the "original state at rest")
                backupConfig = SharedMethods.GetCurrentConfig(_context)

                ' Apply the selected alternate config
                SharedMethods.ApplyModelConfig(_context, _alternateModelConfig)
                appliedAlternate = True

                ' Enforce second API usage for alternate models
                useSecondApi = True
            ElseIf _alternateModelSelected AndAlso _alternateModelConfig Is Nothing AndAlso _context.INI_SecondAPI Then
                ' Legacy toggle: use second API without config swap
                useSecondApi = True
            End If

            ' Check if tooling should be used
            If ShouldUseTooling() AndAlso EnsureToolsSelected() Then
                ' Execute via tooling loop
                Return Await Globals.ThisAddIn.ExecuteToolingLoop(
                    systemPrompt,
                    "",
                    _selectedToolsForChat,
                    useSecondApi,
                    fullPromptOverride:=userPrompt,
                    hideSplash:=True,
                    hideLogWindow:=Not _chkShowToolingLog.Checked).ConfigureAwait(False)
            Else
                ' Standard LLM call
                Return Await LLM(_context,
                                 systemPrompt,
                                 userPrompt,
                                 "",
                                 "",
                                 0,
                                 useSecondApi,
                                 True).ConfigureAwait(False)
            End If

        Finally
            ' Always restore the original config after the call so the rest of the add-in sees the original state.
            If appliedAlternate AndAlso backupConfig IsNot Nothing Then
                SharedMethods.RestoreDefaults(_context, backupConfig)
            End If
            _modelSemaphore.Release()
        End Try
    End Function

#End Region


#Region "Tooling Support"

    ''' <summary>
    ''' Updates enabled state of tooling controls based on current model's tooling support.
    ''' </summary>
    Private Sub UpdateToolingControlsState()
        Dim currentConfig As ModelConfig = Nothing

        If _alternateModelSelected AndAlso _alternateModelConfig IsNot Nothing Then
            currentConfig = _alternateModelConfig
        Else
            currentConfig = SharedMethods.GetCurrentConfig(_context)
        End If

        Dim supportsTooling As Boolean = SharedMethods.ModelSupportsTooling(currentConfig)

        _chkEnableTooling.Enabled = supportsTooling
        _btnTools.Enabled = supportsTooling
        _chkShowToolingLog.Enabled = supportsTooling AndAlso _chkEnableTooling.Checked

        If Not supportsTooling Then
            _chkEnableTooling.Checked = False
            _chkShowToolingLog.Checked = False
            _selectedToolsForChat = Nothing
        End If
    End Sub

    ''' <summary>
    ''' Handles changes to the "Tooling log" checkbox. The checked state is consumed when executing the tooling loop
    ''' to decide whether to show or hide the tooling log window.
    ''' </summary>
    ''' <param name="sender">The event source.</param>
    ''' <param name="e">Event arguments.</param>

    Private Sub OnShowToolingLogChanged(sender As Object, e As EventArgs)
        ' No special handling needed - just uses the Checked state when calling ExecuteToolingLoop
    End Sub

    ''' <summary>
    ''' Handles the Tools button click - opens tool selection dialog.
    ''' </summary>
    Private Sub OnToolsClick(sender As Object, e As EventArgs)
        Try
            Dim selectedTools = Globals.ThisAddIn.SelectToolsForSession(forceDialog:=True, Globals.ThisAddIn.ToolFriendlyName)
            If selectedTools IsNot Nothing Then
                _selectedToolsForChat = selectedTools
            End If
        Catch ex As Exception
            AppendSystemMessage($"Error selecting tools: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Handles the Enable Tooling checkbox change.
    ''' </summary>
    Private Sub OnEnableToolingChanged(sender As Object, e As EventArgs)
        If Not _chkEnableTooling.Checked Then
            _selectedToolsForChat = Nothing
        End If
        ' Update log checkbox enabled state
        _chkShowToolingLog.Enabled = _chkEnableTooling.Checked AndAlso _chkEnableTooling.Enabled
    End Sub

    ''' <summary>
    ''' Determines if tooling should be used for the current call.
    ''' </summary>
    Private Function ShouldUseTooling() As Boolean
        If Not _chkEnableTooling.Checked Then Return False

        Dim currentConfig As ModelConfig = Nothing
        If _alternateModelSelected AndAlso _alternateModelConfig IsNot Nothing Then
            currentConfig = _alternateModelConfig
        Else
            currentConfig = SharedMethods.GetCurrentConfig(_context)
        End If

        Return SharedMethods.ModelSupportsTooling(currentConfig)
    End Function

    ''' <summary>
    ''' Ensures tools are selected for the session if tooling is enabled.
    ''' </summary>
    Private Function EnsureToolsSelected() As Boolean
        If _selectedToolsForChat IsNot Nothing AndAlso _selectedToolsForChat.Count > 0 Then
            Return True
        End If

        _selectedToolsForChat = Globals.ThisAddIn.SelectToolsForSession(forceDialog:=False)
        Return _selectedToolsForChat IsNot Nothing AndAlso _selectedToolsForChat.Count > 0
    End Function

#End Region

#Region "Persona Management"

    ''' <summary>
    ''' Loads persona definitions from configured local and global files into memory.
    ''' </summary>
    Private Sub LoadPersonas()
        _personas.Clear()

        Dim localPath = ExpandEnvironmentVariables(If(_context?.INI_DiscussInkyPathLocal, ""))
        Dim globalPath = ExpandEnvironmentVariables(If(_context?.INI_DiscussInkyPath, ""))

        Dim localLoaded = False
        Dim globalLoaded = False

        ' Load local personas first (marked with (local))
        If Not String.IsNullOrWhiteSpace(localPath) Then
            localLoaded = LoadPersonasFromFile(localPath, isLocal:=True)
        End If

        ' Load global personas
        If Not String.IsNullOrWhiteSpace(globalPath) Then
            globalLoaded = LoadPersonasFromFile(globalPath, isLocal:=False)
        End If

        ' Show error only if both paths are configured but neither loaded any personas
        If _personas.Count = 0 Then
            If Not String.IsNullOrWhiteSpace(localPath) OrElse Not String.IsNullOrWhiteSpace(globalPath) Then
                AppendSystemMessage("No personas could be loaded. Please check your persona configuration files.")
            End If
        End If
    End Sub

    ''' <summary>
    ''' Parses a persona file, appending entries and marking whether they are local.
    ''' </summary>
    Private Function LoadPersonasFromFile(filePath As String, isLocal As Boolean) As Boolean
        ' Must be a file, not a directory
        If String.IsNullOrWhiteSpace(filePath) Then
            Return False
        End If

        If Directory.Exists(filePath) Then
            AppendSystemMessage($"Persona path must be a file, not a directory: {filePath}")
            Return False
        End If

        If Not File.Exists(filePath) Then
            Return False
        End If

        Dim loadedAny = False
        Try
            For Each rawLine In File.ReadAllLines(filePath, Encoding.UTF8)
                Dim line = If(rawLine, "").Trim()

                ' Skip empty lines and comments
                If line.Length = 0 OrElse line.StartsWith(";", StringComparison.Ordinal) Then
                    Continue For
                End If

                ' Parse Name|Prompt format
                Dim pipeIdx = line.IndexOf("|"c)
                If pipeIdx < 1 Then Continue For

                Dim name = line.Substring(0, pipeIdx).Trim()
                Dim prompt = line.Substring(pipeIdx + 1).Trim()

                If name.Length = 0 OrElse prompt.Length = 0 Then Continue For

                ' Create unique display name
                Dim displayName = name & If(isLocal, " (local)", "")
                displayName = MakeUniqueDisplay(displayName, _personas.Select(Function(p) p.DisplayName).ToList())

                _personas.Add(New PersonaEntry With {
                    .Name = name,
                    .Prompt = prompt,
                    .IsLocal = isLocal,
                    .DisplayName = displayName
                })
                loadedAny = True
            Next
        Catch ex As Exception
            AppendSystemMessage($"Error loading persona file: {ex.Message}")
            Return False
        End Try

        Return loadedAny
    End Function

    ''' <summary>
    ''' Ensures persona display names are unique by appending numeric suffixes.
    ''' </summary>
    Private Function MakeUniqueDisplay(baseText As String, existing As ICollection(Of String)) As String
        If Not existing.Contains(baseText) Then Return baseText
        Dim n = 2
        While True
            Dim candidate = baseText & " [" & n.ToString() & "]"
            If Not existing.Contains(candidate) Then Return candidate
            n += 1
        End While
    End Function

    ''' <summary>
    ''' Shows persona picker and applies the chosen persona prompt.
    ''' </summary>
    Private Sub OnSelectPersona(sender As Object, e As EventArgs)
        If _personas.Count = 0 Then
            ShowCustomMessageBox("No personas configured. Please configure INI_DiscussInkyPath or INI_DiscussInkyPathLocal in your settings.",
                                 extraButtonText:="Edit Local Personas",
                                 extraButtonAction:=Sub() OnEditLocalPersona(Nothing, EventArgs.Empty))
            Return
        End If

        ' Build selection items
        Dim items As New List(Of SelectionItem)()
        For i = 0 To _personas.Count - 1
            items.Add(New SelectionItem(_personas(i).DisplayName, i + 1))
        Next

        ' Find current selection
        Dim defaultVal = 1
        For i = 0 To _personas.Count - 1
            If _personas(i).Name.Equals(_currentPersonaName, StringComparison.OrdinalIgnoreCase) Then
                defaultVal = i + 1
                Exit For
            End If
        Next

        Dim result = SelectValue(items, defaultVal, "Select the persona discussing:", AN & " - Select Persona")

        If result > 0 AndAlso result <= _personas.Count Then
            Dim selected = _personas(result - 1)
            _currentPersonaName = selected.Name
            _currentPersonaPrompt = selected.Prompt
            _personaSelectedThisSession = True
            UpdateWindowTitle()
            UpdateSendButtonText()

            Try
                My.Settings.DiscussSelectedPersona = _currentPersonaName
                My.Settings.Save()
            Catch
            End Try

            AppendSystemMessage($"Persona changed to: {_currentPersonaName}")
        End If
    End Sub

    ''' <summary>
    ''' Ensures the local persona file exists and opens it in the shared text editor.
    ''' </summary>
    Private Sub OnEditLocalPersona(sender As Object, e As EventArgs)
        Dim localPath = ExpandEnvironmentVariables(If(_context?.INI_DiscussInkyPathLocal, ""))

        If String.IsNullOrWhiteSpace(localPath) Then
            ShowCustomMessageBox("INI_DiscussInkyPathLocal is not configured in your settings.")
            Return
        End If

        ' Create directory if needed
        Dim dir = Path.GetDirectoryName(localPath)
        If Not String.IsNullOrWhiteSpace(dir) AndAlso Not Directory.Exists(dir) Then
            Try
                Directory.CreateDirectory(dir)
            Catch ex As Exception
                ShowCustomMessageBox($"Cannot create directory: {ex.Message}")
                Return
            End Try
        End If


        ' Create file with sample content if it doesn't exist or contains only whitespace
        Dim needsSampleContent As Boolean = False
        If Not File.Exists(localPath) Then
            needsSampleContent = True
        Else
            Try
                Dim content As String = File.ReadAllText(localPath, System.Text.Encoding.UTF8)
                needsSampleContent = String.IsNullOrWhiteSpace(content)
            Catch
                needsSampleContent = True
            End Try
        End If

        If needsSampleContent Then


            Try
                File.WriteAllText(localPath,
                    "; Discuss This Local Personas" & vbCrLf &
                    "; Format: Name|System Prompt" & vbCrLf &
                    "; Lines starting with ; are comments" & vbCrLf &
                    vbCrLf &
                    "Teacher|You are a teacher and will do an exam with the user based on the knowledge you will be provided. Check the responses and provide feedback." & vbCrLf & vbCrLf &
                    "Summarizer|Summarize the knowledge document for the user in a clear and concise way. Answer follow-up questions about the content." & vbCrLf,
                    Encoding.UTF8)
            Catch ex As Exception
                ShowCustomMessageBox($"Cannot create file: {ex.Message}")
                Return
            End Try
        End If

        ShowTextFileEditor(localPath, $"{AN} - Edit Local Personas (changes active after restart):", False, _context)
    End Sub

#End Region

#Region "Mission Management"

    ''' <summary>
    ''' Derives the mission file path from the persona lib path.
    ''' Format: [personafilename]-missions.txt
    ''' Prefers local path, falls back to global.
    ''' </summary>
    ''' <returns>Full path to the mission file, or empty string if no persona path is configured.</returns>
    Private Function GetMissionFilePath() As String
        ' Prefer local path
        Dim personaPath = ExpandEnvironmentVariables(If(_context?.INI_DiscussInkyPathLocal, ""))

        ' Fall back to global path
        If String.IsNullOrWhiteSpace(personaPath) Then
            personaPath = ExpandEnvironmentVariables(If(_context?.INI_DiscussInkyPath, ""))
        End If

        If String.IsNullOrWhiteSpace(personaPath) Then
            Return ""
        End If

        ' Build mission file path: [name]-missions.txt
        Dim dir = Path.GetDirectoryName(personaPath)
        Dim nameWithoutExt = Path.GetFileNameWithoutExtension(personaPath)
        Dim missionFileName = nameWithoutExt & "-missions.txt"

        Return Path.Combine(If(dir, ""), missionFileName)
    End Function

    ''' <summary>
    ''' Loads mission definitions from the mission file into memory.
    ''' </summary>
    Private Sub LoadMissions()
        _missions.Clear()

        Dim missionPath = GetMissionFilePath()
        If String.IsNullOrWhiteSpace(missionPath) Then
            Return
        End If

        If Not File.Exists(missionPath) Then
            ' File doesn't exist yet - that's okay, user can create it via Edit
            Return
        End If

        Try
            For Each rawLine In File.ReadAllLines(missionPath, Encoding.UTF8)
                Dim line = If(rawLine, "").Trim()

                ' Skip empty lines and comments
                If line.Length = 0 OrElse line.StartsWith(";", StringComparison.Ordinal) Then
                    Continue For
                End If

                ' Parse Name|Prompt format
                Dim pipeIdx = line.IndexOf("|"c)
                If pipeIdx < 1 Then Continue For

                Dim name = line.Substring(0, pipeIdx).Trim()
                Dim prompt = line.Substring(pipeIdx + 1).Trim()

                If name.Length = 0 OrElse prompt.Length = 0 Then Continue For

                ' Create unique display name
                Dim displayName = MakeUniqueDisplay(name, _missions.Select(Function(m) m.DisplayName).ToList())

                _missions.Add(New MissionEntry With {
                    .Name = name,
                    .Prompt = prompt,
                    .DisplayName = displayName
                })
            Next
        Catch ex As Exception
            AppendSystemMessage($"Error loading mission file: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Creates a sample mission file if it doesn't exist or is empty.
    ''' </summary>
    ''' <param name="missionPath">Path to the mission file.</param>
    Private Sub EnsureMissionFileExists(missionPath As String)
        If String.IsNullOrWhiteSpace(missionPath) Then Return

        ' Create directory if needed
        Dim dir = Path.GetDirectoryName(missionPath)
        If Not String.IsNullOrWhiteSpace(dir) AndAlso Not Directory.Exists(dir) Then
            Try
                Directory.CreateDirectory(dir)
            Catch
                Return
            End Try
        End If

        ' Check if file needs sample content
        Dim needsSampleContent = False
        If Not File.Exists(missionPath) Then
            needsSampleContent = True
        Else
            Try
                Dim content = File.ReadAllText(missionPath, Encoding.UTF8)
                needsSampleContent = String.IsNullOrWhiteSpace(content)
            Catch
                needsSampleContent = True
            End Try
        End If

        If needsSampleContent Then
            Try
                File.WriteAllText(missionPath,
                    "; Discuss This Missions" & vbCrLf &
                    "; Format: Name|Mission Prompt" & vbCrLf &
                    "; Lines starting with ; are comments" & vbCrLf &
                    "; Missions provide specific behavioral targets for the conversation." & vbCrLf &
                    vbCrLf &
                    "Devil's Advocate|Challenge every argument presented. Find weaknesses, inconsistencies, and counter-arguments. Push back firmly but constructively, forcing a thorough defense of each position." & vbCrLf & vbCrLf &
                    "Problem Solver|Help find a solution to the problem at hand. Ask probing questions to understand the full context. Encourage exploration of alternatives while remaining constructive and focused on actionable outcomes." & vbCrLf & vbCrLf &
                    "Witness Simulation|Defend the documented position as stated in the knowledge base. Respond as if being questioned, staying consistent with the documented facts. Do not volunteer information beyond what is documented." & vbCrLf & vbCrLf &
                    "Cross-Examination|Systematically question the documented statements to test their credibility and consistency. Look for gaps, contradictions, or areas requiring clarification. Press for specifics and challenge vague assertions." & vbCrLf & vbCrLf &
                    "Only One Paragraph|Limit your response always to a maximum of one paragraph." & vbCrLf & vbCrLf &
                    "Only One Sentence|Limit your response always to a maximum of one sentence.",
                    Encoding.UTF8)
            Catch
                ' Silently fail
            End Try
        End If
    End Sub

    ''' <summary>
    ''' Shows mission picker with "No mission" as first option and "Edit mission library" as last.
    ''' </summary>
    Private Sub OnSelectMission(sender As Object, e As EventArgs)
        Dim missionPath = GetMissionFilePath()

        If String.IsNullOrWhiteSpace(missionPath) Then
            ShowCustomMessageBox("No persona library is configured. Missions require a persona library path (INI_DiscussInkyPathLocal or INI_DiscussInkyPath).")
            Return
        End If

        ' Ensure mission file exists with samples if needed
        EnsureMissionFileExists(missionPath)

        ' Reload missions to pick up any changes
        LoadMissions()

        ' Build selection items
        Dim items As New List(Of SelectionItem)()

        ' First item: "No mission"
        Const NoMissionValue As Integer = -1
        items.Add(New SelectionItem("No mission", NoMissionValue))

        ' Mission items
        For i = 0 To _missions.Count - 1
            items.Add(New SelectionItem(_missions(i).DisplayName, i + 1))
        Next

        ' Last item: "Edit mission library"
        Const EditMissionValue As Integer = -2
        items.Add(New SelectionItem("Edit mission library...", EditMissionValue))

        ' Find current selection
        Dim defaultVal = NoMissionValue
        If Not String.IsNullOrEmpty(_currentMissionName) Then
            For i = 0 To _missions.Count - 1
                If _missions(i).Name.Equals(_currentMissionName, StringComparison.OrdinalIgnoreCase) Then
                    defaultVal = i + 1
                    Exit For
                End If
            Next
        End If

        Dim result = SelectValue(items, defaultVal, "Select a mission (optional behavioral target):", AN & " - Select Mission")

        If result = NoMissionValue Then
            ' User selected "No mission"
            If Not String.IsNullOrEmpty(_currentMissionName) Then
                _currentMissionName = ""
                _currentMissionPrompt = ""
                UpdateWindowTitle()

                Try
                    My.Settings.DiscussSelectedMission = ""
                    My.Settings.Save()
                Catch
                End Try

                AppendSystemMessage("Mission cleared.")
            End If
        ElseIf result = EditMissionValue Then
            ' User selected "Edit mission library"
            ShowTextFileEditor(missionPath, $"{AN} - Edit Missions (changes active after reload):", False, _context)

            ' Reload and show selection again
            OnSelectMission(sender, e)
        ElseIf result > 0 AndAlso result <= _missions.Count Then
            ' User selected a mission
            Dim selected = _missions(result - 1)
            _currentMissionName = selected.Name
            _currentMissionPrompt = selected.Prompt
            UpdateWindowTitle()

            Try
                My.Settings.DiscussSelectedMission = _currentMissionName
                My.Settings.Save()
            Catch
            End Try

            AppendSystemMessage($"Mission set to: {_currentMissionName}")
        End If
        ' If result = 0 (cancelled), do nothing - keep current mission
    End Sub

#End Region

#Region "Knowledge File Management"

    ''' <summary>
    ''' Button handler that launches the knowledge file/directory picker.
    ''' </summary>
    Private Async Sub OnLoadKnowledge(sender As Object, e As EventArgs)
        Await PromptForKnowledgeAsync()
    End Sub

    ''' <summary>
    ''' Prompts the user for a knowledge file or directory, loads content, caches it, and updates state.
    ''' Supports loading multiple files from a directory with unified document numbering.
    ''' </summary>
    Private Async Function PromptForKnowledgeAsync() As Task
        Try
            Globals.ThisAddIn.DragDropFormLabel = "... a document you want to use as a knowledge file or folder to use all documents contained therein, or click Browse"
            Globals.ThisAddIn.DragDropFormFilter = ""

            Dim selectedPath As String = ""

            Using frm As New DragDropForm(DragDropMode.FileOrDirectory)
                If frm.ShowDialog() = DialogResult.OK Then
                    selectedPath = frm.SelectedFilePath
                End If
            End Using

            Globals.ThisAddIn.DragDropFormLabel = ""
            Globals.ThisAddIn.DragDropFormFilter = ""

            If String.IsNullOrWhiteSpace(selectedPath) Then
                ' No file selected - check if there's existing knowledge to delete
                If Not String.IsNullOrWhiteSpace(_knowledgeContent) Then
                    Dim answer = ShowCustomYesNoBox(
                        "No file was selected. Do you want to delete the currently loaded knowledge?",
                        "Yes, delete knowledge", "No, keep it")

                    If answer = 1 Then
                        ' Delete knowledge
                        _knowledgeContent = Nothing
                        _knowledgeFilePath = Nothing
                        _cachedKnowledgeContent = Nothing
                        _cachedKnowledgeFilePath = Nothing

                        ' Delete persisted knowledge file if it exists
                        Try
                            Dim persistPath = GetPersistedKnowledgeFilePath()
                            If File.Exists(persistPath) Then
                                File.Delete(persistPath)
                            End If
                        Catch
                        End Try

                        ' Clear the saved path in settings
                        Try
                            My.Settings.DiscussKnowledgePath = ""
                            My.Settings.Save()
                        Catch
                        End Try

                        UpdateWindowTitle()
                        AppendSystemMessage("Knowledge deleted.")
                    End If
                End If
                Return
            End If

            ' Determine if it's a file or directory
            Dim isDirectory = Directory.Exists(selectedPath)
            Dim isFile = File.Exists(selectedPath)

            If Not isFile AndAlso Not isDirectory Then
                AppendSystemMessage("Selected path does not exist.")
                Return
            End If

            ' Create loading context
            Dim ctx As New KnowledgeLoadingContext()

            ' Collect files to process
            Dim filesToProcess As New List(Of String)()

            If isFile Then
                filesToProcess.Add(selectedPath)
                ' Check if it's a PDF
                If Path.GetExtension(selectedPath).Equals(".pdf", StringComparison.OrdinalIgnoreCase) Then
                    ctx.HasPdfFiles = True
                End If
            Else
                ' It's a directory - collect supported files
                Dim allFiles = Directory.GetFiles(selectedPath, "*.*", SearchOption.TopDirectoryOnly)
                Dim ignoredCount = 0

                For Each f In allFiles
                    Dim ext = Path.GetExtension(f).ToLowerInvariant()
                    If SupportedKnowledgeExtensions.Contains(ext) Then
                        filesToProcess.Add(f)
                        If ext = ".pdf" Then
                            ctx.HasPdfFiles = True
                        End If
                    Else
                        ignoredCount += 1
                    End If
                Next

                If ignoredCount > 0 Then
                    ctx.IgnoredFilesPerDir(selectedPath) = ignoredCount
                End If

                ' Check file count limits
                If filesToProcess.Count > KnowledgeLoadingContext.MaxFilesPerDirectory Then
                    Dim truncateAnswer = ShowCustomYesNoBox(
                        $"The directory contains {filesToProcess.Count} supported files, but the maximum is {KnowledgeLoadingContext.MaxFilesPerDirectory}." & vbCrLf & vbCrLf &
                        $"Only the first {KnowledgeLoadingContext.MaxFilesPerDirectory} files will be loaded. Continue?",
                        "Yes, continue", "No, abort")
                    If truncateAnswer <> 1 Then
                        Return
                    End If
                    filesToProcess = filesToProcess.Take(KnowledgeLoadingContext.MaxFilesPerDirectory).ToList()
                ElseIf filesToProcess.Count > KnowledgeLoadingContext.ConfirmDirectoryFileCount Then
                    Dim confirmAnswer = ShowCustomYesNoBox(
                        $"The directory contains {filesToProcess.Count} files to load. Continue?",
                        "Yes, continue", "No, abort")
                    If confirmAnswer <> 1 Then
                        Return
                    End If
                End If

                If filesToProcess.Count = 0 Then
                    AppendSystemMessage($"No supported files found in directory '{selectedPath}'.")
                    Return
                End If
            End If

            ' Ask about OCR if there are PDF files AND OCR is available
            If ctx.HasPdfFiles Then
                If SharedMethods.IsOcrAvailable(_context) Then
                    Dim ocrAnswer = ShowCustomYesNoBox(
                        "Some files may require OCR (optical character recognition) to extract text. Enable OCR for PDF processing?" & vbCrLf & vbCrLf &
                        "Note: OCR may take longer but allows reading scanned documents and images.",
                        "Yes, enable OCR", "No, skip OCR")
                    ctx.EnableOCR = (ocrAnswer = 1)
                Else
                    ' OCR not available - will extract what text is possible
                    ctx.EnableOCR = False
                End If
            End If

            ' Load all files
            ShowAssistantThinking()

            Dim resultBuilder As New StringBuilder()
            Dim useDocumentTags = (filesToProcess.Count > 1)

            For Each filePath In filesToProcess
                Try
                    Dim result = Await LoadSingleKnowledgeFileAsync(filePath, ctx.EnableOCR, True)
                    Dim content = result.Content

                    ' Track PDFs that may have incomplete content
                    If result.PdfMayBeIncomplete Then
                        ctx.PdfsWithPossibleImages.Add(filePath)
                    End If

                    If String.IsNullOrWhiteSpace(content) Then
                        ctx.FailedFiles.Add(filePath)
                        Continue For
                    End If

                    ctx.GlobalDocumentCounter += 1
                    ctx.LoadedFiles.Add(Tuple.Create(filePath, content.Length))

                    If useDocumentTags Then
                        Dim docNum = ctx.GlobalDocumentCounter
                        Dim fileName = Path.GetFileName(filePath)
                        Dim openTag = $"<document{docNum} name=""{fileName}"">"
                        Dim closeTag = $"</document{docNum}>"
                        resultBuilder.Append(openTag).Append(content).Append(closeTag)
                    Else
                        resultBuilder.Append(content)
                    End If

                Catch ex As Exception
                    ctx.FailedFiles.Add(filePath)
                End Try
            Next

            RemoveAssistantThinking()

            ' Show summary
            Dim combinedContent = resultBuilder.ToString()

            If ctx.LoadedFiles.Count > 0 OrElse ctx.FailedFiles.Count > 0 OrElse ctx.IgnoredFilesPerDir.Count > 0 OrElse ctx.PdfsWithPossibleImages.Count > 0 Then
                Dim summary As New StringBuilder()
                summary.AppendLine("Knowledge loading summary:")
                summary.AppendLine("")

                If ctx.LoadedFiles.Count > 0 Then
                    summary.AppendLine($"Successfully loaded ({ctx.LoadedFiles.Count} files):")
                    Dim totalChars = 0
                    For Each item In ctx.LoadedFiles
                        summary.AppendLine($"  • {Path.GetFileName(item.Item1)} ({item.Item2:N0} chars)")
                        totalChars += item.Item2
                    Next
                    summary.AppendLine($"  Total: {totalChars:N0} characters")
                    summary.AppendLine("")
                End If

                If ctx.FailedFiles.Count > 0 Then
                    summary.AppendLine($"Failed to load ({ctx.FailedFiles.Count} items):")
                    For Each f In ctx.FailedFiles
                        summary.AppendLine($"  • {Path.GetFileName(f)}")
                    Next
                    summary.AppendLine("")
                End If

                If ctx.PdfsWithPossibleImages.Count > 0 Then
                    summary.AppendLine($"⚠ PDFs that may contain images/scans ({ctx.PdfsWithPossibleImages.Count} file(s)):")
                    For Each f In ctx.PdfsWithPossibleImages
                        summary.AppendLine($"  • {Path.GetFileName(f)}")
                    Next
                    summary.AppendLine("  (Text extraction may be incomplete - OCR was not available or not performed)")
                    summary.AppendLine("")
                End If

                If ctx.IgnoredFilesPerDir.Count > 0 Then
                    summary.AppendLine("Ignored unsupported files:")
                    For Each kvp In ctx.IgnoredFilesPerDir
                        summary.AppendLine($"  • {kvp.Key}: {kvp.Value} file(s)")
                    Next
                    summary.AppendLine("")
                End If

                Dim proceedAnswer = ShowCustomYesNoBox(
                    summary.ToString().TrimEnd() & vbCrLf & vbCrLf & "Do you want to use this knowledge?",
                    "Yes, proceed", "No, retry")

                If proceedAnswer <> 1 Then
                    ' User chose to retry
                    Await PromptForKnowledgeAsync()
                    Return
                End If
            End If

            If String.IsNullOrWhiteSpace(combinedContent) Then
                AppendSystemMessage("Failed to load knowledge or all files are empty.")
                Return
            End If

            ' Update state
            _knowledgeContent = combinedContent
            _knowledgeFilePath = If(isFile, selectedPath, selectedPath & " (directory)")

            ' Update runtime cache
            _cachedKnowledgeContent = _knowledgeContent
            _cachedKnowledgeFilePath = _knowledgeFilePath

            ' Persist if checkbox is checked
            If _chkPersistKnowledge.Checked Then
                Try
                    Dim persistPath = GetPersistedKnowledgeFilePath()
                    File.WriteAllText(persistPath, _knowledgeContent, Encoding.UTF8)
                    AppendSystemMessage($"Knowledge loaded and persisted ({_knowledgeContent.Length:N0} characters from {ctx.LoadedFiles.Count} file(s)).")
                Catch ex As Exception
                    AppendSystemMessage($"Knowledge loaded ({_knowledgeContent.Length:N0} characters) but failed to persist: {ex.Message}")
                End Try
            Else
                AppendSystemMessage($"Knowledge loaded: {ctx.LoadedFiles.Count} file(s), {_knowledgeContent.Length:N0} characters total.")
            End If

            UpdateWindowTitle()

            Try
                My.Settings.DiscussKnowledgePath = If(isFile, selectedPath, "")
                My.Settings.Save()
            Catch
            End Try

        Catch ex As Exception
            RemoveAssistantThinking()
            AppendSystemMessage($"Error loading knowledge: {ex.Message}")
        End Try
    End Function

    ''' <summary>
    ''' Loads a single knowledge file, optionally with OCR for PDFs.
    ''' </summary>
    ''' <param name="filePath">Path to the file to load.</param>
    ''' <param name="enableOCR">Whether to enable OCR for PDF files.</param>
    ''' <param name="silent">Whether to suppress error messages.</param>
    ''' <returns>Tuple of (content, pdfMayBeIncomplete) where pdfMayBeIncomplete is True if PDF heuristics suggest images but OCR was not performed.</returns>
    Private Async Function LoadSingleKnowledgeFileAsync(filePath As String, enableOCR As Boolean, silent As Boolean) As Task(Of (Content As String, PdfMayBeIncomplete As Boolean))
        If String.IsNullOrWhiteSpace(filePath) OrElse Not File.Exists(filePath) Then
            Return ("", False)
        End If

        Try
            Dim ext = Path.GetExtension(filePath).ToLowerInvariant()

            Select Case ext
                Case ".txt", ".md", ".log", ".ini", ".csv", ".json", ".xml", ".html", ".htm",
                     ".vb", ".cs", ".js", ".ts", ".py", ".java", ".cpp", ".c", ".h", ".sql", ".yaml", ".yml"
                    Return (File.ReadAllText(filePath, Encoding.UTF8), False)

                Case ".rtf"
                    Return (ReadRtfAsText(filePath), False)

                Case ".doc", ".docx"
                    Return (ReadWordDocument(filePath), False)

                Case ".pdf"
                    ' Use the extended version that reports if OCR was skipped
                    Dim pdfResult = Await ReadPdfAsTextEx(filePath, True, enableOCR, False, _context)
                    Return (pdfResult.Content, pdfResult.OcrWasSkippedDueToHeuristics)

                Case ".pptx"
                    Return (Globals.ThisAddIn.GetPresentationJson(filePath), False)

                Case Else
                    ' Try to read as text
                    Return (File.ReadAllText(filePath, Encoding.UTF8), False)
            End Select

        Catch ex As Exception
            If Not silent Then
                AppendSystemMessage($"Error loading {Path.GetFileName(filePath)}: {ex.Message}")
            End If
            Return ("", False)
        End Try
    End Function

#End Region

#Region "Chat Actions"

    ''' <summary>
    ''' Captures the user's message, adds it to history, and starts asynchronous LLM processing.
    ''' </summary>
    Private Sub OnSend(sender As Object, e As EventArgs)
        Dim userText = _txtInput.Text.Trim()
        If userText.Length = 0 Then Return

        AppendUserHtml(userText)
        _history.Add(("user", userText))
        _txtInput.Clear()
        ShowAssistantThinking()
        Dim __ = SendAsync(userText)
    End Sub

    ''' <summary>
    ''' Clears transcript and history, then regenerates the welcome sequence.
    ''' </summary>
    Private Async Sub OnClear(sender As Object, e As EventArgs)
        Try
            _history.Clear()
            InitializeChatHtml()
            My.Settings.DiscussLastChat = ""
            My.Settings.DiscussLastChatHtml = ""
            My.Settings.Save()
            Await SafeGenerateWelcomeAsync().ConfigureAwait(False)
        Catch
        Finally
            Ui(Sub() _txtInput.Focus())
        End Try
    End Sub

    ' Replace the OnSendToDoc method to properly handle autoresponder messages:

    ''' <summary>
    ''' Creates a new Word document with the chat transcript, excluding system messages.
    ''' Converts markdown to HTML for proper formatting.
    ''' </summary>
    Private Sub OnSendToDoc(sender As Object, e As EventArgs)
        Try
            If _history.Count = 0 Then
                AppendSystemMessage("No conversation to export.")
                Return
            End If

            Dim app = Globals.ThisAddIn.Application
            If app Is Nothing Then
                AppendSystemMessage("Word application is not available.")
                Return
            End If

            ' Create new document first
            Dim newDoc As Microsoft.Office.Interop.Word.Document = app.Documents.Add()
            Dim sel As Microsoft.Office.Interop.Word.Selection = app.Selection

            ' Build markdown content for the conversation
            Dim mdBuilder As New StringBuilder()

            ' Title
            mdBuilder.AppendLine($"# Discussion with {_currentPersonaName}")
            mdBuilder.AppendLine()

            ' Metadata
            mdBuilder.Append($"*Exported: {DateTime.Now:g}")
            If Not String.IsNullOrEmpty(_currentMissionName) Then
                mdBuilder.Append($" | Mission: {_currentMissionName}")
            End If
            If Not String.IsNullOrEmpty(_knowledgeFilePath) Then
                mdBuilder.Append($" | Knowledge: {Path.GetFileName(_knowledgeFilePath)}")
            End If
            mdBuilder.AppendLine("*")
            mdBuilder.AppendLine()
            mdBuilder.AppendLine("---")
            mdBuilder.AppendLine()

            ' Conversation
            For Each msg In _history
                Select Case msg.Role
                    Case "user"
                        mdBuilder.AppendLine("**You:**")
                        mdBuilder.AppendLine()
                        mdBuilder.AppendLine(msg.Content)
                        mdBuilder.AppendLine()

                    Case "assistant"
                        ' Check if content has an embedded display name (from Sort it Out mode)
                        Dim content = msg.Content
                        Dim colonIdx = content.IndexOf(": ", StringComparison.Ordinal)
                        Dim displayName = _currentPersonaName
                        Dim messageText = content

                        ' Check for Sort It Out style naming (e.g., "PersonaName (Advocate): message")
                        If colonIdx > 0 Then
                            Dim potentialName = content.Substring(0, colonIdx)
                            If potentialName.Contains("(Advocate)") OrElse potentialName.Contains("(Challenger)") OrElse potentialName.Contains("(2nd)") Then
                                displayName = potentialName
                                messageText = content.Substring(colonIdx + 2)
                            End If
                        End If

                        mdBuilder.AppendLine($"**{displayName}:**")
                        mdBuilder.AppendLine()
                        mdBuilder.AppendLine(messageText)
                        mdBuilder.AppendLine()

                    Case "autoresponder"
                        ' Autoresponder content is stored as "PersonaName: message"
                        ' We need to extract the persona name and format it properly
                        Dim content = msg.Content
                        Dim colonIdx = content.IndexOf(": ", StringComparison.Ordinal)
                        If colonIdx > 0 Then
                            Dim responderName = content.Substring(0, colonIdx)
                            Dim responderMessage = content.Substring(colonIdx + 2)
                            mdBuilder.AppendLine($"**{responderName}:**")
                            mdBuilder.AppendLine()
                            mdBuilder.AppendLine(responderMessage)
                            mdBuilder.AppendLine()
                        Else
                            ' Fallback: just output the content with a generic label
                            mdBuilder.AppendLine("**Autoresponder:**")
                            mdBuilder.AppendLine()
                            mdBuilder.AppendLine(content)
                            mdBuilder.AppendLine()
                        End If

                    Case Else
                        ' Skip system messages or unknown roles
                End Select
            Next

            ' Use the shared InsertTextWithMarkdown method which handles HTML/paste properly
            sel.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseStart)
            InsertTextWithMarkdown(sel, mdBuilder.ToString(), True)

            ' Move cursor to start
            newDoc.Content.Paragraphs(1).Range.Select()
            app.Selection.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseStart)

            AppendSystemMessage($"Chat exported to new document ({_history.Count} messages).")

        Catch ex As Exception
            AppendSystemMessage($"Error exporting to document: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Closes the DiscussInky form.
    ''' </summary>
    Private Sub OnClose(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    ''' <summary>
    ''' Handles Enter/Escape shortcuts for sending and closing.
    ''' </summary>
    Private Sub OnInputKeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter AndAlso Not e.Shift Then
            e.SuppressKeyPress = True
            OnSend(Me, EventArgs.Empty)
        ElseIf e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

#End Region

#Region "Welcome Message"

    ''' <summary>
    ''' Serializes welcome generation and surfaces any failures in the chat.
    ''' </summary>
    Private Async Function SafeGenerateWelcomeAsync() As Task
        If Interlocked.CompareExchange(_welcomeInProgress, 1, 0) <> 0 Then
            Return
        End If
        Try
            ' Show current session info before welcome
            ShowSessionInfo()
            Await GenerateWelcomeAsync()
        Catch ex As Exception
            RemoveAssistantThinking()
            AppendAssistantMarkdown("*(Welcome failed: " & System.Security.SecurityElement.Escape(ex.Message) & ")*")
        Finally
            Interlocked.Exchange(_welcomeInProgress, 0)
        End Try
    End Function

    ''' <summary>
    ''' Posts a system message summarizing the active persona, mission, and knowledge file.
    ''' </summary>
    Private Sub ShowSessionInfo()
        Dim sb As New StringBuilder()

        ' Persona info
        sb.Append($"Persona: {_currentPersonaName}")

        ' Mission info
        If Not String.IsNullOrEmpty(_currentMissionName) Then
            sb.Append($" | Mission: {_currentMissionName}")
        Else
            sb.Append(" | Mission: None")
        End If

        ' Knowledge document info
        If Not String.IsNullOrEmpty(_knowledgeFilePath) Then
            sb.Append($" | Knowledge: {Path.GetFileName(_knowledgeFilePath)}")
        Else
            sb.Append(" | Knowledge: None loaded")
        End If

        AppendSystemMessage(sb.ToString())
    End Sub

    ''' <summary>
    ''' Requests a short persona-aware welcome message from the LLM.
    ''' </summary>
    Private Async Function GenerateWelcomeAsync() As Task
        Dim langName = System.Globalization.CultureInfo.CurrentUICulture.DisplayName
        Dim partOfDay = GetPartOfDay()
        Dim dateContext = GetDateContext()
        Dim randomWord = GetRandomModifier()
        Dim locationContext = GetLocationContext()
        Dim languageInstruction = GetLanguageInstruction()

        Dim systemPrompt As String

        If String.IsNullOrWhiteSpace(_knowledgeContent) Then
            systemPrompt = $"{dateContext} Generate a brief, friendly {langName} welcome that {randomWord} references it is {partOfDay} now. " &
                           "Tell the user they should load a knowledge document using the 'Load Knowledge' button (button name always in English) to start a discussion. " &
                           $"You are ready to discuss any knowledge they provide. One short sentence, not talkative. {languageInstruction}"
        Else
            ' Use persona prompt to shape the welcome message
            Dim personaContext = ""
            If Not String.IsNullOrEmpty(_currentPersonaPrompt) Then
                personaContext = $" Your persona and role is defined as: '{_currentPersonaPrompt}'."
            End If

            ' Include mission context if active
            Dim missionContext = ""
            If Not String.IsNullOrEmpty(_currentMissionPrompt) Then
                missionContext = $" Your current mission is: '{_currentMissionPrompt}'."
            End If

            systemPrompt = $"{dateContext} {locationContext} Generate a brief, friendly {langName} welcome that {randomWord} references it is {partOfDay} now. " &
                           $"A knowledge base has been loaded (it may contain multiple documents or sections).{personaContext}{missionContext} " &
                           $"Generate a welcome that fits this persona and mission. One or two short sentences, stay in character. {languageInstruction}"
        End If

        Dim answer = ""
        Try
            Dim sw = Stopwatch.StartNew()
            answer = Await CallLlmWithSelectedModelAsync(systemPrompt, "")
            sw.Stop()
        Catch ex As Exception
            answer = $"Good {partOfDay.ToLower()}! How can I help you today?"
        End Try

        answer = If(answer, "").Trim()
        AppendAssistantMarkdown(answer)
        _history.Add(("assistant", answer))

        PersistChatHtml()
        PersistTranscriptLimited()
    End Function

#End Region

#Region "Send Message"

    ''' <summary>
    ''' Builds the full prompt (persona, mission, knowledge, history, document) and sends it to the LLM.
    ''' </summary>
    ''' <param name="userText">User's message text.</param>
    Private Async Function SendAsync(userText As String) As Task
        Try
            ' Build system prompt from persona or default
            Dim dateContext = GetDateContext()
            Dim randomWord = GetRandomModifier()
            Dim locationContext = GetLocationContext()
            Dim languageInstruction = GetLanguageInstruction()

            Dim basePrompt = If(Not String.IsNullOrEmpty(_currentPersonaPrompt),
                                _currentPersonaPrompt,
                                $"You are {_currentPersonaName}, a helpful assistant. Discuss the provided knowledge with the user.")

            ' Append mission if active
            Dim missionClause = ""
            If Not String.IsNullOrEmpty(_currentMissionPrompt) Then
                missionClause = $" Your mission: {_currentMissionPrompt}"
            End If

            Dim systemPrompt = $"{basePrompt}{missionClause}. In your response, be {randomWord}. Do not start with a greeting or salutation. " &
                               "The knowledge provided may consist of multiple documents or sections combined into one. " &
                               $"Refer to it as 'the knowledge' or 'the materials' rather than 'the document' when appropriate. {dateContext} {locationContext} {languageInstruction}"


            ' Build user prompt with knowledge and context
            Dim sb As New StringBuilder()

            sb.AppendLine("User message:")
            sb.AppendLine(userText)
            sb.AppendLine()

            ' Include full knowledge document without truncation for smaller docs
            If Not String.IsNullOrWhiteSpace(_knowledgeContent) Then
                sb.AppendLine("<Knowledge Base>")
                Dim knowledgeText = _knowledgeContent
                sb.AppendLine(knowledgeText)
                sb.AppendLine("</Knowledge Base>")
                sb.AppendLine()
            End If

            ' Include active document if checkbox checked
            If _chkIncludeActiveDoc.Checked Then
                Dim activeDocContent = GetActiveDocumentContent()
                If Not String.IsNullOrWhiteSpace(activeDocContent) Then
                    sb.AppendLine("<User's Active Document>")
                    sb.AppendLine(activeDocContent)
                    sb.AppendLine("</User's Active Document>")
                    sb.AppendLine()
                End If
            End If

            ' Include conversation history (supports user, assistant, and autoresponder roles)
            Dim convo = BuildConversationForAutoResponder()
            If Not String.IsNullOrWhiteSpace(convo) Then
                sb.AppendLine("Conversation so far:")
                sb.AppendLine(convo)
            End If

            Dim sw = Stopwatch.StartNew()
            Dim answer = Await CallLlmWithSelectedModelAsync(systemPrompt, sb.ToString())
            sw.Stop()

            answer = If(answer, "").Trim()

            RemoveAssistantThinking()
            AppendAssistantMarkdown(answer)
            _history.Add(("assistant", answer))

            PersistChatHtml()
            PersistTranscriptLimited()

        Catch ex As Exception
            RemoveAssistantThinking()
            AppendAssistantMarkdown("*(Error: " & System.Security.SecurityElement.Escape(ex.Message) & ")*")
        End Try
    End Function


#End Region

#Region "HTML Chat Display"

    ''' <summary>
    ''' Creates the base HTML document and CSS used by the WebBrowser control.
    ''' </summary>
    Private Sub InitializeChatHtml()
        Ui(Sub()
               _htmlQueue.Clear()
               _htmlReady = False
               Dim baseSize = If(Me.Font IsNot Nothing, Me.Font.SizeInPoints, 9.0F)
               Dim fontPt = Math.Max(CSng(baseSize + 1.0F), 10.0F)
               ' Replace the entire CSS variable in InitializeChatHtml with this:
               Dim css =
                   $"html,body{{height:100%;margin:0;padding:0;background:#fff;color:#000;}}
                    body{{font-family:'Segoe UI',Tahoma,Arial,sans-serif;font-size:{fontPt}pt;line-height:1.45;}}
                    #chat{{padding:8px;}}
                    .msg{{margin:8px 0;word-wrap:break-word;}}
                    .msg .who{{font-weight:600;margin-right:4px;}}
                    .msg.user{{background:#e8f4fc;border-left:3px solid #0078d4;padding:8px 10px;border-radius:4px;margin-right:40px;}}
                    .msg.user .who{{color:#0078d4;}}
                    .msg.assistant{{padding:8px 0;margin-left:0;}}
                    .msg.assistant .who{{color:#003366;}}
                    .msg.autoresponder{{background:#f3e8ff;border-left:3px solid #8b5cf6;padding:8px 10px;border-radius:4px;margin-right:40px;}}
                    .msg.autoresponder .who{{color:#6d28d9;}}
                    .msg.system{{color:#666;font-style:italic;background:#f9f9f9;padding:4px 8px;border-radius:4px;}}
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
    ''' Flushes queued HTML fragments once the browser document is ready.
    ''' </summary>
    Private Sub Chat_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs)
        _htmlReady = True
        If _htmlQueue.Count > 0 Then
            Try
                For Each frag In _htmlQueue
                    _chat.Document.InvokeScript("appendMessage", New Object() {frag})
                Next
            Catch
            Finally
                _htmlQueue.Clear()
            End Try
        End If
    End Sub

    ''' <summary>
    ''' Intercepts navigation to open http/https/mailto links externally.
    ''' </summary>
    Private Sub Chat_Navigating(sender As Object, e As WebBrowserNavigatingEventArgs)
        Try
            Dim scheme = e.Url?.Scheme?.ToLowerInvariant()
            If scheme = "http" OrElse scheme = "https" OrElse scheme = "mailto" Then
                e.Cancel = True
                Process.Start(New ProcessStartInfo(e.Url.ToString()) With {.UseShellExecute = True})
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Prevents the WebBrowser control from spawning new windows.
    ''' </summary>
    Private Sub Chat_NewWindow(sender As Object, e As CancelEventArgs)
        e.Cancel = True
    End Sub

    ''' <summary>
    ''' Appends HTML to the chat DOM, queuing if the document is not ready.
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
    ''' Adds a user message block to the transcript and persists HTML.
    ''' </summary>
    Private Sub AppendUserHtml(text As String)
        Dim encoded = WebUtility.HtmlEncode(text).Replace(vbCrLf, "<br>").Replace(vbLf, "<br>").Replace(vbCr, "<br>")
        AppendHtml($"<div class='msg user'><span class='who'>You:</span><span class='content'>{encoded}</span></div>")
        PersistChatHtml()
    End Sub

    ''' <summary>
    ''' Adds a system message block and persists HTML.
    ''' </summary>
    Private Sub AppendSystemMessage(text As String)
        Dim encoded = WebUtility.HtmlEncode(text)
        AppendHtml($"<div class='msg system'>{encoded}</div>")
        PersistChatHtml()
    End Sub

    ''' <summary>
    ''' Inserts a temporary 'thinking' placeholder for the assistant.
    ''' </summary>
    Private Sub ShowAssistantThinking()
        _lastThinkingId = "thinking-" & Guid.NewGuid().ToString("N")
        AppendHtml($"<div id=""{_lastThinkingId}"" class='msg assistant thinking'><span class='who'>{WebUtility.HtmlEncode(_currentPersonaName)}:</span><span class='content'>Thinking...</span></div>")
    End Sub

    ''' <summary>
    ''' Removes the current thinking placeholder if present.
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
    ''' Converts assistant markdown to HTML and appends it to the transcript.
    ''' </summary>
    Private Sub AppendAssistantMarkdown(md As String)
        AppendAssistantMarkdownWithName(md, _currentPersonaName)
    End Sub


    ''' <summary>
    ''' Converts assistant markdown to HTML and appends it to the transcript with a custom display name.
    ''' </summary>
    Private Sub AppendAssistantMarkdownWithName(md As String, displayName As String)
        md = If(md, "")
        Dim body = Markdig.Markdown.ToHtml(md, _mdPipeline)
        Dim t = body.Trim()
        Dim isSingle = Regex.IsMatch(t, "^\s*<p>[\s\S]*?</p>\s*$", RegexOptions.IgnoreCase) AndAlso
                   Not Regex.IsMatch(t, "<(ul|ol|pre|table|h[1-6]|blockquote|hr|div)\b", RegexOptions.IgnoreCase)

        Dim whoHtml = WebUtility.HtmlEncode(displayName)

        If isSingle Then
            Dim inlineHtml = Regex.Replace(t, "^\s*<p>|</p>\s*$", "", RegexOptions.IgnoreCase)
            AppendHtml($"<div class='msg assistant'><span class='who'>{whoHtml}:</span><span class='content'>{inlineHtml}</span></div>")
        Else
            Dim m = Regex.Match(t, "^\s*<p>([\s\S]*?)</p>\s*", RegexOptions.IgnoreCase)
            If m.Success Then
                Dim firstInline = m.Groups(1).Value
                Dim rest = t.Substring(m.Index + m.Length).Trim()
                Dim sb As New StringBuilder()
                sb.Append("<div class='msg assistant'>")
                sb.Append("<span class='who'>").Append(whoHtml).Append(":</span>")
                sb.Append("<span class='content'>").Append(firstInline).Append("</span>")
                If rest.Length > 0 Then
                    sb.Append("<div class='content'>").Append(rest).Append("</div>")
                End If
                sb.Append("</div>")
                AppendHtml(sb.ToString())
            Else
                AppendHtml($"<div class='msg assistant'><span class='who'>{whoHtml}:</span><div class='content'>{t}</div></div>")
            End If
        End If
    End Sub

#End Region

#Region "Persistence"

    ''' <summary>
    ''' Saves the current chat DOM fragment to settings for restoration.
    ''' </summary>
    Private Sub PersistChatHtml()
        Ui(Sub()
               Try
                   If _chat.Document Is Nothing Then Return
                   Dim root = _chat.Document.GetElementById("chat")
                   If root Is Nothing Then Return
                   My.Settings.DiscussLastChatHtml = root.InnerHtml
                   My.Settings.Save()
               Catch
               End Try
           End Sub)
    End Sub

    ''' <summary>
    ''' Rebuilds the history list from the plain-text transcript copy.
    ''' </summary>
    Private Sub RestoreHistoryFromTranscript(transcript As String)
        _history.Clear()
        If String.IsNullOrEmpty(transcript) Then Return

        Dim lines = transcript.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split({vbLf}, StringSplitOptions.None)
        Dim currentRole As String = Nothing
        Dim content As New StringBuilder()

        Dim flush =
        Sub()
            If content.Length = 0 OrElse String.IsNullOrEmpty(currentRole) Then
                content.Clear() : currentRole = Nothing : Return
            End If
            _history.Add((currentRole, content.ToString().Trim()))
            content.Clear()
            currentRole = Nothing
        End Sub

        For Each ln In lines
            ' Check for user message marker
            If ln.StartsWith("You: ", StringComparison.OrdinalIgnoreCase) Then
                flush()
                currentRole = "user"
                content.Append(ln.Substring(5).TrimStart())
            ElseIf ln.StartsWith(_currentPersonaName & ": ", StringComparison.OrdinalIgnoreCase) Then
                flush()
                currentRole = "assistant"
                content.Append(ln.Substring((_currentPersonaName & ": ").Length).TrimStart())
            ElseIf ln.StartsWith(AssistantName & ": ", StringComparison.OrdinalIgnoreCase) Then
                flush()
                currentRole = "assistant"
                content.Append(ln.Substring((AssistantName & ": ").Length).TrimStart())
            Else
                ' Continuation line - only append if we're already in a message
                If currentRole IsNot Nothing Then
                    If content.Length > 0 Then content.AppendLine()
                    content.Append(ln)
                End If
            End If
        Next
        flush()
    End Sub

    ''' <summary>
    ''' Recreates chat HTML from the stored transcript text.
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
            ElseIf ln.StartsWith(_currentPersonaName & ":", StringComparison.OrdinalIgnoreCase) Then
                flush() : currentRole = "assistant" : content.Append(ln.Substring((_currentPersonaName & ":").Length).TrimStart())
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
    ''' Truncates and saves the plain transcript respecting the configured cap.
    ''' </summary>
    Private Sub PersistTranscriptLimited()
        Dim transcript = BuildTranscriptPlain()
        Dim cap = Math.Max(5000, If(_context IsNot Nothing, _context.INI_ChatCap, 0))
        If transcript.Length > cap Then
            transcript = transcript.Substring(transcript.Length - cap)
        End If
        My.Settings.DiscussLastChat = transcript
    End Sub

    ''' <summary>
    ''' Returns the current chat history in 'You:/Persona:' text format.
    ''' </summary>
    Private Function BuildTranscriptPlain() As String
        Dim sb As New StringBuilder()
        For Each m In _history
            If m.Role = "user" Then
                sb.AppendLine("You: " & m.Content)
            Else
                sb.AppendLine(_currentPersonaName & ": " & m.Content)
            End If
        Next
        Return sb.ToString()
    End Function


#End Region


#Region "Autorespond Feature"

    ''' <summary>
    ''' Handles the Autorespond button click - shows configuration dialog and starts auto-response loop.
    ''' </summary>
    Private Async Sub OnAutoRespondClick(sender As Object, e As EventArgs)
        ' Prevent running if Sort It Out is in progress
        If _sortOutInProgress Then
            AppendSystemMessage("Cannot start Autorespond while Sort It Out is in progress.")
            Return
        End If
        ' Only allow when input is enabled (not during processing)
        If Not _txtInput.Enabled Then
            AppendSystemMessage("Cannot start autorespond while a response is in progress.")
            Return
        End If

        ' Show configuration dialog
        If Not ShowAutoRespondConfigDialog() Then
            Return ' User cancelled
        End If

        ' Start the autorespond loop
        Await RunAutoRespondLoopAsync()
    End Sub

    ' Replace the ShowAutoRespondConfigDialog function with this version that handles mission editing properly:

    ''' <summary>
    ''' Shows the autorespond configuration dialog using ShowCustomVariableInputForm.
    ''' </summary>
    ''' <returns>True if user confirmed, False if cancelled.</returns>
    Private Function ShowAutoRespondConfigDialog() As Boolean
        ' Build persona options
        Dim personaOptions As New List(Of String)()
        For Each p In _personas
            personaOptions.Add(p.DisplayName)
        Next
        If personaOptions.Count = 0 Then
            ShowCustomMessageBox("No personas configured. Please configure personas first.")
            Return False
        End If

        ' Build mission options (including "No mission")
        Dim missionOptions As New List(Of String)()
        missionOptions.Add("No mission")
        For Each m In _missions
            missionOptions.Add(m.DisplayName)
        Next

        ' Build round count options (1 to MaxAutoRespondRounds)
        Dim roundOptions As New List(Of String)()
        For i = 1 To MaxAutoRespondRounds
            roundOptions.Add(i.ToString())
        Next

        ' Restore persisted values or use defaults
        Dim savedPersona = ""
        Dim savedMission = ""
        Dim savedRounds = DefaultRespondRounds
        Dim savedBreakOff = DefaultAutoRespondBreakOff
        Try
            savedPersona = My.Settings.AutoRespondPersona
            savedMission = My.Settings.AutoRespondMission
            savedRounds = My.Settings.AutoRespondMaxRounds
            If savedRounds < 1 OrElse savedRounds > MaxAutoRespondRounds Then savedRounds = 5
            savedBreakOff = My.Settings.AutoRespondBreakOff
            If String.IsNullOrWhiteSpace(savedBreakOff) Then savedBreakOff = DefaultAutoRespondBreakOff
        Catch
        End Try

        ' Find default values
        Dim defaultPersonaDisplay = If(personaOptions.Count > 0, personaOptions(0), "")
        For i = 0 To _personas.Count - 1
            If _personas(i).DisplayName.Equals(savedPersona, StringComparison.OrdinalIgnoreCase) OrElse
               _personas(i).Name.Equals(savedPersona, StringComparison.OrdinalIgnoreCase) Then
                defaultPersonaDisplay = _personas(i).DisplayName
                Exit For
            End If
        Next

        Dim defaultMissionDisplay = "No mission"
        If Not String.IsNullOrEmpty(savedMission) Then
            For i = 0 To _missions.Count - 1
                If _missions(i).DisplayName.Equals(savedMission, StringComparison.OrdinalIgnoreCase) OrElse
                   _missions(i).Name.Equals(savedMission, StringComparison.OrdinalIgnoreCase) Then
                    defaultMissionDisplay = _missions(i).DisplayName
                    Exit For
                End If
            Next
        End If

        ' Build InputParameter array
        Dim p0 As New SharedMethods.InputParameter("Responder Persona", defaultPersonaDisplay, personaOptions)
        Dim p1 As New SharedMethods.InputParameter("Responder Mission", defaultMissionDisplay, missionOptions)
        Dim p2 As New SharedMethods.InputParameter("Maximum Rounds", savedRounds.ToString(), roundOptions)
        Dim p3 As New SharedMethods.InputParameter("Break-off Instruction", savedBreakOff)

        Dim params() As SharedMethods.InputParameter = {p0, p1, p2, p3}

        ' Prepare extra button for editing mission file
        Dim missionPath = GetMissionFilePath()
        Dim extraButtonText = If(Not String.IsNullOrWhiteSpace(missionPath), "Edit Missions...", Nothing)
        Dim extraButtonAction As System.Action = Nothing
        Dim shouldReopenDialog As Boolean = False

        If Not String.IsNullOrWhiteSpace(missionPath) Then
            extraButtonAction = Sub()
                                    EnsureMissionFileExists(missionPath)
                                    ShowTextFileEditor(missionPath, $"{AN} - Edit Missions:", False, _context)
                                    ' Reload missions after editing
                                    LoadMissions()
                                    ' Flag to reopen the dialog
                                    shouldReopenDialog = True
                                End Sub
        End If

        ' Show the dialog - CloseAfterExtra = True to close after Edit Missions button
        Dim result = ShowCustomVariableInputForm(
            "Configure the AI responder that will continue the conversation:",
            $"{AN} - Configure Autorespond",
            params,
            extraButtonText,
            extraButtonAction,
            CloseAfterExtra:=True)

        ' If user clicked Edit Missions, reopen the dialog with updated missions
        If shouldReopenDialog Then
            Return ShowAutoRespondConfigDialog()
        End If

        If Not result Then
            Return False ' User cancelled
        End If

        ' Parse results
        Dim selectedPersonaDisplay = CStr(params(0).Value)
        Dim selectedMissionDisplay = CStr(params(1).Value)
        Dim selectedRounds = 5
        Integer.TryParse(CStr(params(2).Value), selectedRounds)
        Dim breakOffText = CStr(params(3).Value)

        ' Find the selected persona
        Dim foundPersona = _personas.FirstOrDefault(Function(p) p.DisplayName.Equals(selectedPersonaDisplay, StringComparison.OrdinalIgnoreCase))
        If String.IsNullOrEmpty(foundPersona.Name) Then
            ' Fallback to first persona
            foundPersona = _personas(0)
        End If
        _autoRespondPersonaName = foundPersona.Name
        _autoRespondPersonaPrompt = foundPersona.Prompt

        ' Find the selected mission
        If selectedMissionDisplay.Equals("No mission", StringComparison.OrdinalIgnoreCase) Then
            _autoRespondMissionName = ""
            _autoRespondMissionPrompt = ""
        Else
            Dim foundMission = _missions.FirstOrDefault(Function(m) m.DisplayName.Equals(selectedMissionDisplay, StringComparison.OrdinalIgnoreCase))
            If Not String.IsNullOrEmpty(foundMission.Name) Then
                _autoRespondMissionName = foundMission.Name
                _autoRespondMissionPrompt = foundMission.Prompt
            Else
                _autoRespondMissionName = ""
                _autoRespondMissionPrompt = ""
            End If
        End If

        _autoRespondMaxRounds = If(selectedRounds >= 1 AndAlso selectedRounds <= MaxAutoRespondRounds, selectedRounds, DefaultRespondRounds)
        _autoRespondBreakOff = If(String.IsNullOrWhiteSpace(breakOffText), DefaultAutoRespondBreakOff, breakOffText)

        ' Persist settings
        Try
            My.Settings.AutoRespondPersona = _autoRespondPersonaName
            My.Settings.AutoRespondMission = _autoRespondMissionName
            My.Settings.AutoRespondMaxRounds = _autoRespondMaxRounds
            My.Settings.AutoRespondBreakOff = _autoRespondBreakOff
            My.Settings.Save()
        Catch
        End Try

        Return True
    End Function


    ''' <summary>
    ''' Runs the autorespond loop, alternating between the responder and the chatbot.
    ''' </summary>
    Private Async Function RunAutoRespondLoopAsync() As Task
        _autoRespondInProgress = True
        _autoRespondCancelled = False

        ' Disable input during autorespond
        Ui(Sub()
               _txtInput.Enabled = False
               _btnSend.Enabled = False
               _btnAutoRespond.Enabled = False
           End Sub)

        ' Determine display name for responder (add "(2nd)" if same as chatbot persona)
        Dim responderDisplayName = _autoRespondPersonaName
        If _autoRespondPersonaName.Equals(_currentPersonaName, StringComparison.OrdinalIgnoreCase) Then
            responderDisplayName = _autoRespondPersonaName & " (2nd)"
        End If

        ' Show progress bar if more than 1 round
        Dim useProgressBar = (_autoRespondMaxRounds > 1)
        If useProgressBar Then
            ShowProgressBarInSeparateThread($"{AN} Autorespond", $"{responderDisplayName} responding...")
            ProgressBarModule.CancelOperation = False
            ProgressBarModule.GlobalProgressMax = _autoRespondMaxRounds
            ProgressBarModule.GlobalProgressValue = 0
            ProgressBarModule.GlobalProgressLabel = "Starting..."
        End If

        ' Notify start
        AppendSystemMessage($"Autorespond started: {responderDisplayName}" &
                           If(Not String.IsNullOrEmpty(_autoRespondMissionName), $" [{_autoRespondMissionName}]", "") &
                           $" for up to {_autoRespondMaxRounds} round(s).")

        Try
            Dim roundCount = 0
            Dim stopRequested = False

            While roundCount < _autoRespondMaxRounds AndAlso Not _autoRespondCancelled AndAlso Not stopRequested
                roundCount += 1

                If useProgressBar Then
                    ProgressBarModule.GlobalProgressValue = roundCount
                    ProgressBarModule.GlobalProgressLabel = $"Round {roundCount} of {_autoRespondMaxRounds}..."
                    If ProgressBarModule.CancelOperation Then
                        _autoRespondCancelled = True
                        Exit While
                    End If
                End If

                ' Step 1: Get response from the autoresponder (simulating user input)
                ShowAutoResponderThinking(responderDisplayName)
                Dim responderMessage = Await GenerateAutoResponderMessageAsync(responderDisplayName)
                RemoveAssistantThinking()

                ' Check for stop word
                If responderMessage.Contains(AutoRespondStopWord) Then
                    stopRequested = True
                    responderMessage = responderMessage.Replace(AutoRespondStopWord, "").Trim()
                End If

                ' Display and record the responder's message
                If Not String.IsNullOrWhiteSpace(responderMessage) Then
                    AppendAutoResponderHtml(responderDisplayName, responderMessage)
                    _history.Add(("autoresponder", $"{responderDisplayName}: {responderMessage}"))
                End If

                If stopRequested OrElse _autoRespondCancelled Then
                    Exit While
                End If

                ' Step 2: Get response from the chatbot
                ShowAssistantThinking()
                Dim chatbotResponse = Await GenerateChatbotResponseToAutoResponderAsync(responderDisplayName)
                RemoveAssistantThinking()

                ' Check if chatbot also wants to stop (unlikely but possible)
                If chatbotResponse.Contains(AutoRespondStopWord) Then
                    stopRequested = True
                    chatbotResponse = chatbotResponse.Replace(AutoRespondStopWord, "").Trim()
                End If

                ' Display and record the chatbot's response
                If Not String.IsNullOrWhiteSpace(chatbotResponse) Then
                    AppendAssistantMarkdown(chatbotResponse)
                    _history.Add(("assistant", chatbotResponse))
                End If

                PersistChatHtml()
                PersistTranscriptLimited()

                ' Small delay to prevent rate limiting and allow UI updates
                Await Task.Delay(500)
            End While

            ' Summary message
            If _autoRespondCancelled Then
                AppendSystemMessage($"Autorespond cancelled after {roundCount} round(s).")
            ElseIf stopRequested Then
                AppendSystemMessage($"Autorespond completed after {roundCount} round(s) - responder indicated conversation should stop.")
            Else
                AppendSystemMessage($"Autorespond completed - maximum of {roundCount} round(s) reached.")
            End If

            ' Offer summary if enough rounds completed
            Await ShowDiscussionSummaryAsync(roundCount)

        Catch ex As Exception
            AppendSystemMessage($"Autorespond error: {ex.Message}")
        Finally
            If useProgressBar Then
                ProgressBarModule.CancelOperation = True
            End If

            _autoRespondInProgress = False
            _autoRespondCancelled = False

            ' Re-enable input
            Ui(Sub()
                   _txtInput.Enabled = True
                   _btnSend.Enabled = True
                   _btnAutoRespond.Enabled = True
                   _txtInput.Focus()
               End Sub)

            PersistChatHtml()
            PersistTranscriptLimited()
        End Try
    End Function

    ''' <summary>
    ''' Generates a message from the autoresponder persona.
    ''' </summary>
    Private Async Function GenerateAutoResponderMessageAsync(responderDisplayName As String) As Task(Of String)
        Dim dateContext = GetDateContext()
        Dim randomWord = GetRandomModifier()
        Dim locationContext = GetLocationContext()
        Dim languageInstruction = GetLanguageInstruction()

        ' Build system prompt for the responder
        Dim basePrompt = If(Not String.IsNullOrEmpty(_autoRespondPersonaPrompt),
                            _autoRespondPersonaPrompt,
                            $"You are {_autoRespondPersonaName}, participating in a discussion.")

        Dim missionClause = ""
        If Not String.IsNullOrEmpty(_autoRespondMissionPrompt) Then
            missionClause = $" Your mission: {_autoRespondMissionPrompt}"
        End If

        Dim systemPrompt = $"{basePrompt}{missionClause}. In your response, be {randomWord}. Do not start with a greeting or salutation. " &
                           $"You are responding to {_currentPersonaName} in an ongoing discussion. " &
                           $"{_autoRespondBreakOff} {dateContext} {locationContext} {languageInstruction}"


        ' Build the conversation context
        Dim sb As New StringBuilder()
        sb.AppendLine($"You are {responderDisplayName}, responding to {_currentPersonaName}.")
        sb.AppendLine()

        ' Include knowledge if available
        If Not String.IsNullOrWhiteSpace(_knowledgeContent) Then
            sb.AppendLine("<Knowledge Base>")
            sb.AppendLine(_knowledgeContent)
            sb.AppendLine("</Knowledge Base>")
            sb.AppendLine()
        End If

        ' Include active document if checkbox checked (same as main chatbot)
        If _chkIncludeActiveDoc.Checked Then
            Dim activeDocContent = GetActiveDocumentContent()
            If Not String.IsNullOrWhiteSpace(activeDocContent) Then
                sb.AppendLine("<User's Active Document>")
                sb.AppendLine(activeDocContent)
                sb.AppendLine("</User's Active Document>")
                sb.AppendLine()
            End If
        End If

        ' Include conversation history with clear role identification
        sb.AppendLine("Conversation so far:")
        Dim convo = BuildConversationForAutoResponder()
        sb.AppendLine(convo)
        sb.AppendLine()
        sb.AppendLine($"Now respond as {responderDisplayName}:")

        Dim answer = Await CallLlmWithSelectedModelAsync(systemPrompt, sb.ToString())
        Return If(answer, "").Trim()
    End Function

    ''' <summary>
    ''' Generates the chatbot's response to the autoresponder's message.
    ''' </summary>
    Private Async Function GenerateChatbotResponseToAutoResponderAsync(responderDisplayName As String) As Task(Of String)
        Dim dateContext = GetDateContext()
        Dim randomWord = GetRandomModifier()
        Dim locationContext = GetLocationContext()
        Dim languageInstruction = GetLanguageInstruction()

        ' Use the main chatbot's persona and mission
        Dim basePrompt = If(Not String.IsNullOrEmpty(_currentPersonaPrompt),
                            _currentPersonaPrompt,
                            $"You are {_currentPersonaName}, a helpful assistant.")

        Dim missionClause = ""
        If Not String.IsNullOrEmpty(_currentMissionPrompt) Then
            missionClause = $" Your mission: {_currentMissionPrompt}"
        End If

        Dim systemPrompt = $"{basePrompt}{missionClause}. In your response, be {randomWord}. Do not start with a greeting or salutation. " &
                           $"You are discussing with {responderDisplayName}. The knowledge provided may consist of multiple documents or sections. " &
                           $"{dateContext} {locationContext} {languageInstruction}"


        ' Build the conversation context
        Dim sb As New StringBuilder()
        sb.AppendLine($"You are {_currentPersonaName}, discussing with {responderDisplayName}.")
        sb.AppendLine()

        ' Include knowledge if available
        If Not String.IsNullOrWhiteSpace(_knowledgeContent) Then
            sb.AppendLine("<Knowledge Base>")
            sb.AppendLine(_knowledgeContent)
            sb.AppendLine("</Knowledge Base>")
            sb.AppendLine()
        End If

        ' Include active document if checkbox checked
        If _chkIncludeActiveDoc.Checked Then
            Dim activeDocContent = GetActiveDocumentContent()
            If Not String.IsNullOrWhiteSpace(activeDocContent) Then
                sb.AppendLine("<User's Active Document>")
                sb.AppendLine(activeDocContent)
                sb.AppendLine("</User's Active Document>")
                sb.AppendLine()
            End If
        End If

        ' Include conversation history
        sb.AppendLine("Conversation so far:")
        Dim convo = BuildConversationForAutoResponder()
        sb.AppendLine(convo)
        sb.AppendLine()
        sb.AppendLine($"Now respond as {_currentPersonaName}:")

        Dim answer = Await CallLlmWithSelectedModelAsync(systemPrompt, sb.ToString())
        Return If(answer, "").Trim()
    End Function

    ''' <summary>
    ''' Builds conversation history with proper role identification for autorespond context.
    ''' </summary>
    Private Function BuildConversationForAutoResponder() As String
        Dim sb As New StringBuilder()
        Dim cap = Math.Max(5000, If(_context IsNot Nothing, _context.INI_ChatCap, 0))
        Dim acc = 0

        For i = _history.Count - 1 To 0 Step -1
            Dim role = _history(i).Role
            Dim content = _history(i).Content
            Dim line As String

            Select Case role
                Case "user"
                    line = "User: " & content & Environment.NewLine
                Case "assistant"
                    ' Check if content already has an embedded display name (from Sort It Out mode)
                    If content.Contains("(Advocate):") OrElse content.Contains("(Challenger):") OrElse content.Contains("(2nd):") Then
                        ' Content already includes the persona name prefix
                        line = content & Environment.NewLine
                    Else
                        line = _currentPersonaName & ": " & content & Environment.NewLine
                    End If
                Case "autoresponder"
                    ' Content already includes the persona name prefix
                    line = content & Environment.NewLine
                Case Else
                    line = content & Environment.NewLine
            End Select

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
    ''' Shows a thinking placeholder for the autoresponder.
    ''' </summary>
    Private Sub ShowAutoResponderThinking(responderName As String)
        _lastThinkingId = "thinking-" & Guid.NewGuid().ToString("N")
        AppendHtml($"<div id=""{_lastThinkingId}"" class='msg autoresponder thinking'><span class='who'>{WebUtility.HtmlEncode(responderName)}:</span><span class='content'>Thinking...</span></div>")
    End Sub

    ''' <summary>
    ''' Appends an autoresponder message with distinct styling.
    ''' </summary>
    Private Sub AppendAutoResponderHtml(responderName As String, text As String)
        Dim body = Markdig.Markdown.ToHtml(text, _mdPipeline)
        Dim t = body.Trim()
        Dim whoHtml = WebUtility.HtmlEncode(responderName)

        Dim isSingle = Regex.IsMatch(t, "^\s*<p>[\s\S]*?</p>\s*$", RegexOptions.IgnoreCase) AndAlso
                       Not Regex.IsMatch(t, "<(ul|ol|pre|table|h[1-6]|blockquote|hr|div)\b", RegexOptions.IgnoreCase)

        If isSingle Then
            Dim inlineHtml = Regex.Replace(t, "^\s*<p>|</p>\s*$", "", RegexOptions.IgnoreCase)
            AppendHtml($"<div class='msg autoresponder'><span class='who'>{whoHtml}:</span><span class='content'>{inlineHtml}</span></div>")
        Else
            Dim m = Regex.Match(t, "^\s*<p>([\s\S]*?)</p>\s*", RegexOptions.IgnoreCase)
            If m.Success Then
                Dim firstInline = m.Groups(1).Value
                Dim rest = t.Substring(m.Index + m.Length).Trim()
                Dim htmlSb As New StringBuilder()
                htmlSb.Append("<div class='msg autoresponder'>")
                htmlSb.Append("<span class='who'>").Append(whoHtml).Append(":</span>")
                htmlSb.Append("<span class='content'>").Append(firstInline).Append("</span>")
                If rest.Length > 0 Then
                    htmlSb.Append("<div class='content'>").Append(rest).Append("</div>")
                End If
                htmlSb.Append("</div>")
                AppendHtml(htmlSb.ToString())
            Else
                AppendHtml($"<div class='msg autoresponder'><span class='who'>{whoHtml}:</span><div class='content'>{t}</div></div>")
            End If
        End If
    End Sub

#End Region

#Region "Sort It Out Feature"

    ''' <summary>
    ''' Handles the Sort It Out button click - prompts for instruction and starts a structured discussion.
    ''' </summary>
    Private Async Sub OnSortOutClick(sender As Object, e As EventArgs)
        ' Prevent running if autorespond or Sort It Out is already in progress
        If _autoRespondInProgress Then
            AppendSystemMessage("Cannot start Sort It Out while Autorespond is in progress.")
            Return
        End If
        If _sortOutInProgress Then
            AppendSystemMessage("Sort It Out is already in progress.")
            Return
        End If
        If Not _txtInput.Enabled Then
            AppendSystemMessage("Cannot start Sort It Out while a response is in progress.")
            Return
        End If

        ' Run the Sort It Out flow
        Await RunSortOutFlowAsync()
    End Sub

    ''' <summary>
    ''' Main flow for the Sort It Out feature.
    ''' </summary>
    Private Async Function RunSortOutFlowAsync() As Task
        ' Step 1: Get user instruction
        Dim userInstruction = ShowCustomInputBox(
            "Enter your instruction for the discussion. The two bots will sort out this issue based on the conversation so far and the loaded knowledge." & vbCrLf & vbCrLf &
            "Example: ""In the discussion so far, I received the advice to cancel the contract. Now, please discuss whether this really makes sense.""" & vbCrLf,
            $"{AN} - Sort It Out Discussion", False)

        If String.IsNullOrWhiteSpace(userInstruction) Or userInstruction = "ESC" Then
            Return ' User cancelled
        End If

        userInstruction = userInstruction.Trim()

        ' Step 2: Get maximum rounds
        Dim maxRounds = ShowSortOutRoundsDialog()
        If maxRounds < 1 Then
            Return ' User cancelled
        End If

        ' Variables to hold the mission prompts
        Dim mainMission As String = ""
        Dim responderMission As String = ""
        Dim missionsGenerated As Boolean = False

        ' Check if we have stored missions from a previous Sort Out
        Dim hasStoredMissions = False
        Try
            Dim storedMain = My.Settings.SortOutMainMission
            Dim storedResponder = My.Settings.SortOutResponderMission
            hasStoredMissions = Not String.IsNullOrWhiteSpace(storedMain) AndAlso Not String.IsNullOrWhiteSpace(storedResponder)
        Catch
        End Try

        Dim UserSelectMission As Boolean = False

        If hasStoredMissions Then
            ' Ask user if they want to reuse stored missions
            Dim reuseAnswer = ShowCustomYesNoBox(
                "Previously generated mission statements are available. Do you want to reuse them?" & vbCrLf & vbCrLf &
                "Click 'Yes' to reuse the previous missions, or 'No' to generate new ones.",
                "Yes, reuse", "No, generate new")

            If reuseAnswer = 1 Then
                Try
                    mainMission = My.Settings.SortOutMainMission
                    responderMission = My.Settings.SortOutResponderMission
                    missionsGenerated = True
                    AppendSystemMessage("Reusing previously generated mission statements.")
                Catch
                End Try
            End If
            If reuseAnswer = 0 Then
                Dim abort = ShowCustomYesNoBox(
                    "Do you really want to abort, or do you want to select the missions manually?",
                    "Yes, abort", "No, select manually")
                If abort <> 2 Then Return
                UserSelectMission = True
            End If
        End If

        ' Step 3: Generate or select missions
        If Not missionsGenerated AndAlso Not String.IsNullOrWhiteSpace(userInstruction) AndAlso Not UserSelectMission Then
            ' Try to generate missions using LLM
            Dim generatedMissions = Await GenerateSortOutMissionsAsync(userInstruction, maxRounds)

            If generatedMissions.Success Then
                ' Always persist the generated missions
                mainMission = generatedMissions.MainMission
                responderMission = generatedMissions.ResponderMission

                Try
                    My.Settings.SortOutMainMission = mainMission
                    My.Settings.SortOutResponderMission = responderMission
                    My.Settings.Save()
                Catch
                End Try

                If ShowGeneratedMissionsConfirmation Then
                    ' Show the generated missions and ask for confirmation
                    Dim confirmMsg = $"Generated mission statements:" & vbCrLf & vbCrLf &
                                     $"**Advocate (Main Bot):**" & vbCrLf &
                                     $"{generatedMissions.MainMission}" & vbCrLf & vbCrLf &
                                     $"**Challenger (Responder Bot):**" & vbCrLf &
                                     $"{generatedMissions.ResponderMission}" & vbCrLf & vbCrLf &
                                     "Proceed with these missions?"

                    Dim confirmAnswer = ShowCustomYesNoBox(confirmMsg, "Yes, proceed", "No, select manually")

                    If confirmAnswer = 1 Then
                        missionsGenerated = True
                    End If
                Else
                    ' Skip confirmation, use generated missions directly
                    missionsGenerated = True
                    AppendSystemMessage("Mission statements generated for Advocate and Challenger.")
                End If
            Else
                ' LLM failed - notify user
                ShowCustomMessageBox("Could not automatically generate mission statements. Please select missions manually.")
            End If
        End If

        ' Step 4: If missions not generated, let user select manually
        If Not missionsGenerated Then
            ' Select mission for main bot
            mainMission = ShowSortOutMissionSelector("Select mission for the Main Bot (Advocate):", "SortOutMainMissionManual")
            If mainMission Is Nothing Then
                Return ' User cancelled
            End If

            ' Select mission for responder bot
            responderMission = ShowSortOutMissionSelector("Select mission for the Responder Bot (Challenger):", "SortOutResponderMissionManual")
            If responderMission Is Nothing Then
                Return ' User cancelled
            End If
        End If

        ' Step 5: Store original mission and set up temporary missions
        _sortOutOriginalMissionName = _currentMissionName
        _sortOutOriginalMissionPrompt = _currentMissionPrompt

        ' Temporarily set the main bot's mission
        _currentMissionPrompt = mainMission

        ' Set up autoresponder with same persona but different mission
        _autoRespondPersonaName = _currentPersonaName
        _autoRespondPersonaPrompt = _currentPersonaPrompt
        _autoRespondMissionPrompt = responderMission
        _autoRespondMaxRounds = maxRounds
        _autoRespondBreakOff = DefaultAutoRespondBreakOff

        ' Store the sort out mission prompts for reference
        _sortOutMainMissionPrompt = mainMission
        _sortOutResponderMissionPrompt = responderMission

        ' Step 6: Inject user instruction as a user message if provided
        If Not String.IsNullOrWhiteSpace(userInstruction) Then
            AppendUserHtml(userInstruction)
            _history.Add(("user", userInstruction))
        End If

        ' Step 7: Run the discussion loop (reusing autorespond infrastructure)
        Await RunSortOutLoopAsync(maxRounds)

        ' Step 8: Restore original mission
        _currentMissionName = _sortOutOriginalMissionName
        _currentMissionPrompt = _sortOutOriginalMissionPrompt
        UpdateWindowTitle()
    End Function

    ''' <summary>
    ''' Shows a dialog to select the number of rounds for Sort Out.
    ''' </summary>
    ''' <returns>Selected number of rounds, or 0 if cancelled.</returns>
    Private Function ShowSortOutRoundsDialog() As Integer
        ' Build round count options
        Dim roundOptions As New List(Of String)()
        For i = 1 To MaxAutoRespondRounds
            roundOptions.Add(i.ToString())
        Next

        ' Restore persisted value or use default
        Dim savedRounds = DefaultRespondRounds
        Try
            Dim stored = My.Settings.SortOutMaxRounds
            If stored >= 1 AndAlso stored <= MaxAutoRespondRounds Then
                savedRounds = stored
            End If
        Catch
        End Try

        Dim p0 As New SharedMethods.InputParameter("Maximum Rounds", savedRounds.ToString(), roundOptions)
        Dim params() As SharedMethods.InputParameter = {p0}

        Dim result = ShowCustomVariableInputForm(
            "How many rounds (back-and-forth exchanges) should the discussion have at most?",
            $"{AN} - Sort It Out Rounds",
            params)

        If Not result Then
            Return 0 ' Cancelled
        End If

        Dim selectedRounds = DefaultRespondRounds
        Integer.TryParse(CStr(params(0).Value), selectedRounds)

        If selectedRounds < 1 OrElse selectedRounds > MaxAutoRespondRounds Then
            selectedRounds = DefaultRespondRounds
        End If

        ' Persist
        Try
            My.Settings.SortOutMaxRounds = selectedRounds
            My.Settings.Save()
        Catch
        End Try

        Return selectedRounds
    End Function

    ''' <summary>
    ''' Generates mission statements for Sort It Out using LLM.
    ''' </summary>
    Private Async Function GenerateSortOutMissionsAsync(userInstruction As String, maxRounds As Integer) As Task(Of (Success As Boolean, MainMission As String, ResponderMission As String))
        Try
            ShowAssistantThinking()

            ' Build the discussion context
            Dim discussionContext = BuildConversationForAutoResponder()
            If String.IsNullOrWhiteSpace(discussionContext) Then
                discussionContext = "(No discussion yet)"
            End If

            ' Build the knowledge context
            Dim knowledgeContext As New StringBuilder()
            If Not String.IsNullOrWhiteSpace(_knowledgeContent) Then
                knowledgeContext.AppendLine(_knowledgeContent)
            End If

            ' Include active document if checkbox checked
            If _chkIncludeActiveDoc.Checked Then
                Dim activeDocContent = GetActiveDocumentContent()
                If Not String.IsNullOrWhiteSpace(activeDocContent) Then
                    If knowledgeContext.Length > 0 Then
                        knowledgeContext.AppendLine()
                        knowledgeContext.AppendLine("--- User's Active Document ---")
                    End If
                    knowledgeContext.AppendLine(activeDocContent)
                End If
            End If

            Dim knowledgeText = knowledgeContext.ToString()
            If String.IsNullOrWhiteSpace(knowledgeText) Then
                knowledgeText = "(No knowledge loaded)"
            End If

            ' Build the prompt with placeholders replaced
            Dim prompt = ThisAddIn.SP_DiscussThis_SortOut
            prompt = prompt.Replace("{MaxRounds}", maxRounds.ToString())
            prompt = prompt.Replace("{Persona}", If(_currentPersonaPrompt, _currentPersonaName))
            prompt = prompt.Replace("{Location}", If(_context?.INI_Location, "Unknown"))
            prompt = prompt.Replace("{OtherPrompt}", userInstruction)
            prompt = prompt.Replace("{dateContext}", GetDateContext())
            prompt = prompt.Replace("{Discussion}", discussionContext)
            prompt = prompt.Replace("{Knowledge}", knowledgeText)

            ' Call LLM (not using alternate model for mission generation)
            Dim response = Await LLM(_context, prompt, "", "", "", 0, False, True)

            RemoveAssistantThinking()

            If String.IsNullOrWhiteSpace(response) Then
                Return (False, "", "")
            End If

            ' Parse the response - expecting two prompts separated by |||
            Dim parts = response.Split(New String() {"|||"}, StringSplitOptions.None)
            If parts.Length >= 2 Then
                Dim mainMission = parts(0).Trim()
                Dim responderMission = parts(1).Trim()

                If Not String.IsNullOrWhiteSpace(mainMission) AndAlso Not String.IsNullOrWhiteSpace(responderMission) Then
                    Return (True, mainMission, responderMission)
                End If
            End If

            Return (False, "", "")

        Catch ex As Exception
            RemoveAssistantThinking()
            AppendSystemMessage($"Error generating missions: {ex.Message}")
            Return (False, "", "")
        End Try
    End Function

    ''' <summary>
    ''' Shows a mission selector for Sort It Out manual selection.
    ''' </summary>
    ''' <param name="prompt">The prompt to show.</param>
    ''' <param name="settingsKey">The settings key for persisting selection.</param>
    ''' <returns>The selected mission prompt, or Nothing if cancelled.</returns>
    Private Function ShowSortOutMissionSelector(prompt As String, settingsKey As String) As String
        Dim missionPath = GetMissionFilePath()

        ' Ensure mission file exists
        If Not String.IsNullOrWhiteSpace(missionPath) Then
            EnsureMissionFileExists(missionPath)
        End If

        ' Reload missions
        LoadMissions()

        ' Build selection items
        Dim items As New List(Of SelectionItem)()

        ' First item: "No mission"
        Const NoMissionValue As Integer = -1
        items.Add(New SelectionItem("No mission", NoMissionValue))

        ' Mission items
        For i = 0 To _missions.Count - 1
            items.Add(New SelectionItem(_missions(i).DisplayName, i + 1))
        Next

        ' Last item: "Edit mission library"
        Const EditMissionValue As Integer = -2
        items.Add(New SelectionItem("Edit mission library...", EditMissionValue))

        ' Try to restore saved selection
        Dim defaultVal = NoMissionValue
        Try
            Dim saved = ""
            Select Case settingsKey
                Case "SortOutMainMissionManual"
                    saved = My.Settings.SortOutMainMissionManual
                Case "SortOutResponderMissionManual"
                    saved = My.Settings.SortOutResponderMissionManual
            End Select
            If Not String.IsNullOrEmpty(saved) Then
                For i = 0 To _missions.Count - 1
                    If _missions(i).Name.Equals(saved, StringComparison.OrdinalIgnoreCase) Then
                        defaultVal = i + 1
                        Exit For
                    End If
                Next
            End If
        Catch
        End Try

        While True
            Dim result = SelectValue(items, defaultVal, prompt, $"{AN} - Select Mission")

            If result = 0 Then
                Return Nothing ' Cancelled
            ElseIf result = NoMissionValue Then
                ' No mission selected
                Return ""
            ElseIf result = EditMissionValue Then
                ' Edit mission library
                If Not String.IsNullOrWhiteSpace(missionPath) Then
                    ShowTextFileEditor(missionPath, $"{AN} - Edit Missions:", False, _context)
                    LoadMissions()

                    ' Rebuild items
                    items.Clear()
                    items.Add(New SelectionItem("No mission", NoMissionValue))
                    For i = 0 To _missions.Count - 1
                        items.Add(New SelectionItem(_missions(i).DisplayName, i + 1))
                    Next
                    items.Add(New SelectionItem("Edit mission library...", EditMissionValue))
                End If
                ' Loop back to show selector again
            ElseIf result > 0 AndAlso result <= _missions.Count Then
                ' Mission selected
                Dim selected = _missions(result - 1)

                ' Persist selection
                Try
                    Select Case settingsKey
                        Case "SortOutMainMissionManual"
                            My.Settings.SortOutMainMissionManual = selected.Name
                        Case "SortOutResponderMissionManual"
                            My.Settings.SortOutResponderMissionManual = selected.Name
                    End Select
                    My.Settings.Save()
                Catch
                End Try

                Return selected.Prompt
            End If
        End While

        Return Nothing
    End Function

    ''' <summary>
    ''' Runs the Sort It Out discussion loop, reusing autorespond infrastructure.
    ''' </summary>
    Private Async Function RunSortOutLoopAsync(maxRounds As Integer) As Task
        _sortOutInProgress = True
        _autoRespondInProgress = True  ' Block autorespond while Sort It Out is running
        _autoRespondCancelled = False

        ' Disable input during Sort It Out
        Ui(Sub()
               _txtInput.Enabled = False
               _btnSend.Enabled = False
               _btnAutoRespond.Enabled = False
               _btnSortOut.Enabled = False
           End Sub)

        ' Determine display names
        Dim mainDisplayName = _currentPersonaName & " (Advocate)"
        Dim responderDisplayName = _currentPersonaName & " (Challenger)"

        ' Show progress bar
        Dim useProgressBar = (maxRounds > 1)
        If useProgressBar Then
            ShowProgressBarInSeparateThread($"{AN} Sort It Out", "Discussion in progress...")
            ProgressBarModule.CancelOperation = False
            ProgressBarModule.GlobalProgressMax = maxRounds
            ProgressBarModule.GlobalProgressValue = 0
            ProgressBarModule.GlobalProgressLabel = "Starting discussion..."
        End If

        ' Notify start
        AppendSystemMessage($"Sort It Out discussion started between {mainDisplayName} and {responderDisplayName} for up to {maxRounds} round(s).")

        Try
            Dim roundCount = 0
            Dim stopRequested = False

            ' First, get the main bot's initial response to the user's instruction
            ShowAssistantThinking()
            Dim mainResponse = Await GenerateSortOutMainBotResponseAsync(mainDisplayName, responderDisplayName)
            RemoveAssistantThinking()

            ' Check for stop word
            If mainResponse.Contains(AutoRespondStopWord) Then
                stopRequested = True
                mainResponse = mainResponse.Replace(AutoRespondStopWord, "").Trim()
            End If

            If Not String.IsNullOrWhiteSpace(mainResponse) Then
                AppendAssistantMarkdownWithName(mainResponse, mainDisplayName)
                ' Store with display name prefix for Sort It Out mode (like autoresponder)
                _history.Add(("assistant", $"{mainDisplayName}: {mainResponse}"))
            End If

            PersistChatHtml()
            PersistTranscriptLimited()

            ' Now alternate between responder and main bot
            While roundCount < maxRounds AndAlso Not _autoRespondCancelled AndAlso Not stopRequested
                roundCount += 1

                If useProgressBar Then
                    ProgressBarModule.GlobalProgressValue = roundCount
                    ProgressBarModule.GlobalProgressLabel = $"Round {roundCount} of {maxRounds}..."
                    If ProgressBarModule.CancelOperation Then
                        _autoRespondCancelled = True
                        Exit While
                    End If
                End If

                ' Responder (Challenger) responds
                ShowAutoResponderThinking(responderDisplayName)
                Dim responderMessage = Await GenerateSortOutResponderMessageAsync(mainDisplayName, responderDisplayName)
                RemoveAssistantThinking()

                ' Check for stop word
                If responderMessage.Contains(AutoRespondStopWord) Then
                    stopRequested = True
                    responderMessage = responderMessage.Replace(AutoRespondStopWord, "").Trim()
                End If

                If Not String.IsNullOrWhiteSpace(responderMessage) Then
                    AppendAutoResponderHtml(responderDisplayName, responderMessage)
                    _history.Add(("autoresponder", $"{responderDisplayName}: {responderMessage}"))
                End If

                If stopRequested OrElse _autoRespondCancelled Then
                    Exit While
                End If

                ' Main bot (Advocate) responds
                ShowAssistantThinking()
                Dim mainBotResponse = Await GenerateSortOutMainBotResponseAsync(mainDisplayName, responderDisplayName)
                RemoveAssistantThinking()

                ' Check for stop word
                If mainBotResponse.Contains(AutoRespondStopWord) Then
                    stopRequested = True
                    mainBotResponse = mainBotResponse.Replace(AutoRespondStopWord, "").Trim()
                End If

                If Not String.IsNullOrWhiteSpace(mainBotResponse) Then
                    AppendAssistantMarkdownWithName(mainBotResponse, mainDisplayName)
                    ' Store with display name prefix for Sort It Out mode (like autoresponder)
                    _history.Add(("assistant", $"{mainDisplayName}: {mainBotResponse}"))
                End If

                PersistChatHtml()
                PersistTranscriptLimited()

                ' Small delay
                Await Task.Delay(500)
            End While

            ' Summary message
            If _autoRespondCancelled Then
                AppendSystemMessage($"Sort It Out discussion cancelled after {roundCount} round(s).")
            ElseIf stopRequested Then
                AppendSystemMessage($"Sort It Out discussion completed after {roundCount} round(s) - participants came to an end.")
            Else
                AppendSystemMessage($"Sort It Out discussion completed - maximum of {roundCount} round(s) reached.")
            End If

            ' Offer summary if enough rounds completed
            Await ShowDiscussionSummaryAsync(roundCount)

        Catch ex As Exception
            AppendSystemMessage($"Sort It Out error: {ex.Message}")
        Finally
            If useProgressBar Then
                ProgressBarModule.CancelOperation = True
            End If

            _sortOutInProgress = False
            _autoRespondInProgress = False
            _autoRespondCancelled = False

            ' Re-enable input
            Ui(Sub()
                   _txtInput.Enabled = True
                   _btnSend.Enabled = True
                   _btnAutoRespond.Enabled = True
                   _btnSortOut.Enabled = True
                   _txtInput.Focus()
               End Sub)

            PersistChatHtml()
            PersistTranscriptLimited()
        End Try
    End Function

    ''' <summary>
    ''' Generates the main bot's response in Sort It Out mode.
    ''' </summary>
    Private Async Function GenerateSortOutMainBotResponseAsync(mainDisplayName As String, responderDisplayName As String) As Task(Of String)
        Dim dateContext = GetDateContext()
        Dim randomWord = GetRandomModifier()
        Dim locationContext = GetLocationContext()
        Dim languageInstruction = GetLanguageInstruction()

        ' Use the main bot's persona with the Sort It Out mission
        Dim basePrompt = If(Not String.IsNullOrEmpty(_currentPersonaPrompt),
                            _currentPersonaPrompt,
                            $"You are {_currentPersonaName}, participating in a structured discussion.")

        Dim missionClause = ""
        If Not String.IsNullOrEmpty(_sortOutMainMissionPrompt) Then
            missionClause = $" Your mission: {_sortOutMainMissionPrompt}"
        End If

        Dim systemPrompt = $"{basePrompt}{missionClause}. In your response, be {randomWord}. Do not start with a greeting or salutation. " &
                           $"You are {mainDisplayName}, discussing with {responderDisplayName}. " &
                           $"{DefaultAutoRespondBreakOff} {dateContext} {locationContext} {languageInstruction}"

        ' Build context
        Dim sb As New StringBuilder()
        sb.AppendLine($"You are {mainDisplayName}, in a structured discussion with {responderDisplayName}.")
        sb.AppendLine()

        If Not String.IsNullOrWhiteSpace(_knowledgeContent) Then
            sb.AppendLine("<Knowledge Base>")
            sb.AppendLine(_knowledgeContent)
            sb.AppendLine("</Knowledge Base>")
            sb.AppendLine()
        End If

        If _chkIncludeActiveDoc.Checked Then
            Dim activeDocContent = GetActiveDocumentContent()
            If Not String.IsNullOrWhiteSpace(activeDocContent) Then
                sb.AppendLine("<User's Active Document>")
                sb.AppendLine(activeDocContent)
                sb.AppendLine("</User's Active Document>")
                sb.AppendLine()
            End If
        End If

        sb.AppendLine("Conversation so far:")
        Dim convo = BuildConversationForAutoResponder()
        sb.AppendLine(convo)
        sb.AppendLine()
        sb.AppendLine($"Now respond as {mainDisplayName}:")

        Dim answer = Await CallLlmWithSelectedModelAsync(systemPrompt, sb.ToString())
        Return If(answer, "").Trim()
    End Function

    ''' <summary>
    ''' Generates the responder's message in Sort It Out mode.
    ''' </summary>
    Private Async Function GenerateSortOutResponderMessageAsync(mainDisplayName As String, responderDisplayName As String) As Task(Of String)
        Dim dateContext = GetDateContext()
        Dim randomWord = GetRandomModifier()
        Dim locationContext = GetLocationContext()
        Dim languageInstruction = GetLanguageInstruction()

        ' Use same persona but responder mission
        Dim basePrompt = If(Not String.IsNullOrEmpty(_autoRespondPersonaPrompt),
                            _autoRespondPersonaPrompt,
                            $"You are {_autoRespondPersonaName}, participating in a structured discussion.")

        Dim missionClause = ""
        If Not String.IsNullOrEmpty(_sortOutResponderMissionPrompt) Then
            missionClause = $" Your mission: {_sortOutResponderMissionPrompt}"
        End If

        Dim systemPrompt = $"{basePrompt}{missionClause}. In your response, be {randomWord}. Do not start with a greeting or salutation. " &
                           $"You are {responderDisplayName}, responding to {mainDisplayName}. " &
                           $"{DefaultAutoRespondBreakOff} {dateContext} {locationContext} {languageInstruction}"

        ' Build context
        Dim sb As New StringBuilder()
        sb.AppendLine($"You are {responderDisplayName}, responding to {mainDisplayName} in a structured discussion.")
        sb.AppendLine()

        If Not String.IsNullOrWhiteSpace(_knowledgeContent) Then
            sb.AppendLine("<Knowledge Base>")
            sb.AppendLine(_knowledgeContent)
            sb.AppendLine("</Knowledge Base>")
            sb.AppendLine()
        End If

        If _chkIncludeActiveDoc.Checked Then
            Dim activeDocContent = GetActiveDocumentContent()
            If Not String.IsNullOrWhiteSpace(activeDocContent) Then
                sb.AppendLine("<User's Active Document>")
                sb.AppendLine(activeDocContent)
                sb.AppendLine("</User's Active Document>")
                sb.AppendLine()
            End If
        End If

        sb.AppendLine("Conversation so far:")
        Dim convo = BuildConversationForAutoResponder()
        sb.AppendLine(convo)
        sb.AppendLine()
        sb.AppendLine($"Now respond as {responderDisplayName}:")

        Dim answer = Await CallLlmWithSelectedModelAsync(systemPrompt, sb.ToString())
        Return If(answer, "").Trim()
    End Function

#End Region

#Region "Discussion Summary"

    ''' <summary>
    ''' Generates and displays a summary of the discussion after autorespond or sort out completes.
    ''' </summary>
    ''' <param name="roundCount">Number of rounds completed.</param>
    Private Async Function ShowDiscussionSummaryAsync(roundCount As Integer) As Task
        If roundCount < MinRoundsForAutoSummary Then Return

        ' Ensure progress bar is closed before showing summary dialog
        ProgressBarModule.CancelOperation = True

        Try
            ' Build the discussion transcript
            Dim discussionText = BuildConversationForAutoResponder()
            If String.IsNullOrWhiteSpace(discussionText) Then Return

            ' Ask user if they want a summary
            Dim answer = ShowCustomYesNoBox(
                $"The discussion completed {roundCount} rounds. Would you like to generate a summary of the key points?",
                "Yes, summarize", "No, skip")

            If answer <> 1 Then Return

            ShowAssistantThinking()

            ' Call LLM with the summary prompt
            Dim summaryResult = Await CallLlmWithSelectedModelAsync(
                _context.SP_DiscussThis_SumUp,
                "<TEXTTOPROCESS>" & discussionText & "</TEXTTOPROCESS>")

            RemoveAssistantThinking()

            If String.IsNullOrWhiteSpace(summaryResult) Then
                AppendSystemMessage("Could not generate summary.")
                Return
            End If

            ' Convert Markdown to HTML and display
            ShowDiscussionSummaryHtml(summaryResult)

        Catch ex As Exception
            RemoveAssistantThinking()
            AppendSystemMessage($"Error generating summary: {ex.Message}")
        End Try
    End Function

    ''' <summary>
    ''' Displays the discussion summary in an HTML window.
    ''' </summary>
    Private Sub ShowDiscussionSummaryHtml(summaryMarkdown As String)
        Try
            Dim htmlText As String = Markdig.Markdown.ToHtml(summaryMarkdown, _mdPipeline)

            Dim fullHtml As String =
                "<!DOCTYPE html>" &
                "<html><head>" &
                "  <meta charset=""utf-8"" />" &
                "  <style>" &
                "    body { font-family: 'Segoe UI', Tahoma, Arial, sans-serif; font-size: 10pt; line-height: 1.5; padding: 10px; }" &
                "    h1, h2, h3 { color: #003366; margin-top: 0.8em; margin-bottom: 0.4em; }" &
                "    ul, ol { margin-left: 1.5em; padding-left: 0.5em; }" &
                "    li { margin-bottom: 0.3em; }" &
                "    p { margin: 0.5em 0; }" &
                "  </style>" &
                "</head><body>" &
                htmlText &
                "</body></html>"

            ShowHTMLCustomMessageBox(fullHtml, $"{SharedMethods.AN} Discussion Summary")

        Catch ex As Exception
            ' Fallback to plain text
            ShowCustomMessageBox(summaryMarkdown, $"{SharedMethods.AN} Discussion Summary")
        End Try
    End Sub

#End Region

#Region "Active Document Context (with selection/cursor + bubbles)"

    ''' <summary>
    ''' Number of characters to capture before and after the cursor when there is no explicit selection.
    ''' </summary>
    Private Const CursorContextCharCount As Integer = 25

    ''' <summary>
    ''' Extracts the active Word document content for prompt inclusion, including:
    ''' - document name + full text
    ''' - either selected text OR cursor context (if selection is empty)
    ''' - Word bubble comments (when available) via `ThisAddIn.BubblesExtract`
    ''' </summary>
    ''' <returns>Formatted string suitable for direct prompt inclusion.</returns>
    Private Function GetActiveDocumentContent() As String
        Try
            Dim app = Globals.ThisAddIn.Application
            If app Is Nothing Then
                Debug.WriteLine("GetActiveDocumentContent: app is Nothing")
                Return ""
            End If
            If app.Documents Is Nothing OrElse app.Documents.Count = 0 Then
                Debug.WriteLine("GetActiveDocumentContent: No documents open")
                Return ""
            End If
            If app Is Nothing OrElse app.Documents Is Nothing OrElse app.Documents.Count = 0 Then Return ""

            Dim doc = app.ActiveDocument
            If doc Is Nothing Then Return ""

            Dim sb As New StringBuilder()
            sb.AppendLine($"Document: {doc.Name}")

            Dim fullText As String = ""
            Dim docBubbles As String = ""

            Dim haveWindow As Boolean = False
            Dim originalRevisionsView As Microsoft.Office.Interop.Word.WdRevisionsView = Nothing
            Dim originalShowRevisions As Boolean = False

            Try
                haveWindow = (app.ActiveWindow IsNot Nothing AndAlso app.ActiveWindow.View IsNot Nothing)
                If haveWindow Then
                    originalRevisionsView = app.ActiveWindow.View.RevisionsView
                    originalShowRevisions = app.ActiveWindow.View.ShowRevisionsAndComments

                    With app.ActiveWindow.View
                        .RevisionsView = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal
                        .ShowRevisionsAndComments = False
                    End With
                End If

                fullText = doc.Content.Text

                Try
                    docBubbles = ThisAddIn.BubblesExtract(doc.Content, True)
                Catch
                    docBubbles = ""
                End Try

            Finally
                If haveWindow Then
                    With app.ActiveWindow.View
                        .RevisionsView = originalRevisionsView
                        .ShowRevisionsAndComments = originalShowRevisions
                    End With
                End If
            End Try

            Dim selectionBlock As String = BuildSelectionOrCursorContextWithBubbles(doc)
            If Not String.IsNullOrWhiteSpace(selectionBlock) Then
                sb.AppendLine()
                sb.AppendLine(selectionBlock.TrimEnd())
            End If

            sb.AppendLine()
            sb.AppendLine("Full document text:")
            sb.AppendLine(fullText)

            If Not String.IsNullOrWhiteSpace(docBubbles) Then
                sb.AppendLine()
                sb.AppendLine("Comments / bubbles:")
                sb.AppendLine(docBubbles)
            End If

            Return sb.ToString()

        Catch ex As Exception
            Debug.WriteLine($"GetActiveDocumentContent exception: {ex.Message}")
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Builds a context block for either the current selection (if any) or a cursor-context window.
    ''' Includes bubble comments found within that selection/cursor range.
    ''' </summary>
    ''' <param name="doc">Active document.</param>
    ''' <returns>Formatted block or empty string.</returns>
    Private Function BuildSelectionOrCursorContextWithBubbles(doc As Microsoft.Office.Interop.Word.Document) As String
        Try
            Dim app = doc.Application
            Dim sel = app.Selection
            If sel Is Nothing Then Return ""

            Dim bubbles As String = ""

            ' If actual selection exists, use it
            If sel.Start <> sel.End AndAlso Not String.IsNullOrWhiteSpace(sel.Text) Then

                Try
                    bubbles = ThisAddIn.BubblesExtract(sel.Range, True)
                Catch
                    bubbles = ""
                End Try

                Dim sb As New StringBuilder()
                sb.AppendLine("User selection:")
                sb.AppendLine(sel.Text.Trim())

                If Not String.IsNullOrWhiteSpace(bubbles) Then
                    sb.AppendLine()
                    sb.AppendLine("Selection comments / bubbles:")
                    sb.AppendLine(bubbles)
                End If

                Return sb.ToString()
            End If

            ' Otherwise: capture cursor context (N chars before/after) + bubbles in that range
            Dim cursorPos As Integer = sel.Start
            Dim docStart As Integer = doc.Content.Start
            Dim docEnd As Integer = doc.Content.End

            Dim contextStart As Integer = Math.Max(docStart, cursorPos - CursorContextCharCount)
            Dim contextEnd As Integer = Math.Min(docEnd, cursorPos + CursorContextCharCount)

            Dim beforeRange = doc.Range(contextStart, cursorPos)
            Dim afterRange = doc.Range(cursorPos, contextEnd)

            Dim contextText As String = beforeRange.Text & "[cursor is here]" & afterRange.Text

            Dim cursorRange = doc.Range(contextStart, contextEnd)
            bubbles = ""
            Try
                bubbles = ThisAddIn.BubblesExtract(cursorRange, True)
            Catch
                bubbles = ""
            End Try

            Dim sb2 As New StringBuilder()
            sb2.AppendLine("Cursor context:")
            sb2.AppendLine(contextText.Trim())

            If Not String.IsNullOrWhiteSpace(bubbles) Then
                sb2.AppendLine()
                sb2.AppendLine("Cursor-range comments / bubbles:")
                sb2.AppendLine(bubbles)
            End If

            Return sb2.ToString()

        Catch
            Return ""
        End Try
    End Function

#End Region

#Region "Helpers"

    ''' <summary>
    ''' Determines 'Morning/Afternoon/Evening' from the current hour.
    ''' </summary>
    Private Shared Function GetPartOfDay() As String
        Dim h = DateTime.Now.Hour
        If h < 12 Then Return "Morning"
        If h < 18 Then Return "Afternoon"
        Return "Evening"
    End Function

    ''' <summary>
    ''' Detects whether the restored HTML ended on an alternate-model state by checking for model switch messages.
    ''' </summary>
    Private Function ChatHtmlIndicatesAlternateModel(html As String) As Boolean
        If String.IsNullOrEmpty(html) Then Return False

        Try
            Dim switchedToAlternateIdx = html.LastIndexOf("Switched to alternate model", StringComparison.OrdinalIgnoreCase)
            Dim switchedToSecondaryIdx = html.LastIndexOf("Switched to secondary model", StringComparison.OrdinalIgnoreCase)
            Dim switchedBackIdx = html.LastIndexOf("Switched back to primary model", StringComparison.OrdinalIgnoreCase)

            Dim lastSwitchToIdx = Math.Max(switchedToAlternateIdx, switchedToSecondaryIdx)

            If lastSwitchToIdx < 0 Then Return False

            If switchedBackIdx < 0 OrElse switchedBackIdx < lastSwitchToIdx Then
                Return True
            End If

            Return False
        Catch
            Return False
        End Try
    End Function

#End Region

End Class