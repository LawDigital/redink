' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: DiscussInky.vb
' Purpose:
'   Word WinForms surface for persona/mission-based discussion workflows, including
'   optional knowledge/document context, multi-participant automation and tooling.
'
' Architecture / Function:
'   - Maintains its own bounded discussion history and renders Markdown conversation
'     output while persisting only user-facing session/preferences needed for reopening.
'   - Resolves personas/missions and optional file/directory knowledge, then composes that
'     context with the active Word document before invoking the shared LLM bridge.
'   - Autorespond and "Sort It Out" are orchestration modes over the same conversation
'     engine; alternate/special-task model scopes are temporary and restored after calls.
'   - Optional tools are executed through Globals.ThisAddIn.ExecuteToolingLoop rather than
'     reimplemented here; tool permissions/log display remain explicit session choices.
'   - Export back to Word delegates document creation/Markdown insertion to host helpers.
'   - File extraction, model configuration and common dialogs come from SharedLibrary.
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
Imports System.Xml.Linq
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
    Private Const AutoPersistKnowledgeThresholdChars As Integer = 25000
    Private Const DialogueArchiveFolderName As String = "redink-discuss-dialogues"
    Private Const DialogueArchiveFileExtension As String = ".dialogue.xml"
    ' Semantic-search index storage (durable, under %AppData%\redink; short names for Windows path limits).
    Private Const SessionIndexFolderName As String = "discuss-this"       ' per-session persisted index copies: di\<sid>\
    Private Const ArchiveIndexFolderSuffix As String = ".ix"    ' per-archive index copies: <archiveName>.ix\
    Private Const IndexCopyFileExtension As String = ".txt"
    ' Durable, My.Settings-independent pointers under %AppData%\redink\di (crash-safe).
    Private Const SessionIndexPointerFileName As String = "cur.id"    ' stores the stable session id
    Private Const SessionIndexStateFileName As String = "last.xml"    ' stores running-session index references
    Private Const ToolTrigger As String = "(ag)"
    Private Const KBTrigger As String = "(kb)"  ' Trigger to supplement with knowledge store results.
    Private Const DiscussLastCompactPromptSettingName As String = "DiscussLastCompactPrompt"

    ' Default fallback persona used when no persona library is configured
    Private Const DefaultPersonaName As String = "Discussion Partner"
    Private Const DefaultPersonaPrompt As String = "You are a wise, thoughtful and critical discussion partner. You analyze topics from multiple angles, challenge assumptions constructively, and help the user arrive at well-reasoned conclusions. You are knowledgeable across many domains and provide balanced, nuanced perspectives while being direct and honest in your assessments."

    Private _currentPersonaName As String = DefaultPersonaName
    Private _currentPersonaPrompt As String = DefaultPersonaPrompt

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
        ".txt", ".rtf", ".ini", ".csv", ".log",
        ".json", ".xml", ".html", ".htm",
        ".md", ".yaml", ".yml",
        ".vb", ".cs", ".js", ".ts", ".py", ".java", ".cpp", ".c", ".h", ".sql",
        ".doc", ".docx", ".xlsx", ".pptx",
        ".pdf",
        ".eml", ".msg",
        ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".tif", ".webp", ".svg",
        ".mp3", ".wav", ".ogg", ".flac", ".m4a", ".aac", ".wma", ".opus", ".webm",
        ".mp4", ".avi", ".mkv", ".mov", ".wmv"
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

    ''' <summary>
    ''' SplitContainer separating the chat transcript (Panel1) from the user input (Panel2).
    ''' The splitter bar allows the user to resize the input area by dragging.
    ''' </summary>
    Private ReadOnly _splitChat As New SplitContainer() With {
        .Dock = DockStyle.Fill,
        .Orientation = Orientation.Horizontal,
        .FixedPanel = FixedPanel.Panel2,
        .SplitterWidth = 6,
        .Panel2MinSize = 40,
        .Panel1MinSize = 100
    }

    Private ReadOnly _txtInput As TextBox = New TextBox() With {
        .Dock = DockStyle.Fill,
        .Multiline = True,
        .AcceptsReturn = True,
        .WordWrap = True,
        .ScrollBars = ScrollBars.Vertical
    }

    Private ReadOnly _toolTip As ToolTip = New ToolTip() With {
    .AutoPopDelay = 10000,
    .InitialDelay = 500,
    .ReshowDelay = 200
}

    Private ReadOnly _btnClear As Button = New Button() With {.Text = "Clear", .AutoSize = True}
    Private ReadOnly _btnSendToDoc As Button = New Button() With {.Text = "Send to Doc", .AutoSize = True}
    Private ReadOnly _btnInsertSelectionToDoc As Button = New Button() With {.Text = "Insert in Doc", .AutoSize = True}
    Private ReadOnly _btnClose As Button = New Button() With {.Text = "Close", .AutoSize = True}
    Private ReadOnly _btnSend As Button = New Button() With {.Text = $"Send", .AutoSize = True}
    Private ReadOnly _btnPersona As Button = New Button() With {.Text = "Persona", .AutoSize = True}
    Private ReadOnly _btnMission As Button = New Button() With {.Text = "Mission", .AutoSize = True}
    Private ReadOnly _btnEditPersona As Button = New Button() With {.Text = "Edit Local Persona Lib", .AutoSize = True}
    Private ReadOnly _btnKnowledge As Button = New Button() With {.Text = "Load Knowledge (Docs, Indexes)", .AutoSize = True}
    Private ReadOnly _btnManageDocs As Button = New Button() With {.Text = "Manage Docs", .AutoSize = True, .Enabled = False}
    Private ReadOnly _btnManageIndexes As Button = New Button() With {.Text = "Manage Indexes", .AutoSize = True, .Enabled = False}
    Private ReadOnly _btnArchive As Button = New Button() With {.Text = "Archive", .AutoSize = True}
    Private ReadOnly _btnAlternateModel As Button = New Button() With {.Text = "Alternate Model", .AutoSize = True}
    Private ReadOnly _chkIncludeActiveDoc As System.Windows.Forms.CheckBox = New System.Windows.Forms.CheckBox() With {.Text = "Include active document", .AutoSize = True}
    Private ReadOnly _chkPersistKnowledge As System.Windows.Forms.CheckBox = New System.Windows.Forms.CheckBox() With {.Text = "Persist knowledge temporarily", .AutoSize = True}
    Private ReadOnly _btnAutoRespond As Button = New Button() With {.Text = "Autorespond", .AutoSize = True}
    Private ReadOnly _btnSortOut As Button = New Button() With {.Text = "Sort It Out", .AutoSize = True}
    Private ReadOnly _btnTools As Button = New Button() With {.Text = Globals.ThisAddIn.ToolFriendlyName, .AutoSize = True}
    Private ReadOnly _btnTalkToMe As Button = New Button() With {.Text = "", .AutoSize = True}
    Private ReadOnly _chkEnableTooling As System.Windows.Forms.CheckBox = New System.Windows.Forms.CheckBox() With {.Text = $"Enable {Globals.ThisAddIn.ToolFriendlyName.ToLower}", .AutoSize = True}
    Private ReadOnly _chkAdvancedTools As System.Windows.Forms.CheckBox = New System.Windows.Forms.CheckBox() With {.Text = "Advanced tools", .AutoSize = True}
    Private ReadOnly _chkShowToolingLog As System.Windows.Forms.CheckBox = New System.Windows.Forms.CheckBox() With {.Text = "Tooling log", .AutoSize = True, .Checked = True}
    Private ReadOnly _chkInkyMemory As System.Windows.Forms.CheckBox = New System.Windows.Forms.CheckBox() With {.Text = "Inky Memory", .AutoSize = True, .Checked = My.Settings.DiscussInkyMemory}
    Private ReadOnly _lnkEditMemory As New LinkLabel() With {
        .Text = "Edit",
        .AutoSize = True,
        .Visible = My.Settings.DiscussInkyMemory,
        .Margin = New Padding(0, 5, 0, 0)
    }

    ' State
    Private _htmlReady As Boolean = False
    Private ReadOnly _htmlQueue As New List(Of String)()
    Private _persistAfterHtmlFlush As Boolean = False
    Private _lastThinkingId As String = Nothing
    Private ReadOnly _history As New List(Of (Role As String, Content As String))()
    Private _knowledgeContent As String = Nothing
    Private _knowledgeFilePath As String = Nothing
    Private _welcomeInProgress As Integer = 0
    Private _personaSelectedThisSession As Boolean = False
    Private _isUpdatingPersistCheckbox As Boolean = False ' Prevents recursive event handling    
    Private _toolingControlsInitialized As Boolean = False
    Private _suppressToolingLogPreferenceSync As Boolean = False
    Private _noPersonaLibraryConfigured As Boolean = False ' True when no persona path is defined
    Private _suppressTalkToMeForwarding As Boolean = False
    Private _activeDialogueArchiveName As String = ""
    Private _activeDialogueArchiveFilePath As String = ""
    Private _activeDialogueArchiveBaselineHash As String = ""
    Private _persistedKnowledgeCloseWarningAcknowledged As Boolean = False

    ' Serializes all operations that touch the LLM or the knowledge/index collections (sending a
    ' message, loading knowledge, converting/creating an index). Prevents parallel, conflicting
    ' tasks (e.g. iterating _attachedIndexes while another task mutates it) and the resulting
    ' "collection was modified" errors, and blocks overlapping LLM usage.
    Private _exclusiveBusy As Boolean = False
    Private _workingIndicatorId As String = Nothing

    ''' <summary>
    ''' Attempts to enter the single exclusive operation slot. Returns False (and notifies the user)
    ''' when another operation is already running. On success, disables the chat input and the
    ''' knowledge/index buttons until <see cref="EndExclusive"/> is called.
    ''' </summary>
    Private Function TryBeginExclusive(taskLabel As String) As Boolean
        If _exclusiveBusy Then
            AppendSystemMessage($"Please wait — {taskLabel} cannot start while another operation is still running.")
            Return False
        End If

        _exclusiveBusy = True
        Ui(Sub()
               _txtInput.Enabled = False
               _btnSend.Enabled = False
               _btnKnowledge.Enabled = False
               _btnManageDocs.Enabled = False
               _btnManageIndexes.Enabled = False
           End Sub)
        Return True
    End Function

    ''' <summary>
    ''' Releases the exclusive operation slot and re-enables the input and knowledge/index buttons.
    ''' Button availability is recomputed from the current state via <see cref="UpdateWindowTitle"/>.
    ''' </summary>
    Private Sub EndExclusive()
        _exclusiveBusy = False
        Ui(Sub()
               _txtInput.Enabled = True
               _btnSend.Enabled = True
               _btnKnowledge.Enabled = True
               _txtInput.Focus()
           End Sub)
        UpdateWindowTitle()
    End Sub

    ''' <summary>
    ''' Shows a neutral, system-styled progress line in the transcript (not attributed to the
    ''' persona), used for file/index operations where the persona is not actually "thinking".
    ''' </summary>
    Private Sub ShowWorkingIndicator(text As String)
        _workingIndicatorId = "working-" & Guid.NewGuid().ToString("N")
        AppendHtml($"<div id=""{_workingIndicatorId}"" class='msg system'><span class='content'>{WebUtility.HtmlEncode(text)}</span></div>")
    End Sub

    ''' <summary>Updates the current working indicator text, if one is shown.</summary>
    Private Sub UpdateWorkingIndicator(text As String)
        If String.IsNullOrEmpty(_workingIndicatorId) Then Return
        Ui(Sub()
               Try
                   If _chat.Document IsNot Nothing Then
                       _chat.Document.InvokeScript("setThinkingText", New Object() {_workingIndicatorId, text})
                   End If
               Catch
               End Try
           End Sub)
    End Sub

    ''' <summary>Removes the current working indicator, if one is shown.</summary>
    Private Sub RemoveWorkingIndicator()
        If String.IsNullOrEmpty(_workingIndicatorId) Then Return
        Ui(Sub()
               Try
                   If _chat.Document IsNot Nothing Then
                       _chat.Document.InvokeScript("removeById", New Object() {_workingIndicatorId})
                   End If
               Catch
               Finally
                   _workingIndicatorId = Nothing
               End Try
           End Sub)
    End Sub

    ' Attached semantic-search indexes for the current session (queried alongside plain knowledge).
    Private ReadOnly _attachedIndexes As New List(Of DiscussIndexRef)()
    ' Short, stable per-session id used only for the persisted index folder name (di\<sid>\).
    Private _sessionIndexId As String = Nothing
    ' Ensures the "persist this index?" prompt is raised at most once per knowledge set, so loading
    ' several files/indexes in one operation does not ask repeatedly. Reset on delete/replace.
    Private _indexPersistenceOffered As Boolean = False
    ' Per-index semantic conversation state (previously used segment ids), keyed by index id.
    Private ReadOnly _indexConversationState As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)

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
    ''' Reference to an attached semantic-search index for the current session.
    ''' The index file is queried standalone via the retriever; byte offsets and the
    ''' SHA-256 guard require that the exact bytes are never altered or re-encoded.
    ''' </summary>
    Private Class DiscussIndexRef
        Public Property Id As String = ""              ' short id, e.g. "i0"
        Public Property DisplayName As String = ""     ' original file name (for citations)
        Public Property ActivePath As String = ""      ' path currently queried (in place or a durable copy)
        Public Property OriginalPath As String = ""    ' external source file this index was attached from (empty for indexes created in-session)
        Public Property ContentSha256 As String = ""   ' from the index's own JSON header
    End Class

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

    Private Enum KnowledgeDocumentManagerAction
        None = 0
        CompactSelected = 1
        DeleteSelected = 2
        EditSelected = 3
        IndexSelected = 4
    End Enum

    Private Structure KnowledgeDocumentEntry
        Public Number As Integer
        Public Name As String
        Public Content As String
        Public StartIndex As Integer
        Public Length As Integer
        Public IsTagged As Boolean

        Public ReadOnly Property DisplayText As String
            Get
                Dim displayName As String = If(String.IsNullOrWhiteSpace(Name), "Knowledge", Name)
                Dim prefix As String = If(IsTagged, $"document{Number}", "knowledge")
                Return $"{prefix} - {displayName} ({If(Content, "").Length:N0} chars)"
            End Get
        End Property
    End Structure

    Private Structure KnowledgeDocumentSelectionItem
        Public Number As Integer
        Public DisplayText As String

        Public Sub New(number As Integer, displayText As String)
            Me.Number = number
            Me.DisplayText = displayText
        End Sub

        Public Overrides Function ToString() As String
            Return DisplayText
        End Function
    End Structure

    Private Structure KnowledgeDocumentManagerResult
        Public Action As KnowledgeDocumentManagerAction
        Public SelectedDocumentNumbers As List(Of Integer)
    End Structure

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
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0F, 96.0F)
        Me.AutoScaleMode = AutoScaleMode.Dpi
        Me.FormBorderStyle = FormBorderStyle.Sizable
        Me.StartPosition = FormStartPosition.Manual
        Me.MinimumSize = New System.Drawing.Size(780, 480)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0F)
        _btnTalkToMe.Text = Char.ConvertFromUtf32(&H1F5E3) & ChrW(&HFE0F)
        _btnTalkToMe.Font = New System.Drawing.Font("Segoe UI Emoji", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        _btnTalkToMe.AutoSize = False
        _btnTalkToMe.Padding = New Padding(5, 0, 5, 0)
        _btnTalkToMe.Margin = _btnTools.Margin
        _btnTalkToMe.Visible = Globals.ThisAddIn.IsTalkToMeAvailable()
        _btnTalkToMe.Enabled = _btnTalkToMe.Visible
        ' Do NOT set Me.TopMost = True.
        ' Child dialogs are parented via SharedMethods.PushDialogOwner(Me) and the
        ' shared Show* helpers already re-assert TopMost themselves on Shown,
        ' so they will always come to the foreground even over Word.
        Try
            Me.Icon = Icon.FromHandle(New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard)).GetHicon())
        Catch
        End Try

        ' Layout
        Dim table As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 2,
            .Padding = New Padding(10)
        }
        table.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        table.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        table.RowStyles.Add(New RowStyle(SizeType.AutoSize))

        _txtInput.Margin = New Padding(0, 0, 0, 0)
        _txtInput.Font = New System.Drawing.Font(_txtInput.Font.FontFamily, 10.0F, _txtInput.Font.Style)

        ' Place chat and input into the SplitContainer
        _splitChat.Panel1.Controls.Add(_chat)
        _splitChat.Panel2.Controls.Add(_txtInput)
        _splitChat.SplitterDistance = 300 ' Default: generous space for chat transcript

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
        pnlButtons.Controls.Add(_btnManageDocs)
        pnlButtons.Controls.Add(_btnManageIndexes)
        pnlButtons.Controls.Add(_btnArchive)

        ' Show alternate model button if either second API is configured or an alternate INI exists
        If _context.INI_SecondAPI OrElse Not String.IsNullOrWhiteSpace(_context.INI_AlternateModelPath) Then
            UpdateAlternateModelButtonText()
            pnlButtons.Controls.Add(_btnAlternateModel)
        End If

        pnlButtons.Controls.Add(_btnClear)
        pnlButtons.Controls.Add(_btnSendToDoc)
        pnlButtons.Controls.Add(_btnInsertSelectionToDoc)
        pnlButtons.Controls.Add(_btnClose)
        pnlButtons.Controls.Add(_btnAutoRespond)
        pnlButtons.Controls.Add(_btnSortOut)
        pnlButtons.Controls.Add(_btnTools)
        pnlButtons.Controls.Add(_btnTalkToMe)
        pnlButtons.Controls.Add(_chkEnableTooling)
        pnlButtons.Controls.Add(_chkAdvancedTools)
        pnlButtons.Controls.Add(_chkIncludeActiveDoc)
        pnlButtons.Controls.Add(_chkPersistKnowledge)
        pnlButtons.Controls.Add(_chkShowToolingLog)
        pnlButtons.Controls.Add(_chkInkyMemory)
        pnlButtons.Controls.Add(_lnkEditMemory)


        table.Controls.Add(_splitChat, 0, 0)
        table.Controls.Add(pnlButtons, 0, 1)
        Me.Controls.Add(table)

        pnlButtons.PerformLayout()
        Dim talkToMeButtonWidth As Integer =
            TextRenderer.MeasureText(_btnTalkToMe.Text, _btnTalkToMe.Font).Width +
            _btnTalkToMe.Padding.Left +
            _btnTalkToMe.Padding.Right +
            2
        _btnTalkToMe.Size = New System.Drawing.Size(talkToMeButtonWidth, _btnTools.Height)
        _btnTalkToMe.MinimumSize = _btnTalkToMe.Size
        _btnTalkToMe.MaximumSize = _btnTalkToMe.Size

        InitializeButtonToolTips()

        _mdPipeline = New MarkdownPipelineBuilder().
            UseAdvancedExtensions().
            UseSoftlineBreakAsHardlineBreak().
            Build()

        ' Event handlers
        AddHandler Me.Load, AddressOf OnLoadForm
        AddHandler Me.FormClosing, AddressOf OnFormClosing
        AddHandler _btnSend.Click, AddressOf OnSend
        AddHandler _btnClear.Click, AddressOf OnClear
        AddHandler _btnSendToDoc.Click, AddressOf OnSendToDoc
        AddHandler _btnInsertSelectionToDoc.Click, AddressOf OnInsertSelectionToDoc
        AddHandler _btnClose.Click, AddressOf OnClose
        AddHandler _btnPersona.Click, AddressOf OnSelectPersona
        AddHandler _btnMission.Click, AddressOf OnSelectMission
        AddHandler _btnEditPersona.Click, AddressOf OnEditLocalPersona
        AddHandler _btnKnowledge.Click, AddressOf OnLoadKnowledge
        AddHandler _btnManageDocs.Click, AddressOf OnManageKnowledgeDocumentsClick
        AddHandler _btnManageIndexes.Click, AddressOf OnManageIndexesClick
        AddHandler _btnArchive.Click, AddressOf OnArchiveClick
        AddHandler _btnAlternateModel.Click, AddressOf OnAlternateModelClick
        AddHandler _txtInput.KeyDown, AddressOf OnInputKeyDown
        AddHandler _txtInput.KeyPress, AddressOf OnInputKeyPress
        AddHandler _chat.DocumentCompleted, AddressOf Chat_DocumentCompleted
        AddHandler _chat.Navigating, AddressOf Chat_Navigating
        AddHandler _chat.NewWindow, AddressOf Chat_NewWindow
        AddHandler _chkIncludeActiveDoc.CheckedChanged, AddressOf OnIncludeActiveDocChanged
        AddHandler _chkPersistKnowledge.CheckedChanged, AddressOf OnPersistKnowledgeChanged
        AddHandler _chkPersistKnowledge.MouseEnter, AddressOf OnPersistKnowledgeTooltipRefresh
        AddHandler _btnAutoRespond.Click, AddressOf OnAutoRespondClick
        AddHandler _btnSortOut.Click, AddressOf OnSortOutClick
        AddHandler _btnTools.Click, AddressOf OnToolsClick
        AddHandler _btnTalkToMe.Click, AddressOf OnTalkToMeClick
        AddHandler _chkEnableTooling.CheckedChanged, AddressOf OnEnableToolingChanged
        AddHandler _chkAdvancedTools.CheckedChanged, AddressOf OnAdvancedToolsChanged
        AddHandler _chkShowToolingLog.CheckedChanged, AddressOf OnShowToolingLogChanged
        AddHandler _chkInkyMemory.CheckedChanged, AddressOf OnInkyMemoryChanged
        AddHandler _lnkEditMemory.LinkClicked, AddressOf OnEditMemoryClicked
        AddHandler _txtInput.MouseWheel, AddressOf OnInputMouseWheel
        AddHandler Microsoft.Win32.SystemEvents.DisplaySettingsChanged, AddressOf OnDisplaySettingsChanged

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
        Return "Always respond in the same language the user uses in their messages, regardless of the language of these instructions or the knowledge base. However, generally follow language instructions in your mission and persona description."
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

    Private Sub BringDiscussFormToFront()
        If Me.IsDisposed Then Return

        Try
            If Me.InvokeRequired Then
                Me.BeginInvoke(New System.Windows.Forms.MethodInvoker(AddressOf BringDiscussFormToFront))
                Return
            End If

            If Me.WindowState = System.Windows.Forms.FormWindowState.Minimized Then
                Me.WindowState = System.Windows.Forms.FormWindowState.Normal
            End If

            SharedMethods.EnsureVisibleOnScreen(Me)

            Me.Show()
            Me.Activate()
            Me.BringToFront()
            _txtInput.Focus()

        Catch
        End Try
    End Sub

    Private Sub OnTalkToMeClick(sender As Object, e As EventArgs)
        Try
            Globals.ThisAddIn.ShowTalkToMeWidget(AddressOf RestoreFocusAfterTalkToMeStart)
        Catch ex As Exception
            AppendSystemMessage($"Could not open TalkToMe: {ex.Message}")
        End Try
    End Sub

    Private Sub RestoreFocusAfterTalkToMeStart()
        If Me.IsDisposed Then
            Return
        End If

        Ui(
            Sub()
                Try
                    If Me.WindowState = FormWindowState.Minimized Then
                        Me.WindowState = FormWindowState.Normal
                    End If

                    SharedMethods.EnsureVisibleOnScreen(Me)
                    Me.Show()
                    Me.Activate()
                    Me.BringToFront()
                    _txtInput.Focus()
                Catch
                End Try
            End Sub)
    End Sub

    Private Sub ForwardOutputToTalkToMe(speakerName As String, outputText As String)
        If _suppressTalkToMeForwarding Then
            Return
        End If

        Try
            Globals.ThisAddIn.SubmitTalkToMeExternalSpeech(speakerName, outputText)
        Catch
        End Try
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

        If Not String.IsNullOrWhiteSpace(_activeDialogueArchiveName) Then
            title &= $" {{Archive: {_activeDialogueArchiveName}}}"
        End If

        If Not String.IsNullOrEmpty(_knowledgeFilePath) Then
            title &= $" - {Path.GetFileName(_knowledgeFilePath)}"
        End If

        ' Show current model in title if alternate is selected
        If _alternateModelSelected AndAlso Not String.IsNullOrWhiteSpace(_alternateModelDisplayName) Then
            title &= $" (using {_alternateModelDisplayName})"
        End If

        Ui(
            Sub()
                Me.Text = title
                _btnManageDocs.Enabled = Not String.IsNullOrWhiteSpace(_knowledgeContent)
                _btnManageIndexes.Enabled = _attachedIndexes.Count > 0
            End Sub)
    End Sub

    ''' <summary>
    ''' Refreshes the Send button label with the current persona name.
    ''' </summary>
    Private Sub UpdateSendButtonText()
        Ui(Sub() _btnSend.Text = $"Send to {_currentPersonaName}")
    End Sub

    ''' <summary>
    ''' Initializes tooltips for the action buttons.
    ''' </summary>
    Private Sub InitializeButtonToolTips()
        _toolTip.SetToolTip(_btnSend, "Send the current prompt to the selected discussion persona.")
        _toolTip.SetToolTip(_btnPersona, "Select the persona for this discussion.")
        _toolTip.SetToolTip(_btnMission, "Select or clear the current mission.")
        _toolTip.SetToolTip(_btnEditPersona, "Open the local persona library for editing.")
        _toolTip.SetToolTip(_btnKnowledge, "Load a knowledge file or a folder of knowledge files.")
        _toolTip.SetToolTip(_btnManageDocs, "Compact, delete, or edit knowledge documents already loaded into the current discussion.")
        _toolTip.SetToolTip(_btnManageIndexes, "Remove attached searchable indexes, or convert a file into a new searchable index.")
        _toolTip.SetToolTip(_btnArchive, "Store, restore, update, or delete archived discussions.")
        _toolTip.SetToolTip(_btnAlternateModel, "Switch between the primary model and an alternate or secondary model.")
        _toolTip.SetToolTip(_btnClear, "Clear the current discussion and start a new one.")
        _toolTip.SetToolTip(_btnSendToDoc, "Export the full discussion to a new Word document.")
        _toolTip.SetToolTip(_btnInsertSelectionToDoc, "Insert the selected chat text into the active Word document at the current selection or cursor.")
        _toolTip.SetToolTip(_btnClose, "Close this discussion window.")
        _toolTip.SetToolTip(_btnAutoRespond, "Start an automated back-and-forth discussion.")
        _toolTip.SetToolTip(_btnSortOut, "Run a structured Advocate versus Challenger discussion.")
        _toolTip.SetToolTip(_btnTools, $"Select the {Globals.ThisAddIn.ToolFriendlyName.ToLower} available for this discussion.")
        _toolTip.SetToolTip(_btnTalkToMe, "Show TalkToMe and return focus here after listening starts.")
    End Sub

    ''' <summary>
    ''' Gets the currently selected text from the discussion thread.
    ''' </summary>
    Private Function GetSelectedChatText() As String
        Try
            If _chat.Document Is Nothing Then
                Return ""
            End If

            Dim selectedTextObject As Object = _chat.Document.InvokeScript("getSelectedText")
            Dim selectedText As String = If(selectedTextObject, "").ToString()

            If String.IsNullOrWhiteSpace(selectedText) Then
                Return ""
            End If

            selectedText = selectedText.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Replace(vbLf, vbCrLf)
            Return selectedText.Trim()
        Catch
            Return ""
        End Try
    End Function

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
    ''' Gets the Red Ink storage directory in the user's application data folder.
    ''' </summary>
    Private Function GetRedInkStorageDirectoryPath() As String
        Dim storageDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "redink")
        Try
            If Not Directory.Exists(storageDir) Then
                Directory.CreateDirectory(storageDir)
            End If
        Catch
        End Try
        Return storageDir
    End Function

    ''' <summary>
    ''' Gets the full path to the persisted knowledge file under %AppData%\redink.
    ''' For backward compatibility, a legacy file in the temp folder is moved into the
    ''' durable location on first access if no durable copy exists yet.
    ''' </summary>
    ''' <returns>Full path to the persisted knowledge file.</returns>
    Private Function GetPersistedKnowledgeFilePath() As String
        Dim durablePath As String = Path.Combine(GetRedInkStorageDirectoryPath(), PersistedKnowledgeFileName)

        Try
            Dim legacyPath As String = Path.Combine(Path.GetTempPath(), PersistedKnowledgeFileName)
            If Not File.Exists(durablePath) AndAlso File.Exists(legacyPath) Then
                File.Move(legacyPath, durablePath)
            End If
        Catch
            ' Migration is best-effort; the durable path is still returned.
        End Try

        Return durablePath
    End Function

    Private Function HasPersistedKnowledgeForCloseWarning() As Boolean
        Try
            If Not _chkPersistKnowledge.Checked Then
                Return False
            End If

            Dim persistPath As String = GetPersistedKnowledgeFilePath()

            Return File.Exists(persistPath) AndAlso
                   (Not String.IsNullOrWhiteSpace(_knowledgeContent) OrElse
                    Not String.IsNullOrWhiteSpace(_cachedKnowledgeContent))
        Catch
            Return False
        End Try
    End Function

    Private Function IsCurrentKnowledgeBackedByLinkedArchive() As Boolean
        If Not HasTrackedDialogueArchive() Then
            Return False
        End If

        If String.IsNullOrWhiteSpace(_activeDialogueArchiveFilePath) OrElse
           Not File.Exists(_activeDialogueArchiveFilePath) Then
            Return False
        End If

        Try
            Dim archiveDoc As XDocument = XDocument.Load(_activeDialogueArchiveFilePath)
            Dim root As XElement = archiveDoc.Root

            If root Is Nothing Then
                Return False
            End If

            Dim archiveNameElement As XElement = root.Element("ArchiveName")
            Dim archiveName As String =
                If(
                    archiveNameElement IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(archiveNameElement.Value),
                    archiveNameElement.Value.Trim(),
                    GetArchiveNameFromFilePath(_activeDialogueArchiveFilePath))

            Dim archiveIndexDirectory As String = GetArchiveIndexDirectoryPath(archiveName)
            Dim archiveKnowledgeElement As XElement = root.Element("Knowledge")
            Dim archiveHasKnowledge As Boolean =
                archiveKnowledgeElement IsNot Nothing AndAlso
                Not String.IsNullOrWhiteSpace(archiveKnowledgeElement.Value)

            Dim currentHasKnowledge As Boolean =
                Not String.IsNullOrWhiteSpace(_knowledgeContent) OrElse
                Not String.IsNullOrWhiteSpace(_cachedKnowledgeContent)

            ' If plain knowledge is currently loaded, the linked archive must also contain plain knowledge.
            If currentHasKnowledge AndAlso Not archiveHasKnowledge Then
                Return False
            End If

            Dim archiveIndexHashes As New HashSet(Of String)(StringComparer.Ordinal)
            Dim archiveIndexesElement As XElement = root.Element("Indexes")

            If archiveIndexesElement IsNot Nothing Then
                For Each indexElement As XElement In archiveIndexesElement.Elements("Index")
                    Dim fileValue As String = GetXmlAttributeValue(indexElement, "file", "")
                    Dim hashValue As String = GetXmlAttributeValue(indexElement, "sha", "")
                    Dim archived As Boolean = GetXmlAttributeBoolean(indexElement, "archived", False)

                    If String.IsNullOrWhiteSpace(hashValue) Then
                        Dim resolvedPath As String = fileValue

                        If archived Then
                            resolvedPath = Path.Combine(archiveIndexDirectory, fileValue)
                        End If

                        hashValue = ComputeFileSha256(resolvedPath)
                    End If

                    If Not String.IsNullOrWhiteSpace(hashValue) Then
                        archiveIndexHashes.Add(hashValue.Trim())
                    End If
                Next
            End If

            For Each indexRef As DiscussIndexRef In _attachedIndexes
                If indexRef Is Nothing Then
                    Return False
                End If

                Dim activePath As String = If(indexRef.ActivePath, "").Trim()

                ' If the active path already points into the linked archive sidecar, it is safe.
                If Not String.IsNullOrWhiteSpace(activePath) Then
                    Dim activeDirectory As String = Path.GetDirectoryName(activePath)

                    If Not String.IsNullOrWhiteSpace(activeDirectory) AndAlso
                       String.Equals(activeDirectory, archiveIndexDirectory, StringComparison.OrdinalIgnoreCase) Then
                        Continue For
                    End If
                End If

                Dim currentHash As String = If(indexRef.ContentSha256, "").Trim()

                If String.IsNullOrWhiteSpace(currentHash) AndAlso
                   Not String.IsNullOrWhiteSpace(activePath) AndAlso
                   File.Exists(activePath) Then

                    currentHash = ComputeFileSha256(activePath).Trim()
                End If

                If String.IsNullOrWhiteSpace(currentHash) Then
                    Return False
                End If

                If Not archiveIndexHashes.Contains(currentHash) Then
                    Return False
                End If
            Next

            Return True
        Catch
            Return False
        End Try
    End Function

    Private Function ConfirmCloseWhenKnowledgePersisted() As Boolean
        If _persistedKnowledgeCloseWarningAcknowledged Then
            Return True
        End If

        ' Nothing to warn about when neither plain knowledge nor any attached index is present.
        If String.IsNullOrWhiteSpace(_knowledgeContent) AndAlso
           String.IsNullOrWhiteSpace(_cachedKnowledgeContent) AndAlso
           _attachedIndexes.Count = 0 Then
            Return True
        End If

        ' If the currently loaded knowledge and attached indexes are already backed by the linked
        ' DiscussThis archive, no knowledge-loss warning is needed on close.
        If IsCurrentKnowledgeBackedByLinkedArchive() Then
            Return True
        End If

        ' Determine if every currently loaded knowledge source is persisted durably.
        Dim persistPath As String = GetPersistedKnowledgeFilePath()

        Dim currentHasKnowledge As Boolean =
            Not String.IsNullOrWhiteSpace(_knowledgeContent) OrElse
            Not String.IsNullOrWhiteSpace(_cachedKnowledgeContent)

        Dim currentHasIndexes As Boolean = _attachedIndexes.Count > 0

        Dim knowledgePersisted As Boolean =
            Not currentHasKnowledge OrElse
            (_chkPersistKnowledge.Checked AndAlso File.Exists(persistPath))

        Dim indexesPersisted As Boolean =
            Not currentHasIndexes OrElse
            (_chkPersistKnowledge.Checked AndAlso Directory.Exists(GetSessionIndexDirectoryPath()))

        Dim isPersistedToFile As Boolean = knowledgePersisted AndAlso indexesPersisted

        If isPersistedToFile Then
            Return True
        End If

        Dim message As String =
            "DiscussInky is about to close." &
            vbCrLf &
            vbCrLf &
            "⚠ The loaded knowledge is NOT persisted and will not be available when you return, " &
            "unless the original source documents still exist in their original location or the current knowledge/indexes are already stored in the archive." &
            vbCrLf &
            vbCrLf &
            "The current chat will be stored." &
            vbCrLf &
            vbCrLf &
            "You can activate persistence with the applicable checkbox or use 'Archive' to store the chat. Do you want to close now? "

        Dim closeButtonText As String = "Close"
        Dim keepOpenButtonText As String = "Keep open"

        Dim answer As Integer = ShowCustomYesNoBox(message, closeButtonText, keepOpenButtonText)

        If answer = 1 Then
            _persistedKnowledgeCloseWarningAcknowledged = True
            Return True
        End If

        Return False
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
        ' Ensure window state is normal (not minimized or hidden)
        If Me.WindowState = FormWindowState.Minimized Then Me.WindowState = FormWindowState.Normal

        ' Ensure visible on at least one screen
        SharedMethods.EnsureVisibleOnScreen(Me)

        If Not Me.Visible Then
            If owner IsNot Nothing Then Me.Show(owner) Else Me.Show()
        End If

        Me.Activate()
        _txtInput.Focus()
        _txtInput.SelectAll()
    End Sub

    ' Ambient dialog-owner scope. Lifetime is bound to the form's window handle
    ' (NOT to Activated/Deactivate) because those events are pumped asynchronously:
    ' when a child modal returns, the next user line of code can call ShowCustom...
    ' BEFORE Activated has fired, which would leave the stack empty and cause the
    ' new dialog to be parented to the Office host (Word) instead of this form.
    Private _ownerScope As IDisposable

    ''' <summary>
    ''' Pushes this form onto the SharedLibrary dialog-owner stack as soon as it
    ''' has a window handle, so every shared modal dialog opened from this form
    ''' (or after a child modal returns) is correctly parented here.
    ''' </summary>
    Protected Overrides Sub OnHandleCreated(e As EventArgs)
        MyBase.OnHandleCreated(e)
        If _ownerScope Is Nothing Then
            _ownerScope = SharedMethods.PushDialogOwner(Me)
        End If
    End Sub

    ''' <summary>
    ''' Pops this form from the SharedLibrary dialog-owner stack when its handle
    ''' is destroyed (form closing/disposing).
    ''' </summary>
    Protected Overrides Sub OnHandleDestroyed(e As EventArgs)
        Dim scope = _ownerScope
        _ownerScope = Nothing
        If scope IsNot Nothing Then
            Try : scope.Dispose() : Catch : End Try
        End If
        MyBase.OnHandleDestroyed(e)
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
                        AppendSystemMessage($"Knowledge persisted to durable storage ({_cachedKnowledgeContent.Length:N0} characters).")
                    Catch ex As Exception
                        AppendSystemMessage($"Failed to persist knowledge: {ex.Message}")
                        ' Revert checkbox state
                        _isUpdatingPersistCheckbox = True
                        _chkPersistKnowledge.Checked = False
                        _isUpdatingPersistCheckbox = False
                        Return
                    End Try
                    PersistAttachedIndexes()
                ElseIf _attachedIndexes.Count > 0 Then
                    PersistAttachedIndexes()
                    AppendSystemMessage($"{_attachedIndexes.Count:N0} attached index/indexes persisted to durable storage.")
                Else
                    AppendSystemMessage("No knowledge loaded to persist. Load knowledge first, then check this box.")
                End If
            Else
                ' User unchecked the box. Persistence keeps durable copies of the loaded knowledge and
                ' attached indexes; turning it off removes those copies. The current session keeps its
                ' in-memory knowledge, and indexes revert to their original source files where possible.
                Dim sessionIndexDir As String = GetSessionIndexDirectoryPath()
                Dim hasPersistedKnowledgeFile As Boolean = File.Exists(persistPath)
                Dim hasPersistedIndexes As Boolean = _attachedIndexes.Count > 0 AndAlso Directory.Exists(sessionIndexDir)

                If hasPersistedKnowledgeFile OrElse hasPersistedIndexes Then
                    ' The plain knowledge can be reloaded from disk only if its original source still exists;
                    ' otherwise it simply stays in memory for the rest of this session.
                    Dim originalKnowledgePath As String = NormalizeKnowledgePathForSettings(_knowledgeFilePath)
                    Dim knowledgeRecoverable As Boolean =
                        String.IsNullOrWhiteSpace(_knowledgeContent) OrElse
                        (Not String.IsNullOrWhiteSpace(originalKnowledgePath) AndAlso
                         (File.Exists(originalKnowledgePath) OrElse Directory.Exists(originalKnowledgePath)))

                    ' Indexes that still have a reachable original file can revert; the rest exist only
                    ' as durable copies and would be removed.
                    Dim indexesRevertable As Integer =
                        _attachedIndexes.Where(Function(x) Not String.IsNullOrWhiteSpace(x.OriginalPath) AndAlso File.Exists(x.OriginalPath)).Count()
                    Dim indexesLost As Integer = _attachedIndexes.Count - indexesRevertable

                    Dim warning As New StringBuilder()
                    warning.AppendLine("Turning off persistence deletes the durable copies of your knowledge and any attached indexes.")
                    warning.AppendLine()

                    If indexesRevertable > 0 Then
                        warning.AppendLine($"{indexesRevertable:N0} attached index(es) will revert to their original files.")
                    End If
                    If indexesLost > 0 Then
                        warning.AppendLine(
                            $"{indexesLost:N0} attached index(es) exist only as durable copies and will be removed. " &
                            "Archive the dialogue first if you want to keep them.")
                    End If
                    If indexesRevertable > 0 OrElse indexesLost > 0 Then
                        warning.AppendLine()
                    End If

                    If knowledgeRecoverable Then
                        warning.AppendLine("The loaded knowledge will be reloaded from its original source file when needed.")
                    ElseIf Not String.IsNullOrWhiteSpace(_knowledgeContent) Then
                        warning.AppendLine(
                            "The loaded knowledge has no original source file, but it stays available in this session. " &
                            "It will be lost when Word restarts unless you archive the dialogue.")
                    End If

                    warning.AppendLine()
                    warning.Append("Do you want to proceed?")

                    Dim answer = ShowCustomYesNoBox(
                        warning.ToString(),
                        "Yes, delete persisted files",
                        "No, keep persistence on",
                        $"{AN} - Turn Off Persistence")

                    If answer = 1 Then
                        Try
                            If hasPersistedKnowledgeFile Then
                                File.Delete(persistPath)
                            End If
                        Catch ex As Exception
                            AppendSystemMessage($"Failed to delete persisted knowledge: {ex.Message}")
                        End Try

                        DeletePersistedSessionIndexes()

                        ' Durable index copies are now gone. Repoint each attached index back to its
                        ' original external file where that still exists; otherwise detach it. Plain
                        ' knowledge is deliberately left in memory so the chat continues this session.
                        Dim removedIndexes As Integer = 0
                        Dim revertedIndexes As Integer = 0
                        For Each indexRef In _attachedIndexes.ToList()
                            Dim durableCopyGone As Boolean =
                                String.IsNullOrWhiteSpace(indexRef.ActivePath) OrElse Not File.Exists(indexRef.ActivePath)

                            If durableCopyGone Then
                                If Not String.IsNullOrWhiteSpace(indexRef.OriginalPath) AndAlso File.Exists(indexRef.OriginalPath) Then
                                    indexRef.ActivePath = indexRef.OriginalPath
                                    revertedIndexes += 1
                                Else
                                    _attachedIndexes.Remove(indexRef)
                                    _indexConversationState.Remove(indexRef.Id)
                                    removedIndexes += 1
                                End If
                            End If
                        Next

                        ' Keep the in-memory knowledge alive for this session; do not clear _knowledgeContent.
                        _indexPersistenceOffered = False
                        UpdateWindowTitle()

                        Dim summary As New StringBuilder()
                        summary.Append("Persistence turned off. Durable copies removed.")
                        If revertedIndexes > 0 Then
                            summary.Append($" {revertedIndexes:N0} index(es) reverted to their original files.")
                        End If
                        If removedIndexes > 0 Then
                            summary.Append($" {removedIndexes:N0} index(es) detached.")
                        End If
                        AppendSystemMessage(summary.ToString())
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
    ''' Refreshes the persist-knowledge tooltip on demand (e.g., when the pointer enters the
    ''' checkbox), so it always reflects the current durable-store contents rather than a stale
    ''' snapshot. This is cheaper and simpler than eagerly refreshing on every state change.
    ''' </summary>
    Private Sub OnPersistKnowledgeTooltipRefresh(sender As Object, e As EventArgs)
        UpdatePersistKnowledgeTooltip()
    End Sub

    ''' <summary>
    ''' Updates the tooltip for the persist knowledge checkbox to point at the durable storage
    ''' directory and summarize what is currently persisted there (knowledge file and indexes).
    ''' </summary>
    Private Sub UpdatePersistKnowledgeTooltip()
        If Not _chkPersistKnowledge.Checked Then
            _toolTip.SetToolTip(_chkPersistKnowledge, "")
            Return
        End If

        Dim storageDir As String = GetRedInkStorageDirectoryPath()
        Dim sb As New StringBuilder()
        sb.Append("Persisted knowledge is stored in: ")
        sb.Append(storageDir)

        Dim items As New List(Of String)()

        Try
            Dim knowledgePath As String = GetPersistedKnowledgeFilePath()
            If File.Exists(knowledgePath) Then
                items.Add($"knowledge file ({New FileInfo(knowledgePath).Length:N0} bytes)")
            End If
        Catch
        End Try

        Try
            Dim sessionDir As String = GetSessionIndexDirectoryPath()
            If Directory.Exists(sessionDir) Then
                Dim indexCount As Integer = Directory.GetFiles(sessionDir, "*" & IndexCopyFileExtension).Length
                If indexCount > 0 Then
                    items.Add($"{indexCount:N0} index file(s)")
                End If
            End If
        Catch
        End Try

        If items.Count > 0 Then
            sb.Append(". Currently persisted: ")
            sb.Append(String.Join(", ", items))
            sb.Append(".")
        Else
            sb.Append(". Nothing is persisted yet.")
        End If

        _toolTip.SetToolTip(_chkPersistKnowledge, sb.ToString())
    End Sub

    ''' <summary>
    ''' Automatically enables temporary knowledge persistence when loaded knowledge is large enough.
    ''' </summary>
    ''' <param name="loadedFileCount">Number of knowledge files loaded for the current operation.</param>
    ''' <returns>True if knowledge was persisted or persistence was already enabled; otherwise False.</returns>
    Private Function AutoEnablePersistKnowledgeIfLarge(loadedFileCount As Integer) As Boolean
        If String.IsNullOrWhiteSpace(_knowledgeContent) Then Return False

        If _knowledgeContent.Length <= AutoPersistKnowledgeThresholdChars Then
            Return _chkPersistKnowledge.Checked
        End If

        Try
            Dim persistPath As String = GetPersistedKnowledgeFilePath()

            System.IO.File.WriteAllText(persistPath, _knowledgeContent, System.Text.Encoding.UTF8)

            _isUpdatingPersistCheckbox = True
            _chkPersistKnowledge.Checked = True
            _isUpdatingPersistCheckbox = False

            My.Settings.DiscussPersistKnowledge = True
            My.Settings.Save()

            PersistAttachedIndexes()

            UpdatePersistKnowledgeTooltip()

            ShowCustomMessageBox(
            $"Knowledge is large ({_knowledgeContent.Length:N0} characters, threshold {AutoPersistKnowledgeThresholdChars:N0})." &
            vbCrLf & vbCrLf &
            "Temporary knowledge persistence has been turned on automatically for this session." &
            vbCrLf & vbCrLf &
            $"Stored in: {persistPath}")

            AppendSystemMessage($"Knowledge loaded and persisted automatically ({_knowledgeContent.Length:N0} characters from {loadedFileCount} file(s)).")
            Return True

        Catch ex As System.Exception
            _isUpdatingPersistCheckbox = True
            _chkPersistKnowledge.Checked = False
            _isUpdatingPersistCheckbox = False

            Try
                My.Settings.DiscussPersistKnowledge = False
                My.Settings.Save()
            Catch
            End Try

            UpdatePersistKnowledgeTooltip()
            AppendSystemMessage($"Knowledge loaded ({_knowledgeContent.Length:N0} characters) but failed to auto-persist: {ex.Message}")
            Return False
        End Try
    End Function

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
            SharedMethods.EnsureVisibleOnScreen(Me)
        Catch
        End Try

        ' Set input panel to double the original designer height (63px × 2 = 126px)
        Try
            Dim desiredInputHeight As Integer = 126
            Dim newDistance As Integer = _splitChat.Height - desiredInputHeight - _splitChat.SplitterWidth
            If newDistance >= _splitChat.Panel1MinSize Then
                _splitChat.SplitterDistance = newDistance
            End If
        Catch
            ' Layout not ready yet; keep default SplitterDistance
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
        _chkAdvancedTools.Checked = Globals.ThisAddIn.GetDiscussInkyAdvancedToolsEnabled()

        ' Tooling log checkbox reflects the effective local override when present, otherwise the INI default.
        SyncToolingLogPreferenceFromSettings()

        ' Load personas
        LoadPersonas()

        ' Load missions
        LoadMissions()

        ' Update tooling controls based on current model
        UpdateToolingControlsState()

        ' Check if persona was previously saved - if not, use default
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

        ' If no persona was restored from settings, apply the default persona
        If Not personaRestoredFromSettings Then
            _currentPersonaName = DefaultPersonaName
            _currentPersonaPrompt = DefaultPersonaPrompt
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

        ' Restore the running chat only from the normal last-chat storage.
        ' Do NOT call RestoreSessionStateFromXml here: that is reserved for explicit archive/session restore.
        Dim hasChat As Boolean = False
        Dim restoredHtmlHadAlternateModel As Boolean = False

        Try
            ' First, restore _history from plain transcript. This keeps the LLM context intact.
            Dim savedTranscript As String = My.Settings.DiscussLastChat
            If Not String.IsNullOrEmpty(savedTranscript) Then
                RestoreHistoryFromTranscript(savedTranscript)
            End If

            ' Then restore the visible HTML transcript.
            Dim savedHtml As String = My.Settings.DiscussLastChatHtml
            If Not String.IsNullOrEmpty(savedHtml) Then
                restoredHtmlHadAlternateModel = ChatHtmlIndicatesAlternateModel(savedHtml)
                AppendHtml(savedHtml)
                hasChat = True
            ElseIf Not String.IsNullOrEmpty(savedTranscript) Then
                AppendTranscriptToHtml(savedTranscript)
                hasChat = True
            End If

        Catch ex As System.Exception
            AppendSystemMessage($"Failed to restore previous chat: {ex.Message}")
        End Try

        If hasChat AndAlso restoredHtmlHadAlternateModel Then
            _alternateModelSelected = False
            _alternateModelConfig = Nothing
            _alternateModelDisplayName = Nothing
            UpdateAlternateModelButtonText()
            AppendSystemMessage($"Previous chat restored. Now using primary model ({_context.INI_Model}).")
        End If

        Await RestoreKnowledgeAsync()

        ' Restore attached semantic indexes for a non-archived running session (durable, crash-safe),
        ' then remove orphaned index folders left by sessions that are no longer referenced.
        RestoreSessionIndexStateFromDurableFile()
        SweepOrphanSessionIndexes()

        ' Make persistence visible on re-entry: when persistence is on, confirm that the loaded
        ' knowledge and any attached indexes are being retained in durable storage.
        If _chkPersistKnowledge.Checked AndAlso
           (Not String.IsNullOrWhiteSpace(_knowledgeContent) OrElse _attachedIndexes.Count > 0) Then
            AppendSystemMessage(
                $"Knowledge persistence is on. {If(_attachedIndexes.Count > 0, $"{_attachedIndexes.Count:N0} attached index(es) and ", "")}loaded knowledge are retained in durable storage.")
        End If

        ' Only force persona selection if there are custom personas beyond the default
        ' (i.e., a persona library is configured and has entries)
        If Not personaRestoredFromSettings AndAlso _personas.Count > 1 AndAlso Not _personaSelectedThisSession Then
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
    ''' Deletes orphaned per-session index folders under %AppData%\redink\di. A folder is kept
    ''' when it is the current session id or when any currently attached index still points into
    ''' it; everything else is removed to avoid accumulating copies of potentially sensitive data.
    ''' </summary>
    Private Sub SweepOrphanSessionIndexes()
        Dim rootPath As String = GetSessionIndexRootPath()
        If Not Directory.Exists(rootPath) Then
            Return
        End If

        Dim currentId As String = GetOrCreateSessionIndexId()

        Dim referencedDirs As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each idx In _attachedIndexes
            Dim indexRef As DiscussIndexRef = idx
            If Not String.IsNullOrWhiteSpace(indexRef.ActivePath) Then
                Dim dir As String = Path.GetDirectoryName(indexRef.ActivePath)
                If Not String.IsNullOrWhiteSpace(dir) Then
                    referencedDirs.Add(dir)
                End If
            End If
        Next

        Try
            For Each subDir In Directory.GetDirectories(rootPath)
                Dim folderName As String = Path.GetFileName(subDir)

                If String.Equals(folderName, currentId, StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If
                If referencedDirs.Contains(subDir) Then
                    Continue For
                End If

                DeleteIndexDirectorySafe(subDir)
            Next
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try

        SweepOrphanArchiveIndexes()
    End Sub

    ''' <summary>
    ''' Deletes archive index sidecar folders (&lt;name&gt;.ix) whose matching dialogue archive
    ''' (&lt;name&gt;.dialogue.xml) no longer exists, e.g. after an archive was deleted outside the app.
    ''' </summary>
    Private Sub SweepOrphanArchiveIndexes()
        Dim archiveDir As String = GetDialogueArchiveDirectoryPath()
        If Not Directory.Exists(archiveDir) Then
            Return
        End If

        Try
            For Each sidecarDir In Directory.GetDirectories(archiveDir, "*" & ArchiveIndexFolderSuffix, SearchOption.TopDirectoryOnly)
                Dim folderName As String = Path.GetFileName(sidecarDir)
                If Not folderName.EndsWith(ArchiveIndexFolderSuffix, StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If

                Dim baseName As String = folderName.Substring(0, folderName.Length - ArchiveIndexFolderSuffix.Length)
                Dim archiveFilePath As String = Path.Combine(archiveDir, baseName & DialogueArchiveFileExtension)

                If Not File.Exists(archiveFilePath) Then
                    DeleteIndexDirectorySafe(sidecarDir)
                End If
            Next
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Restores knowledge from various sources in priority order:
    ''' 1. Runtime cache (if Word hasn't been restarted)
    ''' 2. Persisted temp file (if checkbox is checked)
    ''' 3. Previously saved file or directory path from settings
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
                    AppendSystemMessage($"Knowledge restored from persisted storage ({GetKnowledgeSummaryText()}).")
                    Return
                Catch ex As Exception
                    AppendSystemMessage($"Failed to restore persisted knowledge: {ex.Message}")
                End Try
            End If
        End If

        ' 3. Try to reload from saved file or directory path in settings
        Dim savedPath As String = ""
        Try
            savedPath = My.Settings.DiscussKnowledgePath
        Catch
            Return
        End Try

        If String.IsNullOrEmpty(savedPath) Then Return

        Dim isFile = File.Exists(savedPath)
        Dim isDirectory = Directory.Exists(savedPath)

        If Not isFile AndAlso Not isDirectory Then
            ' Path no longer exists - clear it from settings
            Try
                My.Settings.DiscussKnowledgePath = ""
                My.Settings.Save()
            Catch
            End Try
            Return
        End If

        Try
            ShowAssistantThinking()

            If isFile Then
                ' Single file - use existing logic
                Dim result = Await LoadSingleKnowledgeFileAsync(savedPath, False, False, askWorksheetSelection:=True)
                _knowledgeContent = result.Content
                _knowledgeFilePath = savedPath

                RemoveAssistantThinking()

                If Not String.IsNullOrWhiteSpace(_knowledgeContent) Then
                    AppendSystemMessage($"Knowledge restored from file: {Path.GetFileName(savedPath)} ({_knowledgeContent.Length:N0} characters).")
                End If
            Else
                ' Directory - reload all supported files
                Dim ctx As New KnowledgeLoadingContext()
                Dim filesToProcess As New List(Of String)()

                Dim allFiles = Directory.GetFiles(savedPath, "*.*", SearchOption.TopDirectoryOnly)
                For Each f In allFiles
                    Dim ext = Path.GetExtension(f).ToLowerInvariant()
                    If SupportedKnowledgeExtensions.Contains(ext) Then
                        filesToProcess.Add(f)
                        If ext = ".pdf" Then
                            ctx.HasPdfFiles = True
                        End If
                    End If
                Next

                ' Apply same limits as initial load
                If filesToProcess.Count > KnowledgeLoadingContext.MaxFilesPerDirectory Then
                    filesToProcess = filesToProcess.Take(KnowledgeLoadingContext.MaxFilesPerDirectory).ToList()
                End If

                If filesToProcess.Count = 0 Then
                    RemoveAssistantThinking()
                    AppendSystemMessage($"No supported files found in previously saved directory: {savedPath}")
                    Return
                End If

                ' Load all files
                Dim resultBuilder As New StringBuilder()
                Dim useDocumentTags = (filesToProcess.Count > 1)
                Dim loadedCount = 0
                For Each filePath In filesToProcess
                    Try
                        UpdateWorkingIndicator($"Reading '{System.IO.Path.GetFileName(filePath)}' ...")

                        Dim askWorksheetSelection As Boolean =
                        isFile AndAlso
                        filesToProcess.Count = 1 AndAlso
                        Path.GetExtension(filePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)

                        Dim result = Await LoadSingleKnowledgeFileAsync(
                        filePath,
                        ctx.EnableOCR,
                        silent:=Not askWorksheetSelection,
                        askWorksheetSelection:=askWorksheetSelection)

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
                        loadedCount += 1

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

                If loadedCount > 0 Then
                    _knowledgeContent = resultBuilder.ToString()
                    _knowledgeFilePath = savedPath & " (directory)"
                    AppendSystemMessage($"Knowledge restored from directory: {loadedCount} file(s), {_knowledgeContent.Length:N0} characters.")
                Else
                    AppendSystemMessage($"Failed to load any files from directory: {savedPath}")
                    Return
                End If
            End If

            If Not String.IsNullOrWhiteSpace(_knowledgeContent) Then
                ' Update runtime cache
                _cachedKnowledgeContent = _knowledgeContent
                _cachedKnowledgeFilePath = _knowledgeFilePath

                ' Persist if checkbox is checked
                If _chkPersistKnowledge.Checked Then
                    PersistKnowledgeToTempFile()
                End If

                UpdateWindowTitle()
            End If

        Catch ex As Exception
            RemoveAssistantThinking()
            AppendSystemMessage($"Error restoring knowledge: {ex.Message}")
        End Try
    End Function

    ''' <summary>
    ''' Persists the current knowledge content to the temp file.
    ''' </summary>
    Private Sub PersistKnowledgeToTempFile()
        If String.IsNullOrWhiteSpace(_knowledgeContent) AndAlso _attachedIndexes.Count = 0 Then Return

        Try
            If Not String.IsNullOrWhiteSpace(_knowledgeContent) Then
                Dim persistPath As String = GetPersistedKnowledgeFilePath()
                System.IO.File.WriteAllText(persistPath, _knowledgeContent, System.Text.Encoding.UTF8)
            End If
        Catch
            ' Silently fail - not critical
        End Try

        If _chkPersistKnowledge.Checked Then
            PersistAttachedIndexes()
        End If
    End Sub

    ''' <summary>
    ''' Repositions the form after monitor/resolution changes.
    ''' </summary>
    Private Sub OnDisplaySettingsChanged(sender As Object, e As EventArgs)
        If Me.IsDisposed Then Return

        Try
            If Me.InvokeRequired Then
                Me.BeginInvoke(New MethodInvoker(
                    Sub()
                        If Not Me.IsDisposed Then SharedMethods.EnsureVisibleOnScreen(Me)
                    End Sub))
            Else
                SharedMethods.EnsureVisibleOnScreen(Me)
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Persists geometry, transcript, persona, mission, knowledge path, and checkbox state on close.
    ''' </summary>
    Private Sub OnFormClosing(sender As Object, e As FormClosingEventArgs)
        Try
            If e.CloseReason = CloseReason.UserClosing AndAlso
               Not ConfirmCloseWhenKnowledgePersisted() Then
                e.Cancel = True
                Return
            End If

            Dim scope = _ownerScope
            _ownerScope = Nothing
            If scope IsNot Nothing Then
                Try : scope.Dispose() : Catch : End Try
            End If
            PersistCurrentSessionSettings(saveImmediately:=False)
            Try
                RemoveHandler Microsoft.Win32.SystemEvents.DisplaySettingsChanged, AddressOf OnDisplaySettingsChanged
            Catch
            End Try
            If Me.WindowState = FormWindowState.Normal Then
                My.Settings.DiscussFormLocation = Me.Location
                My.Settings.DiscussFormSize = Me.Size
            Else
                My.Settings.DiscussFormLocation = Me.RestoreBounds.Location
                My.Settings.DiscussFormSize = Me.RestoreBounds.Size
            End If

            Globals.ThisAddIn.PersistDiscussInkyToolSelection(
                Globals.ThisAddIn.SplitPersistedToolNames(CStr(My.Settings("SelectedMainToolNames"))),
                Globals.ThisAddIn.SplitPersistedToolNames(CStr(My.Settings("SelectedAdvancedToolNames"))),
                _chkAdvancedTools.Checked)

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
        ' Capture UI state before leaving the UI thread
        Dim hideLog As Boolean = Not _chkShowToolingLog.Checked
        Dim shouldUseTool As Boolean = ShouldUseTooling()
        Dim toolsReady As Boolean = If(shouldUseTool, EnsureToolsSelected(), False)

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
            If shouldUseTool AndAlso toolsReady Then
                ' Execute via tooling loop
                Return Await Globals.ThisAddIn.ExecuteToolingLoop(
                    systemPrompt,
                    "",
                    _selectedToolsForChat,
                    useSecondApi,
                    fullPromptOverride:=userPrompt,
                    hideSplash:=True,
                    hideLogWindow:=hideLog,
                    progressSink:=Sub(status) UpdateAssistantThinking(status)).ConfigureAwait(False)
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
    ''' Updates enabled state of tooling controls based on current model support and "(t)" availability.
    ''' </summary>
    Public Sub SyncToolingLogPreferenceFromSettings()
        If Me.IsDisposed Then
            Return
        End If

        Dim effectiveSetting As Boolean = Globals.ThisAddIn.GetEffectiveToolingLogWindowSetting()

        If _chkShowToolingLog.Checked = effectiveSetting Then
            Return
        End If

        _suppressToolingLogPreferenceSync = True

        Try
            _chkShowToolingLog.Checked = effectiveSetting
        Finally
            _suppressToolingLogPreferenceSync = False
        End Try
    End Sub

    Private Sub UpdateToolingControlsState()
        Dim currentConfig As ModelConfig = Nothing

        If _alternateModelSelected AndAlso _alternateModelConfig IsNot Nothing Then
            currentConfig = _alternateModelConfig
        Else
            currentConfig = SharedMethods.GetCurrentConfig(_context)
        End If

        Dim supportsCurrentModelTooling As Boolean = SharedMethods.ModelSupportsTooling(currentConfig)
        Dim supportsToolTrigger As Boolean =
            SharedMethods.HasToolingCapableSpecialTaskModel(_context, _context.INI_AlternateModelPath, "ToolDefaultModel")

        Dim toolingUiAvailable As Boolean = supportsCurrentModelTooling OrElse supportsToolTrigger

        _chkEnableTooling.Enabled = toolingUiAvailable
        _btnTools.Enabled = toolingUiAvailable
        _chkAdvancedTools.Enabled = toolingUiAvailable AndAlso _chkEnableTooling.Checked
        _chkShowToolingLog.Enabled = toolingUiAvailable

        If Not toolingUiAvailable Then
            _chkEnableTooling.Checked = False
            _selectedToolsForChat = Nothing
        End If

        If Not _toolingControlsInitialized Then
            SyncToolingLogPreferenceFromSettings()
            _toolingControlsInitialized = True
        End If
    End Sub

    ''' <summary>
    ''' Handles changes to the "Tooling log" checkbox. The checked state is consumed when executing the tooling loop
    ''' to decide whether to show or hide the tooling log window.
    ''' </summary>
    ''' <param name="sender">The event source.</param>
    ''' <param name="e">Event arguments.</param>

    Private Sub OnShowToolingLogChanged(sender As Object, e As EventArgs)
        If _suppressToolingLogPreferenceSync Then
            Return
        End If

        Globals.ThisAddIn.SetToolingLogWindowOverride(_chkShowToolingLog.Checked)
        Globals.ThisAddIn.RefreshOpenToolingLogPreferenceWindows()
    End Sub

    ''' <summary>
    ''' Handles the Inky Memory checkbox change. Persists preference and toggles edit link.
    ''' </summary>
    Private Sub OnInkyMemoryChanged(sender As Object, e As EventArgs)
        My.Settings.DiscussInkyMemory = _chkInkyMemory.Checked
        My.Settings.Save()
        _lnkEditMemory.Visible = _chkInkyMemory.Checked
    End Sub

    ''' <summary>
    ''' Opens the Inky Memory file for manual editing.
    ''' </summary>
    Private Sub OnEditMemoryClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        SharedMethods.EditInkyMemoryFile()
    End Sub

    ''' <summary>
    ''' Handles the Tools button click - opens tool selection dialog.
    ''' </summary>
    Private Sub OnToolsClick(sender As Object, e As EventArgs)
        Try
            Dim selectedTools = Globals.ThisAddIn.SelectDiscussInkyToolsForSession(forceDialog:=True)

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
        Try
            My.Settings.DiscussEnableTooling = _chkEnableTooling.Checked
            My.Settings.Save()
        Catch
        End Try

        _selectedToolsForChat = Nothing
        _chkAdvancedTools.Enabled = _chkEnableTooling.Checked
        UpdateToolingControlsState()
    End Sub

    Private Sub OnAdvancedToolsChanged(sender As Object, e As EventArgs)
        Try
            Globals.ThisAddIn.PersistDiscussInkyToolSelection(
                Globals.ThisAddIn.GetDiscussInkyEffectiveMainToolNames(),
                Globals.ThisAddIn.GetDiscussInkyEffectiveAdvancedToolNames(),
                _chkAdvancedTools.Checked)
        Catch
        End Try

        _selectedToolsForChat = Nothing
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

        _selectedToolsForChat = Globals.ThisAddIn.SelectDiscussInkyToolsForSession(forceDialog:=False)
        Return _selectedToolsForChat IsNot Nothing AndAlso _selectedToolsForChat.Count > 0
    End Function


#End Region

#Region "Persona Management"

    ''' <summary>
    ''' Loads persona definitions from configured local and global files into memory.
    ''' Always ensures at least the default fallback persona is available.
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

        ' Always ensure the default fallback persona is available
        ' Add it at the beginning so it's always the first option
        Dim defaultDisplay = MakeUniqueDisplay(DefaultPersonaName, _personas.Select(Function(p) p.DisplayName).ToList())
        _personas.Insert(0, New PersonaEntry With {
            .Name = DefaultPersonaName,
            .Prompt = DefaultPersonaPrompt,
            .IsLocal = False,
            .DisplayName = defaultDisplay
        })

        ' Track whether persona library is configured (message shown later in ShowSessionInfo
        ' after the HTML chat is initialized)
        _noPersonaLibraryConfigured = String.IsNullOrWhiteSpace(localPath) AndAlso String.IsNullOrWhiteSpace(globalPath)
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
            ' Should not happen since we always have the default, but guard anyway
            _currentPersonaName = DefaultPersonaName
            _currentPersonaPrompt = DefaultPersonaPrompt
            UpdateWindowTitle()
            UpdateSendButtonText()
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
    ''' Reloads personas after editing if the file was modified.
    ''' </summary>
    Private Sub OnEditLocalPersona(sender As Object, e As EventArgs)
        Dim localPath = ExpandEnvironmentVariables(If(_context?.INI_DiscussInkyPathLocal, ""))

        If String.IsNullOrWhiteSpace(localPath) Then
            ShowCustomMessageBox("'DiscussInkyPathLocal' is not configured in your settings." & vbCrLf & vbCrLf &
                                 "To create a local persona library, configure this path in your configuration file. " &
                                 "Sample files are available via 'Get Sample Files' in the settings menu.")
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

        ' Capture file hash before editing for reliable change detection
        Dim hashBefore As String = GetFileHash(localPath)

        ' ShowTextFileEditor is expected to be synchronous (modal dialog)
        ShowTextFileEditor(localPath, $"{AN} - Edit Local Personas:", False, _context)

        ' Check if file content actually changed (hash comparison is more reliable than timestamp)
        Dim hashAfter As String = GetFileHash(localPath)

        If Not String.Equals(hashBefore, hashAfter, StringComparison.Ordinal) Then
            LoadPersonas()
            UpdateWindowTitle()
            UpdateSendButtonText()
            AppendSystemMessage("Local personas reloaded.")
        End If
    End Sub

    ''' <summary>
    ''' Computes a simple hash of file contents for change detection.
    ''' Returns empty string if file doesn't exist or can't be read.
    ''' </summary>
    Private Shared Function GetFileHash(filePath As String) As String
        Try
            If Not File.Exists(filePath) Then Return ""
            Dim bytes = File.ReadAllBytes(filePath)
            Using sha = System.Security.Cryptography.SHA256.Create()
                Dim hash = sha.ComputeHash(bytes)
                Return System.Convert.ToBase64String(hash)
            End Using
        Catch
            Return ""
        End Try
    End Function

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
            ShowCustomMessageBox("No persona library is configured. Missions require a persona library path ('DiscussInkyPathLocal' or 'DiscussInkyPath').")
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

    Private Function GetKnowledgeDocumentCount(Optional content As String = Nothing) As Integer
        Dim source As String = If(content, _knowledgeContent)

        If String.IsNullOrWhiteSpace(source) Then
            Return 0
        End If

        Dim matches As System.Text.RegularExpressions.MatchCollection =
            System.Text.RegularExpressions.Regex.Matches(
                source,
                "<document\d+\b",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase)

        If matches.Count > 0 Then
            Return matches.Count
        End If

        Return 1
    End Function

    Private Function GetNextKnowledgeDocumentNumber(existingContent As String) As Integer
        If String.IsNullOrWhiteSpace(existingContent) Then
            Return 1
        End If

        Dim maxNumber As Integer = 0

        Dim matches As System.Text.RegularExpressions.MatchCollection =
            System.Text.RegularExpressions.Regex.Matches(
                existingContent,
                "<document(?<n>\d+)\b",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase)

        For Each m As System.Text.RegularExpressions.Match In matches
            Dim n As Integer = 0
            If Integer.TryParse(m.Groups("n").Value, n) Then
                If n > maxNumber Then
                    maxNumber = n
                End If
            End If
        Next

        Return maxNumber + 1
    End Function

    Private Function GetKnowledgeSummaryText(Optional content As String = Nothing) As String
        Dim source As String = If(content, _knowledgeContent)

        If String.IsNullOrWhiteSpace(source) Then
            Return "0 item(s), 0 characters"
        End If

        Return $"{GetKnowledgeDocumentCount(source):N0} item(s), {source.Length:N0} characters"
    End Function

    Private Function GetKnowledgeDisplayName() As String
        If String.IsNullOrWhiteSpace(_knowledgeFilePath) Then
            Return "None loaded"
        End If

        Return $"{System.IO.Path.GetFileName(_knowledgeFilePath)} ({GetKnowledgeSummaryText()})"
    End Function

    Private Function GetKnowledgePathLabelAfterLoad(selectedPath As String, isFile As Boolean, appendToExisting As Boolean) As String
        If appendToExisting Then
            Return "(Combined Knowledge)"
        End If

        If isFile Then
            Return selectedPath
        End If

        Return selectedPath & " (directory)"
    End Function


    ''' <summary>
    ''' Clears only the inlined plain-knowledge content, its persisted file, and the saved path.
    ''' Attached semantic indexes are a separate channel and are left untouched. Used by internal
    ''' mutations that empty the plain content without discarding the whole knowledge.
    ''' </summary>
    Private Sub ClearPlainKnowledge()
        _knowledgeContent = Nothing
        _knowledgeFilePath = Nothing
        _cachedKnowledgeContent = Nothing
        _cachedKnowledgeFilePath = Nothing

        Try
            Dim persistPath = GetPersistedKnowledgeFilePath()
            If File.Exists(persistPath) Then
                File.Delete(persistPath)
            End If
        Catch
        End Try

        Try
            My.Settings.DiscussKnowledgePath = ""
            My.Settings.Save()
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Detaches all attached semantic indexes from the session and deletes their durable copies
    ''' (under di\&lt;sid&gt;). Index files referenced in place from an external location are left
    ''' on disk; only copies this session created are removed.
    ''' </summary>
    Private Sub DetachAllAttachedIndexes()
        _attachedIndexes.Clear()
        _indexConversationState.Clear()
        DeleteIndexDirectorySafe(GetSessionIndexDirectoryPath())
        ' A new knowledge set may be loaded next; allow the persistence prompt to be offered again.
        _indexPersistenceOffered = False
    End Sub

    ''' <summary>
    ''' User-facing knowledge deletion. Because attached indexes are part of "Knowledge", this
    ''' also detaches them and removes their durable copies, then clears the plain content.
    ''' </summary>
    Private Sub DeleteCurrentKnowledge()
        DetachAllAttachedIndexes()
        ClearPlainKnowledge()

        UpdateWindowTitle()
        PersistCurrentSessionSettings()
        AppendSystemMessage("Knowledge deleted (including any attached indexes).")
    End Sub

    ''' <summary>
    ''' Attaches an already-indexed text file as a standalone searchable source. The file is
    ''' referenced in place (no copy) until the session is persisted or archived; its exact
    ''' bytes must never be altered so content-relative offsets and the SHA-256 guard stay valid.
    ''' </summary>
    Private Function AttachIndexFromFile(indexPath As String) As Boolean
        If String.IsNullOrWhiteSpace(indexPath) OrElse Not File.Exists(indexPath) Then
            AppendSystemMessage("The selected index file does not exist.")
            Return False
        End If

        ' Refuse loading the same index twice: compare exact-content SHA-256 against every
        ' currently attached index. Identical content means the same knowledge source.
        Dim newHash As String = ComputeFileSha256(indexPath)
        If Not String.IsNullOrEmpty(newHash) Then
            For Each existing In _attachedIndexes
                Dim existingRef As DiscussIndexRef = existing
                Dim existingHash As String = ""
                If Not String.IsNullOrWhiteSpace(existingRef.ActivePath) AndAlso File.Exists(existingRef.ActivePath) Then
                    existingHash = ComputeFileSha256(existingRef.ActivePath)
                End If

                If String.Equals(existingHash, newHash, StringComparison.OrdinalIgnoreCase) Then
                    AppendSystemMessage(
                        $"The index '{System.IO.Path.GetFileName(indexPath)}' is already attached (identical content). It was not added again.")
                    Return False
                End If
            Next
        End If

        Dim indexRef As New DiscussIndexRef() With {
            .Id = "i" & Guid.NewGuid().ToString("N").Substring(0, 4),
            .DisplayName = System.IO.Path.GetFileName(indexPath),
            .ActivePath = indexPath,
            .OriginalPath = indexPath,
            .ContentSha256 = newHash
        }

        _attachedIndexes.Add(indexRef)

        ' An index is just another knowledge source: if persistence is already on, copy the newly
        ' attached index into durable storage immediately so it is retained alongside other knowledge.
        ' If persistence is off, offer to turn it on so the index survives temp cleanup and restarts.
        If _chkPersistKnowledge.Checked Then
            PersistAttachedIndexes()
        Else
            OfferPersistenceForAttachedIndex(indexRef.DisplayName)
        End If

        AppendSystemMessage(
            $"Attached semantic index '{indexRef.DisplayName}'. It will be searched for each message alongside any loaded knowledge.")
        UpdateWindowTitle()
        PersistCurrentSessionSettings()
        Return True
    End Function

    ''' <summary>
    ''' Prompts the user to enable durable persistence after an index is attached while persistence
    ''' is off. When accepted, turns the persist checkbox on (which copies the current knowledge and
    ''' all attached indexes to durable storage). No-op when the user declines.
    ''' </summary>
    Private Sub OfferPersistenceForAttachedIndex(indexDisplayName As String)
        ' Ask at most once per knowledge set; loading multiple files/indexes must not re-prompt.
        If _indexPersistenceOffered Then
            Return
        End If
        _indexPersistenceOffered = True

        Dim answer = ShowCustomYesNoBox(
            $"The searchable index '{indexDisplayName}' is currently referenced from its original location only. " &
            "Do you want to persist it (and any loaded knowledge) to durable storage so it is retained across restarts?",
            "Yes, persist",
            "No, keep temporary",
            $"{AN} - Persist Index")

        If answer <> 1 Then
            Return
        End If

        ' Setting Checked fires OnPersistKnowledgeChanged, which persists knowledge and indexes.
        _chkPersistKnowledge.Checked = True
    End Sub

    ''' <summary>
    ''' Informs the user that a newly created index lives only inside DiscussInky's durable storage
    ''' and offers to save a reusable copy (ending in '.index.txt') to the Desktop for later use.
    ''' </summary>
    Private Sub OfferDesktopIndexCopy(sourceIndexPath As String, baseName As String)
        Try
            If String.IsNullOrWhiteSpace(sourceIndexPath) OrElse Not File.Exists(sourceIndexPath) Then
                Return
            End If

            Dim answer = ShowCustomYesNoBox(
                "The searchable index has been created, but it lives only inside DiscussInky's durable storage " &
                "(it is retained with this session and its archives, not as a normal file you can reuse elsewhere)." &
                vbCrLf & vbCrLf &
                "Do you also want to save a reusable copy (ending in '.index.txt') to your Desktop for later use?",
                "Yes, save a copy",
                "No, thanks",
                $"{AN} - Save Index Copy")

            If answer <> 1 Then
                Return
            End If

            Dim safeBase As String = Path.GetFileNameWithoutExtension(If(baseName, "index"))
            For Each ch In Path.GetInvalidFileNameChars()
                safeBase = safeBase.Replace(ch, "_"c)
            Next
            If String.IsNullOrWhiteSpace(safeBase) Then
                safeBase = "index"
            End If

            Dim desktop As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            Dim destination As String = Path.Combine(desktop, safeBase & ".index.txt")

            Dim counter As Integer = 1
            While File.Exists(destination)
                destination = Path.Combine(desktop, $"{safeBase} ({counter}).index.txt")
                counter += 1
            End While

            ' Exact-byte copy so content offsets and the SHA-256 guard stay valid in the copy.
            File.Copy(sourceIndexPath, destination, overwrite:=False)
            AppendSystemMessage($"A reusable copy of the index was saved to: {destination}")
        Catch ex As Exception
            AppendSystemMessage($"Could not save a Desktop copy of the index: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Returns the most recent user message as the retrieval query, falling back to the last
    ''' message when no user turn is available (used by automated response modes).
    ''' </summary>
    Private Function GetRetrievalQueryFromHistory() As String
        For i As Integer = _history.Count - 1 To 0 Step -1
            If String.Equals(_history(i).Role, "user", StringComparison.OrdinalIgnoreCase) Then
                Return _history(i).Content
            End If
        Next
        If _history.Count > 0 Then
            Return _history(_history.Count - 1).Content
        End If
        Return ""
    End Function

    ''' <summary>
    ''' Queries each attached semantic index for the given message and returns the merged,
    ''' most-relevant original excerpts (wrapped per source for citation). Updates per-index
    ''' conversation state so follow-up turns can reuse previously selected segments.
    ''' </summary>
    Private Async Function BuildIndexExcerptsAsync(queryText As String,
                                                   conversation As String,
                                                   Optional reportStatus As System.Action(Of String) = Nothing) As Task(Of String)
        If _attachedIndexes.Count = 0 OrElse String.IsNullOrWhiteSpace(queryText) Then
            Return ""
        End If

        Dim sb As New StringBuilder()

        ' Snapshot the collection so a concurrent change cannot break enumeration, and to give
        ' stable index positions in the user-facing progress messages.
        Dim snapshot As List(Of DiscussIndexRef) = _attachedIndexes.ToList()
        Dim total As Integer = snapshot.Count
        Dim position As Integer = 0
        Dim indexesWithHits As Integer = 0
        Dim totalSegments As Integer = 0

        For Each indexRef In snapshot
            position += 1

            If String.IsNullOrWhiteSpace(indexRef.ActivePath) OrElse Not File.Exists(indexRef.ActivePath) Then
                reportStatus?.Invoke($"Index {position} of {total}: '{indexRef.DisplayName}' is unavailable and was skipped.")
                Continue For
            End If

            reportStatus?.Invoke($"Searching index {position} of {total}: '{indexRef.DisplayName}' ...")

            Dim previousIds As List(Of String) = Nothing
            _indexConversationState.TryGetValue(indexRef.Id, previousIds)

            Try
                Dim options As New SharedMethods.SemanticSearchRetrievalOptions() With {
                    .SpecialTaskName = "Indexer"
                }

                Dim retrieval As SharedMethods.SemanticSearchRetrievalResult =
                    Await SharedMethods.RetrieveSemanticSearchAsync(
                        indexRef.ActivePath,
                        _context,
                        queryText,
                        If(conversation, ""),
                        previousIds,
                        options).ConfigureAwait(False)

                Dim matchCount As Integer =
                    If(retrieval IsNot Nothing AndAlso retrieval.SelectedEntryIds IsNot Nothing,
                       retrieval.SelectedEntryIds.Count, 0)

                If retrieval IsNot Nothing AndAlso
                   retrieval.IsIndexed AndAlso
                   Not String.IsNullOrWhiteSpace(retrieval.ReducedSourceText) Then

                    sb.AppendLine($"<document name=""{indexRef.DisplayName}"">")
                    sb.AppendLine(retrieval.ReducedSourceText)
                    sb.AppendLine("</document>")
                    sb.AppendLine()

                    indexesWithHits += 1
                    totalSegments += matchCount

                    If retrieval.SelectedEntryIds IsNot Nothing AndAlso retrieval.SelectedEntryIds.Count > 0 Then
                        _indexConversationState(indexRef.Id) = New List(Of String)(retrieval.SelectedEntryIds)
                    End If

                    reportStatus?.Invoke($"Index {position} of {total}: '{indexRef.DisplayName}' — {matchCount:N0} relevant segment(s) found.")
                Else
                    reportStatus?.Invoke($"Index {position} of {total}: '{indexRef.DisplayName}' — no relevant material found.")
                End If
            Catch ex As Exception
                AppendSystemMessage($"Index retrieval failed for '{indexRef.DisplayName}': {ex.Message}")
            End Try
        Next

        reportStatus?.Invoke($"Index search complete: {totalSegments:N0} relevant segment(s) from {indexesWithHits:N0} of {total:N0} index(es).")

        Return sb.ToString().TrimEnd()
    End Function

    ''' <summary>
    ''' Button handler that launches the knowledge file/directory picker.
    ''' </summary>
    Private Async Sub OnLoadKnowledge(sender As Object, e As EventArgs)
        If Not TryBeginExclusive("loading knowledge") Then
            Return
        End If
        Try
            Await PromptForKnowledgeAsync()
        Finally
            EndExclusive()
        End Try
        BringDiscussFormToFront()
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
                If frm.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
                    selectedPath = frm.SelectedFilePath
                End If
            End Using

            Globals.ThisAddIn.DragDropFormLabel = ""
            Globals.ThisAddIn.DragDropFormFilter = ""

            If String.IsNullOrWhiteSpace(selectedPath) Then
                ' No file selected - check if there's existing knowledge or attached indexes to delete
                If HasLoadedKnowledgeOrIndexes() Then
                    Dim answer = ShowCustomYesNoBox(
                        "No file was selected. Do you want to delete the currently loaded knowledge and any attached indexes?",
                        "Yes, delete knowledge",
                        "No, keep it")

                    If answer = 1 Then
                        DeleteCurrentKnowledge()
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

            ' If the selected file is already a semantic-search index, attach it as a standalone
            ' searchable source instead of inlining its bytes into the plain knowledge blob.
            If isFile AndAlso SharedMethods.IsPotentiallySemanticSearchIndexedTextFile(selectedPath) Then
                ' Indexes are another knowledge source, so honor the same add-vs-replace choice.
                If HasLoadedKnowledgeOrIndexes() Then
                    Dim indexAppendAnswer = ShowCustomYesNoBox(
                        "There is already knowledge loaded:" &
                        vbCrLf & vbCrLf &
                        GetKnowledgeSummaryText() &
                        If(_attachedIndexes.Count > 0, vbCrLf & $"Attached searchable index(es): {_attachedIndexes.Count:N0}", "") &
                        vbCrLf & vbCrLf &
                        "Do you want to add the selected index to the existing knowledge, or replace the existing knowledge?",
                        "Add to existing", "Replace existing")

                    If indexAppendAnswer = 0 Then
                        ' User cancelled.
                        Return
                    End If

                    If indexAppendAnswer = 2 Then
                        ' Replace: discard existing plain knowledge and detach existing indexes
                        ' (only session copies are removed; archived copies stay intact).
                        DetachAllAttachedIndexes()
                        ClearPlainKnowledge()
                    End If
                End If

                AttachIndexFromFile(selectedPath)
                Try
                    My.Settings.DiscussKnowledgePath = selectedPath
                    My.Settings.Save()
                Catch
                End Try
                BringDiscussFormToFront()
                Return
            End If

            Dim appendToExisting As Boolean = False
            Dim existingKnowledgeSummary As String = GetKnowledgeSummaryText()

            If HasLoadedKnowledgeOrIndexes() Then
                Dim appendAnswer = ShowCustomYesNoBox(
                    "There is already knowledge loaded:" &
                    vbCrLf & vbCrLf &
                    existingKnowledgeSummary &
                    If(_attachedIndexes.Count > 0, vbCrLf & $"Attached searchable index(es): {_attachedIndexes.Count:N0}", "") &
                    vbCrLf & vbCrLf &
                    "Do you want to add the selected knowledge to the existing knowledge, or replace the existing knowledge?",
                    "Add to existing", "Replace existing")
                If appendAnswer = 0 Then
                    ' User cancelled
                    Return
                End If
                appendToExisting = (appendAnswer = 1)

                ' Replacing knowledge also discards attached indexes, because an index is just
                ' another knowledge source. Only session copies are removed; archived index
                ' sidecars are left untouched.
                If Not appendToExisting Then
                    DetachAllAttachedIndexes()
                End If
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
            ShowWorkingIndicator("Reading knowledge document(s) ...")

            Dim resultBuilder As New StringBuilder()
            Dim normalizedExistingKnowledgeContent As String = _knowledgeContent

            If appendToExisting Then
                normalizedExistingKnowledgeContent = PrepareKnowledgeContentForAppending(normalizedExistingKnowledgeContent)
            End If

            Dim useDocumentTags As Boolean = appendToExisting OrElse filesToProcess.Count > 1
            Dim firstDocumentNumber As Integer = If(appendToExisting, GetNextKnowledgeDocumentNumber(normalizedExistingKnowledgeContent), 1)

            For Each filePath In filesToProcess
                Try
                    Dim askWorksheetSelection As Boolean =
                        isFile AndAlso
                        filesToProcess.Count = 1 AndAlso
                        Path.GetExtension(filePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)

                    Dim result = Await LoadSingleKnowledgeFileAsync(
                        filePath,
                        ctx.EnableOCR,
                        silent:=Not askWorksheetSelection,
                        askWorksheetSelection:=askWorksheetSelection)

                    If result.UserCancelled Then
                        RemoveWorkingIndicator()

                        If Not String.IsNullOrWhiteSpace(_knowledgeContent) Then
                            Dim answer = ShowCustomYesNoBox(
                                "No worksheet was selected. Do you want to delete the currently loaded knowledge?",
                                "Yes, delete knowledge", "No, keep it")

                            If answer = 1 Then
                                DeleteCurrentKnowledge()
                            End If
                        End If

                        Return
                    End If

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
                        Dim docNum As Integer = If(appendToExisting, firstDocumentNumber + ctx.GlobalDocumentCounter - 1, ctx.GlobalDocumentCounter)
                        Dim fileName As String = System.IO.Path.GetFileName(filePath)
                        Dim openTag As String = $"<document{docNum} name=""{fileName}"">"
                        Dim closeTag As String = $"</document{docNum}>"

                        resultBuilder.
                            Append(openTag).
                            AppendLine().
                            Append(content).
                            AppendLine().
                            Append(closeTag).
                            AppendLine()
                    Else
                        resultBuilder.Append(content)
                    End If

                Catch ex As Exception
                    ctx.FailedFiles.Add(filePath)
                End Try
            Next

            RemoveWorkingIndicator()

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
            If appendToExisting Then
                _knowledgeContent =
                    If(normalizedExistingKnowledgeContent, "").TrimEnd() &
                    vbCrLf & vbCrLf &
                    combinedContent.TrimStart()
            Else
                _knowledgeContent = combinedContent
            End If

            _knowledgeFilePath = GetKnowledgePathLabelAfterLoad(selectedPath, isFile, appendToExisting)

            ' Update runtime cache
            _cachedKnowledgeContent = _knowledgeContent
            _cachedKnowledgeFilePath = _knowledgeFilePath

            ' Auto-enable persistence for large knowledge; otherwise persist only if user already enabled it.
            Dim autoPersisted As Boolean = AutoEnablePersistKnowledgeIfLarge(ctx.LoadedFiles.Count)

            If _chkPersistKnowledge.Checked AndAlso Not autoPersisted Then
                Try
                    PersistKnowledgeToTempFile()
                    AppendSystemMessage(
                        $"Knowledge {If(appendToExisting, "added and persisted", "loaded and persisted")} " &
                        $"({GetKnowledgeSummaryText()}, {ctx.LoadedFiles.Count:N0} new file(s)).")
                Catch ex As System.Exception
                    AppendSystemMessage(
                        $"Knowledge {If(appendToExisting, "added", "loaded")} ({GetKnowledgeSummaryText()}) but failed to persist: {ex.Message}")
                End Try
            ElseIf Not _chkPersistKnowledge.Checked Then
                AppendSystemMessage(
                    $"Knowledge {If(appendToExisting, "added", "loaded")}: " &
                    $"{ctx.LoadedFiles.Count:N0} new file(s), {GetKnowledgeSummaryText()} total.")
            End If

            UpdateWindowTitle()
            BringDiscussFormToFront()

            Try
                My.Settings.DiscussKnowledgePath = selectedPath  ' Save both files AND directories
                My.Settings.Save()
            Catch
            End Try

        Catch ex As Exception
            RemoveWorkingIndicator()
            AppendSystemMessage($"Error loading knowledge: {ex.Message}")
            BringDiscussFormToFront()
        End Try
    End Function

    ''' <summary>
    ''' Loads a single knowledge file via the shared file importer used by Freestyle.
    ''' This aligns DiscussInky with sandboxed readers and shared file-type support.
    ''' </summary>
    ''' <param name="filePath">Path to the file to load.</param>
    ''' <param name="enableOCR">Whether to enable OCR for PDF files.</param>
    ''' <param name="silent">Whether to suppress error messages.</param>
    ''' <param name="askWorksheetSelection">
    ''' For Excel files, whether to prompt the user to select one worksheet or all worksheets.
    ''' </param>
    ''' <returns>
    ''' Tuple of (content, pdfMayBeIncomplete) where pdfMayBeIncomplete is True if PDF
    ''' heuristics suggest images/scans but OCR was not performed.
    ''' </returns>
    Private Async Function LoadSingleKnowledgeFileAsync(filePath As String,
                                                        enableOCR As Boolean,
                                                        silent As Boolean,
                                                        Optional askWorksheetSelection As Boolean = False) As Task(Of (Content As String, PdfMayBeIncomplete As Boolean, UserCancelled As Boolean))
        If String.IsNullOrWhiteSpace(filePath) OrElse Not File.Exists(filePath) Then
            Return ("", False, False)
        End If

        Try
            Dim result = Await Globals.ThisAddIn.GetFileContentEx(
                optionalFilePath:=filePath,
                Silent:=silent,
                DoOCR:=enableOCR,
                AskUser:=False,
                AskWorksheetSelection:=askWorksheetSelection)

            Return (If(result.Content, ""), result.PdfMayBeIncomplete, result.UserCancelled)

        Catch ex As Exception
            If Not silent Then
                AppendSystemMessage($"Error loading {Path.GetFileName(filePath)}: {ex.Message}")
            End If
            Return ("", False, False)
        End Try
    End Function

    Private Async Sub OnManageKnowledgeDocumentsClick(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(_knowledgeContent) Then
            AppendSystemMessage("No knowledge is currently loaded.")
            Return
        End If

        Await ManageKnowledgeDocumentsAsync(
            "Select one or more knowledge documents, then choose Compact Selected, Delete Selected, Edit Selected, or Convert to Index.")

        BringDiscussFormToFront()
    End Sub

    ''' <summary>
    ''' Opens the attached-index manager, where the user can remove attached searchable indexes
    ''' or convert a file into a new searchable index. Indexes are treated as knowledge sources,
    ''' so changes are persisted immediately (and copied to durable storage when persistence is on).
    ''' </summary>
    Private Sub OnManageIndexesClick(sender As Object, e As EventArgs)
        ShowManageIndexesDialog()
        BringDiscussFormToFront()
    End Sub

    ''' <summary>
    ''' Removes an attached index from the session. If the index file lives in this session's
    ''' durable copy folder, that copy is deleted too; externally referenced files are left alone.
    ''' Archived index sidecars are never touched here.
    ''' </summary>
    Private Sub RemoveAttachedIndex(indexRef As DiscussIndexRef)
        If indexRef Is Nothing Then
            Return
        End If

        _attachedIndexes.Remove(indexRef)
        _indexConversationState.Remove(indexRef.Id)

        Try
            Dim sessionDir As String = GetSessionIndexDirectoryPath()
            If Not String.IsNullOrWhiteSpace(indexRef.ActivePath) AndAlso
               File.Exists(indexRef.ActivePath) AndAlso
               String.Equals(Path.GetDirectoryName(indexRef.ActivePath), sessionDir, StringComparison.OrdinalIgnoreCase) Then
                File.Delete(indexRef.ActivePath)
            End If
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try

        AppendSystemMessage($"Removed searchable index '{indexRef.DisplayName}'.")
        UpdateWindowTitle()
        PersistCurrentSessionSettings()
    End Sub

    ''' <summary>
    ''' Prompts for a file and converts it into a new standalone searchable index that is attached
    ''' to the session. Reuses the shared file importer and the semantic-index generator. If the
    ''' picked file is already an index, it is attached in place instead.
    ''' </summary>
    Private Async Function ConvertFileToIndexAsync() As Task(Of Boolean)
        Dim selectedPath As String = ""

        Globals.ThisAddIn.DragDropFormLabel = "... a document to convert into a searchable index, or click Browse"
        Globals.ThisAddIn.DragDropFormFilter = ""

        Using frm As New DragDropForm(DragDropMode.FileOrDirectory)
            If frm.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
                selectedPath = frm.SelectedFilePath
            End If
        End Using

        Globals.ThisAddIn.DragDropFormLabel = ""
        Globals.ThisAddIn.DragDropFormFilter = ""

        If String.IsNullOrWhiteSpace(selectedPath) OrElse Not File.Exists(selectedPath) Then
            Return False
        End If

        ' Already an index: just attach it (with the standard SHA-256 dedupe guard).
        If SharedMethods.IsPotentiallySemanticSearchIndexedTextFile(selectedPath) Then
            Return AttachIndexFromFile(selectedPath)
        End If

        If Not TryBeginExclusive("converting a file into a searchable index") Then
            Return False
        End If

        Dim fileName As String = Path.GetFileName(selectedPath)
        Dim indexDir As String = GetSessionIndexDirectoryPath()
        Dim shortId As String = "i" & Guid.NewGuid().ToString("N").Substring(0, 4)
        Dim outputPath As String = Path.Combine(indexDir, shortId & IndexCopyFileExtension)

        Try
            Using progressScope As New ThisAddIn.ProgressScope(
                $"{AN} - Creating searchable index",
                $"Reading '{fileName}' ...",
                1)

                Dim loaded = Await LoadSingleKnowledgeFileAsync(selectedPath, enableOCR:=False, silent:=False)

                If loaded.UserCancelled OrElse String.IsNullOrWhiteSpace(loaded.Content) Then
                    AppendSystemMessage($"Could not read '{fileName}' for indexing.")
                    Return False
                End If

                ' Wrap in the canonical combine wrapper so the generator records source attribution.
                Dim combined As String =
                    $"<document1 name=""{fileName}"">" & vbCrLf & loaded.Content & vbCrLf & "</document1>"

                Directory.CreateDirectory(indexDir)

                Dim generationProgress As New System.Progress(Of SharedMethods.SemanticSearchIndexGenerationProgress)(
                    Sub(update As SharedMethods.SemanticSearchIndexGenerationProgress)
                        Dim segmentCount As Integer = System.Math.Max(1, update.SegmentCount)
                        Dim segmentNumber As Integer = System.Math.Max(0, System.Math.Min(update.SegmentNumber, segmentCount))
                        Dim statusMessage As String = If(update.Message, "").Trim()
                        If String.IsNullOrWhiteSpace(statusMessage) Then
                            statusMessage = "Generating semantic metadata"
                        End If
                        ThisAddIn.ProgressScope.Report(
                            segmentNumber,
                            segmentCount,
                            $"{statusMessage} ({segmentNumber}/{segmentCount})")
                    End Sub)

                Dim generationResult As SharedMethods.SemanticSearchIndexGenerationResult =
                    Await SharedMethods.CreateSemanticSearchIndexFromTextAsync(
                        combined,
                        outputPath,
                        _context,
                        New SharedMethods.SemanticSearchIndexGeneratorOptions() With {
                            .SpecialTaskName = "Indexer",
                            .OverwriteOutput = True
                        },
                        generationProgress,
                        progressScope.Token).ConfigureAwait(False)

                _attachedIndexes.Add(New DiscussIndexRef() With {
                    .Id = shortId,
                    .DisplayName = fileName,
                    .ActivePath = outputPath,
                    .ContentSha256 = If(generationResult IsNot Nothing, generationResult.ContentSha256, "")
                })

                ' Persistence applies to indexes: if it is already on, copy the new index to durable storage.
                If _chkPersistKnowledge.Checked Then
                    PersistAttachedIndexes()
                End If

                ThisAddIn.ProgressScope.Report(
                    System.Math.Max(1, If(generationResult IsNot Nothing, generationResult.SegmentCount, 1)),
                    System.Math.Max(1, If(generationResult IsNot Nothing, generationResult.SegmentCount, 1)),
                    "Completed successfully.")

                Dim segmentInfo As String =
                    If(generationResult IsNot Nothing,
                       $"{generationResult.SegmentCount:N0} segment(s)",
                       "an index")

                AppendSystemMessage($"Converted '{fileName}' into a searchable index ({segmentInfo}).")
                OfferDesktopIndexCopy(outputPath, fileName)
                UpdateWindowTitle()
                PersistCurrentSessionSettings()
                Return True
            End Using

        Catch ex As System.OperationCanceledException
            AppendSystemMessage($"Index creation for '{fileName}' was cancelled.")

            Try
                If File.Exists(outputPath) Then
                    File.Delete(outputPath)
                End If
            Catch cleanupEx As Exception
                System.Diagnostics.Debug.WriteLine(cleanupEx.Message)
            End Try

            Return False

        Catch ex As Exception
            AppendSystemMessage($"Failed to convert '{fileName}' to an index: {ex.Message}")

            Try
                If File.Exists(outputPath) Then
                    File.Delete(outputPath)
                End If
            Catch cleanupEx As Exception
                System.Diagnostics.Debug.WriteLine(cleanupEx.Message)
            End Try

            Return False
        Finally
            EndExclusive()
        End Try
    End Function

    ''' <summary>
    ''' Modal manager for attached searchable indexes: remove selected indexes or convert a file
    ''' into a new index. The list reflects the current attachments and refreshes after each action.
    ''' </summary>
    Private Sub ShowManageIndexesDialog()
        Using frm As New Form() With {
            .Text = $"{AN} - Manage Indexes",
            .StartPosition = FormStartPosition.CenterParent,
            .Size = New System.Drawing.Size(680, 420),
            .MinimumSize = New System.Drawing.Size(520, 320),
            .FormBorderStyle = FormBorderStyle.Sizable,
            .Font = New System.Drawing.Font("Segoe UI", 9.0F),
            .AutoScaleDimensions = New System.Drawing.SizeF(96.0F, 96.0F),
            .AutoScaleMode = AutoScaleMode.Dpi,
            .ShowInTaskbar = False
        }
            Try
                frm.Icon = Me.Icon
            Catch
            End Try

            Dim layout As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 3,
                .Padding = New Padding(12)
            }
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            layout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            frm.Controls.Add(layout)

            Dim lblInfo As New Label() With {
                .AutoSize = True,
                .Dock = DockStyle.Top,
                .Margin = New Padding(0, 0, 0, 8),
                .Text = "Attached searchable indexes are queried per message alongside any loaded knowledge. " &
                        "Remove indexes you no longer need, or convert a file into a new index."
            }

            Dim lstIndexes As New ListBox() With {
                .Dock = DockStyle.Fill,
                .IntegralHeight = False,
                .SelectionMode = SelectionMode.MultiExtended
            }

            Dim buttonBar As New FlowLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.LeftToRight,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .WrapContents = True,
                .Padding = New Padding(0, 6, 0, 0),
                .Margin = New Padding(0)
            }

            Dim btnRemove As New Button() With {.Text = "Remove Selected", .AutoSize = True}
            Dim btnConvert As New Button() With {.Text = "Convert File to Index...", .AutoSize = True}
            Dim btnClose As New Button() With {.Text = "Close", .AutoSize = True}

            buttonBar.Controls.Add(btnRemove)
            buttonBar.Controls.Add(btnConvert)
            buttonBar.Controls.Add(btnClose)

            layout.Controls.Add(lblInfo, 0, 0)
            layout.Controls.Add(lstIndexes, 0, 1)
            layout.Controls.Add(buttonBar, 0, 2)

            frm.CancelButton = btnClose

            Dim refreshList As System.Action =
                Sub()
                    lstIndexes.BeginUpdate()
                    lstIndexes.Items.Clear()
                    For Each idx In _attachedIndexes
                        lstIndexes.Items.Add(idx)
                    Next
                    lstIndexes.EndUpdate()
                    btnRemove.Enabled = lstIndexes.Items.Count > 0
                End Sub

            lstIndexes.DisplayMember = "DisplayName"

            AddHandler btnClose.Click, Sub() frm.Close()

            AddHandler btnRemove.Click,
                Sub()
                    Dim selected As New List(Of DiscussIndexRef)()
                    For Each item As Object In lstIndexes.SelectedItems
                        Dim indexRef = TryCast(item, DiscussIndexRef)
                        If indexRef IsNot Nothing Then
                            selected.Add(indexRef)
                        End If
                    Next

                    If selected.Count = 0 Then
                        ShowCustomMessageBox("Select at least one index to remove.")
                        Return
                    End If

                    Dim confirm = ShowCustomYesNoBox(
                        $"Remove {selected.Count:N0} attached index(es) from this discussion?",
                        "Yes, remove",
                        "No, keep",
                        $"{AN} - Remove Indexes")

                    If confirm <> 1 Then
                        Return
                    End If

                    For Each indexRef In selected
                        RemoveAttachedIndex(indexRef)
                    Next

                    refreshList.Invoke()
                End Sub

            AddHandler btnConvert.Click,
                Async Sub()
                    btnConvert.Enabled = False
                    Try
                        If Await ConvertFileToIndexAsync() Then
                            refreshList.Invoke()
                        End If
                    Finally
                        btnConvert.Enabled = True
                    End Try
                End Sub

            refreshList.Invoke()
            frm.ShowDialog(Me)
        End Using
    End Sub

    Private Shared Function TrimSingleBoundaryLineBreak(value As String) As String
        Dim result As String = If(value, "")

        If result.StartsWith(vbCrLf, StringComparison.Ordinal) Then
            result = result.Substring(2)
        ElseIf result.StartsWith(vbLf, StringComparison.Ordinal) OrElse result.StartsWith(vbCr, StringComparison.Ordinal) Then
            result = result.Substring(1)
        End If

        If result.EndsWith(vbCrLf, StringComparison.Ordinal) Then
            result = result.Substring(0, result.Length - 2)
        ElseIf result.EndsWith(vbLf, StringComparison.Ordinal) OrElse result.EndsWith(vbCr, StringComparison.Ordinal) Then
            result = result.Substring(0, result.Length - 1)
        End If

        Return result
    End Function

    Private Function GetSyntheticKnowledgeDocumentName(fragmentIndex As Integer) As String
        If fragmentIndex = 1 AndAlso Not String.IsNullOrWhiteSpace(_knowledgeFilePath) Then
            Try
                Dim fileName As String = Path.GetFileName(_knowledgeFilePath)

                If Not String.IsNullOrWhiteSpace(fileName) Then
                    Return fileName
                End If
            Catch
            End Try

            Return _knowledgeFilePath
        End If

        If fragmentIndex <= 1 Then
            Return "Knowledge"
        End If

        Return $"Knowledge fragment {fragmentIndex}"
    End Function

    Private Function ParseKnowledgeDocuments(Optional content As String = Nothing) As List(Of KnowledgeDocumentEntry)
        Dim source As String = If(content, _knowledgeContent)
        Dim result As New List(Of KnowledgeDocumentEntry)()

        If String.IsNullOrWhiteSpace(source) Then
            Return result
        End If

        Dim pattern As String = "<document(?<n>\d+)(?:\s+name=""(?<name>[^""]*)"")?\s*>(?<body>[\s\S]*?)</document\k<n>\s*>"
        Dim matches As MatchCollection = Regex.Matches(source, pattern, RegexOptions.IgnoreCase)

        If matches.Count = 0 Then
            result.Add(
                New KnowledgeDocumentEntry With {
                    .Number = 1,
                    .Name = GetSyntheticKnowledgeDocumentName(1),
                    .Content = source,
                    .StartIndex = 0,
                    .Length = source.Length,
                    .IsTagged = False
                })

            Return result
        End If

        Dim position As Integer = 0
        Dim fragmentIndex As Integer = 0

        For Each match As Match In matches
            If match.Index > position Then
                Dim rawSegment As String = source.Substring(position, match.Index - position)

                If Not String.IsNullOrWhiteSpace(rawSegment) Then
                    fragmentIndex += 1

                    result.Add(
                        New KnowledgeDocumentEntry With {
                            .Number = -fragmentIndex,
                            .Name = GetSyntheticKnowledgeDocumentName(fragmentIndex),
                            .Content = TrimSingleBoundaryLineBreak(rawSegment),
                            .StartIndex = position,
                            .Length = match.Index - position,
                            .IsTagged = False
                        })
                End If
            End If

            Dim number As Integer = 0
            Integer.TryParse(match.Groups("n").Value, number)

            result.Add(
                New KnowledgeDocumentEntry With {
                    .Number = number,
                    .Name = match.Groups("name").Value,
                    .Content = TrimSingleBoundaryLineBreak(match.Groups("body").Value),
                    .StartIndex = match.Index,
                    .Length = match.Length,
                    .IsTagged = True
                })

            position = match.Index + match.Length
        Next

        If position < source.Length Then
            Dim rawTail As String = source.Substring(position)

            If Not String.IsNullOrWhiteSpace(rawTail) Then
                fragmentIndex += 1

                result.Add(
                    New KnowledgeDocumentEntry With {
                        .Number = -fragmentIndex,
                        .Name = GetSyntheticKnowledgeDocumentName(fragmentIndex),
                        .Content = TrimSingleBoundaryLineBreak(rawTail),
                        .StartIndex = position,
                        .Length = source.Length - position,
                        .IsTagged = False
                    })
            End If
        End If

        Return result.OrderBy(Function(x) x.StartIndex).ThenBy(Function(x) x.Number).ToList()
    End Function

    Private Function RequiresKnowledgeDocumentNormalization(source As String,
                                                           Optional forceTagSingleUntaggedDocument As Boolean = False) As Boolean
        If String.IsNullOrWhiteSpace(source) Then
            Return False
        End If

        Dim entries As List(Of KnowledgeDocumentEntry) = ParseKnowledgeDocuments(source)

        If entries.Count = 0 Then
            Return False
        End If

        If forceTagSingleUntaggedDocument AndAlso
           entries.Count = 1 AndAlso
           Not entries(0).IsTagged Then
            Return True
        End If

        If entries.Any(Function(x) Not x.IsTagged) Then
            Return True
        End If

        Dim seenNumbers As New HashSet(Of Integer)()

        For Each entry In entries
            If Not entry.IsTagged Then
                Continue For
            End If

            If entry.Number <= 0 Then
                Return True
            End If

            If Not seenNumbers.Add(entry.Number) Then
                Return True
            End If
        Next

        Return False
    End Function

    Private Function NormalizeKnowledgeDocumentsToTaggedContent(source As String) As String
        If String.IsNullOrWhiteSpace(source) Then
            Return source
        End If

        Dim entries As List(Of KnowledgeDocumentEntry) = ParseKnowledgeDocuments(source)

        If entries.Count = 0 Then
            Return source
        End If

        Dim normalizedEntries As New List(Of KnowledgeDocumentEntry)()

        For i As Integer = 0 To entries.Count - 1
            Dim entry As KnowledgeDocumentEntry = entries(i)

            entry.Number = i + 1
            entry.IsTagged = True

            If String.IsNullOrWhiteSpace(entry.Name) Then
                entry.Name = GetSyntheticKnowledgeDocumentName(i + 1)
            End If

            normalizedEntries.Add(entry)
        Next

        Return BuildKnowledgeContentFromEntries(normalizedEntries)
    End Function

    Private Function PrepareKnowledgeContentForAppending(existingContent As String) As String
        If String.IsNullOrWhiteSpace(existingContent) Then
            Return existingContent
        End If

        If Not RequiresKnowledgeDocumentNormalization(existingContent, forceTagSingleUntaggedDocument:=True) Then
            Return existingContent
        End If

        Return NormalizeKnowledgeDocumentsToTaggedContent(existingContent)
    End Function

    Private Sub NormalizeKnowledgeDocumentsForManagementIfNeeded()
        If String.IsNullOrWhiteSpace(_knowledgeContent) Then
            Return
        End If

        If Not RequiresKnowledgeDocumentNormalization(_knowledgeContent) Then
            Return
        End If

        Dim splash As New SharedMethods.SplashScreen("Please wait ...   ")

        Try
            splash.Show()
            System.Windows.Forms.Application.DoEvents()

            Dim normalizedContent As String = NormalizeKnowledgeDocumentsToTaggedContent(_knowledgeContent)

            If Not String.Equals(normalizedContent, _knowledgeContent, StringComparison.Ordinal) Then
                ApplyKnowledgeContentMutation(normalizedContent)
                AppendSystemMessage("Knowledge documents were normalized for management.")
            End If
        Finally
            Try
                splash.Close()
            Catch
            End Try

            Try
                splash.Dispose()
            Catch
            End Try
        End Try
    End Sub

    Private Function BuildKnowledgeContentFromEntries(entries As IEnumerable(Of KnowledgeDocumentEntry)) As String
        If entries Is Nothing Then
            Return Nothing
        End If

        Dim orderedEntries As List(Of KnowledgeDocumentEntry) =
            entries.
                OrderBy(Function(x) x.StartIndex).
                ThenBy(Function(x) x.Number).
                ToList()

        If orderedEntries.Count = 0 Then
            Return Nothing
        End If

        If orderedEntries.Count = 1 AndAlso Not orderedEntries(0).IsTagged Then
            Return orderedEntries(0).Content
        End If

        Dim sb As New StringBuilder()

        For Each entry In orderedEntries
            If entry.IsTagged Then
                Dim safeName As String = If(entry.Name, "").Replace("""", "'")

                sb.Append($"<document{entry.Number}")

                If safeName.Length > 0 Then
                    sb.Append($" name=""{safeName}""")
                End If

                sb.Append(">").AppendLine()

                Dim documentContent As String = If(entry.Content, "")
                sb.Append(documentContent)

                If documentContent.Length > 0 AndAlso
                   Not documentContent.EndsWith(vbCrLf, StringComparison.Ordinal) AndAlso
                   Not documentContent.EndsWith(vbLf, StringComparison.Ordinal) AndAlso
                   Not documentContent.EndsWith(vbCr, StringComparison.Ordinal) Then
                    sb.AppendLine()
                End If

                sb.Append($"</document{entry.Number}>").AppendLine()
            Else
                sb.Append(If(entry.Content, ""))
            End If
        Next

        Return sb.ToString().TrimEnd()
    End Function

    Private Sub ApplyKnowledgeContentMutation(newKnowledgeContent As String)
        If String.IsNullOrWhiteSpace(newKnowledgeContent) Then
            ClearPlainKnowledge()
            UpdateWindowTitle()
            PersistCurrentSessionSettings()
            Return
        End If

        _knowledgeContent = newKnowledgeContent
        _cachedKnowledgeContent = _knowledgeContent
        _cachedKnowledgeFilePath = _knowledgeFilePath

        If _chkPersistKnowledge.Checked Then
            PersistKnowledgeToTempFile()
        End If

        UpdateWindowTitle()
        PersistCurrentSessionSettings()
    End Sub

    Private Function IsKnowledgeRateLimitResponse(response As String) As Boolean
        If String.IsNullOrWhiteSpace(response) Then
            Return False
        End If

        Return response.IndexOf("HTTP Error 429", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               response.IndexOf("Too Many Requests", StringComparison.OrdinalIgnoreCase) >= 0
    End Function

    Private Function ShouldOfferKnowledgeCompaction(response As String) As Boolean
        If String.IsNullOrWhiteSpace(_knowledgeContent) Then
            Return False
        End If

        If String.IsNullOrWhiteSpace(response) Then
            Return True
        End If

        Return IsKnowledgeRateLimitResponse(response)
    End Function

    Private Function GetStoredCompactPrompt() As String
        Try
            Return If(My.Settings.DiscussLastCompactPrompt, "")
        Catch
            Return ""
        End Try
    End Function

    Private Function PromptForCompactPrompt() As String
        Dim previousPrompt As String = GetStoredCompactPrompt()
        Dim proposedPrompt As String = If(String.IsNullOrWhiteSpace(_context.SP_Compact), previousPrompt, _context.SP_Compact)

        Dim promptText As String =
            ShowCustomInputBox(
                "Enter the prompt used to compact the selected knowledge document(s)." & vbCrLf & vbCrLf &
                "Ctrl+P inserts your last saved compact prompt. The dialog always proposes the default prompt first.",
                $"{AN} - Compact Knowledge",
                False,
                proposedPrompt,
                previousPrompt,
                Context:=_context)

        If promptText = "ESC" Then
            Return Nothing
        End If

        promptText = If(promptText, "").Trim()

        If promptText.Length = 0 Then
            ShowCustomMessageBox("No compact prompt was entered.")
            Return Nothing
        End If

        If promptText <> _context.SP_Compact.Trim() Then
            Try
                My.Settings.DiscussLastCompactPrompt = promptText
                My.Settings.Save()
            Catch
            End Try
        End If

        Return promptText
    End Function

    Private Function ShowKnowledgeDocumentManagerDialog(entries As IReadOnlyList(Of KnowledgeDocumentEntry),
                                                        preselectedNumbers As IEnumerable(Of Integer),
                                                        instruction As String) As KnowledgeDocumentManagerResult
        Dim result As New KnowledgeDocumentManagerResult With {
            .Action = KnowledgeDocumentManagerAction.None,
            .SelectedDocumentNumbers = New List(Of Integer)()
        }

        If entries Is Nothing OrElse entries.Count = 0 Then
            Return result
        End If

        Using dlg As New Form() With {
            .Text = $"{AN} - Manage Knowledge Documents",
            .StartPosition = FormStartPosition.CenterParent,
            .FormBorderStyle = FormBorderStyle.Sizable,
            .MinimizeBox = False,
            .MaximizeBox = True,
            .ShowInTaskbar = False,
            .TopMost = True,
            .Font = New System.Drawing.Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point),
            .AutoScaleMode = AutoScaleMode.Dpi,
            .Size = New System.Drawing.Size(820, 620),
            .MinimumSize = New System.Drawing.Size(640, 420)
        }
            Try
                dlg.Icon = Me.Icon
            Catch
            End Try

            Dim outer As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 4,
                .Padding = New Padding(16, 12, 16, 12)
            }
            outer.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            outer.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            outer.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            outer.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            dlg.Controls.Add(outer)

            Dim lblInstruction As New Label() With {
                .AutoSize = True,
                .Dock = DockStyle.Top,
                .Text = instruction & vbCrLf & vbCrLf &
                        $"Current knowledge: {entries.Count:N0} document(s), {If(_knowledgeContent, "").Length:N0} characters.",
                .MaximumSize = New System.Drawing.Size(760, 0),
                .Margin = New Padding(0, 0, 0, 8)
            }
            outer.Controls.Add(lblInstruction, 0, 0)

            Dim txtFilter As New TextBox() With {
                .Dock = DockStyle.Top,
                .Margin = New Padding(0, 0, 0, 8)
            }
            outer.Controls.Add(txtFilter, 0, 1)

            Dim chkList As New CheckedListBox() With {
                .Dock = DockStyle.Fill,
                .CheckOnClick = True,
                .IntegralHeight = False
            }
            outer.Controls.Add(chkList, 0, 2)

            Dim pnlButtons As New FlowLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.RightToLeft,
                .WrapContents = True,
                .AutoSize = True,
                .Padding = New Padding(0, 8, 0, 0),
                .Margin = New Padding(0)
            }
            outer.Controls.Add(pnlButtons, 0, 3)

            Dim btnClose As New Button() With {.Text = "Close", .AutoSize = True}
            Dim btnEdit As New Button() With {.Text = "Edit Selected", .AutoSize = True}
            Dim btnDelete As New Button() With {.Text = "Delete Selected", .AutoSize = True}
            Dim btnCompact As New Button() With {.Text = "Compact Selected", .AutoSize = True}
            Dim btnIndex As New Button() With {.Text = "Convert to Index", .AutoSize = True}
            Dim btnToggleAll As New Button() With {.Text = "Select All", .AutoSize = True}

            pnlButtons.Controls.Add(btnClose)
            pnlButtons.Controls.Add(btnEdit)
            pnlButtons.Controls.Add(btnDelete)
            pnlButtons.Controls.Add(btnCompact)
            pnlButtons.Controls.Add(btnIndex)
            pnlButtons.Controls.Add(btnToggleAll)

            dlg.CancelButton = btnClose

            Dim selectedNumbers As New HashSet(Of Integer)()

            If preselectedNumbers IsNot Nothing Then
                For Each number In preselectedNumbers
                    selectedNumbers.Add(number)
                Next
            End If

            Dim allItems As New List(Of KnowledgeDocumentSelectionItem)()

            For Each entry In entries.OrderBy(Function(x) x.StartIndex).ThenBy(Function(x) x.Number)
                allItems.Add(New KnowledgeDocumentSelectionItem(entry.Number, entry.DisplayText))
            Next

            Dim isUpdating As Boolean = False

            Dim rebuildList As System.Action =
                Sub()
                    Dim filter As String = If(txtFilter.Text, "").Trim()

                    isUpdating = True
                    chkList.BeginUpdate()

                    Try
                        chkList.Items.Clear()

                        For Each item In allItems
                            If filter.Length = 0 OrElse
                               item.DisplayText.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0 Then
                                chkList.Items.Add(item, selectedNumbers.Contains(item.Number))
                            End If
                        Next
                    Finally
                        chkList.EndUpdate()
                        isUpdating = False
                    End Try
                End Sub

            Dim areAllVisibleItemsChecked As Func(Of Boolean) =
                Function() As Boolean
                    If chkList.Items.Count = 0 Then
                        Return False
                    End If

                    For i As Integer = 0 To chkList.Items.Count - 1
                        If Not chkList.GetItemChecked(i) Then
                            Return False
                        End If
                    Next

                    Return True
                End Function

            Dim updateToggleText As System.Action =
                Sub()
                    btnToggleAll.Text = If(areAllVisibleItemsChecked.Invoke(), "Unselect All", "Select All")
                End Sub

            Dim getSelectedNumbers As Func(Of List(Of Integer)) =
                Function() As List(Of Integer)
                    Return selectedNumbers.OrderBy(Function(x) x).ToList()
                End Function

            Dim acceptAction As Action(Of KnowledgeDocumentManagerAction) =
                Sub(requestedAction As KnowledgeDocumentManagerAction)
                    Dim chosen As List(Of Integer) = getSelectedNumbers.Invoke()

                    If chosen.Count = 0 Then
                        ShowCustomMessageBox("Select at least one knowledge document first.")
                        Return
                    End If

                    result.Action = requestedAction
                    result.SelectedDocumentNumbers = chosen
                    dlg.DialogResult = DialogResult.OK
                    dlg.Close()
                End Sub

            AddHandler txtFilter.TextChanged,
                Sub()
                    rebuildList.Invoke()
                    updateToggleText.Invoke()
                End Sub

            AddHandler chkList.ItemCheck,
                Sub(sender As Object, args As ItemCheckEventArgs)
                    If isUpdating Then
                        Return
                    End If

                    Dim item As KnowledgeDocumentSelectionItem = DirectCast(chkList.Items(args.Index), KnowledgeDocumentSelectionItem)

                    If args.NewValue = CheckState.Checked Then
                        selectedNumbers.Add(item.Number)
                    Else
                        selectedNumbers.Remove(item.Number)
                    End If

                    chkList.BeginInvoke(
                        New MethodInvoker(
                            Sub()
                                updateToggleText.Invoke()
                            End Sub))
                End Sub

            AddHandler chkList.DoubleClick,
                Sub()
                    Dim idx As Integer = chkList.SelectedIndex
                    If idx >= 0 Then
                        chkList.SetItemChecked(idx, Not chkList.GetItemChecked(idx))
                    End If
                End Sub

            AddHandler btnToggleAll.Click,
                Sub()
                    Dim shouldCheck As Boolean = Not areAllVisibleItemsChecked.Invoke()

                    isUpdating = True
                    chkList.BeginUpdate()

                    Try
                        For i As Integer = 0 To chkList.Items.Count - 1
                            Dim item As KnowledgeDocumentSelectionItem = DirectCast(chkList.Items(i), KnowledgeDocumentSelectionItem)
                            chkList.SetItemChecked(i, shouldCheck)

                            If shouldCheck Then
                                selectedNumbers.Add(item.Number)
                            Else
                                selectedNumbers.Remove(item.Number)
                            End If
                        Next
                    Finally
                        chkList.EndUpdate()
                        isUpdating = False
                    End Try

                    updateToggleText.Invoke()
                End Sub

            AddHandler btnCompact.Click, Sub() acceptAction.Invoke(KnowledgeDocumentManagerAction.CompactSelected)
            AddHandler btnDelete.Click, Sub() acceptAction.Invoke(KnowledgeDocumentManagerAction.DeleteSelected)
            AddHandler btnEdit.Click, Sub() acceptAction.Invoke(KnowledgeDocumentManagerAction.EditSelected)
            AddHandler btnIndex.Click, Sub() acceptAction.Invoke(KnowledgeDocumentManagerAction.IndexSelected)
            AddHandler btnClose.Click,
                Sub()
                    result.Action = KnowledgeDocumentManagerAction.None
                    result.SelectedDocumentNumbers = getSelectedNumbers.Invoke()
                    dlg.DialogResult = DialogResult.Cancel
                    dlg.Close()
                End Sub

            rebuildList.Invoke()
            updateToggleText.Invoke()

            dlg.ShowDialog(Me)
        End Using

        Return result
    End Function

    Private Function DeleteKnowledgeDocuments(selectedDocumentNumbers As IEnumerable(Of Integer)) As Integer
        If selectedDocumentNumbers Is Nothing Then
            Return 0
        End If

        Dim selectedSet As New HashSet(Of Integer)(selectedDocumentNumbers)
        If selectedSet.Count = 0 Then
            Return 0
        End If

        Dim confirmDelete As Integer =
            ShowCustomYesNoBox(
                $"Are you sure you want to delete {selectedSet.Count:N0} selected knowledge document(s) from the current knowledge?",
                "Yes, delete",
                "No, keep",
                $"{AN} - Delete Knowledge Documents")

        If confirmDelete <> 1 Then
            Return 0
        End If

        Dim existingEntries As List(Of KnowledgeDocumentEntry) = ParseKnowledgeDocuments()
        Dim remainingEntries As List(Of KnowledgeDocumentEntry) =
            existingEntries.
                Where(Function(x) Not selectedSet.Contains(x.Number)).
                ToList()

        Dim removedCount As Integer = existingEntries.Count - remainingEntries.Count

        If removedCount <= 0 Then
            Return 0
        End If

        If remainingEntries.Count = 0 Then
            DeleteCurrentKnowledge()
        Else
            ApplyKnowledgeContentMutation(BuildKnowledgeContentFromEntries(remainingEntries))
            AppendSystemMessage($"Deleted {removedCount:N0} knowledge document(s).")
        End If

        Return removedCount
    End Function

    ''' <summary>
    ''' Converts the selected knowledge documents into a single standalone semantic index,
    ''' attaches it as a searchable source, and removes the originals from the inlined knowledge
    ''' to reduce prompt size. The index is stored durably under %AppData%\redink so it survives
    ''' temp cleanup and can be persisted/archived. The user is informed of the outcome.
    ''' </summary>
    Private Async Function MakeKnowledgeDocumentsSearchableAsync(selectedDocumentNumbers As IEnumerable(Of Integer)) As Task(Of Integer)
        If selectedDocumentNumbers Is Nothing Then
            Return 0
        End If

        Dim selectedSet As New HashSet(Of Integer)(selectedDocumentNumbers)
        If selectedSet.Count = 0 Then
            Return 0
        End If

        Dim allEntries As List(Of KnowledgeDocumentEntry) = ParseKnowledgeDocuments()
        Dim targetEntries As List(Of KnowledgeDocumentEntry) =
            allEntries.
                Where(Function(x) selectedSet.Contains(x.Number)).
                OrderBy(Function(x) x.StartIndex).
                ThenBy(Function(x) x.Number).
                ToList()

        If targetEntries.Count = 0 Then
            Return 0
        End If

        Dim confirm As Integer =
            ShowCustomYesNoBox(
                $"Create a searchable index from {targetEntries.Count:N0} selected knowledge document(s)?" & vbCrLf & vbCrLf &
                "The selected documents will be removed from the inlined knowledge and replaced by a compact index that is searched per message. " &
                "This reduces the amount of text sent to the model for large or numerous documents.",
                "Yes, make searchable",
                "No, cancel",
                $"{AN} - Make Knowledge Searchable")

        If confirm <> 1 Then
            Return 0
        End If

        ' Combine the selected documents using the canonical wrapper separator so the generator
        ' can align segments to document boundaries.
        Dim combined As String = BuildKnowledgeContentFromEntries(targetEntries)
        If String.IsNullOrWhiteSpace(combined) Then
            AppendSystemMessage("The selected knowledge documents produced no content to index.")
            Return 0
        End If

        Dim indexDir As String = GetSessionIndexDirectoryPath()
        Directory.CreateDirectory(indexDir)

        Dim shortId As String = "i" & Guid.NewGuid().ToString("N").Substring(0, 4)
        Dim outputPath As String = Path.Combine(indexDir, shortId & IndexCopyFileExtension)

        If Not TryBeginExclusive("creating a searchable index") Then
            Return 0
        End If

        Try
            Using progressScope As New ThisAddIn.ProgressScope(
                $"{AN} - Creating searchable index",
                $"Preparing {targetEntries.Count:N0} document(s) ...",
                1)

                Dim generationProgress As New System.Progress(Of SharedMethods.SemanticSearchIndexGenerationProgress)(
                    Sub(update As SharedMethods.SemanticSearchIndexGenerationProgress)
                        Dim segmentCount As Integer = System.Math.Max(1, update.SegmentCount)
                        Dim segmentNumber As Integer = System.Math.Max(0, System.Math.Min(update.SegmentNumber, segmentCount))
                        Dim statusMessage As String = If(update.Message, "").Trim()
                        If String.IsNullOrWhiteSpace(statusMessage) Then
                            statusMessage = "Generating semantic metadata"
                        End If
                        ThisAddIn.ProgressScope.Report(
                            segmentNumber,
                            segmentCount,
                            $"{statusMessage} ({segmentNumber}/{segmentCount})")
                    End Sub)

                Dim generationResult As SharedMethods.SemanticSearchIndexGenerationResult =
                    Await SharedMethods.CreateSemanticSearchIndexFromTextAsync(
                        combined,
                        outputPath,
                        _context,
                        New SharedMethods.SemanticSearchIndexGeneratorOptions() With {
                            .SpecialTaskName = "Indexer",
                            .OverwriteOutput = True
                        },
                        generationProgress,
                        progressScope.Token).ConfigureAwait(False)

                Dim displayName As String =
                    If(targetEntries.Count = 1 AndAlso Not String.IsNullOrWhiteSpace(targetEntries(0).Name),
                       targetEntries(0).Name,
                       $"{targetEntries.Count} documents")

                _attachedIndexes.Add(New DiscussIndexRef() With {
                    .Id = shortId,
                    .DisplayName = displayName,
                    .ActivePath = outputPath,
                    .ContentSha256 = If(generationResult IsNot Nothing, generationResult.ContentSha256, "")
                })

                ' Remove the indexed originals from the inlined knowledge.
                Dim remaining As List(Of KnowledgeDocumentEntry) =
                    allEntries.Where(Function(x) Not selectedSet.Contains(x.Number)).ToList()

                If remaining.Count = 0 Then
                    ClearPlainKnowledge()
                    UpdateWindowTitle()
                    PersistCurrentSessionSettings()
                Else
                    ApplyKnowledgeContentMutation(BuildKnowledgeContentFromEntries(remaining))
                End If

                ThisAddIn.ProgressScope.Report(
                    System.Math.Max(1, If(generationResult IsNot Nothing, generationResult.SegmentCount, 1)),
                    System.Math.Max(1, If(generationResult IsNot Nothing, generationResult.SegmentCount, 1)),
                    "Completed successfully.")

                Dim segmentInfo As String =
                    If(generationResult IsNot Nothing,
                       $"{generationResult.SegmentCount:N0} segment(s)",
                       "an index")

                AppendSystemMessage(
                    $"Created a searchable index '{displayName}' ({segmentInfo}). " &
                    $"The {targetEntries.Count:N0} selected document(s) were removed from the inlined knowledge and will now be searched per message. " &
                    "The index is stored durably and will be included when you persist or archive this session.")

                OfferDesktopIndexCopy(outputPath, displayName)
                UpdateWindowTitle()
                Return targetEntries.Count
            End Using

        Catch ex As System.OperationCanceledException
            AppendSystemMessage("Index creation was cancelled. The selected knowledge documents were left unchanged.")

            Try
                If File.Exists(outputPath) Then
                    File.Delete(outputPath)
                End If
            Catch cleanupEx As Exception
                System.Diagnostics.Debug.WriteLine(cleanupEx.Message)
            End Try

            Return 0

        Catch ex As Exception
            AppendSystemMessage($"Failed to make the selected knowledge searchable: {ex.Message}")

            Try
                If File.Exists(outputPath) Then
                    File.Delete(outputPath)
                End If
            Catch cleanupEx As Exception
                System.Diagnostics.Debug.WriteLine(cleanupEx.Message)
            End Try

            Return 0
        Finally
            EndExclusive()
        End Try
    End Function

    Private Function EditKnowledgeDocuments(selectedDocumentNumbers As IEnumerable(Of Integer)) As Integer
        If selectedDocumentNumbers Is Nothing Then
            Return 0
        End If

        Dim selectedSet As New HashSet(Of Integer)(selectedDocumentNumbers)
        If selectedSet.Count = 0 Then
            Return 0
        End If

        Dim entries As List(Of KnowledgeDocumentEntry) = ParseKnowledgeDocuments()
        Dim editedCount As Integer = 0

        For i As Integer = 0 To entries.Count - 1
            If Not selectedSet.Contains(entries(i).Number) Then
                Continue For
            End If

            Dim entry As KnowledgeDocumentEntry = entries(i)
            Dim tempPath As String =
                Path.Combine(
                    Path.GetTempPath(),
                    $"redink-discuss-docedit-{Guid.NewGuid():N}-document{entry.Number}.txt")

            Try
                File.WriteAllText(tempPath, If(entry.Content, ""), Encoding.UTF8)

                Dim wasSaved As Boolean? = Nothing
                ShowTextFileEditor(
                    tempPath,
                    $"Edit knowledge document document{entry.Number}" &
                    If(String.IsNullOrWhiteSpace(entry.Name), "", $" ({entry.Name})"),
                    False,
                    _context,
                    wasSaved,
                    Me.Handle)

                If wasSaved.HasValue AndAlso wasSaved.Value Then
                    entry.Content = File.ReadAllText(tempPath, Encoding.UTF8)
                    entries(i) = entry
                    editedCount += 1
                End If
            Catch ex As Exception
                AppendSystemMessage($"Could not edit document{entry.Number}: {ex.Message}")
            Finally
                Try
                    If File.Exists(tempPath) Then
                        File.Delete(tempPath)
                    End If
                Catch
                End Try
            End Try
        Next

        If editedCount > 0 Then
            ApplyKnowledgeContentMutation(BuildKnowledgeContentFromEntries(entries))
            AppendSystemMessage($"Edited {editedCount:N0} knowledge document(s).")
        End If

        Return editedCount
    End Function

    Private Async Function CompactKnowledgeDocumentsAsync(selectedDocumentNumbers As IEnumerable(Of Integer)) As Task(Of Integer)
        If selectedDocumentNumbers Is Nothing Then
            Return 0
        End If

        Dim selectedSet As New HashSet(Of Integer)(selectedDocumentNumbers)
        If selectedSet.Count = 0 Then
            Return 0
        End If

        Dim compactPrompt As String = PromptForCompactPrompt()
        If String.IsNullOrWhiteSpace(compactPrompt) Then
            Return 0
        End If

        Dim entries As List(Of KnowledgeDocumentEntry) = ParseKnowledgeDocuments()
        Dim targetEntries As List(Of KnowledgeDocumentEntry) =
            entries.
                Where(Function(x) selectedSet.Contains(x.Number)).
                OrderBy(Function(x) x.StartIndex).
                ThenBy(Function(x) x.Number).
                ToList()

        If targetEntries.Count = 0 Then
            Return 0
        End If

        Dim compactedCount As Integer = 0
        Dim failedCount As Integer = 0
        Dim cancelled As Boolean = False

        ShowProgressBarInSeparateThread($"{AN} Compact Knowledge", "Compacting selected knowledge documents...")
        ProgressBarModule.CancelOperation = False
        ProgressBarModule.GlobalProgressMax = targetEntries.Count
        ProgressBarModule.GlobalProgressValue = 0
        ProgressBarModule.GlobalProgressLabel = "Starting..."

        Try
            For i As Integer = 0 To targetEntries.Count - 1
                If ProgressBarModule.CancelOperation Then
                    cancelled = True
                    Exit For
                End If

                Dim entry As KnowledgeDocumentEntry = targetEntries(i)

                ProgressBarModule.GlobalProgressValue = i + 1
                ProgressBarModule.GlobalProgressLabel = $"Compacting document{entry.Number}..."

                Dim sb As New StringBuilder()
                sb.AppendLine($"Document tag: document{entry.Number}")

                If Not String.IsNullOrWhiteSpace(entry.Name) Then
                    sb.AppendLine($"Document name: {entry.Name}")
                End If

                sb.AppendLine("<DOCUMENT_TO_COMPACT>")
                sb.AppendLine(If(entry.Content, ""))
                sb.AppendLine("</DOCUMENT_TO_COMPACT>")

                Dim compactedText As String = Await CallLlmWithSelectedModelAsync(compactPrompt, sb.ToString())
                compactedText = If(compactedText, "").Trim()

                If String.IsNullOrWhiteSpace(compactedText) Then
                    failedCount += 1
                    AppendSystemMessage($"Compaction returned an empty response for document{entry.Number}. The original content was kept.")
                    Continue For
                End If

                If IsKnowledgeRateLimitResponse(compactedText) Then
                    failedCount += 1
                    AppendSystemMessage($"Compaction returned a 429-style response for document{entry.Number}. The original content was kept.")
                    Continue For
                End If

                For j As Integer = 0 To entries.Count - 1
                    If entries(j).Number = entry.Number Then
                        Dim updatedEntry As KnowledgeDocumentEntry = entries(j)
                        updatedEntry.Content = compactedText
                        entries(j) = updatedEntry
                        compactedCount += 1
                        Exit For
                    End If
                Next
            Next
        Finally
            ProgressBarModule.CancelOperation = True
        End Try

        If compactedCount > 0 Then
            ApplyKnowledgeContentMutation(BuildKnowledgeContentFromEntries(entries))
            AppendSystemMessage($"Compacted {compactedCount:N0} knowledge document(s).")
        End If

        If failedCount > 0 Then
            AppendSystemMessage($"The original content was preserved for {failedCount:N0} knowledge document(s).")
        End If

        If cancelled Then
            AppendSystemMessage("Compaction cancelled by user.")
        End If

        Return compactedCount
    End Function

    Private Async Function ManageKnowledgeDocumentsAsync(Optional instruction As String = "") As Task(Of Integer)
        Dim totalAffectedCount As Integer = 0
        Dim selectedDocumentNumbers As New List(Of Integer)()
        Dim nextInstruction As String =
            If(
                String.IsNullOrWhiteSpace(instruction),
                "Select one or more knowledge documents, then choose Compact Selected, Delete Selected, or Edit Selected.",
                instruction)

        NormalizeKnowledgeDocumentsForManagementIfNeeded()

        Do
            Dim entries As List(Of KnowledgeDocumentEntry) = ParseKnowledgeDocuments()

            If entries.Count = 0 Then
                Exit Do
            End If

            Dim selection As KnowledgeDocumentManagerResult =
                ShowKnowledgeDocumentManagerDialog(entries, selectedDocumentNumbers, nextInstruction)

            If selection.Action = KnowledgeDocumentManagerAction.None Then
                Exit Do
            End If

            Dim affectedCount As Integer = 0

            Select Case selection.Action
                Case KnowledgeDocumentManagerAction.CompactSelected
                    affectedCount = Await CompactKnowledgeDocumentsAsync(selection.SelectedDocumentNumbers)

                Case KnowledgeDocumentManagerAction.DeleteSelected
                    affectedCount = DeleteKnowledgeDocuments(selection.SelectedDocumentNumbers)

                Case KnowledgeDocumentManagerAction.EditSelected
                    affectedCount = EditKnowledgeDocuments(selection.SelectedDocumentNumbers)

                Case KnowledgeDocumentManagerAction.IndexSelected
                    affectedCount = Await MakeKnowledgeDocumentsSearchableAsync(selection.SelectedDocumentNumbers)
            End Select

            totalAffectedCount += affectedCount

            If String.IsNullOrWhiteSpace(_knowledgeContent) Then
                Exit Do
            End If

            Dim remainingNumbers As New HashSet(Of Integer)(ParseKnowledgeDocuments().Select(Function(x) x.Number))
            selectedDocumentNumbers =
                selection.SelectedDocumentNumbers.
                    Where(Function(x) remainingNumbers.Contains(x)).
                    ToList()

            nextInstruction = "Choose another action, or close the selector when you are finished."
        Loop

        Return totalAffectedCount
    End Function

    Private Async Function TryHandleKnowledgeCompactionOpportunityAsync(response As String,
                                                                        originalUserText As String,
                                                                        toolTriggerDetected As Boolean) As Task(Of Boolean)
        If Not ShouldOfferKnowledgeCompaction(response) Then
            Return False
        End If

        RemoveAssistantThinking()

        If String.IsNullOrWhiteSpace(response) Then
            AppendSystemMessage("The AI returned an empty response. This can indicate that the local knowledge is too large.")
        Else
            AppendSystemMessage("The AI returned a 429-style response. This can indicate that the local knowledge is too large.")
        End If

        Dim openManager As Integer =
            ShowCustomYesNoBox(
                "Do you want to open the knowledge document manager now? You can compact, edit, or delete selected knowledge documents.",
                "Yes, manage knowledge",
                "No, not now",
                $"{AN} - Manage Knowledge")

        If openManager <> 1 Then
            Return True
        End If

        Dim affectedCount As Integer =
            Await ManageKnowledgeDocumentsAsync(
                "Select one or more knowledge documents. Compact uses the same prompt for all selected documents.")

        BringDiscussFormToFront()

        If affectedCount <= 0 Then
            Return True
        End If

        Dim retryPrompt As Integer =
            ShowCustomYesNoBox(
                "Knowledge management is complete. Do you want to retry your last prompt now?",
                "Yes, retry",
                "No, not now",
                $"{AN} - Retry Prompt")

        If retryPrompt = 1 Then
            ShowAssistantThinking()
            Await SendAsync(originalUserText, toolTriggerDetected)
        End If

        Return True
    End Function

#End Region


#Region "Chat Actions"

    ''' <summary>
    ''' Captures the user's message, detects (t) trigger, adds it to history, and starts asynchronous LLM processing.
    ''' </summary>
    Private Async Sub OnSend(sender As Object, e As EventArgs)
        Dim userText = _txtInput.Text.Trim()
        If userText.Length = 0 Then Return

        ' Detect and strip explicit ToolTrigger "(t)" from user prompt
        Dim explicitToolTriggerDetected As Boolean = False
        If userText.IndexOf(ToolTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
            explicitToolTriggerDetected = True
            userText = userText.Replace(ToolTrigger, "").Trim()

            If String.IsNullOrWhiteSpace(userText) Then
                _txtInput.Text = ToolTrigger
                Return
            End If
        End If

        If Not TryBeginExclusive("sending a message") Then
            Return
        End If

        Try
            Dim promptToStore As String = If(explicitToolTriggerDetected, $"{ToolTrigger} {userText}".Trim(), userText)

            Try
                My.Settings.LastPromptDiscussInky = promptToStore
                My.Settings.Save()
            Catch
            End Try

            AppendUserHtml(userText)
            _history.Add(("user", userText))
            _txtInput.Clear()
            ShowAssistantThinking()

            Await SendAsync(userText, explicitToolTriggerDetected)
        Finally
            EndExclusive()
        End Try
    End Sub

    ''' <summary>
    ''' Clears transcript and history, then regenerates the welcome sequence.
    ''' </summary>
    Private Async Sub OnClear(sender As Object, e As EventArgs)
        Try
            _history.Clear()
            ClearCurrentActiveDialogueArchive()
            InitializeChatHtml()
            My.Settings.DiscussLastChat = ""
            My.Settings.DiscussLastChatHtml = ""
            My.Settings.DiscussLastSessionStateXml = ""
            My.Settings.Save()
            UpdateWindowTitle()
            Await SafeGenerateWelcomeAsync().ConfigureAwait(False)
        Catch
        Finally
            Ui(Sub() _txtInput.Focus())
        End Try
    End Sub

    ''' <summary>
    ''' Inserts the selected chat text into the active Word document at the current selection or cursor.
    ''' </summary>
    Private Sub OnInsertSelectionToDoc(sender As Object, e As EventArgs)
        Try
            Dim selectedChatText As String = GetSelectedChatText()
            If String.IsNullOrWhiteSpace(selectedChatText) Then
                AppendSystemMessage("Select text in the discussion thread first.")
                Return
            End If

            Dim app As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
            If app Is Nothing OrElse Not Globals.ThisAddIn.IsDocumentEditable(silent:=True) Then
                AppendSystemMessage("Open an editable Word document first.")
                Return
            End If

            Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
            Dim sel As Microsoft.Office.Interop.Word.Selection = Nothing

            Try
                doc = app.ActiveDocument
                sel = app.Selection
            Catch
            End Try

            If doc Is Nothing OrElse sel Is Nothing Then
                AppendSystemMessage("Open an editable Word document first.")
                Return
            End If

            If app.ActiveWindow Is Nothing OrElse
               app.ActiveWindow.Type <> Microsoft.Office.Interop.Word.WdWindowType.wdWindowDocument Then
                AppendSystemMessage("Open an editable Word document first.")
                Return
            End If

            sel.TypeText(selectedChatText)
            AppendSystemMessage("Selected chat text inserted into the active document.")
        Catch ex As Exception
            AppendSystemMessage($"Error inserting selected chat text into document: {ex.Message}")
        End Try
    End Sub

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
        If Not ConfirmCloseWhenKnowledgePersisted() Then
            Return
        End If

        Me.Close()
    End Sub


    Private Sub OnInputMouseWheel(sender As Object, e As System.Windows.Forms.MouseEventArgs)
        Try
            If System.Windows.Forms.Control.ModifierKeys <> System.Windows.Forms.Keys.Control Then
                Return
            End If

            Dim oldFont As System.Drawing.Font = _txtInput.Font
            Dim newSize As Single = oldFont.Size

            If e.Delta > 0 Then
                newSize += 1.0F
            ElseIf e.Delta < 0 Then
                newSize -= 1.0F
            End If

            newSize = System.Math.Max(8.0F, System.Math.Min(24.0F, newSize))

            If newSize <> oldFont.Size Then
                _txtInput.Font = New System.Drawing.Font(oldFont.FontFamily, newSize, oldFont.Style, oldFont.Unit)
                oldFont.Dispose()
            End If

        Catch ex As System.Exception
            AppendSystemMessage($"Error changing input font size: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Handles slash-triggered prompt library insertion for the DiscussInky input box.
    ''' </summary>
    Private Sub OnInputKeyPress(sender As Object, e As KeyPressEventArgs)
        If e.KeyChar <> "/"c Then Return
        If Not _context.INI_PromptLib Then Return

        Dim slashAction As SharedMethods.PromptLibrarySlashAction =
            SharedMethods.HandlePromptLibrarySlash(
                _txtInput,
                _context.INI_PromptLibPath,
                _context.INI_PromptLibPathLocal,
                _context,
                My.Settings.LastPromptDiscussInky
            )

        If slashAction <> SharedMethods.PromptLibrarySlashAction.NotTriggered Then
            e.Handled = True
        End If
    End Sub

    ''' <summary>
    ''' Handles Enter/Escape shortcuts for sending and closing.
    ''' </summary>
    Private Sub OnInputKeyDown(sender As Object, e As KeyEventArgs)
        If e.Control AndAlso e.KeyCode = Keys.P Then
            Dim lastPrompt As String = My.Settings.LastPromptDiscussInky

            If Not String.IsNullOrWhiteSpace(lastPrompt) Then
                Dim insertionIndex As Integer = _txtInput.SelectionStart
                Dim selectionLength As Integer = _txtInput.SelectionLength

                Dim newText As String =
                    _txtInput.Text.Remove(insertionIndex, selectionLength).Insert(insertionIndex, lastPrompt)

                _txtInput.Text = newText
                _txtInput.SelectionStart = insertionIndex + lastPrompt.Length
                _txtInput.SelectionLength = 0
            End If

            e.SuppressKeyPress = True
            e.Handled = True
            Return
        End If

        If e.KeyCode = Keys.Enter AndAlso Not e.Shift Then
            e.SuppressKeyPress = True
            OnSend(Me, EventArgs.Empty)
        ElseIf e.KeyCode = Keys.Escape Then
            e.SuppressKeyPress = True
            e.Handled = True
            OnClose(Me, EventArgs.Empty)
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
            sb.Append($" | Knowledge: {GetKnowledgeDisplayName()}")
        ElseIf _attachedIndexes.Count = 0 Then
            sb.Append(" | Knowledge: None loaded")
        End If

        ' Attached searchable indexes are an equivalent knowledge source, so list them here too.
        If _attachedIndexes.Count > 0 Then
            Dim indexNames As String =
                String.Join(", ", _attachedIndexes.Select(Function(x) DirectCast(x, DiscussIndexRef).DisplayName))
            sb.Append($" | Searchable index(es): {indexNames}")
        End If

        ' Knowledge store hint
        If Not String.IsNullOrEmpty(Globals.ThisAddIn._context.INI_KnowledgeStorePath) OrElse
           Not String.IsNullOrEmpty(Globals.ThisAddIn._context.INI_KnowledgeStorePathLocal) Then
            sb.Append($" | Type '(kb)' to search all stores, '(kb:storename)' for a specific store, or '(kb:tag:...)' for tagged documents")
        End If

        ' ToolTrigger hint
        Dim toolTriggerAvailable As Boolean =
            SharedMethods.HasToolingCapableSpecialTaskModel(_context, _context.INI_AlternateModelPath, "ToolDefaultModel")

        If toolTriggerAvailable Then
            sb.Append($" | Type '{ToolTrigger}' in your prompt to use the configured {Globals.ThisAddIn.ToolFriendlyName.ToLower} model for a single request.")
        End If

        If _context.INI_PromptLib Then
            sb.Append(" | Type '/' at the start of a prompt or after whitespace to insert a prompt from the prompt library.")
        End If

        AppendSystemMessage(sb.ToString())

        ' Show persona library hint if no paths are configured
        If _noPersonaLibraryConfigured Then
            AppendSystemMessage("No persona library is configured — using the default discussion partner. " &
                               "To add custom personas, define 'DiscussInkyPath' or 'DiscussInkyPathLocal' in your configuration file " &
                               "(sample files are available via 'Get Sample Files' in the settings menu).")
        End If
    End Sub

    ''' <summary>
    ''' Requests a short persona-aware welcome message from the LLM.
    ''' </summary>
    Private Async Function GenerateWelcomeAsync() As Task
        Dim langName = Globals.ThisAddIn.GetWordDefaultInterfaceLanguage()
        Dim partOfDay = GetPartOfDay()
        Dim dateContext = GetDateContext()
        Dim randomWord = GetRandomModifier()
        Dim locationContext = GetLocationContext()
        Dim languageInstruction = GetLanguageInstruction()

        Dim systemPrompt As String

        If String.IsNullOrWhiteSpace(_knowledgeContent) AndAlso _attachedIndexes.Count = 0 Then
            systemPrompt = $"{dateContext} Generate a brief, friendly {langName} welcome that {randomWord} references it is {partOfDay} now. " &
                           "Tell the user they should load a knowledge document using the 'Load Knowledge' button (button name always in English) to start a discussion. " &
                           $"You are ready to discuss any knowledge they provide. One short sentence, not talkative. {languageInstruction} "
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

            Dim knowledgeStatusText As String
            If Not String.IsNullOrWhiteSpace(_knowledgeContent) AndAlso _attachedIndexes.Count > 0 Then
                knowledgeStatusText = "A knowledge base has been loaded (it may contain multiple documents or sections) and one or more searchable indexes are attached."
            ElseIf _attachedIndexes.Count > 0 Then
                knowledgeStatusText = "One or more searchable indexes are attached and will be searched for each message."
            Else
                knowledgeStatusText = "A knowledge base has been loaded (it may contain multiple documents or sections)."
            End If

            systemPrompt = $"{dateContext} {locationContext} Generate a brief, friendly {langName} welcome that {randomWord} references it is {partOfDay} now. " &
                           $"{knowledgeStatusText}{personaContext}{missionContext} " &
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
    ''' Supports one-shot ToolTrigger "(t)" for a single request using the ToolDefaultModel.
    ''' Also supports implicit "(t)" behavior when Enable Tooling is checked, the current model
    ''' does not support tooling, and a tooling-capable ToolDefaultModel exists.
    ''' </summary>
    ''' <param name="userText">User's message text.</param>
    ''' <param name="toolTriggerDetected">True if the user included "(t)" in their prompt.</param>
    Private Async Function SendAsync(userText As String, Optional toolTriggerDetected As Boolean = False) As Task
        Try
            Dim explicitToolTriggerDetected As Boolean = toolTriggerDetected
            Dim restoreUserText As String = If(explicitToolTriggerDetected, $"{ToolTrigger} {userText}".Trim(), userText)

            Dim currentConfig As ModelConfig = Nothing
            If _alternateModelSelected AndAlso _alternateModelConfig IsNot Nothing Then
                currentConfig = _alternateModelConfig
            Else
                currentConfig = SharedMethods.GetCurrentConfig(_context)
            End If

            Dim supportsCurrentModelTooling As Boolean = SharedMethods.ModelSupportsTooling(currentConfig)
            Dim supportsToolTrigger As Boolean =
                SharedMethods.HasToolingCapableSpecialTaskModel(_context, _context.INI_AlternateModelPath, "ToolDefaultModel")

            Dim autoToolTriggerFromCheckbox As Boolean =
                _chkEnableTooling.Checked AndAlso
                Not supportsCurrentModelTooling AndAlso
                supportsToolTrigger

            toolTriggerDetected = explicitToolTriggerDetected OrElse autoToolTriggerFromCheckbox

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

            ' Inject InkyMemory into system prompt if enabled
            If _chkInkyMemory.Checked Then
                Dim memoryContent = SharedMethods.ReadInkyMemory(_context.INI_InkyMemoryCap)
                systemPrompt &= vbLf & _context.SP_Add_InkyMemory
                If Not String.IsNullOrWhiteSpace(memoryContent) Then
                    systemPrompt &= vbLf & "<INKY_MEMORY_CURRENT>" & vbLf & memoryContent & vbLf & "</INKY_MEMORY_CURRENT>"
                End If
            End If

            ' (kb) / (kb:...) trigger: Supplement with knowledge store results
            Dim kbContext As String = Nothing
            Dim cleanedUserText = userText
            If KnowledgeTriggerHelper.HasKnowledgeTrigger(cleanedUserText) Then
                Try
                    Dim kbRequest = KnowledgeTriggerHelper.TryParseKnowledgeTrigger(cleanedUserText)
                    If kbRequest IsNot Nothing Then
                        Dim strippedUserText = KnowledgeTriggerHelper.StripKnowledgeTrigger(cleanedUserText, kbRequest)
                        Dim knowledgeTaskPrompt As String = strippedUserText.Trim()

                        If String.IsNullOrWhiteSpace(strippedUserText) Then
                            If Not String.IsNullOrWhiteSpace(kbRequest.SearchQuery) Then
                                cleanedUserText = kbRequest.SearchQuery.Trim()
                            ElseIf kbRequest.Tags IsNot Nothing AndAlso kbRequest.Tags.Length > 0 Then
                                cleanedUserText = "Answer based on the provided Knowledge Store content, focusing on: " &
                                      String.Join(", ", kbRequest.Tags)
                            ElseIf Not String.IsNullOrWhiteSpace(kbRequest.StoreName) Then
                                cleanedUserText = "Answer based on the provided Knowledge Store content from store '" &
                                      kbRequest.StoreName & "'."
                            Else
                                cleanedUserText = "Answer based on the provided Knowledge Store content."
                            End If
                        Else
                            cleanedUserText = strippedUserText
                        End If

                        Dim kbResolveOptions As KnowledgeTriggerHelper.KnowledgeResolveOptions = Nothing
                        If Not String.IsNullOrWhiteSpace(knowledgeTaskPrompt) Then
                            kbResolveOptions = New KnowledgeTriggerHelper.KnowledgeResolveOptions With {
                    .TaskPrompt = knowledgeTaskPrompt,
                    .IncludeRelevantExtracts = True,
                    .IncludeFullDocumentContent = False
                }
                        End If

                        Dim kbSplash As New SharedMethods.SplashScreen("Querying Knowledge Store...   ")
                        kbSplash.Show()
                        System.Windows.Forms.Application.DoEvents()

                        Dim kbResolved As (Content As String, StatusMessage As String)
                        Try
                            kbResolved = Await KnowledgeTriggerHelper.ResolveKnowledgeAsync(kbRequest, _context, kbResolveOptions)
                        Finally
                            If kbSplash.InvokeRequired Then
                                kbSplash.Invoke(Sub()
                                                    kbSplash.Close()
                                                    kbSplash.Dispose()
                                                End Sub)
                            Else
                                kbSplash.Close()
                                kbSplash.Dispose()
                            End If
                        End Try

                        If Not String.IsNullOrWhiteSpace(kbResolved.Content) Then
                            kbContext = kbResolved.Content

                            systemPrompt &= " The following documents from the user's knowledge store are provided as reference material. " &
                                "Use them to answer the user's question. " &
                                "When citing information, ALWAYS prefer the original source file link over the wiki page link. " &
                                "If a KSDOCUMENT element provides a sourcePath attribute and it is non-empty, ALWAYS cite it as [Source](sourcePath). " &
                                "Only fall back to wikiPath if no sourcePath is available for that document. " &
                                "Do not invent links and do not fabricate paths. Use only the paths explicitly provided in the KSDOCUMENT metadata."

                            AppendSystemMessage($"Knowledge store: {kbResolved.StatusMessage}")
                        Else
                            AppendSystemMessage(If(String.IsNullOrWhiteSpace(kbResolved.StatusMessage),
                                       "No documents found in the Knowledge Store.",
                                       $"Knowledge store: {kbResolved.StatusMessage}"))
                        End If
                    End If
                Catch ex As Exception
                    AppendSystemMessage($"Knowledge store query failed: {ex.Message}")
                End Try
            End If

            ' Build user prompt with knowledge and context
            Dim sb As New StringBuilder()

            sb.AppendLine("User message:")
            sb.AppendLine(cleanedUserText)
            sb.AppendLine()

            ' Include full knowledge document without truncation for smaller docs
            If Not String.IsNullOrWhiteSpace(_knowledgeContent) Then
                sb.AppendLine("<Knowledge Base>")
                Dim knowledgeText = _knowledgeContent
                sb.AppendLine(knowledgeText)
                sb.AppendLine("</Knowledge Base>")
                sb.AppendLine()
            End If

            ' Include the most relevant excerpts retrieved from attached semantic indexes, giving the
            ' user detailed, per-index feedback while the (potentially slow) retrieval runs.
            Dim indexExcerpts As String
            If _attachedIndexes.Count > 0 Then
                UpdateAssistantThinking($"Searching {_attachedIndexes.Count:N0} attached index(es) for relevant material ...")
                indexExcerpts = Await BuildIndexExcerptsAsync(
                    cleanedUserText,
                    BuildConversationForAutoResponder(),
                    Sub(status As String) UpdateAssistantThinking(status))

                If Not String.IsNullOrWhiteSpace(indexExcerpts) Then
                    UpdateAssistantThinking($"Relevant material found. {_currentPersonaName} is now thinking ...")
                Else
                    UpdateAssistantThinking($"No matching material found in the attached index(es). {_currentPersonaName} is now thinking ...")
                End If
            Else
                indexExcerpts = Await BuildIndexExcerptsAsync(cleanedUserText, BuildConversationForAutoResponder())
            End If
            If Not String.IsNullOrWhiteSpace(indexExcerpts) Then
                sb.AppendLine("<Indexed Sources>")
                sb.AppendLine("The following are the most relevant excerpts retrieved from attached indexed documents for this message. " &
                              "Treat them as authoritative source material, do not invent content beyond them, and cite the document names where appropriate.")
                sb.AppendLine(indexExcerpts)
                sb.AppendLine("</Indexed Sources>")
                sb.AppendLine()
            End If

            ' Append knowledge store results (supplemental to manually loaded knowledge)
            If Not String.IsNullOrWhiteSpace(kbContext) Then
                sb.AppendLine("<Knowledge Store Results>")
                sb.AppendLine("The following documents from the user's knowledge store are provided as reference material. " &
                              "Use them as additional reference material alongside any loaded knowledge. " &
                              "When citing information, ALWAYS prefer the original source file link over the wiki page link. " &
                              "If a KSDOCUMENT element provides a sourcePath attribute and it is non-empty, ALWAYS cite it as [Source](sourcePath). " &
                              "Only fall back to wikiPath if no sourcePath is available for that document. " &
                              "Do not invent links and do not fabricate paths. Use only the paths explicitly provided in the KSDOCUMENT metadata.")
                sb.AppendLine(kbContext)
                sb.AppendLine("</Knowledge Store Results>")
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

            ' ──────────────────────────────────────────────────────────────
            ' ToolTrigger "(t)" - One-Shot Tooling Model
            ' Also used implicitly when Enable Tooling is checked and only ToolDefaultModel supports tooling
            ' ──────────────────────────────────────────────────────────────
            Dim toolTriggerConfig As ModelConfig = Nothing

            If toolTriggerDetected Then
                If Not SharedMethods.TryGetSpecialTaskModelConfig(
                    _context,
                    _context.INI_AlternateModelPath,
                    "ToolDefaultModel",
                    toolTriggerConfig) Then

                    RemoveAssistantThinking()
                    AppendSystemMessage($"The {ToolTrigger} trigger was requested, but no model with 'ToolDefaultModel=True' was found in the alternate model configuration. Please add a ToolDefaultModel entry to your configuration file.")
                    Ui(Sub() _txtInput.Text = restoreUserText)
                    Return
                End If

                If Not SharedMethods.ModelSupportsTooling(toolTriggerConfig) Then
                    RemoveAssistantThinking()
                    AppendSystemMessage($"The {ToolTrigger} trigger found a ToolDefaultModel, but it does not support {Globals.ThisAddIn.ToolFriendlyName.ToLower}. Please check the model's APICall_ToolInstructions setting.")
                    Ui(Sub() _txtInput.Text = restoreUserText)
                    Return
                End If

                ' Ensure tools are selected
                If _selectedToolsForChat Is Nothing OrElse _selectedToolsForChat.Count = 0 Then
                    _selectedToolsForChat = Globals.ThisAddIn.SelectDiscussInkyToolsForSession(forceDialog:=False)

                    If _selectedToolsForChat Is Nothing OrElse _selectedToolsForChat.Count = 0 Then
                        RemoveAssistantThinking()
                        AppendSystemMessage($"The {ToolTrigger} trigger requires {Globals.ThisAddIn.ToolFriendlyName.ToLower} to be selected. Please select at least one tool and try again.")
                        Ui(Sub() _txtInput.Text = restoreUserText)
                        Return
                    End If
                End If

                ' Execute via tooling loop with one-shot ToolDefaultModel config
                Dim hideLog As Boolean = Not _chkShowToolingLog.Checked
                Await _modelSemaphore.WaitAsync().ConfigureAwait(False)
                Dim backupConfig As ModelConfig = Nothing
                Try
                    backupConfig = SharedMethods.GetCurrentConfig(_context)
                    SharedMethods.ApplyModelConfig(_context, toolTriggerConfig)

                    Dim answer = Await Globals.ThisAddIn.ExecuteToolingLoop(
                        systemPrompt,
                        userText,
                        _selectedToolsForChat,
                        True,
                        fullPromptOverride:=sb.ToString(),
                        hideSplash:=True,
                        hideLogWindow:=hideLog).ConfigureAwait(False)

                    answer = If(answer, "").Trim()

                    If Await TryHandleKnowledgeCompactionOpportunityAsync(answer, userText, toolTriggerDetected) Then
                        Return
                    End If

                    ' Guard against an empty/whitespace response so the chat does not render a
                    ' bare persona line. Surface it as a system message and restore the prompt.
                    If String.IsNullOrWhiteSpace(answer) Then
                        RemoveAssistantThinking()
                        AppendSystemMessage("The model returned an empty response. Please try again.")
                        Ui(Sub() _txtInput.Text = restoreUserText)

                        If _history.Count > 0 AndAlso _history(_history.Count - 1).Role = "user" Then
                            _history.RemoveAt(_history.Count - 1)
                        End If

                        Return
                    End If

                    ' Process InkyMemory updates from LLM response (if enabled)
                    If _chkInkyMemory.Checked Then
                        answer = SharedMethods.ProcessInkyMemoryResponse(answer, _context.INI_InkyMemoryCap)
                    End If

                    RemoveAssistantThinking()
                    AppendAssistantMarkdown(answer)
                    _history.Add(("assistant", answer))

                    PersistChatHtml()
                    PersistTranscriptLimited()
                Finally
                    If backupConfig IsNot Nothing Then
                        SharedMethods.RestoreDefaults(_context, backupConfig)
                    End If
                    _modelSemaphore.Release()
                End Try

                Return
            End If

            ' ──────────────────────────────────────────────────────────────
            ' Standard LLM call (existing behavior)
            ' If the current model supports tooling and Enable Tooling is checked,
            ' CallLlmWithSelectedModelAsync already handles that path.
            ' ──────────────────────────────────────────────────────────────
            Dim sw = Stopwatch.StartNew()
            Dim stdAnswer = Await CallLlmWithSelectedModelAsync(systemPrompt, sb.ToString())
            sw.Stop()

            stdAnswer = If(stdAnswer, "").Trim()

            If Await TryHandleKnowledgeCompactionOpportunityAsync(stdAnswer, userText, toolTriggerDetected) Then
                Return
            End If

            ' Guard against an empty/whitespace response so the chat does not render a
            ' bare persona line. Surface it as a system message and restore the prompt.
            If String.IsNullOrWhiteSpace(stdAnswer) Then
                RemoveAssistantThinking()
                AppendSystemMessage("The model returned an empty response. Please try again.")
                Ui(Sub() _txtInput.Text = restoreUserText)

                If _history.Count > 0 AndAlso _history(_history.Count - 1).Role = "user" Then
                    _history.RemoveAt(_history.Count - 1)
                End If

                Return
            End If

            ' Process InkyMemory updates from LLM response (if enabled)
            If _chkInkyMemory.Checked Then
                stdAnswer = SharedMethods.ProcessInkyMemoryResponse(stdAnswer, _context.INI_InkyMemoryCap)
            End If

            RemoveAssistantThinking()
            AppendAssistantMarkdown(stdAnswer)
            _history.Add(("assistant", stdAnswer))

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
                    var lastSelectedText = '';
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
                    function setThinkingText(id, text) {{
                      var el=document.getElementById(id); if(!el) return;
                      var c=el.getElementsByClassName('content');
                      if(c && c.length>0){{ c[0].innerText = text; }}
                      window.scrollTo(0, document.body.scrollHeight);
                    }}
                    function getWindowSelectionText() {{
                      try {{
                        if (window.getSelection) {{
                          return window.getSelection().toString();
                        }}
                        if (document.selection) {{
                          return document.selection.createRange().text;
                        }}
                      }} catch (e) {{
                      }}
                      return '';
                    }}
                    function captureSelection() {{
                      var text = getWindowSelectionText();
                      if (text && text.replace(/\s+/g, ' ').replace(/^\s+|\s+$/g, '').length > 0) {{
                        lastSelectedText = text;
                      }}
                    }}
                    document.onmouseup = captureSelection;
                    document.onkeyup = captureSelection;
                    function getSelectedText() {{
                      captureSelection();
                      return lastSelectedText || '';
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
    Private Sub Chat_DocumentCompleted(sender As System.Object, e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs)
        Try
            If _chat.Document Is Nothing Then Return

            Dim chatRoot As System.Windows.Forms.HtmlElement = _chat.Document.GetElementById("chat")
            If chatRoot Is Nothing Then Return

            _htmlReady = True

            If _htmlQueue.Count > 0 Then
                Dim queuedFragments As New System.Collections.Generic.List(Of String)(_htmlQueue)

                Try
                    For Each frag As String In queuedFragments
                        _chat.Document.InvokeScript("appendMessage", New System.Object() {frag})
                    Next

                    _htmlQueue.Clear()

                Catch ex As System.Exception
                    _htmlReady = False
                    Return
                End Try
            End If

            If _persistAfterHtmlFlush Then
                _persistAfterHtmlFlush = False
                PersistCurrentSessionSettings()
            End If

        Catch ex As System.Exception
            _htmlReady = False
        End Try
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
    ''' Updates the text of the current thinking placeholder, so the user gets meaningful feedback
    ''' (e.g., that attached indexes are being searched) instead of a static 'Thinking...' label.
    ''' No-op when no placeholder is currently shown.
    ''' </summary>
    Private Sub UpdateAssistantThinking(statusText As String)
        If String.IsNullOrEmpty(_lastThinkingId) Then Return
        Ui(Sub()
               Try
                   If _chat.Document IsNot Nothing Then
                       _chat.Document.InvokeScript("setThinkingText", New Object() {_lastThinkingId, statusText})
                   End If
               Catch
               End Try
           End Sub)
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
    Private Sub AppendAssistantMarkdownWithName(md As String,
                                               displayName As String,
                                               Optional forwardToTalkToMe As Boolean = True)
        md = If(md, "")
        Dim body = Markdig.Markdown.ToHtml(Global.SharedLibrary.SharedLibrary.SharedMethods.NormalizeMarkdownForHtmlDisplay(md), _mdPipeline)
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

        If forwardToTalkToMe Then
            ForwardOutputToTalkToMe(displayName, md)
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
                   My.Settings.DiscussLastSessionStateXml = BuildSessionStateXml()
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

        _suppressTalkToMeForwarding = True

        Try
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
        Finally
            _suppressTalkToMeForwarding = False
        End Try
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
                    AppendAutoResponderHtml(responderDisplayName, responderMessage, forwardToTalkToMe:=False)
                    ForwardOutputToTalkToMe(responderDisplayName, responderMessage)
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
                    AppendAssistantMarkdownWithName(chatbotResponse, _currentPersonaName, forwardToTalkToMe:=False)
                    ForwardOutputToTalkToMe(_currentPersonaName, chatbotResponse)
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

        ' Include the most relevant excerpts retrieved from attached semantic indexes.
        Dim indexExcerpts As String = Await BuildIndexExcerptsAsync(GetRetrievalQueryFromHistory(), BuildConversationForAutoResponder())
        If Not String.IsNullOrWhiteSpace(indexExcerpts) Then
            sb.AppendLine("<Indexed Sources>")
            sb.AppendLine("The following are the most relevant excerpts retrieved from attached indexed documents. " &
                          "Treat them as authoritative source material and cite the document names where appropriate.")
            sb.AppendLine(indexExcerpts)
            sb.AppendLine("</Indexed Sources>")
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

        ' Include the most relevant excerpts retrieved from attached semantic indexes.
        Dim indexExcerpts As String = Await BuildIndexExcerptsAsync(GetRetrievalQueryFromHistory(), BuildConversationForAutoResponder())
        If Not String.IsNullOrWhiteSpace(indexExcerpts) Then
            sb.AppendLine("<Indexed Sources>")
            sb.AppendLine("The following are the most relevant excerpts retrieved from attached indexed documents. " &
                          "Treat them as authoritative source material and cite the document names where appropriate.")
            sb.AppendLine(indexExcerpts)
            sb.AppendLine("</Indexed Sources>")
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
    Private Sub AppendAutoResponderHtml(responderName As String,
                                        text As String,
                                        Optional forwardToTalkToMe As Boolean = True)
        Dim body = Markdig.Markdown.ToHtml(Global.SharedLibrary.SharedLibrary.SharedMethods.NormalizeMarkdownForHtmlDisplay(text), _mdPipeline)
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

        If forwardToTalkToMe Then
            ForwardOutputToTalkToMe(responderName, text)
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
            $"{AN} - Sort It Out Discussion", False, Context:=_context)

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
                AppendAssistantMarkdownWithName(mainResponse, mainDisplayName, forwardToTalkToMe:=False)
                ForwardOutputToTalkToMe(mainDisplayName, mainResponse)
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
                    AppendAutoResponderHtml(responderDisplayName, responderMessage, forwardToTalkToMe:=False)
                    ForwardOutputToTalkToMe(responderDisplayName, responderMessage)
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
                    AppendAssistantMarkdownWithName(mainBotResponse, mainDisplayName, forwardToTalkToMe:=False)
                    ForwardOutputToTalkToMe(mainDisplayName, mainBotResponse)
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
    ''' Bypasses tooling since summary is a simple text generation task.
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

            ' Call LLM directly WITHOUT tooling - summary is a simple text task
            ' that should never go through the tooling loop to avoid JSON responses
            Await _modelSemaphore.WaitAsync().ConfigureAwait(False)
            Dim backupConfig As ModelConfig = Nothing
            Dim appliedAlternate As Boolean = False
            Dim useSecondApi As Boolean = False

            Try
                If _alternateModelSelected AndAlso _alternateModelConfig IsNot Nothing Then
                    backupConfig = SharedMethods.GetCurrentConfig(_context)
                    SharedMethods.ApplyModelConfig(_context, _alternateModelConfig)
                    appliedAlternate = True
                    useSecondApi = True
                ElseIf _alternateModelSelected AndAlso _alternateModelConfig Is Nothing AndAlso _context.INI_SecondAPI Then
                    useSecondApi = True
                End If

                ' Direct LLM call - explicitly bypass tooling for summaries
                Dim summaryResult = Await LLM(_context,
                    _context.SP_DiscussThis_SumUp,
                    "<TEXTTOPROCESS>" & discussionText & "</TEXTTOPROCESS>",
                    "",
                    "",
                    0,
                    useSecondApi,
                    True).ConfigureAwait(False)

                RemoveAssistantThinking()

                If String.IsNullOrWhiteSpace(summaryResult) Then
                    AppendSystemMessage("Could not generate summary.")
                    Return
                End If

                ' Convert Markdown to HTML and display
                ShowDiscussionSummaryHtml(summaryResult)

            Finally
                If appliedAlternate AndAlso backupConfig IsNot Nothing Then
                    SharedMethods.RestoreDefaults(_context, backupConfig)
                End If
                _modelSemaphore.Release()
            End Try

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
            Dim htmlText As String = Markdig.Markdown.ToHtml(Global.SharedLibrary.SharedLibrary.SharedMethods.NormalizeMarkdownForHtmlDisplay(summaryMarkdown), _mdPipeline)

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

#Region "Dialogue Archive and Full Session State"

    Private Class DialogueArchiveInfo
        Public Property Name As String
        Public Property FilePath As String
        Public Property SavedAtLocal As DateTime

        Public Overrides Function ToString() As String
            Return $"{Name} - {SavedAtLocal:yyyy-MM-dd HH:mm}"
        End Function
    End Class

    Private Function BuildDialogueArchiveManagerInfoText() As String
        Dim sb As New StringBuilder()
        sb.Append("Stored dialogues and persisted knowledge are kept in: ")
        sb.Append(GetRedInkStorageDirectoryPath())
        sb.Append(".")

        If Not String.IsNullOrWhiteSpace(_activeDialogueArchiveName) Then
            sb.Append(" Current linked archive: ")
            sb.Append(_activeDialogueArchiveName)
            sb.Append(If(IsCurrentDialogueArchiveDirty(), " (modified).", " (unchanged)."))
        End If

        sb.Append(" Select one to restore or delete, store the current dialogue as a new archive, or update the linked archive.")
        Return sb.ToString()
    End Function

    Private Function GetDialogueArchiveDirectoryPath() As String
        Return GetRedInkStorageDirectoryPath()
    End Function

    ''' <summary>
    ''' Root folder for all per-session index copies: %AppData%\redink\di.
    ''' </summary>
    Private Function GetSessionIndexRootPath() As String
        Return Path.Combine(GetRedInkStorageDirectoryPath(), SessionIndexFolderName)
    End Function

    ''' <summary>
    ''' Returns a short, stable per-session id used as the persisted index folder name.
    ''' The id is stored in a durable pointer file (independent of My.Settings, which may be
    ''' cleared) so it remains stable across restarts and crashes. Kept short to respect
    ''' Windows path-length limits.
    ''' </summary>
    Private Function GetOrCreateSessionIndexId() As String
        If Not String.IsNullOrEmpty(_sessionIndexId) Then
            Return _sessionIndexId
        End If

        Dim rootPath As String = GetSessionIndexRootPath()
        Dim pointerPath As String = Path.Combine(rootPath, SessionIndexPointerFileName)

        Try
            If File.Exists(pointerPath) Then
                Dim stored As String = File.ReadAllText(pointerPath, Encoding.UTF8).Trim()
                If stored.Length > 0 AndAlso stored.Length <= 16 Then
                    _sessionIndexId = stored
                    Return _sessionIndexId
                End If
            End If
        Catch
            ' Fall through to generate a fresh id.
        End Try

        _sessionIndexId = Guid.NewGuid().ToString("N").Substring(0, 8)

        Try
            Directory.CreateDirectory(rootPath)
            File.WriteAllText(pointerPath, _sessionIndexId, New System.Text.UTF8Encoding(False))
        Catch
            ' Best-effort; an in-memory id is still usable for this session.
        End Try

        Return _sessionIndexId
    End Function

    ''' <summary>
    ''' Directory holding this session's persisted index copies: %AppData%\redink\di\&lt;sid&gt;\.
    ''' </summary>
    Private Function GetSessionIndexDirectoryPath() As String
        Return Path.Combine(GetSessionIndexRootPath(), GetOrCreateSessionIndexId())
    End Function

    ''' <summary>
    ''' Directory holding a specific archive's index copies: %AppData%\redink\&lt;archiveName&gt;.ix\.
    ''' </summary>
    Private Function GetArchiveIndexDirectoryPath(archiveName As String) As String
        Dim safeName As String = If(archiveName, "").Trim()
        For Each ch In Path.GetInvalidFileNameChars()
            safeName = safeName.Replace(ch, "_"c)
        Next
        If String.IsNullOrWhiteSpace(safeName) Then
            safeName = "dialogue"
        End If
        Return Path.Combine(GetDialogueArchiveDirectoryPath(), safeName & ArchiveIndexFolderSuffix)
    End Function

    ''' <summary>
    ''' Copies the given attached indexes into a target directory using short ordinal file
    ''' names (i0.txt, i1.txt, ...) with an exact-byte copy so offsets and the SHA-256 guard
    ''' stay valid. Returns the mapping of index id to the written relative file name.
    ''' </summary>
    Private Function CopyAttachedIndexesToDirectory(indexes As IEnumerable(Of DiscussIndexRef),
                                                    targetDirectory As String,
                                                    repointActivePath As Boolean) As Dictionary(Of String, String)

        Dim writtenByIndexId As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        If indexes Is Nothing Then Return writtenByIndexId

        Directory.CreateDirectory(targetDirectory)

        Dim ordinal As Integer = 0
        For Each idx In indexes
            Dim indexRef As DiscussIndexRef = idx
            ' Name each copy by the index's unique id so incremental persistence cannot collide
            ' (positional ordinals restart per call and would overwrite earlier copies). Fall back
            ' to the ordinal only when an id is unusable as a file name.
            Dim safeId As String = If(indexRef.Id, "")
            For Each ch In Path.GetInvalidFileNameChars()
                safeId = safeId.Replace(ch, "_"c)
            Next
            If String.IsNullOrWhiteSpace(safeId) Then
                safeId = "i" & ordinal.ToString(System.Globalization.CultureInfo.InvariantCulture)
            End If
            Dim fileName As String = safeId & IndexCopyFileExtension
            Dim destinationPath As String = Path.Combine(targetDirectory, fileName)

            Try
                If Not String.IsNullOrWhiteSpace(indexRef.ActivePath) AndAlso File.Exists(indexRef.ActivePath) Then
                    ' Exact-byte copy: never re-encode index content.
                    File.Copy(indexRef.ActivePath, destinationPath, overwrite:=True)
                    writtenByIndexId(indexRef.Id) = fileName
                    If repointActivePath Then
                        indexRef.ActivePath = destinationPath
                    End If
                End If
            Catch ex As Exception
                AppendSystemMessage($"Failed to copy index '{indexRef.DisplayName}': {ex.Message}")
            End Try

            ordinal += 1
        Next

        Return writtenByIndexId
    End Function

    ''' <summary>
    ''' Deletes an index copy directory (best-effort), used when persistence is turned off
    ''' or an archive is removed, to avoid orphaned copies of potentially sensitive data.
    ''' </summary>
    Private Sub DeleteIndexDirectorySafe(directoryPath As String)
        Try
            If Not String.IsNullOrWhiteSpace(directoryPath) AndAlso Directory.Exists(directoryPath) Then
                Directory.Delete(directoryPath, recursive:=True)
            End If
        Catch ex As Exception
            AppendSystemMessage($"Failed to delete index storage: {ex.Message}")
        End Try
    End Sub

    Private Function GetDialogueArchiveFilePath(archiveName As String) As String
        Dim safeName = If(archiveName, "").Trim()
        For Each ch In Path.GetInvalidFileNameChars()
            safeName = safeName.Replace(ch, "_"c)
        Next
        If String.IsNullOrWhiteSpace(safeName) Then
            safeName = "dialogue"
        End If
        Return Path.Combine(GetDialogueArchiveDirectoryPath(), safeName & DialogueArchiveFileExtension)
    End Function

    Private Function GetArchiveNameFromFilePath(filePath As String) As String
        Dim fileName = Path.GetFileName(filePath)
        If fileName.EndsWith(DialogueArchiveFileExtension, StringComparison.OrdinalIgnoreCase) Then
            Return fileName.Substring(0, fileName.Length - DialogueArchiveFileExtension.Length)
        End If
        Return Path.GetFileNameWithoutExtension(filePath)
    End Function

    Private Function NormalizeKnowledgePathForSettings(pathValue As String) As String
        Dim normalized = If(pathValue, "")
        If normalized.EndsWith(" (directory)", StringComparison.OrdinalIgnoreCase) Then
            normalized = normalized.Substring(0, normalized.Length - " (directory)".Length)
        End If
        If normalized.Equals("(Persisted Knowledge)", StringComparison.OrdinalIgnoreCase) Then
            Return ""
        End If
        Return normalized
    End Function

    Private Function GetSettingStringSafe(settingName As String) As String
        Try
            Dim value = My.Settings(settingName)
            Return If(value, "").ToString()
        Catch
            Return ""
        End Try
    End Function

    Private Shared Function GetXmlAttributeValue(element As XElement,
                                                 attributeName As String,
                                                 Optional defaultValue As String = "") As String
        If element Is Nothing Then Return defaultValue
        Dim attr = element.Attribute(attributeName)
        If attr Is Nothing Then Return defaultValue
        Return attr.Value
    End Function

    Private Shared Function GetXmlAttributeBoolean(element As XElement,
                                                   attributeName As String,
                                                   Optional defaultValue As Boolean = False) As Boolean
        Dim raw = GetXmlAttributeValue(element, attributeName, "")
        Dim result As Boolean
        If Boolean.TryParse(raw, result) Then
            Return result
        End If
        Return defaultValue
    End Function

    Private Shared Function GetXmlAttributeInteger(element As XElement,
                                                   attributeName As String,
                                                   Optional defaultValue As Integer = 0) As Integer
        Dim raw = GetXmlAttributeValue(element, attributeName, "")
        Dim result As Integer
        If Integer.TryParse(raw, result) Then
            Return result
        End If
        Return defaultValue
    End Function

    Private Shared Function ComputeTextHash(text As String) As String
        Try
            Dim raw = Encoding.UTF8.GetBytes(If(text, ""))
            Using sha = System.Security.Cryptography.SHA256.Create()
                Return System.Convert.ToBase64String(sha.ComputeHash(raw))
            End Using
        Catch
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Computes a SHA-256 hash over the exact bytes of a file, used to refuse loading the same
    ''' semantic index twice. Returns an empty string on any failure so callers can skip the guard.
    ''' </summary>
    Private Shared Function ComputeFileSha256(filePath As String) As String
        Try
            If String.IsNullOrWhiteSpace(filePath) OrElse Not File.Exists(filePath) Then
                Return ""
            End If

            Using stream As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Using sha = System.Security.Cryptography.SHA256.Create()
                    Return System.Convert.ToBase64String(sha.ComputeHash(stream))
                End Using
            End Using
        Catch
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' True when there is any knowledge source loaded: inlined plain knowledge or at least one
    ''' attached semantic index (indexes are treated as an equivalent knowledge source).
    ''' </summary>
    Private Function HasLoadedKnowledgeOrIndexes() As Boolean
        Return Not String.IsNullOrWhiteSpace(_knowledgeContent) OrElse _attachedIndexes.Count > 0
    End Function

    Private Function NormalizeSessionStateXmlForComparison(stateXml As String) As String
        If String.IsNullOrWhiteSpace(stateXml) Then Return ""

        Try
            Dim doc = XDocument.Parse(stateXml)
            Dim root = doc.Root
            If root Is Nothing Then
                Return stateXml.Trim()
            End If

            Dim savedAtElement = root.Element("SavedAtUtc")
            If savedAtElement IsNot Nothing Then
                savedAtElement.Remove()
            End If

            Dim archiveNameElement = root.Element("ArchiveName")
            If archiveNameElement IsNot Nothing Then
                archiveNameElement.Remove()
            End If

            Dim activeArchiveElement = root.Element("ActiveArchive")
            If activeArchiveElement IsNot Nothing Then
                activeArchiveElement.Remove()
            End If

            Dim transcriptHtmlElement = root.Element("TranscriptHtml")
            If transcriptHtmlElement IsNot Nothing Then
                transcriptHtmlElement.Remove()
            End If

            Return doc.ToString(SaveOptions.DisableFormatting)
        Catch
            Return stateXml.Trim()
        End Try
    End Function

    Private Function GetCurrentSessionComparisonHash() As String
        Return ComputeTextHash(NormalizeSessionStateXmlForComparison(BuildSessionStateXml()))
    End Function

    Private Sub SetCurrentActiveDialogueArchive(archiveName As String, archiveFilePath As String, baselineHash As String)
        _activeDialogueArchiveName = If(archiveName, "").Trim()
        _activeDialogueArchiveFilePath = If(archiveFilePath, "").Trim()
        _activeDialogueArchiveBaselineHash = If(baselineHash, "").Trim()
    End Sub

    Private Sub ClearCurrentActiveDialogueArchive()
        SetCurrentActiveDialogueArchive("", "", "")
    End Sub

    Private Function HasTrackedDialogueArchive() As Boolean
        Return Not String.IsNullOrWhiteSpace(_activeDialogueArchiveName) AndAlso
               Not String.IsNullOrWhiteSpace(_activeDialogueArchiveFilePath)
    End Function

    Private Function IsCurrentDialogueArchiveDirty() As Boolean
        If Not HasTrackedDialogueArchive() Then Return False
        If String.IsNullOrWhiteSpace(_activeDialogueArchiveBaselineHash) Then Return True
        Return Not String.Equals(_activeDialogueArchiveBaselineHash, GetCurrentSessionComparisonHash(), StringComparison.Ordinal)
    End Function

    Private Function GetCurrentChatInnerHtml() As String
        If Me.IsDisposed Then Return ""

        If Me.InvokeRequired Then
            Try
                Return CStr(Me.Invoke(New Func(Of String)(AddressOf GetCurrentChatInnerHtml)))
            Catch
                Return ""
            End Try
        End If

        Try
            If _chat.Document Is Nothing Then Return ""
            Dim root = _chat.Document.GetElementById("chat")
            If root Is Nothing Then Return ""
            Return If(root.InnerHtml, "")
        Catch
            Return ""
        End Try
    End Function

    Private Function BuildSessionStateXml(Optional archiveName As String = "") As String
        Dim historyElement As New XElement("History")

        For Each entry In _history
            historyElement.Add(
                New XElement("Message",
                    New XAttribute("role", If(entry.Role, "")),
                    If(entry.Content, "")))
        Next

        Dim alternateMode As String = "Primary"
        If _alternateModelSelected Then
            If _alternateModelConfig Is Nothing AndAlso _context.INI_SecondAPI Then
                alternateMode = "LegacySecondApi"
            Else
                alternateMode = "AlternateConfig"
            End If
        End If

        Dim root As New XElement(
            "DiscussInkyState",
            New XAttribute("version", "1"),
            New XElement("ArchiveName", If(archiveName, "")),
            New XElement("SavedAtUtc", DateTime.UtcNow.ToString("o")),
            New XElement(
                "ActiveArchive",
                New XAttribute("name", If(_activeDialogueArchiveName, "")),
                New XAttribute("filePath", If(_activeDialogueArchiveFilePath, "")),
                New XAttribute("baselineHash", If(_activeDialogueArchiveBaselineHash, ""))),
            New XElement(
                "Persona",
                New XAttribute("name", If(_currentPersonaName, "")),
                New XElement("Prompt", If(_currentPersonaPrompt, ""))),
            New XElement(
                "Mission",
                New XAttribute("name", If(_currentMissionName, "")),
                New XElement("Prompt", If(_currentMissionPrompt, ""))),
            New XElement(
                "Knowledge",
                New XAttribute("path", If(_knowledgeFilePath, "")),
                New XAttribute("persisted", _chkPersistKnowledge.Checked),
                If(_knowledgeContent, "")),
            New XElement(
                "Flags",
                New XAttribute("includeActiveDoc", _chkIncludeActiveDoc.Checked),
                New XAttribute("enableTooling", _chkEnableTooling.Checked),
                New XAttribute("advancedTools", _chkAdvancedTools.Checked),
                New XAttribute("showToolingLog", _chkShowToolingLog.Checked),
                New XAttribute("inkyMemory", _chkInkyMemory.Checked)),
            New XElement(
                "AlternateModel",
                New XAttribute("selected", _alternateModelSelected),
                New XAttribute("mode", alternateMode),
                New XAttribute("displayName", If(_alternateModelDisplayName, ""))),
            New XElement(
                "ToolSelection",
                New XAttribute("main", GetSettingStringSafe("SelectedMainToolNames")),
                New XAttribute("advanced", GetSettingStringSafe("SelectedAdvancedToolNames"))),
                       New XElement(
                "Ui",
                New XAttribute("splitterDistance", _splitChat.SplitterDistance)),
            New XElement("TranscriptHtml", GetCurrentChatInnerHtml()),
            BuildIndexesElement(archiveName),
            historyElement)

        Return New XDocument(root).ToString(SaveOptions.DisableFormatting)
    End Function

    ''' <summary>
    ''' Builds the &lt;Indexes&gt; element describing attached semantic indexes. When an archive
    ''' name is supplied, the index files are copied into that archive's sidecar folder
    ''' (&lt;archiveName&gt;.ix) so the archive is self-contained; the recorded file names are
    ''' relative to that folder. Otherwise (session-settings persistence) the current active
    ''' paths are recorded so the live session can be restored in place.
    ''' </summary>
    Private Function BuildIndexesElement(archiveName As String) As XElement
        Dim indexesElement As New XElement("Indexes")
        If _attachedIndexes.Count = 0 Then
            Return indexesElement
        End If

        Dim copyToArchive As Boolean = Not String.IsNullOrWhiteSpace(archiveName)
        Dim writtenByIndexId As Dictionary(Of String, String) = Nothing

        If copyToArchive Then
            Dim archiveIndexDir As String = GetArchiveIndexDirectoryPath(archiveName)
            ' Rewrite the sidecar from scratch so stale copies are not left behind.
            DeleteIndexDirectorySafe(archiveIndexDir)
            writtenByIndexId = CopyAttachedIndexesToDirectory(_attachedIndexes, archiveIndexDir, repointActivePath:=False)
        End If

        For Each idx In _attachedIndexes
            Dim indexRef As DiscussIndexRef = idx
            Dim fileValue As String

            If copyToArchive Then
                Dim relativeName As String = Nothing
                If writtenByIndexId Is Nothing OrElse Not writtenByIndexId.TryGetValue(indexRef.Id, relativeName) Then
                    Continue For
                End If
                fileValue = relativeName
            Else
                fileValue = If(indexRef.ActivePath, "")
            End If

            indexesElement.Add(
                New XElement("Index",
                    New XAttribute("id", If(indexRef.Id, "")),
                    New XAttribute("name", If(indexRef.DisplayName, "")),
                    New XAttribute("file", fileValue),
                    New XAttribute("originalPath", If(indexRef.OriginalPath, "")),
                    New XAttribute("sha", If(indexRef.ContentSha256, "")),
                    New XAttribute("archived", copyToArchive)))
        Next

        Return indexesElement
    End Function

    ''' <summary>
    ''' Writes the running-session index references to a durable file under %AppData%\redink\di,
    ''' independent of My.Settings (which may be cleared). Enables restoring attached indexes for
    ''' a non-archived session after a restart or crash.
    ''' </summary>
    Private Sub SaveSessionIndexStateDurably()
        Dim rootPath As String = GetSessionIndexRootPath()
        Dim statePath As String = Path.Combine(rootPath, SessionIndexStateFileName)

        Try
            Directory.CreateDirectory(rootPath)

            If _attachedIndexes.Count = 0 Then
                If File.Exists(statePath) Then
                    File.Delete(statePath)
                End If
                Return
            End If

            ' archiveName empty => records current absolute ActivePath values (archived=false).
            Dim indexesElement As XElement = BuildIndexesElement("")
            Dim stateDoc As New XDocument(indexesElement)
            stateDoc.Save(statePath)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Restores attached indexes for a non-archived running session from the durable state file.
    ''' Missing files are skipped by RestoreAttachedIndexesFromXml.
    ''' </summary>
    Private Sub RestoreSessionIndexStateFromDurableFile()
        Dim statePath As String = Path.Combine(GetSessionIndexRootPath(), SessionIndexStateFileName)

        If Not File.Exists(statePath) Then
            Return
        End If

        Try
            Dim doc As XDocument = XDocument.Load(statePath)
            RestoreAttachedIndexesFromXml(doc.Root)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try
    End Sub

    Private Sub PersistCurrentSessionSettings(Optional saveImmediately As Boolean = True)
        Try
            PersistTranscriptLimited()
            SaveSessionIndexStateDurably()
            My.Settings.DiscussLastChatHtml = GetCurrentChatInnerHtml()
            My.Settings.DiscussLastSessionStateXml = BuildSessionStateXml()
            My.Settings.DiscussIncludeActiveDoc = _chkIncludeActiveDoc.Checked
            My.Settings.DiscussPersistKnowledge = _chkPersistKnowledge.Checked
            My.Settings.DiscussSelectedPersona = _currentPersonaName
            My.Settings.DiscussSelectedMission = _currentMissionName
            My.Settings.DiscussKnowledgePath = NormalizeKnowledgePathForSettings(_knowledgeFilePath)
            My.Settings.DiscussEnableTooling = _chkEnableTooling.Checked
            If saveImmediately Then
                My.Settings.Save()
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Copies all attached indexes into this session's durable persist folder (di\&lt;sid&gt;)
    ''' and repoints their active paths to the copies, so persisted sessions keep working even
    ''' if the original source files are moved or deleted. Indexes created via "Make Searchable"
    ''' already live in this folder, so copying is a no-op for those.
    ''' </summary>
    Private Sub PersistAttachedIndexes()
        If _attachedIndexes.Count = 0 Then
            Return
        End If

        Dim sessionDir As String = GetSessionIndexDirectoryPath()
        Dim toCopy As List(Of DiscussIndexRef) =
            _attachedIndexes.
                Where(Function(x) Not String.IsNullOrWhiteSpace(x.ActivePath) AndAlso
                                  Not String.Equals(System.IO.Path.GetDirectoryName(x.ActivePath),
                                                    sessionDir, StringComparison.OrdinalIgnoreCase)).
                ToList()

        If toCopy.Count = 0 Then
            Return
        End If

        CopyAttachedIndexesToDirectory(toCopy, sessionDir, repointActivePath:=True)
    End Sub

    ''' <summary>
    ''' Removes this session's durable persist folder (di\&lt;sid&gt;). Only called when the user
    ''' turns persistence off and confirms deletion; guards against removing indexes that are
    ''' still attached by clearing them from the active session first.
    ''' </summary>
    Private Sub DeletePersistedSessionIndexes()
        DeleteIndexDirectorySafe(GetSessionIndexDirectoryPath())
    End Sub

    Private Sub ApplyKnowledgePersistenceFromCurrentState()
        If String.IsNullOrWhiteSpace(_knowledgeContent) Then
            _knowledgeContent = Nothing
            _cachedKnowledgeContent = Nothing
            _cachedKnowledgeFilePath = Nothing
        Else
            _cachedKnowledgeContent = _knowledgeContent
            _cachedKnowledgeFilePath = _knowledgeFilePath
        End If

        Dim persistPath = GetPersistedKnowledgeFilePath()

        Try
            If _chkPersistKnowledge.Checked AndAlso Not String.IsNullOrWhiteSpace(_knowledgeContent) Then
                File.WriteAllText(persistPath, _knowledgeContent, Encoding.UTF8)
            ElseIf File.Exists(persistPath) Then
                File.Delete(persistPath)
            End If
        Catch
        End Try

        UpdatePersistKnowledgeTooltip()
    End Sub

    Private Function TryRestoreArchivedAlternateModel(displayName As String) As Boolean
        If String.IsNullOrWhiteSpace(displayName) Then Return False
        If String.IsNullOrWhiteSpace(_context.INI_AlternateModelPath) Then Return False

        Try
            Dim models = SharedMethods.LoadAlternativeModels(_context.INI_AlternateModelPath, _context, "Alternate Model")
            Dim found = models.FirstOrDefault(
                Function(m) String.Equals(If(m.ModelDescription, ""), displayName, StringComparison.OrdinalIgnoreCase))

            If found Is Nothing Then
                Return False
            End If

            _alternateModelSelected = True
            _alternateModelConfig = found
            _alternateModelDisplayName = displayName
            UpdateAlternateModelButtonText()
            Return True
        Catch
            Return False
        End Try
    End Function

    Private Sub RenderHistoryToCurrentChat()
        For Each msg In _history
            Select Case msg.Role
                Case "user"
                    AppendUserHtml(msg.Content)

                Case "assistant"
                    Dim displayName = _currentPersonaName
                    Dim messageText = msg.Content
                    Dim colonIdx = messageText.IndexOf(": ", StringComparison.Ordinal)

                    If colonIdx > 0 Then
                        Dim possibleDisplayName = messageText.Substring(0, colonIdx)
                        If possibleDisplayName.Contains("(Advocate)") OrElse
                           possibleDisplayName.Contains("(Challenger)") OrElse
                           possibleDisplayName.Contains("(2nd)") Then
                            displayName = possibleDisplayName
                            messageText = messageText.Substring(colonIdx + 2)
                        End If
                    End If

                    AppendAssistantMarkdownWithName(messageText, displayName)

                Case "autoresponder"
                    Dim responderName = "Autoresponder"
                    Dim responderText = msg.Content
                    Dim colonIdx = responderText.IndexOf(": ", StringComparison.Ordinal)

                    If colonIdx > 0 Then
                        responderName = responderText.Substring(0, colonIdx)
                        responderText = responderText.Substring(colonIdx + 2)
                    End If

                    AppendAutoResponderHtml(responderName, responderText)
            End Select
        Next
    End Sub

    Private Function RestoreSessionStateFromXml(stateXml As String,
                                            sourceLabel As String,
                                            Optional announceRestore As Boolean = True,
                                            Optional resetChatHtml As Boolean = True) As Boolean
        If String.IsNullOrWhiteSpace(stateXml) Then Return False

        Try
            Dim doc = XDocument.Parse(stateXml)
            Dim root = doc.Root
            If root Is Nothing OrElse Not root.Name.LocalName.Equals("DiscussInkyState", StringComparison.OrdinalIgnoreCase) Then
                Return False
            End If

            Dim activeArchiveElement = root.Element("ActiveArchive")
            Dim personaElement = root.Element("Persona")
            Dim missionElement = root.Element("Mission")
            Dim knowledgeElement = root.Element("Knowledge")
            Dim flagsElement = root.Element("Flags")
            Dim alternateElement = root.Element("AlternateModel")
            Dim toolSelectionElement = root.Element("ToolSelection")
            Dim uiElement = root.Element("Ui")
            Dim transcriptHtmlElement = root.Element("TranscriptHtml")
            Dim historyElement = root.Element("History")

            SetCurrentActiveDialogueArchive(
                GetXmlAttributeValue(activeArchiveElement, "name", ""),
                GetXmlAttributeValue(activeArchiveElement, "filePath", ""),
                GetXmlAttributeValue(activeArchiveElement, "baselineHash", ""))

            _currentPersonaName = GetXmlAttributeValue(personaElement, "name", DefaultPersonaName)
            Dim restoredPersonaPrompt = ""
            If personaElement IsNot Nothing Then
                Dim promptElement = personaElement.Element("Prompt")
                restoredPersonaPrompt = If(promptElement IsNot Nothing, promptElement.Value, "")
            End If
            _currentPersonaPrompt = If(String.IsNullOrWhiteSpace(restoredPersonaPrompt), DefaultPersonaPrompt, restoredPersonaPrompt)

            _currentMissionName = GetXmlAttributeValue(missionElement, "name", "")
            Dim restoredMissionPrompt = ""
            If missionElement IsNot Nothing Then
                Dim promptElement = missionElement.Element("Prompt")
                restoredMissionPrompt = If(promptElement IsNot Nothing, promptElement.Value, "")
            End If
            _currentMissionPrompt = restoredMissionPrompt

            _history.Clear()
            If historyElement IsNot Nothing Then
                For Each messageElement In historyElement.Elements("Message")
                    Dim role = GetXmlAttributeValue(messageElement, "role", "")
                    If String.IsNullOrWhiteSpace(role) Then Continue For
                    _history.Add((role, messageElement.Value))
                Next
            End If

            _chkIncludeActiveDoc.Checked = GetXmlAttributeBoolean(flagsElement, "includeActiveDoc", False)

            _isUpdatingPersistCheckbox = True
            _chkPersistKnowledge.Checked = GetXmlAttributeBoolean(knowledgeElement, "persisted", False)
            _isUpdatingPersistCheckbox = False

            _chkEnableTooling.Checked = GetXmlAttributeBoolean(flagsElement, "enableTooling", False)
            _chkAdvancedTools.Checked = GetXmlAttributeBoolean(flagsElement, "advancedTools", False)
            _chkShowToolingLog.Checked = GetXmlAttributeBoolean(flagsElement, "showToolingLog", _context.INI_ToolingLogWindow)
            _chkInkyMemory.Checked = GetXmlAttributeBoolean(flagsElement, "inkyMemory", My.Settings.DiscussInkyMemory)
            _lnkEditMemory.Visible = _chkInkyMemory.Checked

            _knowledgeFilePath = GetXmlAttributeValue(knowledgeElement, "path", "")
            _knowledgeContent = If(knowledgeElement IsNot Nothing, knowledgeElement.Value, Nothing)
            ApplyKnowledgePersistenceFromCurrentState()

            RestoreAttachedIndexesFromXml(root.Element("Indexes"))

            Try
                Dim splitterDistance = GetXmlAttributeInteger(uiElement, "splitterDistance", _splitChat.SplitterDistance)
                If splitterDistance > 0 Then
                    _splitChat.SplitterDistance = splitterDistance
                End If
            Catch
            End Try

            Try
                Globals.ThisAddIn.PersistDiscussInkyToolSelection(
                    Globals.ThisAddIn.SplitPersistedToolNames(GetXmlAttributeValue(toolSelectionElement, "main", "")),
                    Globals.ThisAddIn.SplitPersistedToolNames(GetXmlAttributeValue(toolSelectionElement, "advanced", "")),
                    _chkAdvancedTools.Checked)
                _selectedToolsForChat = Nothing
            Catch
            End Try

            Dim alternateRestoreNotice As String = ""
            _alternateModelSelected = False
            _alternateModelConfig = Nothing
            _alternateModelDisplayName = Nothing

            Dim alternateMode = GetXmlAttributeValue(alternateElement, "mode", "Primary")
            Dim alternateDisplayName = GetXmlAttributeValue(alternateElement, "displayName", "")

            If String.Equals(alternateMode, "LegacySecondApi", StringComparison.OrdinalIgnoreCase) Then
                If _context.INI_SecondAPI Then
                    _alternateModelSelected = True
                    _alternateModelConfig = Nothing
                    _alternateModelDisplayName = If(String.IsNullOrWhiteSpace(alternateDisplayName), _context.INI_Model_2, alternateDisplayName)
                Else
                    alternateRestoreNotice = "The archived dialogue used the secondary model, but that model is no longer configured. Primary model is active."
                End If
            ElseIf GetXmlAttributeBoolean(alternateElement, "selected", False) Then
                If Not TryRestoreArchivedAlternateModel(alternateDisplayName) Then
                    alternateRestoreNotice = $"The archived dialogue requested alternate model '{alternateDisplayName}', but it could not be restored. Primary model is active."
                End If
            End If

            UpdateAlternateModelButtonText()
            UpdateWindowTitle()
            UpdateSendButtonText()
            UpdatePersistKnowledgeTooltip()
            UpdateToolingControlsState()

            If resetChatHtml Then
                InitializeChatHtml()
            End If

            Dim transcriptHtml = If(transcriptHtmlElement IsNot Nothing, transcriptHtmlElement.Value, "")
            If Not String.IsNullOrWhiteSpace(transcriptHtml) Then
                AppendHtml(transcriptHtml)
            Else
                RenderHistoryToCurrentChat()
            End If

            If _htmlReady AndAlso _chat.Document IsNot Nothing Then
                PersistCurrentSessionSettings()
            Else
                _persistAfterHtmlFlush = True
            End If

            If announceRestore AndAlso Not String.IsNullOrWhiteSpace(sourceLabel) Then
                AppendSystemMessage($"Dialogue restored: {sourceLabel}")
            End If

            If Not String.IsNullOrWhiteSpace(alternateRestoreNotice) Then
                AppendSystemMessage(alternateRestoreNotice)
            End If

            Return True

        Catch ex As Exception
            AppendSystemMessage($"Failed to restore dialogue archive: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Restores the attached semantic indexes from a session/archive &lt;Indexes&gt; element.
    ''' Archived indexes are resolved relative to the archive's sidecar folder; session indexes
    ''' use their recorded absolute path. Missing files are reported and skipped so restore
    ''' remains robust.
    ''' </summary>
    Private Sub RestoreAttachedIndexesFromXml(indexesElement As XElement)
        _attachedIndexes.Clear()
        _indexConversationState.Clear()

        If indexesElement Is Nothing Then
            Return
        End If

        Dim missing As New List(Of String)()

        For Each indexElement In indexesElement.Elements("Index")
            Dim id As String = GetXmlAttributeValue(indexElement, "id", "")
            Dim displayName As String = GetXmlAttributeValue(indexElement, "name", "")
            Dim fileValue As String = GetXmlAttributeValue(indexElement, "file", "")
            Dim originalPath As String = GetXmlAttributeValue(indexElement, "originalPath", "")
            Dim sha As String = GetXmlAttributeValue(indexElement, "sha", "")
            Dim archived As Boolean = GetXmlAttributeBoolean(indexElement, "archived", False)

            If String.IsNullOrWhiteSpace(fileValue) Then
                Continue For
            End If

            Dim resolvedPath As String
            If archived Then
                Dim archiveName As String = If(_activeDialogueArchiveName, "")
                resolvedPath = System.IO.Path.Combine(GetArchiveIndexDirectoryPath(archiveName), fileValue)
            Else
                resolvedPath = fileValue
            End If

            If String.IsNullOrWhiteSpace(resolvedPath) OrElse Not File.Exists(resolvedPath) Then
                missing.Add(If(String.IsNullOrWhiteSpace(displayName), fileValue, displayName))
                Continue For
            End If

            _attachedIndexes.Add(New DiscussIndexRef() With {
                .Id = If(String.IsNullOrWhiteSpace(id), "i" & Guid.NewGuid().ToString("N").Substring(0, 4), id),
                .DisplayName = displayName,
                .ActivePath = resolvedPath,
                .OriginalPath = originalPath,
                .ContentSha256 = sha
            })
        Next

        If missing.Count > 0 Then
            AppendSystemMessage(
                $"{missing.Count:N0} attached index file(s) could not be found and were skipped: {String.Join(", ", missing)}.")
        End If
    End Sub

    Private Function GetDialogueArchives() As List(Of DialogueArchiveInfo)
        Dim result As New List(Of DialogueArchiveInfo)()
        Dim archiveDir = GetDialogueArchiveDirectoryPath()

        If Not Directory.Exists(archiveDir) Then
            Return result
        End If

        For Each filePath In Directory.GetFiles(archiveDir, "*" & DialogueArchiveFileExtension, SearchOption.TopDirectoryOnly)
            Try
                Dim doc = XDocument.Load(filePath)
                Dim root = doc.Root

                Dim name = GetArchiveNameFromFilePath(filePath)
                Dim savedAtLocal = File.GetLastWriteTime(filePath)

                If root IsNot Nothing Then
                    Dim archiveNameElement = root.Element("ArchiveName")
                    If archiveNameElement IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(archiveNameElement.Value) Then
                        name = archiveNameElement.Value.Trim()
                    End If

                    Dim savedAtElement = root.Element("SavedAtUtc")
                    Dim parsedUtc As DateTime
                    If savedAtElement IsNot Nothing AndAlso DateTime.TryParse(savedAtElement.Value, parsedUtc) Then
                        savedAtLocal = parsedUtc.ToLocalTime()
                    End If
                End If

                result.Add(New DialogueArchiveInfo With {
                    .Name = name,
                    .FilePath = filePath,
                    .SavedAtLocal = savedAtLocal
                })
            Catch
                result.Add(New DialogueArchiveInfo With {
                    .Name = GetArchiveNameFromFilePath(filePath),
                    .FilePath = filePath,
                    .SavedAtLocal = File.GetLastWriteTime(filePath)
                })
            End Try
        Next

        Return result.OrderByDescending(Function(x) x.SavedAtLocal).ToList()
    End Function

    Private Function HasDialogueStateToArchive() As Boolean
        If _history.Count > 0 Then Return True
        If Not String.IsNullOrWhiteSpace(_knowledgeContent) Then Return True
        If Not String.IsNullOrWhiteSpace(_currentMissionName) Then Return True
        If Not String.Equals(_currentPersonaName, DefaultPersonaName, StringComparison.OrdinalIgnoreCase) Then Return True
        Return False
    End Function

    Private Function SaveCurrentDialogueArchive(archiveName As String,
                                               Optional overwriteWithoutPrompt As Boolean = False) As Boolean
        If String.IsNullOrWhiteSpace(archiveName) Then Return False

        Dim trimmedArchiveName = archiveName.Trim()
        Dim filePath = GetDialogueArchiveFilePath(trimmedArchiveName)
        Dim fileAlreadyExists = File.Exists(filePath)

        If fileAlreadyExists AndAlso Not overwriteWithoutPrompt Then
            Dim overwriteAnswer = ShowCustomYesNoBox(
                $"An archived dialogue named '{trimmedArchiveName}' already exists. Do you want to overwrite it?",
                "Yes, overwrite",
                "No, keep existing",
                $"{AN} - Overwrite Dialogue Archive")

            If overwriteAnswer <> 1 Then
                Return False
            End If
        End If

        Dim previousArchiveName = _activeDialogueArchiveName
        Dim previousArchiveFilePath = _activeDialogueArchiveFilePath
        Dim previousBaselineHash = _activeDialogueArchiveBaselineHash

        Try
            Dim archiveDir = GetDialogueArchiveDirectoryPath()
            If Not Directory.Exists(archiveDir) Then
                Directory.CreateDirectory(archiveDir)
            End If

            SetCurrentActiveDialogueArchive(trimmedArchiveName, filePath, "")
            Dim xmlToSave = BuildSessionStateXml(trimmedArchiveName)
            File.WriteAllText(filePath, xmlToSave, Encoding.UTF8)

            PersistCurrentSessionSettings()
            UpdateWindowTitle()

            If fileAlreadyExists Then
                AppendSystemMessage($"Dialogue archive updated: '{trimmedArchiveName}'.")
            Else
                AppendSystemMessage($"Dialogue archived as '{trimmedArchiveName}'.")
            End If

            ' Capture the baseline from the LIVE session state (the same representation the dirty
            ' check uses) AFTER all side effects, so a freshly stored dialogue is never considered
            ' dirty. The archive-format XML written to disk uses a different index representation
            ' (relative sidecar paths), so it must not be used as the comparison baseline.
            SetCurrentActiveDialogueArchive(
                trimmedArchiveName,
                filePath,
                GetCurrentSessionComparisonHash())

            Return True

        Catch ex As Exception
            SetCurrentActiveDialogueArchive(previousArchiveName, previousArchiveFilePath, previousBaselineHash)
            AppendSystemMessage($"Failed to archive dialogue: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function PromptAndSaveCurrentDialogueArchive() As Boolean
        If Not HasDialogueStateToArchive() Then
            AppendSystemMessage("There is no dialogue state to archive.")
            Return False
        End If

        Dim defaultArchiveName = If(_activeDialogueArchiveName, "")

        Dim archiveName = ShowCustomInputBox(
            "Enter a name for the archived dialogue:",
            $"{AN} - Store Dialogue Archive",
            True,
            defaultArchiveName)

        If String.IsNullOrWhiteSpace(archiveName) OrElse archiveName = "ESC" Then
            Return False
        End If

        Return SaveCurrentDialogueArchive(archiveName.Trim())
    End Function

    Private Function PromptToPersistDirtyArchiveBackedDialogue() As Boolean
        If HasTrackedDialogueArchive() Then
            Dim answer = ShowCustomYesNoBox(
                $"The current dialogue is linked to archive '{_activeDialogueArchiveName}' and has changed. Do you want to update that archive or store the dialogue as a new archive?",
                "Update existing archive",
                "Store as new archive",
                $"{AN} - Save Dialogue Changes",
                extraButtonText:="Cancel",
                extraButtonAction:=Sub()
                                   End Sub,
                CloseAfterExtra:=True)

            If answer = 1 Then
                Return SaveCurrentDialogueArchive(_activeDialogueArchiveName, overwriteWithoutPrompt:=True)
            End If

            If answer = 2 Then
                Return PromptAndSaveCurrentDialogueArchive()
            End If

            Return False
        End If

        Return PromptAndSaveCurrentDialogueArchive()
    End Function

    Private Function PromptToSaveCurrentDialogueBeforeSwitch() As Boolean
        If Not HasTrackedDialogueArchive() Then
            Return True
        End If

        If Not IsCurrentDialogueArchiveDirty() Then
            Return True
        End If

        Dim answer = ShowCustomYesNoBox(
            "The current dialogue has changed since it was restored or stored. Do you want to save those changes before switching?",
            "Yes, save changes",
            "No, switch without saving",
            $"{AN} - Switch Dialogue",
            extraButtonText:="Cancel",
            extraButtonAction:=Sub()
                               End Sub,
            CloseAfterExtra:=True)

        If answer = 1 Then
            Return PromptToPersistDirtyArchiveBackedDialogue()
        End If

        If answer = 2 Then
            Return True
        End If

        Return False
    End Function

    Private Function RestoreDialogueArchiveFromFile(filePath As String, displayName As String) As Boolean
        Try
            If String.IsNullOrWhiteSpace(filePath) OrElse Not File.Exists(filePath) Then
                AppendSystemMessage("The selected dialogue archive no longer exists.")
                Return False
            End If

            Dim stateXml = File.ReadAllText(filePath, Encoding.UTF8)
            Dim restored = RestoreSessionStateFromXml(stateXml, displayName, announceRestore:=True)

            If restored Then
                PersistCurrentSessionSettings()
                UpdateWindowTitle()

                ' Capture the baseline from the LIVE restored session (same representation as the
                ' dirty check) AFTER all side effects, including the restore announcement. Using the
                ' on-disk archive XML here would mismatch (different index path representation) and
                ' make the restored dialogue look changed immediately.
                SetCurrentActiveDialogueArchive(
                    displayName,
                    filePath,
                    GetCurrentSessionComparisonHash())
            End If

            Return restored

        Catch ex As Exception
            AppendSystemMessage($"Failed to load dialogue archive: {ex.Message}")
            Return False
        End Try
    End Function

    Private Sub ShowDialogueArchiveManager()
        Using frm As New Form() With {
            .Text = $"{AN} - Dialogue Archive",
            .StartPosition = FormStartPosition.CenterParent,
            .Size = New System.Drawing.Size(780, 460),
            .MinimumSize = New System.Drawing.Size(560, 360),
            .FormBorderStyle = FormBorderStyle.Sizable,
            .Font = New System.Drawing.Font("Segoe UI", 9.0F),
            .AutoScaleDimensions = New System.Drawing.SizeF(96.0F, 96.0F),
            .AutoScaleMode = AutoScaleMode.Dpi,
            .ShowInTaskbar = False
        }
            Try
                frm.Icon = Me.Icon
            Catch
            End Try

            Dim layout As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 3,
                .Padding = New Padding(12)
            }
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            layout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            frm.Controls.Add(layout)

            Dim lblInfo As New Label() With {
                .AutoSize = True,
                .Dock = DockStyle.Top,
                .Margin = New Padding(0, 0, 0, 8)
            }

            Dim listPanel As New Panel() With {
                .Dock = DockStyle.Fill,
                .Margin = New Padding(0)
            }

            Dim lstArchives As New ListBox() With {
                .Dock = DockStyle.Fill,
                .IntegralHeight = False
            }
            listPanel.Controls.Add(lstArchives)

            Dim buttonBar As New FlowLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.LeftToRight,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .WrapContents = True,
                .Padding = New Padding(0, 6, 0, 0),
                .Margin = New Padding(0)
            }

            Dim btnSave As New Button() With {.Text = "Store Current", .AutoSize = True}
            Dim btnUpdate As New Button() With {.Text = "Update Current Archive", .AutoSize = True}
            Dim btnRestore As New Button() With {.Text = "Restore Selected", .AutoSize = True}
            Dim btnDelete As New Button() With {.Text = "Delete Selected", .AutoSize = True}
            Dim btnClose As New Button() With {.Text = "Close", .AutoSize = True}

            buttonBar.Controls.Add(btnSave)
            buttonBar.Controls.Add(btnUpdate)
            buttonBar.Controls.Add(btnRestore)
            buttonBar.Controls.Add(btnDelete)
            buttonBar.Controls.Add(btnClose)

            layout.Controls.Add(lblInfo, 0, 0)
            layout.Controls.Add(listPanel, 0, 1)
            layout.Controls.Add(buttonBar, 0, 2)

            frm.CancelButton = btnClose

            Dim updateInfoLabel As System.Action =
                Sub()
                    Dim wrapWidth = Math.Max(250, layout.ClientSize.Width - layout.Padding.Horizontal)
                    lblInfo.MaximumSize = New System.Drawing.Size(wrapWidth, 0)
                    lblInfo.Text = BuildDialogueArchiveManagerInfoText()
                End Sub

            Dim refreshArchives As System.Action =
                Sub()
                    Dim previouslySelectedPath = ""
                    Dim currentlySelected = TryCast(lstArchives.SelectedItem, DialogueArchiveInfo)
                    If currentlySelected IsNot Nothing Then
                        previouslySelectedPath = currentlySelected.FilePath
                    End If

                    Dim archives = GetDialogueArchives()

                    lstArchives.BeginUpdate()
                    lstArchives.Items.Clear()
                    For Each archive In archives
                        lstArchives.Items.Add(archive)
                    Next
                    lstArchives.EndUpdate()

                    Dim restoredSelection As DialogueArchiveInfo = Nothing
                    If previouslySelectedPath.Length > 0 Then
                        For Each item As Object In lstArchives.Items
                            Dim archive = TryCast(item, DialogueArchiveInfo)
                            If archive IsNot Nothing AndAlso
                               archive.FilePath.Equals(previouslySelectedPath, StringComparison.OrdinalIgnoreCase) Then
                                restoredSelection = archive
                                Exit For
                            End If
                        Next
                    End If

                    If restoredSelection IsNot Nothing Then
                        lstArchives.SelectedItem = restoredSelection
                    ElseIf lstArchives.Items.Count > 0 Then
                        lstArchives.SelectedIndex = 0
                    End If

                    Dim hasSelection = (TryCast(lstArchives.SelectedItem, DialogueArchiveInfo) IsNot Nothing)
                    btnRestore.Enabled = hasSelection
                    btnDelete.Enabled = hasSelection
                    btnUpdate.Enabled = Not String.IsNullOrWhiteSpace(_activeDialogueArchiveName)

                    updateInfoLabel.Invoke()
                End Sub

            AddHandler frm.SizeChanged,
                Sub()
                    updateInfoLabel.Invoke()
                End Sub

            AddHandler lstArchives.SelectedIndexChanged,
                Sub()
                    Dim hasSelection = (TryCast(lstArchives.SelectedItem, DialogueArchiveInfo) IsNot Nothing)
                    btnRestore.Enabled = hasSelection
                    btnDelete.Enabled = hasSelection
                End Sub

            AddHandler btnClose.Click,
                Sub()
                    frm.Close()
                End Sub

            AddHandler btnSave.Click,
                Sub()
                    If PromptAndSaveCurrentDialogueArchive() Then
                        refreshArchives.Invoke()
                    End If
                End Sub

            AddHandler btnUpdate.Click,
                Sub()
                    If String.IsNullOrWhiteSpace(_activeDialogueArchiveName) Then
                        AppendSystemMessage("There is no linked archive to update.")
                        Return
                    End If

                    If SaveCurrentDialogueArchive(_activeDialogueArchiveName, overwriteWithoutPrompt:=True) Then
                        refreshArchives.Invoke()
                    End If
                End Sub

            AddHandler btnDelete.Click,
                Sub()
                    Dim selected = TryCast(lstArchives.SelectedItem, DialogueArchiveInfo)
                    If selected Is Nothing Then Return

                    Dim deleteAnswer = ShowCustomYesNoBox(
                        $"Do you want to delete the archived dialogue '{selected.Name}'?",
                        "Yes, delete",
                        "No, keep",
                        $"{AN} - Delete Dialogue Archive")

                    If deleteAnswer <> 1 Then
                        Return
                    End If

                    Try
                        File.Delete(selected.FilePath)
                        DeleteIndexDirectorySafe(GetArchiveIndexDirectoryPath(selected.Name))

                        If String.Equals(selected.FilePath, _activeDialogueArchiveFilePath, StringComparison.OrdinalIgnoreCase) Then
                            ClearCurrentActiveDialogueArchive()
                            PersistCurrentSessionSettings()
                            UpdateWindowTitle()
                        End If

                        AppendSystemMessage($"Archived dialogue deleted: {selected.Name}")
                        refreshArchives.Invoke()
                    Catch ex As Exception
                        AppendSystemMessage($"Failed to delete archived dialogue: {ex.Message}")
                    End Try
                End Sub

            AddHandler btnRestore.Click,
                Sub()
                    Dim selected = TryCast(lstArchives.SelectedItem, DialogueArchiveInfo)
                    If selected Is Nothing Then Return

                    If Not PromptToSaveCurrentDialogueBeforeSwitch() Then
                        Return
                    End If

                    If RestoreDialogueArchiveFromFile(selected.FilePath, selected.Name) Then
                        frm.DialogResult = DialogResult.OK
                        frm.Close()
                    End If
                End Sub

            AddHandler lstArchives.DoubleClick,
                Sub()
                    If btnRestore.Enabled Then
                        btnRestore.PerformClick()
                    End If
                End Sub

            refreshArchives.Invoke()
            frm.ShowDialog(Me)
        End Using
    End Sub

    Private Sub OnArchiveClick(sender As Object, e As EventArgs)
        ShowDialogueArchiveManager()
    End Sub

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
