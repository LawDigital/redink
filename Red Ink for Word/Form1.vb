' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: Form1.vb
' PURPOSE
'   Interactive chat assistant UI (Inky) embedded in Microsoft Word.
'   Provides conversational interface to LLM with optional document manipulation
'   commands and Markdown-rendered responses.
'
' =============================================================================
' ARCHITECTURE OVERVIEW
' =============================================================================
'
' UI Components
'   • Primary Controls:
'     - txtChatHistory:  Plain text transcript (hidden when HTML active)
'     - txtUserInput:    Multiline user input box
'     - wbChat:          WebBrowser control for Markdown-rendered HTML chat
'     - lblInstructions: Usage instructions displayed at top
'
'   • Buttons:
'     - btnSend:             Submit user message to LLM
'     - btnCopy:             Copy entire conversation to clipboard
'     - btnCopyLastAnswer:   Copy most recent assistant response
'     - btnClear:            Clear conversation history
'     - btnSwitchModel:      Toggle between primary/secondary/alternate models
'     - btnTools:            Select which tools are available for the chat session
'     - btnExit:             Close chat window (saves state)
'
'   • Checkboxes (User Configuration):
'     - chkIncludeDocText:    Include complete active document content
'     - chkIncludeSelection:  Include current selection or cursor context
'     - chkIncludeOtherDocs:  Include all other open Word documents
'     - chkPermitCommands:    Allow LLM to execute write operations
'     - chkEnableTooling:     Enable tool calling loop (when model supports tools)
'     - chkShowToolingLog:    Show/hide the tooling log window during tool runs
'     - Passed to ExecuteToolingLoop via hideLogWindow:=Not chkShowToolingLog.Checked
'     - Set from INI_ToolingLogWindow on first session load; not persisted
'     - chkStayOnTop:         Toggle always-on-top window behavior
'     - chkConvertMarkdown:   Apply Markdown formatting to inserted text
'
' Prompt Construction Pipeline
'   1. System Prompt Assembly:
'      - Base template from SharedContext.SP_ChatWord()
'      - User language interpolation ({UserLanguage})
'      - Assistant name and current timestamp
'      - Conditional capability declarations:
'        * "You have access to the user's active document" (if chkIncludeDocText)
'        * "You have access to a selection..." (if chkIncludeSelection)
'        * "You also have access to all other open Word documents" (if chkIncludeOtherDocs)
'      - Command permission block (SP_Add_ChatWord_Commands or SP_Add_Chat_NoCommands)
'
'   2. User Prompt Construction:
'      - Active document text (if enabled, extracted in Final view mode)
'      - Selection text or cursor context (25 chars before/after with "[cursor is here]" marker)
'      - Other open documents (via GatherSelectedDocuments with <DOCUMENTn> tags)
'      - User's current message
'      - Recent conversation history (trimmed to INI_ChatCap limit)
'
'   3. Context Enrichment:
'      - Comments/bubbles extracted via BubblesExtract (when available)
'      - Revision view temporarily set to wdRevisionsViewFinal (excludes deleted content)
'      - Document names included for multi-document context
'
' Model Selection System
'   • Primary Model:  Default model from INI_Model
'   • Secondary API:  Optional second model (INI_SecondAPI / INI_Model_2)
'   • Alternate Model: User-selectable from external INI file (INI_AlternateModelPath)
'
'   Model Switching Behavior:
'     - btnSwitchModel toggles between available models
'     - Alternate model selection uses ShowModelSelection dialog
'     - Configuration snapshot/restore pattern preserves global context integrity
'     - Secondary/alternate models disable document-related checkboxes (UpdateDocumentCheckboxesState)
'     - CallLlmWithSelectedModelAsync handles temporary config application
'
' Tooling (Optional)
'   • Tool selection:
'     - btnTools opens SelectToolsForSession dialog
'     - _selectedToolsForChat caches selected tool set per chat session
'
'   • Tool calling loop:
'     - When chkEnableTooling is checked and current model supports tools,
'       ExecuteToolingLoop is used instead of a direct LLM() call.
'
'   • Tooling log window:
'     - chkShowToolingLog toggles visibility of the tooling log while tools run.
'     - Passed to ExecuteToolingLoop via hideLogWindow:=Not chkShowToolingLog.Checked
'     - Persisted to My.Settings.ChatShowToolingLog
'
'   • ToolTrigger "(t)" - One-Shot Tooling Model:
'     - Users can type "(t)" anywhere in their prompt to invoke a one-shot tooling request.
'     - The "(t)" token is always stripped from the prompt before it is sent to the LLM.
'     - On detection, btnSend_Click loads the model marked ToolDefaultModel=True from the
'       alternate models INI via GetSpecialTaskModel, without permanently altering context.
'     - The ToolDefaultModel config is captured (snapshot), the global context is immediately
'       restored, and the snapshot is applied temporarily only for the ExecuteToolingLoop call.
'     - If no tools are currently selected, the tool selection dialog is shown (forceDialog).
'     - Validation checks: INI path configured, ToolDefaultModel found, model supports tooling,
'       tools selected. Each failure reports an error to chat and restores the original prompt
'       (prefixed with "(t)") into txtUserInput so the user can resend without retyping.
'     - Availability is checked at form load via HasToolingCapableSpecialTaskModel(). When 
'       available, lblInstructions includes "(t)" usage hint.
'     - The one-shot model is never persisted — after the request completes, subsequent messages
'       use whatever model was previously active (primary/secondary/alternate).
'
' Bot Command Execution (Optional)
'   Pattern: [#verb: @@argument1@@ §§argument2§§ #]
'
'   Supported Commands:
'     • find:          Locate and highlight text (ExecuteFindCommand)
'     • replace:       Replace or delete text with tracked changes (ExecuteReplaceCommand)
'     • insert:        Insert at current cursor position (ExecuteInsertCommand)
'     • insertbefore:  Insert text before anchor text (ExecuteInsertBeforeAfterCommand)
'     • insertafter:   Insert text after anchor text (ExecuteInsertBeforeAfterCommand)
'     • addcomment:    Add Word comment to matched text (ExecuteAddComment)
'     • replycomment:  Add threaded reply to existing comment (ExecuteReplyToCommentByIdToken)
'
'   Command Processing:
'     - ParseCommands extracts commands using tempered-greedy regex
'     - Supports single @ or § inside arguments (only @@ or §§ terminate)
'     - Second argument optional (empty for delete operations)
'     - DecodeParagraphMarks normalizes line breaks (vbCr, vbLf, \r\n, ^p, ^13)
'     - FindLongTextInChunks ensures reliability with large documents
'     - MarkerChar (U+E000) prevents infinite loops during replacement
'     - TOC ranges skipped to prevent corruption (TocEndIfInside)
'     - ESC key aborts execution mid-operation
'     - Failed commands reported to chat with error formatting
'
' Markdown & HTML Rendering
'   • Markdig Pipeline:
'     - UseAdvancedExtensions (tables, footnotes, task lists, etc.)
'     - UseEmojiAndSmiley (emoji shortcode support)
'     - UseSoftlineBreakAsHardlineBreak (single newlines render as <br>)
'
'   • User Messages:  HTML-encoded plain text with line breaks preserved
'   • Assistant Messages: Markdown → HTML conversion with link instrumentation
'   • Link Handling: All anchor tags instrumented to open in default browser
'     via BrowserBridge COM-visible class (prevents internal WebBrowser navigation)
'   • Thinking Indicator: Temporary DOM element removed on LLM response
'   • Inline Optimization: Single-paragraph responses rendered as <span> instead of <div>
'
' Persistence & State Management
'   • Chat History:
'     - Plain text:  My.Settings.LastChatHistory (fallback, trimmed to INI_ChatCap)
'     - HTML:        My.Settings.LastChatHistoryHtml (preferred, preserves formatting)
'     - In-memory:   _chatHistory list of (Role, Content) tuples
'
'   • Window State:
'     - Position:    My.Settings.FormLocation
'     - Size:        My.Settings.FormSize
'     - RestoreBounds saved when minimized/maximized
'
'   • User Preferences:
'     - IncludeDocument, IncludeSelection, DoCommands
'     - ChatEnableTooling (persisted), ChatShowToolingLog (session-only, from INI)
'     - NotAlwaysOnTop, ConvertMarkdownInChat
'     - All persisted via My.Settings except ChatShowToolingLog
'
' Safety & COM Interop
'   • Word View Management:
'     - Temporarily switches to wdRevisionsViewFinal for clean text extraction
'     - ShowRevisionsAndComments toggled to hide markups during extraction
'     - Original view settings restored in Finally blocks
'
'   • Selection Restoration:
'     - Original selection bounds saved before command execution
'     - Restored with boundary guards (Math.Min/Max to prevent out-of-range)
'     - Main story enforcement prevents caret stuck in headers/footnotes/comments
'
'   • Abort Mechanism:
'     - GetAsyncKeyState polls ESC key during long operations
'     - InfoBox provides user feedback during command execution
'
'   • COM Object Cleanup:
'     - Marshal.ReleaseComObject called on Word.Application references
'     - Try-Catch-Finally patterns ensure cleanup on errors
'
' Dependencies
'   • SharedLibrary.SharedContext:  INI configuration and model settings
'   • SharedLibrary.SharedMethods:  LLM invocation, UI helpers, model selection
'   • Markdig:                      Markdown parsing and HTML generation
'   • Microsoft.Office.Interop.Word: Word automation and COM interop
'   • HtmlAgilityPack:              HTML parsing (ConvertHtmlToPlainText)
'
' Known Limitations
'   • WebBrowser control uses legacy IE rendering engine (no modern CSS support)
'   • FindLongTextInChunks may fail on complex table structures
'   • TOC detection relies on TablesOfContents collection (custom TOCs may be missed)
'   • Comment ID tokens must match specific formats (see TryParseCommentIdToken)
'
' =============================================================================


Imports System.ComponentModel
Imports System.Data
Imports System.Diagnostics
Imports System.Drawing
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Markdig
Imports Microsoft.Office.Interop.Word
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

''' <summary>
''' Main form class for the AI chat interface (Inky) embedded in Microsoft Word.
''' Provides conversational LLM interaction with optional document manipulation capabilities.
''' </summary>
''' <remarks>
''' This form manages all aspects of the chat UI including:
''' - System and user prompt construction based on document context
''' - LLM model selection and switching (primary/secondary/alternate)
''' - Bot command parsing and execution
''' - Markdown rendering via WebBrowser control
''' - Chat history persistence (plain text and HTML)
''' See file header for complete architecture documentation.
''' </remarks>
Public Class frmAIChat

    ' =========================================================================
    ' Windows API Imports
    ' =========================================================================

    ''' <summary>
    ''' Windows API function to check the current state of a virtual key.
    ''' Used to detect ESC key press during long-running command operations.
    ''' </summary>
    ''' <param name="vKey">Virtual-key code (e.g., Keys.Escape = 27)</param>
    ''' <returns>
    ''' High-order bit indicates key is down.
    ''' Low-order bit indicates key was pressed after previous GetAsyncKeyState call.
    ''' </returns>
    <DllImport("user32.dll")>
    Private Shared Function GetAsyncKeyState(vKey As Integer) As Short
    End Function

    ' =========================================================================
    ' Constants
    ' =========================================================================

    ''' <summary>Full application name displayed in messages and credits.</summary>
    Const AN As String = "Red Ink"

    ''' <summary>Chat assistant name shown in conversation and window title.</summary>
    Const AN5 As String = "Inky"

    ''' <summary>Abbreviated name prefixed to comment replies (e.g., "RI: Reply text").</summary>
    Const AN6 As String = "RI"

    Const ToolTrigger As String = "(ag)"

    ''' <summary>
    ''' Special Unicode private-use character (U+E000) inserted during text replacement.
    ''' Acts as temporary marker to prevent infinite loops when searching for replaced text.
    ''' Removed after all replacements complete via ReplaceSpecialCharacter().
    ''' </summary>
    Const MarkerChar As String = ChrW(&HE000)

    ''' <summary>
    ''' Number of characters to extract before and after cursor position for context.
    ''' Used by GetCursorContext() to provide localized document context when no selection exists.
    ''' </summary>
    Const CursorPositionCount As Integer = 25

    ' =========================================================================
    ' Private Fields - Chat State
    ' =========================================================================

    ''' <summary>
    ''' Tracks whether a newline prefix should be added to next chat history entry.
    ''' Ensures proper spacing between conversation blocks in plain text transcript.
    ''' </summary>
    Private PreceedingNewline As String = ""

    ''' <summary>
    ''' Stores older chat history preserved during model switches.
    ''' Appended to conversationSoFar in btnSend_Click to maintain context across model changes.
    ''' Cleared after first use to prevent duplication.
    ''' </summary>
    Private OldChat As String = ""

    ''' <summary>
    ''' Word's default UI language code (e.g., "en-US", "de-DE").
    ''' Retrieved from Globals.ThisAddIn.GetWordDefaultInterfaceLanguage().
    ''' Interpolated into system prompt as {UserLanguage} to guide LLM response localization.
    ''' </summary>
    Private UserLanguage As String = Globals.ThisAddIn.GetWordDefaultInterfaceLanguage()

    ''' <summary>
    ''' Current system prompt assembled dynamically based on active checkboxes.
    ''' Includes base template (SP_ChatWord), assistant name, timestamp, capability declarations,
    ''' and command permission block (SP_Add_ChatWord_Commands or SP_Add_Chat_NoCommands).
    ''' Rebuilt on each Send button click.
    ''' </summary>
    Private SystemPrompt As String = ""

    ''' <summary>
    ''' True when the most recent command batch intentionally moved focus to Word.
    ''' Prevents the chat input box from immediately stealing focus back.
    ''' </summary>
    Private _keepFocusOnDocumentAfterCommands As Boolean = False
    Private _dialogOwnerScope As System.IDisposable = Nothing

    ' =========================================================================
    ' Private Fields - Model Configuration
    ' =========================================================================

    ''' <summary>
    ''' True when user has selected an alternate model via btnSwitchModel.
    ''' When true, CallLlmWithSelectedModelAsync applies _alternateModelConfig temporarily.
    ''' </summary>
    Private _alternateModelSelected As Boolean = False

    ''' <summary>
    ''' Snapshot of alternate model configuration captured after user selection.
    ''' Applied temporarily during LLM calls when _alternateModelSelected = True.
    ''' Original config restored immediately after call to keep global context pristine.
    ''' </summary>
    Private _alternateModelConfig As ModelConfig = Nothing

    ''' <summary>
    ''' Display name of alternate model shown in window title and button text.
    ''' Retrieved from SharedMethods.LastAlternateModel after user selection.
    ''' </summary>
    Private _alternateModelDisplayName As String = Nothing

    ''' <summary>
    ''' Cached list of currently selected tools for the chat session.
    ''' Populated via SelectToolsForSession when tooling is used.
    ''' </summary>
    Private _selectedToolsForChat As List(Of ModelConfig) = Nothing

    ''' <summary>
    ''' Guards first-time initialization of the tooling log checkbox from INI_ToolingLogWindow.
    ''' After the first call to UpdateToolingControlsState, mid-session user toggles are preserved.
    ''' </summary>
    Private _toolingControlsInitialized As Boolean = False

    ' =========================================================================
    ' UI Controls - Buttons
    ' =========================================================================

    ''' <summary>Copies entire plain text conversation to clipboard.</summary>
    Private WithEvents btnCopy As New Button() With {.Text = "Copy All", .AutoSize = True}

    ''' <summary>Copies most recent assistant response to clipboard.</summary>
    Private WithEvents btnCopyLastAnswer As New Button() With {.Text = "Copy Last Answer", .AutoSize = True}

    ''' <summary>Clears conversation history and displays new welcome message.</summary>
    Private WithEvents btnClear As New Button() With {.Text = "Clear", .AutoSize = True}

    ''' <summary>Closes chat window after saving conversation and window state.</summary>
    Private WithEvents btnExit As New Button() With {.Text = "Close", .AutoSize = True}

    ''' <summary>Submits user message to LLM and displays response.</summary>
    Private WithEvents btnSend As New Button() With {.Text = "Send", .AutoSize = True}

    ''' <summary>
    ''' Toggles between primary/secondary/alternate models.
    ''' Visible only when INI_SecondAPI = true or INI_AlternateModelPath is configured.
    ''' Text changes based on _alternateModelSelected state: "Primary model" or "Alternate Model".
    ''' </summary>
    Private WithEvents btnSwitchModel As New Button() With {.Text = "Switch Model", .AutoSize = True}

    ''' <summary>
    ''' Opens tool selection dialog to configure which tools are available.
    ''' Disabled when current model does not support tooling.
    ''' </summary>
    Private WithEvents btnTools As New Button() With {.Text = Globals.ThisAddIn.ToolFriendlyName, .AutoSize = True}

    ''' <summary>
    ''' Clickable label to open the Inky Memory file for manual editing.
    ''' Styled as a link to save space alongside checkboxes.
    ''' </summary>
    Private WithEvents lnkEditMemory As New LinkLabel() With {
        .Text = "Edit",
        .AutoSize = True,
        .Visible = My.Settings.ChatInkyMemory,
        .Margin = New Padding(0, 5, 0, 0)
    }


    ' =========================================================================
    ' UI Controls - Checkboxes
    ' =========================================================================

    ''' <summary>
    ''' When checked, includes complete active document content in prompt.
    ''' Extracted in Final view mode (excludes tracked deletions) via GetActiveDocumentText().
    ''' Mutually exclusive with chkIncludeSelection in UI logic.
    ''' Persisted to My.Settings.IncludeDocument.
    ''' </summary>
    Private WithEvents chkIncludeDocText As New System.Windows.Forms.CheckBox() With {
        .Text = "Include document",
        .AutoSize = True,
        .Checked = My.Settings.IncludeDocument
    }

    ''' <summary>
    ''' Maximum active-document size (in characters) allowed when "Include document"
    ''' is checked. Documents above this threshold cause btnSend_Click to abort early
    ''' and automatically uncheck chkIncludeDocText, avoiding UI stalls on very large
    ''' documents. Empirically, roughly 2,000,000 characters still performs well.
    ''' </summary>
    Private Const MaxIncludeDocumentCharacters As Integer = 2000000

    ''' <summary>
    ''' When checked, includes current selection or cursor context in prompt.
    ''' If no selection exists, GetCursorContext() extracts CursorPositionCount chars before/after cursor.
    ''' Mutually exclusive with chkIncludeDocText in UI logic.
    ''' Persisted to My.Settings.IncludeSelection.
    ''' </summary>
    Private WithEvents chkIncludeselection As New System.Windows.Forms.CheckBox() With {
        .Text = "Include selection",
        .AutoSize = True,
        .Checked = If(My.Settings.IncludeDocument, False, My.Settings.IncludeSelection)
    }

    ''' <summary>
    ''' When checked, allows LLM to execute bot commands on the document.
    ''' Requires either chkIncludeDocText or chkIncludeselection to be checked.
    ''' Commands parsed from LLM response and executed via ExecuteAnyCommands().
    ''' Persisted to My.Settings.DoCommands.
    ''' </summary>
    Private WithEvents chkPermitCommands As New System.Windows.Forms.CheckBox() With {
        .Text = "Grant write access",
        .AutoSize = True,
        .Checked = My.Settings.DoCommands
    }

    ''' <summary>
    ''' When checked, enables tool calling via ExecuteToolingLoop instead of direct LLM().
    ''' Only enabled when current model supports tooling.
    ''' Persisted to My.Settings.ChatEnableTooling.
    ''' </summary>
    Private WithEvents chkEnableTooling As New System.Windows.Forms.CheckBox() With {
        .Text = $"Enable {Globals.ThisAddIn.ToolFriendlyName.ToLower}",
        .AutoSize = True,
        .Checked = My.Settings.ChatEnableTooling
    }

    ''' <summary>
    ''' When checked, the selected advanced tools remain callable.
    ''' When unchecked, advanced-tool selections stay persisted but are excluded from the effective tool list.
    ''' Persisted to My.Settings.AdvancedToolsEnabled.
    ''' </summary>
    Private WithEvents chkAdvancedTools As New System.Windows.Forms.CheckBox() With {
        .Text = "Advanced tools",
        .AutoSize = True,
        .Checked = My.Settings.AdvancedToolsEnabled
    }

    ''' <summary>
    ''' When checked, shows the tooling log window during tool execution.
    ''' Not persisted — set from INI_ToolingLogWindow on first session load,
    ''' then preserves user's mid-session toggle. Respected by (t) trigger
    ''' even when the checkbox is disabled (non-tooling model).
    ''' </summary>
    Private WithEvents chkShowToolingLog As New System.Windows.Forms.CheckBox() With {
    .Text = "Tooling log",
    .AutoSize = True,
    .Checked = False
}

    ''' <summary>
    ''' When checked, enables InkyMemory — persistent cross-session learning.
    ''' The LLM is instructed to identify and store user preferences automatically.
    ''' Persisted to My.Settings.ChatInkyMemory.
    ''' </summary>
    Private WithEvents chkInkyMemory As New System.Windows.Forms.CheckBox() With {
        .Text = "Inky Memory",
        .AutoSize = True,
        .Checked = My.Settings.ChatInkyMemory
    }

    ''' <summary>
    ''' Controls window TopMost property.
    ''' Inversely labeled: checked = NOT always on top (TopMost = false).
    ''' Persisted to My.Settings.NotAlwaysOnTop.
    ''' </summary>
    Private WithEvents chkStayOnTop As New System.Windows.Forms.CheckBox() With {
        .Text = "Not always on top",
        .AutoSize = True,
        .Checked = My.Settings.NotAlwaysOnTop
    }

    ''' <summary>
    ''' When checked, applies Markdown formatting to inserted text and comment replies.
    ''' Controls whether ConvertMarkdownToWord() is called after command execution.
    ''' Persisted to My.Settings.ConvertMarkdownInChat.
    ''' </summary>
    Private WithEvents chkConvertMarkdown As New System.Windows.Forms.CheckBox() With {
        .Text = "Do format",
        .AutoSize = True,
        .Checked = My.Settings.ConvertMarkdownInChat
    }

    ''' <summary>
    ''' When checked, silently includes all other open Word documents in prompt.
    ''' Calls GatherSelectedDocuments(IncludeName:=True, ExceptCurrent:=True, SilentAndGetAll:=True).
    ''' Each document wrapped in numbered DOCUMENTn tags with document name.
    ''' Persisted to My.Settings.ChatIncludeOtherOpenWordDocs.
    ''' </summary>
    Private WithEvents chkIncludeOtherDocs As New System.Windows.Forms.CheckBox() With {
        .Text = "Include all other open Word docs",
        .AutoSize = True,
        .Checked = My.Settings.ChatIncludeOtherOpenWordDocs
    }

    ' =========================================================================
    ' UI Controls - Layout Panels
    ' =========================================================================

    ''' <summary>
    ''' FlowLayoutPanel hosting action buttons (Send, Copy, Clear, Switch Model, Exit).
    ''' Docked to bottom of form with left-to-right flow and auto-sizing.
    ''' </summary>
    Dim pnlButtons As New FlowLayoutPanel() With {
        .Dock = DockStyle.Bottom,
        .FlowDirection = FlowDirection.LeftToRight,
        .AutoSize = True,
        .AutoSizeMode = AutoSizeMode.GrowAndShrink,
        .Height = 40
    }

    ''' <summary>
    ''' FlowLayoutPanel hosting configuration checkboxes.
    ''' Docked below user input area with left-to-right flow and auto-sizing.
    ''' </summary>
    Dim pnlCheckboxes As New FlowLayoutPanel() With {
        .Dock = DockStyle.Bottom,
        .FlowDirection = FlowDirection.LeftToRight,
        .AutoSize = True,
        .AutoSizeMode = AutoSizeMode.GrowAndShrink,
        .Height = 40
    }

    ''' <summary>
    ''' SplitContainer separating the chat history (Panel1) from the user input (Panel2).
    ''' The splitter bar allows the user to resize the input area by dragging.
    ''' </summary>
    Private WithEvents splitChat As New SplitContainer() With {
        .Dock = DockStyle.Fill,
        .Orientation = Orientation.Horizontal,
        .FixedPanel = FixedPanel.Panel2,
        .SplitterWidth = 6,
        .Panel2MinSize = 40,
        .Panel1MinSize = 100
    }

    ' =========================================================================
    ' Private Fields - Application State
    ' =========================================================================

    ''' <summary>
    ''' Shared context providing INI configuration and LLM settings.
    ''' Accessed for model names, API endpoints, system prompts, and chat capacity limits.
    ''' </summary>
    Private _context As ISharedContext = New SharedContext()

    ''' <summary>
    ''' True when secondary API model is active (either toggled or alternate selected).
    ''' When true, UpdateDocumentCheckboxesState() disables document/selection/command checkboxes.
    ''' </summary>
    Private _useSecondApi As Boolean = False

    ''' <summary>
    ''' Complete conversation history as (Role As String, Content As String) tuples.
    ''' Used by BuildConversationString() to construct context window trimmed to INI_ChatCap.
    ''' Plain text content only (Markdown stripped for commands/persistence).
    ''' </summary>
    Private _chatHistory As New List(Of (Role As String, Content As String))

    ' Loaded external context (attached via the Load Context button)
    Private Const PersistedContextFileName As String = "redink-wordchatcontext.txt"

    ' Loaded semantic index (attached via the Load Context button, either a document OR an index).
    Private Const PersistedIndexFileName As String = "redink-wordchatindex.index.txt"

    Private Shared ReadOnly SupportedContextExtensions As String() = {
        ".txt", ".rtf", ".doc", ".docx", ".xlsx", ".pdf", ".pptx", ".msg", ".eml",
        ".ini", ".csv", ".log", ".json", ".xml", ".html", ".htm", ".md",
        ".vb", ".cs", ".js", ".ts", ".py", ".java", ".cpp", ".c", ".h", ".sql", ".yaml", ".yml"
    }

    Private Shared _cachedLoadedContextContent As String = Nothing
    Private Shared _cachedLoadedContextPath As String = Nothing

    Private _loadedContextContent As String = Nothing
    Private _loadedContextPath As String = Nothing
    Private _isUpdatingPersistContextCheckbox As Boolean = False
    Private ReadOnly _contextToolTip As New System.Windows.Forms.ToolTip()

    ' Loaded semantic index state (either a document context or an index is active, never both).
    Private _loadedIndexSourcePath As String = Nothing
    Private _loadedIndexDisplayName As String = Nothing
    Private Shared _cachedLoadedIndexPath As String = Nothing
    Private Shared _cachedLoadedIndexDisplayName As String = Nothing
    Private _semanticConversationState As New SharedMethods.SemanticSearchConversationState()

    ''' <summary>Button: attach or remove external context material (files or a folder).</summary>
    Private WithEvents btnLoadContext As New Button() With {.Text = "Load Context", .AutoSize = True}

    ''' <summary>Checkbox: persist the loaded context or index to durable AppData storage.</summary>
    Private WithEvents chkPersistContext As New System.Windows.Forms.CheckBox() With {
        .Text = "Persist context",
        .AutoSize = True
    }

    ' =========================================================================
    ' Constructor
    ' =========================================================================

    ''' <summary>
    ''' Initializes the chat form with shared context and constructs the UI layout.
    ''' Creates a TableLayoutPanel with 4 rows: instructions label, split chat/input area,
    ''' checkboxes panel, and buttons panel. The chat history and user input are separated
    ''' by a draggable splitter so the user can resize the input area.
    ''' </summary>
    ''' <param name="context">Shared context providing INI settings and LLM configuration</param>
    Public Sub New(context As ISharedContext)
        ' Required designer initialization
        InitializeComponent()

        Me.AutoSize = False

        ' Configure text controls for multiline input
        txtChatHistory.Multiline = True
        txtUserInput.Multiline = True
        txtUserInput.ScrollBars = ScrollBars.Vertical
        txtUserInput.WordWrap = True

        ' Create main layout container (4 rows, 1 column)
        Dim mainLayout As New TableLayoutPanel() With {
            .ColumnCount = 1,
            .RowCount = 4,
            .Dock = DockStyle.Fill,
            .AutoSize = False,
            .Padding = New Padding(10)
        }

        ' Set column to stretch to full width
        mainLayout.ColumnStyles.Clear()
        mainLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))

        ' Override padding to add extra space on right edge
        mainLayout.Padding = New Padding(left:=10, top:=10, right:=20, bottom:=10)

        ' Define row sizing behavior:
        ' Row 0 (instructions): Auto-size to content
        ' Row 1 (split container with chat + input): Fill remaining space (100%)
        ' Row 2 (checkboxes): Auto-size to content
        ' Row 3 (buttons): Auto-size to content
        mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        mainLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))

        ' Configure control docking behavior
        lblInstructions.AutoSize = True
        lblInstructions.Dock = DockStyle.Top
        txtChatHistory.Dock = DockStyle.Fill
        txtUserInput.Dock = DockStyle.Fill

        ' Configure the SplitContainer panels
        ' Panel1 = chat history (top), Panel2 = user input (bottom, resizable via splitter)
        splitChat.Panel1.Controls.Add(txtChatHistory)
        splitChat.Panel2.Controls.Add(txtUserInput)
        splitChat.SplitterDistance = 300 ' Default: generous space for chat history

        ' Add controls to layout (column 0, respective rows)
        mainLayout.Controls.Add(lblInstructions, 0, 0)
        mainLayout.Controls.Add(splitChat, 0, 1)
        mainLayout.Controls.Add(pnlCheckboxes, 0, 2)
        mainLayout.Controls.Add(pnlButtons, 0, 3)

        ' Initialize HTML chat UI (WebBrowser control overlay)
        InitChatHtmlUI(mainLayout)

        ' Replace form's control collection with new layout
        Me.Controls.Clear()
        Me.Controls.Add(mainLayout)

        ' Store shared context reference
        _context = context
    End Sub

    ' =========================================================================
    ' Form Load Event
    ' =========================================================================

    Protected Overrides Sub OnHandleCreated(e As System.EventArgs)
        MyBase.OnHandleCreated(e)

        If _dialogOwnerScope Is Nothing Then
            _dialogOwnerScope = SharedMethods.PushDialogOwner(Me)
        End If
    End Sub

    Protected Overrides Sub OnHandleDestroyed(e As System.EventArgs)
        Dim scope As System.IDisposable = _dialogOwnerScope
        _dialogOwnerScope = Nothing

        If scope IsNot Nothing Then
            Try
                scope.Dispose()
            Catch
            End Try
        End If

        MyBase.OnHandleDestroyed(e)
    End Sub

    ''' <summary>
    ''' Handles form initialization after all controls are created.
    ''' Restores previous chat history (HTML preferred, plain text fallback),
    ''' positions window from saved settings, configures UI elements,
    ''' and displays welcome message if no prior chat exists.
    ''' </summary>
    Private Async Sub frmAIChat_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Configure form positioning and keyboard handling
        Me.StartPosition = FormStartPosition.Manual
        Me.KeyPreview = True  ' Enable form-level key event handling (for ESC to close)

        ' Restore saved plain text chat history from settings
        Dim previousChat As String = My.Settings.LastChatHistory
        If Not String.IsNullOrEmpty(previousChat) Then
            txtChatHistory.Text = previousChat
            OldChat = previousChat  ' Preserve for context in first message after load
            PreceedingNewline = Environment.NewLine
        End If

        ' Initialize HTML rendering engine
        InitializeChatHtml()

        ' Restore chat transcript (prefer HTML format for rich rendering)
        Dim previousChatHtml As String = My.Settings.LastChatHistoryHtml
        Dim hasExistingChat As Boolean = False

        If Not String.IsNullOrEmpty(previousChatHtml) Then
            ' Restore HTML transcript (links auto-wired via wireLinks JavaScript)
            AppendHtml(previousChatHtml)
            hasExistingChat = True
        ElseIf Not String.IsNullOrEmpty(previousChat) Then
            ' Fallback: convert plain text to HTML format
            AppendTranscriptToHtml(previousChat)
            hasExistingChat = True
        End If

        ' Configure form appearance
        Me.Font = New System.Drawing.Font("Segoe UI", 9)
        Me.FormBorderStyle = FormBorderStyle.Sizable
        Me.Icon = Icon.FromHandle(New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard)).GetHicon())
        Me.TopMost = True
        Me.MinimumSize = New Size(830, 521)

        ' Restore window position and size from settings
        If My.Settings.FormLocation <> System.Drawing.Point.Empty AndAlso My.Settings.FormSize <> Size.Empty Then
            Me.Location = My.Settings.FormLocation
            Me.Size = My.Settings.FormSize
        Else
            Me.StartPosition = FormStartPosition.CenterScreen
        End If
        SharedMethods.EnsureVisibleOnScreen(Me)

        ' Set input panel to double the original designer height (63px × 2 = 126px)
        Try
            Dim desiredInputHeight As Integer = 126
            Dim newDistance As Integer = splitChat.Height - desiredInputHeight - splitChat.SplitterWidth
            If newDistance >= splitChat.Panel1MinSize Then
                splitChat.SplitterDistance = newDistance
            End If
        Catch
            ' Layout not ready yet; keep default SplitterDistance
        End Try

        ' Attach input handlers
        AddHandler txtUserInput.KeyDown, AddressOf UserInput_KeyDown
        AddHandler txtUserInput.KeyPress, AddressOf UserInput_KeyPress
        AddHandler Microsoft.Win32.SystemEvents.DisplaySettingsChanged, AddressOf OnDisplaySettingsChanged

        ' Configure instructions label
        Dim baseInstructions As String = "Enter your question and Enter (or 'Send'). You can allow the chatbot to do actions on your document (search, replace, delete, insert text and add or reply to comments). It does not see deletions, markups as such or formatting."

        If _context.INI_PromptLib Then
            baseInstructions &= " Type '/' at the start of a prompt or after whitespace to insert a prompt from the prompt library."
        End If

        Dim toolTriggerAvailable As Boolean =
            SharedMethods.HasToolingCapableSpecialTaskModel(_context, _context.INI_AlternateModelPath, "ToolDefaultModel")

        If toolTriggerAvailable Then
            baseInstructions &= $" Type '{ToolTrigger}' in your prompt to use the configured {Globals.ThisAddIn.ToolFriendlyName.ToLower} model for a single request."
        End If

        lblInstructions.Text = baseInstructions
        lblInstructions.AutoSize = True
        lblInstructions.Height = 50
        lblInstructions.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        lblInstructions.TextAlign = ContentAlignment.MiddleLeft

        ' Populate button panel
        pnlButtons.Padding = New Padding(0, 2, 8, 12)
        pnlButtons.Controls.Add(btnSend)
        pnlButtons.Controls.Add(btnLoadContext)
        pnlButtons.Controls.Add(btnCopyLastAnswer)
        pnlButtons.Controls.Add(btnCopy)
        pnlButtons.Controls.Add(btnClear)

        ' Show model switch button only if secondary API or alternate INI configured
        If _context.INI_SecondAPI OrElse Not String.IsNullOrWhiteSpace(_context.INI_AlternateModelPath) Then
            UpdateModelButtonText()
            pnlButtons.Controls.Add(btnSwitchModel)
        End If

        pnlButtons.Controls.Add(btnTools)
        pnlButtons.Controls.Add(btnExit)

        ' Populate checkbox panel
        pnlCheckboxes.Padding = New Padding(0, 1, 8, 1)
        pnlCheckboxes.Controls.Add(chkIncludeselection)
        pnlCheckboxes.Controls.Add(chkIncludeDocText)
        pnlCheckboxes.Controls.Add(chkPermitCommands)
        pnlCheckboxes.Controls.Add(chkEnableTooling)
        pnlCheckboxes.Controls.Add(chkAdvancedTools)
        pnlCheckboxes.Controls.Add(chkShowToolingLog)
        pnlCheckboxes.Controls.Add(chkStayOnTop)
        pnlCheckboxes.Controls.Add(chkConvertMarkdown)
        pnlCheckboxes.Controls.Add(chkIncludeOtherDocs)
        pnlCheckboxes.Controls.Add(chkInkyMemory)
        pnlCheckboxes.Controls.Add(lnkEditMemory)
        pnlCheckboxes.Controls.Add(chkPersistContext)

        ' Attach event handlers to buttons
        AddHandler btnCopy.Click, AddressOf btnCopy_Click
        AddHandler btnClear.Click, AddressOf btnClear_Click
        AddHandler btnSend.Click, AddressOf btnSend_Click
        AddHandler btnCopyLastAnswer.Click, AddressOf btnCopyLastAnswer_Click
        AddHandler btnSwitchModel.Click, AddressOf btnSwitchModel_Click
        AddHandler btnExit.Click, AddressOf btnExit_Click

        ' Attach event handlers to checkboxes
        AddHandler chkIncludeselection.Click, AddressOf chkIncludeselection_Click
        AddHandler chkIncludeDocText.Click, AddressOf chkIncludeDocText_Click
        AddHandler chkPermitCommands.Click, AddressOf chkPermitCommands_Click
        AddHandler chkStayOnTop.Click, AddressOf chkStayontop_Click
        AddHandler chkConvertMarkdown.Click, AddressOf chkConvertMarkdown_Click
        AddHandler chkIncludeOtherDocs.Click, AddressOf chkIncludeOtherDocs_Click
        AddHandler chkInkyMemory.Click, AddressOf chkInkyMemory_Click
        AddHandler lnkEditMemory.LinkClicked, AddressOf lnkEditMemory_LinkClicked
        AddHandler btnLoadContext.Click, AddressOf btnLoadContext_Click
        AddHandler chkPersistContext.CheckedChanged, AddressOf chkPersistContext_CheckedChanged

        ' Attach event handlers for tooling controls
        AddHandler chkEnableTooling.Click, AddressOf chkEnableTooling_Click
        AddHandler chkAdvancedTools.Click, AddressOf chkAdvancedTools_Click
        AddHandler chkShowToolingLog.CheckedChanged, AddressOf chkShowToolingLog_CheckedChanged
        AddHandler btnTools.Click, AddressOf btnTools_Click

        _isUpdatingPersistContextCheckbox = True
        Try : chkPersistContext.Checked = My.Settings.ChatPersistContext : Catch : chkPersistContext.Checked = False : End Try
        _isUpdatingPersistContextCheckbox = False
        UpdatePersistContextTooltip()

        If Not chkPersistContext.Checked Then
            DeletePersistedContextFile(False)
            DeletePersistedIndexFile(False)
        End If

        Await RestoreLoadedContextAsync()
        UpdateLoadContextButtonText()

        RestoreAlternateModelFromSettings()

        ' Update window title with active model name
        UpdateTitle()

        ' Either restore existing chat or show welcome message
        If hasExistingChat Then
            txtChatHistory.SelectionStart = txtChatHistory.Text.Length
            txtChatHistory.ScrollToCaret()
        Else
            Dim result = Await WelcomeMessage()
        End If


        ' Update tooling controls based on current model support
        UpdateToolingControlsState()

        ' Set focus to user input if empty
        If String.IsNullOrEmpty(txtUserInput.Text) Then txtUserInput.Focus()

    End Sub

    ''' <summary>
    ''' Repositions the form after monitor or resolution changes.
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

    ' =========================================================================
    ' Title and Model Management
    ' =========================================================================

    ''' <summary>
    ''' Updates form title to show currently active model name.
    ''' Priority: alternate model display name > second API model > primary model.
    ''' </summary>
    Private Sub UpdateTitle()
        Dim titleModel As String

        ' Determine which model name to display (priority order)
        If Not String.IsNullOrWhiteSpace(_context.INI_AlternateModelPath) AndAlso
           _alternateModelSelected AndAlso
           Not String.IsNullOrWhiteSpace(_alternateModelDisplayName) Then
            ' Alternate model selected and configured
            titleModel = _alternateModelDisplayName
        Else
            ' Primary or secondary model active
            titleModel = If(_useSecondApi, _context.INI_Model_2, _context.INI_Model)
        End If

        Me.Text = $"Chat (using {titleModel})"
    End Sub

    ' =========================================================================
    ' LLM Invocation with Model Configuration
    ' =========================================================================

    ''' <summary>
    ''' Executes LLM call with temporary alternate model configuration if selected.
    ''' Backs up current config, applies alternate, runs LLM, then restores original.
    ''' Ensures global context remains pristine between calls.
    ''' </summary>
    ''' <param name="systemPrompt">System prompt with capabilities and instructions</param>
    ''' <param name="fullPrompt">Complete user prompt including context and conversation</param>
    ''' <returns>LLM response text</returns>
    ''' <remarks>
    ''' This snapshot/restore pattern prevents alternate model config from polluting
    ''' the global SharedContext used by other add-in features. The backup config
    ''' captures the current state, alternate config is applied only for the duration
    ''' of the LLM call, then original config is restored in the Finally block.
    ''' </remarks>
    Private Async Function CallLlmWithSelectedModelAsync(systemPrompt As String, fullPrompt As String) As Task(Of String)
        Dim backupConfig As ModelConfig = Nothing
        Dim appliedAlternate As Boolean = False

        Try
            ' Apply alternate model configuration if user selected one
            If _alternateModelSelected AndAlso _alternateModelConfig IsNot Nothing Then
                ' Snapshot current configuration (the "original state at rest")
                backupConfig = SharedMethods.GetCurrentConfig(_context)

                ' Apply the user-selected alternate configuration
                SharedMethods.ApplyModelConfig(_context, _alternateModelConfig)
                appliedAlternate = True

                ' Enforce secondary API usage for alternate models
                _useSecondApi = True
            End If

            ' Execute the LLM call with current (possibly modified) config
            Return Await SharedMethods.LLM(_context, systemPrompt, fullPrompt, "", "", 0, _useSecondApi, True)

        Finally
            ' Always restore the original config so the rest of the add-in sees pristine state
            If appliedAlternate AndAlso backupConfig IsNot Nothing Then
                SharedMethods.RestoreDefaults(_context, backupConfig)
            End If
        End Try
    End Function

    ' =========================================================================
    ' Send Button Handler - Main Message Flow
    ' =========================================================================

    ''' <summary>
    ''' Main handler for Send button. Constructs system and user prompts based on selected
    ''' checkboxes (document, selection, other docs, commands), calls LLM, displays response,
    ''' executes any bot commands, and updates chat history. Handles errors and reports them.
    ''' </summary>
    ''' <remarks>
    ''' Execution flow:
    ''' 1. Validate user input (non-empty)
    ''' 2. Build SystemPrompt with conditional capability declarations
    ''' 3. Gather document context (active doc, selection, other docs, conversation history)
    ''' 4. Construct fullPrompt with all context elements
    ''' 5. Display "Thinking..." placeholder
    ''' 6. Call LLM asynchronously
    ''' 7. Process response (strip Markdown for commands, render HTML for display)
    ''' 8. Execute bot commands if permitted and present
    ''' 9. Update chat history (plain text and HTML)
    ''' 10. Report any errors via ReportCommandExecutionError
    ''' </remarks>
    Private Async Sub btnSend_Click(sender As Object, e As EventArgs)
        Dim userPrompt As String = txtUserInput.Text.Trim()
        If userPrompt = "" Then Return

        Dim errorOccurred As Boolean = False
        Dim errorMessage As String = ""

        ' ──────────────────────────────────────────────────────────────
        ' STEP 0: Detect and strip explicit ToolTrigger "(t)" from user prompt
        ' ──────────────────────────────────────────────────────────────
        Dim explicitToolTriggerDetected As Boolean = False
        If userPrompt.IndexOf(ToolTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
            explicitToolTriggerDetected = True
            userPrompt = userPrompt.Replace(ToolTrigger, "").Trim()

            ' If the prompt is now empty after stripping, restore for user to fix
            If String.IsNullOrWhiteSpace(userPrompt) Then
                txtUserInput.Text = ToolTrigger
                Return
            End If
        End If

        Dim promptToRestore As String = If(explicitToolTriggerDetected, $"{ToolTrigger} {userPrompt}".Trim(), userPrompt)

        Try
            My.Settings.LastPromptChat = promptToRestore
            My.Settings.Save()
        Catch
        End Try

        Try
            ' ──────────────────────────────────────────────────────────────
            ' STEP 1: Build System Prompt with Conditional Capabilities
            ' ──────────────────────────────────────────────────────────────
            ' Note: SystemPrompt is assigned twice here (legacy code pattern).
            ' The second assignment overrides the first. Keeping both for
            ' compatibility but second one is the active version.

            SystemPrompt = _context.SP_ChatWord().
                        Replace("{UserLanguage}", UserLanguage).
                        Replace("{Location}", ThisAddIn.Location) &
                        $" Your name is '{AN5}'. The current date and time is: {DateTime.Now.ToString("MMMM dd, yyyy hh:mm tt")}." &
                        If(chkIncludeDocText.Checked, vbLf & "You have access to the user's active document." & vbLf, "") &
                        If(chkIncludeselection.Checked Or chkIncludeDocText.Checked, vbLf & "You have access to the current selection or cursor context in the active document." & vbLf, "") &
                        If(chkIncludeOtherDocs.Checked, vbLf & "You also have read-only access to all other open Word documents for context only. Commands must never target those other documents." & vbLf, "") & If(My.Settings.DoCommands And (chkIncludeDocText.Checked Or chkIncludeselection.Checked),
                           _context.SP_Add_ChatWord_Commands,
                           _context.SP_Add_Chat_NoCommands)

            ' Inject InkyMemory into system prompt if enabled
            If chkInkyMemory.Checked Then
                Dim memoryContent = SharedMethods.ReadInkyMemory(_context.INI_InkyMemoryCap)
                SystemPrompt &= vbLf & _context.SP_Add_InkyMemory
                If Not String.IsNullOrWhiteSpace(memoryContent) Then
                    SystemPrompt &= vbLf & "<INKY_MEMORY_CURRENT>" & vbLf & memoryContent & vbLf & "</INKY_MEMORY_CURRENT>"
                End If
            End If

            ' Tell the model that per-message index excerpts may be supplied when an index is loaded.
            If HasLoadedIndex() Then
                SystemPrompt &= vbLf &
                    "The user has loaded a searchable index. The most relevant original excerpts for the current message " &
                    "will be provided inside <INDEX_EXCERPTS> if available. Answer also based on those excerpts, preserve exact terms, " &
                    "and never expose internal source IDs."
            End If

            If My.Settings.DoCommands AndAlso (chkIncludeDocText.Checked Or chkIncludeselection.Checked) Then
                Dim activeDocumentNameForCommands As String = ""

                Try
                    activeDocumentNameForCommands = Globals.ThisAddIn.Application.ActiveDocument.Name
                Catch
                    activeDocumentNameForCommands = "the currently active Word document"
                End Try

                SystemPrompt &= vbLf &
                    $"Command scope: All commands are executed only against the currently active Word document '{activeDocumentNameForCommands}'. " &
                    "Other open documents, if provided, are read-only context. Never issue a command for text that appears only in another document. " &
                    "If the user asks to work on another document, tell the user to activate that document first."
            End If

            ' ──────────────────────────────────────────────────────────────
            ' STEP 2: Build Conversation Context
            ' ──────────────────────────────────────────────────────────────
            Dim conversationSoFar As String = BuildConversationString(_chatHistory)

            ' Append OldChat if present (preserved from model switch or previous session)
            If Not String.IsNullOrWhiteSpace(OldChat) Then
                conversationSoFar += "\n" & OldChat
                OldChat = ""  ' Clear after use to prevent duplication
            End If

            ' ──────────────────────────────────────────────────────────────
            ' STEP 3: Validate Word Application State
            ' ──────────────────────────────────────────────────────────────
            Dim appGuard As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application

            ' If user requested document context but no document is active, abort
            If (chkIncludeDocText.Checked Or chkIncludeselection.Checked) AndAlso
               (appGuard Is Nothing OrElse
                appGuard.Documents Is Nothing OrElse
                appGuard.Documents.Count = 0 OrElse
                appGuard.ActiveDocument Is Nothing OrElse
                appGuard.ActiveWindow Is Nothing) Then

                ShowCustomMessageBox("There is no active Word document. Please open or activate a document, then try again.")
                Return
            End If

            ' ──────────────────────────────────────────────────────────────
            ' STEP 3b: Guard Against Oversized Documents (Include Document)
            ' ──────────────────────────────────────────────────────────────
            ' Including the entire document for very large files can stall Word.
            ' Read the character count cheaply via Content.End (O(1), no text
            ' extraction), and if it exceeds the threshold, automatically uncheck
            ' "Include document" and abort with a message before any expensive work.
            If chkIncludeDocText.Checked Then
                Dim documentCharacterCount As Integer = 0
                Try
                    documentCharacterCount = appGuard.ActiveDocument.Content.End
                Catch
                    documentCharacterCount = 0
                End Try

                If documentCharacterCount > MaxIncludeDocumentCharacters Then
                    chkIncludeDocText.Checked = False
                    ShowCustomMessageBox($"The active document is too large to include in full ({documentCharacterCount:N0} characters, limit is {MaxIncludeDocumentCharacters:N0}). ""Include document"" has been unchecked. Please use 'Include selection' instead after selecting the relevant portion of the document, and try again.")
                    Return
                End If
            End If

            ' ──────────────────────────────────────────────────────────────
            ' STEP 4: Gather Document Context
            ' ──────────────────────────────────────────────────────────────
            ' Extract active document text (if checkbox enabled)
            Dim docText As String = If(chkIncludeDocText.Checked, GetActiveDocumentText(), "")

            ' Extract selection text or cursor context when either selection or full document access is enabled
            Dim selectionText As String = ""

            Dim sel As Microsoft.Office.Interop.Word.Selection = Globals.ThisAddIn.Application.Selection
            If chkIncludeselection.Checked Or chkIncludeDocText.Checked Then
                selectionText = GetCurrentSelectionText()

                If sel IsNot Nothing AndAlso sel.Start = sel.End Then
                    selectionText = GetCursorContext(CursorPositionCount)
                End If
            End If

            ' Gather other open Word documents (if checkbox enabled)
            Dim otherDocs As String = ""
            If chkIncludeOtherDocs.Checked Then
                otherDocs = Globals.ThisAddIn.GatherSelectedDocuments(
                    IncludeName:=True,
                    IncludeNone:=False,
                    ExceptCurrent:=True,
                    SilentAndGetAll:=True)
            End If

            ' ──────────────────────────────────────────────────────────────
            ' STEP 5: Construct Full User Prompt
            ' ──────────────────────────────────────────────────────────────
            Dim fullPrompt As New StringBuilder()

            ' Add active document content if present
            If Not String.IsNullOrEmpty(docText) Then
                fullPrompt.AppendLine($"The user's document has the name '{Globals.ThisAddIn.Application.ActiveDocument.Name}' and has the following content: '{docText}'")

                ' Provide read-only automatic paragraph/margin numbering for the active document.
                ' Word does not expose these generated numbers via Content.Text, so we supply them
                ' separately (numbered elements only; bullets are excluded by the builder).
                Try
                    Dim activeNumbering As String = ThisAddIn.BuildParagraphNumberingContext(Globals.ThisAddIn.Application.ActiveDocument.Content)
                    If Not String.IsNullOrEmpty(activeNumbering) Then
                        fullPrompt.AppendLine($"The following is read-only automatic numbering information for the document '{Globals.ThisAddIn.Application.ActiveDocument.Name}' (it is not part of the editable text shown above; the references to TEXTTOPROCESS below mean this document):")
                        fullPrompt.AppendLine(activeNumbering)
                    End If
                Catch
                End Try
            End If

            ' Add selection or cursor context if present
            If Not String.IsNullOrEmpty(selectionText) Then
                If sel IsNot Nothing AndAlso sel.Start = sel.End Then
                    fullPrompt.AppendLine($"In the user's document '{Globals.ThisAddIn.Application.ActiveDocument.Name}' the cursor is currently positioned in the following context: '{selectionText}'")
                Else
                    fullPrompt.AppendLine($"In the user's document '{Globals.ThisAddIn.Application.ActiveDocument.Name}' the user has selected the following text: '{selectionText}'")
                End If
            End If

            ' Add other open documents if available and valid
            If chkIncludeOtherDocs.Checked AndAlso
               Not String.IsNullOrEmpty(otherDocs) AndAlso
               Not otherDocs.Equals("NONE", StringComparison.OrdinalIgnoreCase) AndAlso
               Not otherDocs.StartsWith("ERROR", StringComparison.OrdinalIgnoreCase) Then

                fullPrompt.AppendLine("The following are the other open Word documents (each enclosed in <DOCUMENTn> tags, including their name so you can refer to them):")
                fullPrompt.AppendLine(otherDocs)

                ' Provide read-only automatic numbering for each other open Word document.
                ' Mirrors the de-duplication/exclusion logic of GatherSelectedDocuments so the
                ' numbering aligns with the documents whose content was just added above.
                Try
                    Dim appDocs As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
                    Dim activeDoc As Microsoft.Office.Interop.Word.Document = Nothing
                    Try
                        activeDoc = appDocs.ActiveDocument
                    Catch
                        activeDoc = Nothing
                    End Try

                    Dim seenDocs As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                    For Each d As Microsoft.Office.Interop.Word.Document In appDocs.Documents
                        If activeDoc IsNot Nothing AndAlso Object.ReferenceEquals(d, activeDoc) Then Continue For

                        Dim key As String = If(Not String.IsNullOrEmpty(d.FullName), d.FullName, d.Name)
                        If seenDocs.Contains(key) Then Continue For
                        seenDocs.Add(key)

                        Dim otherNumbering As String = ThisAddIn.BuildParagraphNumberingContext(d.Content)
                        If Not String.IsNullOrEmpty(otherNumbering) Then
                            fullPrompt.AppendLine($"The following is read-only automatic numbering information for the open document '{d.Name}' (it is not part of the editable text; the references to TEXTTOPROCESS below mean this document):")
                            fullPrompt.AppendLine(otherNumbering)
                        End If
                    Next
                Catch
                End Try
            End If

            ' Add loaded external context if present
            If Not String.IsNullOrWhiteSpace(_loadedContextContent) Then
                fullPrompt.AppendLine("The user also loaded the following external context material:")
                fullPrompt.AppendLine("<LOADED_CONTEXT>")
                fullPrompt.AppendLine(_loadedContextContent)
                fullPrompt.AppendLine("</LOADED_CONTEXT>")
            End If

            ' The user message and conversation history are appended after any index retrieval
            ' (see Step 8b) so the index excerpts can be inserted before them.

            ' ──────────────────────────────────────────────────────────────
            ' STEP 6: Update UI - Show User Message
            ' ──────────────────────────────────────────────────────────────
            Await UpdateUIAsync(Sub()
                                    AppendToChatHistory(PreceedingNewline & "You: " & userPrompt.TrimEnd() & Environment.NewLine & Environment.NewLine)
                                    txtUserInput.Clear()
                                    PreceedingNewline = Environment.NewLine
                                End Sub)

            Await UpdateUIAsync(Sub()
                                    AppendUserHtml(userPrompt.TrimEnd())
                                End Sub)

            ' Add to in-memory history
            _chatHistory.Add(("user", userPrompt.TrimEnd()))

            ' ──────────────────────────────────────────────────────────────
            ' STEP 7: Determine if Tooling Should Be Used
            ' ──────────────────────────────────────────────────────────────
            Dim aiResponseOriginal As String

            ' Check if tooling should be used
            Dim useTooling As Boolean = False
            Dim currentConfig As ModelConfig = Nothing
            Dim toolTriggerConfig As ModelConfig = Nothing

            If _alternateModelSelected AndAlso _alternateModelConfig IsNot Nothing Then
                currentConfig = _alternateModelConfig
            Else
                currentConfig = SharedMethods.GetCurrentConfig(_context)
            End If

            Dim supportsCurrentModelTooling As Boolean = SharedMethods.ModelSupportsTooling(currentConfig)
            Dim supportsToolTrigger As Boolean =
                SharedMethods.HasToolingCapableSpecialTaskModel(_context, _context.INI_AlternateModelPath, "ToolDefaultModel")

            Dim autoToolTriggerFromCheckbox As Boolean =
                chkEnableTooling.Checked AndAlso
                Not supportsCurrentModelTooling AndAlso
                supportsToolTrigger

            Dim toolTriggerDetected As Boolean = explicitToolTriggerDetected OrElse autoToolTriggerFromCheckbox

            If toolTriggerDetected Then
                If Not SharedMethods.TryGetSpecialTaskModelConfig(
                    _context,
                    _context.INI_AlternateModelPath,
                    "ToolDefaultModel",
                    toolTriggerConfig) Then

                    Await UpdateUIAsync(Sub()
                                            ReportCommandExecutionError(
                                                $"The {ToolTrigger} trigger was requested, but no model with 'ToolDefaultModel=True' was found in the alternate model configuration. Please add a ToolDefaultModel entry to your configuration file.")
                                            txtUserInput.Text = promptToRestore
                                        End Sub)
                    Return
                End If

                If Not SharedMethods.ModelSupportsTooling(toolTriggerConfig) Then
                    Await UpdateUIAsync(Sub()
                                            ReportCommandExecutionError(
                                                $"The {ToolTrigger} trigger found a ToolDefaultModel, but it does not support {Globals.ThisAddIn.ToolFriendlyName.ToLower}. Please check the model's APICall_ToolInstructions setting.")
                                            txtUserInput.Text = promptToRestore
                                        End Sub)
                    Return
                End If

                useTooling = True
                currentConfig = toolTriggerConfig

            ElseIf chkEnableTooling.Checked AndAlso supportsCurrentModelTooling Then
                ' Standard tooling path with the currently selected model
                useTooling = True
            End If

            If useTooling Then
                If _selectedToolsForChat Is Nothing OrElse _selectedToolsForChat.Count = 0 Then
                    Dim wasTopMost As Boolean = Me.TopMost
                    Try
                        Me.TopMost = False
                        _selectedToolsForChat = Globals.ThisAddIn.SelectToolsForSession(
                            forceDialog:=toolTriggerDetected)
                    Finally
                        Me.TopMost = wasTopMost
                    End Try

                    If _selectedToolsForChat Is Nothing OrElse _selectedToolsForChat.Count = 0 Then
                        If toolTriggerDetected Then
                            Await UpdateUIAsync(Sub()
                                                    ReportCommandExecutionError(
                                                        $"The {ToolTrigger} trigger requires {Globals.ThisAddIn.ToolFriendlyName.ToLower} to be selected. Please select at least one tool and try again.")
                                                    txtUserInput.Text = promptToRestore
                                                End Sub)
                            Return
                        Else
                            useTooling = False
                        End If
                    End If
                End If
            End If

            ' ──────────────────────────────────────────────────────────────
            ' STEP 8: Display "Thinking..." Placeholder
            ' ──────────────────────────────────────────────────────────────
            Dim thinkingMessage As String = If(useTooling,
                $"{AN5}: Thinking (using {Globals.ThisAddIn.ToolFriendlyName.ToLower})...",
                $"{AN5}: Thinking...")

            Await UpdateUIAsync(Sub()
                                    AppendToChatHistory(thinkingMessage)
                                End Sub)

            Await UpdateUIAsync(Sub()
                                    ShowAssistantThinking(useTooling)
                                End Sub)

            ' ──────────────────────────────────────────────────────────────
            ' STEP 8b: Query the loaded semantic index (if any) with live progress
            ' ──────────────────────────────────────────────────────────────
            If HasLoadedIndex() Then
                Dim indexExcerpt As String = Await BuildIndexExcerptAsync(
                    userPrompt,
                    conversationSoFar,
                    Sub(status)
                        Try
                            Me.BeginInvoke(New MethodInvoker(Sub() UpdateAssistantThinking(status)))
                        Catch
                        End Try
                    End Sub)

                If Not String.IsNullOrWhiteSpace(indexExcerpt) Then
                    fullPrompt.AppendLine("The user loaded a searchable index. The following are the most relevant original excerpts for this message:")
                    fullPrompt.AppendLine("<INDEX_EXCERPTS>")
                    fullPrompt.AppendLine(indexExcerpt)
                    fullPrompt.AppendLine("</INDEX_EXCERPTS>")
                End If

                ' Restore the neutral thinking caption once retrieval progress is complete.
                Await UpdateUIAsync(Sub()
                                        UpdateAssistantThinking(If(useTooling,
                                            $"Thinking (using {Globals.ThisAddIn.ToolFriendlyName.ToLower})...",
                                            "Thinking..."))
                                    End Sub)
            End If

            ' Finalize the prompt with the current user message and conversation history.
            fullPrompt.AppendLine("User: " & userPrompt)
            fullPrompt.AppendLine($"The conversation so far (not including any previously added text document):{vbLf}{conversationSoFar}")
            Debug.WriteLine(fullPrompt.ToString())

            ' ──────────────────────────────────────────────────────────────
            ' STEP 9: Call LLM Asynchronously (with optional Tooling)
            ' ──────────────────────────────────────────────────────────────


            If useTooling AndAlso _selectedToolsForChat IsNot Nothing AndAlso _selectedToolsForChat.Count > 0 Then
                ' Apply model config temporarily for the tooling call
                Dim backupConfig As ModelConfig = Nothing
                Dim appliedOverride As Boolean = False

                Try
                    If toolTriggerDetected AndAlso toolTriggerConfig IsNot Nothing Then
                        ' (t) trigger: apply the one-shot ToolDefaultModel config
                        backupConfig = SharedMethods.GetCurrentConfig(_context)
                        SharedMethods.ApplyModelConfig(_context, toolTriggerConfig)
                        appliedOverride = True
                    ElseIf _alternateModelSelected AndAlso _alternateModelConfig IsNot Nothing Then
                        ' Standard alternate model
                        backupConfig = SharedMethods.GetCurrentConfig(_context)
                        SharedMethods.ApplyModelConfig(_context, _alternateModelConfig)
                        appliedOverride = True
                    End If

                    ' Call ExecuteToolingLoop with the same fullPrompt as non-tooling calls
                    ' hideSplash:=True suppresses splash during chat
                    ' hideLogWindow:=True suppresses log window for chat integration
                    aiResponseOriginal = Await Globals.ThisAddIn.ExecuteToolingLoop(
                        SystemPrompt,
                        userPrompt,
                        _selectedToolsForChat,
                        If(toolTriggerDetected, True, _useSecondApi),
                        fullPromptOverride:=fullPrompt.ToString(),
                        hideSplash:=True,
                        hideLogWindow:=Not chkShowToolingLog.Checked,
                        progressSink:=Sub(status)
                                          Try
                                              Me.BeginInvoke(New MethodInvoker(Sub() UpdateAssistantThinking(status)))
                                          Catch
                                          End Try
                                      End Sub)
                Finally
                    If appliedOverride AndAlso backupConfig IsNot Nothing Then
                        SharedMethods.RestoreDefaults(_context, backupConfig)
                    End If
                End Try
            Else
                ' Standard LLM call (normal behavior)
                aiResponseOriginal = Await CallLlmWithSelectedModelAsync(SystemPrompt, fullPrompt.ToString())
            End If

            ' ──────────────────────────────────────────────────────────────
            ' STEP 9: Process LLM Response
            ' ──────────────────────────────────────────────────────────────

            ' Guard against an empty/whitespace response so the chat does not
            ' render a bare "Inky:" line. Surface it as an error and restore
            ' the prompt so the user can resend without retyping.
            If String.IsNullOrWhiteSpace(aiResponseOriginal) Then
                Await UpdateUIAsync(Sub()
                                        RemoveLastLineFromChatHistory()
                                        RemoveAssistantThinking()
                                        ReportCommandExecutionError("The model returned an empty response. Please try again.")
                                        txtUserInput.Text = promptToRestore
                                    End Sub)

                ' Remove the user turn we optimistically added so history stays consistent.
                If _chatHistory.Count > 0 AndAlso _chatHistory(_chatHistory.Count - 1).Role = "user" Then
                    _chatHistory.RemoveAt(_chatHistory.Count - 1)
                End If

                Return
            End If

            ' Process InkyMemory updates from LLM response (if enabled)
            If chkInkyMemory.Checked Then
                aiResponseOriginal = SharedMethods.ProcessInkyMemoryResponse(
                    aiResponseOriginal, _context.INI_InkyMemoryCap)
            End If

            ' Keep original Markdown for HTML rendering
            Dim aiResponseMd As String = (If(aiResponseOriginal, "")).TrimEnd()

            ' Create plain text version for command parsing and persistence
            Dim aiResponsePlain As String = aiResponseMd
            ' Convert Markdown list bullets to Unicode bullet (U+2022)
            aiResponsePlain = aiResponsePlain.Replace($"{vbCrLf}* ", vbCrLf & ChrW(8226) & " ")
            aiResponsePlain = aiResponsePlain.Replace($"{vbCr}* ", vbCr & ChrW(8226) & " ")
            aiResponsePlain = aiResponsePlain.Replace($"{vbLf}* ", vbLf & ChrW(8226) & " ")
            aiResponsePlain = aiResponsePlain.Replace($"  *  ", "  " & ChrW(8226) & "  ")
            ' Strip remaining Markdown formatting
            aiResponsePlain = RemoveMarkdownFormatting(aiResponsePlain)

            ' ──────────────────────────────────────────────────────────────
            ' STEP 10: Extract and Remove Bot Commands
            ' ──────────────────────────────────────────────────────────────
            Dim CommandsString As String = ""
            If My.Settings.DoCommands And (chkIncludeselection.Checked Or chkIncludeDocText.Checked) Then
                CommandsString = aiResponsePlain
                ' Remove commands from display text
                aiResponsePlain = RemoveCommands(aiResponsePlain)
                aiResponsePlain = Regex.Replace(aiResponsePlain, "[\r\n\s]+$", "")
            End If

            ' Also remove commands from Markdown version for display
            Dim aiResponseMdDisplay As String = RemoveCommands(aiResponseMd)
            aiResponseMdDisplay = Regex.Replace(aiResponseMdDisplay, "[\r\n\s]+$", "")

            Debug.WriteLine("AI response: " & CommandsString)

            ' ──────────────────────────────────────────────────────────────
            ' STEP 11: Update UI - Show Assistant Response
            ' ──────────────────────────────────────────────────────────────
            Await UpdateUIAsync(Sub()
                                    ' Remove "Thinking..." placeholder from both views
                                    RemoveLastLineFromChatHistory()
                                    RemoveAssistantThinking()

                                    ' Append assistant answer to plain text transcript
                                    AppendToChatHistory(Environment.NewLine & $"{AN5}: " &
                                                       aiResponsePlain.TrimStart().TrimEnd().
                                                       Replace(vbCrLf, Environment.NewLine).
                                                       Replace(vbLf, Environment.NewLine) &
                                                       Environment.NewLine)

                                    ' Append assistant answer as Markdown-rendered HTML
                                    AppendAssistantMarkdown(aiResponseMdDisplay.TrimStart())

                                    ' Execute bot commands if present and permitted
                                    _keepFocusOnDocumentAfterCommands = False

                                    If My.Settings.DoCommands And Not String.IsNullOrWhiteSpace(CommandsString) Then
                                        Try
                                            ExecuteAnyCommands(CommandsString, chkIncludeselection.Checked)
                                        Catch cmdEx As Exception
                                            ' Report command execution error to chat
                                            ReportCommandExecutionError(cmdEx.Message)
                                        End Try
                                    End If

                                    ' Clear user input and restore focus unless a goto/jump/show/select
                                    ' command intentionally moved focus back to Word.
                                    txtUserInput.Text = ""
                                    If String.IsNullOrEmpty(txtUserInput.Text) AndAlso Not _keepFocusOnDocumentAfterCommands Then
                                        txtUserInput.Focus()
                                    End If
                                End Sub)

            ' Add to in-memory history
            _chatHistory.Add(("assistant", aiResponsePlain.TrimEnd()))

        Catch ex As System.Exception
            ' Capture error without performing async work inside catch block
            errorOccurred = True
            errorMessage = $"Error processing request: {ex.Message}"
        End Try

        ' ──────────────────────────────────────────────────────────────
        ' STEP 12: Handle Errors Outside Try-Catch
        ' ──────────────────────────────────────────────────────────────
        If errorOccurred Then
            Await UpdateUIAsync(Sub()
                                    ReportCommandExecutionError(errorMessage)
                                    ' Restore user input so they can try again
                                    txtUserInput.Text = userPrompt
                                End Sub)
        End If

    End Sub

    ''' <summary>
    ''' Reports command execution or LLM error to chat in both plain text and HTML formats.
    ''' Adds error message to _chatHistory so LLM can see failures in subsequent messages.
    ''' </summary>
    ''' <param name="errorMessage">Error description to display</param>
    ''' <remarks>
    ''' Error is rendered with orange/amber styling (#ff9800) to distinguish from
    ''' regular assistant messages (blue) and command failures (red).
    ''' </remarks>
    Private Sub ReportCommandExecutionError(errorMessage As String)
        If String.IsNullOrWhiteSpace(errorMessage) Then Return

        Dim errorText As String = $"⚠ Error: {errorMessage}"

        ' Add to plain text chat history with visual separator
        AppendToChatHistory(Environment.NewLine & "─────────────────────────────────────" & Environment.NewLine)
        AppendToChatHistory(errorText & Environment.NewLine)
        AppendToChatHistory("─────────────────────────────────────" & Environment.NewLine)

        ' Add to HTML chat with amber styling and inline CSS
        Dim htmlError As String = $"<div class='msg assistant error' style='border-left: 3px solid #ff9800; padding-left: 10px; margin: 10px 0; background-color: #fff3e0;'>
            <span class='who' style='color: #ff9800;'>System:</span>
            <div class='content'>
                <hr style='border: none; border-top: 1px solid #ff9800; margin: 8px 0;' />
                <strong>⚠ {HtmlEncode(errorMessage)}</strong>
                <hr style='border: none; border-top: 1px solid #ff9800; margin: 8px 0;' />
            </div>
        </div>"

        AppendHtml(htmlError)
        PersistChatHtml()

        ' Add to chat history so AI can see the error in future context
        _chatHistory.Add(("assistant", $"System Error: {errorMessage}"))
    End Sub

    ' =========================================================================
    ' Document Context Extraction
    ' =========================================================================

    ''' <summary>
    ''' Extracts text context around cursor position when no selection exists.
    ''' Returns specified number of characters before/after cursor with "[cursor is here]" marker.
    ''' Includes comments/bubbles if available in the context range.
    ''' </summary>
    ''' <param name="charCount">Number of characters to extract before and after cursor</param>
    ''' <returns>Context string with cursor marker, or empty if selection exists</returns>
    ''' <remarks>
    ''' This function provides localized document context when the user has not selected
    ''' any text. The marker "[cursor is here]" allows the LLM to understand the exact
    ''' position of user focus within the extracted context window.
    ''' 
    ''' If BubblesExtract succeeds, appends any comments/replies found in the context range.
    ''' All exceptions are silently caught to ensure function never throws.
    ''' </remarks>
    Private Function GetCursorContext(charCount As Integer) As String
        Try
            Dim activeDoc As Microsoft.Office.Interop.Word.Document = Globals.ThisAddIn.Application.ActiveDocument
            Dim sel As Microsoft.Office.Interop.Word.Selection = activeDoc.Application.Selection

            ' If actual selection exists (not just cursor position), return empty
            If Not String.IsNullOrEmpty(sel.Text) AndAlso sel.Start <> sel.End Then
                Return ""
            End If

            ' Get cursor position and document boundaries
            Dim cursorPos As Integer = sel.Start
            Dim docStart As Integer = activeDoc.Content.Start
            Dim docEnd As Integer = activeDoc.Content.End

            ' Calculate context window boundaries (clamped to document range)
            Dim contextStart As Integer = Math.Max(docStart, cursorPos - charCount)
            Dim contextEnd As Integer = Math.Min(docEnd, cursorPos + charCount)

            ' Extract text before cursor
            Dim beforeRange As Microsoft.Office.Interop.Word.Range = activeDoc.Range(contextStart, cursorPos)
            Dim textBefore As String = beforeRange.Text

            ' Extract text after cursor
            Dim afterRange As Microsoft.Office.Interop.Word.Range = activeDoc.Range(cursorPos, contextEnd)
            Dim textAfter As String = afterRange.Text

            ' Combine with cursor position marker
            Dim contextText As String = textBefore & "[cursor is here]" & textAfter

            ' Attempt to extract comments/bubbles from entire context range
            Dim bubbles As String = ""
            Try
                Dim fullContextRange As Microsoft.Office.Interop.Word.Range = activeDoc.Range(contextStart, contextEnd)
                bubbles = ThisAddIn.BubblesExtract(fullContextRange, True) ' Silent=True (no error dialogs)
            Catch
                ' Silently ignore errors; keep contextText without bubbles
            End Try

            ' Append bubbles if extracted successfully
            If Not String.IsNullOrEmpty(bubbles) Then
                Return contextText & " " & bubbles
            End If

            Return contextText

        Catch ex As Exception
            ' Silently handle any errors; return empty string
            Return ""
        End Try
    End Function


    ' =========================================================================
    ' Helper Methods - Chat UI and Utilities
    ' =========================================================================

    ''' <summary>
    ''' Generates and displays a localized welcome message when chat is first opened.
    ''' Uses current time to determine appropriate greeting (good morning/afternoon/evening).
    ''' </summary>
    ''' <returns>Empty string on success or error</returns>
    ''' <remarks>
    ''' This function constructs a minimal system prompt with assistant name and timestamp,
    ''' then asks the LLM to greet the user appropriately for the current time of day.
    ''' The response is displayed in both plain text and HTML formats and added to chat history.
    ''' All Markdown formatting is stripped from the plain text version.
    ''' </remarks>
    Private Async Function WelcomeMessage() As Task(Of String)
        Try
            ' Build system prompt with assistant identity and current timestamp
            SystemPrompt = _context.SP_ChatWord().Replace("{UserLanguage}", UserLanguage).Replace("{Location}", ThisAddIn.Location) &
                          $" Your name is '{AN5}'. The current date and time is: {DateTime.Now.ToString("F")}."
            txtUserInput.Text = ""

            ' Request localized greeting from LLM based on time of day
            Dim aiResponseRaw As String = Await CallLlmWithSelectedModelAsync(
                SystemPrompt,
                $"Welcome the user in {UserLanguage} by (1) referring to the time of day based on the current time in {UserLanguage}, such as in 'good morning', and (2) asking in {UserLanguage} what you can do, but do not say your name.")

            ' Keep Markdown for HTML display (filter bot-commands if any)
            Dim aiDisplayMd As String = RemoveCommands(If(aiResponseRaw, ""))

            ' Create plain text version by stripping all Markdown formatting
            Dim aiResponseTxt As String = If(aiResponseRaw, "")
            aiResponseTxt = aiResponseTxt.Replace(vbLf, "").Replace(vbCr, "").Replace(vbCrLf, "") & vbCrLf
            aiResponseTxt = aiResponseTxt.Replace("**", "").Replace("_", "").Replace("`", "")

            ' Update UI with both plain text and formatted HTML versions
            Await UpdateUIAsync(Sub()
                                    AppendToChatHistory(Environment.NewLine & $"{AN5}: " &
                                                       aiResponseTxt.Replace(vbCrLf, Environment.NewLine).
                                                                   Replace(vbLf, Environment.NewLine))
                                    AppendAssistantMarkdown(aiDisplayMd)
                                End Sub)

            ' Add to in-memory history for context window
            _chatHistory.Add(("assistant", aiResponseTxt))

            ' Set newline prefix for next message
            PreceedingNewline = Environment.NewLine

            Return ""

        Catch ex As System.Exception
            ' Silently handle errors; empty string indicates failure
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Converts HTML markup to plain text by stripping all tags.
    ''' Uses HtmlAgilityPack to safely parse and extract text content.
    ''' </summary>
    ''' <param name="html">HTML markup to convert</param>
    ''' <returns>Plain text with all HTML tags removed</returns>
    ''' <remarks>
    ''' This is a utility function for processing HTML content.
    ''' Currently not actively used in main workflow but available for future features.
    ''' </remarks>
    Private Function ConvertHtmlToPlainText(html As String) As String
        Dim doc As New HtmlAgilityPack.HtmlDocument()
        doc.LoadHtml(html)
        Return doc.DocumentNode.InnerText
    End Function

    ''' <summary>
    ''' Ensures UI updates occur on the correct thread using Control.Invoke pattern.
    ''' Marshals action to UI thread if called from background thread.
    ''' </summary>
    ''' <param name="action">Action to execute on UI thread</param>
    ''' <returns>Completed task</returns>
    ''' <remarks>
    ''' Essential for updating WinForms controls from async LLM calls.
    ''' Checks InvokeRequired property and uses Invoke if necessary.
    ''' If already on UI thread, executes action directly.
    ''' </remarks>
    Private Async Function UpdateUIAsync(action As System.Action) As System.Threading.Tasks.Task
        If InvokeRequired Then
            Await System.Threading.Tasks.Task.Run(Sub() Me.Invoke(action))
        Else
            action()
        End If
    End Function

    ''' <summary>
    ''' Appends text to plain text chat history control (txtChatHistory).
    ''' Thread-safe: uses Control.Invoke if called from non-UI thread.
    ''' </summary>
    ''' <param name="text">Text to append</param>
    ''' <remarks>
    ''' This maintains the plain text transcript used as fallback when HTML is unavailable.
    ''' Text is appended to end of existing content without overwriting.
    ''' </remarks>
    Private Sub AppendToChatHistory(text As String)
        If txtChatHistory.InvokeRequired Then
            txtChatHistory.Invoke(Sub() txtChatHistory.AppendText(text))
        Else
            txtChatHistory.AppendText(text)
        End If
    End Sub

    ''' <summary>
    ''' Removes the last line from plain text chat history control.
    ''' Used to remove "Thinking..." placeholder after LLM responds.
    ''' Thread-safe: uses Control.Invoke if called from non-UI thread.
    ''' </summary>
    ''' <remarks>
    ''' Recursively calls itself via Invoke if on wrong thread.
    ''' Splits text into lines array, removes last entry, and reassigns.
    ''' Safe to call when Lines array is empty (no operation performed).
    ''' </remarks>
    Private Sub RemoveLastLineFromChatHistory()
        If txtChatHistory.InvokeRequired Then
            txtChatHistory.Invoke(Sub() RemoveLastLineFromChatHistory())
        Else
            Dim lines As String() = txtChatHistory.Lines
            If lines.Length > 0 Then
                txtChatHistory.Lines = lines.Take(lines.Length - 1).ToArray()
            End If
        End If
    End Sub

    ' =========================================================================
    ' Checkbox Event Handlers - User Preferences
    ' =========================================================================

    ''' <summary>
    ''' Handles chkStayOnTop checkbox click. Toggles form's TopMost property.
    ''' </summary>
    ''' <remarks>
    ''' Inversely labeled: checkbox text is "Not always on top", but setting is NotAlwaysOnTop.
    ''' When TopMost = True, form stays on top; when False, normal behavior.
    ''' Setting persisted to My.Settings.NotAlwaysOnTop.
    ''' </remarks>
    Private Sub chkStayontop_Click(sender As Object, e As EventArgs)
        Me.TopMost = Not Me.TopMost
        My.Settings.NotAlwaysOnTop = Me.TopMost
        My.Settings.Save()
    End Sub


    ''' <summary>
    ''' Handles chkShowToolingLog checkbox change. Session-only; not persisted.
    ''' The checked state is consumed by ExecuteToolingLoop (hideLogWindow parameter)
    ''' and by the (t) trigger path, even when the checkbox is disabled.
    ''' </summary>
    Private Sub chkShowToolingLog_CheckedChanged(sender As Object, e As EventArgs)
        ' No persistence — checkbox state lives only for the current session.
        ' Value is read at call time via: hideLogWindow:=Not chkShowToolingLog.Checked
    End Sub


    ''' <summary>
    ''' Handles chkConvertMarkdown checkbox click. Persists preference for Markdown formatting.
    ''' </summary>
    ''' <remarks>
    ''' When checked, bot commands that insert text or add comments will apply Markdown formatting
    ''' via ConvertMarkdownToWord(). Setting persisted to My.Settings.ConvertMarkdownInChat.
    ''' </remarks>
    Private Sub chkConvertMarkdown_Click(sender As Object, e As EventArgs)
        My.Settings.ConvertMarkdownInChat = chkConvertMarkdown.Checked
        My.Settings.Save()
    End Sub

    ''' <summary>
    ''' Handles chkPermitCommands checkbox click. Toggles bot command execution permission.
    ''' Automatically enables document inclusion if commands enabled without selection.
    ''' </summary>
    ''' <remarks>
    ''' Commands require either document or selection to be included.
    ''' If enabling commands when neither is checked, automatically checks chkIncludeDocText.
    ''' Setting persisted to My.Settings.DoCommands.
    ''' </remarks>
    Private Sub chkPermitCommands_Click(sender As Object, e As EventArgs)
        My.Settings.DoCommands = Not My.Settings.DoCommands

        ' Auto-enable document inclusion if commands enabled without context
        If My.Settings.DoCommands And Not chkIncludeselection.Checked Then
            chkIncludeDocText.Checked = True
            My.Settings.IncludeDocument = chkIncludeDocText.Checked
        End If

        My.Settings.Save()
    End Sub

    ''' <summary>
    ''' Handles chkIncludeSelection checkbox click. Manages mutual exclusivity with document checkbox.
    ''' Validates that selection exists before allowing checkbox to remain checked.
    ''' </summary>
    ''' <remarks>
    ''' Checkbox state logic:
    ''' - If selection is empty/whitespace, unchecks itself automatically
    ''' - If document checkbox is checked, unchecks document (mutually exclusive)
    ''' - If neither selection nor document checked, disables commands
    ''' Setting persisted to My.Settings.IncludeSelection.
    ''' </remarks>
    Private Sub chkIncludeselection_Click(sender As Object, e As EventArgs)
        Dim activeDoc As Microsoft.Office.Interop.Word.Document = Globals.ThisAddIn.Application.ActiveDocument
        Dim sel As Microsoft.Office.Interop.Word.Selection = activeDoc.Application.Selection

        ' Validate selection exists
        If String.IsNullOrWhiteSpace(sel.Text) Then
            chkIncludeselection.Checked = False
        ElseIf chkIncludeDocText.Checked Then
            ' Enforce mutual exclusivity
            chkIncludeDocText.Checked = False
        End If

        My.Settings.IncludeSelection = chkIncludeselection.Checked

        ' Auto-disable commands if no context available
        If Not chkIncludeselection.Checked And Not chkIncludeDocText.Checked Then
            My.Settings.DoCommands = False
            chkPermitCommands.Checked = False
        End If

        My.Settings.Save()
    End Sub

    ''' <summary>
    ''' Handles chkIncludeDocText checkbox click. Manages mutual exclusivity with selection checkbox.
    ''' </summary>
    ''' <remarks>
    ''' Checkbox state logic:
    ''' - If selection checkbox is checked, unchecks selection (mutually exclusive)
    ''' - If neither selection nor document checked, disables commands
    ''' Setting persisted to My.Settings.IncludeDocument.
    ''' </remarks>
    Private Sub chkIncludeDocText_Click(sender As Object, e As EventArgs)
        ' Enforce mutual exclusivity
        If chkIncludeselection.Checked Then
            chkIncludeselection.Checked = False
        End If

        My.Settings.IncludeDocument = chkIncludeDocText.Checked

        ' Auto-disable commands if no context available
        If Not chkIncludeselection.Checked And Not chkIncludeDocText.Checked Then
            My.Settings.DoCommands = False
            chkPermitCommands.Checked = False
        End If

        My.Settings.Save()
    End Sub

    ''' <summary>
    ''' Handles chkEnableTooling checkbox click. Persists preference and updates related controls.
    ''' </summary>
    Private Sub chkEnableTooling_Click(sender As Object, e As EventArgs)
        My.Settings.ChatEnableTooling = chkEnableTooling.Checked
        My.Settings.Save()

        ' Clear cached tool selection when tooling disabled
        If Not chkEnableTooling.Checked Then
            _selectedToolsForChat = Nothing
        End If

        UpdateToolingControlsState()
    End Sub

    ''' <summary>
    ''' Handles chkAdvancedTools click. Persists the advanced-tools gate and clears cached effective tool selection.
    ''' </summary>
    Private Sub chkAdvancedTools_Click(sender As Object, e As EventArgs)
        My.Settings.AdvancedToolsEnabled = chkAdvancedTools.Checked
        My.Settings.Save()

        _selectedToolsForChat = Nothing
        UpdateToolingControlsState()
    End Sub

    ''' <summary>
    ''' Handles chkInkyMemory checkbox click. Persists preference and toggles edit link visibility.
    ''' </summary>
    Private Sub chkInkyMemory_Click(sender As Object, e As EventArgs)
        My.Settings.ChatInkyMemory = chkInkyMemory.Checked
        My.Settings.Save()
        lnkEditMemory.Visible = chkInkyMemory.Checked
    End Sub

    ''' <summary>
    ''' Handles chkIncludeOtherDocs checkbox click. Persists preference.
    ''' </summary>
    Private Sub chkIncludeOtherDocs_Click(sender As Object, e As EventArgs)
        My.Settings.ChatIncludeOtherOpenWordDocs = chkIncludeOtherDocs.Checked
        My.Settings.Save()
    End Sub

    ''' <summary>
    ''' Opens the Inky Memory file for manual editing.
    ''' </summary>
    Private Sub lnkEditMemory_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        SharedMethods.EditInkyMemoryFile()
    End Sub

    ' =========================================================================
    ' Button Event Handlers - User Actions
    ' =========================================================================

    ''' <summary>
    ''' Handles btnCopy click. Copies entire plain text conversation to clipboard.
    ''' </summary>
    Private Sub btnCopy_Click(sender As Object, e As EventArgs)
        My.Computer.Clipboard.SetText(txtChatHistory.Text)
    End Sub

    ''' <summary>
    ''' Handles btnCopyLastAnswer click. Copies most recent assistant response to clipboard.
    ''' Shows message box if no assistant messages exist in history.
    ''' </summary>
    ''' <remarks>
    ''' Searches _chatHistory in reverse for last message with Role = "assistant".
    ''' Only copies plain text content (Markdown already stripped).
    ''' </remarks>
    Private Sub btnCopyLastAnswer_Click(sender As Object, e As EventArgs)
        Dim lastAssistantMsg = _chatHistory.Where(Function(x) x.Role = "assistant").LastOrDefault()
        If lastAssistantMsg.Content IsNot Nothing Then
            My.Computer.Clipboard.SetText(lastAssistantMsg.Content)
        Else
            SharedMethods.ShowCustomMessageBox("No last AI answer available.")
        End If
    End Sub

    ''' <summary>
    ''' Handles btnTools click. Opens tool selection dialog and caches the selection for this chat session.
    ''' </summary>
    ''' <param name="sender">Event sender.</param>
    ''' <param name="e">Event arguments.</param>
    ''' <remarks>
    ''' Temporarily disables TopMost so the selection dialog is not blocked by the chat form.
    ''' The selected tool set is stored in _selectedToolsForChat and used by ExecuteToolingLoop.
    ''' </remarks>
    Private Sub btnTools_Click(sender As Object, e As EventArgs)
        Dim wasTopMost As Boolean = Me.TopMost
        Try
            Me.TopMost = False
            Dim selectedTools = Globals.ThisAddIn.SelectToolsForSession(forceDialog:=True, Globals.ThisAddIn.ToolFriendlyName)
            If selectedTools IsNot Nothing Then
                _selectedToolsForChat = selectedTools
            End If
        Finally
            Me.TopMost = wasTopMost
        End Try
    End Sub

    ''' <summary>
    ''' Updates enabled/disabled state for tooling-related controls.
    ''' </summary>
    ''' <remarks>
    ''' The tools button and the enable-tooling checkbox are available when either the
    ''' current model supports tooling or a tooling-capable ToolDefaultModel exists for "(t)".
    ''' 
    ''' When only ToolDefaultModel is available, checking "Enable tooling" means:
    ''' treat every request as if it had "(t)".
    ''' </remarks>
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

        chkEnableTooling.Enabled = toolingUiAvailable
        chkAdvancedTools.Enabled = toolingUiAvailable AndAlso chkEnableTooling.Checked
        btnTools.Enabled = toolingUiAvailable
        chkShowToolingLog.Enabled = toolingUiAvailable

        If Not toolingUiAvailable Then
            chkEnableTooling.Checked = False
            _selectedToolsForChat = Nothing
        End If

        If Not _toolingControlsInitialized Then
            chkShowToolingLog.Checked = _context.INI_ToolingLogWindow
            _toolingControlsInitialized = True
        End If
    End Sub


    ' =========================================================================
    ' Model Switching
    ' =========================================================================

    ''' <summary>
    ''' Handles btnSwitchModel click. Toggles between primary/secondary/alternate models.
    ''' Implements snapshot/restore pattern to keep global context pristine.
    ''' Persists selection to My.Settings for restoration on next session.
    ''' </summary>
    ''' <remarks>
    ''' Behavior depends on configuration:
    ''' 
    ''' When Alternate Model INI configured (_context.INI_AlternateModelPath):
    '''   - If alternate already selected: switches back to primary immediately
    '''   - If primary active: shows model selection dialog
    '''   - After selection: snapshots config, restores original, stores snapshot for later use
    '''   - This pattern prevents alternate config from polluting global SharedContext
    ''' 
    ''' When only Primary/Secondary configured (legacy mode):
    '''   - Simple toggle of _useSecondApi flag
    '''   - No dialog shown
    ''' 
    ''' All model switches trigger:
    '''   - UpdateModelButtonText() to reflect new state
    '''   - UpdateTitle() to show active model in window title
    '''   - UpdateDocumentCheckboxesState() to disable checkboxes if secondary/alternate active
    ''' </remarks>
    Private Sub btnSwitchModel_Click(sender As Object, e As EventArgs)
        If Not String.IsNullOrWhiteSpace(_context.INI_AlternateModelPath) Then
            ' ─────────────────────────────────────────────────────────────
            ' Alternate Model Path Configured
            ' ─────────────────────────────────────────────────────────────

            ' If alternate already active, switch back to primary
            If _alternateModelSelected Then
                _alternateModelSelected = False
                _alternateModelConfig = Nothing
                _alternateModelDisplayName = Nothing
                _useSecondApi = False
                UpdateModelButtonText()
                UpdateTitle()
                UpdateDocumentCheckboxesState()
                PersistAlternateModelToSettings()
                Return
            End If

            ' Temporarily disable TopMost so dialog is not blocked
            Dim wasTopMost As Boolean = Me.TopMost
            Try
                Me.TopMost = False

                ' Show model selection dialog
                SharedMethods.LastAlternateModel = "" ' Sentinel value
                Dim ok As Boolean = SharedMethods.ShowModelSelection(
                    _context,
                    _context.INI_AlternateModelPath,
                    "Alternate Model",
                    "Select the alternate model you want to use:",
                    "",
                    2)

                If Not ok Then
                    ' User cancelled dialog
                    Return
                End If

                ' ─────────────────────────────────────────────────────────────
                ' Snapshot Pattern: Capture alternate config then restore original
                ' ─────────────────────────────────────────────────────────────
                Dim justApplied As ModelConfig = SharedMethods.GetCurrentConfig(_context)

                ' Restore original config immediately
                If SharedMethods.originalConfigLoaded Then
                    SharedMethods.RestoreDefaults(_context, SharedMethods.originalConfig)
                End If
                SharedMethods.originalConfigLoaded = False

                ' Check if user actually selected an alternate (vs. primary)
                Dim userChoseAlternate As Boolean = Not String.IsNullOrWhiteSpace(SharedMethods.LastAlternateModel)

                If userChoseAlternate Then
                    ' Store snapshot for use during LLM calls
                    _alternateModelSelected = True
                    _alternateModelConfig = justApplied
                    _alternateModelDisplayName = SharedMethods.LastAlternateModel
                    _useSecondApi = True
                Else
                    ' User selected primary model from dialog
                    _alternateModelSelected = False
                    _alternateModelConfig = Nothing
                    _alternateModelDisplayName = Nothing
                    _useSecondApi = False
                End If

            Finally
                Me.TopMost = wasTopMost
            End Try

            UpdateModelButtonText()
            UpdateTitle()
            UpdateDocumentCheckboxesState()
            PersistAlternateModelToSettings()
        Else
            ' ─────────────────────────────────────────────────────────────
            ' Legacy Mode: Simple toggle between primary and secondary
            ' ─────────────────────────────────────────────────────────────
            _useSecondApi = Not _useSecondApi
            _alternateModelSelected = False
            _alternateModelConfig = Nothing
            _alternateModelDisplayName = Nothing
            UpdateModelButtonText()
            UpdateTitle()
            UpdateDocumentCheckboxesState()
            PersistAlternateModelToSettings()
        End If
    End Sub

    ''' <summary>
    ''' Updates btnSwitchModel text to reflect current model selection state.
    ''' </summary>
    ''' <remarks>
    ''' When an alternate model INI is configured, this button toggles between primary and an alternate selection:
    ''' - Shows "Primary model" when an alternate model is active.
    ''' - Shows "Alternate Model" when the primary model is active.
    ''' Otherwise (no alternate INI), the button uses a generic "Switch Model" label.
    ''' </remarks>
    Private Sub UpdateModelButtonText()
        If Not String.IsNullOrWhiteSpace(_context.INI_AlternateModelPath) Then
            btnSwitchModel.Text = If(_alternateModelSelected, "Primary model", "Alternate Model")
        Else
            btnSwitchModel.Text = "Switch Model"
        End If
    End Sub


    ' =========================================================================
    ' Alternate Model Persistence
    ' =========================================================================

    ''' <summary>
    ''' Persists the current alternate model selection to My.Settings.
    ''' Only saves the display name - config is reloaded from INI on next session.
    ''' </summary>
    Private Sub PersistAlternateModelToSettings()
        Try
            If _alternateModelSelected AndAlso Not String.IsNullOrWhiteSpace(_alternateModelDisplayName) Then
                My.Settings.ChatAlternateModelName = _alternateModelDisplayName
            Else
                My.Settings.ChatAlternateModelName = ""
            End If
            My.Settings.Save()
        Catch ex As Exception
            Debug.WriteLine($"PersistAlternateModelToSettings error: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Restores the alternate model selection from My.Settings by looking up
    ''' the saved model name in the alternate models INI file.
    ''' Falls back to primary model if saved model is no longer available.
    ''' </summary>
    Private Sub RestoreAlternateModelFromSettings()
        Try
            Dim savedName As String = My.Settings.ChatAlternateModelName

            If String.IsNullOrWhiteSpace(savedName) Then
                ' No saved alternate model - use primary
                Return
            End If

            If String.IsNullOrWhiteSpace(_context.INI_AlternateModelPath) Then
                ' No alternate model INI configured - clear saved setting
                My.Settings.ChatAlternateModelName = ""
                My.Settings.Save()
                Return
            End If

            ' Load all available alternate models from INI
            Dim availableModels As List(Of ModelConfig) = SharedMethods.LoadAlternativeModels(
                _context.INI_AlternateModelPath,
                _context,
                "Chat Alternate Model",
                includeToolOnly:=False,
                toolsOnly:=False)

            If availableModels Is Nothing OrElse availableModels.Count = 0 Then
                ' No models available - clear saved setting and use primary
                My.Settings.ChatAlternateModelName = ""
                My.Settings.Save()
                Return
            End If

            ' Find the saved model by display name (ModelDescription)
            Dim matchedModel As ModelConfig = availableModels.FirstOrDefault(
                Function(m) String.Equals(m.ModelDescription, savedName, StringComparison.OrdinalIgnoreCase))

            If matchedModel Is Nothing Then
                ' Saved model no longer available - clear setting and use primary
                Debug.WriteLine($"RestoreAlternateModelFromSettings: Model '{savedName}' no longer available, using primary")
                My.Settings.ChatAlternateModelName = ""
                My.Settings.Save()
                Return
            End If

            ' Found the model - apply it
            _alternateModelSelected = True
            _alternateModelConfig = matchedModel
            _alternateModelDisplayName = savedName
            _useSecondApi = True

            UpdateModelButtonText()
            UpdateDocumentCheckboxesState()

            Debug.WriteLine($"RestoreAlternateModelFromSettings: Restored alternate model '{savedName}'")

        Catch ex As Exception
            Debug.WriteLine($"RestoreAlternateModelFromSettings error: {ex.Message}")
            ' On error, clear the persisted setting and use primary
            Try
                My.Settings.ChatAlternateModelName = ""
                My.Settings.Save()
            Catch
            End Try
        End Try
    End Sub

    ''' <summary>
    ''' Updates checkbox states when secondary or alternate API is active.
    ''' Disables document-related checkboxes for secondary/alternate models.
    ''' </summary>
    ''' <remarks>
    ''' When _useSecondApi = True:
    ''' - Unchecks and saves: chkIncludeDocText, chkIncludeSelection, 
    '''   chkPermitCommands, chkIncludeOtherDocs
    ''' - Updates My.Settings accordingly
    ''' 
    ''' When _useSecondApi = False:
    ''' - Re-enables checkboxes
    ''' - Does NOT automatically restore previous checked states (commented out)
    ''' 
    ''' Rationale: Secondary/alternate models may not support document context features.
    ''' </remarks>
    Private Sub UpdateDocumentCheckboxesState()
        If _useSecondApi Then
            ' Disable document-related features for secondary/alternate models
            chkIncludeDocText.Checked = False
            chkIncludeselection.Checked = False
            chkPermitCommands.Checked = False
            chkIncludeOtherDocs.Checked = False

            ' Persist disabled state
            My.Settings.IncludeDocument = False
            My.Settings.IncludeSelection = False
            My.Settings.DoCommands = False
            My.Settings.Save()
        Else
            ' Re-enable checkboxes when switching back to primary model
            chkIncludeDocText.Enabled = True
            chkIncludeselection.Enabled = True
            chkPermitCommands.Enabled = True
        End If

        ' Update tooling controls based on new model
        UpdateToolingControlsState()
    End Sub

    ' =========================================================================
    ' Conversation Management
    ' =========================================================================

    ''' <summary>
    ''' Handles btnClear click. Clears all conversation history and displays fresh welcome message.
    ''' </summary>
    ''' <remarks>
    ''' Clears:
    ''' - In-memory _chatHistory list
    ''' - Plain text txtChatHistory control
    ''' - OldChat preserved context
    ''' - PreceedingNewline formatting state
    ''' - Both My.Settings.LastChatHistory and LastChatHistoryHtml
    ''' - HTML WebBrowser content via ClearChatHtml()
    ''' 
    ''' Then displays new welcome message asynchronously.
    ''' </remarks>
    Private Async Sub btnClear_Click(sender As Object, e As EventArgs)
        _chatHistory.Clear()
        txtChatHistory.Clear()
        OldChat = ""
        PreceedingNewline = ""
        My.Settings.LastChatHistory = ""
        My.Settings.LastChatHistoryHtml = ""
        My.Settings.Save()

        ClearChatHtml()

        Await WelcomeMessage()
    End Sub

    ' =========================================================================
    ' Form Closing and Keyboard Handlers
    ' =========================================================================

    ''' <summary>
    ''' Handles form-level KeyDown event. Closes form when ESC pressed.
    ''' Saves conversation history (trimmed to INI_ChatCap) and HTML before closing.
    ''' </summary>
    ''' <remarks>
    ''' Requires KeyPreview = True (set in Load event).
    ''' Conversation trimming ensures settings don't grow unbounded.
    ''' Both plain text and HTML history persisted before close.
    ''' </remarks>
    Private Sub frmAIChat_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        If e.KeyCode = Keys.Escape Then

            Try
                RemoveHandler Microsoft.Win32.SystemEvents.DisplaySettingsChanged, AddressOf OnDisplaySettingsChanged
            Catch
            End Try

            ' Trim conversation to capacity limit
            Dim conversation As String = txtChatHistory.Text
            If conversation.Length > _context.INI_ChatCap Then
                conversation = conversation.Substring(conversation.Length - _context.INI_ChatCap)
            End If

            My.Settings.LastChatHistory = conversation
            PersistChatHtml()
            My.Settings.Save()
            Close()
        End If
    End Sub

    ''' <summary>
    ''' Handles btnExit click. Same behavior as ESC key (saves state and closes).
    ''' </summary>
    Private Sub btnExit_Click(sender As Object, e As EventArgs)

        Try
            RemoveHandler Microsoft.Win32.SystemEvents.DisplaySettingsChanged, AddressOf OnDisplaySettingsChanged
        Catch
        End Try

        ' Trim conversation to capacity limit
        Dim conversation As String = txtChatHistory.Text
        If conversation.Length > _context.INI_ChatCap Then
            conversation = conversation.Substring(conversation.Length - _context.INI_ChatCap)
        End If

        My.Settings.LastChatHistory = conversation
        PersistChatHtml()
        My.Settings.Save()
        Close()
    End Sub

    ''' <summary>
    ''' Handles FormClosing event. Persists conversation and window state before form closes.
    ''' </summary>
    ''' <remarks>
    ''' Saves:
    ''' - Conversation history (trimmed to INI_ChatCap)
    ''' - Form location and size (uses RestoreBounds if minimized/maximized)
    ''' - HTML chat content via PersistChatHtml()
    ''' 
    ''' RestoreBounds ensures correct position/size restored when user maximized/minimized.
    ''' </remarks>
    Private Sub frmAIChat_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        Try
            RemoveHandler Microsoft.Win32.SystemEvents.DisplaySettingsChanged, AddressOf OnDisplaySettingsChanged
        Catch
        End Try

        ' Trim and save conversation
        Dim conversation As String = txtChatHistory.Text
        If conversation.Length > _context.INI_ChatCap Then
            conversation = conversation.Substring(conversation.Length - _context.INI_ChatCap)
        End If
        My.Settings.LastChatHistory = conversation

        ' Save window position and size
        If Me.WindowState = FormWindowState.Normal Then
            My.Settings.FormLocation = Me.Location
            My.Settings.FormSize = Me.Size
        Else
            ' Use RestoreBounds for minimized/maximized states
            My.Settings.FormLocation = Me.RestoreBounds.Location
            My.Settings.FormSize = Me.RestoreBounds.Size
        End If

        PersistChatHtml()
        My.Settings.Save()
    End Sub

    ' =========================================================================
    ' Input Keyboard Handlers
    ' =========================================================================


    ''' <summary>
    ''' Handles KeyDown event for txtUserInput. Sends message on Enter, allows Shift+Enter for newline.
    ''' </summary>
    ''' <remarks>
    ''' Keyboard behavior:
    ''' - Enter alone: Triggers Send button (e.SuppressKeyPress prevents actual newline insertion)
    ''' - Shift+Enter: Inserts newline (default TextBox behavior, no action taken)
    ''' 
    ''' This provides familiar chat UI pattern matching modern messaging apps.
    ''' Handler attached in frmAIChat_Load event.
    ''' </remarks>
    Private Sub UserInput_KeyDown(sender As Object, e As KeyEventArgs)
        If e.Control AndAlso e.KeyCode = Keys.P Then
            Dim lastPrompt As String = My.Settings.LastPromptChat

            If Not String.IsNullOrWhiteSpace(lastPrompt) Then
                Dim insertionIndex As Integer = txtUserInput.SelectionStart
                Dim selectionLength As Integer = txtUserInput.SelectionLength

                Dim newText As String =
                    txtUserInput.Text.Remove(insertionIndex, selectionLength).Insert(insertionIndex, lastPrompt)

                txtUserInput.Text = newText
                txtUserInput.SelectionStart = insertionIndex + lastPrompt.Length
                txtUserInput.SelectionLength = 0
            End If

            e.SuppressKeyPress = True
            e.Handled = True
            Return
        End If

        If e.KeyCode = Keys.Enter Then
            If e.Shift Then
                Return
            Else
                e.SuppressKeyPress = True
                btnSend.PerformClick()
                e.Handled = True
            End If
        End If
    End Sub

    ''' <summary>
    ''' Handles slash-triggered prompt library insertion for the chat input box.
    ''' </summary>
    Private Sub UserInput_KeyPress(sender As Object, e As KeyPressEventArgs)
        If e.KeyChar <> "/"c Then Return
        If Not _context.INI_PromptLib Then Return

        Dim slashAction As SharedMethods.PromptLibrarySlashAction =
            SharedMethods.HandlePromptLibrarySlash(
                txtUserInput,
                _context.INI_PromptLibPath,
                _context.INI_PromptLibPathLocal,
                _context,
                My.Settings.LastPromptChat
            )

        If slashAction <> SharedMethods.PromptLibrarySlashAction.NotTriggered Then
            e.Handled = True
        End If
    End Sub

    ' =========================================================================
    ' Document Text Extraction
    ' =========================================================================

    ''' <summary>
    ''' Extracts complete text content from active Word document.
    ''' Temporarily switches to Final view mode to exclude tracked deletions.
    ''' Optionally appends comments/bubbles if BubblesExtract succeeds.
    ''' </summary>
    ''' <returns>Document text with optional comments, or empty string on error</returns>
    ''' <remarks>
    ''' View management:
    ''' - Saves original RevisionsView and ShowRevisionsAndComments settings
    ''' - Temporarily sets to wdRevisionsViewFinal with ShowRevisionsAndComments = False
    ''' - Restores original settings in Finally block
    ''' 
    ''' This ensures LLM sees only accepted text, not deleted content or markup.
    ''' Comments/bubbles appended as separate block if extraction succeeds (Silent=True).
    ''' All exceptions caught and return empty string.
    ''' </remarks>
    Private Function GetActiveDocumentText() As String
        Try
            Dim doc As Microsoft.Office.Interop.Word.Document = Globals.ThisAddIn.Application.ActiveDocument
            Dim wordApp As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application

            ' Capture the specific window we modify so the Finally restore targets the SAME
            ' window even if the active window changes meanwhile (e.g., a document-management
            ' add-in switches documents during the operation).
            Dim targetWindow As Word.Window = wordApp.ActiveWindow

            ' Save current view settings for restoration
            Dim originalRevisionsView As Word.WdRevisionsView = targetWindow.View.RevisionsView
            Dim originalShowRevisions As Boolean = targetWindow.View.ShowRevisionsAndComments

            Try
                ' Temporarily show only final text (no tracked deletions)
                With targetWindow.View
                    .RevisionsView = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal
                    .ShowRevisionsAndComments = False
                End With

                ' Extract document text (excludes deleted content)
                Dim baseText As String = doc.Content.Text

                ' Attempt to extract comments/bubbles
                Dim bubbles As String = ""
                Try
                    bubbles = ThisAddIn.BubblesExtract(doc.Content, True) ' Silent=True
                Catch
                    ' Silently ignore errors; keep baseText only
                End Try

                ' Append bubbles if available
                If Not String.IsNullOrEmpty(bubbles) Then
                    Return baseText & vbCr & vbCr & bubbles
                End If

                Return baseText

            Finally
                ' Restore original view settings on the SAME window we changed
                Try
                    With targetWindow.View
                        .RevisionsView = originalRevisionsView
                        .ShowRevisionsAndComments = originalShowRevisions
                    End With
                Catch
                    ' Best-effort restore; the window may no longer be valid
                End Try
            End Try

        Catch ex As Exception
            ' Silently handle all errors
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Extracts text from current Word selection.
    ''' Temporarily switches to Final view mode to exclude tracked deletions.
    ''' Optionally appends comments/bubbles if BubblesExtract succeeds.
    ''' </summary>
    ''' <returns>Selection text with optional comments, or empty string if no selection or error</returns>
    ''' <remarks>
    ''' Behavior:
    ''' - If selection is empty/null: unchecks chkIncludeSelection and returns empty string
    ''' - If selection exists: extracts text in Final view mode (no deletions)
    ''' - Attempts to extract comments/bubbles from selection range
    ''' - Appends bubbles inline (space-separated) if available
    ''' 
    ''' View management identical to GetActiveDocumentText (save/restore pattern).
    ''' All exceptions caught and return empty string.
    ''' </remarks>
    Private Function GetCurrentSelectionText() As String
        Try
            Dim activeDoc As Microsoft.Office.Interop.Word.Document = Globals.ThisAddIn.Application.ActiveDocument
            Dim wordApp As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
            Dim sel As Microsoft.Office.Interop.Word.Selection = activeDoc.Application.Selection

            ' Validate selection exists
            If String.IsNullOrEmpty(sel.Text) Then
                chkIncludeselection.Checked = False
                Return ""
            Else
                ' Capture the specific window we modify so the Finally restore targets the SAME
                ' window even if the active window changes meanwhile.
                Dim targetWindow As Word.Window = wordApp.ActiveWindow

                ' Save current view settings
                Dim originalRevisionsView As Word.WdRevisionsView = targetWindow.View.RevisionsView
                Dim originalShowRevisions As Boolean = targetWindow.View.ShowRevisionsAndComments

                Try
                    ' Temporarily show only final text
                    With targetWindow.View
                        .RevisionsView = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal
                        .ShowRevisionsAndComments = False
                    End With

                    ' Extract selection text
                    Dim baseText As String = sel.Text

                    ' Attempt to extract comments/bubbles from selection
                    Dim bubbles As String = ""
                    Try
                        bubbles = ThisAddIn.BubblesExtract(sel.Range, True) ' Silent=True
                    Catch
                        ' Silently ignore errors
                    End Try

                    ' Append bubbles inline if available
                    If Not String.IsNullOrEmpty(bubbles) Then
                        Return baseText & " " & bubbles
                    End If

                    Return baseText

                Finally
                    ' Restore original view settings on the SAME window we changed
                    Try
                        With targetWindow.View
                            .RevisionsView = originalRevisionsView
                            .ShowRevisionsAndComments = originalShowRevisions
                        End With
                    Catch
                        ' Best-effort restore; the window may no longer be valid
                    End Try
                End Try
            End If
        Catch ex As Exception
            ' Silently handle all errors
            Return ""
        End Try
    End Function

    ' =========================================================================
    ' Conversation History Management
    ' =========================================================================

    ''' <summary>
    ''' Builds conversation history string from in-memory _chatHistory list.
    ''' Trims to INI_ChatCap character limit, keeping most recent messages.
    ''' </summary>
    ''' <param name="history">List of (Role, Content) tuples representing conversation</param>
    ''' <returns>Formatted conversation string with "User:" and assistant name prefixes</returns>
    ''' <remarks>
    ''' Processing:
    ''' 1. Iterates history in reverse (most recent first)
    ''' 2. Formats each message with "User:" or "{AN5}:" prefix
    ''' 3. Accumulates messages until INI_ChatCap limit reached
    ''' 4. If adding message would exceed limit, truncates it to fit
    ''' 5. Uses StringBuilder.Insert(0, ...) to maintain chronological order
    ''' 
    ''' This ensures LLM always sees most recent context within token limits.
    ''' Older messages dropped when capacity exceeded.
    ''' </remarks>
    Private Function BuildConversationString(history As List(Of (Role As String, Content As String))) As String
        Dim sb As New StringBuilder()
        Dim totalLength As Integer = 0
        Dim maxLength As Integer = _context.INI_ChatCap

        ' Iterate in reverse to prioritize recent messages
        For Each msg In history.AsEnumerable().Reverse()
            ' Format message with role prefix
            Dim message As String
            If msg.Role = "user" Then
                message = $"User: {msg.Content}{Environment.NewLine}"
            Else
                message = $"{AN5}: {msg.Content}{Environment.NewLine}"
            End If

            ' Check if adding message exceeds capacity
            If totalLength + message.Length > maxLength Then
                ' Truncate message to fit within remaining space
                Dim remainingLength As Integer = maxLength - totalLength
                If remainingLength > 0 Then
                    sb.Insert(0, message.Substring(0, remainingLength))
                End If
                Exit For
            Else
                ' Add full message (Insert at position 0 maintains order)
                sb.Insert(0, message)
                totalLength += message.Length
            End If
        Next

        Return sb.ToString()
    End Function



    ' =========================================================================
    ' Text Normalization Utilities
    ' =========================================================================

    ''' <summary>
    ''' Normalizes various paragraph mark encodings to Word's native format (vbCr).
    ''' Handles: actual control chars, Word Find tokens (^p, ^13), literal escape sequences (\r\n, \n, \r).
    ''' Used to ensure LLM-generated text matches Word's internal paragraph representation.
    ''' </summary>
    ''' <param name="raw">Text potentially containing mixed paragraph mark encodings</param>
    ''' <returns>Normalized text with consistent vbCr paragraph marks</returns>
    ''' <remarks>
    ''' Processing order (critical for correct behavior):
    ''' <para>
    ''' 1. Unify actual control characters first:
    ''' vbCrLf → vbCr, vbLf → vbCr
    ''' </para>
    ''' <para>
    ''' 2. Word Find tokens to vbCr:
    ''' ^p → vbCr (case-insensitive)
    ''' ^13 or ^013 → vbCr (optional leading zeros)
    ''' </para>
    ''' <para>
    ''' 3. Convert literal escape sequences from LLM output:
    ''' \r\n → vbCr (treat as single paragraph)
    ''' \r → vbCr
    ''' \n → vbCr
    ''' Only when NOT double-escaped (negative lookbehind (?&lt;!\\) ignores \\r, \\n)
    ''' </para>
    ''' <para>
    ''' 4. Optional collapse multiple consecutive paragraphs (commented out by default)
    ''' </para>
    ''' <para>
    ''' This handles mixed encodings from LLMs that may output \n, Word Find that uses ^p,
    ''' and actual control characters from clipboard/other sources.
    ''' </para>
    ''' </remarks>
    Private Function DecodeParagraphMarks(raw As String) As String
        If String.IsNullOrEmpty(raw) Then Return ""

        ' Step 1: Unify actual control characters
        raw = raw.Replace(vbCrLf, vbCr).Replace(vbLf, vbCr)

        ' Step 2: Word Find tokens → vbCr
        raw = Regex.Replace(raw, "\^p", vbCr, RegexOptions.IgnoreCase)
        raw = Regex.Replace(raw, "\^0*13", vbCr, RegexOptions.IgnoreCase)

        ' Step 3: Convert literal (escaped) sequences from LLM output
        ' Only when NOT double-escaped (negative lookbehind prevents matching \\r, \\n)
        raw = Regex.Replace(raw, "(?<!\\)\\r\\n", vbCr, RegexOptions.IgnoreCase)
        raw = Regex.Replace(raw, "(?<!\\)\\r", vbCr, RegexOptions.IgnoreCase)
        raw = Regex.Replace(raw, "(?<!\\)\\n", vbCr, RegexOptions.IgnoreCase)

        ' Step 4: Optional collapse multiple consecutive paragraphs
        ' Commented out to preserve intentional empty lines
        ' Uncomment if you want to collapse: vbCr & vbCr & vbCr → vbCr & vbCr
        ' raw = Regex.Replace(raw, vbCr & "{2,}", vbCr & vbCr)

        Return raw
    End Function

    ''' <summary>
    ''' Ensures text has properly decoded paragraph marks by calling DecodeParagraphMarks.
    ''' Wrapper function for clarity when intent is to ensure proper formatting.
    ''' </summary>
    ''' <param name="text">Text to normalize</param>
    ''' <returns>Normalized text</returns>
    Private Function EnsureParagraphs(text As String) As String
        If String.IsNullOrEmpty(text) Then Return ""
        Return DecodeParagraphMarks(text)
    End Function

    ''' <summary>
    ''' Cleans bot command arguments by normalizing paragraph marks and trimming spaces/tabs.
    ''' Preserves intentional leading/trailing paragraph marks.
    ''' </summary>
    ''' <param name="arg">Command argument to clean</param>
    ''' <returns>Cleaned argument</returns>
    ''' <remarks>
    ''' Processing:
    ''' 1. Returns empty string if arg is Nothing
    ''' 2. Decodes paragraph marks to vbCr
    ''' 3. Trims only spaces/tabs (regex ^[ \t]+ and [ \t]+$)
    ''' 4. Intentionally preserves leading/trailing vbCr if present
    ''' 
    ''' This allows LLM to specify arguments with intentional leading/trailing newlines
    ''' while removing accidental whitespace from formatting.
    ''' </remarks>
    Private Function CleanArgument(arg As String) As String
        If arg Is Nothing Then Return ""
        arg = DecodeParagraphMarks(arg)
        ' Strip Word cell end marker Chr(7) — appears as a dot in table cell text
        arg = arg.TrimStart(ChrW(7)).TrimEnd(ChrW(7))
        ' Trim only spaces/tabs, preserve paragraph marks
        Return Regex.Replace(arg, "^[ \t]+|[ \t]+$", "")
    End Function


    ' =========================================================================
    ' Bot Command Parsing
    ' =========================================================================

    ''' <summary>
    ''' Nested class representing a parsed bot command with verb and arguments.
    ''' </summary>
    ''' <remarks>
    ''' Properties:
    ''' - Command: Verb (e.g., "replace", "insert", "find")
    ''' - Argument1: First argument (e.g., search term for replace)
    ''' - Argument2: Second argument (optional, e.g., replacement text)
    ''' 
    ''' Second argument may be empty for delete operations (replace with nothing).
    ''' </remarks>
    Public Class ParsedCommand
        Public Property Command As String
        Public Property Argument1 As String
        Public Property Argument2 As String
    End Class

    ''' <summary>
    ''' Parses embedded bot commands from LLM response using pattern: [#verb: @@arg1@@ §§arg2§§ #]
    ''' Supports tempered-greedy matching: single @ or § allowed inside args, only @@ or §§ terminate.
    ''' Second argument is optional (defaults to empty string for delete operations).
    ''' </summary>
    ''' <param name="input">Text potentially containing command blocks</param>
    ''' <returns>List of parsed commands with verb and arguments</returns>
    ''' <remarks>
    ''' Regex Pattern Explanation:
    ''' 
    ''' Pattern: \[#(?&lt;cmd&gt;[^:]+):\s*@@(?&lt;arg1&gt;(?:[^@]|@(?!@))*?)@@\s*(?:§§(?&lt;arg2&gt;(?:[^§]|§(?!§))*?)§§)?\s*#\]
    ''' 
    ''' Breakdown (left to right):
    ''' 
    ''' \[#                          - Literal "[#" opens command block
    ''' (?&lt;cmd&gt;[^:]+)                 - Named group "cmd": one or more chars except ":"
    '''                                 Ends at first colon
    ''' :\s*@@                       - Literal ":" + whitespace + exactly two @
    ''' (?&lt;arg1&gt;(?:[^@]|@(?!@))*?)  - Named group "arg1" with tempered greedy token:
    '''                                 (?:[^@]|@(?!@))* matches any char except @,
    '''                                 OR single @ not followed by another @
    '''                                 Stops only at @@
    ''' @@\s*                        - End delimiter for arg1
    ''' (?:§§(?&lt;arg2&gt;...)§§)?        - Optional arg2 block (same tempered greedy logic)
    ''' \s*#\]                       - Close delimiter
    ''' 
    ''' Tempered Greedy Token Pattern:
    ''' The pattern (?:[^@]|@(?!@))*? allows:
    ''' - Any character except @
    ''' - Single @ if NOT followed by another @ (negative lookahead (?!@))
    ''' This permits email addresses (user@domain) inside arguments while still
    ''' treating @@ as the terminator.
    ''' 
    ''' Duplicate Detection:
    ''' Results.Any check prevents duplicate commands with identical verb and arguments.
    ''' 
    ''' Error Handling:
    ''' MsgBox shown on regex errors (shouldn't happen with valid pattern).
    ''' </remarks>
    Private Function ParseCommands(input As String) As List(Of ParsedCommand)
        Dim results As New List(Of ParsedCommand)
        Try
            ' Tempered-greedy regex pattern for command parsing
            ' See function remarks for detailed explanation
            Dim pattern As String = "\[#(?<cmd>[^:]+):\s*@@(?<arg1>(?:[^@]|@(?!@))*?)@@\s*(?:§§(?<arg2>(?:[^§]|§(?!§))*?)§§)?\s*#\]"
            Dim regex As New Regex(pattern, RegexOptions.Singleline)

            For Each m As Match In regex.Matches(input)
                Dim pc As New ParsedCommand()
                pc.Command = m.Groups("cmd").Value.Trim()

                ' Extract raw arguments
                Dim raw1 As String = m.Groups("arg1").Value
                Dim raw2 As String = If(m.Groups("arg2") IsNot Nothing, m.Groups("arg2").Value, "")

                ' Clean arguments (normalize paragraphs, trim spaces)
                pc.Argument1 = CleanArgument(raw1)
                pc.Argument2 = CleanArgument(raw2)

                ' Add if not duplicate
                If Not results.Any(Function(x) x.Command.Equals(pc.Command, StringComparison.OrdinalIgnoreCase) _
                                        AndAlso x.Argument1 = pc.Argument1 AndAlso x.Argument2 = pc.Argument2) Then
                    results.Add(pc)
                End If
            Next
        Catch ex As Exception
            MsgBox("Error in ParseCommands: " & ex.Message, MsgBoxStyle.Critical)
        End Try
        Return results
    End Function


    ' =========================================================================
    ' Command Removal and Execution
    ' =========================================================================

    ''' <summary>
    ''' Removes bot command blocks from LLM response text.
    ''' Pattern: [#command: @@argument1@@ §§argument2§§ #]
    ''' Also collapses multiple consecutive line breaks to single newline.
    ''' </summary>
    ''' <param name="input">Text potentially containing command blocks</param>
    ''' <returns>Text with commands removed and whitespace normalized</returns>
    Public Function RemoveCommands(input As String) As String
        If input Is Nothing Then Return ""

        Dim output As String = input

        Try
            ' Keep this pattern aligned with ParseCommands:
            ' - single @ is allowed inside @@...@@
            ' - single § is allowed inside §§...§§
            ' - only @@ and §§ close their respective arguments
            Dim commandPattern As String = "\s*[\r\n]*\s*\[#(?<cmd>[^:]+):\s*@@(?<arg1>(?:[^@]|@(?!@))*?)@@\s*(?:§§(?<arg2>(?:[^§]|§(?!§))*?)§§)?\s*#\]\s*[\r\n]*\s*"
            Dim regex As New Regex(commandPattern, RegexOptions.Singleline)
            output = regex.Replace(input, "")

            ' Collapse 3+ consecutive line breaks to single newline
            Dim whitespacePattern As String = "[\r\n]{3,}"
            Dim collapseRegex As New Regex(whitespacePattern)
            output = collapseRegex.Replace(output, Environment.NewLine)

        Catch ex As System.Exception
            MsgBox("Error in RemoveCommands: " & ex.Message, MsgBoxStyle.Critical)
        End Try

        Return output
    End Function

    ' =========================================================================
    ' Command Execution State Fields
    ' =========================================================================

    ''' <summary>Accumulates descriptions of commands being executed for display to user</summary>
    Private CommandsList As String = ""

    ''' <summary>Tracks commands that failed execution for error reporting to chat</summary>
    Private FailedCommandsList As New List(Of String)()

    ''' <summary>Start position of the last document action performed by chat commands.</summary>
    Private _lastActionStart As Integer = -1

    ''' <summary>End position of the last document action performed by chat commands.</summary>
    Private _lastActionEnd As Integer = -1

    ' =========================================================================
    ' Main Command Execution Orchestrator
    ' =========================================================================

    ''' <summary>
    ''' Stores the last active-document range changed or commented by a chat command.
    ''' </summary>
    Private Sub RememberLastActionRange(startPos As Integer, endPos As Integer)
        Try
            Dim doc As Microsoft.Office.Interop.Word.Document = Globals.ThisAddIn.Application.ActiveDocument
            If doc Is Nothing Then Return

            Dim safeStart As Integer = Math.Max(doc.Content.Start, Math.Min(startPos, doc.Content.End))
            Dim safeEnd As Integer = Math.Max(doc.Content.Start, Math.Min(endPos, doc.Content.End))

            If safeEnd < safeStart Then
                safeEnd = safeStart
            End If

            _lastActionStart = safeStart
            _lastActionEnd = safeEnd
        Catch
            _lastActionStart = -1
            _lastActionEnd = -1
        End Try
    End Sub

    ''' <summary>
    ''' Selects a range in the currently active document and scrolls it into view.
    ''' </summary>
    Private Function SelectAndShowActiveDocumentRange(startPos As Integer, endPos As Integer) As Boolean
        Try
            Dim app As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
            If app Is Nothing OrElse app.Documents Is Nothing OrElse app.Documents.Count = 0 Then Return False

            Dim doc As Microsoft.Office.Interop.Word.Document = app.ActiveDocument
            If doc Is Nothing Then Return False

            Dim safeStart As Integer = Math.Max(doc.Content.Start, Math.Min(startPos, doc.Content.End))
            Dim safeEnd As Integer = Math.Max(doc.Content.Start, Math.Min(endPos, doc.Content.End))

            If safeEnd < safeStart Then
                safeEnd = safeStart
            End If

            app.Activate()
            doc.Activate()

            Dim targetRange As Microsoft.Office.Interop.Word.Range = doc.Range(safeStart, safeEnd)
            targetRange.Select()

            Try
                Dim scrollTarget As Object = targetRange
                app.ActiveWindow.ScrollIntoView(scrollTarget, True)
            Catch
                ' Selection is already sufficient if ScrollIntoView is unavailable.
            End Try

            Return True
        Catch ex As Exception
            Debug.WriteLine($"SelectAndShowActiveDocumentRange failed: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Executes a navigation command in the currently active document.
    ''' Use "last" to jump to the last chat-command action.
    ''' </summary>
    Private Function ExecuteGotoCommand(targetText As String, Optional onlySelection As Boolean = False) As Boolean
        Try
            targetText = CleanArgument(targetText)

            If String.IsNullOrWhiteSpace(targetText) OrElse
           targetText.Equals("last", StringComparison.OrdinalIgnoreCase) OrElse
           targetText.Equals("lastaction", StringComparison.OrdinalIgnoreCase) OrElse
           targetText.Equals("last action", StringComparison.OrdinalIgnoreCase) OrElse
           targetText.Equals("last point of action", StringComparison.OrdinalIgnoreCase) Then

                If _lastActionStart >= 0 AndAlso _lastActionEnd >= _lastActionStart Then
                    Return SelectAndShowActiveDocumentRange(_lastActionStart, _lastActionEnd)
                End If

                Debug.WriteLine("Goto: No last action range is available.")
                Return False
            End If

            Dim app As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
            If app Is Nothing OrElse app.Documents Is Nothing OrElse app.Documents.Count = 0 Then Return False

            Dim doc As Microsoft.Office.Interop.Word.Document = app.ActiveDocument
            If doc Is Nothing Then Return False

            targetText = DecodeParagraphMarks(targetText)

            Dim sel As Microsoft.Office.Interop.Word.Selection = doc.Application.Selection
            Dim scopeStart As Integer
            Dim scopeEnd As Integer

            If onlySelection AndAlso sel IsNot Nothing AndAlso Not String.IsNullOrEmpty(sel.Text) Then
                scopeStart = sel.Range.Start
                scopeEnd = sel.Range.End
            Else
                scopeStart = doc.Content.Start
                scopeEnd = doc.Content.End
            End If

            Try
                sel.SetRange(scopeStart, scopeEnd)
            Catch exScope As System.Exception
                Debug.WriteLine($"ExecuteGotoCommand: pre-find SetRange failed: {exScope.Message}")
            End Try

            If Globals.ThisAddIn.FindLongTextInChunks(targetText, sel, True) Then
                Dim foundStart As Integer = sel.Start
                Dim foundEnd As Integer = sel.End
                Return SelectAndShowActiveDocumentRange(foundStart, foundEnd)
            End If

            Debug.WriteLine($"Goto: Target text not found: '{targetText}'.")
            Return False

        Catch ex As Exception
            Debug.WriteLine($"ExecuteGotoCommand failed: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Executes parsed bot commands on the active Word document.
    ''' Ensures cursor is in main text story, sets revisions view to Final,
    ''' tracks success/failure for each command, removes marker characters,
    ''' and reports failures to chat. Supports ESC to abort.
    ''' </summary>
    ''' <param name="teststring">LLM response containing embedded commands</param>
    ''' <param name="OnlySelection">True to restrict operations to current selection</param>
    ''' <remarks>
    ''' Execution flow:
    ''' 1. Parse commands from teststring
    ''' 2. Ensure selection is in main document story (not header/footer/comment)
    ''' 3. Set Word view to Final (hide deletions)
    ''' 4. Iterate commands: find, replace, insert, insertbefore, insertafter, addcomment, replycomment
    ''' 5. Track success/failure for each command
    ''' 6. Remove MarkerChar cleanup markers
    ''' 7. Restore view settings
    ''' 8. Report failures to chat via ReportFailedCommands
    ''' 
    ''' ESC key polling via GetAsyncKeyState allows user to abort mid-execution.
    ''' InfoBox displays progress for operations that modify document (replace, insert*).
    ''' </remarks>
    Public Sub ExecuteAnyCommands(teststring As String, OnlySelection As Boolean)

        Dim commands = ParseCommands(teststring)
        Dim topmost As Boolean = Me.TopMost

        Me.TopMost = False

        CommandsList = ""
        FailedCommandsList.Clear()
        Dim LastCommandsList As String = ""
        Dim activateDocumentAfterCommands As Boolean = False

        Dim wordApp As Microsoft.Office.Interop.Word.Application
        Dim doc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument

        ' ═════════════════════════════════════════════════════════════════════════════
        ' ENSURE CURSOR IN MAIN STORY (NOT HEADER/FOOTER/COMMENT/FOOTNOTE)
        ' ═════════════════════════════════════════════════════════════════════════════
        Try
            wordApp = Globals.ThisAddIn.Application

            If wordApp IsNot Nothing AndAlso wordApp.ActiveDocument IsNot Nothing AndAlso wordApp.Selection IsNot Nothing Then
                Dim currentDoc As Microsoft.Office.Interop.Word.Document = wordApp.ActiveDocument
                Dim currentSel As Microsoft.Office.Interop.Word.Selection = wordApp.Selection
                Dim currentStory As Word.WdStoryType = currentSel.StoryType

                If currentStory <> Word.WdStoryType.wdMainTextStory Then
                    wordApp.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView

                    Dim mainStoryRange As Word.Range = currentDoc.StoryRanges(Word.WdStoryType.wdMainTextStory)
                    mainStoryRange.Collapse(Word.WdCollapseDirection.wdCollapseStart)
                    mainStoryRange.Select()

                    currentSel.Collapse(Word.WdCollapseDirection.wdCollapseStart)
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine($"Warning: Could not reset to main story: {ex.Message}")
        End Try

        If commands.Count() > 0 Then
            Globals.ThisAddIn.Application.Activate()
            System.Threading.Thread.Sleep(200)

            wordApp = Globals.ThisAddIn.Application
            With wordApp.ActiveWindow.View
                .RevisionsView = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal
                .ShowRevisionsAndComments = False
            End With
        End If

        For Each pc In commands
            Debug.WriteLine($"Command: '{pc.Command}' with '{pc.Argument1}' '{pc.Argument2}'")

            If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                Exit For
            End If

            Dim commandSuccess As Boolean = True
            Dim commandDescription As String = ""

            If Not OnlySelection Then
                Select Case pc.Command.ToLower()
                    Case "find", "addcomment", "replace", "insertafter", "insertbefore", "goto", "jump", "show", "select"
                        Try
                            If wordApp IsNot Nothing AndAlso wordApp.ActiveDocument IsNot Nothing AndAlso wordApp.Selection IsNot Nothing Then
                                wordApp.Selection.SetRange(wordApp.ActiveDocument.Content.Start, wordApp.ActiveDocument.Content.End)
                            End If
                        Catch exScope As System.Exception
                            Debug.WriteLine($"ExecuteAnyCommands: pre-command SetRange failed: {exScope.Message}")
                        End Try
                End Select
            End If

            Select Case pc.Command.ToLower()
                Case "find"
                    commandDescription = $"Finding '{pc.Argument1}'"
                    CommandsList = commandDescription & Environment.NewLine & CommandsList
                    LastCommandsList = CommandsList
                    System.Threading.Thread.Sleep(500)
                    commandSuccess = ExecuteFindCommand(pc.Argument1, OnlySelection)

                Case "addcomment"
                    commandDescription = $"Adding comment '{pc.Argument2}' to the text '{pc.Argument1}'"
                    CommandsList = commandDescription & Environment.NewLine & CommandsList
                    LastCommandsList = CommandsList
                    System.Threading.Thread.Sleep(500)
                    commandSuccess = ExecuteAddComment(pc.Argument1, pc.Argument2, OnlySelection)

                Case "replycomment"
                    commandDescription = $"Replying to comment '{pc.Argument1}' with '{pc.Argument2}'"
                    CommandsList = commandDescription & Environment.NewLine & CommandsList
                    LastCommandsList = CommandsList
                    System.Threading.Thread.Sleep(500)
                    commandSuccess = ExecuteReplyToCommentByIdToken(pc.Argument1, pc.Argument2)

                Case "replace"
                    If String.IsNullOrEmpty(pc.Argument2) Then
                        commandDescription = $"Deleting '{pc.Argument1}'"
                    Else
                        commandDescription = $"Replacing '{pc.Argument1}' with '{pc.Argument2}'"
                    End If
                    CommandsList = commandDescription & Environment.NewLine & CommandsList
                    LastCommandsList = CommandsList
                    InfoBox.ShowInfoBox("Executing bot commands ('Esc' to abort):" & Environment.NewLine & Environment.NewLine & CommandsList)
                    System.Threading.Thread.Sleep(500)
                    commandSuccess = ExecuteReplaceCommand(pc.Argument1, pc.Argument2, OnlySelection, MarkerChar)

                Case "insertafter"
                    commandDescription = $"Inserting '{pc.Argument2}' after '{pc.Argument1}'"
                    CommandsList = commandDescription & Environment.NewLine & CommandsList
                    LastCommandsList = CommandsList
                    InfoBox.ShowInfoBox("Executing bot commands ('Esc' to abort):" & Environment.NewLine & Environment.NewLine & CommandsList)
                    System.Threading.Thread.Sleep(500)
                    commandSuccess = ExecuteInsertBeforeAfterCommand(pc.Argument1, pc.Argument2, OnlySelection, False)

                Case "insertbefore"
                    commandDescription = $"Inserting '{pc.Argument2}' before '{pc.Argument1}'"
                    CommandsList = commandDescription & Environment.NewLine & CommandsList
                    LastCommandsList = CommandsList
                    InfoBox.ShowInfoBox("Executing bot commands ('Esc' to abort):" & Environment.NewLine & Environment.NewLine & CommandsList)
                    System.Threading.Thread.Sleep(500)
                    commandSuccess = ExecuteInsertBeforeAfterCommand(pc.Argument1, pc.Argument2, OnlySelection, True)

                Case "insert"
                    commandDescription = $"Inserting '{pc.Argument1}'"
                    CommandsList = commandDescription & Environment.NewLine & CommandsList
                    LastCommandsList = CommandsList
                    InfoBox.ShowInfoBox("Executing bot commands ('Esc' to abort):" & Environment.NewLine & Environment.NewLine & CommandsList)
                    System.Threading.Thread.Sleep(500)
                    Debug.WriteLine("ExecuteInsert")
                    commandSuccess = ExecuteInsertCommand(pc.Argument1)

                Case "goto", "jump", "show", "select"
                    commandDescription = $"Showing '{pc.Argument1}'"
                    CommandsList = commandDescription & Environment.NewLine & CommandsList
                    LastCommandsList = CommandsList
                    System.Threading.Thread.Sleep(250)
                    commandSuccess = ExecuteGotoCommand(pc.Argument1, OnlySelection)
                    activateDocumentAfterCommands = activateDocumentAfterCommands OrElse commandSuccess

                Case Else
                    commandDescription = $"Unknown command: '{pc.Command}'"
                    commandSuccess = False
            End Select

            If Not commandSuccess AndAlso Not String.IsNullOrWhiteSpace(commandDescription) Then
                FailedCommandsList.Add($"Failed: {commandDescription}")
            End If

            If LastCommandsList <> CommandsList Then
                System.Threading.Thread.Sleep(500)
            End If
        Next

        If commands.Count() > 0 Then
            ReplaceSpecialCharacter(OnlySelection)

            InfoBox.ShowInfoBox("")

            With wordApp.ActiveWindow.View
                .RevisionsView = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal
                .ShowRevisionsAndComments = True
            End With
        End If

        If wordApp IsNot Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp)
            wordApp = Nothing
        End If

        Me.TopMost = topmost

        Dim documentActivated As Boolean = False

        If activateDocumentAfterCommands Then
            Try
                Dim activeApp As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
                If activeApp IsNot Nothing Then
                    activeApp.Activate()

                    If activeApp.ActiveDocument IsNot Nothing Then
                        activeApp.ActiveDocument.Activate()
                    End If

                    documentActivated = True
                End If
            Catch
                documentActivated = False
            End Try
        End If

        _keepFocusOnDocumentAfterCommands = documentActivated

        If Not documentActivated Then
            Me.Focus()
        End If

        If FailedCommandsList.Count > 0 Then
            ReportFailedCommands()
        End If

    End Sub

    ' =========================================================================
    ' TOC Detection Helpers
    ' =========================================================================

    ''' <summary>
    ''' Determines whether a range overlaps a table of contents and returns the TOC end position if so.
    ''' </summary>
    ''' <param name="foundRange">The candidate range that was found by a search.</param>
    ''' <param name="doc">The document containing the table(s) of contents.</param>
    ''' <returns>
    ''' The end position of the overlapping TOC range, or 0 when the specified range does not overlap a TOC.
    ''' </returns>
    ''' <remarks>
    ''' Command execution skips TOCs to avoid corrupting generated fields. Any overlap is treated as "inside" for safety.
    ''' </remarks>
    Private Function TocEndIfInside(foundRange As Word.Range, doc As Word.Document) As Integer
        If foundRange Is Nothing OrElse doc Is Nothing Then Return 0

        For Each toc As Word.TableOfContents In doc.TablesOfContents
            Dim tr As Word.Range = toc.Range
            ' Treat any overlap with TOC as "inside" for skipping
            If foundRange.Start < tr.End AndAlso foundRange.End > tr.Start Then
                Return tr.End
            End If
        Next

        Return 0
    End Function

    ''' <summary>
    ''' Indicates whether a range overlaps a table of contents.
    ''' </summary>
    ''' <param name="range">The range to test.</param>
    ''' <param name="doc">The document containing the table(s) of contents.</param>
    ''' <returns><see langword="True"/> when the range overlaps a TOC; otherwise <see langword="False"/>.</returns>

    Private Function IsInsideToc(range As Word.Range, doc As Word.Document) As Boolean
        Return TocEndIfInside(range, doc) > 0
    End Function

    ' =========================================================================
    ' Command Failure Reporting
    ' =========================================================================

    ''' <summary>
    ''' Reports failed commands to chat in both plain text and HTML formats.
    ''' Adds failures to _chatHistory so LLM sees them in subsequent messages.
    ''' </summary>
    ''' <remarks>
    ''' Error rendered with red styling (#d93025) to distinguish from warnings (orange).
    ''' Failures formatted as bulleted list in HTML view.
    ''' </remarks>
    Private Sub ReportFailedCommands()
        If FailedCommandsList Is Nothing OrElse FailedCommandsList.Count = 0 Then Return

        Dim errorMessage As New System.Text.StringBuilder()
        errorMessage.AppendLine()
        errorMessage.AppendLine("─────────────────────────────────────")
        errorMessage.AppendLine("⚠ Some commands could not be executed:")
        errorMessage.AppendLine()

        For Each failedCmd In FailedCommandsList
            errorMessage.AppendLine($"  • {failedCmd}")
        Next

        errorMessage.AppendLine()
        errorMessage.AppendLine("─────────────────────────────────────")

        ' Add to plain text chat history
        AppendToChatHistory(errorMessage.ToString())

        ' Add to HTML chat with red error styling
        Dim htmlError As String = $"<div class='msg assistant error' style='border-left: 3px solid #d93025; padding-left: 10px; margin: 10px 0; background-color: #fef1f0;'>
            <span class='who' style='color: #d93025;'>System:</span>
            <div class='content'>
                <hr style='border: none; border-top: 1px solid #d93025; margin: 8px 0;' />
                <strong>⚠ Some commands could not be executed:</strong><br/>
                <ul style='margin: 8px 0;'>"

        For Each failedCmd In FailedCommandsList
            htmlError += $"<li>{HtmlEncode(failedCmd)}</li>"
        Next

        htmlError += "</ul><hr style='border: none; border-top: 1px solid #d93025; margin: 8px 0;' /></div></div>"

        AppendHtml(htmlError)
        PersistChatHtml()

        ' Add to chat history so AI can see failures in future context
        _chatHistory.Add(("assistant", $"System: Some commands failed - {String.Join("; ", FailedCommandsList)}"))
    End Sub


    ' =========================================================================
    ' Marker Character Cleanup
    ' =========================================================================

    ''' <summary>
    ''' Removes all MarkerChar (U+E000) instances from document or selection.
    ''' MarkerChar inserted during replace operations to prevent infinite loops.
    ''' </summary>
    ''' <param name="OnlySelection">True to clean only selection, False for entire document</param>
    ''' <remarks>
    ''' Uses Word Find/Replace with tracked changes enabled.
    ''' Original TrackRevisions state restored in Finally block.
    ''' </remarks>
    Private Sub ReplaceSpecialCharacter(Optional OnlySelection As Boolean = False)

        Dim doc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument
        Dim trackChangesEnabled = doc.TrackRevisions

        Try
            doc.TrackRevisions = True

            ' Determine search range
            Dim rng As Word.Range =
                If(OnlySelection AndAlso Not String.IsNullOrEmpty(doc.Application.Selection.Text),
                   doc.Application.Selection.Range.Duplicate,
                   doc.Content.Duplicate)

            ' Find and replace all MarkerChar instances
            Using ThisAddIn.BeginMarkupAuthorScope(doc.Application)
                With rng.Find
                    .ClearFormatting()
                    .Text = MarkerChar
                    .Replacement.ClearFormatting()
                    .Replacement.Text = ""
                    .Forward = True
                    .Wrap = Word.WdFindWrap.wdFindStop
                    Do While .Execute(Replace:=Word.WdReplace.wdReplaceOne)
                    Loop
                End With
            End Using
        Catch ex As Exception
            MsgBox("Error in ReplaceSpecialCharacter: " & ex.Message, MsgBoxStyle.Critical)
        Finally
            doc.TrackRevisions = trackChangesEnabled
        End Try
    End Sub

    ' =========================================================================
    ' Comment Reply Command
    ' =========================================================================

    ''' <summary>
    ''' Adds threaded reply to existing Word comment using LLM-friendly token formats.
    ''' Accepts formats: "id|hash", "id=123;hash=abc", "wid:123 ph:abc", "123", "abcdef".
    ''' </summary>
    ''' <param name="idToken">Combined identifier token for target comment</param>
    ''' <param name="replyText">Reply text to add (prefixed with AN6 constant)</param>
    ''' <returns>True if reply added successfully</returns>
    ''' <remarks>
    ''' Restores selection to main story after operation to avoid leaving caret in comment.
    ''' Uses TryParseCommentIdToken to extract Word comment Index and/or PseudoHash.
    ''' Calls ThisAddIn.ReplyToWordComment with formatted flag from chkConvertMarkdown.
    ''' </remarks>
    Private Function ExecuteReplyToCommentByIdToken(ByVal idToken As String, ByVal replyText As String) As Boolean

        Dim app As Microsoft.Office.Interop.Word.Application = Nothing
        Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
        Dim hadSel As Boolean = False
        Dim origStart As Integer = -1
        Dim origEnd As Integer = -1

        Try
            app = Globals.ThisAddIn.Application
            If app IsNot Nothing AndAlso app.Documents IsNot Nothing AndAlso app.Documents.Count > 0 Then
                doc = app.ActiveDocument
                If app.Selection IsNot Nothing Then
                    origStart = app.Selection.Start
                    origEnd = app.Selection.End
                    hadSel = True
                End If
            End If

            ' Validate inputs
            If String.IsNullOrWhiteSpace(idToken) Then
                Debug.WriteLine("Add-Reply: Missing ID token.")
                Return False
            End If
            If String.IsNullOrWhiteSpace(replyText) Then
                Debug.WriteLine("Add-Reply: Reply text is empty.")
                Return False
            End If

            ' Parse comment identifier token
            Dim wordId As System.Nullable(Of Integer) = Nothing
            Dim pseudoHash As String = Nothing

            If Not TryParseCommentIdToken(idToken, wordId, pseudoHash) Then
                Debug.WriteLine("Add-Reply: Could not parse ID token (expected formats like '1234|abcdef' or 'id=1234;hash=abcdef').")
                Return False
            End If

            Debug.WriteLine($"Add-Reply: Parsed token '{idToken}' -> WordId={If(wordId.HasValue, wordId.Value.ToString(), "null")}, Hash={If(pseudoHash, "null")}")

            ' Execute reply with Markdown formatting if enabled
            Dim formatted As Boolean = chkConvertMarkdown.Checked
            Dim ok As Boolean = ThisAddIn.ReplyToWordComment(wordId, pseudoHash, AN6 & ": " & replyText, formatted)

            If ok Then
                Debug.WriteLine($"Add-Reply: Successfully added reply to comment {If(wordId.HasValue, wordId.Value.ToString(), pseudoHash)}")
            Else
                Debug.WriteLine($"Add-Reply: Failed to add reply to comment {If(wordId.HasValue, wordId.Value.ToString(), pseudoHash)} (target not found).")
            End If

            Return ok

        Catch ex As Exception
            Debug.WriteLine($"Add-Reply Error: {ex.Message}")
            Return False
        Finally
            ' Restore selection to main text story to avoid leaving caret in comment
            Try
                If app IsNot Nothing AndAlso doc IsNot Nothing AndAlso hadSel Then
                    app.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
                    Dim s As Integer = Math.Max(doc.Content.Start, Math.Min(origStart, doc.Content.End))
                    Dim e As Integer = Math.Max(doc.Content.Start, Math.Min(origEnd, doc.Content.End))
                    doc.Range(s, e).Select()
                End If
            Catch
                ' Best-effort restore; ignore failures
            End Try
        End Try
    End Function

    ' =========================================================================
    ' Comment ID Token Parsing
    ' =========================================================================

    ''' <summary>
    ''' Parses combined comment ID token into Word comment Index and/or PseudoHash.
    ''' Supports formats: "id|hash", "id=123;hash=abc", "wid:123 ph:abc", "123", "abcdef".
    ''' </summary>
    ''' <param name="raw">Token string to parse</param>
    ''' <param name="wordId">Output: Word comment index if found</param>
    ''' <param name="pseudoHash">Output: Pseudo-hash identifier if found</param>
    ''' <returns>True if at least one identifier extracted</returns>
    ''' <remarks>
    ''' Parsing priority:
    ''' 1. Pipe-separated: "123|abcdef"
    ''' 2. Labeled: "id=123;hash=abc" or "wid:123 ph:abc"
    ''' 3. Plain number: "123" → treated as wordId
    ''' 4. Plain text: "abcdef" (6+ chars) → treated as pseudoHash
    ''' </remarks>
    Private Function TryParseCommentIdToken(ByVal raw As String, ByRef wordId As System.Nullable(Of Integer), ByRef pseudoHash As String) As Boolean
        wordId = Nothing
        pseudoHash = Nothing
        If String.IsNullOrWhiteSpace(raw) Then Return False

        Dim s As String = raw.Trim()
        Debug.WriteLine($"TryParseCommentIdToken: Parsing '{s}'")

        ' ═════════════════════════════════════════════════════════════════════════════
        ' 1. PIPE-SEPARATED FORMAT: "id|hash"
        ' ═════════════════════════════════════════════════════════════════════════════
        Dim pipeParts = s.Split(New Char() {"|"c}, 2, StringSplitOptions.None)
        If pipeParts.Length = 2 Then
            Dim left = pipeParts(0).Trim()
            Dim right = pipeParts(1).Trim()
            Dim idVal As Integer
            If Integer.TryParse(left, idVal) Then wordId = idVal
            If Not String.IsNullOrWhiteSpace(right) Then pseudoHash = right
            Debug.WriteLine($"TryParseCommentIdToken: Pipe format -> WordId={If(wordId.HasValue, wordId.Value.ToString(), "null")}, Hash={If(pseudoHash, "null")}")
            Return (wordId.HasValue OrElse Not String.IsNullOrWhiteSpace(pseudoHash))
        End If

        ' ═════════════════════════════════════════════════════════════════════════════
        ' 2. LABELED FORMAT: "id=123;hash=abc" or "wid:123 ph:abc"
        ' ═════════════════════════════════════════════════════════════════════════════
        Dim idMatch = System.Text.RegularExpressions.Regex.Match(s, "(?:\bwid|\bid|\bwordid)\s*[:=]\s*(?<id>-?\d+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        If idMatch.Success Then
            Dim idVal As Integer
            If Integer.TryParse(idMatch.Groups("id").Value, idVal) Then
                wordId = idVal
                Debug.WriteLine($"TryParseCommentIdToken: Found WordId={wordId.Value} from labeled format")
            End If
        End If

        Dim hashMatch = System.Text.RegularExpressions.Regex.Match(s, "(?:\bph|\bhash|\bpseudohash)\s*[:=]\s*(?<hash>[A-Za-z0-9_-]{6,})", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        If hashMatch.Success Then
            pseudoHash = hashMatch.Groups("hash").Value.Trim()
            Debug.WriteLine($"TryParseCommentIdToken: Found Hash={pseudoHash} from labeled format")
        End If

        If wordId.HasValue OrElse Not String.IsNullOrWhiteSpace(pseudoHash) Then
            Debug.WriteLine($"TryParseCommentIdToken: Labeled format -> WordId={If(wordId.HasValue, wordId.Value.ToString(), "null")}, Hash={If(pseudoHash, "null")}")
            Return True
        End If

        ' ═════════════════════════════════════════════════════════════════════════════
        ' 3. PLAIN TOKEN FALLBACK: all digits → id, otherwise → hash
        ' ═════════════════════════════════════════════════════════════════════════════
        Dim onlyDigits As Boolean = s.All(Function(ch) Char.IsDigit(ch))
        If onlyDigits Then
            Dim idVal As Integer
            If Integer.TryParse(s, idVal) Then
                wordId = idVal
                Debug.WriteLine($"TryParseCommentIdToken: Plain number -> WordId={wordId.Value}")
                Return True
            End If
        Else
            ' Accept as hash if 6+ characters
            If s.Length >= 6 Then
                pseudoHash = s
                Debug.WriteLine($"TryParseCommentIdToken: Plain text -> Hash={pseudoHash}")
                Return True
            End If
        End If

        Debug.WriteLine("TryParseCommentIdToken: Failed to parse")
        Return False
    End Function

    ' =========================================================================
    ' Add Comment Command
    ' =========================================================================

    ''' <summary>
    ''' Adds a Word comment to the first occurrence of search term in document or selection.
    ''' Uses FindLongTextInChunks for reliable matching in large documents.
    ''' </summary>
    ''' <param name="searchTerm">Text to search for as comment anchor</param>
    ''' <param name="commentText">Comment body text (prefixed with AN6)</param>
    ''' <param name="onlySelection">True to restrict to current selection</param>
    ''' <returns>True if one comment was added</returns>
    ''' <remarks>
    ''' Chat-generated addcomment commands are expected to target the first matching anchor.
    ''' This avoids adding duplicate comments to repeated short phrases and prevents loops
    ''' caused by Word moving the active selection into the comment after Comments.Add.
    ''' </remarks>
    Private Function ExecuteAddComment(
        ByVal searchTerm As String,
        ByVal commentText As String,
        Optional ByVal onlySelection As Boolean = False) As Boolean

        Dim app As Microsoft.Office.Interop.Word.Application = Nothing
        Dim doc As Microsoft.Office.Interop.Word.Document = Nothing

        If String.IsNullOrWhiteSpace(searchTerm) Then
            Debug.WriteLine("AddComments: Search term is empty.")
            Return False
        End If

        If String.IsNullOrWhiteSpace(commentText) Then
            Debug.WriteLine("AddComments: Comment text is empty.")
            Return False
        End If

        searchTerm = DecodeParagraphMarks(searchTerm)

        Try
            Try
                app = CType(System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application"), Microsoft.Office.Interop.Word.Application)
            Catch
                app = Globals.ThisAddIn.Application
            End Try
        Catch ex As System.Exception
            Debug.WriteLine("AddComments: Unable to access Word Application instance.")
            Return False
        End Try

        Try
            doc = app.ActiveDocument
        Catch
            Debug.WriteLine("AddComments: No active document found.")
            Return False
        End Try

        If doc Is Nothing Then
            Debug.WriteLine("AddComments: No active document found.")
            Return False
        End If

        Dim sel As Microsoft.Office.Interop.Word.Selection = doc.Application.Selection
        Dim originalSelStart As Integer = sel.Start
        Dim originalSelEnd As Integer = sel.End

        Dim scopeStart As Integer
        Dim scopeEnd As Integer
        If onlySelection AndAlso sel IsNot Nothing AndAlso Not String.IsNullOrEmpty(sel.Text) Then
            scopeStart = sel.Range.Start
            scopeEnd = sel.Range.End
        Else
            scopeStart = doc.Content.Start
            scopeEnd = doc.Content.End
        End If

        Try
            ' Search only inside the intended working range.
            Try
                sel.SetRange(scopeStart, scopeEnd)
            Catch exScope As System.Exception
                Debug.WriteLine($"ExecuteAddComment: pre-find SetRange failed: {exScope.Message}")
            End Try

            If Not Globals.ThisAddIn.FindLongTextInChunks(searchTerm, sel) Then
                Debug.WriteLine($"AddComments: Search term not found: '{searchTerm}'.")
                Return False
            End If

            If sel Is Nothing OrElse sel.Start >= sel.End Then
                Debug.WriteLine($"AddComments: Invalid found range for term '{searchTerm}'.")
                Return False
            End If

            Dim anchor As Microsoft.Office.Interop.Word.Range = sel.Range.Duplicate
            Dim anchorStart As Integer = anchor.Start
            Dim anchorEnd As Integer = anchor.End

            If anchorStart < scopeStart OrElse anchorEnd > scopeEnd Then
                Debug.WriteLine($"AddComments: Found range outside working range for term '{searchTerm}'.")
                Return False
            End If

            Using ThisAddIn.BeginMarkupAuthorScope(app)
                Dim newComment As Microsoft.Office.Interop.Word.Comment = doc.Comments.Add(anchor, String.Empty)

                If chkConvertMarkdown.Checked Then
                    ThisAddIn.InsertMarkdownToComment(newComment.Range, AN6 & ": " & commentText)
                Else
                    newComment.Range.Text = AN6 & ": " & commentText
                End If
            End Using

            Debug.WriteLine($"AddComments: Added one comment for term '{searchTerm}' at [{anchorStart},{anchorEnd}].")

            RememberLastActionRange(anchorStart, anchorEnd)

            Return True

        Catch ex As System.Exception
            Debug.WriteLine($"AddComments failed: {ex.Message}")
            Return False

        Finally
            Try
                Dim s As Integer = Math.Max(doc.Content.Start, Math.Min(originalSelStart, doc.Content.End))
                Dim e As Integer = Math.Max(doc.Content.Start, Math.Min(originalSelEnd, doc.Content.End))
                doc.Range(s, e).Select()
            Catch
            End Try
        End Try
    End Function


    ' =========================================================================
    ' Find Command
    ' =========================================================================

    ''' <summary>
    ''' Finds and highlights all occurrences of search term with yellow highlighting.
    ''' Supports ESC key abort and handles table cell boundaries.
    ''' </summary>
    ''' <param name="searchTerm">Text to find (normalized via DecodeParagraphMarks)</param>
    ''' <param name="OnlySelection">True to restrict search to current selection</param>
    ''' <returns>True if at least one match found</returns>
    ''' <remarks>
    ''' Uses FindLongTextInChunks for reliability with large text.
    ''' Tracks position to detect stuck state (exits after 2 consecutive stuck positions).
    ''' Handles table navigation to avoid infinite loops at cell boundaries.
    ''' Restores original selection and TrackRevisions state in Finally block.
    ''' </remarks>
    Private Function ExecuteFindCommand(searchTerm As String, Optional OnlySelection As Boolean = False) As Boolean
        Dim doc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument
        Dim trackChangesEnabled As Boolean = doc.TrackRevisions
        Dim originalAuthor As String = doc.Application.UserName
        Dim selectionStart As Integer = doc.Application.Selection.Start
        Dim selectionEnd As Integer = doc.Application.Selection.End
        Dim found As Boolean = False

        Try
            doc.Application.Activate()
            doc.Activate()

            doc.TrackRevisions = True

            ' Normalize paragraph marks
            searchTerm = DecodeParagraphMarks(searchTerm)
            If String.IsNullOrWhiteSpace(searchTerm) Then
                CommandsList = $"Note: Empty search term (ignored)." & Environment.NewLine & CommandsList
                Return False
            End If

            ' Define starting selection
            If OnlySelection Then
                If doc.Application.Selection Is Nothing OrElse doc.Application.Selection.Range.Text = "" Then
                    OnlySelection = False
                    doc.Application.Selection.SetRange(doc.Content.Start, doc.Content.End)
                End If
            Else
                doc.Application.Selection.SetRange(doc.Content.Start, doc.Content.End)
            End If

            Dim lastSelectionStart As Integer = -1
            Dim stuckCounter As Integer = 0
            Dim maxStuckLimit As Integer = 2

            ' Find and highlight all instances
            Do While Globals.ThisAddIn.FindLongTextInChunks(searchTerm, doc.Application.Selection, True) = True

                If doc.Application.Selection Is Nothing Then Exit Do

                System.Windows.Forms.Application.DoEvents()
                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then
                    CommandsList = $"Operation cancelled by user (ESC)." & Environment.NewLine & CommandsList
                    Exit Do
                End If

                found = True

                ' Highlight found text with yellow
                doc.Application.Selection.Range.HighlightColorIndex = Word.WdColorIndex.wdYellow

                ' Detect stuck state (same position multiple times)
                If doc.Application.Selection.Start = lastSelectionStart Then
                    stuckCounter += 1
                    If stuckCounter >= maxStuckLimit Then
                        Exit Do
                    End If
                Else
                    stuckCounter = 0
                End If
                lastSelectionStart = doc.Application.Selection.Start

                ' Collapse to end of match
                doc.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                ' Handle table cell boundaries to avoid infinite loops
                If doc.Application.Selection.Range.Tables.Count > 0 Then
                    Try
                        Dim currentCell As Word.Cell = doc.Application.Selection.Cells(1)
                        If doc.Application.Selection.End >= currentCell.Range.End - 1 Then
                            doc.Application.Selection.MoveRight(Unit:=Word.WdUnits.wdCell, Count:=1, Extend:=Word.WdMovementType.wdMove)
                        End If
                    Catch ex As System.Exception
                        ' Not in valid cell; ignore and continue
                    End Try
                End If

                ' Ensure not stuck in empty cell
                If doc.Application.Selection.Range.Text = vbCr Or doc.Application.Selection.Range.Text = "" Then
                    doc.Application.Selection.Move(Unit:=Word.WdUnits.wdCharacter, Count:=1)
                End If

                ' Check if reached end of search range
                If OnlySelection Then
                    If doc.Application.Selection.Start >= selectionEnd Then Exit Do
                    doc.Application.Selection.SetRange(doc.Application.Selection.Start, selectionEnd)
                Else
                    If doc.Application.Selection.Start >= doc.Content.End Then Exit Do
                    doc.Application.Selection.SetRange(doc.Application.Selection.Start, doc.Content.End)
                End If
            Loop

            If Not found Then
                CommandsList = $"Note: The search term was not found." & Environment.NewLine & CommandsList
            End If

            Return found

        Catch ex As System.Exception
            MsgBox("Error in ExecuteFindCommand: " & ex.Message)
            Return False

        Finally
            ' Restore original state
            doc.TrackRevisions = trackChangesEnabled
            doc.Application.Selection.SetRange(selectionStart, selectionEnd)
            doc.Application.Selection.Select()
        End Try
    End Function

    ' =========================================================================
    ' Replace Command
    ' =========================================================================

    ''' <summary>
    ''' Finds and replaces all occurrences of oldText with newText using tracked changes.
    ''' Uses two-pass strategy: forward scan to collect match positions, then reverse-order replacement.
    ''' </summary>
    ''' <param name="oldText">Text to find (normalized via DecodeParagraphMarks)</param>
    ''' <param name="newText">Replacement text (empty for delete)</param>
    ''' <param name="OnlySelection">True to restrict to current selection</param>
    ''' <param name="Marker">MarkerChar (U+E000) — unused in two-pass approach but kept for API compat</param>
    ''' <returns>True if at least one replacement made</returns>
    ''' <remarks>
    ''' Two-pass strategy rationale:
    ''' 
    ''' Single-pass approaches fail because:
    ''' - With TrackRevisions=True, Range.Delete() does NOT remove characters from the
    '''   position space — it only marks them as tracked deletions. Position arithmetic
    '''   that assumes deletion shifts positions is therefore completely wrong.
    ''' - MarkerChar (U+E000) is not filtered by FindLongTextInChunks' canonical search
    '''   (Strategy 4), so it cannot prevent re-matching.
    ''' - FindLongTextInChunks does not support backward searching.
    ''' 
    ''' The two-pass approach avoids all of these issues:
    ''' - Pass 1 collects match positions as integer pairs (Start, End) — not Range objects,
    '''   which avoids COM stale-reference problems in tables.
    ''' - Pass 2 replaces in reverse document order (last match first). Since each
    '''   Selection.Text assignment only affects positions AT or AFTER the replacement
    '''   site, earlier match positions (lower indices) remain valid.
    ''' - Selection.Text = newText creates an atomic tracked change (the old text is
    '''   marked as deleted and the new text as inserted in a single operation).
    '''   This is how Word's own Find/Replace works internally.
    ''' 
    ''' Table handling:
    ''' - Using integer positions instead of Range.Duplicate avoids the stale COM reference
    '''   problem that corrupted tables in the original implementation.
    ''' - Selection.Text assignment handles table cell boundaries correctly because Word
    '''   manages the cell markers internally for Selection operations.
    ''' - The cell-boundary advancement in Pass 1 prevents the scanner from getting stuck
    '''   at end-of-cell markers.
    ''' </remarks>
    Private Function ExecuteReplaceCommand(oldText As String, newText As String, OnlySelection As Boolean, Marker As String) As Boolean
        Dim doc As Word.Document = Nothing
        Dim view As Word.View = Nothing
        Dim trackChangesEnabled As Boolean = False
        Dim originalRevisionsView As Word.WdRevisionsView = Word.WdRevisionsView.wdRevisionsViewFinal
        Dim originalShowRevisions As Boolean = False

        Try
            Debug.WriteLine("ExecuteReplaceCommand: START")
            LogReplaceDiag(New String("-"c, 100))
            LogReplaceDiag("START")

            Try
                doc = Globals.ThisAddIn.Application.ActiveDocument
            Catch ex As Exception
                Debug.WriteLine($"ExecuteReplaceCommand: FAILED to get ActiveDocument: {ex.Message}")
                Return False
            End Try

            trackChangesEnabled = doc.TrackRevisions

            Try
                view = doc.Application.ActiveWindow.View
                originalRevisionsView = view.RevisionsView
                originalShowRevisions = view.ShowRevisionsAndComments
            Catch ex As Exception
                Debug.WriteLine($"ExecuteReplaceCommand: FAILED to get view settings: {ex.Message}")
                Return False
            End Try

            ' Normalize inputs
            oldText = DecodeParagraphMarks(oldText)
            newText = DecodeParagraphMarks(newText)
            oldText = If(oldText, String.Empty)
            newText = If(newText, String.Empty)

            Debug.WriteLine($"ExecuteReplaceCommand: oldText='{oldText}' ({oldText.Length} chars), newText='{newText}' ({newText.Length} chars)")
            LogReplaceDiag($"Inputs: OnlySelection={OnlySelection}; oldTextLen={oldText.Length}; newTextLen={newText.Length}; oldText='{PreviewForLog(oldText)}'; newText='{PreviewForLog(newText)}'")

            If String.IsNullOrWhiteSpace(oldText) Then
                CommandsList = $"Note: Empty search term (ignored)." & Environment.NewLine & CommandsList
                Return False
            End If

            doc.Application.Activate()
            doc.Activate()

            doc.TrackRevisions = True

            ' Show markup during replacement for visibility
            view.RevisionsView = Word.WdRevisionsView.wdRevisionsViewFinal
            view.ShowRevisionsAndComments = True

            Dim savedSelectionStart As Integer = doc.Application.Selection.Start
            Dim savedSelectionEnd As Integer = doc.Application.Selection.End
            Debug.WriteLine($"ExecuteReplaceCommand: savedSelection=[{savedSelectionStart},{savedSelectionEnd}]")
            LogReplaceDiag($"Saved selection=[{savedSelectionStart},{savedSelectionEnd}] {DescribeSelectionState(doc.Application.Selection)}")

            ' Define search boundaries
            Dim searchEnd As Integer
            If OnlySelection AndAlso Not String.IsNullOrWhiteSpace(doc.Application.Selection.Text) Then
                searchEnd = doc.Application.Selection.End
                Debug.WriteLine($"ExecuteReplaceCommand: searching within selection, searchEnd={searchEnd}")
            Else
                OnlySelection = False
                doc.Application.Selection.SetRange(doc.Content.Start, doc.Content.End)
                searchEnd = doc.Content.End
                Debug.WriteLine($"ExecuteReplaceCommand: searching whole document, searchEnd={searchEnd}")
            End If
            ' ─────────────────────────────────────────────────────────────────────
            ' PASS 1: COLLECT ALL MATCH POSITIONS (forward scan)
            ' ─────────────────────────────────────────────────────────────────────
            Dim matchPositions As New List(Of (Start As Integer, [End] As Integer))
            Dim maxIterations As Integer = 5000
            Dim iterationCount As Integer = 0
            Dim lastFoundEnd As Integer = -1
            Dim requestedSearchStart As Integer = doc.Application.Selection.Start

            Debug.WriteLine("ExecuteReplaceCommand: PASS 1 - scanning for matches...")

            Do
                LogReplaceDiag($"PASS1 before Find: nextIteration={iterationCount + 1}; searchEnd={searchEnd}; requestedSearchStart={requestedSearchStart}; lastFoundEnd={lastFoundEnd}; {DescribeSelectionState(doc.Application.Selection)}")

                Dim findReturned As Boolean = False
                Try
                    findReturned = Globals.ThisAddIn.FindLongTextInChunks(oldText, doc.Application.Selection, True)
                Catch ex As Exception
                    LogReplaceDiag($"PASS1 FindLongTextInChunks THREW: {ex.GetType().Name}: {ex.Message}")
                    Throw
                End Try

                LogReplaceDiag($"PASS1 after Find: returned={findReturned}; {DescribeSelectionState(doc.Application.Selection)}")

                If Not findReturned Then
                    Exit Do
                End If

                If doc.Application.Selection Is Nothing Then
                    Debug.WriteLine("ExecuteReplaceCommand: Selection is Nothing after Find, exiting loop")
                    LogReplaceDiag("PASS1 Selection is Nothing after Find")
                    Exit Do
                End If

                ' Check for user abort
                System.Windows.Forms.Application.DoEvents()
                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then
                    CommandsList = $"Operation cancelled by user (ESC)." & Environment.NewLine & CommandsList
                    Debug.WriteLine("ExecuteReplaceCommand: ESC pressed, aborting")
                    LogReplaceDiag("PASS1 ESC pressed, aborting")
                    Return False
                End If

                iterationCount += 1
                If iterationCount > maxIterations Then
                    CommandsList = $"Warning: Max search iterations ({maxIterations}) reached." & Environment.NewLine & CommandsList
                    Debug.WriteLine("ExecuteReplaceCommand: Max iterations reached")
                    LogReplaceDiag($"PASS1 maxIterations reached ({maxIterations})")
                    Exit Do
                End If

                Dim selStart As Integer = doc.Application.Selection.Start
                Dim selEnd As Integer = doc.Application.Selection.End
                Debug.WriteLine($"ExecuteReplaceCommand: PASS1 iteration {iterationCount}, found at [{selStart},{selEnd}] (requestedSearchStart={requestedSearchStart})")
                LogReplaceDiag($"PASS1 iteration={iterationCount}; found=[{selStart},{selEnd}]; requestedSearchStart={requestedSearchStart}; lastFoundEnd={lastFoundEnd}")

                ' Validate match is within search bounds
                If selStart >= searchEnd Then
                    Debug.WriteLine($"ExecuteReplaceCommand: match start {selStart} >= searchEnd {searchEnd}, exiting")
                    LogReplaceDiag($"PASS1 exiting because selStart {selStart} >= searchEnd {searchEnd}")
                    Exit Do
                End If

                ' Reject matches that cross a table-cell boundary.
                ' The canonical fallback in FindLongTextAnchoredFast (Strategy 4)
                ' joins adjacent cell text into a single canonical stream and can
                ' return a hit that spans cells. Word's Selection then auto-
                ' expands to include both cells, which corrupts the replacement.
                Try
                    Dim probeRange As Word.Range = doc.Range(selStart, Math.Min(selStart + 1, doc.Content.End))
                    Dim startInTable As Boolean = False
                    Try
                        startInTable = CBool(probeRange.Information(Word.WdInformation.wdWithInTable))
                    Catch
                        startInTable = False
                    End Try

                    If startInTable Then
                        Dim startCellEnd As Integer = -1
                        Try
                            If probeRange.Cells.Count > 0 Then
                                startCellEnd = probeRange.Cells(1).Range.End
                            End If
                        Catch
                            startCellEnd = -1
                        End Try

                        If startCellEnd > 0 AndAlso selEnd > startCellEnd Then
                            LogReplaceDiag($"PASS1 REJECTING cell-crossing match [{selStart},{selEnd}]; startCellEnd={startCellEnd}")
                            Debug.WriteLine($"ExecuteReplaceCommand: rejecting cell-crossing match [{selStart},{selEnd}], cellEnd={startCellEnd}")

                            Dim escapePos As Integer = Math.Min(startCellEnd, searchEnd)
                            If escapePos <= requestedSearchStart Then
                                Exit Do
                            End If

                            doc.Application.Selection.SetRange(escapePos, escapePos)
                            requestedSearchStart = escapePos
                            lastFoundEnd = Math.Max(lastFoundEnd, escapePos - 1)
                            Continue Do
                        End If
                    End If
                Catch ex As Exception
                    LogReplaceDiag($"PASS1 cell-containment check failed: {ex.Message}")
                End Try

                ' Reject hits that move backwards or that do not advance.
                ' In tables Word can re-return the same logical cell match with a
                ' different End position because of cell markers, so checking End
                ' alone is not sufficient.
                Dim hitWentBackwards As Boolean = selStart < requestedSearchStart
                Dim hitDidNotAdvance As Boolean = selEnd <= lastFoundEnd

                If hitWentBackwards OrElse hitDidNotAdvance Then
                    Debug.WriteLine($"ExecuteReplaceCommand: rejecting stale/backward hit [{selStart},{selEnd}] (requestedSearchStart={requestedSearchStart}, lastFoundEnd={lastFoundEnd})")
                    LogReplaceDiag($"PASS1 rejecting stale/backward hit; hitWentBackwards={hitWentBackwards}; hitDidNotAdvance={hitDidNotAdvance}; {DescribeSelectionState(doc.Application.Selection)}")

                    Dim forcePos As Integer = Math.Max(requestedSearchStart + 1, lastFoundEnd + 1)

                    If forcePos >= searchEnd Then
                        Debug.WriteLine("ExecuteReplaceCommand: forced position past searchEnd, exiting")
                        LogReplaceDiag($"PASS1 forcePos {forcePos} >= searchEnd {searchEnd}; exiting")
                        Exit Do
                    End If

                    doc.Application.Selection.SetRange(forcePos, searchEnd)
                    requestedSearchStart = forcePos
                    Debug.WriteLine($"ExecuteReplaceCommand: forced selection to [{forcePos},{searchEnd}]")
                    LogReplaceDiag($"PASS1 forced selection=[{forcePos},{searchEnd}] {DescribeSelectionState(doc.Application.Selection)}")
                    Continue Do
                End If

                ' Record this match
                lastFoundEnd = selEnd
                matchPositions.Add((selStart, selEnd))
                Debug.WriteLine($"ExecuteReplaceCommand: stored match #{matchPositions.Count} at [{selStart},{selEnd}]")
                LogReplaceDiag($"PASS1 stored match #{matchPositions.Count} at [{selStart},{selEnd}]")

                ' Advance past current match
                doc.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                Debug.WriteLine($"ExecuteReplaceCommand: collapsed to {doc.Application.Selection.Start}")
                LogReplaceDiag($"PASS1 after collapse: {DescribeSelectionState(doc.Application.Selection)}")

                ' Handle table cell boundaries to avoid getting stuck at cell end marker
                Try
                    Dim isInTable As Boolean = False
                    Try
                        isInTable = CBool(doc.Application.Selection.Information(Word.WdInformation.wdWithInTable))
                    Catch ex As Exception
                        Debug.WriteLine($"ExecuteReplaceCommand: wdWithInTable check failed: {ex.Message}")
                        LogReplaceDiag($"PASS1 wdWithInTable check failed: {ex.Message}")
                    End Try

                    If isInTable Then
                        Debug.WriteLine("ExecuteReplaceCommand: in table, checking cell boundary")
                        LogReplaceDiag($"PASS1 in table before boundary handling: {DescribeSelectionState(doc.Application.Selection)}")

                        Dim cel As Word.Cell = Nothing
                        Try
                            If doc.Application.Selection.Cells.Count > 0 Then
                                cel = doc.Application.Selection.Cells(1)
                            End If
                        Catch ex As Exception
                            Debug.WriteLine($"ExecuteReplaceCommand: Cells access failed: {ex.Message}")
                            LogReplaceDiag($"PASS1 Cells access failed: {ex.Message}")
                        End Try

                        If cel IsNot Nothing Then
                            Dim selEndPos As Integer = doc.Application.Selection.End
                            Dim celRangeEnd As Integer = cel.Range.End
                            LogReplaceDiag($"PASS1 table boundary check: selEndPos={selEndPos}; cellEnd={celRangeEnd}")

                            If selEndPos >= celRangeEnd - 1 Then
                                doc.Application.Selection.SetRange(celRangeEnd, celRangeEnd)
                                Debug.WriteLine($"ExecuteReplaceCommand: jumped past cell end to {celRangeEnd}")
                                LogReplaceDiag($"PASS1 jumped past cell end to {celRangeEnd}; {DescribeSelectionState(doc.Application.Selection)}")
                            End If
                        End If
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"ExecuteReplaceCommand: table navigation failed: {ex.Message}")
                    LogReplaceDiag($"PASS1 table navigation failed: {ex.Message}")
                End Try

                ' Check if past search boundary
                Dim currentPos As Integer = doc.Application.Selection.Start
                Debug.WriteLine($"ExecuteReplaceCommand: after advance, position={currentPos}, searchEnd={searchEnd}")
                LogReplaceDiag($"PASS1 after advance: currentPos={currentPos}; searchEnd={searchEnd}; {DescribeSelectionState(doc.Application.Selection)}")

                If currentPos >= searchEnd Then
                    Debug.WriteLine("ExecuteReplaceCommand: past searchEnd, exiting loop")
                    LogReplaceDiag($"PASS1 exiting because currentPos {currentPos} >= searchEnd {searchEnd}")
                    Exit Do
                End If

                ' Extend selection to remaining search scope.
                ' In whole-document searches, keep the selection collapsed.
                ' A non-collapsed Selection that crosses table cells can snap back
                ' to the start of the previous cell, which causes the same hit to
                ' be returned again and again.
                Try
                    If OnlySelection Then
                        doc.Application.Selection.SetRange(currentPos, searchEnd)
                    Else
                        doc.Application.Selection.SetRange(currentPos, currentPos)
                    End If

                    requestedSearchStart = currentPos
                    LogReplaceDiag($"PASS1 prepared next search range=[{currentPos},{If(OnlySelection, searchEnd, currentPos)}] {DescribeSelectionState(doc.Application.Selection)}")
                Catch ex As Exception
                    Debug.WriteLine($"ExecuteReplaceCommand: SetRange({currentPos},{searchEnd}) failed: {ex.Message}")
                    LogReplaceDiag($"PASS1 SetRange({currentPos},{searchEnd}) failed: {ex.Message}")
                    Exit Do
                End Try
            Loop

            Debug.WriteLine($"ExecuteReplaceCommand: PASS 1 complete, found {matchPositions.Count} matches")

            If matchPositions.Count = 0 Then
                CommandsList = $"Note: The search term '{oldText}' was not found." & Environment.NewLine & CommandsList
                Try
                    doc.Application.Selection.SetRange(savedSelectionStart, savedSelectionEnd)
                Catch
                End Try
                Return False
            End If

            ' ─────────────────────────────────────────────────────────────────────
            ' PASS 2: REPLACE IN REVERSE ORDER (last match first)
            ' ─────────────────────────────────────────────────────────────────────
            Debug.WriteLine("ExecuteReplaceCommand: PASS 2 - replacing in reverse order...")
            Dim replaceCount As Integer = 0

            Using ThisAddIn.BeginMarkupAuthorScope(doc.Application)
                For i As Integer = matchPositions.Count - 1 To 0 Step -1
                    Dim mStart As Integer = matchPositions(i).Start
                    Dim mEnd As Integer = matchPositions(i).End

                    ' Avoid replacing the cell's terminal paragraph mark.
                    ' If the match ends one character before the cell end (i.e. on
                    ' the cell's last vbCr), shrink the match so the cell marker
                    ' is not touched, and strip the trailing paragraph mark from
                    ' newText so the replacement does not push content into the
                    ' next cell.
                    Dim adjustedNewText As String = newText
                    Try
                        Dim probe As Word.Range = doc.Range(mStart, mEnd)
                        Dim isInCell As Boolean = False
                        Try
                            isInCell = CBool(probe.Information(Word.WdInformation.wdWithInTable))
                        Catch
                            isInCell = False
                        End Try

                        If isInCell AndAlso probe.Cells.Count > 0 Then
                            Dim cellRangeEnd As Integer = probe.Cells(1).Range.End
                            ' Word stores cell-end marker as the last character of cell range
                            If mEnd = cellRangeEnd - 1 Then
                                Dim tailChar As String = doc.Range(mEnd - 1, mEnd).Text
                                If tailChar = vbCr OrElse tailChar = vbLf Then
                                    mEnd -= 1
                                    LogReplaceDiag($"PASS2 shrunk match to avoid cell paragraph mark; new mEnd={mEnd}")
                                End If
                                If adjustedNewText.EndsWith(vbCrLf, StringComparison.Ordinal) Then
                                    adjustedNewText = adjustedNewText.Substring(0, adjustedNewText.Length - 2)
                                ElseIf adjustedNewText.EndsWith(vbCr, StringComparison.Ordinal) OrElse
                                       adjustedNewText.EndsWith(vbLf, StringComparison.Ordinal) Then
                                    adjustedNewText = adjustedNewText.Substring(0, adjustedNewText.Length - 1)
                                End If
                                LogReplaceDiag($"PASS2 trimmed trailing paragraph mark from newText; len={adjustedNewText.Length}")
                            End If
                        End If
                    Catch ex As Exception
                        LogReplaceDiag($"PASS2 cell-boundary trim failed: {ex.Message}")
                    End Try


                    Debug.WriteLine($"ExecuteReplaceCommand: PASS2 replacing match #{i} at [{mStart},{mEnd}]")

                    Try
                        ' Select the matched text using stored positions
                        Debug.WriteLine($"ExecuteReplaceCommand: calling SetRange({mStart},{mEnd})")
                        doc.Application.Selection.SetRange(mStart, mEnd)
                        Debug.WriteLine($"ExecuteReplaceCommand: SetRange OK, selection=[{doc.Application.Selection.Start},{doc.Application.Selection.End}]")

                        If chkConvertMarkdown.Checked AndAlso newText.Length > 0 Then
                            ' Old replacement path so Markdown can be converted afterwards.
                            Debug.WriteLine($"ExecuteReplaceCommand: assigning Selection.Text = '{adjustedNewText}'")
                            doc.Application.Selection.Text = adjustedNewText
                            Debug.WriteLine($"ExecuteReplaceCommand: Selection.Text assigned OK, selection=[{doc.Application.Selection.Start},{doc.Application.Selection.End}]")

                            Dim actionEnd As Integer = Math.Min(mStart + newText.Length, doc.Content.End)
                            RememberLastActionRange(mStart, actionEnd)

                            Try
                                Debug.WriteLine("ExecuteReplaceCommand: applying ConvertMarkdownToWord")
                                Globals.ThisAddIn.ConvertMarkdownToWord()
                                Debug.WriteLine("ExecuteReplaceCommand: ConvertMarkdownToWord OK")
                            Catch ex As Exception
                                Debug.WriteLine($"ExecuteReplaceCommand: ConvertMarkdownToWord failed: {ex.Message}")
                            End Try
                        Else
                            ' Surgical replacement path for plain-text tracked edits.
                            Dim patchRange As Word.Range = doc.Range(mStart, mEnd)
                            Dim originalPatchText As String = patchRange.Text
                            Dim preserveTrailingCr As Boolean =
                                originalPatchText.EndsWith(vbCr, StringComparison.Ordinal) OrElse
                                originalPatchText.EndsWith(vbLf, StringComparison.Ordinal)

                            Debug.WriteLine("ExecuteReplaceCommand: applying surgical replacement")
                            Globals.ThisAddIn.ApplySurgicalReplacement(
                                originalPatchText,
                                adjustedNewText,
                                patchRange,
                                preserveTrailingCr)

                            RememberLastActionRange(patchRange.Start, patchRange.End)
                        End If

                        replaceCount += 1
                    Catch ex As Exception
                        Debug.WriteLine($"ExecuteReplaceCommand: ERROR replacing match #{i} at ({mStart},{mEnd}): {ex.GetType().Name}: {ex.Message}")
                        Debug.WriteLine($"ExecuteReplaceCommand: StackTrace: {ex.StackTrace}")
                    End Try
                Next
            End Using

            Debug.WriteLine($"ExecuteReplaceCommand: PASS 2 complete, replaced {replaceCount} of {matchPositions.Count}")

            ' ─────────────────────────────────────────────────────────────────────
            ' RESTORE SELECTION
            ' ─────────────────────────────────────────────────────────────────────
            Try
                Dim safeStart As Integer = Math.Max(doc.Content.Start, Math.Min(savedSelectionStart, doc.Content.End))
                Dim safeEnd As Integer = Math.Max(doc.Content.Start, Math.Min(savedSelectionEnd, doc.Content.End))
                Debug.WriteLine($"ExecuteReplaceCommand: restoring selection to [{safeStart},{safeEnd}]")
                doc.Application.Selection.SetRange(safeStart, safeEnd)
                doc.Application.Selection.Select()
            Catch ex As Exception
                Debug.WriteLine($"ExecuteReplaceCommand: restore selection failed: {ex.Message}")
                Try
                    doc.Application.Selection.SetRange(doc.Content.Start, doc.Content.Start)
                Catch
                End Try
            End Try

            Debug.WriteLine($"ExecuteReplaceCommand: END (success={replaceCount > 0})")
            Return replaceCount > 0

        Catch ex As System.Exception
            Debug.WriteLine($"ExecuteReplaceCommand: OUTER CATCH: {ex.GetType().Name}: {ex.Message}")
            Debug.WriteLine($"ExecuteReplaceCommand: StackTrace: {ex.StackTrace}")
#If DEBUG Then
            System.Diagnostics.Debugger.Break()
#End If
            MsgBox("Error in ExecuteReplaceCommand: " & ex.Message, MsgBoxStyle.Critical)
            Return False

        Finally
            Try
                If view IsNot Nothing Then
                    view.RevisionsView = originalRevisionsView
                    view.ShowRevisionsAndComments = originalShowRevisions
                End If
            Catch ex As Exception
                Debug.WriteLine($"ExecuteReplaceCommand: FINALLY view restore failed: {ex.Message}")
            End Try
            Try
                If doc IsNot Nothing Then
                    doc.TrackRevisions = trackChangesEnabled
                End If
            Catch ex As Exception
                Debug.WriteLine($"ExecuteReplaceCommand: FINALLY TrackRevisions restore failed: {ex.Message}")
            End Try
        End Try
    End Function

    Private Shared Function PreviewForLog(value As String, Optional maxLen As Integer = 120) As String
        If value Is Nothing Then Return "<null>"

        Dim s As String = value.
            Replace(vbCr, "\r").
            Replace(vbLf, "\n").
            Replace(ChrW(7), "\cell")

        If s.Length > maxLen Then
            s = s.Substring(0, maxLen) & "…"
        End If

        Return s
    End Function

    <System.Diagnostics.Conditional("DEBUG")>
    Private Sub LogReplaceDiag(message As String)
        If String.IsNullOrWhiteSpace(message) Then Return
        Debug.WriteLine($"{DateTime.Now:HH:mm:ss.fff} [ExecuteReplaceCommand] {message}")
    End Sub

    Private Function DescribeSelectionState(sel As Word.Selection) As String
        Try
            If sel Is Nothing Then
                Return "selection=<null>"
            End If

            Dim preview As String = "<unavailable>"
            Try
                preview = PreviewForLog(If(sel.Text, String.Empty), 80)
            Catch
            End Try

            Dim isInTable As Boolean = False
            Dim cellInfo As String = ""

            Try
                isInTable = CBool(sel.Information(Word.WdInformation.wdWithInTable))

                If isInTable AndAlso sel.Cells.Count > 0 Then
                    Dim currentCell As Word.Cell = sel.Cells(1)
                    cellInfo =
                        $" cell=[row={currentCell.RowIndex},col={currentCell.ColumnIndex}] cellRange=[{currentCell.Range.Start},{currentCell.Range.End}]"
                End If
            Catch ex As Exception
                cellInfo = $" cellInfoError='{ex.Message}'"
            End Try

            Return $"selection=[{sel.Start},{sel.End}] len={Math.Max(0, sel.End - sel.Start)} inTable={isInTable}{cellInfo} text='{preview}'"
        Catch ex As Exception
            Return $"selectionStateError='{ex.Message}'"
        End Try
    End Function


    ' =========================================================================
    ' Insert Before/After Command
    ' =========================================================================

    ''' <summary>
    ''' Inserts newText before or after all occurrences of searchText anchor.
    ''' Tries multiple search variants (original, trimmed) for flexibility.
    ''' Skips TOC ranges to prevent corruption.
    ''' </summary>
    ''' <param name="searchText">Anchor text to find (normalized via DecodeParagraphMarks)</param>
    ''' <param name="newText">Text to insert (normalized via DecodeParagraphMarks)</param>
    ''' <param name="OnlySelection">True to restrict to current selection</param>
    ''' <param name="InsertBefore">True for insertbefore, False for insertafter</param>
    ''' <returns>True if at least one insertion made</returns>
    ''' <remarks>
    ''' Search variants tried in order:
    ''' 1. Original searchText
    ''' 2. TrimEnd if has trailing spaces
    ''' 3. TrimStart if has leading spaces
    ''' 4. Fully trimmed if has both
    ''' 
    ''' Safety measures:
    ''' - Max 1000 iterations per variant
    ''' - Position tracking to detect stuck state
    ''' - TOC detection via TocEndIfInside (skips to end of TOC)
    ''' - Document end boundary guards (End-1 for insertion)
    ''' - Fallback to Selection.Text if Range creation fails
    ''' </remarks>
    Private Function ExecuteInsertBeforeAfterCommand(searchText As String, newText As String, Optional OnlySelection As Boolean = False, Optional InsertBefore As Boolean = False) As Boolean
        Dim doc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument

        Dim trackChangesEnabled As Boolean = doc.TrackRevisions
        Dim originalAuthor As String = doc.Application.UserName

        Try
            ' Normalize inputs
            searchText = DecodeParagraphMarks(searchText)
            newText = DecodeParagraphMarks(newText)

            If String.IsNullOrWhiteSpace(searchText) Then
                CommandsList = $"Note: Empty insertion anchor (ignored)." & Environment.NewLine & CommandsList
                Return False
            End If

            doc.Application.Activate()
            doc.Activate()
            doc.TrackRevisions = True

            ' Determine working range
            Dim workrange As Word.Range
            If OnlySelection Then
                If doc.Application.Selection Is Nothing OrElse doc.Application.Selection.Range.Text = "" Then
                    OnlySelection = False
                    workrange = doc.Content
                Else
                    workrange = doc.Application.Selection.Range
                End If
            Else
                workrange = doc.Content
            End If

            Dim found As Boolean = False
            Dim selectionStart As Integer = doc.Application.Selection.Start
            Dim selectionEnd As Integer = doc.Application.Selection.End

            ' Build list of search variants (handle whitespace issues)
            Dim searchAttempts As New List(Of String)
            searchAttempts.Add(searchText)
            If searchText.EndsWith(" ") Then searchAttempts.Add(searchText.TrimEnd())
            If searchText.StartsWith(" ") Then searchAttempts.Add(searchText.TrimStart())
            If searchText.StartsWith(" ") OrElse searchText.EndsWith(" ") Then searchAttempts.Add(searchText.Trim())
            searchAttempts = searchAttempts.Distinct().ToList()

            Using ThisAddIn.BeginMarkupAuthorScope(doc.Application)
                ' Try each search variant until match found
                For Each currentSearchText In searchAttempts
                    If found Then Exit For

                    Debug.WriteLine($"Trying search variant: '{currentSearchText}'")
                    doc.Application.Selection.SetRange(workrange.Start, workrange.End)

                    Dim maxIterations As Integer = 1000
                    Dim iterationCount As Integer = 0
                    Dim lastProcessedPosition As Integer = -1

                    Do While Globals.ThisAddIn.FindLongTextInChunks(currentSearchText, doc.Application.Selection, True) = True

                        If doc.Application.Selection Is Nothing Then Exit Do

                        System.Windows.Forms.Application.DoEvents()
                        If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then
                            CommandsList = $"Operation cancelled by user (ESC)." & Environment.NewLine & CommandsList
                            Exit Do
                        End If

                        ' Safety: prevent infinite loops
                        iterationCount += 1
                        If iterationCount > maxIterations Then
                            Debug.WriteLine($"ExecuteInsertBeforeAfterCommand: Max iterations ({maxIterations}) reached")
                            Exit Do
                        End If

                        ' Detect stuck state (same position)
                        If doc.Application.Selection.Start = lastProcessedPosition Then
                            Debug.WriteLine("ExecuteInsertBeforeAfterCommand: Stuck at same position, advancing")
                            doc.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                            doc.Application.Selection.Move(Word.WdUnits.wdCharacter, 1)
                            Continue Do
                        End If
                        lastProcessedPosition = doc.Application.Selection.Start

                        Dim foundRange As Word.Range = doc.Application.Selection.Range.Duplicate

                        ' Skip TOC ranges to prevent corruption
                        Dim tocEnd As Integer = TocEndIfInside(foundRange, doc)
                        If tocEnd > 0 Then
                            Debug.WriteLine("ExecuteInsertBeforeAfterCommand: Match in TOC -> skipping")
                            Dim searchLimit As Integer = If(OnlySelection, selectionEnd, doc.Content.End)
                            Dim continuePos As Integer = Math.Min(tocEnd, searchLimit)
                            If continuePos >= searchLimit Then
                                Exit Do
                            Else
                                doc.Application.Selection.SetRange(continuePos, searchLimit)
                                Continue Do
                            End If
                        End If

                        found = True
                        Debug.WriteLine($"Found match at position {foundRange.Start}")

                        Dim foundStart As Integer = foundRange.Start
                        Dim foundEnd As Integer = foundRange.End
                        Dim insertPosition As Integer = If(InsertBefore, foundStart, foundEnd)

                        ' Handle document end boundary (End includes final paragraph mark)
                        Dim docContentEnd As Integer = doc.Content.End
                        If insertPosition >= docContentEnd Then
                            If Not InsertBefore Then
                                insertPosition = docContentEnd - 1
                            End If
                        End If
                        insertPosition = Math.Max(doc.Content.Start, Math.Min(insertPosition, docContentEnd - 1))

                        Try
                            ' Primary method: create Range and insert
                            Dim insertRange As Word.Range = doc.Range(insertPosition, insertPosition)
                            insertRange.Text = newText

                            RememberLastActionRange(insertPosition, Math.Min(insertPosition + Len(newText), doc.Content.End))

                            ' Apply Markdown if enabled
                            If chkConvertMarkdown.Checked AndAlso newText.Length > 0 Then
                                Try
                                    Dim conversionStart As Integer = insertPosition
                                    Dim conversionEnd As Integer = Math.Min(insertPosition + Len(newText), doc.Content.End)
                                    doc.Range(conversionStart, conversionEnd).Select()
                                    Globals.ThisAddIn.ConvertMarkdownToWord()
                                Catch
                                    ' Best effort
                                End Try
                            End If
                        Catch rangeEx As Exception
                            ' Fallback: use Selection to insert
                            Debug.WriteLine($"Range creation failed at {insertPosition}, trying alternative")
                            Try
                                doc.Application.Selection.SetRange(insertPosition, insertPosition)
                                doc.Application.Selection.Text = newText
                                If chkConvertMarkdown.Checked AndAlso newText.Length > 0 Then
                                    Globals.ThisAddIn.ConvertMarkdownToWord()
                                End If
                            Catch altEx As Exception
                                Debug.WriteLine($"Alternative insertion failed: {altEx.Message}")
                                Continue Do
                            End Try
                        End Try

                        ' Calculate next search position
                        Dim continuePosition As Integer
                        If InsertBefore Then
                            continuePosition = insertPosition + Len(newText) + (foundEnd - foundStart)
                        Else
                            continuePosition = insertPosition + Len(newText)
                        End If

                        ' Ensure forward progress
                        If continuePosition <= lastProcessedPosition Then
                            continuePosition = lastProcessedPosition + 1
                        End If

                        ' Adjust selection end if text inserted
                        If OnlySelection Then
                            selectionEnd = selectionEnd + Len(newText)
                        End If

                        ' Check if reached end of search range
                        If OnlySelection Then
                            If continuePosition >= selectionEnd Then Exit Do
                            Dim safeEnd As Integer = Math.Min(selectionEnd, doc.Content.End)
                            doc.Application.Selection.SetRange(continuePosition, safeEnd)
                        Else
                            If continuePosition >= doc.Content.End Then Exit Do
                            doc.Application.Selection.SetRange(continuePosition, doc.Content.End)
                        End If
                    Loop
                Next
            End Using

            If Not found Then
                CommandsList = $"Note: The search term was not found." & Environment.NewLine & CommandsList
            End If

            ' Restore original selection with boundary guards
            Try
                Dim safeStart As Integer = Math.Max(doc.Content.Start, Math.Min(selectionStart, doc.Content.End))
                Dim safeEnd As Integer = Math.Max(doc.Content.Start, Math.Min(selectionEnd, doc.Content.End))
                doc.Application.Selection.SetRange(safeStart, safeEnd)
                doc.Application.Selection.Select()
            Catch
                doc.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseStart)
            End Try

            Return found

        Catch ex As System.Exception
#If DEBUG Then
            Debug.WriteLine("Error: " & ex.Message)
            Debug.WriteLine("Stacktrace: " & ex.StackTrace)
            System.Diagnostics.Debugger.Break()
#End If
            MsgBox("Error in ExecuteInsertBeforeAfterCommand: " & ex.Message, MsgBoxStyle.Critical)
            Return False

        Finally
            doc.TrackRevisions = trackChangesEnabled
        End Try
    End Function

    ' =========================================================================
    ' Insert Command (at cursor)
    ' =========================================================================

    ''' <summary>
    ''' Inserts newText at current cursor position with tracked changes.
    ''' Collapses selection to start before insertion.
    ''' Applies Markdown formatting if chkConvertMarkdown enabled.
    ''' </summary>
    ''' <param name="newText">Text to insert (normalized via DecodeParagraphMarks)</param>
    ''' <returns>True on success, False on error</returns>
    ''' <remarks>
    ''' Simplest command - no search, just inserts at current caret position.
    ''' Normalizes line breaks to vbCr (Word's internal format).
    ''' Restores original TrackRevisions state in Finally block.
    ''' </remarks>
    Private Function ExecuteInsertCommand(newText As String) As Boolean
        Dim doc = Globals.ThisAddIn.Application.ActiveDocument
        Dim trackChangesEnabled = doc.TrackRevisions

        Try
            ' Normalize line breaks to Word format
            newText = DecodeParagraphMarks(newText)
            newText = newText.Replace(vbCrLf, vbCr).Replace(vbLf, vbCr)

            doc.TrackRevisions = True
            Using ThisAddIn.BeginMarkupAuthorScope(doc.Application)
                Dim selection = doc.Application.Selection
                selection.Collapse(Word.WdCollapseDirection.wdCollapseStart)

                Dim insertStart As Integer = selection.Start
                selection.Text = newText

                RememberLastActionRange(insertStart, Math.Min(insertStart + newText.Length, doc.Content.End))

                If chkConvertMarkdown.Checked Then
                    Globals.ThisAddIn.ConvertMarkdownToWord()
                End If
            End Using

            Return True
        Catch ex As Exception
            MsgBox("Error in ExecuteInsertCommand: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        Finally
            doc.TrackRevisions = trackChangesEnabled
        End Try
    End Function


End Class

' =========================================================================
' HTML/Markdown Rendering - WebBrowser Chat Display
' =========================================================================

''' <summary>
''' Partial class extension for HTML/Markdown rendering functionality.
''' Manages WebBrowser control for rich chat display with Markdig pipeline.
''' </summary>
''' <remarks>
''' This section handles:
''' - WebBrowser control initialization and event handling
''' - Markdown-to-HTML conversion via Markdig
''' - Link instrumentation for external browser opening
''' - Chat message queuing and rendering
''' - "Thinking..." placeholder management
''' - HTML persistence to My.Settings
''' 
''' Uses legacy IE rendering engine (WebBrowser control limitation).
''' </remarks>
Partial Public Class frmAIChat

    ' =========================================================================
    ' Private Fields - HTML Rendering State
    ' =========================================================================

    ''' <summary>Tracks whether document-level click handler has been wired to prevent duplicates</summary>
    Private _docClickHooked As Boolean = False

    ''' <summary>
    ''' WebBrowser control for rendering chat with Markdown-formatted HTML.
    ''' Overlays txtChatHistory when HTML mode active. Uses legacy IE rendering engine.
    ''' </summary>
    Private ReadOnly wbChat As New WebBrowser() With {
        .Dock = DockStyle.Fill,
        .AllowWebBrowserDrop = False,
        .IsWebBrowserContextMenuEnabled = True,
        .WebBrowserShortcutsEnabled = True,
        .ScriptErrorsSuppressed = True
    }

    ''' <summary>True when WebBrowser document is ready to receive HTML fragments</summary>
    Private _htmlReady As Boolean = False

    ''' <summary>Queue of HTML fragments waiting to be appended when WebBrowser becomes ready</summary>
    Private ReadOnly _htmlQueue As New List(Of String)()

    ''' <summary>
    ''' Markdig pipeline for Markdown-to-HTML conversion.
    ''' Configured with advanced extensions (tables, footnotes), emoji support, and soft line breaks.
    ''' </summary>
    Private ReadOnly _mdPipeline As MarkdownPipeline =
        New MarkdownPipelineBuilder().
            UseAdvancedExtensions().
            UseSoftlineBreakAsHardlineBreak().
            Build()

    ''' <summary>DOM ID of current "Thinking..." placeholder for removal when LLM responds</summary>
    Private _lastThinkingId As String = Nothing

    ' =========================================================================
    ' Link Click Handler
    ' =========================================================================

    ''' <summary>
    ''' Wires document-level click handler for external link opening.
    ''' Called when WebBrowser document is ready. Prevents duplicate handler attachment.
    ''' </summary>
    Private Sub WireDocumentClick()
        If wbChat Is Nothing OrElse wbChat.Document Is Nothing Then Return
        Try
            ' Remove existing handler to prevent duplicates
            RemoveHandler wbChat.Document.Click, AddressOf Doc_Click
        Catch
            ' Ignore if handler not already attached
        End Try
        AddHandler wbChat.Document.Click, AddressOf Doc_Click
        _docClickHooked = True
    End Sub

    ''' <summary>
    ''' Handles click events in HTML document. Finds nearest anchor tag and opens externally.
    ''' </summary>
    ''' <param name="sender">Event source (HTML document)</param>
    ''' <param name="e">Click event args</param>
    ''' <remarks>
    ''' Walks up DOM tree from clicked element to find nearest anchor tag.
    ''' Only opens external links (http://, https://, mailto:).
    ''' Prevents internal WebBrowser navigation by setting ReturnValue=False.
    ''' </remarks>
    Private Sub Doc_Click(sender As Object, e As HtmlElementEventArgs)
        Try
            Dim el As HtmlElement = wbChat.Document.ActiveElement

            ' Walk up DOM tree to find nearest anchor
            While el IsNot Nothing AndAlso Not String.Equals(el.TagName, "A", StringComparison.OrdinalIgnoreCase)
                el = el.Parent
            End While

            If el Is Nothing Then Return

            Dim href As String = el.GetAttribute("href")
            If String.IsNullOrWhiteSpace(href) Then Return

            ' Only handle external protocols
            Dim lower = href.Trim().ToLowerInvariant()
            If lower.StartsWith("http://") OrElse lower.StartsWith("https://") OrElse lower.StartsWith("mailto:") Then
                Process.Start(New ProcessStartInfo(href) With {.UseShellExecute = True})
                ' Prevent internal WebBrowser navigation
                If e IsNot Nothing Then
                    e.ReturnValue = False
                    e.BubbleEvent = False
                End If
            End If
        Catch
            ' Silently ignore errors
        End Try
    End Sub

    ' =========================================================================
    ' COM Bridge for JavaScript Interaction
    ' =========================================================================

    ''' <summary>
    ''' COM-visible bridge class for JavaScript-to-.NET interaction.
    ''' Exposed via WebBrowser.ObjectForScripting to allow JavaScript calls.
    ''' </summary>
    ''' <remarks>
    ''' JavaScript in HTML document calls window.external.OpenLink(url) to open links.
    ''' This avoids internal WebBrowser navigation and forces external browser.
    ''' </remarks>
    <System.Runtime.InteropServices.ComVisible(True)>
    Public Class BrowserBridge
        ''' <summary>Opens URL in default external browser</summary>
        Public Sub OpenLink(url As String)
            Try
                If String.IsNullOrEmpty(url) Then Return
                Process.Start(New ProcessStartInfo(url) With {.UseShellExecute = True})
            Catch
                ' Silently ignore errors
            End Try
        End Sub
    End Class

    ' =========================================================================
    ' Load Context and File Handling
    ' =========================================================================

    Private Sub AppendSystemMessage(message As String)
        Try
            AppendToChatHistory(Environment.NewLine & "[System] " & message & Environment.NewLine)
        Catch
        End Try
        Try
            AppendHtml($"<div class='msg system'><span class='content'>{HtmlEncode("[System] " & message)}</span></div>")
            PersistChatHtml()
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Gets the Red Ink storage directory in the user's application data folder
    ''' (same location convention used by DiscussInky).
    ''' </summary>
    Private Function GetRedInkStorageDirectoryPath() As String
        Dim storageDir = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "redink")
        Try
            If Not System.IO.Directory.Exists(storageDir) Then
                System.IO.Directory.CreateDirectory(storageDir)
            End If
        Catch
        End Try
        Return storageDir
    End Function

    Private Function GetPersistedContextFilePath() As String
        Return System.IO.Path.Combine(GetRedInkStorageDirectoryPath(), PersistedContextFileName)
    End Function

    Private Function GetContextDragDropFilter() As String
        If ThisAddIn.INI_AllowLegacyDocFiles Then
            Return "Supported Context Files|*.txt;*.rtf;*.doc;*.docx;*.xlsx;*.pdf;*.pptx;*.msg;*.eml;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm;*.md;*.vb;*.cs;*.js;*.ts;*.py;*.java;*.cpp;*.c;*.h;*.sql;*.yaml;*.yml|All Files (*.*)|*.*"
        End If

        Return "Supported Context Files|*.txt;*.rtf;*.docx;*.xlsx;*.pdf;*.pptx;*.msg;*.eml;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm;*.md;*.vb;*.cs;*.js;*.ts;*.py;*.java;*.cpp;*.c;*.h;*.sql;*.yaml;*.yml|All Files (*.*)|*.*"
    End Function


    Private Sub UpdatePersistContextTooltip()
        If chkPersistContext.Checked Then
            If HasLoadedIndex() Then
                _contextToolTip.SetToolTip(chkPersistContext, "Index currently stored in: " & GetPersistedIndexFilePath())
            Else
                _contextToolTip.SetToolTip(chkPersistContext, "Currently stored in: " & GetPersistedContextFilePath())
            End If
        Else
            _contextToolTip.SetToolTip(chkPersistContext, "")
        End If
    End Sub

    Private Sub DeletePersistedContextFile(showMessage As Boolean)
        Try
            Dim persistPath = GetPersistedContextFilePath()
            If System.IO.File.Exists(persistPath) Then
                System.IO.File.Delete(persistPath)
                If showMessage Then
                    AppendSystemMessage("Persisted context file deleted.")
                End If
            End If
        Catch ex As System.Exception
            If showMessage Then
                AppendSystemMessage($"Failed to delete persisted context: {ex.Message}")
            End If
        End Try
    End Sub

    Private Sub PersistLoadedContextToTempFile()
        If String.IsNullOrWhiteSpace(_loadedContextContent) Then Return
        System.IO.File.WriteAllText(GetPersistedContextFilePath(), _loadedContextContent, System.Text.Encoding.UTF8)
    End Sub

    ' =========================================================================
    ' Loaded Semantic Index (either a document context or an index is active)
    ' =========================================================================

    ''' <summary>Full path of the durably persisted index copy under %AppData%\redink\.</summary>
    Private Function GetPersistedIndexFilePath() As String
        Return System.IO.Path.Combine(GetRedInkStorageDirectoryPath(), PersistedIndexFileName)
    End Function

    ''' <summary>True when a semantic index is currently attached to the session.</summary>
    Private Function HasLoadedIndex() As Boolean
        Return Not String.IsNullOrWhiteSpace(_loadedIndexSourcePath) OrElse
               Not String.IsNullOrWhiteSpace(_loadedIndexDisplayName)
    End Function

    ''' <summary>
    ''' Resolves the index file to search: the original source when available, otherwise the
    ''' durably persisted copy. Returns Nothing when neither is present on disk.
    ''' </summary>
    Private Function GetActiveIndexPath() As String
        If Not String.IsNullOrWhiteSpace(_loadedIndexSourcePath) AndAlso System.IO.File.Exists(_loadedIndexSourcePath) Then
            Return _loadedIndexSourcePath
        End If

        Dim persisted As String = GetPersistedIndexFilePath()
        If System.IO.File.Exists(persisted) Then
            Return persisted
        End If

        Return Nothing
    End Function

    ''' <summary>Attaches a semantic index, replacing any loaded document context.</summary>
    Private Function AttachLoadedIndex(indexPath As String) As Boolean
        If String.IsNullOrWhiteSpace(indexPath) OrElse Not System.IO.File.Exists(indexPath) Then
            AppendSystemMessage("The selected index file does not exist.")
            Return False
        End If

        ' Loading an index replaces any plain document context (either a document or an index).
        ClearLoadedContextOnly()

        _loadedIndexSourcePath = indexPath
        _loadedIndexDisplayName = System.IO.Path.GetFileName(indexPath)
        _cachedLoadedIndexPath = _loadedIndexSourcePath
        _cachedLoadedIndexDisplayName = _loadedIndexDisplayName
        _semanticConversationState = New SharedMethods.SemanticSearchConversationState()

        Try
            My.Settings.ChatContextPath = indexPath
            My.Settings.Save()
        Catch
        End Try

        Return True
    End Function

    ''' <summary>Clears the in-memory document context and its persisted file, leaving the index intact.</summary>
    Private Sub ClearLoadedContextOnly()
        _loadedContextContent = Nothing
        _loadedContextPath = Nothing
        _cachedLoadedContextContent = Nothing
        _cachedLoadedContextPath = Nothing
        DeletePersistedContextFile(False)
    End Sub

    ''' <summary>Clears the loaded index, its cache, retrieval state, and persisted copy.</summary>
    Private Sub ClearLoadedIndexOnly()
        _loadedIndexSourcePath = Nothing
        _loadedIndexDisplayName = Nothing
        _cachedLoadedIndexPath = Nothing
        _cachedLoadedIndexDisplayName = Nothing
        _semanticConversationState = New SharedMethods.SemanticSearchConversationState()
        DeletePersistedIndexFile(False)
    End Sub

    ''' <summary>Deletes the durably persisted index copy under %AppData%\redink\.</summary>
    Private Sub DeletePersistedIndexFile(showMessage As Boolean)
        Try
            Dim persistPath As String = GetPersistedIndexFilePath()
            If System.IO.File.Exists(persistPath) Then
                System.IO.File.Delete(persistPath)
                If showMessage Then
                    AppendSystemMessage("Persisted index file deleted.")
                End If
            End If
        Catch ex As System.Exception
            If showMessage Then
                AppendSystemMessage($"Failed to delete persisted index: {ex.Message}")
            End If
        End Try
    End Sub

    ''' <summary>Copies the active index (exact bytes) into durable %AppData%\redink\ storage.</summary>
    Private Sub PersistLoadedIndexToAppData()
        Dim active As String = GetActiveIndexPath()
        If String.IsNullOrWhiteSpace(active) OrElse Not System.IO.File.Exists(active) Then Return

        Dim persistPath As String = GetPersistedIndexFilePath()
        If String.Equals(System.IO.Path.GetFullPath(active), System.IO.Path.GetFullPath(persistPath), StringComparison.OrdinalIgnoreCase) Then
            Return
        End If

        System.IO.File.Copy(active, persistPath, True)
    End Sub

    ''' <summary>Offers to persist a newly attached index when persistence is currently off.</summary>
    Private Sub OfferIndexPersistence()
        Dim answer = ShowCustomYesNoBox(
            $"The index '{_loadedIndexDisplayName}' is currently referenced from its original location only. " &
            "Do you want to persist a copy to durable storage so it is retained across restarts?",
            "Yes, persist",
            "No, keep temporary")

        If answer <> 1 Then Return

        _isUpdatingPersistContextCheckbox = True
        chkPersistContext.Checked = True
        _isUpdatingPersistContextCheckbox = False

        Try
            PersistLoadedIndexToAppData()
            AppendSystemMessage($"Index '{_loadedIndexDisplayName}' persisted to durable storage.")
        Catch ex As System.Exception
            AppendSystemMessage($"Failed to persist index: {ex.Message}")
        End Try

        Try
            My.Settings.ChatPersistContext = True
            My.Settings.Save()
        Catch
        End Try

        UpdatePersistContextTooltip()
    End Sub

    ''' <summary>
    ''' Handles the persist checkbox for the index channel: persists a copy when enabled, or deletes
    ''' the persisted copy when disabled (falling back to the original source, or warning if gone).
    ''' </summary>
    Private Sub HandleIndexPersistenceToggle()
        Dim persistPath As String = GetPersistedIndexFilePath()

        If chkPersistContext.Checked Then
            Dim active As String = GetActiveIndexPath()
            If Not String.IsNullOrWhiteSpace(active) AndAlso System.IO.File.Exists(active) Then
                Try
                    PersistLoadedIndexToAppData()
                    AppendSystemMessage($"Index '{_loadedIndexDisplayName}' persisted to durable storage.")
                Catch ex As System.Exception
                    AppendSystemMessage($"Failed to persist index: {ex.Message}")
                End Try
            Else
                AppendSystemMessage("No index available to persist.")
            End If
        Else
            If System.IO.File.Exists(persistPath) Then
                Dim answer = ShowCustomYesNoBox(
                    "Do you want to delete the persisted index file? The chatbot will then rely on the original index file, if still available.",
                    "Yes, delete",
                    "No, keep it")

                If answer = 1 Then
                    DeletePersistedIndexFile(True)
                    If String.IsNullOrWhiteSpace(_loadedIndexSourcePath) OrElse Not System.IO.File.Exists(_loadedIndexSourcePath) Then
                        _loadedIndexSourcePath = Nothing
                        _loadedIndexDisplayName = Nothing
                        _cachedLoadedIndexPath = Nothing
                        _cachedLoadedIndexDisplayName = Nothing
                        _semanticConversationState = New SharedMethods.SemanticSearchConversationState()

                        Try
                            My.Settings.ChatContextPath = ""
                            My.Settings.Save()
                        Catch
                        End Try

                        AppendSystemMessage("The original index file is no longer available. The loaded index was removed.")
                        UpdateLoadContextButtonText()
                    End If
                Else
                    _isUpdatingPersistContextCheckbox = True
                    chkPersistContext.Checked = True
                    _isUpdatingPersistContextCheckbox = False
                    Return
                End If
            End If
        End If

        My.Settings.ChatPersistContext = chkPersistContext.Checked
        My.Settings.Save()
        UpdatePersistContextTooltip()
    End Sub

    ''' <summary>
    ''' Retrieves the most relevant original excerpts from the loaded index for the current message,
    ''' reporting per-step progress through the supplied callback (mirrors DiscussInky's behavior).
    ''' </summary>
    Private Async Function BuildIndexExcerptAsync(queryText As String,
                                                  conversation As String,
                                                  reportStatus As System.Action(Of String)) As Task(Of String)
        Dim activePath As String = GetActiveIndexPath()
        If String.IsNullOrWhiteSpace(activePath) OrElse Not System.IO.File.Exists(activePath) Then
            reportStatus?.Invoke("The loaded index is unavailable.")
            Return ""
        End If

        If String.IsNullOrWhiteSpace(queryText) Then Return ""

        reportStatus?.Invoke($"Searching index '{_loadedIndexDisplayName}' ...")

        Try
            Dim options As New SharedMethods.SemanticSearchRetrievalOptions() With {
                .SpecialTaskName = "Indexer"
            }

            Dim retrieval As SharedMethods.SemanticSearchRetrievalResult =
                Await SharedMethods.RetrieveSemanticSearchAsync(
                    activePath,
                    _context,
                    queryText,
                    If(conversation, ""),
                    _semanticConversationState,
                    options).ConfigureAwait(False)

            If retrieval IsNot Nothing AndAlso
               retrieval.IsIndexed AndAlso
               Not String.IsNullOrWhiteSpace(retrieval.ReducedSourceText) Then

                Dim matchCount As Integer =
                    If(retrieval.SelectedEntryIds IsNot Nothing, retrieval.SelectedEntryIds.Count, 0)

                reportStatus?.Invoke($"Index '{_loadedIndexDisplayName}' — {matchCount:N0} relevant segment(s) found.")

                Dim sb As New StringBuilder()
                sb.AppendLine($"<document name=""{_loadedIndexDisplayName}"">")
                sb.AppendLine(retrieval.ReducedSourceText)
                sb.AppendLine("</document>")
                Return sb.ToString().TrimEnd()
            Else
                reportStatus?.Invoke($"Index '{_loadedIndexDisplayName}' — no relevant material found.")
                Return ""
            End If
        Catch ex As System.Exception
            reportStatus?.Invoke($"Index retrieval failed: {ex.Message}")
            Return ""
        End Try
    End Function

    Private Async Function LoadSingleContextFileAsync(filePath As String, askUser As Boolean) As Task(Of (Content As String, PdfMayBeIncomplete As Boolean))
        If String.IsNullOrWhiteSpace(filePath) OrElse Not System.IO.File.Exists(filePath) Then
            Return ("", False)
        End If

        ' Silent suppresses per-file error boxes; AskUser lets GetFileContentEx handle the OCR prompt itself.
        Dim result = Await Globals.ThisAddIn.GetFileContentEx(filePath, True, False, askUser)
        Return (result.Content, result.PdfMayBeIncomplete)
    End Function

    Private Async Function LoadContextFromPathAsync(selectedPath As String, interactive As Boolean) As Task(Of (CombinedContent As String, DisplayPath As String, LoadedCount As Integer, Summary As String))
        Dim isFile As Boolean = System.IO.File.Exists(selectedPath)
        Dim isDirectory As Boolean = System.IO.Directory.Exists(selectedPath)

        If Not isFile AndAlso Not isDirectory Then
            Return ("", "", 0, "Selected path does not exist.")
        End If

        Dim filesToProcess As New List(Of String)
        Dim failedFiles As New List(Of String)
        Dim loadedFiles As New List(Of Tuple(Of String, Integer))
        Dim pdfsWithPossibleImages As New List(Of String)
        Dim ignoredCount As Integer = 0

        If isFile Then
            Dim ext = System.IO.Path.GetExtension(selectedPath).ToLowerInvariant()
            If Array.IndexOf(SupportedContextExtensions, ext) < 0 Then
                Return ("", "", 0, $"The selected file type '{ext}' is not supported for loaded context.")
            End If

            filesToProcess.Add(selectedPath)
        Else
            Dim allFiles = System.IO.Directory.GetFiles(selectedPath, "*.*", System.IO.SearchOption.TopDirectoryOnly)

            For Each f In allFiles
                Dim ext = System.IO.Path.GetExtension(f).ToLowerInvariant()
                If Array.IndexOf(SupportedContextExtensions, ext) >= 0 Then
                    filesToProcess.Add(f)
                Else
                    ignoredCount += 1
                End If
            Next

            If filesToProcess.Count > 50 Then
                If interactive Then
                    Dim truncateAnswer = ShowCustomYesNoBox(
                        $"The directory contains {filesToProcess.Count} supported files, but the maximum is 50." & vbCrLf & vbCrLf &
                        "Only the first 50 files will be loaded. Continue?",
                        "Yes, continue",
                        "No, abort")

                    If truncateAnswer <> 1 Then
                        Return ("", "", 0, "")
                    End If
                End If

                filesToProcess = filesToProcess.GetRange(0, 50)
            ElseIf interactive AndAlso filesToProcess.Count > 10 Then
                Dim confirmAnswer = ShowCustomYesNoBox(
                    $"The directory contains {filesToProcess.Count} files to load. Continue?",
                    "Yes, continue",
                    "No, abort")

                If confirmAnswer <> 1 Then
                    Return ("", "", 0, "")
                End If
            End If

            If filesToProcess.Count = 0 Then
                Return ("", "", 0, $"No supported files found in directory '{selectedPath}'.")
            End If
        End If

        Dim resultBuilder As New System.Text.StringBuilder()
        Dim useDocumentTags As Boolean = (filesToProcess.Count > 1)
        Dim documentCounter As Integer = 0

        For Each filePath In filesToProcess
            Dim result = Await LoadSingleContextFileAsync(filePath, interactive)
            Dim content = result.Content

            If result.PdfMayBeIncomplete Then
                pdfsWithPossibleImages.Add(filePath)
            End If

            If String.IsNullOrWhiteSpace(content) OrElse content.StartsWith("Error:", StringComparison.OrdinalIgnoreCase) Then
                failedFiles.Add(filePath)
                Continue For
            End If

            documentCounter += 1
            loadedFiles.Add(Tuple.Create(filePath, content.Length))

            If useDocumentTags Then
                Dim fileName = System.IO.Path.GetFileName(filePath)
                resultBuilder.Append($"<document{documentCounter} name=""{fileName}"">")
                resultBuilder.Append(content)
                resultBuilder.Append($"</document{documentCounter}>")
            Else
                resultBuilder.Append(content)
            End If
        Next

        Dim summary As New System.Text.StringBuilder()

        If loadedFiles.Count > 0 Then
            summary.AppendLine($"Successfully loaded ({loadedFiles.Count} file(s)):")
            Dim totalChars As Integer = 0
            For Each item In loadedFiles
                summary.AppendLine($"  • {System.IO.Path.GetFileName(item.Item1)} ({item.Item2:N0} chars)")
                totalChars += item.Item2
            Next
            summary.AppendLine($"  Total: {totalChars:N0} characters")
            summary.AppendLine()
        End If

        If failedFiles.Count > 0 Then
            summary.AppendLine($"Failed to load ({failedFiles.Count} item(s)):")
            For Each item In failedFiles
                summary.AppendLine($"  • {System.IO.Path.GetFileName(item)}")
            Next
            summary.AppendLine()
        End If

        If pdfsWithPossibleImages.Count > 0 Then
            summary.AppendLine($"PDFs that may contain images/scans ({pdfsWithPossibleImages.Count} file(s)):")
            For Each item In pdfsWithPossibleImages
                summary.AppendLine($"  • {System.IO.Path.GetFileName(item)}")
            Next
            summary.AppendLine("  (Text extraction may be incomplete because OCR was not performed)")
            summary.AppendLine()
        End If

        If ignoredCount > 0 Then
            summary.AppendLine($"Ignored unsupported files: {ignoredCount}")
            summary.AppendLine()
        End If

        Return (
            resultBuilder.ToString(),
            If(isFile, selectedPath, selectedPath & " (directory)"),
            loadedFiles.Count,
            summary.ToString().TrimEnd()
        )
    End Function

    Private Async Function RestoreLoadedContextAsync() As System.Threading.Tasks.Task
        ' Restore a previously loaded semantic index first (either a document OR an index is active).
        If Not String.IsNullOrWhiteSpace(_cachedLoadedIndexPath) OrElse Not String.IsNullOrWhiteSpace(_cachedLoadedIndexDisplayName) Then
            _loadedIndexSourcePath = _cachedLoadedIndexPath
            _loadedIndexDisplayName = If(_cachedLoadedIndexDisplayName, "(Persisted Index)")
            AppendSystemMessage("Index restored from cache.")
            Return
        End If

        Dim persistedIndexPath As String = GetPersistedIndexFilePath()

        Dim savedIndexPath As String = ""
        Try
            savedIndexPath = My.Settings.ChatContextPath
        Catch
        End Try

        Dim savedIndexPathIsIndex As Boolean =
            Not String.IsNullOrWhiteSpace(savedIndexPath) AndAlso
            System.IO.File.Exists(savedIndexPath) AndAlso
            SharedMethods.IsPotentiallySemanticSearchIndexedTextFile(savedIndexPath)

        If chkPersistContext.Checked AndAlso System.IO.File.Exists(persistedIndexPath) Then
            _loadedIndexSourcePath = If(savedIndexPathIsIndex, savedIndexPath, Nothing)
            _loadedIndexDisplayName =
                If(Not String.IsNullOrWhiteSpace(_loadedIndexSourcePath),
                   System.IO.Path.GetFileName(_loadedIndexSourcePath),
                   "(Persisted Index)")
            _cachedLoadedIndexPath = _loadedIndexSourcePath
            _cachedLoadedIndexDisplayName = _loadedIndexDisplayName
            AppendSystemMessage("Index restored from persisted storage.")
            Return
        End If

        If savedIndexPathIsIndex Then
            _loadedIndexSourcePath = savedIndexPath
            _loadedIndexDisplayName = System.IO.Path.GetFileName(savedIndexPath)
            _cachedLoadedIndexPath = _loadedIndexSourcePath
            _cachedLoadedIndexDisplayName = _loadedIndexDisplayName
            AppendSystemMessage($"Index restored from saved path: {_loadedIndexDisplayName}.")
            Return
        End If

        If Not String.IsNullOrWhiteSpace(_cachedLoadedContextContent) AndAlso Not String.IsNullOrWhiteSpace(_cachedLoadedContextPath) Then
            _loadedContextContent = _cachedLoadedContextContent
            _loadedContextPath = _cachedLoadedContextPath
            AppendSystemMessage("Context restored from cache.")
            Return
        End If

        If chkPersistContext.Checked Then
            Dim persistPath = GetPersistedContextFilePath()
            If System.IO.File.Exists(persistPath) Then
                Try
                    _loadedContextContent = System.IO.File.ReadAllText(persistPath, System.Text.Encoding.UTF8)
                    _loadedContextPath = "(Persisted Context)"
                    _cachedLoadedContextContent = _loadedContextContent
                    _cachedLoadedContextPath = _loadedContextPath
                    AppendSystemMessage($"Context restored from persisted storage ({_loadedContextContent.Length:N0} characters).")
                    Return
                Catch ex As System.Exception
                    AppendSystemMessage($"Failed to restore persisted context: {ex.Message}")
                End Try
            End If
        End If

        Dim savedPath As String = ""
        Try
            savedPath = My.Settings.ChatContextPath
        Catch
            Return
        End Try

        If String.IsNullOrWhiteSpace(savedPath) Then Return

        If Not System.IO.File.Exists(savedPath) AndAlso Not System.IO.Directory.Exists(savedPath) Then
            Try
                My.Settings.ChatContextPath = ""
                My.Settings.Save()
            Catch
            End Try
            Return
        End If

        Dim restored = Await LoadContextFromPathAsync(savedPath, False)
        If String.IsNullOrWhiteSpace(restored.CombinedContent) Then Return

        _loadedContextContent = restored.CombinedContent
        _loadedContextPath = restored.DisplayPath
        _cachedLoadedContextContent = _loadedContextContent
        _cachedLoadedContextPath = _loadedContextPath

        If chkPersistContext.Checked Then
            Try
                PersistLoadedContextToTempFile()
            Catch
            End Try
        End If

        AppendSystemMessage($"Context restored from saved path: {System.IO.Path.GetFileName(savedPath)}.")
    End Function

    Private Async Sub btnLoadContext_Click(sender As Object, e As EventArgs)
        If Not String.IsNullOrWhiteSpace(_loadedContextContent) OrElse HasLoadedIndex() Then
            RemoveLoadedContext()
        Else
            Await PromptForLoadedContextAsync()
        End If
    End Sub

    ''' <summary>
    ''' Updates the Load Context button caption to reflect whether external context is loaded.
    ''' </summary>
    Private Sub UpdateLoadContextButtonText()
        btnLoadContext.Text = If(String.IsNullOrWhiteSpace(_loadedContextContent) AndAlso Not HasLoadedIndex(), "Load Context", "Remove Context")
    End Sub

    ''' <summary>
    ''' Removes any loaded external context (in-memory, cache, persisted file, and saved path).
    ''' </summary>
    Private Sub RemoveLoadedContext()
        Dim hadIndex As Boolean = HasLoadedIndex()

        _loadedContextContent = Nothing
        _loadedContextPath = Nothing
        _cachedLoadedContextContent = Nothing
        _cachedLoadedContextPath = Nothing
        DeletePersistedContextFile(False)

        ' Removing the context also removes any loaded index and its persisted files.
        ClearLoadedIndexOnly()

        Try
            My.Settings.ChatContextPath = ""
            My.Settings.Save()
        Catch
        End Try

        AppendSystemMessage(If(hadIndex, "Loaded index removed.", "Loaded context removed."))
        UpdateLoadContextButtonText()
    End Sub

    Private Async Function PromptForLoadedContextAsync() As System.Threading.Tasks.Task
        Try
            Globals.ThisAddIn.DragDropFormLabel = "... a file or folder you want to use as external context, or click Browse"
            Globals.ThisAddIn.DragDropFormFilter = GetContextDragDropFilter()

            Dim selectedPath As String = ""

            Using frm As New DragDropForm(DragDropMode.FileOrDirectory)
                If frm.ShowDialog() = DialogResult.OK Then
                    selectedPath = frm.SelectedFilePath
                End If
            End Using

            Globals.ThisAddIn.DragDropFormLabel = ""
            Globals.ThisAddIn.DragDropFormFilter = ""

            If String.IsNullOrWhiteSpace(selectedPath) Then
                Return
            End If

            ' If the selected file is already a semantic-search index, attach it as an index source
            ' instead of inlining its bytes (the Word Chatbot uses either a document OR an index).
            If System.IO.File.Exists(selectedPath) AndAlso
               SharedMethods.IsPotentiallySemanticSearchIndexedTextFile(selectedPath) Then

                If AttachLoadedIndex(selectedPath) Then
                    If chkPersistContext.Checked Then
                        Try
                            PersistLoadedIndexToAppData()
                            AppendSystemMessage($"Index '{_loadedIndexDisplayName}' loaded and persisted. It will be searched for each message.")
                        Catch ex As System.Exception
                            AppendSystemMessage($"Index '{_loadedIndexDisplayName}' loaded but failed to persist: {ex.Message}")
                        End Try
                    Else
                        AppendSystemMessage($"Index '{_loadedIndexDisplayName}' loaded. It will be searched for each message.")
                        OfferIndexPersistence()
                    End If
                End If

                UpdateLoadContextButtonText()
                Return
            End If

            Dim loaded = Await LoadContextFromPathAsync(selectedPath, True)

            If String.IsNullOrWhiteSpace(loaded.Summary) AndAlso String.IsNullOrWhiteSpace(loaded.CombinedContent) Then
                Return
            End If

            If Not String.IsNullOrWhiteSpace(loaded.Summary) Then
                Dim proceedAnswer = ShowCustomYesNoBox(
                    loaded.Summary & vbCrLf & vbCrLf & "Do you want to use this context?",
                    "Yes, proceed",
                    "No, retry")

                If proceedAnswer <> 1 Then
                    Await PromptForLoadedContextAsync()
                    Return
                End If
            End If

            If String.IsNullOrWhiteSpace(loaded.CombinedContent) Then
                AppendSystemMessage("Failed to load context or all files are empty.")
                Return
            End If

            ' Loading a document context replaces any previously loaded index (either a document or an index).
            ClearLoadedIndexOnly()

            _loadedContextContent = loaded.CombinedContent
            _loadedContextPath = loaded.DisplayPath
            _cachedLoadedContextContent = _loadedContextContent
            _cachedLoadedContextPath = _loadedContextPath

            If chkPersistContext.Checked Then
                Try
                    PersistLoadedContextToTempFile()
                    AppendSystemMessage($"Context loaded and persisted ({_loadedContextContent.Length:N0} characters from {loaded.LoadedCount} file(s)).")
                Catch ex As System.Exception
                    AppendSystemMessage($"Context loaded ({_loadedContextContent.Length:N0} characters) but failed to persist: {ex.Message}")
                End Try
            Else
                AppendSystemMessage($"Context loaded: {loaded.LoadedCount} file(s), {_loadedContextContent.Length:N0} characters total.")
            End If

            Try
                My.Settings.ChatContextPath = selectedPath
                My.Settings.Save()
            Catch
            End Try

        Catch ex As System.Exception
            AppendSystemMessage($"Error loading context: {ex.Message}")
        Finally
            Globals.ThisAddIn.DragDropFormLabel = ""
            Globals.ThisAddIn.DragDropFormFilter = ""
            UpdateLoadContextButtonText()
        End Try
    End Function

    Private Sub chkPersistContext_CheckedChanged(sender As Object, e As EventArgs)
        If _isUpdatingPersistContextCheckbox Then Return

        ' When an index is loaded, persistence applies to the index file rather than inlined context.
        If HasLoadedIndex() Then
            Try
                HandleIndexPersistenceToggle()
            Catch ex As System.Exception
                AppendSystemMessage($"Error handling persist index setting: {ex.Message}")
            End Try
            Return
        End If

        Try
            Dim persistPath = GetPersistedContextFilePath()

            If chkPersistContext.Checked Then
                If Not String.IsNullOrWhiteSpace(_cachedLoadedContextContent) Then
                    System.IO.File.WriteAllText(persistPath, _cachedLoadedContextContent, System.Text.Encoding.UTF8)
                    AppendSystemMessage($"Context persisted to durable storage ({_cachedLoadedContextContent.Length:N0} characters).")
                Else
                    AppendSystemMessage("No context loaded to persist. Load context first, then check this box.")
                End If
            Else
                If System.IO.File.Exists(persistPath) Then
                    Dim answer = ShowCustomYesNoBox(
                        "Do you want to delete the persisted context file? This cannot be undone if you quit Word.",
                        "Yes, delete",
                        "No, keep it")

                    If answer = 1 Then
                        DeletePersistedContextFile(True)
                    Else
                        _isUpdatingPersistContextCheckbox = True
                        chkPersistContext.Checked = True
                        _isUpdatingPersistContextCheckbox = False
                        Return
                    End If
                End If
            End If

            My.Settings.ChatPersistContext = chkPersistContext.Checked
            My.Settings.Save()
            UpdatePersistContextTooltip()

        Catch ex As System.Exception
            AppendSystemMessage($"Error handling persist context setting: {ex.Message}")
        End Try
    End Sub

    ' =========================================================================
    ' Persistence
    ' =========================================================================

    ''' <summary>
    ''' Persists inner HTML of #chat container to My.Settings.LastChatHistoryHtml.
    ''' Called after each message append to preserve chat across sessions.
    ''' </summary>
    Private Sub PersistChatHtml()
        Try
            If wbChat Is Nothing OrElse wbChat.Document Is Nothing Then Return
            Dim chat = wbChat.Document.GetElementById("chat")
            If chat Is Nothing Then Return
            My.Settings.LastChatHistoryHtml = chat.InnerHtml
            My.Settings.Save()
        Catch
            ' Best-effort; ignore errors
        End Try
    End Sub

    ' =========================================================================
    ' Initialization
    ' =========================================================================

    ''' <summary>
    ''' Initializes WebBrowser control and adds to the SplitContainer's Panel1.
    ''' Called from constructor after txtChatHistory placement.
    ''' </summary>
    ''' <param name="host">TableLayoutPanel containing chat controls (unused but kept for API compat)</param>
    ''' <remarks>
    ''' Hides txtChatHistory (plain text fallback), adds wbChat to Panel1 of splitChat,
    ''' sets up BrowserBridge for JavaScript interaction, and wires event handlers.
    ''' </remarks>
    Public Sub InitChatHtmlUI(host As TableLayoutPanel)
        If host Is Nothing Then Return

        txtChatHistory.Visible = False
        splitChat.Panel1.Controls.Add(wbChat)
        wbChat.BringToFront()

        ' Expose COM bridge for JavaScript interaction
        wbChat.ObjectForScripting = New BrowserBridge()

        ' Wire navigation prevention handlers
        AddHandler wbChat.DocumentCompleted, AddressOf WbChat_DocumentCompleted
        AddHandler wbChat.Navigating, AddressOf WbChat_Navigating
        AddHandler wbChat.NewWindow, AddressOf WbChat_NewWindow
    End Sub

    ''' <summary>
    ''' Handles Navigating event to prevent internal navigation.
    ''' Cancels navigation and opens URL externally if http/https/mailto.
    ''' </summary>
    Private Sub WbChat_Navigating(sender As Object, e As WebBrowserNavigatingEventArgs)
        Try
            If e.Url IsNot Nothing Then
                Dim scheme = e.Url.Scheme.ToLowerInvariant()
                If scheme = "http" OrElse scheme = "https" OrElse scheme = "mailto" Then
                    e.Cancel = True
                    ' Launch outside the browser navigation event to avoid starting a process
                    ' from within the WebBrowser COM callback (a re-entrancy / crash risk).
                    Dim urlToOpen As String = e.Url.ToString()
                    Me.BeginInvoke(Sub()
                                       Try
                                           Process.Start(New ProcessStartInfo(urlToOpen) With {.UseShellExecute = True})
                                       Catch
                                           ' Silently ignore errors
                                       End Try
                                   End Sub)
                End If
            End If
        Catch
            ' Silently ignore errors
        End Try
    End Sub

    ''' <summary>
    ''' Handles NewWindow event (popup attempt) to prevent popups and open link externally.
    ''' </summary>
    Private Sub WbChat_NewWindow(sender As Object, e As CancelEventArgs)
        e.Cancel = True
        Try
            Dim doc = wbChat.Document
            If doc IsNot Nothing AndAlso doc.ActiveElement IsNot Nothing Then
                Dim href = doc.ActiveElement.GetAttribute("href")
                If Not String.IsNullOrWhiteSpace(href) Then
                    ' Launch outside the browser NewWindow event to avoid starting a process
                    ' from within the WebBrowser COM callback (a re-entrancy / crash risk).
                    Dim urlToOpen As String = href
                    Me.BeginInvoke(Sub()
                                       Try
                                           Process.Start(New ProcessStartInfo(urlToOpen) With {.UseShellExecute = True})
                                       Catch
                                           ' Silently ignore errors
                                       End Try
                                   End Sub)
                End If
            End If
        Catch
            ' Silently ignore errors
        End Try
    End Sub

    ''' <summary>
    ''' Initializes HTML document in WebBrowser with CSS styling and JavaScript utilities.
    ''' Called once during form load to set up empty chat container.
    ''' </summary>
    ''' <remarks>
    ''' Builds complete HTML document with:
    ''' - CSS: Segoe UI font, message styling, Markdown element formatting
    ''' - JavaScript: wireLinks() for link instrumentation, appendMessage() for adding chat items,
    '''   removeById() for removing "Thinking..." placeholder
    ''' - Empty #chat div container for messages
    ''' 
    ''' Font size calculated from form font + 1pt (min 10pt).
    ''' </remarks>
    Public Sub InitializeChatHtml()
        Dim baseSize As Single = If(Me IsNot Nothing AndAlso Me.Font IsNot Nothing, Me.Font.SizeInPoints, 9.0F)
        Dim fontPt As Single = System.Math.Max(baseSize + 1.0F, 10.0F)

        ' Build CSS stylesheet
        Dim css As String =
$"html,body{{height:100%;margin:0;padding:0;background:#fff;color:#000;}}
body{{font-family:'Segoe UI',Tahoma,Arial,sans-serif;font-size:{fontPt}pt;line-height:1.45;}}
#chat{{padding:6px 8px;}}
.msg{{margin:6px 0;word-wrap:break-word;}}
.msg:after{{content:'';display:block;clear:both;}}
.msg .who{{font-weight:600;margin-right:4px;float:left;}}
.msg .content{{display:block;overflow:hidden;}}
.msg.user .who{{color:#333;}}
.msg.assistant .who{{color:#003366;}}
.msg.thinking .content{{opacity:.75;font-style:italic;}}
/* No top/bottom gap when content is block-rendered */
.msg .content > *:first-child{{margin-top:0;}}
.msg .content > *:last-child{{margin-bottom:0;}}
a{{color:#0068c9;text-decoration:underline;cursor:pointer;}}
a:visited{{color:#5a3694;}}
ul,ol{{margin:6px 0 6px 22px;}}
pre,code,kbd,samp{{font-family:Consolas,'Courier New',monospace;}}
pre{{white-space:pre-wrap;background:#f6f8fa;border:1px solid #e1e4e8;border-radius:4px;padding:6px;}}
blockquote{{border-left:4px solid #e1e4e8;margin:6px 0;padding:6px 10px;background:#fafbfc;color:#333;}}
table{{border-collapse:collapse;margin:6px 0;}}
td,th{{border:1px solid #ddd;padding:4px 6px;}}"

        ' Build complete HTML document with JavaScript utilities
        Dim html As String =
$"<!DOCTYPE html>
<html>
<head>
<meta http-equiv=""X-UA-Compatible"" content=""IE=edge"" />
<meta charset=""utf-8"">
<style>{css}</style>
<script type=""text/javascript"">
function wireLinks(root) {{
  var links = root.getElementsByTagName('a');
  for (var i = 0; i < links.length; i++) {{
    (function(a) {{
      a.setAttribute('target', '_self');    // avoid NewWindow for old IE
      a.setAttribute('rel', 'noopener');
      a.onclick = function() {{
        try {{ if (window.external && window.external.OpenLink) window.external.OpenLink(a.href); }} catch (e) {{}}
        if (window.event) window.event.returnValue = false; // IE8-
        return false;
      }};
    }})(links[i]);
  }}
}}
function appendMessage(html) {{
  var c = document.getElementById('chat');
  if (!c) return;
  var temp = document.createElement('div');
  temp.innerHTML = html;
  wireLinks(temp);
  while (temp.firstChild) {{
    c.appendChild(temp.firstChild);
  }}
  window.scrollTo(0, document.body.scrollHeight);
}}
function removeById(id) {{
  var el = document.getElementById(id);
  if (!el || !el.parentNode) return;
  el.parentNode.removeChild(el);
}}
function setThinking(id, text) {{
  var el = document.getElementById(id);
  if (!el) return;
  var parts = el.getElementsByClassName ? el.getElementsByClassName('content') : null;
  if (parts && parts.length > 0) {{ parts[0].innerHTML = text; }}
  window.scrollTo(0, document.body.scrollHeight);
}}
</script>
</head>
<body>
  <div id=""chat""></div>
</body>
</html>"
        _htmlReady = False
        wbChat.DocumentText = html
    End Sub

    ''' <summary>
    ''' Clears all HTML chat content and reinitializes empty document.
    ''' </summary>
    Public Sub ClearChatHtml()
        _htmlQueue.Clear()
        _htmlReady = False
        InitializeChatHtml()
    End Sub

    ' =========================================================================
    ' HTML Utility Functions
    ' =========================================================================

    ''' <summary>
    ''' HTML-encodes plain text by escaping special characters.
    ''' </summary>
    ''' <param name="s">Text to encode</param>
    ''' <returns>HTML-safe text</returns>
    Private Shared Function HtmlEncode(s As String) As String
        If s Is Nothing Then Return ""
        Return s.Replace("&", "&amp;").
                 Replace("<", "&lt;").
                 Replace(">", "&gt;").
                 Replace("""", "&quot;")
    End Function

    ''' <summary>
    ''' Instruments anchor tags in HTML to open externally via BrowserBridge.
    ''' </summary>
    ''' <param name="html">HTML fragment potentially containing anchor tags</param>
    ''' <returns>HTML with instrumented links</returns>
    ''' <remarks>
    ''' Uses regex to find anchor tags and add:
    ''' - onclick handler calling window.external.OpenLink(href)
    ''' - target="_self" to avoid popup behavior in IE
    ''' - return false to prevent default navigation
    ''' 
    ''' Skips links already instrumented (contains "OpenLink").
    ''' </remarks>
    Private Shared Function InstrumentLinks(html As String) As String
        If String.IsNullOrEmpty(html) Then Return html
        Try
            Return System.Text.RegularExpressions.Regex.Replace(
                html,
                "(?is)<a\s+([^>]*?)\bhref\s*=\s*(?:'([^']*)'|""([^""]*)""|([^\s>]+))([^>]*)>",
                Function(m As System.Text.RegularExpressions.Match)
                    Dim pre = m.Groups(1).Value
                    Dim href = If(m.Groups(2).Success, m.Groups(2).Value, If(m.Groups(3).Success, m.Groups(3).Value, m.Groups(4).Value))
                    Dim post = m.Groups(5).Value
                    If String.IsNullOrWhiteSpace(href) Then Return m.Value
                    ' Skip if already instrumented
                    If m.Value.IndexOf("OpenLink", StringComparison.OrdinalIgnoreCase) >= 0 Then Return m.Value
                    Dim safeHref = href.Replace("""", "&quot;")
                    Dim onclickAttr = " onclick=""try{if(window.external&&window.external.OpenLink)window.external.OpenLink(this.href);}catch(e){};return false;"""
                    Dim targetAttr = If(m.Value.IndexOf("target=", StringComparison.OrdinalIgnoreCase) >= 0, "", " target=""_self""")
                    Return $"<a {pre} href=""{safeHref}""{targetAttr}{onclickAttr}{post}>"
                End Function)
        Catch
            Return html
        End Try
    End Function

    ' =========================================================================
    ' Message Appending Functions
    ' =========================================================================

    ''' <summary>
    ''' Converts plain text transcript to HTML and appends to chat display.
    ''' Parses "You:" and "{AN5}:" prefixes to determine message roles.
    ''' </summary>
    ''' <param name="transcript">Plain text chat transcript</param>
    ''' <remarks>
    ''' Processing:
    ''' 1. Splits transcript into lines (normalized to vbLf)
    ''' 2. Detects role changes via "You:" or "Inky:" line prefixes
    ''' 3. Accumulates content until role change
    ''' 4. Flushes accumulated content as HTML message div
    ''' 
    ''' User messages: HTML-encoded plain text with &lt;br&gt; for line breaks
    ''' Assistant messages: Markdown-to-HTML via Markdig with link instrumentation
    ''' Single-paragraph assistant messages inlined as &lt;span&gt; instead of &lt;div&gt;
    ''' </remarks>
    Public Sub AppendTranscriptToHtml(transcript As String)
        If String.IsNullOrEmpty(transcript) Then Return

        Dim lines = transcript.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split(New String() {vbLf}, StringSplitOptions.None)
        Dim currentRole As String = Nothing
        Dim content As New System.Text.StringBuilder()

        ' Flush accumulated content as HTML message
        Dim SubFlush As System.Action =
            Sub()
                If content.Length = 0 OrElse String.IsNullOrEmpty(currentRole) Then
                    content.Clear() : currentRole = Nothing : Return
                End If
                Dim htmlFrag As String
                If currentRole = "user" Then
                    ' User message: plain HTML-encoded text
                    Dim encoded = HtmlEncode(content.ToString()).Replace(vbLf, "<br>")
                    htmlFrag = $"<div class='msg user'><span class='who'>You:</span><span class='content'>{encoded}</span></div>"
                Else
                    ' Assistant message: convert Markdown to HTML
                    Dim md = RemoveCommands(content.ToString())
                    Dim body = Markdown.ToHtml(md, _mdPipeline)
                    body = InstrumentLinks(body)
                    Dim t = If(body, "").Trim()

                    ' Detect single-paragraph responses for inline rendering
                    Dim isSingleParagraph As Boolean =
                        System.Text.RegularExpressions.Regex.IsMatch(t, "^\s*<p>[\s\S]*?</p>\s*$", RegexOptions.IgnoreCase) AndAlso
                        Not System.Text.RegularExpressions.Regex.IsMatch(t, "<(ul|ol|pre|table|h[1-6]|blockquote|hr|div)\b", RegexOptions.IgnoreCase)

                    If isSingleParagraph Then
                        ' Inline as span (no extra vertical spacing)
                        Dim inlineHtml As String = System.Text.RegularExpressions.Regex.Replace(t, "^\s*<p>|</p>\s*$", "", RegexOptions.IgnoreCase)
                        htmlFrag = $"<div class='msg assistant'><span class='who'>{HtmlEncode(AN5)}:</span><span class='content'>{inlineHtml}</span></div>"
                    Else
                        ' Block rendering for multi-element responses
                        htmlFrag = $"<div class='msg assistant'><span class='who'>{HtmlEncode(AN5)}:</span><div class='content'>{body}</div></div>"
                    End If
                End If
                AppendHtml(htmlFrag)
                content.Clear()
                currentRole = Nothing
            End Sub

        ' Parse lines and accumulate content by role
        For Each ln In lines
            If ln.StartsWith("You:", StringComparison.OrdinalIgnoreCase) Then
                SubFlush()
                currentRole = "user"
                content.Append(ln.Substring(4).TrimStart())
            ElseIf ln.StartsWith(AN5 & ":", StringComparison.OrdinalIgnoreCase) Then
                SubFlush()
                currentRole = "assistant"
                content.Append(ln.Substring((AN5 & ":").Length).TrimStart())
            Else
                If content.Length > 0 Then content.AppendLine()
                content.Append(ln)
            End If
        Next
        SubFlush()
        PersistChatHtml()
    End Sub

    ''' <summary>
    ''' Appends user message as HTML-encoded plain text (no Markdown processing).
    ''' </summary>
    ''' <param name="text">User message text</param>
    Public Sub AppendUserHtml(text As String)
        Dim encoded = HtmlEncode(text).
                      Replace(vbCrLf, "<br>").
                      Replace(vbLf, "<br>").
                      Replace(vbCr, "<br>")
        AppendHtml($"<div class='msg user'><span class='who'>You:</span><span class='content'>{encoded}</span></div>")
        PersistChatHtml()
    End Sub

    ''' <summary>
    ''' Shows "Thinking..." placeholder while waiting for LLM response.
    ''' Generates unique DOM ID for later removal.
    ''' </summary>
    ''' <param name="isTooling">When True, displays tooling-specific message.</param>
    Public Sub ShowAssistantThinking(Optional isTooling As Boolean = False)
        _lastThinkingId = "thinking-" & Guid.NewGuid().ToString("N")
        Dim thinkingText As String = If(isTooling,
            $"Thinking (using {Globals.ThisAddIn.ToolFriendlyName.ToLower})...",
            "Thinking...")
        AppendHtml($"<div id=""{_lastThinkingId}"" class='msg assistant thinking'><span class='who'>{HtmlEncode(AN5)}:</span><span class='content'>{thinkingText}</span></div>")
    End Sub

    ''' <summary>
    ''' Removes "Thinking..." placeholder from DOM after LLM responds.
    ''' Uses JavaScript removeById() function.
    ''' </summary>
    Public Sub RemoveAssistantThinking()
        If String.IsNullOrEmpty(_lastThinkingId) Then Return
        Try
            If wbChat.Document IsNot Nothing Then
                wbChat.Document.InvokeScript("removeById", New Object() {_lastThinkingId})
            End If
        Catch
            ' Best-effort; ignore errors
        Finally
            _lastThinkingId = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Updates the text of the current "Thinking..." placeholder to show retrieval progress
    ''' (used while the loaded semantic index is being searched).
    ''' </summary>
    Public Sub UpdateAssistantThinking(statusText As String)
        If String.IsNullOrEmpty(_lastThinkingId) Then Return
        Try
            If wbChat IsNot Nothing AndAlso wbChat.Document IsNot Nothing Then
                wbChat.Document.InvokeScript("setThinking", New Object() {_lastThinkingId, HtmlEncode(statusText)})
            End If
        Catch
            ' Best-effort; ignore errors
        End Try
    End Sub

    ''' <summary>
    ''' Appends assistant message by converting Markdown to HTML using Markdig.
    ''' Detects single-paragraph responses for inline rendering optimization.
    ''' </summary>
    ''' <param name="md">Markdown text from LLM response</param>
    ''' <remarks>
    ''' Single-paragraph detection prevents unnecessary vertical spacing for short responses.
    ''' Checks for absence of block-level elements (ul, ol, pre, table, headings, blockquote, hr, div).
    ''' </remarks>
    Public Sub AppendAssistantMarkdown(md As String)
        If md Is Nothing Then md = ""
        Dim body As String = Markdown.ToHtml(md, _mdPipeline)
        body = InstrumentLinks(body)
        Dim t As String = If(body, "").Trim()

        ' Detect single-paragraph response
        Dim isSingleParagraph As Boolean =
            System.Text.RegularExpressions.Regex.IsMatch(t, "^\s*<p>[\s\S]*?</p>\s*$", RegexOptions.IgnoreCase) AndAlso
            Not System.Text.RegularExpressions.Regex.IsMatch(t, "<(ul|ol|pre|table|h[1-6]|blockquote|hr|div)\b", RegexOptions.IgnoreCase)

        If isSingleParagraph Then
            ' Inline rendering (strip <p> tags)
            Dim inlineHtml As String = System.Text.RegularExpressions.Regex.Replace(t, "^\s*<p>|</p>\s*$", "", RegexOptions.IgnoreCase)
            AppendHtml($"<div class='msg assistant'><span class='who'>{HtmlEncode(AN5)}:</span><span class='content'>{inlineHtml}</span></div>")
        Else
            ' Block rendering
            AppendHtml($"<div class='msg assistant'><span class='who'>{HtmlEncode(AN5)}:</span><div class='content'>{body}</div></div>")
        End If

        PersistChatHtml()
    End Sub

    ''' <summary>
    ''' Appends HTML fragment to chat display. Queues if WebBrowser not ready.
    ''' </summary>
    ''' <param name="fragment">HTML fragment to append</param>
    ''' <remarks>
    ''' If _htmlReady=False, adds to _htmlQueue for later flushing in WbChat_DocumentCompleted.
    ''' Uses JavaScript appendMessage() function to add to #chat container and scroll.
    ''' </remarks>
    Private Sub AppendHtml(fragment As String)
        If String.IsNullOrEmpty(fragment) Then Return

        ' Queue if WebBrowser not ready
        If Not _htmlReady OrElse wbChat.Document Is Nothing Then
            _htmlQueue.Add(fragment)
            Return
        End If

        Try
            wbChat.Document.InvokeScript("appendMessage", New Object() {fragment})
        Catch
            ' Timing edge: queue and wait for next ready cycle
            _htmlQueue.Add(fragment)
        End Try
    End Sub

    ''' <summary>
    ''' Handles DocumentCompleted event to flush queued HTML fragments.
    ''' Wires document click handler for link opening.
    ''' </summary>
    Private Sub WbChat_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs)
        _htmlReady = True

        WireDocumentClick()

        ' Flush any queued messages
        If _htmlQueue.Count > 0 Then
            Try
                For Each frag In _htmlQueue
                    wbChat.Document.InvokeScript("appendMessage", New Object() {frag})
                Next
            Catch
                ' Ignore errors during flush
            Finally
                _htmlQueue.Clear()
            End Try
        End If
    End Sub

End Class
