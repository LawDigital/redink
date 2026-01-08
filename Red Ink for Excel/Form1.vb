' Part of "Red Ink for Excel"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: Form1.vb (frmAIChat)
' Part of: Red Ink for Excel
' Purpose: Provides a Windows Forms chat interface to an LLM for an Excel add-in.
'          Supports sending user prompts, including worksheet content or a cell
'          selection, persisting chat history, switching between two models,
'          and executing structured commands returned by the LLM against the
'          active Excel workbook.
'
' Architecture:
' - UI Composition: Dynamically builds a TableLayoutPanel containing an instructions
'   label, chat history textbox, user input textbox, a checkbox panel, and a button panel.
'   Two FlowLayoutPanels host action buttons and option checkboxes.
' - State Persistence: Uses My.Settings to store last chat history snippet (bounded by
'   INI_ChatCap), window size/location, and user option selections.
' - Conversation Handling: Maintains an in‑memory List of (Role, Content) tuples for
'   current session; previous persisted chat appended once when sending first message.
' - Prompt Construction: Builds a system prompt from SharedContext plus optional worksheet
'   (entire UsedRange) or selection data. Additional worksheets may be injected via a trigger.
' - Worksheet Access: Uses Microsoft.Office.Interop.Excel to read ActiveWorkbook / ActiveSheet,
'   UsedRange, and current Selection; verifies availability before including content.
' - Command Execution: Parses embedded command blocks in LLM responses and applies resulting
'   instructions to Excel via Globals.ThisAddIn.ApplyLLMInstructions; supports undo state tracking.
' - Model Switching: Toggles between two configured models (INI_Model / INI_Model_2) if second
'   API is enabled (INI_SecondAPI).
' - Asynchronous Flow: Uses async/await for LLM calls; inserts a temporary “Thinking...” line
'   and replaces it with the final response.
' - Formatting: Normalizes bullet points, strips Markdown emphasis, and optionally executes
'   commands if write access is granted.
' - UI Thread Safety: Marshals updates through UpdateUIAsync / Invoke checks.
' - Clipboard Support: Copy entire chat or last assistant answer to clipboard.
' - Lifecycle: Restores previous chat, saves state on close, and sends a time‑aware welcome
'   message on first load if history is empty.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.Drawing
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports System.Text.RegularExpressions
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel
Imports System.Globalization
Imports Microsoft.Office.Core

''' <summary>
''' Chat form integrating an LLM with Excel context (worksheet / selection) and optional command execution.
''' </summary>
Public Class frmAIChat

    ''' <summary>Brings the specified window handle to the foreground via user32.dll.</summary>
    <DllImport("user32.dll")>
    Private Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
    End Function

    ''' <summary>Application name.</summary>
    Const AN As String = "Red Ink"
    ''' <summary>Alias used for chatbot display.</summary>
    Const AN5 As String = "Inky"   ' for Chatbox
    ''' <summary>Trigger token to include additional worksheets.</summary>
    Private Const ExtWSTrigger As String = "(addws)"

    ''' <summary>Tracks whether a newline precedes next appended user text.</summary>
    Private PreceedingNewline As String = ""
    ''' <summary>Persisted prior chat restored once at start of session.</summary>
    Private OldChat As String = ""
    ''' <summary>User interface culture name derived from Excel UI language.</summary>
    Private UserLanguage As String = New CultureInfo(Globals.ThisAddIn.Application.LanguageSettings.LanguageID(MsoAppLanguageID.msoLanguageIDUI)).Name
    ''' <summary>System (instruction) prompt sent with each LLM call.</summary>
    Private SystemPrompt As String = ""

    ''' <summary>Button: copy entire chat history.</summary>
    Private WithEvents btnCopy As New System.Windows.Forms.Button() With {.Text = "Copy All", .AutoSize = True}
    ''' <summary>Button: copy last assistant answer.</summary>
    Private WithEvents btnCopyLastAnswer As New System.Windows.Forms.Button() With {.Text = "Copy Last Answer", .AutoSize = True}
    ''' <summary>Button: clear current conversation.</summary>
    Private WithEvents btnClear As New System.Windows.Forms.Button() With {.Text = "Clear", .AutoSize = True}
    ''' <summary>Button: close form.</summary>
    Private WithEvents btnExit As New System.Windows.Forms.Button() With {.Text = "Close", .AutoSize = True}
    ''' <summary>Button: send prompt to LLM.</summary>
    Private WithEvents btnSend As New System.Windows.Forms.Button() With {.Text = "Send", .AutoSize = True}
    ''' <summary>Button: toggle between two configured models.</summary>
    Private WithEvents btnSwitchModel As New System.Windows.Forms.Button() With {.Text = "Switch Model", .AutoSize = True}
    ''' <summary>Checkbox: include entire worksheet UsedRange.</summary>
    Private WithEvents chkIncludeDocText As New System.Windows.Forms.CheckBox() With {.Text = "Include worksheet", .AutoSize = True, .Checked = My.Settings.IncludeDocument}
    ''' <summary>Checkbox: include only current selection (if not including entire sheet).</summary>
    Private WithEvents chkIncludeselection As New System.Windows.Forms.CheckBox() With {.Text = "Include selection", .AutoSize = True, .Checked = If(My.Settings.IncludeDocument, False, My.Settings.IncludeSelection)}
    ''' <summary>Checkbox: permit execution of commands returned by LLM.</summary>
    Private WithEvents chkPermitCommands As New System.Windows.Forms.CheckBox() With {.Text = "Grant write access", .AutoSize = True, .Checked = My.Settings.DoCommands}
    ''' <summary>Checkbox: control TopMost behavior.</summary>
    Private WithEvents chkStayOnTop As New System.Windows.Forms.CheckBox() With {.Text = "Do not stay on top", .AutoSize = True, .Checked = My.Settings.NotAlwaysOnTop}

    ''' <summary>Panel hosting action buttons.</summary>
    Dim pnlButtons As New FlowLayoutPanel() With {
        .Dock = DockStyle.Bottom,
        .FlowDirection = FlowDirection.LeftToRight,
        .AutoSize = True,
        .AutoSizeMode = AutoSizeMode.GrowAndShrink,
        .Height = 40
    }

    ''' <summary>Panel hosting option checkboxes.</summary>
    Dim pnlCheckboxes As New FlowLayoutPanel() With {
        .Dock = DockStyle.Bottom,
        .FlowDirection = FlowDirection.LeftToRight,
        .AutoSize = True,
        .AutoSizeMode = AutoSizeMode.GrowAndShrink,
        .Height = 40
    }

    ''' <summary>Shared context providing configuration and prompts.</summary>
    Private _context As ISharedContext = New SharedContext()

    ''' <summary>Indicates whether second model/API is active.</summary>
    Private _useSecondApi As Boolean = False

    ''' <summary>Current session chat messages (role, content).</summary>
    Private _chatHistory As New List(Of (Role As String, Content As String))

    ''' <summary>
    ''' Initializes the chat form, constructing dynamic layout and binding provided context.
    ''' </summary>
    ''' <param name="context">Shared context instance used for prompt and settings.</param>
    Public Sub New(context As ISharedContext)
        ' This call is required by the designer.
        InitializeComponent()

        Me.AutoSize = False

        txtChatHistory.Multiline = True
        txtUserInput.Multiline = True

        ' 1) Create TableLayoutPanel
        Dim mainLayout As New TableLayoutPanel() With {
            .ColumnCount = 1,
            .RowCount = 5,
            .Dock = DockStyle.Fill,
            .AutoSize = False,
            .Padding = New Padding(10)   ' will be overridden
        }

        ' 2) Set column width to 100%
        mainLayout.ColumnStyles.Clear()
        mainLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))

        ' 3) Right inner padding 20 px
        mainLayout.Padding = New Padding(left:=10, top:=10, right:=20, bottom:=10)

        ' 4) Define rows
        mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        mainLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))

        ' 5) Configure controls
        lblInstructions.AutoSize = True
        lblInstructions.Dock = DockStyle.Top
        txtChatHistory.Dock = DockStyle.Fill
        txtUserInput.Dock = DockStyle.Fill

        ' 6) Add controls to table
        mainLayout.Controls.Add(lblInstructions, 0, 0)
        mainLayout.Controls.Add(txtChatHistory, 0, 1)
        mainLayout.Controls.Add(txtUserInput, 0, 2)
        mainLayout.Controls.Add(pnlCheckboxes, 0, 3)
        mainLayout.Controls.Add(pnlButtons, 0, 4)

        ' 7) Refill form
        Me.Controls.Clear()
        Me.Controls.Add(mainLayout)

        _context = context
    End Sub

    ''' <summary>
    ''' Handles form load: restores state, initializes UI, attaches handlers, sends welcome if no history.
    ''' </summary>
    Private Async Sub frmAIChat_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.StartPosition = FormStartPosition.Manual
        Me.KeyPreview = True

        ' Restore saved chat text from My.Settings
        Dim previousChat As String = My.Settings.LastChatHistory
        If Not String.IsNullOrEmpty(previousChat) Then
            txtChatHistory.Text = previousChat
            OldChat = previousChat
            PreceedingNewline = Environment.NewLine
        End If

        ' Set the form's title and custom icon
        Me.Text = $"Chat (using " & If(_useSecondApi, _context.INI_Model_2, _context.INI_Model) & ")"
        Me.Font = New System.Drawing.Font("Segoe UI", 9)
        Me.FormBorderStyle = FormBorderStyle.Sizable ' Ensure border supports icons
        Me.Icon = Icon.FromHandle(New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard)).GetHicon())
        Me.TopMost = True ' Always on top

        ' Set the initial and minimum size of the form
        Me.MinimumSize = New Size(830, 521)

        If My.Settings.FormLocation <> System.Drawing.Point.Empty AndAlso My.Settings.FormSize <> Size.Empty Then
            Me.Location = My.Settings.FormLocation
            Me.Size = My.Settings.FormSize
        Else
            Me.StartPosition = FormStartPosition.CenterScreen
        End If

        AddHandler txtUserInput.KeyDown, AddressOf UserInput_KeyDown

        ' Set up instructions label
        lblInstructions.Text = $"Enter your question and click 'Send' or press Enter. Add '{ExtWSTrigger}' to pass along other open worksheets in your question. You can allow the chatbot to perform actions on your worksheet (change or comment cells): you can undo the last action, if needed."
        lblInstructions.AutoSize = True
        lblInstructions.Height = 50
        lblInstructions.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        lblInstructions.TextAlign = ContentAlignment.MiddleLeft

        ' FlowLayoutPanel for buttons
        pnlButtons.Padding = New Padding(0, 2, 8, 12)
        pnlButtons.Controls.Add(btnSend)
        pnlButtons.Controls.Add(btnCopyLastAnswer)
        pnlButtons.Controls.Add(btnCopy)
        pnlButtons.Controls.Add(btnClear)
        If _context.INI_SecondAPI Then pnlButtons.Controls.Add(btnSwitchModel)
        pnlButtons.Controls.Add(btnExit)

        pnlCheckboxes.Padding = New Padding(0, 1, 8, 1)
        pnlCheckboxes.Controls.Add(chkIncludeselection)
        pnlCheckboxes.Controls.Add(chkIncludeDocText)
        pnlCheckboxes.Controls.Add(chkPermitCommands)
        pnlCheckboxes.Controls.Add(chkStayOnTop)

        AddHandler btnCopy.Click, AddressOf btnCopy_Click
        AddHandler btnClear.Click, AddressOf btnClear_Click
        AddHandler btnSend.Click, AddressOf btnSend_Click
        AddHandler btnCopyLastAnswer.Click, AddressOf btnCopyLastAnswer_Click
        AddHandler btnSwitchModel.Click, AddressOf btnSwitchModel_Click
        AddHandler btnExit.Click, AddressOf btnExit_Click
        AddHandler chkIncludeselection.Click, AddressOf chkIncludeselection_Click
        AddHandler chkIncludeDocText.Click, AddressOf chkIncludeDocText_Click
        AddHandler chkPermitCommands.Click, AddressOf chkPermitCommands_Click
        AddHandler chkStayOnTop.Click, AddressOf chkStayontop_Click

        If String.IsNullOrWhiteSpace(txtChatHistory.Text) Then
            Dim result = Await WelcomeMessage()
        Else
            txtChatHistory.SelectionStart = txtChatHistory.Text.Length
            txtChatHistory.ScrollToCaret()
        End If

        If Globals.ThisAddIn.SizeOfWorksheet() > Globals.ThisAddIn.LargeWorksheetSize And chkIncludeDocText.Checked Then
            ShowCustomMessageBox($"Because this worksheet is large (a range of {Globals.ThisAddIn.SizeOfWorksheet()} cells, even if not all are used), it may slow down your interaction with the chatbot, because each time you send a question, the entire worksheet will be passed to {AN5}. If you want to speed up, include only a selection only.")
        End If

        If String.IsNullOrEmpty(txtUserInput.Text) Then txtUserInput.Focus()

    End Sub

    ''' <summary>
    ''' Sends user prompt: builds context, includes worksheet/selection, calls LLM, appends response, executes commands if allowed.
    ''' </summary>
    Private Async Sub btnSend_Click(sender As Object, e As EventArgs)
        Dim userPrompt As String = txtUserInput.Text.Trim()
        If userPrompt = "" Then Return

        Try
            ' Build entire conversation so far into one string for context
            SystemPrompt = _context.SP_ChatExcel().Replace("{UserLanguage}", UserLanguage).
                Replace("{Location}", ThisAddIn.Location) &
                $" Your name is '{AN5}'. The current date and time is: {DateTime.Now.ToString("MMMM dd, yyyy hh:mm tt")}. " &
                If(chkIncludeDocText.Checked, vbLf & "You have access to the user's document. " & vbLf, "") &
                If(chkIncludeselection.Checked, vbLf & "You have access to a selection of user's document. " & vbLf & " ", "") &
                If(My.Settings.DoCommands, _context.SP_Add_ChatExcel_Commands, _context.SP_Add_Chat_NoCommands)

            Dim conversationSoFar As String = BuildConversationString(_chatHistory)
            If Not String.IsNullOrWhiteSpace(OldChat) Then
                conversationSoFar += vbLf & OldChat
                OldChat = ""
            End If

            Dim appGuard As Microsoft.Office.Interop.Excel.Application = Globals.ThisAddIn.Application
            If (chkIncludeDocText.Checked Or chkIncludeselection.Checked) AndAlso
               (appGuard Is Nothing _
               OrElse appGuard.Workbooks Is Nothing _
               OrElse appGuard.Workbooks.Count = 0 _
               OrElse appGuard.ActiveWorkbook Is Nothing _
               OrElse appGuard.ActiveSheet Is Nothing) Then

                ShowCustomMessageBox("There is no active Excel worksheet. Please open or activate a workbook, then try again.")
                Return
            End If

            ' Optionally include Excel worksheet cells or selection
            Dim docText As String = ""
            Dim selectiontext As String = ""
            Dim selectedcells As String = ""
            Dim InsertWS As String = ""

            If chkIncludeDocText.Checked Then
                Dim ws As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
                Dim usedRange As Excel.Range = ws.UsedRange
                docText = Globals.ThisAddIn.ConvertRangeToString(usedRange, True)
            End If

            ' Always determine selection or active cell content/address
            Dim appx As Excel.Application = Globals.ThisAddIn.Application
            Dim used As Excel.Range = appx.ActiveSheet.UsedRange
            Dim selection As Excel.Range = TryCast(appx.Selection, Excel.Range)
            Dim intersectedRange As Excel.Range = Nothing

            If selection IsNot Nothing Then
                Try
                    intersectedRange = appx.Intersect(selection, used)
                Catch
                    intersectedRange = Nothing
                End Try
            End If

            If intersectedRange IsNot Nothing AndAlso intersectedRange.Cells.Count > 0 Then
                ' Non-empty selection: include its content
                selectiontext = Globals.ThisAddIn.ConvertRangeToString(intersectedRange, True, True)
                selectedcells = intersectedRange.Address(False, False)
            Else
                ' No selection or empty selection: fall back to active cell (if within UsedRange)
                Dim activeCell As Excel.Range = appx.ActiveCell
                Dim activeInUsed As Excel.Range = Nothing
                Try
                    activeInUsed = appx.Intersect(activeCell, used)
                Catch
                    activeInUsed = Nothing
                End Try
                If activeInUsed IsNot Nothing AndAlso activeInUsed.Cells.Count > 0 Then
                    selectiontext = Globals.ThisAddIn.ConvertRangeToString(activeInUsed, True, True)
                    selectedcells = activeInUsed.Address(False, False)
                Else
                    ' Active cell outside UsedRange: still report its address and value
                    selectiontext = Globals.ThisAddIn.ConvertRangeToString(activeCell, True, True)
                    selectedcells = activeCell.Address(False, False)
                End If
            End If

            If Not String.IsNullOrEmpty(userPrompt) AndAlso userPrompt.IndexOf(ExtWSTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                If Not chkIncludeDocText.Checked AndAlso String.IsNullOrEmpty(selectiontext) Then
                    ShowCustomMessageBox("You cannot use the " & ExtWSTrigger & " trigger if you do not includ the worksheet or a selection of it - trigger ignored.")
                    InsertWS = ""
                Else
                    InsertWS = Globals.ThisAddIn.GatherSelectedWorksheets()
                    Debug.WriteLine($"GatherSelectedWorksheets returned: {Microsoft.VisualBasic.Left(InsertWS, 3000)}")
                    If String.IsNullOrWhiteSpace(InsertWS) Then
                        ShowCustomMessageBox("No content was found or an error occurred in gathering the additional worksheet(s) - doing without them.")
                        InsertWS = ""
                    ElseIf InsertWS.StartsWith("ERROR", StringComparison.OrdinalIgnoreCase) Then
                        ShowCustomMessageBox($"An error occured gathering the additional worksheet(s) ({InsertWS.Substring(6).Trim()}) - doing without them.")
                        InsertWS = ""
                    ElseIf InsertWS.StartsWith("NONE", StringComparison.OrdinalIgnoreCase) Then
                        ShowCustomMessageBox($"There are no other worksheets to add - doing without them.")
                        InsertWS = ""
                    End If
                End If
                userPrompt = Regex.Replace(userPrompt, Regex.Escape(ExtWSTrigger), "", RegexOptions.IgnoreCase)
            End If

            ' Construct the full prompt
            Dim fullPrompt As New StringBuilder()

            Dim app As Excel.Application = Globals.ThisAddIn.Application
            Dim workbookName As String = app.ActiveWorkbook.Name
            Dim worksheetName As String = app.ActiveSheet.Name
            Dim combinedName As String = workbookName & " - " & worksheetName

            If Not String.IsNullOrEmpty(docText) Then
                fullPrompt.AppendLine("You have access to the user's worksheet. The user's current worksheet is '" & combinedName & "' and has the following content: <RANGEOFCELLS>" & docText & "</RANGEOFCELLS>")
            ElseIf Not String.IsNullOrEmpty(selectiontext) Then
                fullPrompt.AppendLine("You have access to the user's worksheet. The user's current worksheet is '" & combinedName & "'.")
            ElseIf chkIncludeselection.Checked Then
                fullPrompt.AppendLine("The user has granted you access to a selection of the worksheet '" & combinedName & "' but it is empty.")
            ElseIf chkIncludeDocText.Checked Then
                fullPrompt.AppendLine("The user has granted you access to the worksheet '" & combinedName & "' but the entire worksheet is empty.")
            End If

            ' Always include where the user stands or what is selected
            If Not String.IsNullOrEmpty(selectedcells) Then
                fullPrompt.AppendLine("The user is focused on or has selected the following cells: " & selectedcells)
            End If
            If Not String.IsNullOrEmpty(selectiontext) Then
                fullPrompt.AppendLine("Focused/selected cells content: <RANGEOFCELLS>" & selectiontext & "</RANGEOFCELLS>")
            End If

            If Not InsertWS.IsNullOrWhiteSpace(InsertWS) Then
                fullPrompt.AppendLine("The user also provided you access to the following additional worksheet(s): " & InsertWS)
            End If

            fullPrompt.AppendLine("User: " & userPrompt)
            fullPrompt.AppendLine("The conversation so far (not including any previously added worksheet content):" & vbLf & conversationSoFar)

            ' Update UI on the UI thread
            Await UpdateUIAsync(Sub()
                                    AppendToChatHistory(PreceedingNewline & "You: " & userPrompt.TrimEnd() & Environment.NewLine & Environment.NewLine)
                                    txtUserInput.Clear()
                                    PreceedingNewline = Environment.NewLine
                                End Sub)

            _chatHistory.Add(("user", userPrompt.TrimEnd()))

            ' Add a placeholder for AI response while waiting
            Await UpdateUIAsync(Sub()
                                    AppendToChatHistory($"{AN5}: Thinking...")
                                End Sub)

            ' Call the LLM function asynchronously
            Dim aiResponse As String = Await SharedMethods.LLM(_context, SystemPrompt, fullPrompt.ToString(), "", "", 0, _useSecondApi, True)
            aiResponse = aiResponse.TrimEnd()
            aiResponse = aiResponse.Replace($"{vbCrLf}* ", vbCrLf & ChrW(8226) & " ").Replace($"{vbCr}* ", vbCr & ChrW(8226) & " ").Replace($"{vbLf}* ", vbLf & ChrW(8226) & " ")
            aiResponse = aiResponse.Replace($"  *  ", "  " & ChrW(8226) & "  ")
            aiResponse = RemoveMarkdownFormatting(aiResponse)

            Dim CommandsString As String = ""
            If My.Settings.DoCommands Then
                CommandsString = aiResponse
            End If

            Await UpdateUIAsync(Sub()
                                    RemoveLastLineFromChatHistory()
                                    AppendToChatHistory(Environment.NewLine & $"{AN5}: " & aiResponse.TrimEnd().Replace(vbCrLf, Environment.NewLine).Replace(vbLf, Environment.NewLine) & Environment.NewLine)
                                    If My.Settings.DoCommands And Not String.IsNullOrWhiteSpace(CommandsString) Then
                                        ExecuteAnyCommands(CommandsString)
                                    End If
                                    txtUserInput.Text = ""
                                    If String.IsNullOrEmpty(txtUserInput.Text) Then txtUserInput.Focus()
                                End Sub)

            _chatHistory.Add(("assistant", aiResponse.TrimEnd()))

        Catch ex As System.Exception
            MsgBox("Error in btnSend_Click: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    ''' <summary>
    ''' Sends a localized welcome message referencing current time of day.
    ''' </summary>
    Private Async Function WelcomeMessage() As Task(Of String)
        Try
            ' Build entire conversation so far into one string for context
            SystemPrompt = _context.SP_ChatExcel().Replace("{UserLanguage}", UserLanguage).Replace("{Location}", ThisAddIn.Location) &
                $" Your name is '{AN5}'. The current date and time is: {DateTime.Now.ToString("F")}. "
            txtUserInput.Text = ""

            ' Call the LLM function asynchronously
            Dim aiResponse As String = Await SharedMethods.LLM(_context, SystemPrompt,
                $"Welcome the user in {UserLanguage} by (1) referring to the time of day based on the current time in {UserLanguage} , such as in 'good morning', and (2) asking in {UserLanguage} what you can do, but do not say your name.",
                "", "", 0, _useSecondApi, True)

            aiResponse = aiResponse.Replace(vbLf, "").Replace(vbCr, "").Replace(vbCrLf, "") & vbCrLf
            aiResponse = aiResponse.Replace("**", "").Replace("_", "").Replace("`", "")

            ' Remove the "Thinking..." placeholder and update AI response on the UI thread
            Await UpdateUIAsync(Sub()
                                    AppendToChatHistory(Environment.NewLine & $"{AN5}: " & aiResponse)
                                End Sub)

            _chatHistory.Add(("assistant", aiResponse))

            PreceedingNewline = Environment.NewLine
            Return ""

        Catch ex As System.Exception
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Safely executes a UI update action on the UI thread.
    ''' </summary>
    Private Async Function UpdateUIAsync(action As System.Action) As System.Threading.Tasks.Task
        If InvokeRequired Then
            Await System.Threading.Tasks.Task.Run(Sub() Me.Invoke(action))
        Else
            action()
        End If
    End Function

    ''' <summary>
    ''' Appends text to the chat history textbox.
    ''' </summary>
    Private Sub AppendToChatHistory(text As String)
        If txtChatHistory.InvokeRequired Then
            txtChatHistory.Invoke(Sub() txtChatHistory.AppendText(text))
        Else
            txtChatHistory.AppendText(text)
        End If
    End Sub

    ''' <summary>
    ''' Removes the last line from chat history textbox.
    ''' </summary>
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

    ''' <summary>
    ''' Toggles TopMost state and persists setting.
    ''' </summary>
    Private Sub chkStayontop_Click(sender As Object, e As EventArgs)
        Me.TopMost = Not Me.TopMost
        My.Settings.NotAlwaysOnTop = Me.TopMost
        My.Settings.Save()
    End Sub

    ''' <summary>
    ''' Toggles command execution permission and saves setting.
    ''' </summary>
    Private Sub chkPermitCommands_Click(sender As Object, e As EventArgs)
        My.Settings.DoCommands = Not My.Settings.DoCommands
        My.Settings.Save()
    End Sub

    ''' <summary>
    ''' Handles inclusion of selection: validates selection, toggles worksheet inclusion, persists setting.
    ''' </summary>
    Private Sub chkIncludeselection_Click(sender As Object, e As EventArgs) Handles chkIncludeselection.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim selection As Excel.Range = TryCast(app.Selection, Excel.Range)

        ' Check if selection is valid and contains data
        If selection Is Nothing OrElse IsSelectionEmpty(selection) Then
            chkIncludeselection.Checked = False
        ElseIf chkIncludeDocText.Checked Then
            chkIncludeDocText.Checked = False
        End If

        My.Settings.IncludeSelection = chkIncludeselection.Checked
        My.Settings.Save()
    End Sub

    ''' <summary>
    ''' Determines if selection intersects with meaningful UsedRange content.
    ''' </summary>
    Private Function IsSelectionEmpty(selection As Excel.Range) As Boolean
        Dim ws As Excel.Worksheet = selection.Worksheet
        Dim app As Excel.Application = ws.Application

        ' build the range of all cells that "mean something"
        Dim infoRange As Excel.Range = ws.UsedRange

        ' see if any of those intersect the user’s selection
        Dim intersected As Excel.Range = Nothing
        Try
            intersected = app.Intersect(selection, infoRange)
        Catch ex As System.Exception
            ' should never really happen, but just in case
            Return True
        End Try

        ' if nothing in common, it's empty
        Return (intersected Is Nothing) OrElse (intersected.Cells.Count = 0)
    End Function

    ''' <summary>
    ''' Handles inclusion of entire worksheet and warns if large; mutually exclusive with selection.
    ''' </summary>
    Private Sub chkIncludeDocText_Click(sender As Object, e As EventArgs)

        If chkIncludeselection.Checked Then
            chkIncludeselection.Checked = False
        End If
        My.Settings.IncludeDocument = chkIncludeDocText.Checked
        My.Settings.Save()

        If Globals.ThisAddIn.SizeOfWorksheet() > Globals.ThisAddIn.LargeWorksheetSize And chkIncludeDocText.Checked Then
            ShowCustomMessageBox($"Because this worksheet is large (a range of {Globals.ThisAddIn.SizeOfWorksheet()} cells, even if not all are used), it may slow down your interaction with the chatbot, because each time you send a question, the entire worksheet will be passed to {AN5}. If you want to speed up, include only a selection only.")
        End If

    End Sub

    ''' <summary>
    ''' Copies entire chat history to clipboard.
    ''' </summary>
    Private Sub btnCopy_Click(sender As Object, e As EventArgs)
        My.Computer.Clipboard.SetText(txtChatHistory.Text)
    End Sub

    ''' <summary>
    ''' Copies last assistant message content to clipboard.
    ''' </summary>
    Private Sub btnCopyLastAnswer_Click(sender As Object, e As EventArgs)
        Dim lastAssistantMsg = _chatHistory.Where(Function(x) x.Role = "assistant").LastOrDefault()
        If lastAssistantMsg.Content IsNot Nothing Then
            My.Computer.Clipboard.SetText(lastAssistantMsg.Content)
        Else
            SharedMethods.ShowCustomMessageBox("No last AI answer available.")
        End If
    End Sub

    ''' <summary>
    ''' Toggles active model between primary and secondary and updates title text.
    ''' </summary>
    Private Sub btnSwitchModel_Click(sender As Object, e As EventArgs)
        _useSecondApi = Not _useSecondApi
        Me.Text = $"Chat (using " & If(_useSecondApi, _context.INI_Model_2, _context.INI_Model) & ")"
    End Sub

    ''' <summary>
    ''' Clears chat history (memory + UI + persisted), sends a new welcome message.
    ''' </summary>
    Private Sub btnClear_Click(sender As Object, e As EventArgs)

        _chatHistory.Clear()
        txtChatHistory.Clear()
        OldChat = ""
        PreceedingNewline = ""
        My.Settings.LastChatHistory = ""
        My.Settings.Save()
        Dim result = WelcomeMessage()
    End Sub

    ''' <summary>
    ''' Handles Escape key to save bounded chat snippet and close form.
    ''' </summary>
    Private Sub frmAIChat_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Dim conversation As String = txtChatHistory.Text
            If conversation.Length > _context.INI_ChatCap Then
                conversation = conversation.Substring(conversation.Length - _context.INI_ChatCap)
            End If
            My.Settings.LastChatHistory = conversation
            My.Settings.Save()
            Close()
        End If
    End Sub

    ''' <summary>
    ''' Button-driven exit: saves bounded chat snippet then closes.
    ''' </summary>
    Private Sub btnExit_Click(sender As Object, e As EventArgs)
        Dim conversation As String = txtChatHistory.Text
        If conversation.Length > _context.INI_ChatCap Then
            conversation = conversation.Substring(conversation.Length - _context.INI_ChatCap)
        End If
        My.Settings.LastChatHistory = conversation
        My.Settings.Save()
        Close()
    End Sub

    ''' <summary>
    ''' Persists chat snippet, window bounds, and settings on form closing.
    ''' </summary>
    Private Sub frmAIChat_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        ' Save the chat history before the form closes
        Dim conversation As String = txtChatHistory.Text
        If conversation.Length > _context.INI_ChatCap Then
            conversation = conversation.Substring(conversation.Length - _context.INI_ChatCap)
        End If
        My.Settings.LastChatHistory = conversation

        ' Save the form's location and size to My.Settings
        If Me.WindowState = FormWindowState.Normal Then
            My.Settings.FormLocation = Me.Location
            My.Settings.FormSize = Me.Size
        Else
            ' If the form is minimized or maximized, save the restored bounds
            My.Settings.FormLocation = Me.RestoreBounds.Location
            My.Settings.FormSize = Me.RestoreBounds.Size
        End If
        My.Settings.Save()
    End Sub

    ''' <summary>
    ''' Legacy handler for Ctrl+Enter send (unused).
    ''' </summary>
    Private Sub oldUserInput_KeyDown(sender As Object, e As KeyEventArgs)
        If e.Control AndAlso e.KeyCode = Keys.Enter Then
            btnSend.PerformClick()
            e.Handled = True
        End If
    End Sub

    ''' <summary>
    ''' Handles Enter to send; Shift+Enter inserts newline; suppresses default Enter behavior when sending.
    ''' </summary>
    Private Sub UserInput_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            If e.Shift Then
                ' Allow Shift+Enter to insert a new line (default behavior)
                Return
            Else
                ' Enter alone sends the message
                e.SuppressKeyPress = True
                btnSend.PerformClick()
                e.Handled = True
            End If
        End If
    End Sub

    ''' <summary>
    ''' Builds bounded conversation context string prefixed with role identifiers.
    ''' </summary>
    Private Function BuildConversationString(history As List(Of (Role As String, Content As String))) As String
        Dim sb As New StringBuilder()
        Dim totalLength As Integer = 0
        Dim maxLength As Integer = _context.INI_ChatCap

        ' Iterate through the history in reverse order (most recent messages first)
        For Each msg In history.AsEnumerable().Reverse()
            Dim message As String
            If msg.Role = "user" Then
                message = $"User: {msg.Content}{Environment.NewLine}"
            Else
                message = $"{AN5}: {msg.Content}{Environment.NewLine}"
            End If

            ' Check if adding this message will exceed the limit
            If totalLength + message.Length > maxLength Then
                ' If so, truncate the message to fit within the limit
                Dim remainingLength As Integer = maxLength - totalLength
                If remainingLength > 0 Then
                    sb.Insert(0, message.Substring(0, remainingLength))
                End If
                Exit For
            Else
                ' Otherwise, append the full message
                sb.Insert(0, message)
                totalLength += message.Length
            End If
        Next

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Paint handler (unused placeholder) for checkbox panel.
    ''' </summary>
    Private Sub pnlCheckboxes_Paint(sender As Object, e As PaintEventArgs)
    End Sub

    ''' <summary>Represents one parsed embedded command.</summary>
    Public Class ParsedCommand
        Public Property Command As String
        Public Property Argument1 As String
        Public Property Argument2 As String
    End Class

    ''' <summary>
    ''' Parses embedded commands of format [#command: @@argument1@@ §§argument2§§ #] into a list.
    ''' </summary>
    Private Function ParseCommands(input As String) As List(Of ParsedCommand)
        Dim results As New List(Of ParsedCommand)

        Try
            Dim pattern As String = "\[#(?<cmd>[^:]+):\s*@@(?<arg1>[^@]+)@@\s*(?:§§(?<arg2>[^§]*)§§)?\s*#\]"
            Dim regex As New Regex(pattern)

            For Each m As Match In regex.Matches(input)
                Dim pc As New ParsedCommand()

                pc.Command = m.Groups("cmd").Value.Trim()
                pc.Argument1 = m.Groups("arg1").Value.Trim()

                ' If arg2 wasn't found, it might be blank
                If m.Groups("arg2") IsNot Nothing Then
                    pc.Argument2 = m.Groups("arg2").Value.Trim().Replace("\r\n", vbCrLf).Replace("\n", vbCrLf).Replace("\r", vbCrLf)
                End If

                If String.IsNullOrEmpty(pc.Argument2) Then
                    pc.Argument2 = ""
                Else
                    pc.Argument1 = pc.Argument1.Replace("\r\n", ".*").Replace("\n", ".*").Replace("\r", ".*")
                    pc.Argument1 = pc.Argument1.Replace(vbCrLf, ".*").Replace(vbCr, ".*").Replace(vbLf, ".*")
                End If

                If Not results.Any(Function(x) x.Command = pc.Command AndAlso x.Argument1 = pc.Argument1 AndAlso x.Argument2 = pc.Argument2) Then
                    results.Add(pc)
                End If
            Next

        Catch ex As System.Exception
            MsgBox("Error in ParseCommands: " & ex.Message, MsgBoxStyle.Critical)
        End Try

        Return results
    End Function

    ''' <summary>
    ''' Removes embedded command blocks and collapses excessive line breaks.
    ''' </summary>
    Public Function RemoveCommands(input As String) As String
        Dim output As String = input
        Try
            Dim commandPattern As String = "\s*[\r\n]*\s*\[#[^:]+:\s*@@[^@]+@@\s*(?:§§[^§]*§§)?\s*#\]\s*[\r\n]*\s*"
            Dim regex As New Regex(commandPattern)
            output = regex.Replace(input, "")

            Dim whitespacePattern As String = "[\r\n]{3,}"
            Dim collapseRegex As New Regex(whitespacePattern)
            output = collapseRegex.Replace(output, Environment.NewLine)

        Catch ex As System.Exception
            MsgBox("Error in RemoveCommands: " & ex.Message, MsgBoxStyle.Critical)
        End Try

        Return output
    End Function

    ''' <summary>Holds concatenated commands (unused placeholder).</summary>
    Private CommandsList As String = ""

    ''' <summary>
    ''' Executes parsed commands by bringing Excel to foreground and applying instructions.
    ''' </summary>
    Public Sub ExecuteAnyCommands(commands As String)
        Dim topmost As Boolean = Me.TopMost
        Me.TopMost = False

        Dim instructions As New List(Of String)
        instructions = Globals.ThisAddIn.ParseLLMResponse(commands)

        If instructions.Count > 0 Then
            ' Bring Excel window to front (instead of Application.Activate())
            Dim hwnd As IntPtr = CType(Globals.ThisAddIn.Application.Hwnd, IntPtr)
            SetForegroundWindow(hwnd)

            System.Threading.Thread.Sleep(200)

            Globals.ThisAddIn.undoStates.Clear()
            Globals.ThisAddIn.ApplyLLMInstructions(instructions, True)
            Dim result = Globals.Ribbons.Ribbon1.UpdateUndoButton()
        End If

        Me.TopMost = topmost
        Me.Focus()
    End Sub

End Class