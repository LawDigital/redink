' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ConfigWizardForm.vb
' Purpose: WinForms modal dialog for the script-driven Configuration Wizard.
'          Presents a sidebar ListBox for group navigation with a scrollable
'          content panel showing field controls, descriptions, and branch selectors.
'
' Architecture / Responsibilities:
'   - Wizard UI shell: Builds a fixed-dialog form with group list navigation, header,
'     description, scrollable content panel, footer notes, and action buttons.
'   - DPI-aware layout: Uses a DPI scale factor to size and position controls and
'     applies post-handle scaling for HiDPI environments.
'   - Group rendering: Renders fields for a selected group with dynamic label sizing,
'     typed input controls, optional branch selector, and descriptive text.
'   - Default indicators: Shows "(default)" or "(default — will not be saved)" based on
'     built-in defaults and whether the key exists in the on-disk INI.
'   - Change tracking: Tracks original values and loaded defaults to distinguish true
'     edits from pre-populated values.
'   - Validation and save: Validates fields using `ConfigWizardEngine`, writes changes
'     to the target INI via `WriteIniValues`, and prompts the user for confirmation.
'
' External dependencies:
'   - `ConfigWizardEngine`: condition evaluation, validation, and INI write operations.
'   - `SharedMethods`: default values dictionary and utilities (e.g., registry access).
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

Namespace SharedLibrary

    ''' <summary>
    ''' Modal wizard form that guides administrators through INI configuration
    ''' using a JSON-defined group/field structure.
    ''' </summary>
    Public Class ConfigWizardForm
        Inherits Form

        Protected Overrides Sub OnShown(e As System.EventArgs)
            MyBase.OnShown(e)
            Me.TopMost = True
            Global.SharedLibrary.SharedLibrary.SharedMethods.ForceDialogToForeground(Me)
            Global.SharedLibrary.SharedLibrary.SharedMethods.AttachForeignForegroundWatchdog(Me)
        End Sub


        Private ReadOnly _context As ISharedContext
        Private ReadOnly _definition As WizardDefinition
        Private _iniPath As String

        ''' <summary>Working copy of all INI values being edited.</summary>
        Private ReadOnly _currentValues As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        ''' <summary>Tracks which fields have been modified from their original on-disk value.</summary>
        Private ReadOnly _originalValues As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        ''' <summary>Canonical defaults dictionary used to detect whether a value matches its built-in default.</summary>
        Private ReadOnly _keysToSkipWhenDefault As Dictionary(Of String, Object) = GetKeysToSkipWhenDefault()

        ''' <summary>Maps field key to its "(default)" indicator label for the currently displayed group.</summary>
        Private _defaultIndicators As New Dictionary(Of String, Label)(StringComparer.OrdinalIgnoreCase)

        ''' <summary>
        ''' Records the value each field had when the group was loaded (from on-disk INI or
        ''' wizard JSON default). Used to distinguish real user edits from pre-populated defaults.
        ''' </summary>
        Private _loadedValues As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        ''' <summary>
        ''' Maps group list indices to the original definition indices. When groups are
        ''' filtered by their <see cref="WizardGroup.Condition"/>, the sidebar only
        ''' contains visible groups; this list maps each sidebar position to the
        ''' corresponding index in <see cref="WizardDefinition.Groups"/>.
        ''' </summary>
        Private ReadOnly _visibleGroupIndices As New List(Of Integer)

        ' UI elements
        Private WithEvents _groupList As ListBox
        Private _contentPanel As Panel
        Private _descriptionLabel As Label
        Private _titleLabel As Label
        Private WithEvents _btnBack As Button
        Private WithEvents _btnNext As Button
        Private WithEvents _btnSave As Button
        Private WithEvents _btnCancel As Button
        Private WithEvents _btnInfo As Button
        Private WithEvents _btnIniPath As Button
        Private WithEvents _btnReload As Button
        Private WithEvents _btnEncodeKey As Button

        ''' <summary>Label showing the path of the INI file currently being edited.</summary>
        Private _lblPath As Label

        ''' <summary>Maps field key to its input control for the currently displayed group.</summary>
        Private _fieldControls As New Dictionary(Of String, Control)(StringComparer.OrdinalIgnoreCase)

        ''' <summary>Maps branch field key to its ComboBox for the currently displayed group.</summary>
        Private _branchCombo As ComboBox = Nothing
        Private _currentBranchField As String = ""

        ''' <summary>Tracks the currently selected template per branch field until it is explicitly applied.</summary>
        Private ReadOnly _pendingBranchSelections As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        ''' <summary>Prevents change handlers from re-entrancy during group rendering.</summary>
        Private _isLoadingGroup As Boolean = False

        ' --- #1: Sidebar width computed from actual group titles ---
        Private _sidebarWidth As Integer = 200

        ''' <summary>Stores the bottom Y coordinate of the footer notes so LoadGroup can compute the content panel height.</summary>
        Private _contentPanelBottomLimit As Integer = 0

        ''' <summary>DPI scale factor relative to 96 DPI (1.0 = 100%, 1.25 = 125%, 1.5 = 150%, etc.).</summary>
        Private _dpiScale As Single = 1.0F

        ''' <summary>
        ''' Initializes a new instance of the wizard form with the given definition and INI values.
        ''' </summary>
        ''' <param name="context">Shared context used for environment details and defaults.</param>
        ''' <param name="definition">Wizard definition containing groups and fields.</param>
        ''' <param name="iniPath">Path to the INI file being edited.</param>
        ''' <param name="iniValues">Current INI key/value pairs loaded from disk.</param>
        Public Sub New(context As ISharedContext, definition As WizardDefinition, iniPath As String, iniValues As Dictionary(Of String, String))
            _context = context
            _definition = definition
            _iniPath = iniPath

            ' Copy INI values into working set
            For Each kvp In iniValues
                _currentValues(kvp.Key) = kvp.Value
                _originalValues(kvp.Key) = kvp.Value
            Next

            ' --- Enable DPI and font scaling (must be set before any control creation) ---
            Me.AutoScaleDimensions = New SizeF(96.0F, 96.0F)
            Me.AutoScaleMode = AutoScaleMode.Dpi

            InitializeWizardUI()
        End Sub

        ''' <summary>
        ''' Scales a base pixel value by the current DPI factor.
        ''' </summary>
        ''' <param name="basePixels">Pixel value at 96 DPI.</param>
        ''' <returns>Scaled pixel value for the current DPI.</returns>
        Private Function Dpi(basePixels As Integer) As Integer
            Return CInt(Math.Ceiling(basePixels * _dpiScale))
        End Function

        ''' <summary>
        ''' Returns the standard height used for inline action buttons.
        ''' </summary>
        Private Function GetInlineButtonHeight() As Integer
            Using btn As New Button()
                btn.Text = "Apply Template"
                btn.AutoSize = True
                Return btn.PreferredSize.Height
            End Using
        End Function

        ''' <summary>
        ''' Builds and initializes the wizard UI controls and layout.
        ''' </summary>
        Private Sub InitializeWizardUI()
            ' Compute DPI scale factor
            Using g As Graphics = Me.CreateGraphics()
                _dpiScale = g.DpiX / 96.0F
            End Using

            Me.Text = $"{AN} {_definition.Title}"
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ShowInTaskbar = False

            Dim standardFont As New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            Me.Font = standardFont

            Try
                Dim bmp As New Bitmap(GetLogoBitmap(LogoType.Standard))
                Me.Icon = Icon.FromHandle(bmp.GetHicon())
            Catch
            End Try

            ' Build the list of visible groups (evaluate group-level conditions)
            _visibleGroupIndices.Clear()
            For i As Integer = 0 To _definition.Groups.Count - 1
                Dim grp = _definition.Groups(i)
                If ConfigWizardEngine.EvaluateCondition(grp.Condition, _currentValues) Then
                    _visibleGroupIndices.Add(i)
                End If
            Next

            ' --- #1: Compute sidebar width from the longest visible group title ---
            _sidebarWidth = Dpi(200)
            For Each idx In _visibleGroupIndices
                Dim tw As Integer = TextRenderer.MeasureText(_definition.Groups(idx).Title, standardFont).Width + Dpi(30)
                If tw > _sidebarWidth Then _sidebarWidth = tw
            Next
            _sidebarWidth = Math.Min(_sidebarWidth, Dpi(350)) ' cap

            ' Responsive sizing
            Dim workArea = Screen.FromPoint(Cursor.Position).WorkingArea
            Dim formWidth As Integer = Math.Min(Dpi(1050), CInt(workArea.Width * 0.8))
            Dim formHeight As Integer = Math.Min(Dpi(700), CInt(workArea.Height * 0.8))
            ' Ensure the form is wide enough for the sidebar + content
            formWidth = Math.Max(formWidth, _sidebarWidth + Dpi(500))
            Me.ClientSize = New Size(formWidth, formHeight)

            ' Sidebar ListBox — only visible groups
            _groupList = New ListBox() With {
                .Location = New Point(Dpi(10), Dpi(10)),
                .Size = New Size(_sidebarWidth, formHeight - Dpi(80)),
                .Font = standardFont,
                .IntegralHeight = False
            }
            For Each idx In _visibleGroupIndices
                _groupList.Items.Add(_definition.Groups(idx).Title)
            Next
            Me.Controls.Add(_groupList)

            ' Title label
            _titleLabel = New Label() With {
                .Location = New Point(_sidebarWidth + Dpi(20), Dpi(10)),
                .AutoSize = True,
                .Font = New Font("Segoe UI", 11.0F, FontStyle.Bold, GraphicsUnit.Point)
            }
            Me.Controls.Add(_titleLabel)

            ' Description label with padding and dynamic height; AutoSize + MaximumSize handles word-wrap
            _descriptionLabel = New Label() With {
                .Location = New Point(_sidebarWidth + Dpi(20), Dpi(36)),
                .MaximumSize = New Size(formWidth - _sidebarWidth - Dpi(50), 0),
                .AutoSize = True,
                .Font = standardFont,
                .ForeColor = Color.FromArgb(80, 80, 80),
                .Padding = New Padding(0, Dpi(4), 0, Dpi(4))
            }
            Me.Controls.Add(_descriptionLabel)

            ' --- Footer area: buttons, path label, default note ---
            Dim btnY As Integer = formHeight - Dpi(50)
            Dim btnSpacing As Integer = Dpi(10)
            Dim rightMargin As Integer = Dpi(20)

            ' Path label above the buttons with generous spacing
            _lblPath = New Label() With {
                .Text = $"Editing: {_iniPath}",
                .AutoSize = True,
                .ForeColor = Color.Gray,
                .Font = New Font("Segoe UI", 7.5F, FontStyle.Italic, GraphicsUnit.Point),
                .Location = New Point(_sidebarWidth + Dpi(20), btnY - Dpi(28)),
                .MaximumSize = New Size(formWidth - _sidebarWidth - Dpi(50), 0)
            }
            Me.Controls.Add(_lblPath)

            ' Note explaining the default/will-not-save indicator — more space above the path label
            Dim lblDefaultNote As New Label() With {
                .Text = "Fields marked ""(default — will not be saved)"" match the built-in default and will be omitted from the .ini file on save.",
                .AutoSize = True,
                .ForeColor = Color.FromArgb(140, 140, 140),
                .Font = New Font("Segoe UI", 7.0F, FontStyle.Italic, GraphicsUnit.Point),
                .Location = New Point(_sidebarWidth + Dpi(20), _lblPath.Top - Dpi(26)),
                .MaximumSize = New Size(formWidth - _sidebarWidth - Dpi(50), 0)
            }
            Me.Controls.Add(lblDefaultNote)

            ' Remember the bottom limit for the content panel (generous gap above footer notes)
            _contentPanelBottomLimit = lblDefaultNote.Top - Dpi(16)

            ' Scrollable content panel — top adapts to description height via LoadGroup
            _contentPanel = New Panel() With {
                .Location = New Point(_sidebarWidth + Dpi(20), Dpi(80)),
                .Size = New Size(formWidth - _sidebarWidth - Dpi(40), _contentPanelBottomLimit - Dpi(80)),
                .AutoScroll = True,
                .BorderStyle = BorderStyle.None
            }
            Me.Controls.Add(_contentPanel)

            ' Buttons (left side: Cancel, Back, Next, Save & Close)
            _btnCancel = New Button() With {.Text = "Cancel", .AutoSize = True, .Location = New Point(Dpi(10), btnY)}
            _btnBack = New Button() With {.Text = "< Back", .AutoSize = True, .Location = New Point(_btnCancel.Right + btnSpacing, btnY)}
            _btnNext = New Button() With {.Text = "Next >", .AutoSize = True, .Location = New Point(_btnBack.Right + btnSpacing, btnY)}
            _btnSave = New Button() With {.Text = "Save && Close", .AutoSize = True, .Location = New Point(_btnNext.Right + btnSpacing + Dpi(20), btnY)}

            ' "INI Path" button (right-aligned, rightmost — with right margin)
            _btnIniPath = New Button() With {.Text = "INI Path", .AutoSize = True}
            _btnIniPath.Location = New Point(formWidth - _btnIniPath.PreferredSize.Width - rightMargin, btnY)
            Dim iniPathTip As New ToolTip()
            iniPathTip.SetToolTip(_btnIniPath, "Shows the registry INI path and copies it to the clipboard.")

            ' "Reload INI" button (right-aligned, left of INI Path — with spacing between)
            _btnReload = New Button() With {.Text = "Reload INI", .AutoSize = True}
            _btnReload.Location = New Point(_btnIniPath.Left - _btnReload.PreferredSize.Width - btnSpacing * 2, btnY)
            Dim reloadTip As New ToolTip()
            reloadTip.SetToolTip(_btnReload, "Reload values from a local or central redink.ini file, discarding unsaved changes.")

            ' "Encode Key" button (right-aligned, left of Reload INI — with spacing between)
            _btnEncodeKey = New Button() With {.Text = "Encode Key", .AutoSize = True}
            _btnEncodeKey.Location = New Point(_btnReload.Left - _btnEncodeKey.PreferredSize.Width - btnSpacing * 2, btnY)
            Dim encodeKeyTip As New ToolTip()
            encodeKeyTip.SetToolTip(_btnEncodeKey, "Encodes an API key with a user-provided secret key and copies the result to the clipboard.")

            ' "Client Name" button (right-aligned, left of Encode Key — with spacing between)
            _btnInfo = New Button() With {.Text = "Client Name", .AutoSize = True}
            _btnInfo.Location = New Point(_btnEncodeKey.Left - _btnInfo.PreferredSize.Width - btnSpacing * 2, btnY)
            Dim infoTip As New ToolTip()
            infoTip.SetToolTip(_btnInfo, "Shows the client name (computer name) and copies it to the clipboard.")

            Me.Controls.AddRange({_btnCancel, _btnBack, _btnNext, _btnSave, _btnInfo, _btnEncodeKey, _btnReload, _btnIniPath})

            ' Select first group
            If _groupList.Items.Count > 0 Then
                _groupList.SelectedIndex = 0
            End If
        End Sub

        ''' <summary>
        ''' Applies additional DPI scaling after the window handle is created,
        ''' consistent with the pattern used in <see cref="ModelSelectorForm"/>.
        ''' </summary>
        ''' <param name="e">Event data for handle creation.</param>
        Protected Overrides Sub OnHandleCreated(e As EventArgs)
            MyBase.OnHandleCreated(e)
            Dim handleDpiScale As Single = Me.DeviceDpi / 96.0F
            If handleDpiScale <> 1.0F Then
                Me.Scale(New SizeF(handleDpiScale, handleDpiScale))
            End If
        End Sub

        ''' <summary>
        ''' Updates the "(default)" indicator label for a given field key.
        ''' Shows "(default — will not be saved)" if the current value matches the built-in default
        ''' and the key is NOT already present in the on-disk INI file.
        ''' Shows "(modified)" only if the user actually changed the value from what was loaded.
        ''' </summary>
        ''' <param name="fieldKey">The INI key for the field.</param>
        ''' <param name="currentValue">The current string value in the control.</param>
        ''' <param name="fieldType">The wizard field type (e.g. "boolean", "string", "integer").</param>
        Private Sub UpdateDefaultIndicator(fieldKey As String, currentValue As String, Optional fieldType As String = "")
            If Not _defaultIndicators.ContainsKey(fieldKey) Then Return

            Dim indicator As Label = _defaultIndicators(fieldKey)

            ' Determine whether the value matches its built-in default.
            ' First check the canonical defaults dictionary.
            Dim isDefault As Boolean = IsDefaultValue(fieldKey, currentValue, _keysToSkipWhenDefault)

            ' Fallback for boolean fields not in the defaults dictionary:
            ' the INI convention is that absent booleans default to False.
            If Not isDefault AndAlso
               fieldType.Equals("boolean", StringComparison.OrdinalIgnoreCase) AndAlso
               Not _keysToSkipWhenDefault.ContainsKey(fieldKey) Then
                isDefault = String.Equals(currentValue, "False", StringComparison.OrdinalIgnoreCase)
            End If

            If isDefault Then
                If _originalValues.ContainsKey(fieldKey) Then
                    indicator.Text = "(default)"
                    indicator.ForeColor = Color.FromArgb(140, 140, 140)
                Else
                    indicator.Text = "(default " & ChrW(8212) & " will not be saved)"
                    indicator.ForeColor = Color.FromArgb(180, 140, 60)
                End If
            Else
                ' Compare against the value loaded into the control (which may be
                ' from the on-disk INI or from the wizard JSON field default).
                ' This avoids falsely showing "(modified)" for pre-populated defaults.
                Dim loadedVal As String = ""
                If _loadedValues.ContainsKey(fieldKey) Then
                    loadedVal = _loadedValues(fieldKey)
                ElseIf _originalValues.ContainsKey(fieldKey) Then
                    loadedVal = _originalValues(fieldKey)
                End If

                If Not String.Equals(currentValue, loadedVal, StringComparison.OrdinalIgnoreCase) Then
                    indicator.Text = "(modified)"
                    indicator.ForeColor = Color.FromArgb(0, 120, 60)
                Else
                    indicator.Text = ""
                End If
            End If
        End Sub


#Region "Group Navigation"

        ''' <summary>
        ''' Returns the definition-level group index for the currently selected sidebar item,
        ''' or -1 if nothing is selected.
        ''' </summary>
        Private Function GetCurrentDefinitionGroupIndex() As Integer
            If _groupList.SelectedIndex < 0 OrElse _groupList.SelectedIndex >= _visibleGroupIndices.Count Then
                Return -1
            End If
            Return _visibleGroupIndices(_groupList.SelectedIndex)
        End Function

        ''' <summary>
        ''' Saves current group values and loads the newly selected group.
        ''' </summary>
        Private Sub GroupList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles _groupList.SelectedIndexChanged
            If _groupList.SelectedIndex < 0 Then Return
            SaveCurrentGroupValues()
            Dim defIndex As Integer = GetCurrentDefinitionGroupIndex()
            If defIndex >= 0 Then LoadGroup(defIndex)
            UpdateNavigationButtons()
        End Sub

        ''' <summary>
        ''' Navigates to the previous group in the list.
        ''' </summary>
        Private Sub BtnBack_Click(sender As Object, e As EventArgs) Handles _btnBack.Click
            If _groupList.SelectedIndex > 0 Then
                _groupList.SelectedIndex -= 1
            End If
        End Sub

        ''' <summary>
        ''' Validates the current group and navigates to the next group if valid.
        ''' </summary>
        Private Sub BtnNext_Click(sender As Object, e As EventArgs) Handles _btnNext.Click
            If _groupList.SelectedIndex < _groupList.Items.Count - 1 Then
                Dim errors = ValidateCurrentGroup()
                If errors.Count > 0 Then
                    ShowCustomMessageBox("Please fix the following issues:" & vbCrLf & vbCrLf & String.Join(vbCrLf, errors))
                    Return
                End If
                _groupList.SelectedIndex += 1
            End If
        End Sub

        ''' <summary>
        ''' Validates all visible groups and persists changes to the target INI file.
        ''' </summary>
        Private Sub BtnSave_Click(sender As Object, e As EventArgs) Handles _btnSave.Click
            SaveCurrentGroupValues()

            ' Validate all visible groups using their current _currentValues snapshot
            Dim allErrors As New List(Of String)
            For Each defIdx In _visibleGroupIndices
                Dim grp = _definition.Groups(defIdx)
                For Each field In grp.Fields
                    If Not ConfigWizardEngine.EvaluateCondition(field.Condition, _currentValues) Then Continue For
                    Dim val As String = ""
                    If _currentValues.ContainsKey(field.Key) Then val = _currentValues(field.Key)
                    Dim err = ConfigWizardEngine.ValidateField(field, val)
                    If Not String.IsNullOrWhiteSpace(err) Then
                        allErrors.Add($"[{grp.Title}] {err}")
                    End If
                Next
            Next

            If allErrors.Count > 0 Then
                ShowCustomMessageBox("Please fix the following issues:" & vbCrLf & vbCrLf & String.Join(vbCrLf, allErrors))
                Return
            End If

            ' Build a lookup of wizard JSON defaults and branch defaults keyed by field key
            Dim wizardDefaults As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            For Each grp In _definition.Groups
                For Each field In grp.Fields
                    If Not String.IsNullOrWhiteSpace(field.Default) Then
                        wizardDefaults(field.Key) = field.Default
                    End If
                Next
                ' Include branch defaults so that selecting a branch template without further
                ' edits does not count as a change when the key was absent from the on-disk INI.
                For Each branch In grp.Branches
                    For Each kvp In branch.Defaults
                        If Not wizardDefaults.ContainsKey(kvp.Key) Then
                            wizardDefaults(kvp.Key) = kvp.Value
                        End If
                    Next
                Next
            Next

            ' Collect all branch-selector keys (UI-only template pickers, e.g. "_Provider").
            ' These are never valid INI keys and must not be persisted to the file.
            Dim branchFieldKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each grp In _definition.Groups
                If Not String.IsNullOrWhiteSpace(grp.BranchField) Then
                    branchFieldKeys.Add(grp.BranchField)
                End If
            Next

            Dim changedValues As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            For Each kvp In _currentValues
                ' Skip branch-selector pseudo-keys so they are never written to the INI.
                If branchFieldKeys.Contains(kvp.Key) Then Continue For

                Dim origVal As String = ""
                Dim hadOriginal As Boolean = _originalValues.ContainsKey(kvp.Key)
                If hadOriginal Then origVal = _originalValues(kvp.Key)

                If Not String.Equals(kvp.Value, origVal, StringComparison.Ordinal) Then
                    ' If the key was NOT in the on-disk INI and the current value matches
                    ' the wizard JSON default, the user never actually changed it — skip.
                    If Not hadOriginal Then
                        Dim wizDefault As String = ""
                        If wizardDefaults.ContainsKey(kvp.Key) Then wizDefault = wizardDefaults(kvp.Key)
                        If String.Equals(kvp.Value, wizDefault, StringComparison.Ordinal) Then
                            Continue For
                        End If
                    End If
                    changedValues(kvp.Key) = kvp.Value
                End If
            Next

            If changedValues.Count = 0 Then
                ShowCustomMessageBox("No changes were made.")
                Me.DialogResult = DialogResult.Cancel
                Me.Close()
                Return
            End If

            Dim confirm = ShowCustomYesNoBox(
                $"Save {changedValues.Count} changed value(s) to '{Path.GetFileName(_iniPath)}'? A timestamped backup will be created first.",
                "Save", "Cancel")
            If confirm <> 1 Then Return

            Try
                ConfigWizardEngine.WriteIniValues(_iniPath, changedValues)
                ShowCustomMessageBox($"{changedValues.Count} value(s) saved successfully. A backup was created in the same directory.")
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Catch ex As Exception
                ShowCustomMessageBox($"Error saving configuration: {ex.Message}")
            End Try
        End Sub


        ''' <summary>
        ''' Cancels the wizard without saving.
        ''' </summary>
        Private Sub BtnCancel_Click(sender As Object, e As EventArgs) Handles _btnCancel.Click
            Me.DialogResult = DialogResult.Cancel
            Me.Close()
        End Sub

        ''' <summary>
        ''' Displays and copies the client name to the clipboard.
        ''' </summary>
        Private Sub BtnInfo_Click(sender As Object, e As EventArgs) Handles _btnInfo.Click
            Dim clientName As String = GetCurrentClientIdentifier()
            Dim displayName As String = If(String.IsNullOrWhiteSpace(clientName), "(not available)", clientName)
            Dim infoText As String = $"Client Name: {displayName}"

            Try
                Clipboard.SetText(infoText)
                ShowCustomMessageBox(infoText & vbCrLf & vbCrLf & "(Copied to clipboard.)")
            Catch
                ShowCustomMessageBox(infoText)
            End Try
        End Sub

        ''' <summary>
        ''' Displays and copies the registry INI path to the clipboard.
        ''' </summary>
        Private Sub BtnIniPath_Click(sender As Object, e As EventArgs) Handles _btnIniPath.Click
            Dim regPath As String = ""
            Try
                regPath = GetFromRegistry(RegPath_Base, RegPath_IniPath, True)
            Catch
            End Try

            Dim displayPath As String = If(String.IsNullOrWhiteSpace(regPath), "(not configured)", regPath)
            Dim infoText As String = $"Registry INI Path: {displayPath}"

            Try
                Clipboard.SetText(infoText)
                ShowCustomMessageBox(infoText & vbCrLf & vbCrLf & "(Copied to clipboard.)")
            Catch
                ShowCustomMessageBox(infoText)
            End Try
        End Sub

        ''' <summary>
        ''' Encodes an API key using a user-provided secret key and copies the result to the clipboard.
        ''' Prompts for the API key, an optional prefix, and the secret key. The prefix portion
        ''' (if present) is preserved unencoded in the output.
        ''' </summary>
        Private Sub BtnEncodeKey_Click(sender As Object, e As EventArgs) Handles _btnEncodeKey.Click
            ' Prompt for the API key to encode
            Dim apiKey As String = ShowCustomInputBox("Please enter the API key to encode:", "API Key Encryptor", True)
            If String.IsNullOrEmpty(apiKey) OrElse apiKey = "ESC" Then Return
            apiKey = apiKey.Trim()

            If String.IsNullOrEmpty(apiKey) Then
                ShowCustomMessageBox("No API key entered.")
                Return
            End If

            ' Prompt for the prefix (pre-populated from the current context if available)
            Dim prefixDefault As String = If(_context IsNot Nothing, _context.INI_APIKeyPrefix, "")
            Dim prefixValue As String = ShowCustomInputBox(
                "Please enter the API key prefix (as used in the configuration file, if any):",
                "API Key Encryptor", True, prefixDefault)
            If prefixValue = "ESC" Then Return

            ' Prompt for the secret key
            Dim secretKey As String = ShowCustomInputBox("Please enter the secret key:", "API Key Encryptor", True)
            If String.IsNullOrEmpty(secretKey) OrElse secretKey = "ESC" Then
                ShowCustomMessageBox("No secret key entered.")
                Return
            End If

            ' Encode: strip prefix before encoding, re-attach afterwards
            Dim hadPrefix As Boolean = False
            Dim modifiedKey As String

            If Not String.IsNullOrEmpty(prefixValue) AndAlso apiKey.StartsWith(prefixValue) Then
                hadPrefix = True
                modifiedKey = apiKey.Substring(prefixValue.Length)
            Else
                modifiedKey = apiKey
            End If

            Dim encChoice As Integer = ShowCustomYesNoBox(
                "Which encryption do you want to use for this key?",
                "Strong (recommended)", "Standard (legacy XOR)",
                "API Key Encryptor")
            If encChoice <> 1 AndAlso encChoice <> 2 Then Return

            Dim resultKey As String = CodeString(modifiedKey, secretKey, encChoice = 1)

            If hadPrefix Then
                resultKey = prefixValue & resultKey
            End If

            ' Copy to clipboard and notify
            Try
                Clipboard.SetText(resultKey)
                ShowCustomMessageBox("Encoded key (also copied to clipboard):" & vbCrLf & vbCrLf & resultKey)
            Catch
                ShowCustomMessageBox("Encoded key:" & vbCrLf & vbCrLf & resultKey)
            End Try
        End Sub

        ''' <summary>
        ''' Reloads INI values from a user-selected configuration file (local, central, or Word shared).
        ''' Follows the same path resolution logic as <see cref="InitializeConfig"/> to discover all
        ''' candidate INI files. When the current host is Excel or Outlook and no host-specific INI
        ''' exists, the Word shared INI is offered as a candidate.
        ''' Discards all unsaved changes in the wizard and refreshes the current group.
        ''' </summary>
        Private Sub BtnReload_Click(sender As Object, e As EventArgs) Handles _btnReload.Click
            Dim candidates As Dictionary(Of String, String) = ConfigWizardEngine.BuildIniCandidates(_context)

            Dim selectedPath As String = Nothing

            If candidates.Count = 0 Then
                ShowCustomMessageBox("No configuration files were found.")
                Return
            ElseIf candidates.Count = 1 Then
                ' Only one candidate — use it directly
                selectedPath = candidates.Values.First()
            Else
                ' Multiple candidates — let the user choose
                Dim choice As String = ShowSelectionForm(
                    $"Select the '{AN2}.ini' file to reload into the wizard:",
                    "Reload Configuration File",
                    candidates.Keys)

                If String.IsNullOrWhiteSpace(choice) OrElse Not candidates.ContainsKey(choice) Then
                    Return ' User cancelled
                End If

                selectedPath = candidates(choice)
            End If

            If String.IsNullOrWhiteSpace(selectedPath) OrElse Not File.Exists(selectedPath) Then
                ShowCustomMessageBox($"The selected configuration file does not exist: {selectedPath}")
                Return
            End If

            ' Confirm discard of unsaved changes
            Dim confirm = ShowCustomYesNoBox(
                $"Reload all values from '{Path.GetFileName(selectedPath)}'? Any unsaved changes in the wizard will be lost.",
                "Reload", "Cancel")
            If confirm <> 1 Then Return

            ' Re-read the selected INI file
            Dim newValues As Dictionary(Of String, String) = ConfigWizardEngine.ReadIniValues(selectedPath)

            ' Replace the working set, original values, and target path
            _currentValues.Clear()
            _originalValues.Clear()
            _pendingBranchSelections.Clear()
            For Each kvp In newValues
                _currentValues(kvp.Key) = kvp.Value
                _originalValues(kvp.Key) = kvp.Value
            Next

            _iniPath = selectedPath
            _lblPath.Text = $"Editing: {_iniPath}"

            ' Reload the current group to reflect the new values
            Dim defIndex As Integer = GetCurrentDefinitionGroupIndex()
            If defIndex >= 0 Then
                LoadGroup(defIndex)
            End If
        End Sub

        ''' <summary>
        ''' Updates the enabled state of Back/Next navigation buttons.
        ''' </summary>
        Private Sub UpdateNavigationButtons()
            _btnBack.Enabled = _groupList.SelectedIndex > 0
            _btnNext.Enabled = _groupList.SelectedIndex < _groupList.Items.Count - 1
        End Sub

#End Region

#Region "Group Rendering"

        ''' <summary>
        ''' Determines whether a field should be visible based on its own condition AND any
        ''' <see cref="WizardBranch.ShowExtra"/> list for the currently selected branch.
        ''' A field is visible when:
        '''   1. Its <see cref="WizardField.Condition"/> is satisfied (or empty), AND
        '''   2. Either the current group has no branch selector, or the field's key appears in
        '''      the selected branch's <see cref="WizardBranch.ShowExtra"/> list, or the list is empty.
        ''' </summary>
        Private Function IsFieldVisible(field As WizardField, grp As WizardGroup) As Boolean
            ' Standard condition check
            If Not ConfigWizardEngine.EvaluateCondition(field.Condition, _currentValues) Then
                Return False
            End If

            ' ShowExtra filtering — only applies when a branch selector is active
            If Not String.IsNullOrWhiteSpace(grp.BranchField) AndAlso grp.Branches.Count > 0 AndAlso
               _branchCombo IsNot Nothing AndAlso _branchCombo.SelectedIndex >= 0 Then
                Dim selectedBranch = grp.Branches(_branchCombo.SelectedIndex)
                If selectedBranch.ShowExtra IsNot Nothing AndAlso selectedBranch.ShowExtra.Count > 0 Then
                    ' The branch explicitly lists extra fields to show. Fields NOT in this list
                    ' are only shown if they don't have a condition that ties them to the branch.
                    ' Fields whose condition references the branch field itself are branch-specific
                    ' and must appear in ShowExtra to be visible.
                    If Not String.IsNullOrWhiteSpace(field.Condition) AndAlso
                       field.Condition.Contains(grp.BranchField) Then
                        ' Condition references the branch field — require ShowExtra listing
                        If Not selectedBranch.ShowExtra.Any(
                            Function(k) k.Equals(field.Key, StringComparison.OrdinalIgnoreCase)) Then
                            Return False
                        End If
                    End If
                End If
            End If

            Return True
        End Function

        ''' <summary>
        ''' Builds and displays the controls for the specified definition-level group index.
        ''' </summary>
        ''' <param name="groupIndex">Index into <see cref="WizardDefinition.Groups"/>.</param>
        Private Sub LoadGroup(groupIndex As Integer)
            _isLoadingGroup = True
            _contentPanel.SuspendLayout()
            _contentPanel.Controls.Clear()
            _fieldControls.Clear()
            _defaultIndicators.Clear()
            _loadedValues.Clear()
            _branchCombo = Nothing
            _currentBranchField = ""

            Dim grp = _definition.Groups(groupIndex)
            _titleLabel.Text = grp.Title

            Dim hasDescription As Boolean = Not String.IsNullOrWhiteSpace(grp.Description)
            If hasDescription Then
                _descriptionLabel.Text = grp.Description
                _descriptionLabel.Visible = True
                ' Reset size so AutoSize recalculates from scratch for the new text
                _descriptionLabel.Size = New Size(0, 0)
                Me.PerformLayout()
                ' For word-wrapped AutoSize labels, a second pass may be needed
                ' after the width constraint from MaximumSize triggers re-flow.
                Me.PerformLayout()
            Else
                _descriptionLabel.Text = ""
                _descriptionLabel.Visible = False
            End If

            Dim contentTop As Integer
            If hasDescription Then
                contentTop = _descriptionLabel.Bottom + Dpi(6)
            Else
                contentTop = _titleLabel.Bottom + Dpi(8)
            End If
            _contentPanel.Top = contentTop
            _contentPanel.Height = _contentPanelBottomLimit - contentTop

            Dim contentWidth As Integer = _contentPanel.Width - SystemInformation.VerticalScrollBarWidth - Dpi(10)
            Dim standardFont As Font = Me.Font
            Dim lineHeight As Integer = CInt(standardFont.Height * 1.8)
            Dim descFont As New Font("Segoe UI", 8.0F, FontStyle.Italic, GraphicsUnit.Point)
            Dim defaultIndicatorFont As New Font("Segoe UI", 7.5F, FontStyle.Italic, GraphicsUnit.Point)
            Dim inlineButtonHeight As Integer = GetInlineButtonHeight()

            ' Pre-measure all visible labels to determine the dynamic FIELD_LABEL_WIDTH
            Dim dynamicLabelWidth As Integer = Dpi(120)
            For Each field In grp.Fields
                If Not ConfigWizardEngine.EvaluateCondition(field.Condition, _currentValues) Then Continue For
                Dim tw As Integer = TextRenderer.MeasureText(field.Label & ":", standardFont).Width + Dpi(10)
                If tw > dynamicLabelWidth Then dynamicLabelWidth = tw
            Next
            If Not String.IsNullOrWhiteSpace(grp.BranchField) AndAlso grp.Branches.Count > 0 Then
                Dim bw As Integer = TextRenderer.MeasureText("Provider template:", New Font(standardFont, FontStyle.Bold)).Width + Dpi(10)
                If bw > dynamicLabelWidth Then dynamicLabelWidth = bw
            End If
            dynamicLabelWidth = Math.Min(dynamicLabelWidth, CInt(contentWidth * 0.45))
            Dim fieldWidth As Integer = contentWidth - dynamicLabelWidth - Dpi(20)

            Dim yPos As Integer = 0

            ' Branch selector (if applicable)          
            If Not String.IsNullOrWhiteSpace(grp.BranchField) AndAlso grp.Branches.Count > 0 Then
                _currentBranchField = grp.BranchField

                Dim lbl As New Label() With {
                    .Text = "Provider template:",
                    .AutoSize = True,
                    .Font = New Font(standardFont, FontStyle.Bold),
                    .Location = New Point(0, yPos)
                }
                _contentPanel.Controls.Add(lbl)

                Dim btnApplyTemplate As New Button() With {
                    .Text = "Apply Template",
                    .AutoSize = False,
                    .Font = standardFont
                }

                Dim applyTemplateButtonWidth As Integer =
                    TextRenderer.MeasureText(btnApplyTemplate.Text, btnApplyTemplate.Font).Width + Dpi(18)

                btnApplyTemplate.Size = New Size(applyTemplateButtonWidth, inlineButtonHeight)

                Dim branchComboWidth As Integer = Math.Max(Dpi(160), fieldWidth - btnApplyTemplate.Width - Dpi(8))

                _branchCombo = New ComboBox() With {
                    .DropDownStyle = ComboBoxStyle.DropDownList,
                    .Width = branchComboWidth,
                    .Location = New Point(dynamicLabelWidth, yPos)
                }
                For Each branch In grp.Branches
                    _branchCombo.Items.Add(branch.Label)
                Next

                Dim selectedBranchLabel As String = ""
                If _pendingBranchSelections.ContainsKey(grp.BranchField) Then
                    selectedBranchLabel = _pendingBranchSelections(grp.BranchField)
                ElseIf _currentValues.ContainsKey(grp.BranchField) Then
                    selectedBranchLabel = _currentValues(grp.BranchField)
                End If

                Dim idx = _branchCombo.Items.IndexOf(selectedBranchLabel)
                _branchCombo.SelectedIndex = If(idx >= 0, idx, 0)

                AddHandler _branchCombo.SelectedIndexChanged,
                    Sub(s, ev)
                        If _isLoadingGroup Then Return
                        _pendingBranchSelections(_currentBranchField) =
                            If(_branchCombo.SelectedItem IsNot Nothing, _branchCombo.SelectedItem.ToString(), "")
                    End Sub

                btnApplyTemplate.Location = New Point(_branchCombo.Right + Dpi(8), yPos - Dpi(1))
                AddHandler btnApplyTemplate.Click,
                    Sub(s, ev)
                        ApplyBranchDefaults(grp)
                    End Sub

                Dim applyTemplateNote As New Label() With {
                    .Text = "Select a template and click ""Apply Template"" to overwrite the values in this section and clear unused template-specific entries.",
                    .AutoSize = True,
                    .MaximumSize = New Size(fieldWidth, 0),
                    .Font = descFont,
                    .ForeColor = Color.FromArgb(100, 100, 100),
                    .Location = New Point(dynamicLabelWidth, yPos + _branchCombo.Height + Dpi(4))
                }

                _contentPanel.Controls.Add(_branchCombo)
                _contentPanel.Controls.Add(btnApplyTemplate)
                _contentPanel.Controls.Add(applyTemplateNote)

                yPos = applyTemplateNote.Bottom + Dpi(8)
            End If

            ' Fields
            For Each field In grp.Fields
                If Not IsFieldVisible(field, grp) Then Continue For

                Dim fieldLabel As New Label() With {
                    .Text = field.Label & ":",
                    .AutoSize = True,
                    .Font = standardFont,
                    .Location = New Point(0, yPos + Dpi(2))
                }
                _contentPanel.Controls.Add(fieldLabel)

                Dim currentVal As String = ""
                If _currentValues.ContainsKey(field.Key) Then
                    currentVal = _currentValues(field.Key)
                ElseIf Not String.IsNullOrWhiteSpace(field.Default) Then
                    currentVal = field.Default
                End If

                ' Record the value as loaded — baseline for "(modified)" detection
                _loadedValues(field.Key) = currentVal

                Dim fieldType = If(field.Type, "string").ToLowerInvariant()

                If fieldType = "boolean" Then
                    Dim chk As New CheckBox() With {
                        .Checked = String.Equals(currentVal, "True", StringComparison.OrdinalIgnoreCase),
                        .Location = New Point(dynamicLabelWidth, yPos),
                        .AutoSize = True
                    }
                    _contentPanel.Controls.Add(chk)
                    _fieldControls(field.Key) = chk

                    Dim boolIndicator As New Label() With {
                        .AutoSize = True,
                        .Font = defaultIndicatorFont,
                        .ForeColor = Color.FromArgb(140, 140, 140),
                        .Location = New Point(chk.Right + Dpi(4), yPos + Dpi(3)),
                        .Text = ""
                    }
                    _contentPanel.Controls.Add(boolIndicator)
                    _defaultIndicators(field.Key) = boolIndicator
                    UpdateDefaultIndicator(field.Key, currentVal, fieldType)

                    Dim capturedKey As String = field.Key
                    Dim capturedType As String = fieldType
                    AddHandler chk.CheckedChanged, Sub(s, ev)
                                                       UpdateDefaultIndicator(capturedKey, DirectCast(s, CheckBox).Checked.ToString(), capturedType)
                                                       If _isLoadingGroup Then Return
                                                       _currentValues(capturedKey) = DirectCast(s, CheckBox).Checked.ToString()
                                                       Dim scrollPos As Integer = _contentPanel.VerticalScroll.Value
                                                       SaveCurrentGroupValues()
                                                       LoadGroup(groupIndex)
                                                       _contentPanel.AutoScrollPosition = New Point(0, scrollPos)
                                                   End Sub

                ElseIf fieldType = "multiline" Then
                    Dim txt As New TextBox() With {
                        .Text = currentVal,
                        .Multiline = True,
                        .ScrollBars = ScrollBars.Vertical,
                        .Size = New Size(fieldWidth, lineHeight * 3),
                        .Location = New Point(dynamicLabelWidth, yPos)
                    }
                    _contentPanel.Controls.Add(txt)
                    _fieldControls(field.Key) = txt

                    Dim mlIndicator As New Label() With {
                        .AutoSize = True,
                        .Font = defaultIndicatorFont,
                        .ForeColor = Color.FromArgb(140, 140, 140),
                        .Location = New Point(dynamicLabelWidth, yPos + txt.Height + Dpi(1)),
                        .Text = ""
                    }
                    _contentPanel.Controls.Add(mlIndicator)
                    _defaultIndicators(field.Key) = mlIndicator
                    UpdateDefaultIndicator(field.Key, currentVal, fieldType)

                    Dim capturedKey As String = field.Key
                    Dim capturedType As String = fieldType
                    AddHandler txt.TextChanged, Sub(s, ev)
                                                    UpdateDefaultIndicator(capturedKey, DirectCast(s, TextBox).Text, capturedType)
                                                End Sub

                    ' Advance past the TextBox + indicator label so the description does not overlap
                    yPos += lineHeight * 2 + Dpi(3) + mlIndicator.PreferredHeight

                ElseIf fieldType = "password" Then
                    Dim showButtonWidth As Integer = Dpi(70)

                    Dim txt As New TextBox() With {
                        .Text = currentVal,
                        .UseSystemPasswordChar = True,
                        .Size = New Size(fieldWidth - showButtonWidth - Dpi(10), Dpi(20)),
                        .Location = New Point(dynamicLabelWidth, yPos)
                    }
                    Dim btnShow As New Button() With {
                        .Text = "Show",
                        .Size = New Size(showButtonWidth, inlineButtonHeight),
                        .Location = New Point(txt.Right + Dpi(5), yPos - Math.Max(0, (inlineButtonHeight - txt.Height) \ 2))
                    }
                    AddHandler btnShow.Click, Sub(s, ev)
                                                  txt.UseSystemPasswordChar = Not txt.UseSystemPasswordChar
                                                  btnShow.Text = If(txt.UseSystemPasswordChar, "Show", "Hide")
                                              End Sub
                    _contentPanel.Controls.Add(txt)
                    _contentPanel.Controls.Add(btnShow)
                    _fieldControls(field.Key) = txt

                    Dim pwIndicator As New Label() With {
                        .AutoSize = True,
                        .Font = defaultIndicatorFont,
                        .ForeColor = Color.FromArgb(140, 140, 140),
                        .Location = New Point(btnShow.Right + Dpi(6), yPos + Dpi(4)),
                        .Text = ""
                    }
                    _contentPanel.Controls.Add(pwIndicator)
                    _defaultIndicators(field.Key) = pwIndicator
                    UpdateDefaultIndicator(field.Key, currentVal, fieldType)

                    Dim capturedKey As String = field.Key
                    Dim capturedType As String = fieldType
                    AddHandler txt.TextChanged, Sub(s, ev)
                                                    UpdateDefaultIndicator(capturedKey, DirectCast(s, TextBox).Text, capturedType)
                                                End Sub

                Else
                    ' string, integer, path
                    Dim txt As New TextBox() With {
                        .Text = currentVal,
                        .Size = New Size(fieldWidth, Dpi(20)),
                        .Location = New Point(dynamicLabelWidth, yPos)
                    }
                    _contentPanel.Controls.Add(txt)
                    _fieldControls(field.Key) = txt

                    Dim txtIndicator As New Label() With {
                        .AutoSize = True,
                        .Font = defaultIndicatorFont,
                        .ForeColor = Color.FromArgb(140, 140, 140),
                        .Location = New Point(dynamicLabelWidth, yPos + txt.Height + Dpi(1)),
                        .Text = ""
                    }
                    _contentPanel.Controls.Add(txtIndicator)
                    _defaultIndicators(field.Key) = txtIndicator
                    UpdateDefaultIndicator(field.Key, currentVal, fieldType)

                    Dim capturedKey As String = field.Key
                    Dim capturedType As String = fieldType
                    AddHandler txt.TextChanged, Sub(s, ev)
                                                    UpdateDefaultIndicator(capturedKey, DirectCast(s, TextBox).Text, capturedType)
                                                End Sub

                    yPos += Dpi(14)
                End If

                yPos += lineHeight

                If Not String.IsNullOrWhiteSpace(field.Description) Then
                    Dim descGap As Integer = If(fieldType = "multiline", 0, Dpi(-4))
                    Dim descLabel As New Label() With {
                        .Text = field.Description,
                        .AutoSize = True,
                        .MaximumSize = New Size(fieldWidth, 0),
                        .Font = descFont,
                        .ForeColor = Color.FromArgb(100, 100, 100),
                        .Location = New Point(dynamicLabelWidth, yPos + descGap)
                    }
                    _contentPanel.Controls.Add(descLabel)
                    yPos += descLabel.PreferredHeight + Dpi(4)
                End If
            Next

            ' Group note
            If Not String.IsNullOrWhiteSpace(grp.Note) Then
                yPos += Dpi(10)
                Dim noteLabel As New Label() With {
                    .Text = grp.Note,
                    .AutoSize = True,
                    .MaximumSize = New Size(contentWidth, 0),
                    .Font = New Font("Segoe UI", 8.5F, FontStyle.Italic),
                    .ForeColor = Color.FromArgb(60, 60, 60),
                    .Location = New Point(0, yPos)
                }
                _contentPanel.Controls.Add(noteLabel)
            End If

            _contentPanel.ResumeLayout(True)
            _isLoadingGroup = False
        End Sub

        ''' <summary>
        ''' Returns the set of field keys controlled by provider templates in the current group.
        ''' These keys are cleared before the selected template defaults are applied.
        ''' </summary>
        Private Function GetBranchManagedFieldKeys(grp As WizardGroup) As HashSet(Of String)
            Dim managedKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each branch In grp.Branches
                For Each kvp In branch.Defaults
                    managedKeys.Add(kvp.Key)
                Next

                For Each key In branch.ShowExtra
                    If Not String.IsNullOrWhiteSpace(key) Then
                        managedKeys.Add(key)
                    End If
                Next
            Next

            For Each field In grp.Fields
                If Not String.IsNullOrWhiteSpace(field.Condition) AndAlso
                   field.Condition.IndexOf(grp.BranchField, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    managedKeys.Add(field.Key)
                End If
            Next

            managedKeys.Remove(grp.BranchField)

            Return managedKeys
        End Function

        ''' <summary>
        ''' Applies the defaults associated with the selected branch in the current group.
        ''' Existing template-controlled values are overwritten, and unused template values
        ''' are cleared or removed.
        ''' </summary>
        ''' <param name="grp">The current wizard group containing branches.</param>
        Private Sub ApplyBranchDefaults(grp As WizardGroup)
            If _branchCombo Is Nothing OrElse _branchCombo.SelectedIndex < 0 Then Return

            SaveCurrentGroupValues()

            Dim selectedBranch = grp.Branches(_branchCombo.SelectedIndex)
            Dim managedKeys As HashSet(Of String) = GetBranchManagedFieldKeys(grp)

            For Each key In managedKeys
                If _originalValues.ContainsKey(key) Then
                    _currentValues(key) = ""
                Else
                    _currentValues.Remove(key)
                End If
            Next

            _currentValues(_currentBranchField) = selectedBranch.Label
            _pendingBranchSelections(_currentBranchField) = selectedBranch.Label

            For Each kvp In selectedBranch.Defaults
                _currentValues(kvp.Key) = kvp.Value
            Next

            Dim defIndex As Integer = GetCurrentDefinitionGroupIndex()
            If defIndex >= 0 Then LoadGroup(defIndex)
        End Sub

#End Region

#Region "Value Sync"

        ''' <summary>
        ''' Captures the current UI values into the working values dictionary.
        ''' </summary>
        Private Sub SaveCurrentGroupValues()
            For Each kvp In _fieldControls
                If TypeOf kvp.Value Is CheckBox Then
                    _currentValues(kvp.Key) = DirectCast(kvp.Value, CheckBox).Checked.ToString()
                ElseIf TypeOf kvp.Value Is TextBox Then
                    _currentValues(kvp.Key) = DirectCast(kvp.Value, TextBox).Text
                End If
            Next

            If _branchCombo IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(_currentBranchField) Then
                _pendingBranchSelections(_currentBranchField) =
                    If(_branchCombo.SelectedItem IsNot Nothing, _branchCombo.SelectedItem.ToString(), "")
            End If
        End Sub

        ''' <summary>
        ''' Validates the current group and returns any validation errors.
        ''' </summary>
        ''' <returns>List of validation error messages.</returns>
        Private Function ValidateCurrentGroup() As List(Of String)
            SaveCurrentGroupValues()
            Dim errors As New List(Of String)

            Dim defIndex As Integer = GetCurrentDefinitionGroupIndex()
            If defIndex < 0 Then Return errors
            Dim grp = _definition.Groups(defIndex)

            For Each field In grp.Fields
                If Not ConfigWizardEngine.EvaluateCondition(field.Condition, _currentValues) Then Continue For
                Dim val As String = ""
                If _currentValues.ContainsKey(field.Key) Then val = _currentValues(field.Key)
                Dim err = ConfigWizardEngine.ValidateField(field, val)
                If Not String.IsNullOrWhiteSpace(err) Then
                    errors.Add(err)
                End If
            Next

            Return errors
        End Function

#End Region

    End Class

End Namespace
