' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.Config.vb
' Purpose: Configuration data class and startup dialog for Inky AutoPilot.
'          Allows the user to select model, define filter rules, set restrictions,
'          choose tools, and confirm settings before starting an AutoPilot session.
'
' Architecture:
'  - Dialog flow collects model selection, mailbox, filters, subject trigger,
'    whitelist, rate limits, and tool selection (when supported).
'  - Additional session options:
'      * EnableWebGrounding — tells the model about its native web-search capability.
'      * EnableScheduler — activates the AutoPilot Task Scheduler for recurring tasks.
'      * EnableVoicemailProcessing — transcribes and processes incoming voicemails
'        (requires VoicemailSenderAddress and VoicemailCallerIdMapPath).
'      * ReprocessLookbackHours — re-proposes already-processed mails in catch-up.
'      * IsUnattended — auto-start mode (skips catch-up preview, auto-approves replies).
'  - Settings are persisted to My.Settings and loaded as defaults for subsequent runs.
'  - Filter rule parsing is shared between the dialog and settings loader.
' =============================================================================


Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Configuration for an AutoPilot session.
    ''' </summary>
    Public Class AutoPilotConfig
        ''' <summary>Filter rules (domain/sender/folder wildcards).</summary>
        Public Property FilterRules As New List(Of AutoPilotFilterRule)()

        ''' <summary>Whitelisted sender patterns (auto-send without approval).</summary>
        Public Property WhitelistedSenders As New List(Of String)()

        ''' <summary>Subject trigger word (if set, only mails containing this word are processed).</summary>
        Public Property SubjectTriggerWord As String = Nothing

        ''' <summary>Whether to require approval for non-whitelisted senders.</summary>
        Public Property RequireApprovalForNonWhitelisted As Boolean = True

        ''' <summary>Per-sender cooldown in seconds.</summary>
        Public Property CooldownSeconds As Integer = AP_DefaultCooldownSeconds

        ''' <summary>Maximum replies per session.</summary>
        Public Property MaxRepliesPerSession As Integer = AP_DefaultMaxRepliesPerSession

        ''' <summary>Maximum attachment size in bytes.</summary>
        Public Property MaxAttachmentBytes As Long = AP_DefaultMaxAttachmentBytes

        ''' <summary>Whether to use the second API (alternate model applied).</summary>
        Public Property UseSecondApi As Boolean = False

        ''' <summary>
        ''' The ModelDescription key of the selected alternate model (INI section name + note).
        ''' Empty/Nothing means the primary or default secondary model is used.
        ''' </summary>
        Public Property SelectedModelKey As String = ""

        ''' <summary>User-selected external tools (from specialservices/INI).</summary>
        Public Property SelectedExternalTools As List(Of ModelConfig) = Nothing

        ''' <summary>SMTP address of the mailbox to monitor (empty = default/all).</summary>
        Public Property MonitoredMailbox As String = ""

        ''' <summary>Custom footer text appended to every AutoPilot reply.</summary>
        Public Property FooterText As String = ""

        ''' <summary>
        ''' When True, the session was auto-started (timer expired) without explicit user interaction.
        ''' Unattended mode skips the catch-up preview dialog (auto-selects all) and auto-approves
        ''' replies for non-whitelisted senders instead of blocking with an approval dialog.
        ''' </summary>
        Public Property IsUnattended As Boolean = False

        ''' <summary>
        ''' Number of hours to look back and re-propose already-processed mails for reprocessing.
        ''' 0 = disabled (only unprocessed mails are proposed in catch-up).
        ''' When > 0, mails within this window that match filters are shown in the catch-up dialog
        ''' even if they were already processed, allowing the operator to select them for reprocessing.
        ''' </summary>
        Public Property ReprocessLookbackHours As Integer = 0

        ''' <summary>
        ''' Number of hours after which answered incoming mails and all AutoPilot replies
        ''' in the same cleanup group are automatically deleted. 0 = disabled.
        ''' Deletion also covers items already moved to Deleted Items.
        ''' </summary>
        Public Property AutoDeleteAfterHours As Integer = 0

        ''' <summary>
        ''' When True, the system prompt tells the model it has built-in web search / grounding capability.
        ''' Only meaningful when the model natively supports web search (e.g. Gemini with grounding).
        ''' </summary>
        Public Property EnableWebGrounding As Boolean = False

        ''' <summary>
        ''' When True, incoming voicemails (audio attachments from the registered voicemail sender)
        ''' are transcribed and processed as instructions. Requires binary input support.
        ''' </summary>
        Public Property EnableVoicemailProcessing As Boolean = False

        ''' <summary>
        ''' SMTP address of the voicemail system (e.g. noreply@combox.swisscom.ch).
        ''' Only mails from this sender with audio attachments are treated as voicemails.
        ''' </summary>
        Public Property VoicemailSenderAddress As String = ""

        ''' <summary>
        ''' Path to a CSV file mapping caller IDs (phone numbers) to e-mail addresses.
        ''' Format: phone,email[,displayname]  (displayname is optional)
        ''' </summary>
        Public Property VoicemailCallerIdMapPath As String = ""

        ''' <summary>
        ''' When True, the AutoPilot Scheduler is enabled — tasks can be created, managed,
        ''' and executed automatically on schedule with results delivered by e-mail.
        ''' </summary>
        Public Property EnableScheduler As Boolean = False

        ''' <summary>
        ''' When True, per-user InkyMemory is enabled — the model learns from each user
        ''' and personalises responses. Users can opt in/out individually.
        ''' </summary>
        Public Property EnableUserMemory As Boolean = False

        ''' <summary>
        ''' When True, per-user home directory file storage is enabled — users can
        ''' store, retrieve, and manage persistent files for use in future requests.
        ''' </summary>
        Public Property EnableUserFiles As Boolean = False

        ''' <summary>
        ''' When True, privacy protections are enforced for web grounding and internet search tools.
        ''' The model is instructed not to include personal data, confidential information, or
        ''' non-public details in search queries sent to external services.
        ''' When False (default), no such restrictions are applied — the model may use any
        ''' information in search queries at the admin's discretion.
        ''' </summary>
        Public Property EnablePrivacyProtection As Boolean = False

        ''' <summary>
        ''' Full path to the optional per-sender tool policy file. When set, the file hard-limits
        ''' (in code, not via the LLM) which tools/skills/agents/sources individual senders may use.
        ''' Empty = no per-sender policy (all senders may use every selected tool).
        ''' </summary>
        Public Property SenderToolPolicyPath As String = ""

        ''' <summary>
        ''' Number of days that a sender's incoming attachments are retained so that
        ''' follow-up mails from the SAME sender in the SAME conversation can continue
        ''' the discussion with the earlier files still available. 0 = disabled.
        ''' Retention is applied only for whitelisted senders (AutoPilot session only).
        ''' </summary>
        Public Property ThreadRetentionDays As Integer = 0


    End Class

    ''' <summary>
    ''' Shows the AutoPilot configuration dialog. Returns a config object or Nothing if cancelled.
    ''' </summary>
    Private Function ShowAutoPilotConfigDialog() As AutoPilotConfig
        Dim config As New AutoPilotConfig()

        ' Restore the latest persisted AutoPilot configuration before reading defaults.
        TryRestoreAutoPilotSettingsFromRegistry()

        EnsureAutoPilotSettingsRestoredFromRegistry()

        ' Load persisted defaults so we can pre-populate the dialog fields
        Dim saved = LoadAutoPilotConfigDefaults()
        Dim previousMailbox As String = If(My.Settings.AP_MonitoredMailbox, "")

        ' ── Step 1: Model selection ──
        If INI_SecondAPI AndAlso Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
            ' Alternate models defined: pre-select the last used model
            If Not String.IsNullOrWhiteSpace(saved.SelectedModelKey) Then
                LastAlternateModel = saved.SelectedModelKey
            End If

            If Not ShowModelSelection(_context, INI_AlternateModelPath, $"{AN6} AutoPilot", "Select the model for AutoPilot:", "Reset to default model after AutoPilot stops") Then
                originalConfigLoaded = False
                Return Nothing
            End If
            config.UseSecondApi = True
            ' Capture the selected model key from the shared state set by ShowModelSelection
            config.SelectedModelKey = If(LastAlternateModel, "")
        ElseIf INI_SecondAPI Then
            ' Only primary/secondary available
            ' Pre-select based on persisted choice
            Dim modelChoice = ShowCustomYesNoBox(
                "Which AI model should " & AN6 & " AutoPilot use?",
                $"Primary ({INI_Model})",
                $"Secondary ({INI_Model_2})",
                header:=$"{AN6} AutoPilot Setup")
            If modelChoice = 0 Then Return Nothing
            config.UseSecondApi = (modelChoice = 2)
            config.SelectedModelKey = "" ' no alternate model key for primary/secondary toggle
        Else
            ' Only primary model available — no selection needed
            config.UseSecondApi = False
            config.SelectedModelKey = ""
        End If

        ' ── Step 1a: Check if the selected model supports tooling ──
        Dim modelSupportsTools As Boolean = True
        Dim baseConfig = GetCurrentConfig(_context)
        If Not ModelSupportsTooling(baseConfig) Then
            modelSupportsTools = False
            Dim modelLabel = If(Not String.IsNullOrWhiteSpace(config.SelectedModelKey), config.SelectedModelKey,
                             If(config.UseSecondApi, INI_Model_2, INI_Model))
            Dim proceed = ShowCustomYesNoBox(
                $"The selected model ({modelLabel}) does not support tool calling." & vbCrLf & vbCrLf &
                $"{AN6} AutoPilot will still respond to e-mails, but will not be able to " &
                $"use {Globals.ThisAddIn.ToolFriendlyName.ToLower} (databases, web retrieval, etc.) or process attachments " &
                "via built-in document tools." & vbCrLf & vbCrLf &
                "Do you want to proceed without tool support?",
                "Proceed without tools", "Cancel",
                header:=$"{AN6} AutoPilot — No Tool Support")
            If proceed <> 1 Then Return Nothing
        End If

        ' ── Step 1b: Mailbox selection ──
        Dim ns As Microsoft.Office.Interop.Outlook.NameSpace = Globals.ThisAddIn.Application.GetNamespace("MAPI")
        Dim accounts As Microsoft.Office.Interop.Outlook.Accounts = ns.Accounts
        If accounts.Count > 1 Then
            Dim accountItems As New List(Of SelectionItem)()
            accountItems.Add(New SelectionItem("(All mailboxes)", 0))
            For i As Integer = 1 To accounts.Count
                accountItems.Add(New SelectionItem(accounts(i).SmtpAddress, i))
            Next

            ' Pre-select last used mailbox
            Dim defaultValue As Integer = 0
            If Not String.IsNullOrWhiteSpace(previousMailbox) Then
                For i As Integer = 1 To accounts.Count
                    If accounts(i).SmtpAddress.Equals(previousMailbox, StringComparison.OrdinalIgnoreCase) Then
                        defaultValue = i
                        Exit For
                    End If
                Next
            End If

            Dim selected = SelectValue(
                accountItems,
                defaultValue,
                "Select the mailbox AutoPilot should monitor:",
                AN6 & " AutoPilot — Mailbox")

            If selected < 0 Then Return Nothing

            If selected = 0 Then
                config.MonitoredMailbox = ""
            Else
                config.MonitoredMailbox = accounts(selected).SmtpAddress
            End If
        End If

        ' Persist the selected mailbox immediately so subsequent runs and reinstall recovery remember it
        My.Settings.AP_MonitoredMailbox = If(config.MonitoredMailbox, "")
        SaveAutoPilotSettingsWithRegistryBackup()

        ' Determine whether the mailbox changed — if so, reset filter/whitelist to defaults
        Dim mailboxChanged As Boolean = Not String.Equals(
            If(previousMailbox, ""),
            If(config.MonitoredMailbox, ""),
            StringComparison.OrdinalIgnoreCase)

        ' ── Step 2: Filter rules ──
        Dim defaultFilter As String
        If mailboxChanged OrElse String.IsNullOrWhiteSpace(My.Settings.AP_FilterRules) Then
            defaultFilter = ""
            If Not String.IsNullOrWhiteSpace(config.MonitoredMailbox) AndAlso config.MonitoredMailbox.Contains("@") Then
                Dim domain = config.MonitoredMailbox.Substring(config.MonitoredMailbox.IndexOf("@"c))
                defaultFilter = domain
            End If
        Else
            defaultFilter = My.Settings.AP_FilterRules.Replace(vbLf, vbCrLf)
        End If

        Dim filterInput = ShowCustomInputBox(
            "Enter filter rules (one per line):" & vbCrLf & vbCrLf &
            "Formats:" & vbCrLf &
            "  @domain.com           — match sender domain" & vbCrLf &
            "  john@example.com      — match specific sender" & vbCrLf &
            "  *@*.example.com       — wildcard match" & vbCrLf &
            "  EXCLUDE @noreply.com  — exclude domain" & vbCrLf &
            "  FOLDER Inbox\Clients  — match folder" & vbCrLf & vbCrLf &
            "Leave empty to process all incoming mail.",
            $"{AN6} AutoPilot — Filters", False, defaultFilter)

        If filterInput Is Nothing Then Return Nothing

        If Not String.IsNullOrWhiteSpace(filterInput) Then
            config.FilterRules = ParseFilterRulesText(filterInput)
        End If

        ' ── Step 3: Subject trigger word ──
        Dim defaultTrigger As String = If(mailboxChanged, "", If(saved.SubjectTriggerWord, ""))

        Dim triggerInput = ShowCustomInputBox(
            "Enter a subject trigger word (optional):" & vbCrLf & vbCrLf &
            "If set, " & AN6 & " will only respond to mails whose subject contains this word." & vbCrLf &
            "Leave empty to respond to all mails that pass the filters.",
            $"{AN6} AutoPilot — Trigger Word", True, defaultTrigger)
        If triggerInput Is Nothing Then Return Nothing
        config.SubjectTriggerWord = If(String.IsNullOrWhiteSpace(triggerInput), Nothing, triggerInput.Trim())

        ' ── Step 4: Whitelist (auto-send senders) ──
        Dim defaultWhitelist As String
        If mailboxChanged OrElse String.IsNullOrWhiteSpace(My.Settings.AP_WhitelistedSenders) Then
            defaultWhitelist = ""
        Else
            defaultWhitelist = My.Settings.AP_WhitelistedSenders.Replace(vbLf, vbCrLf)
        End If

        Dim whitelistInput = ShowCustomInputBox(
            "Enter whitelisted senders/domains (one per line):" & vbCrLf & vbCrLf &
            "Mails from these senders will be auto-replied without approval." & vbCrLf &
            "All others will require your approval before sending." & vbCrLf & vbCrLf &
            "Examples:" & vbCrLf &
            "  *@mycompany.com              — all senders from this domain" & vbCrLf &
            "  trusted.partner@example.com  — specific sender" & vbCrLf &
            "  *@*.example.com              — all subdomains" & vbCrLf & vbCrLf &
            "Leave empty to require approval for all senders.",
            $"{AN6} AutoPilot — Whitelist", False, defaultWhitelist)
        If whitelistInput Is Nothing Then Return Nothing

        If Not String.IsNullOrWhiteSpace(whitelistInput) Then
            For Each line In whitelistInput.Split({vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)
                Dim trimmed = line.Trim()
                If trimmed.Length > 0 AndAlso Not trimmed.StartsWith(";") Then
                    ' Auto-correct bare @domain patterns: @domain.com → *@domain.com
                    ' (bare @domain would fail WildcardMatch since it anchors with ^ and $)
                    If trimmed.StartsWith("@") Then
                        trimmed = "*" & trimmed
                    End If
                    config.WhitelistedSenders.Add(trimmed)
                End If
            Next
        End If

        config.RequireApprovalForNonWhitelisted = True

        ' ── Step 5: Rate limits ──
        Dim pCooldown As New InputParameter() With {
            .Name = $"Cooldown per sender (seconds, default {AP_DefaultCooldownSeconds})",
            .Value = saved.CooldownSeconds.ToString()
        }
        Dim pMaxReplies As New InputParameter() With {
            .Name = $"Max replies per session (default {AP_DefaultMaxRepliesPerSession}; 0 = unlimited)",
            .Value = saved.MaxRepliesPerSession.ToString()
        }
        Dim pMaxAttachmentMb As New InputParameter() With {
            .Name = "Max attachment size (MB, default 10)",
            .Value = CInt(saved.MaxAttachmentBytes / 1024 / 1024).ToString()
        }
        Dim pReprocessLookback As New InputParameter() With {
            .Name = "Reprocess lookback (hours, 0 = only new mails)",
            .Value = saved.ReprocessLookbackHours.ToString()
        }
        Dim pAutoDeleteAfterHours As New InputParameter() With {
            .Name = "Auto-delete answered mails after N hours (0 = disabled; also empties Deleted Items for tagged mails)",
            .Value = saved.AutoDeleteAfterHours.ToString()
        }
        Dim pEnableWebGrounding As New InputParameter() With {
            .Name = "Add web grounding (if models are configured)",
            .Value = saved.EnableWebGrounding
        }
        Dim pEnableScheduler As New InputParameter() With {
            .Name = "Enable task scheduler (create && run scheduled tasks)",
            .Value = saved.EnableScheduler
        }
        Dim pEnableUserMemory As New InputParameter() With {
            .Name = "Enable per-user memory (learn user preferences)",
            .Value = saved.EnableUserMemory
        }
        Dim pEnableUserFiles As New InputParameter() With {
            .Name = "Enable per-user file storage (home directory)",
            .Value = saved.EnableUserFiles
        }
        Dim pEnablePrivacyProtection As New InputParameter() With {
            .Name = "Enable privacy protection for web/search queries (restrict personal data in queries)",
            .Value = saved.EnablePrivacyProtection
        }
        Dim pThreadRetentionDays As New InputParameter() With {
            .Name = "Retain a sender's attachments for follow-up discussion for N days (0 = disabled; whitelisted senders only)",
            .Value = saved.ThreadRetentionDays.ToString()
        }

        Dim paramsList As New List(Of InputParameter) From {
            pCooldown,
            pMaxReplies,
            pMaxAttachmentMb,
            pReprocessLookback,
            pAutoDeleteAfterHours,
            pEnableWebGrounding,
            pEnableScheduler,
            pEnableUserMemory,
            pEnableUserFiles,
            pEnablePrivacyProtection,
            pThreadRetentionDays
        }

        ' ── Voicemail processing (only if audio transcription is available) ──
        Dim audioTranscriptionAvailable As Boolean = IsAudioTranscriptionAvailable(_context)
        Dim pVoicemail As InputParameter = Nothing
        Dim pVoicemailSender As InputParameter = Nothing
        Dim pVoicemailMapPath As InputParameter = Nothing
        If audioTranscriptionAvailable Then
            pVoicemail = New InputParameter() With {
                .Name = "Process voicemails (transcribe & respond)",
                .Value = saved.EnableVoicemailProcessing
            }
            paramsList.Add(pVoicemail)

            pVoicemailSender = New InputParameter() With {
                .Name = "Voicemail sender address (e.g. comboxmailer@swisscom.com)",
                .Value = If(saved.VoicemailSenderAddress, "")
            }
            paramsList.Add(pVoicemailSender)

            pVoicemailMapPath = New InputParameter() With {
                .Name = "Caller ID → Email map file (CSV path)",
                .Value = If(saved.VoicemailCallerIdMapPath, "")
            }
            paramsList.Add(pVoicemailMapPath)
        End If

        Dim limitsOk = ShowCustomVariableInputForm(
            "Configure rate limits, restrictions, and optional features:",
            $"{AN6} AutoPilot — Limits & Features",
            paramsList.ToArray())

        If Not limitsOk Then Return Nothing

        Dim cooldown As Integer
        If Integer.TryParse(pCooldown.Value?.ToString(), cooldown) AndAlso cooldown >= 0 Then
            config.CooldownSeconds = cooldown
        End If

        Dim maxReplies As Integer
        If Integer.TryParse(pMaxReplies.Value?.ToString(), maxReplies) AndAlso maxReplies >= 0 Then
            config.MaxRepliesPerSession = maxReplies
        End If

        Dim maxMb As Integer
        If Integer.TryParse(pMaxAttachmentMb.Value?.ToString(), maxMb) AndAlso maxMb > 0 Then
            config.MaxAttachmentBytes = CLng(maxMb) * 1024 * 1024
        End If

        Dim reprocessHours As Integer
        If Integer.TryParse(pReprocessLookback.Value?.ToString(), reprocessHours) AndAlso reprocessHours >= 0 Then
            config.ReprocessLookbackHours = reprocessHours
        End If

        Dim autoDeleteHours As Integer
        If Integer.TryParse(pAutoDeleteAfterHours.Value?.ToString(), autoDeleteHours) AndAlso autoDeleteHours >= 0 Then
            config.AutoDeleteAfterHours = autoDeleteHours
        End If

        config.EnableWebGrounding = CBool(If(pEnableWebGrounding.Value, False))
        config.EnableScheduler = CBool(If(pEnableScheduler.Value, False))
        config.EnableUserMemory = CBool(If(pEnableUserMemory.Value, False))
        config.EnableUserFiles = CBool(If(pEnableUserFiles.Value, False))
        config.EnablePrivacyProtection = CBool(If(pEnablePrivacyProtection.Value, False))

        Dim threadRetentionDays As Integer
        If Integer.TryParse(pThreadRetentionDays.Value?.ToString(), threadRetentionDays) AndAlso threadRetentionDays >= 0 Then
            config.ThreadRetentionDays = threadRetentionDays
        End If

        ' Voicemail settings
        If audioTranscriptionAvailable AndAlso pVoicemail IsNot Nothing Then
            config.EnableVoicemailProcessing = CBool(If(pVoicemail.Value, False))
            config.VoicemailSenderAddress = If(pVoicemailSender?.Value?.ToString()?.Trim(), "")
            config.VoicemailCallerIdMapPath = If(pVoicemailMapPath?.Value?.ToString()?.Trim(), "").Trim(""""c)

            If config.EnableVoicemailProcessing Then
                If String.IsNullOrWhiteSpace(config.VoicemailSenderAddress) Then
                    ShowCustomMessageBox("Voicemail processing requires a voicemail sender address.", AN)
                    Return Nothing
                End If
                If String.IsNullOrWhiteSpace(config.VoicemailCallerIdMapPath) OrElse
                   Not IO.File.Exists(config.VoicemailCallerIdMapPath) Then
                    ShowCustomMessageBox("Voicemail processing requires a valid caller ID map file path." & vbCrLf &
                                         "File not found: " & If(config.VoicemailCallerIdMapPath, "(empty)"), AN)
                    Return Nothing
                End If
            End If
        End If

        ' ── Step 6: Source selection (only if model supports tooling) ──
        If modelSupportsTools Then
            Dim availableTools = GetAvailableToolsForAutoPilotSelection()
            If availableTools IsNot Nothing AndAlso availableTools.Count > 0 Then
                ' Load previously persisted tool names for pre-selection
                Dim previousToolNames As New List(Of String)()
                If Not String.IsNullOrWhiteSpace(My.Settings.AP_SelectedExternalToolNames) Then
                    previousToolNames = My.Settings.AP_SelectedExternalToolNames.Split({vbLf}, StringSplitOptions.RemoveEmptyEntries).
                        Select(Function(s) s.Trim()).Where(Function(s) s.Length > 0).ToList()
                End If

                Dim toolChoice = ShowCustomYesNoBox(
                    $"There are {availableTools.Count} {Globals.ThisAddIn.ToolFriendlyName.ToLower} available (web retrieval, etc.)." & vbCrLf &
                    If(previousToolNames.Count > 0,
                       $"Previously selected: {previousToolNames.Count} {Globals.ThisAddIn.ToolFriendlyName.ToLower}." & vbCrLf, "") &
                    "Would you like to enable them for AutoPilot?",
                    $"Yes, select {Globals.ThisAddIn.ToolFriendlyName.ToLower}", "No, skip",
                    header:=$"{AN6} AutoPilot — Optional {ToolFriendlyName}")

                If toolChoice = 1 Then
                    ' Use MultiModelSelectorForm with pre-selection if we have persisted names
                    If previousToolNames.Count > 0 Then
                        Using dlg As New SharedLibrary.SharedLibrary.MultiModelSelectorForm(
                            availableTools, Nothing,
                            title:=$"{AN6} AutoPilot {ToolFriendlyName}",
                            resetChecked:=False,
                            preselectMany:=previousToolNames,
                            instruction:=$"Select the optional {Globals.ThisAddIn.ToolFriendlyName.ToLower} to enable for AutoPilot:")
                            Dim __safeDialogOwner532 As System.Windows.Forms.IWin32Window = SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
                            If If(__safeDialogOwner532 IsNot Nothing, dlg.ShowDialog(__safeDialogOwner532), dlg.ShowDialog()) = DialogResult.OK Then
                                config.SelectedExternalTools = dlg.SelectedModels
                            End If
                        End Using
                    Else
                        config.SelectedExternalTools = ShowToolSelectionDialog(availableTools, preselectAll:=True, FriendlyName:=$"{AN6} AutoPilot {ToolFriendlyName}")
                    End If
                End If
            End If
        End If

        ' ── Step 6b: Per-sender tool policy file (optional) ──
        Dim defaultPolicyPath As String = If(saved.SenderToolPolicyPath, "")
        Dim policyInput = ShowCustomInputBox(
            "Enter the full path to a per-sender tool policy file (optional):" & vbCrLf & vbCrLf &
            "If set, this file hard-limits (in code, not via the AI) which tools, skills, " &
            "agents, and sources individual senders may use. It does NOT affect your mail " &
            "filters or the auto-send whitelist." & vbCrLf & vbCrLf &
            "Leave empty to let every sender use all selected tools.",
            $"{AN6} AutoPilot — Sender Tool Policy", True, defaultPolicyPath)
        If policyInput Is Nothing Then Return Nothing
        config.SenderToolPolicyPath = If(String.IsNullOrWhiteSpace(policyInput), "", policyInput.Trim())

        ' ── Step 7: Confirmation ──
        Dim confirmModelLabel As String
        If Not String.IsNullOrWhiteSpace(config.SelectedModelKey) Then
            confirmModelLabel = config.SelectedModelKey
        ElseIf config.UseSecondApi Then
            confirmModelLabel = INI_Model_2
        Else
            confirmModelLabel = INI_Model
        End If

        Dim summaryBuilder As New System.Text.StringBuilder()
        summaryBuilder.AppendLine($"Model: {confirmModelLabel}")
        If Not modelSupportsTools Then
            summaryBuilder.AppendLine("⚠ Model does not support tool calling — running without tools")
        End If
        summaryBuilder.AppendLine($"Filters: {config.FilterRules.Count} rule(s)")
        summaryBuilder.AppendLine($"Trigger word: {If(String.IsNullOrWhiteSpace(config.SubjectTriggerWord), "(none)", config.SubjectTriggerWord)}")
        summaryBuilder.AppendLine($"Whitelisted senders: {config.WhitelistedSenders.Count}")
        summaryBuilder.AppendLine($"Cooldown: {config.CooldownSeconds}s per sender")
        summaryBuilder.AppendLine($"Max replies: {If(config.MaxRepliesPerSession = 0, "unlimited", config.MaxRepliesPerSession.ToString())} per session")
        summaryBuilder.AppendLine($"Max attachment: {config.MaxAttachmentBytes / 1024 / 1024:F0} MB")
        summaryBuilder.AppendLine($"Auto-delete: {If(config.AutoDeleteAfterHours > 0, $"enabled after {config.AutoDeleteAfterHours}h", "disabled")}")
        summaryBuilder.AppendLine($"Web grounding: {If(config.EnableWebGrounding, "enabled", "disabled")}")
        summaryBuilder.AppendLine($"Task scheduler: {If(config.EnableScheduler, "enabled", "disabled")}")
        summaryBuilder.AppendLine($"User memory: {If(config.EnableUserMemory, "enabled", "disabled")}")
        summaryBuilder.AppendLine($"User file storage: {If(config.EnableUserFiles, "enabled", "disabled")}")
        summaryBuilder.AppendLine($"Privacy protection: {If(config.EnablePrivacyProtection, "enabled (queries sanitized)", "disabled (unrestricted)")}")
        summaryBuilder.AppendLine($"Sender tool policy: {If(String.IsNullOrWhiteSpace(config.SenderToolPolicyPath), "disabled", config.SenderToolPolicyPath)}")
        If config.EnableVoicemailProcessing Then
            summaryBuilder.AppendLine($"Voicemail processing: enabled (from {config.VoicemailSenderAddress})")
        End If
        If config.ReprocessLookbackHours > 0 Then
            summaryBuilder.AppendLine($"Reprocess lookback: {config.ReprocessLookbackHours}h (already-processed mails re-proposed)")
        Else
            summaryBuilder.AppendLine("Reprocess lookback: disabled (only new mails)")
        End If
        summaryBuilder.AppendLine($"{ToolFriendlyName}: {If(config.SelectedExternalTools?.Count, 0)}")
        If modelSupportsTools Then
            Dim internalToolCount As Integer = 0
            Try : internalToolCount = GetAutoPilotInternalTools().Count : Catch : End Try
            summaryBuilder.AppendLine($"Internal tools: {internalToolCount} (document processing, PDF extraction, etc.)")
        End If
        summaryBuilder.AppendLine()
        summaryBuilder.AppendLine($"Mode: Auto-send for {config.WhitelistedSenders.Count} whitelisted pattern(s); approval required for others.")

        Dim confirm = ShowCustomYesNoBox(
            summaryBuilder.ToString().TrimEnd(),
            $"Start {AN6} AutoPilot", "Cancel",
            header:=$"{AN6} AutoPilot — Confirm")

        If confirm <> 1 Then Return Nothing

        Return config
    End Function

    ''' <summary>
    ''' Parses multi-line filter rules text into a list of AutoPilotFilterRule objects.
    ''' Shared between the config dialog and LoadAutoPilotConfigDefaults.
    ''' </summary>
    Private Shared Function ParseFilterRulesText(filterText As String) As List(Of AutoPilotFilterRule)
        Dim rules As New List(Of AutoPilotFilterRule)()
        If String.IsNullOrWhiteSpace(filterText) Then Return rules

        For Each line In filterText.Split({vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)
            Dim trimmed = line.Trim()
            If trimmed.Length = 0 OrElse trimmed.StartsWith(";") Then Continue For

            Dim rule As New AutoPilotFilterRule()
            rule.IsNegative = False

            If trimmed.StartsWith("EXCLUDE ", StringComparison.OrdinalIgnoreCase) Then
                rule.IsNegative = True
                trimmed = trimmed.Substring(8).Trim()
            End If

            If trimmed.StartsWith("FOLDER ", StringComparison.OrdinalIgnoreCase) Then
                rule.RuleType = AutoPilotFilterRuleType.Folder
                rule.Pattern = trimmed.Substring(7).Trim()
            ElseIf trimmed.StartsWith("@") OrElse trimmed.StartsWith("*@") Then
                rule.RuleType = AutoPilotFilterRuleType.Domain
                rule.Pattern = trimmed
            Else
                rule.RuleType = AutoPilotFilterRuleType.Sender
                rule.Pattern = trimmed
            End If

            rules.Add(rule)
        Next

        Return rules
    End Function

    ''' <summary>Persists the AutoPilot config to My.Settings.</summary>
    ''' <summary>Persists the AutoPilot config to My.Settings.</summary>
    Private Sub SaveAutoPilotConfigToSettings(config As AutoPilotConfig)
        My.Settings.AP_FilterRules = String.Join(vbLf, config.FilterRules.Select(
            Function(r) $"{If(r.IsNegative, "EXCLUDE ", "")}{If(r.RuleType = AutoPilotFilterRuleType.Folder, "FOLDER ", "")}{r.Pattern}"))
        My.Settings.AP_WhitelistedSenders = String.Join(vbLf, config.WhitelistedSenders)
        My.Settings.AP_SubjectTriggerWord = If(config.SubjectTriggerWord, "")
        My.Settings.AP_CooldownSeconds = config.CooldownSeconds
        My.Settings.AP_MaxRepliesPerSession = config.MaxRepliesPerSession
        My.Settings.AP_MaxAttachmentMB = CInt(config.MaxAttachmentBytes / 1024 / 1024)
        My.Settings.AP_FooterText = If(config.FooterText, "")
        My.Settings.AP_RequireApproval = config.RequireApprovalForNonWhitelisted
        My.Settings.AP_MonitoredMailbox = If(config.MonitoredMailbox, "")
        My.Settings.AP_SelectedModelKey = If(config.SelectedModelKey, "")
        My.Settings.AP_UseSecondApi = config.UseSecondApi
        My.Settings.AP_ReprocessLookbackHours = config.ReprocessLookbackHours
        My.Settings.AP_EnableWebGrounding = config.EnableWebGrounding
        My.Settings.AP_EnableVoicemailProcessing = config.EnableVoicemailProcessing
        My.Settings.AP_VoicemailSenderAddress = If(config.VoicemailSenderAddress, "")
        My.Settings.AP_VoicemailCallerIdMapPath = If(config.VoicemailCallerIdMapPath, "")
        My.Settings.AP_EnableScheduler = config.EnableScheduler
        My.Settings.AP_EnableUserMemory = config.EnableUserMemory
        My.Settings.AP_EnableUserFiles = config.EnableUserFiles
        My.Settings.AP_EnablePrivacyProtection = config.EnablePrivacyProtection
        My.Settings.AP_AutoDeleteAfterHours = config.AutoDeleteAfterHours
        My.Settings.AP_SenderToolPolicyPath = If(config.SenderToolPolicyPath, "")
        My.Settings.AP_ThreadRetentionDays = config.ThreadRetentionDays

        ' Persist external tool selection by ToolName/ModelDescription
        If config.SelectedExternalTools IsNot Nothing AndAlso config.SelectedExternalTools.Count > 0 Then
            Dim toolNames = config.SelectedExternalTools.Select(
                Function(t) If(Not String.IsNullOrEmpty(t.ToolName), t.ToolName,
                            If(Not String.IsNullOrEmpty(t.ModelDescription), t.ModelDescription, t.Model))).
                Where(Function(n) Not String.IsNullOrEmpty(n))
            My.Settings.AP_SelectedExternalToolNames = String.Join(vbLf, toolNames)
        Else
            My.Settings.AP_SelectedExternalToolNames = ""
        End If

        SaveAutoPilotSettingsWithRegistryBackup()
    End Sub

    ''' <summary>Loads previously saved config as defaults for the dialog.</summary>
    Private Function LoadAutoPilotConfigDefaults() As AutoPilotConfig
        Dim config As New AutoPilotConfig()
        config.SubjectTriggerWord = If(String.IsNullOrWhiteSpace(My.Settings.AP_SubjectTriggerWord), Nothing, My.Settings.AP_SubjectTriggerWord)
        config.CooldownSeconds = If(My.Settings.AP_CooldownSeconds > 0, CInt(My.Settings.AP_CooldownSeconds), AP_DefaultCooldownSeconds)
        config.MaxRepliesPerSession = If(My.Settings.AP_MaxRepliesPerSession >= 0, My.Settings.AP_MaxRepliesPerSession, AP_DefaultMaxRepliesPerSession)
        config.MaxAttachmentBytes = CLng(If(My.Settings.AP_MaxAttachmentMB > 0, My.Settings.AP_MaxAttachmentMB, 10)) * 1024 * 1024
        config.FooterText = If(My.Settings.AP_FooterText, "")
        config.RequireApprovalForNonWhitelisted = My.Settings.AP_RequireApproval
        config.MonitoredMailbox = If(My.Settings.AP_MonitoredMailbox, "")
        config.SelectedModelKey = If(My.Settings.AP_SelectedModelKey, "")
        config.UseSecondApi = My.Settings.AP_UseSecondApi
        config.ReprocessLookbackHours = If(My.Settings.AP_ReprocessLookbackHours >= 0, My.Settings.AP_ReprocessLookbackHours, 0)
        config.EnableWebGrounding = My.Settings.AP_EnableWebGrounding
        config.EnableVoicemailProcessing = My.Settings.AP_EnableVoicemailProcessing
        config.VoicemailSenderAddress = If(My.Settings.AP_VoicemailSenderAddress, "")
        config.VoicemailCallerIdMapPath = If(My.Settings.AP_VoicemailCallerIdMapPath, "")
        config.EnableScheduler = My.Settings.AP_EnableScheduler
        config.EnableUserMemory = My.Settings.AP_EnableUserMemory
        config.EnableUserFiles = My.Settings.AP_EnableUserFiles
        config.EnablePrivacyProtection = My.Settings.AP_EnablePrivacyProtection
        config.AutoDeleteAfterHours = If(My.Settings.AP_AutoDeleteAfterHours >= 0, My.Settings.AP_AutoDeleteAfterHours, 0)
        config.SenderToolPolicyPath = If(My.Settings.AP_SenderToolPolicyPath, "")
        config.ThreadRetentionDays = If(My.Settings.AP_ThreadRetentionDays >= 0, My.Settings.AP_ThreadRetentionDays, 0)

        ' Restore filter rules using the shared parser
        If Not String.IsNullOrWhiteSpace(My.Settings.AP_FilterRules) Then
            config.FilterRules = ParseFilterRulesText(My.Settings.AP_FilterRules)
        End If

        ' Restore whitelist
        If Not String.IsNullOrWhiteSpace(My.Settings.AP_WhitelistedSenders) Then
            config.WhitelistedSenders = My.Settings.AP_WhitelistedSenders.Split({vbLf}, StringSplitOptions.RemoveEmptyEntries).
                Select(Function(s) s.Trim()).Where(Function(s) s.Length > 0).ToList()
        End If

        ' Restore external tools by matching persisted names against currently available tools
        If Not String.IsNullOrWhiteSpace(My.Settings.AP_SelectedExternalToolNames) Then
            Dim savedToolNames = My.Settings.AP_SelectedExternalToolNames.Split({vbLf}, StringSplitOptions.RemoveEmptyEntries).
                Select(Function(s) s.Trim()).Where(Function(s) s.Length > 0).ToList()
            If savedToolNames.Count > 0 Then
                Try
                    Dim availableTools = GetAvailableTools()
                    If availableTools IsNot Nothing AndAlso availableTools.Count > 0 Then
                        Dim matched = availableTools.Where(
                            Function(t) savedToolNames.Any(Function(n)
                                                               Return String.Equals(n, t.ToolName, StringComparison.OrdinalIgnoreCase) OrElse
                                                                      String.Equals(n, t.ModelDescription, StringComparison.OrdinalIgnoreCase) OrElse
                                                                      String.Equals(n, t.Model, StringComparison.OrdinalIgnoreCase)
                                                           End Function)).ToList()
                        If matched.Count > 0 Then config.SelectedExternalTools = matched
                    End If
                Catch
                End Try
            End If
        End If

        Return config
    End Function

    ''' <summary>
    ''' Returns True if there is a previously saved AutoPilot configuration with at least
    ''' a filter rule or monitored mailbox — i.e., the user has run AutoPilot before.
    ''' </summary>
    Private Function HasSavedAutoPilotConfig() As Boolean
        ' We consider a config "saved" if at least the filter rules or monitored mailbox
        ' were persisted, indicating a prior completed config dialog run.
        Return Not String.IsNullOrWhiteSpace(My.Settings.AP_FilterRules) OrElse
               Not String.IsNullOrWhiteSpace(My.Settings.AP_MonitoredMailbox) OrElse
               Not String.IsNullOrWhiteSpace(My.Settings.AP_WhitelistedSenders)
    End Function

    ''' <summary>
    ''' Attempts to auto-start AutoPilot using the last saved configuration.
    ''' Shows a 30-second countdown dialog allowing the user to cancel or configure manually.
    ''' Called from <see cref="DelayedStartupTasks"/> when <c>INI_AutoPilotAutoStart</c> is True.
    ''' </summary>
    Public Sub TryAutoStartAutoPilot()
        Try
            ' Gate: INI flag must be enabled and user must be licensed/permitted
            If Not INI_AutoPilotAutoStart Then Return
            If Not IsAutoPilotLicenseValid() Then Return
            If Not IsAutoPilotPermitted() Then Return
            If _apActive Then Return

            EnsureAutoPilotSettingsRestoredFromRegistry()

            If Not HasSavedAutoPilotConfig() Then Return

            ' Load the last saved configuration
            Dim config = LoadAutoPilotConfigDefaults()

            ' Build a human-readable summary for the countdown dialog
            Dim modelLabel As String
            If Not String.IsNullOrWhiteSpace(config.SelectedModelKey) Then
                modelLabel = config.SelectedModelKey
            ElseIf config.UseSecondApi Then
                modelLabel = INI_Model_2
            Else
                modelLabel = INI_Model
            End If

            Dim summary As New System.Text.StringBuilder()
            summary.AppendLine($"{AN6} AutoPilot will start automatically with the last used settings:")
            summary.AppendLine()
            summary.AppendLine($"Model: {modelLabel}")
            summary.AppendLine($"Filters: {config.FilterRules.Count} rule(s)")
            summary.AppendLine($"Whitelisted senders: {config.WhitelistedSenders.Count}")
            summary.AppendLine($"Cooldown: {config.CooldownSeconds}s | Max replies: {If(config.MaxRepliesPerSession = 0, "unlimited", config.MaxRepliesPerSession.ToString())}")
            summary.AppendLine($"Auto-delete: {If(config.AutoDeleteAfterHours > 0, $"{config.AutoDeleteAfterHours}h", "disabled")}")
            If Not String.IsNullOrWhiteSpace(config.MonitoredMailbox) Then
                summary.AppendLine($"Mailbox: {config.MonitoredMailbox}")
            End If
            summary.AppendLine()
            summary.AppendLine("Press 'Configure' to open the full setup dialog instead.")

            ' Show a 30-second auto-close dialog (returns 1=Start, 2=Configure, 3=auto-close)
            Dim choice = ShowCustomYesNoBox(
                summary.ToString().TrimEnd(),
                $"Start {AN6} AutoPilot", "Configure",
                header:=$"{AN6} AutoPilot — Auto Start",
                autoCloseSeconds:=30,
                Defaulttext:=$"AutoPilot will start automatically")

            If choice = 2 Then
                ' User chose to configure manually — fall through to the normal dialog
                StartAutoPilot()
                Return
            End If

            If choice = 0 Then
                ' Dialog was closed/cancelled — do not start
                Return
            End If

            ' choice = 1 (user clicked Start) → attended; choice = 3 (timer expired) → unattended
            Dim unattended As Boolean = (choice = 3)
            config.IsUnattended = unattended

            ' Apply the saved model configuration to the context before starting.
            ' This replicates what ShowModelSelection does interactively.
            If config.UseSecondApi AndAlso Not String.IsNullOrWhiteSpace(config.SelectedModelKey) Then
                ' Alternate model: find and apply it from the INI file
                If INI_SecondAPI AndAlso Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                    Try
                        Dim models = LoadAlternativeModels(INI_AlternateModelPath, _context)
                        If models IsNot Nothing Then
                            Dim match = models.FirstOrDefault(
                                Function(m) String.Equals(
                                    If(Not String.IsNullOrWhiteSpace(m.ModelDescription), m.ModelDescription, m.Model),
                                    config.SelectedModelKey, StringComparison.OrdinalIgnoreCase))
                            If match IsNot Nothing Then
                                originalConfig = GetCurrentConfig(_context)
                                originalConfigLoaded = True
                                ApplyModelConfig(_context, match)
                                LastAlternateModel = config.SelectedModelKey
                            End If
                        End If
                    Catch
                        ' If model application fails, proceed with default — StartAutoPilotWithConfig
                        ' will capture whatever config is active at that point
                    End Try
                End If
            ElseIf config.UseSecondApi Then
                ' Secondary API without alternate model path — just ensure SecondAPI is active
                ' (the context should already have the secondary config from INI)
            End If

            SaveAutoPilotConfigToSettings(config)
            StartAutoPilotWithConfig(config)

        Catch ex As System.Exception
            ' Auto-start is best-effort; do not block Outlook startup on failure
            Debug.WriteLine($"AutoPilot auto-start failed: {ex.Message}")
        End Try
    End Sub

End Class
