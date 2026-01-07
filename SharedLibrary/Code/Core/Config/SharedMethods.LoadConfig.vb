' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.LoadConfig.vb
' Purpose: Loads and validates runtime configuration from an `.ini` file and `My.Settings`,
'          then initializes the shared execution context used by the add-in.
'
' Architecture:
'  - Configuration Source Resolution:
'     - Determines the `.ini` location based on registry settings and default locations
'       returned by `GetDefaultINIPath(...)`.
'     - On first start, can invoke `InitialConfig` to guide creation of an `.ini` file.
'  - INI Parsing:
'     - Reads the `.ini` file as text, ignores empty lines and lines starting with `;`,
'       parses `key=value` pairs into a case-insensitive dictionary.
'  - Context Population:
'     - Copies values from the dictionary into `context.INI_*` properties, applying defaults
'       for selected prompt templates and optional settings.
'     - Loads some per-user overrides from `My.Settings`.
'  - Optional Feature Blocks:
'     - Internet search (`INI_ISearch`), RAG/library (`INI_Lib`), Second API (`INI_SecondAPI`),
'       OAuth2 (`INI_OAuth2` and `INI_OAuth2_2`).
'  - License Gate:
'     - Calls `LicenseOK(context, configDict)` and aborts initialization when it returns `False`.
'  - Key Handling:
'     - If API key encryption is enabled, uses `Codebasis` and `DecodeString` via `RealAPIKey(...)`
'       to derive usable secrets for runtime.
'  - Missing Mandatory Values:
'     - `INIValuesMissing` checks required settings and can use `MissingSettingsWindow` to allow
'       interactive completion and saving via `UpdateAppConfig(context)`.
' =============================================================================


Option Strict On
Option Explicit On

Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary.SharedContext
Imports UglyToad.PdfPig.Graphics.Operations.PathPainting

Namespace SharedLibrary
    Partial Public Class SharedMethods

        ''' <summary>
        ''' Initializes configuration and populates <paramref name="context"/> with values from the `.ini` file
        ''' and selected values from <c>My.Settings</c>.
        ''' </summary>
        ''' <param name="context">Shared context to populate and mark as loaded.</param>
        ''' <param name="FirstTime">If <c>True</c>, allows showing the initial setup wizard when no `.ini` is found.</param>
        ''' <param name="Reload">If <c>True</c>, forces reload even when <c>context.INIloaded</c> is already set.</param>
        Public Shared Sub InitializeConfig(ByRef context As ISharedContext, FirstTime As Boolean, Reload As Boolean)

            If context.INIloaded AndAlso Not Reload Then Return

            context.GPTSetupError = True

            context.INIloaded = False

            Dim IniFilePath As String = ""
            Dim RegFilePath As String = ""
            Dim DefaultPath As String = ""
            Dim DefaultPath2 As String = ""

            Try

                ' Determine the configuration file path.

                RegFilePath = GetFromRegistry(RegPath_Base, RegPath_IniPath, True)
                DefaultPath = GetDefaultINIPath(context.RDV)
                DefaultPath2 = GetDefaultINIPath("Word")

                If Not String.IsNullOrWhiteSpace(RegFilePath) AndAlso RegPath_IniPrio Then
                    IniFilePath = System.IO.Path.Combine(ExpandEnvironmentVariables(RegFilePath), $"{AN2}.ini")
                ElseIf System.IO.File.Exists(DefaultPath) Then
                    IniFilePath = DefaultPath
                ElseIf System.IO.File.Exists(DefaultPath2) Then
                    IniFilePath = DefaultPath2
                ElseIf Not String.IsNullOrWhiteSpace(RegFilePath) Then
                    IniFilePath = System.IO.Path.Combine(ExpandEnvironmentVariables(RegFilePath), $"{AN2}.ini")
                Else
                    IniFilePath = DefaultPath
                End If

                IniFilePath = RemoveCR(IniFilePath)

                ' Check if the configuration file exists.

                If Not System.IO.File.Exists(IniFilePath) Then
                    If FirstTime Then
                        Using frm As New InitialConfig(context)
                            frm.ShowDialog()
                        End Using
                        IniFilePath = DefaultPath
                        If context.InitialConfigFailed AndAlso Not System.IO.File.Exists(IniFilePath) Then
                            ShowCustomMessageBox($"You have aborted the setup wizard and no configuration file has been found ('{IniFilePath}'). You will have to retry or configure it manually to use {AN}, even if you see the menus (they will disappear once {AN} has been de-installed or de-activated).")
                            Return
                        End If
                        If Not System.IO.File.Exists(IniFilePath) Then
                            ShowCustomMessageBox($"The configuration file is (still) not found ('{IniFilePath}'). There may be an error in the setup assistant. Please configure the configuration file manually.")
                            Return
                        End If
                    Else
                        ShowCustomMessageBox($"The configuration file has not been found ('{IniFilePath}').")
                        Return
                    End If
                End If

                Dim iniContent As String = ""
                Dim configDict As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

                ' Read and parse the .ini file.
                iniContent = System.IO.File.ReadAllText(IniFilePath)
                Dim iniLines As String() = iniContent.Split({vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
                For Each line As String In iniLines
                    Dim trimmedLine = line.Trim()
                    If Not String.IsNullOrEmpty(trimmedLine) AndAlso Not trimmedLine.StartsWith(";") Then ' Skip comments and empty lines.
                        Dim keyValue = trimmedLine.Split(New Char() {"="c}, 2)
                        If keyValue.Length = 2 Then
                            configDict(keyValue(0).Trim()) = keyValue(1).Trim()
                        End If
                    End If
                Next

                ' Assign and validate configuration values.
                context.INI_APIKey = If(configDict.ContainsKey("APIKey"), configDict("APIKey"), "")
                context.INI_Endpoint = If(configDict.ContainsKey("Endpoint"), configDict("Endpoint"), "")
                context.INI_HeaderA = If(configDict.ContainsKey("HeaderA"), configDict("HeaderA"), "")
                context.INI_HeaderB = If(configDict.ContainsKey("HeaderB"), configDict("HeaderB"), "")
                context.INI_Response = If(configDict.ContainsKey("Response"), configDict("Response"), "")
                context.INI_Anon = If(configDict.ContainsKey("Anon"), configDict("Anon"), "")
                context.INI_TokenCount = If(configDict.ContainsKey("TokenCount"), configDict("TokenCount"), "")
                context.INI_APICall = If(configDict.ContainsKey("APICall"), configDict("APICall"), "")
                context.INI_APICall_Object = If(configDict.ContainsKey("APICall_Object"), configDict("APICall_Object"), "")
                context.INI_Timeout = If(configDict.ContainsKey("Timeout"), CLng(configDict("Timeout")), DEFAULT_TIMEOUT_LONG)
                context.INI_MaxOutputToken = If(configDict.ContainsKey("MaxOutputToken"), CInt(configDict("MaxOutputToken")), DEFAULT_MAX_OUTPUT_TOKEN)
                context.INI_Temperature = If(configDict.ContainsKey("Temperature"), configDict("Temperature"), DEFAULT_TEMPERATURE)
                context.INI_Model = If(configDict.ContainsKey("Model"), configDict("Model"), "")

                context.SP_Translate = If(configDict.ContainsKey("SP_Translate"), configDict("SP_Translate"), Default_SP_Translate)
                context.SP_Correct = If(configDict.ContainsKey("SP_Correct"), configDict("SP_Correct"), Default_SP_Correct)
                context.SP_Improve = If(configDict.ContainsKey("SP_Improve"), configDict("SP_Improve"), Default_SP_Improve)
                context.SP_Explain = If(configDict.ContainsKey("SP_Explain"), configDict("SP_Explain"), Default_SP_Explain)
                context.SP_FindClause = If(configDict.ContainsKey("SP_FindClause"), configDict("SP_FindClause"), Default_SP_FindClause)
                context.SP_FindClause_Clean = If(configDict.ContainsKey("SP_FindClause_Clean"), configDict("SP_FindClause_Clean"), Default_SP_FindClause_Clean)
                context.SP_ApplyDocStyle = If(configDict.ContainsKey("SP_ApplyDocStyle"), configDict("SP_ApplyDocStyle"), Default_SP_ApplyDocStyle)
                context.SP_ApplyDocStyle_NumberingHint = If(configDict.ContainsKey("SP_ApplyDocStyle_NumberingHint"), configDict("SP_ApplyDocStyle_NumberingHint"), Default_SP_ApplyDocStyle_NumberingHint)
                context.SP_DocCheck_Clause = If(configDict.ContainsKey("SP_DocCheck_Clause"), configDict("SP_DocCheck_Clause"), Default_SP_DocCheck_Clause)
                context.SP_DocCheck_MultiClause = If(configDict.ContainsKey("SP_DocCheck_MultiClause"), configDict("SP_DocCheck_MultiClause"), Default_SP_DocCheck_MultiClause)
                context.SP_DocCheck_MultiClauseSum = If(configDict.ContainsKey("SP_DocCheck_MultiClauseSum"), configDict("SP_DocCheck_MultiClauseSum"), Default_SP_DocCheck_MulticlauseSum)
                context.SP_DocCheck_MultiClauseSum_Bubbles = If(configDict.ContainsKey("SP_DocCheck_MultiClauseSum_Bubbles"), configDict("SP_DocCheck_MultiClauseSum_Bubbles"), Default_SP_DocCheck_MultiClauseSum_Bubbles)
                context.SP_SuggestTitles = If(configDict.ContainsKey("SP_SuggestTitles"), configDict("SP_SuggestTitles"), Default_SP_SuggestTitles)
                context.SP_Friendly = If(configDict.ContainsKey("SP_Friendly"), configDict("SP_Friendly"), Default_SP_Friendly)
                context.SP_Convincing = If(configDict.ContainsKey("SP_Convincing"), configDict("SP_Convincing"), Default_SP_Convincing)
                context.SP_NoFillers = If(configDict.ContainsKey("SP_NoFillers"), configDict("SP_NoFillers"), Default_SP_NoFillers)
                context.SP_Podcast = If(configDict.ContainsKey("SP_Podcast"), configDict("SP_Podcast"), Default_SP_Podcast)
                context.SP_MyStyle_Word = If(configDict.ContainsKey("SP_MyStyle_Word"), configDict("SP_MyStyle_Word"), Default_SP_MyStyle_Word)
                context.SP_MyStyle_Outlook = If(configDict.ContainsKey("SP_MyStyle_Outlook"), configDict("SP_MyStyle_Outlook"), Default_SP_MyStyle_Outlook)
                context.SP_MyStyle_Apply = If(configDict.ContainsKey("SP_MyStyle_Apply"), configDict("SP_MyStyle_Apply"), Default_SP_MyStyle_Apply)
                context.SP_Shorten = If(configDict.ContainsKey("SP_Shorten"), configDict("SP_Shorten"), Default_SP_Shorten)
                context.SP_Filibuster = If(configDict.ContainsKey("SP_Filibuster"), configDict("SP_Filibuster"), Default_SP_Filibuster)
                context.SP_ArgueAgainst = If(configDict.ContainsKey("SP_ArgueAgainst"), configDict("SP_ArgueAgainst"), Default_SP_ArgueAgainst)
                context.SP_InsertClipboard = If(configDict.ContainsKey("SP_InsertClipboard"), configDict("SP_InsertClipboard"), Default_SP_InsertClipboard)
                context.SP_Summarize = If(configDict.ContainsKey("SP_Summarize"), configDict("SP_Summarize"), Default_SP_Summarize)
                context.SP_Markup = If(configDict.ContainsKey("SP_Markup"), configDict("SP_Markup"), Default_SP_Markup)
                context.SP_FreestyleText = If(configDict.ContainsKey("SP_FreestyleText"), configDict("SP_FreestyleText"), Default_SP_FreestyleText)
                context.SP_FreestyleNoText = If(configDict.ContainsKey("SP_FreestyleNoText"), configDict("SP_FreestyleNoText"), Default_SP_FreestyleNoText)
                context.SP_MailReply = If(configDict.ContainsKey("SP_MailReply"), configDict("SP_MailReply"), Default_SP_MailReply)
                context.SP_MailSumup = If(configDict.ContainsKey("SP_MailSumup"), configDict("SP_MailSumup"), Default_SP_MailSumup)
                context.SP_MailSumup2 = If(configDict.ContainsKey("SP_MailSumup2"), configDict("SP_MailSumup2"), Default_SP_MailSumup2)
                context.SP_SwitchParty = If(configDict.ContainsKey("SP_SwitchParty"), configDict("SP_SwitchParty"), Default_SP_SwitchParty)
                context.SP_Anonymize = If(configDict.ContainsKey("SP_Anonymize"), configDict("SP_Anonymize"), Default_SP_Anonymize)
                context.SP_Rename = If(configDict.ContainsKey("SP_Rename"), configDict("SP_Rename"), Default_SP_Rename)
                context.SP_Redact = If(configDict.ContainsKey("SP_Redact"), configDict("SP_Redact"), Default_SP_Redact)
                context.SP_CheckforII = If(configDict.ContainsKey("SP_CheckforII"), configDict("SP_CheckforII"), Default_SP_CheckforII)
                context.SP_Extract = If(configDict.ContainsKey("SP_Extract"), configDict("SP_Extract"), Default_SP_Extract)
                context.SP_ExtractSchema = If(configDict.ContainsKey("SP_ExtractSchema"), configDict("SP_ExtractSchema"), Default_SP_ExtractSchema)
                context.SP_MergeDateRows = If(configDict.ContainsKey("SP_MergeDateRows"), configDict("SP_MergeDateRows"), Default_SP_MergeDateRows)
                context.SP_ContextSearch = If(configDict.ContainsKey("SP_ContextSearch"), configDict("SP_ContextSearch"), Default_SP_ContextSearch)
                context.SP_ContextSearchMulti = If(configDict.ContainsKey("SP_ContextSearchMulti"), configDict("SP_ContextSearchMulti"), Default_SP_ContextSearchMulti)
                context.SP_RangeOfCells = If(configDict.ContainsKey("SP_RangeOfCells"), configDict("SP_RangeOfCells"), Default_SP_RangeOfCells)
                context.SP_ParseFile = If(configDict.ContainsKey("SP_ParseFile"), configDict("SP_ParseFile"), Default_SP_ParseFile)
                context.SP_WriteNeatly = If(configDict.ContainsKey("SP_WriteNeatly"), configDict("SP_WriteNeatly"), Default_SP_WriteNeatly)
                context.SP_Add_KeepFormulasIntact = If(configDict.ContainsKey("SP_Add_KeepFormulasIntact"), configDict("SP_Add_KeepFormulasIntact"), Default_SP_Add_KeepFormulasIntact)
                context.SP_Add_KeepHTMLIntact = If(configDict.ContainsKey("SP_Add_KeepHTMLIntact"), configDict("SP_Add_KeepHTMLIntact"), Default_SP_Add_KeepHTMLIntact)
                context.SP_Add_KeepInlineIntact = If(configDict.ContainsKey("SP_Add_KeepInlineIntact"), configDict("SP_Add_KeepInlineIntact"), Default_SP_Add_KeepInlineIntact)
                context.SP_Add_Tooling = If(configDict.ContainsKey("SP_Add_Tooling"), configDict("SP_Add_Tooling"), Default_SP_Add_Tooling)
                context.SP_Add_Bubbles = If(configDict.ContainsKey("SP_Add_Bubbles"), configDict("SP_Add_Bubbles"), Default_SP_Add_Bubbles)
                context.SP_Add_BubblesExtract = If(configDict.ContainsKey("SP_Add_BubblesExtract"), configDict("SP_Add_BubblesExtract"), Default_SP_Add_BubblesExtract)
                context.SP_Add_BubblesReply = If(configDict.ContainsKey("SP_Add_BubblesReply"), configDict("SP_Add_BubblesReply"), Default_SP_Add_BubblesReply)
                context.SP_Add_Bubbles_Format = If(configDict.ContainsKey("SP_Add_Bubbles_Format"), configDict("SP_Add_Bubbles_Format"), Default_SP_Add_Bubbles_Format)
                context.SP_Add_Batch = If(configDict.ContainsKey("SP_Add_Batch"), configDict("SP_Add_Batch"), Default_SP_Add_Batch)
                context.SP_Add_Slides = If(configDict.ContainsKey("SP_Add_Slides"), configDict("SP_Add_Slides"), Default_SP_Add_Slides)
                context.SP_BubblesExcel = If(configDict.ContainsKey("SP_BubblesExcel"), configDict("SP_BubblesExcel"), Default_SP_BubblesExcel)
                context.SP_Add_Revisions = If(configDict.ContainsKey("SP_Add_Revisions"), configDict("SP_Add_Revisions"), Default_SP_Add_Revisions)
                context.SP_MarkupRegex = If(configDict.ContainsKey("SP_MarkupRegex"), configDict("SP_MarkupRegex"), Default_SP_MarkupRegex)
                context.SP_ChatWord = If(configDict.ContainsKey("SP_ChatWord"), configDict("SP_ChatWord"), Default_SP_ChatWord)
                context.SP_HelpMe = If(configDict.ContainsKey("SP_HelpMe"), configDict("SP_HelpMe"), Default_SP_HelpMe)
                context.SP_DiscussThis_SortOut = If(configDict.ContainsKey("SP_DiscussThis_SortOut"), configDict("SP_DiscussThis_SortOut"), Default_SP_DiscussThis_SortOut)
                context.SP_Chat = If(configDict.ContainsKey("SP_Chat"), configDict("SP_Chat"), Default_SP_Chat)
                context.SP_Add_ChatWord_Commands = If(configDict.ContainsKey("SP_Add_ChatWord_Commands"), configDict("SP_Add_ChatWord_Commands"), Default_SP_Add_ChatWord_Commands)
                context.SP_Add_Chat_NoCommands = If(configDict.ContainsKey("SP_Add_Chat_NoCommands"), configDict("SP_Add_Chat_NoCommands"), Default_SP_Add_Chat_NoCommands)
                context.SP_ChatExcel = If(configDict.ContainsKey("SP_ChatExcel"), configDict("SP_ChatExcel"), Default_SP_ChatExcel)
                context.SP_Add_ChatExcel_Commands = If(configDict.ContainsKey("SP_Add_ChatExcel_Commands"), configDict("SP_Add_ChatExcel_Commands"), Default_SP_Add_ChatExcel_Commands)
                context.SP_MergePrompt = If(configDict.ContainsKey("SP_MergePrompt"), configDict("SP_MergePrompt"), Default_SP_MergePrompt)
                context.SP_MergePrompt2 = If(configDict.ContainsKey("SP_MergePrompt2"), configDict("SP_MergePrompt2"), Default_SP_MergePrompt2)
                context.SP_Add_MergePrompt = If(configDict.ContainsKey("SP_Add_MergePrompt"), configDict("SP_Add_MergePrompt"), Default_SP_Add_MergePrompt)
                context.SP_FindPrompts = If(configDict.ContainsKey("SP_FindPrompts"), configDict("SP_FindPrompts"), Default_SP_FindPrompts)
                context.SP_Ignore = If(configDict.ContainsKey("SP_Ignore"), configDict("SP_Ignore"), Default_SP_Ignore)

                ' Legacy; was required For Excel Helper.
                ' context.INI_OpenSSLPath = If(configDict.ContainsKey("OpenSSLPath"), configDict("OpenSSLPath"), "%APPDATA%\Microsoft\OpenSSL_Runtime\openssl.exe")

                ' Optional values.
                context.INI_PreCorrection = If(configDict.ContainsKey("PreCorrection"), configDict("PreCorrection"), "")
                context.INI_PostCorrection = If(configDict.ContainsKey("PostCorrection"), configDict("PostCorrection"), "")
                context.INI_APIKeyPrefix = If(configDict.ContainsKey("APIKeyPrefix"), configDict("APIKeyPrefix"), "")
                context.INI_UsageRestrictions = If(configDict.ContainsKey("UsageRestrictions"), configDict("UsageRestrictions"), "")
                context.INI_LogPath = If(configDict.ContainsKey("LogPath"), configDict("LogPath"), "")
                context.INI_Language1 = If(configDict.ContainsKey("Language1"), configDict("Language1"), DEFAULT_LANGUAGE_1)
                context.INI_Language2 = If(configDict.ContainsKey("Language2"), configDict("Language2"), DEFAULT_LANGUAGE_2)
                context.INI_KeepFormatCap = If(configDict.ContainsKey("KeepFormatCap"), CInt(configDict("KeepFormatCap")), DEFAULT_KEEPFORMAT_CAP)
                context.INI_MarkupMethodHelper = If(configDict.ContainsKey("MarkupMethodHelper"), CInt(configDict("MarkupMethodHelper")), DEFAULT_MARKUP_METHOD_HELPER)
                context.INI_MarkupMethodWord = If(configDict.ContainsKey("MarkupMethodWord"), CInt(configDict("MarkupMethodWord")), DEFAULT_MARKUP_METHOD_WORD)
                context.INI_MarkupMethodOutlook = If(configDict.ContainsKey("MarkupMethodOutlook"), CInt(configDict("MarkupMethodOutlook")), DEFAULT_MARKUP_METHOD_OUTLOOK)
                context.INI_MarkupDiffCap = If(configDict.ContainsKey("MarkupDiffCap"), CInt(configDict("MarkupDiffCap")), DEFAULT_MARKUP_DIFF_CAP)
                context.INI_MarkupRegexCap = If(configDict.ContainsKey("MarkupRegexCap"), CInt(configDict("MarkupRegexCap")), DEFAULT_MARKUP_REGEX_CAP)
                context.INI_ChatCap = If(configDict.ContainsKey("ChatCap"), CInt(configDict("ChatCap")), DEFAULT_CHAT_CAP)

                ' Load per-user overrides from My.Settings.
                context.INI_DefaultPrefix = My.Settings.DefaultPrefix
                context.INI_ReplaceText2Override = My.Settings.ReplaceText2Override
                context.INI_MarkupMethodWordOverride = My.Settings.MarkupMethodWordOverride
                context.INI_MarkupMethodOutlookOverride = My.Settings.MarkupMethodOutlookOverride

                ' Boolean parameters.
                context.INI_DoubleS = ParseBoolean(configDict, "DoubleS")
                context.INI_Clean = ParseBoolean(configDict, "Clean")
                context.INI_Ignore = ParseBoolean(configDict, "Ignore")
                context.INI_NoDash = ParseBoolean(configDict, "NoEmDash")
                context.INI_MarkdownBubbles = ParseBoolean(configDict, "MarkdownBubbles")
                context.INI_KeepFormat1 = ParseBoolean(configDict, "KeepFormat1")
                context.INI_ReplaceText1 = ParseBoolean(configDict, "ReplaceText1", DEFAULT_BOOL_REPLACETEXT1)
                context.INI_KeepFormat2 = ParseBoolean(configDict, "KeepFormat2")
                context.INI_MarkdownConvert = ParseBoolean(configDict, "MarkdownConvert", DEFAULT_BOOL_MARKDOWNCONVERT)
                context.INI_KeepParaFormatInline = ParseBoolean(configDict, "KeepParaFormatInline")
                context.INI_ReplaceText2 = ParseBoolean(configDict, "ReplaceText2", DEFAULT_BOOL_REPLACETEXT2)
                context.INI_DoMarkupOutlook = ParseBoolean(configDict, "DoMarkupOutlook", DEFAULT_BOOL_DOMARKUPOUTLOOK)
                context.INI_DoMarkupWord = ParseBoolean(configDict, "DoMarkupWord", DEFAULT_BOOL_DOMARKUPWORD)
                context.INI_RoastMe = ParseBoolean(configDict, "RoastMe", False)
                context.INI_APIDebug = ParseBoolean(configDict, "APIDebug")
                context.INI_APIEncrypted = ParseBoolean(configDict, "APIKeyEncrypted")
                context.INI_ShortcutsWordExcel = If(configDict.ContainsKey("ShortcutsWordExcel"), configDict("ShortcutsWordExcel"), "")
                context.INI_ContextMenu = ParseBoolean(configDict, "ContextMenu", DEFAULT_BOOL_CONTEXTMENU)

                ' Tooling settings

                context.INI_ToolingLogWindow = ParseBoolean(configDict, "ToolingLogWindow", DEFAULT_BOOL_TOOLINGLOGWINDOW)
                context.INI_ToolingDryRun = ParseBoolean(configDict, "ToolingDryRun")
                context.INI_ToolingMaximumIterations = If(configDict.ContainsKey("ToolingMaximumIterations"), CInt(configDict("ToolingMaximumIterations")), DEFAULT_TOOLING_MAXIMUMITERATIONS)

                ' Other parameters.

                context.INI_NoHelperDownload = ParseBoolean(configDict, "NoHelperDownload")
                context.INI_UpdateCheckInterval = If(configDict.ContainsKey("UpdateCheckInterval"), CInt(configDict("UpdateCheckInterval")), DefaultUpdateIntervalDays)
                context.INI_UpdatePath = If(configDict.ContainsKey("UpdatePath"), configDict("UpdatePath"), "")
                context.INI_UpdateIni = ParseBoolean(configDict, "UpdateIni", DEFAULT_BOOL_UPDATEINI)
                context.INI_UpdateIniAllowRemote = ParseBoolean(configDict, "UpdateIniAllowRemote", DEFAULT_BOOL_UPDATEINI_ALLOWREMOTE)
                context.INI_UpdateIniNoSignature = ParseBoolean(configDict, "UpdateIniNoSignature", False)
                context.INI_UpdateSource = If(configDict.ContainsKey("UpdateSource"), configDict("UpdateSource"), "")

                ' First parameter Is legacy
                ' context.INI_UpdateIniClients = If(configDict.ContainsKey("UpdateClients"), configDict("UpdateClients"), "")
                context.INI_UpdateIniClients = If(configDict.ContainsKey("UpdateIniClients"), configDict("UpdateIniClients"), "")

                context.INI_UpdateIniIgnoreOverride = If(configDict.ContainsKey("UpdateIniIgnoreOverride"), configDict("UpdateIniIgnoreOverride"), "")
                context.INI_UpdateIniSilentMode = If(configDict.ContainsKey("UpdateIniSilentMode"), CInt(configDict("UpdateIniSilentMode")), DEFAULT_UPDATE_INI_SILENT_MODE)
                context.INI_UpdateIniSilentLog = ParseBoolean(configDict, "UpdateIniSilentLog", DEFAULT_BOOL_UPDATEINISILENTLOG)

                context.INI_HelpMeInkyPath = If(configDict.ContainsKey("HelpMeInkyPath"), configDict("HelpMeInkyPath"), Default_HelpMeInkyPath)
                context.INI_DiscussInkyPath = If(configDict.ContainsKey("DiscussInkyPath"), configDict("DiscussInkyPath"), "")
                context.INI_DiscussInkyPathLocal = If(configDict.ContainsKey("DiscussInkyPathLocal"), configDict("DiscussInkyPathLocal"), "")
                context.INI_RedactionInstructionsPath = If(configDict.ContainsKey("RedactionInstructionsPath"), configDict("RedactionInstructionsPath"), "")
                context.INI_RedactionInstructionsPathLocal = If(configDict.ContainsKey("RedactionInstructionsPathLocal"), configDict("RedactionInstructionsPathLocal"), "")
                context.INI_ExtractorPath = If(configDict.ContainsKey("ExtractorPath"), configDict("ExtractorPath"), "")
                context.INI_ExtractorPathLocal = If(configDict.ContainsKey("ExtractorPathLocal"), configDict("ExtractorPathLocal"), "")
                context.INI_RenameLibPath = If(configDict.ContainsKey("RenameLibPath"), configDict("RenameLibPath"), "")
                context.INI_RenameLibPathLocal = If(configDict.ContainsKey("RenameLibPathLocal"), configDict("RenameLibPathLocal"), "")

                context.INI_Location = If(configDict.ContainsKey("Location"), configDict("Location"), "")

                context.INI_SpeechModelPath = If(configDict.ContainsKey("SpeechModelPath"), configDict("SpeechModelPath"), "")
                context.INI_TTSEndpoint = If(configDict.ContainsKey("TTSEndpoint"), configDict("TTSEndpoint"), "")
                context.INI_LocalModelPath = If(configDict.ContainsKey("LocalModelPath"), configDict("LocalModelPath"), "")

                context.INI_PromptLibPath = If(configDict.ContainsKey("PromptLib"), configDict("PromptLib"), "")
                context.INI_PromptLibPathLocal = If(configDict.ContainsKey("PromptLibLocal"), configDict("PromptLibLocal"), "")
                context.INI_MyStylePath = If(configDict.ContainsKey("MyStylePath"), configDict("MyStylePath"), "")
                context.INI_AlternateModelPath = If(configDict.ContainsKey("AlternateModelPath"), configDict("AlternateModelPath"), "")
                context.INI_SpecialServicePath = If(configDict.ContainsKey("SpecialServicePath"), configDict("SpecialServicePath"), "")
                context.INI_WebAgentPath = If(configDict.ContainsKey("WebAgentPath"), configDict("WebAgentPath"), "")
                context.INI_WebAgentPathLocal = If(configDict.ContainsKey("WebAgentPathLocal"), configDict("WebAgentPathLocal"), "")
                context.INI_FindClausePath = If(configDict.ContainsKey("FindClausePath"), configDict("FindClausePath"), "")
                context.INI_FindClausePathLocal = If(configDict.ContainsKey("FindClausePathLocal"), configDict("FindClausePathLocal"), "")
                context.INI_DocCheckPath = If(configDict.ContainsKey("DocCheckPath"), configDict("DocCheckPath"), "")
                context.INI_DocCheckPathLocal = If(configDict.ContainsKey("DocCheckPathLocal"), configDict("DocCheckPathLocal"), "")
                context.INI_DocStylePath = If(configDict.ContainsKey("DocStylePath"), configDict("DocStylePath"), "")
                context.INI_DocStylePathLocal = If(configDict.ContainsKey("DocStylePathLocal"), configDict("DocStylePathLocal"), "")
                context.INI_PromptLibPath_Transcript = If(configDict.ContainsKey("PromptLib_Transcript"), configDict("PromptLib_Transcript"), "")

                ' Process Internet search if enabled.
                context.INI_ISearch = ParseBoolean(configDict, "ISearch", DEFAULT_BOOL_ISEARCH_ENABLED)
                If context.INI_ISearch Then
                    context.INI_ISearch_Approve = ParseBoolean(configDict, "ISearch_Approve", False)
                    context.INI_ISearch_URL = If(configDict.ContainsKey("ISearch_URL"), configDict("ISearch_URL"), DEFAULT_ISEARCH_URL)
                    context.INI_ISearch_ResponseMask1 = If(configDict.ContainsKey("ISearch_ResponseMask1"), configDict("ISearch_ResponseMask1"), DEFAULT_ISEARCH_RESPONSE_MASK_1)
                    context.INI_ISearch_ResponseMask2 = If(configDict.ContainsKey("ISearch_ResponseMask2"), configDict("ISearch_ResponseMask2"), DEFAULT_ISEARCH_RESPONSE_MASK_2)
                    context.INI_ISearch_Name = If(configDict.ContainsKey("ISearch_Name"), configDict("ISearch_Name"), DEFAULT_ISEARCH_NAME)
                    context.INI_ISearch_Tries = If(configDict.ContainsKey("ISearch_Tries"), CInt(configDict("ISearch_Tries")), ISearch_DefTries)
                    context.INI_ISearch_Results = If(configDict.ContainsKey("ISearch_Results"), CInt(configDict("ISearch_Results")), ISearch_DefResults)
                    context.INI_ISearch_MaxDepth = If(configDict.ContainsKey("ISearch_MaxDepth"), CInt(configDict("ISearch_MaxDepth")), ISearch_DefMaxDepth)
                    context.INI_ISearch_Timeout = If(configDict.ContainsKey("ISearch_Timeout"), CLng(configDict("ISearch_Timeout")), ISearch_DefSearchTimeout)
                    context.INI_ISearch_SearchTerm_SP = If(configDict.ContainsKey("ISearch_SearchTerm_SP"), configDict("ISearch_SearchTerm_SP"), Default_INI_ISearch_SearchTerm_SP)
                    context.INI_ISearch_Apply_SP = If(configDict.ContainsKey("ISearch_Apply_SP"), configDict("ISearch_Apply_SP"), Default_INI_ISearch_Apply_SP)
                    context.INI_ISearch_Apply_SP_Markup = If(configDict.ContainsKey("ISearch_Apply_SP_Markup"), configDict("ISearch_Apply_SP_Markup"), Default_INI_ISearch_Apply_SP_Markup)
                    If context.INI_ISearch_Tries > ISearch_MaxTries Then context.INI_ISearch_Tries = ISearch_MaxTries
                    If context.INI_ISearch_Results > ISearch_MaxResults Then context.INI_ISearch_Results = ISearch_MaxResults
                    If context.INI_ISearch_MaxDepth > ISearch_MaxMaxDepth Then context.INI_ISearch_MaxDepth = ISearch_MaxMaxDepth
                    If context.INI_ISearch_Timeout > ISearch_MaxSearchTimeout Then context.INI_ISearch_Timeout = ISearch_MaxSearchTimeout
                    If context.INI_ISearch_Results > ISearch_MaxResults Then context.INI_ISearch_Results = ISearch_MaxResults
                End If

                ' Process RAG if enabled.
                context.INI_Lib = ParseBoolean(configDict, "Lib")
                If context.INI_Lib Then
                    context.INI_Lib_File = If(configDict.ContainsKey("Lib_File"), configDict("Lib_File"), "")
                    context.INI_Lib_Timeout = If(configDict.ContainsKey("Lib_Timeout"), CLng(configDict("Lib_Timeout")), DEFAULT_TIMEOUT_LIB)
                    context.INI_Lib_Find_SP = If(configDict.ContainsKey("Lib_Find_SP"), configDict("Lib_Find_SP"), Default_Lib_Find_SP)
                    context.INI_Lib_Apply_SP = If(configDict.ContainsKey("Lib_Apply_SP"), configDict("Lib_Apply_SP"), Default_Lib_Apply_SP)
                    context.INI_Lib_Apply_SP_Markup = If(configDict.ContainsKey("Lib_Apply_SP_Markup"), configDict("Lib_Apply_SP_Markup"), Default_Lib_Apply_SP_Markup)

                End If

                ' Process SecondAPI configuration if enabled.
                context.INI_Endpoint_2 = "" ' Necessary for googleapi check (should not be null).
                context.INI_SecondAPI = ParseBoolean(configDict, "SecondAPI")
                If context.INI_SecondAPI Then
                    context.INI_APIKey_2 = If(configDict.ContainsKey("APIKey_2"), configDict("APIKey_2"), "")
                    context.INI_Endpoint_2 = If(configDict.ContainsKey("Endpoint_2"), configDict("Endpoint_2"), "")
                    context.INI_HeaderA_2 = If(configDict.ContainsKey("HeaderA_2"), configDict("HeaderA_2"), "")
                    context.INI_HeaderB_2 = If(configDict.ContainsKey("HeaderB_2"), configDict("HeaderB_2"), "")
                    context.INI_Response_2 = If(configDict.ContainsKey("Response_2"), configDict("Response_2"), "")
                    context.INI_Anon_2 = If(configDict.ContainsKey("Anon_2"), configDict("Anon_2"), "")
                    context.INI_TokenCount_2 = If(configDict.ContainsKey("TokenCount_2"), configDict("TokenCount_2"), "")
                    context.INI_APICall_2 = If(configDict.ContainsKey("APICall_2"), configDict("APICall_2"), "")
                    context.INI_APICall_Object_2 = If(configDict.ContainsKey("APICall_Object_2"), configDict("APICall_Object_2"), "")
                    context.INI_Timeout_2 = If(configDict.ContainsKey("Timeout_2"), CLng(configDict("Timeout_2")), DEFAULT_TIMEOUT_2_LONG)
                    context.INI_MaxOutputToken_2 = If(configDict.ContainsKey("MaxOutputToken_2"), CInt(configDict("MaxOutputToken_2")), DEFAULT_MAX_OUTPUT_TOKEN_2)
                    context.INI_Temperature_2 = If(configDict.ContainsKey("Temperature_2"), configDict("Temperature_2"), DEFAULT_TEMPERATURE)
                    context.INI_Model_2 = If(configDict.ContainsKey("Model_2"), configDict("Model_2"), "")
                    context.INI_APIEncrypted_2 = ParseBoolean(configDict, "APIKeyEncrypted_2")
                    context.INI_APIKeyPrefix_2 = If(configDict.ContainsKey("APIKeyPrefix_2"), configDict("APIKeyPrefix_2"), "")
                End If

                ' Process OAuth2 configuration if enabled.
                context.INI_OAuth2 = ParseBoolean(configDict, "OAuth2")
                If context.INI_OAuth2 Then
                    context.INI_OAuth2ClientMail = If(configDict.ContainsKey("OAuth2ClientMail"), configDict("OAuth2ClientMail"), "")
                    context.INI_OAuth2Scopes = If(configDict.ContainsKey("OAuth2Scopes"), configDict("OAuth2Scopes"), "")
                    context.INI_OAuth2Endpoint = If(configDict.ContainsKey("OAuth2Endpoint"), configDict("OAuth2Endpoint"), "")
                    context.INI_OAuth2ATExpiry = If(configDict.ContainsKey("OAuth2ATExpiry"), CLng(configDict("OAuth2ATExpiry")), DEFAULT_OAUTH2_AT_EXPIRY)

                End If

                If context.INI_SecondAPI Then
                    context.INI_OAuth2_2 = ParseBoolean(configDict, "OAuth2_2")
                    If context.INI_OAuth2_2 Then
                        context.INI_OAuth2ClientMail_2 = If(configDict.ContainsKey("OAuth2ClientMail_2"), configDict("OAuth2ClientMail_2"), "")
                        context.INI_OAuth2Scopes_2 = If(configDict.ContainsKey("OAuth2Scopes_2"), configDict("OAuth2Scopes_2"), "")
                        context.INI_OAuth2Endpoint_2 = If(configDict.ContainsKey("OAuth2Endpoint_2"), configDict("OAuth2Endpoint_2"), "")
                        context.INI_OAuth2ATExpiry_2 = If(configDict.ContainsKey("OAuth2ATExpiry_2"), CLng(configDict("OAuth2ATExpiry_2")), DEFAULT_OAUTH2_AT_EXPIRY_2)
                    End If
                End If

                ' Set runtime ignore prompt based on INI_Ignore. Same with Location.
                If context.INI_Ignore Then context.Ignore = context.SP_Ignore Else context.Ignore = ""
                context.Location = context.INI_Location.Trim()


                ' Resolve Codebasis (used to decode encrypted API keys) if required.
                If context.INI_APIEncrypted OrElse context.INI_APIEncrypted_2 Then
                    If IsEmptyOrBlank(Int_CodeBasis) Then
                        context.Codebasis = GetFromRegistry(RegPath_Base, RegPath_CodeBasis, False)
                    Else
                        context.Codebasis = Int_CodeBasis
                    End If
                End If

                ' Keep backups of configured API keys (as loaded from INI) before decoding.
                context.INI_APIKeyBack = context.INI_APIKey
                context.INI_APIKeyBack_2 = context.INI_APIKey_2

                ' Enforce licensing before continuing setup.
                If Not LicenseOK(context, configDict) Then
                    ShowCustomMessageBox($"{AN} disabled due to invalid or expired license.")
                    Return
                End If

                ' Abort if required INI values are missing and not supplied by the user.
                If INIValuesMissing(context) Then
                    Return
                End If

                ' Additional configurations for OAuth2.
                context.TokenExpiry = Microsoft.VisualBasic.DateAndTime.DateAdd(Microsoft.VisualBasic.DateInterval.Year, -1, DateTime.Now)
                context.DecodedAPI = ""
                context.INI_APIKeyBack = context.INI_APIKey

                context.TokenExpiry_2 = Microsoft.VisualBasic.DateAndTime.DateAdd(Microsoft.VisualBasic.DateInterval.Year, -1, DateTime.Now)
                context.DecodedAPI_2 = ""
                context.INI_APIKeyBack_2 = context.INI_APIKey_2

                ' Set PromptLib if a path is configured.
                If context.INI_PromptLibPath = "" And context.INI_PromptLibPathLocal = "" Then context.INI_PromptLib = False Else context.INI_PromptLib = True

                ' Check and decrypt API keys for the primary model.
                If context.INI_OAuth2 Then
                    context.INI_APIKey = Trim(Replace(RealAPIKey(context.INI_APIKey, False, True, context), "\n", ""))
                    If String.IsNullOrWhiteSpace(context.INI_APIKey) Then
                        ShowCustomMessageBox("Internal error: Could not determine private key (likely a decryption error).")
                        Return
                    End If
                Else
                    context.DecodedAPI = RealAPIKey(context.INI_APIKey, False, False, context)
                    If String.IsNullOrWhiteSpace(context.DecodedAPI) Then
                        ShowCustomMessageBox("Internal error: Could not determine API key (likely a decryption error).")
                        Return
                    End If
                End If

                ' Check and decrypt API keys for the secondary model (if configured).
                If context.INI_SecondAPI Then
                    If context.INI_OAuth2_2 Then
                        context.INI_APIKey_2 = Trim(Replace(RealAPIKey(context.INI_APIKey_2, True, True, context), "\n", ""))
                        If String.IsNullOrWhiteSpace(context.INI_APIKey_2) Then
                            ShowCustomMessageBox("Internal error: Could not determine private key (likely a decryption error).")
                            Return
                        End If
                    Else
                        context.DecodedAPI_2 = RealAPIKey(context.INI_APIKey_2, True, False, context)
                        If String.IsNullOrWhiteSpace(context.DecodedAPI_2) Then
                            MessageBox.Show("Internal error: Could not determine API key for second API (likely a decryption error).", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Return
                        End If
                    End If
                End If

                context.GPTSetupError = False
                context.INIloaded = True

            Catch ex As System.Exception
                MessageBox.Show($"Error in InitializeConfig: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        ''' <summary>
        ''' Parses a boolean INI value from <paramref name="configDict"/>.
        ''' Accepts "yes"/"true" and the German equivalents "ja"/"wahr" (case-insensitive).
        ''' </summary>
        ''' <param name="configDict">Configuration dictionary containing INI values.</param>
        ''' <param name="key">Key to read.</param>
        ''' <param name="defaultvalue">Value to return when the key does not exist.</param>
        ''' <returns>The parsed boolean value, or <paramref name="defaultvalue"/> when missing.</returns>
        Public Shared Function ParseBoolean(configDict As Dictionary(Of String, String), key As String, Optional defaultvalue As Boolean = False) As Boolean
            If configDict.ContainsKey(key) Then
                Dim value = configDict(key).Trim().ToLower()
                Return value = "yes" OrElse value = "true" OrElse value = "ja" OrElse value = "wahr"
            End If
            Return defaultvalue
        End Function


        ''' <summary>
        ''' Checks whether mandatory INI values are missing and optionally prompts the user to provide them.
        ''' </summary>
        ''' <param name="context">Shared context containing current configuration values.</param>
        ''' <returns>
        ''' <c>True</c> if required values are missing and were not provided; otherwise <c>False</c>.
        ''' </returns>
        Public Shared Function INIValuesMissing(ByVal context As ISharedContext) As Boolean
            Dim missingSettings As New Dictionary(Of String, String)
            Dim usercompleted As Boolean = False

            Do

                missingSettings.Clear()

                ' Check for missing values.
                If String.IsNullOrEmpty(context.INI_APIKey) Then missingSettings.Add("APIKey", "APIKey (Model 1)")
                ' If String.IsNullOrEmpty(context.INI_Temperature) Then missingSettings.Add("Temperature", "Temperature (Model 1)")
                If context.INI_Timeout = 0 Then missingSettings.Add("Timeout", "Timeout (Model 1)")
                If String.IsNullOrEmpty(context.INI_Model) Then missingSettings.Add("Model", "Model (Model 1)")
                If String.IsNullOrEmpty(context.INI_Endpoint) Then missingSettings.Add("Endpoint", "Endpoint (Model 1)")
                If String.IsNullOrEmpty(context.INI_APICall) Then missingSettings.Add("APICall", "APICall (Model 1)")
                If String.IsNullOrEmpty(context.INI_Response) Then missingSettings.Add("Response", "Response (Model 1)")

                If context.INI_SecondAPI Then
                    If String.IsNullOrEmpty(context.INI_APIKey_2) Then missingSettings.Add("APIKey_2", "APIKey (Model 2)")
                    'If String.IsNullOrEmpty(context.INI_Temperature_2) Then missingSettings.Add("Temperature_2", "Temperature (Model 2)")
                    If context.INI_Timeout_2 = 0 Then missingSettings.Add("Timeout_2", "Timeout (Model 2)")
                    If String.IsNullOrEmpty(context.INI_Model_2) Then missingSettings.Add("Model_2", "Model (Model 2)")
                    If String.IsNullOrEmpty(context.INI_Endpoint_2) Then missingSettings.Add("Endpoint_2", "Endpoint (Model 2)")
                    If String.IsNullOrEmpty(context.INI_APICall_2) Then missingSettings.Add("APICall_2", "APICall (Model 2)")
                    If String.IsNullOrEmpty(context.INI_Response_2) Then missingSettings.Add("Response_2", "Response (Model 2)")
                End If

                If context.INI_OAuth2 Then
                    If String.IsNullOrEmpty(context.INI_OAuth2ClientMail) Then missingSettings.Add("OAuth2ClientMail", "OAuth2Client Mail (Model 1)")
                    If String.IsNullOrEmpty(context.INI_OAuth2Scopes) Then missingSettings.Add("OAuth2Scopes", "OAuth2Scopes (Model 1)")
                    If String.IsNullOrEmpty(context.INI_OAuth2Endpoint) Then missingSettings.Add("OAuth2Endpoint", "OAuth2Endpoint (Model 1)")
                    If context.INI_OAuth2ATExpiry < 0 Then missingSettings.Add("OAuth2ATExpiry", "OAuth2ATExpiry (Model 1)")
                End If

                If context.INI_OAuth2_2 Then
                    If String.IsNullOrEmpty(context.INI_OAuth2ClientMail_2) Then missingSettings.Add("OAuth2ClientMail_2", "OAuth2ClientMail (Model 2)")
                    If String.IsNullOrEmpty(context.INI_OAuth2Scopes_2) Then missingSettings.Add("OAuth2Scopes_2", "OAuth2Scopes (Model 2)")
                    If String.IsNullOrEmpty(context.INI_OAuth2Endpoint_2) Then missingSettings.Add("OAuth2Endpoint_2", "OAuth2Endpoint (Model 2)")
                    If context.INI_OAuth2ATExpiry_2 < 0 Then missingSettings.Add("OAuth2ATExpiry_2", "OAuth2ATExpiry (Model 2)")
                End If

                If context.INI_ISearch AndAlso context.RDV.Substring(0, 4) = "Word" Then
                    If String.IsNullOrEmpty(context.INI_ISearch_URL) Then missingSettings.Add("ISearch_URL", "Search URL")
                    If String.IsNullOrEmpty(context.INI_ISearch_ResponseMask1) Then missingSettings.Add("ISearch_ResponseMask1", "Response Mask 1")
                    If String.IsNullOrEmpty(context.INI_ISearch_ResponseMask2) Then missingSettings.Add("ISearch_ResponseMask2", "Response Mask 2")
                    If String.IsNullOrEmpty(context.INI_ISearch_Name) Then missingSettings.Add("ISearch_Name", "ISearch_Name")
                    If context.INI_ISearch_Tries = 0 Then missingSettings.Add("ISearch_Tries", "ISearch_Tries")
                    If context.INI_ISearch_Results = 0 Then missingSettings.Add("ISearch_Results", "ISearch_Results")
                End If

                If context.INI_Lib AndAlso context.RDV.Substring(0, 4) = "Word" Then
                    If String.IsNullOrEmpty(context.INI_Lib_File) Then missingSettings.Add("Lib_File", "Lib_File")
                    If String.IsNullOrEmpty(context.INI_Lib_Find_SP) Then missingSettings.Add("Lib_Find_SP", "Lib_Find_SP")
                    If String.IsNullOrEmpty(context.INI_Lib_Apply_SP) Then missingSettings.Add("Lib_Apply_SP", "Lib_Apply_SP")
                    If String.IsNullOrEmpty(context.INI_Lib_Apply_SP_Markup) Then missingSettings.Add("Lib_Apply_SP_Markup", "Lib_Apply_SP_Markup")
                End If

                If context.INI_APIEncrypted OrElse context.INI_APIEncrypted_2 Then
                    If String.IsNullOrEmpty(context.Codebasis) Then missingSettings.Add("Codebasis", "CodeBasis (for decryption)")
                End If

                ' If there are missing settings, prompt user to complete them.
                If missingSettings.Count > 0 Then
                    usercompleted = MissingSettingsWindow(missingSettings, context)
                    If Not usercompleted Then
                        ShowCustomMessageBox($"You have not provided all required parameters, which is why {AN} will not operate properly. Update '{AN2}.ini' (all values are described in the manual) before you continue or retry and add the parameters.")
                        Return True
                        Exit Do
                    End If
                Else
                    Return False
                    Exit Do
                End If
            Loop

        End Function


        ''' <summary>
        ''' Shows a modal dialog to collect missing configuration values and persists them via <c>UpdateAppConfig</c>.
        ''' </summary>
        ''' <param name="Settings">Dictionary mapping INI keys to user-facing label texts.</param>
        ''' <param name="context">Shared context to read/write values from.</param>
        ''' <returns><c>True</c> if the user saved values; otherwise <c>False</c>.</returns>
        Public Shared Function MissingSettingsWindow(Settings As Dictionary(Of String, String), context As ISharedContext) As Boolean

            ' Create the form.
            Dim settingsForm As New System.Windows.Forms.Form()
            settingsForm.Text = $"{AN} Settings"
            settingsForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            settingsForm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            settingsForm.MaximizeBox = False
            settingsForm.MinimizeBox = False
            settingsForm.ShowInTaskbar = False
            settingsForm.TopMost = True

            ' Set the icon.
            Dim bmp As New System.Drawing.Bitmap(My.Resources.Red_Ink_Logo)
            settingsForm.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

            ' Set a predefined font for consistent layout.
            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
            settingsForm.Font = standardFont

            ' Add description label.
            Dim descriptionLabel As New System.Windows.Forms.Label()
            descriptionLabel.Text = "Complete missing mandatory values:"
            descriptionLabel.AutoSize = True
            descriptionLabel.Location = New System.Drawing.Point(10, 20)
            settingsForm.Controls.Add(descriptionLabel)

            ' Define controls for labels and inputs.
            Dim labelControls As New Dictionary(Of String, System.Windows.Forms.Label)
            Dim settingControls As New Dictionary(Of String, System.Windows.Forms.Control)

            ' Dynamically calculate label width.
            Dim maxLabelWidth As Integer = 0

            ' Calculate maximum label width.
            For Each setting In Settings
                Dim textSize As System.Drawing.Size = TextRenderer.MeasureText(setting.Value & ":", standardFont)
                maxLabelWidth = Math.Max(maxLabelWidth, textSize.Width)
            Next

            Dim controlXOffset As Integer = maxLabelWidth + 20
            Dim defaultControlWidth As Integer = 240
            Dim lineSpacing As Integer = CInt(TextRenderer.MeasureText("Sample", standardFont).Height * 1.5)
            Dim yPos As Integer = descriptionLabel.Bottom + 20

            ' Add labels and input controls.
            For Each setting In Settings
                Dim label As New System.Windows.Forms.Label()
                If context.INI_SecondAPI Then
                    label.Text = setting.Value.Replace("{model}", context.INI_Model).Replace("{model2}", context.INI_Model_2) & ":"
                Else
                    label.Text = setting.Value.Replace("{model}", context.INI_Model).Replace("{model2}", "2nd model (none)") & ":"
                End If
                label.AutoSize = True
                label.Font = standardFont
                label.Location = New System.Drawing.Point(10, yPos)
                settingsForm.Controls.Add(label)
                labelControls.Add(setting.Key, label)

                If IsBooleanSetting(setting.Key) Then
                    Dim checkBox As New System.Windows.Forms.CheckBox()
                    checkBox.Checked = Boolean.Parse(GetSettingValue(setting.Key, context))
                    checkBox.Location = New System.Drawing.Point(controlXOffset, yPos)
                    settingsForm.Controls.Add(checkBox)
                    settingControls.Add(setting.Key, checkBox)
                Else
                    Dim textBox As New System.Windows.Forms.TextBox()
                    textBox.Text = GetSettingValue(setting.Key, context)
                    textBox.Size = New System.Drawing.Size(defaultControlWidth, 20)
                    textBox.Location = New System.Drawing.Point(controlXOffset, yPos)
                    settingsForm.Controls.Add(textBox)
                    settingControls.Add(setting.Key, textBox)
                End If

                yPos += lineSpacing
            Next

            ' Add buttons.
            Dim buttonYPos As Integer = yPos + 20
            Dim buttonSpacing As Integer = 10

            Dim okButton As New System.Windows.Forms.Button()
            okButton.Text = "Save and continue"
            Dim okButtonSize As System.Drawing.Size = TextRenderer.MeasureText(okButton.Text, standardFont)
            okButton.Size = New System.Drawing.Size(okButtonSize.Width + 20, okButtonSize.Height + 10)
            okButton.Location = New System.Drawing.Point(10, buttonYPos)
            settingsForm.Controls.Add(okButton)

            Dim okButtonToolTip As New System.Windows.Forms.ToolTip()
            okButtonToolTip.SetToolTip(okButton, $"Will save the exisiting values and those you have entered into a local copy of '{AN2}.ini' (overwriting any existing such file).")

            Dim cancelButton As New System.Windows.Forms.Button()
            cancelButton.Text = "Cancel"
            Dim cancelButtonSize As System.Drawing.Size = TextRenderer.MeasureText(cancelButton.Text, standardFont)
            cancelButton.Size = New System.Drawing.Size(cancelButtonSize.Width + 20, cancelButtonSize.Height + 10)
            cancelButton.Location = New System.Drawing.Point(okButton.Right + buttonSpacing, buttonYPos)
            settingsForm.Controls.Add(cancelButton)

            Dim cancelButtonToolTip As New System.Windows.Forms.ToolTip()
            cancelButtonToolTip.SetToolTip(cancelButton, $"{AN} will not operate properly until you have provided the necessary configuration parameters. You can retry later.")

            ' Flag to track whether the user completed the form.
            Dim userCompleted As Boolean = False

            ' Attach handlers to buttons.
            AddHandler okButton.Click, Sub(sender, e)
                                           For Each settingKey In settingControls.Keys
                                               Dim control = settingControls(settingKey)
                                               If TypeOf control Is System.Windows.Forms.TextBox Then
                                                   ' Handle TextBox settings.
                                                   Dim textValue As String = DirectCast(control, System.Windows.Forms.TextBox).Text
                                                   SetSettingValue(settingKey, textValue, context)
                                               ElseIf TypeOf control Is System.Windows.Forms.CheckBox Then
                                                   ' Handle CheckBox settings.
                                                   Dim boolValue As Boolean = DirectCast(control, System.Windows.Forms.CheckBox).Checked
                                                   SetSettingValue(settingKey, boolValue.ToString(), context)
                                               Else
                                                   MessageBox.Show($"Error in MissingSettingsWindow - unsupported control type for setting '{settingKey}' in MissingSettingsWindow.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                               End If
                                           Next
                                           UpdateAppConfig(context) ' Save the configuration.
                                           userCompleted = True
                                           settingsForm.Close()
                                       End Sub

            AddHandler cancelButton.Click, Sub(sender, e)
                                               settingsForm.Close()
                                           End Sub

            ' Adjust form size dynamically.
            settingsForm.ClientSize = New System.Drawing.Size(controlXOffset + defaultControlWidth + 40, cancelButton.Bottom + 20)

            ' Show the form and wait for user input.
            settingsForm.ShowDialog()

            ' Return whether the user completed the form.
            Return userCompleted
        End Function


        ''' <summary>
        ''' Returns the runtime API key value derived from INI input by optionally removing a configured prefix
        ''' and decoding encrypted values using <c>DecodeString</c>.
        ''' </summary>
        ''' <param name="APIInput">Input value (as stored in the INI file or UI).</param>
        ''' <param name="SecondAPI"><c>True</c> to use second-API settings; otherwise primary settings.</param>
        ''' <param name="IgnorePrefix"><c>True</c> to skip prefix removal/addition; otherwise applies configured prefix.</param>
        ''' <param name="context">Shared context providing encryption flags, prefixes, and codebasis.</param>
        ''' <returns>The final API key value (with prefix applied) after optional decoding.</returns>
        Public Shared Function RealAPIKey(ByVal APIInput As String, ByVal SecondAPI As Boolean, ByVal IgnorePrefix As Boolean, ByVal context As ISharedContext) As String

            APIInput = Trim(RemoveCR(APIInput))

            Dim Prefix As String = ""
            Dim Result As String = APIInput

            ' Determine the prefix based on whether it's the second API and IgnorePrefix is false.
            If Not SecondAPI Then
                If Not IgnorePrefix Then
                    Prefix = context.INI_APIKeyPrefix

                    If Not String.IsNullOrWhiteSpace(Prefix) Then
                        ' Remove the prefix if present.
                        If APIInput.StartsWith(Prefix) Then
                            APIInput = APIInput.Substring(Prefix.Length)
                        End If
                    End If
                End If

                Result = APIInput

                ' Decode the API key if encryption is enabled for the main API.
                If context.INI_APIEncrypted Then
                    Result = DecodeString(APIInput, context.Codebasis)
                End If
            Else
                If Not IgnorePrefix Then
                    Prefix = context.INI_APIKeyPrefix_2

                    If Not String.IsNullOrWhiteSpace(Prefix) Then
                        ' Remove the prefix if present.
                        If APIInput.StartsWith(Prefix) Then
                            APIInput = APIInput.Substring(Prefix.Length)
                        End If
                    End If
                End If

                Result = APIInput

                ' Decode the API key if encryption is enabled for the second API.
                If context.INI_APIEncrypted_2 Then
                    Result = DecodeString(APIInput, context.Codebasis)
                End If
            End If

            ' Remove any carriage return characters.
            Result = RemoveCR(Result)

            ' Add the prefix back and return the final result.
            Result = Prefix & Result

            Return Result
        End Function

        ''' <summary>
        ''' Equivalent of <see cref="RealAPIKey(String, Boolean, Boolean, ISharedContext)"/> for a <see cref="ModelConfig"/>.
        ''' Optionally removes/applies the model prefix and decodes an encrypted key using <paramref name="context2"/>.Codebasis.
        ''' </summary>
        ''' <param name="APIInput">Input value (as stored in configuration).</param>
        ''' <param name="IgnorePrefix"><c>True</c> to skip prefix removal/addition; otherwise applies configured prefix.</param>
        ''' <param name="context">Model configuration providing encryption flags and prefix.</param>
        ''' <param name="context2">Shared context providing <c>Codebasis</c> for decoding.</param>
        ''' <returns>The final API key value (with prefix applied) after optional decoding.</returns>
        Public Shared Function RealAPIKeyMC(ByVal APIInput As String, ByVal IgnorePrefix As Boolean, ByVal context As ModelConfig, context2 As ISharedContext) As String

            APIInput = Trim(RemoveCR(APIInput))

            Dim Prefix As String = ""
            Dim Result As String = APIInput

            ' Determine the prefix based on whether it's the second API and IgnorePrefix is false.

            If Not IgnorePrefix Then
                Prefix = context.APIKeyPrefix

                If Not String.IsNullOrWhiteSpace(Prefix) Then
                    ' Remove the prefix if present.
                    If APIInput.StartsWith(Prefix) Then
                        APIInput = APIInput.Substring(Prefix.Length)
                    End If
                End If
            End If

            Result = APIInput

            ' Decode the API key if encryption is enabled for the main API.
            If context.APIEncrypted Then
                Result = DecodeString(APIInput, context2.Codebasis)
            End If

            ' Remove any carriage return characters.
            Result = RemoveCR(Result)

            ' Add the prefix back and return the final result.
            Result = Prefix & Result

            Return Result
        End Function


    End Class
End Namespace