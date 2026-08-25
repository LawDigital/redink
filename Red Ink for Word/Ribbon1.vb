' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: Ribbon1.vb
' Purpose:
'   Word Ribbon callback surface: routes user commands to ThisAddIn workflows,
'   manages model/menu state, and applies feature/configuration visibility.
'
' Architecture:
'   Thin VSTO UI adapter over the Word add-in services. Control layout is generated
'   in Ribbon1.Designer.vb; this file owns callbacks, dynamic menus, theme-aware
'   icons, and visibility/availability policy.
' =============================================================================

Imports System.Diagnostics
Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Win32
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary

Public Class Ribbon1

    Private Enum OfficeTheme
        Unknown
        Light
        Dark
    End Enum

    Private Sub ApplyThemeAwareMenuIcon()
        Try
            Dim theme = DetectOfficeTheme()
            Select Case theme
                Case OfficeTheme.Light
                    Menu1.Image = SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Large)
                Case Else
                    Menu1.Image = SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Large)
            End Select
            Menu1.ShowImage = True
        Catch
            Menu1.Image = SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Large)
            Menu1.ShowImage = True
        End Try
    End Sub


    Private Function DetectOfficeTheme() As OfficeTheme
        Const registryPath As String = "Software\Microsoft\Office\16.0\Common"
        Const valueName As String = "UI Theme"

        Try
            Using key = Registry.CurrentUser.OpenSubKey(registryPath)
                If key Is Nothing Then Return OfficeTheme.Unknown

                Dim raw = key.GetValue(valueName)
                If raw Is Nothing Then Return OfficeTheme.Unknown

                Dim value As Integer
                If Integer.TryParse(raw.ToString(), value) Then
                    Select Case value
                        Case 0 ' Colorful
                            Return OfficeTheme.Light
                        Case 1, 2 ' Dark Gray, Black
                            Return OfficeTheme.Dark
                        Case 3 ' White
                            Return OfficeTheme.Light
                        Case 4 ' Use system setting -> resolve via Windows app theme
                            Return If(IsWindowsAppsLightTheme(), OfficeTheme.Light, OfficeTheme.Dark)
                    End Select
                End If
            End Using
        Catch
            ' fall through
        End Try

        Return OfficeTheme.Unknown
    End Function

    Private Function IsWindowsAppsLightTheme() As Boolean
        Const personalizePath As String = "Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        Const appsUseLightTheme As String = "AppsUseLightTheme"
        Try
            Using key = Registry.CurrentUser.OpenSubKey(personalizePath)
                If key Is Nothing Then Return True ' default to light if unknown
                Dim raw = key.GetValue(appsUseLightTheme)
                If raw Is Nothing Then Return True
                Dim v As Integer
                If Integer.TryParse(raw.ToString(), v) Then
                    Return v <> 0 ' 1=Light, 0=Dark
                End If
            End Using
        Catch
            ' default to light on error
        End Try
        Return True
    End Function

    ' =========================================================================
    ' Model Selection
    ' =========================================================================

    Public Sub UpdateModelsMenu()
        Try
            Dim available = PrimaryModelManager.GetAvailableModels()
            Dim current = PrimaryModelManager.GetCurrentModelNumber()

            ' Hide the entire model menu if there are fewer than 2 models.
            ' Use the availability helper so a subsequent ApplyRibbonVisibilityConfiguration()
            ' does not force the menu visible again when no additional models are defined.
            SetRibbonControlVisibleByAvailability(Menu6, available.Count > 1)

            For i = 1 To 10
                Dim btn = GetModelButton(i)
                If btn Is Nothing Then Continue For

                If available.Contains(i) Then
                    SetRibbonControlVisibleByAvailability(btn, True)
                    Dim label = PrimaryModelManager.GetModelDisplayName(i)
                    btn.Label = If(i = current, $"{label} (active)", label)
                Else
                    SetRibbonControlVisibleByAvailability(btn, False)
                End If
            Next

            ApplyRibbonVisibilityConfiguration()
        Catch
            ' non-critical
        End Try
    End Sub

    Private Function GetModelButton(i As Integer) As RibbonButton
        Select Case i
            Case 1 : Return RI_Model1
            Case 2 : Return RI_Model2
            Case 3 : Return RI_Model3
            Case 4 : Return RI_Model4
            Case 5 : Return RI_Model5
            Case 6 : Return RI_Model6
            Case 7 : Return RI_Model7
            Case 8 : Return RI_Model8
            Case 9 : Return RI_Model9
            Case 10 : Return RI_Model10
        End Select
        Return Nothing
    End Function

    Private Sub RI_Model1_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_Model1.Click
        SelectModelCommand(1)
    End Sub

    Private Sub RI_Model2_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_Model2.Click
        SelectModelCommand(2)
    End Sub

    Private Sub RI_Model3_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_Model3.Click
        SelectModelCommand(3)
    End Sub

    Private Sub RI_Model4_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_Model4.Click
        SelectModelCommand(4)
    End Sub

    Private Sub RI_Model5_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_Model5.Click
        SelectModelCommand(5)
    End Sub

    Private Sub RI_Model6_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_Model6.Click
        SelectModelCommand(6)
    End Sub

    Private Sub RI_Model7_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_Model7.Click
        SelectModelCommand(7)
    End Sub

    Private Sub RI_Model8_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_Model8.Click
        SelectModelCommand(8)
    End Sub

    Private Sub RI_Model9_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_Model9.Click
        SelectModelCommand(9)
    End Sub

    Private Sub RI_Model10_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_Model10.Click
        SelectModelCommand(10)
    End Sub

    Private _voiceCommands As System.Collections.Generic.Dictionary(Of String, System.Action)

    Private ReadOnly Property VoiceCommands As System.Collections.Generic.Dictionary(Of String, System.Action)
        Get
            If _voiceCommands Is Nothing Then
                _voiceCommands = CreateVoiceCommands()
            End If

            Return _voiceCommands
        End Get
    End Property

    Private NotInheritable Class TalkToMeRibbonCommandEntry
        Public Property Name As String = ""
        Public Property Category As String = ""
        Public Property Button As Microsoft.Office.Tools.Ribbon.RibbonButton = Nothing
        Public Property Execute As System.Action = Nothing
    End Class

    Private Function GetTalkToMeRibbonCommandEntries() As List(Of TalkToMeRibbonCommandEntry)
        Return New List(Of TalkToMeRibbonCommandEntry) From {
            New TalkToMeRibbonCommandEntry With {.Name = "translate_language1", .Category = "Task", .Button = RI_Primlang, .Execute = AddressOf RunPrimaryLanguageCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "translate_language2", .Category = "Task", .Button = RI_SecLang, .Execute = AddressOf RunSecondaryLanguageCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "translate", .Category = "Task", .Button = RI_Translate, .Execute = AddressOf RunTranslateCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "correct", .Category = "Task", .Button = RI_Correct, .Execute = AddressOf RunCorrectCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "improve", .Category = "Improve", .Button = RI_Improve, .Execute = AddressOf RunImproveCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "apply_my_style", .Category = "Improve", .Button = RI_ApplyMyStyle, .Execute = AddressOf RunApplyMyStyleCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "no_fillers", .Category = "Improve", .Button = RI_NoFillers, .Execute = AddressOf RunNoFillersCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "friendly", .Category = "Improve", .Button = RI_Friendly, .Execute = AddressOf RunFriendlyCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "convincing", .Category = "Improve", .Button = RI_Convincing, .Execute = AddressOf RunConvincingCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "apply_comment", .Category = "Improve", .Button = RI_BalloonMergePart, .Execute = AddressOf RunBalloonMergePartCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "apply_comment_justify", .Category = "Improve", .Button = RI_BalloonMergePartJustify, .Execute = AddressOf RunBalloonMergePartJustifyCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "apply_comment_to_para", .Category = "Improve", .Button = RI_BalloonMergeFull, .Execute = AddressOf RunBalloonMergeFullCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "apply_comment_to_para_justify", .Category = "Improve", .Button = RI_BalloonMergeFullJustify, .Execute = AddressOf RunBalloonMergeFullJustifyCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "apply_comment_edit", .Category = "Improve", .Button = RI_BalloonMergePartPrompt, .Execute = AddressOf RunBalloonMergePartPromptCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "apply_comment_to_para_edit", .Category = "Improve", .Button = RI_BalloonMergeFullPrompt, .Execute = AddressOf RunBalloonMergeFullPromptCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "store_original_clause", .Category = "Improve", .Button = RI_StoreOriginalClause, .Execute = AddressOf RunStoreOriginalClauseCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "justify_markup", .Category = "Improve", .Button = RI_JustifyMarkup, .Execute = AddressOf RunJustifyMarkupCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "filibuster", .Category = "Improve", .Button = RI_Filibuster, .Execute = AddressOf RunFilibusterCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "apply_doc_style", .Category = "Improve", .Button = RI_ApplyDocStyle, .Execute = AddressOf RunApplyDocStyleCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "learn_doc_style", .Category = "Improve", .Button = RI_LearnDocStyle, .Execute = AddressOf RunLearnDocStyleCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "shorten", .Category = "Task", .Button = RI_Shorten, .Execute = AddressOf RunShortenCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "anonymize_ai", .Category = "Anonymize", .Button = RI_Anonymize, .Execute = AddressOf RunAnonymizeCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "anonymize_terms", .Category = "Anonymize", .Button = RI_Anonymization, .Execute = AddressOf RunAnonymizationCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "prepare_redactions", .Category = "Anonymize", .Button = RI_PrepareRedactions, .Execute = AddressOf RunPrepareRedactionsCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "finalize_redactions", .Category = "Anonymize", .Button = RI_FinalizeRedactions, .Execute = AddressOf RunFinalizeRedactionsCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "edit_redaction_instructions", .Category = "Anonymize", .Button = RI_EditRedact, .Execute = AddressOf RunEditRedactCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "check_document_ii", .Category = "Anonymize", .Button = RI_CheckDocumentsII, .Execute = AddressOf RunCheckDocumentsIICommand},
            New TalkToMeRibbonCommandEntry With {.Name = "switch_party", .Category = "Task", .Button = RI_SwitchParty, .Execute = AddressOf RunSwitchPartyCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "summarize", .Category = "Analyze", .Button = RI_Summarize, .Execute = AddressOf RunSummarizeCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "explain", .Category = "Analyze", .Button = RI_Explain, .Execute = AddressOf RunExplainCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "argue_against", .Category = "Analyze", .Button = RI_ArgueAgainst, .Execute = AddressOf RunArgueAgainstCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "suggest_titles", .Category = "Analyze", .Button = RI_SuggestTitles, .Execute = AddressOf RunSuggestTitlesCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "revisions_summary", .Category = "Analyze", .Button = RI_RevisionsSummary, .Execute = AddressOf RunRevisionSummaryCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "special_service", .Category = "Analyze", .Button = RI_SpecialModel, .Execute = AddressOf RunSpecialModelCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "document_check", .Category = "Analyze", .Button = RI_DocCheck, .Execute = AddressOf RunDocCheckCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "find_clause", .Category = "Analyze", .Button = RI_FindClause, .Execute = AddressOf RunFindClauseCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "add_clause", .Category = "Analyze", .Button = RI_AddClause, .Execute = AddressOf RunAddClauseCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "tabular_overview", .Category = "Analyze", .Button = RI_Tabular, .Execute = AddressOf RunTabularCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "create_podcast", .Category = "Analyze", .Button = RI_CreatePodcast, .Execute = AddressOf RunCreatePodcastCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "create_audio", .Category = "Analyze", .Button = RI_CreateAudio, .Execute = AddressOf RunCreateAudioCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "define_my_style", .Category = "Analyze", .Button = RI_DefineMyStyle, .Execute = AddressOf RunDefineMyStyleCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "run_web_agent", .Category = "Analyze", .Button = RI_WebAgent, .Execute = AddressOf RunWebAgentCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "edit_web_agent", .Category = "Analyze", .Button = RI_EditWebAgent, .Execute = AddressOf RunEditWebAgentCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "snapshot_compare", .Category = "Analyze", .Button = RI_Snapshot, .Execute = AddressOf RunSnapshotCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "find_hidden_prompts", .Category = "Analyze", .Button = RI_FindHidden, .Execute = AddressOf RunFindHiddenCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "freestyle", .Category = "Task", .Button = RI_FreestyleNM, .Execute = AddressOf RunFreestyleNmCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "freestyle_second", .Category = "Task", .Button = RI_FreestyleAM, .Execute = AddressOf RunFreestyleAmCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "freestyle_redo", .Category = "Task", .Button = RI_FreestyleRepeat, .Execute = AddressOf RunFreestyleRepeatCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "chat", .Category = "Task", .Button = RI_Chat2, .Execute = AddressOf RunChatCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "discuss_inky", .Category = "Analyze", .Button = RI_DiscussInky, .Execute = AddressOf RunDiscussInkyCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "context_search", .Category = "Task", .Button = RI_Search, .Execute = AddressOf RunSearchCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "self_compare_selection", .Category = "Word Helpers", .Button = RI_Halves, .Execute = AddressOf RunHalvesCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "compare_active_docs", .Category = "Word Helpers", .Button = RI_LiveCompare, .Execute = AddressOf RunLiveCompareCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "accept_format_changes", .Category = "Word Helpers", .Button = RI_AcceptFormat, .Execute = AddressOf RunAcceptFormatCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "markup_time_span", .Category = "Word Helpers", .Button = RI_TimeSpan, .Execute = AddressOf RunTimeSpanCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "regex_search_replace", .Category = "Word Helpers", .Button = RI_Regex, .Execute = AddressOf RunRegexCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "edit_create_diagrams", .Category = "Word Helpers", .Button = RI_Charting, .Execute = AddressOf RunChartingCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "flowchart_to_webapp", .Category = "Word Helpers", .Button = RI_WebApp, .Execute = AddressOf RunWebAppCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "import_text_file", .Category = "Word Helpers", .Button = RI_Import, .Execute = AddressOf RunImportCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "flatten_pdfs_to_images", .Category = "Word Helpers", .Button = RI_FlattenPDF, .Execute = AddressOf RunFlattenPdfCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "convert_files_to_text", .Category = "Word Helpers", .Button = RI_ConvertDocToTxt, .Execute = AddressOf RunConvertDocToTxtCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "stamp_pdf_exhibits", .Category = "Word Helpers", .Button = RI_Stamper, .Execute = AddressOf RunStamperCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "split_pdf_with_ai", .Category = "Word Helpers", .Button = RI_SplitPDF, .Execute = AddressOf RunSplitPdfCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "convert_markdown", .Category = "Word Helpers", .Button = RI_Markdown, .Execute = AddressOf RunMarkdownCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "reset_spacing", .Category = "Word Helpers", .Button = RI_ResetSpacing, .Execute = AddressOf RunResetSpacingCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "remove_content_controls", .Category = "Word Helpers", .Button = RI_ContentControls, .Execute = AddressOf RunContentControlsCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "remove_ri_reference", .Category = "Word Helpers", .Button = RI_Remove, .Execute = AddressOf RunRemoveCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "clipboard_to_text", .Category = "Word Helpers", .Button = RI_InsertClipboard, .Execute = AddressOf RunInsertClipboardCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "generate_image", .Category = "Task", .Button = RI_Image, .Execute = AddressOf RunImageCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "transcriptor", .Category = "Task", .Button = RI_Transcriptor, .Execute = AddressOf RunTranscriptorCommand},
            New TalkToMeRibbonCommandEntry With {.Name = "help_me", .Category = "Task", .Button = RI_HelpMe, .Execute = AddressOf RunHelpMeCommand}
        }
    End Function


    Public Function GetTalkToMeCommandDefinitions() As List(Of TalkToMeCommandDefinition)
        Dim result As New List(Of TalkToMeCommandDefinition)()

        For Each entry As TalkToMeRibbonCommandEntry In GetTalkToMeRibbonCommandEntries()
            If entry Is Nothing OrElse entry.Button Is Nothing Then
                Continue For
            End If

            If Not entry.Button.Visible OrElse Not entry.Button.Enabled Then
                Continue For
            End If

            result.Add(New TalkToMeCommandDefinition With {
                .Name = entry.Name,
                .Label = If(entry.Button.Label, "").Trim(),
                .Category = entry.Category,
                .Description = If(entry.Button.ScreenTip, "").Trim(),
                .Aliases = New List(Of String)()
            })
        Next

        Return result
    End Function

    Public Function TryExecuteTalkToMeCommand(commandName As String) As Boolean
        Dim normalizedName As String = If(commandName, "").Trim()

        If String.IsNullOrWhiteSpace(normalizedName) Then
            Return False
        End If

        For Each entry As TalkToMeRibbonCommandEntry In GetTalkToMeRibbonCommandEntries()
            If entry Is Nothing OrElse entry.Execute Is Nothing Then
                Continue For
            End If

            If normalizedName.Equals(entry.Name, StringComparison.OrdinalIgnoreCase) Then
                entry.Execute.Invoke()
                Return True
            End If
        Next

        Return False
    End Function

    Private Function CreateVoiceCommands() As System.Collections.Generic.Dictionary(Of String, System.Action)
        Return New System.Collections.Generic.Dictionary(Of String, System.Action)(System.StringComparer.OrdinalIgnoreCase) From {
            {"correct", AddressOf RunCorrectCommand},
            {"summarize", AddressOf RunSummarizeCommand},
            {"summarise", AddressOf RunSummarizeCommand},
            {"shorten", AddressOf RunShortenCommand},
            {"translate", AddressOf RunTranslateCommand},
            {"primary language", AddressOf RunPrimaryLanguageCommand},
            {"secondary language", AddressOf RunSecondaryLanguageCommand},
            {"improve", AddressOf RunImproveCommand},
            {"chat", AddressOf RunChatCommand},
            {"search", AddressOf RunSearchCommand},
            {"help me", AddressOf RunHelpMeCommand},
            {"transcriptor", AddressOf RunTranscriptorCommand},
            {"explain", AddressOf RunExplainCommand},
            {"suggest titles", AddressOf RunSuggestTitlesCommand},
            {"create podcast", AddressOf RunCreatePodcastCommand},
            {"create audio", AddressOf RunCreateAudioCommand},
            {"no fillers", AddressOf RunNoFillersCommand},
            {"friendly", AddressOf RunFriendlyCommand},
            {"convincing", AddressOf RunConvincingCommand},
            {"apply my style", AddressOf RunApplyMyStyleCommand},
            {"define my style", AddressOf RunDefineMyStyleCommand},
            {"doc check", AddressOf RunDocCheckCommand},
            {"find clause", AddressOf RunFindClauseCommand},
            {"add clause", AddressOf RunAddClauseCommand},
            {"web agent", AddressOf RunWebAgentCommand},
            {"edit web agent", AddressOf RunEditWebAgentCommand},
            {"markdown", AddressOf RunMarkdownCommand},
            {"find hidden", AddressOf RunFindHiddenCommand},
            {"content controls", AddressOf RunContentControlsCommand},
            {"prepare redactions", AddressOf RunPrepareRedactionsCommand},
            {"finalize redactions", AddressOf RunFinalizeRedactionsCommand},
            {"check documents", AddressOf RunCheckDocumentsIICommand},
            {"discuss inky", AddressOf RunDiscussInkyCommand},
            {"convert doc to text", AddressOf RunConvertDocToTxtCommand},
            {"flatten pdf", AddressOf RunFlattenPdfCommand},
            {"charting", AddressOf RunChartingCommand},
            {"snapshot", AddressOf RunSnapshotCommand},
            {"remove comments", AddressOf RunRemoveCommand},
            {"web app", AddressOf RunWebAppCommand},
            {"split pdf", AddressOf RunSplitPdfCommand},
            {"store original clause", AddressOf RunStoreOriginalClauseCommand},
            {"justify markup", AddressOf RunJustifyMarkupCommand},
            {"stamper", AddressOf RunStamperCommand},
            {"image", AddressOf RunImageCommand},
            {"tabular", AddressOf RunTabularCommand}
        }
    End Function

    Private Function NormalizeVoiceCommand(spokenCommand As String) As String
        Dim normalizedCommand As String = spokenCommand.Trim()

        If normalizedCommand.StartsWith("inky", System.StringComparison.OrdinalIgnoreCase) Then
            normalizedCommand = normalizedCommand.Substring(4).TrimStart(" "c, ","c, ":"c, ";"c, "."c, "-"c)
        End If

        normalizedCommand = System.Text.RegularExpressions.Regex.Replace(normalizedCommand, "\s+", " ").Trim()

        Return normalizedCommand
    End Function

    Public Sub ExecuteVoiceCommand(spokenCommand As String)
        If String.IsNullOrWhiteSpace(spokenCommand) Then
            Return
        End If

        Dim normalizedCommand As String = NormalizeVoiceCommand(spokenCommand)
        Dim action As System.Action = Nothing

        If String.IsNullOrWhiteSpace(normalizedCommand) Then
            SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Unknown voice command: " & spokenCommand)
            Return
        End If

        If Not VoiceCommands.TryGetValue(normalizedCommand, action) OrElse action Is Nothing Then
            SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Unknown voice command: " & spokenCommand)
            Return
        End If

        Try
            action.Invoke()
        Catch ex As System.Exception
            SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Voice command failed: " & ex.Message)
        End Try
    End Sub

    Private ReadOnly RibbonControlsHiddenByAvailability As New System.Collections.Generic.HashSet(Of String)(
        System.StringComparer.OrdinalIgnoreCase
    )

    Private ReadOnly RibbonControlsHiddenByConfiguration As New System.Collections.Generic.HashSet(Of String)(
        System.StringComparer.OrdinalIgnoreCase
    )

    Private Function GetRibbonControlLabel(ByVal ribbonControl As Microsoft.Office.Tools.Ribbon.RibbonControl) As String
        If ribbonControl Is Nothing Then
            Return ""
        End If

        Try
            Dim propertyInfo As System.Reflection.PropertyInfo =
                ribbonControl.GetType().GetProperty(
                    "Label",
                    System.Reflection.BindingFlags.Instance Or
                    System.Reflection.BindingFlags.Public Or
                    System.Reflection.BindingFlags.IgnoreCase
                )

            If propertyInfo Is Nothing Then
                Return ""
            End If

            Dim value As Object = propertyInfo.GetValue(ribbonControl, Nothing)
            Return If(TryCast(value, String), "")
        Catch
            Return ""
        End Try
    End Function

    Private Function SplitRibbonControlNames(ByVal controlNames As String) As String()
        If System.String.IsNullOrWhiteSpace(controlNames) Then
            Return New String() {}
        End If

        Dim separators As Char() = {","c, ";"c}

        Return controlNames.Split(
            separators,
            System.StringSplitOptions.RemoveEmptyEntries
        )
    End Function

    Private Function IsRibbonControlHiddenByConfiguration(
        ByVal ribbonControl As Microsoft.Office.Tools.Ribbon.RibbonControl
    ) As Boolean
        If ribbonControl Is Nothing Then
            Return False
        End If

        If System.String.IsNullOrWhiteSpace(ribbonControl.Name) Then
            Return False
        End If

        Return RibbonControlsHiddenByConfiguration.Contains(ribbonControl.Name)
    End Function


    Private Function GetRibbonControlByName(ByVal controlName As String) As Microsoft.Office.Tools.Ribbon.RibbonControl
        If System.String.IsNullOrWhiteSpace(controlName) Then
            Return Nothing
        End If

        Dim normalizedControlName As String = controlName.Trim()
        Dim flags As System.Reflection.BindingFlags =
            System.Reflection.BindingFlags.Instance Or
            System.Reflection.BindingFlags.NonPublic Or
            System.Reflection.BindingFlags.Public Or
            System.Reflection.BindingFlags.IgnoreCase

        Dim directField As System.Reflection.FieldInfo =
            Me.GetType().GetField(normalizedControlName, flags)

        If directField IsNot Nothing Then
            Dim directControl As Microsoft.Office.Tools.Ribbon.RibbonControl =
                TryCast(directField.GetValue(Me), Microsoft.Office.Tools.Ribbon.RibbonControl)

            If directControl IsNot Nothing Then
                Return directControl
            End If
        End If

        For Each fieldInfo As System.Reflection.FieldInfo In Me.GetType().GetFields(flags)
            Dim ribbonControl As Microsoft.Office.Tools.Ribbon.RibbonControl =
                TryCast(fieldInfo.GetValue(Me), Microsoft.Office.Tools.Ribbon.RibbonControl)

            If ribbonControl Is Nothing Then
                Continue For
            End If

            If normalizedControlName.Equals(fieldInfo.Name, System.StringComparison.OrdinalIgnoreCase) OrElse
               normalizedControlName.Equals(If(ribbonControl.Name, ""), System.StringComparison.OrdinalIgnoreCase) OrElse
               normalizedControlName.Equals(GetRibbonControlLabel(ribbonControl), System.StringComparison.OrdinalIgnoreCase) Then
                Return ribbonControl
            End If
        Next

        Return Nothing
    End Function

    Private Sub SetRibbonControlVisibleByAvailability(
        ByVal ribbonControl As Microsoft.Office.Tools.Ribbon.RibbonControl,
        ByVal visible As Boolean
    )
        If ribbonControl Is Nothing Then
            System.Diagnostics.Debug.WriteLine("SetRibbonControlVisibleByAvailability: ribbonControl ist Nothing.")
            Return
        End If

        If System.String.IsNullOrWhiteSpace(ribbonControl.Name) Then
            System.Diagnostics.Debug.WriteLine("SetRibbonControlVisibleByAvailability: Ribbon-Control ohne Name.")
            ribbonControl.Visible = visible
            Return
        End If

        If visible Then
            RibbonControlsHiddenByAvailability.Remove(ribbonControl.Name)
        Else
            RibbonControlsHiddenByAvailability.Add(ribbonControl.Name)
        End If

        If visible AndAlso IsRibbonControlHiddenByConfiguration(ribbonControl) Then
            ribbonControl.Visible = False
            Return
        End If

        ribbonControl.Visible = visible
    End Sub

    Private Sub SetRibbonControlVisibleByConfiguration(
        ByVal ribbonControl As Microsoft.Office.Tools.Ribbon.RibbonControl,
        ByVal visible As Boolean
    )
        If ribbonControl Is Nothing Then
            System.Diagnostics.Debug.WriteLine("SetRibbonControlVisibleByConfiguration: ribbonControl is Nothing.")
            Return
        End If

        If System.String.IsNullOrWhiteSpace(ribbonControl.Name) Then
            System.Diagnostics.Debug.WriteLine("SetRibbonControlVisibleByConfiguration: Ribbon control without name.")
            ribbonControl.Visible = visible
            Return
        End If

        If visible Then
            RibbonControlsHiddenByConfiguration.Remove(ribbonControl.Name)

            If Not RibbonControlsHiddenByAvailability.Contains(ribbonControl.Name) Then
                ribbonControl.Visible = True
            End If
        Else
            RibbonControlsHiddenByConfiguration.Add(ribbonControl.Name)
            ribbonControl.Visible = False
        End If
    End Sub

    Public Sub ApplyRibbonVisibilityConfiguration()
        ApplySimpleModeToRibbonControls(
            ThisAddIn.INI_SimpleMenuHide,
            ThisAddIn.INI_MenuBlock,
            ThisAddIn.INI_SimpleMenuOverride
        )
    End Sub

    Public Sub ApplySimpleModeToRibbonControls(ByVal controlNames As String,
        ByVal blockedControlNames As String,
        ByVal simpleMode As Boolean
    )

        Dim names As String() = SplitRibbonControlNames(controlNames)
        Dim blockedNames As String() = SplitRibbonControlNames(blockedControlNames)

        If names.Length = 0 AndAlso blockedNames.Length = 0 Then
            System.Diagnostics.Debug.WriteLine("ApplySimpleModeToRibbonControls: No control names provided.")
            Return
        End If

        RibbonControlsHiddenByConfiguration.Clear()

        For Each rawName As String In names
            Dim controlName As String = rawName.Trim()

            If System.String.IsNullOrWhiteSpace(controlName) Then
                Continue For
            End If

            Try
                Dim ribbonControl As Microsoft.Office.Tools.Ribbon.RibbonControl =
                    GetRibbonControlByName(controlName)

                If ribbonControl Is Nothing Then
                    System.Diagnostics.Debug.WriteLine("ApplySimpleModeToRibbonControls: Ribbon control not found: " & controlName)
                    Continue For
                End If

                If simpleMode Then
                    ribbonControl.Visible = False
                Else
                    If RibbonControlsHiddenByAvailability.Contains(ribbonControl.Name) Then
                        System.Diagnostics.Debug.WriteLine(
                            "ApplySimpleModeToRibbonControls: Control remains hidden due to availability: " & ribbonControl.Name
                        )
                    Else
                        ribbonControl.Visible = True
                    End If
                End If

            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("ApplySimpleModeToRibbonControls: Error in '" & controlName & "': " & ex.Message)
            End Try
        Next

        For Each rawName As String In blockedNames
            Dim controlName As String = rawName.Trim()

            If System.String.IsNullOrWhiteSpace(controlName) Then
                Continue For
            End If

            Try
                Dim ribbonControl As Microsoft.Office.Tools.Ribbon.RibbonControl =
                    GetRibbonControlByName(controlName)

                If ribbonControl Is Nothing Then
                    System.Diagnostics.Debug.WriteLine("ApplySimpleModeToRibbonControls: Ribbon control to block not found: " & controlName)
                    Continue For
                End If

                SetRibbonControlVisibleByConfiguration(ribbonControl, False)

            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("ApplySimpleModeToRibbonControls: Error blocking '" & controlName & "': " & ex.Message)
            End Try
        Next
    End Sub

    Public Function GetRibbonControlNamesAsString() As String
        Dim result As New System.Text.StringBuilder()
        Dim flags As System.Reflection.BindingFlags =
            System.Reflection.BindingFlags.Instance Or
            System.Reflection.BindingFlags.NonPublic Or
            System.Reflection.BindingFlags.Public

        Try
            For Each fieldInfo As System.Reflection.FieldInfo In Me.GetType().GetFields(flags)
                Dim ribbonControl As Microsoft.Office.Tools.Ribbon.RibbonControl =
                    TryCast(fieldInfo.GetValue(Me), Microsoft.Office.Tools.Ribbon.RibbonControl)

                If ribbonControl Is Nothing Then
                    Continue For
                End If

                result.AppendLine(
                    fieldInfo.Name &
                    " | " &
                    fieldInfo.FieldType.Name &
                    " | Name=" & If(ribbonControl.Name, "") &
                    " | Label=" & GetRibbonControlLabel(ribbonControl)
                )
            Next

        Catch ex As System.Exception
            System.Diagnostics.Debug.WriteLine("GetRibbonControlNamesAsString: Error: " & ex.Message)
        End Try

        Return result.ToString()
    End Function


    Private Sub SelectModelCommand(modelNumber As Integer)
        ThisAddIn.SelectModel(modelNumber)
    End Sub

    Private Sub ExecuteLoggedCommand(logMessage As String, command As System.Action)
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, logMessage)
        command.Invoke()
    End Sub

    Private Sub RunCorrectCommand()
        ExecuteLoggedCommand("Correct_Word invoked", Sub() Globals.ThisAddIn.Correct())
    End Sub

    Private Sub RunSummarizeCommand()
        ExecuteLoggedCommand("Summarize_Word invoked", Sub() Globals.ThisAddIn.Summarize())
    End Sub

    Private Sub RunShortenCommand()
        ExecuteLoggedCommand("Shorten_Word invoked", Sub() Globals.ThisAddIn.Shorten())
    End Sub

    Private Sub RunPrimaryLanguageCommand()
        ExecuteLoggedCommand("PrimLang_Word invoked", Sub() Globals.ThisAddIn.InLanguage1())
    End Sub

    Private Sub RunSecondaryLanguageCommand()
        ExecuteLoggedCommand("SecLang_Word invoked", Sub() Globals.ThisAddIn.InLanguage2())
    End Sub

    Private Sub RunImproveCommand()
        ExecuteLoggedCommand("Improve_Word invoked", Sub() Globals.ThisAddIn.Improve())
    End Sub

    Private Sub RunFreestyleNmCommand()
        ExecuteLoggedCommand("FreestyleNM_Word invoked", Sub() Globals.ThisAddIn.FreeStyleNM())
    End Sub

    Private Sub RunAnonymizeCommand()
        ExecuteLoggedCommand("Anonymize_Word invoked", Sub() Globals.ThisAddIn.Anonymize())
    End Sub

    Private Sub RunChatCommand()
        ExecuteLoggedCommand("Chat_Word invoked", Sub() Globals.ThisAddIn.ShowChatForm())
    End Sub

    Private Sub RunTimeSpanCommand()
        ExecuteLoggedCommand("TimeSpan_Word invoked", Sub() Globals.ThisAddIn.CalculateUserMarkupTimeSpan())
    End Sub

    Private Sub RunAcceptFormatCommand()
        ExecuteLoggedCommand("AcceptFormat_Word invoked", Sub() Globals.ThisAddIn.AcceptFormatting())
    End Sub

    Private Sub RunTranslateCommand()
        ExecuteLoggedCommand("Translate_Word invoked", Sub() Globals.ThisAddIn.InOther())
    End Sub

    Private Sub RunSettingsCommand()
        ExecuteLoggedCommand("Settings_Word invoked", Sub() Globals.ThisAddIn.ShowSettings())
    End Sub

    Private Sub RunFreestyleAmCommand()
        ExecuteLoggedCommand("FreestyleAM_Word invoked", Sub() Globals.ThisAddIn.FreeStyleAM())
    End Sub

    Private Sub RunSwitchPartyCommand()
        ExecuteLoggedCommand("SwitchParty_Word invoked", Sub() Globals.ThisAddIn.SwitchParty())
    End Sub

    Private Sub RunRegexCommand()
        ExecuteLoggedCommand("Regex_Word invoked", Sub() Globals.ThisAddIn.RegexSearchReplace())
    End Sub

    Private Sub RunImportCommand()
        ExecuteLoggedCommand("Import_Word invoked", Sub() Globals.ThisAddIn.ImportTextFile())
    End Sub

    Private Sub RunHalvesCommand()
        ExecuteLoggedCommand("Halves_Word invoked", Sub() Globals.ThisAddIn.CompareSelectionHalves())
    End Sub

    Private Sub RunSearchCommand()
        ExecuteLoggedCommand("Search_Word invoked", Sub() Globals.ThisAddIn.ContextSearch())
    End Sub

    Private Sub RunEastereggCommand()
        ExecuteLoggedCommand("Easteregg_Word invoked", Sub() Globals.ThisAddIn.EasterEgg())
    End Sub

    Private Sub RunTranscriptorCommand()
        ExecuteLoggedCommand("Transcriptor_Word invoked", Sub() Globals.ThisAddIn.Transcriptor())
    End Sub

    Private Sub RunTalkToMeCommand()
        ExecuteLoggedCommand("TalkToMe_Word invoked", Sub() Globals.ThisAddIn.ShowTalkToMeWidget())
    End Sub

    Private Sub RunExplainCommand()
        ExecuteLoggedCommand("Explain_Word invoked", Sub() Globals.ThisAddIn.Explain())
    End Sub

    Private Sub RunSuggestTitlesCommand()
        ExecuteLoggedCommand("SuggestTitles_Word invoked", Sub() Globals.ThisAddIn.SuggestTitles())
    End Sub

    Private Sub RunCreatePodcastCommand()
        ExecuteLoggedCommand("CreatePodcast_Word invoked", Sub() Globals.ThisAddIn.CreatePodcast())
    End Sub

    Private Sub RunCreateAudioCommand()
        ExecuteLoggedCommand("CreateAudio_Word invoked", Sub() Globals.ThisAddIn.CreateAudio())
    End Sub

    Private Sub RunNoFillersCommand()
        ExecuteLoggedCommand("NoFillers_Word invoked", Sub() Globals.ThisAddIn.NoFillers())
    End Sub

    Private Sub RunFriendlyCommand()
        ExecuteLoggedCommand("Friendly_Word invoked", Sub() Globals.ThisAddIn.Friendly())
    End Sub

    Private Sub RunConvincingCommand()
        ExecuteLoggedCommand("Convincing_Word invoked", Sub() Globals.ThisAddIn.Convincing())
    End Sub

    Private Sub RunSpecialModelCommand()
        ExecuteLoggedCommand("SpecialModel_Word invoked", Sub() Globals.ThisAddIn.SpecialModel())
    End Sub

    Private Sub RunAnonymizationCommand()
        ExecuteLoggedCommand("Anonymization_Word invoked", Sub() Globals.ThisAddIn.AnonymizeSelection())
    End Sub

    Private Sub RunInsertClipboardCommand()
        ExecuteLoggedCommand("InsertClipboard_Word invoked", Sub() Globals.ThisAddIn.InsertClipboard())
    End Sub

    Private Sub RunBalloonMergePartCommand()
        ExecuteLoggedCommand("BallooMergePart_Word invoked", Sub() Globals.ThisAddIn.BalloonMerge(False, True))
    End Sub

    Private Sub RunBalloonMergeFullCommand()
        ExecuteLoggedCommand("BalloonMergeFull_Word invoked", Sub() Globals.ThisAddIn.BalloonMerge(True, True))
    End Sub

    Private Sub RunBalloonMergePartPromptCommand()
        ExecuteLoggedCommand("BalloonMergePartPrompt_Word invoked", Sub() Globals.ThisAddIn.BalloonMerge(False, False))
    End Sub

    Private Sub RunBalloonMergeFullPromptCommand()
        ExecuteLoggedCommand("BalloonMergeFullPrompt_Word invoked", Sub() Globals.ThisAddIn.BalloonMerge(True, False))
    End Sub

    Private Sub RunFreestyleRepeatCommand()
        ExecuteLoggedCommand("FreestyleRepeat_Word invoked", Sub() Globals.ThisAddIn.FreeStyleRepeat())
    End Sub

    Private Sub RunApplyMyStyleCommand()
        ExecuteLoggedCommand("ApplyMyStyle_Word invoked", Sub() Globals.ThisAddIn.ApplyMyStyle())
    End Sub

    Private Sub RunDefineMyStyleCommand()
        ExecuteLoggedCommand("DefineMyStyle_Word invoked", Sub() Globals.ThisAddIn.DefineMyStyle())
    End Sub

    Private Sub RunDocCheckCommand()
        ExecuteLoggedCommand("DocCheck_Word invoked", Sub() Globals.ThisAddIn.RunDocCheck())
    End Sub

    Private Sub RunFindClauseCommand()
        ExecuteLoggedCommand("FindClause_Word invoked", Sub() Globals.ThisAddIn.FindClause())
    End Sub

    Private Sub RunAddClauseCommand()
        ExecuteLoggedCommand("AddClause_Word invoked", Sub() Globals.ThisAddIn.AddClause())
    End Sub

    Private Sub RunWebAgentCommand()
        ExecuteLoggedCommand("WebAgent_Word invoked", Sub() Globals.ThisAddIn.WebAgent())
    End Sub

    Private Sub RunEditWebAgentCommand()
        ExecuteLoggedCommand("EditWebAgent_Word invoked", Sub() Globals.ThisAddIn.CreateModifyWebAgentScript())
    End Sub

    Private Sub RunMarkdownCommand()
        ExecuteLoggedCommand("Markdown_Word invoked", Sub() Globals.ThisAddIn.ConvertMarkdownToWord())
    End Sub

    Private Sub RunResetSpacingCommand()
        ExecuteLoggedCommand("ResetSpacing_Word invoked", Sub() SharedMethods.ResetSelectedTextParagraphSpacing())
    End Sub

    Private Sub RunFindHiddenCommand()
        ExecuteLoggedCommand("FindHidden_Word invoked", Sub() Globals.ThisAddIn.FindHiddenPrompts())
    End Sub

    Private Sub RunContentControlsCommand()
        ExecuteLoggedCommand("ContentControls_Word invoked", Sub() Globals.ThisAddIn.RemoveContentControlsRespectSelection())
    End Sub

    Private Sub RunHelpMeCommand()
        ExecuteLoggedCommand("HelpMe_Word invoked", Sub() Globals.ThisAddIn.HelpMeInky())
    End Sub

    Private Sub RunPrepareRedactionsCommand()
        ExecuteLoggedCommand("PrepareRedactions_Word invoked", Sub() Globals.ThisAddIn.PrepareRedactedPDF())
    End Sub

    Private Sub RunFinalizeRedactionsCommand()
        ExecuteLoggedCommand("FinalizeRedactions_Word invoked", Sub() Globals.ThisAddIn.FlattenRedactedPDF())
    End Sub

    Private Sub RunCheckDocumentsIICommand()
        ExecuteLoggedCommand("CheckDocumentsII_Word invoked", Sub() Globals.ThisAddIn.CheckDocumentII())
    End Sub

    Private Sub RunEditRedactCommand()
        ExecuteLoggedCommand("EditRedact_Word invoked", Sub() Globals.ThisAddIn.EditRedactionInstructions())
    End Sub

    Private Sub RunFilibusterCommand()
        ExecuteLoggedCommand("Filibuster_Word invoked", Sub() Globals.ThisAddIn.Filibuster())
    End Sub

    Private Sub RunArgueAgainstCommand()
        ExecuteLoggedCommand("ArgueAgainst_Word invoked", Sub() Globals.ThisAddIn.ArgueAgainst())
    End Sub

    Private Sub RunLiveCompareCommand()
        ExecuteLoggedCommand("LiveCompare_Word invoked", Sub() Globals.ThisAddIn.CompareActiveDocWithOtherOpenDoc())
    End Sub

    Private Sub RunRevisionSummaryCommand()
        ExecuteLoggedCommand("RevisionSummary_Word invoked", Sub() Globals.ThisAddIn.SummarizeDocumentChanges())
    End Sub

    Private Sub RunDiscussInkyCommand()
        ExecuteLoggedCommand("DiscussInky_Word invoked", Sub() Globals.ThisAddIn.DiscussInky())
    End Sub

    Private Sub RunLearnDocStyleCommand()
        ExecuteLoggedCommand("LearnDocStyle_Word invoked", Sub() Globals.ThisAddIn.ExtractParagraphStylesToJson())
    End Sub

    Private Sub RunApplyDocStyleCommand()
        ExecuteLoggedCommand("ApplyDocStyle_Word invoked", Sub() Globals.ThisAddIn.ApplyStyleTemplate())
    End Sub

    Private Sub RunConvertDocToTxtCommand()
        ExecuteLoggedCommand("ConvertDocToTxt_Word invoked", Sub() Globals.ThisAddIn.ExportFileContentToText())
    End Sub

    Private Sub RunFlattenPdfCommand()
        ExecuteLoggedCommand("FlattenPDF_Word invoked", Sub() Globals.ThisAddIn.FlattenPdfToImages())
    End Sub

    Private Sub RunChartingCommand()
        ExecuteLoggedCommand("Charting_Word invoked", Sub() Globals.ThisAddIn.OpenExistingDrawioFileForEditing())
    End Sub

    Private Sub RunSnapshotCommand()
        ExecuteLoggedCommand("Snapshot_Word invoked", Sub() Globals.ThisAddIn.SelectSnapshotDocument())
    End Sub

    Private Sub RunRemoveCommand()
        ExecuteLoggedCommand("Remove_Word invoked", Sub() Globals.ThisAddIn.RemoveRIPrefixFromComments())
    End Sub

    Private Sub RunWebAppCommand()
        ExecuteLoggedCommand("WebApp_Word invoked", Sub() Globals.ThisAddIn.ConvertDrawioToHtml())
    End Sub

    Private Sub RunSplitPdfCommand()
        ExecuteLoggedCommand("SplitPDF_Word invoked", Sub() Globals.ThisAddIn.SplitPdfByExhibits())
    End Sub

    Private Sub RunStoreOriginalClauseCommand()
        ExecuteLoggedCommand("StoreOriginalClause_Word invoked", Sub() Globals.ThisAddIn.StoreOriginalClause())
    End Sub

    Private Sub RunJustifyMarkupCommand()
        ExecuteLoggedCommand("JustifyMarkup_Word invoked", Sub() Globals.ThisAddIn.JustifyMarkup())
    End Sub

    Private Sub RunBalloonMergePartJustifyCommand()
        ExecuteLoggedCommand("BalloonMergePartJustify_Word invoked", Sub() Globals.ThisAddIn.BalloonMergeWithJustification(False, True))
    End Sub

    Private Sub RunBalloonMergeFullJustifyCommand()
        ExecuteLoggedCommand("BalloonMergeFullJustify_Word invoked", Sub() Globals.ThisAddIn.BalloonMergeWithJustification(True, True))
    End Sub

    Private Sub RunStamperCommand()
        ExecuteLoggedCommand("Stamper_Word invoked", Sub() Globals.ThisAddIn.StampExhibitPDF())
    End Sub

    Private Sub RunImageCommand()
        ExecuteLoggedCommand("GenerateImage_Word invoked", Sub() Globals.ThisAddIn.GenerateImage())
    End Sub

    Private Sub RunTabularCommand()
        ExecuteLoggedCommand("Tabular_Word invoked", Sub() Globals.ThisAddIn.TabularOverview())
    End Sub

    Public Sub RI_Correct_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunCorrectCommand()
    End Sub

    Public Sub RI_Correct2_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunCorrectCommand()
    End Sub

    Public Sub RI_Summarize_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunSummarizeCommand()
    End Sub

    Public Sub RI_Shorten_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunShortenCommand()
    End Sub

    Public Sub RI_PrimLang_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunPrimaryLanguageCommand()
    End Sub

    Public Sub RI_PrimLang2_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunPrimaryLanguageCommand()
    End Sub

    Public Sub RI_SecLang_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunSecondaryLanguageCommand()
    End Sub

    Public Sub RI_Improve_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunImproveCommand()
    End Sub

    Public Sub RI_FreestyleNM_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunFreestyleNmCommand()
    End Sub

    Public Sub RI_Anonymize_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunAnonymizeCommand()
    End Sub

    Public Sub RI_Chat_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunChatCommand()
    End Sub

    Public Sub RI_Chat2_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunChatCommand()
    End Sub

    Public Sub RI_TimeSpan_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunTimeSpanCommand()
    End Sub

    Public Sub RI_AcceptFormat_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunAcceptFormatCommand()
    End Sub

    Private Sub RI_Translate_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunTranslateCommand()
    End Sub

    Private Sub Settings_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunSettingsCommand()
    End Sub

    Private Sub RI_FreestyleAM_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunFreestyleAmCommand()
    End Sub

    Private Sub RI_SwitchParty_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunSwitchPartyCommand()
    End Sub

    Private Sub RI_Regex_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunRegexCommand()
    End Sub

    Private Sub RI_Import_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunImportCommand()
    End Sub

    Private Sub RI_Halves_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunHalvesCommand()
    End Sub

    Private Sub RI_Search_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunSearchCommand()
    End Sub

    Private Sub Easteregg_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunEastereggCommand()
    End Sub

    Private Sub RI_Transcriptor_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunTranscriptorCommand()
    End Sub

    Private Sub RI_TalkToMe_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunTalkToMeCommand()
    End Sub

    Private Sub RI_Explain_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunExplainCommand()
    End Sub

    Private Sub RI_SuggestTitles_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunSuggestTitlesCommand()
    End Sub

    Private Sub RI_CreatePodcast_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunCreatePodcastCommand()
    End Sub

    Private Sub RI_CreateAudio_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunCreateAudioCommand()
    End Sub

    Private Sub RI_NoFillers_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunNoFillersCommand()
    End Sub

    Private Sub RI_Friendly_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunFriendlyCommand()
    End Sub

    Private Sub RI_Convincing_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunConvincingCommand()
    End Sub

    Private Sub RI_SpecialModel_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunSpecialModelCommand()
    End Sub

    Private Sub RI_Anonymization_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        RunAnonymizationCommand()
    End Sub

    Private Sub RI_InsertClipboard_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_InsertClipboard.Click
        RunInsertClipboardCommand()
    End Sub

    Private Sub RI_BallooMergePart_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_BalloonMergePart.Click
        RunBalloonMergePartCommand()
    End Sub

    Private Sub RI_BalloonMergeFull_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_BalloonMergeFull.Click
        RunBalloonMergeFullCommand()
    End Sub

    Private Sub RI_BalloonMergePartPrompt_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_BalloonMergePartPrompt.Click
        RunBalloonMergePartPromptCommand()
    End Sub

    Private Sub RI_BalloonMergeFullPrompt_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_BalloonMergeFullPrompt.Click
        RunBalloonMergeFullPromptCommand()
    End Sub

    Private Sub RI_FreestyleRepeat_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_FreestyleRepeat.Click
        RunFreestyleRepeatCommand()
    End Sub

    Private Sub RI_ApplyMyStyle_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_ApplyMyStyle.Click
        RunApplyMyStyleCommand()
    End Sub

    Private Sub RI_DefineMyStyle_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_DefineMyStyle.Click
        RunDefineMyStyleCommand()
    End Sub

    Private Sub RI_DocCheck_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_DocCheck.Click
        RunDocCheckCommand()
    End Sub

    Private Sub RI_FindClause_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_FindClause.Click
        RunFindClauseCommand()
    End Sub

    Private Sub RI_AddClause_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_AddClause.Click
        RunAddClauseCommand()
    End Sub

    Private Sub RI_WebAgent_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_WebAgent.Click
        RunWebAgentCommand()
    End Sub

    Private Sub RI_EditWebAgent_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_EditWebAgent.Click
        RunEditWebAgentCommand()
    End Sub

    Private Sub RI_Markdown_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_Markdown.Click
        RunMarkdownCommand()
    End Sub

    Private Sub RI_ResetSpacing_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_ResetSpacing.Click
        RunResetSpacingCommand()
    End Sub

    Private Sub RI_FindHidden_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_FindHidden.Click
        RunFindHiddenCommand()
    End Sub

    Private Sub RI_ContentControls_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_ContentControls.Click
        RunContentControlsCommand()
    End Sub

    Private Sub RI_HelpMe_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_HelpMe.Click
        RunHelpMeCommand()
    End Sub

    Private Sub RI_PrepareRedactions_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_PrepareRedactions.Click
        RunPrepareRedactionsCommand()
    End Sub

    Private Sub RI_FinalizeRedactions_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_FinalizeRedactions.Click
        RunFinalizeRedactionsCommand()
    End Sub

    Private Sub RI_CheckDocumentsII_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_CheckDocumentsII.Click
        RunCheckDocumentsIICommand()
    End Sub

    Private Sub RI_EditRedact_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_EditRedact.Click
        RunEditRedactCommand()
    End Sub

    Private Sub RI_Filibuster_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_Filibuster.Click
        RunFilibusterCommand()
    End Sub

    Private Sub RI_ArgueAgainst_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_ArgueAgainst.Click
        RunArgueAgainstCommand()
    End Sub

    Private Sub RI_LiveCompare_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_LiveCompare.Click
        RunLiveCompareCommand()
    End Sub

    Private Sub RI_RevisionSummary_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_RevisionsSummary.Click
        RunRevisionSummaryCommand()
    End Sub

    Private Sub RI_DiscussInky_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_DiscussInky.Click
        RunDiscussInkyCommand()
    End Sub

    Private Sub RI_LearnDocStyle_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_LearnDocStyle.Click
        RunLearnDocStyleCommand()
    End Sub

    Private Sub RI_ApplyDocStyle_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_ApplyDocStyle.Click
        RunApplyDocStyleCommand()
    End Sub

    Private Sub RI_ConvertDocToTxt_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_ConvertDocToTxt.Click
        RunConvertDocToTxtCommand()
    End Sub

    Private Sub RI_FlattenPDF_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_FlattenPDF.Click
        RunFlattenPdfCommand()
    End Sub

    Private Sub RI_Charting_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_Charting.Click
        RunChartingCommand()
    End Sub

    Private Sub RI_Snapshot_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_Snapshot.Click
        RunSnapshotCommand()
    End Sub

    Private Sub RI_Remove_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_Remove.Click
        RunRemoveCommand()
    End Sub

    Private Sub RI_WebApp_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_WebApp.Click
        RunWebAppCommand()
    End Sub

    Private Sub RI_SplitPDF_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_SplitPDF.Click
        RunSplitPdfCommand()
    End Sub

    Private Sub RI_StoreOriginalClause_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_StoreOriginalClause.Click
        RunStoreOriginalClauseCommand()
    End Sub

    Private Sub RI_JustifyMarkup_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_JustifyMarkup.Click
        RunJustifyMarkupCommand()
    End Sub

    Private Sub RI_BalloonMergePartJustify_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_BalloonMergePartJustify.Click
        RunBalloonMergePartJustifyCommand()
    End Sub

    Private Sub RI_BalloonMergeFullJustify_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_BalloonMergeFullJustify.Click
        RunBalloonMergeFullJustifyCommand()
    End Sub

    Private Sub RI_Stamper_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_Stamper.Click
        RunStamperCommand()
    End Sub

    Private Sub RI_Image_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_Image.Click
        RunImageCommand()
    End Sub

    Private Sub RI_Tabular_Click(sender As Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles RI_Tabular.Click
        RunTabularCommand()
    End Sub
End Class
