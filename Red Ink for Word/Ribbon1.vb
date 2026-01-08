' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

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
                    Menu1.Image = SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Medium)
                Case Else
                    Menu1.Image = SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard)
            End Select
            Menu1.ShowImage = True
        Catch
            Menu1.Image = SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard)
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

    Public Sub RI_Correct_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Correct.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Correct_Click invoked")
        Globals.ThisAddIn.Correct()
    End Sub

    Public Sub RI_Correct2_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Correct2.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Correct_Click2 invoked")
        Globals.ThisAddIn.Correct()
    End Sub

    Public Sub RI_Summarize_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Summarize.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Summarize_Click invoked")
        Globals.ThisAddIn.Summarize()
    End Sub

    Public Sub RI_Shorten_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Shorten.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Shorten_Click invoked")
        Globals.ThisAddIn.Shorten()
    End Sub

    Public Sub RI_PrimLang_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Primlang.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_PrimLang_Click invoked")
        Globals.ThisAddIn.InLanguage1()
    End Sub

    Public Sub RI_PrimLang2_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_PrimLang2.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_PrimLang2_Click invoked")
        Globals.ThisAddIn.InLanguage1()
    End Sub

    Public Sub RI_SecLang_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_SecLang.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_SecLang_Click invoked")
        Globals.ThisAddIn.InLanguage2()
    End Sub
    Public Sub RI_Improve_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Improve.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Improve_Click invoked")
        Globals.ThisAddIn.Improve()
    End Sub

    Public Sub RI_FreestyleNM_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_FreestyleNM.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_FreestyleNM_Click invoked")
        Globals.ThisAddIn.FreeStyleNM()
    End Sub

    Public Sub RI_Anonymize_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Anonymize.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Anonymize_Click invoked")
        Globals.ThisAddIn.Anonymize()
    End Sub

    Public Sub RI_Chat_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Chat.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Chat_Click invoked")
        Globals.ThisAddIn.ShowChatForm()
    End Sub

    Public Sub RI_Chat2_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Chat2.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Chat2_Click invoked")
        Globals.ThisAddIn.ShowChatForm()
    End Sub

    Public Sub RI_TimeSpan_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_TimeSpan.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_TimeSpan_Click invoked")
        Globals.ThisAddIn.CalculateUserMarkupTimeSpan()
    End Sub

    Public Sub RI_AcceptFormat_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_AcceptFormat.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_AcceptFormat_Click invoked")
        Globals.ThisAddIn.AcceptFormatting()
    End Sub

    Private Sub RI_Translate_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Translate.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Translate_Click invoked")
        Globals.ThisAddIn.InOther()
    End Sub

    Private Sub Settings_Click(sender As Object, e As RibbonControlEventArgs) 'Handles Settings.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Settings_Click invoked")
        Globals.ThisAddIn.ShowSettings()
    End Sub

    Private Sub RI_FreestyleAM_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_FreestyleAM.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_FreestyleAM_Click invoked")
        Globals.ThisAddIn.FreeStyleAM()
    End Sub

    Private Sub RI_SwitchParty_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_SwitchParty.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_SwitchParty_Click invoked")
        Globals.ThisAddIn.SwitchParty()
    End Sub

    Private Sub RI_Regex_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Regex.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Regex_Click invoked")
        Globals.ThisAddIn.RegexSearchReplace()
    End Sub

    Private Sub RI_Import_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Import.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Import_Click invoked")
        Globals.ThisAddIn.ImportTextFile()
    End Sub

    Private Sub RI_Halves_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Halves.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Halves_Click invoked")
        Globals.ThisAddIn.CompareSelectionHalves()
    End Sub

    Private Sub RI_Search_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Import.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Search_Click invoked")
        Globals.ThisAddIn.ContextSearch()
    End Sub

    Private Sub Easteregg_Click(sender As Object, e As RibbonControlEventArgs)
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Easteregg_Click invoked")
        Globals.ThisAddIn.EasterEgg()
    End Sub

    Private Sub RI_Transcriptor_Click(sender As Object, e As RibbonControlEventArgs)
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Transcriptor_Click invoked")
        Globals.ThisAddIn.Transcriptor()
    End Sub

    Private Sub RI_Explain_Click(sender As Object, e As RibbonControlEventArgs)
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Explain_Click invoked")
        Globals.ThisAddIn.Explain()
    End Sub

    Private Sub RI_SuggestTitles_Click(sender As Object, e As RibbonControlEventArgs)
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_SuggestTitles_Click invoked")
        Globals.ThisAddIn.SuggestTitles()
    End Sub

    Private Sub RI_CreatePodcast_Click(sender As Object, e As RibbonControlEventArgs)
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_CreatePodcast_Click invoked")
        Globals.ThisAddIn.CreatePodcast()
    End Sub

    Private Sub RI_CreateAudio_Click(sender As Object, e As RibbonControlEventArgs)
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_CreateAudio_Click invoked")
        Globals.ThisAddIn.CreateAudio()
    End Sub

    Private Sub RI_NoFillers_Click(sender As Object, e As RibbonControlEventArgs)
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_NoFillers_Click invoked")
        Globals.ThisAddIn.NoFillers()
    End Sub

    Private Sub RI_Friendly_Click(sender As Object, e As RibbonControlEventArgs)
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Friendly_Click invoked")
        Globals.ThisAddIn.Friendly()
    End Sub

    Private Sub RI_Convincing_Click(sender As Object, e As RibbonControlEventArgs)
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Convincing_Click invoked")
        Globals.ThisAddIn.Convincing()
    End Sub

    Private Sub RI_SpecialModel_Click(sender As Object, e As RibbonControlEventArgs)
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_SpecialModel_Click invoked")
        Globals.ThisAddIn.SpecialModel()
    End Sub

    Private Sub RI_Anonymization_Click(sender As Object, e As RibbonControlEventArgs)
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Anonymization_Click invoked")
        Globals.ThisAddIn.AnonymizeSelection()
    End Sub

    Private Sub RI_InsertClipboard_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_InsertClipboard.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_InsertClipboard_Click invoked")
        Globals.ThisAddIn.InsertClipboard()
    End Sub

    Private Sub RI_BallooMergePart_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_BalloonMergePart.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_BallooMergePart_Click invoked")
        Globals.ThisAddIn.BalloonMerge(False, True)
    End Sub

    Private Sub RI_BalloonMergeFull_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_BalloonMergeFull.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_BalloonMergeFull_Click invoked")
        Globals.ThisAddIn.BalloonMerge(True, True)
    End Sub

    Private Sub RI_BalloonMergePartPrompt_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_BalloonMergePartPrompt.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_BalloonMergePartPrompt_Click invoked")
        Globals.ThisAddIn.BalloonMerge(False, False)
    End Sub

    Private Sub RI_BalloonMergeFullPrompt_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_BalloonMergeFullPrompt.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_BalloonMergeFullPrompt_Click invoked")
        Globals.ThisAddIn.BalloonMerge(True, False)
    End Sub

    Private Sub RI_FreestyleRepeat_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_FreestyleRepeat.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_FreestyleRepeat_Click invoked")
        Globals.ThisAddIn.FreeStyleRepeat()
    End Sub

    Private Sub RI_ApplyMyStyle_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_ApplyMyStyle.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_ApplyMyStyle_Click invoked")
        Globals.ThisAddIn.ApplyMyStyle()
    End Sub

    Private Sub RI_DefineMyStyle_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_DefineMyStyle.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_DefineMyStyle_Click invoked")
        Globals.ThisAddIn.DefineMyStyle()
    End Sub

    Private Sub RI_DocCheck_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_DocCheck.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_DocCheck_Click invoked")
        Globals.ThisAddIn.RunDocCheck()
    End Sub

    Private Sub RI_FindClause_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_FindClause.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_FindClause_Click invoked")
        Globals.ThisAddIn.FindClause()
    End Sub

    Private Sub RI_AddClause_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_AddClause.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_AddClause_Click invoked")
        Globals.ThisAddIn.AddClause()
    End Sub

    Private Sub RI_WebAgent_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_WebAgent.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_WebAgent_Click invoked")
        Globals.ThisAddIn.WebAgent()
    End Sub

    Private Sub RI_EditWebAgent_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_EditWebAgent.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_EditWebAgent_Click invoked")
        Globals.ThisAddIn.CreateModifyWebAgentScript()
    End Sub

    Private Sub RI_Markdown_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Markdown.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Markdown_Click invoked")
        Globals.ThisAddIn.ConvertMarkdownToWord()
    End Sub

    Private Sub RI_FindHidden_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_FindHidden.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_FindHidden_Click invoked")
        Globals.ThisAddIn.FindHiddenPrompts()
    End Sub

    Private Sub RI_ContentControls_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_ContentControls.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_ContentControls_Click invoked")
        Globals.ThisAddIn.RemoveContentControlsRespectSelection()
    End Sub

    Private Sub RI_HelpMe_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_HelpMe.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_HelpMe_Click invoked")
        Globals.ThisAddIn.HelpMeInky()
    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_PrepareRedactions.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_PrepareRedactions_Click invoked")
        Globals.ThisAddIn.PrepareRedactedPDF()
    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_FinalizeRedactions.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_FinalizeRedactions_Click invoked")
        Globals.ThisAddIn.FlattenRedactedPDF()
    End Sub

    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_CheckDocumentsII.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_CheckDocumentsII_Click invoked")
        Globals.ThisAddIn.CheckDocumentII()
    End Sub

    Private Sub RI_EditRedact_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_EditRedact.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_EditRedact_Click invoked")
        Globals.ThisAddIn.EditRedactionInstructions()
    End Sub

    Private Sub RI_Filibuster_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Filibuster.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Filibuster_Click invoked")
        Globals.ThisAddIn.Filibuster()
    End Sub

    Private Sub RI_ArgueAgainst_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_ArgueAgainst.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_ArgueAgainst_Click invoked")
        Globals.ThisAddIn.ArgueAgainst()
    End Sub

    Private Sub RI_LiveCompare_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_LiveCompare.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_LiveCompare_Click invoked")
        Globals.ThisAddIn.CompareActiveDocWithOtherOpenDoc()
    End Sub

    Private Sub RI_RevisionSummary_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_RevisionsSummary.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_RevisionSummary_Click invoked")
        Globals.ThisAddIn.SummarizeDocumentChanges()
    End Sub

    Private Sub RI_DiscussInky_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_DiscussInky.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_DiscussInky_Click invoked")
        Globals.ThisAddIn.DiscussInky()
    End Sub

    Private Sub RI_LearnDocStyle_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_LearnDocStyle.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_LearnDocStyle_Click invoked")
        Globals.ThisAddIn.ExtractParagraphStylesToJson()
    End Sub

    Private Sub RI_ApplyDocStyle_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_ApplyDocStyle.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_ApplyDocStyle_Click invoked")
        Globals.ThisAddIn.ApplyStyleTemplate()
    End Sub

    Private Sub RI_ConvertDocToTxt_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_ConvertDocToTxt.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_ConvertDocToTxt_Click invoked")
        Globals.ThisAddIn.ExportFileContentToText()
    End Sub
End Class