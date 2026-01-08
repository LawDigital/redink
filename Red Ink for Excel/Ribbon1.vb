' Part of "Red Ink for Excel"
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

    Public Async Function RI_Correct_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_Correct.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Correct_Click invoked")
        Dim Result As Boolean = Await Globals.ThisAddIn.Correct()
    End Function

    Public Async Function RI_Shorten_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_Shorten.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Shorten_Click invoked")
        Dim Result As Boolean = Await Globals.ThisAddIn.Shorten()
    End Function

    Public Async Function RI_PrimLang_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_Primlang.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_PrimLang_Click invoked")
        Dim Result As Boolean = Await Globals.ThisAddIn.InLanguage1()
    End Function

    Public Async Function RI_PrimLang2_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_PrimLang2.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_PrimLang2_Click invoked")
        Dim Result As Boolean = Await Globals.ThisAddIn.InLanguage1()
    End Function

    Public Async Function RI_SecLang_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_SecLang.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_SecLang_Click invoked")
        Dim Result As Boolean = Await Globals.ThisAddIn.InLanguage2()
    End Function

    Public Async Function RI_Improve_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_Improve.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Improve_Click invoked")
        Dim Result As Boolean = Await Globals.ThisAddIn.Improve()
    End Function

    Public Async Function RI_FreestyleNM_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_FreestyleNM.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_FreestyleNM_Click invoked")
        Dim Result As Boolean = Await Globals.ThisAddIn.FreestyleNM()
    End Function

    Public Async Function RI_FreestyleNM2_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_FreestyleNM2.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_FreestyleNM2_Click invoked")
        Dim Result As Boolean = Await Globals.ThisAddIn.FreestyleNM()
    End Function

    Public Async Function RI_Anonymize_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_Anonymize.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Anonymize_Click invoked")
        Dim Result As Boolean = Await Globals.ThisAddIn.Anonymize()
    End Function

    Public Sub RI_AdjustHeight_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_AdjustHeight.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_AdjustHeight_Click invoked")
        Globals.ThisAddIn.AdjustHeight()
    End Sub

    Public Sub RI_AdjustLegacyNotes_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_AdjustLegacyNotes.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_AdjustLegacyNotes_Click invoked")
        Globals.ThisAddIn.AdjustLegacyNotes()
    End Sub

    Private Async Function RI_Translate_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_Translate.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Translate_Click invoked")
        Dim Result As Boolean = Await Globals.ThisAddIn.InOther()
    End Function

    Private Async Function RI_TranslateF_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_TranslateF.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_TranslateF_Click invoked")
        Dim Result As Boolean = Await Globals.ThisAddIn.InOtherFormulas()
    End Function

    Private Sub Settings_Click(sender As Object, e As RibbonControlEventArgs) Handles Settings.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Settings_Click invoked")
        Globals.ThisAddIn.ShowSettings()
    End Sub

    Private Async Function RI_FreestyleAM_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_FreestyleAM.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_FreestyleAM_Click invoked")
        Dim Result As Boolean = Await Globals.ThisAddIn.FreestyleAM()
    End Function

    Private Async Function RI_SwitchParty_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_SwitchParty.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_SwitchParty_Click invoked")
        Dim Result As Boolean = Await Globals.ThisAddIn.SwitchParty()
    End Function

    Private Sub RI_Regex_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Regex.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Regex_Click invoked")
        Globals.ThisAddIn.RegexSearchReplace()
    End Sub

    Private Sub RI_Undo_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Undo.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Undo_Click invoked")
        Globals.ThisAddIn.UndoAction()
    End Sub

    Public Sub RI_Chat_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Chat.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Chat_Click invoked")
        Globals.ThisAddIn.ShowChatForm()
    End Sub

    Public Sub RI_Chat2_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Chat2.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Chat2_Click invoked")
        Globals.ThisAddIn.ShowChatForm()
    End Sub

    Private Sub RI_CSVAnalyze_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_CSVAnalyze.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_CSVAnalyze_Click invoked")
        Globals.ThisAddIn.AnalyzeCsvWithLLM()
    End Sub

    Private Sub RI_HelpMe_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_HelpMe.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_HelpMe_Click invoked")
        Globals.ThisAddIn.HelpMeInky()
    End Sub

    Private Sub RI_Extractor_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Extractor.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Extractor_Click invoked")
        Globals.ThisAddIn.FactExtraction()
    End Sub

    Private Sub RI_Renamer_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Renamer.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Renamer_Click invoked")
        Globals.ThisAddIn.RenameDocumentsWithAi()
    End Sub

End Class