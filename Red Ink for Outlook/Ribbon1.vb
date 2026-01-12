' Part of "Red Ink for Outlook"
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
                Case OfficeTheme.Dark
                    Menu1.Image = SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard)
                Case Else
                    Menu1.Image = SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Medium)
            End Select
            Menu1.ShowImage = True
        Catch
            Menu1.Image = SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard)
            Menu1.ShowImage = True
        End Try
    End Sub

    Private Sub Ribbon1_Load(sender As Object, e As RibbonUIEventArgs) Handles MyBase.Load
        ApplyThemeAwareMenuIcon()
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

    Public Sub RI_Correct_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Correct.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Correct_Click invoked")
        Globals.ThisAddIn.MainMenu("Correct")
    End Sub

    Public Sub RI_Correct2_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Correct2.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Correct2_Click invoked")
        Globals.ThisAddIn.MainMenu("Correct")
    End Sub

    Public Sub RI_Summarize_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Summarize.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Summarize_Click invoked")
        Globals.ThisAddIn.MainMenu("Summarize")
    End Sub

    Public Sub RI_Shorten_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Shorten.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Shorten_Click invoked")
        Globals.ThisAddIn.MainMenu("Shorten")
    End Sub

    Public Sub RI_PrimLang_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Primlang.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_PrimLang_Click invoked")
        Globals.ThisAddIn.MainMenu("PrimLang")
    End Sub

    Public Sub RI_PrimLang2_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_PrimLang2.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_PrimLang2_Click invoked")
        Globals.ThisAddIn.MainMenu("PrimLang")
    End Sub

    Public Sub RI_Improve_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Improve.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Improve_Click invoked")
        Globals.ThisAddIn.MainMenu("Improve")
    End Sub

    Public Sub RI_Freestyle_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Freestyle.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Freestyle_Click invoked")
        Globals.ThisAddIn.MainMenu("Freestyle")
    End Sub

    Public Sub RI_Answers_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Answers.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Answers_Click invoked")
        Globals.ThisAddIn.MainMenu("Answers")
    End Sub

    Private Sub RI_Translate_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Translate.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Translate_Click invoked")
        Globals.ThisAddIn.MainMenu("Translate")
    End Sub

    Private Sub RI_QuickTranslate_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_QuickTranslate.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_QuickTranslate_Click invoked")
        Globals.ThisAddIn.ShowQuickTranslate()
    End Sub



    Private Sub Settings_Click(sender As Object, e As RibbonControlEventArgs) Handles Settings.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Settings_Click invoked")
        Globals.ThisAddIn.ShowSettings()
    End Sub

    Private Sub Sumup_Click(sender As Object, e As RibbonControlEventArgs) Handles Sumup.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Sumup_Click invoked")
        Globals.ThisAddIn.MainMenu("Sumup")
    End Sub

    Private Sub Sumup2_Click(sender As Object, e As RibbonControlEventArgs) Handles Sumup2.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Sumup2_Click invoked")
        Globals.ThisAddIn.MainMenu("Sumup")
    End Sub

    Private Sub RI_NoFillers_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_NoFillers.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_NoFillers_Click invoked")
        Globals.ThisAddIn.MainMenu("NoFillers")
    End Sub

    Private Sub RI_Friendly_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Friendly.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Friendly_Click invoked")
        Globals.ThisAddIn.MainMenu("Friendly")
    End Sub

    Private Sub RI_Convincing_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Convincing.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Convincing_Click invoked")
        Globals.ThisAddIn.MainMenu("Convincing")
    End Sub

    Private Sub RI_Clipboard_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Clipboard.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Clipboard_Click invoked")
        Globals.ThisAddIn.MainMenu("InsertClipboard")
    End Sub

    Private Sub RI_ApplyMyStyle_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_ApplyMyStyle.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_ApplyMyStyle_Click invoked")
        Globals.ThisAddIn.MainMenu("ApplyMyStyle")
    End Sub

    Private Sub RI_DefineMyStyle_Click_1(sender As Object, e As RibbonControlEventArgs) Handles RI_DefineMyStyle.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_DefineMyStyle_Click_1 invoked")
        Globals.ThisAddIn.DefineMyStyle()
    End Sub

    Private Sub RI_HelpMe_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_HelpMe.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_HelpMe_Click invoked")
        Globals.ThisAddIn.HelpMeInky()
    End Sub

    Private Sub RI_CompareSelected_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_CompareSelected.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_CompareSelected_Click invoked")
        Globals.ThisAddIn.CompareSelectedTextRangesOutlook()
    End Sub

End Class

Public Class Ribbon2


    Private Enum OfficeTheme
        Unknown
        Light
        Dark
    End Enum

    Public Sub ApplyThemeAwareMenuIcon()
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

    Private Sub Ribbon2_Load(sender As Object, e As RibbonUIEventArgs) Handles MyBase.Load
        ApplyThemeAwareMenuIcon()
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
                If key Is Nothing Then Return True
                Dim raw = key.GetValue(appsUseLightTheme)
                If raw Is Nothing Then Return True
                Dim v As Integer
                If Integer.TryParse(raw.ToString(), v) Then
                    Return v <> 0
                End If
            End Using
        Catch
        End Try
        Return True
    End Function

    Public Sub RI_Correct_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Correct.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Correct_Click invoked")
        Globals.ThisAddIn.MainMenu("Correct")
    End Sub

    Public Sub RI_Correct2_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Correct2.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Correct2_Click invoked")
        Globals.ThisAddIn.MainMenu("Correct")
    End Sub

    Public Sub RI_Summarize_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Summarize.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Summarize_Click invoked")
        Globals.ThisAddIn.MainMenu("Summarize")
    End Sub

    Public Sub RI_Shorten_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Shorten.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Shorten_Click invoked")
        Globals.ThisAddIn.MainMenu("Shorten")
    End Sub

    Public Sub RI_PrimLang_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Primlang.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_PrimLang_Click invoked")
        Globals.ThisAddIn.MainMenu("PrimLang")
    End Sub

    Public Sub RI_PrimLang2_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_PrimLang2.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_PrimLang2_Click invoked")
        Globals.ThisAddIn.MainMenu("PrimLang")
    End Sub

    Public Sub RI_Improve_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Improve.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Improve_Click invoked")
        Globals.ThisAddIn.MainMenu("Improve")
    End Sub

    Public Sub RI_Freestyle_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Freestyle.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Freestyle_Click invoked")
        Globals.ThisAddIn.MainMenu("Freestyle")
    End Sub

    Public Sub RI_Answers_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Answers.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Answers_Click invoked")
        Globals.ThisAddIn.MainMenu("Answers")
    End Sub

    Private Sub RI_Translate_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Translate.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Translate_Click invoked")
        Globals.ThisAddIn.MainMenu("Translate")
    End Sub

    Private Sub RI_QuickTranslate_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_QuickTranslate.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_QuickTranslate_Click invoked")
        Globals.ThisAddIn.ShowQuickTranslate()
    End Sub


    Private Sub Settings_Click(sender As Object, e As RibbonControlEventArgs) Handles Settings.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Settings_Click invoked")
        Globals.ThisAddIn.ShowSettings()
    End Sub

    Private Sub Sumup_Click(sender As Object, e As RibbonControlEventArgs) Handles Sumup.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Sumup_Click invoked")
        Globals.ThisAddIn.MainMenu("Sumup")
    End Sub

    Private Sub Sumup2_Click(sender As Object, e As RibbonControlEventArgs) Handles Sumup2.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Sumup2_Click invoked")
        Globals.ThisAddIn.MainMenu("Sumup")
    End Sub

    Private Sub RI_NoFillers_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_NoFillers.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_NoFillers_Click invoked")
        Globals.ThisAddIn.MainMenu("NoFillers")
    End Sub

    Private Sub RI_Friendly_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Friendly.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Friendly_Click invoked")
        Globals.ThisAddIn.MainMenu("Friendly")
    End Sub

    Private Sub RI_Convincing_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Convincing.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Convincing_Click invoked")
        Globals.ThisAddIn.MainMenu("Convincing")
    End Sub

    Private Sub RI_Clipboard_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Clipboard.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_Clipboard_Click invoked")
        Globals.ThisAddIn.MainMenu("InsertClipboard")
    End Sub

    Private Sub RI_ApplyMyStyle_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_ApplyMyStyle.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_ApplyMyStyle_Click invoked")
        Globals.ThisAddIn.MainMenu("ApplyMyStyle")
    End Sub

    Private Sub RI_DefineMyStyle_Click_1(sender As Object, e As RibbonControlEventArgs) Handles RI_DefineMyStyle.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_DefineMyStyle_Click_1 invoked")
        Globals.ThisAddIn.DefineMyStyle()
    End Sub

    Private Sub RI_HelpMe_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_HelpMe.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_HelpMe_Click invoked")
        Globals.ThisAddIn.HelpMeInky()
    End Sub

    Private Sub RI_CompareSelected_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_CompareSelected.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RI_CompareSelected_Click invoked")
        Globals.ThisAddIn.CompareSelectedTextRangesOutlook()
    End Sub

End Class