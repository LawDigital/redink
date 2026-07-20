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
            System.Diagnostics.Debug.WriteLine("SetRibbonControlVisibleByAvailability: ribbonControl is Nothing.")
            Return
        End If

        If System.String.IsNullOrWhiteSpace(ribbonControl.Name) Then
            System.Diagnostics.Debug.WriteLine("SetRibbonControlVisibleByAvailability: Ribbon control without name.")
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

    Public Sub ApplySimpleModeToRibbonControls(
        ByVal controlNames As String,
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



    ' =========================================================================
    ' Model Selection
    ' =========================================================================

    Public Sub UpdateModelsMenu()
        Try
            Dim available = PrimaryModelManager.GetAvailableModels()
            Dim current = PrimaryModelManager.GetCurrentModelNumber()

            SetRibbonControlVisibleByAvailability(Menu4, available.Count > 1)

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

    Private Sub RI_Model1_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Model1.Click
        ThisAddIn.SelectModel(1)
    End Sub

    Private Sub RI_Model2_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Model2.Click
        ThisAddIn.SelectModel(2)
    End Sub

    Private Sub RI_Model3_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Model3.Click
        ThisAddIn.SelectModel(3)
    End Sub

    Private Sub RI_Model4_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Model4.Click
        ThisAddIn.SelectModel(4)
    End Sub

    Private Sub RI_Model5_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Model5.Click
        ThisAddIn.SelectModel(5)
    End Sub

    Private Sub RI_Model6_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Model6.Click
        ThisAddIn.SelectModel(6)
    End Sub

    Private Sub RI_Model7_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Model7.Click
        ThisAddIn.SelectModel(7)
    End Sub

    Private Sub RI_Model8_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Model8.Click
        ThisAddIn.SelectModel(8)
    End Sub

    Private Sub RI_Model9_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Model9.Click
        ThisAddIn.SelectModel(9)
    End Sub

    Private Sub RI_Model10_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Model10.Click
        ThisAddIn.SelectModel(10)
    End Sub


    Public Async Function RI_Correct_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_Correct.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Correct_Excel invoked")
        Dim Result As Boolean = Await Globals.ThisAddIn.Correct()
    End Function

    Public Async Function RI_Shorten_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_Shorten.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Shorten_Excel invoked")
        Dim Result As Boolean = Await Globals.ThisAddIn.Shorten()
    End Function

    Public Async Function RI_PrimLang_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_Primlang.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "PrimLang_Excel invoked")
        Dim Result As Boolean = Await Globals.ThisAddIn.InLanguage1()
    End Function

    Public Async Function RI_PrimLang2_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_PrimLang2.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "PrimLang2_Excel invoked")
        Dim Result As Boolean = Await Globals.ThisAddIn.InLanguage1()
    End Function

    Public Async Function RI_SecLang_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_SecLang.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "SecLang_Excel invoked")
        Dim Result As Boolean = Await Globals.ThisAddIn.InLanguage2()
    End Function

    Public Async Function RI_Improve_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_Improve.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Improve_Excel invoked")
        Dim Result As Boolean = Await Globals.ThisAddIn.Improve()
    End Function

    Public Async Function RI_FreestyleNM_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_FreestyleNM.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "FreestyleNM_Excel invoked")
        Dim Result As Boolean = Await Globals.ThisAddIn.FreestyleNM()
    End Function

    Public Async Function RI_FreestyleNM2_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_FreestyleNM2.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "FreestyleNM_Excel invoked")
        Dim Result As Boolean = Await Globals.ThisAddIn.FreestyleNM()
    End Function

    Public Async Function RI_Anonymize_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_Anonymize.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Anonymize_Excel invoked")
        Dim Result As Boolean = Await Globals.ThisAddIn.Anonymize()
    End Function

    Public Sub RI_AdjustHeight_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_AdjustHeight.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "AdjustHeight_Excel invoked")
        Globals.ThisAddIn.AdjustHeight()
    End Sub

    Public Sub RI_AdjustLegacyNotes_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_AdjustLegacyNotes.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "AdjustLegacyNotes_Excel invoked")
        Globals.ThisAddIn.AdjustLegacyNotes()
    End Sub

    Private Async Function RI_Translate_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_Translate.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Translate_Excel invoked")
        Dim Result As Boolean = Await Globals.ThisAddIn.InOther()
    End Function

    Private Async Function RI_TranslateF_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_TranslateF.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Translate_Excel invoked")
        Dim Result As Boolean = Await Globals.ThisAddIn.InOtherFormulas()
    End Function

    Private Sub Settings_Click(sender As Object, e As RibbonControlEventArgs) Handles Settings.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Settings_Excel invoked")
        Globals.ThisAddIn.ShowSettings()
    End Sub

    Private Async Function RI_FreestyleAM_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_FreestyleAM.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "FreestyleAM_Excel invoked")
        Dim Result As Boolean = Await Globals.ThisAddIn.FreestyleAM()
    End Function

    Private Async Function RI_SwitchParty_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_SwitchParty.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "SwitchParty_Excel invoked")
        Dim Result As Boolean = Await Globals.ThisAddIn.SwitchParty()
    End Function

    Private Sub RI_Regex_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Regex.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Regex_Excel invoked")
        Globals.ThisAddIn.RegexSearchReplace()
    End Sub

    Private Sub RI_Undo_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Undo.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Undo_Excel invoked")
        Globals.ThisAddIn.UndoAction()
    End Sub

    Public Sub RI_Chat_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Chat.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Chat_Excel invoked")
        Globals.ThisAddIn.ShowChatForm()
    End Sub

    Public Sub RI_Chat2_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Chat2.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Chat_Excel invoked")
        Globals.ThisAddIn.ShowChatForm()
    End Sub

    Private Sub RI_CSVAnalyze_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_CSVAnalyze.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "CSVAnalyze_Excel invoked")
        Globals.ThisAddIn.AnalyzeCsvWithLLM()
    End Sub

    Private Sub RI_HelpMe_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_HelpMe.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "HelpMe_Excel invoked")
        Globals.ThisAddIn.HelpMeInky()
    End Sub

    Private Sub RI_Extractor_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Extractor.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Extractor_Excel invoked")
        Globals.ThisAddIn.FactExtraction()
    End Sub

    Private Sub RI_Renamer_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Renamer.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Renamer_Excel invoked")
        Globals.ThisAddIn.RenameDocumentsWithAi()
    End Sub

    Private Sub RI_RIRemove_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_RIRemove.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RIRemove_Excel invoked")
        Globals.ThisAddIn.RemoveRIPrefixFromComments()
    End Sub

    Private Sub RI_TXTAnalyze_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_TXTAnalyze.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "TXTAnalyze_Excel invoked")
        Globals.ThisAddIn.AnalyzeDirectoryWithLLM()
    End Sub


End Class
