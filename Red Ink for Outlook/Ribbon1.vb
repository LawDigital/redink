' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

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
        ApplyRibbonVisibilityConfiguration()
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

    Public Sub RI_Correct_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Correct.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Correct_Outlook invoked")
        Globals.ThisAddIn.MainMenu("Correct")
    End Sub

    Public Sub RI_Correct2_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Correct2.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Correct_Outlook invoked")
        Globals.ThisAddIn.MainMenu("Correct")
    End Sub

    Public Sub RI_Summarize_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Summarize.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Summarize_Outlook invoked")
        Globals.ThisAddIn.MainMenu("Summarize")
    End Sub

    Public Sub RI_Shorten_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Shorten.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Shorten_Outlook invoked")
        Globals.ThisAddIn.MainMenu("Shorten")
    End Sub

    Public Sub RI_PrimLang_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Primlang.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "PrimLang_Outlook invoked")
        Globals.ThisAddIn.MainMenu("PrimLang")
    End Sub

    Public Sub RI_PrimLang2_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_PrimLang2.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "PrimLang_Outlook invoked")
        Globals.ThisAddIn.MainMenu("PrimLang")
    End Sub

    Public Sub RI_Improve_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Improve.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Improve_Outlook invoked")
        Globals.ThisAddIn.MainMenu("Improve")
    End Sub

    Public Sub RI_Freestyle_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Freestyle.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Freestyle_Outlook invoked")
        Globals.ThisAddIn.MainMenu("Freestyle")
    End Sub

    Public Sub RI_Answers_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Answers.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Answers_Outlook invoked")
        Globals.ThisAddIn.MainMenu("Answers")
    End Sub

    Private Sub RI_Translate_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Translate.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Translate_Outlook invoked")
        Globals.ThisAddIn.MainMenu("Translate")
    End Sub

    Private Sub RI_QuickTranslate_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_QuickTranslate.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "QuickTranslate_Outlook invoked")
        Globals.ThisAddIn.ShowQuickTranslate()
    End Sub



    Private Sub Settings_Click(sender As Object, e As RibbonControlEventArgs) Handles Settings.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Settings_Outlook invoked")
        Globals.ThisAddIn.ShowSettings()
    End Sub

    Private Sub RI_Sumup_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Sumup.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Sumup_Outlook invoked")
        Globals.ThisAddIn.MainMenu("Sumup")
    End Sub

    Private Sub RI_Sumup2_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Sumup2.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Sumup_Outlook invoked")
        Globals.ThisAddIn.MainMenu("Sumup")
    End Sub

    Private Sub RI_NoFillers_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_NoFillers.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "NoFillers_Outlook invoked")
        Globals.ThisAddIn.MainMenu("NoFillers")
    End Sub

    Private Sub RI_Friendly_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Friendly.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Friendly_Outlook invoked")
        Globals.ThisAddIn.MainMenu("Friendly")
    End Sub

    Private Sub RI_Convincing_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Convincing.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Convincing_Outlook invoked")
        Globals.ThisAddIn.MainMenu("Convincing")
    End Sub

    Private Sub RI_Clipboard_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Clipboard.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Clipboard_Outlook invoked")
        Globals.ThisAddIn.MainMenu("InsertClipboard")
    End Sub

    Private Sub RI_ApplyMyStyle_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_ApplyMyStyle.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "ApplyMyStyle_Outlook invoked")
        Globals.ThisAddIn.MainMenu("ApplyMyStyle")
    End Sub

    Private Sub RI_DefineMyStyle_Click_1(sender As Object, e As RibbonControlEventArgs) Handles RI_DefineMyStyle.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "DefineMyStyle_Click_1 invoked")
        Globals.ThisAddIn.DefineMyStyle()
    End Sub

    Private Sub RI_HelpMe_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_HelpMe.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "HelpMe_Outlook invoked")
        Globals.ThisAddIn.HelpMeInky()
    End Sub

    Private Sub RI_CompareSelected_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_CompareSelected.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "CompareSelected_Outlook invoked")
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
        ApplyRibbonVisibilityConfiguration()
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

    Public Sub RI_Correct_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Correct.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Correct_Outlook invoked")
        Globals.ThisAddIn.MainMenu("Correct")
    End Sub

    Public Sub RI_Correct2_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Correct2.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Correct_Outlook invoked")
        Globals.ThisAddIn.MainMenu("Correct")
    End Sub

    Public Sub RI_Summarize_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Summarize.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Summarize_Outlook invoked")
        Globals.ThisAddIn.MainMenu("Summarize")
    End Sub

    Public Sub RI_Shorten_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Shorten.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Shorten_Outlook invoked")
        Globals.ThisAddIn.MainMenu("Shorten")
    End Sub

    Public Sub RI_PrimLang_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Primlang.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "PrimLang_Outlook invoked")
        Globals.ThisAddIn.MainMenu("PrimLang")
    End Sub

    Public Sub RI_PrimLang2_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_PrimLang2.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "PrimLang_Outlook invoked")
        Globals.ThisAddIn.MainMenu("PrimLang")
    End Sub

    Public Sub RI_Improve_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Improve.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Improve_Outlook invoked")
        Globals.ThisAddIn.MainMenu("Improve")
    End Sub

    Public Sub RI_Freestyle_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Freestyle.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Freestyle_Outlook invoked")
        Globals.ThisAddIn.MainMenu("Freestyle")
    End Sub

    Public Sub RI_Answers_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Answers.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Answers_Outlook invoked")
        Globals.ThisAddIn.MainMenu("Answers")
    End Sub

    Private Sub RI_Translate_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Translate.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Translate_Outlook invoked")
        Globals.ThisAddIn.MainMenu("Translate")
    End Sub

    Private Sub RI_QuickTranslate_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_QuickTranslate.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "QuickTranslate_Outlook invoked")
        Globals.ThisAddIn.ShowQuickTranslate()
    End Sub


    Private Sub Settings_Click(sender As Object, e As RibbonControlEventArgs) Handles Settings.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Settings_Outlook invoked")
        Globals.ThisAddIn.ShowSettings()
    End Sub

    Private Sub RI_Sumup_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Sumup.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Sumup_Outlook invoked")
        Globals.ThisAddIn.MainMenu("Sumup")
    End Sub

    Private Sub RI_Sumup2_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Sumup2.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Sumup_Outlook invoked")
        Globals.ThisAddIn.MainMenu("Sumup")
    End Sub

    Private Sub RI_NoFillers_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_NoFillers.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "NoFillers_Outlook invoked")
        Globals.ThisAddIn.MainMenu("NoFillers")
    End Sub

    Private Sub RI_Friendly_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Friendly.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Friendly_Outlook invoked")
        Globals.ThisAddIn.MainMenu("Friendly")
    End Sub

    Private Sub RI_Convincing_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Convincing.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Convincing_Outlook invoked")
        Globals.ThisAddIn.MainMenu("Convincing")
    End Sub

    Private Sub RI_Clipboard_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Clipboard.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Clipboard_Outlook invoked")
        Globals.ThisAddIn.MainMenu("InsertClipboard")
    End Sub

    Private Sub RI_ApplyMyStyle_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_ApplyMyStyle.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "ApplyMyStyle_Outlook invoked")
        Globals.ThisAddIn.MainMenu("ApplyMyStyle")
    End Sub

    Private Sub RI_M365_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_M365.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "M365_Outlook invoked")
        Globals.ThisAddIn.MainMenu("M365")
    End Sub

    Private Sub RI_MailMover_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_MailMover.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "MailMover_Outlook invoked")
        Globals.ThisAddIn.MainMenu("MailMover")
    End Sub

    Private Sub RI_InboxBoard_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_InboxBoard.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "InboxBoard_Outlook invoked")
        Globals.ThisAddIn.MainMenu("InboxBoard")
    End Sub

    Private Sub RI_DefineMyStyle_Click_1(sender As Object, e As RibbonControlEventArgs) Handles RI_DefineMyStyle.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "DefineMyStyle_Outlook invoked")
        Globals.ThisAddIn.DefineMyStyle()
    End Sub

    Private Sub RI_KnowledgeStores_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_KnowledgeStores.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "KnowledgeStores_Outlook invoked")
        Using frm As New KnowledgeStoreForm(ThisAddIn._context)
            frm.ShowDialog()
        End Using
    End Sub

    Private Sub RI_HelpMe_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_HelpMe.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "HelpMe_Outlook invoked")
        Globals.ThisAddIn.HelpMeInky()
    End Sub

    Private Sub RI_CompareSelected_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_CompareSelected.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "CompareSelected_Outlook invoked")
        Globals.ThisAddIn.CompareSelectedTextRangesOutlook()
    End Sub

    Private Sub RI_AutoPilot_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_AutoPilot.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "AutoPilot_Outlook invoked")
        Globals.ThisAddIn.StartAutoPilot()
    End Sub

    Private Sub RI_OpenChat_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_OpenChat.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "OpenChat_Outlook invoked")
        Dim startInfo As New ProcessStartInfo(SharedMethods.AN7) With {
            .UseShellExecute = True
        }
        Process.Start(startInfo)
    End Sub

    Private Sub RI_PSTExport_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_PSTExport.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "PSTExport_Outlook invoked")
        Globals.ThisAddIn.ExportPstContentToText()
    End Sub


End Class
