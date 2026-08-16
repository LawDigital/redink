' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.RegistryBackup.vb
' Purpose: Registry-based backup/restore of AutoPilot configuration to survive
'          My.Settings deletions caused by VSTO/M365 updates or profile resets.
'
' Architecture / How it works:
'  - After AutoPilot config is saved to My.Settings, the same values are written
'    as an encoded JSON blob to the registry under:
'       HKEY_CURRENT_USER\Software\Red Ink\AutoPilot
'  - On startup, if no saved AutoPilot config exists in My.Settings, the registry
'    backup is read and restored silently before auto-start logic continues.
'  - All operations are silent and fail-safe; exceptions are swallowed with
'    Debug logging only.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports Microsoft.Win32
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

#Region "Registry AutoPilot Backup"

    ''' <summary>
    ''' Full registry path for the AutoPilot backup value.
    ''' Resolves to: HKEY_CURRENT_USER\Software\Red Ink\AutoPilot
    ''' </summary>
    Private Shared ReadOnly RegistryAutoPilotFullPath As String = RegPath_Base & "AutoPilot"

    ''' <summary>
    ''' Saves the current AutoPilot My.Settings payload to the registry as an encoded JSON blob.
    ''' Fully silent — never shows UI, never throws.
    ''' </summary>
    Private Sub BackupAutoPilotSettingsToRegistry()
        Try
            Dim json As New JObject From {
                {"T", "AutoPilot"},
                {"V", 5},
                {"AP_FilterRules", If(My.Settings.AP_FilterRules, "")},
                {"AP_WhitelistedSenders", If(My.Settings.AP_WhitelistedSenders, "")},
                {"AP_SubjectTriggerWord", If(My.Settings.AP_SubjectTriggerWord, "")},
                {"AP_CooldownSeconds", My.Settings.AP_CooldownSeconds},
                {"AP_MaxRepliesPerSession", My.Settings.AP_MaxRepliesPerSession},
                {"AP_MaxAttachmentMB", My.Settings.AP_MaxAttachmentMB},
                {"AP_FooterText", If(My.Settings.AP_FooterText, "")},
                {"AP_RequireApproval", My.Settings.AP_RequireApproval},
                {"AP_MonitoredMailbox", If(My.Settings.AP_MonitoredMailbox, "")},
                {"AP_SelectedModelKey", If(My.Settings.AP_SelectedModelKey, "")},
                {"AP_UseSecondApi", My.Settings.AP_UseSecondApi},
                {"AP_ReprocessLookbackHours", My.Settings.AP_ReprocessLookbackHours},
                {"AP_AutoDeleteAfterHours", My.Settings.AP_AutoDeleteAfterHours},
                {"AP_EnableWebGrounding", My.Settings.AP_EnableWebGrounding},
                {"AP_EnableVoicemailProcessing", My.Settings.AP_EnableVoicemailProcessing},
                {"AP_VoicemailSenderAddress", If(My.Settings.AP_VoicemailSenderAddress, "")},
                {"AP_VoicemailCallerIdMapPath", If(My.Settings.AP_VoicemailCallerIdMapPath, "")},
                {"AP_EnableScheduler", My.Settings.AP_EnableScheduler},
                {"AP_EnableUserMemory", My.Settings.AP_EnableUserMemory},
                {"AP_EnableUserFiles", My.Settings.AP_EnableUserFiles},
                {"AP_EnablePrivacyProtection", My.Settings.AP_EnablePrivacyProtection},
                {"AP_SelectedExternalToolNames", If(My.Settings.AP_SelectedExternalToolNames, "")},
                {"AP_DashboardWindowX", My.Settings.AP_DashboardWindowX},
                {"AP_DashboardWindowY", My.Settings.AP_DashboardWindowY},
                {"AP_DashboardWindowW", My.Settings.AP_DashboardWindowW},
                {"AP_DashboardWindowH", My.Settings.AP_DashboardWindowH},
                {"AP_SenderToolPolicyPath", If(My.Settings.AP_SenderToolPolicyPath, "")},
                {"AP_ThreadRetentionDays", My.Settings.AP_ThreadRetentionDays},
                {"D", Date.UtcNow.ToString("o")}
            }

            Dim encoded As String = CodeString(json.ToString(Newtonsoft.Json.Formatting.None), SK)
            WriteAutoPilotRegistryValue(encoded)

            Debug.WriteLine("[AutoPilot] Registry backup saved.")
        Catch ex As Exception
            Debug.WriteLine($"[AutoPilot] Failed to save registry backup: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Persists AutoPilot settings to My.Settings and refreshes the registry backup.
    ''' Keeps the reinstall-safe backup in sync with the latest saved values.
    ''' </summary>
    Private Sub SaveAutoPilotSettingsWithRegistryBackup()
        My.Settings.Save()
        BackupAutoPilotSettingsToRegistry()
    End Sub

    ''' <summary>
    ''' Ensures the registry-backed AutoPilot settings are restored at most once per session.
    ''' This avoids duplicate restore attempts when both auto-start and the manual dialog path
    ''' touch the same recovery logic during the same Outlook session.
    ''' </summary>
    Private _autoPilotRegistrySettingsRestoredThisSession As Boolean = False

    Private Sub EnsureAutoPilotSettingsRestoredFromRegistry()
        If _autoPilotRegistrySettingsRestoredThisSession Then
            Return
        End If

        TryRestoreAutoPilotSettingsFromRegistry()
        _autoPilotRegistrySettingsRestoredThisSession = True
    End Sub



    ''' <summary>
    ''' Attempts to restore AutoPilot configuration from the registry backup into My.Settings.
    ''' Fully silent — never shows UI, never throws.
    ''' </summary>
    ''' <returns>
    ''' <see langword="True"/> if a valid backup was found and restored; otherwise <see langword="False"/>.
    ''' </returns>
    Private Function TryRestoreAutoPilotSettingsFromRegistry() As Boolean
        Try
            Dim encoded As String = ReadAutoPilotRegistryValue()
            If String.IsNullOrWhiteSpace(encoded) Then Return False

            Dim decoded As String = DecodeString(encoded, SK)
            If String.IsNullOrWhiteSpace(decoded) OrElse
               decoded.StartsWith("Error:", StringComparison.OrdinalIgnoreCase) Then
                Return False
            End If

            Dim json As JObject = JObject.Parse(decoded)

            Dim backupType As String = GetJsonString(json, "T")
            If Not "AutoPilot".Equals(backupType, StringComparison.OrdinalIgnoreCase) Then
                Return False
            End If

            Dim filterRules As String = GetJsonString(json, "AP_FilterRules")
            Dim monitoredMailbox As String = GetJsonString(json, "AP_MonitoredMailbox")
            Dim whitelistedSenders As String = GetJsonString(json, "AP_WhitelistedSenders")

            My.Settings.AP_FilterRules = filterRules
            My.Settings.AP_WhitelistedSenders = whitelistedSenders
            My.Settings.AP_SubjectTriggerWord = GetJsonString(json, "AP_SubjectTriggerWord")
            My.Settings.AP_CooldownSeconds = GetJsonInteger(json, "AP_CooldownSeconds", My.Settings.AP_CooldownSeconds)
            My.Settings.AP_MaxRepliesPerSession = GetJsonInteger(json, "AP_MaxRepliesPerSession", My.Settings.AP_MaxRepliesPerSession)
            My.Settings.AP_MaxAttachmentMB = GetJsonInteger(json, "AP_MaxAttachmentMB", My.Settings.AP_MaxAttachmentMB)
            My.Settings.AP_FooterText = GetJsonString(json, "AP_FooterText")
            My.Settings.AP_RequireApproval = GetJsonBoolean(json, "AP_RequireApproval", My.Settings.AP_RequireApproval)
            My.Settings.AP_MonitoredMailbox = monitoredMailbox
            My.Settings.AP_SelectedModelKey = GetJsonString(json, "AP_SelectedModelKey")
            My.Settings.AP_UseSecondApi = GetJsonBoolean(json, "AP_UseSecondApi", My.Settings.AP_UseSecondApi)
            My.Settings.AP_ReprocessLookbackHours = GetJsonInteger(json, "AP_ReprocessLookbackHours", My.Settings.AP_ReprocessLookbackHours)
            My.Settings.AP_AutoDeleteAfterHours = GetJsonInteger(json, "AP_AutoDeleteAfterHours", My.Settings.AP_AutoDeleteAfterHours)
            My.Settings.AP_EnableWebGrounding = GetJsonBoolean(json, "AP_EnableWebGrounding", My.Settings.AP_EnableWebGrounding)
            My.Settings.AP_EnableVoicemailProcessing = GetJsonBoolean(json, "AP_EnableVoicemailProcessing", My.Settings.AP_EnableVoicemailProcessing)
            My.Settings.AP_VoicemailSenderAddress = GetJsonString(json, "AP_VoicemailSenderAddress")
            My.Settings.AP_VoicemailCallerIdMapPath = GetJsonString(json, "AP_VoicemailCallerIdMapPath")
            My.Settings.AP_EnableScheduler = GetJsonBoolean(json, "AP_EnableScheduler", My.Settings.AP_EnableScheduler)
            My.Settings.AP_EnableUserMemory = GetJsonBoolean(json, "AP_EnableUserMemory", My.Settings.AP_EnableUserMemory)
            My.Settings.AP_EnableUserFiles = GetJsonBoolean(json, "AP_EnableUserFiles", My.Settings.AP_EnableUserFiles)
            My.Settings.AP_EnablePrivacyProtection = GetJsonBoolean(json, "AP_EnablePrivacyProtection", My.Settings.AP_EnablePrivacyProtection)
            My.Settings.AP_SelectedExternalToolNames = GetJsonString(json, "AP_SelectedExternalToolNames")
            My.Settings.AP_DashboardWindowX = GetJsonInteger(json, "AP_DashboardWindowX", My.Settings.AP_DashboardWindowX)
            My.Settings.AP_DashboardWindowY = GetJsonInteger(json, "AP_DashboardWindowY", My.Settings.AP_DashboardWindowY)
            My.Settings.AP_DashboardWindowW = GetJsonInteger(json, "AP_DashboardWindowW", My.Settings.AP_DashboardWindowW)
            My.Settings.AP_DashboardWindowH = GetJsonInteger(json, "AP_DashboardWindowH", My.Settings.AP_DashboardWindowH)
            My.Settings.AP_SenderToolPolicyPath = GetJsonString(json, "AP_SenderToolPolicyPath")
            My.Settings.AP_ThreadRetentionDays = GetJsonInteger(json, "AP_ThreadRetentionDays", My.Settings.AP_ThreadRetentionDays)

            My.Settings.Save()

            Debug.WriteLine("[AutoPilot] Settings restored from registry backup.")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"[AutoPilot] Failed to restore registry backup: {ex.Message}")
            Return False
        End Try
    End Function

#End Region

#Region "Registry Low-Level Helpers"

    ''' <summary>
    ''' Writes a string to the registry AutoPilot backup location (default value).
    ''' Uses direct registry access to avoid UI side effects.
    ''' </summary>
    Private Shared Sub WriteAutoPilotRegistryValue(value As String)
        Dim fullPath As String = RegistryAutoPilotFullPath
        Dim hiveName As String = fullPath.Split("\"c)(0)
        Dim subKeyPath As String = fullPath.Substring(hiveName.Length + 1)

        Using subKey As RegistryKey = Registry.CurrentUser.CreateSubKey(subKeyPath, True)
            If subKey IsNot Nothing Then
                subKey.SetValue("", If(value, ""), RegistryValueKind.String)
            End If
        End Using
    End Sub

    ''' <summary>
    ''' Reads the string from the registry AutoPilot backup location (default value).
    ''' Returns an empty string if the key or value does not exist.
    ''' </summary>
    Private Shared Function ReadAutoPilotRegistryValue() As String
        Dim fullPath As String = RegistryAutoPilotFullPath
        Dim hiveName As String = fullPath.Split("\"c)(0)
        Dim subKeyPath As String = fullPath.Substring(hiveName.Length + 1)

        Using subKey As RegistryKey = Registry.CurrentUser.OpenSubKey(subKeyPath)
            If subKey Is Nothing Then Return ""
            Dim val As Object = subKey.GetValue("", Nothing)
            Return If(val?.ToString(), "")
        End Using
    End Function

#End Region

#Region "JSON Helpers"

    Private Shared Function GetJsonString(json As JObject, propertyName As String, Optional defaultValue As String = "") As String
        Dim token As JToken = json(propertyName)
        If token Is Nothing OrElse token.Type = JTokenType.Null Then Return defaultValue
        Return If(token.ToString(), defaultValue)
    End Function

    Private Shared Function GetJsonInteger(json As JObject, propertyName As String, defaultValue As Integer) As Integer
        Dim token As JToken = json(propertyName)
        If token Is Nothing OrElse token.Type = JTokenType.Null Then Return defaultValue

        Dim parsed As Integer
        If Integer.TryParse(token.ToString(), parsed) Then
            Return parsed
        End If

        Return defaultValue
    End Function

    Private Shared Function GetJsonBoolean(json As JObject, propertyName As String, defaultValue As Boolean) As Boolean
        Dim token As JToken = json(propertyName)
        If token Is Nothing OrElse token.Type = JTokenType.Null Then Return defaultValue

        Dim parsed As Boolean
        If Boolean.TryParse(token.ToString(), parsed) Then
            Return parsed
        End If

        Dim numericValue As Integer
        If Integer.TryParse(token.ToString(), numericValue) Then
            Return (numericValue <> 0)
        End If

        Return defaultValue
    End Function

#End Region

End Class
