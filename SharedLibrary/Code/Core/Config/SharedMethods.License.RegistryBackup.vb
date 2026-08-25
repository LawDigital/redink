' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.License.RegistryBackup.vb
' Purpose: Registry-based backup/restore of Pro license credentials to survive
'          user.config (My.Settings) deletions caused by VSTO updates, profile
'          resets, or roaming profile sync issues.
'
' Architecture / How it works:
'  - On every successful Pro license save to My.Settings (via SaveProLicenseToSettings),
'    the same credentials are written as an XOR-encoded JSON blob to the registry
'    under RegPath_Base & RegPath_License (default value).
'  - On startup, if My.Settings has no stored Pro license, the registry backup
'    is read and restored silently before prompting the user.
'  - On license clear/deactivation, the registry backup is also cleared.
'  - All operations are fully silent (no UI) and fail-safe (exceptions are swallowed
'    with logging only).
'
' Notes:
'  - Private licenses are NOT backed up (by design — they are trivial to recreate).
'  - The existing WriteToRegistry helper is NOT used because it shows a MessageBox
'    on success. Direct Microsoft.Win32.Registry access is used instead.
'  - Encoding uses the existing CodeString/DecodeString (XOR + Base64) with SK as key.
' =============================================================================

Option Strict On
Option Explicit On

Imports Microsoft.Win32
Imports Newtonsoft.Json.Linq

Namespace SharedLibrary

    Partial Public Class SharedMethods

#Region "Registry License Backup"

        ''' <summary>
        ''' Full registry path for the license backup value.
        ''' Resolves to: HKEY_CURRENT_USER\Software\Red Ink\License
        ''' </summary>
        Private Shared ReadOnly RegistryLicenseFullPath As String = RegPath_Base & RegPath_License

        ''' <summary>
        ''' Saves a Pro license backup to the registry as an encoded JSON blob.
        ''' Called from <see cref="SaveProLicenseToSettings"/> after writing to My.Settings.
        ''' Fully silent — never shows UI, never throws.
        ''' </summary>
        Friend Shared Sub BackupProLicenseToRegistry(productId As String,
                                                      licenseKey As String,
                                                      userId As String,
                                                      productName As String,
                                                      apiConfirmed As Boolean)
            Try
                Dim json As New JObject From {
                    {"T", "Pro"},
                    {"P", If(productId, "")},
                    {"K", If(licenseKey, "")},
                    {"U", If(userId, "")},
                    {"N", If(productName, "")},
                    {"A", apiConfirmed},
                    {"D", Date.UtcNow.ToString("o")}
                }

                Dim encoded As String = CodeString(json.ToString(Newtonsoft.Json.Formatting.None), SK)
                WriteLicenseRegistryValue(encoded)

                LogLicenseEvent("REGISTRY_BACKUP", "Pro license backup saved to registry.")

            Catch ex As Exception
                LogLicenseEvent("REGISTRY_BACKUP_ERROR",
                                $"Failed to save Pro license backup to registry: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' Attempts to restore Pro license credentials from the registry backup.
        ''' Called from <see cref="LicenseOK"/> when no stored Pro license is found in My.Settings.
        ''' Fully silent — never shows UI, never throws.
        ''' </summary>
        ''' <returns><see langword="True"/> if a backup was found and successfully restored to My.Settings;
        ''' otherwise <see langword="False"/>.</returns>
        Friend Shared Function TryRestoreProLicenseFromRegistry() As Boolean
            Try
                Dim encoded As String = ReadLicenseRegistryValue()
                If String.IsNullOrWhiteSpace(encoded) Then Return False

                Dim decoded As String = DecodeString(encoded, SK)
                If String.IsNullOrWhiteSpace(decoded) OrElse
                   decoded.StartsWith("Error:", StringComparison.OrdinalIgnoreCase) Then
                    Return False
                End If

                Dim json As JObject = JObject.Parse(decoded)

                ' Must be a Pro license backup
                Dim licType As String = json.Value(Of String)("T")
                If Not "Pro".Equals(licType, StringComparison.OrdinalIgnoreCase) Then
                    Return False
                End If

                Dim productId As String = json.Value(Of String)("P")
                Dim licenseKey As String = json.Value(Of String)("K")
                Dim userId As String = json.Value(Of String)("U")
                Dim productName As String = json.Value(Of String)("N")
                Dim apiConfirmed As Boolean = json.Value(Of Boolean)("A")

                ' Validate minimum required fields
                If String.IsNullOrWhiteSpace(productId) OrElse
                   String.IsNullOrWhiteSpace(licenseKey) OrElse
                   String.IsNullOrWhiteSpace(userId) Then
                    LogLicenseEvent("REGISTRY_RESTORE", "Registry backup found but incomplete. Skipping.")
                    Return False
                End If

                ' Restore to My.Settings via the existing save method
                SaveProLicenseToSettings(productId, licenseKey, userId,
                                         If(productName, ""), apiConfirmed)

                _restoredFromRegistryBackup = True

                LogLicenseEvent("REGISTRY_RESTORE",
                                $"Pro license restored from registry backup (Product: {If(productName, "unknown")}).",
                                alwaysLog:=True)
                Return True

            Catch ex As Exception
                ' Corrupted, tampered, or missing registry data — silently ignore
                LogLicenseEvent("REGISTRY_RESTORE_ERROR",
                                $"Failed to restore license from registry: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Clears the license backup from the registry.
        ''' Called from <see cref="ClearStoredLicense"/>.
        ''' Fully silent — never shows UI, never throws.
        ''' </summary>
        Friend Shared Sub ClearLicenseRegistryBackup()
            Try
                WriteLicenseRegistryValue("")
                LogLicenseEvent("REGISTRY_BACKUP", "License registry backup cleared.")
            Catch ex As Exception
                LogLicenseEvent("REGISTRY_BACKUP_ERROR",
                                $"Failed to clear license registry backup: {ex.Message}")
            End Try
        End Sub

#End Region

#Region "Registry Shared User Settings Backup"

        ''' <summary>
        ''' Full registry path for the shared user-settings backup value.
        ''' Resolves to: HKEY_CURRENT_USER\Software\Red Ink\UserSettings
        ''' </summary>
        Private Shared ReadOnly RegistrySharedUserSettingsFullPath As String = RegPath_Base & "UserSettings"

        ''' <summary>
        ''' Saves the shared My.Settings payload to the registry as an encoded JSON blob.
        ''' Fully silent — never shows UI, never throws.
        ''' </summary>
        Friend Shared Sub BackupSharedUserSettingsToRegistry()
            Try
                Dim json As New JObject From {
                    {"T", "UserSettings"},
                    {"V", 1},
                    {"DefaultPrefix", GetMySettingStringValue("DefaultPrefix")},
                    {"ReplaceText2Override", GetMySettingStringValue("ReplaceText2Override")},
                    {"RestrictedModelAccessCode", GetMySettingStringValue("RestrictedModelAccessCode")},
                    {"MarkupMethodWordOverride", GetMySettingStringValue("MarkupMethodWordOverride")},
                    {"MarkupMethodOutlookOverride", GetMySettingStringValue("MarkupMethodOutlookOverride")},
                    {"MarkupAuthor", GetMySettingStringValue("MarkupAuthor")},
                    {"SimpleMenuOverride", GetMySettingBooleanValue("SimpleMenuOverride", False)},
                    {"SimpleMenuOverrideIsSet", GetMySettingBooleanValue("SimpleMenuOverrideIsSet", False)},
                    {"EnableKBBackgroundIndexing", GetMySettingBooleanValue("EnableKBBackgroundIndexing", False)},
                    {"KnowledgeStoreBackgroundIndexingWindow", GetMySettingStringValue("KnowledgeStoreBackgroundIndexingWindow")},
                    {"FormulaInstruction", GetMySettingStringValue("FormulaInstruction")},
                    {"D", Date.UtcNow.ToString("o")}
                }

                Dim encoded As String = CodeString(json.ToString(Newtonsoft.Json.Formatting.None), SK)
                WriteSharedUserSettingsRegistryValue(encoded)
            Catch ex As Exception
                LogLicenseEvent("REGISTRY_BACKUP_ERROR",
                                $"Failed to save shared user-settings backup to registry: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' Restores the shared My.Settings payload from the registry backup when those settings are absent.
        ''' Fully silent — never shows UI, never throws.
        ''' </summary>
        Friend Shared Sub TryRestoreSharedUserSettingsFromRegistry()
            If HasStoredSharedUserSettings() Then
                Return
            End If

            Try
                Dim encoded As String = ReadSharedUserSettingsRegistryValue()
                If String.IsNullOrWhiteSpace(encoded) Then
                    Return
                End If

                Dim decoded As String = DecodeString(encoded, SK)
                If String.IsNullOrWhiteSpace(decoded) OrElse
                   decoded.StartsWith("Error:", StringComparison.OrdinalIgnoreCase) Then
                    Return
                End If

                Dim json As JObject = JObject.Parse(decoded)
                Dim backupType As String = GetJsonStringValue(json, "T")
                If Not "UserSettings".Equals(backupType, StringComparison.OrdinalIgnoreCase) Then
                    Return
                End If

                SetMySettingValue("DefaultPrefix", GetJsonStringValue(json, "DefaultPrefix"))
                SetMySettingValue("ReplaceText2Override", GetJsonStringValue(json, "ReplaceText2Override"))
                SetMySettingValue("RestrictedModelAccessCode", GetJsonStringValue(json, "RestrictedModelAccessCode"))
                SetMySettingValue("MarkupMethodWordOverride", GetJsonStringValue(json, "MarkupMethodWordOverride"))
                SetMySettingValue("MarkupMethodOutlookOverride", GetJsonStringValue(json, "MarkupMethodOutlookOverride"))
                SetMySettingValue("MarkupAuthor", GetJsonStringValue(json, "MarkupAuthor"))
                SetMySettingValue("SimpleMenuOverride", GetJsonBooleanValue(json, "SimpleMenuOverride", False))
                SetMySettingValue("SimpleMenuOverrideIsSet", GetJsonBooleanValue(json, "SimpleMenuOverrideIsSet", False))
                SetMySettingValue("EnableKBBackgroundIndexing", GetJsonBooleanValue(json, "EnableKBBackgroundIndexing", False))
                SetMySettingValue("KnowledgeStoreBackgroundIndexingWindow", GetJsonStringValue(json, "KnowledgeStoreBackgroundIndexingWindow"))
                SetMySettingValue("FormulaInstruction", GetJsonStringValue(json, "FormulaInstruction"))
                My.Settings.Save()
            Catch ex As Exception
                LogLicenseEvent("REGISTRY_RESTORE_ERROR",
                                $"Failed to restore shared user-settings from registry: {ex.Message}")
            End Try
        End Sub

        Private Shared Function HasStoredSharedUserSettings() As Boolean
            If Not String.IsNullOrWhiteSpace(GetMySettingStringValue("DefaultPrefix")) Then Return True
            If Not String.IsNullOrWhiteSpace(GetMySettingStringValue("ReplaceText2Override")) Then Return True
            If Not String.IsNullOrWhiteSpace(GetMySettingStringValue("RestrictedModelAccessCode")) Then Return True
            If Not String.IsNullOrWhiteSpace(GetMySettingStringValue("MarkupMethodWordOverride")) Then Return True
            If Not String.IsNullOrWhiteSpace(GetMySettingStringValue("MarkupMethodOutlookOverride")) Then Return True
            If Not String.IsNullOrWhiteSpace(GetMySettingStringValue("MarkupAuthor")) Then Return True
            If GetMySettingBooleanValue("SimpleMenuOverrideIsSet", False) Then Return True
            If GetMySettingBooleanValue("EnableKBBackgroundIndexing", False) Then Return True
            If Not String.IsNullOrWhiteSpace(GetMySettingStringValue("KnowledgeStoreBackgroundIndexingWindow")) Then Return True
            If Not String.IsNullOrWhiteSpace(GetMySettingStringValue("FormulaInstruction")) Then Return True

            Return False
        End Function

#End Region

#Region "Registry Low-Level Helpers"

        ''' <summary>
        ''' Writes a string to the registry license backup location (default value).
        ''' Uses direct registry access to avoid the MessageBox in <see cref="WriteToRegistry"/>.
        ''' </summary>
        Private Shared Sub WriteLicenseRegistryValue(value As String)
            ' RegPath_Base includes "HKEY_CURRENT_USER\" prefix — strip the hive name to get the subkey path
            Dim fullPath As String = RegistryLicenseFullPath
            Dim hiveName As String = fullPath.Split("\"c)(0)
            Dim subKeyPath As String = fullPath.Substring(hiveName.Length + 1)

            Using subKey As RegistryKey = Registry.CurrentUser.CreateSubKey(subKeyPath, True)
                If subKey IsNot Nothing Then
                    subKey.SetValue("", If(value, ""), RegistryValueKind.String)
                End If
            End Using
        End Sub

        ''' <summary>
        ''' Reads the string from the registry license backup location (default value).
        ''' Returns an empty string if the key or value does not exist.
        ''' </summary>
        Private Shared Function ReadLicenseRegistryValue() As String
            Dim fullPath As String = RegistryLicenseFullPath
            Dim hiveName As String = fullPath.Split("\"c)(0)
            Dim subKeyPath As String = fullPath.Substring(hiveName.Length + 1)

            Using subKey As RegistryKey = Registry.CurrentUser.OpenSubKey(subKeyPath)
                If subKey Is Nothing Then Return ""
                Dim val As Object = subKey.GetValue("", Nothing)
                Return If(val?.ToString(), "")
            End Using
        End Function

        Private Shared Sub WriteSharedUserSettingsRegistryValue(value As String)
            Dim fullPath As String = RegistrySharedUserSettingsFullPath
            Dim hiveName As String = fullPath.Split("\"c)(0)
            Dim subKeyPath As String = fullPath.Substring(hiveName.Length + 1)

            Using subKey As RegistryKey = Registry.CurrentUser.CreateSubKey(subKeyPath, True)
                If subKey IsNot Nothing Then
                    subKey.SetValue("", If(value, ""), RegistryValueKind.String)
                End If
            End Using
        End Sub

        Private Shared Function ReadSharedUserSettingsRegistryValue() As String
            Dim fullPath As String = RegistrySharedUserSettingsFullPath
            Dim hiveName As String = fullPath.Split("\"c)(0)
            Dim subKeyPath As String = fullPath.Substring(hiveName.Length + 1)

            Using subKey As RegistryKey = Registry.CurrentUser.OpenSubKey(subKeyPath)
                If subKey Is Nothing Then Return ""
                Dim val As Object = subKey.GetValue("", Nothing)
                Return If(val?.ToString(), "")
            End Using
        End Function

        Private Shared Function GetMySettingStringValue(settingName As String) As String
            Try
                Dim rawValue As Object = My.Settings.Item(settingName)
                If rawValue Is Nothing Then
                    Return ""
                End If

                Return If(rawValue.ToString(), "")
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function GetMySettingBooleanValue(settingName As String, defaultValue As Boolean) As Boolean
            Try
                Dim rawValue As Object = My.Settings.Item(settingName)
                If rawValue Is Nothing Then
                    Return defaultValue
                End If

                If TypeOf rawValue Is Boolean Then
                    Return CBool(rawValue)
                End If

                Dim parsedValue As Boolean
                If Boolean.TryParse(rawValue.ToString().Trim(), parsedValue) Then
                    Return parsedValue
                End If

                Return defaultValue
            Catch
                Return defaultValue
            End Try
        End Function

        Private Shared Sub SetMySettingValue(settingName As String, value As Object)
            Try
                My.Settings.Item(settingName) = value
            Catch
            End Try
        End Sub

        Private Shared Function GetJsonStringValue(json As JObject, propertyName As String, Optional defaultValue As String = "") As String
            Dim token = json(propertyName)
            If token Is Nothing OrElse token.Type = JTokenType.Null Then
                Return defaultValue
            End If

            Return If(token.ToString(), defaultValue)
        End Function

        Private Shared Function GetJsonBooleanValue(json As JObject, propertyName As String, defaultValue As Boolean) As Boolean
            Dim token = json(propertyName)
            If token Is Nothing OrElse token.Type = JTokenType.Null Then
                Return defaultValue
            End If

            Dim parsedValue As Boolean
            If Boolean.TryParse(token.ToString(), parsedValue) Then
                Return parsedValue
            End If

            Dim numericValue As Integer
            If Integer.TryParse(token.ToString(), numericValue) Then
                Return numericValue <> 0
            End If

            Return defaultValue
        End Function

#End Region

    End Class

End Namespace
