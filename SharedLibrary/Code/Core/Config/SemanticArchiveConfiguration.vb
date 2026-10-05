' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' Canonical SemanticArchive configuration names and personal processing controls.

' =============================================================================
' File: SemanticArchiveConfiguration.vb
' Purpose:
'   Canonical Semantic Archive INI names, shared configuration lifecycle and personal
'   maintenance controls.
'
' Architecture / Function:
'   Keeps global archive configuration distinct from per-user controls and validates
'   catalog activation before publishing configuration changes.
' =============================================================================

Option Strict On
Option Explicit On
Option Infer On

Namespace SharedLibrary
    Public NotInheritable Class SemanticArchiveConfiguration
        Public Const CatalogPathKey As System.String = "SemanticArchiveCatalogPathLocal"
        Public Const LibraryPathKey As System.String = "SemanticArchiveCatalogLibraryPath"
        Public Const BackgroundEnabledSetting As System.String = "SemanticArchiveBackgroundIndexingEnabled"
        Public Const PermissionEnabledSetting As System.String = "SemanticArchivePermissionMaintenanceEnabled"
        Public Const BackgroundWindowSetting As System.String = "SemanticArchiveBackgroundIndexingWindow"
        Public Const PermissionWindowSetting As System.String = "SemanticArchivePermissionMaintenanceWindow"
        Public Const LegacyBackgroundEnabledKey As System.String = "SemanticArchiveBackgroundIndexing"
        Public Const SettingsVersionSetting As System.String = "SemanticArchiveSettingsVersion"
        Public Const CurrentSettingsVersion As System.Int32 = 1
        Friend Shared ReadOnly UserSettingsGate As New System.Object()

        Private Shared ReadOnly ConfiguredStates As New System.Runtime.CompilerServices.ConditionalWeakTable(Of SharedContext.ISharedContext, ConfiguredState)()

        ''' <summary>The INI commit succeeded, but the live catalog location could not be activated reliably.</summary>
        Public NotInheritable Class CatalogActivationException
            Inherits System.Configuration.ConfigurationErrorsException

            Public ReadOnly Property SavedIniPath As System.String

            Public Sub New(savedIniPath As System.String, innerException As System.Exception)
                MyBase.New("The personal INI was saved at '" & savedIniPath & "', but activating the catalog location failed. The saved change has not been rolled back. " &
                    If(innerException Is Nothing, System.String.Empty, innerException.Message), innerException)
                Me.SavedIniPath = savedIniPath
            End Sub
        End Class

        Public NotInheritable Class ControlsSnapshot
            Public ReadOnly Property BackgroundEnabled As System.Boolean
            Public ReadOnly Property BackgroundWindow As System.String
            Public ReadOnly Property PermissionEnabled As System.Boolean
            Public ReadOnly Property PermissionWindow As System.String

            Public Sub New(backgroundEnabled As System.Boolean, backgroundWindow As System.String,
                           permissionEnabled As System.Boolean, permissionWindow As System.String)
                Me.BackgroundEnabled = backgroundEnabled
                Me.BackgroundWindow = If(backgroundWindow, SharedMethods.DEFAULT_SEMANTICARCHIVE_BACKGROUND_INDEXING_WINDOW)
                Me.PermissionEnabled = permissionEnabled
                Me.PermissionWindow = If(permissionWindow, SharedMethods.DEFAULT_SEMANTICARCHIVE_PERMISSION_MAINTENANCE_WINDOW)
            End Sub
        End Class

        Private NotInheritable Class ConfiguredState
            Public Property Controls As ControlsSnapshot = DefaultControls()
            Public ReadOnly PendingPersonalValues As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.Ordinal)
        End Class

        Private Sub New()
        End Sub

        ''' <summary>Saves the personal catalog location through the normal INI path and activates it only after commit.</summary>
        Public Shared Function SavePersonalCatalogPath(context As SharedContext.ISharedContext, path As System.String) As System.String
            If context Is Nothing Then Throw New System.ArgumentNullException(NameOf(context))
            If System.String.IsNullOrWhiteSpace(path) Then Throw New System.ArgumentException("Choose a personal catalog directory.", NameOf(path))
            Dim configuredPath As System.String = path.Trim()
            Dim canonicalPath As System.String = SemanticArchivePathGuard.RequireWindowsCompatiblePath(configuredPath)
            Dim edits As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase) From {
                {CatalogPathKey, configuredPath}}
            Dim savedIniPath As System.String = ConfigWizardEngine.WriteLocalIniValues(context, edits)
            Try
                Dim reloadedContext As SharedContext.ISharedContext = context
                SharedMethods.InitializeConfig(reloadedContext, False, True)
                If Not context.INIloaded OrElse context.GPTSetupError Then
                    Throw New System.Configuration.ConfigurationErrorsException("The normal configuration loader did not complete successfully.")
                End If
                Dim effectivePath As System.String = SemanticArchivePathGuard.RequireWindowsCompatiblePath(context.INI_SemanticArchiveCatalogPathLocal)
                If Not System.String.Equals(canonicalPath, effectivePath, System.StringComparison.OrdinalIgnoreCase) Then
                    Throw New System.Configuration.ConfigurationErrorsException("The active configuration did not select the saved personal catalog location.")
                End If
            Catch failure As System.Exception
                Throw New CatalogActivationException(savedIniPath, failure)
            End Try
            Return savedIniPath
        End Function

        Public Shared Function DefaultControls() As ControlsSnapshot
            Return New ControlsSnapshot(SharedMethods.DEFAULT_SEMANTICARCHIVE_BACKGROUND_INDEXING_ENABLED,
                SharedMethods.DEFAULT_SEMANTICARCHIVE_BACKGROUND_INDEXING_WINDOW,
                SharedMethods.DEFAULT_SEMANTICARCHIVE_PERMISSION_MAINTENANCE_ENABLED,
                SharedMethods.DEFAULT_SEMANTICARCHIVE_PERMISSION_MAINTENANCE_WINDOW)
        End Function

        Public Shared Function GlobalKeys() As System.String()
            Return New System.String() {CatalogPathKey, LibraryPathKey, BackgroundEnabledSetting, BackgroundWindowSetting, PermissionEnabledSetting, PermissionWindowSetting}
        End Function

        Public Shared Function CanonicalizeIniKey(name As System.String) As System.String
            If System.String.Equals(name, LegacyBackgroundEnabledKey, System.StringComparison.OrdinalIgnoreCase) Then Return BackgroundEnabledSetting
            For Each canonical As System.String In GlobalKeys()
                If System.String.Equals(name, canonical, System.StringComparison.OrdinalIgnoreCase) Then Return canonical
            Next
            Return name
        End Function

        Public Shared Function IsGlobalKey(name As System.String) As System.Boolean
            Dim canonical As System.String = CanonicalizeIniKey(name)
            Return canonical = CatalogPathKey OrElse canonical = LibraryPathKey OrElse IsOperationalKey(canonical)
        End Function

        Public Shared Function IsOperationalKey(name As System.String) As System.Boolean
            Dim canonical As System.String = CanonicalizeIniKey(name)
            Return canonical = BackgroundEnabledSetting OrElse canonical = BackgroundWindowSetting OrElse
                canonical = PermissionEnabledSetting OrElse canonical = PermissionWindowSetting
        End Function

        ''' <summary>Normalizes only known SA keys. Canonical values, including blank/False, win over aliases in either order.</summary>
        Public Shared Function NormalizeIniValues(values As System.Collections.Generic.IDictionary(Of System.String, System.String)) As System.Collections.Generic.Dictionary(Of System.String, System.String)
            Dim normalized As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)
            If values Is Nothing Then Return normalized
            For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, System.String) In values
                normalized(CanonicalizeIniKey(pair.Key)) = If(pair.Value, System.String.Empty)
            Next
            For Each canonical As System.String In GlobalKeys()
                Dim value As System.String = Nothing
                If TryReadKey(values, canonical, value) Then normalized(canonical) = If(value, System.String.Empty)
            Next
            For Each booleanKey As System.String In New System.String() {BackgroundEnabledSetting, PermissionEnabledSetting}
                If normalized.ContainsKey(booleanKey) Then normalized(booleanKey) = ParseIniBooleanValue(normalized(booleanKey), booleanKey).ToString()
            Next
            Return normalized
        End Function

        Public Shared Function ReadConfiguredControlValues(values As System.Collections.Generic.IDictionary(Of System.String, System.String)) As ControlsSnapshot
            Dim normalized As System.Collections.Generic.Dictionary(Of System.String, System.String) = NormalizeIniValues(values)
            Return New ControlsSnapshot(ReadBoolean(normalized, BackgroundEnabledSetting, SharedMethods.DEFAULT_SEMANTICARCHIVE_BACKGROUND_INDEXING_ENABLED),
                ReadWindow(normalized, BackgroundWindowSetting, SharedMethods.DEFAULT_SEMANTICARCHIVE_BACKGROUND_INDEXING_WINDOW),
                ReadBoolean(normalized, PermissionEnabledSetting, SharedMethods.DEFAULT_SEMANTICARCHIVE_PERMISSION_MAINTENANCE_ENABLED),
                ReadWindow(normalized, PermissionWindowSetting, SharedMethods.DEFAULT_SEMANTICARCHIVE_PERMISSION_MAINTENANCE_WINDOW))
        End Function

        Private Shared Function ReadBoolean(values As System.Collections.Generic.IDictionary(Of System.String, System.String), name As System.String, fallback As System.Boolean) As System.Boolean
            Dim value As System.String = Nothing
            If Not TryReadKey(values, name, value) Then Return fallback
            Return ParseIniBooleanValue(value, name)
        End Function

        Private Shared Function ParseIniBooleanValue(value As System.String, name As System.String) As System.Boolean
            ' Preserve the established English/German INI synonyms; outputs always use canonical True/False.
            Select Case If(value, System.String.Empty).Trim().ToLowerInvariant()
                Case "true", "yes", "ja", "wahr" : Return True
                Case "false", "no", "nein", "falsch" : Return False
                Case Else : Throw New System.Configuration.ConfigurationErrorsException(name & " must be True or False (yes/no and ja/nein are also accepted).")
            End Select
        End Function

        Private Shared Function ReadWindow(values As System.Collections.Generic.IDictionary(Of System.String, System.String), name As System.String, fallback As System.String) As System.String
            Dim value As System.String = Nothing
            If Not TryReadKey(values, name, value) Then Return fallback
            value = If(value, System.String.Empty).Trim()
            If Not BackgroundProcessingWindow.IsValid(value) Then Throw New System.Configuration.ConfigurationErrorsException(name & " has an invalid processing window.")
            Return value
        End Function

        ''' <summary>Called only when a configuration is loaded. Timer reads use this immutable INI fallback, never edited live controls.</summary>
        Public Shared Sub LoadInto(context As SharedContext.ISharedContext, values As System.Collections.Generic.IDictionary(Of System.String, System.String))
            If context Is Nothing Then Throw New System.ArgumentNullException(NameOf(context))
            Dim configured As ControlsSnapshot = ReadConfiguredControlValues(values)
            SyncLock UserSettingsGate
                Dim state As ConfiguredState = ConfiguredStates.GetValue(context, Function(unused As SharedContext.ISharedContext) New ConfiguredState())
                state.Controls = configured
                state.PendingPersonalValues.Clear()
                Dim oldValue As System.String = Nothing
                Dim newValue As System.String = Nothing
                If values IsNot Nothing AndAlso (TryReadKey(values, "SemanticArchiveCatalogPath", oldValue) OrElse TryReadKey(values, "SemanticIndexArchivePath", oldValue)) AndAlso Not TryReadKey(values, CatalogPathKey, newValue) Then
                    Throw New System.Configuration.ConfigurationErrorsException("Rename the old SemanticArchiveCatalogPath/SemanticIndexArchivePath setting to SemanticArchiveCatalogPathLocal. Its value remains the personal catalog DIRECTORY, not a JSON filename.")
                End If
                context.INI_SemanticArchiveCatalogPathLocal = ReadCatalogPath(values)
                Dim libraryValue As System.String = SharedMethods.DEFAULT_SEMANTICARCHIVE_CATALOG_LIBRARY_PATH
                TryReadKey(values, LibraryPathKey, libraryValue)
                context.INI_SemanticArchiveCatalogLibraryPath = If(libraryValue, SharedMethods.DEFAULT_SEMANTICARCHIVE_CATALOG_LIBRARY_PATH).Trim()
                ApplyControls(context, ReadEffectiveControls(context))
            End SyncLock
        End Sub

        Public Shared Function ReadConfiguredControls(context As SharedContext.ISharedContext) As ControlsSnapshot
            If context Is Nothing Then Return DefaultControls()
            SyncLock UserSettingsGate
                Dim state As ConfiguredState = Nothing
                If ConfiguredStates.TryGetValue(context, state) Then Return state.Controls
                Return DefaultControls()
            End SyncLock
        End Function

        Public Shared Function MergeConfiguredControls(previous As ControlsSnapshot, values As System.Collections.Generic.IDictionary(Of System.String, System.String)) As ControlsSnapshot
            If previous Is Nothing Then Throw New System.ArgumentNullException(NameOf(previous))
            Dim merged As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase) From {
                {BackgroundEnabledSetting, previous.BackgroundEnabled.ToString()}, {BackgroundWindowSetting, previous.BackgroundWindow},
                {PermissionEnabledSetting, previous.PermissionEnabled.ToString()}, {PermissionWindowSetting, previous.PermissionWindow}}
            For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, System.String) In NormalizeIniValues(values)
                If IsOperationalKey(pair.Key) Then merged(pair.Key) = pair.Value
            Next
            Return ReadConfiguredControlValues(merged)
        End Function

        Public Shared Sub SetConfiguredControls(context As SharedContext.ISharedContext, values As System.Collections.Generic.IDictionary(Of System.String, System.String))
            If context Is Nothing Then Throw New System.ArgumentNullException(NameOf(context))
            SyncLock UserSettingsGate
                ConfiguredStates.GetValue(context, Function(unused As SharedContext.ISharedContext) New ConfiguredState()).Controls = MergeConfiguredControls(ReadConfiguredControls(context), values)
            End SyncLock
        End Sub

        Public Shared Function ReadEffectiveControls(context As SharedContext.ISharedContext) As ControlsSnapshot
            SyncLock UserSettingsGate
                ' A fresh provider instance cannot expose unrelated, unsaved My.Settings edits.
                Dim persisted As New Global.SharedLibrary.My.MySettings()
                Return ResolveUserControls(persisted, ReadConfiguredControls(context))
            End SyncLock
        End Function

        ''' <summary>Per-key explicit user values override the configured INI fallback; absent provider defaults do not.</summary>
        Public Shared Function ResolveUserControls(settings As System.Configuration.ApplicationSettingsBase, configured As ControlsSnapshot,
            Optional persistedSettingNames As System.Collections.Generic.IEnumerable(Of System.String) = Nothing) As ControlsSnapshot
            If settings Is Nothing Then Throw New System.ArgumentNullException(NameOf(settings))
            If configured Is Nothing Then Throw New System.ArgumentNullException(NameOf(configured))
            Dim present As New System.Collections.Generic.HashSet(Of System.String)(If(persistedSettingNames, GetPersistedSettingNames(settings)), System.StringComparer.Ordinal)
            Dim version As System.Int32 = System.Convert.ToInt32(settings(SettingsVersionSetting), System.Globalization.CultureInfo.InvariantCulture)
            ' Capture presence BEFORE migration: migration-created values are not new user choices.
            Dim hasBackground As System.Boolean = HasExplicitSetting(settings, present, BackgroundEnabledSetting) OrElse
                (version < CurrentSettingsVersion AndAlso HasExplicitSetting(settings, present, "EnableSABackgroundIndexing"))
            Dim hasPermissions As System.Boolean = HasExplicitSetting(settings, present, PermissionEnabledSetting) OrElse
                (version < CurrentSettingsVersion AndAlso HasExplicitSetting(settings, present, "EnableSAPermissionMaintenance"))
            Dim hasBackgroundWindow As System.Boolean = HasExplicitSetting(settings, present, BackgroundWindowSetting)
            Dim hasPermissionWindow As System.Boolean = HasExplicitSetting(settings, present, PermissionWindowSetting)
            MigrateUserSettings(settings, present)
            Dim result As New ControlsSnapshot(
                If(hasBackground, System.Convert.ToBoolean(settings(BackgroundEnabledSetting), System.Globalization.CultureInfo.InvariantCulture), configured.BackgroundEnabled),
                If(hasBackgroundWindow, System.Convert.ToString(settings(BackgroundWindowSetting), System.Globalization.CultureInfo.InvariantCulture), configured.BackgroundWindow),
                If(hasPermissions, System.Convert.ToBoolean(settings(PermissionEnabledSetting), System.Globalization.CultureInfo.InvariantCulture), configured.PermissionEnabled),
                If(hasPermissionWindow, System.Convert.ToString(settings(PermissionWindowSetting), System.Globalization.CultureInfo.InvariantCulture), configured.PermissionWindow))
            If Not BackgroundProcessingWindow.IsValid(result.BackgroundWindow) OrElse Not BackgroundProcessingWindow.IsValid(result.PermissionWindow) Then
                Throw New System.Configuration.ConfigurationErrorsException("A personal Semantic Archive processing window is invalid.")
            End If
            Return result
        End Function

        Public Shared Function CreateUserSettingsBackup(settings As System.Configuration.ApplicationSettingsBase,
            Optional persistedSettingNames As System.Collections.Generic.IEnumerable(Of System.String) = Nothing) As Newtonsoft.Json.Linq.JObject
            If settings Is Nothing Then Throw New System.ArgumentNullException(NameOf(settings))
            Dim present As New System.Collections.Generic.HashSet(Of System.String)(If(persistedSettingNames, GetPersistedSettingNames(settings)), System.StringComparer.Ordinal)
            Dim version As System.Int32 = System.Convert.ToInt32(settings(SettingsVersionSetting), System.Globalization.CultureInfo.InvariantCulture)
            Dim explicitNames As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            For Each key As System.String In GlobalKeys()
                If IsOperationalKey(key) AndAlso HasExplicitSetting(settings, present, key) Then explicitNames.Add(key)
            Next
            If version < CurrentSettingsVersion Then
                If HasExplicitSetting(settings, present, "EnableSABackgroundIndexing") Then explicitNames.Add(BackgroundEnabledSetting)
                If HasExplicitSetting(settings, present, "EnableSAPermissionMaintenance") Then explicitNames.Add(PermissionEnabledSetting)
            End If
            Dim preserveVersion As System.Boolean = HasExplicitSetting(settings, present, SettingsVersionSetting)
            MigrateUserSettings(settings, present)
            Dim result As New Newtonsoft.Json.Linq.JObject()
            For Each key As System.String In explicitNames
                If key = BackgroundEnabledSetting OrElse key = PermissionEnabledSetting Then
                    result(key) = New Newtonsoft.Json.Linq.JValue(System.Convert.ToBoolean(settings(key), System.Globalization.CultureInfo.InvariantCulture))
                Else
                    result(key) = New Newtonsoft.Json.Linq.JValue(System.Convert.ToString(settings(key), System.Globalization.CultureInfo.InvariantCulture))
                End If
            Next
            If explicitNames.Count > 0 OrElse preserveVersion Then result(SettingsVersionSetting) = New Newtonsoft.Json.Linq.JValue(CurrentSettingsVersion)
            Return result
        End Function

        ''' <summary>Applies only fields actually present in the backup. Missing fields remain inherited INI defaults. Never saves.</summary>
        Public Shared Sub ApplyUserSettingsBackup(settings As System.Configuration.ApplicationSettingsBase, payload As Newtonsoft.Json.Linq.JObject)
            If settings Is Nothing Then Throw New System.ArgumentNullException(NameOf(settings))
            If payload Is Nothing Then Throw New System.ArgumentNullException(NameOf(payload))
            Dim normalized As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)
            Dim versionToken As Newtonsoft.Json.Linq.JToken = payload(SettingsVersionSetting)
            Dim completed As System.Boolean = versionToken IsNot Nothing AndAlso versionToken.Type <> Newtonsoft.Json.Linq.JTokenType.Null AndAlso
                System.Convert.ToInt32(versionToken.ToString(), System.Globalization.CultureInfo.InvariantCulture) >= CurrentSettingsVersion
            For Each key As System.String In GlobalKeys()
                If Not IsOperationalKey(key) Then Continue For
                Dim token As Newtonsoft.Json.Linq.JToken = payload(key)
                If token Is Nothing AndAlso Not completed Then
                    If key = BackgroundEnabledSetting Then token = payload("EnableSABackgroundIndexing")
                    If key = PermissionEnabledSetting Then token = payload("EnableSAPermissionMaintenance")
                End If
                If token Is Nothing Then Continue For
                If key = BackgroundEnabledSetting OrElse key = PermissionEnabledSetting Then
                    If token.Type <> Newtonsoft.Json.Linq.JTokenType.Boolean AndAlso token.Type <> Newtonsoft.Json.Linq.JTokenType.String Then
                        Throw New System.Configuration.ConfigurationErrorsException("Invalid saved Boolean: " & key)
                    End If
                ElseIf token.Type <> Newtonsoft.Json.Linq.JTokenType.String AndAlso token.Type <> Newtonsoft.Json.Linq.JTokenType.Null Then
                    Throw New System.Configuration.ConfigurationErrorsException("Invalid saved processing window: " & key)
                End If
                normalized(key) = If(token.Type = Newtonsoft.Json.Linq.JTokenType.Null, System.String.Empty, token.ToString())
            Next
            normalized = NormalizeIniValues(normalized)
            ReadConfiguredControlValues(normalized)
            For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, System.String) In normalized
                If pair.Key = BackgroundEnabledSetting OrElse pair.Key = PermissionEnabledSetting Then
                    settings(pair.Key) = System.Boolean.Parse(pair.Value)
                Else
                    settings(pair.Key) = pair.Value
                End If
            Next
            If normalized.Count > 0 OrElse completed Then settings(SettingsVersionSetting) = CurrentSettingsVersion
        End Sub

        Private Shared Function HasExplicitSetting(settings As System.Configuration.ApplicationSettingsBase, present As System.Collections.Generic.ISet(Of System.String), name As System.String) As System.Boolean
            Dim materialized As System.Object = settings(name)
            Dim value As System.Configuration.SettingsPropertyValue = settings.PropertyValues(name)
            Return present.Contains(name) OrElse (value IsNot Nothing AndAlso value.IsDirty)
        End Function

        Public Shared Sub ApplyControls(context As SharedContext.ISharedContext, values As ControlsSnapshot)
            If context Is Nothing Then Return
            context.INI_SemanticArchiveBackgroundIndexing = values.BackgroundEnabled
            context.INI_SemanticArchiveBackgroundIndexingWindow = values.BackgroundWindow
            context.INI_SemanticArchivePermissionMaintenanceEnabled = values.PermissionEnabled
            context.INI_SemanticArchivePermissionMaintenanceWindow = values.PermissionWindow
        End Sub

        ''' <summary>Stages only actually changed personal controls. Saving unrelated settings never pins inherited defaults.</summary>
        Public Shared Sub StagePersonalControl(context As SharedContext.ISharedContext, name As System.String, value As System.String)
            If context Is Nothing Then Throw New System.ArgumentNullException(NameOf(context))
            Dim canonical As System.String = CanonicalizeIniKey(name)
            If Not IsOperationalKey(canonical) Then Throw New System.ArgumentException("Not a Semantic Archive personal control.", NameOf(name))
            Dim singleValue As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase) From {{canonical, value}}
            Dim parsed As ControlsSnapshot = ReadConfiguredControlValues(singleValue)
            Dim current As System.String
            Dim normalized As System.String
            SyncLock UserSettingsGate
                Select Case canonical
                    Case BackgroundEnabledSetting
                        current = context.INI_SemanticArchiveBackgroundIndexing.ToString()
                        normalized = parsed.BackgroundEnabled.ToString()
                    Case BackgroundWindowSetting
                        current = context.INI_SemanticArchiveBackgroundIndexingWindow
                        normalized = parsed.BackgroundWindow
                    Case PermissionEnabledSetting
                        current = context.INI_SemanticArchivePermissionMaintenanceEnabled.ToString()
                        normalized = parsed.PermissionEnabled.ToString()
                    Case Else
                        current = context.INI_SemanticArchivePermissionMaintenanceWindow
                        normalized = parsed.PermissionWindow
                End Select
                Dim state As ConfiguredState = ConfiguredStates.GetValue(context, Function(unused As SharedContext.ISharedContext) New ConfiguredState())
                ' A timer can reload the effective value between edits; a newer explicit edit must replace an older pending value.
                If System.String.Equals(If(current, System.String.Empty), normalized, System.StringComparison.Ordinal) AndAlso Not state.PendingPersonalValues.ContainsKey(canonical) Then Return
                state.PendingPersonalValues(canonical) = normalized
                Select Case canonical
                    Case BackgroundEnabledSetting : context.INI_SemanticArchiveBackgroundIndexing = parsed.BackgroundEnabled
                    Case BackgroundWindowSetting : context.INI_SemanticArchiveBackgroundIndexingWindow = parsed.BackgroundWindow
                    Case PermissionEnabledSetting : context.INI_SemanticArchivePermissionMaintenanceEnabled = parsed.PermissionEnabled
                    Case Else : context.INI_SemanticArchivePermissionMaintenanceWindow = parsed.PermissionWindow
                End Select
            End SyncLock
        End Sub

        ''' <summary>Closes a settings draft. Clears pending writes before any reload, and never saves settings.</summary>
        Public Shared Sub DiscardPendingUserSettings(context As SharedContext.ISharedContext, Optional restoreEffectiveControls As System.Boolean = True)
            If context Is Nothing Then Return
            SyncLock UserSettingsGate
                Dim state As ConfiguredState = Nothing
                If Not ConfiguredStates.TryGetValue(context, state) OrElse state.PendingPersonalValues.Count = 0 Then Return
                state.PendingPersonalValues.Clear()
                If restoreEffectiveControls Then ApplyControls(context, ReadEffectiveControls(context))
            End SyncLock
        End Sub

        Public Shared Sub SavePendingUserSettings(context As SharedContext.ISharedContext)
            If context Is Nothing Then Return
            SyncLock UserSettingsGate
                Dim state As ConfiguredState = Nothing
                If Not ConfiguredStates.TryGetValue(context, state) OrElse state.PendingPersonalValues.Count = 0 Then Return
                SaveUserControls(context, state.PendingPersonalValues)
                state.PendingPersonalValues.Clear()
            End SyncLock
        End Sub

        Public Shared Sub SaveUserControls(context As SharedContext.ISharedContext, values As System.Collections.Generic.IDictionary(Of System.String, System.String))
            If values Is Nothing Then Throw New System.ArgumentNullException(NameOf(values))
            Dim normalized As System.Collections.Generic.Dictionary(Of System.String, System.String) = NormalizeIniValues(values)
            For Each key As System.String In normalized.Keys
                If Not IsOperationalKey(key) Then Throw New System.ArgumentException("Only personal Semantic Archive controls can be saved here.", NameOf(values))
            Next
            ReadConfiguredControlValues(normalized) ' Validate every value before any persistence.
            SyncLock UserSettingsGate
                Dim persisted As New Global.SharedLibrary.My.MySettings()
                MigrateUserSettings(persisted)
                For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, System.String) In normalized
                    If pair.Key = BackgroundEnabledSetting OrElse pair.Key = PermissionEnabledSetting Then
                        persisted(pair.Key) = System.Boolean.Parse(pair.Value)
                    Else
                        persisted(pair.Key) = If(pair.Value, System.String.Empty).Trim()
                    End If
                Next
                persisted.Save()
                For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, System.String) In normalized
                    My.Settings(pair.Key) = persisted(pair.Key)
                Next
                My.Settings.SemanticArchiveSettingsVersion = CurrentSettingsVersion
                If context IsNot Nothing Then ApplyControls(context, ReadEffectiveControls(context))
            End SyncLock
            SharedMethods.BackupSharedUserSettingsToRegistry()
        End Sub

        Public Shared Function ReadCatalogPath(values As System.Collections.Generic.IDictionary(Of System.String, System.String)) As System.String
            If values Is Nothing Then Return SharedMethods.DEFAULT_SEMANTICARCHIVE_CATALOG_PATH_LOCAL
            Dim value As System.String = Nothing
            If TryReadKey(values, CatalogPathKey, value) Then Return If(value, System.String.Empty)
            Return SharedMethods.DEFAULT_SEMANTICARCHIVE_CATALOG_PATH_LOCAL
        End Function

        Private Shared Function TryReadKey(values As System.Collections.Generic.IDictionary(Of System.String, System.String), name As System.String, ByRef value As System.String) As System.Boolean
            For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, System.String) In values
                If System.String.Equals(pair.Key, name, System.StringComparison.OrdinalIgnoreCase) Then
                    value = pair.Value
                    Return True
                End If
            Next
            Return False
        End Function

        ''' <summary>Copies legacy values only when the canonical user setting is absent and unmodified.
        ''' Never saves: callers persist the canonical values and version together only on an explicit save.</summary>
        Public Shared Function MigrateUserSettings(settings As System.Configuration.ApplicationSettingsBase,
                                                   Optional persistedSettingNames As System.Collections.Generic.IEnumerable(Of System.String) = Nothing) As System.Boolean
            If settings Is Nothing Then Throw New System.ArgumentNullException(NameOf(settings))
            Dim version As System.Int32 = System.Convert.ToInt32(settings(SettingsVersionSetting), System.Globalization.CultureInfo.InvariantCulture)
            If version >= CurrentSettingsVersion Then Return False
            Dim present As New System.Collections.Generic.HashSet(Of System.String)(
                If(persistedSettingNames, GetPersistedSettingNames(settings)), System.StringComparer.Ordinal)
            CopyLegacyWhenAbsent(settings, present, BackgroundEnabledSetting, "EnableSABackgroundIndexing")
            CopyLegacyWhenAbsent(settings, present, PermissionEnabledSetting, "EnableSAPermissionMaintenance")
            settings(SettingsVersionSetting) = CurrentSettingsVersion
            Return True
        End Function

        Public Shared Function HasUserSettings(settings As System.Configuration.ApplicationSettingsBase,
                                              Optional persistedSettingNames As System.Collections.Generic.IEnumerable(Of System.String) = Nothing) As System.Boolean
            If settings Is Nothing Then Throw New System.ArgumentNullException(NameOf(settings))
            If System.Convert.ToInt32(settings(SettingsVersionSetting), System.Globalization.CultureInfo.InvariantCulture) >= CurrentSettingsVersion Then Return True
            Dim present As New System.Collections.Generic.HashSet(Of System.String)(
                If(persistedSettingNames, GetPersistedSettingNames(settings)), System.StringComparer.Ordinal)
            For Each name As System.String In New System.String() {BackgroundEnabledSetting, PermissionEnabledSetting,
                    "EnableSABackgroundIndexing", "EnableSAPermissionMaintenance", "SemanticArchiveBackgroundIndexingWindow", "SemanticArchivePermissionMaintenanceWindow"}
                If present.Contains(name) Then Return True
                Dim value As System.Object = settings(name)
                If settings.PropertyValues(name) IsNot Nothing AndAlso settings.PropertyValues(name).IsDirty Then Return True
            Next
            Return False
        End Function

        Private Shared Sub CopyLegacyWhenAbsent(settings As System.Configuration.ApplicationSettingsBase,
                                               present As System.Collections.Generic.ISet(Of System.String),
                                               canonical As System.String, legacy As System.String)
            Dim current As System.Object = settings(canonical) ' Materialize provider values before checking IsDirty.
            Dim propertyValue As System.Configuration.SettingsPropertyValue = settings.PropertyValues(canonical)
            If present.Contains(canonical) OrElse (propertyValue IsNot Nothing AndAlso propertyValue.IsDirty) Then Return
            ' An inherited legacy default is not an explicit preference and must not become a dirty override.
            If Not HasExplicitSetting(settings, present, legacy) Then Return
            settings(canonical) = settings(legacy)
        End Sub

        Private Shared Function GetPersistedSettingNames(settings As System.Configuration.ApplicationSettingsBase) As System.Collections.Generic.IEnumerable(Of System.String)
            Dim groupName As System.String = System.Convert.ToString(settings.Context("GroupName"), System.Globalization.CultureInfo.InvariantCulture)
            If System.String.IsNullOrWhiteSpace(groupName) Then groupName = settings.GetType().FullName
            Dim settingsKey As System.String = settings.SettingsKey
            If Not System.String.IsNullOrEmpty(settingsKey) Then groupName &= "." & settingsKey
            Dim paths As New System.Collections.Generic.List(Of System.String)()
            For Each level As System.Configuration.ConfigurationUserLevel In New System.Configuration.ConfigurationUserLevel() {
                    System.Configuration.ConfigurationUserLevel.PerUserRoaming,
                    System.Configuration.ConfigurationUserLevel.PerUserRoamingAndLocal}
                Dim configuration As System.Configuration.Configuration = System.Configuration.ConfigurationManager.OpenExeConfiguration(level)
                If Not paths.Contains(configuration.FilePath) Then paths.Add(configuration.FilePath)
            Next
            Return ReadPersistedSettingNames(paths, groupName)
        End Function

        ''' <summary>Reads raw user-config files, never inherited application defaults. Unknown access or malformed XML fails visibly.</summary>
        Public Shared Function ReadPersistedSettingNames(configurationPaths As System.Collections.Generic.IEnumerable(Of System.String),
                                                        groupName As System.String) As System.Collections.Generic.HashSet(Of System.String)
            If configurationPaths Is Nothing Then Throw New System.ArgumentNullException(NameOf(configurationPaths))
            If System.String.IsNullOrWhiteSpace(groupName) Then Throw New System.ArgumentException("A settings group is required.", NameOf(groupName))
            Dim result As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            Dim encodedGroup As System.String = System.Xml.XmlConvert.EncodeLocalName(groupName)
            For Each path As System.String In configurationPaths
                Dim input As System.IO.FileStream
                Try
                    input = New System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite Or System.IO.FileShare.Delete)
                Catch ex As System.IO.FileNotFoundException
                    Continue For
                Catch ex As System.IO.DirectoryNotFoundException
                    Continue For
                End Try
                Using input
                    Dim readerSettings As New System.Xml.XmlReaderSettings With {
                        .DtdProcessing = System.Xml.DtdProcessing.Prohibit, .XmlResolver = Nothing,
                        .MaxCharactersInDocument = 16777216, .IgnoreComments = True}
                    Using reader As System.Xml.XmlReader = System.Xml.XmlReader.Create(input, readerSettings)
                        Dim document As New System.Xml.XmlDocument() With {.XmlResolver = Nothing}
                        document.Load(reader)
                        Dim configuration As System.Xml.XmlElement = document.DocumentElement
                        If configuration Is Nothing OrElse configuration.Name <> "configuration" Then Throw New System.Configuration.ConfigurationErrorsException("User settings XML has an invalid root.")
                        For Each child As System.Xml.XmlNode In configuration.ChildNodes
                            If child.Name <> "userSettings" Then Continue For
                            For Each group As System.Xml.XmlNode In child.ChildNodes
                                If group.Name <> encodedGroup Then Continue For
                                If group.Attributes IsNot Nothing AndAlso group.Attributes("configSource") IsNot Nothing Then Throw New System.Configuration.ConfigurationErrorsException("External user-settings sections require an explicit migration before changing canonical preferences.")
                                For Each setting As System.Xml.XmlNode In group.ChildNodes
                                    If setting.Name <> "setting" OrElse setting.Attributes Is Nothing Then Continue For
                                    Dim name As System.Xml.XmlAttribute = setting.Attributes("name")
                                    If name IsNot Nothing Then result.Add(name.Value)
                                Next
                            Next
                        Next
                    End Using
                End Using
            Next
            Return result
        End Function
    End Class
End Namespace
