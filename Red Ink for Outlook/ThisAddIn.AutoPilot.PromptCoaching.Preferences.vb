' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.AutoPilot.PromptCoaching.Preferences.vb
' PURPOSE
'   Stores per-sender recommendation preferences and enforces deterministic eligibility and opt-out.
'
' KEY ARCHITECTURE
'   • Identity: Normalized per-sender preferences are separate from deletable user memory.
'   • Eligibility: Explicit opt-out and sender-bound controls do not trust model recipients.
' =============================================================================

'
' Inky prompt recommendations: explicit per-user preferences and deterministic eligibility.
' The preference store is separate from deletable memory/home files so deleting those
' files cannot silently undo an opt-out. No model-provided recipient is accepted.

Option Explicit On
Option Strict On

Partial Public Class ThisAddIn

    Private Shared ReadOnly _apPromptCoachingPreferenceGate As New System.Object()
    Private Const AP_PromptCoachingPreferenceMaxBytes As System.Int32 = 65536

    Private Shared Function NormalizeAutoPilotPromptCoachingEmail(value As System.String) As System.String
        If System.String.IsNullOrWhiteSpace(value) Then Return System.String.Empty
        Dim candidate As System.String = value.Trim()
        If candidate.Length > 320 OrElse candidate.IndexOfAny(New System.Char() {
                ";"c, ","c, "<"c, ">"c, "*"c, "?"c, System.Convert.ToChar(13), System.Convert.ToChar(10)}) >= 0 Then Return System.String.Empty
        For Each ch As System.Char In candidate
            If System.Char.IsControl(ch) OrElse System.Char.IsWhiteSpace(ch) Then Return System.String.Empty
        Next
        Try
            Dim address As New System.Net.Mail.MailAddress(candidate)
            If address.DisplayName.Length > 0 OrElse
               Not System.String.Equals(address.Address, candidate, System.StringComparison.OrdinalIgnoreCase) OrElse
               address.User.Length = 0 OrElse address.Host.Length = 0 OrElse candidate.IndexOf("@"c) <= 0 Then Return System.String.Empty
            Return address.Address.ToLowerInvariant()
        Catch ex As System.FormatException
            Return System.String.Empty
        End Try
    End Function

    ''' <summary>Accepts exact addresses separated by newlines, semicolons or commas; never wildcard rules.</summary>
    Private Shared Function ParseAutoPilotPromptCoachingIncludedUsers(value As System.String) As System.Collections.Generic.List(Of System.String)
        Dim result As New System.Collections.Generic.List(Of System.String)()
        Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
        For Each item As System.String In If(value, System.String.Empty).Split(
                New System.Char() {System.Convert.ToChar(13), System.Convert.ToChar(10), ";"c, ","c}, System.StringSplitOptions.RemoveEmptyEntries)
            If System.String.IsNullOrWhiteSpace(item) Then Continue For
            Dim address As System.String = NormalizeAutoPilotPromptCoachingEmail(item)
            If address.Length = 0 Then Throw New System.ArgumentException("Inky recommendation users must be exact e-mail addresses; display names, domain rules and wildcards are not supported.")
            If seen.Add(address) Then result.Add(address)
        Next
        result.Sort(System.StringComparer.OrdinalIgnoreCase)
        Return result
    End Function

    Private Function GetAutoPilotPromptCoachingPreferencesDir() As System.String
        Return System.IO.Path.Combine(System.IO.Path.GetDirectoryName(GetUserStorageRootDir()), "autopilot_prompt_coaching", "preferences")
    End Function

    Private Function GetAutoPilotPromptCoachingPreferencePath(email As System.String) As System.String
        Dim canonical As System.String = NormalizeAutoPilotPromptCoachingEmail(email)
        If canonical.Length = 0 Then Throw New System.ArgumentException("A valid single sender e-mail address is required.", NameOf(email))
        Return System.IO.Path.Combine(GetAutoPilotPromptCoachingPreferencesDir(), GetCanonicalUserFolderName(canonical) & ".json")
    End Function

    Private Shared Sub RejectAutoPilotPromptCoachingPreferenceLinks(path As System.String)
        ' Operator-selected INI parent is trusted; descendants owned by this feature must
        ' not redirect preference reads/writes through filesystem junctions or symlinks.
        If (System.IO.Directory.Exists(path) OrElse System.IO.File.Exists(path)) AndAlso
           (System.IO.File.GetAttributes(path) And System.IO.FileAttributes.ReparsePoint) <> 0 Then
            Throw New System.IO.InvalidDataException("The Inky recommendation preference path must not be a link or junction.")
        End If
    End Sub

    Private Sub ValidateAutoPilotPromptCoachingPreferencePath(path As System.String)
        Dim preferencesDir As System.String = GetAutoPilotPromptCoachingPreferencesDir()
        If Not IsPathContained(path, preferencesDir) Then Throw New System.IO.InvalidDataException("Invalid Inky recommendation preference path.")
        RejectAutoPilotPromptCoachingPreferenceLinks(System.IO.Path.GetDirectoryName(preferencesDir))
        RejectAutoPilotPromptCoachingPreferenceLinks(preferencesDir)
        RejectAutoPilotPromptCoachingPreferenceLinks(path)
    End Sub

    Private Function ReadAutoPilotPromptCoachingPreference(email As System.String) As Newtonsoft.Json.Linq.JObject
        Dim path As System.String = GetAutoPilotPromptCoachingPreferencePath(email)
        ValidateAutoPilotPromptCoachingPreferencePath(path)
        Dim content As System.String
        Try
            ' File.Exists also returns False on access errors. Opening explicitly keeps
            ' an unreadable opt-out from being mistaken for a missing preference.
            Using stream As New System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                If stream.Length > AP_PromptCoachingPreferenceMaxBytes Then Throw New System.IO.InvalidDataException("The Inky recommendation preference file is too large.")
                Using reader As New System.IO.StreamReader(stream, System.Text.Encoding.UTF8, True)
                    content = reader.ReadToEnd()
                End Using
            End Using
        Catch ex As System.IO.FileNotFoundException
            Return Nothing
        Catch ex As System.IO.DirectoryNotFoundException
            Return Nothing
        End Try
        Dim settings As New Newtonsoft.Json.Linq.JsonLoadSettings() With {
            .DuplicatePropertyNameHandling = Newtonsoft.Json.Linq.DuplicatePropertyNameHandling.Error
        }
        Dim data As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(content, settings)
        If data("schemaVersion") Is Nothing OrElse data("schemaVersion").Type <> Newtonsoft.Json.Linq.JTokenType.Integer OrElse
           data.Value(Of System.Int32)("schemaVersion") <> 1 OrElse
           data("ownerEmail") Is Nothing OrElse data("ownerEmail").Type <> Newtonsoft.Json.Linq.JTokenType.String OrElse
           Not System.String.Equals(data.Value(Of System.String)("ownerEmail"), email, System.StringComparison.Ordinal) Then
            Throw New System.IO.InvalidDataException("The Inky recommendation preference identity or schema is invalid.")
        End If
        Dim preference As System.String = data.Value(Of System.String)("preference")
        If preference <> "enabled" AndAlso preference <> "disabled" AndAlso preference <> "default" Then
            Throw New System.IO.InvalidDataException("The Inky recommendation preference value is invalid.")
        End If
        Dim source As System.String = data.Value(Of System.String)("source")
        If source <> "user" AndAlso source <> "admin" AndAlso source <> "observed" Then Throw New System.IO.InvalidDataException("The Inky recommendation preference source is invalid.")
        If source = "observed" AndAlso preference <> "default" Then Throw New System.IO.InvalidDataException("An observed user cannot have an implicit explicit preference.")
        Dim delivery As System.String = ReadAutoPilotPromptCoachingDeliveryPreference(data)
        If source = "observed" AndAlso delivery <> "default" Then Throw New System.IO.InvalidDataException("An observed user cannot have an explicit delivery preference.")
        If data("deliverySource") IsNot Nothing Then
            If data("deliverySource").Type <> Newtonsoft.Json.Linq.JTokenType.String Then Throw New System.IO.InvalidDataException("Invalid delivery preference source type.")
            Dim deliverySource As System.String = data.Value(Of System.String)("deliverySource")
            If deliverySource <> "user" AndAlso deliverySource <> "admin" AndAlso deliverySource <> "default" Then Throw New System.IO.InvalidDataException("Invalid delivery preference source.")
        End If
        Return data
    End Function

    Private Shared Function ReadAutoPilotPromptCoachingDeliveryPreference(data As Newtonsoft.Json.Linq.JObject) As System.String
        If data Is Nothing OrElse data("deliveryPreference") Is Nothing Then Return "default"
        Dim token As Newtonsoft.Json.Linq.JToken = data("deliveryPreference")
        If token.Type <> Newtonsoft.Json.Linq.JTokenType.String Then Throw New System.IO.InvalidDataException("Invalid recommendation delivery preference type.")
        Dim value As System.String = token.ToObject(Of System.String)()
        If value <> "default" AndAlso value <> "daily" AndAlso value <> "immediate" Then Throw New System.IO.InvalidDataException("Invalid recommendation delivery preference value.")
        Return value
    End Function

    ''' <summary>Fail-closed policy; explicit user/admin disable always wins over both list and defaults.</summary>
    Private Function GetAutoPilotPromptCoachingPreferenceStatus(email As System.String) As Newtonsoft.Json.Linq.JObject
        Dim canonical As System.String = NormalizeAutoPilotPromptCoachingEmail(email)
        Dim status As New Newtonsoft.Json.Linq.JObject From {
            {"preference", "default"}, {"effectiveEnabled", False},
            {"globalEnabled", _apConfig IsNot Nothing AndAlso _apConfig.PromptCoachingEnabled},
            {"deliveryMode", If(_apConfig IsNot Nothing AndAlso _apConfig.PromptCoachingImmediateDelivery, "immediate", "daily_20_local")},
            {"deliveryPreference", "default"}, {"deliverySource", "default"},
            {"systemDeliveryMode", If(_apConfig IsNot Nothing AndAlso _apConfig.PromptCoachingImmediateDelivery, "immediate", "daily_20_local")},
            {"source", "default"}, {"reason", "invalid_sender"}
        }
        If canonical.Length = 0 Then Return status
        SyncLock _apPromptCoachingPreferenceGate
            Try
                Dim data As Newtonsoft.Json.Linq.JObject = ReadAutoPilotPromptCoachingPreference(canonical)
                Dim preference As System.String = If(data Is Nothing, "default", data.Value(Of System.String)("preference"))
                status("preference") = preference
                Dim delivery As System.String = ReadAutoPilotPromptCoachingDeliveryPreference(data)
                status("deliveryPreference") = delivery
                If delivery <> "default" Then status("deliveryMode") = If(delivery = "daily", "daily_20_local", "immediate")
                If data IsNot Nothing Then
                    status("source") = data.Value(Of System.String)("source")
                    status("deliverySource") = If(data.Value(Of System.String)("deliverySource"), "default")
                End If
                If _apConfig Is Nothing OrElse Not _apConfig.PromptCoachingEnabled Then
                    status("reason") = "globally_disabled"
                ElseIf preference = "disabled" Then
                    status("reason") = "explicitly_disabled"
                ElseIf preference = "enabled" Then
                    status("effectiveEnabled") = True
                    status("reason") = "explicitly_enabled"
                Else
                    Dim included As System.Collections.Generic.List(Of System.String) = ParseAutoPilotPromptCoachingIncludedUsers(_apConfig.PromptCoachingIncludedUsers)
                    If included.Contains(canonical) Then
                        status("effectiveEnabled") = True
                        status("reason") = "included_user"
                    ElseIf _apConfig.PromptCoachingDefaultForAllUsers Then
                        status("effectiveEnabled") = True
                        status("reason") = "default_all_users"
                    Else
                        status("reason") = "opt_in_required"
                    End If
                End If
            Catch ex As System.Exception
                status("preference") = "unavailable"
                status("effectiveEnabled") = False
                status("reason") = "preference_unavailable"
                status("deliveryPreference") = "unavailable"
                LogAutoPilotLogDiagnosticsStatus("coaching-preferences", "Preference could not be verified; recommendations withheld. " & ex.GetType().Name & ": " & ex.Message)
            End Try
        End SyncLock
        Return status
    End Function

    Private Function IsAutoPilotPromptCoachingEnabledForUser(email As System.String) As System.Boolean
        Return GetAutoPilotPromptCoachingPreferenceStatus(email).Value(Of System.Boolean)("effectiveEnabled")
    End Function

    Private Sub SetAutoPilotPromptCoachingPreference(email As System.String, preference As System.String, source As System.String,
                                                     Optional deliveryPreference As System.String = Nothing)
        Dim canonical As System.String = NormalizeAutoPilotPromptCoachingEmail(email)
        If canonical.Length = 0 Then Throw New System.ArgumentException("A valid single sender e-mail address is required.", NameOf(email))
        If preference IsNot Nothing AndAlso preference <> "enabled" AndAlso preference <> "disabled" AndAlso preference <> "default" Then Throw New System.ArgumentException("Unknown Inky recommendation preference.", NameOf(preference))
        If source <> "user" AndAlso source <> "admin" AndAlso source <> "observed" Then Throw New System.ArgumentException("Unknown Inky recommendation preference source.", NameOf(source))
        If deliveryPreference IsNot Nothing AndAlso deliveryPreference <> "default" AndAlso deliveryPreference <> "daily" AndAlso deliveryPreference <> "immediate" Then
            Throw New System.ArgumentException("Unknown Inky recommendation delivery preference.", NameOf(deliveryPreference))
        End If
        If preference Is Nothing AndAlso deliveryPreference Is Nothing Then Throw New System.ArgumentException("No recommendation preference change was supplied.")
        If source = "observed" AndAlso (preference <> "default" OrElse deliveryPreference IsNot Nothing) Then Throw New System.ArgumentException("An observed user can only have the default preference.", NameOf(preference))
        SyncLock _apPromptCoachingPreferenceGate
            Dim path As System.String = GetAutoPilotPromptCoachingPreferencePath(canonical)
            ValidateAutoPilotPromptCoachingPreferencePath(path)
            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path))
            Dim tempPath As System.String = path & ".tmp." & System.Guid.NewGuid().ToString("N")
            Try
                Dim data As Newtonsoft.Json.Linq.JObject = Nothing
                Try
                    data = ReadAutoPilotPromptCoachingPreference(canonical)
                Catch ex As System.IO.InvalidDataException When preference = "disabled" OrElse (preference IsNot Nothing AndAlso deliveryPreference IsNot Nothing)
                    ' An explicit opt-out or explicit replacement of BOTH settings may
                    ' repair malformed state. A timing-only change never resets an opt-out.
                    LogAutoPilotLogDiagnosticsStatus("coaching-preferences", "Malformed preference replaced by an explicit complete choice or opt-out; unknown delivery choice reset to the system default.")
                Catch ex As Newtonsoft.Json.JsonException When preference = "disabled" OrElse (preference IsNot Nothing AndAlso deliveryPreference IsNot Nothing)
                    LogAutoPilotLogDiagnosticsStatus("coaching-preferences", "Malformed preference JSON replaced by an explicit complete choice or opt-out.")
                End Try
                If data Is Nothing Then
                    data = New Newtonsoft.Json.Linq.JObject From {
                        {"schemaVersion", 1}, {"ownerEmail", canonical}, {"preference", "default"},
                        {"source", "observed"}, {"deliveryPreference", "default"}, {"deliverySource", "default"}
                    }
                End If
                Dim nowUtc As System.String = System.DateTime.UtcNow.ToString("o", System.Globalization.CultureInfo.InvariantCulture)
                If preference IsNot Nothing Then
                    data("preference") = preference
                    data("source") = source
                    data("updatedUtc") = nowUtc
                End If
                If deliveryPreference IsNot Nothing Then
                    data("deliveryPreference") = deliveryPreference
                    data("deliverySource") = source
                    data("deliveryUpdatedUtc") = nowUtc
                    If data.Value(Of System.String)("source") = "observed" Then data("source") = source
                End If
                Dim bytes As System.Byte() = New System.Text.UTF8Encoding(False).GetBytes(data.ToString(Newtonsoft.Json.Formatting.Indented))
                Using stream As New System.IO.FileStream(tempPath, System.IO.FileMode.CreateNew, System.IO.FileAccess.Write, System.IO.FileShare.None)
                    stream.Write(bytes, 0, bytes.Length)
                    stream.Flush(True)
                End Using
                If source = "observed" Then
                    ' Never replace a concurrent preference, including an unreadable one.
                    Try
                        System.IO.File.Move(tempPath, path)
                    Catch ex As System.IO.IOException
                        If Not System.IO.File.Exists(path) Then Throw
                    End Try
                ElseIf System.IO.File.Exists(path) Then
                    System.IO.File.Replace(tempPath, path, Nothing)
                Else
                    System.IO.File.Move(tempPath, path)
                End If
            Finally
                Try
                    If System.IO.File.Exists(tempPath) Then System.IO.File.Delete(tempPath)
                Catch ex As System.Exception
                    System.Diagnostics.Debug.WriteLine("Inky recommendation preference temp cleanup failed: " & ex.GetType().Name)
                End Try
            End Try
        End SyncLock
    End Sub

    Private Sub RememberAutoPilotPromptCoachingUser(email As System.String)
        Dim canonical As System.String = NormalizeAutoPilotPromptCoachingEmail(email)
        If canonical.Length = 0 Then Return
        SyncLock _apPromptCoachingPreferenceGate
            Try
                If ReadAutoPilotPromptCoachingPreference(canonical) Is Nothing Then
                    SetAutoPilotPromptCoachingPreference(canonical, "default", "observed")
                End If
            Catch ex As System.Exception
                ' A corrupt/unreadable file must retain its fail-closed meaning; observing
                ' another request is never permission to replace or reset a preference.
                LogAutoPilotLogDiagnosticsStatus("coaching-preferences", "Could not remember the recommendation participant; existing preference retained. " & ex.GetType().Name)
            End Try
        End SyncLock
    End Sub

    Private Function ExecuteAutoPilotManagePromptCoaching(args As System.Collections.Generic.Dictionary(Of System.String, System.Object)) As System.String
        If args Is Nothing Then Throw New System.ArgumentException("The action parameter is required.")
        For Each key As System.String In args.Keys
            If Not System.String.Equals(key, "action", System.StringComparison.Ordinal) Then
                Throw New System.ArgumentException("Only action is supported. Recommendation preferences always belong to the current sender.")
            End If
        Next
        Dim action As System.String = If(GetArgString(args, "action"), System.String.Empty).Trim().ToLowerInvariant()
        Dim info As AutoPilotMailInfo = _apCurrentMailInfo
        If info Is Nothing OrElse System.String.IsNullOrWhiteSpace(info.EntryID) OrElse
           info.EntryID.StartsWith("scheduler_", System.StringComparison.OrdinalIgnoreCase) Then
            Throw New System.InvalidOperationException("Inky recommendation preferences can only be managed from a direct user request, not by a scheduled task.")
        End If
        Dim sender As System.String = NormalizeAutoPilotPromptCoachingEmail(info.SenderEmail)
        If sender.Length = 0 Then Throw New System.InvalidOperationException("The current sender could not be verified.")
        Select Case action
            Case "enable"
                SetAutoPilotPromptCoachingPreference(sender, "enabled", "user")
            Case "disable"
                SetAutoPilotPromptCoachingPreference(sender, "disabled", "user")
            Case "reset"
                SetAutoPilotPromptCoachingPreference(sender, "default", "user", "default")
            Case "daily"
                SetAutoPilotPromptCoachingPreference(sender, "enabled", "user", "daily")
            Case "immediate"
                SetAutoPilotPromptCoachingPreference(sender, "enabled", "user", "immediate")
            Case "delivery_default"
                SetAutoPilotPromptCoachingPreference(sender, Nothing, "user", "default")
            Case "status"
                ' Read only.
            Case Else
                Throw New System.ArgumentException("Use action='status', 'enable', 'disable', 'reset', 'daily', 'immediate', or 'delivery_default'.")
        End Select
        Dim status As Newtonsoft.Json.Linq.JObject = GetAutoPilotPromptCoachingPreferenceStatus(sender)
        Dim preference As System.String = status.Value(Of System.String)("preference")
        If preference = "unavailable" Then
            status("message") = "Your Inky recommendation preference could not be read safely. Recommendations are currently withheld. You can set your preference again or ask your administrator."
        ElseIf Not status.Value(Of System.Boolean)("globalEnabled") Then
            status("message") = "Your preference is stored (" & preference & "). Inky recommendations are currently switched off by the administrator, so no recommendation e-mails will be sent."
        ElseIf status.Value(Of System.Boolean)("effectiveEnabled") Then
            status("message") = "Inky may send you useful, concrete suggestions for future requests " &
                If(status.Value(Of System.String)("deliveryMode") = "immediate", "after the analysis of a completed request.", "in a daily digest from 20:00 local time.") &
                " Only highly significant, request-specific improvements that pass a second quality review are mailed; minor style suggestions do not trigger a message. You can ask Inky to turn these recommendations off at any time."
        Else
            status("message") = "Inky recommendation e-mails are off for you. You can ask Inky to enable them at any time."
        End If
        Dim timingChoice As System.String = status.Value(Of System.String)("deliveryPreference")
        Dim timingDescription As System.String = If(timingChoice = "default", "follow the administrator's schedule",
            If(timingChoice = "daily", "once daily", If(timingChoice = "immediate", "after a qualifying request is analysed", "could not be verified")))
        status("message") = status.Value(Of System.String)("message") & " Your saved timing choice: " & timingDescription &
            ". ON/OFF preserve this choice; DAILY/IMMEDIATE enable tips and select the corresponding schedule; SYSTEM changes only the schedule back to the system default; DEFAULT resets both personal choices."
        Return status.ToString(Newtonsoft.Json.Formatting.None)
    End Function

    Private Function GetAutoPilotPromptCoachingDashboardEmail(folderName As System.String) As System.String
        If System.String.IsNullOrWhiteSpace(folderName) Then Return System.String.Empty
        Try
            Dim path As System.String = System.IO.Path.GetFullPath(System.IO.Path.Combine(GetUserStorageRootDir(), folderName))
            If Not IsPathContained(path, GetCanonicalUserStorageRootDir()) Then Return System.String.Empty
            RejectAutoPilotPromptCoachingPreferenceLinks(path)
            Dim marker As System.String = System.IO.Path.Combine(path, AP_UserOwnerIdentityFileName)
            If Not System.IO.File.Exists(marker) OrElse New System.IO.FileInfo(marker).Length > 1024 Then Return System.String.Empty
            RejectAutoPilotPromptCoachingPreferenceLinks(marker)
            Dim canonical As System.String = NormalizeAutoPilotPromptCoachingEmail(System.IO.File.ReadAllText(marker, System.Text.Encoding.UTF8))
            If canonical.Length = 0 OrElse Not System.String.Equals(path, System.IO.Path.GetFullPath(GetCanonicalUserDirPath(canonical)), System.StringComparison.OrdinalIgnoreCase) Then Return System.String.Empty
            Return canonical
        Catch ex As System.Exception
            Return System.String.Empty
        End Try
    End Function

    ''' <summary>Known identities for the admin grid, including opt-outs retained after memory/file deletion.</summary>
    Private Function ListAutoPilotPromptCoachingKnownUsers() As System.Collections.Generic.List(Of System.String)
        Dim users As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
        SyncLock _apPromptCoachingPreferenceGate
            Try
                Dim directory As System.String = GetAutoPilotPromptCoachingPreferencesDir()
                RejectAutoPilotPromptCoachingPreferenceLinks(System.IO.Path.GetDirectoryName(directory))
                RejectAutoPilotPromptCoachingPreferenceLinks(directory)
                If System.IO.Directory.Exists(directory) Then
                    For Each path As System.String In System.IO.Directory.GetFiles(directory, "*.json", System.IO.SearchOption.TopDirectoryOnly)
                        Try
                            ValidateAutoPilotPromptCoachingPreferencePath(path)
                            If New System.IO.FileInfo(path).Length > AP_PromptCoachingPreferenceMaxBytes Then Continue For
                            Dim data As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(System.IO.File.ReadAllText(path, System.Text.Encoding.UTF8))
                            Dim canonical As System.String = NormalizeAutoPilotPromptCoachingEmail(data.Value(Of System.String)("ownerEmail"))
                            If canonical.Length > 0 AndAlso System.String.Equals(path, GetAutoPilotPromptCoachingPreferencePath(canonical), System.StringComparison.OrdinalIgnoreCase) Then users.Add(canonical)
                        Catch ex As System.Exception
                            LogAutoPilotLogDiagnosticsStatus("coaching-preferences-list", "An unreadable recommendation preference was excluded from the admin identity list; recommendations remain fail-closed. " & ex.GetType().Name)
                        End Try
                    Next
                End If
                If _apConfig IsNot Nothing Then
                    For Each canonical As System.String In ParseAutoPilotPromptCoachingIncludedUsers(_apConfig.PromptCoachingIncludedUsers)
                        users.Add(canonical)
                    Next
                End If
            Catch ex As System.Exception
                LogAutoPilotLogDiagnosticsStatus("coaching-preferences-list", "Recommendation user list could not be fully read: " & ex.GetType().Name)
            End Try
        End SyncLock
        Dim result As New System.Collections.Generic.List(Of System.String)(users)
        result.Sort(System.StringComparer.OrdinalIgnoreCase)
        Return result
    End Function

End Class
