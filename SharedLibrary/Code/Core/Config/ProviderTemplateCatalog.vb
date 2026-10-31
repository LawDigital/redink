' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ProviderTemplateCatalog.vb
' PURPOSE
'   Loads a shared embedded provider baseline for configuration and installation wizards.
'
' KEY ARCHITECTURE
'   • Baseline: Loads an immutable snapshot from the embedded provider catalog.
'   • Update: Optional online refresh is explicitly user-confirmed and does not rewrite INI.
' =============================================================================

'
' Canonical embedded provider templates shared by the Installation Wizard and ConfigWizard.
' The embedded catalog is the offline baseline. Both wizards use the same optional,
' user-confirmed online update workflow; it never mutates the installed INI.

Option Strict On
Option Explicit On

Namespace SharedLibrary

    Public NotInheritable Class ProviderTemplateCatalogSnapshot
        Public ReadOnly Configurations As System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of AppConfigurationVariable))
        Public ReadOnly Notes As System.Collections.Generic.Dictionary(Of System.String, System.String)
        Public ReadOnly ProviderOrder As System.Collections.Generic.IReadOnlyList(Of System.String)

        Public Sub New(configurations As System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of AppConfigurationVariable)),
                       notes As System.Collections.Generic.Dictionary(Of System.String, System.String),
                       providerOrder As System.Collections.Generic.IList(Of System.String))
            Me.Configurations = configurations
            Me.Notes = notes
            Me.ProviderOrder = New System.Collections.Generic.List(Of System.String)(providerOrder).AsReadOnly()
        End Sub
    End Class

    Public NotInheritable Class ProviderTemplateCatalog
        Private Sub New()
        End Sub

        Public Shared Function LoadEmbedded() As ProviderTemplateCatalogSnapshot
            Dim assembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
            Dim resourceName As System.String = assembly.GetManifestResourceNames().FirstOrDefault(
                Function(name As System.String) name.EndsWith("redink-defaultconfig.ini", System.StringComparison.OrdinalIgnoreCase))
            If System.String.IsNullOrWhiteSpace(resourceName) Then
                Throw New System.Configuration.ConfigurationErrorsException("Embedded provider catalog 'redink-defaultconfig.ini' is missing.")
            End If
            Using resourceStream As System.IO.Stream = assembly.GetManifestResourceStream(resourceName)
                If resourceStream Is Nothing Then
                    Throw New System.Configuration.ConfigurationErrorsException("Embedded provider catalog cannot be read: " & resourceName)
                End If
                Using reader As New System.IO.StreamReader(resourceStream, System.Text.Encoding.UTF8, True)
                    Return Parse(reader.ReadToEnd())
                End Using
            End Using
        End Function

        ''' <summary>
        ''' Parses the pipe-delimited provider catalog. Field order and complete default values
        ''' (including embedded JSON, punctuation, and any additional '|' characters) survive.
        ''' Malformed FieldN records are rejected instead of silently losing configurations.
        ''' </summary>
        Public Shared Function Parse(content As System.String) As ProviderTemplateCatalogSnapshot
            If System.String.IsNullOrWhiteSpace(content) Then
                Throw New System.Configuration.ConfigurationErrorsException("Provider catalog is empty.")
            End If

            Dim configurations As New System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of AppConfigurationVariable))(
                System.StringComparer.OrdinalIgnoreCase)
            Dim notes As New System.Collections.Generic.Dictionary(Of System.String, System.String)(
                System.StringComparer.OrdinalIgnoreCase)
            Dim providerOrder As New System.Collections.Generic.List(Of System.String)()
            Dim sectionName As System.String = Nothing
            Dim fields As New System.Collections.Generic.SortedDictionary(Of System.Int32, System.String)()
            Dim sectionNote As System.String = System.String.Empty

            Dim flush As System.Action =
                Sub()
                    If System.String.IsNullOrWhiteSpace(sectionName) Then Return
                    If fields.Count = 0 Then
                        Throw New System.Configuration.ConfigurationErrorsException("Provider catalog section has no FieldN entries: " & sectionName)
                    End If
                    Dim variables As New System.Collections.Generic.List(Of AppConfigurationVariable)()
                    Dim knownKeys As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
                    For Each item As System.Collections.Generic.KeyValuePair(Of System.Int32, System.String) In fields
                        Dim parts As System.String() = item.Value.Split(New System.Char() {"|"c}, 5, System.StringSplitOptions.None)
                        If parts.Length <> 5 OrElse System.String.IsNullOrWhiteSpace(parts(1)) OrElse
                           Not parts(1).Trim().StartsWith("INI_", System.StringComparison.OrdinalIgnoreCase) Then
                            Throw New System.Configuration.ConfigurationErrorsException(
                                "Invalid provider catalog field " & sectionName & ".Field" & item.Key.ToString(System.Globalization.CultureInfo.InvariantCulture))
                        End If
                        Dim varName As System.String = parts(1).Trim()
                        If Not knownKeys.Add(varName) Then
                            Throw New System.Configuration.ConfigurationErrorsException("Duplicate provider configuration key in " & sectionName & ": " & varName)
                        End If
                        variables.Add(New AppConfigurationVariable With {
                            .DisplayName = parts(0).Trim(),
                            .VarName = varName,
                            .VarType = parts(2).Trim(),
                            .ValidationRule = parts(3).Trim(),
                            .DefaultValue = parts(4),
                            .CurrentValue = parts(4)
                        })
                    Next
                    configurations.Add(sectionName, variables)
                    providerOrder.Add(sectionName)
                    notes(sectionName) = sectionNote
                    fields.Clear()
                    sectionNote = System.String.Empty
                End Sub

            Dim lineNumber As System.Int32 = 0
            For Each raw As System.String In content.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split(New System.Char() {ChrW(10)}, System.StringSplitOptions.None)
                lineNumber += 1
                Dim line As System.String = raw.Trim()
                If line.Length = 0 OrElse line.StartsWith(";", System.StringComparison.Ordinal) OrElse
                   line.StartsWith("#", System.StringComparison.Ordinal) Then Continue For
                If line.StartsWith("[", System.StringComparison.Ordinal) AndAlso line.EndsWith("]", System.StringComparison.Ordinal) Then
                    flush()
                    sectionName = line.Substring(1, line.Length - 2).Trim()
                    If sectionName.Length = 0 OrElse configurations.ContainsKey(sectionName) Then
                        Throw New System.Configuration.ConfigurationErrorsException("Invalid or duplicate provider catalog section at line " & lineNumber.ToString())
                    End If
                    Continue For
                End If
                If System.String.IsNullOrWhiteSpace(sectionName) Then
                    Throw New System.Configuration.ConfigurationErrorsException("Provider catalog entry outside section at line " & lineNumber.ToString())
                End If
                Dim equalsPos As System.Int32 = line.IndexOf("="c)
                If equalsPos <= 0 Then
                    Throw New System.Configuration.ConfigurationErrorsException("Invalid provider catalog entry at line " & lineNumber.ToString())
                End If
                Dim key As System.String = line.Substring(0, equalsPos).Trim()
                Dim value As System.String = line.Substring(equalsPos + 1).Trim()
                If System.String.Equals(key, "Note", System.StringComparison.OrdinalIgnoreCase) Then
                    sectionNote = value
                ElseIf key.StartsWith("Field", System.StringComparison.OrdinalIgnoreCase) Then
                    Dim index As System.Int32
                    If Not System.Int32.TryParse(key.Substring(5), index) OrElse index <= 0 OrElse fields.ContainsKey(index) Then
                        Throw New System.Configuration.ConfigurationErrorsException("Invalid or duplicate provider catalog field at line " & lineNumber.ToString())
                    End If
                    fields.Add(index, value)
                Else
                    Throw New System.Configuration.ConfigurationErrorsException("Unexpected provider catalog key at line " & lineNumber.ToString() & ": " & key)
                End If
            Next
            flush()
            If configurations.Count = 0 Then
                Throw New System.Configuration.ConfigurationErrorsException("Provider catalog does not define any providers.")
            End If
            Return New ProviderTemplateCatalogSnapshot(configurations, notes, providerOrder)
        End Function

        ''' <summary>
        ''' Adds canonical provider connection values to both primary and secondary wizard branches.
        ''' The JSON owns the wizard layout and wizard-only defaults (e.g. OAuth2 flag or
        ''' APIKeyEncrypted); provider connection values live exclusively in the INI catalog.
        ''' </summary>
        ' <summary>
        ' Both configuration wizards use this identical opt-in update workflow.
        ' A download is never made before consent; online changes are only used for
        ' this wizard session after a second confirmation, not persisted to user INI.
        ' </summary>
        Public Shared Function OfferOnlineUpdate(current As ProviderTemplateCatalogSnapshot,
                                                 wizardName As System.String) As ProviderTemplateCatalogSnapshot
            If current Is Nothing Then Throw New System.ArgumentNullException(NameOf(current))
            Dim url As System.String = SharedMethods.RemoteDefaultsUrl
            Dim choice As System.Int32 = SharedMethods.ShowCustomYesNoBox(
                "Do you want to check " & url & " for updated provider templates for " & wizardName & "?",
                "Check online", "Keep embedded defaults")
            If choice <> 1 Then Return current

            Dim updated As ProviderTemplateCatalogSnapshot
            Try
                updated = DownloadAndParse(url, 10000)
                ValidateSupportedOnlineFields(current, updated)
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("Provider template update failed: " & ex.ToString())
                SharedMethods.ShowCustomMessageBox(
                    "Online provider templates could not be loaded or validated. The embedded catalog remains in use." &
                    System.Environment.NewLine & ex.Message)
                Return current
            End Try

            If AreEquivalent(current, updated) Then
                SharedMethods.ShowCustomMessageBox("Provider templates are up to date. The current catalog remains in use.")
                Return current
            End If

            Dim accept As System.Int32 = SharedMethods.ShowCustomYesNoBox(
                "Updated provider templates are available online. Use them for this wizard session?" &
                System.Environment.NewLine & "Loading templates will not itself change or save any configuration file.",
                "Use online templates", "Keep embedded defaults")
            Return If(accept = 1, updated, current)
        End Function

        ' <summary>
        ' Downloads the canonical INI schema with a bounded timeout. Parsing rejects
        ' invalid remote content before it can affect either wizard.
        ' </summary>
        Public Shared Function DownloadAndParse(url As System.String,
                                                timeoutMs As System.Int32) As ProviderTemplateCatalogSnapshot
            If System.String.IsNullOrWhiteSpace(url) OrElse
               Not System.Uri.IsWellFormedUriString(url, System.UriKind.Absolute) OrElse
               Not url.StartsWith("https://", System.StringComparison.OrdinalIgnoreCase) Then
                Throw New System.Configuration.ConfigurationErrorsException("A secure HTTPS provider template URL is required.")
            End If
            System.Net.ServicePointManager.SecurityProtocol =
                System.Net.ServicePointManager.SecurityProtocol Or System.Net.SecurityProtocolType.Tls12
            Using handler As New System.Net.Http.HttpClientHandler() With {
                .AutomaticDecompression = System.Net.DecompressionMethods.GZip Or System.Net.DecompressionMethods.Deflate
            }
                Using client As New System.Net.Http.HttpClient(handler)
                    client.Timeout = System.TimeSpan.FromMilliseconds(System.Math.Max(1000, timeoutMs))
                    Dim ini As System.String = client.GetStringAsync(url).GetAwaiter().GetResult()
                    Return Parse(ini)
                End Using
            End Using
        End Function

        ' Reject remote fields neither wizard could safely persist. New provider names
        ' and new values for supported fields are supported without a code update;
        ' unknown fields require a matching host/schema update first.
        Public Shared Sub ValidateSupportedOnlineFields(current As ProviderTemplateCatalogSnapshot,
                                                       remote As ProviderTemplateCatalogSnapshot)
            If current Is Nothing OrElse remote Is Nothing Then
                Throw New System.Configuration.ConfigurationErrorsException("Provider catalog validation requires two catalogs.")
            End If
            Dim supported As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            For Each variables As System.Collections.Generic.List(Of AppConfigurationVariable) In current.Configurations.Values
                For Each variable As AppConfigurationVariable In variables
                    supported.Add(variable.VarName)
                Next
            Next
            For Each provider As System.Collections.Generic.KeyValuePair(Of System.String, System.Collections.Generic.List(Of AppConfigurationVariable)) In remote.Configurations
                For Each variable As AppConfigurationVariable In provider.Value
                    If Not supported.Contains(variable.VarName) Then
                        Throw New System.Configuration.ConfigurationErrorsException(
                            "Online provider '" & provider.Key & "' defines a configuration field not supported by the installed application: " & variable.VarName)
                    End If
                Next
            Next
        End Sub

        ' <summary>Semantically compares provider entries including notes and every field.</summary>
        Public Shared Function AreEquivalent(first As ProviderTemplateCatalogSnapshot,
                                             second As ProviderTemplateCatalogSnapshot) As System.Boolean
            If first Is Nothing OrElse second Is Nothing Then Return first Is second
            If first.Configurations.Count <> second.Configurations.Count Then Return False
            If first.ProviderOrder.Count <> second.ProviderOrder.Count Then Return False
            For i As System.Int32 = 0 To first.ProviderOrder.Count - 1
                If Not System.String.Equals(first.ProviderOrder(i), second.ProviderOrder(i), System.StringComparison.OrdinalIgnoreCase) Then Return False
            Next
            For Each item As System.Collections.Generic.KeyValuePair(Of System.String, System.Collections.Generic.List(Of AppConfigurationVariable)) In first.Configurations
                Dim other As System.Collections.Generic.List(Of AppConfigurationVariable) = Nothing
                If Not second.Configurations.TryGetValue(item.Key, other) OrElse other.Count <> item.Value.Count Then Return False
                Dim aNote As System.String = Nothing
                Dim bNote As System.String = Nothing
                first.Notes.TryGetValue(item.Key, aNote)
                second.Notes.TryGetValue(item.Key, bNote)
                If Not ValuesEqual(aNote, bNote) Then Return False
                Dim byKey As New System.Collections.Generic.Dictionary(Of System.String, AppConfigurationVariable)(System.StringComparer.OrdinalIgnoreCase)
                For Each v As AppConfigurationVariable In other
                    byKey.Add(v.VarName, v)
                Next
                For Each v As AppConfigurationVariable In item.Value
                    Dim b As AppConfigurationVariable = Nothing
                    If Not byKey.TryGetValue(v.VarName, b) Then Return False
                    If Not ValuesEqual(v.DisplayName, b.DisplayName) OrElse
                       Not ValuesEqual(v.VarType, b.VarType) OrElse
                       Not ValuesEqual(v.ValidationRule, b.ValidationRule) OrElse
                       Not ValuesEqual(v.DefaultValue, b.DefaultValue) OrElse
                       Not ValuesEqual(v.CurrentValue, b.CurrentValue) Then Return False
                Next
            Next
            Return True
        End Function

        Private Shared Function ValuesEqual(a As System.String, b As System.String) As System.Boolean
            If System.Object.ReferenceEquals(a, b) Then Return True
            If a Is Nothing OrElse b Is Nothing Then Return System.String.IsNullOrEmpty(a) AndAlso System.String.IsNullOrEmpty(b)
            Dim left As System.String = a.Replace(vbCrLf, vbLf).Trim()
            Dim right As System.String = b.Replace(vbCrLf, vbLf).Trim()
            If System.String.Equals(left, right, System.StringComparison.Ordinal) Then Return True
            Dim jsonLeft As System.Boolean = left.Contains(":") AndAlso (left.Contains("{") OrElse left.Contains("["))
            Dim jsonRight As System.Boolean = right.Contains(":") AndAlso (right.Contains("{") OrElse right.Contains("["))
            Return jsonLeft AndAlso jsonRight AndAlso System.String.Equals(RemoveJsonWhitespace(left), RemoveJsonWhitespace(right), System.StringComparison.Ordinal)
        End Function

        Private Shared Function RemoveJsonWhitespace(source As System.String) As System.String
            Dim output As New System.Text.StringBuilder(source.Length)
            Dim inString As System.Boolean = False
            Dim escaped As System.Boolean = False
            For Each ch As System.Char In source
                If inString Then
                    output.Append(ch)
                    If escaped Then
                        escaped = False
                    ElseIf ch = "\"c Then
                        escaped = True
                    ElseIf ch = """"c Then
                        inString = False
                    End If
                ElseIf ch = """"c Then
                    inString = True
                    output.Append(ch)
                ElseIf Not System.Char.IsWhiteSpace(ch) Then
                    output.Append(ch)
                End If
            Next
            Return output.ToString()
        End Function

        Public Shared Sub ApplyToWizard(definition As WizardDefinition, catalog As ProviderTemplateCatalogSnapshot)
            If definition Is Nothing OrElse catalog Is Nothing Then
                Throw New System.Configuration.ConfigurationErrorsException("Wizard definition or provider catalog is missing.")
            End If

            ' Build every branch first and commit only after complete validation.
            ' This ensures an incomplete online update cannot partially mutate UI state.
            Dim replacements As New System.Collections.Generic.Dictionary(Of WizardGroup, System.Collections.Generic.List(Of WizardBranch))()
            For Each group As WizardGroup In definition.Groups
                If System.String.IsNullOrWhiteSpace(group.BranchField) OrElse group.Branches Is Nothing Then Continue For
                Dim suffix As System.String = If(group.KeySuffix, System.String.Empty)
                Dim allowed As New System.Collections.Generic.HashSet(Of System.String)(
                    group.Fields.Select(Function(f As WizardField) f.Key), System.StringComparer.OrdinalIgnoreCase)
                Dim existing As New System.Collections.Generic.Dictionary(Of System.String, WizardBranch)(System.StringComparer.OrdinalIgnoreCase)
                For Each branch As WizardBranch In group.Branches
                    existing.Add(branch.Label, branch)
                Next
                Dim newBranches As New System.Collections.Generic.List(Of WizardBranch)()
                For Each providerName As System.String In catalog.ProviderOrder
                    Dim providerVariables As System.Collections.Generic.List(Of AppConfigurationVariable) = catalog.Configurations(providerName)
                    Dim oldBranch As WizardBranch = Nothing
                    existing.TryGetValue(providerName, oldBranch)
                    Dim nextBranch As New WizardBranch With {.Label = providerName}
                    ' Preserve only wizard-specific defaults. The actual connection values
                    ' ALWAYS come from the one catalog (including after repeated updates).
                    If oldBranch IsNot Nothing Then
                        For Each entry As System.Collections.Generic.KeyValuePair(Of System.String, System.String) In oldBranch.Defaults
                            Dim keyWithoutSuffix As System.String = entry.Key
                            If suffix.Length > 0 AndAlso keyWithoutSuffix.EndsWith(suffix, System.StringComparison.OrdinalIgnoreCase) Then
                                keyWithoutSuffix = keyWithoutSuffix.Substring(0, keyWithoutSuffix.Length - suffix.Length)
                            End If
                            If System.String.Equals(keyWithoutSuffix, "APIKeyEncrypted", System.StringComparison.OrdinalIgnoreCase) OrElse
                               System.String.Equals(keyWithoutSuffix, "APIKeyPrefix", System.StringComparison.OrdinalIgnoreCase) Then
                                nextBranch.Defaults(entry.Key) = entry.Value
                            End If
                        Next
                        If oldBranch.ShowExtra IsNot Nothing Then
                            For Each extra As System.String In oldBranch.ShowExtra
                                Dim basicExtra As System.String = extra
                                If suffix.Length > 0 AndAlso basicExtra.EndsWith(suffix, System.StringComparison.OrdinalIgnoreCase) Then
                                    basicExtra = basicExtra.Substring(0, basicExtra.Length - suffix.Length)
                                End If
                                If Not basicExtra.StartsWith("OAuth2", System.StringComparison.OrdinalIgnoreCase) Then
                                    nextBranch.ShowExtra.Add(extra)
                                End If
                            Next
                        End If
                    End If
                    Dim hasOAuth As System.Boolean = False
                    For Each variable As AppConfigurationVariable In providerVariables
                        Dim key As System.String = variable.VarName.Substring("INI_".Length) & suffix
                        If Not allowed.Contains(key) Then
                            Throw New System.Configuration.ConfigurationErrorsException(
                                "Provider catalog field is not defined in wizard: " & group.Id & "." & key)
                        End If
                        nextBranch.Defaults(key) = variable.DefaultValue
                        If variable.VarName.StartsWith("INI_OAuth2", System.StringComparison.OrdinalIgnoreCase) Then
                            hasOAuth = True
                            If Not nextBranch.ShowExtra.Contains(key) Then nextBranch.ShowExtra.Add(key)
                        End If
                    Next
                    If hasOAuth AndAlso allowed.Contains("OAuth2" & suffix) Then
                        nextBranch.Defaults("OAuth2" & suffix) = "True"
                        If Not nextBranch.ShowExtra.Contains("OAuth2" & suffix) Then nextBranch.ShowExtra.Add("OAuth2" & suffix)
                    End If
                    newBranches.Add(nextBranch)
                Next
                Dim customBranch As WizardBranch = Nothing
                If existing.TryGetValue("Custom / Other", customBranch) Then
                    newBranches.Add(customBranch)
                End If
                replacements.Add(group, newBranches)
            Next
            For Each entry As System.Collections.Generic.KeyValuePair(Of WizardGroup, System.Collections.Generic.List(Of WizardBranch)) In replacements
                entry.Key.Branches = entry.Value
            Next
        End Sub
    End Class
End Namespace
