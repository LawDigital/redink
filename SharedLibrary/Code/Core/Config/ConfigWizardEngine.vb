' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ConfigWizardEngine.vb
' Purpose: Script-driven configuration wizard engine for guided INI editing.
'          Loads a JSON wizard definition from an embedded resource, reads current
'          values from the target INI file, evaluates conditional field visibility,
'          validates input, and writes changes back with a timestamped backup.
'
' Architecture / Responsibilities:
'   - Data model: WizardDefinition / WizardGroup / WizardField / WizardBranch
'     represent the JSON-defined wizard structure.
'   - Access control: Restricts wizard use via CentralConfigClients (same pattern
'     as UpdateIniClients) and current client identifier.
'   - INI processing: Reads and writes INI values, preserving comments and layout,
'     and applies default-skip/write-eligibility rules shared with UpdateAppConfig.
'   - Backup: Creates timestamped .bak files using the same convention as other
'     config update workflows (CommitDryRunPlan / CreateTimestampedBackup).
'   - Validation: Enforces simple rules (NotEmpty, Hyperlink, E-Mail, >0, range).
'
' External dependencies:
'   - `SharedMethods`: INI path resolution, default comparisons, write rules, and
'     helper utilities (e.g., GetActiveConfigFilePath).
'   - `ConfigWizardForm`: UI that consumes these engine APIs.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

Namespace SharedLibrary

#Region "Data Model"

    ''' <summary>
    ''' Represents the entire wizard definition deserialized from ConfigWizard.json.
    ''' </summary>
    Public Class WizardDefinition
        ''' <summary>Definition version string from the JSON resource.</summary>
        Public Property Version As String = ""

        ''' <summary>Wizard title displayed by the UI.</summary>
        Public Property Title As String = ""

        ''' <summary>List of groups (pages) contained in the wizard.</summary>
        Public Property Groups As New List(Of WizardGroup)
    End Class

    ''' <summary>
    ''' Represents a single group (page) in the wizard, containing fields, optional branches, and a description.
    ''' </summary>
    Public Class WizardGroup
        ''' <summary>Group identifier used by the wizard definition.</summary>
        Public Property Id As String = ""

        ''' <summary>Display title for the group page.</summary>
        Public Property Title As String = ""

        ''' <summary>Optional description displayed under the group title.</summary>
        Public Property Description As String = ""

        ''' <summary>Optional condition expression controlling group visibility.</summary>
        Public Property Condition As String = ""

        ''' <summary>Optional suffix appended to field keys in this group.</summary>
        Public Property KeySuffix As String = ""

        ''' <summary>Optional field key that represents the active branch selector.</summary>
        Public Property BranchField As String = ""

        ''' <summary>Optional list of branches providing default values.</summary>
        Public Property Branches As New List(Of WizardBranch)

        ''' <summary>Fields defined for this group.</summary>
        Public Property Fields As New List(Of WizardField)

        ''' <summary>Optional note displayed at the bottom of the group.</summary>
        Public Property Note As String = ""
    End Class

    ''' <summary>
    ''' Represents a selectable branch within a group (e.g. provider template) that supplies default values.
    ''' </summary>
    Public Class WizardBranch
        ''' <summary>Display label for the branch option.</summary>
        Public Property Label As String = ""

        ''' <summary>Default values applied when this branch is selected.</summary>
        Public Property Defaults As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        ''' <summary>List of extra fields to show when this branch is selected.</summary>
        Public Property ShowExtra As New List(Of String)
    End Class

    ''' <summary>
    ''' Represents a single configurable field in a wizard group.
    ''' </summary>
    Public Class WizardField
        ''' <summary>INI key for the field.</summary>
        Public Property Key As String = ""

        ''' <summary>Display label for the field.</summary>
        Public Property Label As String = ""

        ''' <summary>Field type (e.g., string, boolean, integer, multiline, password).</summary>
        Public Property Type As String = "string"

        ''' <summary>Validation rule name or range expression.</summary>
        Public Property Validation As String = ""

        ''' <summary>Default value specified in the wizard definition.</summary>
        Public Property [Default] As String = ""

        ''' <summary>Optional descriptive text displayed below the field.</summary>
        Public Property Description As String = ""

        ''' <summary>Optional condition expression controlling field visibility.</summary>
        Public Property Condition As String = ""
    End Class

#End Region

#Region "Engine"

    ''' <summary>
    ''' Provides static methods for loading the wizard definition, checking access control,
    ''' reading/writing INI values, and creating timestamped backups.
    ''' </summary>
    Public NotInheritable Class ConfigWizardEngine

        ''' <summary>
        ''' Prevents instantiation; this class provides only shared members.
        ''' </summary>
        Private Sub New()
        End Sub

        ''' <summary>
        ''' Case-insensitive check for whether a file path already exists as a value
        ''' in the candidates dictionary. <see cref="Dictionary(Of String, String).ContainsValue"/>
        ''' uses ordinal comparison regardless of the key comparer, so this helper
        ''' ensures Windows file paths that differ only in casing are detected as duplicates.
        ''' </summary>
        Private Shared Function CandidatesContainPath(candidates As Dictionary(Of String, String), path As String) As Boolean
            Return candidates.Values.Any(Function(v) String.Equals(v, path, StringComparison.OrdinalIgnoreCase))
        End Function



        ''' <summary>
        ''' Resolves the target INI path for the wizard. When both a central and local INI exist,
        ''' prompts the user to choose which file to edit.
        ''' </summary>
        ''' <param name="context">Shared context providing application identifiers and path resolution.</param>
        ''' <param name="ownerForm">Optional owner form for the selection dialog.</param>
        ''' <returns>The selected INI file path, or Nothing if the user cancels.</returns>
        Public Shared Function ResolveWizardTargetPath(context As ISharedContext, Optional ownerForm As Form = Nothing) As String
            Dim activePath As String = GetActiveConfigFilePath(context)
            Dim localPath As String = GetDefaultINIPath(context.RDV)

            ' Determine the central path (registry-directed)
            Dim regPath As String = GetFromRegistry(RegPath_Base, RegPath_IniPath, True)
            Dim centralPath As String = ""
            If Not String.IsNullOrWhiteSpace(regPath) Then
                centralPath = Path.Combine(ExpandEnvironmentVariables(regPath), $"{AN2}.ini")
            End If

            ' Also check the Word default as a potential shared/central path
            Dim wordPath As String = GetDefaultINIPath("Word")

            ' Build list of distinct existing INI files
            Dim candidates As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            If Not String.IsNullOrWhiteSpace(centralPath) AndAlso File.Exists(centralPath) Then
                candidates($"Central ({centralPath})") = centralPath
            End If

            If File.Exists(localPath) AndAlso Not CandidatesContainPath(candidates, localPath) Then
                candidates($"Local / {context.RDV} ({localPath})") = localPath
            End If

            If File.Exists(wordPath) AndAlso Not CandidatesContainPath(candidates, wordPath) AndAlso
               Not String.Equals(wordPath, localPath, StringComparison.OrdinalIgnoreCase) Then
                candidates($"Word shared ({wordPath})") = wordPath
            End If

            ' If only one file exists, use it directly
            If candidates.Count <= 1 Then
                Return activePath
            End If

            ' Multiple files exist — ask the user which to edit
            Dim choice As String = ShowSelectionForm(
                $"Multiple '{AN2}.ini' files were found. Select the one you want to edit with the Configuration Wizard:",
                "Select Configuration File",
                candidates.Keys)

            If String.IsNullOrWhiteSpace(choice) OrElse Not candidates.ContainsKey(choice) Then
                Return Nothing ' User cancelled
            End If

            Return candidates(choice)
        End Function

        ''' <summary>
        ''' Builds a dictionary of all discoverable INI file candidates for the current context.
        ''' Follows the same resolution logic as <see cref="InitializeConfig"/>: registry path,
        ''' host-specific local path (<c>context.RDV</c>), and Word shared fallback.
        ''' When the current host is Excel or Outlook and no host-specific INI exists, the Word
        ''' shared path is included as a candidate because <c>InitializeConfig</c> falls back to it.
        ''' </summary>
        ''' <param name="context">Shared context providing the current application identifier (<c>RDV</c>).</param>
        ''' <returns>
        ''' An ordered dictionary mapping display labels to full file paths.
        ''' Only paths where the file actually exists on disk are included.
        ''' </returns>
        Public Shared Function BuildIniCandidates(context As ISharedContext) As Dictionary(Of String, String)
            Dim candidates As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            ' 1. Central / registry-directed path
            Dim regPath As String = GetFromRegistry(RegPath_Base, RegPath_IniPath, True)
            Dim centralPath As String = ""
            If Not String.IsNullOrWhiteSpace(regPath) Then
                centralPath = Path.Combine(ExpandEnvironmentVariables(regPath), $"{AN2}.ini")
            End If
            If Not String.IsNullOrWhiteSpace(centralPath) AndAlso File.Exists(centralPath) Then
                candidates($"Central ({centralPath})") = centralPath
            End If

            ' 2. Host-specific local path (e.g. %AppData%\Microsoft\Excel\redink.ini)
            Dim localPath As String = GetDefaultINIPath(context.RDV)
            If File.Exists(localPath) AndAlso Not CandidatesContainPath(candidates, localPath) Then
                candidates($"Local / {context.RDV} ({localPath})") = localPath
            End If

            ' 3. Word shared path — always checked because Excel and Outlook fall back to it
            '    when they don't have their own INI (same logic as InitializeConfig: DefaultPath2)
            Dim wordPath As String = GetDefaultINIPath("Word")
            If File.Exists(wordPath) AndAlso Not CandidatesContainPath(candidates, wordPath) AndAlso
               Not String.Equals(wordPath, localPath, StringComparison.OrdinalIgnoreCase) Then
                candidates($"Word shared ({wordPath})") = wordPath
            End If

            Return candidates
        End Function

        ''' <summary>
        ''' Loads the wizard definition from the embedded resource ConfigWizard.json.
        ''' </summary>
        ''' <returns>The deserialized wizard definition.</returns>
        Public Shared Function LoadWizardDefinition() As WizardDefinition
            Dim asm = System.Reflection.Assembly.GetExecutingAssembly()
            Dim resourceName As String = ""

            For Each name In asm.GetManifestResourceNames()
                If name.EndsWith("ConfigWizard.json", StringComparison.OrdinalIgnoreCase) Then
                    resourceName = name
                    Exit For
                End If
            Next

            If String.IsNullOrWhiteSpace(resourceName) Then
                Throw New FileNotFoundException("Embedded resource 'ConfigWizard.json' not found.")
            End If

            Using stream = asm.GetManifestResourceStream(resourceName)
                Using reader As New StreamReader(stream, System.Text.Encoding.UTF8)
                    Dim json As String = reader.ReadToEnd()
                    Return JsonConvert.DeserializeObject(Of WizardDefinition)(json)
                End Using
            End Using
        End Function

        ''' <summary>
        ''' Reads the active INI file into a case-insensitive key/value dictionary.
        ''' </summary>
        ''' <param name="iniPath">Path to the INI file.</param>
        ''' <returns>Dictionary of INI key/value pairs.</returns>
        Public Shared Function ReadIniValues(iniPath As String) As Dictionary(Of String, String)
            Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            If String.IsNullOrWhiteSpace(iniPath) OrElse Not File.Exists(iniPath) Then Return result

            For Each line In File.ReadAllLines(iniPath)
                Dim trimmed = line.Trim()
                If String.IsNullOrEmpty(trimmed) OrElse trimmed.StartsWith(";") Then Continue For
                Dim parts = trimmed.Split({"="c}, 2)
                If parts.Length = 2 Then
                    result(parts(0).Trim()) = parts(1).Trim()
                End If
            Next
            Return result
        End Function

        ''' <summary>
        ''' Writes changed key/value pairs back to the INI file, preserving comments and structure.
        ''' Creates a timestamped backup before writing. Uses the same default-skip and write-eligibility
        ''' logic as <see cref="SharedMethods.UpdateAppConfig"/> via <see cref="SharedMethods.GetKeysToSkipWhenDefault"/>,
        ''' <see cref="SharedMethods.IsDefaultValue"/>, and <see cref="SharedMethods.ShouldWriteKey"/>.
        ''' </summary>
        ''' <param name="iniPath">Path to the INI file to update.</param>
        ''' <param name="editedValues">Only the changed key/value pairs to write.</param>
        Public Shared Sub WriteIniValues(iniPath As String, editedValues As Dictionary(Of String, String))
            If String.IsNullOrWhiteSpace(iniPath) Then
                Throw New ArgumentNullException(NameOf(iniPath))
            End If

            ' Create timestamped backup (same convention as CommitDryRunPlan)
            CreateWizardBackup(iniPath)

            ' Retrieve the canonical default-skip dictionary (same source as UpdateAppConfig)
            Dim defaults As Dictionary(Of String, Object) = GetKeysToSkipWhenDefault()

            Dim existingLines As New List(Of String)
            If File.Exists(iniPath) Then
                existingLines.AddRange(File.ReadAllLines(iniPath))
            End If

            Dim updatedContent As New System.Text.StringBuilder()
            Dim foundKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each line In existingLines
                Dim trimmed = line.Trim()

                ' Preserve comments and empty lines
                If String.IsNullOrEmpty(trimmed) OrElse trimmed.StartsWith(";") Then
                    updatedContent.AppendLine(line)
                    Continue For
                End If

                Dim parts = trimmed.Split({"="c}, 2)
                If parts.Length = 2 Then
                    Dim key = parts(0).Trim()
                    foundKeys.Add(key)

                    If editedValues.ContainsKey(key) Then
                        ' Key already exists in the file: always update it even if it matches default
                        ' (the user explicitly set it to this value; removing it would be surprising)
                        updatedContent.AppendLine($"{key} = {editedValues(key)}")
                    Else
                        updatedContent.AppendLine(line)
                    End If
                Else
                    updatedContent.AppendLine(line)
                End If
            Next

            ' Append new keys not found in the existing file — apply the same skip logic as UpdateAppConfig
            For Each kvp In editedValues
                If Not foundKeys.Contains(kvp.Key) Then
                    Dim value As String = kvp.Value

                    ' Skip if the value matches its built-in default (same as UpdateAppConfig)
                    If IsDefaultValue(kvp.Key, value, defaults) Then
                        Continue For
                    End If

                    ' Skip empty, whitespace-only, or implicitly-False boolean values (same as UpdateAppConfig)
                    If Not ShouldWriteKey(kvp.Key, value, defaults) Then
                        Continue For
                    End If

                    updatedContent.AppendLine($"{kvp.Key} = {value}")
                End If
            Next

            ' Atomic write via temp file + replace (same pattern as CommitDryRunPlan)
            Dim directory As String = Path.GetDirectoryName(iniPath)
            Dim baseName As String = Path.GetFileNameWithoutExtension(iniPath)
            Dim ext As String = Path.GetExtension(iniPath)
            Dim tmpPath As String = Path.Combine(directory, baseName & "_tmp_" & Guid.NewGuid().ToString("N") & ext)

            File.WriteAllText(tmpPath, updatedContent.ToString(), System.Text.Encoding.UTF8)

            Try
                If File.Exists(iniPath) Then
                    File.Replace(tmpPath, iniPath, Nothing, True)
                Else
                    File.Move(tmpPath, iniPath)
                End If
            Catch
                If File.Exists(iniPath) Then
                    Try : File.Delete(iniPath) : Catch : End Try
                End If
                File.Move(tmpPath, iniPath)
            End Try
        End Sub

        ''' <summary>
        ''' Creates a timestamped backup using the same naming convention as
        ''' CommitDryRunPlan and CreateTimestampedBackup: filename.ini_yyyyMMdd_HHmmss_fff.bak
        ''' Appends a short GUID suffix to avoid collisions when called within the same millisecond.
        ''' </summary>
        ''' <param name="iniPath">Path to the INI file to back up.</param>
        ''' <returns>Full path of the backup file, or Nothing if it could not be created.</returns>
        Public Shared Function CreateWizardBackup(iniPath As String) As String
            Try
                If String.IsNullOrWhiteSpace(iniPath) OrElse Not File.Exists(iniPath) Then
                    Return Nothing
                End If

                Dim directory As String = Path.GetDirectoryName(iniPath)
                Dim fileName As String = Path.GetFileName(iniPath)
                Dim timestamp As String = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff")
                Dim uniqueSuffix As String = Guid.NewGuid().ToString("N").Substring(0, 6)
                Dim backupFileName As String = $"{fileName}_{timestamp}_{uniqueSuffix}.bak"
                Dim backupPath As String = Path.Combine(directory, backupFileName)

                File.Copy(iniPath, backupPath, overwrite:=False)
                Return backupPath

            Catch ex As Exception
                Debug.WriteLine($"Failed to create wizard backup for {iniPath}: {ex.Message}")
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Evaluates a simple condition expression like "key == value" against the current values.
        ''' Returns True if the condition is met or if the condition is empty.
        ''' </summary>
        ''' <param name="condition">Condition expression to evaluate.</param>
        ''' <param name="currentValues">Current INI values dictionary.</param>
        ''' <returns>True when the condition is satisfied or empty.</returns>
        Public Shared Function EvaluateCondition(condition As String, currentValues As Dictionary(Of String, String)) As Boolean
            If String.IsNullOrWhiteSpace(condition) Then Return True

            ' Support "key == value" format
            Dim parts = condition.Split(New String() {"=="}, 2, StringSplitOptions.None)
            If parts.Length <> 2 Then Return True

            Dim key = parts(0).Trim()
            Dim expected = parts(1).Trim()

            Dim actual As String = ""
            If currentValues.ContainsKey(key) Then
                actual = currentValues(key)
            End If

            Return String.Equals(actual, expected, StringComparison.OrdinalIgnoreCase)
        End Function

        ''' <summary>
        ''' Validates a field value against the field's validation rule.
        ''' Returns an error message or empty string if valid.
        ''' </summary>
        ''' <param name="field">Field definition with validation metadata.</param>
        ''' <param name="value">Value to validate.</param>
        ''' <returns>Error message when invalid; otherwise an empty string.</returns>
        Public Shared Function ValidateField(field As WizardField, value As String) As String
            If String.IsNullOrWhiteSpace(field.Validation) Then Return ""

            Dim rule = field.Validation.Trim()

            If rule.Equals("NotEmpty", StringComparison.OrdinalIgnoreCase) Then
                If String.IsNullOrWhiteSpace(value) Then
                    Return $"'{field.Label}' is required."
                End If
            End If

            If rule.Equals("Hyperlink", StringComparison.OrdinalIgnoreCase) Then
                If Not String.IsNullOrWhiteSpace(value) Then
                    If Not value.StartsWith("http://", StringComparison.OrdinalIgnoreCase) AndAlso
                       Not value.StartsWith("https://", StringComparison.OrdinalIgnoreCase) Then
                        Return $"'{field.Label}' must start with http:// or https://."
                    End If
                End If
            End If

            If rule.Equals("E-Mail", StringComparison.OrdinalIgnoreCase) Then
                If Not String.IsNullOrWhiteSpace(value) AndAlso Not value.Contains("@") Then
                    Return $"'{field.Label}' must be a valid email address."
                End If
            End If

            If rule.Equals(">0", StringComparison.OrdinalIgnoreCase) Then
                Dim intVal As Integer
                If Not String.IsNullOrWhiteSpace(value) AndAlso Integer.TryParse(value, intVal) Then
                    If intVal <= 0 Then
                        Return $"'{field.Label}' must be greater than 0."
                    End If
                ElseIf Not String.IsNullOrWhiteSpace(value) Then
                    Return $"'{field.Label}' must be a positive integer."
                End If
            End If

            ' Range validation like "0.0-2.0" — use the last hyphen as the separator
            ' so that negative lower bounds (e.g. "-1.0-1.0") are handled correctly.
            If rule.Contains("-") AndAlso Not rule.StartsWith(">") Then
                Dim lastHyphen As Integer = rule.LastIndexOf("-"c)
                If lastHyphen > 0 Then
                    Dim minPart As String = rule.Substring(0, lastHyphen)
                    Dim maxPart As String = rule.Substring(lastHyphen + 1)
                    Dim minVal As Double
                    Dim maxVal As Double
                    Dim currentVal As Double
                    Dim normalizedValue = value.Replace(","c, "."c)
                    If Double.TryParse(minPart.Replace(","c, "."c), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, minVal) AndAlso
                       Double.TryParse(maxPart.Replace(","c, "."c), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, maxVal) Then
                        If Not String.IsNullOrWhiteSpace(value) Then
                            If Double.TryParse(normalizedValue, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, currentVal) Then
                                If currentVal < minVal OrElse currentVal > maxVal Then
                                    Return $"'{field.Label}' must be between {minVal} and {maxVal}."
                                End If
                            Else
                                Return $"'{field.Label}' must be a number between {minVal} and {maxVal}."
                            End If
                        End If
                    End If
                End If
            End If

            Return ""
        End Function

    End Class

#End Region

End Namespace
