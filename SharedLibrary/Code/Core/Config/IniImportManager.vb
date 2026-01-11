' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: IniImportManager.vb
' Purpose: Imports configuration settings into the active per-application INI file and
'          related sectioned INI files, supports downloading sample files, and provides
'          rollback to the most recent recognized backup.
'
' Architecture:
'  - Entry points:
'     * RunImportFromVariableConfigurationWindow / RunInteractiveImportProvidersOnly /
'       RunInteractiveImportOtherParameters / RunAutoImportOtherParameters
'       -> all route to RunImportInternal.
'     * RunDownloadSampleFiles downloads a central list, downloads referenced sample files,
'       writes them locally, and optionally updates INI path parameters.
'     * TryRollbackLastBackup restores the active INI from the latest recognized *.bak and
'       creates a safety backup first.
'
'  - Import processing (RunImportInternal):
'     * Source acquisition: URL (HTTPS only, size-limited) or local/UNC file path.
'     * Optional preview/edit: interactive mode opens a temp viewer and requires "Save" to proceed.
'     * Import kind selection: primary/secondary model, alternate model, special service, other parameters.
'     * Placeholder handling: prompts user for values for [[...]] placeholders and substitutes them.
'     * Target resolution:
'         - Primary/Secondary/OtherParameters -> main INI (activeIniPath).
'         - AlternateModel/SpecialService -> sectioned INIs (AlternateModelPath / SpecialServicePath),
'           may update main INI to store those paths and creates missing files.
'     * Trust enforcement: host-based trust list for HTTPS sources, with additional section marker rules
'       for restricted-trust hosts.
'     * Dry run: builds one or more DryRunPlan instances, summarizes changes, and asks for confirmation.
'     * Commit: writes backups, applies changes via atomic replace when possible, and returns whether the main INI changed.
'
'  - Backup/rollback:
'     * CommitDryRunPlan creates a timestamped full backup (*.bak) and optionally a removed-content backup.
'     * IsRecognizedIniBackupFile limits rollback candidates to module-created full backups (excludes *_removed_*).
'
' Notes:
'  - This module restricts import availability (CanUseImportFeature) based on configuration source rules.
'  - Network access is limited to HTTPS and a maximum download size (MAX_DOWNLOAD_BYTES).
' =============================================================================

Option Explicit On
Option Strict On

Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

''' <summary>
''' Provides configuration import, sample file download, and rollback support for per-application INI configuration.
''' </summary>
Public NotInheritable Class IniImportManager

    Private Sub New()
    End Sub

    ''' <summary>
    ''' Maximum number of bytes allowed when downloading import sources or sample files over HTTPS.
    ''' </summary>
    Private Const MAX_DOWNLOAD_BYTES As System.Int32 = 500 * 1024

    ''' <summary>
    ''' UI title used for import-related dialogs.
    ''' </summary>
    Private Shared Property TITLE_IMPORT As System.String = AN & " Get Settings"

    ''' <summary>
    ''' UI title used for sample-file download dialogs.
    ''' </summary>
    Private Shared Property TITLE_SAMPLE_FILES As System.String = AN & " Download Sample Files"

    ''' <summary>
    ''' Validates backup filenames created by this module for rollback selection.
    ''' Matches: anyname.ini_yyyyMMdd_HHmmss.bak or anyname.ini_yyyyMMdd_HHmmss_fff.bak
    ''' </summary>
    Private Shared ReadOnly BAK_SIGNATURE_REGEX As New System.Text.RegularExpressions.Regex(
    "^.+\.ini_\d{8}_\d{6}(_\d{3})?\.bak$",
    System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.CultureInvariant
)

    ''' <summary>
    ''' Represents a placeholder that was successfully resolved during import processing,
    ''' including the placeholder name and the user-provided replacement value.
    ''' </summary>
    ''' <remarks>
    ''' Instances of this class are collected during placeholder resolution and later used
    ''' to inject explanatory comments into the resulting INI output.
    ''' </remarks>
    Private NotInheritable Class ResolvedPlaceholder
        ''' <summary>
        ''' The placeholder name without surrounding brackets (e.g. "API_KEY").
        ''' </summary>
        Public Property Name As System.String

        ''' <summary>
        ''' The value entered by the user that replaced the placeholder.
        ''' </summary>
        Public Property Value As System.String
    End Class



    ''' <summary>
    ''' When True, user-facing summaries include download/parsing errors for skipped items.
    ''' </summary>
    Private Const InformOnErrors As System.Boolean = True

    ''' <summary>
    ''' Controls which interactive steps are shown to the user for an import run.
    ''' </summary>
    Private Enum ImportMode
        InteractiveAllKinds
        InteractiveWithoutOtherParameters
        DirectOtherParameters
    End Enum

    ''' <summary>
    ''' Represents the semantic category of settings to import.
    ''' </summary>
    Private Enum ImportKind
        PrimaryModel = 1
        SecondaryModel = 2
        AlternateModel = 3
        SpecialService = 4
        OtherParameters = 5
    End Enum

    ''' <summary>
    ''' Options for non-interactive "auto import" runs.
    ''' </summary>
    Private NotInheritable Class AutoImportOptions
        Public Property Source As String
        Public Property ReplaceExisting As Boolean
        Public Property UserConfirmation As Boolean
    End Class

    ' -----------------------------------------------------------------------------------------
    '  ROLLBACK FEATURE (latest .bak by signature)
    ' -----------------------------------------------------------------------------------------

    ''' <summary>
    ''' Scans the active configuration directory, finds the latest backup created by this module (by filename signature),
    ''' asks for confirmation, backs up the current active INI, and rolls back to that backup.
    ''' </summary>
    ''' <param name="context">Shared context used to resolve the active configuration file path.</param>
    ''' <param name="ownerForm">Owner window for UI dialogs.</param>
    ''' <returns>True if rollback changed the main INI file; otherwise False.</returns>
    Public Shared Function TryRollbackLastBackup(
        context As ISharedContext,
        ownerForm As System.Windows.Forms.Form
    ) As System.Boolean

        If context Is Nothing Then
            ShowCustomMessageBox("Internal error: context is missing.")
            Return False
        End If

        Dim activeIniPath As System.String = Nothing
        Try
            activeIniPath = GetActiveConfigFilePath(context)
        Catch ex As System.Exception
            ShowCustomMessageBox("Could not determine active configuration file path: " & ex.Message)
            Return False
        End Try

        If System.String.IsNullOrWhiteSpace(activeIniPath) Then
            ShowCustomMessageBox("No active configuration file path found.")
            Return False
        End If

        Dim disableReason As System.String = Nothing
        If Not CanUseImportFeature(context, activeIniPath, disableReason) Then
            ShowCustomMessageBox(disableReason)
            Return False
        End If

        If Not System.IO.File.Exists(activeIniPath) Then
            ShowCustomMessageBox("The main configuration file does not exist: " & activeIniPath)
            Return False
        End If

        Dim configDir As System.String = System.IO.Path.GetDirectoryName(activeIniPath)
        If System.String.IsNullOrWhiteSpace(configDir) OrElse Not System.IO.Directory.Exists(configDir) Then
            ShowCustomMessageBox("Could not determine configuration directory.")
            Return False
        End If

        Dim latestBak As System.IO.FileInfo = Nothing
        Try
            Dim di As New System.IO.DirectoryInfo(configDir)
            For Each fi As System.IO.FileInfo In di.GetFiles("*.bak", System.IO.SearchOption.TopDirectoryOnly)
                If fi Is Nothing Then Continue For
                If Not IsRecognizedIniBackupFile(fi.Name) Then Continue For

                If latestBak Is Nothing OrElse fi.LastWriteTimeUtc > latestBak.LastWriteTimeUtc Then
                    latestBak = fi
                End If
            Next
        Catch ex As System.Exception
            ShowCustomMessageBox("Could not scan configuration directory: " & ex.Message)
            Return False
        End Try

        If latestBak Is Nothing Then
            ShowCustomMessageBox("No recognizable backup (.bak) file was found in:" & System.Environment.NewLine & configDir)
            Return False
        End If

        ' Determine the original INI filename from the backup name
        Dim originalIniName As System.String = ExtractOriginalIniName(latestBak.Name)
        If System.String.IsNullOrWhiteSpace(originalIniName) Then
            ShowCustomMessageBox("Could not determine the original file name from backup: " & latestBak.Name)
            Return False
        End If

        ' Build the target path (the file to restore)
        Dim targetIniPath As System.String = System.IO.Path.Combine(configDir, originalIniName)

        ' Check if the target file exists; if not, we can still restore (creates the file)
        Dim targetExists As System.Boolean = System.IO.File.Exists(targetIniPath)

        ' Return True only if the main configuration file content changes.
        Dim mainIniWasRolledBack As System.Boolean = False

        Dim beforeFingerprint As System.String = Nothing
        If targetExists Then
            Try
                beforeFingerprint = GetFileFingerprint(targetIniPath)
            Catch ex As System.Exception
                ShowCustomMessageBox("Could not read current configuration file: " & ex.Message)
                Return False
            End Try
        End If

        Dim msg As System.String =
            "A configuration rollback was requested." & System.Environment.NewLine & System.Environment.NewLine &
            "Target file to restore:" & System.Environment.NewLine &
            "  " & targetIniPath & System.Environment.NewLine & System.Environment.NewLine &
            "Latest backup found:" & System.Environment.NewLine &
            "  " & latestBak.FullName & System.Environment.NewLine &
            "  (" & latestBak.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss") & ")" & System.Environment.NewLine & System.Environment.NewLine &
            If(targetExists, "Proceed with rollback? The current file will be backed up first.", "Proceed with rollback? The file will be created from the backup.")

        Dim decision As System.Int32 = ShowCustomYesNoBox(msg, "Yes, rollback", "No, cancel", TITLE_IMPORT)
        If decision <> 1 Then
            Return False
        End If

        Dim ts As System.String = System.DateTime.Now.ToString("yyyyMMdd_HHmmss_fff")
        Dim safetyBackup As System.String = System.IO.Path.Combine(configDir, originalIniName & "_rollback_" & ts & ".bak")

        If targetExists Then
            Try
                System.IO.File.Copy(targetIniPath, safetyBackup, overwrite:=False)
            Catch ex As System.Exception
                ShowCustomMessageBox("Could not create safety backup before rollback: " & ex.Message)
                Return False
            End Try
        End If

        Dim tmpPath As System.String = targetIniPath & ".tmp"

        Try
            System.IO.File.Copy(latestBak.FullName, tmpPath, overwrite:=True)

            Try
                If targetExists Then
                    System.IO.File.Replace(tmpPath, targetIniPath, Nothing, True)
                Else
                    System.IO.File.Move(tmpPath, targetIniPath)
                End If
            Catch
                Try
                    If System.IO.File.Exists(targetIniPath) Then System.IO.File.Delete(targetIniPath)
                Catch
                End Try
                System.IO.File.Move(tmpPath, targetIniPath)
            End Try

        Catch ex As System.Exception
            Try
                If System.IO.File.Exists(tmpPath) Then System.IO.File.Delete(tmpPath)
            Catch
            End Try

            ShowCustomMessageBox("Rollback failed: " & ex.Message)
            Return False
        End Try

        ' Determine if the main INI was affected
        Dim isMainIni As System.Boolean = System.String.Equals(
            targetIniPath, activeIniPath, System.StringComparison.OrdinalIgnoreCase)

        If isMainIni AndAlso targetExists Then
            Try
                Dim afterFingerprint As System.String = GetFileFingerprint(targetIniPath)
                mainIniWasRolledBack =
                    Not System.String.Equals(beforeFingerprint, afterFingerprint, System.StringComparison.Ordinal)
            Catch
                mainIniWasRolledBack = False
            End Try
        ElseIf isMainIni AndAlso Not targetExists Then
            mainIniWasRolledBack = True
        End If

        Dim summaryMsg As System.String =
            "Rollback completed." & System.Environment.NewLine & System.Environment.NewLine &
            "Restored file:" & System.Environment.NewLine &
            targetIniPath & System.Environment.NewLine & System.Environment.NewLine &
            "From backup:" & System.Environment.NewLine &
            latestBak.FullName & System.Environment.NewLine

        If targetExists Then
            summaryMsg &= System.Environment.NewLine &
                "Safety backup of the previous file:" & System.Environment.NewLine &
                safetyBackup & System.Environment.NewLine
        End If

        ShowCustomMessageBox(summaryMsg)

        Return mainIniWasRolledBack

    End Function

    ''' <summary>
    ''' Extracts the original INI filename from a recognized backup filename.
    ''' For example: "allmodels.ini_20251230_120000.bak" returns "allmodels.ini"
    ''' </summary>
    ''' <param name="bakFileName">Backup file name (no directory).</param>
    ''' <returns>Original INI filename, or Nothing if not recognized.</returns>
    Private Shared Function ExtractOriginalIniName(bakFileName As System.String) As System.String
        If System.String.IsNullOrWhiteSpace(bakFileName) Then Return Nothing
        If Not IsRecognizedIniBackupFile(bakFileName) Then Return Nothing

        ' Pattern: anyname.ini_yyyyMMdd_HHmmss.bak or anyname.ini_yyyyMMdd_HHmmss_fff.bak
        ' We need to extract "anyname.ini" from this.
        Dim rx As New System.Text.RegularExpressions.Regex(
            "^(.+\.ini)_\d{8}_\d{6}(_\d{3})?\.bak$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.CultureInvariant
        )

        Dim m As System.Text.RegularExpressions.Match = rx.Match(bakFileName)
        If m.Success AndAlso m.Groups.Count >= 2 Then
            Return m.Groups(1).Value
        End If

        Return Nothing
    End Function

    ''' <summary>
    ''' Computes a stable fingerprint for a file based on length, timestamp, and SHA-256 hash.
    ''' </summary>
    ''' <param name="path">Path to the file.</param>
    ''' <returns>Fingerprint string.</returns>
    Private Shared Function GetFileFingerprint(path As System.String) As System.String
        Dim fi As New System.IO.FileInfo(path)
        Dim bytes As System.Byte() = System.IO.File.ReadAllBytes(path)

        Using sha As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
            Dim hash As System.Byte() = sha.ComputeHash(bytes)
            Dim hashHex As System.String = BitConverter.ToString(hash).Replace("-", "")
            Return fi.Length.ToString() & "|" & fi.LastWriteTimeUtc.Ticks.ToString() & "|" & hashHex
        End Using
    End Function

    ''' <summary>
    ''' Returns True if a file name matches the full-backup naming convention used by this module.
    ''' Rollback safety backups (*_rollback_*) are intentionally excluded to prevent restoring
    ''' the state that was just rolled back from.
    ''' </summary>
    ''' <param name="fileName">Backup file name (no directory).</param>
    Private Shared Function IsRecognizedIniBackupFile(fileName As System.String) As System.Boolean
        If System.String.IsNullOrWhiteSpace(fileName) Then Return False

        ' Only full backups are rollback candidates; removed-content backups and rollback safety
        ' backups are excluded. The regex requires the pattern: name.ini_YYYYMMDD_HHMMSS[_fff].bak
        ' which excludes *_rollback_* and *_removed_* patterns.
        If Not fileName.EndsWith(".bak", System.StringComparison.OrdinalIgnoreCase) Then Return False

        Return BAK_SIGNATURE_REGEX.IsMatch(fileName)
    End Function


    ' =========================================================================================
    '  SAMPLE FILE DOWNLOAD FEATURE
    ' =========================================================================================

    ''' <summary>
    ''' Represents a parsed entry from the sample files list.
    ''' </summary>
    Private NotInheritable Class SampleFileEntry
        Public Property FriendlyName As System.String
        Public Property SourceURL As System.String
        Public Property Apps As System.String()
        Public Property ParameterForPath As System.String
        Public Property IsFullPath As System.Boolean
        Public Property DefaultDir As System.String
        Public Property DefaultFileName As System.String

        Public Property DownloadedContent As System.String
        Public Property TargetPath As System.String
        Public Property TargetPathForConfig As System.String
        Public Property WillOverwrite As System.Boolean
        Public Property NeedsConfigUpdate As System.Boolean
        Public Property ConfigValueToStore As System.String
    End Class

    ''' <summary>
    ''' Downloads sample files from the central list, writes them locally, and optionally updates INI path parameters.
    ''' </summary>
    ''' <param name="context">Shared context used to resolve the active configuration file path.</param>
    ''' <param name="ownerForm">Owner window for UI dialogs.</param>
    ''' <returns>True if the main configuration file was updated.</returns>
    Public Shared Function RunDownloadSampleFiles(
        context As ISharedContext,
        ownerForm As System.Windows.Forms.Form
    ) As System.Boolean

        Dim mainIniChanged As System.Boolean = False
        Dim errors As New System.Collections.Generic.List(Of System.String)()

        If context Is Nothing Then
            ShowCustomMessageBox("Internal error: context is missing.")
            Return False
        End If

        Dim activeIniPath As System.String = Nothing
        Try
            activeIniPath = GetActiveConfigFilePath(context)
        Catch ex As System.Exception
            ShowCustomMessageBox("Could not determine active configuration file path: " & ex.Message)
            Return False
        End Try

        If System.String.IsNullOrWhiteSpace(activeIniPath) Then
            ShowCustomMessageBox("No active configuration file path found.")
            Return False
        End If

        Dim disableReason As System.String = Nothing
        If Not CanUseImportFeature(context, activeIniPath, disableReason) Then
            ShowCustomMessageBox(disableReason)
            Return False
        End If

        If Not System.IO.File.Exists(activeIniPath) Then
            ShowCustomMessageBox("The main configuration file does not exist: " & activeIniPath)
            Return False
        End If

        Dim currentApp As System.String = ExtractAppFromRDV(context.RDV)
        If System.String.IsNullOrWhiteSpace(currentApp) Then
            ShowCustomMessageBox("Could not determine the current application.")
            Return False
        End If

        Dim listText As System.String = Nothing
        Try
            Debug.WriteLine("Downloading sample files list from: " & SampleFilesListURL)
            Dim listUri As New System.Uri(SampleFilesListURL)
            listText = DownloadHttpsTextWithLimit(listUri, MAX_DOWNLOAD_BYTES)
        Catch ex As System.Exception
            ShowCustomMessageBox($"Could not download the sample files list ({SampleFilesListURL}): " & ex.Message)
            Return False
        End Try

        If System.String.IsNullOrWhiteSpace(listText) Then
            ShowCustomMessageBox("The sample files list is empty.")
            Return False
        End If

        Dim allEntries As System.Collections.Generic.List(Of SampleFileEntry) =
            ParseSampleFilesList(listText, currentApp, errors)

        If allEntries.Count = 0 Then
            Dim msg As System.String = "No sample files available for " & currentApp & "."
            If InformOnErrors AndAlso errors.Count > 0 Then
                msg &= System.Environment.NewLine & System.Environment.NewLine &
                       "Errors encountered:" & System.Environment.NewLine &
                       System.String.Join(System.Environment.NewLine, errors)
            End If
            ShowCustomMessageBox(msg)
            Return False
        End If

        Dim validEntries As New System.Collections.Generic.List(Of SampleFileEntry)()

        For Each entry As SampleFileEntry In allEntries
            Try
                Dim sourceUri As New System.Uri(entry.SourceURL)
                entry.DownloadedContent = DownloadHttpsTextWithLimit(sourceUri, MAX_DOWNLOAD_BYTES)

                If System.String.IsNullOrWhiteSpace(entry.DownloadedContent) Then
                    errors.Add("Empty content from: " & entry.SourceURL)
                    Continue For
                End If
            Catch ex As System.Exception
                errors.Add("Download failed for '" & entry.FriendlyName & "': " & ex.Message)
                Continue For
            End Try

            If Not TryDetermineTargetPath(entry, activeIniPath, errors) Then
                Continue For
            End If

            entry.WillOverwrite = System.IO.File.Exists(entry.TargetPath)

            DetermineConfigUpdateNeeded(entry, activeIniPath)

            validEntries.Add(entry)
        Next

        If validEntries.Count = 0 Then
            Dim msg As System.String = "No sample files could be processed."
            If InformOnErrors AndAlso errors.Count > 0 Then
                msg &= System.Environment.NewLine & System.Environment.NewLine &
                       "Errors encountered:" & System.Environment.NewLine &
                       System.String.Join(System.Environment.NewLine, errors)
            End If
            ShowCustomMessageBox(msg)
            Return False
        End If

        Dim sbSummary As New System.Text.StringBuilder()
        sbSummary.AppendLine("The following sample files will be downloaded:")
        sbSummary.AppendLine()

        Dim newFiles As New System.Collections.Generic.List(Of SampleFileEntry)()
        Dim overwriteFiles As New System.Collections.Generic.List(Of SampleFileEntry)()

        For Each entry As SampleFileEntry In validEntries
            If entry.WillOverwrite Then
                overwriteFiles.Add(entry)
            Else
                newFiles.Add(entry)
            End If
        Next

        If newFiles.Count > 0 Then
            sbSummary.AppendLine("NEW FILES:")
            For Each entry As SampleFileEntry In newFiles
                sbSummary.AppendLine("  • " & entry.FriendlyName)
                sbSummary.AppendLine("    → " & entry.TargetPath)
            Next
            sbSummary.AppendLine()
        End If

        If overwriteFiles.Count > 0 Then
            sbSummary.AppendLine("FILES THAT WILL BE REPLACED:")
            For Each entry As SampleFileEntry In overwriteFiles
                sbSummary.AppendLine("  • " & entry.FriendlyName)
                sbSummary.AppendLine("    → " & entry.TargetPath)
            Next
            sbSummary.AppendLine()
        End If

        If InformOnErrors AndAlso errors.Count > 0 Then
            sbSummary.AppendLine("ERRORS (these files will be skipped):")
            For Each err As System.String In errors
                sbSummary.AppendLine("  • " & err)
            Next
            sbSummary.AppendLine()
        End If

        sbSummary.AppendLine("How do you want to proceed?")

        Dim userChoice As System.Int32
        If overwriteFiles.Count > 0 Then
            userChoice = ShowCustomYesNoBox(sbSummary.ToString(), "Replace existing files", "Skip existing files", TITLE_SAMPLE_FILES)
        Else
            userChoice = ShowCustomYesNoBox(sbSummary.ToString(), "Proceed", "Abort", TITLE_SAMPLE_FILES)
            If userChoice = 0 Then
                Return False
            End If
            userChoice = 2
        End If

        If userChoice = 0 Then
            Return False
        End If

        Dim replaceExisting As System.Boolean = (userChoice = 1)

        Dim entriesToProcess As New System.Collections.Generic.List(Of SampleFileEntry)()
        For Each entry As SampleFileEntry In validEntries
            If entry.WillOverwrite AndAlso Not replaceExisting Then
                Continue For
            End If
            entriesToProcess.Add(entry)
        Next

        If entriesToProcess.Count = 0 Then
            ShowCustomMessageBox("No files to process after applying your choice.")
            Return False
        End If

        Dim configUpdates As New System.Collections.Generic.Dictionary(Of System.String, System.String)(
            System.StringComparer.OrdinalIgnoreCase)

        For Each entry As SampleFileEntry In entriesToProcess
            If entry.NeedsConfigUpdate AndAlso Not System.String.IsNullOrWhiteSpace(entry.ConfigValueToStore) Then
                configUpdates(entry.ParameterForPath) = entry.ConfigValueToStore
            End If
        Next

        If configUpdates.Count > 0 Then
            Dim sbConfig As New System.Text.StringBuilder()
            sbConfig.AppendLine("The following configuration parameters will be added or updated:")
            sbConfig.AppendLine()

            For Each kvp As System.Collections.Generic.KeyValuePair(Of System.String, System.String) In configUpdates
                sbConfig.AppendLine("  " & kvp.Key & " = " & kvp.Value)
            Next

            sbConfig.AppendLine()
            sbConfig.AppendLine("Proceed?")

            Dim configDecision As System.Int32 = ShowCustomYesNoBox(
                sbConfig.ToString(),
                "Yes, update configuration",
                "No, abort"
            )

            If configDecision <> 1 Then
                Return False
            End If
        End If

        Dim ts As System.String = System.DateTime.Now.ToString("yyyyMMdd_HHmmss")

        If configUpdates.Count > 0 Then
            Try
                Dim configDir As System.String = System.IO.Path.GetDirectoryName(activeIniPath)
                Dim configBaseName As System.String = System.IO.Path.GetFileNameWithoutExtension(activeIniPath)
                Dim configBaseExt As System.String = System.IO.Path.GetExtension(activeIniPath)
                Dim configBackup As System.String = System.IO.Path.Combine(configDir, configBaseName & configBaseExt & "_" & ts & ".bak")
                System.IO.File.Copy(activeIniPath, configBackup, overwrite:=False)
            Catch ex As System.Exception
                ShowCustomMessageBox("Could not create backup of configuration file: " & ex.Message)
                Return False
            End Try
        End If

        For Each entry As SampleFileEntry In entriesToProcess
            Try
                Dim targetDir As System.String = System.IO.Path.GetDirectoryName(entry.TargetPath)
                If Not System.String.IsNullOrWhiteSpace(targetDir) AndAlso Not System.IO.Directory.Exists(targetDir) Then
                    System.IO.Directory.CreateDirectory(targetDir)
                End If

                If entry.WillOverwrite Then
                    Dim fileBaseName As System.String = System.IO.Path.GetFileNameWithoutExtension(entry.TargetPath)
                    Dim configBaseExt As System.String = System.IO.Path.GetExtension(activeIniPath)
                    Dim fileBackup As System.String = System.IO.Path.Combine(targetDir, fileBaseName & configBaseExt & "_" & ts & ".bak")
                    System.IO.File.Copy(entry.TargetPath, fileBackup, overwrite:=False)
                End If

                System.IO.File.WriteAllText(entry.TargetPath, entry.DownloadedContent, System.Text.Encoding.UTF8)

            Catch ex As System.Exception
                errors.Add("Could not save '" & entry.FriendlyName & "': " & ex.Message)
            End Try
        Next

        If configUpdates.Count > 0 Then
            Try
                Dim existingLines As System.Collections.Generic.List(Of System.String) = ReadAllLinesPreserve(activeIniPath)

                Dim plan As New DryRunPlan() With {
                    .TargetIniPath = activeIniPath,
                    .Kind = ImportKind.OtherParameters,
                    .TargetSectionName = Nothing,
                    .NewFileLines = New System.Collections.Generic.List(Of System.String)(existingLines),
                    .RemovedLinesBackup = New System.Collections.Generic.List(Of System.String)(),
                    .OverwrittenKeys = New System.Collections.Generic.List(Of System.String)(),
                    .WillCreateRemovedBackup = False
                }

                ApplyMainIniKeyReplaceAppend(plan, existingLines, configUpdates, replaceExisting:=True)

                Dim tmpPath As System.String = activeIniPath & ".tmp"
                System.IO.File.WriteAllText(
                    tmpPath,
                    System.String.Join(System.Environment.NewLine, plan.NewFileLines),
                    System.Text.Encoding.UTF8
                )

                Try
                    System.IO.File.Replace(tmpPath, activeIniPath, Nothing, True)
                Catch
                    Try
                        System.IO.File.Delete(activeIniPath)
                    Catch
                    End Try
                    System.IO.File.Move(tmpPath, activeIniPath)
                End Try

                mainIniChanged = True

            Catch ex As System.Exception
                ShowCustomMessageBox("Could not update configuration file: " & ex.Message)
            End Try
        End If

        Dim sbFinal As New System.Text.StringBuilder()
        sbFinal.AppendLine("Sample files download completed.")

        If mainIniChanged Then
            sbFinal.AppendLine()
            sbFinal.AppendLine("The configuration file was updated. A restart may be required for changes to take effect.")
        End If

        If InformOnErrors AndAlso errors.Count > 0 Then
            sbFinal.AppendLine()
            sbFinal.AppendLine("Some errors occurred:")
            For Each err As System.String In errors
                sbFinal.AppendLine("  • " & err)
            Next
        End If

        ShowCustomMessageBox(sbFinal.ToString())

        Return mainIniChanged

    End Function

    ''' <summary>
    ''' Extracts the application name from the RDV string (e.g., "Word (V231312)" -> "Word").
    ''' </summary>
    ''' <param name="rdv">RDV identifier string.</param>
    ''' <returns>Application name or Nothing.</returns>
    Private Shared Function ExtractAppFromRDV(rdv As System.String) As System.String
        If System.String.IsNullOrWhiteSpace(rdv) Then Return Nothing

        Dim trimmed As System.String = rdv.Trim()
        Dim spaceIdx As System.Int32 = trimmed.IndexOf(" "c)

        If spaceIdx > 0 Then
            Return trimmed.Substring(0, spaceIdx)
        End If

        Return trimmed
    End Function

    ''' <summary>
    ''' Parses the sample files list, filters entries by the current application, and records parse errors.
    ''' </summary>
    ''' <param name="listText">Raw list text.</param>
    ''' <param name="currentApp">Current application name used for filtering.</param>
    ''' <param name="errors">Collector for parse errors.</param>
    ''' <returns>Filtered list of sample file entries.</returns>
    Private Shared Function ParseSampleFilesList(
        listText As System.String,
        currentApp As System.String,
        errors As System.Collections.Generic.List(Of System.String)
    ) As System.Collections.Generic.List(Of SampleFileEntry)

        Dim result As New System.Collections.Generic.List(Of SampleFileEntry)()
        Dim lines As System.Collections.Generic.List(Of System.String) = SplitToLinesPreserve(listText)

        For Each line As System.String In lines
            If line Is Nothing Then Continue For

            Dim trimmed As System.String = line.Trim()

            If System.String.IsNullOrWhiteSpace(trimmed) Then Continue For
            If trimmed.StartsWith(";", System.StringComparison.Ordinal) Then Continue For

            Dim parts As System.String() = trimmed.Split(New System.Char() {";"c}, System.StringSplitOptions.None)

            If parts.Length < 7 Then
                errors.Add("Malformed line (expected 7 fields): " & trimmed)
                Continue For
            End If

            Dim friendlyName As System.String = parts(0).Trim()
            Dim sourceURL As System.String = ExpandSourceUrlPlaceholders(parts(1).Trim())
            Dim appList As System.String = parts(2).Trim()
            Dim parameterForPath As System.String = parts(3).Trim()
            Dim isFullPathStr As System.String = parts(4).Trim()
            Dim defaultDir As System.String = parts(5).Trim()
            Dim defaultFileName As System.String = parts(6).Trim()

            If System.String.IsNullOrWhiteSpace(friendlyName) OrElse
               System.String.IsNullOrWhiteSpace(sourceURL) OrElse
               System.String.IsNullOrWhiteSpace(parameterForPath) OrElse
               System.String.IsNullOrWhiteSpace(defaultDir) OrElse
               System.String.IsNullOrWhiteSpace(defaultFileName) Then
                errors.Add("Missing required fields in line: " & trimmed)
                Continue For
            End If

            Dim isFullPath As System.Boolean
            If System.String.Equals(isFullPathStr, "True", System.StringComparison.OrdinalIgnoreCase) Then
                isFullPath = True
            ElseIf System.String.Equals(isFullPathStr, "False", System.StringComparison.OrdinalIgnoreCase) Then
                isFullPath = False
            Else
                errors.Add("Invalid IsFullPath value in line: " & trimmed)
                Continue For
            End If

            Dim apps As System.String() = appList.Split(New System.Char() {","c}, System.StringSplitOptions.RemoveEmptyEntries)
            Dim appMatches As System.Boolean = False

            For i As System.Int32 = 0 To apps.Length - 1
                apps(i) = apps(i).Trim()
                If System.String.Equals(apps(i), currentApp, System.StringComparison.OrdinalIgnoreCase) Then
                    appMatches = True
                End If
            Next

            If Not appMatches Then Continue For

            If Not sourceURL.StartsWith("https://", System.StringComparison.OrdinalIgnoreCase) Then
                errors.Add("Invalid URL (must be HTTPS): " & sourceURL)
                Continue For
            End If

            Dim entry As New SampleFileEntry() With {
                .FriendlyName = friendlyName,
                .SourceURL = sourceURL,
                .Apps = apps,
                .ParameterForPath = parameterForPath,
                .IsFullPath = isFullPath,
                .DefaultDir = defaultDir,
                .DefaultFileName = defaultFileName
            }

            result.Add(entry)
        Next

        Return result
    End Function

    ''' <summary>
    ''' Resolves the download target path for a sample file entry based on config and defaults.
    ''' </summary>
    ''' <param name="entry">Entry to update.</param>
    ''' <param name="activeIniPath">Active main INI file path.</param>
    ''' <param name="errors">Collector for path resolution errors.</param>
    ''' <returns>True if a target path could be determined; otherwise False.</returns>
    Private Shared Function TryDetermineTargetPath(
        entry As SampleFileEntry,
        activeIniPath As System.String,
        errors As System.Collections.Generic.List(Of System.String)
    ) As System.Boolean

        Try
            Dim existingValue As System.String = TryReadIniKeyValue(activeIniPath, entry.ParameterForPath)
            Dim expandedExisting As System.String = Nothing

            If Not System.String.IsNullOrWhiteSpace(existingValue) Then
                Try
                    expandedExisting = ExpandEnvironmentVariables(existingValue)
                Catch
                    expandedExisting = existingValue
                End Try
            End If

            Dim expandedDefaultDir As System.String = Nothing
            Try
                expandedDefaultDir = ExpandEnvironmentVariables(entry.DefaultDir)
            Catch
                expandedDefaultDir = entry.DefaultDir
            End Try

            If entry.IsFullPath Then
                If Not System.String.IsNullOrWhiteSpace(expandedExisting) Then
                    Dim existingFileName As System.String = System.IO.Path.GetFileName(expandedExisting)

                    If Not System.String.IsNullOrWhiteSpace(existingFileName) AndAlso existingFileName.Contains(".") Then
                        entry.TargetPath = expandedExisting
                        entry.TargetPathForConfig = existingValue
                    Else
                        entry.TargetPath = System.IO.Path.Combine(expandedExisting, entry.DefaultFileName)
                        entry.TargetPathForConfig = System.IO.Path.Combine(existingValue, entry.DefaultFileName)
                    End If
                Else
                    entry.TargetPath = System.IO.Path.Combine(expandedDefaultDir, entry.DefaultFileName)
                    entry.TargetPathForConfig = System.IO.Path.Combine(entry.DefaultDir, entry.DefaultFileName)
                End If
            Else
                If Not System.String.IsNullOrWhiteSpace(expandedExisting) Then
                    entry.TargetPath = System.IO.Path.Combine(expandedExisting, entry.DefaultFileName)
                    entry.TargetPathForConfig = existingValue
                Else
                    entry.TargetPath = System.IO.Path.Combine(expandedDefaultDir, entry.DefaultFileName)
                    entry.TargetPathForConfig = entry.DefaultDir
                End If
            End If

            Return True

        Catch ex As System.Exception
            errors.Add("Could not determine target path for '" & entry.FriendlyName & "': " & ex.Message)
            Return False
        End Try

    End Function

    ''' <summary>
    ''' Determines whether a sample file entry requires a configuration update and populates ConfigValueToStore.
    ''' </summary>
    ''' <param name="entry">Entry to update.</param>
    ''' <param name="activeIniPath">Active main INI file path.</param>
    Private Shared Sub DetermineConfigUpdateNeeded(
        entry As SampleFileEntry,
        activeIniPath As System.String
    )

        Dim existingValue As System.String =
            TryReadIniKeyValue(activeIniPath, entry.ParameterForPath)

        Dim expandedExisting As System.String = Nothing
        If Not System.String.IsNullOrWhiteSpace(existingValue) Then
            Try
                expandedExisting = ExpandEnvironmentVariables(existingValue)
            Catch
                expandedExisting = existingValue
            End Try
        End If

        Dim expandedTarget As System.String = entry.TargetPath

        If entry.IsFullPath Then
            If System.String.IsNullOrWhiteSpace(expandedExisting) OrElse
               Not System.String.Equals(
                   NormalizePath(expandedExisting),
                   NormalizePath(expandedTarget),
                   System.StringComparison.OrdinalIgnoreCase
               ) Then

                entry.NeedsConfigUpdate = True
                entry.ConfigValueToStore = entry.TargetPathForConfig
            Else
                entry.NeedsConfigUpdate = False
            End If

        Else
            Dim expandedTargetDir As System.String =
                System.IO.Path.GetDirectoryName(expandedTarget)

            If System.String.IsNullOrWhiteSpace(expandedTargetDir) Then
                entry.NeedsConfigUpdate = False
                Return
            End If

            If System.String.IsNullOrWhiteSpace(expandedExisting) OrElse
               Not System.String.Equals(
                   NormalizePath(expandedExisting),
                   NormalizePath(expandedTargetDir),
                   System.StringComparison.OrdinalIgnoreCase
               ) Then

                entry.NeedsConfigUpdate = True
                entry.ConfigValueToStore = entry.TargetPathForConfig
            Else
                entry.NeedsConfigUpdate = False
            End If
        End If

    End Sub

    ''' <summary>
    ''' Normalizes a path for case-insensitive comparison and stable directory formatting.
    ''' </summary>
    ''' <param name="path">Input path.</param>
    ''' <returns>Normalized path.</returns>
    Private Shared Function NormalizePath(path As System.String) As System.String
        If System.String.IsNullOrWhiteSpace(path) Then Return path

        Try
            Return System.IO.Path.GetFullPath(path).TrimEnd("\"c)
        Catch
            Return path.TrimEnd("\"c)
        End Try
    End Function

    ' ENTRY POINTS

    ''' <summary>
    ''' Runs an interactive import limited to provider settings (excludes "OtherParameters").
    ''' </summary>
    ''' <param name="context">Shared context.</param>
    ''' <param name="ownerForm">Owner window for UI dialogs.</param>
    ''' <returns>True if the main INI file changed; otherwise False.</returns>
    Public Shared Function RunInteractiveImportProvidersOnly(
        context As ISharedContext,
        ownerForm As System.Windows.Forms.Form
    ) As Boolean

        Return RunImportInternal(
            context,
            ownerForm,
            ImportMode.InteractiveWithoutOtherParameters,
            Nothing
        )

    End Function

    ''' <summary>
    ''' Runs an interactive import limited to "OtherParameters".
    ''' </summary>
    ''' <param name="context">Shared context.</param>
    ''' <param name="ownerForm">Owner window for UI dialogs.</param>
    ''' <returns>True if the main INI file changed; otherwise False.</returns>
    Public Shared Function RunInteractiveImportOtherParameters(
        context As ISharedContext,
        ownerForm As System.Windows.Forms.Form
    ) As Boolean

        Return RunImportInternal(
            context,
            ownerForm,
            ImportMode.DirectOtherParameters,
            Nothing
        )

    End Function

    ''' <summary>
    ''' Runs the full interactive import flow where the user selects the import kind.
    ''' </summary>
    ''' <param name="context">Shared context.</param>
    ''' <param name="ownerForm">Owner window for UI dialogs.</param>
    ''' <returns>True if the main INI file changed; otherwise False.</returns>
    Public Shared Function RunImportFromVariableConfigurationWindow(
        context As ISharedContext,
        ownerForm As System.Windows.Forms.Form
    ) As Boolean
        Return RunImportInternal(context, ownerForm, ImportMode.InteractiveAllKinds, Nothing)
    End Function

    ''' <summary>
    ''' Runs a non-interactive import of "OtherParameters" from a provided URL/path, with optional overwrite and confirmation.
    ''' </summary>
    ''' <param name="context">Shared context.</param>
    ''' <param name="ownerForm">Owner window for UI dialogs.</param>
    ''' <param name="source">HTTPS URL or local/UNC path.</param>
    ''' <param name="replaceExisting">True to replace existing keys; False to only add missing keys.</param>
    ''' <param name="userConfirmation">True to prompt the user for overwrite behavior.</param>
    ''' <returns>True if the main INI file changed; otherwise False.</returns>
    Public Shared Function RunAutoImportOtherParameters(
        context As ISharedContext,
        ownerForm As System.Windows.Forms.Form,
        source As String,
        replaceExisting As Boolean,
        userConfirmation As Boolean
    ) As Boolean

        Dim opts As New AutoImportOptions With {
            .Source = source,
            .ReplaceExisting = replaceExisting,
            .UserConfirmation = userConfirmation
        }

        Return RunImportInternal(
            context,
            ownerForm,
            ImportMode.DirectOtherParameters,
            opts
        )
    End Function

    ''' <summary>
    ''' Implements the shared import flow used by all import entry points.
    ''' </summary>
    ''' <param name="context">Shared context.</param>
    ''' <param name="ownerForm">Owner window for UI dialogs.</param>
    ''' <param name="mode">Import interaction mode.</param>
    ''' <param name="autoOptions">Options for non-interactive runs; Nothing for interactive runs.</param>
    ''' <returns>True if the main INI file changed; otherwise False.</returns>
    Private Shared Function RunImportInternal(
        context As ISharedContext,
        ownerForm As System.Windows.Forms.Form,
        mode As ImportMode,
        autoOptions As AutoImportOptions
    ) As Boolean

        Dim mainIniChanged As System.Boolean = False
        Dim trustWarningShown As System.Boolean = False

        If context Is Nothing Then
            ShowCustomMessageBox("Internal error: context is missing.")
            Return False
        End If

        Dim activeIniPath As System.String = Nothing
        Try
            activeIniPath = GetActiveConfigFilePath(context)
        Catch ex As System.Exception
            ShowCustomMessageBox("Could not determine active configuration file path: " & ex.Message)
            Return False
        End Try

        If System.String.IsNullOrWhiteSpace(activeIniPath) Then
            ShowCustomMessageBox("No active configuration file path found.")
            Return False
        End If

        Dim disableReason As System.String = Nothing
        If Not CanUseImportFeature(context, activeIniPath, disableReason) Then
            ShowCustomMessageBox(disableReason)
            Return False
        End If

        If Not System.IO.File.Exists(activeIniPath) Then
            ShowCustomMessageBox("The main configuration file does not exist: " & activeIniPath)
            Return False
        End If

        Dim sourceText As System.String = Nothing
        Dim sourceLabel As System.String = Nothing

        If autoOptions IsNot Nothing AndAlso Not System.String.IsNullOrWhiteSpace(autoOptions.Source) Then
            If Not TryGetImportSourceTextFromProvidedSource(ownerForm, autoOptions.Source, sourceText, sourceLabel, trustWarningShown) Then
                Return False
            End If
        Else
            If Not TryGetImportSourceText(ownerForm, sourceText, sourceLabel, trustWarningShown) Then
                Return False
            End If
        End If

        If System.String.IsNullOrWhiteSpace(sourceText) Then
            ShowCustomMessageBox("No content loaded.")
            Return False
        End If

        If autoOptions Is Nothing Then
            Dim editedText As System.String = Nothing

            If Not ShowTextAsViewer(ownerForm, sourceText, "The following settings will be imported (press 'Save' to proceed)", editedText) OrElse String.IsNullOrEmpty(editedText) Then
                ShowCustomMessageBox("Import aborted.")
                Return False
            End If

            sourceText = editedText
        End If

        Dim kind As ImportKind

        Dim hostForm As System.Windows.Forms.Form = TryCast(ownerForm, System.Windows.Forms.Form)
        Dim wasTopMost As System.Boolean = False
        Dim hadHost As System.Boolean = (hostForm IsNot Nothing)

        If hadHost Then
            wasTopMost = hostForm.TopMost
            hostForm.TopMost = False
            hostForm.Enabled = False
            System.Windows.Forms.Application.DoEvents()
        End If

        Dim kindSelected As System.Boolean = True

        Try
            Select Case mode
                Case ImportMode.DirectOtherParameters
                    kind = ImportKind.OtherParameters

                Case ImportMode.InteractiveWithoutOtherParameters
                    If Not TryChooseImportKind(ownerForm, kind, True) Then
                        kindSelected = False
                    End If

                Case Else
                    If Not TryChooseImportKind(ownerForm, kind, False) Then
                        kindSelected = False
                    End If
            End Select
        Finally
            If hadHost Then
                hostForm.Enabled = True
                hostForm.TopMost = wasTopMost
                hostForm.Activate()
            End If
        End Try

        If Not kindSelected Then
            Return False
        End If

        Dim normalizedImportText As System.String = NormalizeImportTextForKind(sourceText, kind)

        Dim substitutedText As System.String = normalizedImportText
        Dim placeholderWarnings As New System.Collections.Generic.List(Of System.String)()

        ' Declare target path variables early
        Dim mainIniPath As System.String = activeIniPath
        Dim altIniPath As System.String = Nothing
        Dim svcIniPath As System.String = Nothing
        Dim targetIniPath As System.String = Nothing
        Dim targetSectionName As System.String = Nothing

        ' Determine targetIniPath and targetSectionName BEFORE placeholder resolution
        If kind = ImportKind.AlternateModel Then
            Dim pathUpdated As System.Boolean = False

            If Not TryEnsureSectionedIniPath(ownerForm,
                                             context,
                                             mainIniPath,
                                             isAlternate:=True,
                                             targetIniPath:=altIniPath,
                                             mainIniWasUpdated:=pathUpdated) Then
                Return False
            End If

            If pathUpdated Then
                mainIniChanged = True
            End If

            targetIniPath = altIniPath
            If Not TryGetSectionNameFromImportText(ownerForm, substitutedText, targetSectionName, "alternate model") Then
                Return False
            End If

        ElseIf kind = ImportKind.SpecialService Then
            Dim pathUpdated As System.Boolean = False

            If Not TryEnsureSectionedIniPath(ownerForm,
                                             context,
                                             mainIniPath,
                                             isAlternate:=False,
                                             targetIniPath:=svcIniPath,
                                             mainIniWasUpdated:=pathUpdated) Then
                Return False
            End If

            If pathUpdated Then
                mainIniChanged = True
            End If

            targetIniPath = svcIniPath
            If Not TryGetSectionNameFromImportText(ownerForm, substitutedText, targetSectionName, "special service") Then
                Return False
            End If

        Else
            targetIniPath = mainIniPath
        End If

        ' NOW look up existing placeholder values - targetIniPath is available
        Dim lookupSectionName As System.String = Nothing
        If kind = ImportKind.AlternateModel OrElse kind = ImportKind.SpecialService Then
            lookupSectionName = targetSectionName
        End If

        Dim existingPlaceholderValues As System.Collections.Generic.Dictionary(Of System.String, System.String) =
            ExtractExistingPlaceholderValues(GetLinesForPlaceholderLookup(targetIniPath, lookupSectionName))

        Dim resolvedPlaceholders As New System.Collections.Generic.List(Of ResolvedPlaceholder)()

        If Not TryResolvePlaceholders(
                ownerForm,
                substitutedText,
                placeholderWarnings,
                resolvedPlaceholders,
                existingPlaceholderValues
            ) Then
            Return False
        End If

        If placeholderWarnings.Count > 0 Then
            ShowCustomMessageBox(System.String.Join(System.Environment.NewLine & System.Environment.NewLine, placeholderWarnings))
        End If

        Dim importedLines As System.Collections.Generic.List(Of System.String) = SplitToLinesPreserveNonEmpty(substitutedText)

        If importedLines.Count = 0 Then
            ShowCustomMessageBox("The import content is empty after processing.")
            Return False
        End If

        Dim sourceUri As System.Uri = Nothing
        If Not System.String.IsNullOrWhiteSpace(sourceLabel) Then
            System.Uri.TryCreate(sourceLabel, System.UriKind.Absolute, sourceUri)
        End If

        If sourceUri IsNot Nothing AndAlso sourceUri.Scheme.Equals("https", System.StringComparison.OrdinalIgnoreCase) Then

            Dim host As System.String = sourceUri.Host

            Dim isFullyTrusted As System.Boolean =
                System.Linq.Enumerable.Any(TRUSTED_HOSTS_FOR_GETSETTINGS,
                                          Function(h) System.String.Equals(h, host, System.StringComparison.OrdinalIgnoreCase))

            If Not isFullyTrusted Then

                Dim allowedMarkers As System.String() = Nothing

                If RESTRICTED_TRUSTED_HOSTS_FOR_GETSETTINGS.TryGetValue(host, allowedMarkers) Then

                    Dim needWarning As System.Boolean = False
                    Dim warningText As New System.Text.StringBuilder()

                    Dim trustedMarkerText As System.String =
                        If(allowedMarkers IsNot Nothing AndAlso allowedMarkers.Length > 0,
                           System.String.Join(", ",
                               System.Linq.Enumerable.Select(
                                   allowedMarkers,
                                   Function(m As System.String) "'" & m & "'"
                               )
                           ),
                           "(no trusted section markers configured)"
                        )

                    If kind <> ImportKind.AlternateModel AndAlso kind <> ImportKind.SpecialService Then
                        needWarning = True
                        warningText.AppendLine("Warning: The configuration source '" & host & "' is a host with restricted trust.")
                        warningText.AppendLine()
                        warningText.AppendLine("You are about to import settings into the main configuration file, but this source is trusted only for alternate AI models or Special Services.")
                        warningText.AppendLine()
                        warningText.AppendLine("Proceed anyway?")
                    Else
                        Dim segments = ParseIniSegments(importedLines)

                        If segments Is Nothing OrElse segments.Count = 0 Then
                            needWarning = True
                            warningText.AppendLine("Warning: The configuration source '" & host & "' is a host with restricted trust.")
                            warningText.AppendLine()
                            warningText.AppendLine($"You are about to import settings for an AI model or Special Service for which this source has not been trusted (this host is only trusted for markers {trustedMarkerText} in the Section names).")
                            warningText.AppendLine()
                            warningText.AppendLine("Proceed anyway?")
                        Else
                            Dim invalidSections As New System.Collections.Generic.List(Of System.String)()

                            Dim markersOk As System.Boolean = (allowedMarkers IsNot Nothing AndAlso allowedMarkers.Length > 0)

                            For Each sectionName As System.String In segments.Keys

                                Dim sectionAllowed As System.Boolean = False

                                If markersOk Then
                                    sectionAllowed =
                                        System.Linq.Enumerable.Any(allowedMarkers,
                                            Function(marker As System.String) As System.Boolean
                                                If System.String.IsNullOrWhiteSpace(marker) Then Return False
                                                Return sectionName.IndexOf(marker, System.StringComparison.OrdinalIgnoreCase) >= 0
                                            End Function
                                        )
                                End If

                                If Not sectionAllowed Then
                                    invalidSections.Add(sectionName)
                                End If
                            Next

                            If invalidSections.Count > 0 Then
                                needWarning = True
                                warningText.AppendLine("Warning: The configuration source '" & host & "' is a host with restricted trust.")
                                warningText.AppendLine()
                                warningText.AppendLine($"For this host, each [Section] name must contain at least one trusted marker ({trustedMarkerText}).")
                                warningText.AppendLine()
                                warningText.AppendLine("The following section names did not match:")
                                warningText.AppendLine()

                                For Each s As System.String In invalidSections
                                    warningText.AppendLine("  - " & s)
                                Next

                                warningText.AppendLine()
                                warningText.AppendLine("Proceed anyway?")
                            End If
                        End If
                    End If

                    If needWarning AndAlso Not trustWarningShown Then
                        trustWarningShown = True

                        Dim decisionTrust As System.Int32 =
                            ShowCustomYesNoBox(warningText.ToString(), "Yes, continue", "No, abort import")

                        If decisionTrust <> 1 Then
                            Return False
                        End If
                    End If

                End If
            End If

        End If

        If autoOptions IsNot Nothing AndAlso autoOptions.UserConfirmation Then
            Dim msg As System.String =
                "How should existing keys be handled?" & System.Environment.NewLine & System.Environment.NewLine &
                "Yes = Replace existing keys" & System.Environment.NewLine &
                "No = Only add missing keys" & System.Environment.NewLine &
                "Cancel = Abort import"

            Dim decisionAuto As System.Int32 = ShowCustomYesNoBox(msg, "Replace", "Add only")
            If decisionAuto = 0 Then Return False

            autoOptions.ReplaceExisting = (decisionAuto = 1)
        End If

        Dim plans As New System.Collections.Generic.List(Of DryRunPlan)
        Dim allRemovedLines As New System.Collections.Generic.List(Of System.String)

        Try
            If kind = ImportKind.AlternateModel OrElse kind = ImportKind.SpecialService Then

                Dim segments = ParseIniSegments(importedLines)

                If segments.Count = 0 Then
                    Throw New System.Exception("No valid sections found.")
                End If

                Dim combined As DryRunPlan = BuildSectionedCombinedPlan(context, kind, targetIniPath, segments)

                If combined.WillCreateRemovedBackup Then
                    allRemovedLines.AddRange(combined.RemovedLinesBackup)
                    combined.WillCreateRemovedBackup = False
                End If

                plans.Add(combined)

            Else
                Dim replaceExisting As System.Boolean = True
                If autoOptions IsNot Nothing Then
                    replaceExisting = autoOptions.ReplaceExisting
                End If

                Dim singlePlan As DryRunPlan =
                    BuildDryRunPlan(context, kind, targetIniPath, targetSectionName, importedLines, replaceExisting)

                If singlePlan.WillCreateRemovedBackup Then
                    allRemovedLines.AddRange(singlePlan.RemovedLinesBackup)
                    singlePlan.WillCreateRemovedBackup = False
                End If

                plans.Add(singlePlan)
            End If

        Catch ex As System.Exception
            ShowCustomMessageBox("Could not build dry run plan: " & ex.Message)
            Return False
        End Try

        If plans.Count = 0 Then
            ShowCustomMessageBox("Nothing to import.")
            Return False
        End If

        Dim sb As New System.Text.StringBuilder()

        sb.AppendLine("Dry run – review before importing")
        sb.AppendLine()

        For Each p As DryRunPlan In plans
            sb.AppendLine("Target file: " & p.TargetIniPath)

            If p.Kind = ImportKind.AlternateModel OrElse p.Kind = ImportKind.SpecialService Then

                Dim names As System.Collections.Generic.List(Of String) =
                    If(p.SectionNames, New System.Collections.Generic.List(Of String)())

                For Each sectionName As String In names
                    sb.AppendLine("Section: [" & sectionName & "]")

                    Dim added As System.Collections.Generic.List(Of String) = Nothing
                    If p.AddedKeysBySection IsNot Nothing Then p.AddedKeysBySection.TryGetValue(sectionName, added)

                    Dim overwritten As System.Collections.Generic.List(Of String) = Nothing
                    If p.OverwrittenKeysBySection IsNot Nothing Then p.OverwrittenKeysBySection.TryGetValue(sectionName, overwritten)

                    Dim addedCount As Integer = If(added Is Nothing, 0, added.Count)
                    Dim overwrittenCount As Integer = If(overwritten Is Nothing, 0, overwritten.Count)

                    If addedCount > 0 Then
                        sb.AppendLine("Keys that will be added (" & addedCount.ToString() & "): " & System.String.Join(", ", added))
                    Else
                        sb.AppendLine("Keys that will be added (0): (none)")
                    End If

                    If overwrittenCount > 0 Then
                        sb.AppendLine("Keys that will be overwritten (" & overwrittenCount.ToString() & "): " & System.String.Join(", ", overwritten))
                    Else
                        sb.AppendLine("Keys that will be overwritten (0): (none)")
                    End If

                    sb.AppendLine()
                Next

                sb.AppendLine()
            Else
                If Not System.String.IsNullOrWhiteSpace(p.TargetSectionName) Then
                    sb.AppendLine("Section: [" & p.TargetSectionName & "]")
                End If

                If p.OverwrittenKeys IsNot Nothing AndAlso p.OverwrittenKeys.Count > 0 Then
                    sb.AppendLine(
                        "Keys that will be overwritten (" &
                        p.OverwrittenKeys.Count.ToString() &
                        "): " &
                        System.String.Join(", ", p.OverwrittenKeys)
                    )
                Else
                    If p.Kind = ImportKind.PrimaryModel OrElse
                       p.Kind = ImportKind.SecondaryModel OrElse
                       p.Kind = ImportKind.OtherParameters Then

                        sb.AppendLine("No existing keys will be overwritten (new keys will be appended at the end of the file).")
                    End If
                End If

                sb.AppendLine()
            End If
        Next

        If allRemovedLines IsNot Nothing AndAlso allRemovedLines.Count > 0 Then
            sb.AppendLine("A full backup of the target file will be created, and all removed content will be stored in one backup file.")
        Else
            sb.AppendLine("A full backup of the target file will be created.")
        End If

        sb.AppendLine()
        sb.AppendLine("Proceed with import?")

        If autoOptions Is Nothing Then
            Dim decision As System.Int32 = ShowCustomYesNoBox(sb.ToString(), "Yes, continue", "No, abort import")
            If decision <> 1 Then Return False
        End If

        If resolvedPlaceholders.Count > 0 Then
            For Each p As DryRunPlan In plans
                InjectPlaceholderCommentsIntoPlan(p, resolvedPlaceholders)
            Next
        End If

        For Each p As DryRunPlan In plans
            CommitDryRunPlan(p)
        Next

        For Each p As DryRunPlan In plans
            If System.String.Equals(p.TargetIniPath,
                                    mainIniPath,
                                    System.StringComparison.OrdinalIgnoreCase) Then
                mainIniChanged = True
                Exit For
            End If
        Next

        If allRemovedLines.Count > 0 Then
            Dim ts As System.String = System.DateTime.Now.ToString("yyyyMMdd_HHmmss")
            Dim baseName As System.String = System.IO.Path.GetFileNameWithoutExtension(targetIniPath)
            Dim baseExt As System.String = System.IO.Path.GetExtension(targetIniPath)
            Dim dir As System.String = System.IO.Path.GetDirectoryName(targetIniPath)

            Dim removedBackup As System.String =
                System.IO.Path.Combine(dir, baseName & baseExt & "_removed_" & ts & ".bak")

            System.IO.File.WriteAllText(
                removedBackup,
                System.String.Join(System.Environment.NewLine, allRemovedLines),
                System.Text.Encoding.UTF8
            )
        End If

        If Not mainIniChanged Then ShowCustomMessageBox("Import completed. No reloading or restarting required.")

        Return mainIniChanged

    End Function


    ' =========================================================================================
    '  FEATURE ENABLEMENT RULES
    ' =========================================================================================

    ''' <summary>
    ''' Determines whether the import feature can be used for the current configuration context and active INI path.
    ''' </summary>
    ''' <param name="context">Shared context.</param>
    ''' <param name="activeIniPath">Active INI path resolved at runtime.</param>
    ''' <param name="disableReason">If disabled, contains a user-friendly reason.</param>
    ''' <returns>True if import is allowed; otherwise False.</returns>
    Public Shared Function CanUseImportFeature(context As ISharedContext,
                                               activeIniPath As System.String,
                                               ByRef disableReason As System.String) As System.Boolean

        disableReason = Nothing

        If context Is Nothing Then
            disableReason = "Import is not available (missing context)."
            Return False
        End If

        If System.String.IsNullOrWhiteSpace(activeIniPath) Then
            disableReason = "Import is not available (no active .ini path)."
            Return False
        End If

        Try
            If RegPath_IniPrio Then
                disableReason = "Import is not available when the configuration is controlled via registry/network setup."
                Return False
            End If
        Catch
        End Try

        Dim rdv As System.String = Nothing
        Try
            rdv = context.RDV
        Catch
        End Try

        If Not System.String.IsNullOrWhiteSpace(rdv) Then
            Dim defaultPathThisApp As System.String = Nothing
            Dim defaultPathWord As System.String = Nothing

            Try
                defaultPathThisApp = GetDefaultINIPath(rdv)
            Catch
            End Try

            Try
                defaultPathWord = GetDefaultINIPath("Word")
            Catch
            End Try

            If (System.String.Equals(rdv, "Excel", System.StringComparison.OrdinalIgnoreCase) OrElse
                System.String.Equals(rdv, "Outlook", System.StringComparison.OrdinalIgnoreCase)) AndAlso
               Not System.String.IsNullOrWhiteSpace(defaultPathWord) AndAlso
               System.IO.File.Exists(defaultPathWord) AndAlso
               System.String.Equals(activeIniPath, defaultPathWord, System.StringComparison.OrdinalIgnoreCase) Then

                disableReason = "Import is not available here because this application is using Word's configuration file. Please use Word to import settings."
                Return False
            End If

            If Not System.String.IsNullOrWhiteSpace(defaultPathThisApp) AndAlso
               System.IO.File.Exists(defaultPathThisApp) Then

                If Not System.String.Equals(activeIniPath, defaultPathThisApp, System.StringComparison.OrdinalIgnoreCase) Then
                    disableReason = "Import is only available when using the local per-application configuration file."
                    Return False
                End If
            End If
        End If

        Return True

    End Function

    ' =========================================================================================
    '  SOURCE ACQUISITION (URL or FILE)
    ' =========================================================================================

    ''' <summary>
    ''' Prompts the user for an import source (HTTPS URL or local/UNC path) and loads the source text.
    ''' </summary>
    ''' <param name="ownerForm">Owner window for UI dialogs.</param>
    ''' <param name="sourceText">Loaded source text.</param>
    ''' <param name="sourceLabel">URL or path label used for reporting and trust logic.</param>
    ''' <param name="trustWarningShown">Tracks whether a trust warning has already been shown for the current run.</param>
    ''' <returns>True if text was loaded; otherwise False.</returns>
    Private Shared Function TryGetImportSourceText(
        ownerForm As System.Windows.Forms.Form,
        ByRef sourceText As System.String,
        ByRef sourceLabel As System.String,
        ByRef trustWarningShown As System.Boolean
    ) As System.Boolean

        sourceText = Nothing
        sourceLabel = Nothing

        Dim wasTopMost As Boolean = False
        Dim hadOwner As Boolean = (ownerForm IsNot Nothing)
        Dim input As System.String

        If hadOwner Then
            wasTopMost = ownerForm.TopMost
            ownerForm.TopMost = False
            ownerForm.Enabled = False
            System.Windows.Forms.Application.DoEvents()
        End If

        Try
            input = ShowCustomInputBox("Enter the source URL (https://...) or file / UNC path:", TITLE_IMPORT, True)
        Finally
            If hadOwner Then
                ownerForm.Enabled = True
                ownerForm.TopMost = wasTopMost
                ownerForm.Activate()
            End If
        End Try

        If System.String.IsNullOrWhiteSpace(input) Then
            ShowCustomMessageBox("No source provided.")
            Return False
        End If

        input = input.Trim()

        If input.StartsWith("https://", System.StringComparison.OrdinalIgnoreCase) Then

            Dim u As System.Uri = Nothing
            Try
                u = New System.Uri(input)
            Catch
                ShowCustomMessageBox("Invalid URL.")
                Return False
            End Try

            Dim host As System.String = u.Host

            Dim isFullyTrusted As System.Boolean =
                System.Linq.Enumerable.Any(TRUSTED_HOSTS_FOR_GETSETTINGS,
                                          Function(h) System.String.Equals(h, host, System.StringComparison.OrdinalIgnoreCase))

            Dim isRestrictedTrusted As System.Boolean = RESTRICTED_TRUSTED_HOSTS_FOR_GETSETTINGS.ContainsKey(host)

            Dim isAllowedWithoutWarning As System.Boolean = (isFullyTrusted OrElse isRestrictedTrusted)

            If Not isAllowedWithoutWarning AndAlso Not trustWarningShown Then
                trustWarningShown = True

                Dim warnDecision As System.Int32 = ShowCustomYesNoBox(
                    "Warning: This URL host is not on the built-in trust list:" & System.Environment.NewLine &
                    host & System.Environment.NewLine & System.Environment.NewLine &
                    "Importing configuration from unknown hosts can be dangerous." & System.Environment.NewLine & System.Environment.NewLine &
                    "Do you want to continue?",
                    "Yes, continue",
                    "No, abort import"
                )

                If warnDecision <> 1 Then
                    Return False
                End If
            End If

            Dim downloaded As System.String = Nothing
            Try
                downloaded = DownloadHttpsTextWithLimit(u, MAX_DOWNLOAD_BYTES)
            Catch ex As System.Exception
                ShowCustomMessageBox("Download failed: " & ex.Message)
                Return False
            End Try

            sourceText = downloaded
            sourceLabel = u.ToString()
            Return True

        Else
            Dim path As System.String = input

            Try
                path = ExpandEnvironmentVariables(path)
            Catch
            End Try

            If Not System.IO.File.Exists(path) Then
                ShowCustomMessageBox("File not found: " & path)
                Return False
            End If

            Try
                Dim fi As New System.IO.FileInfo(path)
                If fi.Length > MAX_DOWNLOAD_BYTES Then
                    ShowCustomMessageBox("The file is larger than the allowed limit (" & MAX_DOWNLOAD_BYTES.ToString() & " bytes).")
                    Return False
                End If
            Catch
            End Try

            Try
                sourceText = System.IO.File.ReadAllText(path, System.Text.Encoding.UTF8)
            Catch ex As System.Exception
                ShowCustomMessageBox("Could not read file: " & ex.Message)
                Return False
            End Try

            sourceLabel = path
            Return True
        End If

    End Function

    ''' <summary>
    ''' Checks whether a host is included in the trusted-host list.
    ''' </summary>
    ''' <param name="host">Host name.</param>
    ''' <returns>True if the host is trusted; otherwise False.</returns>
    Private Shared Function IsHostAllowed(host As System.String) As System.Boolean
        If System.String.IsNullOrWhiteSpace(host) Then Return False
        For Each h As System.String In TRUSTED_HOSTS_FOR_GETSETTINGS
            If System.String.Equals(host, h, System.StringComparison.OrdinalIgnoreCase) Then Return True
        Next
        Return False
    End Function

    ''' <summary>
    ''' Downloads UTF-8 text from an HTTPS URL with a hard maximum size limit.
    ''' </summary>
    ''' <param name="u">HTTPS URL.</param>
    ''' <param name="maxBytes">Maximum allowed download size.</param>
    ''' <returns>Downloaded text.</returns>
    Private Shared Function DownloadHttpsTextWithLimit(u As System.Uri, maxBytes As System.Int32) As System.String

        If u Is Nothing Then Throw New System.ArgumentNullException(NameOf(u))
        If Not System.String.Equals(u.Scheme, "https", System.StringComparison.OrdinalIgnoreCase) Then
            Throw New System.Exception("Only HTTPS is allowed.")
        End If

        Try
            System.Net.ServicePointManager.SecurityProtocol =
                System.Net.SecurityProtocolType.Tls12 Or
                CType(0, System.Net.SecurityProtocolType)
        Catch
        End Try

        Dim req As System.Net.HttpWebRequest = CType(System.Net.WebRequest.Create(u), System.Net.HttpWebRequest)
        req.Method = "GET"
        req.AllowAutoRedirect = True
        req.UserAgent = "RedInk-Importer"
        req.Timeout = 15000
        req.ReadWriteTimeout = 15000

        Using resp As System.Net.HttpWebResponse = CType(req.GetResponse(), System.Net.HttpWebResponse)
            Using s As System.IO.Stream = resp.GetResponseStream()
                If s Is Nothing Then Throw New System.Exception("No response stream.")
                Using ms As New System.IO.MemoryStream()
                    Dim buffer(4095) As System.Byte
                    Dim total As System.Int32 = 0
                    While True
                        Dim read As System.Int32 = s.Read(buffer, 0, buffer.Length)
                        If read <= 0 Then Exit While
                        total += read
                        If total > maxBytes Then
                            Throw New System.Exception("Download exceeds the maximum allowed size (" & maxBytes.ToString() & " bytes).")
                        End If
                        ms.Write(buffer, 0, read)
                    End While
                    Dim data As System.Byte() = ms.ToArray()
                    Return System.Text.Encoding.UTF8.GetString(data)
                End Using
            End Using
        End Using

    End Function

    ''' <summary>
    ''' Loads import source text from a provided HTTPS URL or local/UNC path.
    ''' </summary>
    ''' <param name="ownerForm">Owner window for UI dialogs.</param>
    ''' <param name="providedSource">HTTPS URL or local/UNC path.</param>
    ''' <param name="sourceText">Loaded source text.</param>
    ''' <param name="sourceLabel">URL or path label used for reporting.</param>
    ''' <param name="trustWarningShown">Tracks whether a trust warning has already been shown for the current run.</param>
    ''' <returns>True if text was loaded; otherwise False.</returns>
    Private Shared Function TryGetImportSourceTextFromProvidedSource(
        ownerForm As System.Windows.Forms.Form,
        providedSource As System.String,
        ByRef sourceText As System.String,
        ByRef sourceLabel As System.String,
        ByRef trustWarningShown As System.Boolean
    ) As System.Boolean

        sourceText = Nothing
        sourceLabel = Nothing

        If System.String.IsNullOrWhiteSpace(providedSource) Then
            ShowCustomMessageBox("No source provided.")
            Return False
        End If

        Dim input As System.String = providedSource.Trim()

        If input.StartsWith("https://", System.StringComparison.OrdinalIgnoreCase) Then

            Dim u As System.Uri = Nothing
            Try
                u = New System.Uri(input)
            Catch ex As System.Exception
                ShowCustomMessageBox("Invalid URL: " & ex.Message)
                Return False
            End Try

            Dim host As System.String = u.Host

            Dim isFullyTrusted As System.Boolean =
                System.Linq.Enumerable.Any(TRUSTED_HOSTS_FOR_GETSETTINGS,
                                          Function(h) System.String.Equals(h, host, System.StringComparison.OrdinalIgnoreCase))

            Dim isRestrictedTrusted As System.Boolean = RESTRICTED_TRUSTED_HOSTS_FOR_GETSETTINGS.ContainsKey(host)

            Dim isAllowedWithoutWarning As System.Boolean = (isFullyTrusted OrElse isRestrictedTrusted)

            If Not isAllowedWithoutWarning AndAlso Not trustWarningShown Then
                trustWarningShown = True

                Dim warnDecision As System.Int32 = ShowCustomYesNoBox(
                    "Warning: This URL host is not on the built-in trust list:" & System.Environment.NewLine &
                    host & System.Environment.NewLine & System.Environment.NewLine &
                    "Importing configuration from unknown hosts can be dangerous." & System.Environment.NewLine & System.Environment.NewLine &
                    "Do you want to continue?",
                    "Yes, continue",
                    "No, abort import"
                )
                If warnDecision <> 1 Then
                    Return False
                End If
            End If

            Try
                sourceText = DownloadHttpsTextWithLimit(u, MAX_DOWNLOAD_BYTES)
            Catch ex As System.Exception
                ShowCustomMessageBox("Download failed: " & ex.Message)
                Return False
            End Try

            sourceLabel = u.ToString()
            Return True

        Else
            Dim path As System.String = input

            Try
                path = ExpandEnvironmentVariables(path)
            Catch
            End Try

            If Not System.IO.File.Exists(path) Then
                ShowCustomMessageBox("File not found: " & path)
                Return False
            End If

            Try
                Dim fi As New System.IO.FileInfo(path)
                If fi.Length > MAX_DOWNLOAD_BYTES Then
                    ShowCustomMessageBox("The file is larger than the allowed limit (" & MAX_DOWNLOAD_BYTES.ToString() & " bytes).")
                    Return False
                End If
            Catch
            End Try

            Try
                sourceText = System.IO.File.ReadAllText(path, System.Text.Encoding.UTF8)
            Catch ex As System.Exception
                ShowCustomMessageBox("Could not read file: " & ex.Message)
                Return False
            End Try

            sourceLabel = path
            Return True
        End If

    End Function

    ''' <summary>
    ''' Shows a text file preview/editor and returns updated text only when the user saves.
    ''' </summary>
    ''' <param name="ownerForm">Owner window.</param>
    ''' <param name="text">Initial text.</param>
    ''' <param name="title">Viewer title.</param>
    ''' <param name="finalText">Final text after save; Nothing if not saved.</param>
    ''' <returns>True if user saved; otherwise False.</returns>
    Private Shared Function ShowTextAsViewer(
        ownerForm As System.Windows.Forms.Form,
        text As System.String,
        title As System.String,
        ByRef finalText As System.String
    ) As Boolean

        Dim normalized As String = text
        If normalized IsNot Nothing Then
            normalized = normalized.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Replace(vbLf, vbCrLf)
        End If

        Dim tmp As System.String =
            System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                "RedInk_ImportPreview_" & System.Guid.NewGuid().ToString("N") & ".txt"
            )

        System.IO.File.WriteAllText(tmp, normalized, System.Text.Encoding.UTF8)

        Dim wasTopMost As System.Boolean = False
        Dim hadOwner As System.Boolean = (ownerForm IsNot Nothing)

        If hadOwner Then
            wasTopMost = ownerForm.TopMost
            ownerForm.TopMost = False
            ownerForm.Enabled = False
            System.Windows.Forms.Application.DoEvents()
        End If

        Try
            Dim wasSaved As System.Nullable(Of System.Boolean) = Nothing

            ShowTextFileEditor(
                tmp,
                title,
                False,
                Nothing,
                wasSaved
            )

            If wasSaved.HasValue AndAlso wasSaved.Value Then
                finalText = System.IO.File.ReadAllText(tmp, System.Text.Encoding.UTF8)
                Return True
            End If

            finalText = Nothing
            Return False

        Finally
            If hadOwner Then
                ownerForm.Enabled = True
                ownerForm.TopMost = wasTopMost
                ownerForm.Activate()
            End If

            Try
                System.IO.File.Delete(tmp)
            Catch
            End Try
        End Try

    End Function

    ' =========================================================================================
    '  IMPORT KIND SELECTION
    ' =========================================================================================

    ''' <summary>
    ''' Prompts the user to choose the kind of settings to import.
    ''' </summary>
    ''' <param name="ownerForm">Owner window.</param>
    ''' <param name="kind">Selected import kind.</param>
    ''' <param name="excludeOtherParameters">If True, hides the "OtherParameters" option.</param>
    ''' <returns>True if a choice was made; otherwise False.</returns>
    Private Shared Function TryChooseImportKind(
        ownerForm As System.Windows.Forms.Form,
        ByRef kind As ImportKind,
        excludeOtherParameters As System.Boolean
    ) As System.Boolean

        kind = ImportKind.PrimaryModel

        Dim options As New List(Of String) From {
            "For the primary model",
            "For the secondary model",
            "For an alternate model",
            "For a special service"
        }

        If Not excludeOtherParameters Then
            options.Add("For other parameters")
        End If

        Dim choice As String =
            ShowSelectionForm("Which settings do you want to import?", TITLE_IMPORT, options)

        If String.IsNullOrWhiteSpace(choice) Then Return False

        If choice.StartsWith("For the primary", StringComparison.OrdinalIgnoreCase) Then
            kind = ImportKind.PrimaryModel
        ElseIf choice.StartsWith("For the secondary", StringComparison.OrdinalIgnoreCase) Then
            kind = ImportKind.SecondaryModel
        ElseIf choice.StartsWith("For an alternate", StringComparison.OrdinalIgnoreCase) Then
            kind = ImportKind.AlternateModel
        ElseIf choice.StartsWith("For a special service", StringComparison.OrdinalIgnoreCase) Then
            kind = ImportKind.SpecialService
        Else
            kind = ImportKind.OtherParameters
        End If

        Return True
    End Function

    ''' <summary>
    ''' Normalizes import text for kinds that import into the main INI by removing section headers.
    ''' </summary>
    ''' <param name="sourceText">Original import text.</param>
    ''' <param name="kind">Import kind.</param>
    ''' <returns>Normalized import text.</returns>
    Private Shared Function NormalizeImportTextForKind(sourceText As System.String, kind As ImportKind) As System.String
        If System.String.IsNullOrWhiteSpace(sourceText) Then Return ""

        If kind = ImportKind.PrimaryModel OrElse kind = ImportKind.SecondaryModel OrElse kind = ImportKind.OtherParameters Then
            Dim lines As System.Collections.Generic.List(Of System.String) = SplitToLinesPreserve(sourceText)
            Dim kept As New System.Collections.Generic.List(Of System.String)()

            For Each line As System.String In lines
                Dim t As System.String = line.Trim()
                If t.StartsWith("[") AndAlso t.EndsWith("]") AndAlso t.Length >= 2 Then
                    Continue For
                End If
                kept.Add(line)
            Next

            Return System.String.Join(System.Environment.NewLine, kept)
        End If

        Return sourceText
    End Function

    ' =========================================================================================
    '  PLACEHOLDERS [[...]] -> prompt user
    ' =========================================================================================

    ''' <summary>
    ''' Prompts the user for values for each [[...]] placeholder found in the input text and
    ''' substitutes non-empty entries directly into the text.
    ''' </summary>
    ''' <param name="ownerForm">Owner window used for modal UI prompts.</param>
    ''' <param name="text">
    ''' Input/output text containing placeholders; updated in-place with resolved values.
    ''' </param>
    ''' <param name="warnings">
    ''' Collector for warnings about placeholders that were left unresolved and remain in the text.
    ''' </param>
    ''' <param name="resolved">
    ''' Collector for placeholders that were successfully resolved, including the placeholder
    ''' name and the user-provided value. This list is later used to annotate the output INI file.
    ''' </param>
    ''' <returns>
    ''' True if placeholder processing completed and import should continue; otherwise False
    ''' if the user aborted the placeholder input dialog.
    ''' </returns>
    Private Shared Function TryResolvePlaceholders(
    ownerForm As System.Windows.Forms.Form,
    ByRef text As System.String,
    warnings As System.Collections.Generic.List(Of System.String),
    ByRef resolved As System.Collections.Generic.List(Of ResolvedPlaceholder),
    Optional existingValues As System.Collections.Generic.Dictionary(Of System.String, System.String) = Nothing
) As System.Boolean

        If resolved Is Nothing Then
            resolved = New System.Collections.Generic.List(Of ResolvedPlaceholder)()
        End If

        If warnings Is Nothing Then warnings = New System.Collections.Generic.List(Of System.String)()
        If System.String.IsNullOrWhiteSpace(text) Then Return True

        Dim rx As New System.Text.RegularExpressions.Regex("\[\[(.+?)\]\]",
                                                      System.Text.RegularExpressions.RegexOptions.Singleline)

        Dim matches As System.Text.RegularExpressions.MatchCollection = rx.Matches(text)
        If matches Is Nothing OrElse matches.Count = 0 Then Return True

        Dim unique As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)
        For Each m As System.Text.RegularExpressions.Match In matches
            If m Is Nothing OrElse Not m.Success Then Continue For
            Dim key As System.String = m.Groups(1).Value
            If System.String.IsNullOrWhiteSpace(key) Then Continue For
            If Not unique.ContainsKey(key) Then unique.Add(key, "")
        Next

        If unique.Count = 0 Then Return True

        Dim paramList As New System.Collections.Generic.List(Of InputParameter)()
        For Each k As System.String In unique.Keys
            ' Pre-fill with existing value if available
            Dim defaultValue As System.String = ""
            If existingValues IsNot Nothing AndAlso existingValues.ContainsKey(k) Then
                defaultValue = existingValues(k)
            End If
            paramList.Add(New InputParameter(k, defaultValue))
        Next

        Dim params() As InputParameter = paramList.ToArray()

        Dim wasTopMost As Boolean = False
        Dim hadOwner As Boolean = (ownerForm IsNot Nothing)

        If hadOwner Then
            wasTopMost = ownerForm.TopMost
            ownerForm.TopMost = False
            ownerForm.Enabled = False
            System.Windows.Forms.Application.DoEvents()
        End If

        Try
            If ShowCustomVariableInputForm("The settings require your to enter individual values. Please enter them (leave empty to keep a placeholder and edit later): ",
                                       TITLE_IMPORT,
                                       params) = False Then
                Return False
            End If

        Finally
            If hadOwner Then
                ownerForm.Enabled = True
                ownerForm.TopMost = wasTopMost
                ownerForm.Activate()
            End If
        End Try

        For Each p As InputParameter In params
            Dim name As System.String = System.Convert.ToString(p.Name)
            Dim value As System.String = System.Convert.ToString(p.Value)

            If System.String.IsNullOrWhiteSpace(name) Then Continue For

            If Not System.String.IsNullOrWhiteSpace(value) Then
                Dim keyRx As New System.Text.RegularExpressions.Regex("\[\[" & System.Text.RegularExpressions.Regex.Escape(name) & "\]\]",
                                                                 System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                text = keyRx.Replace(text, value)
                resolved.Add(New ResolvedPlaceholder With {
                    .Name = name,
                    .Value = value
                })

            Else
                warnings.Add("Warning: Placeholder '[[ " & name & " ]]' was left empty and remains in the configuration." & System.Environment.NewLine &
                         "You can later fill it using the 'Edit .ini Files' feature or directly access the file.")
            End If
        Next

        Return True
    End Function

    ''' <summary>
    ''' Extracts previously stored placeholder values from comment lines in the format "; [[name]] = value".
    ''' </summary>
    ''' <param name="lines">Lines to search (either full file or section lines).</param>
    ''' <returns>Dictionary mapping placeholder names to their stored values.</returns>
    Private Shared Function ExtractExistingPlaceholderValues(
    lines As System.Collections.Generic.List(Of System.String)
) As System.Collections.Generic.Dictionary(Of System.String, System.String)

        Dim result As New System.Collections.Generic.Dictionary(Of System.String, System.String)(
        System.StringComparer.OrdinalIgnoreCase)

        If lines Is Nothing Then Return result

        Dim rx As New System.Text.RegularExpressions.Regex(
        "^\s*;\s*\[\[(.+?)\]\]\s*=\s*(.*)$",
        System.Text.RegularExpressions.RegexOptions.IgnoreCase)

        For Each line As System.String In lines
            If line Is Nothing Then Continue For

            Dim m As System.Text.RegularExpressions.Match = rx.Match(line)
            If m.Success AndAlso m.Groups.Count >= 3 Then
                Dim name As System.String = m.Groups(1).Value.Trim()
                Dim value As System.String = m.Groups(2).Value.Trim()

                If Not System.String.IsNullOrWhiteSpace(name) Then
                    result(name) = value
                End If
            End If
        Next

        Return result
    End Function

    ''' <summary>
    ''' Gets lines for a specific section from the target INI file, or all lines for global imports.
    ''' </summary>
    ''' <param name="targetIniPath">Path to the target INI file.</param>
    ''' <param name="sectionName">Section name for sectioned imports; Nothing for global imports.</param>
    ''' <returns>Relevant lines to search for existing placeholders.</returns>
    Private Shared Function GetLinesForPlaceholderLookup(
    targetIniPath As System.String,
    sectionName As System.String
) As System.Collections.Generic.List(Of System.String)

        If System.String.IsNullOrWhiteSpace(targetIniPath) OrElse
       Not System.IO.File.Exists(targetIniPath) Then
            Return New System.Collections.Generic.List(Of System.String)()
        End If

        Dim allLines As System.Collections.Generic.List(Of System.String) = ReadAllLinesPreserve(targetIniPath)

        If System.String.IsNullOrWhiteSpace(sectionName) Then
            ' Global import: return all lines
            Return allLines
        End If

        ' Sectioned import: return only lines within the target section
        Dim startIndex As System.Int32 = -1
        Dim endIndex As System.Int32 = -1
        FindSectionRange(allLines, sectionName, startIndex, endIndex)

        If startIndex < 0 Then
            Return New System.Collections.Generic.List(Of System.String)()
        End If

        Dim sectionLines As New System.Collections.Generic.List(Of System.String)()
        For i As System.Int32 = startIndex To endIndex
            sectionLines.Add(allLines(i))
        Next

        Return sectionLines
    End Function


    ' =========================================================================================
    '  ENSURE AlternateModelPath / SpecialServicePath exist in redink.ini
    ' =========================================================================================

    ''' <summary>
    ''' Resolves the sectioned INI path for alternate models or special services and ensures the file exists.
    ''' If not configured, prompts the user for a path and stores it into the main INI.
    ''' </summary>
    ''' <param name="ownerForm">Owner window.</param>
    ''' <param name="context">Shared context.</param>
    ''' <param name="mainIniPath">Main INI path.</param>
    ''' <param name="isAlternate">True for AlternateModelPath; False for SpecialServicePath.</param>
    ''' <param name="targetIniPath">Resolved target INI path (expanded for access).</param>
    ''' <param name="mainIniWasUpdated">True if the main INI was updated with the path key.</param>
    ''' <returns>True if the target INI path is available; otherwise False.</returns>
    Private Shared Function TryEnsureSectionedIniPath(
        ownerForm As System.Windows.Forms.Form,
        context As ISharedContext,
        mainIniPath As System.String,
        isAlternate As System.Boolean,
        ByRef targetIniPath As System.String,
        Optional ByRef mainIniWasUpdated As System.Boolean = False
    ) As System.Boolean

        targetIniPath = Nothing
        mainIniWasUpdated = False

        Dim currentSetting As System.String = Nothing
        Dim settingKey As System.String = If(isAlternate, "AlternateModelPath", "SpecialServicePath")
        Dim defaultFileName As System.String = If(isAlternate, "allmodels.ini", "specialservices.ini")

        Try
            currentSetting = If(isAlternate, context.INI_AlternateModelPath, context.INI_SpecialServicePath)
        Catch
        End Try

        Dim expandedCurrent As System.String = Nothing
        If Not System.String.IsNullOrWhiteSpace(currentSetting) Then
            Try
                expandedCurrent = ExpandEnvironmentVariables(currentSetting)
            Catch
                expandedCurrent = currentSetting
            End Try
        End If

        If Not System.String.IsNullOrWhiteSpace(expandedCurrent) Then
            targetIniPath = expandedCurrent

            Try
                EnsureIniFileExists(targetIniPath)
            Catch ex As System.Exception
                ShowCustomMessageBox("Could not create file: " & ex.Message)
                Return False
            End Try

            Return True
        End If

        Dim baseDir As System.String = System.IO.Path.GetDirectoryName(mainIniPath)
        Dim suggested As System.String = System.IO.Path.Combine(baseDir, defaultFileName)

        Dim wasTopMost As Boolean = False
        Dim hadOwner As Boolean = (ownerForm IsNot Nothing)
        Dim chosenPath As String = ""

        If hadOwner Then
            wasTopMost = ownerForm.TopMost
            ownerForm.TopMost = False
            ownerForm.Enabled = False
            System.Windows.Forms.Application.DoEvents()
        End If

        Try
            chosenPath =
                ShowCustomInputBox(
                    "Please confirm or change the file path for " & settingKey & " (if unsure, just confirm):",
                    TITLE_IMPORT,
                    True,
                    suggested
                )

            If System.String.IsNullOrWhiteSpace(chosenPath) Then
                ShowCustomMessageBox("No path provided. Import aborted.")
                Return False
            End If

        Finally
            If hadOwner Then
                ownerForm.Enabled = True
                ownerForm.TopMost = wasTopMost
                ownerForm.Activate()
            End If
        End Try

        chosenPath = chosenPath.Trim()

        Dim expandedChosen As System.String = chosenPath
        Try
            expandedChosen = ExpandEnvironmentVariables(chosenPath)
        Catch
        End Try

        targetIniPath = expandedChosen

        Try
            EnsureIniFileExists(targetIniPath)
        Catch ex As System.Exception
            ShowCustomMessageBox("Could not create file: " & ex.Message)
            Return False
        End Try

        Dim valueToStore As System.String = chosenPath

        Try
            Dim plan As DryRunPlan = BuildDryRunPlanForSingleKey(mainIniPath, settingKey, valueToStore)
            CommitDryRunPlan(plan)
            mainIniWasUpdated = True
        Catch ex As System.Exception
            ShowCustomMessageBox("Could not update main configuration with " & settingKey & ": " & ex.Message)
            Return False
        End Try

        Return True

    End Function

    ''' <summary>
    ''' Ensures that an INI file exists at the given path and creates its directory if required.
    ''' </summary>
    ''' <param name="path">Target file path.</param>
    Private Shared Sub EnsureIniFileExists(path As System.String)
        Dim dir As System.String = System.IO.Path.GetDirectoryName(path)
        If Not System.String.IsNullOrWhiteSpace(dir) AndAlso Not System.IO.Directory.Exists(dir) Then
            System.IO.Directory.CreateDirectory(dir)
        End If

        If Not System.IO.File.Exists(path) Then
            System.IO.File.WriteAllText(
                path,
                "; created by Red Ink Settings Importer on " &
                System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") &
                System.Environment.NewLine,
                System.Text.Encoding.UTF8
            )
        End If
    End Sub

    ' =========================================================================================
    '  SECTION NAME EXTRACTION (Alternate/Special)
    ' =========================================================================================

    ''' <summary>
    ''' Extracts the first section name from import text, or prompts the user if none exists.
    ''' </summary>
    ''' <param name="ownerForm">Owner window.</param>
    ''' <param name="text">Import text.</param>
    ''' <param name="sectionName">The resolved section name.</param>
    ''' <param name="friendlyType">Friendly kind description used in the prompt.</param>
    ''' <returns>True if a section name is available; otherwise False.</returns>
    Private Shared Function TryGetSectionNameFromImportText(ownerForm As System.Windows.Forms.Form,
                                                           text As System.String,
                                                           ByRef sectionName As System.String,
                                                           friendlyType As System.String) As System.Boolean

        sectionName = Nothing

        If System.String.IsNullOrWhiteSpace(text) Then
            ShowCustomMessageBox("Import content is empty.")
            Return False
        End If

        Dim lines As System.Collections.Generic.List(Of System.String) = SplitToLinesPreserve(text)
        For Each line As System.String In lines
            Dim t As System.String = line.Trim()
            If t.StartsWith("[") AndAlso t.EndsWith("]") AndAlso t.Length >= 2 Then
                sectionName = t.Substring(1, t.Length - 2).Trim()
                Exit For
            End If
        Next

        If System.String.IsNullOrWhiteSpace(sectionName) Then

            Dim wasTopMost As Boolean = False
            Dim hadOwner As Boolean = (ownerForm IsNot Nothing)

            If hadOwner Then
                wasTopMost = ownerForm.TopMost
                ownerForm.TopMost = False
                ownerForm.Enabled = False
                System.Windows.Forms.Application.DoEvents()
            End If

            Try
                sectionName = ShowCustomInputBox(
                    $"You wish to import settings that require a Section header (a user friendly name of the model or service, e.g., 'LexiSearch' or 'Gemini 3 Pro with minimum reasoning'). It can be changed later. Please enter a section name for the {friendlyType}:",
                    TITLE_IMPORT,
                    True,
                    "Name")
            Finally
                If hadOwner Then
                    ownerForm.Enabled = True
                    ownerForm.TopMost = wasTopMost
                    ownerForm.Activate()
                End If
            End Try

            If System.String.IsNullOrWhiteSpace(sectionName) Then
                ShowCustomMessageBox("No section name provided. Import aborted.")
                Return False
            End If
            sectionName = sectionName.Trim()
        End If

        Return True

    End Function

    ' =========================================================================================
    '  DRY RUN PLAN + COMMIT
    ' =========================================================================================

    ''' <summary>
    ''' Represents a planned INI update, including generated content and backup/audit information.
    ''' </summary>
    Private NotInheritable Class DryRunPlan
        Public Property TargetIniPath As System.String
        Public Property Kind As ImportKind
        Public Property TargetSectionName As System.String
        Public Property NewFileLines As System.Collections.Generic.List(Of System.String)
        Public Property RemovedLinesBackup As System.Collections.Generic.List(Of System.String)
        Public Property OverwrittenKeys As System.Collections.Generic.List(Of System.String)
        Public Property WillCreateRemovedBackup As System.Boolean

        Public Property SectionNames As System.Collections.Generic.List(Of System.String)
        Public Property AddedKeysBySection As System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of System.String))
        Public Property OverwrittenKeysBySection As System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of System.String))

        ''' <summary>
        ''' Builds a user-facing summary text for a single plan.
        ''' </summary>
        Public Function GetUserSummary() As System.String
            Dim sb As New System.Text.StringBuilder()

            sb.AppendLine("Dry run – review before importing")
            sb.AppendLine()
            sb.AppendLine("Target file: " & TargetIniPath)
            sb.AppendLine()

            If Kind = ImportKind.AlternateModel OrElse Kind = ImportKind.SpecialService Then
                sb.AppendLine("Target section: [" & TargetSectionName & "]")
                sb.AppendLine()
            End If

            If OverwrittenKeys IsNot Nothing AndAlso OverwrittenKeys.Count > 0 Then
                sb.AppendLine(
                    "Keys that will be overwritten (" &
                    OverwrittenKeys.Count.ToString() &
                    "): " &
                    System.String.Join(", ", OverwrittenKeys)
                )
                sb.AppendLine()
            Else
                If Kind = ImportKind.PrimaryModel OrElse
                   Kind = ImportKind.SecondaryModel OrElse
                   Kind = ImportKind.OtherParameters Then

                    sb.AppendLine("No existing keys will be overwritten (new keys will be appended at the end of the file).")
                    sb.AppendLine()
                End If
            End If

            If WillCreateRemovedBackup Then
                sb.AppendLine("A full backup of the target file and a backup of the removed content will always be created in the same directory.")
            Else
                sb.AppendLine("A full backup of the target file will always be created in the same directory.")
            End If

            sb.AppendLine()
            sb.AppendLine("Proceed with import?")

            Return sb.ToString()
        End Function

    End Class

    ''' <summary>
    ''' Builds a dry-run plan for a target INI file based on import kind and imported content.
    ''' </summary>
    Private Shared Function BuildDryRunPlan(context As ISharedContext,
                                           kind As ImportKind,
                                           targetIniPath As System.String,
                                           targetSectionName As System.String,
                                           importedLines As System.Collections.Generic.List(Of System.String),
                                           Optional replaceExisting As System.Boolean = True) As DryRunPlan

        If System.String.IsNullOrWhiteSpace(targetIniPath) Then Throw New System.ArgumentNullException(NameOf(targetIniPath))
        If importedLines Is Nothing OrElse importedLines.Count = 0 Then Throw New System.Exception("No imported lines.")

        Dim existingLines As System.Collections.Generic.List(Of System.String) = ReadAllLinesPreserve(targetIniPath)

        Dim plan As New DryRunPlan() With {
            .TargetIniPath = targetIniPath,
            .Kind = kind,
            .TargetSectionName = targetSectionName,
            .NewFileLines = New System.Collections.Generic.List(Of System.String)(existingLines),
            .RemovedLinesBackup = New System.Collections.Generic.List(Of System.String)(),
            .OverwrittenKeys = New System.Collections.Generic.List(Of System.String)(),
            .WillCreateRemovedBackup = False,
            .SectionNames = New System.Collections.Generic.List(Of System.String)(),
            .AddedKeysBySection = New System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of System.String))(System.StringComparer.OrdinalIgnoreCase),
            .OverwrittenKeysBySection = New System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of System.String))(System.StringComparer.OrdinalIgnoreCase)
        }

        If kind = ImportKind.AlternateModel OrElse kind = ImportKind.SpecialService Then
            Dim newSectionLines As System.Collections.Generic.List(Of System.String) = BuildSectionLines(targetSectionName, importedLines)
            ApplySectionReplace(plan, existingLines, newSectionLines)
            Return plan
        End If

        Dim kv As System.Collections.Generic.Dictionary(Of System.String, System.String) = ParseKeyValueLines(importedLines)

        If kind = ImportKind.SecondaryModel Then
            kv = ConvertKeysToSecondary(kv)
            If Not kv.ContainsKey("SecondAPI") Then
                kv.Add("SecondAPI", "True")
            Else
                kv("SecondAPI") = "True"
            End If
        End If

        ApplyMainIniKeyReplaceAppend(plan, existingLines, kv, replaceExisting)

        Return plan

    End Function

    ''' <summary>
    ''' Builds a dry-run plan that replaces/appends a single key/value in the main INI.
    ''' </summary>
    Private Shared Function BuildDryRunPlanForSingleKey(mainIniPath As System.String, key As System.String, value As System.String) As DryRunPlan
        Dim existingLines As System.Collections.Generic.List(Of System.String) = ReadAllLinesPreserve(mainIniPath)

        Dim kv As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)
        kv(key) = value

        Dim plan As New DryRunPlan() With {
            .TargetIniPath = mainIniPath,
            .Kind = ImportKind.OtherParameters,
            .TargetSectionName = Nothing,
            .NewFileLines = New System.Collections.Generic.List(Of System.String)(existingLines),
            .RemovedLinesBackup = New System.Collections.Generic.List(Of System.String)(),
            .OverwrittenKeys = New System.Collections.Generic.List(Of System.String)(),
            .WillCreateRemovedBackup = False
        }

        ApplyMainIniKeyReplaceAppend(plan, existingLines, kv)
        Return plan
    End Function

    ''' <summary>
    ''' Injects explanatory comment lines for resolved placeholders into a dry-run plan's
    ''' generated output content, replacing any existing placeholder comments with the same names.
    ''' </summary>
    ''' <remarks>
    ''' For sectioned INI files (alternate models or special services), existing placeholder comments
    ''' within the section are removed first, then new comments are inserted at the beginning of each
    ''' affected section, preceded and followed by exactly one empty line.
    ''' For global (non-sectioned) INI files, existing placeholder comments anywhere in the file are
    ''' removed first, then new comments are appended at the end of the file, also surrounded by
    ''' exactly one empty line before and after the comment block.
    '''
    ''' The injected comments use the format:
    '''   ; [[placeholder]] = uservalue
    '''
    ''' This method modifies <see cref="DryRunPlan.NewFileLines"/> in-place and does not affect
    ''' parsing, backup generation, or commit semantics.
    ''' </remarks>
    ''' <param name="plan">
    ''' The dry-run plan whose generated output lines will be annotated.
    ''' </param>
    ''' <param name="resolved">
    ''' The list of placeholders that were resolved during import processing.
    ''' If empty or Nothing, no changes are made.
    ''' </param>
    Private Shared Sub InjectPlaceholderCommentsIntoPlan(
        plan As DryRunPlan,
        resolved As System.Collections.Generic.List(Of ResolvedPlaceholder)
    )
        If plan Is Nothing OrElse
           resolved Is Nothing OrElse
           resolved.Count = 0 OrElse
           plan.NewFileLines Is Nothing Then
            Return
        End If

        ' Build a set of placeholder names being resolved (for matching existing comments)
        Dim resolvedNames As New System.Collections.Generic.HashSet(Of System.String)(
            System.StringComparer.OrdinalIgnoreCase)
        For Each rp As ResolvedPlaceholder In resolved
            If Not System.String.IsNullOrWhiteSpace(rp.Name) Then
                resolvedNames.Add(rp.Name)
            End If
        Next

        ' Regex to match existing placeholder comment lines: ; [[name]] = value
        Dim placeholderCommentRx As New System.Text.RegularExpressions.Regex(
            "^\s*;\s*\[\[(.+?)\]\]\s*=",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase)

        ' Build the new comment lines
        Dim commentLines As New System.Collections.Generic.List(Of System.String)()
        For Each rp As ResolvedPlaceholder In resolved
            commentLines.Add("; [[" & rp.Name & "]] = " & rp.Value)
        Next

        ' ==========================================================
        ' SECTIONED INI (AlternateModel / SpecialService)
        ' ==========================================================
        If plan.Kind = ImportKind.AlternateModel OrElse
           plan.Kind = ImportKind.SpecialService Then

            For Each sectionName As System.String In plan.SectionNames

                Dim startIndex As System.Int32 = -1
                Dim endIndex As System.Int32 = -1

                FindSectionRange(plan.NewFileLines, sectionName, startIndex, endIndex)
                If startIndex < 0 Then Continue For

                ' Remove existing placeholder comments for the resolved placeholders within this section
                Dim i As System.Int32 = startIndex + 1
                While i <= endIndex AndAlso i < plan.NewFileLines.Count
                    Dim line As System.String = plan.NewFileLines(i)
                    Dim m As System.Text.RegularExpressions.Match = placeholderCommentRx.Match(If(line, ""))
                    If m.Success Then
                        Dim existingName As System.String = m.Groups(1).Value.Trim()
                        If resolvedNames.Contains(existingName) Then
                            plan.NewFileLines.RemoveAt(i)
                            endIndex -= 1
                            Continue While
                        End If
                    End If
                    i += 1
                End While

                ' Recalculate section range after removals
                FindSectionRange(plan.NewFileLines, sectionName, startIndex, endIndex)
                If startIndex < 0 Then Continue For

                Dim insertIndex As System.Int32 = startIndex + 1

                ' Remove all empty lines directly after the section header
                While insertIndex < plan.NewFileLines.Count AndAlso
                      plan.NewFileLines(insertIndex).Trim().Length = 0
                    plan.NewFileLines.RemoveAt(insertIndex)
                End While

                ' Insert ONE empty line before comments
                plan.NewFileLines.Insert(insertIndex, "")
                insertIndex += 1

                ' Insert comments
                For Each c As System.String In commentLines
                    plan.NewFileLines.Insert(insertIndex, c)
                    insertIndex += 1
                Next

                ' Remove all empty lines after comments
                While insertIndex < plan.NewFileLines.Count AndAlso
                      plan.NewFileLines(insertIndex).Trim().Length = 0
                    plan.NewFileLines.RemoveAt(insertIndex)
                End While

                ' Insert ONE empty line after comments
                plan.NewFileLines.Insert(insertIndex, "")
            Next

        Else
            ' ==========================================================
            ' GLOBAL INI
            ' ==========================================================

            ' Remove existing placeholder comments for the resolved placeholders anywhere in the file
            Dim i As System.Int32 = 0
            While i < plan.NewFileLines.Count
                Dim line As System.String = plan.NewFileLines(i)
                Dim m As System.Text.RegularExpressions.Match = placeholderCommentRx.Match(If(line, ""))
                If m.Success Then
                    Dim existingName As System.String = m.Groups(1).Value.Trim()
                    If resolvedNames.Contains(existingName) Then
                        plan.NewFileLines.RemoveAt(i)
                        Continue While
                    End If
                End If
                i += 1
            End While

            ' Remove trailing empty lines
            While plan.NewFileLines.Count > 0 AndAlso
                  plan.NewFileLines(plan.NewFileLines.Count - 1).Trim().Length = 0
                plan.NewFileLines.RemoveAt(plan.NewFileLines.Count - 1)
            End While

            ' ONE empty line before comments
            plan.NewFileLines.Add("")

            ' Comments
            For Each c As System.String In commentLines
                plan.NewFileLines.Add(c)
            Next

            ' ONE empty line after comments
            plan.NewFileLines.Add("")
        End If
    End Sub


    ''' <summary>
    ''' Commits a dry-run plan by writing backups and replacing the target file content (atomic replace when possible).
    ''' </summary>
    ''' <param name="plan">Plan to commit.</param>
    Private Shared Sub CommitDryRunPlan(plan As DryRunPlan)

        If plan Is Nothing Then Throw New System.ArgumentNullException(NameOf(plan))
        If System.String.IsNullOrWhiteSpace(plan.TargetIniPath) Then Throw New System.Exception("Target ini path missing.")
        If plan.NewFileLines Is Nothing Then Throw New System.Exception("No new content.")

        Dim targetPath As System.String = plan.TargetIniPath

        Dim targetDir As System.String = System.IO.Path.GetDirectoryName(targetPath)
        If System.String.IsNullOrWhiteSpace(targetDir) Then
            Throw New System.Exception("Invalid target path.")
        End If

        If Not System.IO.Directory.Exists(targetDir) Then
            Throw New System.Exception("Target directory does not exist: " & targetDir)
        End If

        Dim ts As System.String = System.DateTime.Now.ToString("yyyyMMdd_HHmmss_fff")

        Dim baseName As System.String = System.IO.Path.GetFileNameWithoutExtension(targetPath)
        Dim extensionWithDot As System.String = System.IO.Path.GetExtension(targetPath)

        Dim fullBackup As System.String =
            System.IO.Path.Combine(targetDir, baseName & extensionWithDot & "_" & ts & ".bak")

        If System.IO.File.Exists(targetPath) Then
            System.IO.File.Copy(targetPath, fullBackup, overwrite:=False)
        End If

        If plan.WillCreateRemovedBackup AndAlso
           plan.RemovedLinesBackup IsNot Nothing AndAlso
           plan.RemovedLinesBackup.Count > 0 Then

            Dim removedBackup As System.String =
                System.IO.Path.Combine(targetDir, baseName & extensionWithDot & "_removed_" & ts & ".bak")

            System.IO.File.WriteAllText(
                removedBackup,
                System.String.Join(System.Environment.NewLine, plan.RemovedLinesBackup),
                System.Text.Encoding.UTF8
            )
        End If

        Dim tmpPath As System.String =
            System.IO.Path.Combine(
                targetDir,
                baseName & "_tmp_" & System.Guid.NewGuid().ToString("N") & extensionWithDot
            )

        System.IO.File.WriteAllText(
            tmpPath,
            System.String.Join(System.Environment.NewLine, plan.NewFileLines),
            System.Text.Encoding.UTF8
        )

        Try
            If System.IO.File.Exists(targetPath) Then
                System.IO.File.Replace(tmpPath, targetPath, Nothing, True)
            Else
                System.IO.File.Move(tmpPath, targetPath)
            End If
        Catch
            If System.IO.File.Exists(targetPath) Then
                Try
                    System.IO.File.Delete(targetPath)
                Catch
                End Try
            End If

            System.IO.File.Move(tmpPath, targetPath)
        End Try

    End Sub

    ''' <summary>
    ''' Computes a combined sectioned-import plan that replaces/imports multiple sections in a single target file.
    ''' </summary>
    Private Shared Function BuildSectionedCombinedPlan(
        context As ISharedContext,
        kind As ImportKind,
        targetIniPath As System.String,
        segments As System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of System.String))
    ) As DryRunPlan

        If segments Is Nothing OrElse segments.Count = 0 Then Throw New System.Exception("No valid sections found.")
        If System.String.IsNullOrWhiteSpace(targetIniPath) Then Throw New System.ArgumentNullException(NameOf(targetIniPath))

        Dim combinedPlan As DryRunPlan = Nothing

        Dim removedAll As New System.Collections.Generic.List(Of System.String)()
        Dim overwrittenAll As New System.Collections.Generic.List(Of System.String)()

        Dim sectionNames As New System.Collections.Generic.List(Of System.String)()
        Dim addedBySection As New System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of System.String))(System.StringComparer.OrdinalIgnoreCase)
        Dim overwrittenBySection As New System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of System.String))(System.StringComparer.OrdinalIgnoreCase)

        For Each kvp As System.Collections.Generic.KeyValuePair(Of System.String, System.Collections.Generic.List(Of System.String)) In segments

            Dim sectionName As System.String = kvp.Key
            Dim sectionBodyLines As System.Collections.Generic.List(Of System.String) = kvp.Value

            If System.String.IsNullOrWhiteSpace(sectionName) Then Continue For

            sectionNames.Add(sectionName)

            Dim importKeys As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            If sectionBodyLines IsNot Nothing Then
                For Each l As String In sectionBodyLines
                    Dim k As String = Nothing
                    If TryParseIniKey(l, k) AndAlso Not System.String.IsNullOrWhiteSpace(k) Then
                        importKeys.Add(k)
                    End If
                Next
            End If

            Dim currentFileLines As System.Collections.Generic.List(Of System.String) =
                If(combinedPlan Is Nothing, ReadAllLinesPreserve(targetIniPath), combinedPlan.NewFileLines)

            Dim currentSectionKeys As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            Dim startIndex As Integer = -1
            Dim endIndex As Integer = -1
            FindSectionRange(currentFileLines, sectionName, startIndex, endIndex)

            If startIndex >= 0 AndAlso endIndex >= startIndex Then
                For i As Integer = startIndex + 1 To endIndex
                    Dim k As String = Nothing
                    If TryParseIniKey(currentFileLines(i), k) AndAlso Not System.String.IsNullOrWhiteSpace(k) Then
                        currentSectionKeys.Add(k)
                    End If
                Next
            End If

            Dim addedKeys As New System.Collections.Generic.List(Of System.String)()
            Dim overwrittenKeys As New System.Collections.Generic.List(Of System.String)()

            For Each k As String In importKeys
                If currentSectionKeys.Contains(k) Then
                    overwrittenKeys.Add(k)
                Else
                    addedKeys.Add(k)
                End If
            Next

            addedBySection(sectionName) = UniquePreserveOrder(addedKeys)
            overwrittenBySection(sectionName) = UniquePreserveOrder(overwrittenKeys)

            Dim newSectionLines As System.Collections.Generic.List(Of System.String) = BuildSectionLines(sectionName, sectionBodyLines)

            If combinedPlan Is Nothing Then
                combinedPlan = BuildDryRunPlan(context, kind, targetIniPath, sectionName, sectionBodyLines)
            Else
                combinedPlan.TargetSectionName = sectionName
                ApplySectionReplace(combinedPlan, combinedPlan.NewFileLines, newSectionLines)
            End If

            If combinedPlan.WillCreateRemovedBackup AndAlso combinedPlan.RemovedLinesBackup IsNot Nothing AndAlso combinedPlan.RemovedLinesBackup.Count > 0 Then
                removedAll.AddRange(combinedPlan.RemovedLinesBackup)
            End If

            If combinedPlan.OverwrittenKeys IsNot Nothing AndAlso combinedPlan.OverwrittenKeys.Count > 0 Then
                overwrittenAll.AddRange(combinedPlan.OverwrittenKeys)
            End If

        Next

        If combinedPlan Is Nothing Then Throw New System.Exception("No sections to import.")

        combinedPlan.RemovedLinesBackup = removedAll
        combinedPlan.OverwrittenKeys = UniquePreserveOrder(overwrittenAll)
        combinedPlan.WillCreateRemovedBackup = (removedAll.Count > 0)

        combinedPlan.SectionNames = sectionNames
        combinedPlan.AddedKeysBySection = addedBySection
        combinedPlan.OverwrittenKeysBySection = overwrittenBySection

        combinedPlan.TargetSectionName = Nothing

        Return combinedPlan

    End Function

    ' =========================================================================================
    '  APPLY: redink.ini key overwrite (remove old lines, append new at end)
    ' =========================================================================================

    ''' <summary>
    ''' Applies key updates to a main INI file by removing existing key lines (optional) and appending new lines at the end.
    ''' </summary>
    ''' <param name="plan">Plan to update.</param>
    ''' <param name="existingLines">Existing file lines.</param>
    ''' <param name="newKeyValues">Key/value pairs to apply.</param>
    ''' <param name="replaceExisting">True to replace existing key lines; False to only add new keys.</param>
    Private Shared Sub ApplyMainIniKeyReplaceAppend(
        plan As DryRunPlan,
        existingLines As List(Of String),
        newKeyValues As Dictionary(Of String, String),
        Optional replaceExisting As Boolean = True
    )

        Dim keys As New System.Collections.Generic.HashSet(Of System.String)(newKeyValues.Keys, System.StringComparer.OrdinalIgnoreCase)

        Dim newLines As New System.Collections.Generic.List(Of System.String)()
        Dim removed As New System.Collections.Generic.List(Of System.String)()
        Dim overwritten As New System.Collections.Generic.List(Of System.String)()

        For Each line As System.String In existingLines
            Dim parsedKey As System.String = Nothing
            Dim isKeyLine As System.Boolean = TryParseIniKey(line, parsedKey)

            If isKeyLine AndAlso Not System.String.IsNullOrWhiteSpace(parsedKey) AndAlso keys.Contains(parsedKey) Then
                If replaceExisting Then
                    removed.Add(line)
                    overwritten.Add(parsedKey)
                    Continue For
                Else
                    newLines.Add(line)
                    Continue For
                End If
            End If

            newLines.Add(line)
        Next

        If newLines.Count > 0 Then
            Dim last As System.String = newLines(newLines.Count - 1)
            If last IsNot Nothing AndAlso last.Trim().Length > 0 Then
                newLines.Add("")
            End If
        End If

        Dim existingKeys As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
        For Each line As System.String In newLines
            Dim k0 As System.String = Nothing
            If TryParseIniKey(line, k0) AndAlso Not System.String.IsNullOrWhiteSpace(k0) Then
                existingKeys.Add(k0)
            End If
        Next

        For Each kvp As System.Collections.Generic.KeyValuePair(Of System.String, System.String) In newKeyValues
            If replaceExisting Then
                newLines.Add(kvp.Key & " = " & kvp.Value)
            Else
                If Not existingKeys.Contains(kvp.Key) Then
                    newLines.Add(kvp.Key & " = " & kvp.Value)
                    existingKeys.Add(kvp.Key)
                End If
            End If
        Next

        plan.NewFileLines = newLines

        If removed.Count > 0 Then
            plan.WillCreateRemovedBackup = True
            plan.RemovedLinesBackup = removed
            plan.OverwrittenKeys = UniquePreserveOrder(overwritten)
        Else
            plan.WillCreateRemovedBackup = False
            plan.RemovedLinesBackup = New System.Collections.Generic.List(Of System.String)()
            plan.OverwrittenKeys = New System.Collections.Generic.List(Of System.String)()
        End If

    End Sub

    ''' <summary>
    ''' Returns a case-insensitive unique list preserving original order.
    ''' </summary>
    ''' <param name="items">Input items.</param>
    ''' <returns>Unique items in original order.</returns>
    Private Shared Function UniquePreserveOrder(items As System.Collections.Generic.List(Of System.String)) As System.Collections.Generic.List(Of System.String)
        Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
        Dim res As New System.Collections.Generic.List(Of System.String)()
        For Each s As System.String In items
            If System.String.IsNullOrWhiteSpace(s) Then Continue For
            If seen.Add(s) Then res.Add(s)
        Next
        Return res
    End Function

    ' =========================================================================================
    '  APPLY: section replace
    ' =========================================================================================

    ''' <summary>
    ''' Replaces the target section in-place if it exists; otherwise appends the section at end of file.
    ''' </summary>
    ''' <param name="plan">Plan to update (TargetSectionName must be set).</param>
    ''' <param name="existingLines">Existing file lines.</param>
    ''' <param name="newSectionLines">New section lines including section header.</param>
    Private Shared Sub ApplySectionReplace(plan As DryRunPlan,
                                          existingLines As System.Collections.Generic.List(Of System.String),
                                          newSectionLines As System.Collections.Generic.List(Of System.String))

        Dim sectionName As System.String = plan.TargetSectionName
        If System.String.IsNullOrWhiteSpace(sectionName) Then Throw New System.Exception("Section name missing.")

        Dim startIndex As System.Int32 = -1
        Dim endIndex As System.Int32 = -1

        FindSectionRange(existingLines, sectionName, startIndex, endIndex)

        Dim newFile As New System.Collections.Generic.List(Of System.String)()

        If startIndex >= 0 AndAlso endIndex >= startIndex Then
            Dim removed As New System.Collections.Generic.List(Of System.String)()
            For i As System.Int32 = startIndex To endIndex
                removed.Add(existingLines(i))
            Next

            For i As System.Int32 = 0 To startIndex - 1
                newFile.Add(existingLines(i))
            Next

            For Each l As System.String In newSectionLines
                newFile.Add(l)
            Next

            If endIndex + 1 < existingLines.Count Then
                Dim nextLine As System.String = existingLines(endIndex + 1)

                If nextLine IsNot Nothing Then
                    Dim trimmed As System.String = nextLine.Trim()

                    If trimmed.StartsWith("[") AndAlso trimmed.EndsWith("]") Then
                        If newFile.Count > 0 AndAlso newFile(newFile.Count - 1).Trim().Length > 0 Then
                            newFile.Add("")
                        End If
                    End If
                End If
            End If

            For i As System.Int32 = endIndex + 1 To existingLines.Count - 1
                newFile.Add(existingLines(i))
            Next

            plan.NewFileLines = newFile
            plan.WillCreateRemovedBackup = True
            plan.RemovedLinesBackup = removed

            Dim overwrittenKeys As New System.Collections.Generic.List(Of System.String)()

            For Each line As String In removed
                Dim key As String = Nothing
                If TryParseIniKey(line, key) Then
                    overwrittenKeys.Add(key)
                End If
            Next

            plan.OverwrittenKeys = UniquePreserveOrder(overwrittenKeys)

        Else
            newFile = New System.Collections.Generic.List(Of System.String)(existingLines)

            If newFile.Count > 0 Then
                Dim last As System.String = newFile(newFile.Count - 1)
                If last IsNot Nothing AndAlso last.Trim().Length > 0 Then
                    newFile.Add("")
                End If
            End If

            For Each l As System.String In newSectionLines
                newFile.Add(l)
            Next

            plan.NewFileLines = newFile
            plan.WillCreateRemovedBackup = False
            plan.RemovedLinesBackup = New System.Collections.Generic.List(Of System.String)()
            plan.OverwrittenKeys = New System.Collections.Generic.List(Of System.String)()
        End If

    End Sub

    ''' <summary>
    ''' Finds the start and end line indices of a named section within an INI file.
    ''' </summary>
    ''' <param name="lines">INI file lines.</param>
    ''' <param name="sectionName">Section name (without brackets).</param>
    ''' <param name="startIndex">Index of section header line; -1 if not found.</param>
    ''' <param name="endIndex">Index of last line belonging to the section.</param>
    Private Shared Sub FindSectionRange(lines As System.Collections.Generic.List(Of System.String),
                                        sectionName As System.String,
                                        ByRef startIndex As System.Int32,
                                        ByRef endIndex As System.Int32)

        startIndex = -1
        endIndex = -1
        If lines Is Nothing OrElse lines.Count = 0 Then Return

        Dim targetHeader As System.String = "[" & sectionName & "]"

        For i As System.Int32 = 0 To lines.Count - 1
            Dim t As System.String = lines(i)
            If t Is Nothing Then Continue For
            Dim trimmed As System.String = t.Trim()
            If trimmed.StartsWith("[") AndAlso trimmed.EndsWith("]") Then
                If System.String.Equals(trimmed, targetHeader, System.StringComparison.OrdinalIgnoreCase) Then
                    startIndex = i
                    Exit For
                End If
            End If
        Next

        If startIndex < 0 Then Return

        endIndex = lines.Count - 1
        For i As System.Int32 = startIndex + 1 To lines.Count - 1
            Dim trimmed As System.String = If(lines(i), "").Trim()
            If trimmed.StartsWith("[") AndAlso trimmed.EndsWith("]") AndAlso trimmed.Length >= 2 Then
                endIndex = i - 1
                Exit For
            End If
        Next

    End Sub

    ''' <summary>
    ''' Builds a section block including header and a blank line, skipping any section headers in imported content.
    ''' </summary>
    ''' <param name="sectionName">Section name.</param>
    ''' <param name="importedLines">Imported lines.</param>
    ''' <returns>Section block lines.</returns>
    Private Shared Function BuildSectionLines(sectionName As System.String,
                                              importedLines As System.Collections.Generic.List(Of System.String)) As System.Collections.Generic.List(Of System.String)

        Dim res As New System.Collections.Generic.List(Of System.String)()
        res.Add("[" & sectionName & "]")
        res.Add("")

        For Each line As System.String In importedLines
            Dim t As System.String = line.Trim()
            If t.StartsWith("[") AndAlso t.EndsWith("]") AndAlso t.Length >= 2 Then
                Continue For
            End If
            res.Add(line)
        Next

        Return res
    End Function

    ' =========================================================================================
    '  PARSING UTILITIES
    ' =========================================================================================

    ''' <summary>
    ''' Splits text into lines and preserves empty lines; normalizes CR/LF combinations.
    ''' </summary>
    Private Shared Function SplitToLinesPreserve(text As System.String) As System.Collections.Generic.List(Of System.String)
        Dim res As New System.Collections.Generic.List(Of System.String)()
        If text Is Nothing Then Return res

        Dim normalized As System.String = text.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
        Dim parts As System.String() = normalized.Split(New System.Char() {ControlChars.Lf}, System.StringSplitOptions.None)
        res.AddRange(parts)
        Return res
    End Function

    ''' <summary>
    ''' Splits text into lines and preserves all lines (including blank/comment lines).
    ''' </summary>
    Private Shared Function SplitToLinesPreserveNonEmpty(text As System.String) As System.Collections.Generic.List(Of System.String)
        Dim all As System.Collections.Generic.List(Of System.String) = SplitToLinesPreserve(text)
        Dim res As New System.Collections.Generic.List(Of System.String)()
        For Each l As System.String In all
            If l Is Nothing Then Continue For
            res.Add(l)
        Next
        Return res
    End Function

    ''' <summary>
    ''' Reads an INI file as UTF-8 and returns lines preserving original line boundaries.
    ''' </summary>
    Private Shared Function ReadAllLinesPreserve(path As System.String) As System.Collections.Generic.List(Of System.String)
        Dim text As System.String = System.IO.File.ReadAllText(path, System.Text.Encoding.UTF8)
        Return SplitToLinesPreserve(text)
    End Function

    ''' <summary>
    ''' Tries to parse an INI key from a line in the form "key = value".
    ''' </summary>
    ''' <param name="line">Line text.</param>
    ''' <param name="key">Parsed key on success.</param>
    ''' <returns>True if the line contains a key assignment; otherwise False.</returns>
    Private Shared Function TryParseIniKey(line As System.String, ByRef key As System.String) As System.Boolean
        key = Nothing
        If line Is Nothing Then Return False

        Dim trimmed As System.String = line.TrimStart()
        If trimmed.StartsWith(";", System.StringComparison.Ordinal) Then Return False
        If trimmed.StartsWith("[", System.StringComparison.Ordinal) Then Return False

        Dim idx As System.Int32 = line.IndexOf("="c)
        If idx <= 0 Then Return False

        Dim left As System.String = line.Substring(0, idx).Trim()
        If System.String.IsNullOrWhiteSpace(left) Then Return False

        key = left
        Return True
    End Function

    ''' <summary>
    ''' Reads a specific key's value from an INI file (case-insensitive key match).
    ''' Returns Nothing if the key is not found or on any error.
    ''' </summary>
    Private Shared Function TryReadIniKeyValue(iniPath As System.String, key As System.String) As System.String
        If System.String.IsNullOrWhiteSpace(iniPath) Then Return Nothing
        If System.String.IsNullOrWhiteSpace(key) Then Return Nothing
        If Not System.IO.File.Exists(iniPath) Then Return Nothing

        Try
            Dim lines As System.Collections.Generic.List(Of System.String) = ReadAllLinesPreserve(iniPath)

            For Each line As System.String In lines
                If line Is Nothing Then Continue For

                Dim trimmedStart As System.String = line.TrimStart()
                If trimmedStart.StartsWith(";", System.StringComparison.Ordinal) Then Continue For
                If trimmedStart.StartsWith("[", System.StringComparison.Ordinal) Then Continue For

                Dim idx As System.Int32 = line.IndexOf("="c)
                If idx <= 0 Then Continue For

                Dim k As System.String = line.Substring(0, idx).Trim()

                If System.String.Equals(k, key, System.StringComparison.OrdinalIgnoreCase) Then
                    Return line.Substring(idx + 1).Trim()
                End If
            Next

        Catch
        End Try

        Return Nothing
    End Function

    ''' <summary>
    ''' Parses key/value pairs from INI lines and returns the last occurrence per key.
    ''' </summary>
    Private Shared Function ParseKeyValueLines(lines As System.Collections.Generic.List(Of System.String)) As System.Collections.Generic.Dictionary(Of System.String, System.String)
        Dim kv As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)

        For Each line As System.String In lines
            If line Is Nothing Then Continue For

            Dim trimmedStart As System.String = line.TrimStart()
            If trimmedStart.StartsWith(";", System.StringComparison.Ordinal) Then Continue For
            If trimmedStart.StartsWith("[", System.StringComparison.Ordinal) Then Continue For

            Dim idx As System.Int32 = line.IndexOf("="c)
            If idx <= 0 Then Continue For

            Dim k As System.String = line.Substring(0, idx).Trim()
            Dim v As System.String = line.Substring(idx + 1).Trim()

            If System.String.IsNullOrWhiteSpace(k) Then Continue For

            kv(k) = v
        Next

        Return kv
    End Function

    ' =========================================================================================
    '  MULTI-SEGMENT PARSER (SectionName -> Lines)
    ' =========================================================================================

    ''' <summary>
    ''' Parses an INI-like text stream into named section segments.
    ''' </summary>
    ''' <param name="lines">INI lines including section headers.</param>
    ''' <returns>Dictionary mapping section name to section body lines.</returns>
    Private Shared Function ParseIniSegments(
        lines As System.Collections.Generic.List(Of System.String)
    ) As System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of System.String))

        Dim result As New System.Collections.Generic.Dictionary(
            Of System.String,
            System.Collections.Generic.List(Of System.String)
        )(System.StringComparer.OrdinalIgnoreCase)

        Dim currentSection As System.String = Nothing
        Dim currentLines As System.Collections.Generic.List(Of System.String) = Nothing

        For Each line As System.String In lines
            Dim t As System.String = If(line, "").Trim()

            If t.StartsWith("[") AndAlso t.EndsWith("]") AndAlso t.Length > 2 Then
                currentSection = t.Substring(1, t.Length - 2).Trim()
                currentLines = New System.Collections.Generic.List(Of System.String)()
                result(currentSection) = currentLines
            ElseIf currentSection IsNot Nothing Then
                currentLines.Add(line)
            End If
        Next

        Return result
    End Function

    ''' <summary>
    ''' Converts main-model INI keys into secondary-model keys by applying a "_2" suffix, except for already suffixed keys.
    ''' </summary>
    ''' <param name="kv">Input key/value pairs.</param>
    ''' <returns>Converted key/value pairs.</returns>
    Private Shared Function ConvertKeysToSecondary(kv As System.Collections.Generic.Dictionary(Of System.String, System.String)) As System.Collections.Generic.Dictionary(Of System.String, System.String)
        Dim out As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)

        For Each kvp As System.Collections.Generic.KeyValuePair(Of System.String, System.String) In kv
            Dim k As System.String = kvp.Key
            Dim v As System.String = kvp.Value

            If System.String.IsNullOrWhiteSpace(k) Then Continue For

            If System.String.Equals(k, "SecondAPI", System.StringComparison.OrdinalIgnoreCase) Then
                out("SecondAPI") = "True"
            ElseIf k.EndsWith("_2", System.StringComparison.OrdinalIgnoreCase) Then
                out(k) = v
            Else
                out(k & "_2") = v
            End If
        Next

        Return out
    End Function

    ''' <summary>
    ''' Expands supported placeholders in sample list URLs.
    ''' </summary>
    ''' <param name="sourceUrl">Source URL text.</param>
    ''' <returns>Expanded URL text.</returns>
    Private Shared Function ExpandSourceUrlPlaceholders(
        sourceUrl As System.String
    ) As System.String

        If System.String.IsNullOrWhiteSpace(sourceUrl) Then Return sourceUrl

        Dim result As System.String = sourceUrl

        result = System.Text.RegularExpressions.Regex.Replace(
            result,
            "\{AppsUrl\}",
            AppsUrl,
            System.Text.RegularExpressions.RegexOptions.IgnoreCase
        )

        result = System.Text.RegularExpressions.Regex.Replace(
            result,
            "\{AppsUrlDir\}",
            AppsUrlDir,
            System.Text.RegularExpressions.RegexOptions.IgnoreCase
        )

        Return result
    End Function

End Class