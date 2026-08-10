' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File:    SharedMethods.UpdateIni.vb
' Purpose: Automated INI configuration update subsystem with optional Ed25519
'          signature validation, interactive approval UI, silent (policy-driven)
'          application modes, persistent ignore rules, placeholder preservation,
'          and audit logging.
'
' Architectural Overview
' ----------------------
' This module is implemented as a `SharedLibrary.SharedMethods` partial class and is
' driven by a context object (`ISharedContext`) supplied at runtime.
'
' It follows a "detect → filter → decide → apply" pipeline:
'   1) Detect differences between local INI files and configured update sources.
'   2) Collect signature/download diagnostics and track failed sources.
'   3) Filter out ignored changes (with override rules).
'   4) Resolve placeholders in remote values using local definitions.
'   5) Decide which changes may be applied (silent security level or interactive UI).
'   6) Create a timestamped backup and write updates back into the local INI file(s).
'   7) Log all relevant actions (security events are forced to log).
'
' Primary Entry Point (startup integration)
' ----------------------------------------
' - `Public Shared Function CheckForIniUpdates(ByRef context As ISharedContext) As Boolean`
'
' Placeholder Preservation System
' -------------------------------
' Remote update templates may contain placeholders like `[[ Your Tenant ]]` which are
' substituted with locally-defined values stored as INI comments:
'   `;  [[ Your Tenant ]] = contoso`
'
' Key functions:
'   - `ParsePlaceholderDefinitions` - reads `;  [[ name ]] = value` comments from local file
'   - `ExtractPlaceholders` / `ContainsPlaceholders` - detects `[[ ... ]]` patterns
'   - `ApplyPlaceholders` - substitutes placeholders with stored definitions
'   - `FindMissingPlaceholders` - identifies placeholders without local definitions
'   - `ShowPlaceholderInputDialog` - interactive prompt for missing values
'   - `GeneratePendingUpdatesBlock` - creates commented block for silent mode fallback
'   - `WriteNewPlaceholderDefinitions` / `WritePendingUpdatesBlock` - file writers
'
' Supported INI Targets
' ---------------------
' - Main INI: `GetDefaultINIPath(context.RDV)` with `_iniUpdateContext.INI_UpdateSource`
' - Alternate model INI: `context.INI_AlternateModelPath` (segmented; per-segment UpdateSource)
' - Special service INI: `context.INI_SpecialServicePath` (segmented; per-segment UpdateSource)
'
' Update Source Format
' --------------------
' Format: `path; keys; publicKey`
' Keys: `all`, `new`, explicit keys, `-key` exclusions
'
' Change Detection
' ----------------
' - `CheckSingleIniFile` - non-segmented files with placeholder support
' - `CheckSegmentedIniFile` - segmented files with per-segment placeholder support
'
' Signature Verification (Ed25519)
' --------------------------------
' - `VerifyEd25519Signature` - core verifier
' - `GenerateEd25519KeyPair` / `SignUpdateFile` / `VerifySignatureFile` - admin tools
' - `ShowSignatureManagementDialog` / `ShowBatchSigningDialog` - UI tools
'
' Silent vs Interactive Modes
' ---------------------------
' - `SilentUpdateSecurityLevel` enum (Disabled, SafeOnly, SignedOnly, LocalTrusted, All)
' - `ProcessSilentUpdates` - applies updates without UI based on security level
' - `ShowUpdateApprovalDialog` / `ShowIgnoreConfirmationDialog` - interactive UI
'
' Ignore List Management
' ----------------------
' - `ParseIgnoreOverrideRules` / `FilterIgnoredParameters` - rule parsing and filtering
' - `GetIgnoreListForFile` / `SaveIgnoreListForFile` / `AddToIgnoreList` - persistence
' - `ShowIgnoredParametersDialog` - UI for managing ignored parameters
'
' Backup and Apply
' ----------------
' - `CreateTimestampedBackup` - creates backup before modification
' - `ApplyApprovedUpdates` - writes changes to disk
'
' Client Authorization
' --------------------
' - `GetCurrentClientIdentifier` / `IsClientAllowedToUpdate`
'
' Logging
' -------
' - `LogIniUpdateEvent` - central logger with forced security event logging
'
' Dialog Positioning
' ------------------
' - `GetDialogOwner` / `CenterFormOnOwnerScreen` - ensures dialogs appear on correct monitor
'
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Text
Imports System.Windows.Forms
Imports Org.BouncyCastle.Crypto.Generators
Imports Org.BouncyCastle.Crypto.Parameters
Imports Org.BouncyCastle.Crypto.Signers
Imports Org.BouncyCastle.Security
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary
    Partial Public Class SharedMethods

        ''' <summary>
        ''' Freestyle command name used to manage ignored INI update parameters.
        ''' </summary>
        Private Const IniUpdateIgnored As String = "iniupdateignored"   ' the name of the Freestyle command

        ''' <summary>
        ''' Stores the active shared context for the current INI update check run.
        ''' </summary>
        Private Shared _iniUpdateContext As ISharedContext

        ''' <summary>
        ''' Indicates that the user temporarily unlocked the expert/central configuration in the
        ''' current settings session (e.g. via the central configuration password) even though
        ''' NoLocalConfig is active. When True, <see cref="CheckForIniUpdates"/> may offer to check
        ''' for INI updates even if the client is not in the allowed update client list.
        ''' </summary>
        Public Shared NoLocalConfigSessionUnlocked As Boolean = False

        ''' <summary>
        ''' Defines the security level for silent (unattended) updates.
        ''' </summary>
        Public Enum SilentUpdateSecurityLevel
            Disabled = 0
            SafeOnly = 1
            SignedOnly = 2
            LocalTrusted = 3
            All = 4
        End Enum

        ''' <summary>
        ''' Gets the configured silent update mode as a typed <see cref="SilentUpdateSecurityLevel"/> value.
        ''' </summary>
        Private Shared ReadOnly Property SilentMode As SilentUpdateSecurityLevel
            Get
                Return CType(_iniUpdateContext.INI_UpdateIniSilentMode, SilentUpdateSecurityLevel)
            End Get
        End Property

        ' Overrides local ignore settings with file-specific and segment-specific rules.
        ' Format: Comma-separated rules with + (ignore) or - (include) prefix
        ' 
        ' Rule formats:
        '   +Key or -Key                    → Match parameter in any file/segment
        '   +file.ini|Key                   → Match parameter in specific file (any segment)
        '   +file.ini|Segment|Key           → Match parameter in specific file and segment
        '   +*|Segment|Key                  → Match parameter in specific segment of any file
        '   +file.ini|*|Key                 → Same as +file.ini|Key
        '
        ' Special values:
        '   +all                            → Ignore all updates globally
        '   -all                            → Clear all ignores, process all updates
        '
        ' Examples:
        '   "-all"                          → Process all updates (default)
        '   "+all,-redink.ini|ApiKey"       → Ignore all except ApiKey in redink.ini
        '   "+allmodels.ini|*|Model"        → Ignore Model parameter in all segments of allmodels.ini
        '   "-redink.ini|Endpoint,+Model"   → Force Endpoint in redink.ini, ignore Model everywhere

#Region "Data Structures"

        ''' <summary>
        ''' Represents a single parameter change detected during update check.
        ''' </summary>
        Public Class IniParameterChange
            Public Property IniFile As String           ' "redink.ini", "allmodels.ini", "specialservices.ini"
            Public Property SegmentName As String       ' Empty for redink.ini, segment name for others
            Public Property ParameterKey As String      ' The key name (e.g., "Model", "Endpoint")
            Public Property OldValue As String          ' Current value in local file
            Public Property NewValue As String          ' Value from remote source
            Public Property IsSelected As Boolean       ' User's approval choice
            Public Property IsSuspicious As Boolean     ' True if contains URL/path that changed
            Public Property RemoteTemplateValue As String ' The original remote value with placeholders (before substitution).
            Public Property Placeholders As List(Of String)  ' List of placeholders found in the remote value.

            Public Overrides Function ToString() As String
                Return $"[{IniFile}]{If(String.IsNullOrEmpty(SegmentName), "", $"[{SegmentName}]")}.{ParameterKey}"
            End Function
        End Class

        ''' <summary>
        ''' Represents an INI file segment with its parameters and update source.
        ''' </summary>
        Public Class IniSegment
            Public Property Name As String                                      ' Segment name (from [Name])
            Public Property Parameters As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Public Property UpdateSource As String                              ' Full UpdateSource value
            Public Property UpdatePath As String                                ' Extracted path/URL
            Public Property UpdateKeys As List(Of String)                       ' Keys to update ("all" or specific list)
            Public Property PublicKey As String                                 ' Base64 public key for signature
        End Class

        ''' <summary>
        ''' Represents a parsed INI file with optional segments.
        ''' </summary>
        Public Class ParsedIniFile
            Public Property FilePath As String
            Public Property FileName As String
            Public Property GlobalParameters As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Public Property Segments As New List(Of IniSegment)()
            Public Property RawContent As String
        End Class

#End Region

#Region "Keys That Cannot Be Updated"

        ''' <summary>
        ''' Returns True if the specified key should never be updated automatically.
        ''' </summary>
        Private Shared Function IsProtectedKey(key As String) As Boolean
            If String.IsNullOrWhiteSpace(key) Then Return True
            ' Keys starting with "Update" are protected
            Return key.Trim().StartsWith("Update", StringComparison.OrdinalIgnoreCase)
        End Function

#End Region


#Region "Main Entry Point"

        ''' <summary>
        ''' Determines whether an INI-governed update workflow may run for the current context.
        ''' Uses the same master switch, client allow-list and NoLocalConfig session override as the
        ''' INI update flow so related update checks stay aligned.
        ''' </summary>
        Friend Shared Function CanRunIniGovernedUpdate(ByRef context As ISharedContext,
                                                       updateDisplayName As String) As Boolean
            If context Is Nothing Then
                Return False
            End If

            _iniUpdateContext = context

            If Not _iniUpdateContext.INI_UpdateIni Then
                Debug.WriteLine($"{updateDisplayName} Update: Disabled via _iniUpdateContext.INI_UpdateIni")
                Return False
            End If

            If IsClientAllowedToUpdate() Then
                Return True
            End If

            If NoLocalConfigSessionUnlocked Then
                Dim overrideAnswer As Integer = ShowCustomYesNoBox(
                    $"Your client is not allowed to check for or apply {updateDisplayName} updates. " &
                    "However, because you unlocked the expert configuration with the central configuration password, " &
                    $"you may still check for {updateDisplayName} updates now. Do you want to check for updates?",
                    $"Yes, check for {updateDisplayName} updates", "No")
                If overrideAnswer <> 1 Then
                    Debug.WriteLine($"{updateDisplayName} Update: Client not authorized; user declined the override check")
                    Return False
                End If
                LogIniUpdateEvent("Client Check",
                    $"Client not authorized, but expert configuration was unlocked in this session - user chose to check for {updateDisplayName} updates")
            Else
                Debug.WriteLine($"{updateDisplayName} Update: This client is not authorized to perform updates")
                Return False
            End If

            Return True
        End Function

        ''' <summary>
        ''' Main entry point for INI update checking. Called from UpdateHandler at startup.
        ''' </summary>
        ''' <param name="context">The shared context containing configuration.</param>
        ''' <returns>True if updates were applied, False otherwise.</returns>
        Public Shared Function CheckForIniUpdates(ByRef context As ISharedContext,
                                                  Optional skipAuthorizationGate As Boolean = False) As Boolean

            ' Store context for use by helper methods
            _iniUpdateContext = context

            Try
                If Not skipAuthorizationGate AndAlso Not CanRunIniGovernedUpdate(context, "INI configuration") Then
                    Return False
                End If

                ' Check registry for silent update permission - override silent mode if not permitted
                ' Only enforce registry check if noSilentIniUpdatesWithoutRegistryFlag is True
                If _iniUpdateContext.INI_UpdateIniSilentMode <> SilentUpdateSecurityLevel.Disabled AndAlso
                   noSilentIniUpdatesWithoutRegistryFlag Then
                    Dim permitSilentUpdates As String = GetFromRegistry(RegPath_Base, RegPath_PermitSilentInitUpdates, True)
                    If String.IsNullOrWhiteSpace(permitSilentUpdates) OrElse
                       Not (permitSilentUpdates.Equals("yes", StringComparison.OrdinalIgnoreCase) OrElse
                            permitSilentUpdates.Equals("true", StringComparison.OrdinalIgnoreCase) OrElse
                            permitSilentUpdates.Equals("1", StringComparison.OrdinalIgnoreCase)) Then
                        Debug.WriteLine("INI Update: Silent mode disabled via registry (PermitSilentIniUpdates not set or false)")
                        LogIniUpdateEvent("Silent Mode Override", "Registry key PermitSilentIniUpdates not set or false - forcing interactive mode")
                        _iniUpdateContext.INI_UpdateIniSilentMode = SilentUpdateSecurityLevel.Disabled
                    End If
                End If

                LogIniUpdateEvent("Check Started", $"SilentMode={_iniUpdateContext.INI_UpdateIniSilentMode}")

                ' Collect all changes from all three INI files
                Dim allChanges As New List(Of IniParameterChange)()
                Dim signatureErrors As New List(Of String)()
                Dim failedSources As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                Dim updateSources As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

                ' 1. Check redink.ini (main config)
                Dim mainIniPath As String = GetDefaultINIPath(context.RDV)
                If File.Exists(mainIniPath) AndAlso Not String.IsNullOrWhiteSpace(_iniUpdateContext.INI_UpdateSource) Then
                    Dim mainFileName = Path.GetFileName(mainIniPath)
                    Dim sourceInfo = ParseGlobalUpdateSource(_iniUpdateContext.INI_UpdateSource)
                    updateSources(mainFileName) = sourceInfo.UpdatePath
                    Dim mainChanges = CheckSingleIniFile(mainIniPath, mainFileName, _iniUpdateContext.INI_UpdateSource, Nothing, signatureErrors, failedSources)
                    If mainChanges IsNot Nothing Then allChanges.AddRange(mainChanges)
                End If

                ' 2. Check AlternateModelPath (user-defined filename)
                If Not String.IsNullOrWhiteSpace(context.INI_AlternateModelPath) Then
                    Dim altPath = ExpandEnvironmentVariables(context.INI_AlternateModelPath)
                    If File.Exists(altPath) Then
                        Dim altFileName = Path.GetFileName(altPath)
                        Dim altChanges = CheckSegmentedIniFile(altPath, altFileName, signatureErrors, updateSources, failedSources)
                        If altChanges IsNot Nothing Then allChanges.AddRange(altChanges)
                    End If
                End If

                ' 3. Check SpecialServicePath (user-defined filename)
                If Not String.IsNullOrWhiteSpace(context.INI_SpecialServicePath) Then
                    Dim svcPath = ExpandEnvironmentVariables(context.INI_SpecialServicePath)
                    If File.Exists(svcPath) Then
                        Dim svcFileName = Path.GetFileName(svcPath)
                        Dim svcChanges = CheckSegmentedIniFile(svcPath, svcFileName, signatureErrors, updateSources, failedSources)
                        If svcChanges IsNot Nothing Then allChanges.AddRange(svcChanges)
                    End If
                End If

                ' Handle signature errors - ALWAYS log security events
                If signatureErrors.Count > 0 Then
                    LogIniUpdateEvent("SECURITY: Signature Errors",
                        $"{signatureErrors.Count} signature error(s) detected:" & vbCrLf &
                        String.Join(vbCrLf, signatureErrors), alwaysLog:=True)

                    If _iniUpdateContext.INI_UpdateIniSilentMode = SilentUpdateSecurityLevel.Disabled Then
                        ShowSignatureErrorDialog(signatureErrors)
                    End If
                End If

                ' Filter out ignored parameters
                allChanges = FilterIgnoredParameters(allChanges)

                ' Check for missing placeholders
                Dim changesWithMissingPlaceholders = allChanges.Where(
                    Function(c) c.Placeholders IsNot Nothing AndAlso
                                ContainsPlaceholders(c.NewValue)).ToList()

                ' Determine blocked scopes (file + segment)
                Dim blockedScopes As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                For Each c As IniParameterChange In changesWithMissingPlaceholders
                    blockedScopes.Add($"{c.IniFile}|{If(c.SegmentName, "")}")
                Next


                If changesWithMissingPlaceholders.Count > 0 Then
                    ' Collect all missing placeholder info
                    Dim missingPlaceholderInfos As New List(Of MissingPlaceholderInfo)()
                    Dim seenPlaceholders As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                    For Each change In changesWithMissingPlaceholders
                        Dim missingInChange = ExtractPlaceholders(change.NewValue)
                        For Each placeholder In missingInChange
                            Dim key = $"{change.IniFile}|{change.SegmentName}|{placeholder}"
                            If Not seenPlaceholders.Contains(key) Then
                                seenPlaceholders.Add(key)
                                missingPlaceholderInfos.Add(New MissingPlaceholderInfo() With {
                                    .Placeholder = placeholder,
                                    .SegmentName = change.SegmentName,
                                    .FileName = change.IniFile,
                                    .SampleContext = $"{change.ParameterKey} = {change.RemoteTemplateValue}"
                                })
                            End If
                        Next
                    Next

                    If _iniUpdateContext.INI_UpdateIniSilentMode <> SilentUpdateSecurityLevel.Disabled Then
                        ' SILENT MODE: Write pending updates as comments
                        ' Group by file and segment
                        Dim byFileAndSegment = changesWithMissingPlaceholders.GroupBy(
                            Function(c) $"{c.IniFile}|{c.SegmentName}")

                        For Each group In byFileAndSegment
                            Dim parts = group.Key.Split("|"c)
                            Dim fileNamePart = parts(0)
                            Dim segmentPart = If(parts.Length > 1, parts(1), "")

                            ' Get all changes for this file/segment (not just those with missing placeholders)
                            Dim allChangesForSegment = allChanges.Where(
                                Function(c) c.IniFile.Equals(fileNamePart, StringComparison.OrdinalIgnoreCase) AndAlso
                                            (c.SegmentName = segmentPart OrElse
                                             (String.IsNullOrEmpty(c.SegmentName) AndAlso String.IsNullOrEmpty(segmentPart)))).ToList()

                            ' Collect missing placeholders for this segment
                            Dim missingForSegment = missingPlaceholderInfos.Where(
                                Function(m) m.FileName.Equals(fileNamePart, StringComparison.OrdinalIgnoreCase) AndAlso
                                            (m.SegmentName = segmentPart OrElse
                                             (String.IsNullOrEmpty(m.SegmentName) AndAlso String.IsNullOrEmpty(segmentPart)))).
                                Select(Function(m) m.Placeholder).Distinct().ToList()

                            ' Resolve file path
                            Dim filePath = ResolveIniFilePath(fileNamePart, context)
                            If Not String.IsNullOrEmpty(filePath) Then
                                Dim pendingBlock = GeneratePendingUpdatesBlock(allChangesForSegment, missingForSegment, segmentPart)
                                WritePendingUpdatesBlock(filePath, segmentPart, pendingBlock)

                                ' Remove ALL changes for this segment from the apply list                                
                                allChanges.RemoveAll(
                                    Function(c)
                                        blockedScopes.Contains($"{c.IniFile}|{If(c.SegmentName, "")}")
                                    End Function
                                )

                            End If
                        Next

                        LogIniUpdateEvent("Missing Placeholders",
                            $"Silent mode: {changesWithMissingPlaceholders.Count} change(s) written as pending due to missing placeholders",
                            alwaysLog:=True)
                    Else
                        ' INTERACTIVE MODE: Prompt user for placeholder values
                        If ShowPlaceholderInputDialog(missingPlaceholderInfos) Then
                            ' User provided values - apply them
                            ' Group by file and segment to write definitions
                            Dim byFileAndSegment = missingPlaceholderInfos.GroupBy(
                                Function(m) $"{m.FileName}|{m.SegmentName}")

                            For Each group In byFileAndSegment
                                Dim parts = group.Key.Split("|"c)
                                Dim fileNamePart = parts(0)
                                Dim segmentPart = If(parts.Length > 1, parts(1), "")

                                Dim filePath = ResolveIniFilePath(fileNamePart, context)
                                If Not String.IsNullOrEmpty(filePath) Then
                                    ' Write placeholder definitions
                                    Dim newDefs As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                                    For Each mp In group
                                        If Not String.IsNullOrEmpty(mp.UserValue) Then
                                            newDefs(mp.Placeholder) = mp.UserValue
                                        End If
                                    Next

                                    If newDefs.Count > 0 Then
                                        WriteNewPlaceholderDefinitions(filePath, segmentPart, newDefs)

                                        ' Update change NewValue with substituted values
                                        For Each change In allChanges.Where(
                                            Function(c) c.IniFile.Equals(fileNamePart, StringComparison.OrdinalIgnoreCase) AndAlso
                                                        (c.SegmentName = segmentPart OrElse
                                                         (String.IsNullOrEmpty(c.SegmentName) AndAlso String.IsNullOrEmpty(segmentPart))))
                                            If ContainsPlaceholders(change.NewValue) Then
                                                change.NewValue = ApplyPlaceholders(change.NewValue, newDefs)
                                            End If
                                        Next
                                    End If
                                End If
                            Next
                        Else
                            ' User cancelled - skip these changes entirely                            
                            allChanges.RemoveAll(
                                        Function(c)
                                            blockedScopes.Contains($"{c.IniFile}|{If(c.SegmentName, "")}")
                                        End Function
                                    )
                            LogIniUpdateEvent("Placeholder Input",
                                $"User skipped {changesWithMissingPlaceholders.Count} change(s) requiring placeholder values")
                        End If
                    End If
                End If

                If allChanges.Count = 0 Then
                    Debug.WriteLine("INI Update: No changes detected")
                    LogIniUpdateEvent("Check Complete", "No changes detected")
                    Return False
                End If

                ' Log detected changes summary
                Dim suspiciousCount As Integer = allChanges.Where(Function(c) c.IsSuspicious).Count()
                LogIniUpdateEvent("Changes Detected",
                    $"Total={allChanges.Count}, Suspicious={suspiciousCount}")

                ' === SILENT MODE HANDLING ===
                If _iniUpdateContext.INI_UpdateIniSilentMode <> SilentUpdateSecurityLevel.Disabled Then
                    Return ProcessSilentUpdates(allChanges, context, failedSources, updateSources)
                End If

                ' === INTERACTIVE MODE (existing behavior) ===
                ' Show approval dialog
                Dim approvalResult = ShowUpdateApprovalDialog(allChanges)
                If approvalResult = UpdateApprovalResult.Reject Then
                    LogIniUpdateEvent("User Action", "User rejected all changes")
                    ShowIgnoreConfirmationDialog(allChanges)
                    Return False
                ElseIf approvalResult = UpdateApprovalResult.Cancel Then
                    LogIniUpdateEvent("User Action", "User cancelled update dialog")
                    Return False
                End If

                ' Get approved and rejected changes
                Dim approvedChanges = allChanges.Where(Function(c) c.IsSelected).ToList()
                Dim rejectedChanges = allChanges.Where(Function(c) Not c.IsSelected).ToList()

                ' Log user decisions
                If approvedChanges.Count > 0 OrElse rejectedChanges.Count > 0 Then
                    LogIniUpdateEvent("User Action",
                        $"Approved={approvedChanges.Count}, Rejected={rejectedChanges.Count}")
                End If

                ' Show ignore dialog for rejected items
                If rejectedChanges.Count > 0 Then
                    ShowIgnoreConfirmationDialog(rejectedChanges)
                End If

                ' Apply approved changes
                If approvedChanges.Count > 0 Then
                    ApplyApprovedUpdates(approvedChanges, context)
                    LogIniUpdateEvent("Updates Applied",
                        String.Join(vbCrLf, approvedChanges.Select(Function(c) $"{c.ToString()}: {c.OldValue} → {c.NewValue}")),
                        alwaysLog:=True)
                    ShowCustomMessageBox($"{approvedChanges.Count} configuration parameter(s) have been updated. Changes will be active upon next reload.")
                    Return True
                End If

                Return False

            Catch ex As Exception
                Debug.WriteLine($"INI Update Error: {ex.Message}")
                LogIniUpdateEvent("ERROR", $"Unexpected error: {ex.Message}", alwaysLog:=True)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Resolves an INI file name to its full path using the context.
        ''' </summary>
        Private Shared Function ResolveIniFilePath(fileName As String, context As ISharedContext) As String
            ' Check main INI
            Dim mainIniPath = GetDefaultINIPath(context.RDV)
            If Path.GetFileName(mainIniPath).Equals(fileName, StringComparison.OrdinalIgnoreCase) Then
                Return mainIniPath
            End If

            ' Check AlternateModelPath
            If Not String.IsNullOrWhiteSpace(context.INI_AlternateModelPath) Then
                Dim altPath = ExpandEnvironmentVariables(context.INI_AlternateModelPath)
                If Path.GetFileName(altPath).Equals(fileName, StringComparison.OrdinalIgnoreCase) Then
                    Return altPath
                End If
            End If

            ' Check SpecialServicePath
            If Not String.IsNullOrWhiteSpace(context.INI_SpecialServicePath) Then
                Dim svcPath = ExpandEnvironmentVariables(context.INI_SpecialServicePath)
                If Path.GetFileName(svcPath).Equals(fileName, StringComparison.OrdinalIgnoreCase) Then
                    Return svcPath
                End If
            End If

            Return Nothing
        End Function

        ''' <summary>
        ''' Defines the outcome of the update approval dialog.
        ''' </summary>
        Private Enum UpdateApprovalResult
            ''' <summary>Approve the selected changes.</summary>
            Approve
            ''' <summary>Reject all changes.</summary>
            Reject
            ''' <summary>Cancel the dialog without applying changes.</summary>
            Cancel
        End Enum

        ''' <summary>
        ''' Determines if a source path is local (file system or network share, not HTTPS).
        ''' </summary>
        Private Shared Function IsLocalSource(sourcePath As String) As Boolean
            If String.IsNullOrWhiteSpace(sourcePath) Then Return False

            ' Remote sources start with http:// or https://
            Return Not sourcePath.StartsWith("http://", StringComparison.OrdinalIgnoreCase) AndAlso
           Not sourcePath.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
        End Function


        ''' <summary>
        ''' Applies updates without UI interaction according to the configured silent update security level.
        ''' </summary>
        ''' <param name="changes">The detected parameter changes.</param>
        ''' <param name="context">The shared context used to resolve local INI file locations.</param>
        ''' <param name="failedSources">Update sources that failed signature validation or required verification data.</param>
        ''' <param name="updateSources">Mapping of INI file or INI file/segment identifiers to their update source path.</param>
        ''' <returns><c>True</c> if at least one change was applied; otherwise <c>False</c>.</returns>
        Private Shared Function ProcessSilentUpdates(changes As List(Of IniParameterChange),
                                      context As ISharedContext,
                                      failedSources As HashSet(Of String),
                                      updateSources As Dictionary(Of String, String)) As Boolean
            Dim changesToApply As New List(Of IniParameterChange)()

            Select Case _iniUpdateContext.INI_UpdateIniSilentMode
                Case SilentUpdateSecurityLevel.SafeOnly
                    changesToApply = changes.Where(Function(c) Not c.IsSuspicious).ToList()
                    LogIniUpdateEvent("SafeOnly mode",
                $"Applying {changesToApply.Count} safe changes, skipping {changes.Count - changesToApply.Count} suspicious changes")

                Case SilentUpdateSecurityLevel.SignedOnly
                    ' Only apply changes from sources that passed signature validation
                    For Each change As IniParameterChange In changes
                        Dim sourcePath As String = Nothing

                        ' Get source path for this change
                        If Not String.IsNullOrEmpty(change.SegmentName) Then
                            Dim segmentKey As String = $"{change.IniFile}|{change.SegmentName}"
                            If updateSources IsNot Nothing Then
                                updateSources.TryGetValue(segmentKey, sourcePath)
                            End If
                        End If

                        If sourcePath Is Nothing AndAlso updateSources IsNot Nothing Then
                            updateSources.TryGetValue(change.IniFile, sourcePath)
                        End If

                        ' Only add if source didn't fail signature validation
                        If String.IsNullOrWhiteSpace(sourcePath) OrElse Not failedSources.Contains(sourcePath) Then
                            changesToApply.Add(change)
                        End If
                    Next

                    Dim skippedCount = changes.Count - changesToApply.Count
                    If skippedCount > 0 Then
                        LogIniUpdateEvent("SignedOnly mode",
                    $"Applying {changesToApply.Count} changes from valid sources, skipping {skippedCount} from sources with signature errors",
                    alwaysLog:=True)
                    Else
                        LogIniUpdateEvent("SignedOnly mode", $"All signatures valid - applying {changesToApply.Count} changes")
                    End If

                Case SilentUpdateSecurityLevel.LocalTrusted
                    For Each change As IniParameterChange In changes
                        Dim sourcePath As String = Nothing

                        If Not String.IsNullOrEmpty(change.SegmentName) Then
                            Dim segmentKey As String = $"{change.IniFile}|{change.SegmentName}"
                            If updateSources IsNot Nothing Then
                                updateSources.TryGetValue(segmentKey, sourcePath)
                            End If
                        End If

                        If sourcePath Is Nothing AndAlso updateSources IsNot Nothing Then
                            updateSources.TryGetValue(change.IniFile, sourcePath)
                        End If

                        If Not String.IsNullOrWhiteSpace(sourcePath) AndAlso IsLocalSource(sourcePath) Then
                            changesToApply.Add(change)
                        End If
                    Next

                    If changesToApply.Count > 0 Then
                        LogIniUpdateEvent("LocalTrusted mode",
                            $"Applying {changesToApply.Count} changes from local/network sources, " &
                            $"skipping {changes.Count - changesToApply.Count} remote changes")
                    Else
                        LogIniUpdateEvent("LocalTrusted mode", "No local source changes to apply")
                    End If

                Case SilentUpdateSecurityLevel.All
                    changesToApply = changes.ToList()
                    LogIniUpdateEvent("All mode", $"Applying all {changesToApply.Count} changes (including suspicious and remote)")

                Case Else
                    Return False
            End Select

            If changesToApply.Count = 0 Then
                Return False
            End If

            For Each change As IniParameterChange In changesToApply
                change.IsSelected = True
            Next

            ApplyApprovedUpdates(changesToApply, context)
            LogIniUpdateEvent("Updates Applied",
                String.Join(vbCrLf, changesToApply.Select(Function(c) $"{c.ToString()}: {c.OldValue} → {c.NewValue}")),
                alwaysLog:=True)

            Return True
        End Function

        ''' <summary>
        ''' Writes an INI update event to the update log (subject to logging settings unless forced).
        ''' </summary>
        ''' <param name="eventType">Short event category string.</param>
        ''' <param name="details">Event details text.</param>
        ''' <param name="alwaysLog">If <c>True</c>, logs even when silent logging is disabled.</param>
        Private Shared Sub LogIniUpdateEvent(eventType As String, details As String, Optional alwaysLog As Boolean = False)
            ' For silent mode, respect the INI_UpdateIniSilentLog setting unless alwaysLog is True
            If Not alwaysLog AndAlso Not _iniUpdateContext.INI_UpdateIniSilentLog Then Return

            Try
                Dim message As String = $"[INI Update] [{eventType}]"
                If Not String.IsNullOrWhiteSpace(details) Then
                    message &= vbCrLf & "  " & details.Replace(vbCrLf, vbCrLf & "  ")
                End If
                UpdateHandler.WriteUpdateLog(message)
            Catch ex As Exception
                Debug.WriteLine($"Failed to log INI update event: {ex.Message}")
            End Try
        End Sub

#End Region

#Region "Signature Error Reporting"

        ''' <summary>
        ''' Represents a signature validation error for reporting.
        ''' </summary>
        Public Class SignatureError
            Public Property SourcePath As String
            Public Property ErrorType As SignatureErrorType
            Public Property Details As String

            Public Overrides Function ToString() As String
                Return $"[{ErrorType}] {SourcePath}: {Details}"
            End Function
        End Class

        ''' <summary>
        ''' Defines the type of signature validation error for reporting.
        ''' </summary>
        Public Enum SignatureErrorType
            ''' <summary>The signature file (.sig) could not be found.</summary>
            SignatureFileMissing
            ''' <summary>No public key was provided for signature verification.</summary>
            PublicKeyMissing
            ''' <summary>The provided signature did not match the content and public key.</summary>
            SignatureInvalid
            ''' <summary>Downloading the update source or signature failed.</summary>
            DownloadFailed
            ''' <summary>Any other signature-related failure.</summary>
            Other
        End Enum

        ''' <summary>
        ''' Displays signature validation errors to the user in interactive mode.
        ''' </summary>
        ''' <param name="errors">The formatted error messages to display.</param>
        Private Shared Sub ShowSignatureErrorDialog(errors As List(Of String))
            If errors Is Nothing OrElse errors.Count = 0 Then Return

            Dim form As New Form() With {
                .Text = $"{AN} - Signature Validation Errors",
                .Size = New Size(850, 520),
                .StartPosition = FormStartPosition.CenterParent,
                .FormBorderStyle = FormBorderStyle.Sizable,
                .Font = New Font("Segoe UI", 9.0F),
                .MinimumSize = New Size(750, 450),
                .AutoScaleMode = AutoScaleMode.Dpi,
                .TopMost = True
            }


            Try
                Dim bmp As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                form.Icon = Icon.FromHandle(bmp.GetHicon())
            Catch
            End Try

            ' Use TableLayoutPanel for proper scaling and layout
            Dim tblMain As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(15),
                .ColumnCount = 1,
                .RowCount = 4,
                .AutoSize = False
            }
            tblMain.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            tblMain.RowStyles.Add(New RowStyle(SizeType.AutoSize))  ' Warning header
            tblMain.RowStyles.Add(New RowStyle(SizeType.Percent, 100))  ' Error list (fills remaining space)
            tblMain.RowStyles.Add(New RowStyle(SizeType.AutoSize))  ' Info label
            tblMain.RowStyles.Add(New RowStyle(SizeType.AutoSize))  ' Buttons
            form.Controls.Add(tblMain)

            ' Row 0: Warning header
            Dim pnlHeader As New Panel() With {
                .Dock = DockStyle.Fill,
                .BackColor = Color.FromArgb(255, 250, 230),
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .MinimumSize = New Size(0, 60)
            }
            tblMain.Controls.Add(pnlHeader, 0, 0)

            Dim lblWarning As New Label() With {
                .Text = "⚠ SECURITY WARNING: The following update sources could not be verified." & vbCrLf &
                        "This may indicate tampering, misconfiguration, or missing signature files.",
                .Dock = DockStyle.Fill,
                .Font = New Font("Segoe UI", 10.0F, FontStyle.Bold),
                .ForeColor = Color.DarkOrange,
                .TextAlign = ContentAlignment.MiddleLeft,
                .Padding = New Padding(10),
                .AutoSize = True
            }
            pnlHeader.Controls.Add(lblWarning)

            ' Row 1: Error list (takes remaining space)
            Dim txtErrors As New TextBox() With {
                .Dock = DockStyle.Fill,
                .Multiline = True,
                .ReadOnly = True,
                .ScrollBars = ScrollBars.Both,
                .BackColor = SystemColors.Window,
                .Font = New Font("Consolas", 9.0F),
                .Text = String.Join(vbCrLf & vbCrLf, errors),
                .Margin = New Padding(0, 10, 0, 10)
            }
            tblMain.Controls.Add(txtErrors, 0, 1)

            ' Row 2: Info label (auto-sized)
            Dim lblInfo As New Label() With {
                .Text = "These update sources were skipped. Contact your administrator if this is unexpected." & vbCrLf &
                        "Administrators: Use the Signature Management tool to diagnose and fix signature issues.",
                .Dock = DockStyle.Fill,
                .AutoSize = True,
                .Padding = New Padding(0, 5, 0, 5),
                .BackColor = Color.FromArgb(240, 240, 240),
                .Margin = New Padding(0, 0, 0, 10)
            }
            tblMain.Controls.Add(lblInfo, 0, 2)

            ' Row 3: Buttons panel
            Dim pnlButtons As New FlowLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.LeftToRight,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .WrapContents = False,
                .Margin = New Padding(0, 0, 0, 5)
            }
            tblMain.Controls.Add(pnlButtons, 0, 3)

            Dim btnClose As New Button() With {
                .Text = "Close",
                .AutoSize = True,
                .Padding = New Padding(10, 5, 10, 5),
                .Margin = New Padding(0, 0, 10, 0)
            }
            Dim btnCopy As New Button() With {
                .Text = "Copy to Clipboard",
                .AutoSize = True,
                .Padding = New Padding(10, 5, 10, 5),
                .Margin = New Padding(0, 0, 10, 0)
            }
            Dim btnDiagnose As New Button() With {
                .Text = "Open Signature Tool...",
                .AutoSize = True,
                .Padding = New Padding(10, 5, 10, 5)
            }

            pnlButtons.Controls.Add(btnClose)
            pnlButtons.Controls.Add(btnCopy)
            pnlButtons.Controls.Add(btnDiagnose)

            AddHandler btnClose.Click, Sub() form.Close()

            AddHandler btnCopy.Click, Sub()
                                          Try
                                              Dim report = $"=== {AN} Signature Validation Report ==={vbCrLf}" &
                                                           $"Date: {DateTime.Now:yyyy-MM-dd HH:mm:ss}{vbCrLf}{vbCrLf}" &
                                                           String.Join(vbCrLf & vbCrLf, errors)
                                              Clipboard.SetText(report)
                                              ShowCustomMessageBox("Error report copied to clipboard.")
                                          Catch ex As Exception
                                              ShowCustomMessageBox("Failed to copy to clipboard: " & ex.Message)
                                          End Try
                                      End Sub

            AddHandler btnDiagnose.Click, Sub()
                                              form.Close()
                                              ShowSignatureManagementDialog()
                                          End Sub

            form.ShowDialog()
        End Sub

#End Region

#Region "INI File Parsing"

        ''' <summary>
        ''' Parses an INI file from disk into global parameters and named segments.
        ''' </summary>
        ''' <param name="filePath">Path to the INI file.</param>
        ''' <returns>A parsed representation of the INI file; empty if the file does not exist or parsing fails.</returns>
        Private Shared Function ParseIniFile(filePath As String) As ParsedIniFile
            Dim result As New ParsedIniFile() With {
                .FilePath = filePath,
                .FileName = Path.GetFileName(filePath)
            }

            If Not File.Exists(filePath) Then Return result

            Try
                result.RawContent = File.ReadAllText(filePath, Encoding.UTF8)
                Dim lines = result.RawContent.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)

                Dim currentSegment As IniSegment = Nothing

                For Each rawLine In lines
                    Dim line = rawLine.Trim()

                    ' Skip empty lines and comments
                    If String.IsNullOrEmpty(line) OrElse line.StartsWith(";") Then Continue For

                    ' Check for segment header [Name]
                    If line.StartsWith("[") AndAlso line.EndsWith("]") Then
                        ' Save previous segment if exists
                        If currentSegment IsNot Nothing Then
                            ParseUpdateSource(currentSegment)
                            result.Segments.Add(currentSegment)
                        End If

                        ' Start new segment
                        currentSegment = New IniSegment() With {
                            .Name = line.Substring(1, line.Length - 2).Trim()
                        }
                        Continue For
                    End If

                    ' Parse key = value
                    Dim eqIndex = line.IndexOf("="c)
                    If eqIndex > 0 Then
                        Dim key = line.Substring(0, eqIndex).Trim()
                        Dim value = line.Substring(eqIndex + 1).Trim()

                        If currentSegment IsNot Nothing Then
                            currentSegment.Parameters(key) = value
                        Else
                            result.GlobalParameters(key) = value
                        End If
                    End If
                Next

                ' Save last segment
                If currentSegment IsNot Nothing Then
                    ParseUpdateSource(currentSegment)
                    result.Segments.Add(currentSegment)
                End If

            Catch ex As Exception
                Debug.WriteLine($"Error parsing INI file {filePath}: {ex.Message}")
            End Try

            Return result
        End Function

        ''' <summary>
        ''' Parses the UpdateSource parameter of a segment into its components.
        ''' Format: "path; key1,key2,key3; base64_public_key" or "path; all; base64_public_key"
        ''' Supported key specifiers:
        '''   - "all" = update all keys from remote (including new keys)
        '''   - "new" = only propose new keys not in local file
        '''   - "key1,key2" = only update specific listed keys
        '''   - "-key1" = do not update specific listed key
        '''   - "all,new" or "key1,key2,new" = combine behaviors
        ''' </summary>
        Private Shared Sub ParseUpdateSource(segment As IniSegment)
            If segment Is Nothing Then Return
            If Not segment.Parameters.ContainsKey("UpdateSource") Then Return

            segment.UpdateSource = segment.Parameters("UpdateSource")
            Dim parts = segment.UpdateSource.Split(";"c)

            If parts.Length >= 1 Then
                segment.UpdatePath = parts(0).Trim()
            End If

            If parts.Length >= 2 Then
                Dim keysPart = parts(1).Trim()
                segment.UpdateKeys = keysPart.Split(","c).
                    Select(Function(k) k.Trim()).
                    Where(Function(k) Not String.IsNullOrEmpty(k)).
                    ToList()
            End If

            If parts.Length >= 3 Then
                segment.PublicKey = parts(2).Trim()
            End If
        End Sub

        ''' <summary>
        ''' Parses an <c>UpdateSource</c> string for a non-segmented INI file into an <see cref="IniSegment"/> container.
        ''' </summary>
        ''' <param name="updateSource">The raw update source string.</param>
        ''' <returns>An <see cref="IniSegment"/> containing the parsed update path, keys, and optional public key.</returns>
        Private Shared Function ParseGlobalUpdateSource(updateSource As String) As IniSegment
            Dim segment As New IniSegment() With {
                .Name = "",
                .UpdateSource = updateSource
            }

            If String.IsNullOrWhiteSpace(updateSource) Then Return segment

            Dim parts = updateSource.Split(";"c)

            If parts.Length >= 1 Then
                segment.UpdatePath = parts(0).Trim()
            End If

            If parts.Length >= 2 Then
                Dim keysPart = parts(1).Trim()
                segment.UpdateKeys = keysPart.Split(","c).
                    Select(Function(k) k.Trim()).
                    Where(Function(k) Not String.IsNullOrEmpty(k)).
                    ToList()
            End If

            If parts.Length >= 3 Then
                segment.PublicKey = parts(2).Trim()
            End If

            Return segment
        End Function

#End Region

#Region "Placeholder Preservation"

        ''' <summary>
        ''' Regex pattern to detect placeholders in values (e.g., [[ Your API Key ]], [[ Region ]])
        ''' </summary>
        Private Shared ReadOnly PlaceholderPattern As New System.Text.RegularExpressions.Regex(
            "\[\[\s*[^\]]+?\s*\]\]",
            System.Text.RegularExpressions.RegexOptions.Compiled)

        ''' <summary>
        ''' Parses placeholder definition comments from INI file lines.
        ''' Format: ; [[ Placeholder Name ]] = value
        ''' </summary>
        ''' <param name="lines">Lines from the INI file.</param>
        ''' <param name="segmentName">Segment name to scope the search, or empty for global scope.</param>
        ''' <returns>Dictionary mapping placeholder strings (including brackets) to their values.</returns>
        Private Shared Function ParsePlaceholderDefinitions(lines As IEnumerable(Of String), segmentName As String) As Dictionary(Of String, String)
            Dim definitions As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim inTargetSegment = String.IsNullOrEmpty(segmentName)
            Dim currentSegment As String = ""

            For Each rawLine In lines
                Dim line = rawLine.Trim()

                ' Track segment changes
                If line.StartsWith("[") AndAlso line.EndsWith("]") Then
                    currentSegment = line.Substring(1, line.Length - 2).Trim()
                    inTargetSegment = String.IsNullOrEmpty(segmentName) OrElse
                                      currentSegment.Equals(segmentName, StringComparison.OrdinalIgnoreCase)
                    Continue For
                End If

                ' Parse placeholder definition comments: ; [[ Name ]] = value
                If inTargetSegment AndAlso line.StartsWith(";") Then
                    Dim commentContent = line.Substring(1).Trim()

                    ' Check if this comment defines a placeholder
                    Dim match = PlaceholderPattern.Match(commentContent)
                    If match.Success AndAlso match.Index = 0 Then
                        ' Found a placeholder at the start of the comment
                        Dim placeholder = match.Value
                        Dim afterPlaceholder = commentContent.Substring(match.Length).Trim()

                        ' Check for " = value" pattern
                        If afterPlaceholder.StartsWith("=") Then
                            Dim value = afterPlaceholder.Substring(1).Trim()
                            ' Only store if value is non-empty
                            If Not String.IsNullOrEmpty(value) Then
                                definitions(placeholder) = value
                            End If
                        End If
                    End If
                End If

                ' Stop if we've passed the target segment in segmented mode
                If Not String.IsNullOrEmpty(segmentName) AndAlso
                   Not String.IsNullOrEmpty(currentSegment) AndAlso
                   Not currentSegment.Equals(segmentName, StringComparison.OrdinalIgnoreCase) AndAlso
                   inTargetSegment = False AndAlso definitions.Count > 0 Then
                    Exit For
                End If
            Next

            Return definitions
        End Function

        ''' <summary>
        ''' Determines if a value contains any placeholders.
        ''' </summary>
        Private Shared Function ContainsPlaceholders(value As String) As Boolean
            If String.IsNullOrWhiteSpace(value) Then Return False
            Return PlaceholderPattern.IsMatch(value)
        End Function

        ''' <summary>
        ''' Extracts all unique placeholder names from a value.
        ''' </summary>
        Private Shared Function ExtractPlaceholders(value As String) As List(Of String)
            Dim result As New List(Of String)()
            If String.IsNullOrWhiteSpace(value) Then Return result

            For Each match As System.Text.RegularExpressions.Match In PlaceholderPattern.Matches(value)
                If Not result.Contains(match.Value, StringComparer.OrdinalIgnoreCase) Then
                    result.Add(match.Value)
                End If
            Next

            Return result
        End Function

        ''' <summary>
        ''' Applies placeholder substitutions to a value.
        ''' </summary>
        ''' <param name="value">The value containing [[ placeholder ]] markers.</param>
        ''' <param name="definitions">Dictionary mapping placeholder strings to their values.</param>
        ''' <returns>The value with all known placeholders replaced.</returns>
        Private Shared Function ApplyPlaceholders(value As String, definitions As Dictionary(Of String, String)) As String
            If String.IsNullOrWhiteSpace(value) OrElse definitions Is Nothing OrElse definitions.Count = 0 Then
                Return value
            End If

            Dim result = value
            For Each kvp In definitions
                result = result.Replace(kvp.Key, kvp.Value)
            Next

            Return result
        End Function

        ''' <summary>
        ''' Finds placeholders in a value that are not defined in the given definitions.
        ''' </summary>
        Private Shared Function FindMissingPlaceholders(value As String, definitions As Dictionary(Of String, String)) As List(Of String)
            Dim missing As New List(Of String)()
            Dim allPlaceholders = ExtractPlaceholders(value)

            For Each placeholder In allPlaceholders
                If Not definitions.ContainsKey(placeholder) Then
                    missing.Add(placeholder)
                End If
            Next

            Return missing
        End Function

        ''' <summary>
        ''' Represents a missing placeholder that needs user input.
        ''' </summary>
        Private Class MissingPlaceholderInfo
            Public Property Placeholder As String       ' e.g., "[[ Your Tenant ]]"
            Public Property SegmentName As String       ' e.g., "GPT-4" or empty for global
            Public Property FileName As String          ' e.g., "allmodels.ini"
            Public Property SampleContext As String     ' e.g., "Endpoint = https://api.[[ Your Tenant ]].example.com"
            Public Property UserValue As String         ' User-provided value (filled by dialog)
        End Class

        ''' <summary>
        ''' Shows a dialog for the user to provide values for missing placeholders.
        ''' </summary>
        ''' <param name="missingPlaceholders">List of missing placeholder info.</param>
        ''' <returns>True if user provided values and clicked OK; False if cancelled.</returns>
        Private Shared Function ShowPlaceholderInputDialog(missingPlaceholders As List(Of MissingPlaceholderInfo)) As Boolean
            If missingPlaceholders Is Nothing OrElse missingPlaceholders.Count = 0 Then Return True

            Dim form As New Form() With {
                .Text = $"{AN} - Placeholder Values Required",
                .Size = New Size(700, 450),
                .StartPosition = FormStartPosition.CenterParent,
                .FormBorderStyle = FormBorderStyle.Sizable,
                .Font = New Font("Segoe UI", 9.0F),
                .MinimumSize = New Size(550, 350),
                .AutoScaleMode = AutoScaleMode.Dpi,
                .TopMost = True
            }



            Try
                Dim bmp As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                form.Icon = Icon.FromHandle(bmp.GetHicon())
            Catch
            End Try

            Dim tblMain As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(15),
                .ColumnCount = 1,
                .RowCount = 3,
                .AutoSize = False
            }
            tblMain.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            tblMain.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            tblMain.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
            tblMain.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            form.Controls.Add(tblMain)

            ' Header
            Dim lblHeader As New Label() With {
                .Text = "The following configuration values contain placeholders that need to be filled in." & vbCrLf &
                        "Please provide the values for your environment:",
                .Dock = DockStyle.Fill,
                .AutoSize = True,
                .Padding = New Padding(0, 0, 0, 10)
            }
            tblMain.Controls.Add(lblHeader, 0, 0)

            ' DataGridView for placeholder input
            Dim dgv As New DataGridView() With {
                .Dock = DockStyle.Fill,
                .AllowUserToAddRows = False,
                .AllowUserToDeleteRows = False,
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                .SelectionMode = DataGridViewSelectionMode.CellSelect,
                .RowHeadersVisible = False,
                .Margin = New Padding(0, 5, 0, 10),
                .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
                .EnableHeadersVisualStyles = True
            }

            Dim colSegment As New DataGridViewTextBoxColumn() With {
                .HeaderText = "Segment",
                .Name = "colSegment",
                .ReadOnly = True,
                .Width = 100
            }
            Dim colPlaceholder As New DataGridViewTextBoxColumn() With {
                .HeaderText = "Placeholder",
                .Name = "colPlaceholder",
                .ReadOnly = True,
                .Width = 150
            }
            Dim colValue As New DataGridViewTextBoxColumn() With {
                .HeaderText = "Your Value",
                .Name = "colValue",
                .ReadOnly = False,
                .Width = 200
            }
            Dim colContext As New DataGridViewTextBoxColumn() With {
                .HeaderText = "Used In",
                .Name = "colContext",
                .ReadOnly = True,
                .Width = 200
            }

            dgv.Columns.AddRange(colSegment, colPlaceholder, colValue, colContext)

            For Each mp In missingPlaceholders
                Dim segmentDisplay = If(String.IsNullOrEmpty(mp.SegmentName), "(global)", mp.SegmentName)
                dgv.Rows.Add(segmentDisplay, mp.Placeholder, "", TruncateValue(mp.SampleContext, 60))
            Next

            ' Ensure column headers are properly sized
            dgv.AutoResizeColumnHeadersHeight()
            tblMain.Controls.Add(dgv, 0, 1)

            ' Buttons
            Dim pnlButtons As New FlowLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.LeftToRight,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Margin = New Padding(0, 5, 0, 5)
            }
            tblMain.Controls.Add(pnlButtons, 0, 2)

            Dim dialogResult As Boolean = False

            Dim btnOK As New Button() With {
                .Text = "Apply Values",
                .AutoSize = True,
                .Padding = New Padding(10, 5, 10, 5),
                .Margin = New Padding(0, 0, 10, 0)
            }
            Dim btnCancel As New Button() With {
                .Text = "Skip These Updates",
                .AutoSize = True,
                .Padding = New Padding(10, 5, 10, 5)
            }

            pnlButtons.Controls.Add(btnOK)
            pnlButtons.Controls.Add(btnCancel)

            AddHandler btnCancel.Click, Sub() form.Close()

            AddHandler btnOK.Click, Sub()
                                        ' Validate all values are filled
                                        Dim allFilled = True
                                        For i As Integer = 0 To dgv.Rows.Count - 1
                                            Dim cellValue = dgv.Rows(i).Cells("colValue").Value
                                            If cellValue Is Nothing OrElse String.IsNullOrWhiteSpace(cellValue.ToString()) Then
                                                allFilled = False
                                                dgv.Rows(i).Cells("colValue").Style.BackColor = Color.FromArgb(255, 230, 230)
                                            Else
                                                dgv.Rows(i).Cells("colValue").Style.BackColor = SystemColors.Window
                                            End If
                                        Next

                                        If Not allFilled Then
                                            ShowCustomMessageBox("Please provide values for all placeholders (highlighted in red).")
                                            Return
                                        End If

                                        ' Collect values
                                        For i As Integer = 0 To dgv.Rows.Count - 1
                                            missingPlaceholders(i).UserValue = dgv.Rows(i).Cells("colValue").Value?.ToString()
                                        Next

                                        dialogResult = True
                                        form.Close()
                                    End Sub

            ' Surface any same-process native host prompt (e.g. Word's "Save changes?")
            ' that would otherwise stay hidden behind this TopMost modal dialog.
            AttachForeignForegroundWatchdog(form)

            form.ShowDialog()

            Return dialogResult
        End Function

        ''' <summary>
        ''' Generates the pending updates comment block for silent mode when placeholders are missing.
        ''' </summary>
        ''' <param name="changes">All changes for the segment/file.</param>
        ''' <param name="missingPlaceholders">Placeholders that need values.</param>
        ''' <param name="segmentName">Segment name for context.</param>
        ''' <returns>List of comment lines to insert.</returns>
        Private Shared Function GeneratePendingUpdatesBlock(
            changes As List(Of IniParameterChange),
            missingPlaceholders As List(Of String),
            segmentName As String) As List(Of String)

            Dim lines As New List(Of String)()
            Dim timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

            lines.Add("")
            lines.Add($"; === PENDING UPDATES ({timestamp}) ===")
            lines.Add("; The following updates could not be applied automatically because")
            lines.Add("; placeholder values are missing. To apply these updates:")
            lines.Add(";   1. Fill in the placeholder values below")
            lines.Add(";   2. Uncomment the parameter lines you want to apply")
            lines.Add(";   3. Comment out or DELETE the existing parameter lines below that are being replaced")
            lines.Add(";   4. Remove this comment block when done")
            lines.Add(";")
            lines.Add("; Required placeholder values:")

            For Each placeholder In missingPlaceholders.Distinct(StringComparer.OrdinalIgnoreCase)
                lines.Add($"; {placeholder} = ")
            Next

            lines.Add(";")
            lines.Add("; Pending parameter changes:")

            For Each change In changes
                Dim changeInfo = If(change.OldValue = "(new key)", "(NEW)", $"(replaces existing value)")
                lines.Add($"; {change.ParameterKey} = {change.NewValue}  {changeInfo}")
            Next

            lines.Add("; === END PENDING UPDATES ===")
            lines.Add("")  ' Empty line after block

            Return lines
        End Function

        ''' <summary>
        ''' Writes placeholder definition comments to the INI file.
        ''' </summary>
        ''' <param name="filePath">Path to the INI file.</param>
        ''' <param name="segmentName">Segment name (empty for global).</param>
        ''' <param name="placeholderValues">Dictionary of placeholder → value to write.</param>
        Private Shared Sub WriteNewPlaceholderDefinitions(
            filePath As String,
            segmentName As String,
            placeholderValues As Dictionary(Of String, String))

            If Not File.Exists(filePath) OrElse placeholderValues Is Nothing OrElse placeholderValues.Count = 0 Then
                Return
            End If

            Try
                Dim lines = File.ReadAllLines(filePath, Encoding.UTF8).ToList()
                Dim insertIndex As Integer = -1

                If String.IsNullOrEmpty(segmentName) Then
                    ' Global scope - insert at the beginning (after any existing comments/blank lines at top)
                    insertIndex = 0
                    For i As Integer = 0 To lines.Count - 1
                        Dim line = lines(i).Trim()
                        If line.StartsWith("[") Then
                            insertIndex = i
                            Exit For
                        ElseIf Not String.IsNullOrEmpty(line) AndAlso Not line.StartsWith(";") Then
                            insertIndex = i
                            Exit For
                        End If
                    Next
                    If insertIndex = -1 Then insertIndex = lines.Count
                Else
                    ' Find the segment header and insert right after it
                    For i As Integer = 0 To lines.Count - 1
                        Dim line = lines(i).Trim()
                        If line.StartsWith("[") AndAlso line.EndsWith("]") Then
                            Dim name = line.Substring(1, line.Length - 2).Trim()
                            If name.Equals(segmentName, StringComparison.OrdinalIgnoreCase) Then
                                insertIndex = i + 1
                                Exit For
                            End If
                        End If
                    Next
                End If

                If insertIndex = -1 Then Return

                ' Build placeholder comment lines
                Dim newLines As New List(Of String)()
                For Each kvp In placeholderValues
                    newLines.Add($"; {kvp.Key} = {kvp.Value}")
                Next

                ' Insert the lines
                lines.InsertRange(insertIndex, newLines)

                File.WriteAllLines(filePath, lines, Encoding.UTF8)

            Catch ex As Exception
                Debug.WriteLine($"Error writing placeholder definitions: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' Writes the pending updates comment block to the INI file.
        ''' </summary>
        ''' <param name="filePath">Path to the INI file.</param>
        ''' <param name="segmentName">Segment name (empty for global).</param>
        ''' <param name="commentLines">Lines to insert.</param>
        Private Shared Sub WritePendingUpdatesBlock(
            filePath As String,
            segmentName As String,
            commentLines As List(Of String))

            If Not File.Exists(filePath) OrElse commentLines Is Nothing OrElse commentLines.Count = 0 Then
                Return
            End If

            Try
                Dim lines = File.ReadAllLines(filePath, Encoding.UTF8).ToList()
                Dim insertIndex As Integer = -1

                If String.IsNullOrEmpty(segmentName) Then
                    ' Global scope - insert at the end of the file
                    insertIndex = lines.Count
                Else
                    ' Find the segment header and insert right after it
                    For i As Integer = 0 To lines.Count - 1
                        Dim line = lines(i).Trim()
                        If line.StartsWith("[") AndAlso line.EndsWith("]") Then
                            Dim name = line.Substring(1, line.Length - 2).Trim()
                            If name.Equals(segmentName, StringComparison.OrdinalIgnoreCase) Then
                                insertIndex = i + 1
                                Exit For
                            End If
                        End If
                    Next
                End If

                If insertIndex = -1 Then Return

                ' Insert the comment block
                lines.InsertRange(insertIndex, commentLines)

                File.WriteAllLines(filePath, lines, Encoding.UTF8)

                LogIniUpdateEvent("Pending Updates Written",
                    $"Wrote pending updates block to {Path.GetFileName(filePath)}" &
                    If(String.IsNullOrEmpty(segmentName), "", $" [{segmentName}]"),
                    alwaysLog:=True)

            Catch ex As Exception
                Debug.WriteLine($"Error writing pending updates block: {ex.Message}")
            End Try
        End Sub

#End Region

#Region "Update Source Loading"

        ''' <summary>
        ''' Loads update content from a local path, network share, or HTTP(S) URL and optionally verifies its signature.
        ''' </summary>
        ''' <param name="sourcePath">Path or URL to the update source.</param>
        ''' <param name="publicKey">Base64-encoded Ed25519 public key used for signature verification.</param>
        ''' <param name="signatureErrors">Collector for formatted signature and download errors.</param>
        ''' <param name="failedSources">Collector for sources that failed signature validation or required verification inputs.</param>
        ''' <returns>The loaded content text, or <c>Nothing</c> if loading or required signature verification fails.</returns>
        Private Shared Function LoadUpdateSourceContent(sourcePath As String, publicKey As String, signatureErrors As List(Of String), Optional failedSources As HashSet(Of String) = Nothing) As String

            If String.IsNullOrWhiteSpace(sourcePath) Then Return Nothing

            Try
                Dim isRemote = sourcePath.StartsWith("https://", StringComparison.OrdinalIgnoreCase) OrElse
                       sourcePath.StartsWith("http://", StringComparison.OrdinalIgnoreCase)

                ' Check if remote sources are allowed
                If isRemote AndAlso Not _iniUpdateContext.INI_UpdateIniAllowRemote Then
                    Debug.WriteLine($"Remote update source blocked by policy: {sourcePath}")
                    Return Nothing
                End If

                Dim contentBytes As Byte() = Nothing
                Dim signatureContent As String = Nothing
                Dim expandedPath As String = Nothing

                If isRemote Then
                    ' Enable TLS 1.2 for HTTPS
                    ServicePointManager.SecurityProtocol = ServicePointManager.SecurityProtocol Or SecurityProtocolType.Tls12

                    Using client As New HttpClient()
                        client.Timeout = TimeSpan.FromSeconds(30)

                        ' Download main content as exact bytes
                        Try
                            Dim contentTask = client.GetByteArrayAsync(sourcePath)
                            contentTask.Wait()
                            contentBytes = contentTask.Result
                        Catch ex As Exception
                            signatureErrors?.Add($"SOURCE: {sourcePath}" & vbCrLf &
                                         $"ERROR: Failed to download update source" & vbCrLf &
                                         $"DETAILS: {ex.Message}")
                            failedSources?.Add(sourcePath)
                            Return Nothing
                        End Try

                        ' Download signature file (.sig) - only if signature verification is required AND public key exists
                        If Not _iniUpdateContext.INI_UpdateIniNoSignature AndAlso Not String.IsNullOrWhiteSpace(publicKey) Then
                            Try
                                Dim sigTask = client.GetStringAsync(sourcePath & ".sig")
                                sigTask.Wait()
                                signatureContent = sigTask.Result?.Trim()
                            Catch ex As Exception
                                signatureErrors?.Add($"SOURCE: {sourcePath}" & vbCrLf &
                                             $"ERROR: Signature file not found or inaccessible" & vbCrLf &
                                             $"EXPECTED: {sourcePath}.sig" & vbCrLf &
                                             $"DETAILS: {ex.Message}" & vbCrLf &
                                             $"ACTION: Ensure the .sig file exists alongside the update file, or contact your administrator.")
                                failedSources?.Add(sourcePath)
                            End Try
                        End If
                    End Using
                Else
                    ' Local or network path - expand environment variables
                    expandedPath = ExpandEnvironmentVariables(sourcePath)
                    If Not File.Exists(expandedPath) Then
                        Debug.WriteLine($"Update source file not found: {expandedPath}")
                        Return Nothing
                    End If

                    Try
                        contentBytes = File.ReadAllBytes(expandedPath)
                    Catch ex As Exception
                        signatureErrors?.Add($"SOURCE: {expandedPath}" & vbCrLf &
                                     $"ERROR: Failed to read update source file" & vbCrLf &
                                     $"DETAILS: {ex.Message}")
                        failedSources?.Add(sourcePath)
                        Return Nothing
                    End Try

                    ' Check for signature file - only if signature verification is required AND public key exists
                    If Not _iniUpdateContext.INI_UpdateIniNoSignature AndAlso Not String.IsNullOrWhiteSpace(publicKey) Then
                        Dim sigPath = expandedPath & ".sig"
                        If File.Exists(sigPath) Then
                            Try
                                signatureContent = File.ReadAllText(sigPath, Encoding.UTF8)?.Trim()
                            Catch ex As Exception
                                signatureErrors?.Add($"SOURCE: {expandedPath}" & vbCrLf &
                                             $"ERROR: Failed to read signature file" & vbCrLf &
                                             $"SIGNATURE FILE: {sigPath}" & vbCrLf &
                                             $"DETAILS: {ex.Message}")
                                failedSources?.Add(sourcePath)
                            End Try
                        Else
                            signatureErrors?.Add($"SOURCE: {expandedPath}" & vbCrLf &
                                         $"ERROR: Signature file not found" & vbCrLf &
                                         $"EXPECTED: {sigPath}" & vbCrLf &
                                         $"ACTION: Create a .sig file using the Signature Management tool, or contact your administrator.")
                            failedSources?.Add(sourcePath)
                        End If
                    End If
                End If

                ' Verify signature if required
                If Not _iniUpdateContext.INI_UpdateIniNoSignature Then
                    Dim displayPath = If(expandedPath, sourcePath)

                    ' If no public key is configured, skip signature verification but add warning
                    If String.IsNullOrWhiteSpace(publicKey) Then
                        signatureErrors?.Add($"SOURCE: {displayPath}" & vbCrLf &
                                     $"WARNING: No public key configured - signature verification skipped" & vbCrLf &
                                     $"NOTE: Updates will proceed without cryptographic verification." & vbCrLf &
                                     $"ACTION: For better security, add a public key as the third parameter in UpdateSource:" & vbCrLf &
                                     $"        UpdateSource = path; keys; PUBLIC_KEY_HERE")
                        failedSources?.Add(sourcePath)
                        Return DecodeContentText(contentBytes)
                    End If

                    If String.IsNullOrWhiteSpace(signatureContent) Then
                        Return Nothing
                    End If

                    If Not VerifyEd25519SignatureCompatible(contentBytes, signatureContent, publicKey) Then
                        signatureErrors?.Add($"SOURCE: {displayPath}" & vbCrLf &
                                     $"ERROR: SIGNATURE VERIFICATION FAILED" & vbCrLf &
                                     $"⚠ This may indicate the file has been tampered with!" & vbCrLf &
                                     $"POSSIBLE CAUSES:" & vbCrLf &
                                     $"  - File was modified after signing" & vbCrLf &
                                     $"  - Wrong public key configured" & vbCrLf &
                                     $"  - Signature file corrupted or for different file" & vbCrLf &
                                     $"ACTION: Contact your administrator immediately.")
                        failedSources?.Add(sourcePath)
                        Return Nothing
                    End If
                End If

                Return DecodeContentText(contentBytes)

            Catch ex As Exception
                signatureErrors?.Add($"SOURCE: {sourcePath}" & vbCrLf &
                             $"ERROR: Unexpected error during update check" & vbCrLf &
                             $"DETAILS: {ex.Message}")
                failedSources?.Add(sourcePath)
                Return Nothing
            End Try
        End Function

        Private Shared Function DecodeContentText(contentBytes As Byte()) As String
            If contentBytes Is Nothing Then Return Nothing

            Using ms As New MemoryStream(contentBytes)
                Using reader As New StreamReader(ms, Encoding.UTF8, detectEncodingFromByteOrderMarks:=True)
                    Return reader.ReadToEnd()
                End Using
            End Using
        End Function

#End Region

#Region "Change Detection"

        ''' <summary>
        ''' Compares a local non-segmented INI file with its update source and returns the detected parameter changes.
        ''' </summary>
        ''' <param name="localPath">Local INI file path.</param>
        ''' <param name="fileName">INI file name used for reporting.</param>
        ''' <param name="updateSource">Update source definition string (<c>path; keys; publicKey</c>).</param>
        ''' <param name="segmentName">Segment name used for reporting; empty for global scope.</param>
        ''' <param name="signatureErrors">Collector for signature-related diagnostics.</param>
        ''' <param name="failedSources">Collector for sources that failed signature validation or required verification inputs.</param>
        ''' <returns>List of detected changes; empty if none are found or update content cannot be loaded.</returns>
        Private Shared Function CheckSingleIniFile(localPath As String, fileName As String,
                                                   updateSource As String, segmentName As String,
                                                   signatureErrors As List(Of String),
                                                   Optional failedSources As HashSet(Of String) = Nothing) As List(Of IniParameterChange)

            Dim changes As New List(Of IniParameterChange)()

            Try
                ' Parse update source
                Dim sourceInfo = ParseGlobalUpdateSource(updateSource)
                If String.IsNullOrWhiteSpace(sourceInfo.UpdatePath) Then Return changes
                If sourceInfo.UpdateKeys Is Nothing OrElse sourceInfo.UpdateKeys.Count = 0 Then Return changes

                ' Circular reference check
                If IsSameFile(localPath, sourceInfo.UpdatePath) Then
                    Debug.WriteLine($"INI Update: Skipping self-referencing update source for {fileName}")
                    LogIniUpdateEvent("Skipped", $"Self-referencing UpdateSource detected for {fileName} - update source points to itself")
                    Return changes
                End If

                ' Load remote content (with signature validation)
                Dim remoteContent = LoadUpdateSourceContent(sourceInfo.UpdatePath, sourceInfo.PublicKey, signatureErrors, failedSources)
                If String.IsNullOrWhiteSpace(remoteContent) Then Return changes

                ' Parse local and remote files
                Dim localIni = ParseIniFile(localPath)
                Dim remoteParams = ParseIniContentToDict(remoteContent)

                ' Check for special key modifiers
                Dim hasAll = sourceInfo.UpdateKeys.Any(Function(k) k.Equals("all", StringComparison.OrdinalIgnoreCase))
                Dim hasNew = sourceInfo.UpdateKeys.Any(Function(k) k.Equals("new", StringComparison.OrdinalIgnoreCase))
                Dim excludedKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                For Each k In sourceInfo.UpdateKeys
                    If k.StartsWith("-", StringComparison.Ordinal) AndAlso k.Length > 1 Then
                        excludedKeys.Add(k.Substring(1).Trim())
                    End If
                Next

                Dim specificKeys = sourceInfo.UpdateKeys.Where(
                            Function(k) Not k.Equals("all", StringComparison.OrdinalIgnoreCase) AndAlso
                                        Not k.Equals("new", StringComparison.OrdinalIgnoreCase) AndAlso
                                        Not k.StartsWith("-", StringComparison.Ordinal)
                        ).ToList()

                ' Determine which keys to check
                Dim keysToCheck As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                If hasAll Then
                    ' "all" - check all keys from remote
                    For Each key In remoteParams.Keys
                        keysToCheck.Add(key)
                    Next
                Else
                    ' Add specific keys
                    For Each key In specificKeys
                        keysToCheck.Add(key)
                    Next
                End If

                ' If "new" is specified, also include keys that exist in remote but not in local
                If hasNew Then
                    For Each key In remoteParams.Keys
                        If Not localIni.GlobalParameters.ContainsKey(key) Then
                            keysToCheck.Add(key)
                        End If
                    Next
                End If

                ' Remove explicitly excluded keys
                For Each excludedKey In excludedKeys
                    keysToCheck.Remove(excludedKey)
                Next

                ' Parse placeholder definitions from local file
                Dim localLines = File.ReadAllLines(localPath, Encoding.UTF8)
                Dim placeholderDefs = ParsePlaceholderDefinitions(localLines, If(segmentName, ""))

                ' Compare values
                For Each key In keysToCheck
                    ' Skip protected keys
                    If IsProtectedKey(key) Then Continue For

                    If Not remoteParams.ContainsKey(key) Then Continue For

                    Dim remoteValue = remoteParams(key)
                    Dim localValue As String = Nothing
                    localIni.GlobalParameters.TryGetValue(key, localValue)

                    ' Apply placeholder substitutions to remote value
                    Dim effectiveRemoteValue = remoteValue
                    Dim placeholdersInValue = ExtractPlaceholders(remoteValue)

                    If placeholdersInValue.Count > 0 Then
                        effectiveRemoteValue = ApplyPlaceholders(remoteValue, placeholderDefs)
                    End If

                    ' Check if values differ (or key is new)
                    If Not String.Equals(localValue, effectiveRemoteValue, StringComparison.Ordinal) Then
                        Dim isNewKey = localValue Is Nothing
                        Dim change As New IniParameterChange() With {
                            .IniFile = fileName,
                            .SegmentName = If(segmentName, ""),
                            .ParameterKey = key,
                            .OldValue = If(localValue, "(new key)"),
                            .NewValue = effectiveRemoteValue,
                            .IsSelected = True,
                            .IsSuspicious = IsPathOrUrlChange(localValue, effectiveRemoteValue),
                            .RemoteTemplateValue = If(placeholdersInValue.Count > 0, remoteValue, Nothing),
                            .Placeholders = If(placeholdersInValue.Count > 0, placeholdersInValue, Nothing)
                        }

                        ' Suspicious changes and new keys with URLs/paths are not selected by default
                        If change.IsSuspicious Then change.IsSelected = False
                        If isNewKey AndAlso IsPathOrUrlChange("", effectiveRemoteValue) Then change.IsSelected = False

                        changes.Add(change)
                    End If
                Next

            Catch ex As Exception
                Debug.WriteLine($"Error checking INI file {localPath}: {ex.Message}")
            End Try

            Return changes
        End Function

        ''' <summary>
        ''' Compares each local segment in a segmented INI file with its segment-specific update source.
        ''' </summary>
        ''' <param name="localPath">Local INI file path.</param>
        ''' <param name="fileName">INI file name used for reporting.</param>
        ''' <param name="signatureErrors">Collector for signature-related diagnostics.</param>
        ''' <param name="updateSources">Optional map updated with resolved update source paths per file/segment.</param>
        ''' <param name="failedSources">Collector for sources that failed signature validation or required verification inputs.</param>
        ''' <returns>List of detected changes; empty if none are found.</returns>
        Private Shared Function CheckSegmentedIniFile(localPath As String, fileName As String,
                                              signatureErrors As List(Of String),
                                              Optional updateSources As Dictionary(Of String, String) = Nothing,
                                              Optional failedSources As HashSet(Of String) = Nothing) As List(Of IniParameterChange)

            Dim changes As New List(Of IniParameterChange)()

            Try
                Dim localIni = ParseIniFile(localPath)

                For Each segment In localIni.Segments
                    ' Skip segments without UpdateSource
                    If String.IsNullOrWhiteSpace(segment.UpdatePath) Then Continue For
                    If segment.UpdateKeys Is Nothing OrElse segment.UpdateKeys.Count = 0 Then Continue For

                    ' Circular reference Check
                    If IsSameFile(localPath, segment.UpdatePath) Then
                        Debug.WriteLine($"INI Update: Skipping self-referencing update source for {fileName}[{segment.Name}]")
                        LogIniUpdateEvent("Skipped", $"Self-referencing UpdateSource in segment [{segment.Name}] of {fileName}")
                        Continue For
                    End If

                    ' Track update source for LocalTrusted mode
                    ' Use segment-specific key: "filename|segmentname" to handle multiple segments
                    If updateSources IsNot Nothing Then
                        Dim sourceKey As String = $"{fileName}|{segment.Name}"
                        updateSources(sourceKey) = segment.UpdatePath
                    End If

                    ' Load remote content for this segment (with signature validation)
                    Dim remoteContent As String = LoadUpdateSourceContent(segment.UpdatePath, segment.PublicKey, signatureErrors, failedSources)
                    If String.IsNullOrWhiteSpace(remoteContent) Then Continue For

                    ' Parse remote content - look for matching segment
                    Dim remoteSegment As IniSegment = FindSegmentInContent(remoteContent, segment.Name)
                    If remoteSegment Is Nothing Then Continue For

                    ' Check for special key modifiers
                    Dim hasAll As Boolean = segment.UpdateKeys.Any(Function(k) k.Equals("all", StringComparison.OrdinalIgnoreCase))
                    Dim hasNew As Boolean = segment.UpdateKeys.Any(Function(k) k.Equals("new", StringComparison.OrdinalIgnoreCase))

                    Dim excludedKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                    For Each k In segment.UpdateKeys
                        If k.StartsWith("-", StringComparison.Ordinal) AndAlso k.Length > 1 Then
                            excludedKeys.Add(k.Substring(1).Trim())
                        End If
                    Next

                    Dim specificKeys As List(Of String) = segment.UpdateKeys.Where(
                                Function(k) Not k.Equals("all", StringComparison.OrdinalIgnoreCase) AndAlso
                                            Not k.Equals("new", StringComparison.OrdinalIgnoreCase) AndAlso
                                            Not k.StartsWith("-", StringComparison.Ordinal)
                            ).ToList()

                    ' Determine which keys to check
                    Dim keysToCheck As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                    If hasAll Then
                        For Each key As String In remoteSegment.Parameters.Keys
                            keysToCheck.Add(key)
                        Next
                    Else
                        For Each key As String In specificKeys
                            keysToCheck.Add(key)
                        Next
                    End If

                    If hasNew Then
                        For Each key As String In remoteSegment.Parameters.Keys
                            If Not segment.Parameters.ContainsKey(key) Then
                                keysToCheck.Add(key)
                            End If
                        Next
                    End If

                    ' Remove explicitly excluded keys
                    For Each excludedKey In excludedKeys
                        keysToCheck.Remove(excludedKey)
                    Next

                    ' Parse placeholder definitions for this segment
                    Dim localLines = File.ReadAllLines(localPath, Encoding.UTF8)
                    Dim placeholderDefs = ParsePlaceholderDefinitions(localLines, segment.Name)

                    ' Compare values
                    For Each key As String In keysToCheck
                        If IsProtectedKey(key) Then Continue For
                        If Not remoteSegment.Parameters.ContainsKey(key) Then Continue For

                        Dim remoteValue As String = remoteSegment.Parameters(key)
                        Dim localValue As String = Nothing
                        segment.Parameters.TryGetValue(key, localValue)

                        ' Apply placeholder substitutions to remote value
                        Dim effectiveRemoteValue = remoteValue
                        Dim placeholdersInValue = ExtractPlaceholders(remoteValue)

                        If placeholdersInValue.Count > 0 Then
                            effectiveRemoteValue = ApplyPlaceholders(remoteValue, placeholderDefs)
                        End If

                        If Not String.Equals(localValue, effectiveRemoteValue, StringComparison.Ordinal) Then
                            Dim isNewKey As Boolean = localValue Is Nothing
                            Dim change As New IniParameterChange() With {
                                .IniFile = fileName,
                                .SegmentName = segment.Name,
                                .ParameterKey = key,
                                .OldValue = If(localValue, "(new key)"),
                                .NewValue = effectiveRemoteValue,
                                .IsSelected = True,
                                .IsSuspicious = IsPathOrUrlChange(localValue, effectiveRemoteValue),
                                .RemoteTemplateValue = If(placeholdersInValue.Count > 0, remoteValue, Nothing),
                                .Placeholders = If(placeholdersInValue.Count > 0, placeholdersInValue, Nothing)
                            }

                            If change.IsSuspicious Then change.IsSelected = False
                            If isNewKey AndAlso IsPathOrUrlChange("", effectiveRemoteValue) Then change.IsSelected = False

                            changes.Add(change)
                        End If
                    Next

                Next

            Catch ex As Exception
                Debug.WriteLine($"Error checking segmented INI file {localPath}: {ex.Message}")
            End Try

            Return changes
        End Function

        ''' <summary>
        ''' Determines whether two paths refer to the same file (handles environment variables, relative paths, UNC variations).
        ''' </summary>
        ''' <param name="path1">First path to compare.</param>
        ''' <param name="path2">Second path to compare.</param>
        ''' <returns><c>True</c> if both paths resolve to the same file; otherwise <c>False</c>.</returns>
        Private Shared Function IsSameFile(path1 As String, path2 As String) As Boolean
            If String.IsNullOrWhiteSpace(path1) OrElse String.IsNullOrWhiteSpace(path2) Then
                Return False
            End If

            ' Skip comparison for remote URLs - they can't be the same as local files
            If path1.StartsWith("http://", StringComparison.OrdinalIgnoreCase) OrElse
       path1.StartsWith("https://", StringComparison.OrdinalIgnoreCase) OrElse
       path2.StartsWith("http://", StringComparison.OrdinalIgnoreCase) OrElse
       path2.StartsWith("https://", StringComparison.OrdinalIgnoreCase) Then
                Return False
            End If

            Try
                ' Expand environment variables and resolve to full paths
                Dim expanded1 = ExpandEnvironmentVariables(path1)
                Dim expanded2 = ExpandEnvironmentVariables(path2)

                ' Get canonical full paths (resolves relative paths, normalizes separators)
                Dim fullPath1 = Path.GetFullPath(expanded1)
                Dim fullPath2 = Path.GetFullPath(expanded2)

                ' Compare case-insensitively (Windows file system)
                Return String.Equals(fullPath1, fullPath2, StringComparison.OrdinalIgnoreCase)
            Catch
                ' If path resolution fails, fall back to simple string comparison
                Return String.Equals(path1, path2, StringComparison.OrdinalIgnoreCase)
            End Try
        End Function

        ''' <summary>
        ''' Parses INI text content into a flat key/value dictionary, ignoring segments and comments.
        ''' </summary>
        ''' <param name="content">INI content text.</param>
        ''' <returns>Dictionary of parsed keys and values.</returns>
        Private Shared Function ParseIniContentToDict(content As String) As Dictionary(Of String, String)
            Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            If String.IsNullOrWhiteSpace(content) Then Return result

            Dim lines = content.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
            For Each rawLine In lines
                Dim line = rawLine.Trim()
                If String.IsNullOrEmpty(line) OrElse line.StartsWith(";") OrElse line.StartsWith("[") Then Continue For

                Dim eqIndex = line.IndexOf("="c)
                If eqIndex > 0 Then
                    Dim key = line.Substring(0, eqIndex).Trim()
                    Dim value = line.Substring(eqIndex + 1).Trim()
                    result(key) = value
                End If
            Next

            Return result
        End Function

        ''' <summary>
        ''' Extracts a named segment and its key/value pairs from INI content text.
        ''' </summary>
        ''' <param name="content">INI content text.</param>
        ''' <param name="segmentName">The segment name to locate.</param>
        ''' <returns>The parsed segment, or <c>Nothing</c> if not found.</returns>
        Private Shared Function FindSegmentInContent(content As String, segmentName As String) As IniSegment
            If String.IsNullOrWhiteSpace(content) OrElse String.IsNullOrWhiteSpace(segmentName) Then Return Nothing

            Dim lines = content.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
            Dim inTargetSegment = False
            Dim segment As IniSegment = Nothing

            For Each rawLine In lines
                Dim line = rawLine.Trim()

                If line.StartsWith("[") AndAlso line.EndsWith("]") Then
                    ' End previous segment if we found our target
                    If inTargetSegment AndAlso segment IsNot Nothing Then
                        Return segment
                    End If

                    Dim name = line.Substring(1, line.Length - 2).Trim()
                    If name.Equals(segmentName, StringComparison.OrdinalIgnoreCase) Then
                        inTargetSegment = True
                        segment = New IniSegment() With {.Name = name}
                    Else
                        inTargetSegment = False
                    End If
                    Continue For
                End If

                If inTargetSegment AndAlso segment IsNot Nothing Then
                    If String.IsNullOrEmpty(line) OrElse line.StartsWith(";") Then Continue For

                    Dim eqIndex = line.IndexOf("="c)
                    If eqIndex > 0 Then
                        Dim key = line.Substring(0, eqIndex).Trim()
                        Dim value = line.Substring(eqIndex + 1).Trim()
                        segment.Parameters(key) = value
                    End If
                End If
            Next

            Return segment
        End Function

        ''' <summary>
        ''' Determines whether a value change is treated as suspicious based on URL/path-like patterns.
        ''' </summary>
        ''' <param name="oldValue">Existing value in the local INI file, or empty for new keys.</param>
        ''' <param name="newValue">Proposed value from the update source.</param>
        ''' <returns><c>True</c> if the change is classified as suspicious; otherwise <c>False</c>.</returns>
        Private Shared Function IsPathOrUrlChange(oldValue As String, newValue As String) As Boolean
            ' Check if either contains URL patterns
            Dim urlPatterns = {"http://", "https://", "ftp://", "file://"}
            Dim pathPatterns = {":\", ":\\", "/"}

            Dim oldHasUrl = Not String.IsNullOrWhiteSpace(oldValue) AndAlso
                    urlPatterns.Any(Function(p) oldValue.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0)
            Dim newHasUrl = Not String.IsNullOrWhiteSpace(newValue) AndAlso
                    urlPatterns.Any(Function(p) newValue.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0)
            Dim oldHasPath = Not String.IsNullOrWhiteSpace(oldValue) AndAlso
                     pathPatterns.Any(Function(p) oldValue.Contains(p))
            Dim newHasPath = Not String.IsNullOrWhiteSpace(newValue) AndAlso
                     pathPatterns.Any(Function(p) newValue.Contains(p))

            ' Suspicious if:
            ' 1. New key with URL/path in the value
            ' 2. Existing key where URL/path is present and values differ
            If String.IsNullOrWhiteSpace(oldValue) Then
                ' New key - suspicious if new value contains URL or path
                Return newHasUrl OrElse newHasPath
            End If

            ' Existing key - suspicious if URL or path is present and values differ
            If oldHasUrl OrElse newHasUrl OrElse oldHasPath OrElse newHasPath Then
                Return Not String.Equals(oldValue, newValue, StringComparison.OrdinalIgnoreCase)
            End If

            Return False
        End Function

#End Region

#Region "Ignored Parameters Management"

        ''' <summary>
        ''' Represents a parsed override rule for filtering ignored parameters.
        ''' </summary>
        Private Class IgnoreOverrideRule
            Public Property IsForceIgnore As Boolean    ' True = +, False = -
            Public Property FileName As String          ' "*" = any file, or specific filename
            Public Property SegmentName As String       ' "*" = any segment, "" = global, or specific segment
            Public Property ParameterKey As String      ' "*" = any key, "all" = special, or specific key

            ''' <summary>
            ''' Returns whether this override rule matches the specified change (file/segment/key).
            ''' </summary>
            ''' <param name="change">The change candidate.</param>
            ''' <returns><c>True</c> if the rule matches; otherwise <c>False</c>.</returns>            
            Public Function Matches(change As IniParameterChange) As Boolean
                ' Check filename match
                If FileName <> "*" AndAlso Not FileName.Equals(change.IniFile, StringComparison.OrdinalIgnoreCase) Then
                    Return False
                End If

                ' Check segment match
                If SegmentName <> "*" Then
                    Dim changeSegment = If(change.SegmentName, "")
                    If Not SegmentName.Equals(changeSegment, StringComparison.OrdinalIgnoreCase) Then
                        Return False
                    End If
                End If

                ' Check parameter key match
                If ParameterKey <> "*" AndAlso Not ParameterKey.Equals(change.ParameterKey, StringComparison.OrdinalIgnoreCase) Then
                    Return False
                End If

                Return True
            End Function
        End Class


        ''' <summary>
        ''' Parses ignore override rules from <c>INI_UpdateIniIgnoreOverride</c>.
        ''' </summary>
        ''' <param name="overrideValue">Raw rule string.</param>
        ''' <returns>List of parsed rules; empty if the input is empty or invalid.</returns>
        ''' <emarks>
        ''' Supported formats:
        '''   Simple:     "+Key" or "-Key" (matches any file/segment)
        '''   File:       "+filename.ini|Key" or "-filename.ini|Key"
        '''   Full:       "+filename.ini|Segment|Key" or "-filename.ini|Segment|Key"
        '''   Wildcards:  "+*|Segment|Key" or "+filename.ini|*|Key"
        '''   Special:    "+all" (ignore everything), "-all" (include everything)
        ''' </remarks>
        Private Shared Function ParseIgnoreOverrideRules(overrideValue As String) As List(Of IgnoreOverrideRule)
            Dim rules As New List(Of IgnoreOverrideRule)()

            If String.IsNullOrWhiteSpace(overrideValue) Then Return rules

            For Each item In overrideValue.Split(","c)
                Dim trimmed = item.Trim()
                If String.IsNullOrEmpty(trimmed) Then Continue For

                Dim isForceIgnore As Boolean
                Dim ruleBody As String

                If trimmed.StartsWith("+") Then
                    isForceIgnore = True
                    ruleBody = trimmed.Substring(1).Trim()
                ElseIf trimmed.StartsWith("-") Then
                    isForceIgnore = False
                    ruleBody = trimmed.Substring(1).Trim()
                Else
                    Continue For ' Invalid format, skip
                End If

                If String.IsNullOrEmpty(ruleBody) Then Continue For

                Dim rule As New IgnoreOverrideRule() With {.IsForceIgnore = isForceIgnore}

                ' Parse the rule body: can be "key", "file|key", or "file|segment|key"
                Dim parts = ruleBody.Split("|"c)

                Select Case parts.Length
                    Case 1
                        ' Simple format: just the key (or "all")
                        rule.FileName = "*"
                        rule.SegmentName = "*"
                        rule.ParameterKey = parts(0).Trim()

                    Case 2
                        ' File|Key format
                        rule.FileName = parts(0).Trim()
                        rule.SegmentName = "*"
                        rule.ParameterKey = parts(1).Trim()

                    Case 3
                        ' File|Segment|Key format
                        rule.FileName = parts(0).Trim()
                        rule.SegmentName = parts(1).Trim()
                        rule.ParameterKey = parts(2).Trim()

                    Case Else
                        Continue For ' Invalid format
                End Select

                ' Validate: empty components become wildcards
                If String.IsNullOrEmpty(rule.FileName) Then rule.FileName = "*"
                If String.IsNullOrEmpty(rule.SegmentName) Then rule.SegmentName = "*"
                If String.IsNullOrEmpty(rule.ParameterKey) Then Continue For ' Key is required

                rules.Add(rule)
            Next

            Return rules
        End Function

        ''' <summary>
        ''' Filters out parameters that are in the ignore list, applying any overrides.
        ''' Supports file-specific and segment-specific override rules.
        ''' More specific rules take precedence over less specific ones.
        ''' </summary>
        Private Shared Function FilterIgnoredParameters(changes As List(Of IniParameterChange)) As List(Of IniParameterChange)
            Dim result As New List(Of IniParameterChange)()

            ' Parse override rules
            Dim rules = ParseIgnoreOverrideRules(_iniUpdateContext.INI_UpdateIniIgnoreOverride)

            ' Check for global "all" rules
            Dim hasIgnoreAll = rules.Any(Function(r) r.IsForceIgnore AndAlso
                                                  r.FileName = "*" AndAlso
                                                  r.SegmentName = "*" AndAlso
                                                  r.ParameterKey.Equals("all", StringComparison.OrdinalIgnoreCase))

            Dim hasIncludeAll = rules.Any(Function(r) Not r.IsForceIgnore AndAlso
                                                   r.FileName = "*" AndAlso
                                                   r.SegmentName = "*" AndAlso
                                                   r.ParameterKey.Equals("all", StringComparison.OrdinalIgnoreCase))

            ' Apply filtering with overrides
            For Each change In changes
                Dim ignoreKey = GetIgnoreKey(change)

                ' Find matching override rules (excluding "all" special rules)
                Dim matchingRules = rules.Where(Function(r) r.Matches(change) AndAlso
                                                         Not r.ParameterKey.Equals("all", StringComparison.OrdinalIgnoreCase)).ToList()

                ' Find the most specific matching rule
                ' Specificity: file match = 4, segment match = 2, key match = 1
                Dim bestRule As IgnoreOverrideRule = Nothing
                Dim bestSpecificity As Integer = -1

                For Each rule In matchingRules
                    Dim specificity = 0
                    If rule.FileName <> "*" Then specificity += 4
                    If rule.SegmentName <> "*" Then specificity += 2
                    If rule.ParameterKey <> "*" Then specificity += 1

                    ' Higher specificity wins; if equal, later rule wins
                    If specificity >= bestSpecificity Then
                        bestSpecificity = specificity
                        bestRule = rule
                    End If
                Next

                ' Apply the most specific rule if found
                If bestRule IsNot Nothing Then
                    If Not bestRule.IsForceIgnore Then
                        result.Add(change) ' Force included
                    End If
                    ' If IsForceIgnore, skip adding (force ignored)
                    Continue For
                End If

                ' Apply global "all" rules if no specific rule matched
                If hasIgnoreAll Then
                    Continue For ' Ignore all
                End If

                If hasIncludeAll Then
                    result.Add(change)
                    Continue For
                End If

                ' Fall back to local ignore list
                Dim ignoreList = GetIgnoreListForFile(change.IniFile)
                If Not ignoreList.Contains(ignoreKey) Then
                    result.Add(change)
                End If
            Next

            Return result
        End Function


        ''' <summary>
        ''' Builds the ignore-list identifier string for a change (file[/segment]/key).
        ''' </summary>
        ''' <param name="change">The change.</param>
        ''' <returns>Ignore key string used for persistence and matching.</returns>
        Private Shared Function GetIgnoreKey(change As IniParameterChange) As String
            If String.IsNullOrEmpty(change.SegmentName) Then
                Return $"{change.IniFile}|{change.ParameterKey}"
            Else
                Return $"{change.IniFile}|{change.SegmentName}|{change.ParameterKey}"
            End If
        End Function

        ''' <summary>
        ''' Loads the persisted ignore list for the specified INI file name.
        ''' </summary>
        ''' <param name="fileName">The INI file name (not full path).</param>
        ''' <returns>A case-insensitive set of ignore keys.</returns>
        Private Shared Function GetIgnoreListForFile(fileName As String) As HashSet(Of String)
            Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim settingsValue As String = ""

            Try
                ' Normalize filename to lowercase for consistent storage key
                Dim normalizedName = fileName.ToLowerInvariant()

                ' Map known config types to their settings keys
                ' For the main INI, use "RedInk" setting
                ' For other files, derive a key from the filename
                Dim settingsKey As String

                If normalizedName.EndsWith("redink.ini") OrElse normalizedName.EndsWith($"{AN2.ToLowerInvariant()}.ini") Then
                    settingsKey = "IgnoredUpdates_RedInk"
                Else
                    ' For user-defined filenames, use a hash-based approach or store in a single collection
                    ' For simplicity, we'll use a combined setting with filename prefix
                    settingsKey = "IgnoredUpdates_Custom"
                End If

                settingsValue = CStr(My.Settings.Item(settingsKey))
            Catch
                ' Settings may not exist yet
            End Try

            If Not String.IsNullOrWhiteSpace(settingsValue) Then
                For Each item In settingsValue.Split(";"c)
                    If Not String.IsNullOrWhiteSpace(item) Then
                        ' For custom files, items are stored as "filename|segment|key" or "filename|key"
                        result.Add(item.Trim())
                    End If
                Next
            End If

            Return result
        End Function

        ''' <summary>
        ''' Persists the ignore list for the specified INI file name.
        ''' </summary>
        ''' <param name="fileName">The INI file name (not full path).</param>
        ''' <param name="ignoreList">Ignore keys to store.</param>
        Private Shared Sub SaveIgnoreListForFile(fileName As String, ignoreList As HashSet(Of String))
            Try
                Dim value = String.Join(";", ignoreList)
                Dim normalizedName = fileName.ToLowerInvariant()

                Dim settingsKey As String
                If normalizedName.EndsWith("redink.ini") OrElse normalizedName.EndsWith($"{AN2.ToLowerInvariant()}.ini") Then
                    settingsKey = "IgnoredUpdates_RedInk"
                Else
                    settingsKey = "IgnoredUpdates_Custom"
                End If

                My.Settings.Item(settingsKey) = value
                My.Settings.Save()
            Catch ex As Exception
                Debug.WriteLine($"Error saving ignore list: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' Adds the specified changes to the persisted ignore list.
        ''' </summary>
        ''' <param name="changes">Changes to persist as ignored.</param>
        Private Shared Sub AddToIgnoreList(changes As List(Of IniParameterChange))
            ' Group by file
            Dim byFile = changes.GroupBy(Function(c) c.IniFile)

            For Each fileGroup In byFile
                Dim ignoreList = GetIgnoreListForFile(fileGroup.Key)

                For Each change In fileGroup
                    ignoreList.Add(GetIgnoreKey(change))
                Next

                SaveIgnoreListForFile(fileGroup.Key, ignoreList)
            Next
        End Sub

        ''' <summary>
        ''' Shows the UI for editing the persisted ignore list entries.
        ''' </summary>
        Public Shared Sub ShowIgnoredParametersDialog()
            Dim form As New Form() With {
                .Text = $"{AN} - Manage Ignored Update Parameters",
                .Size = New Size(750, 520),
                .StartPosition = FormStartPosition.CenterParent,
                .FormBorderStyle = FormBorderStyle.Sizable,
                .Font = New Font("Segoe UI", 9.0F),
                .MinimumSize = New Size(550, 400),
                .AutoScaleMode = AutoScaleMode.Dpi,
                .TopMost = True
            }



            Try
                Dim bmp As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                form.Icon = Icon.FromHandle(bmp.GetHicon())
            Catch
            End Try

            ' Use TableLayoutPanel for proper scaling and layout
            Dim tblMain As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(15),
                .ColumnCount = 1,
                .RowCount = 3,
                .AutoSize = False
            }
            tblMain.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            tblMain.RowStyles.Add(New RowStyle(SizeType.AutoSize))  ' Info label
            tblMain.RowStyles.Add(New RowStyle(SizeType.Percent, 100))  ' CheckedListBox
            tblMain.RowStyles.Add(New RowStyle(SizeType.AutoSize))  ' Buttons
            form.Controls.Add(tblMain)

            ' Row 0: Info label
            Dim lblInfo As New Label() With {
                .Text = $"The following parameters are ignored during automatic updates. Uncheck items to remove them from the ignore list:",
                .Dock = DockStyle.Fill,
                .AutoSize = True,
                .Padding = New Padding(0, 0, 0, 5),
                .Margin = New Padding(0, 0, 0, 10)
            }
            tblMain.Controls.Add(lblInfo, 0, 0)

            ' Row 1: CheckedListBox
            Dim clb As New CheckedListBox() With {
                .Dock = DockStyle.Fill,
                .CheckOnClick = True,
                .Margin = New Padding(0, 5, 0, 10)
            }
            tblMain.Controls.Add(clb, 0, 1)

            ' Load all ignored items from both settings keys
            ' Items are stored as "filename|key" or "filename|segment|key"
            Dim allIgnored As New List(Of String)()

            ' Load from IgnoredUpdates_RedInk (main INI file)
            Try
                Dim redInkValue = CStr(My.Settings.Item("IgnoredUpdates_RedInk"))
                If Not String.IsNullOrWhiteSpace(redInkValue) Then
                    For Each item In redInkValue.Split(";"c)
                        If Not String.IsNullOrWhiteSpace(item) Then
                            allIgnored.Add(item.Trim())
                            clb.Items.Add(item.Trim(), True)
                        End If
                    Next
                End If
            Catch
            End Try

            ' Load from IgnoredUpdates_Custom (user-defined INI files)
            Try
                Dim customValue = CStr(My.Settings.Item("IgnoredUpdates_Custom"))
                If Not String.IsNullOrWhiteSpace(customValue) Then
                    For Each item In customValue.Split(";"c)
                        If Not String.IsNullOrWhiteSpace(item) Then
                            allIgnored.Add(item.Trim())
                            clb.Items.Add(item.Trim(), True)
                        End If
                    Next
                End If
            Catch
            End Try

            ' Row 2: Buttons panel
            Dim pnlButtons As New FlowLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.LeftToRight,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .WrapContents = False,
                .Margin = New Padding(0, 5, 0, 5)
            }
            tblMain.Controls.Add(pnlButtons, 0, 2)

            Dim btnSave As New Button() With {
                .Text = "Save Changes",
                .AutoSize = True,
                .Padding = New Padding(10, 5, 10, 5),
                .Margin = New Padding(0, 0, 10, 0)
            }
            Dim btnCancel As New Button() With {
                .Text = "Cancel",
                .AutoSize = True,
                .Padding = New Padding(10, 5, 10, 5)
            }

            pnlButtons.Controls.Add(btnSave)
            pnlButtons.Controls.Add(btnCancel)

            AddHandler btnCancel.Click, Sub() form.Close()

            AddHandler btnSave.Click, Sub()
                                          ' Rebuild ignore lists based on checked items
                                          ' Group by settings key based on filename in the ignore key
                                          Dim redInkItems As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                                          Dim customItems As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                                          For i As Integer = 0 To clb.Items.Count - 1
                                              If clb.GetItemChecked(i) Then
                                                  Dim item = allIgnored(i)
                                                  ' Parse the filename from the ignore key (format: "filename|key" or "filename|segment|key")
                                                  Dim parts = item.Split("|"c)
                                                  If parts.Length >= 2 Then
                                                      Dim fileName = parts(0).ToLowerInvariant()
                                                      ' Determine which settings key this belongs to
                                                      If fileName.EndsWith("redink.ini") OrElse fileName.EndsWith($"{AN2.ToLowerInvariant()}.ini") Then
                                                          redInkItems.Add(item)
                                                      Else
                                                          customItems.Add(item)
                                                      End If
                                                  End If
                                              End If
                                          Next

                                          ' Save to settings
                                          Try
                                              My.Settings.Item("IgnoredUpdates_RedInk") = String.Join(";", redInkItems)
                                              My.Settings.Item("IgnoredUpdates_Custom") = String.Join(";", customItems)
                                              My.Settings.Save()
                                          Catch ex As Exception
                                              Debug.WriteLine($"Error saving ignore lists: {ex.Message}")
                                          End Try

                                          ShowCustomMessageBox($"Ignore list updated successfully.")
                                          form.Close()
                                      End Sub

            ' Surface any same-process native host prompt (e.g. Word's "Save changes?")
            ' that would otherwise stay hidden behind this TopMost modal dialog.
            AttachForeignForegroundWatchdog(form)

            form.ShowDialog()
        End Sub

#End Region

#Region "Update Approval Dialog"

        ''' <summary>
        ''' Shows a dialog allowing the user to approve or reject individual detected changes.
        ''' </summary>
        ''' <param name="changes">Detected changes whose selection state is edited by the dialog.</param>
        ''' <returns>The dialog result indicating approve/reject/cancel.</returns>
        Private Shared Function ShowUpdateApprovalDialog(changes As List(Of IniParameterChange)) As UpdateApprovalResult
            Dim result As UpdateApprovalResult = UpdateApprovalResult.Cancel

            Dim form As New Form() With {
                .Text = $"{AN} - Configuration Updates Available",
                .Size = New Size(950, 600),
                .StartPosition = FormStartPosition.CenterParent,
                .FormBorderStyle = FormBorderStyle.Sizable,
                .Font = New Font("Segoe UI", 9.0F),
                .MinimumSize = New Size(800, 450),
                .AutoScaleMode = AutoScaleMode.Dpi,
                .TopMost = True
            }




            Try
                Dim bmp As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                form.Icon = Icon.FromHandle(bmp.GetHicon())
            Catch
            End Try

            ' Use TableLayoutPanel for proper scaling and layout
            Dim tblMain As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(15),
                .ColumnCount = 1,
                .RowCount = 3,
                .AutoSize = False
            }
            tblMain.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            tblMain.RowStyles.Add(New RowStyle(SizeType.AutoSize))  ' Header
            tblMain.RowStyles.Add(New RowStyle(SizeType.Percent, 100))  ' DataGridView
            tblMain.RowStyles.Add(New RowStyle(SizeType.AutoSize))  ' Buttons
            form.Controls.Add(tblMain)

            ' Row 0: Header
            Dim lblHeader As New Label() With {
                .Text = "The following configuration updates are available. Review and select which changes to apply:" & vbCrLf &
                        "(Items shown in red contain URL or path changes and are not selected by default for security reasons)",
                .Dock = DockStyle.Fill,
                .AutoSize = True,
                .Padding = New Padding(0, 0, 0, 10),
                .Margin = New Padding(0, 0, 0, 5)
            }
            tblMain.Controls.Add(lblHeader, 0, 0)

            ' Row 1: DataGridView for changes
            Dim dgv As New DataGridView() With {
    .Dock = DockStyle.Fill,
    .AllowUserToAddRows = False,
    .AllowUserToDeleteRows = False,
    .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
    .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
    .MultiSelect = False,
    .RowHeadersVisible = False,
    .Margin = New Padding(0, 5, 0, 10),
    .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
    .EnableHeadersVisualStyles = True
}

            ' Columns with proper minimum widths to show full header text
            Dim colApply As New DataGridViewCheckBoxColumn() With {
    .HeaderText = "Apply",
    .Name = "colApply",
    .Width = 50,
    .MinimumWidth = 50,
    .AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
}
            Dim colFile As New DataGridViewTextBoxColumn() With {
    .HeaderText = "File",
    .Name = "colFile",
    .ReadOnly = True,
    .Width = 120,
    .MinimumWidth = 60,
    .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
}
            Dim colSegment As New DataGridViewTextBoxColumn() With {
    .HeaderText = "Segment",
    .Name = "colSegment",
    .ReadOnly = True,
    .Width = 120,
    .MinimumWidth = 70,
    .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
}
            Dim colKey As New DataGridViewTextBoxColumn() With {
    .HeaderText = "Parameter",
    .Name = "colKey",
    .ReadOnly = True,
    .Width = 130,
    .MinimumWidth = 80,
    .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
}
            Dim colOld As New DataGridViewTextBoxColumn() With {
    .HeaderText = "Current Value",
    .Name = "colOld",
    .ReadOnly = True,
    .MinimumWidth = 100
}
            Dim colNew As New DataGridViewTextBoxColumn() With {
    .HeaderText = "New Value",
    .Name = "colNew",
    .ReadOnly = True,
    .MinimumWidth = 100
}

            dgv.Columns.AddRange(colApply, colFile, colSegment, colKey, colOld, colNew)

            ' Set column header height after columns are added to ensure headers fit
            AddHandler dgv.DataBindingComplete, Sub(sender, e)
                                                    dgv.AutoResizeColumnHeadersHeight()
                                                End Sub

            ' Add rows
            For Each change In changes
                Dim rowIndex = dgv.Rows.Add(
                change.IsSelected,
                change.IniFile,
                If(change.SegmentName, ""),
                change.ParameterKey,
                TruncateValue(change.OldValue, 100),
                TruncateValue(change.NewValue, 100)
            )

                ' Store the suspicious flag in the row's Tag for use in CellFormatting
                dgv.Rows(rowIndex).Tag = change.IsSuspicious

                ' Set default cell style for suspicious rows (will be overridden when selected)
                If change.IsSuspicious Then
                    dgv.Rows(rowIndex).DefaultCellStyle.ForeColor = Color.DarkRed
                    dgv.Rows(rowIndex).DefaultCellStyle.BackColor = Color.FromArgb(255, 230, 230)
                End If
            Next

            ' Handle CellFormatting to maintain suspicious highlighting even when row is selected
            AddHandler dgv.CellFormatting, Sub(sender, e)
                                               If e.RowIndex < 0 OrElse e.RowIndex >= dgv.Rows.Count Then Return

                                               Dim row = dgv.Rows(e.RowIndex)
                                               Dim isSuspicious = TypeOf row.Tag Is Boolean AndAlso CBool(row.Tag)

                                               If isSuspicious Then
                                                   If row.Selected Then
                                                       ' Use darker red colors when selected
                                                       e.CellStyle.BackColor = Color.FromArgb(220, 150, 150)
                                                       e.CellStyle.ForeColor = Color.DarkRed
                                                       e.CellStyle.SelectionBackColor = Color.FromArgb(220, 150, 150)
                                                       e.CellStyle.SelectionForeColor = Color.DarkRed
                                                   Else
                                                       ' Use light red colors when not selected
                                                       e.CellStyle.BackColor = Color.FromArgb(255, 230, 230)
                                                       e.CellStyle.ForeColor = Color.DarkRed
                                                   End If
                                               End If
                                           End Sub

            ' Ensure column headers are sized after rows are added
            dgv.AutoResizeColumnHeadersHeight()

            ' Handle checkbox changes
            AddHandler dgv.CellValueChanged, Sub(sender, e)
                                                 If e.ColumnIndex = 0 AndAlso e.RowIndex >= 0 Then
                                                     changes(e.RowIndex).IsSelected = CBool(dgv.Rows(e.RowIndex).Cells(0).Value)
                                                 End If
                                             End Sub

            AddHandler dgv.CurrentCellDirtyStateChanged, Sub()
                                                             If dgv.IsCurrentCellDirty Then
                                                                 dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
                                                             End If
                                                         End Sub

            tblMain.Controls.Add(dgv, 0, 1)

            ' Row 2: Buttons panel
            Dim pnlButtons As New FlowLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.LeftToRight,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .WrapContents = False,
                .Margin = New Padding(0, 5, 0, 5)
            }
            tblMain.Controls.Add(pnlButtons, 0, 2)

            Dim btnApprove As New Button() With {
                .Text = "Approve Selected",
                .AutoSize = True,
                .Padding = New Padding(10, 5, 10, 5),
                .Margin = New Padding(0, 0, 10, 0)
            }
            Dim btnReject As New Button() With {
                .Text = "Reject All",
                .AutoSize = True,
                .Padding = New Padding(10, 5, 10, 5)
            }

            pnlButtons.Controls.Add(btnApprove)
            pnlButtons.Controls.Add(btnReject)

            AddHandler btnApprove.Click, Sub()
                                             result = UpdateApprovalResult.Approve
                                             form.Close()
                                         End Sub

            AddHandler btnReject.Click, Sub()
                                            ' Deselect all
                                            For Each change In changes
                                                change.IsSelected = False
                                            Next
                                            result = UpdateApprovalResult.Reject
                                            form.Close()
                                        End Sub

            ' Surface any same-process native host prompt (e.g. Word's "Save changes?")
            ' that would otherwise stay hidden behind this TopMost modal dialog.
            AttachForeignForegroundWatchdog(form)

            form.ShowDialog()

            Return result
        End Function


        ''' <summary>
        ''' Shows a dialog offering to add rejected changes to the persisted ignore list.
        ''' </summary>
        ''' <param name="rejectedChanges">Changes that were rejected by the user.</param>
        Private Shared Sub ShowIgnoreConfirmationDialog(rejectedChanges As List(Of IniParameterChange))
            If rejectedChanges Is Nothing OrElse rejectedChanges.Count = 0 Then Return

            Dim form As New Form() With {
        .Text = $"{AN} - Ignore Future Updates?",
        .Size = New Size(750, 480),
        .StartPosition = FormStartPosition.CenterParent,
        .FormBorderStyle = FormBorderStyle.Sizable,
        .Font = New Font("Segoe UI", 9.0F),
        .MinimumSize = New Size(550, 350),
        .AutoScaleMode = AutoScaleMode.Dpi,
        .TopMost = True
            }


            Try
                Dim bmp As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                form.Icon = Icon.FromHandle(bmp.GetHicon())
            Catch
            End Try

            ' Use TableLayoutPanel for proper scaling and layout
            Dim tblMain As New TableLayoutPanel() With {
        .Dock = DockStyle.Fill,
        .Padding = New Padding(15),
        .ColumnCount = 1,
        .RowCount = 3,
        .AutoSize = False
    }
            tblMain.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            tblMain.RowStyles.Add(New RowStyle(SizeType.AutoSize))  ' Info label
            tblMain.RowStyles.Add(New RowStyle(SizeType.Percent, 100))  ' CheckedListBox
            tblMain.RowStyles.Add(New RowStyle(SizeType.AutoSize))  ' Buttons
            form.Controls.Add(tblMain)

            ' Row 0: Info label
            Dim lblInfo As New Label() With {
        .Text = $"The following parameters were not approved. Select which ones to ignore in future update checks (this will only result in additions to the Ignore list, not any removal from it; use the Freestyle command '{IniUpdateIgnored}' to manage it):",
        .Dock = DockStyle.Fill,
        .AutoSize = True,
        .Padding = New Padding(0, 0, 0, 5),
        .Margin = New Padding(0, 0, 0, 10)
    }
            tblMain.Controls.Add(lblInfo, 0, 0)

            ' Row 1: CheckedListBox
            Dim clb As New CheckedListBox() With {
        .Dock = DockStyle.Fill,
        .CheckOnClick = True,
        .Margin = New Padding(0, 5, 0, 10)
    }
            tblMain.Controls.Add(clb, 0, 1)

            For Each change In rejectedChanges
                clb.Items.Add($"[{change.IniFile}]{If(String.IsNullOrEmpty(change.SegmentName), "", $"[{change.SegmentName}]")}.{change.ParameterKey}", False)
            Next

            ' Row 2: Buttons panel
            Dim pnlButtons As New FlowLayoutPanel() With {
        .Dock = DockStyle.Fill,
        .FlowDirection = FlowDirection.LeftToRight,
        .AutoSize = True,
        .AutoSizeMode = AutoSizeMode.GrowAndShrink,
        .WrapContents = False,
        .Margin = New Padding(0, 5, 0, 5)
    }
            tblMain.Controls.Add(pnlButtons, 0, 2)

            Dim btnIgnore As New Button() With {
        .Text = "Ignore Selected",
        .AutoSize = True,
        .Padding = New Padding(10, 5, 10, 5),
        .Margin = New Padding(0, 0, 10, 0)
    }
            Dim btnIgnoreAll As New Button() With {
        .Text = "Ignore All",
        .AutoSize = True,
        .Padding = New Padding(10, 5, 10, 5),
        .Margin = New Padding(0, 0, 10, 0)
    }
            Dim btnAbort As New Button() With {
        .Text = "Don't Ignore Any",
        .AutoSize = True,
        .Padding = New Padding(10, 5, 10, 5)
    }

            pnlButtons.Controls.Add(btnIgnore)
            pnlButtons.Controls.Add(btnIgnoreAll)
            pnlButtons.Controls.Add(btnAbort)

            AddHandler btnAbort.Click, Sub() form.Close()

            AddHandler btnIgnoreAll.Click, Sub()
                                               ' Add all rejected changes to ignore list
                                               AddToIgnoreList(rejectedChanges)
                                               ShowCustomMessageBox($"{rejectedChanges.Count} parameter(s) will be ignored in future update checks (use the Freestyle command '{IniUpdateIgnored}' to manage).")
                                               form.Close()
                                           End Sub

            AddHandler btnIgnore.Click, Sub()
                                            Dim toIgnore As New List(Of IniParameterChange)()
                                            For i As Integer = 0 To clb.Items.Count - 1
                                                If clb.GetItemChecked(i) Then
                                                    toIgnore.Add(rejectedChanges(i))
                                                End If
                                            Next

                                            If toIgnore.Count > 0 Then
                                                AddToIgnoreList(toIgnore)
                                                ShowCustomMessageBox($"{toIgnore.Count} parameter(s) will be ignored in future update checks (use the Freestyle command '{IniUpdateIgnored}' to manage).")
                                            End If

                                            form.Close()
                                        End Sub

            form.ShowDialog()
        End Sub


        ''' <summary>
        ''' Truncates a string for display in UI lists.
        ''' </summary>
        ''' <param name="value">Value to truncate.</param>
        ''' <param name="maxLen">Maximum length of the returned string.</param>
        ''' <returns>The original string if short enough; otherwise a truncated string ending with <c>...</c>.</returns>
        Private Shared Function TruncateValue(value As String, maxLen As Integer) As String
            If String.IsNullOrEmpty(value) Then Return ""
            If value.Length <= maxLen Then Return value
            Return value.Substring(0, maxLen - 3) & "..."
        End Function

#End Region

#Region "Apply Updates"

        ''' <summary>
        ''' Writes approved changes to the corresponding local INI files and creates timestamped backups.
        ''' </summary>
        ''' <param name="changes">Approved changes to apply.</param>
        ''' <param name="context">Shared context used to resolve INI file paths.</param>
        Private Shared Sub ApplyApprovedUpdates(changes As List(Of IniParameterChange), context As ISharedContext)
            ' Group changes by file
            Dim byFile = changes.GroupBy(Function(c) c.IniFile)

            For Each fileGroup In byFile
                Dim filePath As String = Nothing
                Dim fileName = fileGroup.Key

                ' Determine the actual file path based on the filename
                Dim mainIniPath = GetDefaultINIPath(context.RDV)
                Dim mainIniName = Path.GetFileName(mainIniPath)

                ' Check main INI file
                If fileName.Equals(mainIniName, StringComparison.OrdinalIgnoreCase) Then
                    filePath = mainIniPath
                End If

                ' Check AlternateModelPath (independent check, not ElseIf)
                If filePath Is Nothing AndAlso Not String.IsNullOrWhiteSpace(context.INI_AlternateModelPath) Then
                    Dim altPath = ExpandEnvironmentVariables(context.INI_AlternateModelPath)
                    If Path.GetFileName(altPath).Equals(fileName, StringComparison.OrdinalIgnoreCase) Then
                        filePath = altPath
                    End If
                End If

                ' Check SpecialServicePath (independent check)
                If filePath Is Nothing AndAlso Not String.IsNullOrWhiteSpace(context.INI_SpecialServicePath) Then
                    Dim svcPath = ExpandEnvironmentVariables(context.INI_SpecialServicePath)
                    If Path.GetFileName(svcPath).Equals(fileName, StringComparison.OrdinalIgnoreCase) Then
                        filePath = svcPath
                    End If
                End If

                If String.IsNullOrWhiteSpace(filePath) OrElse Not File.Exists(filePath) Then
                    Debug.WriteLine($"ApplyApprovedUpdates: Could not resolve path for file '{fileName}' - skipping")
                    Continue For
                End If

                Try
                    ' Create timestamped backup (preserves original extension)
                    Dim backupPath = CreateTimestampedBackup(filePath)
                    If String.IsNullOrWhiteSpace(backupPath) Then
                        Debug.WriteLine($"Warning: Could not create backup for {filePath}")
                        LogIniUpdateEvent("Backup Warning", $"Could not create backup for {filePath}")
                    Else
                        LogIniUpdateEvent("Backup Created", $"{filePath} → {backupPath}")
                    End If

                    ' Read current content from original file
                    Dim lines = File.ReadAllLines(filePath, Encoding.UTF8).ToList()
                    Dim updatedLines As New List(Of String)()

                    ' Build lookup of changes by segment
                    Dim changesBySegment = fileGroup.GroupBy(Function(c) If(c.SegmentName, "")).
                ToDictionary(Function(g) g.Key,
                             Function(g) g.ToDictionary(Function(c) c.ParameterKey, StringComparer.OrdinalIgnoreCase))

                    Dim currentSegment As String = ""
                    Dim usedKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                    Dim segmentInsertionPoints As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

                    For i As Integer = 0 To lines.Count - 1
                        Dim line = lines(i)
                        Dim trimmed = line.Trim()

                        ' Track segment changes
                        If trimmed.StartsWith("[") AndAlso trimmed.EndsWith("]") Then
                            ' Before leaving current segment, record insertion point for new keys
                            If Not String.IsNullOrEmpty(currentSegment) OrElse i > 0 Then
                                segmentInsertionPoints(currentSegment) = updatedLines.Count
                            End If

                            currentSegment = trimmed.Substring(1, trimmed.Length - 2).Trim()
                            usedKeys.Clear()
                            updatedLines.Add(line)
                            Continue For
                        End If

                        ' Check if this is a key=value line
                        Dim eqIndex = trimmed.IndexOf("="c)
                        If eqIndex > 0 AndAlso Not trimmed.StartsWith(";") Then
                            Dim key = trimmed.Substring(0, eqIndex).Trim()

                            ' Check if this key needs updating
                            If changesBySegment.ContainsKey(currentSegment) AndAlso
                       changesBySegment(currentSegment).ContainsKey(key) Then
                                Dim change = changesBySegment(currentSegment)(key)
                                updatedLines.Add($"{key} = {change.NewValue}")
                                usedKeys.Add(key)
                                Continue For
                            End If
                        End If

                        updatedLines.Add(line)
                    Next

                    ' Record final segment insertion point
                    segmentInsertionPoints(currentSegment) = updatedLines.Count

                    ' Now add NEW keys that were not found in the file
                    For Each segmentGroup In changesBySegment
                        Dim segName As String = segmentGroup.Key
                        Dim segChanges As Dictionary(Of String, IniParameterChange) = segmentGroup.Value

                        ' Find keys that are new (OldValue = "(new key)")
                        Dim newKeysList As New List(Of KeyValuePair(Of String, IniParameterChange))()
                        For Each kvp In segChanges
                            ' Find the original change to check OldValue
                            Dim originalChange As IniParameterChange = Nothing
                            For Each chg In fileGroup
                                If (If(chg.SegmentName, "") = segName) AndAlso
                           chg.ParameterKey.Equals(kvp.Key, StringComparison.OrdinalIgnoreCase) Then
                                    originalChange = chg
                                    Exit For
                                End If
                            Next
                            If originalChange IsNot Nothing AndAlso originalChange.OldValue = "(new key)" Then
                                newKeysList.Add(kvp)
                            End If
                        Next

                        If newKeysList.Count > 0 Then
                            ' Determine insertion point for this segment
                            Dim insertAt As Integer = updatedLines.Count ' Default: end of file
                            If segmentInsertionPoints.ContainsKey(segName) Then
                                insertAt = segmentInsertionPoints(segName)
                            End If

                            ' Add a blank line before new keys if the previous line is not already blank
                            If insertAt > 0 AndAlso Not String.IsNullOrWhiteSpace(updatedLines(insertAt - 1)) Then
                                updatedLines.Insert(insertAt, "")
                                insertAt += 1
                                ' Adjust all insertion points since we added a line
                                For Each key In segmentInsertionPoints.Keys.ToList()
                                    If segmentInsertionPoints(key) >= insertAt - 1 Then
                                        segmentInsertionPoints(key) += 1
                                    End If
                                Next
                            End If

                            ' Insert new keys
                            Dim offset As Integer = 0
                            For Each kvp In newKeysList
                                Dim newLine = $"{kvp.Key} = {kvp.Value.NewValue}"
                                updatedLines.Insert(insertAt + offset, newLine)
                                offset += 1
                            Next

                            ' Add a blank line after new keys if the next line exists and is not already blank
                            Dim afterInsertPos = insertAt + offset
                            If afterInsertPos < updatedLines.Count AndAlso Not String.IsNullOrWhiteSpace(updatedLines(afterInsertPos)) Then
                                updatedLines.Insert(afterInsertPos, "")
                                offset += 1
                            End If

                            ' Adjust subsequent insertion points
                            For Each key In segmentInsertionPoints.Keys.ToList()
                                If segmentInsertionPoints(key) >= insertAt Then
                                    segmentInsertionPoints(key) += offset
                                End If
                            Next
                        End If
                    Next

                    ' Write updated content
                    File.WriteAllLines(filePath, updatedLines, Encoding.UTF8)

                Catch ex As Exception
                    Debug.WriteLine($"Error applying updates to {filePath}: {ex.Message}")
                    ShowCustomMessageBox($"Failed to update {fileName}: {ex.Message}")
                End Try
            Next
        End Sub

#End Region

#Region "Client Identification"

        ''' <summary>
        ''' Returns the current client identifier used for UpdateIniClients matching.
        ''' This uses the Windows computer name (Environment.MachineName).
        ''' </summary>
        ''' <returns>The client identifier string (computer name).</returns>
        Public Shared Function GetCurrentClientIdentifier() As String
            Try
                Return Environment.MachineName
            Catch
                Return String.Empty
            End Try
        End Function

        ''' <summary>
        ''' Determines whether the current client is allowed to perform INI updates
        ''' based on the UpdateIniClients parameter.
        ''' </summary>
        ''' <returns>True if this client can update; False otherwise.</returns>
        Private Shared Function IsClientAllowedToUpdate() As Boolean
            Try
                ' Get the UpdateIniClients setting from context
                Dim UpdateIniClients As String = _iniUpdateContext.INI_UpdateIniClients

                ' If not configured, any client can update
                If String.IsNullOrWhiteSpace(UpdateIniClients) Then
                    Return True
                End If

                Dim currentClient As String = GetCurrentClientIdentifier()
                If String.IsNullOrWhiteSpace(currentClient) Then
                    LogIniUpdateEvent("Client Check", "Could not determine current client identifier - allowing update")
                    Return True
                End If

                ' Parse the comma-separated list of allowed clients
                Dim allowedClients = UpdateIniClients.Split(","c).
                    Select(Function(c) c.Trim()).
                    Where(Function(c) Not String.IsNullOrWhiteSpace(c)).
                    ToList()

                If allowedClients.Count = 0 Then
                    Return True
                End If

                ' Check if current client is in the allowed list (case-insensitive)
                Dim isAllowed = allowedClients.Any(Function(c) c.Equals(currentClient, StringComparison.OrdinalIgnoreCase))

                If Not isAllowed Then
                    LogIniUpdateEvent("Client Check",
                        $"Client '{currentClient}' is not in UpdateIniClients list: {UpdateIniClients} - skipping update")
                End If

                Return isAllowed

            Catch ex As Exception
                Debug.WriteLine($"Error checking client authorization: {ex.Message}")
                Return True ' On error, allow update to proceed
            End Try
        End Function

#End Region

#Region "Timestamped Backup Helper"

        ''' <summary>
        ''' Creates a timestamped backup of the specified file, preserving the original extension.
        ''' Format: filename.ext_yyyyMMdd_HHmmss_fff.bak
        ''' </summary>
        ''' <param name="filePath">Path to the file to back up.</param>
        ''' <returns>The path to the created backup file, or Nothing on failure.</returns>
        Private Shared Function CreateTimestampedBackup(filePath As String) As String
            Try
                If String.IsNullOrWhiteSpace(filePath) OrElse Not File.Exists(filePath) Then
                    Return Nothing
                End If

                Dim directory As String = Path.GetDirectoryName(filePath)
                Dim fileName As String = Path.GetFileName(filePath)
                Dim timestamp As String = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff")
                Dim backupFileName As String = $"{fileName}_{timestamp}.bak"
                Dim backupPath As String = Path.Combine(directory, backupFileName)

                File.Copy(filePath, backupPath, overwrite:=False)

                Return backupPath
            Catch ex As Exception
                Debug.WriteLine($"Failed to create timestamped backup for {filePath}: {ex.Message}")
                Return Nothing
            End Try
        End Function

#End Region

#Region "Ed25519 Digital Signature"

        ''' <summary>
        ''' Generates a new Ed25519 key pair encoded as Base64 strings.
        ''' </summary>
        ''' <returns>
        ''' A tuple containing <c>PublicKey</c> and <c>PrivateKey</c> as Base64 strings, or (<c>Nothing</c>, <c>Nothing</c>) on failure.
        ''' </returns>
        Public Shared Function GenerateEd25519KeyPair() As (PublicKey As String, PrivateKey As String)
            Try
                Dim keyGen As New Ed25519KeyPairGenerator()
                keyGen.Init(New Ed25519KeyGenerationParameters(New SecureRandom()))
                Dim keyPair = keyGen.GenerateKeyPair()

                Dim publicKey = DirectCast(keyPair.Public, Ed25519PublicKeyParameters)
                Dim privateKey = DirectCast(keyPair.Private, Ed25519PrivateKeyParameters)

                Dim pubBytes = publicKey.GetEncoded()
                Dim privBytes = privateKey.GetEncoded()

                Return (System.Convert.ToBase64String(pubBytes), System.Convert.ToBase64String(privBytes))

            Catch ex As Exception
                Debug.WriteLine($"Error generating Ed25519 keypair: {ex.Message}")
                Return (Nothing, Nothing)
            End Try
        End Function

        ''' <summary>
        ''' Signs a file's bytes using an Ed25519 private key and writes the Base64 signature to <c>filePath + ".sig"</c>.
        ''' </summary>
        ''' <param name="filePath">The file to sign.</param>
        ''' <param name="base64PrivateKey">Base64-encoded Ed25519 private key.</param>
        ''' <returns><c>True</c> if the signature file was created; otherwise <c>False</c>.</returns>
        Public Shared Function SignUpdateFile(filePath As String, base64PrivateKey As String) As Boolean
            Try
                If Not File.Exists(filePath) Then
                    ShowCustomMessageBox($"File not found: {filePath}")
                    Return False
                End If

                Dim privKeyBytes = System.Convert.FromBase64String(base64PrivateKey)
                Dim privateKey As New Ed25519PrivateKeyParameters(privKeyBytes, 0)

                Dim content = File.ReadAllBytes(filePath)

                Dim signer As New Ed25519Signer()
                signer.Init(True, privateKey)
                signer.BlockUpdate(content, 0, content.Length)
                Dim signature = signer.GenerateSignature()

                Dim sigPath = filePath & ".sig"
                File.WriteAllText(sigPath, System.Convert.ToBase64String(signature))

                Return True

            Catch ex As Exception
                ShowCustomMessageBox($"Error signing file: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Verifies that the specified Base64 signature matches the specified content and Base64 public key.
        ''' </summary>
        ''' <param name="content">Content string to verify (UTF-8 bytes are used).</param>
        ''' <param name="signatureBase64">Base64-encoded signature bytes.</param>
        ''' <param name="publicKeyBase64">Base64-encoded Ed25519 public key bytes.</param>
        ''' <returns><c>True</c> if the signature is valid; otherwise <c>False</c>.</returns>
        Private Shared Function VerifyEd25519Signature(contentBytes As Byte(), signatureBase64 As String, publicKeyBase64 As String) As Boolean
            Try
                Dim pubKeyBytes = System.Convert.FromBase64String(publicKeyBase64)
                Dim publicKey As New Ed25519PublicKeyParameters(pubKeyBytes, 0)

                Dim signatureBytes = System.Convert.FromBase64String(signatureBase64)

                Dim verifier As New Ed25519Signer()
                verifier.Init(False, publicKey)
                verifier.BlockUpdate(contentBytes, 0, contentBytes.Length)

                Return verifier.VerifySignature(signatureBytes)

            Catch ex As Exception
                Debug.WriteLine($"Signature verification error: {ex.Message}")
                Return False
            End Try
        End Function

        Private Shared Function VerifyEd25519Signature(content As String, signatureBase64 As String, publicKeyBase64 As String) As Boolean
            If content Is Nothing Then Return False
            Return VerifyEd25519Signature(Encoding.UTF8.GetBytes(content), signatureBase64, publicKeyBase64)
        End Function

        Private Shared Function VerifyEd25519SignatureCompatible(contentBytes As Byte(), signatureBase64 As String, publicKeyBase64 As String) As Boolean
            If contentBytes Is Nothing Then Return False

            ' Preferred path: verify the exact bytes that were signed.
            If VerifyEd25519Signature(contentBytes, signatureBase64, publicKeyBase64) Then
                Return True
            End If

            ' Compatibility fallback: preserve support for any legacy text-based signatures.
            Try
                Dim decodedContent = DecodeContentText(contentBytes)
                Return VerifyEd25519Signature(decodedContent, signatureBase64, publicKeyBase64)
            Catch ex As Exception
                Debug.WriteLine($"Legacy signature compatibility check failed: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Verifies <c>filePath</c> against its adjacent <c>.sig</c> file using the provided public key.
        ''' </summary>
        ''' <param name="filePath">Source file path.</param>
        ''' <param name="publicKeyBase64">Base64-encoded Ed25519 public key.</param>
        ''' <returns><c>True</c> if the signature validates; otherwise <c>False</c>.</returns>
        Public Shared Function VerifySignatureFile(filePath As String, publicKeyBase64 As String) As Boolean
            Try
                If Not File.Exists(filePath) Then
                    ShowCustomMessageBox($"File not found: {filePath}")
                    Return False
                End If

                Dim sigPath = filePath & ".sig"
                If Not File.Exists(sigPath) Then
                    ShowCustomMessageBox($"Signature file not found: {sigPath}")
                    Return False
                End If

                Dim contentBytes = File.ReadAllBytes(filePath)
                Dim signature = File.ReadAllText(sigPath, Encoding.UTF8).Trim()

                Return VerifyEd25519SignatureCompatible(contentBytes, signature, publicKeyBase64)

            Catch ex As Exception
                ShowCustomMessageBox($"Error verifying signature: {ex.Message}")
                Return False
            End Try
        End Function

#End Region

#Region "Signature Management UI"

        ''' <summary>
        ''' Shows the signature management UI for generating keys, signing files, and verifying signatures.
        ''' </summary>
        Public Shared Sub ShowSignatureManagementDialog()
            Dim form As New Form() With {
                .Text = $"{AN} - Update Signature Management",
                .Size = New Size(750, 380),
                .StartPosition = FormStartPosition.CenterParent,
                .FormBorderStyle = FormBorderStyle.Sizable,
                .Font = New Font("Segoe UI", 9.0F),
                .MinimumSize = New Size(700, 350),
                .TopMost = True
            }



            Try
                Dim bmp As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                form.Icon = Icon.FromHandle(bmp.GetHicon())
            Catch
            End Try

            Dim tabControl As New TabControl() With {.Dock = DockStyle.Fill}
            form.Controls.Add(tabControl)

            ' =====================================================================
            ' Tab 1: Generate Keypair
            ' =====================================================================
            Dim tabGenerate As New TabPage("Generate Keypair")
            tabControl.TabPages.Add(tabGenerate)

            Dim pnlGenerate As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(15),
                .ColumnCount = 2,
                .RowCount = 5
            }
            pnlGenerate.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            pnlGenerate.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            ' Set row heights to prevent excess space
            pnlGenerate.RowStyles.Add(New RowStyle(SizeType.Absolute, 50))  ' Info
            pnlGenerate.RowStyles.Add(New RowStyle(SizeType.Absolute, 35))  ' Generate button
            pnlGenerate.RowStyles.Add(New RowStyle(SizeType.Absolute, 30))  ' Public key
            pnlGenerate.RowStyles.Add(New RowStyle(SizeType.Absolute, 30))  ' Private key
            pnlGenerate.RowStyles.Add(New RowStyle(SizeType.Absolute, 40))  ' Copy buttons
            tabGenerate.Controls.Add(pnlGenerate)

            ' Row 0: Info label
            Dim lblGenInfo As New Label() With {
                .Text = "Generate a new Ed25519 keypair for signing update files." & vbCrLf &
                        "The public key goes into UpdateSource. Keep the private key secure!",
                .Dock = DockStyle.Fill,
                .AutoSize = False
            }
            pnlGenerate.Controls.Add(lblGenInfo, 0, 0)
            pnlGenerate.SetColumnSpan(lblGenInfo, 2)

            ' Row 1: Generate button
            Dim btnGenerate As New Button() With {
                .Text = "Generate New Keypair",
                .AutoSize = True
            }
            pnlGenerate.Controls.Add(btnGenerate, 0, 1)

            ' Row 2: Public Key label and textbox
            Dim lblPubKey As New Label() With {
                .Text = "Public Key (for UpdateSource):",
                .AutoSize = True,
                .Anchor = AnchorStyles.Left,
                .Padding = New Padding(0, 0, 10, 0)
            }
            pnlGenerate.Controls.Add(lblPubKey, 0, 2)

            Dim txtPubKey As New TextBox() With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .BackColor = SystemColors.Window
            }
            pnlGenerate.Controls.Add(txtPubKey, 1, 2)

            ' Row 3: Private Key label and textbox
            Dim lblPrivKey As New Label() With {
                .Text = "Private Key (KEEP SECRET!):",
                .AutoSize = True,
                .Anchor = AnchorStyles.Left,
                .ForeColor = Color.DarkRed,
                .Padding = New Padding(0, 0, 10, 0)
            }
            pnlGenerate.Controls.Add(lblPrivKey, 0, 3)

            Dim txtPrivKey As New TextBox() With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .BackColor = SystemColors.Window,
                .UseSystemPasswordChar = False
            }
            pnlGenerate.Controls.Add(txtPrivKey, 1, 3)

            ' Row 4: Copy buttons
            Dim pnlCopyButtons As New FlowLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.LeftToRight
            }
            pnlGenerate.Controls.Add(pnlCopyButtons, 1, 4)

            Dim btnCopyPub As New Button() With {.Text = "Copy Public Key", .AutoSize = True}
            Dim btnCopyPriv As New Button() With {.Text = "Copy Private Key", .AutoSize = True}
            Dim btnCopyBoth As New Button() With {.Text = "Copy Both to Clipboard", .AutoSize = True}
            pnlCopyButtons.Controls.AddRange({btnCopyPub, btnCopyPriv, btnCopyBoth})
            AddHandler btnGenerate.Click, Sub()
                                              Dim keys = GenerateEd25519KeyPair()
                                              If keys.PublicKey IsNot Nothing Then
                                                  txtPubKey.Text = keys.PublicKey
                                                  txtPrivKey.Text = keys.PrivateKey
                                                  ShowCustomMessageBox("Keypair generated successfully!" & vbCrLf & vbCrLf &
                                                      "Public Key: Use this in the UpdateSource parameter of your INI segments." & vbCrLf &
                                                      "Private Key: Store securely and use for signing update files.")
                                              End If
                                          End Sub

            AddHandler btnCopyPub.Click, Sub()
                                             If Not String.IsNullOrEmpty(txtPubKey.Text) Then
                                                 Try
                                                     Clipboard.SetText(txtPubKey.Text)
                                                     ShowCustomMessageBox("Public key copied to clipboard.")
                                                 Catch ex As Exception
                                                     ShowCustomMessageBox("Failed to copy: " & ex.Message)
                                                 End Try
                                             End If
                                         End Sub

            AddHandler btnCopyPriv.Click, Sub()
                                              If Not String.IsNullOrEmpty(txtPrivKey.Text) Then
                                                  Try
                                                      Clipboard.SetText(txtPrivKey.Text)
                                                      ShowCustomMessageBox("Private key copied to clipboard. Keep it secure!")
                                                  Catch ex As Exception
                                                      ShowCustomMessageBox("Failed to copy: " & ex.Message)
                                                  End Try
                                              End If
                                          End Sub

            AddHandler btnCopyBoth.Click, Sub()
                                              If Not String.IsNullOrEmpty(txtPubKey.Text) Then
                                                  Try
                                                      Dim text = $"=== Ed25519 KEYPAIR ==={vbCrLf}{vbCrLf}" &
                                                                 $"PUBLIC KEY (for UpdateSource):{vbCrLf}{txtPubKey.Text}{vbCrLf}{vbCrLf}" &
                                                                 $"PRIVATE KEY (KEEP SECRET - for signing):{vbCrLf}{txtPrivKey.Text}{vbCrLf}"
                                                      Clipboard.SetText(text)
                                                      ShowCustomMessageBox("Both keys copied to clipboard.")
                                                  Catch ex As Exception
                                                      ShowCustomMessageBox("Failed to copy: " & ex.Message)
                                                  End Try
                                              End If
                                          End Sub

            ' =====================================================================
            ' Tab 2: Sign File
            ' =====================================================================
            Dim tabSign As New TabPage("Sign File")
            tabControl.TabPages.Add(tabSign)

            Dim pnlSign As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(15),
                .ColumnCount = 3,
                .RowCount = 6
            }
            pnlSign.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120))
            pnlSign.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            pnlSign.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 40))
            tabSign.Controls.Add(pnlSign)

            ' Row 0: Info
            Dim lblSignInfo As New Label() With {
                .Text = "Sign an update INI file to create a .sig signature file alongside it." & vbCrLf &
                        "The signature file must be placed next to the update source file.",
                .Dock = DockStyle.Fill,
                .Height = 50
            }
            pnlSign.Controls.Add(lblSignInfo, 0, 0)
            pnlSign.SetColumnSpan(lblSignInfo, 3)

            ' Row 1: File to sign
            Dim lblSignFile As New Label() With {
                .Text = "File to Sign:",
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleLeft
            }
            pnlSign.Controls.Add(lblSignFile, 0, 1)

            Dim txtSignFile As New TextBox() With {.Dock = DockStyle.Fill}
            pnlSign.Controls.Add(txtSignFile, 1, 1)

            Dim btnBrowseSign As New Button() With {.Text = "...", .Dock = DockStyle.Fill}
            pnlSign.Controls.Add(btnBrowseSign, 2, 1)

            ' Row 2: Private key
            Dim lblSignPrivKey As New Label() With {
                .Text = "Private Key:",
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleLeft
            }
            pnlSign.Controls.Add(lblSignPrivKey, 0, 2)

            Dim txtSignPrivKey As New TextBox() With {
                .Dock = DockStyle.Fill,
                .UseSystemPasswordChar = True
            }
            pnlSign.Controls.Add(txtSignPrivKey, 1, 2)
            pnlSign.SetColumnSpan(txtSignPrivKey, 2)

            ' Row 3: Show/hide password toggle
            Dim chkShowPrivKey As New CheckBox() With {
                .Text = "Show private key",
                .Dock = DockStyle.Left,
                .AutoSize = True
            }
            pnlSign.Controls.Add(chkShowPrivKey, 1, 3)

            AddHandler chkShowPrivKey.CheckedChanged, Sub()
                                                          txtSignPrivKey.UseSystemPasswordChar = Not chkShowPrivKey.Checked
                                                      End Sub

            ' Row 4: Sign button
            Dim btnSign As New Button() With {
                .Text = "Sign File",
                .Width = 120,
                .Height = 35,
                .Dock = DockStyle.Left
            }
            pnlSign.Controls.Add(btnSign, 1, 4)

            ' Row 5: Result
            Dim lblSignResult As New Label() With {
                .Text = "",
                .Dock = DockStyle.Fill,
                .ForeColor = Color.DarkGreen
            }
            pnlSign.Controls.Add(lblSignResult, 0, 5)
            pnlSign.SetColumnSpan(lblSignResult, 3)

            AddHandler btnBrowseSign.Click, Sub()
                                                Using ofd As New OpenFileDialog()
                                                    ofd.Title = "Select File to Sign"
                                                    ofd.Filter = "Text Files|*.txt|INI Files|*.ini|All Files|*.*"
                                                    If ofd.ShowDialog() = DialogResult.OK Then
                                                        txtSignFile.Text = ofd.FileName
                                                    End If
                                                End Using
                                            End Sub

            AddHandler btnSign.Click, Sub()
                                          lblSignResult.Text = ""
                                          lblSignResult.ForeColor = Color.DarkGreen

                                          If String.IsNullOrWhiteSpace(txtSignFile.Text) Then
                                              ShowCustomMessageBox("Please select a file to sign.")
                                              Return
                                          End If

                                          If Not File.Exists(txtSignFile.Text) Then
                                              ShowCustomMessageBox("The selected file does not exist.")
                                              Return
                                          End If

                                          If String.IsNullOrWhiteSpace(txtSignPrivKey.Text) Then
                                              ShowCustomMessageBox("Please enter the private key.")
                                              Return
                                          End If

                                          Try
                                              If SignUpdateFile(txtSignFile.Text, txtSignPrivKey.Text.Trim()) Then
                                                  lblSignResult.Text = $"✓ Signature created: {txtSignFile.Text}.sig"
                                                  lblSignResult.ForeColor = Color.DarkGreen
                                                  ShowCustomMessageBox($"File signed successfully!{vbCrLf}{vbCrLf}" &
                                                      $"Signature saved to:{vbCrLf}{txtSignFile.Text}.sig{vbCrLf}{vbCrLf}" &
                                                      "Upload both the INI file and its .sig file to your update location.")
                                              Else
                                                  lblSignResult.Text = "✗ Signing failed"
                                                  lblSignResult.ForeColor = Color.DarkRed
                                              End If
                                          Catch ex As Exception
                                              lblSignResult.Text = $"✗ Error: {ex.Message}"
                                              lblSignResult.ForeColor = Color.DarkRed
                                          End Try
                                      End Sub

            ' =====================================================================
            ' Tab 3: Verify Signature
            ' =====================================================================
            Dim tabVerify As New TabPage("Verify Signature")
            tabControl.TabPages.Add(tabVerify)

            Dim pnlVerify As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(15),
                .ColumnCount = 3,
                .RowCount = 5
            }
            pnlVerify.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            pnlVerify.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            pnlVerify.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 40))
            ' Set row heights to prevent excess space
            pnlVerify.RowStyles.Add(New RowStyle(SizeType.Absolute, 50))  ' Info
            pnlVerify.RowStyles.Add(New RowStyle(SizeType.Absolute, 30))  ' File to verify
            pnlVerify.RowStyles.Add(New RowStyle(SizeType.Absolute, 30))  ' Public key
            pnlVerify.RowStyles.Add(New RowStyle(SizeType.Absolute, 40))  ' Verify button
            pnlVerify.RowStyles.Add(New RowStyle(SizeType.Absolute, 50))  ' Result
            tabVerify.Controls.Add(pnlVerify)

            ' Row 0: Info
            Dim lblVerifyInfo As New Label() With {
                .Text = "Verify that a file's signature is valid using the public key." & vbCrLf &
                        "The .sig file must exist alongside the file being verified.",
                .Dock = DockStyle.Fill,
                .AutoSize = False
            }
            pnlVerify.Controls.Add(lblVerifyInfo, 0, 0)
            pnlVerify.SetColumnSpan(lblVerifyInfo, 3)

            ' Row 1: File to verify
            Dim lblVerifyFile As New Label() With {
                .Text = "File to Verify:",
                .AutoSize = True,
                .Anchor = AnchorStyles.Left,
                .Padding = New Padding(0, 0, 10, 0)
            }
            pnlVerify.Controls.Add(lblVerifyFile, 0, 1)

            Dim txtVerifyFile As New TextBox() With {.Dock = DockStyle.Fill}
            pnlVerify.Controls.Add(txtVerifyFile, 1, 1)

            Dim btnBrowseVerify As New Button() With {.Text = "...", .Dock = DockStyle.Fill}
            pnlVerify.Controls.Add(btnBrowseVerify, 2, 1)

            ' Row 2: Public key
            Dim lblVerifyPubKey As New Label() With {
                .Text = "Public Key:",
                .AutoSize = True,
                .Anchor = AnchorStyles.Left,
                .Padding = New Padding(0, 0, 10, 0)
            }
            pnlVerify.Controls.Add(lblVerifyPubKey, 0, 2)

            Dim txtVerifyPubKey As New TextBox() With {.Dock = DockStyle.Fill}
            pnlVerify.Controls.Add(txtVerifyPubKey, 1, 2)
            pnlVerify.SetColumnSpan(txtVerifyPubKey, 2)

            ' Row 3: Verify button
            Dim btnVerify As New Button() With {
                .Text = "Verify Signature",
                .AutoSize = True
            }
            pnlVerify.Controls.Add(btnVerify, 1, 3)

            ' Row 4: Result
            Dim lblVerifyResult As New Label() With {
                .Text = "",
                .Dock = DockStyle.Fill,
                .Font = New Font("Segoe UI", 11.0F, FontStyle.Bold),
                .AutoSize = False
            }
            pnlVerify.Controls.Add(lblVerifyResult, 0, 4)
            pnlVerify.SetColumnSpan(lblVerifyResult, 3)

            AddHandler btnBrowseVerify.Click, Sub()
                                                  Using ofd As New OpenFileDialog()
                                                      ofd.Title = "Select File to Verify"
                                                      ofd.Filter = "Text Files|*.txt|INI Files|*.ini|All Files|*.*"
                                                      If ofd.ShowDialog() = DialogResult.OK Then
                                                          txtVerifyFile.Text = ofd.FileName
                                                      End If
                                                  End Using
                                              End Sub

            AddHandler btnVerify.Click, Sub()
                                            lblVerifyResult.Text = ""

                                            If String.IsNullOrWhiteSpace(txtVerifyFile.Text) Then
                                                ShowCustomMessageBox("Please select a file to verify.")
                                                Return
                                            End If

                                            If Not File.Exists(txtVerifyFile.Text) Then
                                                ShowCustomMessageBox("The selected file does not exist.")
                                                Return
                                            End If

                                            Dim sigPath = txtVerifyFile.Text & ".sig"
                                            If Not File.Exists(sigPath) Then
                                                ShowCustomMessageBox($"Signature file not found:{vbCrLf}{sigPath}")
                                                Return
                                            End If

                                            If String.IsNullOrWhiteSpace(txtVerifyPubKey.Text) Then
                                                ShowCustomMessageBox("Please enter the public key.")
                                                Return
                                            End If

                                            Try
                                                Dim isValid = VerifySignatureFile(txtVerifyFile.Text, txtVerifyPubKey.Text.Trim())
                                                If isValid Then
                                                    lblVerifyResult.Text = "✓ SIGNATURE VALID - File is authentic"
                                                    lblVerifyResult.ForeColor = Color.DarkGreen
                                                Else
                                                    lblVerifyResult.Text = "✗ SIGNATURE INVALID - File may have been modified!"
                                                    lblVerifyResult.ForeColor = Color.DarkRed
                                                End If
                                            Catch ex As Exception
                                                lblVerifyResult.Text = $"✗ Verification failed: {ex.Message}"
                                                lblVerifyResult.ForeColor = Color.DarkRed
                                            End Try
                                        End Sub

            ' =====================================================================
            ' Tab 4: Help / Instructions
            ' =====================================================================
            Dim tabHelp As New TabPage("Help")
            tabControl.TabPages.Add(tabHelp)

            Dim txtHelp As New TextBox() With {
                .Dock = DockStyle.Fill,
                .Multiline = True,
                .ReadOnly = True,
                .ScrollBars = ScrollBars.Vertical,
                .BackColor = SystemColors.Window,
                .Font = New Font("Segoe UI", 9.5F),
                .Text = GetSignatureHelpText()
            }
            tabHelp.Controls.Add(txtHelp)

            form.ShowDialog()

        End Sub

        ''' <summary>
        ''' Returns help text for the signature management dialog.
        ''' </summary>
        Private Shared Function GetSignatureHelpText() As String
            Return $"=== {AN} Update Signature System ===" & vbCrLf & vbCrLf &
                "This tool uses Ed25519 digital signatures to ensure the authenticity " &
                "and integrity of configuration updates. Ed25519 is a modern, secure " &
                "signature algorithm that provides strong protection against tampering." & vbCrLf & vbCrLf &
                "=== HOW IT WORKS ===" & vbCrLf & vbCrLf &
                "1. GENERATE KEYPAIR" & vbCrLf &
                "   - Go to 'Generate Keypair' tab and click 'Generate New Keypair'" & vbCrLf &
                "   - You'll receive a Public Key and a Private Key" & vbCrLf &
                "   - Store the Private Key securely (e.g., in a password manager)" & vbCrLf &
                "   - The Public Key will be included in UpdateSource entries" & vbCrLf & vbCrLf &
                "2. CONFIGURE UPDATE SOURCE" & vbCrLf &
                "   In your INI files, configure UpdateSource as:" & vbCrLf &
                "   UpdateSource = path; keys; public_key" & vbCrLf & vbCrLf &
                "   Example:" & vbCrLf &
                "   UpdateSource = https://example.com/updates/models.ini; all; MCow..." & vbCrLf & vbCrLf &
                "   Where:" & vbCrLf &
                "   - path: URL or file path to the update INI file" & vbCrLf &
                "   - keys: 'all' or comma-separated list of keys to update" & vbCrLf &
                "   - public_key: Base64-encoded Ed25519 public key" & vbCrLf & vbCrLf &
                "3. SIGN UPDATE FILES" & vbCrLf &
                "   - Create/modify your update INI file" & vbCrLf &
                "   - Go to 'Sign File' tab" & vbCrLf &
                "   - Select the file and enter your Private Key" & vbCrLf &
                "   - Click 'Sign File' to create the .sig file" & vbCrLf &
                "   - Upload BOTH the INI file and its .sig file to your server" & vbCrLf & vbCrLf &
                "4. VERIFY SIGNATURES (Optional)" & vbCrLf &
                "   - Use the 'Verify Signature' tab to manually test" & vbCrLf &
                "   - This is useful for troubleshooting" & vbCrLf & vbCrLf &
                "=== SECURITY NOTES ===" & vbCrLf & vbCrLf &
                "• NEVER share your Private Key" & vbCrLf &
                "• The Private Key is needed only for signing (administrator)" & vbCrLf &
                "• The Public Key is safe to distribute in UpdateSource" & vbCrLf &
                "• If your Private Key is compromised, generate a new keypair" & vbCrLf &
                "  and update all UpdateSource entries with the new Public Key" & vbCrLf & vbCrLf &
                "=== FILE STRUCTURE ===" & vbCrLf & vbCrLf &
                "When you sign 'updates.ini', the system creates 'updates.ini.sig'" & vbCrLf &
                "Both files must be uploaded to the same location:" & vbCrLf & vbCrLf &
                "   https://example.com/updates/models.ini" & vbCrLf &
                "   https://example.com/updates/models.ini.sig" & vbCrLf & vbCrLf &
                "The update checker downloads both files and verifies the signature " &
                "before applying any changes."
        End Function

#End Region



#Region "Batch Signing Utility"

        ''' <summary>
        ''' Signs multiple files using the same private key and returns per-file results.
        ''' </summary>
        ''' <param name="filePaths">File paths to sign.</param>
        ''' <param name="base64PrivateKey">Base64-encoded Ed25519 private key.</param>
        ''' <returns>Dictionary mapping file paths to a result string.</returns>
        Public Shared Function BatchSignFiles(filePaths As String(), base64PrivateKey As String) As Dictionary(Of String, String)
            Dim results As New Dictionary(Of String, String)()

            If filePaths Is Nothing OrElse filePaths.Length = 0 Then
                Return results
            End If

            For Each filePath In filePaths
                Try
                    If SignUpdateFile(filePath, base64PrivateKey) Then
                        results(filePath) = "Success"
                    Else
                        results(filePath) = "Failed (unknown error)"
                    End If
                Catch ex As Exception
                    results(filePath) = $"Error: {ex.Message}"
                End Try
            Next

            Return results
        End Function

        ''' <summary>
        ''' Shows a dialog for selecting multiple files and signing them with a single private key.
        ''' </summary>
        Public Shared Sub ShowBatchSigningDialog()
            Dim form As New Form() With {
                .Text = $"{AN} - Batch Sign Files",
                .Size = New Size(750, 520),
                .StartPosition = FormStartPosition.CenterParent,
                .FormBorderStyle = FormBorderStyle.Sizable,
                .Font = New Font("Segoe UI", 9.0F),
                .MinimumSize = New Size(600, 400),
                .AutoScaleMode = AutoScaleMode.Dpi,
                .TopMost = True
                            }



            Try
                Dim bmp As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                form.Icon = Icon.FromHandle(bmp.GetHicon())
            Catch
            End Try

            ' Use TableLayoutPanel for proper scaling and layout
            Dim tblMain As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(15),
                .ColumnCount = 1,
                .RowCount = 4,
                .AutoSize = False
            }
            tblMain.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            tblMain.RowStyles.Add(New RowStyle(SizeType.AutoSize))  ' Private key row
            tblMain.RowStyles.Add(New RowStyle(SizeType.AutoSize))  ' Add/Clear buttons row
            tblMain.RowStyles.Add(New RowStyle(SizeType.Percent, 100))  ' ListBox
            tblMain.RowStyles.Add(New RowStyle(SizeType.AutoSize))  ' Action buttons
            form.Controls.Add(tblMain)

            ' Row 0: Private key input
            Dim pnlPrivKey As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 3,
                .RowCount = 1,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Margin = New Padding(0, 0, 0, 10)
            }
            pnlPrivKey.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            pnlPrivKey.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            pnlPrivKey.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            tblMain.Controls.Add(pnlPrivKey, 0, 0)

            Dim lblPrivKey As New Label() With {
                .Text = "Private Key:",
                .AutoSize = True,
                .Anchor = AnchorStyles.Left,
                .Padding = New Padding(0, 0, 10, 0)
            }
            pnlPrivKey.Controls.Add(lblPrivKey, 0, 0)

            Dim txtPrivKey As New TextBox() With {
                .Dock = DockStyle.Fill,
                .UseSystemPasswordChar = True
            }
            pnlPrivKey.Controls.Add(txtPrivKey, 1, 0)

            Dim chkShowKey As New CheckBox() With {
                .Text = "Show",
                .AutoSize = True,
                .Anchor = AnchorStyles.Left,
                .Margin = New Padding(10, 0, 0, 0)
            }
            pnlPrivKey.Controls.Add(chkShowKey, 2, 0)

            AddHandler chkShowKey.CheckedChanged, Sub()
                                                      txtPrivKey.UseSystemPasswordChar = Not chkShowKey.Checked
                                                  End Sub

            ' Row 1: Add/Clear buttons
            Dim pnlFileButtons As New FlowLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.LeftToRight,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .WrapContents = False,
                .Margin = New Padding(0, 0, 0, 10)
            }
            tblMain.Controls.Add(pnlFileButtons, 0, 1)

            Dim btnAddFiles As New Button() With {
                .Text = "Add Files...",
                .AutoSize = True,
                .Padding = New Padding(10, 5, 10, 5),
                .Margin = New Padding(0, 0, 10, 0)
            }
            Dim btnClear As New Button() With {
                .Text = "Clear List",
                .AutoSize = True,
                .Padding = New Padding(10, 5, 10, 5)
            }

            pnlFileButtons.Controls.Add(btnAddFiles)
            pnlFileButtons.Controls.Add(btnClear)

            ' Row 2: ListBox
            Dim lbFiles As New ListBox() With {
                .Dock = DockStyle.Fill,
                .SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended,
                .Margin = New Padding(0, 5, 0, 10)
            }
            tblMain.Controls.Add(lbFiles, 0, 2)

            ' Row 3: Action buttons
            Dim pnlButtons As New FlowLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.LeftToRight,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .WrapContents = False,
                .Margin = New Padding(0, 5, 0, 5)
            }
            tblMain.Controls.Add(pnlButtons, 0, 3)

            Dim btnSignAll As New Button() With {
                .Text = "Sign All Files",
                .AutoSize = True,
                .Padding = New Padding(10, 5, 10, 5),
                .Margin = New Padding(0, 0, 10, 0)
            }
            Dim btnClose As New Button() With {
                .Text = "Close",
                .AutoSize = True,
                .Padding = New Padding(10, 5, 10, 5)
            }

            pnlButtons.Controls.Add(btnSignAll)
            pnlButtons.Controls.Add(btnClose)

            AddHandler btnAddFiles.Click, Sub()
                                              Using ofd As New OpenFileDialog()
                                                  ofd.Title = "Select Files to Sign"
                                                  ofd.Filter = "Text Files|*.txt|INI Files|*.ini|All Files|*.*"
                                                  ofd.Multiselect = True
                                                  If ofd.ShowDialog() = DialogResult.OK Then
                                                      For Each f In ofd.FileNames
                                                          If Not lbFiles.Items.Contains(f) Then
                                                              lbFiles.Items.Add(f)
                                                          End If
                                                      Next
                                                  End If
                                              End Using
                                          End Sub

            AddHandler btnClear.Click, Sub() lbFiles.Items.Clear()

            AddHandler btnClose.Click, Sub() form.Close()

            AddHandler btnSignAll.Click, Sub()
                                             If lbFiles.Items.Count = 0 Then
                                                 ShowCustomMessageBox("No files to sign. Add files first.")
                                                 Return
                                             End If

                                             If String.IsNullOrWhiteSpace(txtPrivKey.Text) Then
                                                 ShowCustomMessageBox("Please enter the private key.")
                                                 Return
                                             End If

                                             Dim files = lbFiles.Items.Cast(Of String)().ToArray()
                                             Dim results = BatchSignFiles(files, txtPrivKey.Text.Trim())

                                             Dim sb As New StringBuilder()
                                             sb.AppendLine("Batch Signing Results:")
                                             sb.AppendLine()

                                             Dim successCount = 0
                                             For Each kvp In results
                                                 Dim fileName = Path.GetFileName(kvp.Key)
                                                 If kvp.Value = "Success" Then
                                                     sb.AppendLine($"✓ {fileName}")
                                                     successCount += 1
                                                 Else
                                                     sb.AppendLine($"✗ {fileName}: {kvp.Value}")
                                                 End If
                                             Next

                                             sb.AppendLine()
                                             sb.AppendLine($"Signed: {successCount} / {results.Count}")

                                             ShowCustomMessageBox(sb.ToString())
                                         End Sub

            form.ShowDialog()
        End Sub

#End Region

#Region "Developer Instructions: Creating Signature Files"

        ' =============================================================================
        ' DEVELOPER INSTRUCTIONS: Creating .sig Signature Files
        ' =============================================================================
        '
        ' This system uses Ed25519 digital signatures (via BouncyCastle) to verify
        ' the authenticity of remote INI configuration files. Each update source
        ' file must have a corresponding .sig file containing a Base64-encoded signature.
        '
        ' -----------------------------------------------------------------------------
        ' TECHNICAL SPECIFICATIONS
        ' -----------------------------------------------------------------------------
        '
        ' Algorithm:        Ed25519 (EdDSA with Curve25519)
        ' Key Format:       Raw 32-byte keys, Base64-encoded for storage
        ' Signature Format: Raw 64-byte signature, Base64-encoded in .sig file
        ' Content Encoding: UTF-8 (file content must be read as UTF-8 bytes before signing)
        '
        ' Key Sizes:
        '   - Public Key:  32 bytes raw → ~44 characters Base64
        '   - Private Key: 32 bytes raw → ~44 characters Base64
        '   - Signature:   64 bytes raw → ~88 characters Base64
        '
        ' -----------------------------------------------------------------------------
        ' OPTION 1: Using the Built-in Signature Management Tool
        ' -----------------------------------------------------------------------------
        '
        ' 1. Call SharedMethods.ShowSignatureManagementDialog() from the application
        ' 2. "Generate Keypair" tab → Click "Generate New Keypair"
        ' 3. Save both keys securely:
        '    - Public Key:  Add to UpdateSource parameter in INI files
        '    - Private Key: Store securely (password manager, secure vault)
        ' 4. "Sign File" tab → Select INI file, enter private key, click "Sign File"
        ' 5. Upload BOTH the INI file and its .sig file to the update location
        '
        ' -----------------------------------------------------------------------------
        ' OPTION 2: External Implementation (Python)
        ' -----------------------------------------------------------------------------
        '
        ' Required: pip install cryptography
        '
        ' # Generate keypair
        ' from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey
        ' import base64
        '
        ' private_key = Ed25519PrivateKey.generate()
        ' public_key = private_key.public_key()
        '
        ' private_key_b64 = base64.b64encode(private_key.private_bytes_raw()).decode('ascii')
        ' public_key_b64 = base64.b64encode(public_key.public_bytes_raw()).decode('ascii')
        '
        ' # Sign a file
        ' def sign_file(file_path: str, private_key_b64: str) -> None:
        '     from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey
        '     priv_bytes = base64.b64decode(private_key_b64)
        '     private_key = Ed25519PrivateKey.from_private_bytes(priv_bytes)
        '
        '     with open(file_path, 'rb') as f:
        '         content = f.read()
        '
        '     signature = private_key.sign(content)
        '     sig_b64 = base64.b64encode(signature).decode('ascii')
        '
        '     with open(file_path + '.sig', 'w') as f:
        '         f.write(sig_b64)
        '
        ' -----------------------------------------------------------------------------
        ' OPTION 3: External Implementation (C# / .NET with BouncyCastle)
        ' -----------------------------------------------------------------------------
        '
        ' Required NuGet: Install-Package BouncyCastle.Cryptography
        '
        ' using Org.BouncyCastle.Crypto.Generators;
        ' using Org.BouncyCastle.Crypto.Parameters;
        ' using Org.BouncyCastle.Crypto.Signers;
        ' using Org.BouncyCastle.Security;
        '
        ' // Generate keypair
        ' var keyGen = new Ed25519KeyPairGenerator();
        ' keyGen.Init(new Ed25519KeyGenerationParameters(new SecureRandom()));
        ' var keyPair = keyGen.GenerateKeyPair();
        '
        ' var publicKey = (Ed25519PublicKeyParameters)keyPair.Public;
        ' var privateKey = (Ed25519PrivateKeyParameters)keyPair.Private;
        '
        ' string pubKeyB64 = Convert.ToBase64String(publicKey.GetEncoded());
        ' string privKeyB64 = Convert.ToBase64String(privateKey.GetEncoded());
        '
        ' // Sign a file
        ' void SignFile(string filePath, string privateKeyB64)
        ' {
        '     var privKeyBytes = Convert.FromBase64String(privateKeyB64);
        '     var privateKey = new Ed25519PrivateKeyParameters(privKeyBytes, 0);
        '
        '     var content = File.ReadAllBytes(filePath);
        '
        '     var signer = new Ed25519Signer();
        '     signer.Init(true, privateKey);
        '     signer.BlockUpdate(content, 0, content.Length);
        '     var signature = signer.GenerateSignature();
        '
        '     File.WriteAllText(filePath + ".sig", Convert.ToBase64String(signature));
        ' }
        '
        ' -----------------------------------------------------------------------------
        ' OPTION 4: External Implementation (Node.js)
        ' -----------------------------------------------------------------------------
        '
        ' // Node.js 16+ has built-in Ed25519 support
        ' const crypto = require('crypto');
        ' const fs = require('fs');
        '
        ' // Generate keypair
        ' const { publicKey, privateKey } = crypto.generateKeyPairSync('ed25519');
        '
        ' // Export as raw bytes then Base64
        ' const pubKeyB64 = publicKey.export({ type: 'raw', format: 'der' }).toString('base64');
        ' const privKeyB64 = privateKey.export({ type: 'raw', format: 'der' }).toString('base64');
        '
        ' // Sign a file (requires reconstructing key in PKCS8 format)
        ' function signFile(filePath, privateKeyB64) {
        '     const privKeyBytes = Buffer.from(privateKeyB64, 'base64');
        '     const privateKey = crypto.createPrivateKey({
        '         key: Buffer.concat([
        '             Buffer.from('302e020100300506032b657004220420', 'hex'),
        '             privKeyBytes
        '         ]),
        '         format: 'der',
        '         type: 'pkcs8'
        '     });
        '
        '     const content = fs.readFileSync(filePath);
        '     const signature = crypto.sign(null, content, privateKey);
        '     fs.writeFileSync(filePath + '.sig', signature.toString('base64'));
        ' }
        '
        ' -----------------------------------------------------------------------------
        ' OPTION 5: External Implementation (OpenSSL Command Line)
        ' -----------------------------------------------------------------------------
        '
        ' # Generate keypair
        ' openssl genpkey -algorithm Ed25519 -out private.pem
        ' openssl pkey -in private.pem -pubout -out public.pem
        '
        ' # Extract raw keys as Base64 (for UpdateSource configuration)
        ' openssl pkey -in private.pem -outform DER | tail -c 32 | base64
        ' openssl pkey -in public.pem -pubin -outform DER | tail -c 32 | base64
        '
        ' # Sign a file
        ' openssl pkeyutl -sign -inkey private.pem -in models.ini -out models.ini.sig.raw
        ' base64 models.ini.sig.raw > models.ini.sig
        '
        ' # Verify (for testing)
        ' openssl pkeyutl -verify -pubin -inkey public.pem -in models.ini -sigfile models.ini.sig.raw
        '
        ' -----------------------------------------------------------------------------
        ' UPDATESOURCE CONFIGURATION FORMAT
        ' -----------------------------------------------------------------------------
        '
        ' Add the public key to your INI file's UpdateSource parameter:
        '
        '   [ModelName]
        '   Model = gpt-4
        '   Endpoint = https://api.example.com
        '   UpdateSource = https://updates.example.com/models.ini; all; MCowBQYDK2VwAyEA...
        '
        ' Format: path; keys; public_key
        '
        '   path       - URL or file path to the update INI file
        '   keys       - "all", "new", or comma-separated key names
        '   public_key - Base64-encoded Ed25519 public key (32 bytes → ~44 chars)
        '
        ' -----------------------------------------------------------------------------
        ' DEPLOYMENT CHECKLIST
        ' -----------------------------------------------------------------------------
        '
        ' [ ] Generate Ed25519 keypair
        ' [ ] Store private key securely (NEVER commit to source control)
        ' [ ] Add public key to UpdateSource in distributed INI files
        ' [ ] Create/update the remote INI file
        ' [ ] Sign the INI file to create .sig file
        ' [ ] Upload BOTH files to the same location:
        '       https://example.com/updates/models.ini
        '       https://example.com/updates/models.ini.sig
        ' [ ] Test verification before deployment
        '
        ' -----------------------------------------------------------------------------
        ' SECURITY NOTES
        ' -----------------------------------------------------------------------------
        '
        ' • NEVER share or commit private keys to source control
        ' • Store private keys in a secure vault or password manager
        ' • If a private key is compromised, generate a new keypair and update
        '   all UpdateSource entries with the new public key
        ' • The .sig file must be re-created whenever the INI file content changes
        ' • Signature verification ensures file has not been tampered with
        '
        ' =============================================================================

#End Region

#Region "Agent Resource Package Update"

        ''' </summary>
        ''' <summary>
        ''' Gets the URL of the downloadable Agent Resources ZIP package.
        ''' </summary>
        Private Shared ReadOnly Property AgentResourcesZipUrl As String
            Get
                Return AppsUrl & AppsUrlDir & DotInkyZip
            End Get
        End Property

        ''' <summary>
        ''' ZIP timestamps commonly have a two-second resolution. Differences within this tolerance
        ''' are treated as the same timestamp when the file size also matches.
        ''' </summary>
        Private Const AgentResourceTimestampToleranceSeconds As Double = 2.0R

        Private Const AgentResourceActionUpdateText As String = "Update"
        Private Const AgentResourceActionDeleteText As String = "Delete"
        Private Const AgentResourceActionLeaveText As String = "Leave as is"

        ''' <summary>
        ''' Describes how a package file differs from the corresponding local Agent Resources file.
        ''' </summary>
        Private Enum AgentResourceChangeKind
            NewInPackage
            LocalOnly
            PackageNewer
            LocalNewer
            DifferentSameTimestamp
        End Enum

        ''' <summary>
        ''' Action that may be applied to a local Agent Resources file.
        ''' </summary>
        Private Enum AgentResourceAction
            Update
            Delete
            LeaveAsIs
        End Enum

        ''' <summary>
        ''' Information about a file discovered below either the extracted package root or local root.
        ''' </summary>
        Private Class AgentResourceFileInfo
            Public Property FullPath As String
            Public Property RelativePath As String
            Public Property Length As Long
            Public Property LastWriteTime As DateTime
        End Class

        ''' <summary>
        ''' One detected resource-file difference together with the user's proposed/selected action.
        ''' </summary>
        Private Class AgentResourceFileChange
            Public Property RelativePath As String
            Public Property ChangeKind As AgentResourceChangeKind
            Public Property PackageFile As AgentResourceFileInfo
            Public Property LocalFile As AgentResourceFileInfo
            Public Property IsMarked As Boolean
            Public Property Action As AgentResourceAction
        End Class

        ''' <summary>
        ''' Result of applying selected Agent Resources file operations.
        ''' </summary>
        Private Class AgentResourceApplyResult
            Private ReadOnly _errors As New System.Collections.Generic.List(Of String)()

            Public Property AppliedCount As Integer
            Public Property UpdatedCount As Integer
            Public Property DeletedCount As Integer
            Public Property BackupRoot As String

            Public ReadOnly Property Errors As System.Collections.Generic.List(Of String)
                Get
                    Return _errors
                End Get
            End Property
        End Class

        ''' <summary>
        ''' Separately callable Agent Resources update workflow.
        ''' Downloads the configured ZIP package, extracts it safely to a temporary directory,
        ''' compares the matching resource directory with INI_AgentResourcesPath using file size and
        ''' last-write time, lets the user choose Update/Delete/Leave-as-is per changed file, and then
        ''' applies only the marked operations.
        ''' </summary>
        ''' <param name="context">Shared context containing INI_AgentResourcesPath.</param>
        ''' <param name="zipUrlOverride">
        ''' Optional URL used instead of AgentResourcesZipUrl. Intended mainly for testing.
        ''' </param>
        ''' <returns>True if at least one file operation was successfully applied; otherwise False.</returns>
        Public Shared Function CheckForAgentResourceUpdates(ByRef context As ISharedContext,
                                                            Optional zipUrlOverride As String = Nothing,
                                                            Optional skipAuthorizationGate As Boolean = False) As Boolean
            If context Is Nothing Then
                ShowCustomMessageBox("Agent Resources update cannot run because the shared context is not available.")
                Return False
            End If

            _iniUpdateContext = context

            If Not skipAuthorizationGate AndAlso Not CanRunIniGovernedUpdate(context, "Agent Resources") Then
                Return False
            End If

            Dim tempRoot As String = Nothing

            Try
                If String.IsNullOrWhiteSpace(context.INI_AgentResourcesPath) Then
                    ShowCustomMessageBox("AgentResourcesPath is not configured.")
                    Return False
                End If

                Dim localRoot As String = ExpandEnvironmentVariables(context.INI_AgentResourcesPath)
                If String.IsNullOrWhiteSpace(localRoot) Then
                    ShowCustomMessageBox("AgentResourcesPath could not be resolved.")
                    Return False
                End If

                localRoot = System.IO.Path.GetFullPath(localRoot)

                Dim packageUrl As String = If(String.IsNullOrWhiteSpace(zipUrlOverride), AgentResourcesZipUrl, zipUrlOverride.Trim())
                If String.IsNullOrWhiteSpace(packageUrl) Then
                    ShowCustomMessageBox("The Agent Resources ZIP URL has not been configured. Set AgentResourcesZipUrl first.")
                    Return False
                End If

                Debug.WriteLine("AgentresourcesURL=" & AgentResourcesZipUrl)

                Dim packageUri As System.Uri = Nothing
                If Not System.Uri.TryCreate(packageUrl, System.UriKind.Absolute, packageUri) OrElse
                   Not (packageUri.Scheme.Equals(System.Uri.UriSchemeHttps, System.StringComparison.OrdinalIgnoreCase) OrElse
                        packageUri.Scheme.Equals(System.Uri.UriSchemeHttp, System.StringComparison.OrdinalIgnoreCase)) Then
                    ShowCustomMessageBox("The Agent Resources ZIP URL is invalid. Only HTTP and HTTPS URLs are supported.")
                    Return False
                End If

                LogAgentResourceUpdateEvent("Check Started", $"LocalRoot={localRoot}; Source={packageUrl}")

                tempRoot = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                                                  "AgentResourcesUpdate_" & System.Guid.NewGuid().ToString("N"))
                Dim zipPath As String = System.IO.Path.Combine(tempRoot, "package.zip")
                Dim extractRoot As String = System.IO.Path.Combine(tempRoot, "extracted")

                System.IO.Directory.CreateDirectory(tempRoot)
                System.IO.Directory.CreateDirectory(extractRoot)

                DownloadAgentResourcesZip(packageUrl, zipPath)
                ExtractAgentResourcesZipSafely(zipPath, extractRoot)

                Dim packageRoot As String = ResolveExtractedAgentResourceRoot(extractRoot, localRoot)
                If String.IsNullOrWhiteSpace(packageRoot) Then
                    Dim expectedDirectoryName As String = New System.IO.DirectoryInfo(localRoot).Name
                    Throw New System.InvalidOperationException(
                        $"The ZIP package does not contain a uniquely identifiable directory named '{expectedDirectoryName}'.")
                End If

                Dim changes As System.Collections.Generic.List(Of AgentResourceFileChange) =
                    BuildAgentResourceChanges(packageRoot, localRoot)

                If changes.Count = 0 Then
                    LogAgentResourceUpdateEvent("Check Complete", "No resource-file differences detected")
                    ShowCustomMessageBox("The Agent Resources directory is already up to date.")
                    Return False
                End If

                Dim newCount As Integer = 0
                Dim localOnlyCount As Integer = 0
                Dim packageNewerCount As Integer = 0
                Dim localNewerCount As Integer = 0
                Dim ambiguousCount As Integer = 0

                For Each change As AgentResourceFileChange In changes
                    Select Case change.ChangeKind
                        Case AgentResourceChangeKind.NewInPackage
                            newCount += 1
                        Case AgentResourceChangeKind.LocalOnly
                            localOnlyCount += 1
                        Case AgentResourceChangeKind.PackageNewer
                            packageNewerCount += 1
                        Case AgentResourceChangeKind.LocalNewer
                            localNewerCount += 1
                        Case AgentResourceChangeKind.DifferentSameTimestamp
                            ambiguousCount += 1
                    End Select
                Next

                LogAgentResourceUpdateEvent(
                    "Changes Detected",
                    $"Total={changes.Count}; New={newCount}; PackageNewer={packageNewerCount}; LocalNewer={localNewerCount}; LocalOnly={localOnlyCount}; Ambiguous={ambiguousCount}")

                If Not ShowAgentResourceUpdateDialog(changes, localRoot, packageUrl) Then
                    LogAgentResourceUpdateEvent("User Action", "User cancelled the Agent Resources update")
                    Return False
                End If

                Dim applyResult As AgentResourceApplyResult = ApplyAgentResourceChanges(changes, localRoot)

                Dim summary As New System.Text.StringBuilder()
                summary.AppendLine($"Applied {applyResult.AppliedCount} file operation(s).")
                summary.AppendLine($"Updated/added: {applyResult.UpdatedCount}")
                summary.AppendLine($"Deleted: {applyResult.DeletedCount}")

                If Not String.IsNullOrWhiteSpace(applyResult.BackupRoot) Then
                    summary.AppendLine()
                    summary.AppendLine("Backup of replaced/deleted local files:")
                    summary.AppendLine(applyResult.BackupRoot)
                End If

                If applyResult.Errors.Count > 0 Then
                    summary.AppendLine()
                    summary.AppendLine($"Errors: {applyResult.Errors.Count}")
                    Dim maxErrorsToShow As Integer = System.Math.Min(10, applyResult.Errors.Count)
                    For i As Integer = 0 To maxErrorsToShow - 1
                        summary.AppendLine("- " & applyResult.Errors(i))
                    Next
                    If applyResult.Errors.Count > maxErrorsToShow Then
                        summary.AppendLine($"... and {applyResult.Errors.Count - maxErrorsToShow} more error(s).")
                    End If
                End If

                LogAgentResourceUpdateEvent(
                    "Apply Complete",
                    $"Applied={applyResult.AppliedCount}; Updated={applyResult.UpdatedCount}; Deleted={applyResult.DeletedCount}; Errors={applyResult.Errors.Count}; Backup={If(applyResult.BackupRoot, "")}")

                ShowCustomMessageBox(summary.ToString().TrimEnd())
                Return applyResult.AppliedCount > 0

            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine($"Agent Resources Update Error: {ex.Message}")
                LogAgentResourceUpdateEvent("ERROR", ex.ToString())
                ShowCustomMessageBox("Agent Resources update failed:" & vbCrLf & vbCrLf & ex.Message)
                Return False
            Finally
                If Not String.IsNullOrWhiteSpace(tempRoot) Then
                    Try
                        If System.IO.Directory.Exists(tempRoot) Then
                            System.IO.Directory.Delete(tempRoot, recursive:=True)
                        End If
                    Catch ex As System.Exception
                        System.Diagnostics.Debug.WriteLine($"Failed to remove Agent Resources temporary directory '{tempRoot}': {ex.Message}")
                    End Try
                End If
            End Try
        End Function

        ''' <summary>
        ''' Downloads the Agent Resources ZIP to a local temporary file.
        ''' </summary>
        Private Shared Sub DownloadAgentResourcesZip(packageUrl As String, destinationZipPath As String)
            System.Net.ServicePointManager.SecurityProtocol =
                System.Net.ServicePointManager.SecurityProtocol Or System.Net.SecurityProtocolType.Tls12

            Using client As New System.Net.Http.HttpClient()
                client.Timeout = System.TimeSpan.FromSeconds(60)

                Using response As System.Net.Http.HttpResponseMessage =
                    client.GetAsync(packageUrl, System.Net.Http.HttpCompletionOption.ResponseHeadersRead).GetAwaiter().GetResult()

                    response.EnsureSuccessStatusCode()

                    Using sourceStream As System.IO.Stream = response.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
                        Using targetStream As New System.IO.FileStream(destinationZipPath,
                                                                     System.IO.FileMode.Create,
                                                                     System.IO.FileAccess.Write,
                                                                     System.IO.FileShare.None)
                            sourceStream.CopyTo(targetStream)
                        End Using
                    End Using
                End Using
            End Using

            Dim downloadedFile As New System.IO.FileInfo(destinationZipPath)
            If Not downloadedFile.Exists OrElse downloadedFile.Length = 0 Then
                Throw New System.IO.InvalidDataException("The downloaded Agent Resources ZIP is empty.")
            End If
        End Sub

        ''' <summary>
        ''' Extracts a ZIP archive while preventing entries from escaping the intended temporary root.
        ''' File last-write times are preserved for subsequent comparison.
        ''' </summary>
        Private Shared Sub ExtractAgentResourcesZipSafely(zipPath As String, extractRoot As String)
            Dim normalizedRoot As String = EnsureTrailingDirectorySeparator(System.IO.Path.GetFullPath(extractRoot))

            Using archiveStream As New System.IO.FileStream(zipPath,
                                                            System.IO.FileMode.Open,
                                                            System.IO.FileAccess.Read,
                                                            System.IO.FileShare.Read)
                Using archive As New System.IO.Compression.ZipArchive(archiveStream,
                                                                      System.IO.Compression.ZipArchiveMode.Read,
                                                                      leaveOpen:=False)
                    For Each entry As System.IO.Compression.ZipArchiveEntry In archive.Entries
                        Dim entryPath As String = entry.FullName.Replace("/"c, System.IO.Path.DirectorySeparatorChar)
                        Dim destinationPath As String = System.IO.Path.GetFullPath(System.IO.Path.Combine(extractRoot, entryPath))

                        If Not destinationPath.StartsWith(normalizedRoot, System.StringComparison.OrdinalIgnoreCase) Then
                            Throw New System.IO.InvalidDataException($"Unsafe ZIP entry rejected: {entry.FullName}")
                        End If

                        If String.IsNullOrEmpty(entry.Name) Then
                            System.IO.Directory.CreateDirectory(destinationPath)
                            Continue For
                        End If

                        Dim destinationDirectory As String = System.IO.Path.GetDirectoryName(destinationPath)
                        If Not String.IsNullOrWhiteSpace(destinationDirectory) Then
                            System.IO.Directory.CreateDirectory(destinationDirectory)
                        End If

                        Using sourceStream As System.IO.Stream = entry.Open()
                            Using targetStream As New System.IO.FileStream(destinationPath,
                                                                         System.IO.FileMode.Create,
                                                                         System.IO.FileAccess.Write,
                                                                         System.IO.FileShare.None)
                                sourceStream.CopyTo(targetStream)
                            End Using
                        End Using

                        Try
                            System.IO.File.SetLastWriteTime(destinationPath, entry.LastWriteTime.LocalDateTime)
                        Catch ex As System.Exception
                            System.Diagnostics.Debug.WriteLine($"Could not preserve ZIP timestamp for '{entry.FullName}': {ex.Message}")
                        End Try
                    Next
                End Using
            End Using
        End Sub

        ''' <summary>
        ''' Finds the extracted directory whose name matches the final directory component of
        ''' INI_AgentResourcesPath. If multiple matches exist, the unique shallowest match is used.
        ''' </summary>
        Private Shared Function ResolveExtractedAgentResourceRoot(extractRoot As String, localRoot As String) As String
            Dim expectedName As String = New System.IO.DirectoryInfo(localRoot).Name
            If String.IsNullOrWhiteSpace(expectedName) Then Return Nothing

            Dim candidates As New System.Collections.Generic.List(Of String)()
            For Each directoryPath As String In System.IO.Directory.GetDirectories(extractRoot,
                                                                                   "*",
                                                                                   System.IO.SearchOption.AllDirectories)
                Dim directoryName As String = New System.IO.DirectoryInfo(directoryPath).Name
                If directoryName.Equals(expectedName, System.StringComparison.OrdinalIgnoreCase) Then
                    candidates.Add(directoryPath)
                End If
            Next

            If candidates.Count = 0 Then Return Nothing
            If candidates.Count = 1 Then Return candidates(0)

            Dim minimumDepth As Integer = System.Int32.MaxValue
            For Each candidate As String In candidates
                Dim relative As String = GetRelativeResourcePath(extractRoot, candidate)
                Dim depth As Integer = relative.Split({System.IO.Path.DirectorySeparatorChar,
                                                       System.IO.Path.AltDirectorySeparatorChar},
                                                      System.StringSplitOptions.RemoveEmptyEntries).Length
                If depth < minimumDepth Then minimumDepth = depth
            Next

            Dim shallowMatches As New System.Collections.Generic.List(Of String)()
            For Each candidate As String In candidates
                Dim relative As String = GetRelativeResourcePath(extractRoot, candidate)
                Dim depth As Integer = relative.Split({System.IO.Path.DirectorySeparatorChar,
                                                       System.IO.Path.AltDirectorySeparatorChar},
                                                      System.StringSplitOptions.RemoveEmptyEntries).Length
                If depth = minimumDepth Then shallowMatches.Add(candidate)
            Next

            If shallowMatches.Count = 1 Then Return shallowMatches(0)
            Return Nothing
        End Function

        ''' <summary>
        ''' Builds the list of changed files. Unchanged files (same size and timestamp within ZIP
        ''' timestamp tolerance) are intentionally omitted from the approval dialog.
        ''' </summary>
        Private Shared Function BuildAgentResourceChanges(packageRoot As String,
                                                          localRoot As String) As System.Collections.Generic.List(Of AgentResourceFileChange)
            Dim packageFiles As System.Collections.Generic.Dictionary(Of String, AgentResourceFileInfo) =
                EnumerateAgentResourceFiles(packageRoot)

            Dim localFiles As New System.Collections.Generic.Dictionary(Of String, AgentResourceFileInfo)(System.StringComparer.OrdinalIgnoreCase)
            If System.IO.Directory.Exists(localRoot) Then
                localFiles = EnumerateAgentResourceFiles(localRoot)
            End If

            Dim allPaths As New System.Collections.Generic.SortedSet(Of String)(System.StringComparer.OrdinalIgnoreCase)
            For Each relativePath As String In packageFiles.Keys
                allPaths.Add(relativePath)
            Next
            For Each relativePath As String In localFiles.Keys
                allPaths.Add(relativePath)
            Next

            Dim result As New System.Collections.Generic.List(Of AgentResourceFileChange)()

            For Each relativePath As String In allPaths
                Dim packageFile As AgentResourceFileInfo = Nothing
                Dim localFile As AgentResourceFileInfo = Nothing
                Dim inPackage As Boolean = packageFiles.TryGetValue(relativePath, packageFile)
                Dim inLocal As Boolean = localFiles.TryGetValue(relativePath, localFile)

                If inPackage AndAlso Not inLocal Then
                    result.Add(New AgentResourceFileChange() With {
                        .RelativePath = relativePath,
                        .ChangeKind = AgentResourceChangeKind.NewInPackage,
                        .PackageFile = packageFile,
                        .LocalFile = Nothing,
                        .IsMarked = True,
                        .Action = AgentResourceAction.Update
                    })
                    Continue For
                End If

                If inLocal AndAlso Not inPackage Then
                    result.Add(New AgentResourceFileChange() With {
                        .RelativePath = relativePath,
                        .ChangeKind = AgentResourceChangeKind.LocalOnly,
                        .PackageFile = Nothing,
                        .LocalFile = localFile,
                        .IsMarked = False,
                        .Action = AgentResourceAction.LeaveAsIs
                    })
                    Continue For
                End If

                If packageFile Is Nothing OrElse localFile Is Nothing Then Continue For

                Dim timeDifferenceSeconds As Double =
                    System.Math.Abs((packageFile.LastWriteTime - localFile.LastWriteTime).TotalSeconds)
                Dim sameTimestamp As Boolean = timeDifferenceSeconds <= AgentResourceTimestampToleranceSeconds
                Dim sameSize As Boolean = packageFile.Length = localFile.Length

                If sameTimestamp AndAlso sameSize Then
                    Continue For
                End If

                Dim changeKind As AgentResourceChangeKind
                Dim defaultMarked As Boolean = False
                Dim defaultAction As AgentResourceAction = AgentResourceAction.LeaveAsIs

                If packageFile.LastWriteTime > localFile.LastWriteTime.AddSeconds(AgentResourceTimestampToleranceSeconds) Then
                    changeKind = AgentResourceChangeKind.PackageNewer
                    defaultMarked = True
                    defaultAction = AgentResourceAction.Update
                ElseIf localFile.LastWriteTime > packageFile.LastWriteTime.AddSeconds(AgentResourceTimestampToleranceSeconds) Then
                    changeKind = AgentResourceChangeKind.LocalNewer
                Else
                    changeKind = AgentResourceChangeKind.DifferentSameTimestamp
                End If

                result.Add(New AgentResourceFileChange() With {
                    .RelativePath = relativePath,
                    .ChangeKind = changeKind,
                    .PackageFile = packageFile,
                    .LocalFile = localFile,
                    .IsMarked = defaultMarked,
                    .Action = defaultAction
                })
            Next

            Return result
        End Function

        ''' <summary>
        ''' Enumerates all files below a resource root, keyed by case-insensitive relative path.
        ''' </summary>
        Private Shared Function EnumerateAgentResourceFiles(rootPath As String) As System.Collections.Generic.Dictionary(Of String, AgentResourceFileInfo)
            Dim result As New System.Collections.Generic.Dictionary(Of String, AgentResourceFileInfo)(System.StringComparer.OrdinalIgnoreCase)
            If Not System.IO.Directory.Exists(rootPath) Then Return result

            For Each filePath As String In System.IO.Directory.GetFiles(rootPath, "*", System.IO.SearchOption.AllDirectories)
                Dim relativePath As String = GetRelativeResourcePath(rootPath, filePath)
                Dim info As New System.IO.FileInfo(filePath)

                result(relativePath) = New AgentResourceFileInfo() With {
                    .FullPath = filePath,
                    .RelativePath = relativePath,
                    .Length = info.Length,
                    .LastWriteTime = info.LastWriteTime
                }
            Next

            Return result
        End Function

        ''' <summary>
        ''' Displays the interactive resource-file approval dialog.
        ''' </summary>
        Private Shared Function ShowAgentResourceUpdateDialog(changes As System.Collections.Generic.List(Of AgentResourceFileChange),
                                                              localRoot As String,
                                                              packageUrl As String) As Boolean
            Dim approved As Boolean = False

            Dim form As New System.Windows.Forms.Form() With {
                .Text = $"{AN} - Agent Resources Updates Available",
                .Size = New System.Drawing.Size(1180, 680),
                .StartPosition = System.Windows.Forms.FormStartPosition.CenterParent,
                .FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable,
                .Font = New System.Drawing.Font("Segoe UI", 9.0F),
                .MinimumSize = New System.Drawing.Size(950, 500),
                .AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi,
                .TopMost = True
            }

            Try
                Dim bmp As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                form.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine($"Could not set Agent Resources dialog icon: {ex.Message}")
            End Try

            Dim tblMain As New System.Windows.Forms.TableLayoutPanel() With {
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .Padding = New System.Windows.Forms.Padding(15),
                .ColumnCount = 1,
                .RowCount = 4,
                .AutoSize = False
            }
            tblMain.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
            tblMain.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            tblMain.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            tblMain.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
            tblMain.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            form.Controls.Add(tblMain)

            Dim lblHeader As New System.Windows.Forms.Label() With {
                .Text = "The ZIP package differs from the local Agent Resources directory. " &
                        "Mark the file operations you want to perform and choose an action for each file." & vbCrLf &
                        "Local-only files are normally user-generated and are therefore left unchanged by default.",
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .AutoSize = True,
                .Padding = New System.Windows.Forms.Padding(0, 0, 0, 6)
            }
            tblMain.Controls.Add(lblHeader, 0, 0)

            Dim lblPaths As New System.Windows.Forms.Label() With {
                .Text = "Local: " & localRoot & vbCrLf & "Package: " & packageUrl,
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .AutoSize = True,
                .Font = New System.Drawing.Font("Segoe UI", 8.25F),
                .ForeColor = System.Drawing.SystemColors.GrayText,
                .Padding = New System.Windows.Forms.Padding(0, 0, 0, 8)
            }
            tblMain.Controls.Add(lblPaths, 0, 1)

            Dim dgv As New System.Windows.Forms.DataGridView() With {
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .AllowUserToAddRows = False,
                .AllowUserToDeleteRows = False,
                .AllowUserToResizeRows = False,
                .SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect,
                .MultiSelect = False,
                .RowHeadersVisible = False,
                .Margin = New System.Windows.Forms.Padding(0, 5, 0, 10),
                .ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize,
                .EnableHeadersVisualStyles = True,
                .AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.None
            }

            Dim colMark As New System.Windows.Forms.DataGridViewCheckBoxColumn() With {
                .HeaderText = "Apply",
                .Name = "colMark",
                .Width = 55,
                .MinimumWidth = 55
            }
            Dim colStatus As New System.Windows.Forms.DataGridViewTextBoxColumn() With {
                .HeaderText = "Difference",
                .Name = "colStatus",
                .ReadOnly = True,
                .Width = 145,
                .MinimumWidth = 120
            }
            Dim colFile As New System.Windows.Forms.DataGridViewTextBoxColumn() With {
                .HeaderText = "File",
                .Name = "colFile",
                .ReadOnly = True,
                .AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill,
                .MinimumWidth = 240
            }
            Dim colAction As New System.Windows.Forms.DataGridViewComboBoxColumn() With {
                .HeaderText = "Action",
                .Name = "colAction",
                .Width = 120,
                .MinimumWidth = 105,
                .FlatStyle = System.Windows.Forms.FlatStyle.Flat
            }
            colAction.Items.AddRange(AgentResourceActionUpdateText, AgentResourceActionDeleteText, AgentResourceActionLeaveText)

            Dim colLocalDate As New System.Windows.Forms.DataGridViewTextBoxColumn() With {
                .HeaderText = "Local date",
                .Name = "colLocalDate",
                .ReadOnly = True,
                .Width = 145
            }
            Dim colPackageDate As New System.Windows.Forms.DataGridViewTextBoxColumn() With {
                .HeaderText = "ZIP date",
                .Name = "colPackageDate",
                .ReadOnly = True,
                .Width = 145
            }
            Dim colLocalSize As New System.Windows.Forms.DataGridViewTextBoxColumn() With {
                .HeaderText = "Local size",
                .Name = "colLocalSize",
                .ReadOnly = True,
                .Width = 95
            }
            Dim colPackageSize As New System.Windows.Forms.DataGridViewTextBoxColumn() With {
                .HeaderText = "ZIP size",
                .Name = "colPackageSize",
                .ReadOnly = True,
                .Width = 95
            }

            dgv.Columns.AddRange(colMark, colStatus, colFile, colAction,
                                 colLocalDate, colPackageDate, colLocalSize, colPackageSize)

            For Each change As AgentResourceFileChange In changes
                Dim rowIndex As Integer = dgv.Rows.Add(
                    change.IsMarked,
                    GetAgentResourceChangeDisplayText(change.ChangeKind),
                    change.RelativePath,
                    GetAgentResourceActionDisplayText(change.Action),
                    FormatAgentResourceFileDate(change.LocalFile),
                    FormatAgentResourceFileDate(change.PackageFile),
                    FormatAgentResourceFileSize(change.LocalFile),
                    FormatAgentResourceFileSize(change.PackageFile))

                Dim actionCell As New System.Windows.Forms.DataGridViewComboBoxCell() With {
                    .FlatStyle = System.Windows.Forms.FlatStyle.Flat
                }

                Select Case change.ChangeKind
                    Case AgentResourceChangeKind.NewInPackage
                        actionCell.Items.Add(AgentResourceActionUpdateText)
                        actionCell.Items.Add(AgentResourceActionLeaveText)
                    Case AgentResourceChangeKind.LocalOnly
                        actionCell.Items.Add(AgentResourceActionDeleteText)
                        actionCell.Items.Add(AgentResourceActionLeaveText)
                    Case Else
                        actionCell.Items.Add(AgentResourceActionUpdateText)
                        actionCell.Items.Add(AgentResourceActionDeleteText)
                        actionCell.Items.Add(AgentResourceActionLeaveText)
                End Select

                actionCell.Value = GetAgentResourceActionDisplayText(change.Action)
                dgv.Rows(rowIndex).Cells("colAction") = actionCell

                If change.ChangeKind = AgentResourceChangeKind.LocalOnly OrElse
                   change.ChangeKind = AgentResourceChangeKind.LocalNewer OrElse
                   change.ChangeKind = AgentResourceChangeKind.DifferentSameTimestamp Then
                    dgv.Rows(rowIndex).DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 250, 230)
                End If
            Next

            AddHandler dgv.CurrentCellDirtyStateChanged,
                Sub()
                    If dgv.IsCurrentCellDirty Then
                        dgv.CommitEdit(System.Windows.Forms.DataGridViewDataErrorContexts.Commit)
                    End If
                End Sub

            AddHandler dgv.CellValueChanged,
                Sub(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs)
                    If e.RowIndex < 0 OrElse e.RowIndex >= changes.Count OrElse e.ColumnIndex < 0 Then Return

                    If dgv.Columns(e.ColumnIndex).Name = "colMark" Then
                        Dim rawValue As Object = dgv.Rows(e.RowIndex).Cells("colMark").Value
                        changes(e.RowIndex).IsMarked = rawValue IsNot Nothing AndAlso System.Convert.ToBoolean(rawValue)
                    ElseIf dgv.Columns(e.ColumnIndex).Name = "colAction" Then
                        Dim rawValue As Object = dgv.Rows(e.RowIndex).Cells("colAction").Value
                        changes(e.RowIndex).Action = ParseAgentResourceAction(If(rawValue, AgentResourceActionLeaveText).ToString())
                    End If
                End Sub

            AddHandler dgv.CellFormatting,
                Sub(sender As Object, e As System.Windows.Forms.DataGridViewCellFormattingEventArgs)
                    If e.RowIndex < 0 OrElse e.RowIndex >= dgv.Rows.Count Then Return
                    Dim actionValue As Object = dgv.Rows(e.RowIndex).Cells("colAction").Value
                    If actionValue IsNot Nothing AndAlso actionValue.ToString().Equals(AgentResourceActionDeleteText, System.StringComparison.OrdinalIgnoreCase) Then
                        e.CellStyle.ForeColor = System.Drawing.Color.DarkRed
                        e.CellStyle.SelectionForeColor = System.Drawing.Color.DarkRed
                    End If
                End Sub

            dgv.AutoResizeColumnHeadersHeight()
            tblMain.Controls.Add(dgv, 0, 2)

            Dim pnlButtons As New System.Windows.Forms.FlowLayoutPanel() With {
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
                .AutoSize = True,
                .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                .WrapContents = True,
                .Margin = New System.Windows.Forms.Padding(0, 5, 0, 5)
            }
            tblMain.Controls.Add(pnlButtons, 0, 3)

            Dim btnApply As New System.Windows.Forms.Button() With {
                .Text = "Apply Selected",
                .AutoSize = True,
                .Padding = New System.Windows.Forms.Padding(10, 5, 10, 5),
                .Margin = New System.Windows.Forms.Padding(0, 0, 10, 0)
            }
            Dim btnSelectRecommended As New System.Windows.Forms.Button() With {
                .Text = "Select Recommended",
                .AutoSize = True,
                .Padding = New System.Windows.Forms.Padding(10, 5, 10, 5),
                .Margin = New System.Windows.Forms.Padding(0, 0, 10, 0)
            }
            Dim btnClear As New System.Windows.Forms.Button() With {
                .Text = "Clear Selection",
                .AutoSize = True,
                .Padding = New System.Windows.Forms.Padding(10, 5, 10, 5),
                .Margin = New System.Windows.Forms.Padding(0, 0, 10, 0)
            }
            Dim btnCancel As New System.Windows.Forms.Button() With {
                .Text = "Cancel",
                .AutoSize = True,
                .Padding = New System.Windows.Forms.Padding(10, 5, 10, 5)
            }

            pnlButtons.Controls.Add(btnApply)
            pnlButtons.Controls.Add(btnSelectRecommended)
            pnlButtons.Controls.Add(btnClear)
            pnlButtons.Controls.Add(btnCancel)

            AddHandler btnSelectRecommended.Click,
                Sub()
                    For i As Integer = 0 To changes.Count - 1
                        Dim change As AgentResourceFileChange = changes(i)
                        Dim recommended As Boolean =
                            change.ChangeKind = AgentResourceChangeKind.NewInPackage OrElse
                            change.ChangeKind = AgentResourceChangeKind.PackageNewer

                        change.IsMarked = recommended
                        change.Action = If(recommended, AgentResourceAction.Update, AgentResourceAction.LeaveAsIs)
                        dgv.Rows(i).Cells("colMark").Value = change.IsMarked
                        dgv.Rows(i).Cells("colAction").Value = GetAgentResourceActionDisplayText(change.Action)
                    Next
                End Sub

            AddHandler btnClear.Click,
                Sub()
                    For i As Integer = 0 To changes.Count - 1
                        changes(i).IsMarked = False
                        dgv.Rows(i).Cells("colMark").Value = False
                    Next
                End Sub

            AddHandler btnCancel.Click, Sub() form.Close()

            AddHandler btnApply.Click,
                Sub()
                    dgv.EndEdit()
                    SyncAgentResourceDialogSelections(dgv, changes)

                    Dim actionableCount As Integer = 0
                    Dim deleteCount As Integer = 0
                    For Each change As AgentResourceFileChange In changes
                        If change.IsMarked AndAlso change.Action <> AgentResourceAction.LeaveAsIs Then
                            actionableCount += 1
                            If change.Action = AgentResourceAction.Delete Then deleteCount += 1
                        End If
                    Next

                    If actionableCount = 0 Then
                        ShowCustomMessageBox("No file operations are marked for application.")
                        Return
                    End If

                    If deleteCount > 0 Then
                        Dim answer As Integer = ShowCustomYesNoBox(
                            $"You selected {deleteCount} file(s) for deletion. Existing files will be backed up before deletion. Continue?",
                            "Yes, apply changes",
                            "No")
                        If answer <> 1 Then Return
                    End If

                    approved = True
                    form.Close()
                End Sub

            ' Surface any same-process native host prompt (e.g. Word's "Save changes?")
            ' that would otherwise stay hidden behind this TopMost modal dialog.
            AttachForeignForegroundWatchdog(form)

            form.ShowDialog()
            Return approved
        End Function

        ''' <summary>
        ''' Synchronizes grid values into the change objects before the dialog closes.
        ''' </summary>
        Private Shared Sub SyncAgentResourceDialogSelections(dgv As System.Windows.Forms.DataGridView,
                                                             changes As System.Collections.Generic.List(Of AgentResourceFileChange))
            For i As Integer = 0 To changes.Count - 1
                Dim markValue As Object = dgv.Rows(i).Cells("colMark").Value
                changes(i).IsMarked = markValue IsNot Nothing AndAlso System.Convert.ToBoolean(markValue)

                Dim actionValue As Object = dgv.Rows(i).Cells("colAction").Value
                changes(i).Action = ParseAgentResourceAction(If(actionValue, AgentResourceActionLeaveText).ToString())
            Next
        End Sub

        ''' <summary>
        ''' Applies all marked file operations and creates a single timestamped sibling backup
        ''' directory for any existing local files that are replaced or deleted.
        ''' </summary>
        Private Shared Function ApplyAgentResourceChanges(changes As System.Collections.Generic.List(Of AgentResourceFileChange),
                                                          localRoot As String) As AgentResourceApplyResult
            Dim result As New AgentResourceApplyResult()

            For Each change As AgentResourceFileChange In changes
                If Not change.IsMarked OrElse change.Action = AgentResourceAction.LeaveAsIs Then Continue For

                Try
                    Dim localPath As String = GetSafeResourceDestinationPath(localRoot, change.RelativePath)

                    Select Case change.Action
                        Case AgentResourceAction.Update
                            If change.PackageFile Is Nothing OrElse Not System.IO.File.Exists(change.PackageFile.FullPath) Then
                                Throw New System.IO.FileNotFoundException("The ZIP source file is not available.", change.RelativePath)
                            End If

                            If System.IO.File.Exists(localPath) Then
                                BackupAgentResourceFile(localRoot, localPath, result)
                            End If

                            Dim localDirectory As String = System.IO.Path.GetDirectoryName(localPath)
                            If Not String.IsNullOrWhiteSpace(localDirectory) Then
                                System.IO.Directory.CreateDirectory(localDirectory)
                            End If

                            System.IO.File.Copy(change.PackageFile.FullPath, localPath, overwrite:=True)
                            System.IO.File.SetLastWriteTime(localPath, change.PackageFile.LastWriteTime)

                            result.AppliedCount += 1
                            result.UpdatedCount += 1
                            LogAgentResourceUpdateEvent("File Updated", change.RelativePath)

                        Case AgentResourceAction.Delete
                            If System.IO.File.Exists(localPath) Then
                                BackupAgentResourceFile(localRoot, localPath, result)
                                System.IO.File.Delete(localPath)
                                result.AppliedCount += 1
                                result.DeletedCount += 1
                                LogAgentResourceUpdateEvent("File Deleted", change.RelativePath)
                            End If
                    End Select

                Catch ex As System.Exception
                    Dim errorText As String = $"{change.RelativePath}: {ex.Message}"
                    result.Errors.Add(errorText)
                    LogAgentResourceUpdateEvent("File Error", errorText)
                End Try
            Next

            Return result
        End Function

        ''' <summary>
        ''' Backs up one existing local file into the run's timestamped backup directory.
        ''' </summary>
        Private Shared Sub BackupAgentResourceFile(localRoot As String,
                                                   localFilePath As String,
                                                   result As AgentResourceApplyResult)
            If String.IsNullOrWhiteSpace(result.BackupRoot) Then
                Dim trimmedLocalRoot As String = localRoot.TrimEnd(System.IO.Path.DirectorySeparatorChar,
                                                                  System.IO.Path.AltDirectorySeparatorChar)
                Dim parentDirectory As String = System.IO.Path.GetDirectoryName(trimmedLocalRoot)
                Dim localDirectoryName As String = New System.IO.DirectoryInfo(trimmedLocalRoot).Name

                If String.IsNullOrWhiteSpace(parentDirectory) Then
                    parentDirectory = System.IO.Path.GetTempPath()
                End If

                result.BackupRoot = System.IO.Path.Combine(
                    parentDirectory,
                    localDirectoryName & "_backup_" & System.DateTime.Now.ToString("yyyyMMdd_HHmmss_fff"))

                System.IO.Directory.CreateDirectory(result.BackupRoot)
                LogAgentResourceUpdateEvent("Backup Created", result.BackupRoot)
            End If

            Dim relativePath As String = GetRelativeResourcePath(localRoot, localFilePath)
            Dim backupPath As String = GetSafeResourceDestinationPath(result.BackupRoot, relativePath)
            Dim backupDirectory As String = System.IO.Path.GetDirectoryName(backupPath)

            If Not String.IsNullOrWhiteSpace(backupDirectory) Then
                System.IO.Directory.CreateDirectory(backupDirectory)
            End If

            System.IO.File.Copy(localFilePath, backupPath, overwrite:=False)
            System.IO.File.SetLastWriteTime(backupPath, System.IO.File.GetLastWriteTime(localFilePath))
        End Sub

        ''' <summary>
        ''' Returns a path below rootPath and rejects any relative path that would escape the root.
        ''' </summary>
        Private Shared Function GetSafeResourceDestinationPath(rootPath As String, relativePath As String) As String
            Dim normalizedRoot As String = EnsureTrailingDirectorySeparator(System.IO.Path.GetFullPath(rootPath))
            Dim fullPath As String = System.IO.Path.GetFullPath(System.IO.Path.Combine(rootPath, relativePath))

            If Not fullPath.StartsWith(normalizedRoot, System.StringComparison.OrdinalIgnoreCase) Then
                Throw New System.IO.InvalidDataException($"Unsafe resource path rejected: {relativePath}")
            End If

            Return fullPath
        End Function

        ''' <summary>
        ''' Returns the normalized path of fullPath relative to rootPath.
        ''' </summary>
        Private Shared Function GetRelativeResourcePath(rootPath As String, fullPath As String) As String
            Dim normalizedRoot As String = EnsureTrailingDirectorySeparator(System.IO.Path.GetFullPath(rootPath))
            Dim normalizedFullPath As String = System.IO.Path.GetFullPath(fullPath)

            If normalizedFullPath.Equals(normalizedRoot.TrimEnd(System.IO.Path.DirectorySeparatorChar,
                                                                 System.IO.Path.AltDirectorySeparatorChar),
                                         System.StringComparison.OrdinalIgnoreCase) Then
                Return String.Empty
            End If

            If Not normalizedFullPath.StartsWith(normalizedRoot, System.StringComparison.OrdinalIgnoreCase) Then
                Throw New System.ArgumentException("The supplied path is outside the specified resource root.", NameOf(fullPath))
            End If

            Return normalizedFullPath.Substring(normalizedRoot.Length).
                Replace(System.IO.Path.AltDirectorySeparatorChar, System.IO.Path.DirectorySeparatorChar)
        End Function

        ''' <summary>
        ''' Ensures a directory path ends with the platform directory separator.
        ''' </summary>
        Private Shared Function EnsureTrailingDirectorySeparator(pathValue As String) As String
            If String.IsNullOrEmpty(pathValue) Then Return pathValue
            If pathValue.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString(), System.StringComparison.Ordinal) OrElse
               pathValue.EndsWith(System.IO.Path.AltDirectorySeparatorChar.ToString(), System.StringComparison.Ordinal) Then
                Return pathValue
            End If
            Return pathValue & System.IO.Path.DirectorySeparatorChar
        End Function

        Private Shared Function GetAgentResourceChangeDisplayText(changeKind As AgentResourceChangeKind) As String
            Select Case changeKind
                Case AgentResourceChangeKind.NewInPackage
                    Return "New in ZIP"
                Case AgentResourceChangeKind.LocalOnly
                    Return "Local only"
                Case AgentResourceChangeKind.PackageNewer
                    Return "ZIP newer"
                Case AgentResourceChangeKind.LocalNewer
                    Return "Local newer"
                Case AgentResourceChangeKind.DifferentSameTimestamp
                    Return "Size differs / same date"
                Case Else
                    Return changeKind.ToString()
            End Select
        End Function

        Private Shared Function GetAgentResourceActionDisplayText(action As AgentResourceAction) As String
            Select Case action
                Case AgentResourceAction.Update
                    Return AgentResourceActionUpdateText
                Case AgentResourceAction.Delete
                    Return AgentResourceActionDeleteText
                Case Else
                    Return AgentResourceActionLeaveText
            End Select
        End Function

        Private Shared Function ParseAgentResourceAction(value As String) As AgentResourceAction
            If value.Equals(AgentResourceActionUpdateText, System.StringComparison.OrdinalIgnoreCase) Then
                Return AgentResourceAction.Update
            End If
            If value.Equals(AgentResourceActionDeleteText, System.StringComparison.OrdinalIgnoreCase) Then
                Return AgentResourceAction.Delete
            End If
            Return AgentResourceAction.LeaveAsIs
        End Function

        Private Shared Function FormatAgentResourceFileDate(fileInfo As AgentResourceFileInfo) As String
            If fileInfo Is Nothing Then Return "—"
            Return fileInfo.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
        End Function

        Private Shared Function FormatAgentResourceFileSize(fileInfo As AgentResourceFileInfo) As String
            If fileInfo Is Nothing Then Return "—"
            Return fileInfo.Length.ToString("N0", System.Globalization.CultureInfo.CurrentCulture)
        End Function

        ''' <summary>
        ''' Writes Agent Resources update activity to the same update log used by the INI updater,
        ''' but with a separate event prefix.
        ''' </summary>
        Private Shared Sub LogAgentResourceUpdateEvent(eventType As String, details As String)
            Try
                Dim message As String = $"[Agent Resources Update] [{eventType}]"
                If Not String.IsNullOrWhiteSpace(details) Then
                    message &= vbCrLf & "  " & details.Replace(vbCrLf, vbCrLf & "  ")
                End If
                UpdateHandler.WriteUpdateLog(message)
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine($"Failed to log Agent Resources update event: {ex.Message}")
            End Try
        End Sub

#End Region


    End Class
End Namespace
