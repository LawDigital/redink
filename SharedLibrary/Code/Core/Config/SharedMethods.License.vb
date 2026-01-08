' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.License.vb
' Purpose: Implements runtime license validation and related user messaging.
'
' Architecture:
'  - Configuration Sources:
'     - `configDict` (central configuration) can provide `LicensedTill`, `LicenseStatus`,
'       `LicenseUsers`, `LicenseContact`, and `LicenseNoWarning`.
'     - `My.Settings` persists per-user values for `LicensedTill`, `LicenseStatus`,
'       `LicenseUsers` and warning counters/dates.
'  - License Modes:
'     - GA (General Audience): Uses `LicensedTill` (config or `My.Settings`) and prompts
'       for license entry when missing or expired (subject to grace period).
'     - Beta: Uses `BetaEndDate` only; license storage via `My.Settings` is not offered.
'  - Warning Throttling:
'     - Expiry warnings are shown at specific day intervals and throttled using
'       `My.Settings.LicenseWarningStartCount`.
'     - Grace period warnings are throttled using `My.Settings.GracePeriodWarningStartcount`.
'     - Beta warnings are throttled using `My.Settings.LastBetaWarningDate` and
'       `My.Settings.BetaWarningStartCount`.
'  - External Dependencies:
'     - UI helpers: `ShowCustomMessageBox`, `ShowCustomYesNoBox`
'     - License types and metadata: `GetLicenseTypes`, `LicenseTypeInfo`
'     - Configuration helpers/constants: `ParseBoolean`, `AppsUrl`, `NewHomeURL`,
'       `AN`, `AN4`, `BetaEndDate`, `BetaUpgradeInstructions`, thresholds/interval constants.
' =============================================================================


Option Strict On
Option Explicit On

Imports System.ComponentModel
Imports System.Net
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary
    Partial Public Class SharedMethods


        ''' <summary>
        ''' Main license check entry point.
        ''' Loads settings, determines GA vs Beta behavior, and returns <c>True</c> if usage is allowed.
        ''' </summary>
        ''' <param name="context">Execution context (used for version display and license type lookup).</param>
        ''' <param name="configDict">Configuration dictionary (typically centrally managed).</param>
        ''' <returns>
        ''' <c>True</c> if the license check allows continuation; otherwise <c>False</c>.
        ''' </returns>
        Public Shared Function LicenseOK(ByVal context As ISharedContext,
                                          ByVal configDict As Dictionary(Of String, String)) As Boolean
            Try
                ' Load license settings from config file and My.Settings.
                LoadLicenseSettings(configDict, context)

                ' Skip all checks if the config explicitly disabled them.
                If LicenseCheckDisabled Then
                    Return True
                End If

                ' Determine if this is GA version or Beta based on AppsUrl prefix.
                Dim isGAVersion As Boolean = AppsUrl.StartsWith(NewHomeURL, StringComparison.OrdinalIgnoreCase)

                If isGAVersion Then
                    Return PerformGALicenseCheck(context)
                Else
                    Return PerformBetaLicenseCheck(context)
                End If

            Catch ex As Exception
                ' Fault tolerance: on unexpected errors, log a message and allow continuation.
                ' (This is a deliberate "fail-open" behavior.)
                Try
                    ShowCustomMessageBox($"License check encountered an error: {ex.Message}. Continuing with limited functionality.", AN)
                Catch
                End Try
                Return True
            End Try
        End Function

        ''' <summary>
        ''' Loads license-related settings from <paramref name="configDict"/> and/or <c>My.Settings</c>.
        ''' Also sets flags such as <c>LicenseFromConfig</c> and <c>LicenseCheckDisabled</c>.
        ''' </summary>
        ''' <param name="configDict">Configuration dictionary containing license keys.</param>
        ''' <param name="context">
        ''' Optional execution context. When provided, enables checking whether the stored license type
        ''' implies a later fixed end date and updates <c>My.Settings.LicensedTill</c> accordingly.
        ''' </param>
        Private Shared Sub LoadLicenseSettings(configDict As Dictionary(Of String, String), Optional context As ISharedContext = Nothing)
            Try
                ' Reset the config flag.
                LicenseFromConfig = False

                ' Load LicenseContact from config.
                LicenseContact = If(configDict.ContainsKey("LicenseContact"), configDict("LicenseContact"), "")

                ' Load LicenseNoWarning from config.
                LicenseNoWarning = ParseBoolean(configDict, "LicenseNoWarning", False)

                ' Determine if this is a Beta version (AppsUrl not pointing to NewHomeURL).
                Dim isBetaVersion As Boolean = Not AppsUrl.StartsWith(NewHomeURL, StringComparison.OrdinalIgnoreCase)

                ' Check if LicensedTill is "False"/"No" or more than LicenseCheckDisabledYears in the future.
                If configDict.ContainsKey("LicensedTill") Then
                    Dim configValue = configDict("LicensedTill").Trim()

                    ' Check for explicit disable values.
                    If configValue.Equals("False", StringComparison.OrdinalIgnoreCase) OrElse
               configValue.Equals("No", StringComparison.OrdinalIgnoreCase) Then
                        LicenseCheckDisabled = True
                        LicenseStatus = If(configDict.ContainsKey("LicenseStatus"), configDict("LicenseStatus"), "")
                        If configDict.ContainsKey("LicenseUsers") Then Integer.TryParse(configDict("LicenseUsers"), LicenseUsers)
                        Return
                    End If

                    ' Try to parse the date from config.
                    Dim configDate As Date
                    If Date.TryParse(configValue, configDate) Then
                        ' Disable checking if the config value is beyond LicenseCheckDisabledYears years.
                        If configDate > Date.Now.AddYears(LicenseCheckDisabledYears) Then
                            LicenseCheckDisabled = True
                            Return
                        End If

                        ' Config file value takes precedence over everything.
                        LicensedTill = configDate
                        LicenseFromConfig = True ' Mark that license came from config.
                    Else
                        ' Could not parse date - fall through to defaults.
                        LicensedTill = If(isBetaVersion, BetaEndDate, Date.MinValue)
                    End If
                Else
                    ' No LicensedTill in config file.
                    If isBetaVersion Then
                        ' Beta version: always use BetaEndDate.
                        LicensedTill = BetaEndDate
                        LicenseUsers = 1
                        LicenseStatus = "Beta Test License"
                    Else
                        ' GA version: use My.Settings if available, otherwise Date.MinValue (triggers entry form).
                        Try
                            If My.Settings.LicensedTill > Date.MinValue AndAlso
                                My.Settings.LicensedTill < Date.MaxValue Then
                                LicensedTill = My.Settings.LicensedTill

                                ' If the stored license type implies a later fixed end date, extend silently.
                                If context IsNot Nothing Then
                                    TryExtendLicenseToFixedEndDate(context)
                                End If
                            Else
                                ' No valid license date - set to MinValue to trigger entry form.
                                LicensedTill = Date.MinValue
                            End If
                        Catch
                            LicensedTill = Date.MinValue
                        End Try
                    End If
                End If

                ' Load LicenseStatus - config takes precedence.
                If configDict.ContainsKey("LicenseStatus") Then
                    LicenseStatus = configDict("LicenseStatus")

                    ' If config has LicenseStatus, also get LicenseUsers from config.
                    If configDict.ContainsKey("LicenseUsers") Then
                        Integer.TryParse(configDict("LicenseUsers"), LicenseUsers)
                    End If
                ElseIf Not isBetaVersion Then
                    ' GA version without config status: use My.Settings.
                    Try
                        LicenseStatus = If(String.IsNullOrEmpty(My.Settings.LicenseStatus), "", My.Settings.LicenseStatus)
                        LicenseUsers = If(My.Settings.LicenseUsers > 0, My.Settings.LicenseUsers, 1)
                    Catch
                        LicenseStatus = ""
                        LicenseUsers = 1
                    End Try
                End If

            Catch ex As Exception
                ' Fault tolerance: determine reasonable defaults based on version.
                LicenseCheckDisabled = False
                LicenseStatus = ""
                LicenseFromConfig = False
                Dim isBeta = Not AppsUrl.StartsWith("https://redink.ai", StringComparison.OrdinalIgnoreCase)
                LicensedTill = If(isBeta, BetaEndDate, Date.MinValue)
            End Try
        End Sub

        ''' <summary>
        ''' If the current stored license type implies a later fixed end date than the currently stored
        ''' <c>LicensedTill</c>, extends <c>LicensedTill</c> and persists the updated date to <c>My.Settings</c>.
        ''' </summary>
        ''' <param name="context">Execution context used to resolve license types based on version date.</param>
        Private Shared Sub TryExtendLicenseToFixedEndDate(context As ISharedContext)
            Try
                ' Only proceed if we have a LicenseStatus from My.Settings.
                Dim storedStatus As String = ""
                Try
                    storedStatus = My.Settings.LicenseStatus
                Catch
                    Return
                End Try

                If String.IsNullOrEmpty(storedStatus) Then Return

                ' Get the version date from context.RDV.
                Dim versionDate As Date = ParseVersionDateFromRDV(context.RDV)

                ' Get the license types using the version date.
                Dim licenseTypes = GetLicenseTypes(versionDate)

                ' Find the matching license type by name.
                Dim matchingType As LicenseTypeInfo = Nothing
                For Each lt In licenseTypes
                    If lt.Name.Equals(storedStatus, StringComparison.OrdinalIgnoreCase) Then
                        matchingType = lt
                        Exit For
                    End If
                Next

                If matchingType Is Nothing Then Return

                ' Only extend if the license type does NOT allow user-defined end dates.
                If matchingType.UserDefinedEndDate Then Return

                ' Require a fixed end date to extend.
                If Not matchingType.FixedEndDate.HasValue Then Return

                Dim fixedDate As Date = matchingType.FixedEndDate.Value

                ' Extend only if the fixed date is beyond the current LicensedTill.
                If fixedDate > LicensedTill Then
                    LicensedTill = fixedDate

                    ' Persist the updated date to My.Settings.
                    Try
                        My.Settings.LicensedTill = LicensedTill
                        My.Settings.Save()
                    Catch
                        ' Ignore save errors; updated value remains in memory.
                    End Try
                End If

            Catch
                ' Fault tolerance: ignore any errors in the extension check.
            End Try
        End Sub

        ''' <summary>
        ''' Performs the GA (General Audience) license check flow.
        ''' Handles config-based licenses, user-stored licenses, expiry/grace period behavior,
        ''' and optional prompting to update missing/expired licenses.
        ''' </summary>
        ''' <param name="context">Execution context used for UI messages and license entry.</param>
        ''' <returns><c>True</c> if the license check allows continuation; otherwise <c>False</c>.</returns>
        Private Shared Function PerformGALicenseCheck(context As ISharedContext) As Boolean
            Dim licenseExpired As Boolean = False
            Dim noLicenseConfigured As Boolean = False

            ' If license came from config file with a valid future date, accept it without further checks.
            If LicenseFromConfig Then
                If Date.Now > LicensedTill Then
                    ' Check if we're within the grace period.
                    If GracePeriodDays > 0 AndAlso Date.Now <= LicensedTill.AddDays(GracePeriodDays) Then
                        ' Within grace period: show warning and allow continuation.
                        CheckGracePeriodWarning(context, LicensedTill, False)
                        Return True
                    End If

                    Dim msg = BuildLicenseMessage(
        $"Your license for {AN} for {context.RDV} has EXPIRED on {LicensedTill:d}." & vbCrLf & vbCrLf &
        "Please contact your administrator to update the license configuration.")
                    ShowCustomMessageBox(msg, $"{AN} License Expired")
                    Return False
                End If

                If Not LicenseNoWarning AndAlso LicensedTill > Date.Now AndAlso LicensedTill < Date.MaxValue Then
                    CheckLicenseExpiryWarnings(context, False)
                End If
                Return True
            End If

            ' Non-config license path: check My.Settings-based license.
            If LicensedTill = Date.MinValue OrElse LicensedTill = Date.MaxValue Then
                noLicenseConfigured = True
            ElseIf Date.Now > LicensedTill Then
                ' Check grace period.
                If GracePeriodDays > 0 AndAlso Date.Now <= LicensedTill.AddDays(GracePeriodDays) Then
                    CheckGracePeriodWarning(context, LicensedTill, False)
                    Return True
                End If
                licenseExpired = True
            End If

            ' Handle expired license (past grace period).
            If licenseExpired Then
                Dim msg = BuildLicenseMessage(
    $"Your license for {AN} for {context.RDV} has EXPIRED on {LicensedTill:d}." & vbCrLf & vbCrLf &
    "Would you like to update your license information now?")

                Dim result = ShowCustomYesNoBox(msg, "Update License", "Cancel", $"{AN} License")
                If result = 1 Then
                    If Not ShowLicenseEntryForm(context) Then
                        Return False
                    End If

                    ' Re-validate after form: check if new license is valid.
                    If LicensedTill = Date.MinValue OrElse LicensedTill = Date.MaxValue OrElse Date.Now > LicensedTill Then
                        Return False
                    End If
                Else
                    Return False
                End If
            End If

            ' Handle no license configured.
            If noLicenseConfigured Then
                Dim msg = BuildLicenseMessage(
    $"No valid license is configured for {AN} for {context.RDV}." & vbCrLf & vbCrLf &
    "Would you like to enter your license information now?")

                Dim result = ShowCustomYesNoBox(msg, "Enter License", "Cancel", $"{AN} License")
                If result = 1 Then
                    If Not ShowLicenseEntryForm(context) Then
                        Return False
                    End If

                    ' Re-validate after form: check if new license is valid.
                    If LicensedTill = Date.MinValue OrElse LicensedTill = Date.MaxValue OrElse Date.Now > LicensedTill Then
                        Return False
                    End If
                Else
                    Return False
                End If
            End If

            ' Check for upcoming expiry warnings.
            If Not LicenseNoWarning AndAlso LicensedTill > Date.Now AndAlso LicensedTill < Date.MaxValue Then
                CheckLicenseExpiryWarnings(context, False)
            End If

            Return True
        End Function

        ''' <summary>
        ''' Shows a grace period warning (when enabled) during the grace period after expiry.
        ''' Warning frequency is throttled via <c>ShouldShowGracePeriodWarning</c>.
        ''' </summary>
        ''' <param name="context">Execution context used for UI messages.</param>
        ''' <param name="expiredDate">The original expiry date that started the grace period.</param>
        ''' <param name="isBeta"><c>True</c> for beta messaging; otherwise GA messaging.</param>
        Private Shared Sub CheckGracePeriodWarning(context As ISharedContext, expiredDate As Date, isBeta As Boolean)
            Try
                ' Skip warning if LicenseNoWarning is True.
                If LicenseNoWarning Then Return

                ' Calculate remaining grace period days.
                Dim gracePeriodEnd As Date = expiredDate.AddDays(GracePeriodDays)
                Dim remainingDays As Integer = CInt((gracePeriodEnd.Date - Date.Now.Date).TotalDays)

                ' Check if we should show the warning based on start count.
                If ShouldShowGracePeriodWarning() Then
                    Dim msg As String
                    If isBeta Then
                        msg = BuildLicenseMessage(
                            $"Your beta test license for {AN} for {context.RDV} EXPIRED on {expiredDate:d}." & vbCrLf & vbCrLf &
                            $"You are currently in a {GracePeriodDays}-day grace period. " &
                            $"The add-in will stop working in {remainingDays} day(s) on {gracePeriodEnd:d}." & vbCrLf & vbCrLf &
                            $"To continue using {AN}, please upgrade to the General Audience or Preview version. " &
                            $"Visit {AN4} for upgrade instructions.")
                        ShowCustomMessageBox(msg, $"{AN} Beta Test Grace Period")
                    Else
                        msg = BuildLicenseMessage(
                            $"Your license for {AN} for {context.RDV} EXPIRED on {expiredDate:d}." & vbCrLf & vbCrLf &
                            $"You are currently in a {GracePeriodDays}-day grace period. " &
                            $"The add-in will stop working in {remainingDays} day(s) on {gracePeriodEnd:d}." & vbCrLf & vbCrLf &
                            If(LicenseFromConfig,
                               "Please contact your administrator to update the license configuration.",
                               "Would you like to update your license information now?"))

                        If LicenseFromConfig Then
                            ShowCustomMessageBox(msg, $"{AN} License Grace Period")
                        Else
                            Dim result = ShowCustomYesNoBox(msg, "Update License", "Later", $"{AN} License Grace Period")
                            If result = 1 Then
                                ShowLicenseEntryForm(context)
                            End If
                        End If
                    End If

                    RecordGracePeriodWarningShown()
                End If

            Catch
                ' Fault tolerance: ignore warning check errors.
            End Try
        End Sub


        ''' <summary>
        ''' Returns whether a grace period warning should be shown on this start.
        ''' Uses and updates <c>My.Settings.GracePeriodWarningStartcount</c>.
        ''' </summary>
        ''' <returns><c>True</c> if the warning should be shown; otherwise <c>False</c>.</returns>
        Private Shared Function ShouldShowGracePeriodWarning() As Boolean
            Try
                ' Increment start count.
                Dim startCount As Integer = 0
                Try
                    startCount = My.Settings.GracePeriodWarningStartcount + 1
                Catch
                    startCount = 1
                End Try

                My.Settings.GracePeriodWarningStartcount = startCount
                My.Settings.Save()

                ' Show warning if start count reaches the interval threshold.
                Return startCount >= GracePeriodWarningIntervals

            Catch
                Return True
            End Try
        End Function

        ''' <summary>
        ''' Records that a grace period warning was shown by resetting the counter in <c>My.Settings</c>.
        ''' </summary>
        Private Shared Sub RecordGracePeriodWarningShown()
            Try
                My.Settings.GracePeriodWarningStartcount = 0
                My.Settings.Save()
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' Performs the Beta license check flow based on <c>BetaEndDate</c>.
        ''' Shows pre-expiry and expiry messaging, supports a grace period, and optionally opens upgrade instructions.
        ''' </summary>
        ''' <param name="context">Execution context used for UI messages.</param>
        ''' <returns><c>True</c> if the beta period (or its grace period) allows continuation; otherwise <c>False</c>.</returns>
        Private Shared Function PerformBetaLicenseCheck(context As ISharedContext) As Boolean
            ' Track if we showed a warning to avoid double-warning from CheckLicenseExpiryWarnings.
            Dim warningAlreadyShown As Boolean = False

            ' Check if BetaUpgradeInstructions URL is available.
            Dim upgradeAvailable = CheckUrlAvailable(BetaUpgradeInstructions)

            Debug.WriteLine("upgradeAvailable = " & upgradeAvailable)

            ' Calculate days until beta end.
            Dim daysUntilBetaEnd As Integer = CInt((BetaEndDate.Date - Date.Now.Date).TotalDays)

            ' Show pre-expiry warning only if:
            ' - upgradeAvailable = True, OR
            ' - BetaWarningNoUpgradeDays or fewer days remain until BetaEndDate.
            If Not Date.Now > BetaEndDate Then
                Dim shouldShowBetaPreWarning As Boolean = upgradeAvailable OrElse daysUntilBetaEnd <= BetaWarningNoUpgradeDays

                If shouldShowBetaPreWarning AndAlso ShouldShowBetaWarning() Then
                    Dim msg As String
                    If upgradeAvailable Then
                        msg = BuildLicenseMessage(
                            $"The beta test for {AN} ends on {BetaEndDate:d}." & vbCrLf & vbCrLf &
                            $"To continue using {AN}, upgrade now to the General Audience or Preview version. " &
                            $"Visit {AN4} for upgrade instructions." & vbCrLf & vbCrLf &
                            "Would you like to open the upgrade instructions page?")

                        Dim result = ShowCustomYesNoBox(msg, "Open Instructions", "Later", $"{AN} Beta Test")
                        If result = 1 Then
                            Try
                                Process.Start(New ProcessStartInfo(BetaUpgradeInstructions) With {.UseShellExecute = True})
                            Catch ex As Exception
                                ShowCustomMessageBox($"Could not open the upgrade instructions page: {ex.Message}", AN)
                            End Try
                        End If
                    Else
                        ' No upgrade available but within threshold days - just inform.
                        msg = BuildLicenseMessage(
                            $"The beta test for {AN} ends in {daysUntilBetaEnd} day(s) on {BetaEndDate:d}." & vbCrLf & vbCrLf &
                            $"To continue using {AN} after that date, you will need to upgrade to the General Audience or Preview version. " &
                            $"Visit {AN4} for more information.")
                        ShowCustomMessageBox(msg, $"{AN} Beta Test")
                    End If

                    RecordBetaWarningShown()
                    warningAlreadyShown = True
                End If
            End If

            ' Check if beta has expired.
            If Date.Now > BetaEndDate Then
                ' Check if we're within the grace period.
                If GracePeriodDays > 0 AndAlso Date.Now <= BetaEndDate.AddDays(GracePeriodDays) Then
                    ' Within grace period: show warning and allow continuation.
                    CheckGracePeriodWarning(context, BetaEndDate, True)
                    Return True
                End If

                ' Past grace period: beta expired.
                If upgradeAvailable Then
                    Dim msg = BuildLicenseMessage(
                        $"Your beta test license for {AN} for {context.RDV} has EXPIRED on {BetaEndDate:d}." & vbCrLf & vbCrLf &
                        $"To continue using {AN}, upgrade to the General Audience or Preview version. " &
                        $"Visit {AN4} for upgrade instructions." & vbCrLf & vbCrLf &
                        "Would you like to open the upgrade instructions page?")

                    Dim result = ShowCustomYesNoBox(msg, "Open Instructions", "Cancel", $"{AN} Beta Test Expired")
                    If result = 1 Then
                        Try
                            Process.Start(New ProcessStartInfo(BetaUpgradeInstructions) With {.UseShellExecute = True})
                        Catch ex As Exception
                            ShowCustomMessageBox($"Could not open the upgrade instructions page: {ex.Message}", AN)
                        End Try
                    End If
                Else
                    Dim msg = BuildLicenseMessage(
                        $"Your beta test license for {AN} for {context.RDV} has EXPIRED on {BetaEndDate:d}." & vbCrLf & vbCrLf &
                        $"To continue using {AN}, upgrade to the General Audience or Preview version at {AN4}.")
                    ShowCustomMessageBox(msg, $"{AN} Beta Test Expired")
                End If

                Return False
            End If

            ' Upcoming expiry warnings for beta - only if no warning was already shown.
            If Not warningAlreadyShown AndAlso Not LicenseNoWarning Then
                ' Only show if upgrade available OR within BetaWarningNoUpgradeDays days.
                Dim shouldShowExpiryWarning As Boolean = upgradeAvailable OrElse daysUntilBetaEnd <= BetaWarningNoUpgradeDays
                If shouldShowExpiryWarning Then
                    CheckLicenseExpiryWarnings(context, True)
                End If
            End If

            Return True
        End Function

        ''' <summary>
        ''' Shows license expiry warnings at configured day thresholds prior to expiry.
        ''' Warning frequency is throttled using <c>My.Settings.LicenseWarningStartCount</c>.
        ''' </summary>
        ''' <param name="context">Execution context used for UI messages.</param>
        ''' <param name="isBeta"><c>True</c> for beta expiry messaging; otherwise GA messaging.</param>
        Private Shared Sub CheckLicenseExpiryWarnings(context As ISharedContext, isBeta As Boolean)
            Try
                Dim expiryDate As Date = If(isBeta, BetaEndDate, LicensedTill)
                Dim daysUntilExpiry = CInt((expiryDate.Date - Date.Now.Date).TotalDays)

                For Each warningDay In LicenseWarningDays
                    If daysUntilExpiry = warningDay Then
                        ' Check if we should show the warning based on start count.
                        If Not ShouldShowLicenseWarning() Then
                            Exit For
                        End If

                        Dim msg As String
                        If isBeta Then
                            msg = BuildLicenseMessage(
                                $"Your beta test license for {AN} for {context.RDV} will EXPIRE in {daysUntilExpiry} day(s) " &
                                $"on {expiryDate:d}." & vbCrLf & vbCrLf &
                                $"To continue using {AN}, please upgrade to the General Audience or Preview version. " &
                                $"Visit {AN4} for upgrade instructions.")
                            ShowCustomMessageBox(msg, $"{AN} Beta Test Warning")
                        Else
                            msg = BuildLicenseMessage(
                                $"Your license for {AN} for {context.RDV} will EXPIRE in {daysUntilExpiry} day(s) " &
                                $"on {expiryDate:d}." & vbCrLf & vbCrLf &
                                If(LicenseFromConfig,
                                   "Your license is configured centrally. Contact your administrator to renew.",
                                   $"Please update your license at {AN4} or contact your administrator. Updating the license information is possible via 'Settings', then 'About {AN}'." & vbCrLf & vbCrLf & "Would you like to update your license information now?"))

                            ' Only offer to update if license is NOT from config.
                            If LicenseFromConfig Then
                                ShowCustomMessageBox(msg, $"{AN} License Warning")
                            Else
                                Dim result = ShowCustomYesNoBox(msg, "Update License", "Later", $"{AN} License Warning")
                                If result = 1 Then
                                    ShowLicenseEntryForm(context)
                                End If
                            End If
                        End If

                        RecordLicenseWarningShown()
                        Exit For
                    End If
                Next
            Catch
                ' Fault tolerance: ignore warning check errors.
            End Try
        End Sub

        ''' <summary>
        ''' Returns whether an expiry warning should be shown on this start.
        ''' Uses and updates <c>My.Settings.LicenseWarningStartCount</c>.
        ''' </summary>
        ''' <returns><c>True</c> if the warning should be shown; otherwise <c>False</c>.</returns>
        Private Shared Function ShouldShowLicenseWarning() As Boolean
            Try
                ' Increment start count.
                Dim startCount As Integer = 0
                Try
                    startCount = My.Settings.LicenseWarningStartCount + 1
                Catch
                    startCount = 1
                End Try

                My.Settings.LicenseWarningStartCount = startCount
                My.Settings.Save()

                ' Show warning if start count reaches the interval threshold.
                Return startCount >= LicenseWarningInterval

            Catch
                Return True
            End Try
        End Function

        ''' <summary>
        ''' Records that a license warning was shown by resetting the counter in <c>My.Settings</c>.
        ''' </summary>
        Private Shared Sub RecordLicenseWarningShown()
            Try
                My.Settings.LicenseWarningStartCount = 0
                My.Settings.Save()
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' Shows the license entry form for selecting a license type and persisting selection to <c>My.Settings</c>.
        ''' Returns <c>False</c> immediately for beta versions (license storage is disabled in beta).
        ''' </summary>
        ''' <param name="context">Execution context used to resolve license types and version date.</param>
        ''' <returns><c>True</c> if the user saved a license successfully; otherwise <c>False</c>.</returns>
        Public Shared Function ShowLicenseEntryForm(context As ISharedContext) As Boolean
            Try
                ' Beta version cannot store license information.
                Dim isBetaVersion As Boolean = Not AppsUrl.StartsWith(NewHomeURL, StringComparison.OrdinalIgnoreCase)
                If isBetaVersion Then
                    ShowCustomMessageBox(
                        $"License configuration is not available during the beta test period." & vbCrLf & vbCrLf &
                        $"To use {AN} after the beta ends on {BetaEndDate:d}, please upgrade to the General Audience or Preview version. " &
                        $"Visit {AN4} for more information.",
                        $"{AN} Beta Test")
                    Return False
                End If

                Dim versionDate = ParseVersionDateFromRDV(context.RDV)
                Dim licenseTypes = GetLicenseTypes(versionDate)

                Using form As New Form()
                    ' Enable DPI awareness.
                    form.AutoScaleMode = AutoScaleMode.Dpi
                    form.AutoScaleDimensions = New System.Drawing.SizeF(96.0F, 96.0F)

                    form.Text = $"{AN} License Configuration"
                    form.FormBorderStyle = FormBorderStyle.FixedDialog
                    form.StartPosition = FormStartPosition.CenterScreen
                    form.MaximizeBox = False
                    form.MinimizeBox = False
                    form.ShowInTaskbar = True
                    form.TopMost = True

                    ' Enable auto-sizing for the form.
                    form.AutoSize = True
                    form.AutoSizeMode = AutoSizeMode.GrowAndShrink

                    ' Set icon.
                    Try
                        Dim bmp As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                        form.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())
                    Catch
                    End Try

                    Dim font As New System.Drawing.Font("Segoe UI", 9.0F)
                    form.Font = font

                    ' Use TableLayoutPanel for DPI-aware layout.
                    Dim mainLayout As New TableLayoutPanel() With {
                        .AutoSize = True,
                        .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                        .ColumnCount = 2,
                        .RowCount = 6,
                        .Padding = New Padding(15, 15, 15, 20)
                    }
                    mainLayout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
                    mainLayout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))

                    ' Row 0: Title label (spans both columns).
                    Dim titleLabel As New Label() With {
                        .Text = "Select your license type:",
                        .AutoSize = True,
                        .Font = New System.Drawing.Font("Segoe UI", 9.0F),
                        .Margin = New Padding(0, 0, 0, 5)
                    }
                    mainLayout.Controls.Add(titleLabel, 0, 0)
                    mainLayout.SetColumnSpan(titleLabel, 2)

                    ' Row 1: License type ComboBox (spans both columns).
                    Dim cboLicenseType As New ComboBox() With {
                        .Width = 450,
                        .DropDownStyle = ComboBoxStyle.DropDownList,
                        .Margin = New Padding(0, 0, 0, 10)
                    }
                    For Each lt In licenseTypes
                        cboLicenseType.Items.Add(lt.Name)
                    Next
                    mainLayout.Controls.Add(cboLicenseType, 0, 1)
                    mainLayout.SetColumnSpan(cboLicenseType, 2)

                    ' Calculate max description height needed.
                    Dim maxDescHeight = 0
                    Using g = form.CreateGraphics()
                        For Each lt In licenseTypes
                            Dim size = g.MeasureString(lt.Description, font, 450)
                            maxDescHeight = Math.Max(maxDescHeight, CInt(Math.Ceiling(size.Height)))
                        Next
                    End Using

                    ' Row 2: Description label (spans both columns).
                    Dim lblDescription As New Label() With {
                        .Width = 450,
                        .Height = maxDescHeight + 10,
                        .ForeColor = System.Drawing.Color.FromArgb(96, 96, 96),
                        .BackColor = System.Drawing.SystemColors.Control,
                        .Margin = New Padding(0, 0, 0, 10)
                    }
                    mainLayout.Controls.Add(lblDescription, 0, 2)
                    mainLayout.SetColumnSpan(lblDescription, 2)

                    ' Row 3: License end date.
                    Dim lblEndDate As New Label() With {
                        .Text = "License valid until:",
                        .AutoSize = True,
                        .Anchor = AnchorStyles.Left,
                        .Margin = New Padding(0, 5, 15, 12)
                    }
                    mainLayout.Controls.Add(lblEndDate, 0, 3)

                    Dim dtpEndDate As New DateTimePicker() With {
                        .Format = DateTimePickerFormat.Short,
                        .Width = 150,
                        .Anchor = AnchorStyles.Left,
                        .Margin = New Padding(0, 0, 0, 12)
                    }
                    mainLayout.Controls.Add(dtpEndDate, 1, 3)

                    ' Row 4: Number of users.
                    Dim lblUsers As New Label() With {
                        .Text = "Number of users:",
                        .AutoSize = True,
                        .Anchor = AnchorStyles.Left,
                        .Margin = New Padding(0, 5, 15, 20)
                    }
                    mainLayout.Controls.Add(lblUsers, 0, 4)

                    Dim nudUsers As New NumericUpDown() With {
                        .Minimum = 1,
                        .Maximum = 10000,
                        .Value = 1,
                        .Width = 80,
                        .Anchor = AnchorStyles.Left,
                        .Margin = New Padding(0, 0, 0, 20)
                    }
                    mainLayout.Controls.Add(nudUsers, 1, 4)

                    ' Row 5: Buttons panel (spans both columns).
                    Dim buttonPanel As New FlowLayoutPanel() With {
                        .AutoSize = True,
                        .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                        .FlowDirection = FlowDirection.LeftToRight,
                        .Margin = New Padding(0, 5, 0, 0)
                    }

                    Dim btnSave As New Button() With {
                        .Text = "Save License",
                        .AutoSize = True,
                        .Padding = New Padding(10, 5, 10, 5),
                        .Margin = New Padding(0, 0, 10, 0)
                    }
                    buttonPanel.Controls.Add(btnSave)

                    Dim btnCancel As New Button() With {
                        .Text = "Cancel",
                        .AutoSize = True,
                        .Padding = New Padding(10, 5, 10, 5),
                        .Margin = New Padding(0, 0, 0, 0)
                    }
                    buttonPanel.Controls.Add(btnCancel)

                    mainLayout.Controls.Add(buttonPanel, 0, 5)
                    mainLayout.SetColumnSpan(buttonPanel, 2)

                    form.Controls.Add(mainLayout)

                    Dim result As Boolean = False

                    ' Update UI based on license type selection.
                    AddHandler cboLicenseType.SelectedIndexChanged, Sub(s, e)
                                                                        If cboLicenseType.SelectedIndex >= 0 AndAlso cboLicenseType.SelectedIndex < licenseTypes.Count Then
                                                                            Dim selectedType = licenseTypes(cboLicenseType.SelectedIndex)
                                                                            lblDescription.Text = selectedType.Description

                                                                            ' End date.
                                                                            If selectedType.FixedEndDate.HasValue Then
                                                                                dtpEndDate.Value = selectedType.FixedEndDate.Value
                                                                                dtpEndDate.Enabled = False
                                                                            ElseIf selectedType.DefaultEndDate.HasValue Then
                                                                                dtpEndDate.Value = selectedType.DefaultEndDate.Value
                                                                                dtpEndDate.Enabled = True
                                                                            ElseIf selectedType.UserDefinedEndDate Then
                                                                                ' User-defined but no default: suggest end of current calendar year.
                                                                                dtpEndDate.Value = New Date(Date.Now.Year, 12, 31)
                                                                                dtpEndDate.Enabled = True
                                                                            Else
                                                                                dtpEndDate.Value = Date.Now.AddYears(1)
                                                                                dtpEndDate.Enabled = True
                                                                            End If

                                                                            ' Users.
                                                                            If selectedType.FixedUsers.HasValue Then
                                                                                nudUsers.Value = selectedType.FixedUsers.Value
                                                                                nudUsers.Enabled = False
                                                                            ElseIf selectedType.DefaultUsers.HasValue Then
                                                                                nudUsers.Value = selectedType.DefaultUsers.Value
                                                                                nudUsers.Enabled = True
                                                                            Else
                                                                                nudUsers.Enabled = True
                                                                            End If
                                                                        End If
                                                                    End Sub

                    ' Save handler.
                    AddHandler btnSave.Click, Sub(s, e)
                                                  Try
                                                      ' Validation.
                                                      If cboLicenseType.SelectedIndex < 0 Then
                                                          ShowCustomMessageBox("Please select a license type.", AN)
                                                          Return
                                                      End If

                                                      Dim endDate = dtpEndDate.Value.Date
                                                      Dim users = CInt(nudUsers.Value)

                                                      If users <= 0 Then
                                                          ShowCustomMessageBox("Number of users must be at least 1.", AN)
                                                          Return
                                                      End If

                                                      If endDate <= Date.Now.Date Then
                                                          ShowCustomMessageBox("License end date must be in the future.", AN)
                                                          Return
                                                      End If

                                                      If endDate > Date.Now.AddYears(MaxLicenseYearsInFuture) Then
                                                          ShowCustomMessageBox($"License end date cannot be more than {MaxLicenseYearsInFuture} years in the future.", AN)
                                                          Return
                                                      End If

                                                      ' Save to My.Settings.
                                                      LicenseStatus = licenseTypes(cboLicenseType.SelectedIndex).Name
                                                      LicenseUsers = users
                                                      LicensedTill = endDate

                                                      My.Settings.LicenseStatus = LicenseStatus
                                                      My.Settings.LicenseUsers = LicenseUsers
                                                      My.Settings.LicensedTill = LicensedTill
                                                      My.Settings.Save()

                                                      result = True
                                                      form.Close()
                                                  Catch ex As Exception
                                                      ShowCustomMessageBox($"Error saving license: {ex.Message}", AN)
                                                  End Try
                                              End Sub

                    AddHandler btnCancel.Click, Sub(s, e)
                                                    form.Close()
                                                End Sub

                    ' Set initial selection if we have a stored license.
                    If Not String.IsNullOrEmpty(LicenseStatus) Then
                        For i = 0 To licenseTypes.Count - 1
                            If licenseTypes(i).Name.Equals(LicenseStatus, StringComparison.OrdinalIgnoreCase) Then
                                cboLicenseType.SelectedIndex = i
                                If LicensedTill > Date.MinValue AndAlso LicensedTill < Date.MaxValue Then
                                    Try : dtpEndDate.Value = LicensedTill : Catch : End Try
                                End If
                                nudUsers.Value = Math.Max(1, LicenseUsers)
                                Exit For
                            End If
                        Next
                    End If

                    form.ShowDialog()
                    Return result
                End Using

            Catch ex As Exception
                ShowCustomMessageBox($"Error showing license form: {ex.Message}", AN)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Returns <c>True</c> if this build is treated as a beta version (license storage disabled).
        ''' </summary>
        Public Shared Function IsBetaVersion() As Boolean
            Return Not AppsUrl.StartsWith(NewHomeURL, StringComparison.OrdinalIgnoreCase)
        End Function

        ''' <summary>
        ''' Parses a version date from an RDV string by extracting 6 digits after "V." in DDMMYY format.
        ''' Returns <see cref="Date.Now"/> if parsing fails.
        ''' </summary>
        ''' <param name="rdv">A string that may contain a version marker such as "V.101225".</param>
        ''' <returns>The parsed version date, or <see cref="Date.Now"/> on failure.</returns>
        Private Shared Function ParseVersionDateFromRDV(rdv As String) As Date
            Try
                ' Find "V." followed by 6 digits (DDMMYY format).
                Dim vIndex = rdv.IndexOf("V.", StringComparison.OrdinalIgnoreCase)
                If vIndex >= 0 AndAlso rdv.Length >= vIndex + 8 Then
                    Dim dateStr = rdv.Substring(vIndex + 2, 6)
                    If dateStr.All(AddressOf Char.IsDigit) Then
                        Dim day = Integer.Parse(dateStr.Substring(0, 2))
                        Dim month = Integer.Parse(dateStr.Substring(2, 2))
                        Dim year = 2000 + Integer.Parse(dateStr.Substring(4, 2))
                        Return New Date(year, month, day)
                    End If
                End If
            Catch
            End Try

            ' Default to current date if parsing fails.
            Return Date.Now
        End Function

        ''' <summary>
        ''' Appends <c>LicenseContact</c> (if present) to the given message.
        ''' </summary>
        ''' <param name="baseMessage">The message to display without contact details.</param>
        ''' <returns>The message including optional contact information.</returns>
        Private Shared Function BuildLicenseMessage(baseMessage As String) As String
            If Not String.IsNullOrEmpty(LicenseContact) Then
                Return baseMessage & vbCrLf & vbCrLf & LicenseContact
            End If
            Return baseMessage
        End Function


        ''' <summary>
        ''' Checks whether the given URL is reachable.
        ''' Issues a HEAD request first (following redirects), and retries with GET if HEAD fails.
        ''' Only 2xx responses are treated as available.
        ''' </summary>
        ''' <param name="url">The URL to probe.</param>
        ''' <returns><c>True</c> if a 2xx response is obtained; otherwise <c>False</c>.</returns>
        Private Shared Function CheckUrlAvailable(url As String) As Boolean
            Try
                ' Enable TLS 1.2 and TLS 1.3 for HTTPS connections.
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls13

                Dim request = DirectCast(WebRequest.Create(url), HttpWebRequest)
                request.Method = "HEAD"
                request.Timeout = 5000
                request.AllowAutoRedirect = True
                request.MaximumAutomaticRedirections = 5
                request.UserAgent = $"{AN}/1.0"

                Try
                    Using response = DirectCast(request.GetResponse(), HttpWebResponse)
                        Dim statusCode = CInt(response.StatusCode)
                        Debug.WriteLine($"{url} returned {response.StatusCode} ({statusCode})")
                        ' Accept only 2xx status codes as available.
                        Return statusCode >= 200 AndAlso statusCode < 300
                    End Using
                Catch webEx As WebException
                    ' Check if this is a 4xx/5xx error - if so, return False immediately.
                    If webEx.Response IsNot Nothing Then
                        Dim httpResponse = TryCast(webEx.Response, HttpWebResponse)
                        If httpResponse IsNot Nothing Then
                            Dim statusCode = CInt(httpResponse.StatusCode)
                            Debug.WriteLine($"{url} HEAD failed with {httpResponse.StatusCode} ({statusCode})")
                            httpResponse.Close()

                            ' 4xx (client errors including 404) and 5xx (server errors) mean URL is not available.
                            If statusCode >= 400 Then
                                Return False
                            End If
                        End If
                    End If

                    ' HEAD might not be supported; retry with GET method.
                    Dim getRequest = DirectCast(WebRequest.Create(url), HttpWebRequest)
                    getRequest.Method = "GET"
                    getRequest.Timeout = 5000
                    getRequest.AllowAutoRedirect = True
                    getRequest.MaximumAutomaticRedirections = 5
                    getRequest.UserAgent = $"{AN}/1.0"

                    Using response = DirectCast(getRequest.GetResponse(), HttpWebResponse)
                        Dim statusCode = CInt(response.StatusCode)
                        Debug.WriteLine($"{url} GET returned {response.StatusCode} ({statusCode})")
                        ' Accept only 2xx status codes as available.
                        Return statusCode >= 200 AndAlso statusCode < 300
                    End Using
                End Try

            Catch ex As WebException
                ' Handle any remaining WebExceptions (including 404, 500, etc.).
                If ex.Response IsNot Nothing Then
                    Dim httpResponse = TryCast(ex.Response, HttpWebResponse)
                    If httpResponse IsNot Nothing Then
                        Dim statusCode = CInt(httpResponse.StatusCode)
                        Debug.WriteLine($"{url} failed with {httpResponse.StatusCode} ({statusCode})")
                        httpResponse.Close()
                        ' Only 2xx means available.
                        Return False
                    End If
                End If
                Debug.WriteLine($"{url} failed: {ex.Message}")
                Return False
            Catch ex As Exception
                Debug.WriteLine($"{url} failed: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Returns whether a beta warning should be shown.
        ''' Triggers when either:
        ''' - <c>BetaWarningDays</c> or more days have passed since the last warning, or
        ''' - <c>BetaWarningInterval</c> or more starts occurred since last reset.
        ''' </summary>
        ''' <returns><c>True</c> if a beta warning should be shown; otherwise <c>False</c>.</returns>
        Private Shared Function ShouldShowBetaWarning() As Boolean
            Try
                ' Check days since last warning.
                Dim lastWarning = My.Settings.LastBetaWarningDate
                Dim daysSinceLastWarning = If(lastWarning = Date.MinValue, Integer.MaxValue, CInt((Date.Now.Date - lastWarning.Date).TotalDays))

                ' Increment and check start count since last warning.
                Dim startCount = My.Settings.BetaWarningStartCount + 1
                My.Settings.BetaWarningStartCount = startCount
                My.Settings.Save()

                ' Show warning if either condition is met.
                Return daysSinceLastWarning >= BetaWarningDays OrElse startCount >= BetaWarningInterval

            Catch
                Return True
            End Try
        End Function

        ''' <summary>
        ''' Records that a beta warning was shown by persisting the current date and resetting the start counter.
        ''' </summary>
        Private Shared Sub RecordBetaWarningShown()
            Try
                My.Settings.LastBetaWarningDate = Date.Now.Date
                My.Settings.BetaWarningStartCount = 0
                My.Settings.Save()
            Catch
            End Try
        End Sub

    End Class
End Namespace