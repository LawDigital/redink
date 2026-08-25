' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.License.Private.vb
' Purpose: Private license functionality for the licensing subsystem.
'
' Architecture:
'  - Stored license detection:
'     - A Private license is identified by `My.Settings.License_Type = "Private"`.
'  - Validation & enforcement (`ProcessStoredPrivateLicense`):
'     - Expiration is determined via `My.Settings.License_ValidUntil`.
'     - An optional grace period is applied using `GracePeriodDays` and `CheckGracePeriodWarning`.
'     - Periodic reconfirmation is required based on:
'         - RDV change (`My.Settings.License_PrivateVersion` vs `context.RDV`), or
'         - Time since last confirmation (`PrivateReconfirmIntervalDays`).
'     - User may postpone reconfirmation up to `PrivateReconfirmDismissMax` times (not allowed on version change).
'  - Activation (`ShowPrivateConfirmationFlow`):
'     - Presents private-use terms, optionally shows full license text, then stores license to settings.
'  - License text display:
'     - Downloads `license.txt` from `LicenseFileUrl` using `HttpClient` with TLS 1.2/1.3 enabled,
'       and renders it via `ShowHTMLCustomMessageBox`.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Net
Imports System.Net.Http
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary
    Partial Public Class SharedMethods

#Region "Private License Constants"

        ''' <summary>
        ''' Compliance message shown in Private license reconfirmation dialogs.
        ''' </summary>
        Private Const PrivateComplianceMessage As String = "This Private License is for personal, non-commercial use only. " &
            "You confirm that you are not using this software for any business, professional, or organizational purposes."

#End Region

#Region "Private License Processing"

        ''' <summary>
        ''' Returns whether stored settings indicate a Private license.
        ''' </summary>
        Private Shared Function HasStoredPrivateLicense() As Boolean
            Try
                Dim licenseType = My.Settings.License_Type
                Return Not String.IsNullOrEmpty(licenseType) AndAlso
                       licenseType.Equals("Private", StringComparison.OrdinalIgnoreCase)
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Processes a stored Private license by checking expiration and reconfirmation requirements.
        ''' </summary>
        ''' <param name="context">Shared execution context used for version information and UI.</param>
        ''' <returns><see langword="True"/> when the Private license flow allows continued use; otherwise <see langword="False"/>.</returns>
        Private Shared Function ProcessStoredPrivateLicense(context As ISharedContext) As Boolean
            Try
                Dim validUntil = My.Settings.License_ValidUntil
                Dim confirmedOn = My.Settings.License_PrivateConfirmedOn
                Dim storedVersion = My.Settings.License_PrivateVersion

                LogLicenseEvent("Private License", $"ValidUntil={validUntil:d}, ConfirmedOn={confirmedOn:d}")

                ' Check if license has expired
                If validUntil > Date.MinValue AndAlso Date.Now > validUntil Then
                    ' Check grace period
                    If GracePeriodDays > 0 AndAlso Date.Now <= validUntil.AddDays(GracePeriodDays) Then
                        _currentLicenseState = LicenseState.PrivateActive
                        CheckGracePeriodWarning(context, validUntil)
                        LicenseStatus = "Private License"
                        Return True
                    End If

                    _currentLicenseState = LicenseState.PrivateExpired
                    LogLicenseEvent("Private License", "Expired")
                    Return HandleExpiredPrivateLicense(context)
                End If

                ' Check if reconfirmation is needed (version change or interval elapsed)
                If NeedsPrivateReconfirmation(context) Then
                    Return HandlePrivateReconfirmation(context)
                End If

                _currentLicenseState = LicenseState.PrivateActive
                LicenseStatus = "Private License"

                ' Show expiry warnings if applicable. My.Settings.License_ValidUntil is
                ' authoritative for Private expiry and takes precedence over any config value.
                Dim expiryDate As Date = My.Settings.License_ValidUntil
                If Not LicenseNoWarning AndAlso expiryDate > Date.Now Then
                    CheckLicenseExpiryWarnings(context, expiryDate)
                End If

                Return True

            Catch ex As Exception
                LogLicenseEvent("Private License Error", ex.Message)
                Return ShowLicenseTypeSelectionDialog(context)
            End Try
        End Function

        ''' <summary>
        ''' Returns whether Private license reconfirmation is required due to version change or elapsed interval.
        ''' </summary>
        ''' <param name="context">Shared execution context providing the current `RDV`.</param>
        Private Shared Function NeedsPrivateReconfirmation(context As ISharedContext) As Boolean
            Try
                Dim confirmedOn = My.Settings.License_PrivateConfirmedOn
                Dim storedVersion = My.Settings.License_PrivateVersion

                ' Check if version has changed
                If Not String.IsNullOrEmpty(storedVersion) AndAlso
                   Not storedVersion.Equals(context.RDV, StringComparison.OrdinalIgnoreCase) Then
                    LogLicenseEvent("Reconfirmation Needed", $"Version changed: {storedVersion} -> {context.RDV}")
                    Return True
                End If

                ' Check if interval has elapsed
                If confirmedOn > Date.MinValue Then
                    Dim daysSinceConfirmation = CInt((Date.Now.Date - confirmedOn.Date).TotalDays)
                    If daysSinceConfirmation >= PrivateReconfirmIntervalDays Then
                        LogLicenseEvent("Reconfirmation Needed", $"Days since confirmation: {daysSinceConfirmation}")
                        Return True
                    End If
                End If

                Return False

            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Handles the reconfirmation flow for a stored Private license.
        ''' </summary>
        ''' <remarks>
        ''' The user can postpone reconfirmation up to `PrivateReconfirmDismissMax` times, except when `context.RDV`
        ''' differs from the stored version.
        ''' </remarks>
        Private Shared Function HandlePrivateReconfirmation(context As ISharedContext) As Boolean
            Try
                _currentLicenseState = LicenseState.PrivateReconfirmNeeded

                Dim dismissCount = My.Settings.License_PrivateDismissCount
                Dim remainingDismissals = PrivateReconfirmDismissMax - dismissCount

                ' Check if version changed - always require confirmation for new versions
                Dim storedVersion = My.Settings.License_PrivateVersion
                Dim isNewVersion = Not String.IsNullOrEmpty(storedVersion) AndAlso
                                   Not storedVersion.Equals(context.RDV, StringComparison.OrdinalIgnoreCase)

                Dim msg As String
                msg = $"Your Private License for {AN} requires periodic reconfirmation." & vbCrLf & vbCrLf &
                      PrivateComplianceMessage & vbCrLf & vbCrLf &
                      "Please confirm that your use remains for private, non-business, " &
                      "and non-professional purposes only."

                ' Add dismissal info if applicable (not for new versions)
                Dim canDismiss = remainingDismissals > 0 AndAlso Not isNewVersion
                If canDismiss Then
                    msg &= vbCrLf & vbCrLf & $"(You may postpone this {remainingDismissals} more time(s))"
                End If

                Dim buttonText = If(canDismiss, "Later", "Cancel")
                Dim result = ShowCustomYesNoBox(
                    msg,
                    "I Confirm",
                    buttonText,
                    $"{AN} - License Reconfirmation",
                    extraButtonText:="Get Pro License",
                    extraButtonAction:=Sub() Process.Start(New ProcessStartInfo(AN4 & ProSubUrl) With {.UseShellExecute = True}),
                    CloseAfterExtra:=False)

                If result = 1 Then
                    ' User confirmed - update timestamps and reset dismiss count
                    UpdatePrivateReconfirmation(context)
                    _currentLicenseState = LicenseState.PrivateActive
                    LogLicenseEvent("Reconfirmation", "User confirmed", alwaysLog:=True)
                    Return True
                Else
                    If canDismiss Then
                        ' User dismissed but has remaining dismissals
                        My.Settings.License_PrivateDismissCount = dismissCount + 1
                        My.Settings.Save()

                        LogLicenseEvent("Reconfirmation Dismissed", $"Remaining: {remainingDismissals - 1}")

                        _currentLicenseState = LicenseState.PrivateActive
                        Return True
                    Else
                        ' User declined and no more dismissals available
                        ShowCustomMessageBox(
                            "You must confirm your private use to continue using " & AN & "." & vbCrLf & vbCrLf &
                            "Alternatively, you can obtain a Professional License at " & AN4,
                            $"{AN} - License Required")

                        Return ShowLicenseTypeSelectionDialog(context)
                    End If
                End If

            Catch ex As Exception
                LogLicenseEvent("Reconfirmation Error", ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Updates the stored Private license confirmation timestamp, version, and dismissal counter.
        ''' When the version has changed, also extends the license validity period.
        ''' </summary>
        ''' <param name="context">Shared execution context providing the current `RDV`.</param>
        Private Shared Sub UpdatePrivateReconfirmation(context As ISharedContext)
            Try
                Dim storedVersion = My.Settings.License_PrivateVersion
                Dim isNewVersion = Not String.IsNullOrEmpty(storedVersion) AndAlso
                           Not storedVersion.Equals(context.RDV, StringComparison.OrdinalIgnoreCase)

                My.Settings.License_PrivateConfirmedOn = Date.Now.Date
                My.Settings.License_PrivateVersion = context.RDV
                My.Settings.License_PrivateDismissCount = 0

                ' Extend license validity when upgrading to a new version
                If isNewVersion Then
                    Dim versionDate = ParseVersionDateFromRDV(context.RDV)
                    Dim newValidUntil = versionDate.AddYears(PrivateLicenseYears)
                    My.Settings.License_ValidUntil = newValidUntil
                    LogLicenseEvent("License Extended", $"New ValidUntil={newValidUntil:d} for version {context.RDV}")
                End If

                My.Settings.Save()
            Catch ex As Exception
                LogLicenseEvent("Update Reconfirmation Error", ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Handles the expired Private license flow by offering update and license switching options.
        ''' </summary>
        ''' <param name="context">Shared execution context used for UI.</param>
        ''' <returns><see langword="True"/> when a subsequent license flow succeeds; otherwise <see langword="False"/>.</returns>
        Private Shared Function HandleExpiredPrivateLicense(context As ISharedContext) As Boolean
            Try
                Dim msg = $"Your Private License for {AN} has expired." & vbCrLf & vbCrLf &
                          "The Private License is valid for " & PrivateLicenseYears & " years from the version release date." & vbCrLf & vbCrLf &
                          "You have the following options:" & vbCrLf &
                          "• Update to a newer version (free for private use)" & vbCrLf &
                          "• Obtain a Professional License for business use" & vbCrLf & vbCrLf &
                          "Would you like to check for updates or change your license?"

                ' Track if extra button was clicked
                Dim extraButtonClicked As Boolean = False

                Dim result = ShowCustomYesNoBox(
                    msg,
                    "Check for Updates",
                    "Cancel",
                    $"{AN} - Private License Expired",
                    extraButtonText:="Get Pro License",
                    extraButtonAction:=Sub()
                                           extraButtonClicked = True
                                           Try
                                               Process.Start(New ProcessStartInfo(AN4 & ProSubUrl) With {.UseShellExecute = True})
                                           Catch
                                           End Try
                                       End Sub,
                    CloseAfterExtra:=True)

                ' If extra button was clicked, show license selection dialog
                If extraButtonClicked Then
                    Return ShowLicenseTypeSelectionDialog(context)
                End If

                Select Case result
                    Case 1 ' Check for Updates
                        ' Open the update/download page
                        Try
                            Process.Start(New ProcessStartInfo(AN4 & UpdateSubUrl) With {.UseShellExecute = True})
                        Catch
                        End Try

                        ' Also offer to switch license type
                        Dim switchMsg = "After updating to a newer version, restart the application." & vbCrLf & vbCrLf &
                                        "If you'd prefer to switch to a Professional License now, click 'Change License'."

                        Dim switchResult = ShowCustomYesNoBox(
                            switchMsg,
                            "Change License",
                            "Close",
                            $"{AN} - Update Available")

                        If switchResult = 1 Then
                            Return ShowLicenseTypeSelectionDialog(context)
                        End If

                        Return False

                    Case Else ' Cancel
                        Return False
                End Select

            Catch ex As Exception
                LogLicenseEvent("Expired Private License Error", ex.Message)
                Return False
            End Try
        End Function

#End Region

#Region "Private License Confirmation Flow"

        ''' <summary>
        ''' Runs the Private license activation flow and stores the resulting license in settings.
        ''' </summary>
        ''' <param name="context">Shared execution context used for version calculation and UI.</param>
        ''' <returns><see langword="True"/> when confirmed and persisted; otherwise <see langword="False"/>.</returns>
        Private Shared Function ShowPrivateConfirmationFlow(context As ISharedContext) As Boolean
            Try
                If Not ShowPrivateUseConfirmation() Then
                    Return False
                End If

                SavePrivateLicenseToSettings(context)

                LogLicenseEvent("Private License", "Confirmed and saved", alwaysLog:=True)

                ShowCustomMessageBox(
                    $"Thank you! Your Private License has been activated. You will be asked to reconfirm your private use status periodically." & vbCrLf & vbCrLf &
                    $"Valid until: {My.Settings.License_ValidUntil:d}",
                    $"{AN} - License Activated")

                Return True

            Catch ex As Exception
                LogLicenseEvent("Private Confirmation Error", ex.Message)
                ShowCustomMessageBox($"Error during license confirmation: {ex.Message}", AN)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Shows the Private use confirmation dialog, including an option to view the full license text.
        ''' </summary>
        ''' <returns><see langword="True"/> when the user confirms; otherwise <see langword="False"/>.</returns>
        Private Shared Function ShowPrivateUseConfirmation() As Boolean
            Dim confirmationText As String =
                $"The Private License for {AN} is free of charge and is exclusively " &
                "for private, non-commercial use." & vbCrLf & vbCrLf &
                "This license may NOT be used for:" & vbCrLf &
                "  • Any business or commercial purposes" & vbCrLf &
                "  • Professional or work-related activities" & vbCrLf &
                "  • Use by or within organizations, companies, or institutions" & vbCrLf & vbCrLf &
                $"For all details, get the full license text from {AN4} or use the button below." & vbCrLf & vbCrLf &
                $"For professional use, please obtain a Professional License at {AN4}{ProSubUrl}." & vbCrLf & vbCrLf &
                "By clicking 'I Confirm', you confirm that your use of this product is private, " &
                "non-commercial and non-professional, and that you accept the license terms."

            Dim result = ShowCustomYesNoBox(
                confirmationText,
                "I Confirm",
                "Cancel",
                $"{AN} - Private License Confirmation",
                extraButtonText:="View License",
                extraButtonAction:=Sub() ShowLicenseText(),
                CloseAfterExtra:=False)

            Return (result = 1)
        End Function

        ''' <summary>
        ''' Saves a newly confirmed Private license to `My.Settings` and updates related globals.
        ''' </summary>
        ''' <param name="context">Shared execution context providing the current `RDV`.</param>
        Private Shared Sub SavePrivateLicenseToSettings(context As ISharedContext)
            Try
                Dim versionDate = ParseVersionDateFromRDV(context.RDV)
                Dim validUntil = versionDate.AddYears(PrivateLicenseYears)

                My.Settings.License_Type = "Private"
                My.Settings.License_PrivateConfirmedOn = Date.Now.Date
                My.Settings.License_PrivateVersion = context.RDV
                My.Settings.License_PrivateDismissCount = 0
                My.Settings.License_ValidUntil = validUntil
                My.Settings.License_State = LicenseState.PrivateActive.ToString()
                My.Settings.Save()

                ' Update global variables
                LicenseStatus = "Private License"
                _currentLicenseState = LicenseState.PrivateActive

                ' Clear any legacy license data
                ClearLegacyLicense()

            Catch ex As Exception
                LogLicenseEvent("Save Private License Error", ex.Message)
            End Try
        End Sub

#End Region

#Region "License Text Display"

        ''' <summary>
        ''' Downloads and displays the license text from `LicenseFileUrl`.
        ''' </summary>
        Private Shared Sub ShowLicenseText()
            Try
                Dim licenseText = TryDownloadLicenseText()

                If Not String.IsNullOrWhiteSpace(licenseText) Then
                    Dim htmlContent = "<pre style=""font-family: Consolas, monospace; font-size: 12px; white-space: pre-wrap; word-wrap: break-word;"">" &
                                      System.Web.HttpUtility.HtmlEncode(licenseText) &
                                      "</pre>"
                    ShowHTMLCustomMessageBox(htmlContent, $"{AN} - License Terms")
                Else
                    ShowCustomMessageBox(
                        $"The license text could not be downloaded from {LicenseFileUrl}." & vbCrLf & vbCrLf &
                        $"Please view the license terms at {AN4}.",
                        $"{AN} - License")
                End If
            Catch ex As Exception
                ShowCustomMessageBox(
                    $"Error loading license: {ex.Message}" & vbCrLf & vbCrLf &
                    $"Please view the license terms at {AN4}.",
                    $"{AN} - License")
            End Try
        End Sub

        ''' <summary>
        ''' Attempts to download license text from `LicenseFileUrl`.
        ''' </summary>
        ''' <returns>The downloaded license text, or <see langword="Nothing"/> if unavailable.</returns>
        Private Shared Function TryDownloadLicenseText() As String
            Try
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls13

                Using client As New HttpClient()
                    client.Timeout = TimeSpan.FromSeconds(10)
                    client.DefaultRequestHeaders.Add("User-Agent", $"{AN}/1.0")

                    Dim text = client.GetStringAsync(LicenseFileUrl).ConfigureAwait(False).GetAwaiter().GetResult()

                    If Not String.IsNullOrWhiteSpace(text) Then
                        LogLicenseEvent("License Text", $"Downloaded from: {LicenseFileUrl}")
                        Return text
                    End If
                End Using

                Return String.Empty

            Catch ex As Exception
                Dim message = If(TypeOf ex Is AggregateException,
                                 DirectCast(ex, AggregateException).InnerException?.Message,
                                 ex.Message)
                LogLicenseEvent("License Download Error", $"URL: {LicenseFileUrl}, Error: {message}")
                Return String.Empty
            End Try
        End Function

#End Region

    End Class
End Namespace
