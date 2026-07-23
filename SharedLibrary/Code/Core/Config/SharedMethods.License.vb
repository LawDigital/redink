' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.License.vb
' Purpose: Retained shared license helpers used by the current license regime.
'
' NOTE: The former legacy GA/Beta license engine (LicenseOK_Legacy and its
'       GA/Beta-specific helpers) has been removed. License validation now flows
'       exclusively through LicenseOK in SharedMethods.License.Core.vb (Private,
'       Pro, and offline-domain licenses).
'
'       The grace-period and expiry-warning helpers below are still used by the
'       current Private license flow (SharedMethods.License.Private.vb).
' =============================================================================

Option Strict On
Option Explicit On

Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary
    Partial Public Class SharedMethods

        ''' <summary>
        ''' Parses a version date from an RDV string by extracting 6 digits after "V." in DDMMYY format.
        ''' Returns <see cref="Date.Now"/> if parsing fails.
        ''' </summary>
        Private Shared Function ParseVersionDateFromRDV(rdv As String) As Date
            Try
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

            Return Date.Now
        End Function

        ''' <summary>
        ''' Opens the current license type selection dialog. Retained because the grace/expiry
        ''' warning flows offer the user a way to update their license.
        ''' </summary>
        Public Shared Function ShowLicenseEntryForm(context As ISharedContext) As Boolean
            Try
                Return ShowLicenseTypeSelectionDialog(context)
            Catch ex As Exception
                ShowCustomMessageBox($"Error showing license form: {ex.Message}", AN)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Shows a grace period warning (when enabled) during the grace period after expiry.
        ''' Warning frequency is throttled via <c>ShouldShowGracePeriodWarning</c>.
        ''' </summary>
        Private Shared Sub CheckGracePeriodWarning(context As ISharedContext, expiredDate As Date)
            Try
                If LicenseNoWarning Then Return

                Dim gracePeriodEnd As Date = expiredDate.AddDays(GracePeriodDays)
                Dim remainingDays As Integer = CInt((gracePeriodEnd.Date - Date.Now.Date).TotalDays)

                If ShouldShowGracePeriodWarning() Then
                    Dim msg = BuildLicenseMessage(
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

                    RecordGracePeriodWarningShown()
                End If
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' Returns whether a grace period warning should be shown on this start.
        ''' </summary>
        Private Shared Function ShouldShowGracePeriodWarning() As Boolean
            Try
                Dim startCount As Integer = 0
                Try
                    startCount = My.Settings.GracePeriodWarningStartcount + 1
                Catch
                    startCount = 1
                End Try

                My.Settings.GracePeriodWarningStartcount = startCount
                My.Settings.Save()

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
        ''' Shows license expiry warnings at configured day thresholds prior to expiry.
        ''' Warning frequency is throttled using <c>My.Settings.LicenseWarningStartCount</c>.
        ''' </summary>
        Private Shared Sub CheckLicenseExpiryWarnings(context As ISharedContext, expiryDate As Date)
            Try
                Dim daysUntilExpiry = CInt((expiryDate.Date - Date.Now.Date).TotalDays)

                For Each warningDay In LicenseWarningDays
                    If daysUntilExpiry = warningDay Then
                        If Not ShouldShowLicenseWarning() Then
                            Exit For
                        End If

                        Dim msg = BuildLicenseMessage(
                            $"Your license for {AN} for {context.RDV} will EXPIRE in {daysUntilExpiry} day(s) " &
                            $"on {expiryDate:d}." & vbCrLf & vbCrLf &
                            If(LicenseFromConfig,
                               "Your license is configured centrally. Contact your administrator to renew.",
                               $"Please update your license at {AN4} or contact your administrator. Updating the license information is possible via 'Settings', then 'About {AN}'." & vbCrLf & vbCrLf & "Would you like to update your license information now?"))

                        If LicenseFromConfig Then
                            ShowCustomMessageBox(msg, $"{AN} License Warning")
                        Else
                            Dim result = ShowCustomYesNoBox(msg, "Update License", "Later", $"{AN} License Warning")
                            If result = 1 Then
                                ShowLicenseEntryForm(context)
                            End If
                        End If

                        RecordLicenseWarningShown()
                        Exit For
                    End If
                Next
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' Returns whether an expiry warning should be shown on this start.
        ''' </summary>
        Private Shared Function ShouldShowLicenseWarning() As Boolean
            Try
                Dim startCount As Integer = 0
                Try
                    startCount = My.Settings.LicenseWarningStartCount + 1
                Catch
                    startCount = 1
                End Try

                My.Settings.LicenseWarningStartCount = startCount
                My.Settings.Save()

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
        ''' Appends <c>LicenseContact</c> (if present) to the given message.
        ''' </summary>
        Private Shared Function BuildLicenseMessage(baseMessage As String) As String
            If Not String.IsNullOrEmpty(LicenseContact) Then
                Return baseMessage & vbCrLf & vbCrLf & LicenseContact
            End If
            Return baseMessage
        End Function

    End Class
End Namespace
