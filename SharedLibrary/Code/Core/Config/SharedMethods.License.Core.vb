' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.License.Core.vb
' Purpose: Core license validation and state management for Private and Pro licenses,
'          including API-based verification/activation and legacy-license migration enforcement.
'
' Architecture:
'  - Flow Priority in `LicenseOK`:
'     1. License disabled via config (`LicenseNoCheck`) → allow use (state set to `ProActive`).
'     2. Config-based Pro license (`LicenseKey` in config) → process/activate via API or UI.
'     3. Stored Pro license → periodic API verification with offline grace handling.
'     4. Legacy Pro license → valid until `LegacyRegimeEndDate`, then enforce migration with grace.
'     5. Stored Private license → processed with periodic compliance confirmation (handled in other partial file).
'     6. No license → clear legacy Private license (if present), show license selection UI.
'  - Stored state:
'     - New regime: persisted in `My.Settings.License_*`.
'     - Legacy regime: persisted in `My.Settings.LicenseStatus`, `My.Settings.LicensedTill`, `My.Settings.LicenseUsers`.
'  - Network failure handling:
'     - If previously API-confirmed: allow continuation for `OfflineGracePeriodDays` with warnings.
'     - If never API-confirmed: fail closed and require connectivity/license update.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary
    Partial Public Class SharedMethods

#Region "License Constants"

        ' API Configuration

        ' Default license source identifier (matches INI value `LicenseSource = redink.ai`).
        Private Const DefaultLicenseSource As String = "redink.ai"

        ' Hardcoded mapping of alternative license sources to their API base URLs.
        ' The key is the value users set in the INI under `LicenseSource`.
        ' "redink.ai" is the built-in default and uses the standard redink.ai endpoint.
        ' Add additional license servers here as needed.
        Private Shared ReadOnly LicenseApiBaseUrls As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {DefaultLicenseSource, "https://redink.ai/?wc-api=wc-am-api"},
            {"alt-server", "https://redink.ai/?wc-api=wc-am-api"}
        }

        ' Currently selected license source (loaded from INI in LoadBasicLicenseSettings).
        Private Shared _licenseSource As String = DefaultLicenseSource

        ''' <summary>
        ''' Resolves the active license API base URL based on the configured `LicenseSource`
        ''' (INI key `LicenseSource`). Defaults to the redink.ai endpoint if unset or unknown.
        ''' </summary>
        Private Shared ReadOnly Property LicenseApiBaseUrl As String
            Get
                Dim source = If(String.IsNullOrWhiteSpace(_licenseSource), DefaultLicenseSource, _licenseSource.Trim())
                Dim url As String = Nothing
                If LicenseApiBaseUrls.TryGetValue(source, url) Then
                    Return url
                End If
                ' Unknown source - log and fall back to default.
                LogLicenseEvent("License Source", $"Unknown LicenseSource '{source}' - falling back to '{DefaultLicenseSource}'", alwaysLog:=True)
                Return LicenseApiBaseUrls(DefaultLicenseSource)
            End Get
        End Property

        Private Const ApiTimeoutMs As Integer = 10000
        Private Const ApiRetryCount As Integer = 3

        ' Private License
        Private Const PrivateReconfirmIntervalDays As Integer = 30
        Private Const PrivateReconfirmDismissMax As Integer = 3
        Private Const PrivateLicenseConfirmationText As String = "I confirm that I will use this software only for private, non-business, and non-professional purposes."

        ' Professional License
        Private Const ProLicenseCheckIntervalDays As Integer = 5
        Private Const OfflineGracePeriodDays As Integer = 3
        Private Const RevocationGracePeriodDays As Integer = 1

        ' Testing Pro License
        Private Const TestingProComplianceIntervalStartups As Integer = 15 ' 0 = every startup
        Private Const TestingProProductIdPrefix As String = "1727"
        Private Const TestingProComplianceMessage As String = "This Pro Test Tec license is for non-productive testing purposes only. " &
                                    "You confirm that you are not using this software for any productive purposes. " &
                                    "For productive use, please obtain another suitable Professional License."

        ' Legacy Transition
        Private Shared ReadOnly LegacyRegimeEndDate As Date = #1/31/2026#
        Private Const LegacyMigrationGracePeriodDays As Integer = 12
        Private Const LegacyMigrationPromptIntervalStartups As Integer = 15

        ' Warning/Prompt Intervals
        Private Const OfflineGraceWarningIntervalStartups As Integer = 5

        ''' <summary>
        ''' Returns the URL of the public `license.txt` file.
        ''' </summary>
        Private Shared ReadOnly Property LicenseFileUrl As String
            Get
                Return $"{AppsUrl}{AppsUrlDir}license.txt"
            End Get
        End Property

        ''' <summary>
        ''' Returns a standard contact/help message for license acquisition issues.
        ''' </summary>
        Private Shared ReadOnly Property StandardLicenseContactMessage As String
            Get
                Dim msg = $"Please obtain a new license at {AN4}{ProSubUrl}"
                If Not String.IsNullOrEmpty(LicenseContact) Then
                    msg &= $" or contact your administrator:{vbCrLf}{LicenseContact}"
                Else
                    msg &= " or contact your administrator."
                End If
                Return msg
            End Get
        End Property

        ''' <summary>
        ''' Returns a standard contact/help message for general support.
        ''' </summary>
        Private Shared ReadOnly Property StandardSupportContactMessage As String
            Get
                Dim msg = $"For support, contact us at {AN4}"
                If Not String.IsNullOrEmpty(LicenseContact) Then
                    msg &= $" or contact your administrator:{vbCrLf}{LicenseContact}"
                Else
                    msg &= " or contact your administrator."
                End If
                Return msg
            End Get
        End Property


#End Region


#Region "License State Enum"

        ''' <summary>
        ''' Defines the current state of the license system.
        ''' </summary>
        Public Enum LicenseState
            ''' <summary>No license configured.</summary>
            None = 0
            ''' <summary>Private license active and valid.</summary>
            PrivateActive = 1
            ''' <summary>Private license expired.</summary>
            PrivateExpired = 2
            ''' <summary>Private license needs periodic reconfirmation.</summary>
            PrivateReconfirmNeeded = 3
            ''' <summary>Professional license activated and valid.</summary>
            ProActive = 4
            ''' <summary>Professional license expired or revoked.</summary>
            ProExpired = 5
            ''' <summary>Has license key but not yet activated.</summary>
            ProPendingActivation = 6
            ''' <summary>Pro license in offline grace period.</summary>
            ProOfflineGrace = 7
            ''' <summary>Legacy Pro license still valid (until `LegacyRegimeEndDate`).</summary>
            LegacyProActive = 8
            ''' <summary>Legacy Pro license - migration required.</summary>
            LegacyProMigrationRequired = 9
            ''' <summary>Testing Pro license active.</summary>
            TestingProActive = 10
            ''' <summary>Backward compatibility: old-regime license still valid.</summary>
            LegacyActive = 11
            ''' <summary>Backward compatibility: old-regime license expired.</summary>
            LegacyExpired = 12
        End Enum

#End Region


#Region "License API Response Classes"

        ''' <summary>
        ''' Represents the parsed response from the license API.
        ''' </summary>
        Public Class LicenseApiResponse
            ''' <summary>Indicates whether the API call was successful.</summary>
            Public Property Success As Boolean

            ''' <summary>Raw status string (for example: "active").</summary>
            Public Property StatusCheck As String = ""

            ''' <summary>Indicates whether the license is activated for the supplied credentials.</summary>
            Public Property Activated As Boolean

            ''' <summary>Optional message returned by the API.</summary>
            Public Property Message As String = ""

            ''' <summary>Remaining value returned by API (field usage depends on API implementation).</summary>
            Public Property Remaining As Integer

            ''' <summary>Total activations purchased (as reported by API).</summary>
            Public Property TotalActivationsPurchased As Integer

            ''' <summary>Total activations used (as reported by API).</summary>
            Public Property TotalActivations As Integer

            ''' <summary>Remaining activations (as reported by API).</summary>
            Public Property ActivationsRemaining As Integer

            ''' <summary>Product title (as reported by API).</summary>
            Public Property ProductTitle As String = ""

            ''' <summary>Product ID (as reported by API).</summary>
            Public Property ProductId As String = ""

            ''' <summary>Next payment date string (as reported by API).</summary>
            Public Property NextPayment As String = ""

            ''' <summary>Error message (if call failed or was rejected).</summary>
            Public Property ErrorMessage As String = ""

            ''' <summary>Raw JSON returned by the API (if any).</summary>
            Public Property RawJson As String = ""
        End Class

#End Region

#Region "License Context Storage"

        ' Stored context for license operations
        Private Shared _licenseContext As ISharedContext
        Private Shared _licenseConfigDict As Dictionary(Of String, String)

        ' Current license state
        Private Shared _currentLicenseState As LicenseState = LicenseState.None

        ' Tracks whether a registry backup restore occurred during this startup
        Private Shared _restoredFromRegistryBackup As Boolean = False

        ''' <summary>
        ''' Gets the current license state as determined by `LicenseOK` and subsequent processing.
        ''' </summary>
        Public Shared ReadOnly Property CurrentLicenseState As LicenseState
            Get
                Return _currentLicenseState
            End Get
        End Property

        ''' <summary>
        ''' Gets the stored Product ID of the current Pro license, or empty if none.
        ''' </summary>
        Public Shared ReadOnly Property StoredProductId As String
            Get
                Try
                    If HasStoredProLicense() Then
                        Return My.Settings.License_ProductID
                    End If
                Catch
                End Try
                Return ""
            End Get
        End Property

#End Region

#Region "Main Entry Point"

        ''' <summary>
        ''' Main license check entry point.
        ''' </summary>
        ''' <param name="context">Shared execution context used for UI and version information.</param>
        ''' <param name="configDict">Configuration dictionary containing license-related settings.</param>
        ''' <returns><see langword="True"/> when license checks allow continued use; otherwise <see langword="False"/>.</returns>
        Public Shared Function LicenseOK(ByVal context As ISharedContext,
                                          ByVal configDict As Dictionary(Of String, String)) As Boolean
            Try
                ' Store context for use by helper methods
                _licenseContext = context
                _licenseConfigDict = configDict

                LogLicenseEvent("Check Started", $"Version: {context.RDV}")

                ' Load basic license settings from config
                LoadBasicLicenseSettings(configDict)

                ' Skip all checks if explicitly disabled
                If LicenseCheckDisabled Then
                    LogLicenseEvent("Check Skipped", "License checking disabled via config")
                    _currentLicenseState = LicenseState.ProActive
                    Return True
                End If

                ' Increment startup counter for compliance checks
                IncrementStartupCounter()

                ' ═══════════════════════════════════════════════════════════════
                ' STEP 1: Config-based Pro License (HIGHEST PRIORITY)
                ' If config has LicenseKey, immediately process Pro license
                ' Skip ALL other license types (legacy, private, etc.)
                ' ═══════════════════════════════════════════════════════════════
                If HasConfigLicenseKey(configDict) Then
                    LogLicenseEvent("Flow", "Config Pro License path (highest priority)")
                    Return ProcessConfigBasedProLicense(context, configDict)
                End If

                ' ═══════════════════════════════════════════════════════════════
                ' STEP 2: Stored Pro License (API-verified)
                ' ═══════════════════════════════════════════════════════════════
                If HasStoredProLicense() Then
                    LogLicenseEvent("Flow", "Stored Pro License path")
                    Return ProcessStoredProLicense(context)
                End If

                ' ═══════════════════════════════════════════════════════════════
                ' STEP 2b: Registry Backup Restore
                ' If My.Settings was wiped (e.g., VSTO update, profile reset),
                ' attempt silent restore from registry backup before continuing.
                ' ═══════════════════════════════════════════════════════════════
                If TryRestoreProLicenseFromRegistry() Then
                    LogLicenseEvent("Flow", "Pro License restored from registry backup")
                    Return ProcessStoredProLicense(context)
                End If


                ' ═══════════════════════════════════════════════════════════════
                ' STEP 3: Legacy Pro License (valid until LegacyRegimeEndDate)
                ' Legacy Private licenses are ignored/cleared, not processed here
                ' ═══════════════════════════════════════════════════════════════
                If HasLegacyProLicense() Then
                    LogLicenseEvent("Flow", "Legacy Pro License path")
                    Return ProcessLegacyProLicense(context)
                End If

                ' ═══════════════════════════════════════════════════════════════
                ' STEP 4: Stored Private License (with compliance confirmation)
                ' ═══════════════════════════════════════════════════════════════
                If HasStoredPrivateLicense() Then
                    LogLicenseEvent("Flow", "Private License path")
                    Return ProcessStoredPrivateLicense(context)
                End If

                ' ═══════════════════════════════════════════════════════════════
                ' STEP 5: No valid license found
                ' Clear any legacy private license and show welcome dialog
                ' ═══════════════════════════════════════════════════════════════
                ClearLegacyPrivateLicenseIfExists()

                LogLicenseEvent("Flow", "No license - showing welcome dialog")
                Return ShowLicenseTypeSelectionDialog(context)

            Catch ex As Exception
                ' Fail closed on unexpected errors
                LogLicenseEvent("ERROR", $"Unexpected error: {ex.Message}", alwaysLog:=True)
                Try
                    ShowCustomMessageBox(
                        $"License check encountered an error: {ex.Message}" & vbCrLf & vbCrLf &
                        "Please restart the application. If the problem persists, " & StandardLicenseContactMessage,
                        AN)
                Catch
                End Try
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Loads basic license settings from config dictionary.
        ''' </summary>
        ''' <param name="configDict">Configuration dictionary containing license-related settings.</param>
        Private Shared Sub LoadBasicLicenseSettings(configDict As Dictionary(Of String, String))
            Try
                ' Reset flags
                LicenseFromConfig = False
                LicenseCheckDisabled = False

                ' Load LicenseContact from config
                LicenseContact = If(configDict.ContainsKey("LicenseContact"), configDict("LicenseContact"), "")

                ' Load LicenseSource from config (INI key: LicenseSource).
                ' Default is "redink.ai", which selects the built-in redink.ai API base URL.
                ' Any other value must match an entry in LicenseApiBaseUrls; unknown values
                ' fall back to the default at resolution time (see LicenseApiBaseUrl property).
                Dim cfgSource As String = If(configDict.ContainsKey("LicenseSource"), configDict("LicenseSource"), "")
                _licenseSource = If(String.IsNullOrWhiteSpace(cfgSource), DefaultLicenseSource, cfgSource.Trim())

                ' Load LicenseNoWarning from config
                LicenseNoWarning = ParseBoolean(configDict, "LicenseNoWarning", False)

                ' Load LicenseNoCheck from config
                LicenseCheckDisabled = ParseBoolean(configDict, "LicenseNoCheck", False)

                ' Check if LicensedTill explicitly disables checking
                If configDict.ContainsKey("LicensedTill") Then
                    Dim configValue = configDict("LicensedTill").Trim()

                    If configValue.Equals("False", StringComparison.OrdinalIgnoreCase) OrElse
                       configValue.Equals("No", StringComparison.OrdinalIgnoreCase) Then
                        LicenseCheckDisabled = False
                        Return
                    End If

                    Dim configDate As Date
                    If Date.TryParse(configValue, configDate) Then
                        If configDate > Date.Now.AddYears(LicenseCheckDisabledYears) Then
                            LicenseCheckDisabled = False
                            Return
                        End If
                    End If
                End If

            Catch ex As Exception
                LogLicenseEvent("Settings Load Error", ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Checks whether a non-empty `LicenseKey` is provided via config.
        ''' </summary>
        ''' <param name="configDict">Configuration dictionary.</param>
        Private Shared Function HasConfigLicenseKey(configDict As Dictionary(Of String, String)) As Boolean
            Return configDict.ContainsKey("LicenseKey") AndAlso
                   Not String.IsNullOrWhiteSpace(configDict("LicenseKey"))
        End Function

#End Region

#Region "License Type Selection Dialog"

        ''' <summary>
        ''' Determines whether a new private license can be created for the current build, based on
        ''' the build date (`RDV`) and `PrivateLicenseYears`.
        ''' </summary>
        Private Shared Function IsPrivateLicenseAvailable(context As ISharedContext) As Boolean
            Try
                Dim versionDate = ParseVersionDateFromRDV(context.RDV)
                Dim privateLicenseEndDate = versionDate.AddYears(PrivateLicenseYears)
                Return Date.Now <= privateLicenseEndDate
            Catch
                ' If date cannot be determined, allow private license
                Return True
            End Try
        End Function

        ''' <summary>
        ''' Shows the license type selection dialog (Private vs Professional).
        ''' </summary>
        ''' <param name="context">Shared execution context.</param>
        ''' <returns><see langword="True"/> when a license flow succeeds; otherwise <see langword="False"/>.</returns>
        Private Shared Function ShowLicenseTypeSelectionDialog(context As ISharedContext) As Boolean

            ' Check if private license is still available for this version
            Dim privateAvailable As Boolean = IsPrivateLicenseAvailable(context)
            Dim DisablePrivateText As String
            If privateAvailable Then
                DisablePrivateText = "For personal, non-commercial use only"
            Else
                DisablePrivateText = $"This version is too old for a new private license. Update at {AN4}{UpdateSubUrl} or get a Pro License."
            End If
            If HasLegacyProLicense() Then
                privateAvailable = False
                DisablePrivateText = "For personal, non-commercial use only (for private use, uninstall the add-in first)"
            End If

            Using form As New Form()
                ' DPI awareness
                form.AutoScaleMode = AutoScaleMode.Dpi
                form.AutoScaleDimensions = New SizeF(96.0F, 96.0F)

                form.Text = $"{AN} - License Selection"
                form.StartPosition = FormStartPosition.CenterScreen
                form.FormBorderStyle = FormBorderStyle.FixedDialog
                form.MaximizeBox = False
                form.MinimizeBox = False
                form.ShowInTaskbar = True
                form.TopMost = True

                Try
                    Dim bmp As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                    form.Icon = Icon.FromHandle(bmp.GetHicon())
                Catch
                End Try

                Dim font As New Font("Segoe UI", 9.5F)
                form.Font = font

                ' Main TableLayoutPanel - AutoSize the dialog from content (then clamp)
                Dim mainLayout As New TableLayoutPanel() With {
                    .Dock = DockStyle.Fill,
                    .ColumnCount = 1,
                    .RowCount = 4,
                    .Padding = New Padding(20, 20, 20, 20),
                    .AutoSize = True,
                    .AutoSizeMode = AutoSizeMode.GrowAndShrink
                }
                mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize)) ' Title
                mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize)) ' Description
                mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize)) ' Options grid
                mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize)) ' Buttons
                form.Controls.Add(mainLayout)

                ' Row 0: Title
                Dim lblTitle As New Label() With {
                    .Text = $"Welcome to {AN}!",
                    .Font = New Font("Segoe UI", 14.0F, FontStyle.Bold),
                    .AutoSize = True,
                    .Margin = New Padding(0, 0, 0, 10)
                }
                mainLayout.Controls.Add(lblTitle, 0, 0)

                ' Row 1: Description
                Dim lblDescription As New Label() With {
                    .Text = "Please select your license type to continue:",
                    .AutoSize = True,
                    .Margin = New Padding(0, 0, 0, 15)
                }
                mainLayout.Controls.Add(lblDescription, 0, 1)

                ' Row 2: Two-column grid for radio options
                Dim optionsGrid As New TableLayoutPanel() With {
                    .AutoSize = True,
                    .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    .ColumnCount = 2,
                    .RowCount = 4,
                    .Margin = New Padding(0, 0, 0, 15)
                }
                ' Radio column has explicit width to add right-padding next to the circle.
                optionsGrid.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 34.0F)) ' Rad + right gap
                optionsGrid.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))       ' Labels

                optionsGrid.RowStyles.Add(New RowStyle(SizeType.AutoSize)) ' Private title row
                optionsGrid.RowStyles.Add(New RowStyle(SizeType.AutoSize)) ' Private desc row
                optionsGrid.RowStyles.Add(New RowStyle(SizeType.AutoSize)) ' Pro title row
                optionsGrid.RowStyles.Add(New RowStyle(SizeType.AutoSize)) ' Pro desc row
                mainLayout.Controls.Add(optionsGrid, 0, 2)

                ' Row 0, Col 0: Private radio button (aligned with title text)
                Dim rbPrivate As New RadioButton() With {
                    .Text = "",
                    .AutoSize = True,
                    .Anchor = AnchorStyles.Left,
                    .Margin = New Padding(0, 2, 0, 0),
                    .Enabled = privateAvailable
                }
                optionsGrid.Controls.Add(rbPrivate, 0, 0)

                ' Row 0, Col 1: Private title label
                Dim privateTitle = If(privateAvailable,
                                      "Private License (Free)",
                                      "Private License (Not Available)")
                Dim lblPrivateTitle As New Label() With {
                    .Text = privateTitle,
                    .Font = New Font("Segoe UI", 10.0F, FontStyle.Bold),
                    .AutoSize = True,
                    .Margin = New Padding(0, 2, 0, 2),
                    .Cursor = If(privateAvailable, Cursors.Hand, Cursors.Default),
                    .ForeColor = If(privateAvailable, SystemColors.ControlText, SystemColors.GrayText)
                }
                optionsGrid.Controls.Add(lblPrivateTitle, 1, 0)

                ' Row 1, Col 1: Private description
                Dim privateDesc = DisablePrivateText
                Dim lblPrivateDesc As New Label() With {
                    .Text = privateDesc,
                    .Font = New Font("Segoe UI", 9.0F),
                    .ForeColor = If(privateAvailable, Color.DimGray, SystemColors.GrayText),
                    .AutoSize = True,
                    .Margin = New Padding(0, 0, 0, 12),
                    .Cursor = If(privateAvailable, Cursors.Hand, Cursors.Default)
                }
                optionsGrid.Controls.Add(lblPrivateDesc, 1, 1)

                ' Row 2, Col 0: Professional radio button (aligned with title text)
                Dim rbProfessional As New RadioButton() With {
                    .Text = "",
                    .AutoSize = True,
                    .Anchor = AnchorStyles.Left,
                    .Margin = New Padding(0, 2, 0, 0)
                }
                optionsGrid.Controls.Add(rbProfessional, 0, 2)

                ' Row 2, Col 1: Professional title label
                Dim lblProTitle As New Label() With {
                    .Text = "Professional License",
                    .Font = New Font("Segoe UI", 10.0F, FontStyle.Bold),
                    .AutoSize = True,
                    .Margin = New Padding(0, 2, 0, 2),
                    .Cursor = Cursors.Hand
                }
                optionsGrid.Controls.Add(lblProTitle, 1, 2)

                ' Row 3, Col 1: Professional description
                Dim lblProDesc As New Label() With {
                    .Text = "For business, professional, or organizational use",
                    .Font = New Font("Segoe UI", 9.0F),
                    .ForeColor = Color.DimGray,
                    .AutoSize = True,
                    .Margin = New Padding(0, 0, 0, 0),
                    .Cursor = Cursors.Hand
                }
                optionsGrid.Controls.Add(lblProDesc, 1, 3)

                ' Click on labels selects radio button (only if enabled)
                If privateAvailable Then
                    AddHandler lblPrivateTitle.Click, Sub() rbPrivate.Checked = True
                    AddHandler lblPrivateDesc.Click, Sub() rbPrivate.Checked = True
                End If
                AddHandler lblProTitle.Click, Sub() rbProfessional.Checked = True
                AddHandler lblProDesc.Click, Sub() rbProfessional.Checked = True

                ' Row 3: Button panel (use default button padding => Padding(0))
                Dim buttonPanel As New FlowLayoutPanel() With {
                    .AutoSize = True,
                    .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    .FlowDirection = FlowDirection.LeftToRight,
                    .Margin = New Padding(0),
                    .Padding = New Padding(0),
                    .WrapContents = False
                }
                mainLayout.Controls.Add(buttonPanel, 0, 3)

                Dim btnContinue As New Button() With {
                    .Text = "Continue",
                    .AutoSize = True,
                    .Padding = New Padding(0),
                    .Margin = New Padding(0, 0, 10, 0),
                    .Enabled = False
                }
                buttonPanel.Controls.Add(btnContinue)

                Dim btnCancel As New Button() With {
                    .Text = "Cancel",
                    .AutoSize = True,
                    .Padding = New Padding(0),
                    .Margin = New Padding(0)
                }
                buttonPanel.Controls.Add(btnCancel)

                Dim dialogResult As Boolean = False
                Dim selectedPrivate As Boolean = False

                ' Enable continue when a selection is made
                AddHandler rbPrivate.CheckedChanged, Sub()
                                                         btnContinue.Enabled = rbPrivate.Checked OrElse rbProfessional.Checked
                                                         selectedPrivate = rbPrivate.Checked
                                                     End Sub

                AddHandler rbProfessional.CheckedChanged, Sub()
                                                              btnContinue.Enabled = rbPrivate.Checked OrElse rbProfessional.Checked
                                                              selectedPrivate = rbPrivate.Checked
                                                          End Sub

                AddHandler btnContinue.Click, Sub()
                                                  form.Close()

                                                  If selectedPrivate Then
                                                      dialogResult = ShowPrivateConfirmationFlow(context)
                                                  Else
                                                      dialogResult = ShowProLicenseEntryForm(context, "", "", "", prefilled:=False)
                                                  End If
                                              End Sub

                AddHandler btnCancel.Click, Sub()
                                                form.Close()
                                            End Sub

                ' Critical: after layout, size the form to content so buttons are fully visible (DPI-safe),
                ' then clamp to screen working area; finally freeze the size.
                AddHandler form.Shown,
                    Sub()
                        form.PerformLayout()
                        mainLayout.PerformLayout()

                        Dim wa As Rectangle = Screen.FromPoint(Cursor.Position).WorkingArea

                        Dim chromeW As Integer = form.Width - form.ClientSize.Width
                        Dim chromeH As Integer = form.Height - form.ClientSize.Height

                        Dim desiredClientW As Integer = mainLayout.PreferredSize.Width
                        Dim desiredClientH As Integer = mainLayout.PreferredSize.Height

                        desiredClientW = Math.Min(desiredClientW, wa.Width - chromeW - 40)
                        desiredClientH = Math.Min(desiredClientH, wa.Height - chromeH - 40)

                        desiredClientW = Math.Max(desiredClientW, 460)
                        desiredClientH = Math.Max(desiredClientH, 260)

                        form.ClientSize = New Size(desiredClientW, desiredClientH)

                        ' Make sure nothing can end up clipped later.
                        form.MinimumSize = form.Size
                        form.MaximumSize = form.Size
                    End Sub

                form.ShowDialog()
                Return dialogResult
            End Using
        End Function

#End Region


#Region "Startup Counter and Compliance"

        ''' <summary>
        ''' Increments the startup counter used for periodic compliance checks.
        ''' </summary>
        Private Shared Sub IncrementStartupCounter()
            Try
                Dim currentCount = My.Settings.License_StartupCount
                My.Settings.License_StartupCount = currentCount + 1
                My.Settings.Save()
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' Gets the current startup count used for compliance and warning intervals.
        ''' </summary>
        Private Shared Function GetStartupCount() As Integer
            Try
                Return My.Settings.License_StartupCount
            Catch
                Return 0
            End Try
        End Function

        ''' <summary>
        ''' Returns whether a compliance check is due based on the current startup count.
        ''' </summary>
        ''' <param name="intervalStartups">Interval in startups; 0 means every startup.</param>
        Private Shared Function IsComplianceCheckDue(intervalStartups As Integer) As Boolean
            ' If interval is 0, always check
            If intervalStartups <= 0 Then
                Return True
            End If

            Dim count = GetStartupCount()
            Return count > 0 AndAlso (count Mod intervalStartups) = 0
        End Function


        ''' <summary>
        ''' Shows the compliance confirmation dialog required for Testing Pro licenses.
        ''' </summary>
        ''' <returns><see langword="True"/> when user confirms; otherwise <see langword="False"/>.</returns>
        Private Shared Function ShowTestingProComplianceConfirmation() As Boolean
            Try
                Dim msg = TestingProComplianceMessage & vbCrLf & vbCrLf &
                          "Do you confirm that you are using this software in compliance with these license terms?"

                Dim result = ShowCustomYesNoBox(
                    msg,
                    "I Confirm",
                    "Cancel",
                    $"{AN} - Testing License Compliance",
                    extraButtonText:="Get Full License",
                    extraButtonAction:=Sub() Process.Start(New ProcessStartInfo(AN4 & ProSubUrl) With {.UseShellExecute = True}),
                    CloseAfterExtra:=False)

                If result = 1 Then
                    LogLicenseEvent("Testing Compliance", "User confirmed compliance", alwaysLog:=True)
                    Return True
                Else
                    LogLicenseEvent("Testing Compliance", "User declined - directing to support", alwaysLog:=True)
                    ShowCustomMessageBox(
                        "To continue using this software for testing purposes, you must confirm compliance with the testing license terms." & vbCrLf & vbCrLf &
                        "For a full Professional License, please contact support@redink.ai or visit " & AN4,
                        $"{AN} - License Required")
                    Return False
                End If

            Catch ex As Exception
                LogLicenseEvent("Testing Compliance Error", ex.Message)
                Return True ' Fail open on UI errors
            End Try
        End Function

        ''' <summary>
        ''' Returns whether the specified product ID identifies a Testing Pro license.
        ''' </summary>
        Private Shared Function IsTestingProLicenseByProductId(productId As String) As Boolean
            Return Not String.IsNullOrEmpty(productId) AndAlso
           productId.StartsWith(TestingProProductIdPrefix, StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function BuildConfigManagedLicenseSuccessMessage(licenseKey As String,
                                                                       productTitle As String,
                                                                       userId As String,
                                                                       totalActivations As Integer,
                                                                       totalPurchased As Integer,
                                                                       activationsRemaining As Integer,
                                                                       alreadyActivated As Boolean) As String
            Dim safeProductTitle = If(String.IsNullOrEmpty(productTitle) OrElse productTitle = "(no product title available)",
                                      "(unknown)",
                                      productTitle)

            Dim successMsg As New StringBuilder()

            If IsOfflineDomainLicenseKey(licenseKey) Then
                If alreadyActivated Then
                    successMsg.AppendLine($"Your offline-domain license has been verified locally for this copy of {AN}.")
                Else
                    successMsg.AppendLine($"Your offline-domain license has been verified locally and stored for this copy of {AN}.")
                End If
            Else
                If alreadyActivated Then
                    successMsg.AppendLine($"Your license is already activated for this User ID and has now been registered for this copy of {AN}.")
                Else
                    successMsg.AppendLine($"Your license has been automatically activated and registered for this copy of {AN}.")
                End If
            End If

            successMsg.AppendLine()
            successMsg.AppendLine($"Product: {safeProductTitle}")
            successMsg.AppendLine($"User ID: {userId}")

            If totalPurchased > 0 Then
                successMsg.AppendLine($"Activations: {totalActivations} of {totalPurchased} used ({activationsRemaining} remaining)")
            ElseIf IsOfflineDomainLicenseKey(licenseKey) Then
                successMsg.AppendLine(BuildOfflineDomainLicenseCountingMessage())
            End If

            successMsg.AppendLine()
            successMsg.AppendLine("License configured by your administrator.")

            If Not String.IsNullOrEmpty(LicenseContact) Then
                successMsg.AppendLine()
                successMsg.AppendLine($"Contact: {LicenseContact}")
            End If

            Return successMsg.ToString()
        End Function

        Private Shared Function BuildOfflineDomainLicenseCountingMessage() As String
            If _licenseContext Is Nothing OrElse String.IsNullOrWhiteSpace(_licenseContext.INI_LicenseCounterPath) Then
                Return "License Counting: not used"
            End If

            Dim counterMethod = _licenseContext.INI_LicenseCounterMethod.Trim()

            If String.IsNullOrWhiteSpace(counterMethod) Then
                Return "License Counting: used"
            End If

            Return $"License Counting: used ({counterMethod})"
        End Function

#End Region


#Region "Config-Based Pro License Processing"

        ''' <summary>
        ''' Processes a Pro license when the configuration provides a license key.
        ''' </summary>
        ''' <remarks>
        ''' Logic:
        ''' - If `LicenseClearAll = False`: apply config only when no valid Private or Pro license exists.
        ''' - If `LicenseClearAll = True`: apply config only when no matching valid Pro license exists where:
        '''   (a) `LicenseKey` + `LicenseProductID` match and config `LicenseUserID` is empty, OR
        '''   (b) `LicenseKey` + `LicenseProductID` + expanded `LicenseUserID` all match stored values.
        ''' - If `LicenseDoNotTouchProductID` or `LicenseDoNotTouchUserID` match stored values: preserve stored license.
        ''' </remarks>
        Private Shared Function ProcessConfigBasedProLicense(context As ISharedContext, configDict As Dictionary(Of String, String)) As Boolean
            Try
                ' Extract and expand config values
                Dim configLicenseKey = configDict("LicenseKey").Trim()
                Dim configProductId = If(configDict.ContainsKey("LicenseProductID"), configDict("LicenseProductID").Trim(), "")
                Dim configUserIdRaw = If(configDict.ContainsKey("LicenseUserID"), configDict("LicenseUserID").Trim(), "")
                Dim configUserIdExpanded = ExpandLicenseEnvironmentVariables(configUserIdRaw)
                Dim configUserIdIsEmpty = String.IsNullOrWhiteSpace(configUserIdExpanded)

                Dim isOfflineDomainLicense = IsOfflineDomainLicenseKey(configLicenseKey)

                If String.IsNullOrWhiteSpace(configProductId) AndAlso isOfflineDomainLicense Then
                    Dim extractedProductId As String = ""
                    If TryExtractOfflineDomainLicenseProductId(configLicenseKey, extractedProductId) Then
                        configProductId = extractedProductId
                    End If
                End If

                If configUserIdIsEmpty AndAlso isOfflineDomainLicense Then
                    configUserIdExpanded = GetOfflineDomainLicenseSyntheticUserId()
                    configUserIdIsEmpty = False
                End If

                Dim licenseClearAll = ParseBoolean(configDict, "LicenseClearAll", False)

                ' ═══════════════════════════════════════════════════════════════
                ' REGISTRY BACKUP RESTORE (before evaluating stored licenses)
                ' If My.Settings was wiped, restore from registry so the
                ' matching/preservation checks below can find the stored license.
                ' ═══════════════════════════════════════════════════════════════
                If Not HasStoredProLicense() AndAlso Not HasStoredPrivateLicense() Then
                    TryRestoreProLicenseFromRegistry()
                End If

                ' ═══════════════════════════════════════════════════════════════
                ' CHECK "DO NOT TOUCH" CONDITIONS FIRST (highest priority)
                ' If stored Pro license matches any DoNotTouch values, preserve it
                ' ═══════════════════════════════════════════════════════════════
                If HasStoredProLicense() Then
                    Dim storedProductId = My.Settings.License_ProductID
                    Dim storedUserId = My.Settings.License_UserID
                    Dim apiConfirmed = My.Settings.License_ApiConfirmed
                    Dim offlineStart = My.Settings.License_OfflineGraceStart

                    ' Check if license is valid (API confirmed or in offline grace)
                    Dim isValid = False
                    If apiConfirmed Then
                        If offlineStart > Date.MinValue Then
                            Dim offlineDays = CInt((Date.Now.Date - offlineStart.Date).TotalDays)
                            isValid = (offlineDays <= OfflineGracePeriodDays)
                        Else
                            isValid = True
                        End If
                    End If

                    If isValid Then
                        ' Check LicenseDoNotTouchProductID
                        If configDict.ContainsKey("LicenseDoNotTouchProductID") Then
                            Dim doNotTouchProductIds = configDict("LicenseDoNotTouchProductID").Split(","c)
                            For Each productIdValue In doNotTouchProductIds
                                Dim trimmedProductId = productIdValue.Trim()
                                If Not String.IsNullOrEmpty(trimmedProductId) AndAlso
                                   storedProductId.Equals(trimmedProductId, StringComparison.OrdinalIgnoreCase) Then
                                    LogLicenseEvent("Config Pro", $"DoNotTouchProductID matched '{trimmedProductId}' - preserving existing license")
                                    Return ProcessStoredProLicense(context)
                                End If
                            Next
                        End If

                        ' Check LicenseDoNotTouchUserID
                        If configDict.ContainsKey("LicenseDoNotTouchUserID") Then
                            Dim doNotTouchUserIds = configDict("LicenseDoNotTouchUserID").Split(","c)
                            For Each userIdValue In doNotTouchUserIds
                                Dim trimmedUserId = ExpandLicenseEnvironmentVariables(userIdValue.Trim())
                                If Not String.IsNullOrEmpty(trimmedUserId) AndAlso
                                   storedUserId.Equals(trimmedUserId, StringComparison.OrdinalIgnoreCase) Then
                                    LogLicenseEvent("Config Pro", $"DoNotTouchUserID matched '{trimmedUserId}' - preserving existing license")
                                    Return ProcessStoredProLicense(context)
                                End If
                            Next
                        End If
                    End If
                End If

                If Not licenseClearAll Then
                    ' ═══════════════════════════════════════════════════════════════
                    ' LicenseClearAll = False
                    ' Only apply config if NO valid Private or Pro license exists
                    ' ═══════════════════════════════════════════════════════════════

                    ' Check for existing valid Pro license
                    If HasStoredProLicense() Then
                        Dim productName = My.Settings.License_ProductName
                        Dim apiConfirmed = My.Settings.License_ApiConfirmed
                        Dim offlineStart = My.Settings.License_OfflineGraceStart

                        ' Check if license is valid (API confirmed or in offline grace)
                        Dim isValid = False

                        If apiConfirmed Then
                            ' Check if in offline grace period
                            If offlineStart > Date.MinValue Then
                                Dim offlineDays = CInt((Date.Now.Date - offlineStart.Date).TotalDays)
                                isValid = (offlineDays <= OfflineGracePeriodDays)
                            Else
                                ' API confirmed and not in offline grace - consider valid
                                isValid = True
                            End If
                        End If

                        If isValid Then
                            LogLicenseEvent("Config Pro", "LicenseClearAll=False - existing valid Pro license found, preserving")
                            Return ProcessStoredProLicense(context)
                        End If
                    End If

                    ' Check for existing valid Private license
                    If HasStoredPrivateLicense() Then
                        Dim confirmedOn = My.Settings.License_PrivateConfirmedOn
                        Dim validUntil = My.Settings.License_ValidUntil

                        If confirmedOn > Date.MinValue AndAlso validUntil > Date.Now Then
                            LogLicenseEvent("Config Pro", "LicenseClearAll=False - existing valid Private license found, preserving")
                            _currentLicenseState = LicenseState.PrivateActive
                            LicenseStatus = "Private License"
                            LicenseFromConfig = False
                            Return True
                        End If
                    End If

                    ' No valid license exists - proceed with config-based activation
                    LogLicenseEvent("Config Pro", "LicenseClearAll=False - no valid license exists, applying config")

                Else
                    ' ═══════════════════════════════════════════════════════════════
                    ' LicenseClearAll = True
                    ' Only apply config if NO matching Pro license exists
                    ' ═══════════════════════════════════════════════════════════════

                    If HasStoredProLicense() Then
                        Dim storedProductId = My.Settings.License_ProductID
                        Dim storedLicenseKey = My.Settings.License_Key
                        Dim storedUserId = My.Settings.License_UserID
                        Dim apiConfirmed = My.Settings.License_ApiConfirmed
                        Dim lastCheck = My.Settings.License_LastCheck
                        Dim offlineStart = My.Settings.License_OfflineGraceStart

                        ' Check if LicenseKey and ProductID match
                        Dim keyAndProductMatch = storedLicenseKey.Equals(configLicenseKey, StringComparison.Ordinal) AndAlso
                                                 storedProductId.Equals(configProductId, StringComparison.Ordinal)

                        If keyAndProductMatch Then
                            ' Check matching conditions:
                            ' (a) Config LicenseUserID is empty, OR
                            ' (b) Expanded LicenseUserID matches stored UserID
                            Dim userIdMatches = configUserIdIsEmpty OrElse
                                                storedUserId.Equals(configUserIdExpanded, StringComparison.OrdinalIgnoreCase)

                            If userIdMatches Then
                                ' Check if license is valid (API confirmed or in offline grace)
                                Dim isValid = False

                                If apiConfirmed Then
                                    If offlineStart > Date.MinValue Then
                                        Dim offlineDays = CInt((Date.Now.Date - offlineStart.Date).TotalDays)
                                        isValid = (offlineDays <= OfflineGracePeriodDays)
                                    Else
                                        isValid = True
                                    End If
                                End If

                                If isValid Then
                                    ' Matching valid Pro license exists - preserve it
                                    LogLicenseEvent("Config Pro", $"LicenseClearAll=True - matching valid Pro license found (UserID empty={configUserIdIsEmpty}), preserving")

                                    ' Still need to verify via API based on check interval
                                    Dim shouldCheckApi = False
                                    If ProLicenseCheckIntervalDays = 0 Then
                                        shouldCheckApi = True
                                    Else
                                        Dim daysSinceCheck = If(lastCheck = Date.MinValue, Integer.MaxValue,
                                                        CInt((Date.Now.Date - lastCheck.Date).TotalDays))
                                        shouldCheckApi = (daysSinceCheck >= ProLicenseCheckIntervalDays)
                                    End If

                                    If Not shouldCheckApi Then
                                        ' Within check interval - use existing license
                                        _currentLicenseState = If(IsTestingProLicenseByProductId(storedProductId), LicenseState.TestingProActive, LicenseState.ProActive)
                                        LicenseStatus = My.Settings.License_ProductName
                                        LicenseFromConfig = True
                                        Return True
                                    End If

                                    ' Need to verify - check API
                                    Dim statusResponse = CallLicenseApi("status", storedProductId, storedLicenseKey, storedUserId)

                                    If statusResponse.Success AndAlso statusResponse.Activated Then
                                        RecordSuccessfulApiCheck()
                                        _currentLicenseState = If(IsTestingProLicenseByProductId(storedProductId), LicenseState.TestingProActive, LicenseState.ProActive)
                                        LicenseStatus = My.Settings.License_ProductName
                                        LicenseFromConfig = True
                                        Return True
                                    End If

                                    ' API check failed - could be network issue
                                    If statusResponse.RawJson Is Nothing OrElse statusResponse.RawJson.Length = 0 Then
                                        If offlineStart = Date.MinValue Then
                                            My.Settings.License_OfflineGraceStart = Date.Now.Date
                                            My.Settings.Save()
                                        End If
                                        Dim offlineDays = CInt((Date.Now.Date - My.Settings.License_OfflineGraceStart.Date).TotalDays)
                                        If offlineDays <= OfflineGracePeriodDays Then
                                            LogLicenseEvent("Config Pro", $"Matching Pro license - network issue, entering offline grace ({OfflineGracePeriodDays - offlineDays}d left)")
                                            _currentLicenseState = LicenseState.ProOfflineGrace
                                            LicenseStatus = My.Settings.License_ProductName
                                            LicenseFromConfig = True
                                            Return True
                                        End If
                                    End If

                                    ' License no longer valid - proceed with config activation
                                    LogLicenseEvent("Config Pro", "Matching Pro license no longer valid - will re-activate")
                                End If
                            End If
                        End If
                    End If

                    ' No matching valid Pro license - clear and proceed with config
                    LogLicenseEvent("Config Pro", "LicenseClearAll=True - no matching valid Pro license, clearing and applying config")
                    ClearStoredLicense(includeLegacy:=True)
                End If

                ' ═══════════════════════════════════════════════════════════════
                ' Proceed with config-based activation
                ' ═══════════════════════════════════════════════════════════════

                LicenseFromConfig = True

                Dim autoMode = ParseBoolean(configDict, "LicenseAutoMode", False)

                LogLicenseEvent("Config Pro", $"AutoMode={autoMode}, ProductID={configProductId}, UserID={If(configUserIdIsEmpty, "(empty)", configUserIdExpanded)}")

                ' Check if we have all required parameters for activation
                Dim missingParams As New List(Of String)
                If String.IsNullOrWhiteSpace(configProductId) Then missingParams.Add("Product ID")
                If configUserIdIsEmpty Then missingParams.Add("User ID")

                If missingParams.Count > 0 Then
                    If autoMode Then
                        ' Auto mode but missing params - show error and fail
                        ShowCustomMessageBox(
                            $"Automatic license activation cannot proceed. Missing: {String.Join(", ", missingParams)}." & vbCrLf & vbCrLf &
                            StandardSupportContactMessage & " You may want to set 'LicenseAutoMode = False' to allow manual license entry.",
                            $"{AN} - License Configuration Error")
                        Return False
                    Else
                        ' Show prefilled form for user to complete
                        Return ShowProLicenseEntryForm(context, configLicenseKey, configProductId, configUserIdExpanded, prefilled:=True)
                    End If
                End If

                ' All params present - check if already activated with matching credentials
                If IsLicenseActivatedWithCredentials(configProductId, configLicenseKey, configUserIdExpanded) Then
                    ' Already activated - verify it's still valid via API
                    LogLicenseEvent("Config Pro", "Already activated with matching credentials - verifying")
                    Return VerifyAndProcessProLicense(context, configProductId, configLicenseKey, configUserIdExpanded)
                End If

                ' Not activated or credentials changed - perform activation
                If autoMode Then
                    Return PerformAutoModeActivation(context, configProductId, configLicenseKey, configUserIdExpanded)
                Else
                    Return ShowProLicenseEntryForm(context, configLicenseKey, configProductId, configUserIdExpanded, prefilled:=True)
                End If

            Catch ex As Exception
                LogLicenseEvent("Config Pro Error", ex.Message, alwaysLog:=True)
                ShowCustomMessageBox(
                    $"License configuration error: {ex.Message}" & vbCrLf & vbCrLf &
                    StandardSupportContactMessage,
                    $"{AN} - License Error")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Performs automatic Pro license activation using the configured Product ID, License Key, and User ID.
        ''' Uses a status call first, then falls back to activation when required.
        ''' </summary>
        Private Shared Function PerformAutoModeActivation(context As ISharedContext, productId As String, licenseKey As String, userId As String) As Boolean
            Try
                Dim autoModeSilent As Boolean = False
                If _licenseConfigDict IsNot Nothing Then
                    autoModeSilent = ParseBoolean(_licenseConfigDict, "LicenseAutoModeSilent", False)
                End If

                ' Check if LicenseClearAll is set - if so, clear all licenses before proceeding
                If _licenseConfigDict IsNot Nothing AndAlso ParseBoolean(_licenseConfigDict, "LicenseClearAll", False) Then
                    LogLicenseEvent("Auto Activation", "LicenseClearAll=True - clearing all licenses including legacy")
                    ClearStoredLicense(includeLegacy:=True)
                End If
                LogLicenseEvent("Auto Activation", $"Checking status for {userId}")

                ' First, check if already activated via status call
                Dim statusResponse = CallLicenseApi("status", productId, licenseKey, userId)

                If statusResponse.Success AndAlso statusResponse.Activated Then
                    ' Already activated - no need to call activate API
                    Dim productTitle = statusResponse.ProductTitle
                    Dim totalActivations = statusResponse.TotalActivations
                    Dim totalPurchased = statusResponse.TotalActivationsPurchased
                    Dim activationsRemaining = statusResponse.ActivationsRemaining

                    ' Save the license
                    SaveProLicenseToSettings(productId, licenseKey, userId, productTitle, apiConfirmed:=True)
                    _currentLicenseState = If(IsTestingProLicenseByProductId(productId), LicenseState.TestingProActive, LicenseState.ProActive)

                    LogLicenseEvent("Auto Activation", $"Already activated - Product: {productTitle}, Activations: {totalActivations}/{totalPurchased}", alwaysLog:=True)

                    Dim successMsg As String = BuildConfigManagedLicenseSuccessMessage(
                        licenseKey,
                        productTitle,
                        userId,
                        totalActivations,
                        totalPurchased,
                        activationsRemaining,
                        alreadyActivated:=True)

                    If Not autoModeSilent AndAlso Not _restoredFromRegistryBackup Then
                        ShowCustomMessageBox(successMsg.ToString(), $"{AN} - License Active")
                    End If

                    Return True
                End If

                ' Not activated yet - attempt activation
                LogLicenseEvent("Auto Activation", $"Attempting activation for {userId}")

                Dim activateResponse = CallLicenseApi("activate", productId, licenseKey, userId)

                If activateResponse.Success AndAlso activateResponse.Activated Then
                    ' Activation successful - get the product title and activation counts via status call
                    Dim productTitle = activateResponse.ProductTitle
                    Dim totalActivations As Integer = activateResponse.TotalActivations
                    Dim totalPurchased As Integer = activateResponse.TotalActivationsPurchased
                    Dim activationsRemaining As Integer = activateResponse.ActivationsRemaining

                    ' Always do a status call to get accurate product title and activation counts
                    Try
                        Dim postActivateStatus = CallLicenseApi("status", productId, licenseKey, userId)
                        If postActivateStatus.Success Then
                            ' Update product title if we got a better one
                            If Not String.IsNullOrEmpty(postActivateStatus.ProductTitle) AndAlso
                               postActivateStatus.ProductTitle <> "(no product title available)" Then
                                productTitle = postActivateStatus.ProductTitle
                            End If
                            ' Update activation counts from status response
                            If postActivateStatus.TotalActivationsPurchased > 0 Then
                                totalPurchased = postActivateStatus.TotalActivationsPurchased
                                totalActivations = postActivateStatus.TotalActivations
                                activationsRemaining = postActivateStatus.ActivationsRemaining
                            End If
                        End If
                    Catch
                        ' Ignore status call errors - activation succeeded
                    End Try

                    ' Save the license with product title
                    SaveProLicenseToSettings(productId, licenseKey, userId, productTitle, apiConfirmed:=True)
                    _currentLicenseState = If(IsTestingProLicenseByProductId(productId), LicenseState.TestingProActive, LicenseState.ProActive)

                    LogLicenseEvent("Auto Activation", $"Success - Product: {productTitle}, Activations: {totalActivations}/{totalPurchased}", alwaysLog:=True)

                    Dim successMsg As String = BuildConfigManagedLicenseSuccessMessage(
                        licenseKey,
                        productTitle,
                        userId,
                        totalActivations,
                        totalPurchased,
                        activationsRemaining,
                        alreadyActivated:=False)

                    If Not autoModeSilent AndAlso Not _restoredFromRegistryBackup Then
                        ShowCustomMessageBox(successMsg.ToString(), $"{AN} - License Activated")
                    End If

                    Return True
                End If

                ' Activation failed - check status again to see why (might have been activated by race condition)
                Dim recheckResponse = CallLicenseApi("status", productId, licenseKey, userId)

                If recheckResponse.Success AndAlso recheckResponse.Activated Then
                    ' Actually is activated (race condition or API quirk)
                    SaveProLicenseToSettings(productId, licenseKey, userId, recheckResponse.ProductTitle, apiConfirmed:=True)
                    _currentLicenseState = If(IsTestingProLicenseByProductId(productId), LicenseState.TestingProActive, LicenseState.ProActive)

                    LogLicenseEvent("Auto Activation", $"Verified activated after failed activate call - Product: {recheckResponse.ProductTitle}", alwaysLog:=True)

                    Dim successMsg As String = BuildConfigManagedLicenseSuccessMessage(
                        licenseKey,
                        recheckResponse.ProductTitle,
                        userId,
                        recheckResponse.TotalActivations,
                        recheckResponse.TotalActivationsPurchased,
                        recheckResponse.ActivationsRemaining,
                        alreadyActivated:=True)

                    If Not autoModeSilent AndAlso Not _restoredFromRegistryBackup Then
                        ShowCustomMessageBox(successMsg.ToString(), $"{AN} - License Active")
                    End If

                    Return True
                ElseIf recheckResponse.Success AndAlso recheckResponse.ActivationsRemaining <= 0 Then
                    ' No slots remaining
                    LogLicenseEvent("Auto Activation Failed", $"No activation slots remaining. Used: {recheckResponse.TotalActivations}/{recheckResponse.TotalActivationsPurchased}", alwaysLog:=True)
                    ShowCustomMessageBox(
                        $"Automatic license activation failed: No activation slots remaining." & vbCrLf & vbCrLf &
                        $"Product: {StripHtmlFromLicenseMessage(recheckResponse.ProductTitle)}" & vbCrLf &
                        $"Activations: {recheckResponse.TotalActivations} of {recheckResponse.TotalActivationsPurchased} used" & vbCrLf & vbCrLf &
                        "You will first need to deactivate another User ID or obtain additional licenses." & vbCrLf & vbCrLf &
                        StandardSupportContactMessage,
                        $"{AN} - License Activation Failed")
                    Return False
                End If

                ' Genuine failure
                LogLicenseEvent("Auto Activation Failed", activateResponse.ErrorMessage, alwaysLog:=True)
                ShowCustomMessageBox(
                    $"Automatic license activation failed: {StripHtmlFromLicenseMessage(activateResponse.ErrorMessage)}" & vbCrLf & vbCrLf &
                    $"Product ID: {productId}" & vbCrLf &
                    $"User ID: {userId}" & vbCrLf & vbCrLf &
                    StandardSupportContactMessage,
                    $"{AN} - License Activation Failed")
                Return False

            Catch ex As Exception
                LogLicenseEvent("Auto Activation Error", ex.Message, alwaysLog:=True)
                ShowCustomMessageBox(
                    $"Automatic license activation error: {ex.Message}" & vbCrLf & vbCrLf &
                    StandardSupportContactMessage,
                    $"{AN} - License Activation Error")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Returns whether the stored Pro license matches the supplied credentials and is API-confirmed.
        ''' </summary>
        Private Shared Function IsLicenseActivatedWithCredentials(productId As String, licenseKey As String, userId As String) As Boolean
            Try
                If Not HasStoredProLicense() Then Return False

                Return My.Settings.License_ProductID.Equals(productId, StringComparison.Ordinal) AndAlso
                       My.Settings.License_Key.Equals(licenseKey, StringComparison.Ordinal) AndAlso
                       My.Settings.License_UserID.Equals(userId, StringComparison.OrdinalIgnoreCase) AndAlso
                       My.Settings.License_ActivatedOn > Date.MinValue AndAlso
                       My.Settings.License_ApiConfirmed
            Catch
                Return False
            End Try
        End Function


#End Region

#Region "Stored Pro License Processing"

        ''' <summary>
        ''' Returns whether stored settings indicate a Pro license.
        ''' </summary>
        Private Shared Function HasStoredProLicense() As Boolean
            Try
                Dim licenseType = My.Settings.License_Type
                Return Not String.IsNullOrEmpty(licenseType) AndAlso
                       licenseType.Equals("Pro", StringComparison.OrdinalIgnoreCase)
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Returns whether the stored Pro license has previously been confirmed via the API.
        ''' </summary>
        Private Shared Function HasApiConfirmedProLicense() As Boolean
            Try
                Return HasStoredProLicense() AndAlso My.Settings.License_ApiConfirmed
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Processes a stored Pro license, including periodic API verification and offline grace handling.
        ''' </summary>
        Private Shared Function ProcessStoredProLicense(context As ISharedContext) As Boolean
            Try
                Dim productId = My.Settings.License_ProductID
                Dim licenseKey = My.Settings.License_Key
                Dim userId = My.Settings.License_UserID
                Dim productName = My.Settings.License_ProductName
                Dim lastCheck = My.Settings.License_LastCheck
                Dim apiConfirmed = My.Settings.License_ApiConfirmed

                LogLicenseEvent("Stored Pro", $"ProductID={productId}, ApiConfirmed={apiConfirmed}, LastCheck={lastCheck:d}")

                If IsOfflineDomainLicenseKey(licenseKey) Then
                    LogLicenseEvent("Stored Pro", "Offline-domain license path")
                    Return VerifyAndProcessProLicense(context, productId, licenseKey, userId)
                End If

                ' Check if this is a Testing Pro License - requires compliance confirmation
                If IsTestingProLicenseByProductId(productId) Then
                    If IsComplianceCheckDue(TestingProComplianceIntervalStartups) Then
                        If Not ShowTestingProComplianceConfirmation() Then
                            Return False
                        End If
                    End If
                End If

                ' Determine if periodic API check is needed
                ' If ProLicenseCheckIntervalDays = 0, always check on every startup
                ' Otherwise, check only if the interval has elapsed
                Dim shouldCheckApi As Boolean = False

                If ProLicenseCheckIntervalDays = 0 Then
                    ' Always check when interval is 0
                    shouldCheckApi = True
                Else
                    ' Check based on interval
                    Dim daysSinceCheck = If(lastCheck = Date.MinValue, Integer.MaxValue,
                                    CInt((Date.Now.Date - lastCheck.Date).TotalDays))
                    shouldCheckApi = (daysSinceCheck >= ProLicenseCheckIntervalDays)
                End If

                If shouldCheckApi Then
                    ' Time for API verification
                    Return VerifyAndProcessProLicense(context, productId, licenseKey, userId)
                End If

                ' Check if we're in offline grace period
                Dim offlineStart = My.Settings.License_OfflineGraceStart
                If offlineStart > Date.MinValue Then
                    Dim offlineDays = CInt((Date.Now.Date - offlineStart.Date).TotalDays)
                    If offlineDays > OfflineGracePeriodDays Then
                        ' Offline grace expired
                        _currentLicenseState = LicenseState.ProExpired
                        ShowOfflineGraceExpiredMessage()
                        Return False
                    End If

                    ' Still in offline grace - show warning periodically
                    _currentLicenseState = LicenseState.ProOfflineGrace
                    If IsComplianceCheckDue(OfflineGraceWarningIntervalStartups) Then
                        ShowOfflineGraceWarning(OfflineGracePeriodDays - offlineDays)
                    End If
                    LicenseStatus = productName
                    Return True
                End If

                ' License is valid
                _currentLicenseState = If(IsTestingProLicenseByProductId(productId), LicenseState.TestingProActive, LicenseState.ProActive)
                LicenseStatus = productName
                Return True

            Catch ex As Exception
                LogLicenseEvent("Stored Pro Error", ex.Message)
                Return False ' Fail closed
            End Try
        End Function

        ''' <summary>
        ''' Verifies a Pro license via the license API and updates local state accordingly.
        ''' </summary>
        Private Shared Function VerifyAndProcessProLicense(context As ISharedContext, productId As String, licenseKey As String, userId As String) As Boolean
            Try
                LogLicenseEvent("API Verify", "Starting verification")

                If IsOfflineDomainLicenseKey(licenseKey) Then
                    Dim offlineResponse As LicenseApiResponse = Nothing

                    If TryCreateOfflineDomainLicenseResponse("status", productId, licenseKey, userId, offlineResponse) Then
                        If offlineResponse.Success AndAlso offlineResponse.Activated Then
                            RecordSuccessfulApiCheck()

                            Dim resolvedProductName = If(String.IsNullOrWhiteSpace(My.Settings.License_ProductName),
                                                         offlineResponse.ProductTitle,
                                                         My.Settings.License_ProductName)

                            If Not resolvedProductName.Equals(My.Settings.License_ProductName, StringComparison.Ordinal) Then
                                My.Settings.License_ProductName = resolvedProductName
                                My.Settings.Save()
                            End If

                            _currentLicenseState = If(IsTestingProLicenseByProductId(productId), LicenseState.TestingProActive, LicenseState.ProActive)
                            LicenseStatus = resolvedProductName

                            LogLicenseEvent("Offline Domain Verify", "License valid")
                            Return True
                        End If

                        Return HandleInvalidLicenseCredentials(productId, licenseKey, userId, offlineResponse.ErrorMessage)
                    End If
                End If

                Dim response = CallLicenseApi("status", productId, licenseKey, userId)

                If response.Success Then
                    If response.Activated OrElse response.StatusCheck.Equals("active", StringComparison.OrdinalIgnoreCase) Then
                        ' License valid - record successful check
                        RecordSuccessfulApiCheck()
                        _currentLicenseState = If(IsTestingProLicenseByProductId(productId), LicenseState.TestingProActive, LicenseState.ProActive)
                        LicenseStatus = My.Settings.License_ProductName
                        LogLicenseEvent("API Verify", "License valid")
                        Return True
                    Else
                        ' License revoked or inactive
                        Return HandleProLicenseRevocation()
                    End If
                End If

                ' API call failed - determine if it's a network issue or an API rejection
                ' If the API returned a response (even an error response), it's not a connectivity issue
                If response.RawJson IsNot Nothing AndAlso response.RawJson.Length > 0 Then
                    ' API responded but indicated the license is invalid/not found
                    ' This is NOT an offline situation - the credentials are wrong or outdated
                    Return HandleInvalidLicenseCredentials(productId, licenseKey, userId, response.ErrorMessage)
                End If

                ' No response received - this is likely a network/connectivity issue
                Return HandleApiCheckFailure(response.ErrorMessage)

            Catch ex As Exception
                Return HandleApiCheckFailure(ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Handles the case where the API responded, but rejected the license credentials.
        ''' </summary>
        Private Shared Function HandleInvalidLicenseCredentials(productId As String, licenseKey As String, userId As String, errorMessage As String) As Boolean
            Try
                LogLicenseEvent("Invalid Credentials", $"API rejected license: {errorMessage}", alwaysLog:=True)

                _currentLicenseState = LicenseState.ProExpired

                ' Build a message showing the credentials that were checked
                Dim credentialsInfo As String = $"The license information could not be verified by the license server." & vbCrLf & vbCrLf &
                    "Credentials checked:" & vbCrLf &
                    $"  Product ID: {productId}" & vbCrLf &
                    $"  License Key: {TruncateLicenseKey(licenseKey)}" & vbCrLf &
                    $"  User ID: {userId}" & vbCrLf & vbCrLf &
                    $"Server response: {StripHtmlFromLicenseMessage(errorMessage)}" & vbCrLf & vbCrLf &
                    "This may indicate:" & vbCrLf &
                    "  • The license key is incorrect or has a typo" & vbCrLf &
                    "  • The Product ID does not match the license" & vbCrLf &
                    "  • The license has been deactivated or expired" & vbCrLf &
                    "  • The User ID is not authorized for this license" & vbCrLf & vbCrLf &
                    StandardLicenseContactMessage & vbCrLf & vbCrLf &
                    "Would you like to update your license information?"

                Dim result = ShowCustomYesNoBox(
                    credentialsInfo,
                    "Update License",
                    "Cancel",
                    $"{AN} - License Verification Failed")

                If result = 1 Then
                    ' User wants to update license - show the license entry form with current values
                    Return ShowProLicenseEntryForm(_licenseContext, licenseKey, productId, userId, prefilled:=False)
                End If

                ' User cancelled - clear stored license and fail
                ClearStoredLicense()
                Return False

            Catch ex As Exception
                LogLicenseEvent("Invalid Credentials Error", ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Truncates a license key for display (first 8 and last 4 characters).
        ''' </summary>
        Private Shared Function TruncateLicenseKey(licenseKey As String) As String
            If String.IsNullOrEmpty(licenseKey) Then Return "(empty)"
            If licenseKey.Length <= 16 Then Return licenseKey
            Return licenseKey.Substring(0, 8) & "..." & licenseKey.Substring(licenseKey.Length - 4)
        End Function


        ''' <summary>
        ''' Handles the case where the API verified the request but indicated the license is inactive/revoked.
        ''' </summary>
        Private Shared Function HandleProLicenseRevocation() As Boolean
            Try
                LogLicenseEvent("Revoked", "License no longer active", alwaysLog:=True)

                ' Check if we have a previously API-confirmed license
                If Not HasApiConfirmedProLicense() Then
                    ' Never confirmed - fail immediately but offer to manage license
                    _currentLicenseState = LicenseState.ProExpired

                    Dim result = ShowCustomYesNoBox(
                        "Your license could not be verified and has never been successfully activated." & vbCrLf & vbCrLf &
                        StandardLicenseContactMessage & vbCrLf & vbCrLf &
                        "Would you like to open the license management dialog to resolve this issue?",
                        "Manage License",
                        "Cancel",
                        $"{AN} - License Invalid")

                    If result = 1 Then
                        ClearStoredLicense()
                        Return ShowProLicenseEntryForm(_licenseContext, "", "", "", prefilled:=False)
                    End If

                    ClearStoredLicense()
                    Return False
                End If

                ' Previously confirmed - apply revocation grace period
                Dim graceStart = My.Settings.License_GracePeriodStart
                If graceStart = Date.MinValue Then
                    My.Settings.License_GracePeriodStart = Date.Now.Date
                    My.Settings.Save()
                    graceStart = Date.Now.Date
                End If

                Dim graceDays = CInt((Date.Now.Date - graceStart.Date).TotalDays)

                If graceDays <= RevocationGracePeriodDays Then
                    ' Within grace period - offer to manage license
                    _currentLicenseState = LicenseState.ProExpired
                    Dim remainingDays = RevocationGracePeriodDays - graceDays

                    Dim result = ShowCustomYesNoBox(
                        $"Your license has been deactivated or revoked." & vbCrLf & vbCrLf &
                        $"You have {remainingDays} day(s) remaining to resolve this issue." & vbCrLf & vbCrLf &
                        StandardLicenseContactMessage & vbCrLf & vbCrLf &
                        "Would you like to open the license management dialog to resolve this issue?",
                        "Manage License",
                        "Continue",
                        $"{AN} - License Revoked")

                    If result = 1 Then
                        ' User wants to manage license
                        Return ShowProLicenseEntryForm(_licenseContext,
                                                       My.Settings.License_Key,
                                                       My.Settings.License_ProductID,
                                                       My.Settings.License_UserID,
                                                       prefilled:=False)
                    End If

                    Return True ' Allow continuation during grace
                End If

                ' Grace period expired - offer to manage license
                _currentLicenseState = LicenseState.ProExpired

                Dim expiredResult = ShowCustomYesNoBox(
                    "Your license has been revoked and the grace period has expired." & vbCrLf & vbCrLf &
                    StandardLicenseContactMessage & vbCrLf & vbCrLf &
                    "Would you like to open the license management dialog to enter new license credentials?",
                    "Manage License",
                    "Cancel",
                    $"{AN} - License Expired")

                If expiredResult = 1 Then
                    ClearStoredLicense()
                    Return ShowProLicenseEntryForm(_licenseContext, "", "", "", prefilled:=False)
                End If

                ClearStoredLicense()
                Return False

            Catch ex As Exception
                LogLicenseEvent("Revocation Error", ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Handles API check failures such as timeouts and connectivity issues.
        ''' </summary>
        Private Shared Function HandleApiCheckFailure(errorMessage As String) As Boolean
            Try
                LogLicenseEvent("API Check Failed", errorMessage)

                ' If never API-confirmed, fail immediately
                If Not HasApiConfirmedProLicense() Then
                    _currentLicenseState = LicenseState.ProExpired
                    ShowCustomMessageBox(
                        "Unable to verify your license and no previous successful activation exists." & vbCrLf & vbCrLf &
                        "Please ensure internet connectivity and try again." & vbCrLf & vbCrLf &
                        StandardSupportContactMessage,
                        $"{AN} - License Verification Failed")
                    Return False
                End If

                ' Previously confirmed - apply offline grace period
                Dim offlineStart = My.Settings.License_OfflineGraceStart
                If offlineStart = Date.MinValue Then
                    My.Settings.License_OfflineGraceStart = Date.Now.Date
                    My.Settings.Save()
                    offlineStart = Date.Now.Date
                End If

                Dim offlineDays = CInt((Date.Now.Date - offlineStart.Date).TotalDays)

                If offlineDays <= OfflineGracePeriodDays Then
                    ' Within offline grace
                    _currentLicenseState = LicenseState.ProOfflineGrace
                    Dim remainingDays = OfflineGracePeriodDays - offlineDays
                    ShowOfflineGraceWarning(remainingDays)
                    Return True
                End If

                ' Offline grace expired
                _currentLicenseState = LicenseState.ProExpired
                ShowOfflineGraceExpiredMessage()
                Return False

            Catch ex As Exception
                LogLicenseEvent("API Failure Handler Error", ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Records a successful API check and resets grace-period tracking.
        ''' </summary>
        Private Shared Sub RecordSuccessfulApiCheck()
            Try
                My.Settings.License_LastCheck = Date.Now.Date
                My.Settings.License_ApiConfirmed = True
                My.Settings.License_OfflineGraceStart = Date.MinValue
                My.Settings.License_GracePeriodStart = Date.MinValue
                My.Settings.Save()
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' Shows an offline grace period warning with remaining days.
        ''' </summary>
        Private Shared Sub ShowOfflineGraceWarning(remainingDays As Integer)
            ShowCustomMessageBox(
                $"Unable to verify your license (no internet connection)." & vbCrLf & vbCrLf &
                $"You have {remainingDays} day(s) remaining to reconnect." & vbCrLf & vbCrLf &
                "Please ensure internet connectivity to continue using " & AN & ".",
                $"{AN} - License Verification")
        End Sub

        ''' <summary>
        ''' Shows an offline grace expired message.
        ''' </summary>
        Private Shared Sub ShowOfflineGraceExpiredMessage()
            ShowCustomMessageBox(
                $"License verification failed: Unable to connect to the license server." & vbCrLf & vbCrLf &
                $"The offline grace period of {OfflineGracePeriodDays} days has expired. To use {AN} again, restore internet connectivity." & vbCrLf & vbCrLf &
                StandardSupportContactMessage,
                $"{AN} - License Verification Failed")
        End Sub

#End Region

#Region "Legacy Pro License Processing"

        ''' <summary>
        ''' Returns whether legacy Pro license settings are present (non-Private) and no new-regime license is set.
        ''' </summary>
        Private Shared Function HasLegacyProLicense() As Boolean
            Try
                ' Must not have new system license already activated
                Dim newLicenseType = My.Settings.License_Type
                If Not String.IsNullOrEmpty(newLicenseType) Then
                    Return False
                End If

                ' Check legacy fields - these are preserved until migration succeeds
                Dim legacyStatus = My.Settings.LicenseStatus
                Dim legacyTill = My.Settings.LicensedTill

                If String.IsNullOrEmpty(legacyStatus) Then Return False
                If legacyTill = Date.MinValue Then Return False

                ' Must NOT be a Private License (those are cleared, not migrated)
                If legacyStatus.Equals("Private License", StringComparison.OrdinalIgnoreCase) Then
                    Return False
                End If

                ' Has legacy Pro license data
                Return True

            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Returns whether a legacy migration flow has been started (tracked in settings).
        ''' </summary>
        Private Shared Function IsLegacyMigrationStarted() As Boolean
            Try
                Return My.Settings.License_LegacyMigrationStarted > Date.MinValue
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Processes a legacy Pro license, including regime end enforcement and migration prompting.
        ''' </summary>
        Private Shared Function ProcessLegacyProLicense(context As ISharedContext) As Boolean
            Try
                Dim legacyStatus = My.Settings.LicenseStatus
                Dim legacyTill = My.Settings.LicensedTill
                Dim legacyUsers = My.Settings.LicenseUsers

                LogLicenseEvent("Legacy Pro", $"Status={legacyStatus}, ValidTill={legacyTill:d}")

                ' Update global variables for display
                LicenseStatus = legacyStatus
                LicensedTill = legacyTill
                LicenseUsers = legacyUsers

                ' Check if legacy license itself is expired
                If Date.Now > legacyTill Then
                    _currentLicenseState = LicenseState.LegacyExpired

                    ' Only show expiration message once (first time we detect expiration)
                    ' Use License_LegacyMigrationStarted to track that user has been notified
                    If Not IsLegacyMigrationStarted() Then
                        ShowCustomMessageBox(
                            $"Your legacy license ({legacyStatus}) expired on {legacyTill:d}." & vbCrLf & vbCrLf &
                            StandardLicenseContactMessage,
                            $"{AN} - License Expired")
                    End If

                    Return HandleLegacyProMigration(context, legacyStatus, legacyTill)
                End If

                ' Check if we're past the legacy regime end date
                If Date.Now >= LegacyRegimeEndDate Then
                    Return HandleLegacyProMigration(context, legacyStatus, legacyTill)
                End If

                ' Before regime end date - legacy license is valid
                _currentLicenseState = LicenseState.LegacyProActive
                LogLicenseEvent("Legacy Pro", "Valid (before regime end)")

                ' Show expiry warnings for legacy license (on warning days, throttled by interval)
                If Not LicenseNoWarning AndAlso legacyTill > Date.Now AndAlso legacyTill < Date.MaxValue Then
                    ' Calculate days until legacy license expires
                    Dim daysUntilExpiry = CInt((legacyTill.Date - Date.Now.Date).TotalDays)

                    ' Check against warning thresholds
                    For Each warningDay In LicenseWarningDays
                        If daysUntilExpiry = warningDay Then
                            ' Check if we should show warning based on startup interval
                            If ShouldShowLicenseWarning() Then
                                Dim result = ShowCustomYesNoBox(
                                    $"Your license ({legacyStatus}) will EXPIRE in {daysUntilExpiry} day(s) on {legacyTill:d}." & vbCrLf & vbCrLf &
                                    "Upgrade to the new licensing scheme to continue using the software." & vbCrLf & vbCrLf &
                                    StandardLicenseContactMessage & vbCrLf & vbCrLf &
                                    "Would you like to enter your new license now?",
                                    "Enter License",
                                    "Later",
                                    $"{AN} - License Warning")

                                RecordLicenseWarningShown()

                                If result = 1 Then
                                    ' User wants to enter license now
                                    Dim answer = ShowLicenseTypeSelectionDialog(context)
                                    Return True ' License is still valid, continue regardless
                                End If
                            End If
                            Exit For
                        End If
                    Next
                End If

                Return True

            Catch ex As Exception
                LogLicenseEvent("Legacy Pro Error", ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Enforces migration requirements for legacy Pro licenses after `LegacyRegimeEndDate`.
        ''' Legacy license data remains stored until a new license is successfully activated.
        ''' </summary>
        Private Shared Function HandleLegacyProMigration(context As ISharedContext, legacyStatus As String, legacyTill As Date) As Boolean
            Try
                ' Calculate days since regime ended
                Dim daysSinceRegimeEnd = CInt((Date.Now.Date - LegacyRegimeEndDate.Date).TotalDays)

                ' Check if within migration grace period
                If daysSinceRegimeEnd <= LegacyMigrationGracePeriodDays Then
                    ' Within grace - allow usage but prompt for migration
                    _currentLicenseState = LicenseState.LegacyProMigrationRequired
                    Dim remainingDays = LegacyMigrationGracePeriodDays - daysSinceRegimeEnd

                    ' Show migration prompt:
                    ' - Immediately on first encounter (migration not yet started)
                    ' - Every X startups thereafter
                    If Not IsLegacyMigrationStarted() OrElse IsComplianceCheckDue(LegacyMigrationPromptIntervalStartups) Then
                        ' Mark that we've shown the migration prompt (first time)
                        If Not IsLegacyMigrationStarted() Then
                            My.Settings.License_LegacyMigrationStarted = Date.Now.Date
                            My.Settings.Save()
                        End If

                        ShowLegacyMigrationPrompt(context, legacyStatus, remainingDays)
                    End If

                    Return True ' Allow usage during grace period
                End If

                ' Grace period expired - migration is mandatory
                _currentLicenseState = LicenseState.LegacyProMigrationRequired

                ' Always show the "grace period expired" message
                ShowCustomMessageBox(
                    $"Your legacy license ({legacyStatus}) requires migration to the new licensing system." & vbCrLf & vbCrLf &
                    "The grace period for migration has expired. You must enter a new license to continue." & vbCrLf & vbCrLf &
                    StandardLicenseContactMessage,
                    $"{AN} - License Migration Required")

                ' Show license selection - returns False if user cancels (blocking app)
                Return ShowLicenseTypeSelectionDialog(context)

            Catch ex As Exception
                LogLicenseEvent("Legacy Migration Error", ex.Message)
                Return False
            End Try
        End Function


        ''' <summary>
        ''' Shows a prompt explaining the legacy migration requirement and optionally opens the license dialog.
        ''' </summary>
        Private Shared Sub ShowLegacyMigrationPrompt(context As ISharedContext, legacyStatus As String, remainingDays As Integer)
            Try
                Dim msg = $"Your legacy license ({legacyStatus}) must be migrated to the new licensing system." & vbCrLf & vbCrLf &
                          $"You have {remainingDays} day(s) remaining to complete migration." & vbCrLf & vbCrLf &
                          "The new license system requires you to enter a license key, a product ID and user ID." & vbCrLf & vbCrLf &
                          StandardLicenseContactMessage & vbCrLf & vbCrLf &
                          "Would you like to enter your new license now?"

                Dim result = ShowCustomYesNoBox(
                    msg,
                    "Enter License",
                    "Later",
                    $"{AN} - License Migration Required")

                If result = 1 Then
                    ' Show license dialog - legacy license cleared only on successful activation
                    ShowLicenseTypeSelectionDialog(context)
                End If

            Catch ex As Exception
                LogLicenseEvent("Migration Prompt Error", ex.Message)
            End Try
        End Sub


        ''' <summary>
        ''' Clears legacy license fields stored in `My.Settings` and resets corresponding globals.
        ''' </summary>
        Private Shared Sub ClearLegacyLicense()
            Try
                My.Settings.LicenseStatus = ""
                My.Settings.LicensedTill = Date.MinValue
                My.Settings.LicenseUsers = 0
                My.Settings.Save()

                LicenseStatus = ""
                LicensedTill = Date.MinValue
                LicenseUsers = 0

            Catch ex As Exception
                LogLicenseEvent("Clear Legacy Error", ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Clears an existing legacy Private license from legacy storage (legacy Private licenses are not migrated).
        ''' </summary>
        Private Shared Sub ClearLegacyPrivateLicenseIfExists()
            Try
                Dim legacyStatus = My.Settings.LicenseStatus
                If Not String.IsNullOrEmpty(legacyStatus) AndAlso
                   legacyStatus.Equals("Private License", StringComparison.OrdinalIgnoreCase) Then
                    LogLicenseEvent("Legacy Private", "Clearing legacy private license")
                    ClearLegacyLicense()
                End If
            Catch
            End Try
        End Sub

#End Region

#Region "License Storage Helpers"

        ''' <summary>
        ''' Saves Pro license data to `My.Settings` and updates related globals.
        ''' </summary>
        Private Shared Sub SaveProLicenseToSettings(productId As String, licenseKey As String, userId As String, productName As String, Optional apiConfirmed As Boolean = False)
            Try
                My.Settings.License_Type = "Pro"
                My.Settings.License_ProductID = productId
                My.Settings.License_Key = licenseKey
                My.Settings.License_UserID = userId
                My.Settings.License_ProductName = productName
                My.Settings.License_ActivatedOn = Date.Now.Date
                My.Settings.License_LastCheck = Date.Now.Date
                My.Settings.License_ApiConfirmed = apiConfirmed
                My.Settings.License_OfflineGraceStart = Date.MinValue
                My.Settings.License_GracePeriodStart = Date.MinValue
                My.Settings.License_State = LicenseState.ProActive.ToString()
                My.Settings.Save()

                ' Update global variables
                LicenseStatus = productName
                _currentLicenseState = If(IsTestingProLicenseByProductId(productId), LicenseState.TestingProActive, LicenseState.ProActive)

                ' Clear any legacy license data
                ClearLegacyLicense()

                ' Backup to registry in case user.config gets deleted
                BackupProLicenseToRegistry(productId, licenseKey, userId, productName, apiConfirmed)

            Catch ex As Exception
                LogLicenseEvent("Save Pro Error", ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Clears stored license data in `My.Settings`.
        ''' </summary>
        ''' <param name="includeLegacy">
        ''' If <see langword="True"/>, also clears legacy storage fields (`LicenseStatus`, `LicensedTill`, `LicenseUsers`).
        ''' </param>
        Private Shared Sub ClearStoredLicense(Optional includeLegacy As Boolean = False)
            Try
                ' Clear new license system settings
                My.Settings.License_Type = ""
                My.Settings.License_ProductID = ""
                My.Settings.License_Key = ""
                My.Settings.License_UserID = ""
                My.Settings.License_ProductName = ""
                My.Settings.License_ActivatedOn = Date.MinValue
                My.Settings.License_ValidUntil = Date.MinValue
                My.Settings.License_LastCheck = Date.MinValue
                My.Settings.License_ApiConfirmed = False
                My.Settings.License_PrivateConfirmedOn = Date.MinValue
                My.Settings.License_PrivateVersion = ""
                My.Settings.License_PrivateDismissCount = 0
                My.Settings.License_OfflineGraceStart = Date.MinValue
                My.Settings.License_GracePeriodStart = Date.MinValue
                My.Settings.License_State = ""
                My.Settings.License_AutoActivationWarningShown = False
                My.Settings.License_LegacyMigrationStarted = Date.MinValue
                My.Settings.License_LastMigrationPrompt = Date.MinValue

                If includeLegacy Then
                    ' Only clear legacy when explicitly requested (after successful new license activation)
                    My.Settings.LicenseStatus = ""
                    My.Settings.LicensedTill = Date.MinValue
                    My.Settings.LicenseUsers = 0
                    LicensedTill = Date.MinValue
                    LicenseUsers = 0
                End If

                My.Settings.Save()

                ' Reset global variables
                LicenseStatus = ""
                _currentLicenseState = LicenseState.None

                ' Clear registry backup as well
                ClearLicenseRegistryBackup()

            Catch ex As Exception
                LogLicenseEvent("Clear License Error", ex.Message)
            End Try
        End Sub

#End Region

#Region "Logging"

        ''' <summary>
        ''' Logs a license event to the update log.
        ''' </summary>
        ''' <param name="eventType">Short event category label.</param>
        ''' <param name="details">Optional event details; line breaks are indented.</param>
        ''' <param name="alwaysLog">Controls whether the event is recorded regardless of other conditions.</param>
        Private Shared Sub LogLicenseEvent(eventType As String, details As String, Optional alwaysLog As Boolean = False)
            Try
                Dim message As String = $"[License] [{eventType}]"
                If Not String.IsNullOrWhiteSpace(details) Then
                    message &= vbCrLf & "  " & details.Replace(vbCrLf, vbCrLf & "  ")
                End If
                UpdateHandler.WriteUpdateLog(message)
            Catch ex As Exception
                Debug.WriteLine($"Failed to log license event: {ex.Message}")
            End Try
        End Sub

#End Region

#Region "License Status Display"

        ''' <summary>
        ''' Shows a dialog with license status information based on the current license state.
        ''' </summary>
        Public Shared Sub ShowLicenseStatusDialog()
            Try
                Dim sb As New StringBuilder()
                Dim canUpdateLicense As Boolean = True

                Select Case _currentLicenseState
                    Case LicenseState.None
                        sb.AppendLine("No license is currently configured.")
                        sb.AppendLine()
                        sb.AppendLine($"Please configure a license to use {AN}.")
                        sb.AppendLine($"Visit {AN4} for licensing options.")

                    Case LicenseState.PrivateActive, LicenseState.PrivateReconfirmNeeded
                        sb.AppendLine("License Type: Private License (Free)")
                        sb.AppendLine()
                        Dim validUntil = My.Settings.License_ValidUntil
                        If validUntil > Date.MinValue AndAlso validUntil < Date.MaxValue Then
                            sb.AppendLine($"Valid Until: {validUntil:d}")
                            Dim daysRemaining = CInt((validUntil.Date - Date.Now.Date).TotalDays)
                            If daysRemaining > 0 Then
                                sb.AppendLine($"Days Remaining: {daysRemaining}")
                            End If
                        End If
                        Dim confirmedOn = My.Settings.License_PrivateConfirmedOn
                        If confirmedOn > Date.MinValue Then
                            sb.AppendLine($"Last Confirmed: {confirmedOn:d}")
                        End If
                        sb.AppendLine()
                        sb.AppendLine("This license is for private, non-commercial use only.")

                    Case LicenseState.PrivateExpired
                        sb.AppendLine("License Type: Private License (EXPIRED)")
                        sb.AppendLine()
                        Dim validUntil = My.Settings.License_ValidUntil
                        If validUntil > Date.MinValue Then
                            sb.AppendLine($"Expired On: {validUntil:d}")
                        End If
                        sb.AppendLine()
                        sb.AppendLine($"Please renew your license at {AN4}.")

                    Case LicenseState.ProActive, LicenseState.ProOfflineGrace, LicenseState.TestingProActive
                        Dim productName = My.Settings.License_ProductName
                        If _currentLicenseState = LicenseState.TestingProActive Then
                            sb.AppendLine("License Type: Testing Pro License")
                            sb.AppendLine()
                            sb.AppendLine("⚠ FOR NON-PRODUCTIVE TESTING ONLY")
                        Else
                            ' Use actual product name, not generic "Professional License"
                            If Not String.IsNullOrEmpty(productName) Then
                                sb.AppendLine($"License Type: {productName}")
                            Else
                                sb.AppendLine("License Type: Pro License")
                            End If
                        End If
                        Dim productId = My.Settings.License_ProductID
                        If Not String.IsNullOrEmpty(productId) Then
                            sb.AppendLine($"Product ID: {productId}")
                        End If
                        Dim userId = My.Settings.License_UserID
                        If Not String.IsNullOrEmpty(userId) Then
                            sb.AppendLine($"User ID: {userId}")
                        End If
                        Dim activatedOn = My.Settings.License_ActivatedOn
                        If activatedOn > Date.MinValue Then
                            sb.AppendLine($"Activated On: {activatedOn:d}")
                        End If
                        Dim lastCheck = My.Settings.License_LastCheck
                        If lastCheck > Date.MinValue Then
                            sb.AppendLine($"Last Verified: {lastCheck:d}")
                        End If
                        If _currentLicenseState = LicenseState.ProOfflineGrace Then
                            sb.AppendLine()
                            Dim offlineStart = My.Settings.License_OfflineGraceStart
                            Dim remainingDays = OfflineGracePeriodDays - CInt((Date.Now.Date - offlineStart.Date).TotalDays)
                            sb.AppendLine($"⚠ Offline mode: {remainingDays} day(s) remaining to reconnect")
                        End If

                    Case LicenseState.ProExpired
                        sb.AppendLine("License Type: Professional License (EXPIRED/REVOKED)")
                        sb.AppendLine()
                        Dim productName = My.Settings.License_ProductName
                        If Not String.IsNullOrEmpty(productName) Then
                            sb.AppendLine($"Product: {productName}")
                        End If
                        sb.AppendLine()
                        sb.AppendLine(StandardLicenseContactMessage)

                    Case LicenseState.ProPendingActivation
                        sb.AppendLine("License Type: Professional License (Pending Activation)")
                        sb.AppendLine()
                        sb.AppendLine("Your license is configured but not yet activated.")
                        sb.AppendLine("Activation will occur on the next application start.")

                    Case LicenseState.LegacyProActive
                        sb.AppendLine($"License Type: {LicenseStatus} (Legacy)")
                        sb.AppendLine()
                        If LicensedTill > Date.MinValue AndAlso LicensedTill < Date.MaxValue Then
                            sb.AppendLine($"Valid Until: {LicensedTill:d}")
                            Dim daysRemaining = CInt((LicensedTill.Date - Date.Now.Date).TotalDays)
                            If daysRemaining > 0 Then
                                sb.AppendLine($"Days Remaining: {daysRemaining}")
                            End If
                        End If
                        If LicenseUsers > 0 Then
                            sb.AppendLine($"Licensed Users: {LicenseUsers}")
                        End If
                        sb.AppendLine()
                        sb.AppendLine("⚠ This legacy license will require migration to the new system.")
                        sb.AppendLine("Please contact your administrator for a new license key.")

                    Case LicenseState.LegacyProMigrationRequired
                        sb.AppendLine($"License Type: {LicenseStatus} (Migration Required)")
                        sb.AppendLine()
                        sb.AppendLine("Your legacy license requires migration to the new licensing system.")
                        sb.AppendLine()
                        sb.AppendLine(StandardLicenseContactMessage)

                    Case Else
                        sb.AppendLine("License status unknown.")
                End Select

                ' Add contact info if available
                If Not String.IsNullOrEmpty(LicenseContact) Then
                    sb.AppendLine()
                    sb.AppendLine("Contact:")
                    sb.AppendLine(LicenseContact)
                End If

                ' Check if license update is allowed
                If LicenseFromConfig Then
                    canUpdateLicense = False
                    sb.AppendLine()
                    sb.AppendLine("(License is centrally managed - contact your administrator to change)")
                End If

                ' Show dialog with option to update license
                If canUpdateLicense Then
                    sb.AppendLine()
                    sb.AppendLine("Would you like to change your license type?")

                    Dim result = ShowCustomYesNoBox(
                        sb.ToString(),
                        "Change License",
                        "Close",
                        $"{AN} - License Information")

                    If result = 1 Then
                        If _licenseContext IsNot Nothing Then
                            ShowLicenseTypeSelectionDialog(_licenseContext)
                        Else
                            ShowCustomMessageBox(
                                "License update requires application restart. Please close and reopen the application to change your license.",
                                $"{AN} - License Update")
                        End If
                    End If
                Else
                    ShowCustomMessageBox(sb.ToString(), $"{AN} - License Information")
                End If

            Catch ex As Exception
                ShowCustomMessageBox($"Error retrieving license information: {ex.Message}", AN)
            End Try
        End Sub

        ''' <summary>
        ''' Returns a short license status string suitable for display.
        ''' Uses the stored product name where available.
        ''' </summary>
        Public Shared Function GetLicenseStatusShort() As String
            Try
                Select Case _currentLicenseState
                    Case LicenseState.None
                        Return "No license configured"

                    Case LicenseState.PrivateActive
                        Dim validUntil = My.Settings.License_ValidUntil
                        If validUntil > Date.MinValue AndAlso validUntil < Date.MaxValue Then
                            Return $"Private License (until {validUntil:d})"
                        End If
                        Return "Private License (active)"

                    Case LicenseState.PrivateExpired
                        Return "Private License (expired)"

                    Case LicenseState.PrivateReconfirmNeeded
                        Return "Private License (confirm needed)"

                    Case LicenseState.TestingProActive, LicenseState.ProActive
                        Dim userId = My.Settings.License_UserID
                        Dim productName = My.Settings.License_ProductName
                        Dim displayName = GetProductDisplayName(productName)
                        If Not String.IsNullOrEmpty(userId) Then
                            Return $"{displayName} - {userId}"
                        End If
                        Return $"{displayName} (active)"

                    Case LicenseState.ProExpired
                        Dim productName = My.Settings.License_ProductName
                        Dim displayName = GetProductDisplayName(productName)
                        Return $"{displayName} (expired)"

                    Case LicenseState.ProPendingActivation
                        Dim productName = My.Settings.License_ProductName
                        Dim displayName = GetProductDisplayName(productName)
                        Return $"{displayName} (pending)"

                    Case LicenseState.ProOfflineGrace
                        Dim offlineStart = My.Settings.License_OfflineGraceStart
                        Dim userId = My.Settings.License_UserID
                        Dim productName = My.Settings.License_ProductName
                        Dim displayName = GetProductDisplayName(productName)
                        If offlineStart > Date.MinValue Then
                            Dim remainingDays = OfflineGracePeriodDays - CInt((Date.Now.Date - offlineStart.Date).TotalDays)
                            If Not String.IsNullOrEmpty(userId) Then
                                Return $"{displayName} - {userId} (offline, {remainingDays}d left)"
                            End If
                            Return $"{displayName} (offline, {remainingDays}d left)"
                        End If
                        Return $"{displayName} (offline)"

                    Case LicenseState.LegacyProActive
                        If Not String.IsNullOrEmpty(LicenseStatus) Then
                            Return $"{LicenseStatus} (legacy)"
                        End If
                        Return "Legacy License (active)"

                    Case LicenseState.LegacyProMigrationRequired
                        If Not String.IsNullOrEmpty(LicenseStatus) Then
                            Return $"{LicenseStatus} (migrate!)"
                        End If
                        Return "Legacy License (migrate!)"

                    Case Else
                        Return "Unknown license state"
                End Select

            Catch
                Return "License status unavailable"
            End Try
        End Function

        ''' <summary>
        ''' Returns a display product name for status rendering, handling missing product titles.
        ''' </summary>
        Private Shared Function GetProductDisplayName(productName As String) As String
            If Not String.IsNullOrEmpty(productName) Then
                Return productName
            End If

            ' No product name available
            If Date.Now >= LegacyRegimeEndDate Then
                ' After new regime - indicate missing product title
                Return "Pro License (no product title available)"
            Else
                ' Before new regime - use generic name
                Return "Pro License"
            End If
        End Function


#End Region

#Region "Utility Functions"

        ''' <summary>
        ''' Expands environment variables in a license User ID string.        
        ''' Supports: %USERNAME%, %USERDOMAIN%, %COMPUTERNAME%, %USERPROFILE%,
        ''' %APPDATA%, %LOCALAPPDATA%, %TEMP%, %TMP%, %HOMEDRIVE%, %HOMEPATH%,
        ''' %LOGONSERVER%, %USERDNSDOMAIN%
        ''' </summary>
        ''' <param name="value">Value that may contain environment variables like %USERNAME%.</param>
        Private Shared Function ExpandLicenseEnvironmentVariables(value As String) As String
            If String.IsNullOrWhiteSpace(value) Then Return value

            Try
                Return Environment.ExpandEnvironmentVariables(value)
            Catch
                Return value
            End Try
        End Function


        ''' <summary>
        ''' Strips HTML tags from a license-related message string so it can be displayed
        ''' in plain-text message boxes. Decodes HTML entities and normalizes whitespace.
        ''' </summary>
        ''' <param name="message">Message that may contain HTML markup from the license API.</param>
        Private Shared Function StripHtmlFromLicenseMessage(message As String) As String
            If String.IsNullOrWhiteSpace(message) Then Return message

            Try
                ' Check if the message actually contains HTML tags
                If Not message.Contains("<") Then Return message

                ' Use the existing RemoveHTML function if available, otherwise do simple strip
                Dim cleaned = RemoveHTML(message)

                ' Normalize multiple whitespace/newlines into single spaces
                cleaned = System.Text.RegularExpressions.Regex.Replace(cleaned, "\s+", " ").Trim()

                Return If(String.IsNullOrWhiteSpace(cleaned), message, cleaned)
            Catch
                Return message
            End Try
        End Function


#End Region


    End Class
End Namespace
