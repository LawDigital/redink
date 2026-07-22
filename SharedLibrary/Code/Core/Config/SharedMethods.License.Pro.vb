' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.License.Pro.vb
' Purpose: Professional license UI and API integration for activation, status checking,
'          and deactivation using the configured license API endpoint.
'
' Architecture:
'  - License management UI (`ShowProLicenseEntryForm`):
'     - Provides fields for License Key, Product ID, and User ID (with environment-variable expansion for User ID).
'     - Supports a "prefilled" mode used when values come from configuration.
'     - Allows status check, activation (including "already activated" detection), deactivation, and clearing local storage.
'     - Uses `My.Settings` as the persistence mechanism for Pro license credentials and state.
'  - API integration (`CallLicenseApi` + `ParseLicenseApiResponse`):
'     - Invokes WooCommerce API Manager actions: `status`, `activate`, `deactivate`.
'     - Uses GET requests with `HttpClient`, TLS 1.2/1.3, a timeout (`ApiTimeoutMs`) and retry loop (`ApiRetryCount`).
'     - Parses JSON responses into `LicenseApiResponse` for unified handling by UI and core logic.
'  - Deactivation (`ShowDeactivationDialog`):
'     - Confirms deactivation, then calls the API to deactivate and clears stored license data.
'  - License status display (`ShowLicenseManagementDialog`, `GetDetailedLicenseStatus`):
'     - Shows a textual summary of current stored license info and allows opening the Pro management dialog.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Net
Imports System.Net.Http
Imports System.Text
Imports System.Web
Imports System.Windows.Forms
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary
    Partial Public Class SharedMethods

#Region "Professional License Entry Form"

        ''' <summary>
        ''' Shows the Pro license management form for entering, verifying, activating, and deactivating credentials.
        ''' </summary>
        ''' <param name="context">Execution context (may be <see langword="Nothing"/> for management-only scenarios).</param>
        ''' <param name="prefilledKey">License key to prefill in the UI.</param>
        ''' <param name="prefilledProductId">Product ID to prefill in the UI.</param>
        ''' <param name="prefilledUserId">User ID to prefill in the UI.</param>
        ''' <param name="prefilled">When <see langword="True"/>, the dialog is presented as a confirmation flow.</param>
        ''' <returns><see langword="True"/> when a Pro license was verified/activated and persisted; otherwise <see langword="False"/>.</returns>
        Private Shared Function ShowProLicenseEntryForm(context As ISharedContext,
                                                         prefilledKey As String,
                                                         prefilledProductId As String,
                                                         prefilledUserId As String,
                                                         prefilled As Boolean) As Boolean
            Using form As New Form()
                ' Let WinForms handle DPI scaling automatically
                ' Set AutoScaleDimensions to 96 DPI (design baseline), then AutoScaleMode.Dpi
                ' WinForms will scale ALL sizes from 96 DPI to current DPI automatically
                ' DO NOT manually multiply by dpiScale - that causes double-scaling!
                form.AutoScaleDimensions = New SizeF(96.0F, 96.0F)
                form.AutoScaleMode = AutoScaleMode.Dpi

                form.Text = $"{AN} - Manage Pro License"
                form.StartPosition = FormStartPosition.CenterScreen
                form.FormBorderStyle = FormBorderStyle.FixedDialog
                form.MaximizeBox = False
                form.MinimizeBox = False
                form.ShowInTaskbar = True
                form.TopMost = True

                form.FormBorderStyle = FormBorderStyle.Sizable
                form.MinimumSize = New Size(900, 720)

                Try
                    Dim bmp As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                    form.Icon = Icon.FromHandle(bmp.GetHicon())
                Catch
                End Try

                form.Font = New Font("Segoe UI", 9.5F)

                ' Main TableLayoutPanel - all values at 96 DPI baseline
                Dim mainLayout As New TableLayoutPanel() With {
                    .Dock = DockStyle.Fill,
                    .Padding = New Padding(20),
                    .ColumnCount = 2,
                    .RowCount = 9
                }

                ' Column styles
                mainLayout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
                mainLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))

                ' Row styles - use AutoSize for content rows, Percent for expandable row
                mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize)) ' 0: Title
                mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize)) ' 1: Description
                mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize)) ' 2: License Key
                mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize)) ' 3: Product ID
                mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize)) ' 4: User ID
                mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize)) ' 5: Activation Status
                mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize)) ' 6: License Info
                mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize)) ' 7: Query
                mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize)) ' 8: Buttons

                form.Controls.Add(mainLayout)

                ' Row 0: Title
                Dim lblTitle As New Label() With {
                    .Text = If(prefilled, "Confirm License Activation", "Manage Pro License"),
                    .Font = New Font("Segoe UI", 12.0F, FontStyle.Bold),
                    .AutoSize = True,
                    .Margin = New Padding(0, 0, 0, 5)
                }
                mainLayout.Controls.Add(lblTitle, 0, 0)
                mainLayout.SetColumnSpan(lblTitle, 2)

                ' Row 1: Description
                Dim descText = If(prefilled,
                    "Your administrator has provided the following license information. " &
                    "Please verify and click Activate to complete the activation.",
                    $"Enter your license credentials to activate or deactivate {AN}. " &
                    $"A license can be activated for one User ID at a time. " &
                    $"You can obtain a license at {AN4}{ProSubUrl}")
                Dim lblDescription As New Label() With {
                    .Text = descText,
                    .AutoSize = True,
                    .MaximumSize = New Size(750, 0),
                    .Margin = New Padding(0, 0, 0, 15)
                }
                mainLayout.Controls.Add(lblDescription, 0, 1)
                mainLayout.SetColumnSpan(lblDescription, 2)

                ' Row 2: License Key
                Dim lblKey As New Label() With {
                    .Text = "License Key:",
                    .AutoSize = True,
                    .Anchor = AnchorStyles.Left,
                    .Margin = New Padding(0, 0, 10, 0)
                }
                mainLayout.Controls.Add(lblKey, 0, 2)

                Dim txtKey As New TextBox() With {
                    .Text = prefilledKey,
                    .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
                    .Margin = New Padding(0, 3, 0, 3)
                }
                mainLayout.Controls.Add(txtKey, 1, 2)

                ' Row 3: Product ID
                Dim lblProductId As New Label() With {
                    .Text = "Product ID:",
                    .AutoSize = True,
                    .Anchor = AnchorStyles.Left,
                    .Margin = New Padding(0, 0, 10, 0)
                }
                mainLayout.Controls.Add(lblProductId, 0, 3)

                Dim txtProductId As New TextBox() With {
                    .Text = prefilledProductId,
                    .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
                    .Margin = New Padding(0, 3, 0, 3)
                }
                mainLayout.Controls.Add(txtProductId, 1, 3)

                ' Row 4: User ID
                Dim lblUserId As New Label() With {
                    .Text = "User ID:",
                    .AutoSize = True,
                    .Anchor = AnchorStyles.Left,
                    .Margin = New Padding(0, 0, 10, 0)
                }
                mainLayout.Controls.Add(lblUserId, 0, 4)

                Dim txtUserId As New TextBox() With {
                    .Text = prefilledUserId,
                    .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
                    .Margin = New Padding(0, 3, 0, 3)
                }
                mainLayout.Controls.Add(txtUserId, 1, 4)

                ' Tooltip for User ID
                Dim userIdToolTip As New ToolTip() With {
                    .AutoPopDelay = 15000,
                    .InitialDelay = 500,
                    .ReshowDelay = 200
                }
                userIdToolTip.SetToolTip(txtUserId, "Enter a unique identifier for your license activation." & vbCrLf &
                                                    "This value ties your license to you personally across all add-ins." & vbCrLf &
                                                    "If your administrator hasn't provided a specific value, use your email address.")

                ' Row 5: Activation Status
                Dim lblActivationStatusLabel As New Label() With {
                    .Text = "Activation Status:",
                    .AutoSize = True,
                    .Anchor = AnchorStyles.Left,
                    .Margin = New Padding(0, 8, 10, 8)
                }
                mainLayout.Controls.Add(lblActivationStatusLabel, 0, 5)

                Dim lblActivationStatus As New Label() With {
                    .Text = "Unknown (click 'Check Status' to verify)",
                    .AutoSize = True,
                    .Anchor = AnchorStyles.Left,
                    .ForeColor = Color.Gray,
                    .Margin = New Padding(0, 8, 0, 8)
                }
                mainLayout.Controls.Add(lblActivationStatus, 1, 5)

                ' Row 6: License Info
                Dim lblProductInfoLabel As New Label() With {
                    .Text = "License Info:",
                    .AutoSize = True,
                    .Anchor = AnchorStyles.Left Or AnchorStyles.Top,
                    .Margin = New Padding(0, 0, 10, 8)
                }
                mainLayout.Controls.Add(lblProductInfoLabel, 0, 6)

                Dim lblProductInfo As New Label() With {
                    .Text = "(not yet retrieved)",
                    .AutoSize = True,
                    .Anchor = AnchorStyles.Left Or AnchorStyles.Top,
                    .ForeColor = Color.Gray,
                    .Margin = New Padding(0, 0, 0, 8)
                }
                mainLayout.Controls.Add(lblProductInfo, 1, 6)

                ' Row 7: Status Details Panel
                Dim pnlStatus As New Panel() With {
                    .Dock = DockStyle.Fill,
                    .BorderStyle = BorderStyle.FixedSingle,
                    .Margin = New Padding(0, 5, 0, 15),
                    .AutoScroll = True
                }
                Dim lblStatus As New Label() With {
                    .Text = "Click 'Check Status' or 'Activate' to retrieve license information." & vbCrLf & vbCrLf,
                    .AutoSize = True,
                    .MaximumSize = New Size(730, 0),
                    .ForeColor = Color.Gray,
                    .Location = New Point(5, 5)
                }
                pnlStatus.Controls.Add(lblStatus)
                mainLayout.Controls.Add(pnlStatus, 0, 7)
                mainLayout.SetColumnSpan(pnlStatus, 2)

                ' Row 8: Buttons
                Dim buttonPanel As New FlowLayoutPanel() With {
                    .AutoSize = True,
                    .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    .FlowDirection = FlowDirection.LeftToRight,
                    .Anchor = AnchorStyles.Left Or AnchorStyles.Bottom,
                    .Margin = New Padding(0, 0, 0, 0),
                    .WrapContents = False
                }

                Dim btnCheckStatus As New Button() With {
                    .Text = "Check Status",
                    .AutoSize = True,
                    .Margin = New Padding(0, 0, 8, 0)
                }
                buttonPanel.Controls.Add(btnCheckStatus)

                Dim btnActivate As New Button() With {
                    .Text = "Activate",
                    .AutoSize = True,
                    .Margin = New Padding(0, 0, 8, 0)
                }
                buttonPanel.Controls.Add(btnActivate)

                Dim btnDeactivate As New Button() With {
                    .Text = "Deactivate",
                    .AutoSize = True,
                    .Margin = New Padding(0, 0, 8, 0)
                }
                buttonPanel.Controls.Add(btnDeactivate)

                Dim btnClearLicense As New Button() With {
                    .Text = "Clear License",
                    .AutoSize = True,
                    .Margin = New Padding(0, 0, 8, 0)
                }
                buttonPanel.Controls.Add(btnClearLicense)

                Dim btnClose As New Button() With {
                    .Text = "Close",
                    .AutoSize = True,
                    .Margin = New Padding(0)
                }
                buttonPanel.Controls.Add(btnClose)

                mainLayout.Controls.Add(buttonPanel, 0, 8)
                mainLayout.SetColumnSpan(buttonPanel, 2)

                ' Force layout recalculation after form is displayed
                ' This is necessary because rows 6 and 7 have empty/minimal content at construction time
                ' and TableLayoutPanel calculates row heights during initial layout
                AddHandler form.Load, Sub()
                                          mainLayout.PerformLayout()
                                      End Sub

                ' State tracking
                Dim dialogResult As Boolean = False
                Dim isCurrentlyActivated As Boolean = False
                Dim statusVerified As Boolean = False
                Dim lastVerifiedUserId As String = ""
                Dim lastVerifiedProductId As String = ""
                Dim lastVerifiedKey As String = ""
                Dim licenseExplicitlyCleared As Boolean = False
                Dim closeHandlerExecuted As Boolean = False

                ' Capture original license state
                Dim originalLicenseType As String = ""
                Dim hadOriginalLicense As Boolean = False
                Try
                    originalLicenseType = My.Settings.License_Type
                    hadOriginalLicense = Not String.IsNullOrEmpty(originalLicenseType)
                Catch
                End Try

                ' Helper: Update status display from API response
                Dim updateStatusDisplay As Action(Of LicenseApiResponse, Boolean) =
                    Sub(response As LicenseApiResponse, isError As Boolean)
                        If response Is Nothing Then Return

                        If response.Success Then
                            Dim productInfo = response.ProductTitle
                            If response.TotalActivationsPurchased > 0 Then
                                productInfo &= $"  ({response.TotalActivations} of {response.TotalActivationsPurchased} activations used)"
                            End If
                            lblProductInfo.Text = productInfo
                            lblProductInfo.ForeColor = Color.DarkBlue

                            If response.Activated Then
                                lblActivationStatus.Text = "ACTIVATED for this User ID"
                                lblActivationStatus.ForeColor = Color.DarkGreen
                                isCurrentlyActivated = True
                            Else
                                If response.ActivationsRemaining > 0 Then
                                    lblActivationStatus.Text = $"Not Activated ({response.ActivationsRemaining} activation slot(s) available)"
                                Else
                                    lblActivationStatus.Text = "Not Activated (no slots remaining)"
                                End If
                                lblActivationStatus.ForeColor = Color.DarkOrange
                                isCurrentlyActivated = False
                            End If
                        Else
                            Dim errorSummary = "Error"
                            If Not String.IsNullOrEmpty(response.ErrorMessage) Then
                                Dim shortError = response.ErrorMessage
                                If shortError.Length > 60 Then shortError = shortError.Substring(0, 57) & "..."
                                errorSummary = $"Error: {shortError}"
                            End If
                            lblActivationStatus.Text = errorSummary
                            lblActivationStatus.ForeColor = Color.DarkRed
                            lblProductInfo.Text = ""
                            isCurrentlyActivated = False
                        End If
                    End Sub

                ' Helper: Show status message
                Dim showStatusMessage As Action(Of String, Color) =
                    Sub(message As String, color As Color)
                        lblStatus.Text = message
                        lblStatus.ForeColor = color
                    End Sub

                Dim buildActivationUsageText As Func(Of LicenseApiResponse, String, String) =
                        Function(response As LicenseApiResponse, enteredKey As String) As String
                            If response Is Nothing Then Return ""

                            If response.TotalActivationsPurchased > 0 Then
                                Return $"{vbCrLf}Activations: {response.TotalActivations} / {response.TotalActivationsPurchased}"
                            End If

                            If IsOfflineDomainLicenseKey(enteredKey) Then
                                Return $"{vbCrLf}Activation Slots: not limited for offline-domain licenses"
                            End If

                            Return ""
                        End Function

                ' Helper: Load stored credentials
                Dim loadStoredCredentials As Action =
                    Sub()
                        Try
                            If HasStoredProLicense() Then
                                If String.IsNullOrWhiteSpace(txtKey.Text) Then txtKey.Text = My.Settings.License_Key
                                If String.IsNullOrWhiteSpace(txtProductId.Text) Then txtProductId.Text = My.Settings.License_ProductID
                                If String.IsNullOrWhiteSpace(txtUserId.Text) Then txtUserId.Text = My.Settings.License_UserID
                            End If
                        Catch
                        End Try
                    End Sub

                loadStoredCredentials()

                ' Check Status button handler
                AddHandler btnCheckStatus.Click, Sub()
                                                     Dim key = txtKey.Text.Trim()
                                                     Dim productId = txtProductId.Text.Trim()
                                                     Dim userId = ExpandLicenseEnvironmentVariables(txtUserId.Text.Trim())

                                                     If IsOfflineDomainLicenseKey(key) Then
                                                         Dim confirmLocalResult = ShowCustomYesNoBox(
                                                            "Are you sure you want to remove the offline-domain license from local storage?" & vbCrLf & vbCrLf &
                                                            "No online deactivation is necessary.",
                                                            "Remove License",
                                                            "Cancel",
                                                            $"{AN} - License Deactivation")

                                                         If confirmLocalResult <> 1 Then Return

                                                         ClearStoredLicense()
                                                         licenseExplicitlyCleared = True
                                                         dialogResult = False
                                                         isCurrentlyActivated = False
                                                         statusVerified = False

                                                         txtKey.Text = ""
                                                         txtProductId.Text = ""
                                                         txtUserId.Text = ""

                                                         lblActivationStatus.Text = "DEACTIVATED"
                                                         lblActivationStatus.ForeColor = Color.DarkOrange
                                                         lblProductInfo.Text = "Offline Domain License"
                                                         showStatusMessage("Offline-domain license removed from local storage.", Color.DarkGreen)

                                                         ShowCustomMessageBox(
                                                            "Offline-domain license removed from local storage." & vbCrLf & vbCrLf &
                                                            "No online deactivation was necessary.",
                                                            $"{AN} - License Deactivated")
                                                         Return
                                                     End If

                                                     If String.IsNullOrWhiteSpace(key) OrElse String.IsNullOrWhiteSpace(productId) OrElse String.IsNullOrWhiteSpace(userId) Then
                                                         showStatusMessage("Please fill in all fields before checking status.", Color.DarkRed)
                                                         lblActivationStatus.Text = "Unknown"
                                                         lblActivationStatus.ForeColor = Color.Gray
                                                         lblProductInfo.Text = ""
                                                         statusVerified = False
                                                         Return
                                                     End If

                                                     showStatusMessage("Checking license status...", Color.DarkBlue)
                                                     lblActivationStatus.Text = "Checking..."
                                                     lblActivationStatus.ForeColor = Color.DarkBlue
                                                     form.Cursor = Cursors.WaitCursor
                                                     form.Refresh()
                                                     Application.DoEvents()

                                                     Try
                                                         Dim response = CallLicenseApi("status", productId, key, userId)

                                                         If response.Success Then
                                                             updateStatusDisplay(response, False)

                                                             Dim statusMsg As New StringBuilder()
                                                             statusMsg.AppendLine($"Status: {response.StatusCheck}")
                                                             statusMsg.AppendLine($"Product: {response.ProductTitle}")
                                                             If response.TotalActivationsPurchased > 0 Then
                                                                 statusMsg.AppendLine($"Activations: {response.TotalActivations} / {response.TotalActivationsPurchased}")
                                                             ElseIf IsOfflineDomainLicenseKey(key) Then
                                                                 statusMsg.AppendLine("Activation Slots: not limited for offline-domain licenses")
                                                             End If
                                                             If response.Activated Then
                                                                 statusMsg.AppendLine()
                                                                 statusMsg.AppendLine("✓ This User ID is ACTIVATED.")
                                                             Else
                                                                 statusMsg.AppendLine()
                                                                 statusMsg.AppendLine("This User ID is not currently activated.")
                                                                 If response.ActivationsRemaining > 0 Then
                                                                     statusMsg.AppendLine($"{response.ActivationsRemaining} activation slot(s) available.")
                                                                 Else
                                                                     statusMsg.AppendLine("No activation slots remaining. Deactivate another User ID first.")
                                                                 End If
                                                             End If
                                                             showStatusMessage(statusMsg.ToString(), Color.DarkGreen)

                                                             statusVerified = True
                                                             lastVerifiedUserId = userId
                                                             lastVerifiedProductId = productId
                                                             lastVerifiedKey = key

                                                             If response.Activated Then
                                                                 SaveProLicenseToSettings(productId, key, userId, response.ProductTitle, True)
                                                                 dialogResult = True
                                                             End If
                                                         Else
                                                             updateStatusDisplay(response, True)
                                                             showStatusMessage($"Error checking status:{vbCrLf}{vbCrLf}{response.ErrorMessage}{vbCrLf}{vbCrLf}Raw response: {TruncateForLog(response.RawJson, 300)}", Color.DarkRed)
                                                             statusVerified = False
                                                         End If
                                                     Catch ex As Exception
                                                         showStatusMessage($"Error: {ex.Message}", Color.DarkRed)
                                                         lblActivationStatus.Text = "Error"
                                                         lblActivationStatus.ForeColor = Color.DarkRed
                                                         lblProductInfo.Text = ""
                                                         statusVerified = False
                                                         isCurrentlyActivated = False
                                                     Finally
                                                         form.Cursor = Cursors.Default
                                                     End Try
                                                 End Sub

                ' Activate button handler
                AddHandler btnActivate.Click, Sub()
                                                  Dim key = txtKey.Text.Trim()
                                                  Dim productId = txtProductId.Text.Trim()
                                                  Dim userId = ExpandLicenseEnvironmentVariables(txtUserId.Text.Trim())

                                                  If String.IsNullOrWhiteSpace(key) Then
                                                      ShowCustomMessageBox("Please enter your License Key.", $"{AN} - License")
                                                      Return
                                                  End If
                                                  If String.IsNullOrWhiteSpace(productId) Then
                                                      ShowCustomMessageBox("Please enter your Product ID.", $"{AN} - License")
                                                      Return
                                                  End If
                                                  If String.IsNullOrWhiteSpace(userId) Then
                                                      ShowCustomMessageBox("Please enter your User ID.", $"{AN} - License")
                                                      Return
                                                  End If

                                                  showStatusMessage("Activating license...", Color.DarkBlue)
                                                  lblActivationStatus.Text = "Activating..."
                                                  lblActivationStatus.ForeColor = Color.DarkBlue
                                                  btnActivate.Enabled = False
                                                  form.Cursor = Cursors.WaitCursor
                                                  form.Refresh()
                                                  Application.DoEvents()

                                                  Try
                                                      Dim statusResponse = CallLicenseApi("status", productId, key, userId)

                                                      If statusResponse.Success AndAlso statusResponse.Activated Then
                                                          SaveProLicenseToSettings(productId, key, userId, statusResponse.ProductTitle, True)
                                                          updateStatusDisplay(statusResponse, False)
                                                          showStatusMessage($"✓ License is already activated for this User ID.{vbCrLf}{vbCrLf}Product: {statusResponse.ProductTitle}{vbCrLf}User ID: {userId}{buildActivationUsageText(statusResponse, key)}", Color.DarkGreen)

                                                          statusVerified = True
                                                          lastVerifiedUserId = userId
                                                          lastVerifiedProductId = productId
                                                          lastVerifiedKey = key
                                                          dialogResult = True

                                                          ShowCustomMessageBox($"License is already activated for this User ID." & vbCrLf & vbCrLf & $"Product: {statusResponse.ProductTitle}" & vbCrLf & $"User ID: {userId}", $"{AN} - License")
                                                          Return
                                                      End If

                                                      Dim activateResponse = CallLicenseApi("activate", productId, key, userId)

                                                      If activateResponse.Success AndAlso activateResponse.Activated Then
                                                          Dim productTitle = activateResponse.ProductTitle

                                                          If String.IsNullOrEmpty(productTitle) OrElse productTitle = "(no product title available)" Then
                                                              Try
                                                                  Dim postActivateStatus = CallLicenseApi("status", productId, key, userId)
                                                                  If postActivateStatus.Success AndAlso Not String.IsNullOrEmpty(postActivateStatus.ProductTitle) AndAlso postActivateStatus.ProductTitle <> "(no product title available)" Then
                                                                      productTitle = postActivateStatus.ProductTitle
                                                                      If postActivateStatus.TotalActivationsPurchased > 0 Then
                                                                          activateResponse.TotalActivationsPurchased = postActivateStatus.TotalActivationsPurchased
                                                                          activateResponse.TotalActivations = postActivateStatus.TotalActivations
                                                                          activateResponse.ActivationsRemaining = postActivateStatus.ActivationsRemaining
                                                                      End If
                                                                  End If
                                                              Catch
                                                              End Try
                                                          End If

                                                          activateResponse.ProductTitle = productTitle
                                                          SaveProLicenseToSettings(productId, key, userId, productTitle, True)
                                                          LogLicenseEvent("Pro License", "Activated successfully", alwaysLog:=True)

                                                          updateStatusDisplay(activateResponse, False)
                                                          showStatusMessage($"✓ License activated successfully!{vbCrLf}{vbCrLf}Product: {productTitle}{vbCrLf}User ID: {userId}{buildActivationUsageText(activateResponse, key)}", Color.DarkGreen)

                                                          statusVerified = True
                                                          lastVerifiedUserId = userId
                                                          lastVerifiedProductId = productId
                                                          lastVerifiedKey = key
                                                          dialogResult = True

                                                          ShowCustomMessageBox($"License activated successfully!" & vbCrLf & vbCrLf & $"Product: {productTitle}" & vbCrLf & $"User ID: {userId}", $"{AN} - License Activated")
                                                      Else
                                                          Dim recheckResponse = CallLicenseApi("status", productId, key, userId)

                                                          If recheckResponse.Success AndAlso recheckResponse.Activated Then
                                                              SaveProLicenseToSettings(productId, key, userId, recheckResponse.ProductTitle, True)
                                                              updateStatusDisplay(recheckResponse, False)
                                                              showStatusMessage($"✓ License is activated for this User ID.{vbCrLf}{vbCrLf}Product: {recheckResponse.ProductTitle}{vbCrLf}User ID: {userId}{buildActivationUsageText(recheckResponse, key)}", Color.DarkGreen)

                                                              statusVerified = True
                                                              lastVerifiedUserId = userId
                                                              lastVerifiedProductId = productId
                                                              lastVerifiedKey = key
                                                              dialogResult = True

                                                              ShowCustomMessageBox($"License is activated for this User ID." & vbCrLf & vbCrLf & $"Product: {recheckResponse.ProductTitle}" & vbCrLf & $"User ID: {userId}", $"{AN} - License")
                                                          ElseIf recheckResponse.Success AndAlso recheckResponse.ActivationsRemaining <= 0 Then
                                                              updateStatusDisplay(recheckResponse, False)
                                                              showStatusMessage($"Activation failed: No activation slots remaining.{vbCrLf}{vbCrLf}Product: {recheckResponse.ProductTitle}{vbCrLf}Activations: {recheckResponse.TotalActivations} / {recheckResponse.TotalActivationsPurchased}{vbCrLf}{vbCrLf}Please deactivate another User ID first, or contact your administrator.", Color.DarkRed)
                                                              statusVerified = True
                                                          Else
                                                              Dim errorMsg = If(Not String.IsNullOrEmpty(activateResponse.ErrorMessage), activateResponse.ErrorMessage, "Unknown error")
                                                              updateStatusDisplay(activateResponse, True)
                                                              showStatusMessage($"Activation failed: {errorMsg}{vbCrLf}{vbCrLf}Raw response: {TruncateForLog(activateResponse.RawJson, 300)}", Color.DarkRed)
                                                              statusVerified = False
                                                          End If
                                                      End If
                                                  Catch ex As Exception
                                                      showStatusMessage($"Error during activation: {ex.Message}", Color.DarkRed)
                                                      lblActivationStatus.Text = "Error"
                                                      lblActivationStatus.ForeColor = Color.DarkRed
                                                      LogLicenseEvent("Activation Error", ex.Message, alwaysLog:=True)
                                                      statusVerified = False
                                                  Finally
                                                      btnActivate.Enabled = True
                                                      form.Cursor = Cursors.Default
                                                  End Try
                                              End Sub

                ' Deactivate button handler
                AddHandler btnDeactivate.Click, Sub()
                                                    Dim key = txtKey.Text.Trim()
                                                    Dim productId = txtProductId.Text.Trim()
                                                    Dim userId = ExpandLicenseEnvironmentVariables(txtUserId.Text.Trim())

                                                    If String.IsNullOrWhiteSpace(key) OrElse String.IsNullOrWhiteSpace(productId) OrElse String.IsNullOrWhiteSpace(userId) Then
                                                        ShowCustomMessageBox("Please fill in all license fields before deactivating.", $"{AN} - License")
                                                        Return
                                                    End If

                                                    Dim confirmResult = ShowCustomYesNoBox($"Are you sure you want to deactivate the license?" & vbCrLf & vbCrLf & $"User ID to deactivate: {userId}" & vbCrLf & vbCrLf & "The activation slot will be freed and can be used for another User ID.", "Deactivate", "Cancel", $"{AN} - License Deactivation")

                                                    If confirmResult <> 1 Then Return

                                                    showStatusMessage("Deactivating license...", Color.DarkBlue)
                                                    lblActivationStatus.Text = "Deactivating..."
                                                    lblActivationStatus.ForeColor = Color.DarkBlue
                                                    btnDeactivate.Enabled = False
                                                    form.Cursor = Cursors.WaitCursor
                                                    form.Refresh()
                                                    Application.DoEvents()

                                                    Try
                                                        Dim response = CallLicenseApi("deactivate", productId, key, userId)

                                                        If response.Success Then
                                                            LogLicenseEvent("License Deactivated", $"ProductID={productId}, UserID={userId}", alwaysLog:=True)

                                                            Dim statusResponse = CallLicenseApi("status", productId, key, userId)
                                                            Dim activationsRemaining = If(statusResponse.Success, statusResponse.ActivationsRemaining, 0)
                                                            Dim totalActivations = If(statusResponse.Success, statusResponse.TotalActivations, 0)
                                                            Dim totalPurchased = If(statusResponse.Success, statusResponse.TotalActivationsPurchased, 0)
                                                            Dim productTitle = If(statusResponse.Success AndAlso Not String.IsNullOrEmpty(statusResponse.ProductTitle), statusResponse.ProductTitle, response.ProductTitle)

                                                            lblActivationStatus.Text = "DEACTIVATED"
                                                            lblActivationStatus.ForeColor = Color.DarkOrange
                                                            lblProductInfo.Text = If(totalPurchased > 0, $"{productTitle}  ({totalActivations} of {totalPurchased} activations used)", productTitle)

                                                            showStatusMessage($"✓ License deactivated successfully.{vbCrLf}{vbCrLf}User ID deactivated: {userId}{vbCrLf}Product: {productTitle}{vbCrLf}Activations now: {totalActivations} / {totalPurchased}{vbCrLf}Slots available: {activationsRemaining}", Color.DarkGreen)

                                                            isCurrentlyActivated = False
                                                            statusVerified = True
                                                            dialogResult = False

                                                            ShowCustomMessageBox($"License deactivated successfully." & vbCrLf & vbCrLf & $"User ID deactivated: {userId}" & vbCrLf & $"Product: {productTitle}" & vbCrLf & vbCrLf & $"Activation slots now available: {activationsRemaining}", $"{AN} - License Deactivated")

                                                            Try
                                                                If HasStoredProLicense() AndAlso My.Settings.License_UserID.Equals(userId, StringComparison.OrdinalIgnoreCase) Then
                                                                    ClearStoredLicense()
                                                                    licenseExplicitlyCleared = True
                                                                End If
                                                            Catch
                                                            End Try

                                                            Try
                                                                If HasStoredProLicense() Then
                                                                    txtKey.Text = My.Settings.License_Key
                                                                    txtProductId.Text = My.Settings.License_ProductID
                                                                    txtUserId.Text = My.Settings.License_UserID
                                                                    form.Refresh()
                                                                    Application.DoEvents()
                                                                    Threading.Thread.Sleep(500)
                                                                    btnCheckStatus.PerformClick()
                                                                End If
                                                            Catch
                                                            End Try
                                                        Else
                                                            showStatusMessage($"Deactivation failed: {response.ErrorMessage}{vbCrLf}{vbCrLf}Raw response: {TruncateForLog(response.RawJson, 300)}", Color.DarkRed)
                                                            lblActivationStatus.Text = "Deactivation Failed"
                                                            lblActivationStatus.ForeColor = Color.DarkRed
                                                        End If
                                                    Catch ex As Exception
                                                        showStatusMessage($"Error during deactivation: {ex.Message}", Color.DarkRed)
                                                        lblActivationStatus.Text = "Error"
                                                        lblActivationStatus.ForeColor = Color.DarkRed
                                                        LogLicenseEvent("Deactivation Error", ex.Message, alwaysLog:=True)
                                                    Finally
                                                        btnDeactivate.Enabled = True
                                                        form.Cursor = Cursors.Default
                                                    End Try
                                                End Sub

                ' Clear License button handler
                AddHandler btnClearLicense.Click, Sub()
                                                      Dim currentLicenseType As String = ""
                                                      Try
                                                          currentLicenseType = My.Settings.License_Type
                                                      Catch
                                                      End Try

                                                      If String.IsNullOrEmpty(currentLicenseType) Then
                                                          ShowCustomMessageBox("No license is currently stored.", $"{AN} - License")
                                                          Return
                                                      End If

                                                      Dim licenseDescription As String
                                                      If currentLicenseType.Equals("Private", StringComparison.OrdinalIgnoreCase) Then
                                                          licenseDescription = "Private (Non-Commercial) License"
                                                      ElseIf currentLicenseType.Equals("Pro", StringComparison.OrdinalIgnoreCase) Then
                                                          licenseDescription = $"Professional License (User ID: {My.Settings.License_UserID})"
                                                      Else
                                                          licenseDescription = $"{currentLicenseType} License"
                                                      End If

                                                      Dim confirmResult = ShowCustomYesNoBox($"Are you sure you want to clear the stored license?" & vbCrLf & vbCrLf & $"Current License: {licenseDescription}" & vbCrLf & vbCrLf & "This will remove the license from local storage. If this is a Pro license, it will NOT be deactivated on the server (use 'Deactivate' button to free the activation slot)." & vbCrLf & vbCrLf & "The add-in may not function properly without a valid license.", "Clear License", "Cancel", $"{AN} - Clear License")

                                                      If confirmResult <> 1 Then Return

                                                      ClearStoredLicense(True)
                                                      licenseExplicitlyCleared = True
                                                      dialogResult = False
                                                      isCurrentlyActivated = False
                                                      statusVerified = False

                                                      txtKey.Text = ""
                                                      txtProductId.Text = ""
                                                      txtUserId.Text = ""

                                                      lblActivationStatus.Text = "No license stored"
                                                      lblActivationStatus.ForeColor = Color.Gray
                                                      lblProductInfo.Text = ""
                                                      showStatusMessage("License cleared from local storage.", Color.DarkOrange)

                                                      LogLicenseEvent("License Cleared", $"Previous type: {currentLicenseType}", alwaysLog:=True)

                                                      ShowCustomMessageBox("License has been cleared from local storage." & vbCrLf & vbCrLf & "You can now enter new license credentials or close this dialog.", $"{AN} - License Cleared")
                                                  End Sub

                ' Close button handler
                AddHandler btnClose.Click, Sub()
                                               closeHandlerExecuted = True

                                               If licenseExplicitlyCleared Then
                                                   form.Close()
                                                   Return
                                               End If

                                               Dim currentKey = txtKey.Text.Trim()
                                               Dim currentProductId = txtProductId.Text.Trim()
                                               Dim currentUserId = ExpandLicenseEnvironmentVariables(txtUserId.Text.Trim())

                                               If dialogResult AndAlso isCurrentlyActivated AndAlso statusVerified AndAlso currentKey.Equals(lastVerifiedKey, StringComparison.Ordinal) AndAlso currentProductId.Equals(lastVerifiedProductId, StringComparison.Ordinal) AndAlso currentUserId.Equals(lastVerifiedUserId, StringComparison.OrdinalIgnoreCase) Then
                                                   form.Close()
                                                   Return
                                               End If

                                               If String.IsNullOrWhiteSpace(currentKey) AndAlso String.IsNullOrWhiteSpace(currentProductId) AndAlso String.IsNullOrWhiteSpace(currentUserId) Then
                                                   form.Close()
                                                   Return
                                               End If

                                               If HasStoredProLicense() Then
                                                   Try
                                                       If currentKey.Equals(My.Settings.License_Key, StringComparison.Ordinal) AndAlso currentProductId.Equals(My.Settings.License_ProductID, StringComparison.Ordinal) AndAlso currentUserId.Equals(My.Settings.License_UserID, StringComparison.OrdinalIgnoreCase) Then
                                                           form.Close()
                                                           Return
                                                       End If
                                                   Catch
                                                   End Try
                                               End If

                                               If Not statusVerified OrElse Not currentKey.Equals(lastVerifiedKey, StringComparison.Ordinal) OrElse Not currentProductId.Equals(lastVerifiedProductId, StringComparison.Ordinal) OrElse Not currentUserId.Equals(lastVerifiedUserId, StringComparison.OrdinalIgnoreCase) Then
                                                   Dim closeResult = ShowCustomYesNoBox("You have entered new license credentials that have not been verified." & vbCrLf & vbCrLf & "Would you like to verify and activate these credentials before closing?" & vbCrLf & vbCrLf & "Click 'Verify Now' to check the license status, or 'Close Without Saving' to discard the changes and keep your existing license.", "Verify Now", "Close Without Saving", $"{AN} - License")

                                                   If closeResult = 1 Then
                                                       btnActivate.PerformClick()
                                                       Return
                                                   End If
                                               End If

                                               form.Close()
                                           End Sub

                ' FormClosing handler
                AddHandler form.FormClosing, Sub(sender As Object, e As FormClosingEventArgs)
                                                 If closeHandlerExecuted Then Return
                                                 If licenseExplicitlyCleared Then Return

                                                 Dim currentKey = txtKey.Text.Trim()
                                                 Dim currentProductId = txtProductId.Text.Trim()
                                                 Dim currentUserId = ExpandLicenseEnvironmentVariables(txtUserId.Text.Trim())

                                                 If dialogResult AndAlso isCurrentlyActivated AndAlso statusVerified AndAlso currentKey.Equals(lastVerifiedKey, StringComparison.Ordinal) AndAlso currentProductId.Equals(lastVerifiedProductId, StringComparison.Ordinal) AndAlso currentUserId.Equals(lastVerifiedUserId, StringComparison.OrdinalIgnoreCase) Then
                                                     Return
                                                 End If

                                                 If String.IsNullOrWhiteSpace(currentKey) AndAlso String.IsNullOrWhiteSpace(currentProductId) AndAlso String.IsNullOrWhiteSpace(currentUserId) Then
                                                     Return
                                                 End If

                                                 If HasStoredProLicense() Then
                                                     Try
                                                         If currentKey.Equals(My.Settings.License_Key, StringComparison.Ordinal) AndAlso currentProductId.Equals(My.Settings.License_ProductID, StringComparison.Ordinal) AndAlso currentUserId.Equals(My.Settings.License_UserID, StringComparison.OrdinalIgnoreCase) Then
                                                             Return
                                                         End If
                                                     Catch
                                                     End Try
                                                 End If

                                                 If Not statusVerified OrElse Not currentKey.Equals(lastVerifiedKey, StringComparison.Ordinal) OrElse Not currentProductId.Equals(lastVerifiedProductId, StringComparison.Ordinal) OrElse Not currentUserId.Equals(lastVerifiedUserId, StringComparison.OrdinalIgnoreCase) Then
                                                     Dim closeResult = ShowCustomYesNoBox("You have entered new license credentials that have not been verified." & vbCrLf & vbCrLf & "Would you like to verify and activate these credentials before closing?" & vbCrLf & vbCrLf & "Click 'Verify Now' to check the license status, or 'Close Without Saving' to discard the changes and keep your existing license.", "Verify Now", "Close Without Saving", $"{AN} - License")

                                                     If closeResult = 1 Then
                                                         e.Cancel = True
                                                         btnActivate.PerformClick()
                                                         Return
                                                     End If
                                                 End If
                                             End Sub

                form.ShowDialog()
                Return dialogResult
            End Using
        End Function

#End Region

#Region "License API Calls"

        ''' <summary>
        ''' Calls the license API using the specified action and credentials.
        ''' </summary>
        ''' <param name="action">API action: "status", "activate", or "deactivate".</param>
        ''' <param name="productId">Product identifier transmitted as `product_id`.</param>
        ''' <param name="licenseKey">License key transmitted as `api_key`.</param>
        ''' <param name="userId">User identifier transmitted as `instance`.</param>
        ''' <returns>A populated <see cref="LicenseApiResponse"/> instance.</returns>
        Private Shared Function CallLicenseApi(action As String, productId As String, licenseKey As String, userId As String) As LicenseApiResponse
            Dim result As New LicenseApiResponse()

            If IsOfflineDomainLicenseKey(licenseKey) Then
                Dim offlineResponse As LicenseApiResponse = Nothing

                If TryCreateOfflineDomainLicenseResponse(action, productId, licenseKey, userId, offlineResponse) Then
                    Return offlineResponse
                End If
            End If

            Try
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls13

                ' Build API URL
                Dim apiAction As String
                Select Case action.ToLowerInvariant()
                    Case "status"
                        apiAction = "status"
                    Case "activate"
                        apiAction = "activate"
                    Case "deactivate"
                        apiAction = "deactivate"
                    Case Else
                        result.ErrorMessage = $"Unknown API action: {action}"
                        Return result
                End Select

                Dim url = $"{LicenseApiBaseUrl}&wc_am_action={apiAction}" &
                          $"&product_id={HttpUtility.UrlEncode(productId)}" &
                          $"&api_key={HttpUtility.UrlEncode(licenseKey)}" &
                          $"&instance={HttpUtility.UrlEncode(userId)}"

                LogLicenseEvent("API Call", $"Action={action}, ProductID={productId}, Instance={userId}")

                ' Make request with retries, bounded by an overall wall-clock budget so the
                ' calling (host UI) thread is never blocked longer than ApiTotalBudgetMs.
                Dim lastError As String = ""
                Dim budget = System.Diagnostics.Stopwatch.StartNew()
                For attempt As Integer = 1 To ApiRetryCount
                    Dim remainingMs = ApiTotalBudgetMs - CInt(budget.ElapsedMilliseconds)
                    If remainingMs <= 0 Then
                        lastError = "License server did not respond within the allotted time budget."
                        LogLicenseEvent("API Budget", $"Overall budget of {ApiTotalBudgetMs}ms exceeded before attempt {attempt}; deferring.")
                        Exit For
                    End If

                    Try
                        Dim httpResponse = SendHttpRequestAsync(
                            New SharedHttpRequest() With {
                                .Url = url,
                                .Method = "GET",
                                .TimeoutMs = System.Math.Min(ApiTimeoutMs, remainingMs),
                                .UserAgent = $"{AN}/1.0",
                                .StackPreference = HttpStackPreference.PreferConfiguredDefault
                            }).GetAwaiter().GetResult()

                        Dim json = httpResponse.GetBodyAsString()
                        result.RawJson = json

                        LogLicenseEvent("API Response", $"Attempt {attempt} via {httpResponse.UsedStack}: {TruncateForLog(json, 500)}")

                        ' Parse response
                        Return ParseLicenseApiResponse(json, action)

                    Catch ex As Exception
                        lastError = ex.Message
                        LogLicenseEvent("API Error", $"Attempt {attempt}: {ex.Message}")

                        If attempt < ApiRetryCount Then
                            Dim backoffMs = 1000 * attempt ' Exponential backoff
                            Dim budgetLeft = ApiTotalBudgetMs - CInt(budget.ElapsedMilliseconds)
                            If budgetLeft <= 0 Then Exit For
                            Threading.Thread.Sleep(System.Math.Min(backoffMs, budgetLeft))
                        End If
                    End Try
                Next

                result.ErrorMessage = $"Failed within {ApiTotalBudgetMs}ms budget: {lastError}"
                Return result

            Catch ex As Exception
                result.ErrorMessage = ex.Message
                LogLicenseEvent("API Exception", ex.Message, alwaysLog:=True)
                Return result
            End Try
        End Function

        ''' <summary>
        ''' Parses the JSON returned by the license API into a <see cref="LicenseApiResponse"/>.
        ''' </summary>
        ''' <param name="json">Raw JSON payload returned by the API.</param>
        ''' <param name="action">Action that produced this response.</param>
        Private Shared Function ParseLicenseApiResponse(json As String, action As String) As LicenseApiResponse
            Dim result As New LicenseApiResponse()
            result.RawJson = json

            Try
                Dim jObj = JObject.Parse(json)

                ' Check for success field at root level
                If jObj("success") IsNot Nothing Then
                    result.Success = jObj("success").Value(Of Boolean)()
                End If

                ' Check for error at root level
                If jObj("error") IsNot Nothing Then
                    result.ErrorMessage = jObj("error").Value(Of String)()
                    result.Success = False
                    Return result
                End If

                ' Check for remaining at root level (common in status responses)
                If jObj("remaining") IsNot Nothing Then
                    result.ActivationsRemaining = jObj("remaining").Value(Of Integer)()
                End If

                ' Extract data object fields (common to all actions)
                If jObj("data") IsNot Nothing Then
                    Dim data = jObj("data")

                    ' Core activation counts
                    If data("total_activations_purchased") IsNot Nothing Then
                        result.TotalActivationsPurchased = data("total_activations_purchased").Value(Of Integer)()
                    End If
                    If data("total_activations") IsNot Nothing Then
                        result.TotalActivations = data("total_activations").Value(Of Integer)()
                    End If
                    If data("activations_remaining") IsNot Nothing Then
                        result.ActivationsRemaining = data("activations_remaining").Value(Of Integer)()
                    End If

                    ' CRITICAL: Check for activated field directly in data object
                    ' The status API returns {"data":{"activated":true,...}} for activated users
                    If data("activated") IsNot Nothing Then
                        result.Activated = data("activated").Value(Of Boolean)()
                    End If

                    ' Also check if this instance is activated via activations array (alternative format)
                    If Not result.Activated AndAlso data("activations") IsNot Nothing AndAlso data("activations").Type = JTokenType.Array Then
                        Dim activationsArray = CType(data("activations"), JArray)
                        If activationsArray.Count > 0 Then
                            result.Activated = True
                        End If
                    End If

                    ' Get product info from api_key_expirations
                    If data("api_key_expirations") IsNot Nothing Then
                        Dim expirations = data("api_key_expirations")

                        ' Check for wc_subs_resources array inside api_key_expirations (subscription format)
                        If expirations("wc_subs_resources") IsNot Nothing AndAlso
                           expirations("wc_subs_resources").Type = JTokenType.Array Then
                            Dim subs = CType(expirations("wc_subs_resources"), JArray)
                            If subs.Count > 0 Then
                                Dim firstSub = subs(0)
                                If firstSub("product_title") IsNot Nothing Then
                                    result.ProductTitle = firstSub("product_title").Value(Of String)()
                                End If
                                If firstSub("product_id") IsNot Nothing Then
                                    result.ProductId = firstSub("product_id").ToString()
                                End If
                                If firstSub("next_payment") IsNot Nothing Then
                                    result.NextPayment = firstSub("next_payment").Value(Of String)()
                                End If
                            End If
                        End If

                        ' Also check if api_key_expirations is directly an array (non-subscription format)
                        If String.IsNullOrEmpty(result.ProductTitle) AndAlso expirations.Type = JTokenType.Array Then
                            Dim expArray = CType(expirations, JArray)
                            If expArray.Count > 0 Then
                                Dim firstExp = expArray(0)
                                If firstExp("product_title") IsNot Nothing Then
                                    result.ProductTitle = firstExp("product_title").Value(Of String)()
                                End If
                                If firstExp("product_id") IsNot Nothing Then
                                    result.ProductId = firstExp("product_id").ToString()
                                End If
                            End If
                        End If
                    End If

                    ' Get subscription info from wc_subs_resources at data level (alternative location)
                    If String.IsNullOrEmpty(result.ProductTitle) AndAlso
                       data("wc_subs_resources") IsNot Nothing AndAlso
                       data("wc_subs_resources").Type = JTokenType.Array Then
                        Dim subs = CType(data("wc_subs_resources"), JArray)
                        If subs.Count > 0 Then
                            Dim firstSub = subs(0)
                            If firstSub("next_payment") IsNot Nothing Then
                                result.NextPayment = firstSub("next_payment").Value(Of String)()
                            End If
                            If firstSub("product_title") IsNot Nothing Then
                                result.ProductTitle = firstSub("product_title").Value(Of String)()
                            End If
                        End If
                    End If
                End If

                ' Parse action-specific fields at root level
                Select Case action.ToLowerInvariant()
                    Case "status"
                        If jObj("status_check") IsNot Nothing Then
                            result.StatusCheck = jObj("status_check").Value(Of String)()
                            result.Success = True
                        End If

                    Case "activate"
                        ' Check activated at root level
                        If jObj("activated") IsNot Nothing Then
                            result.Activated = jObj("activated").Value(Of Boolean)()
                            result.Success = result.Activated
                        End If

                        If jObj("instance") IsNot Nothing Then
                            result.Activated = True
                            result.Success = True
                        End If

                        ' Get product title if available at root
                        If jObj("product_title") IsNot Nothing Then
                            result.ProductTitle = jObj("product_title").Value(Of String)()
                        End If

                        ' Check for activation error message
                        If Not result.Success AndAlso jObj("error") Is Nothing Then
                            If jObj("message") IsNot Nothing Then
                                result.ErrorMessage = jObj("message").Value(Of String)()
                            ElseIf jObj("code") IsNot Nothing Then
                                result.ErrorMessage = $"Activation failed (code: {jObj("code").Value(Of String)()})"
                            Else
                                result.ErrorMessage = "Activation failed for unknown reason"
                            End If
                        End If

                    Case "deactivate"
                        If jObj("deactivated") IsNot Nothing Then
                            result.Success = jObj("deactivated").Value(Of Boolean)()
                        End If

                        ' Get product title if available at root level
                        If jObj("product_title") IsNot Nothing Then
                            result.ProductTitle = jObj("product_title").Value(Of String)()
                        End If

                        If Not result.Success AndAlso jObj("error") Is Nothing Then
                            If jObj("message") IsNot Nothing Then
                                result.ErrorMessage = jObj("message").Value(Of String)()
                            Else
                                result.ErrorMessage = "Deactivation failed for unknown reason"
                            End If
                        End If
                End Select

                ' Default product title if not found
                If String.IsNullOrEmpty(result.ProductTitle) Then
                    result.ProductTitle = "(no product title available)"
                End If

                Return result

            Catch ex As Exception
                result.Success = False
                result.ErrorMessage = $"Failed to parse API response: {ex.Message}"
                LogLicenseEvent("Parse Error", $"{ex.Message} - JSON: {TruncateForLog(json, 200)}")
                Return result
            End Try
        End Function

        ''' <summary>
        ''' Truncates a string for logging, returning at most <paramref name="maxLength"/> characters.
        ''' </summary>
        ''' <param name="value">Value to truncate.</param>
        ''' <param name="maxLength">Maximum number of characters to return before appending an ellipsis.</param>
        Private Shared Function TruncateForLog(value As String, maxLength As Integer) As String
            If String.IsNullOrEmpty(value) Then Return ""
            If value.Length <= maxLength Then Return value
            Return value.Substring(0, maxLength) & "..."
        End Function

#End Region

#Region "License Deactivation"

        ''' <summary>
        ''' Shows a confirmation dialog and deactivates the currently stored Pro license via the API.
        ''' </summary>
        ''' <returns><see langword="True"/> if deactivation succeeded; otherwise <see langword="False"/>.</returns>
        Public Shared Function ShowDeactivationDialog() As Boolean
            Try
                ' Check if we have a pro license to deactivate
                If Not HasStoredProLicense() Then
                    ShowCustomMessageBox(
                        "No active Professional License found to deactivate.",
                        $"{AN} - License")
                    Return False
                End If

                Dim productId = My.Settings.License_ProductID
                Dim licenseKey = My.Settings.License_Key
                Dim userId = My.Settings.License_UserID
                Dim productName = My.Settings.License_ProductName

                If IsOfflineDomainLicenseKey(licenseKey) Then
                    Dim localResult = ShowCustomYesNoBox(
                        "Are you sure you want to remove the offline-domain license from local storage?" & vbCrLf & vbCrLf &
                        "No online deactivation is necessary.",
                        "Remove License",
                        "Cancel",
                        $"{AN} - License Deactivation")

                    If localResult <> 1 Then
                        Return False
                    End If

                    ClearStoredLicense()

                    LogLicenseEvent("License Deactivated", "Offline-domain license removed from local storage", alwaysLog:=True)

                    ShowCustomMessageBox(
                        "Offline-domain license removed from local storage." & vbCrLf & vbCrLf &
                        "No online deactivation was necessary.",
                        $"{AN} - License Deactivated")

                    Return True
                End If

                Dim msg = $"Are you sure you want to deactivate your license?" & vbCrLf & vbCrLf &
                          $"Product: {productName}" & vbCrLf &
                          $"User ID: {userId}" & vbCrLf & vbCrLf &
                          "This will affect all add-ins using this license with this User ID." & vbCrLf &
                          "The activation slot will be freed and can be used for another User ID."

                Dim result = ShowCustomYesNoBox(msg, "Deactivate", "Cancel", $"{AN} - License Deactivation")

                If result <> 1 Then
                    Return False
                End If

                ' Perform deactivation
                Dim response = CallLicenseApi("deactivate", productId, licenseKey, userId)

                If response.Success Then
                    ' Get updated status for remaining activations
                    Dim statusResponse = CallLicenseApi("status", productId, licenseKey, userId)
                    Dim activationsRemaining = If(statusResponse.Success, statusResponse.ActivationsRemaining, 0)

                    ' Clear stored license
                    ClearStoredLicense()

                    LogLicenseEvent("License Deactivated", $"ProductID={productId}, UserID={userId}", alwaysLog:=True)

                    ShowCustomMessageBox(
                        "License deactivated successfully." & vbCrLf & vbCrLf &
                        $"User ID deactivated: {userId}" & vbCrLf &
                        $"Activation slots now available: {activationsRemaining}",
                        $"{AN} - License Deactivated")

                    Return True
                Else
                    ShowCustomMessageBox(
                        $"Failed to deactivate license: {response.ErrorMessage}" & vbCrLf & vbCrLf &
                        "The local license data has been cleared, but the server-side " &
                        "deactivation may not have completed. Contact support if needed.",
                        $"{AN} - Deactivation Warning")

                    ' Clear local data anyway
                    ClearStoredLicense()
                    Return False
                End If

            Catch ex As Exception
                LogLicenseEvent("Deactivation Error", ex.Message, alwaysLog:=True)
                ShowCustomMessageBox($"Error during deactivation: {ex.Message}", $"{AN} - License")
                Return False
            End Try
        End Function

#End Region

#Region "License Management UI"

        ''' <summary>
        ''' Shows a license management dialog and allows opening the Pro license management form.
        ''' </summary>
        Public Shared Sub ShowLicenseManagementDialog()
            Dim detailedStatus = GetDetailedLicenseStatus()
            Dim hasProLicense As Boolean = HasStoredProLicense()

            Dim answer As Integer = ShowCustomYesNoBox(
                "License Status:" & vbCrLf & vbCrLf & detailedStatus,
                "Manage Pro License",
                "Close",
                $"{AN} - License Management",
                extraButtonText:=If(hasProLicense, "Deactivate License", Nothing),
                extraButtonAction:=If(hasProLicense,
                    Sub()
                        ShowDeactivationDialog()
                    End Sub, Nothing),
                CloseAfterExtra:=True)

            If answer = 1 Then
                ' Open license dialog with stored credentials pre-filled
                Dim prefilledKey As String = ""
                Dim prefilledProductId As String = ""
                Dim prefilledUserId As String = ""

                If hasProLicense Then
                    Try
                        prefilledKey = My.Settings.License_Key
                        prefilledProductId = My.Settings.License_ProductID
                        prefilledUserId = My.Settings.License_UserID
                    Catch
                        ' Ignore errors reading stored credentials
                    End Try
                End If

                ShowProLicenseEntryForm(Nothing, prefilledKey, prefilledProductId, prefilledUserId, False)
            End If
        End Sub

        ''' <summary>
        ''' Builds a detailed textual description of the currently stored license data.
        ''' </summary>
        Private Shared Function GetDetailedLicenseStatus() As String
            Dim sb As New StringBuilder()

            sb.AppendLine($"Current State: {GetLicenseStatusShort()}")
            sb.AppendLine()

            Try
                Dim licenseType = My.Settings.License_Type

                If String.IsNullOrEmpty(licenseType) Then
                    sb.AppendLine("No license configured.")
                ElseIf licenseType.Equals("Private", StringComparison.OrdinalIgnoreCase) Then
                    sb.AppendLine("License Type: Private (Non-Commercial)")
                    sb.AppendLine($"Confirmed On: {My.Settings.License_PrivateConfirmedOn:d}")
                    sb.AppendLine($"Valid Until: {My.Settings.License_ValidUntil:d}")
                    sb.AppendLine($"Version: {My.Settings.License_PrivateVersion}")
                ElseIf licenseType.Equals("Pro", StringComparison.OrdinalIgnoreCase) Then
                    sb.AppendLine("License Type: Professional")
                    sb.AppendLine($"Product: {My.Settings.License_ProductName}")
                    sb.AppendLine($"Product ID: {My.Settings.License_ProductID}")
                    sb.AppendLine($"User ID: {My.Settings.License_UserID}")
                    sb.AppendLine($"Activated On: {My.Settings.License_ActivatedOn:d}")
                    sb.AppendLine($"Last Check: {My.Settings.License_LastCheck:d}")
                    sb.AppendLine($"API Confirmed: {My.Settings.License_ApiConfirmed}")

                    If IsOfflineDomainLicenseKey(My.Settings.License_Key) Then
                        Dim validUntilUtc As Date = Date.MinValue
                        Dim allowedDomains As New List(Of String)()
                        Dim offlineProductId As String = ""

                        If TryReadOfflineDomainLicenseMetadata(My.Settings.License_Key, validUntilUtc, allowedDomains, offlineProductId) Then
                            sb.AppendLine($"Valid Until: {validUntilUtc:yyyy-MM-dd} UTC")

                            Dim daysRemaining = CInt((validUntilUtc.Date - Date.UtcNow.Date).TotalDays)
                            sb.AppendLine($"Days Remaining: {daysRemaining}")

                            If allowedDomains.Count > 0 Then
                                sb.AppendLine($"Allowed Network IDs: {String.Join(", ", allowedDomains)}")
                            End If
                        End If
                    End If
                End If

            Catch ex As Exception
                sb.AppendLine($"Error reading license data: {ex.Message}")
            End Try

            Return sb.ToString()
        End Function

#End Region

    End Class
End Namespace
