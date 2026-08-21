' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' Purpose: Debug-only interactive test helpers for license system.

#If DEBUG Then

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms

Namespace SharedLibrary

    Partial Public Class SharedMethods

#Region "Interactive Test Entry Point"

        ''' <summary>
        ''' Main test entry point - shows a test control panel for license system testing.
        ''' Call this from any command handler to access all test functions.
        ''' </summary>
        Public Shared Sub TestLicenseSystem()
            Using form As New Form()
                form.AutoScaleMode = AutoScaleMode.Dpi
                form.AutoScaleDimensions = New SizeF(96.0F, 96.0F)
                form.Text = $"{AN} - License System Test Panel"
                form.Size = New Size(620, 700)
                form.StartPosition = FormStartPosition.CenterScreen
                form.FormBorderStyle = FormBorderStyle.FixedDialog
                form.MaximizeBox = False
                form.MinimizeBox = False
                form.ShowInTaskbar = True
                form.TopMost = True

                Dim font As New Font("Segoe UI", 9.5F)
                form.Font = font

                ' Main layout
                Dim mainLayout As New TableLayoutPanel() With {
                    .Dock = DockStyle.Fill,
                    .Padding = New Padding(15),
                    .ColumnCount = 1,
                    .RowCount = 3
                }
                mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
                mainLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
                mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
                form.Controls.Add(mainLayout)

                ' Title
                Dim lblTitle As New Label() With {
                    .Text = "License System Test Panel",
                    .Font = New Font("Segoe UI", 14.0F, FontStyle.Bold),
                    .AutoSize = True,
                    .Margin = New Padding(0, 0, 0, 15)
                }
                mainLayout.Controls.Add(lblTitle, 0, 0)

                ' Test buttons panel
                Dim buttonPanel As New FlowLayoutPanel() With {
                    .Dock = DockStyle.Fill,
                    .FlowDirection = FlowDirection.TopDown,
                    .AutoScroll = True,
                    .WrapContents = False,
                    .Padding = New Padding(0)
                }
                mainLayout.Controls.Add(buttonPanel, 0, 1)

                ' Status section
                Dim lblStatus As New Label() With {
                    .Text = "Ready",
                    .Dock = DockStyle.Fill,
                    .ForeColor = Color.DarkBlue,
                    .AutoSize = True,
                    .Margin = New Padding(0, 10, 0, 0)
                }

                ' Add test buttons
                AddTestButton(buttonPanel, "📊 Show Current License Status", Sub() TestShowCurrentStatus(lblStatus))
                AddTestButton(buttonPanel, "🗑️ Clear All License Data", Sub() TestClearAllData(lblStatus))
                AddTestButton(buttonPanel, "📝 Set License State", Sub() TestSetLicenseState(lblStatus))
                AddTestButton(buttonPanel, "🔐 Set Private License", Sub() TestSetPrivateLicense(lblStatus))
                AddTestButton(buttonPanel, "💼 Set Pro License", Sub() TestSetProLicense(lblStatus))
                AddTestButton(buttonPanel, "🌐 Test API Call", Sub() TestApiCall(lblStatus))
                AddTestButton(buttonPanel, "📄 Parse API Response", Sub() TestParseApiResponseInteractive(lblStatus))
                AddTestButton(buttonPanel, "⏰ Set Expiry Date", Sub() TestSetExpiryDate(lblStatus))
                AddTestButton(buttonPanel, "📅 Set Offline Grace Start", Sub() TestSetOfflineGraceStart(lblStatus))
                AddTestButton(buttonPanel, "🔄 Test Reconfirmation Check", Sub() TestReconfirmationCheck(lblStatus))
                AddTestButton(buttonPanel, "🚀 Run Full LicenseOK Flow", Sub() TestFullLicenseOKFlow(lblStatus))
                AddTestButton(buttonPanel, "📋 Dump All Settings", Sub() TestDumpAllSettings(lblStatus))
                AddTestButton(buttonPanel, "📁 Open Log File", Sub() TestOpenLogFile(lblStatus))

                mainLayout.Controls.Add(lblStatus, 0, 2)

                AddHandler form.Shown,
                    Sub()
                        form.TopMost = True
                        Global.SharedLibrary.SharedLibrary.SharedMethods.ForceDialogToForeground(form)
                        Global.SharedLibrary.SharedLibrary.SharedMethods.AttachForeignForegroundWatchdog(form)
                    End Sub
                Dim __safeDialogOwner97 As System.Windows.Forms.IWin32Window = Global.SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
                If __safeDialogOwner97 IsNot Nothing Then
                    form.ShowDialog(__safeDialogOwner97)
                Else
                    form.ShowDialog()
                End If
            End Using
        End Sub

        Private Shared Sub AddTestButton(panel As FlowLayoutPanel, text As String, action As Action)
            Dim btn As New Button() With {
                .Text = text,
                .Font = New Font("Segoe UI Emoji", 9.5F),
                .AutoSize = False,
                .Size = New Size(560, 32),
                .Padding = New Padding(0),
                .Margin = New Padding(0, 0, 0, 5),
                .TextAlign = ContentAlignment.MiddleLeft,
                .FlatStyle = FlatStyle.System
            }
            AddHandler btn.Click, Sub(s, e)
                                      Try
                                          action()
                                      Catch ex As Exception
                                          MessageBox.Show($"Error: {ex.Message}", "Test Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                      End Try
                                  End Sub
            panel.Controls.Add(btn)
        End Sub

#End Region

#Region "Individual Test Functions"
        Private Shared Sub TestShowCurrentStatus(lbl As Label)
            Dim sb As New StringBuilder()
            sb.AppendLine("=== CURRENT LICENSE STATUS ===")
            sb.AppendLine()
            sb.AppendLine($"State: {_currentLicenseState}")
            sb.AppendLine($"Status String: {GetLicenseStatusShort()}")
            sb.AppendLine()

            Try
                sb.AppendLine("--- My.Settings ---")
                sb.AppendLine($"License_Type: {My.Settings.License_Type}")
                sb.AppendLine($"License_ProductID: {My.Settings.License_ProductID}")
                sb.AppendLine($"License_Key: {If(String.IsNullOrEmpty(My.Settings.License_Key), "(empty)", "***")}")
                sb.AppendLine($"License_UserID: {My.Settings.License_UserID}")
                sb.AppendLine($"License_ProductName: {My.Settings.License_ProductName}")
                sb.AppendLine($"License_ActivatedOn: {My.Settings.License_ActivatedOn:d}")
                sb.AppendLine($"License_ValidUntil: {My.Settings.License_ValidUntil:d}")
                sb.AppendLine($"License_LastCheck: {My.Settings.License_LastCheck:d}")
                sb.AppendLine($"License_ApiConfirmed: {My.Settings.License_ApiConfirmed}")
                sb.AppendLine($"License_PrivateConfirmedOn: {My.Settings.License_PrivateConfirmedOn:d}")
                sb.AppendLine($"License_PrivateVersion: {My.Settings.License_PrivateVersion}")
                sb.AppendLine($"License_PrivateDismissCount: {My.Settings.License_PrivateDismissCount}")
                sb.AppendLine($"License_OfflineGraceStart: {My.Settings.License_OfflineGraceStart:d}")
                sb.AppendLine($"License_GracePeriodStart: {My.Settings.License_GracePeriodStart:d}")
                sb.AppendLine($"License_State: {My.Settings.License_State}")
                sb.AppendLine($"License_AutoActivationWarningShown: {My.Settings.License_AutoActivationWarningShown}")
                sb.AppendLine($"License_StartupCount: {My.Settings.License_StartupCount}")
            Catch ex As Exception
                sb.AppendLine($"Error reading settings: {ex.Message}")
            End Try

            ShowTestResult("License Status", sb.ToString())
            lbl.Text = "Status displayed"
            lbl.ForeColor = Color.DarkGreen
        End Sub

        Private Shared Sub TestClearAllData(lbl As Label)
            Dim result = MessageBox.Show("This will clear ALL license data. Continue?",
                                          "Confirm Clear", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If result = DialogResult.Yes Then
                ClearStoredLicense()
                LogLicenseEvent("TEST", "All license data cleared", alwaysLog:=True)
                lbl.Text = "All license data cleared"
                lbl.ForeColor = Color.DarkGreen
            Else
                lbl.Text = "Clear cancelled"
                lbl.ForeColor = Color.DarkOrange
            End If
        End Sub

        Private Shared Sub TestSetLicenseState(lbl As Label)
            Dim states = [Enum].GetValues(GetType(LicenseState))
            Dim items As New List(Of String)
            For Each s In states
                items.Add(s.ToString())
            Next

            Dim selected = ShowListSelector("Select License State", "Choose a state to set:", items)
            If Not String.IsNullOrEmpty(selected) Then
                Dim state = DirectCast([Enum].Parse(GetType(LicenseState), selected), LicenseState)
                _currentLicenseState = state
                LogLicenseEvent("TEST", $"State forced to: {state}", alwaysLog:=True)
                lbl.Text = $"State set to: {state}"
                lbl.ForeColor = Color.DarkGreen
            Else
                lbl.Text = "State selection cancelled"
                lbl.ForeColor = Color.DarkOrange
            End If
        End Sub


        Private Shared Sub TestSetPrivateLicense(lbl As Label)
            ' First, ask what scenario to set up
            Dim scenarios = New List(Of String) From {
                "Active (no reconfirmation needed)",
                "Needs reconfirmation (version changed)",
                "Needs reconfirmation (interval elapsed)",
                "Needs reconfirmation (dismissals available)",
                "Needs reconfirmation (no dismissals left)",
                "Expired (within grace period)",
                "Expired (past grace period)"
            }

            Dim selected = ShowListSelector("Private License Scenario", "Select the test scenario:", scenarios)
            If String.IsNullOrEmpty(selected) Then
                lbl.Text = "Private license cancelled"
                lbl.ForeColor = Color.DarkOrange
                Return
            End If

            ' Get current version from context or use default
            Dim currentVersion = If(_licenseContext IsNot Nothing, _licenseContext.RDV, "Word (V.240115)")
            Dim differentVersion = "Word (V.200101)" ' Old version for version-change scenario

            Select Case selected
                Case "Active (no reconfirmation needed)"
                    ' Recently confirmed, same version, no dismissals
                    My.Settings.License_Type = "Private"
                    My.Settings.License_PrivateConfirmedOn = Date.Now.Date
                    My.Settings.License_PrivateVersion = currentVersion
                    My.Settings.License_PrivateDismissCount = 0
                    My.Settings.License_ValidUntil = Date.Now.AddYears(1)
                    _currentLicenseState = LicenseState.PrivateActive

                Case "Needs reconfirmation (version changed)"
                    ' Confirmed recently but version is different
                    My.Settings.License_Type = "Private"
                    My.Settings.License_PrivateConfirmedOn = Date.Now.Date
                    My.Settings.License_PrivateVersion = differentVersion ' Different from current
                    My.Settings.License_PrivateDismissCount = 0
                    My.Settings.License_ValidUntil = Date.Now.AddYears(1)
                    _currentLicenseState = LicenseState.PrivateReconfirmNeeded

                Case "Needs reconfirmation (interval elapsed)"
                    ' Confirmed 31+ days ago, same version
                    My.Settings.License_Type = "Private"
                    My.Settings.License_PrivateConfirmedOn = Date.Now.AddDays(-35) ' Over 30 days
                    My.Settings.License_PrivateVersion = currentVersion
                    My.Settings.License_PrivateDismissCount = 0
                    My.Settings.License_ValidUntil = Date.Now.AddYears(1)
                    _currentLicenseState = LicenseState.PrivateReconfirmNeeded

                Case "Needs reconfirmation (dismissals available)"
                    ' Interval elapsed, 1 dismissal used (2 remaining of 3 max)
                    My.Settings.License_Type = "Private"
                    My.Settings.License_PrivateConfirmedOn = Date.Now.AddDays(-35)
                    My.Settings.License_PrivateVersion = currentVersion
                    My.Settings.License_PrivateDismissCount = 1 ' 2 dismissals remaining
                    My.Settings.License_ValidUntil = Date.Now.AddYears(1)
                    _currentLicenseState = LicenseState.PrivateReconfirmNeeded

                Case "Needs reconfirmation (no dismissals left)"
                    ' Interval elapsed, all dismissals used
                    My.Settings.License_Type = "Private"
                    My.Settings.License_PrivateConfirmedOn = Date.Now.AddDays(-35)
                    My.Settings.License_PrivateVersion = currentVersion
                    My.Settings.License_PrivateDismissCount = 3 ' Max reached
                    My.Settings.License_ValidUntil = Date.Now.AddYears(1)
                    _currentLicenseState = LicenseState.PrivateReconfirmNeeded

                Case "Expired (within grace period)"
                    ' Expired 3 days ago (within 7-day grace)
                    My.Settings.License_Type = "Private"
                    My.Settings.License_PrivateConfirmedOn = Date.Now.AddDays(-60)
                    My.Settings.License_PrivateVersion = currentVersion
                    My.Settings.License_PrivateDismissCount = 0
                    My.Settings.License_ValidUntil = Date.Now.AddDays(-3)
                    _currentLicenseState = LicenseState.PrivateExpired

                Case "Expired (past grace period)"
                    ' Expired 15 days ago (past grace)
                    My.Settings.License_Type = "Private"
                    My.Settings.License_PrivateConfirmedOn = Date.Now.AddDays(-90)
                    My.Settings.License_PrivateVersion = currentVersion
                    My.Settings.License_PrivateDismissCount = 0
                    My.Settings.License_ValidUntil = Date.Now.AddDays(-15)
                    _currentLicenseState = LicenseState.PrivateExpired
            End Select

            My.Settings.Save()

            LogLicenseEvent("TEST", $"Private license scenario set: {selected}", alwaysLog:=True)
            lbl.Text = $"Set: {selected}"
            lbl.ForeColor = Color.DarkGreen
        End Sub

        Private Shared Sub TestSetProLicense(lbl As Label)
            Dim productId = ShowInputBox("Product ID", "Enter Product ID:", "12345")
            If productId Is Nothing Then
                lbl.Text = "Pro license cancelled"
                lbl.ForeColor = Color.DarkOrange
                Return
            End If

            Dim licenseKey = ShowInputBox("License Key", "Enter License Key:", "test-key-12345")
            If licenseKey Is Nothing Then
                lbl.Text = "Pro license cancelled"
                lbl.ForeColor = Color.DarkOrange
                Return
            End If

            Dim userId = ShowInputBox("User ID", "Enter User ID (supports %USERNAME%, etc.):", "%USERNAME%-%COMPUTERNAME%")
            If userId Is Nothing Then
                lbl.Text = "Pro license cancelled"
                lbl.ForeColor = Color.DarkOrange
                Return
            End If

            userId = Environment.ExpandEnvironmentVariables(userId)

            My.Settings.License_Type = "Pro"
            My.Settings.License_ProductID = productId
            My.Settings.License_Key = licenseKey
            My.Settings.License_UserID = userId
            My.Settings.License_ProductName = "Test Pro License"
            My.Settings.License_ActivatedOn = Date.Now.Date
            My.Settings.License_LastCheck = Date.Now.Date
            My.Settings.License_ApiConfirmed = True
            My.Settings.License_OfflineGraceStart = Date.MinValue
            My.Settings.License_GracePeriodStart = Date.MinValue
            My.Settings.Save()

            _currentLicenseState = LicenseState.ProActive
            LogLicenseEvent("TEST", $"Pro license set: {productId} / {userId}", alwaysLog:=True)
            lbl.Text = $"Pro license set: {productId} / {userId}"
            lbl.ForeColor = Color.DarkGreen
        End Sub

        Private Shared Sub TestApiCall(lbl As Label)
            Dim actions = New List(Of String) From {"status", "activate", "deactivate"}
            Dim action = ShowListSelector("Select API Action", "Choose an action:", actions)
            If String.IsNullOrEmpty(action) Then
                lbl.Text = "API call cancelled"
                lbl.ForeColor = Color.DarkOrange
                Return
            End If

            Dim productId = ShowInputBox("Product ID", "Enter Product ID:", My.Settings.License_ProductID)
            If productId Is Nothing Then Return

            Dim licenseKey = ShowInputBox("License Key", "Enter License Key:", My.Settings.License_Key)
            If licenseKey Is Nothing Then Return

            Dim userId = ShowInputBox("User ID", "Enter User ID:", My.Settings.License_UserID)
            If userId Is Nothing Then Return

            lbl.Text = "Calling API..."
            lbl.ForeColor = Color.DarkBlue
            Application.DoEvents()

            Try
                Dim response = CallLicenseApi(action, productId, licenseKey, userId)

                Dim sb As New StringBuilder()
                sb.AppendLine($"=== API RESPONSE ({action}) ===")
                sb.AppendLine()
                sb.AppendLine($"Success: {response.Success}")
                sb.AppendLine($"StatusCheck: {response.StatusCheck}")
                sb.AppendLine($"Activated: {response.Activated}")
                sb.AppendLine($"ProductTitle: {response.ProductTitle}")
                sb.AppendLine($"TotalActivations: {response.TotalActivations}")
                sb.AppendLine($"TotalActivationsPurchased: {response.TotalActivationsPurchased}")
                sb.AppendLine($"ActivationsRemaining: {response.ActivationsRemaining}")
                sb.AppendLine($"ErrorMessage: {response.ErrorMessage}")
                sb.AppendLine()
                sb.AppendLine("--- Raw JSON ---")
                sb.AppendLine(response.RawJson)

                ShowTestResult($"API Response - {action}", sb.ToString())
                lbl.Text = $"API call complete: {If(response.Success, "Success", "Failed")}"
                lbl.ForeColor = If(response.Success, Color.DarkGreen, Color.DarkRed)
            Catch ex As Exception
                lbl.Text = $"API error: {ex.Message}"
                lbl.ForeColor = Color.DarkRed
            End Try
        End Sub

        Private Shared Sub TestParseApiResponseInteractive(lbl As Label)
            Dim actions = New List(Of String) From {"status", "activate", "deactivate"}
            Dim action = ShowListSelector("Select Action Type", "Choose the response type to parse:", actions)
            If String.IsNullOrEmpty(action) Then
                lbl.Text = "Parse cancelled"
                lbl.ForeColor = Color.DarkOrange
                Return
            End If

            Dim sampleJson = ""
            Select Case action
                Case "status"
                    sampleJson = "{""success"":true,""status_check"":""active"",""data"":{""total_activations_purchased"":5,""total_activations"":2,""activations_remaining"":3}}"
                Case "activate"
                    sampleJson = "{""activated"":true,""instance"":""test-user"",""product_title"":""Test Product""}"
                Case "deactivate"
                    sampleJson = "{""deactivated"":true}"
            End Select

            Dim json = ShowMultilineInputBox("Enter JSON", "Enter or paste the JSON response to parse:", sampleJson)
            If json Is Nothing Then
                lbl.Text = "Parse cancelled"
                lbl.ForeColor = Color.DarkOrange
                Return
            End If

            Try
                Dim response = ParseLicenseApiResponse(json, action)

                Dim sb As New StringBuilder()
                sb.AppendLine($"=== PARSED RESULT ({action}) ===")
                sb.AppendLine()
                sb.AppendLine($"Success: {response.Success}")
                sb.AppendLine($"StatusCheck: {response.StatusCheck}")
                sb.AppendLine($"Activated: {response.Activated}")
                sb.AppendLine($"ProductTitle: {response.ProductTitle}")
                sb.AppendLine($"ProductId: {response.ProductId}")
                sb.AppendLine($"TotalActivations: {response.TotalActivations}")
                sb.AppendLine($"TotalActivationsPurchased: {response.TotalActivationsPurchased}")
                sb.AppendLine($"ActivationsRemaining: {response.ActivationsRemaining}")
                sb.AppendLine($"NextPayment: {response.NextPayment}")
                sb.AppendLine($"ErrorMessage: {response.ErrorMessage}")

                ShowTestResult("Parse Result", sb.ToString())
                lbl.Text = $"Parse complete: {If(response.Success, "Valid", "Invalid/Error")}"
                lbl.ForeColor = If(response.Success, Color.DarkGreen, Color.DarkRed)
            Catch ex As Exception
                lbl.Text = $"Parse error: {ex.Message}"
                lbl.ForeColor = Color.DarkRed
            End Try
        End Sub

        Private Shared Sub TestSetExpiryDate(lbl As Label)
            Dim daysStr = ShowInputBox("Days Until Expiry", "Enter days from now (negative = expired):", "30")
            If daysStr Is Nothing Then
                lbl.Text = "Expiry setting cancelled"
                lbl.ForeColor = Color.DarkOrange
                Return
            End If

            Dim days As Integer
            If Not Integer.TryParse(daysStr, days) Then days = 30

            Dim expiryDate = Date.Now.AddDays(days)
            My.Settings.License_ValidUntil = expiryDate
            My.Settings.Save()

            LogLicenseEvent("TEST", $"Expiry set to: {expiryDate:d}", alwaysLog:=True)
            lbl.Text = $"Expiry set to: {expiryDate:d}"
            lbl.ForeColor = Color.DarkGreen
        End Sub

        Private Shared Sub TestSetOfflineGraceStart(lbl As Label)
            Dim daysStr = ShowInputBox("Days Ago", "Enter days ago (0 = today):", "0")
            If daysStr Is Nothing Then
                lbl.Text = "Offline grace cancelled"
                lbl.ForeColor = Color.DarkOrange
                Return
            End If

            Dim days As Integer
            If Not Integer.TryParse(daysStr, days) Then days = 0

            Dim graceStart = Date.Now.AddDays(-days)
            My.Settings.License_OfflineGraceStart = graceStart
            My.Settings.License_ApiConfirmed = False
            My.Settings.Save()

            LogLicenseEvent("TEST", $"Offline grace started: {graceStart:d}", alwaysLog:=True)
            lbl.Text = $"Offline grace started: {graceStart:d} ({days} days ago)"
            lbl.ForeColor = Color.DarkGreen
        End Sub

        Private Shared Sub TestReconfirmationCheck(lbl As Label)
            If _licenseContext Is Nothing Then
                lbl.Text = "No context available - run LicenseOK first"
                lbl.ForeColor = Color.DarkRed
                Return
            End If

            Dim needsReconfirm = NeedsPrivateReconfirmation(_licenseContext)
            lbl.Text = $"Needs reconfirmation: {needsReconfirm}"
            lbl.ForeColor = If(needsReconfirm, Color.DarkOrange, Color.DarkGreen)
            LogLicenseEvent("TEST", $"Reconfirmation check: {needsReconfirm}", alwaysLog:=True)
        End Sub

        Private Shared Sub TestFullLicenseOKFlow(lbl As Label)
            If _licenseContext Is Nothing OrElse _licenseConfigDict Is Nothing Then
                lbl.Text = "No context/config - call from add-in"
                lbl.ForeColor = Color.DarkRed
                Return
            End If

            lbl.Text = "Running LicenseOK..."
            lbl.ForeColor = Color.DarkBlue
            Application.DoEvents()

            Try
                Dim result = LicenseOK(_licenseContext, _licenseConfigDict)
                lbl.Text = $"LicenseOK returned: {result} (State: {_currentLicenseState})"
                lbl.ForeColor = If(result, Color.DarkGreen, Color.DarkRed)
                LogLicenseEvent("TEST", $"Full LicenseOK flow: {result}", alwaysLog:=True)
            Catch ex As Exception
                lbl.Text = $"LicenseOK error: {ex.Message}"
                lbl.ForeColor = Color.DarkRed
            End Try
        End Sub

        Private Shared Sub TestDumpAllSettings(lbl As Label)
            Dim sb As New StringBuilder()
            sb.AppendLine("=== ALL LICENSE SETTINGS ===")
            sb.AppendLine($"Generated: {Date.Now:yyyy-MM-dd HH:mm:ss}")
            sb.AppendLine()

            Try
                For Each prop In GetType(My.MySettings).GetProperties()
                    If prop.Name.StartsWith("License", StringComparison.OrdinalIgnoreCase) Then
                        Try
                            Dim value = prop.GetValue(My.Settings)
                            Dim displayValue = If(prop.Name.Contains("Key"), "***", value?.ToString())
                            sb.AppendLine($"{prop.Name}: {displayValue}")
                        Catch
                            sb.AppendLine($"{prop.Name}: (error reading)")
                        End Try
                    End If
                Next
            Catch ex As Exception
                sb.AppendLine($"Error: {ex.Message}")
            End Try

            ShowTestResult("All Settings Dump", sb.ToString())
            lbl.Text = "Settings dumped"
            lbl.ForeColor = Color.DarkGreen
        End Sub

        Private Shared Sub TestOpenLogFile(lbl As Label)
            Try
                Dim logPath = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "redink",
            LogFileName)

                If System.IO.File.Exists(logPath) Then
                    Process.Start(New ProcessStartInfo(logPath) With {.UseShellExecute = True})
                    lbl.Text = "Log file opened"
                    lbl.ForeColor = Color.DarkGreen
                Else
                    lbl.Text = $"Log file not found: {logPath}"
                    lbl.ForeColor = Color.DarkOrange
                End If
            Catch ex As Exception
                lbl.Text = $"Error: {ex.Message}"
                lbl.ForeColor = Color.DarkRed
            End Try
        End Sub

#End Region

#Region "UI Helpers"

        Private Shared Function ShowInputBox(title As String, prompt As String, defaultValue As String) As String
            Using form As New Form()
                form.AutoScaleMode = AutoScaleMode.Dpi
                form.AutoScaleDimensions = New SizeF(96.0F, 96.0F)
                form.Text = title
                form.ClientSize = New Size(400, 130)
                form.StartPosition = FormStartPosition.CenterScreen
                form.FormBorderStyle = FormBorderStyle.FixedDialog
                form.MaximizeBox = False
                form.MinimizeBox = False
                form.TopMost = True
                form.Font = New Font("Segoe UI", 9.5F)

                Dim lbl As New Label() With {
            .Text = prompt,
            .Location = New Point(15, 15),
            .AutoSize = True
        }
                Dim txt As New TextBox() With {
            .Location = New Point(15, 40),
            .Size = New Size(370, 23),
            .Text = defaultValue,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        }

                Dim pnlButtons As New Panel() With {
            .Dock = DockStyle.Bottom,
            .Height = 45
        }
                Dim btnOk As New Button() With {
            .Text = "OK",
            .Size = New Size(80, 28),
            .Anchor = AnchorStyles.Bottom Or AnchorStyles.Right,
            .DialogResult = DialogResult.OK
        }
                Dim btnCancel As New Button() With {
            .Text = "Cancel",
            .Size = New Size(80, 28),
            .Anchor = AnchorStyles.Bottom Or AnchorStyles.Right,
            .DialogResult = DialogResult.Cancel
        }

                ' Position buttons relative to panel size
                btnCancel.Location = New Point(pnlButtons.ClientSize.Width - btnCancel.Width - 10, 8)
                btnOk.Location = New Point(btnCancel.Left - btnOk.Width - 8, 8)
                pnlButtons.Controls.AddRange({btnOk, btnCancel})

                form.Controls.Add(txt)
                form.Controls.Add(lbl)
                form.Controls.Add(pnlButtons)
                form.AcceptButton = btnOk
                form.CancelButton = btnCancel

                AddHandler form.Shown,
                    Sub()
                        form.TopMost = True
                        Global.SharedLibrary.SharedLibrary.SharedMethods.ForceDialogToForeground(form)
                        Global.SharedLibrary.SharedLibrary.SharedMethods.AttachForeignForegroundWatchdog(form)
                    End Sub
                Dim __safeDialogOwner615 As System.Windows.Forms.IWin32Window = Global.SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
                If If(__safeDialogOwner615 IsNot Nothing, form.ShowDialog(__safeDialogOwner615), form.ShowDialog()) = DialogResult.OK Then
                    Return txt.Text
                End If
                Return Nothing
            End Using
        End Function

        Private Shared Function ShowMultilineInputBox(title As String, prompt As String, defaultValue As String) As String
            Using form As New Form()
                form.AutoScaleMode = AutoScaleMode.Dpi
                form.AutoScaleDimensions = New SizeF(96.0F, 96.0F)
                form.Text = title
                form.ClientSize = New Size(500, 300)
                form.StartPosition = FormStartPosition.CenterScreen
                form.FormBorderStyle = FormBorderStyle.Sizable
                form.MinimumSize = New Size(400, 200)
                form.TopMost = True
                form.Font = New Font("Segoe UI", 9.5F)

                Dim lbl As New Label() With {
                    .Text = prompt,
                    .Dock = DockStyle.Top,
                    .AutoSize = True,
                    .Padding = New Padding(10)
                }

                Dim pnlButtons As New Panel() With {
                    .Dock = DockStyle.Bottom,
                    .Height = 45
                }
                Dim btnOk As New Button() With {
                    .Text = "OK",
                    .Size = New Size(80, 28),
                    .Anchor = AnchorStyles.Bottom Or AnchorStyles.Right,
                    .DialogResult = DialogResult.OK
                }
                Dim btnCancel As New Button() With {
                    .Text = "Cancel",
                    .Size = New Size(80, 28),
                    .Anchor = AnchorStyles.Bottom Or AnchorStyles.Right,
                    .DialogResult = DialogResult.Cancel
                }
                btnCancel.Location = New Point(pnlButtons.ClientSize.Width - btnCancel.Width - 10, 8)
                btnOk.Location = New Point(btnCancel.Left - btnOk.Width - 8, 8)
                pnlButtons.Controls.AddRange({btnOk, btnCancel})

                Dim txt As New TextBox() With {
                    .Dock = DockStyle.Fill,
                    .Multiline = True,
                    .ScrollBars = ScrollBars.Both,
                    .Text = defaultValue,
                    .Font = New Font("Consolas", 9.0F)
                }

                form.Controls.Add(txt)
                form.Controls.Add(lbl)
                form.Controls.Add(pnlButtons)
                form.AcceptButton = btnOk
                form.CancelButton = btnCancel

                AddHandler form.Shown,
                    Sub()
                        form.TopMost = True
                        Global.SharedLibrary.SharedLibrary.SharedMethods.ForceDialogToForeground(form)
                        Global.SharedLibrary.SharedLibrary.SharedMethods.AttachForeignForegroundWatchdog(form)
                    End Sub
                Dim __safeDialogOwner675 As System.Windows.Forms.IWin32Window = Global.SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
                If If(__safeDialogOwner675 IsNot Nothing, form.ShowDialog(__safeDialogOwner675), form.ShowDialog()) = DialogResult.OK Then
                    Return txt.Text
                End If
                Return Nothing
            End Using
        End Function

        Private Shared Function ShowListSelector(title As String, prompt As String, items As List(Of String)) As String
            Using form As New Form()
                form.AutoScaleMode = AutoScaleMode.Dpi
                form.AutoScaleDimensions = New SizeF(96.0F, 96.0F)
                form.Text = title
                form.ClientSize = New Size(350, 300)
                form.StartPosition = FormStartPosition.CenterScreen
                form.FormBorderStyle = FormBorderStyle.FixedDialog
                form.MaximizeBox = False
                form.MinimizeBox = False
                form.TopMost = True
                form.Font = New Font("Segoe UI", 9.5F)

                Dim lbl As New Label() With {
            .Text = prompt,
            .Dock = DockStyle.Top,
            .AutoSize = True,
            .Padding = New Padding(10)
        }

                Dim pnlButtons As New Panel() With {
            .Dock = DockStyle.Bottom,
            .Height = 45
        }
                Dim btnOk As New Button() With {
            .Text = "OK",
            .Size = New Size(80, 28),
            .Anchor = AnchorStyles.Bottom Or AnchorStyles.Right,
            .DialogResult = DialogResult.OK
        }
                Dim btnCancel As New Button() With {
            .Text = "Cancel",
            .Size = New Size(80, 28),
            .Anchor = AnchorStyles.Bottom Or AnchorStyles.Right,
            .DialogResult = DialogResult.Cancel
        }

                ' Position buttons relative to panel size
                btnCancel.Location = New Point(pnlButtons.ClientSize.Width - btnCancel.Width - 10, 8)
                btnOk.Location = New Point(btnCancel.Left - btnOk.Width - 8, 8)
                pnlButtons.Controls.AddRange({btnOk, btnCancel})

                Dim lst As New ListBox() With {
            .Dock = DockStyle.Fill
        }
                For Each item In items
                    lst.Items.Add(item)
                Next
                If lst.Items.Count > 0 Then lst.SelectedIndex = 0

                form.Controls.Add(lst)
                form.Controls.Add(lbl)
                form.Controls.Add(pnlButtons)
                form.AcceptButton = btnOk
                form.CancelButton = btnCancel

                AddHandler lst.DoubleClick, Sub() form.DialogResult = DialogResult.OK

                AddHandler form.Shown,
                    Sub()
                        form.TopMost = True
                        Global.SharedLibrary.SharedLibrary.SharedMethods.ForceDialogToForeground(form)
                        Global.SharedLibrary.SharedLibrary.SharedMethods.AttachForeignForegroundWatchdog(form)
                    End Sub
                Dim __safeDialogOwner740 As System.Windows.Forms.IWin32Window = Global.SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
                If If(__safeDialogOwner740 IsNot Nothing, form.ShowDialog(__safeDialogOwner740), form.ShowDialog()) = DialogResult.OK AndAlso lst.SelectedItem IsNot Nothing Then
                    Return lst.SelectedItem.ToString()
                End If
                Return Nothing
            End Using
        End Function

        Private Shared Sub ShowTestResult(title As String, content As String)
            Using form As New Form()
                form.AutoScaleMode = AutoScaleMode.Dpi
                form.AutoScaleDimensions = New SizeF(96.0F, 96.0F)
                form.Text = title
                form.ClientSize = New Size(600, 500)
                form.StartPosition = FormStartPosition.CenterScreen
                form.FormBorderStyle = FormBorderStyle.Sizable
                form.MinimumSize = New Size(400, 300)
                form.TopMost = True
                form.Font = New Font("Segoe UI", 9.5F)

                Dim pnlButtons As New Panel() With {
            .Dock = DockStyle.Bottom,
            .Height = 45
        }
                Dim btnCopy As New Button() With {
            .Text = "Copy to Clipboard",
            .Size = New Size(120, 28),
            .Anchor = AnchorStyles.Bottom Or AnchorStyles.Right,
            .DialogResult = DialogResult.None
        }
                Dim btnClose As New Button() With {
            .Text = "Close",
            .Size = New Size(80, 28),
            .Anchor = AnchorStyles.Bottom Or AnchorStyles.Right,
            .DialogResult = DialogResult.Cancel
        }

                ' Position buttons relative to panel size
                btnClose.Location = New Point(pnlButtons.ClientSize.Width - btnClose.Width - 10, 8)
                btnCopy.Location = New Point(btnClose.Left - btnCopy.Width - 8, 8)

                AddHandler btnCopy.Click, Sub()
                                              Try
                                                  Clipboard.SetText(content)
                                                  MessageBox.Show("Copied!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                              Catch
                                              End Try
                                          End Sub

                pnlButtons.Controls.AddRange({btnCopy, btnClose})

                Dim txt As New TextBox() With {
            .Dock = DockStyle.Fill,
            .Multiline = True,
            .ReadOnly = True,
            .ScrollBars = ScrollBars.Both,
            .Text = content,
            .Font = New Font("Consolas", 9.5F),
            .BackColor = SystemColors.Window
        }

                form.Controls.Add(txt)
                form.Controls.Add(pnlButtons)
                form.CancelButton = btnClose

                AddHandler form.Shown,
                    Sub()
                        form.TopMost = True
                        Global.SharedLibrary.SharedLibrary.SharedMethods.ForceDialogToForeground(form)
                        Global.SharedLibrary.SharedLibrary.SharedMethods.AttachForeignForegroundWatchdog(form)
                    End Sub
                Dim __safeDialogOwner804 As System.Windows.Forms.IWin32Window = Global.SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
                If __safeDialogOwner804 IsNot Nothing Then
                    form.ShowDialog(__safeDialogOwner804)
                Else
                    form.ShowDialog()
                End If
            End Using
        End Sub

#End Region

    End Class
End Namespace

#End If
