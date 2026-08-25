' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.License.Admin.vb
' Purpose: Organization bulk license-management UI for existing online Pro
'          licenses, driven by pasted backend page text and the existing
'          WooCommerce/Kestrel license API interface.
'
' Architecture:
'  - Entry point: `ShowOrganizationBulkLicenseManagementDialog`
'  - Parser: deterministic extraction of one backend license page containing
'    license key, product rows, and activation instances.
'  - Verification: reuses existing `CallLicenseApi("status" | "activate" | "deactivate")`
'    and `ParseLicenseApiResponse` behavior through the shared partial class.
'  - UI: direct WinForms construction, DPI-aware, DataGridView-centered workflow.
'  - Persistence: only user-ID mapping rules are persisted in `My.Settings`;
'    pasted backend text and imported desired lists remain session-only.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.ComponentModel
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.VisualBasic.FileIO
Imports Newtonsoft.Json.Linq

Namespace SharedLibrary
    Partial Public Class SharedMethods

#Region "Organization Bulk License Management - Models"

        Private NotInheritable Class AdminProductCandidate
            Public Property ProductId As String = ""
            Public Property ProductTitle As String = ""
            Public Property ActivationUsageText As String = ""
            Public Property ParsedUsedCount As Integer = -1
            Public Property ParsedMaxCount As Integer = -1
            Public Property NextPaymentText As String = ""
            Public Property Confidence As String = "Medium"
            Public Property SourceOffset As Integer
        End Class

        Private NotInheritable Class AdminInstanceCandidate
            Public Property InstanceUserId As String = ""
            Public Property ActivatedOnText As String = ""
            Public Property Confidence As String = "Medium"
            Public Property SourceOffset As Integer
            Public Property ProductId As String = ""
            Public Property ProductTitle As String = ""
            Public Property ProductAssociationInferred As Boolean
            Public Property IsAmbiguous As Boolean
        End Class

        Private NotInheritable Class AdminBackendParseResult
            Public Property RawText As String = ""
            Public Property WorkingText As String = ""
            Public Property ExtractedLicenseKey As String = ""
            Public Property LicenseMatched As Boolean
            Public Property BlockingError As String = ""
            Public ReadOnly Property Warnings As New List(Of String)()
            Public ReadOnly Property Products As New List(Of AdminProductCandidate)()
            Public ReadOnly Property Instances As New List(Of AdminInstanceCandidate)()
        End Class

        Private NotInheritable Class AdminUserIdMappingConfig
            Public Property PrimaryField As String = "User ID"
            Public Property UseSecondaryField As Boolean
            Public Property SecondaryField As String = ""
            Public Property PrimaryFieldPart As String = "Whole"
            Public Property SecondaryFieldPart As String = "Whole"
            Public Property Separator As String = "."
            Public Property UseEmailLocalPart As Boolean
            Public Property Lowercase As Boolean = True
            Public Property Uppercase As Boolean
            Public Property TrimWhitespace As Boolean = True
            Public Property RemoveSpaces As Boolean
            Public Property ReplaceSpacesWith As String = ""
            Public Property RemoveDiacritics As Boolean
        End Class

        Private NotInheritable Class AdminPendingDesiredRow
            Public Property ProductId As String = ""
            Public Property ProductTitle As String = ""
            Public Property InstanceUserId As String = ""
            Public Property WarningText As String = ""
            Public Property SourceType As String = ""
            Public Property SourceKind As String = ""
        End Class

        Private NotInheritable Class AdminLicenseRow
            Public Property IncludeRow As Boolean = True
            Public Property ProductId As String = ""
            Public Property ProductTitle As String = ""
            Public Property InstanceUserId As String = ""
            Public Property ParsedActivationDate As String = ""
            Public Property SourceType As String = ""
            Public Property VerificationState As String = "Not verified"
            Public Property DesiredState As String = "Ignore"
            Public Property PlannedAction As String = "Ignore"
            Public Property ApiResult As String = ""
            Public Property WarningText As String = ""
            Public Property Confidence As String = ""
            Public Property SourceKind As String = ""
            Public Property ProductAssociationInferred As Boolean
            Public Property IsAmbiguous As Boolean
            Public Property SourceOffset As Integer
            Public Property IsVerifiedActive As Boolean
            Public Property RequiresReverify As Boolean = True
            Public Property DesiredOccurrences As Integer
            Public Property LastKnownActivationsRemaining As Integer = -1
            Public Property LastKnownTotalActivations As Integer = -1
            Public Property LastKnownTotalPurchased As Integer = -1
            Public Property LastCorrelationId As String = ""
        End Class

#End Region

#Region "Organization Bulk License Management - Entry Point"

        Public Shared Sub ShowOrganizationBulkLicenseManagementDialog()
            Dim storedProductId As String = ""
            Dim storedLicenseKey As String = ""

            If Not TryGetEligibleOrganizationAdminLicense(storedProductId, storedLicenseKey) Then
                Return
            End If

            Dim mappingConfig As AdminUserIdMappingConfig = LoadAdminUserIdMappingConfig()
            Dim rawBackendText As String = ""
            Dim lastApplyCorrelationId As String = ""
            Dim rows As New BindingList(Of AdminLicenseRow)()
            Dim allowedProductIds As New HashSet(Of String)(StringComparer.Ordinal)
            Dim parsedProductTitles As New Dictionary(Of String, String)(StringComparer.Ordinal)
            Dim parsedProductUsage As New Dictionary(Of String, String)(StringComparer.Ordinal)
            Dim changesMadeSinceLastVerify As Boolean = False

            Using form As New System.Windows.Forms.Form()
                form.AutoScaleDimensions = New System.Drawing.SizeF(96.0F, 96.0F)
                form.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
                form.Text = $"{AN} - Organization License Manager"
                form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
                form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
                form.MinimizeBox = False
                form.MaximizeBox = False
                form.ShowInTaskbar = True
                form.TopMost = True
                form.MinimumSize = New System.Drawing.Size(1500, 760)
                form.Size = New System.Drawing.Size(1720, 860)
                form.Font = New System.Drawing.Font("Segoe UI", 9.5F)

                Try
                    Dim bmp As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                    form.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())
                Catch
                End Try

                Dim mainLayout As New System.Windows.Forms.TableLayoutPanel() With {
                    .Dock = System.Windows.Forms.DockStyle.Fill,
                    .Padding = New System.Windows.Forms.Padding(18),
                    .ColumnCount = 1,
                    .RowCount = 6
                }
                mainLayout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
                mainLayout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
                mainLayout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
                mainLayout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
                mainLayout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
                mainLayout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
                form.Controls.Add(mainLayout)

                Dim lblTitle As New System.Windows.Forms.Label() With {
                    .Text = "Organization Bulk License Management",
                    .Font = New System.Drawing.Font("Segoe UI", 12.0F, System.Drawing.FontStyle.Bold),
                    .AutoSize = True,
                    .Margin = New System.Windows.Forms.Padding(0, 0, 0, 6)
                }
                mainLayout.Controls.Add(lblTitle, 0, 0)

                Dim lblInfo As New System.Windows.Forms.Label() With {
                    .Text = $"Stored Product ID: {If(String.IsNullOrWhiteSpace(storedProductId), "(none)", storedProductId)}    Stored License Key: {MaskLicenseKey(storedLicenseKey)}" & vbCrLf &
                            "Paste the website/backend license page first. Extra website noise is acceptable. No pasted backend text is stored.",
                    .AutoSize = True,
                    .MaximumSize = New System.Drawing.Size(1120, 0),
                    .Margin = New System.Windows.Forms.Padding(0, 0, 0, 12)
                }
                mainLayout.Controls.Add(lblInfo, 0, 1)

                Dim filterLayout As New System.Windows.Forms.TableLayoutPanel() With {
                    .Dock = System.Windows.Forms.DockStyle.Fill,
                    .ColumnCount = 10,
                    .AutoSize = True,
                    .Margin = New System.Windows.Forms.Padding(0, 0, 0, 10)
                }
                filterLayout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize))
                filterLayout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20.0F))
                filterLayout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize))
                filterLayout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20.0F))
                filterLayout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize))
                filterLayout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20.0F))
                filterLayout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize))
                filterLayout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20.0F))
                filterLayout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize))
                filterLayout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20.0F))

                Dim lblFilterProduct As New System.Windows.Forms.Label() With {.Text = "Product:", .Anchor = System.Windows.Forms.AnchorStyles.Left, .AutoSize = True}
                Dim cboFilterProduct As New System.Windows.Forms.ComboBox() With {.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList, .Dock = System.Windows.Forms.DockStyle.Fill}
                Dim lblFilterDesired As New System.Windows.Forms.Label() With {.Text = "Desired State:", .Anchor = System.Windows.Forms.AnchorStyles.Left, .AutoSize = True}
                Dim cboFilterDesired As New System.Windows.Forms.ComboBox() With {.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList, .Dock = System.Windows.Forms.DockStyle.Fill}
                Dim lblFilterAction As New System.Windows.Forms.Label() With {.Text = "Action:", .Anchor = System.Windows.Forms.AnchorStyles.Left, .AutoSize = True}
                Dim cboFilterAction As New System.Windows.Forms.ComboBox() With {.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList, .Dock = System.Windows.Forms.DockStyle.Fill}
                Dim lblFilterVerify As New System.Windows.Forms.Label() With {.Text = "Verification:", .Anchor = System.Windows.Forms.AnchorStyles.Left, .AutoSize = True}
                Dim cboFilterVerify As New System.Windows.Forms.ComboBox() With {.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList, .Dock = System.Windows.Forms.DockStyle.Fill}
                Dim lblFilterWarnings As New System.Windows.Forms.Label() With {.Text = "Warnings:", .Anchor = System.Windows.Forms.AnchorStyles.Left, .AutoSize = True}
                Dim cboFilterWarnings As New System.Windows.Forms.ComboBox() With {.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList, .Dock = System.Windows.Forms.DockStyle.Fill}

                filterLayout.Controls.Add(lblFilterProduct, 0, 0)
                filterLayout.Controls.Add(cboFilterProduct, 1, 0)
                filterLayout.Controls.Add(lblFilterDesired, 2, 0)
                filterLayout.Controls.Add(cboFilterDesired, 3, 0)
                filterLayout.Controls.Add(lblFilterAction, 4, 0)
                filterLayout.Controls.Add(cboFilterAction, 5, 0)
                filterLayout.Controls.Add(lblFilterVerify, 6, 0)
                filterLayout.Controls.Add(cboFilterVerify, 7, 0)
                filterLayout.Controls.Add(lblFilterWarnings, 8, 0)
                filterLayout.Controls.Add(cboFilterWarnings, 9, 0)
                mainLayout.Controls.Add(filterLayout, 0, 2)

                Dim grid As New System.Windows.Forms.DataGridView() With {
                    .Dock = System.Windows.Forms.DockStyle.Fill,
                    .AutoGenerateColumns = False,
                    .AllowUserToAddRows = False,
                    .AllowUserToDeleteRows = False,
                    .AllowUserToResizeRows = False,
                    .MultiSelect = True,
                    .SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect,
                    .EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter,
                    .RowHeadersVisible = False,
                    .AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.None,
                    .ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText,
                    .ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
                }

                Dim bindingSource As New System.Windows.Forms.BindingSource() With {
                    .DataSource = rows
                }
                grid.DataSource = bindingSource

                Dim colInclude As New System.Windows.Forms.DataGridViewCheckBoxColumn() With {
                    .DataPropertyName = NameOf(AdminLicenseRow.IncludeRow),
                    .HeaderText = "Include",
                    .Width = 72
                }
                Dim colProductId As New System.Windows.Forms.DataGridViewTextBoxColumn() With {
                    .DataPropertyName = NameOf(AdminLicenseRow.ProductId),
                    .HeaderText = "Product ID",
                    .Width = 95,
                    .ReadOnly = True
                }
                Dim colProductTitle As New System.Windows.Forms.DataGridViewTextBoxColumn() With {
                    .DataPropertyName = NameOf(AdminLicenseRow.ProductTitle),
                    .HeaderText = "Product Title",
                    .Width = 210,
                    .ReadOnly = True
                }
                Dim colInstance As New System.Windows.Forms.DataGridViewTextBoxColumn() With {
                    .DataPropertyName = NameOf(AdminLicenseRow.InstanceUserId),
                    .HeaderText = "Instance / User ID",
                    .Width = 240,
                    .ReadOnly = True
                }
                Dim colActivated As New System.Windows.Forms.DataGridViewTextBoxColumn() With {
                    .DataPropertyName = NameOf(AdminLicenseRow.ParsedActivationDate),
                    .HeaderText = "Parsed Activation Date",
                    .Width = 165,
                    .ReadOnly = True
                }
                Dim colSource As New System.Windows.Forms.DataGridViewTextBoxColumn() With {
                    .DataPropertyName = NameOf(AdminLicenseRow.SourceType),
                    .HeaderText = "Source",
                    .Width = 175,
                    .ReadOnly = True
                }
                Dim colVerify As New System.Windows.Forms.DataGridViewTextBoxColumn() With {
                    .DataPropertyName = NameOf(AdminLicenseRow.VerificationState),
                    .HeaderText = "Verification State",
                    .Width = 155,
                    .ReadOnly = True
                }
                Dim colDesired As New System.Windows.Forms.DataGridViewComboBoxColumn() With {
                    .DataPropertyName = NameOf(AdminLicenseRow.DesiredState),
                    .HeaderText = "Desired State",
                    .Width = 180,
                    .FlatStyle = System.Windows.Forms.FlatStyle.Flat
                }
                colDesired.Items.AddRange(New Object() {"Ignore", "DesiredActive", "DesiredInactive"})

                Dim colPlanned As New System.Windows.Forms.DataGridViewTextBoxColumn() With {
                    .DataPropertyName = NameOf(AdminLicenseRow.PlannedAction),
                    .HeaderText = "Planned Action",
                    .Width = 190,
                    .ReadOnly = True
                }
                Dim colApi As New System.Windows.Forms.DataGridViewTextBoxColumn() With {
                    .DataPropertyName = NameOf(AdminLicenseRow.ApiResult),
                    .HeaderText = "API Result",
                    .Width = 145,
                    .ReadOnly = True
                }
                Dim colWarn As New System.Windows.Forms.DataGridViewTextBoxColumn() With {
                    .DataPropertyName = NameOf(AdminLicenseRow.WarningText),
                    .HeaderText = "Warning / Error",
                    .Width = 320,
                    .ReadOnly = True
                }

                grid.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {
                    colInclude, colProductId, colProductTitle, colInstance, colActivated, colSource, colVerify, colDesired, colPlanned, colApi, colWarn
                })
                grid.ColumnHeadersDefaultCellStyle.WrapMode = System.Windows.Forms.DataGridViewTriState.True
                grid.ColumnHeadersDefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                grid.AutoResizeColumnHeadersHeight()

                mainLayout.Controls.Add(grid, 0, 3)

                Dim pnlStatus As New System.Windows.Forms.Panel() With {
                    .Dock = System.Windows.Forms.DockStyle.Fill,
                    .BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle,
                    .Padding = New System.Windows.Forms.Padding(8),
                    .Height = 96,
                    .Margin = New System.Windows.Forms.Padding(0, 10, 0, 10)
                }
                Dim lblStatus As New System.Windows.Forms.Label() With {
                    .Text = "Paste a backend page and click 'Parse / Reparse'.",
                    .Dock = System.Windows.Forms.DockStyle.Fill,
                    .AutoSize = False,
                    .ForeColor = System.Drawing.Color.DarkSlateBlue
                }
                pnlStatus.Controls.Add(lblStatus)
                mainLayout.Controls.Add(pnlStatus, 0, 4)

                Dim buttonPanel As New System.Windows.Forms.FlowLayoutPanel() With {
                    .Dock = System.Windows.Forms.DockStyle.Fill,
                    .AutoSize = True,
                    .WrapContents = True,
                    .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
                    .Margin = New System.Windows.Forms.Padding(0)
                }

                Dim btnPaste As New System.Windows.Forms.Button() With {.Text = "Paste Backend License Page", .AutoSize = True}
                Dim btnParse As New System.Windows.Forms.Button() With {.Text = "Parse / Reparse", .AutoSize = True}
                Dim btnVerify As New System.Windows.Forms.Button() With {.Text = "Verify Parsed Activations", .AutoSize = True}
                Dim btnAddUser As New System.Windows.Forms.Button() With {.Text = "Add User", .AutoSize = True}
                Dim btnRemoveRows As New System.Windows.Forms.Button() With {.Text = "Mark Selected for Deactivation", .AutoSize = True}
                Dim btnMore As New System.Windows.Forms.Button() With {.Text = "More...", .AutoSize = True}
                Dim btnPreviewSync As New System.Windows.Forms.Button() With {.Text = "Preview Selected Changes", .AutoSize = True}
                Dim btnApply As New System.Windows.Forms.Button() With {.Text = "Apply Selected Changes", .AutoSize = True}
                Dim btnExportTable As New System.Windows.Forms.Button() With {.Text = "Export Current Table", .AutoSize = True}
                Dim btnExportResult As New System.Windows.Forms.Button() With {.Text = "Export Result Report", .AutoSize = True}
                Dim btnClose As New System.Windows.Forms.Button() With {.Text = "Close", .AutoSize = True}

                buttonPanel.Controls.Add(btnPaste)
                buttonPanel.Controls.Add(btnParse)
                buttonPanel.Controls.Add(btnVerify)
                buttonPanel.Controls.Add(btnAddUser)
                buttonPanel.Controls.Add(btnRemoveRows)
                buttonPanel.Controls.Add(btnMore)
                buttonPanel.Controls.Add(btnPreviewSync)
                buttonPanel.Controls.Add(btnApply)
                buttonPanel.Controls.Add(btnClose)
                mainLayout.Controls.Add(buttonPanel, 0, 5)

                Dim busyControls As New List(Of System.Windows.Forms.Control) From {
                    btnPaste, btnParse, btnVerify, btnAddUser, btnRemoveRows, btnMore, btnPreviewSync, btnApply, btnExportTable, btnExportResult, btnClose
                }

                Dim setStatus As Action(Of String, System.Drawing.Color) =
                    Sub(message As String, color As System.Drawing.Color)
                        lblStatus.Text = message
                        lblStatus.ForeColor = color
                    End Sub

                Dim refreshGridState As Action =
                    Sub()
                        bindingSource.ResetBindings(False)
                        ApplyAdminDuplicateWarnings(rows)
                        RefreshAdminFilterChoices(cboFilterProduct, cboFilterDesired, cboFilterAction, cboFilterVerify, cboFilterWarnings, rows)
                        ApplyAdminGridFilters(grid, cboFilterProduct, cboFilterDesired, cboFilterAction, cboFilterVerify, cboFilterWarnings)
                    End Sub

                Dim setBusy As Action(Of Boolean) =
                    Sub(isBusy As Boolean)
                        For Each control As System.Windows.Forms.Control In busyControls
                            control.Enabled = Not isBusy
                        Next
                        form.Cursor = If(isBusy, System.Windows.Forms.Cursors.WaitCursor, System.Windows.Forms.Cursors.Default)
                        form.Refresh()
                        System.Windows.Forms.Application.DoEvents()
                    End Sub

                RefreshAdminFilterChoices(cboFilterProduct, cboFilterDesired, cboFilterAction, cboFilterVerify, cboFilterWarnings, rows)

                AddHandler cboFilterProduct.SelectedIndexChanged, Sub() ApplyAdminGridFilters(grid, cboFilterProduct, cboFilterDesired, cboFilterAction, cboFilterVerify, cboFilterWarnings)
                AddHandler cboFilterDesired.SelectedIndexChanged, Sub() ApplyAdminGridFilters(grid, cboFilterProduct, cboFilterDesired, cboFilterAction, cboFilterVerify, cboFilterWarnings)
                AddHandler cboFilterAction.SelectedIndexChanged, Sub() ApplyAdminGridFilters(grid, cboFilterProduct, cboFilterDesired, cboFilterAction, cboFilterVerify, cboFilterWarnings)
                AddHandler cboFilterVerify.SelectedIndexChanged, Sub() ApplyAdminGridFilters(grid, cboFilterProduct, cboFilterDesired, cboFilterAction, cboFilterVerify, cboFilterWarnings)
                AddHandler cboFilterWarnings.SelectedIndexChanged, Sub() ApplyAdminGridFilters(grid, cboFilterProduct, cboFilterDesired, cboFilterAction, cboFilterVerify, cboFilterWarnings)
                AddHandler grid.DataBindingComplete, Sub() ApplyAdminGridFilters(grid, cboFilterProduct, cboFilterDesired, cboFilterAction, cboFilterVerify, cboFilterWarnings)

                AddHandler grid.CurrentCellDirtyStateChanged,
                    Sub()
                        If grid.IsCurrentCellDirty Then
                            grid.CommitEdit(System.Windows.Forms.DataGridViewDataErrorContexts.Commit)
                        End If
                    End Sub

                AddHandler grid.CellEndEdit,
                    Sub(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs)
                        If e.RowIndex < 0 OrElse e.RowIndex >= grid.Rows.Count Then Return

                        Dim boundRow As AdminLicenseRow = TryCast(grid.Rows(e.RowIndex).DataBoundItem, AdminLicenseRow)
                        If boundRow Is Nothing Then Return

                        If e.ColumnIndex = colProductId.Index OrElse e.ColumnIndex = colInstance.Index OrElse e.ColumnIndex = colDesired.Index Then
                            boundRow.RequiresReverify = True
                            If Not String.IsNullOrWhiteSpace(boundRow.VerificationState) Then
                                boundRow.VerificationState = "Not verified"
                            End If
                            boundRow.IsVerifiedActive = False
                            boundRow.ApiResult = ""

                            If e.ColumnIndex = colDesired.Index Then
                                Select Case If(boundRow.DesiredState, "")
                                    Case "DesiredInactive"
                                        boundRow.IncludeRow = True
                                        boundRow.PlannedAction = "Pending deactivation"
                                    Case "DesiredActive"
                                        boundRow.IncludeRow = True
                                        boundRow.PlannedAction = "Pending activation"
                                    Case Else
                                        boundRow.IncludeRow = False
                                        boundRow.PlannedAction = "Ignore"
                                End Select
                            End If

                            changesMadeSinceLastVerify = True
                            refreshGridState()
                        End If
                    End Sub

                AddHandler grid.CellDoubleClick,
                    Sub(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs)
                        If e.RowIndex < 0 OrElse e.RowIndex >= grid.Rows.Count Then Return
                        If e.ColumnIndex <> colProductId.Index AndAlso e.ColumnIndex <> colInstance.Index Then Return

                        Dim boundRow As AdminLicenseRow = TryCast(grid.Rows(e.RowIndex).DataBoundItem, AdminLicenseRow)
                        If boundRow Is Nothing Then Return

                        Dim currentValue As String = If(e.ColumnIndex = colProductId.Index, If(boundRow.ProductId, ""), If(boundRow.InstanceUserId, ""))
                        Dim promptText As String = If(e.ColumnIndex = colProductId.Index, "Edit Product ID:", "Edit Instance / User ID:")
                        Dim editedValue As String = ShowCustomInputBox(
                            promptText,
                            $"{AN} - Edit Value",
                            True,
                            currentValue).Trim()

                        If editedValue.Length = 0 OrElse editedValue = currentValue Then
                            Return
                        End If

                        If e.ColumnIndex = colProductId.Index Then
                            boundRow.ProductId = editedValue
                        Else
                            boundRow.InstanceUserId = editedValue
                        End If

                        boundRow.RequiresReverify = True
                        boundRow.VerificationState = "Not verified"
                        boundRow.IsVerifiedActive = False
                        boundRow.ApiResult = ""
                        changesMadeSinceLastVerify = True
                        refreshGridState()
                    End Sub

                AddHandler btnPaste.Click,
                    Sub()
                        Dim pastedText As String = ShowCustomInputBox(
                            "Paste the website/backend license page text. Extra website text/noise Is acceptable." & vbCrLf & vbCrLf &
                            "The text Is kept only In memory For this session.",
                            $"{AN} - Paste Backend License Page",
                            False)

                        If pastedText = "ESC" Then
                            Return
                        End If

                        If String.IsNullOrWhiteSpace(pastedText) Then
                            setStatus("Paste canceled Or empty text provided.", System.Drawing.Color.DarkOrange)
                            Return
                        End If

                        rawBackendText = pastedText
                        setStatus("Backend text stored In memory For this session. Click 'Parse / Reparse' to continue.", System.Drawing.Color.DarkGreen)
                    End Sub

                AddHandler btnParse.Click,
                    Sub()
                        If String.IsNullOrWhiteSpace(rawBackendText) Then
                            ShowCustomMessageBox("Please paste the backend license page first.", $"{AN} - Organization License Manager")
                            Return
                        End If

                        setBusy(True)

                        Try
                            Dim parseResult As AdminBackendParseResult = ParseOrganizationBackendLicensePage(rawBackendText, storedLicenseKey)

                            If Not String.IsNullOrWhiteSpace(parseResult.BlockingError) Then
                                setStatus(parseResult.BlockingError, System.Drawing.Color.DarkRed)
                                ShowCustomMessageBox(parseResult.BlockingError, $"{AN} - Organization License Manager")
                                Return
                            End If

                            allowedProductIds.Clear()
                            parsedProductTitles.Clear()
                            parsedProductUsage.Clear()

                            For Each product As AdminProductCandidate In parseResult.Products
                                If Not String.IsNullOrWhiteSpace(product.ProductId) Then
                                    allowedProductIds.Add(product.ProductId)
                                    parsedProductTitles(product.ProductId) = product.ProductTitle
                                    parsedProductUsage(product.ProductId) = product.ActivationUsageText
                                End If
                            Next

                            Dim parsedRowsAdded As Integer = MergeParsedCandidatesIntoRows(rows, parseResult)

                            Dim statusBuilder As New System.Text.StringBuilder()
                            statusBuilder.AppendLine($"Parse successful. Parsed {parseResult.Products.Count} product(s) and {parseResult.Instances.Count} activation candidate(s).")
                            If parseResult.Products.Count = 0 Then
                                statusBuilder.AppendLine("No product IDs were found.")
                            End If
                            If parseResult.Instances.Count = 0 Then
                                statusBuilder.AppendLine("No activation instances were found. This is acceptable when no users are currently activated.")
                            End If
                            If parseResult.Warnings.Count > 0 Then
                                statusBuilder.AppendLine()
                                statusBuilder.AppendLine("Warnings:")
                                For Each warning As String In parseResult.Warnings
                                    statusBuilder.AppendLine($"- {warning}")
                                Next
                            End If

                            changesMadeSinceLastVerify = True
                            refreshGridState()
                            setStatus(statusBuilder.ToString().TrimEnd(), If(parseResult.Warnings.Count > 0, System.Drawing.Color.DarkOrange, System.Drawing.Color.DarkGreen))
                        Catch ex As System.Exception
                            setStatus($"Parse failed: {ex.Message}", System.Drawing.Color.DarkRed)
                            ShowCustomMessageBox($"Parse failed: {ex.Message}", $"{AN} - Organization License Manager")
                        Finally
                            setBusy(False)
                        End Try
                    End Sub

                AddHandler btnVerify.Click,
                    Sub()
                        If allowedProductIds.Count = 0 Then
                            ShowCustomMessageBox("Please parse a matching backend license page first.", $"{AN} - Organization License Manager")
                            Return
                        End If

                        Dim targetRows As List(Of AdminLicenseRow) = rows.ToList()

                        If grid.SelectedRows.Count > 0 Then
                            Dim verifySelected As Integer = ShowCustomYesNoBox(
                                $"You have selected {grid.SelectedRows.Count} row(s)." & vbCrLf & vbCrLf &
                                "Would you like to verify only the selected rows?" & vbCrLf & vbCrLf &
                                "Click 'Verify Selected' to verify only those rows, or 'Verify All' to verify all rows in the table.",
                                "Verify Selected",
                                "Verify All",
                                $"{AN} - Verify Parsed Activations")

                            If verifySelected = 1 Then
                                targetRows = New List(Of AdminLicenseRow)()
                                For Each selectedGridRow As System.Windows.Forms.DataGridViewRow In grid.SelectedRows
                                    Dim boundRow As AdminLicenseRow = TryCast(selectedGridRow.DataBoundItem, AdminLicenseRow)
                                    If boundRow IsNot Nothing AndAlso Not targetRows.Contains(boundRow) Then
                                        targetRows.Add(boundRow)
                                    End If
                                Next
                            End If
                        End If

                        setBusy(True)

                        Try
                            Dim verifiedCount As Integer = VerifyAdminRows(
                                targetRows,
                                storedLicenseKey,
                                allowedProductIds,
                                Sub(currentIndex As Integer, totalCount As Integer, message As String)
                                    setStatus($"Verification progress: {currentIndex} / {totalCount}" & vbCrLf & message, System.Drawing.Color.DarkBlue)
                                    grid.Refresh()
                                    form.Refresh()
                                    System.Windows.Forms.Application.DoEvents()
                                End Sub)

                            changesMadeSinceLastVerify = False
                            refreshGridState()
                            setStatus($"Verification completed for {verifiedCount} row(s).", System.Drawing.Color.DarkGreen)
                        Catch ex As System.Exception
                            setStatus($"Verification failed: {ex.Message}", System.Drawing.Color.DarkRed)
                            ShowCustomMessageBox($"Verification failed: {ex.Message}", $"{AN} - Organization License Manager")
                        Finally
                            setBusy(False)
                        End Try
                    End Sub

                AddHandler btnAddUser.Click,
                    Sub()
                        If allowedProductIds.Count = 0 Then
                            ShowCustomMessageBox("Please parse a matching backend license page first.", $"{AN} - Organization License Manager")
                            Return
                        End If

                        Dim defaultProductId As String = SelectDefaultAdminProductId(allowedProductIds, parsedProductTitles, "Select the product for the new user.")
                        If String.IsNullOrWhiteSpace(defaultProductId) Then
                            Return
                        End If

                        Dim userId As String = ShowCustomInputBox(
                            "Enter the new User ID / instance value.",
                            $"{AN} - Add User",
                            True).Trim()

                        If String.IsNullOrWhiteSpace(userId) Then
                            Return
                        End If

                        Dim row As AdminLicenseRow = UpsertDesiredRow(rows, defaultProductId, GetProductTitleFromLookup(parsedProductTitles, defaultProductId), userId, "Manually added", "Manual")
                        row.PlannedAction = "Activate"
                        row.ApiResult = ""
                        row.RequiresReverify = True

                        changesMadeSinceLastVerify = True
                        refreshGridState()
                        setStatus($"User queued: {userId}", System.Drawing.Color.DarkGreen)
                    End Sub

                AddHandler btnRemoveRows.Click,
                    Sub()
                        If grid.SelectedRows.Count = 0 Then
                            ShowCustomMessageBox("Select one or more rows to mark for deactivation.", $"{AN} - Organization License Manager")
                            Return
                        End If

                        Dim confirm As Integer = ShowCustomYesNoBox(
                            $"Mark {grid.SelectedRows.Count} selected row(s) for server-side deactivation?" & vbCrLf & vbCrLf &
                            "Only verified-active rows can be marked for deactivation." & vbCrLf &
                            "The selected rows will remain in the table, will be re-verified, and will be deactivated only when you later click 'Apply Selected Changes'.",
                            "Mark for Deactivation",
                            "Cancel",
                            $"{AN} - Mark Selected for Deactivation")

                        If confirm <> 1 Then
                            Return
                        End If

                        Dim markedCount As Integer = 0
                        Dim skippedCount As Integer = 0

                        For Each selectedRow As System.Windows.Forms.DataGridViewRow In grid.SelectedRows
                            Dim boundRow As AdminLicenseRow = TryCast(selectedRow.DataBoundItem, AdminLicenseRow)
                            If boundRow Is Nothing Then Continue For

                            If boundRow.IsVerifiedActive OrElse boundRow.VerificationState.Equals("Verified active", StringComparison.OrdinalIgnoreCase) Then
                                boundRow.DesiredState = "DesiredInactive"
                                boundRow.PlannedAction = "Deactivate"
                                boundRow.IncludeRow = True
                                boundRow.RequiresReverify = True
                                boundRow.ApiResult = ""
                                markedCount += 1
                            Else
                                AppendDistinctAdminWarning(boundRow, "Only verified-active rows can be marked for deactivation.")
                                skippedCount += 1
                            End If
                        Next

                        changesMadeSinceLastVerify = True
                        refreshGridState()

                        setStatus(
                            $"Marked {markedCount} selected row(s) for deactivation. " &
                            $"Skipped {skippedCount} row(s). " &
                            "Marked rows will be re-verified before apply.",
                            System.Drawing.Color.DarkOrange)
                    End Sub

                AddHandler btnMore.Click,
                    Sub()
                        Dim choice As String = ShowSelectionForm(
                            "Choose an additional organization-license action:",
                            $"{AN} - Organization License Manager",
                            New String() {
                                "Bulk Add Users",
                                "Bulk Mark for Deactivation",
                                "Sync with Imported List",
                                "Configure User-ID Mapping",
                                "Export Current Table",
                                "Export Result Report"
                            })

                        If choice = "ESC" Then
                            Return
                        End If

                        If choice.StartsWith("Bulk Add Users", StringComparison.OrdinalIgnoreCase) Then
                            If allowedProductIds.Count = 0 Then
                                ShowCustomMessageBox("Please parse a matching backend license page first.", $"{AN} - Organization License Manager")
                                Return
                            End If

                            Dim defaultProductId As String = SelectDefaultAdminProductId(allowedProductIds, parsedProductTitles, "Select the default product for bulk-added users.")
                            If String.IsNullOrWhiteSpace(defaultProductId) Then
                                Return
                            End If

                            Dim bulkText As String = ShowCustomInputBox(
                                "Paste one user ID per line. Spreadsheet-style tab/comma/semicolon-separated values are also accepted.",
                                $"{AN} - Bulk Add Users",
                                False)

                            If bulkText = "ESC" OrElse String.IsNullOrWhiteSpace(bulkText) Then
                                Return
                            End If

                            Dim addedCount As Integer = 0
                            For Each userId As String In ParseOpaqueUserList(bulkText)
                                If String.IsNullOrWhiteSpace(userId) Then Continue For
                                Dim row As AdminLicenseRow = UpsertDesiredRow(rows, defaultProductId, GetProductTitleFromLookup(parsedProductTitles, defaultProductId), userId, "Bulk added", "BulkAdd")
                                row.PlannedAction = "Activate"
                                row.RequiresReverify = True
                                addedCount += 1
                            Next

                            changesMadeSinceLastVerify = True
                            refreshGridState()
                            setStatus($"{addedCount} user(s) queued for activation.", System.Drawing.Color.DarkGreen)
                            Return
                        End If

                        If choice.StartsWith("Bulk Mark for Deactivation", StringComparison.OrdinalIgnoreCase) Then
                            If allowedProductIds.Count = 0 Then
                                ShowCustomMessageBox("Please parse a matching backend license page first.", $"{AN} - Organization License Manager")
                                Return
                            End If

                            Dim matchedCount As Integer = 0
                            Dim useSelectedRows As Boolean = False

                            If grid.SelectedRows.Count > 0 Then
                                Dim selectedAnswer As Integer = ShowCustomYesNoBox(
                                    $"You have selected {grid.SelectedRows.Count} row(s)." & vbCrLf & vbCrLf &
                                    "Would you like to mark the selected rows for deactivation?" & vbCrLf & vbCrLf &
                                    "Click 'Use Selected Rows' to mark only those rows, or 'Paste User List' to paste a list of user IDs instead.",
                                    "Use Selected Rows",
                                    "Paste User List",
                                    $"{AN} - Bulk Mark for Deactivation")

                                useSelectedRows = (selectedAnswer = 1)
                            End If

                            If useSelectedRows Then
                                For Each selectedRow As System.Windows.Forms.DataGridViewRow In grid.SelectedRows
                                    Dim boundRow As AdminLicenseRow = TryCast(selectedRow.DataBoundItem, AdminLicenseRow)
                                    If boundRow Is Nothing Then Continue For

                                    boundRow.DesiredState = "DesiredInactive"
                                    boundRow.PlannedAction = "Pending deactivation"
                                    boundRow.IncludeRow = True
                                    boundRow.RequiresReverify = True
                                    boundRow.ApiResult = ""
                                    matchedCount += 1
                                Next
                            Else
                                Dim removalText As String = ShowCustomInputBox(
                                    "Paste a list of user IDs to mark for deactivation. Exact matching is used by default.",
                                    $"{AN} - Bulk Mark for Deactivation",
                                    False)

                                If removalText = "ESC" OrElse String.IsNullOrWhiteSpace(removalText) Then
                                    Return
                                End If

                                For Each removalUserId As String In ParseOpaqueUserList(removalText)
                                    Dim matchingRows = rows.Where(Function(r) r.InstanceUserId.Equals(removalUserId, StringComparison.Ordinal)).ToList()

                                    If matchingRows.Count = 0 Then
                                        rows.Add(New AdminLicenseRow() With {
                                            .IncludeRow = False,
                                            .ProductId = "",
                                            .ProductTitle = "",
                                            .InstanceUserId = removalUserId,
                                            .ParsedActivationDate = "",
                                            .SourceType = "Bulk deactivation request",
                                            .SourceKind = "BulkRemove",
                                            .VerificationState = "Unmatched",
                                            .DesiredState = "DesiredInactive",
                                            .PlannedAction = "Unmatched",
                                            .ApiResult = "",
                                            .WarningText = "No matching current row found for exact user-ID match.",
                                            .Confidence = "",
                                            .RequiresReverify = False
                                        })
                                    Else
                                        For Each matchingRow As AdminLicenseRow In matchingRows
                                            matchingRow.DesiredState = "DesiredInactive"
                                            matchingRow.PlannedAction = "Pending deactivation"
                                            matchingRow.IncludeRow = True
                                            matchingRow.RequiresReverify = True
                                            matchingRow.ApiResult = ""
                                            matchedCount += 1
                                        Next
                                    End If
                                Next
                            End If

                            changesMadeSinceLastVerify = True
                            refreshGridState()
                            setStatus($"{matchedCount} row(s) marked for deactivation.", System.Drawing.Color.DarkOrange)
                            Return
                        End If

                        If choice.StartsWith("Sync with Imported List", StringComparison.OrdinalIgnoreCase) Then
                            If allowedProductIds.Count = 0 Then
                                ShowCustomMessageBox("Please parse a matching backend license page first.", $"{AN} - Organization License Manager")
                                Return
                            End If

                            Dim markedForDeactivationCount As Integer = 0
                            Dim importedCount As Integer = ImportDesiredRowsFromFile(rows, allowedProductIds, parsedProductTitles, mappingConfig, markedForDeactivationCount)

                            If importedCount > 0 Then
                                changesMadeSinceLastVerify = True
                                refreshGridState()
                                setStatus(
                                    $"Sync completed. {importedCount} imported user(s) were evaluated. {markedForDeactivationCount} existing activation(s) not present in the imported list are marked for deactivation.",
                                    If(markedForDeactivationCount > 0, System.Drawing.Color.DarkOrange, System.Drawing.Color.DarkGreen))
                            End If
                            Return
                        End If

                        If choice.StartsWith("Configure User-ID Mapping", StringComparison.OrdinalIgnoreCase) Then
                            If ShowAdminUserIdMappingDialog(mappingConfig) Then
                                SaveAdminUserIdMappingConfig(mappingConfig)
                                setStatus("User-ID mapping configuration saved to local settings.", System.Drawing.Color.DarkGreen)
                            End If
                            Return
                        End If

                        If choice.StartsWith("Export Current Table", StringComparison.OrdinalIgnoreCase) Then
                            btnExportTable.PerformClick()
                            Return
                        End If

                        If choice.StartsWith("Export Result Report", StringComparison.OrdinalIgnoreCase) Then
                            btnExportResult.PerformClick()
                            Return
                        End If
                    End Sub

                AddHandler btnPreviewSync.Click,
                    Sub()
                        If allowedProductIds.Count = 0 Then
                            ShowCustomMessageBox("Please parse a matching backend license page first.", $"{AN} - Organization License Manager")
                            Return
                        End If

                        setBusy(True)

                        Try
                            Dim summary As String = PreviewAdminSynchronization(rows.ToList(), allowedProductIds)
                            refreshGridState()
                            setStatus(summary, System.Drawing.Color.DarkBlue)
                        Catch ex As System.Exception
                            setStatus($"Preview of selected changes failed: {ex.Message}", System.Drawing.Color.DarkRed)
                            ShowCustomMessageBox($"Preview of selected changes failed: {ex.Message}", $"{AN} - Organization License Manager")
                        Finally
                            setBusy(False)
                        End Try
                    End Sub

                AddHandler btnApply.Click,
                    Sub()
                        If allowedProductIds.Count = 0 Then
                            ShowCustomMessageBox("Please parse a matching backend license page first.", $"{AN} - Organization License Manager")
                            Return
                        End If

                        setBusy(True)

                        Try
                            Dim targetRows = rows.Where(
                                Function(r) r.IncludeRow AndAlso
                                    (r.DesiredState.Equals("DesiredActive", StringComparison.OrdinalIgnoreCase) OrElse
                                     r.DesiredState.Equals("DesiredInactive", StringComparison.OrdinalIgnoreCase) OrElse
                                     r.PlannedAction.Equals("Pending activation", StringComparison.OrdinalIgnoreCase) OrElse
                                     r.PlannedAction.Equals("Pending deactivation", StringComparison.OrdinalIgnoreCase) OrElse
                                     r.PlannedAction.Equals("Activate", StringComparison.OrdinalIgnoreCase) OrElse
                                     r.PlannedAction.Equals("Deactivate", StringComparison.OrdinalIgnoreCase))
                                ).ToList()

                            If targetRows.Count = 0 Then
                                ShowCustomMessageBox("There are no selected rows queued for activation or deactivation.", $"{AN} - Organization License Manager")
                                Return
                            End If

                            If changesMadeSinceLastVerify Then
                                VerifyAdminRows(
                                    targetRows,
                                    storedLicenseKey,
                                    allowedProductIds,
                                    Sub(currentIndex As Integer, totalCount As Integer, message As String)
                                        setStatus($"Re-verifying selected rows before apply: {currentIndex} / {totalCount}" & vbCrLf & message, System.Drawing.Color.DarkBlue)
                                        grid.Refresh()
                                        form.Refresh()
                                        System.Windows.Forms.Application.DoEvents()
                                    End Sub)
                                changesMadeSinceLastVerify = False
                            End If

                            Dim summary As String = PreviewAdminSynchronization(rows.ToList(), allowedProductIds)
                            refreshGridState()

                            Dim applyRows = rows.Where(Function(r) r.IncludeRow AndAlso
                                                                  (r.PlannedAction.Equals("Activate", StringComparison.OrdinalIgnoreCase) OrElse
                                                                   r.PlannedAction.Equals("Deactivate", StringComparison.OrdinalIgnoreCase))).ToList()

                            If applyRows.Count = 0 Then
                                setStatus(summary, System.Drawing.Color.DarkOrange)
                                ShowCustomMessageBox("There are no selected activate/deactivate actions to apply.", $"{AN} - Organization License Manager")
                                Return
                            End If

                            Dim confirmText As String =
                                "Dry-run plan before applying changes:" & vbCrLf & vbCrLf &
                                $"License Key: {MaskLicenseKey(storedLicenseKey)}" & vbCrLf & vbCrLf &
                                summary & vbCrLf & vbCrLf &
                                "Apply these changes now? Only the selected queued rows will be processed. Deactivations will run before activations."

                            Dim confirm As Integer = ShowCustomYesNoBox(confirmText, "Apply", "Cancel", $"{AN} - Apply License Changes")
                            If confirm <> 1 Then
                                setStatus("Apply canceled.", System.Drawing.Color.DarkOrange)
                                Return
                            End If

                            Dim report As String = ApplyAdminPlannedChanges(
                                rows,
                                storedLicenseKey,
                                allowedProductIds,
                                lastApplyCorrelationId,
                                Sub(currentIndex As Integer, totalCount As Integer, message As String)
                                    setStatus($"Apply progress: {currentIndex} / {totalCount}" & vbCrLf & message, System.Drawing.Color.DarkBlue)
                                    grid.Refresh()
                                    form.Refresh()
                                    System.Windows.Forms.Application.DoEvents()
                                End Sub)

                            refreshGridState()
                            setStatus(report, System.Drawing.Color.DarkGreen)
                            ShowCustomMessageBox(report, $"{AN} - Organization License Manager")
                        Catch ex As System.Exception
                            setStatus($"Apply failed: {ex.Message}", System.Drawing.Color.DarkRed)
                            ShowCustomMessageBox($"Apply failed: {ex.Message}", $"{AN} - Organization License Manager")
                        Finally
                            setBusy(False)
                        End Try
                    End Sub

                AddHandler btnExportTable.Click,
                    Sub()
                        Try
                            ExportAdminRowsToCsv(rows, False, lastApplyCorrelationId)
                            setStatus("Current table exported.", System.Drawing.Color.DarkGreen)
                        Catch ex As System.Exception
                            setStatus($"Export failed: {ex.Message}", System.Drawing.Color.DarkRed)
                            ShowCustomMessageBox($"Export failed: {ex.Message}", $"{AN} - Organization License Manager")
                        End Try
                    End Sub

                AddHandler btnExportResult.Click,
                    Sub()
                        Try
                            ExportAdminRowsToCsv(rows, True, lastApplyCorrelationId)
                            setStatus("Result report exported.", System.Drawing.Color.DarkGreen)
                        Catch ex As System.Exception
                            setStatus($"Export failed: {ex.Message}", System.Drawing.Color.DarkRed)
                            ShowCustomMessageBox($"Export failed: {ex.Message}", $"{AN} - Organization License Manager")
                        End Try
                    End Sub

                AddHandler btnClose.Click, Sub() form.Close()

                Dim ownerWnd As System.Windows.Forms.IWin32Window = ResolveSameThreadDialogOwner()
                If ownerWnd IsNot Nothing Then
                    form.ShowDialog(ownerWnd)
                Else
                    form.ShowDialog()
                End If
            End Using
        End Sub

#End Region

#Region "Organization Bulk License Management - Eligibility"

        Private Shared Function TryGetEligibleOrganizationAdminLicense(ByRef storedProductId As String, ByRef storedLicenseKey As String) As Boolean
            storedProductId = ""
            storedLicenseKey = ""

            Try
                If Not HasStoredProLicense() Then
                    ShowCustomMessageBox(
                        "Organization bulk license management requires a stored Professional License first." & vbCrLf & vbCrLf &
                        "Please configure and validate a Pro license before using __licensemanager__.",
                        $"{AN} - Organization License Manager")
                    Return False
                End If

                If Not My.Settings.License_ApiConfirmed Then
                    ShowCustomMessageBox(
                        "Organization bulk license management is available only after the stored Professional License has been validated online.",
                        $"{AN} - Organization License Manager")
                    Return False
                End If

                storedProductId = If(My.Settings.License_ProductID, "").Trim()
                storedLicenseKey = If(My.Settings.License_Key, "").Trim()

                If String.IsNullOrWhiteSpace(storedLicenseKey) Then
                    ShowCustomMessageBox(
                        "No stored license key was found. Please configure and validate a Professional License first.",
                        $"{AN} - Organization License Manager")
                    Return False
                End If

                If IsOfflineDomainLicenseKey(storedLicenseKey) Then
                    ShowCustomMessageBox(
                        "Online organization bulk activation management does not apply to offline-domain licenses.",
                        $"{AN} - Organization License Manager")
                    Return False
                End If

                Return True
            Catch ex As System.Exception
                ShowCustomMessageBox(
                    $"Could not read the stored license configuration: {ex.Message}",
                    $"{AN} - Organization License Manager")
                Return False
            End Try
        End Function

#End Region

#Region "Organization Bulk License Management - Parser"

        Private Shared Function ParseOrganizationBackendLicensePage(rawText As String, storedLicenseKey As String) As AdminBackendParseResult
            Dim result As New AdminBackendParseResult() With {
                .RawText = rawText
            }

            If String.IsNullOrWhiteSpace(rawText) Then
                result.BlockingError = "The pasted backend text is empty."
                Return result
            End If

            Dim normalizedRawText As String = NormalizeAdminLineEndings(rawText)
            Dim workingText As String = PrepareAdminWorkingText(rawText)

            result.WorkingText = workingText

            Dim extractedKeys As New List(Of String)()
            For Each match As Match In Regex.Matches(workingText, "User\s+license\s+key\s*[:\-]?\s*(?<key>[A-Za-z0-9]{16,})", RegexOptions.IgnoreCase)
                Dim keyValue As String = match.Groups("key").Value.Trim()
                If keyValue.Length > 0 AndAlso Not extractedKeys.Any(Function(k) k.Equals(keyValue, StringComparison.Ordinal)) Then
                    extractedKeys.Add(keyValue)
                End If
            Next

            If extractedKeys.Count = 0 Then
                Dim plausibleKeys As New List(Of String)()
                For Each match As Match In Regex.Matches(normalizedRawText, "(?<![A-Za-z0-9])(?<key>[A-Za-z0-9]{20,})(?![A-Za-z0-9])")
                    Dim keyValue As String = match.Groups("key").Value.Trim()
                    If keyValue.Length > 0 AndAlso Not plausibleKeys.Any(Function(k) k.Equals(keyValue, StringComparison.Ordinal)) Then
                        plausibleKeys.Add(keyValue)
                    End If
                Next

                If plausibleKeys.Count > 1 Then
                    result.BlockingError = "Multiple plausible license keys were found. Please paste only one license-key page."
                    Return result
                End If

                If normalizedRawText.IndexOf(storedLicenseKey, StringComparison.Ordinal) >= 0 Then
                    result.ExtractedLicenseKey = storedLicenseKey
                    result.LicenseMatched = True
                    result.Warnings.Add("No explicit license-key anchor was found, but the stored license key appears in the pasted text.")
                ElseIf plausibleKeys.Count = 1 Then
                    result.ExtractedLicenseKey = plausibleKeys(0)
                Else
                    result.BlockingError = "No license key could be found in the pasted backend text."
                    Return result
                End If
            ElseIf extractedKeys.Count = 1 Then
                result.ExtractedLicenseKey = extractedKeys(0)
            Else
                result.BlockingError = "Multiple different license keys were found. Please paste only one license-key page."
                Return result
            End If

            If Not String.IsNullOrWhiteSpace(result.ExtractedLicenseKey) Then
                If Not result.ExtractedLicenseKey.Equals(storedLicenseKey, StringComparison.Ordinal) Then
                    result.BlockingError =
                        "The pasted backend page does not match the stored add-in license key." & vbCrLf & vbCrLf &
                        $"Stored key: {MaskLicenseKey(storedLicenseKey)}" & vbCrLf &
                        $"Pasted key: {MaskLicenseKey(result.ExtractedLicenseKey)}"
                    Return result
                End If
                result.LicenseMatched = True
            End If

            Dim products As List(Of AdminProductCandidate) = ExtractAdminProducts(normalizedRawText)
            For Each product As AdminProductCandidate In products
                result.Products.Add(product)
            Next

            If result.Products.Count = 0 Then
                result.BlockingError = "No product IDs were found in the pasted backend text."
                Return result
            End If

            Dim instances As List(Of AdminInstanceCandidate) = ExtractAdminInstances(workingText)

            AssociateAdminInstancesToProducts(instances, result.Products)

            For Each instanceCandidate As AdminInstanceCandidate In instances
                result.Instances.Add(instanceCandidate)
            Next

            If result.Instances.Count = 0 Then
                result.Warnings.Add("No activation instances were found in the pasted backend text.")
            End If

            Return result
        End Function

        Private Shared Function NormalizeAdminLineEndings(value As String) As String
            If value Is Nothing Then Return ""
            Return value.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
        End Function

        Private Shared Function PrepareAdminWorkingText(value As String) As String
            Dim workingText As String = NormalizeAdminLineEndings(value)

            Dim anchorTokens As String() = {
                "DeleteInstance:",
                "ViewInstance:",
                "Instance:",
                "Product title",
                "Mailing list",
                "About Us",
                "Privacy Notice",
                "Legal Terms",
                "Switch back to",
                "Ask Inky",
                "Questions?×"
            }

            For Each token As String In anchorTokens
                workingText = workingText.Replace(token, vbLf & token)
            Next

            workingText = Regex.Replace(workingText, "[ ]{2,}", " ")
            workingText = workingText.Replace(vbTab, " ")
            Return workingText
        End Function

        Private Shared Function ExtractAdminProducts(normalizedRawText As String) As List(Of AdminProductCandidate)
            Dim products As New List(Of AdminProductCandidate)()
            Dim lines As String() = normalizedRawText.Split({vbLf}, StringSplitOptions.None)
            Dim offset As Integer = 0
            Dim insideProductSection As Boolean = False

            For Each rawLine As String In lines
                Dim line As String = rawLine.Trim()

                If line.IndexOf("Product title", StringComparison.OrdinalIgnoreCase) >= 0 AndAlso
                   line.IndexOf("Product ID", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    insideProductSection = True
                    offset += rawLine.Length + 1
                    Continue For
                End If

                If insideProductSection Then
                    If line.IndexOf("Mailing list", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                       line.IndexOf("About Us", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                       line.IndexOf("Privacy Notice", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                       line.IndexOf("Legal Terms", StringComparison.OrdinalIgnoreCase) >= 0 Then
                        insideProductSection = False
                        offset += rawLine.Length + 1
                        Continue For
                    End If

                    If String.Equals(line, "View", StringComparison.OrdinalIgnoreCase) OrElse
                       line.StartsWith("DeleteInstance:", StringComparison.OrdinalIgnoreCase) OrElse
                       line.StartsWith("ViewInstance:", StringComparison.OrdinalIgnoreCase) OrElse
                       line.StartsWith("Instance:", StringComparison.OrdinalIgnoreCase) Then
                        offset += rawLine.Length + 1
                        Continue For
                    End If

                    Dim product As AdminProductCandidate = TryParseAdminProductLine(rawLine, offset)
                    If product IsNot Nothing Then
                        If Not products.Any(Function(p) p.ProductId.Equals(product.ProductId, StringComparison.Ordinal)) Then
                            products.Add(product)
                        End If
                    End If
                End If

                offset += rawLine.Length + 1
            Next

            If products.Count = 0 Then
                offset = 0
                For Each rawLine As String In lines
                    Dim product As AdminProductCandidate = TryParseAdminProductLine(rawLine, offset)
                    If product IsNot Nothing Then
                        If Not products.Any(Function(p) p.ProductId.Equals(product.ProductId, StringComparison.Ordinal)) Then
                            products.Add(product)
                        End If
                    End If
                    offset += rawLine.Length + 1
                Next
            End If

            Return products
        End Function

        Private Shared Function TryParseAdminProductLine(rawLine As String, offset As Integer) As AdminProductCandidate
            If String.IsNullOrWhiteSpace(rawLine) Then Return Nothing

            Dim line As String = rawLine.Trim()

            If line.StartsWith("DeleteInstance:", StringComparison.OrdinalIgnoreCase) OrElse
               line.StartsWith("ViewInstance:", StringComparison.OrdinalIgnoreCase) OrElse
               line.StartsWith("Instance:", StringComparison.OrdinalIgnoreCase) OrElse
               line.IndexOf("Activated on", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               line.IndexOf("User license key", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               line.IndexOf("Search", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               line.IndexOf("Switch back to", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return Nothing
            End If

            Dim productId As String = ""
            Dim productTitle As String = ""
            Dim activationUsage As String = ""
            Dim nextPayment As String = ""
            Dim confidence As String = "Medium"
            Dim parsedUsedCount As Integer = -1
            Dim parsedMaxCount As Integer = -1

            If rawLine.IndexOf(vbTab, StringComparison.Ordinal) >= 0 Then
                Dim parts As New List(Of String)()
                For Each part As String In rawLine.Split({ControlChars.Tab}, StringSplitOptions.RemoveEmptyEntries)
                    Dim trimmedPart As String = part.Trim()
                    If trimmedPart.Length > 0 Then parts.Add(trimmedPart)
                Next

                If parts.Count >= 2 AndAlso Regex.IsMatch(parts(1), "^\d{2,}$") Then
                    productTitle = parts(0)
                    productId = parts(1)
                    If parts.Count >= 3 Then activationUsage = parts(2)
                    If parts.Count >= 4 Then nextPayment = parts(3)
                    confidence = "High"
                End If
            End If

            If String.IsNullOrWhiteSpace(productId) Then
                Dim match As Match = Regex.Match(line, "^(?<title>.+?)\s+(?<productid>\d{2,})(?<rest>.*)$")
                If match.Success Then
                    productTitle = match.Groups("title").Value.Trim()
                    productId = match.Groups("productid").Value.Trim()

                    Dim rest As String = match.Groups("rest").Value.Trim()
                    Dim activationMatch As Match = Regex.Match(rest, "(?<used>\d+)\s+out\s+of\s+(?<max>\d+)", RegexOptions.IgnoreCase)
                    If activationMatch.Success Then
                        activationUsage = activationMatch.Value.Trim()
                        Integer.TryParse(activationMatch.Groups("used").Value, parsedUsedCount)
                        Integer.TryParse(activationMatch.Groups("max").Value, parsedMaxCount)
                    End If

                    Dim paymentMatch As Match = Regex.Match(rest, "\b\d{1,2}\s+\w+\s+\d{4}\b.*$")
                    If paymentMatch.Success Then
                        nextPayment = paymentMatch.Value.Trim()
                    End If
                End If
            End If

            If String.IsNullOrWhiteSpace(productId) OrElse String.IsNullOrWhiteSpace(productTitle) Then
                Return Nothing
            End If

            If activationUsage.Length = 0 Then
                Dim activationMatch As Match = Regex.Match(line, "(?<used>\d+)\s+out\s+of\s+(?<max>\d+)", RegexOptions.IgnoreCase)
                If activationMatch.Success Then
                    activationUsage = activationMatch.Value.Trim()
                    Integer.TryParse(activationMatch.Groups("used").Value, parsedUsedCount)
                    Integer.TryParse(activationMatch.Groups("max").Value, parsedMaxCount)
                End If
            End If

            Return New AdminProductCandidate() With {
                .ProductId = productId,
                .ProductTitle = productTitle,
                .ActivationUsageText = activationUsage,
                .ParsedUsedCount = parsedUsedCount,
                .ParsedMaxCount = parsedMaxCount,
                .NextPaymentText = nextPayment,
                .Confidence = confidence,
                .SourceOffset = offset
            }
        End Function

        Private Shared Function ExtractAdminInstances(workingText As String) As List(Of AdminInstanceCandidate)
            Dim instances As New List(Of AdminInstanceCandidate)()

            Dim pattern As String =
                "(?is)(?:^|\n)(?<anchor>DeleteInstance:|ViewInstance:|Instance:)\s*(?<body>.*?)(?=(?:\n(?:DeleteInstance:|ViewInstance:|Instance:|Product title|Mailing list|About Us|Privacy Notice|Legal Terms|Switch back to|Ask Inky|Questions\?×))|\z)"

            For Each match As Match In Regex.Matches(workingText, pattern)
                Dim body As String = match.Groups("body").Value.Trim()
                If body.Length = 0 Then Continue For

                Dim activatedOnText As String = ""
                Dim confidence As String = "Medium"

                Dim activatedIndex As Integer = body.IndexOf("Activated on", StringComparison.OrdinalIgnoreCase)
                If activatedIndex >= 0 Then
                    activatedOnText = body.Substring(activatedIndex + "Activated on".Length).Trim()
                    body = body.Substring(0, activatedIndex).Trim()
                    confidence = "High"
                End If

                body = CleanAdminInstanceBoundaryText(body)
                If String.IsNullOrWhiteSpace(body) Then Continue For

                instances.Add(New AdminInstanceCandidate() With {
                    .InstanceUserId = body,
                    .ActivatedOnText = activatedOnText,
                    .Confidence = confidence,
                    .SourceOffset = match.Index
                })
            Next

            Return instances
        End Function

        Private Shared Function CleanAdminInstanceBoundaryText(value As String) As String
            Dim cleaned As String = value.Trim()

            Do While cleaned.StartsWith(":", StringComparison.Ordinal) OrElse cleaned.StartsWith("-", StringComparison.Ordinal)
                cleaned = cleaned.Substring(1).Trim()
            Loop

            cleaned = cleaned.Trim()
            Return cleaned
        End Function

        Private Shared Sub AssociateAdminInstancesToProducts(instances As List(Of AdminInstanceCandidate), products As List(Of AdminProductCandidate))
            If products.Count = 0 Then Return

            Dim orderedProducts = products.OrderBy(Function(p) p.SourceOffset).ToList()

            For Each instanceCandidate As AdminInstanceCandidate In instances
                Dim assignedProduct As AdminProductCandidate = Nothing

                For Each product As AdminProductCandidate In orderedProducts
                    If product.SourceOffset <= instanceCandidate.SourceOffset Then
                        assignedProduct = product
                    Else
                        Exit For
                    End If
                Next

                If assignedProduct IsNot Nothing Then
                    instanceCandidate.ProductId = assignedProduct.ProductId
                    instanceCandidate.ProductTitle = assignedProduct.ProductTitle
                    instanceCandidate.ProductAssociationInferred = False
                    instanceCandidate.IsAmbiguous = False
                ElseIf products.Count = 1 Then
                    instanceCandidate.ProductId = products(0).ProductId
                    instanceCandidate.ProductTitle = products(0).ProductTitle
                    instanceCandidate.ProductAssociationInferred = True
                    instanceCandidate.IsAmbiguous = False
                Else
                    instanceCandidate.ProductId = ""
                    instanceCandidate.ProductTitle = ""
                    instanceCandidate.ProductAssociationInferred = False
                    instanceCandidate.IsAmbiguous = True
                End If
            Next
        End Sub

#End Region

#Region "Organization Bulk License Management - Grid Population"

        Private Shared Function MergeParsedCandidatesIntoRows(rows As BindingList(Of AdminLicenseRow), parseResult As AdminBackendParseResult) As Integer
            For i As Integer = rows.Count - 1 To 0 Step -1
                If rows(i).SourceKind.Equals("Parsed", StringComparison.OrdinalIgnoreCase) Then
                    rows.RemoveAt(i)
                End If
            Next

            Dim addedCount As Integer = 0

            For Each instanceCandidate As AdminInstanceCandidate In parseResult.Instances
                Dim row As AdminLicenseRow = FindAdminRowByProductAndInstance(rows, instanceCandidate.ProductId, instanceCandidate.InstanceUserId)

                If row Is Nothing Then
                    row = New AdminLicenseRow()
                    rows.Add(row)
                    addedCount += 1
                End If

                row.IncludeRow = False
                row.ProductId = instanceCandidate.ProductId
                row.ProductTitle = instanceCandidate.ProductTitle
                row.InstanceUserId = instanceCandidate.InstanceUserId
                row.ParsedActivationDate = instanceCandidate.ActivatedOnText
                row.SourceType = "Parsed backend"
                row.SourceKind = "Parsed"
                row.VerificationState = "Not verified"
                row.DesiredState = "Ignore"
                row.PlannedAction = "Ignore"
                row.ApiResult = ""
                row.WarningText = ""
                row.Confidence = instanceCandidate.Confidence
                row.ProductAssociationInferred = instanceCandidate.ProductAssociationInferred
                row.IsAmbiguous = instanceCandidate.IsAmbiguous
                row.SourceOffset = instanceCandidate.SourceOffset
                row.IsVerifiedActive = False
                row.RequiresReverify = True
                row.DesiredOccurrences = 0
                row.LastKnownActivationsRemaining = -1
                row.LastKnownTotalActivations = -1
                row.LastKnownTotalPurchased = -1

                If instanceCandidate.IsAmbiguous Then
                    AppendDistinctAdminWarning(row, "Ambiguous product association. Review the Product ID before verification.")
                End If

                If instanceCandidate.ProductAssociationInferred Then
                    AppendDistinctAdminWarning(row, "Product association was inferred because only one product ID was parsed.")
                End If
            Next

            Return addedCount
        End Function

        Private Shared Function FindAdminRowByProductAndInstance(rows As IEnumerable(Of AdminLicenseRow), productId As String, instanceUserId As String) As AdminLicenseRow
            Dim safeProductId As String = If(productId, "")
            Dim safeInstanceUserId As String = If(instanceUserId, "")

            Return rows.FirstOrDefault(
                Function(r) String.Equals(If(r.ProductId, ""), safeProductId, StringComparison.Ordinal) AndAlso
                            String.Equals(If(r.InstanceUserId, ""), safeInstanceUserId, StringComparison.Ordinal))
        End Function

        Private Shared Function AdminRowRepresentsExistingActivation(row As AdminLicenseRow) As Boolean
            If row Is Nothing Then
                Return False
            End If

            Return row.IsVerifiedActive OrElse
                   String.Equals(If(row.VerificationState, ""), "Verified active", StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(If(row.SourceKind, ""), "Parsed", StringComparison.OrdinalIgnoreCase) OrElse
                   If(row.SourceType, "").IndexOf("Parsed backend", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                   Not String.IsNullOrWhiteSpace(row.ParsedActivationDate)
        End Function

        Private Shared Function UpsertDesiredRow(rows As BindingList(Of AdminLicenseRow),
                                                 productId As String,
                                                 productTitle As String,
                                                 instanceUserId As String,
                                                 sourceType As String,
                                                 sourceKind As String) As AdminLicenseRow

            Dim row As AdminLicenseRow = FindAdminRowByProductAndInstance(rows, productId, instanceUserId)

            If row Is Nothing Then
                row = New AdminLicenseRow() With {
                    .IncludeRow = True,
                    .ProductId = productId,
                    .ProductTitle = productTitle,
                    .InstanceUserId = instanceUserId,
                    .ParsedActivationDate = "",
                    .SourceType = sourceType,
                    .SourceKind = sourceKind,
                    .VerificationState = "Not verified",
                    .DesiredState = "DesiredActive",
                    .PlannedAction = "Pending activation",
                    .ApiResult = "",
                    .WarningText = "",
                    .Confidence = "",
                    .ProductAssociationInferred = False,
                    .IsAmbiguous = False,
                    .SourceOffset = 0,
                    .IsVerifiedActive = False,
                    .RequiresReverify = True,
                    .DesiredOccurrences = 1
                }
                rows.Add(row)
            Else
                If String.IsNullOrWhiteSpace(row.SourceType) Then
                    row.SourceType = sourceType
                ElseIf row.SourceType.IndexOf(sourceType, StringComparison.OrdinalIgnoreCase) < 0 Then
                    row.SourceType = $"{row.SourceType} + {sourceType}"
                End If

                row.SourceKind = sourceKind
                row.IncludeRow = True
                row.RequiresReverify = True
                row.ApiResult = ""
                row.DesiredOccurrences = Math.Max(1, row.DesiredOccurrences + 1)

                If AdminRowRepresentsExistingActivation(row) Then
                    row.DesiredState = "DesiredInactive"
                    row.PlannedAction = "Pending deactivation"
                Else
                    row.DesiredState = "DesiredActive"
                    row.PlannedAction = "Pending activation"
                End If

                If String.IsNullOrWhiteSpace(row.ProductTitle) AndAlso Not String.IsNullOrWhiteSpace(productTitle) Then
                    row.ProductTitle = productTitle
                End If
            End If

            Return row
        End Function

        Private Shared Sub ApplyAdminDuplicateWarnings(rows As IEnumerable(Of AdminLicenseRow))
            Dim byProductAndInstance As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
            Dim byInstanceProducts As New Dictionary(Of String, HashSet(Of String))(StringComparer.Ordinal)

            For Each row As AdminLicenseRow In rows
                If String.IsNullOrWhiteSpace(row.ProductId) OrElse String.IsNullOrWhiteSpace(row.InstanceUserId) Then Continue For

                Dim key As String = BuildAdminKey(row.ProductId, row.InstanceUserId)
                If Not byProductAndInstance.ContainsKey(key) Then
                    byProductAndInstance(key) = 0
                End If
                byProductAndInstance(key) += 1

                If Not byInstanceProducts.ContainsKey(row.InstanceUserId) Then
                    byInstanceProducts(row.InstanceUserId) = New HashSet(Of String)(StringComparer.Ordinal)
                End If
                byInstanceProducts(row.InstanceUserId).Add(row.ProductId)
            Next

            For Each row As AdminLicenseRow In rows
                If String.IsNullOrWhiteSpace(row.ProductId) OrElse String.IsNullOrWhiteSpace(row.InstanceUserId) Then Continue For

                Dim key As String = BuildAdminKey(row.ProductId, row.InstanceUserId)

                If byProductAndInstance.ContainsKey(key) AndAlso byProductAndInstance(key) > 1 Then
                    AppendDistinctAdminWarning(row, "Duplicate instance under the same product ID.")
                End If

                If byInstanceProducts.ContainsKey(row.InstanceUserId) AndAlso byInstanceProducts(row.InstanceUserId).Count > 1 Then
                    AppendDistinctAdminWarning(row, "This same instance appears under multiple product IDs.")
                End If
            Next
        End Sub

#End Region

#Region "Organization Bulk License Management - Verification / Planning / Apply"

        Private Shared Function VerifyAdminRows(rows As IEnumerable(Of AdminLicenseRow),
                                                licenseKey As String,
                                                allowedProductIds As HashSet(Of String),
                                                Optional progressCallback As Action(Of Integer, Integer, String) = Nothing) As Integer

            Dim verifiedCount As Integer = 0
            Dim responseCache As New Dictionary(Of String, LicenseApiResponse)(StringComparer.Ordinal)
            Dim keysToVerify As New HashSet(Of String)(StringComparer.Ordinal)

            For Each row As AdminLicenseRow In rows
                If Not String.IsNullOrWhiteSpace(row.ProductId) AndAlso
                   Not String.IsNullOrWhiteSpace(row.InstanceUserId) AndAlso
                   Not row.IsAmbiguous AndAlso
                   (allowedProductIds Is Nothing OrElse allowedProductIds.Count = 0 OrElse allowedProductIds.Contains(row.ProductId.Trim())) Then
                    keysToVerify.Add(BuildAdminKey(row.ProductId.Trim(), row.InstanceUserId.Trim()))
                End If
            Next

            Dim totalApiCalls As Integer = keysToVerify.Count
            Dim currentApiCall As Integer = 0

            For Each row As AdminLicenseRow In rows
                row.ApiResult = ""

                If String.IsNullOrWhiteSpace(row.ProductId) Then
                    row.VerificationState = "Invalid/missing product"
                    row.IsVerifiedActive = False
                    AppendDistinctAdminWarning(row, "No Product ID is assigned.")
                    Continue For
                End If

                If allowedProductIds IsNot Nothing AndAlso allowedProductIds.Count > 0 AndAlso Not allowedProductIds.Contains(row.ProductId.Trim()) Then
                    row.VerificationState = "Invalid/missing product"
                    row.IsVerifiedActive = False
                    AppendDistinctAdminWarning(row, "Product ID is not in the parsed backend page and cannot be used.")
                    Continue For
                End If

                If String.IsNullOrWhiteSpace(row.InstanceUserId) Then
                    row.VerificationState = "Invalid/missing user"
                    row.IsVerifiedActive = False
                    AppendDistinctAdminWarning(row, "Instance / User ID is missing.")
                    Continue For
                End If

                If row.IsAmbiguous Then
                    row.VerificationState = "Ambiguous product"
                    row.IsVerifiedActive = False
                    AppendDistinctAdminWarning(row, "Resolve the product assignment before verification.")
                    Continue For
                End If

                Dim cacheKey As String = BuildAdminKey(row.ProductId.Trim(), row.InstanceUserId.Trim())
                Dim response As LicenseApiResponse = Nothing

                If Not responseCache.TryGetValue(cacheKey, response) Then
                    currentApiCall += 1

                    If progressCallback IsNot Nothing Then
                        progressCallback(
                            currentApiCall,
                            totalApiCalls,
                            $"Verifying Product ID {row.ProductId.Trim()}, User ID '{row.InstanceUserId.Trim()}'...")
                    End If

                    response = CallLicenseApi("status", row.ProductId.Trim(), licenseKey, row.InstanceUserId.Trim())
                    responseCache(cacheKey) = response
                End If

                If response.Success Then
                    row.IsVerifiedActive = response.Activated
                    row.VerificationState = If(response.Activated, "Verified active", "Verified not active")
                    row.RequiresReverify = False
                    row.LastKnownActivationsRemaining = response.ActivationsRemaining
                    row.LastKnownTotalActivations = response.TotalActivations
                    row.LastKnownTotalPurchased = response.TotalActivationsPurchased
                    row.ApiResult = If(String.IsNullOrWhiteSpace(response.StatusCheck), If(response.Activated, "active", "inactive"), response.StatusCheck)

                    If Not String.IsNullOrWhiteSpace(response.ProductTitle) AndAlso response.ProductTitle <> "(no product title available)" Then
                        row.ProductTitle = response.ProductTitle
                    End If
                Else
                    row.IsVerifiedActive = False
                    row.VerificationState = ClassifyAdminApiFailure(response.ErrorMessage)
                    row.RequiresReverify = True
                    row.ApiResult = TruncateForLog(response.ErrorMessage, 180)
                    AppendDistinctAdminWarning(row, ExtractShortAdminApiMessage(response.ErrorMessage))
                End If

                verifiedCount += 1
            Next

            Return verifiedCount
        End Function

        Private Shared Function ClassifyAdminApiFailure(errorMessage As String) As String
            Dim message As String = If(errorMessage, "")

            If message.IndexOf("product", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return "Product mismatch"
            End If

            If message.IndexOf("license", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               message.IndexOf("api key", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return "License error"
            End If

            Return "API error"
        End Function

        Private Shared Function ExtractShortAdminApiMessage(errorMessage As String) As String
            If String.IsNullOrWhiteSpace(errorMessage) Then
                Return "API call failed."
            End If

            Dim message As String = errorMessage.Trim()
            If message.Length <= 140 Then
                Return message
            End If

            Return message.Substring(0, 137) & "..."
        End Function

        Private Shared Function PreviewAdminSynchronization(rows As IList(Of AdminLicenseRow),
                                                            allowedProductIds As HashSet(Of String)) As String

            Dim desiredCounts As New Dictionary(Of String, Integer)(StringComparer.Ordinal)

            For Each row As AdminLicenseRow In rows
                If row.IncludeRow AndAlso
                   row.DesiredState.Equals("DesiredActive", StringComparison.OrdinalIgnoreCase) AndAlso
                   Not String.IsNullOrWhiteSpace(row.ProductId) AndAlso
                   Not String.IsNullOrWhiteSpace(row.InstanceUserId) Then
                    Dim key As String = BuildAdminKey(row.ProductId, row.InstanceUserId)
                    If Not desiredCounts.ContainsKey(key) Then
                        desiredCounts(key) = 0
                    End If
                    desiredCounts(key) += 1
                End If
            Next

            For Each row As AdminLicenseRow In rows
                If Not row.IncludeRow Then
                    row.PlannedAction = "Not selected"
                    Continue For
                End If

                If String.IsNullOrWhiteSpace(row.ProductId) Then
                    row.PlannedAction = "Invalid/missing product"
                    Continue For
                End If

                If allowedProductIds IsNot Nothing AndAlso allowedProductIds.Count > 0 AndAlso Not allowedProductIds.Contains(row.ProductId.Trim()) Then
                    row.PlannedAction = "Invalid/missing product"
                    AppendDistinctAdminWarning(row, "Product ID is not part of the parsed backend page.")
                    Continue For
                End If

                If String.IsNullOrWhiteSpace(row.InstanceUserId) Then
                    row.PlannedAction = "Invalid/missing user"
                    Continue For
                End If

                If row.IsAmbiguous Then
                    row.PlannedAction = "Ambiguous product"
                    Continue For
                End If

                If row.RequiresReverify Then
                    row.PlannedAction = "Needs re-verification"
                    Continue For
                End If

                Dim key As String = BuildAdminKey(row.ProductId, row.InstanceUserId)

                If row.DesiredState.Equals("DesiredActive", StringComparison.OrdinalIgnoreCase) AndAlso
                   desiredCounts.ContainsKey(key) AndAlso desiredCounts(key) > 1 Then
                    row.PlannedAction = "Duplicate desired user"
                    AppendDistinctAdminWarning(row, "Duplicate desired user under the same product.")
                    Continue For
                End If

                If row.DesiredState.Equals("DesiredInactive", StringComparison.OrdinalIgnoreCase) Then
                    If row.VerificationState.Equals("Verified active", StringComparison.OrdinalIgnoreCase) Then
                        row.PlannedAction = "Deactivate"
                    ElseIf row.VerificationState.Equals("Verified not active", StringComparison.OrdinalIgnoreCase) Then
                        row.PlannedAction = "Already inactive"
                    Else
                        row.PlannedAction = row.VerificationState
                    End If
                    Continue For
                End If

                If row.DesiredState.Equals("DesiredActive", StringComparison.OrdinalIgnoreCase) Then
                    If row.VerificationState.Equals("Verified active", StringComparison.OrdinalIgnoreCase) Then
                        row.PlannedAction = "Keep active"
                    ElseIf row.VerificationState.Equals("Product mismatch", StringComparison.OrdinalIgnoreCase) OrElse
                           row.VerificationState.Equals("License error", StringComparison.OrdinalIgnoreCase) OrElse
                           row.VerificationState.Equals("API error", StringComparison.OrdinalIgnoreCase) Then
                        row.PlannedAction = row.VerificationState
                    Else
                        row.PlannedAction = "Activate"
                    End If
                    Continue For
                End If

                row.PlannedAction = "Ignore"
            Next

            Dim capacitySummary As String = ApplyAdminCapacityChecks(rows.Where(Function(r) r.IncludeRow))
            Return BuildAdminPlanSummary(rows, capacitySummary)
        End Function

        Private Shared Function ApplyAdminCapacityChecks(rows As IEnumerable(Of AdminLicenseRow)) As String
            Dim warnings As New List(Of String)()
            Dim productIds As List(Of String) = rows.
                Where(Function(r) Not String.IsNullOrWhiteSpace(r.ProductId)).
                Select(Function(r) r.ProductId).
                Distinct(StringComparer.Ordinal).
                ToList()

            For Each productId As String In productIds
                Dim productRows = rows.Where(Function(r) r.ProductId.Equals(productId, StringComparison.Ordinal)).ToList()
                Dim plannedActivations As Integer = Enumerable.Count(productRows, Function(r) r.IncludeRow AndAlso r.PlannedAction.Equals("Activate", StringComparison.OrdinalIgnoreCase))
                Dim plannedDeactivations As Integer = Enumerable.Count(productRows, Function(r) r.IncludeRow AndAlso r.PlannedAction.Equals("Deactivate", StringComparison.OrdinalIgnoreCase))

                If plannedActivations = 0 Then Continue For

                Dim knownRemainingValues = productRows.
                    Where(Function(r) r.LastKnownActivationsRemaining >= 0).
                    Select(Function(r) r.LastKnownActivationsRemaining).
                    ToList()

                If knownRemainingValues.Count = 0 Then
                    warnings.Add($"Product {productId}: capacity could not be verified from status responses.")
                    Continue For
                End If

                Dim availableSlots As Integer = knownRemainingValues.Max() + plannedDeactivations

                If plannedActivations > availableSlots Then
                    warnings.Add($"Product {productId}: planned activations ({plannedActivations}) exceed verified available slots after planned deactivations ({availableSlots}).")

                    For Each row As AdminLicenseRow In productRows.Where(Function(r) r.PlannedAction.Equals("Activate", StringComparison.OrdinalIgnoreCase))
                        row.PlannedAction = "Capacity problem"
                        row.IncludeRow = False
                        AppendDistinctAdminWarning(row, "Planned activations exceed verified available slots for this product.")
                    Next
                End If
            Next

            Return String.Join(vbCrLf, warnings)
        End Function

        Private Shared Function BuildAdminPlanSummary(rows As IEnumerable(Of AdminLicenseRow),
                                                      capacitySummary As String) As String

            Dim sb As New System.Text.StringBuilder()

            Dim selectedRows = rows.Where(Function(r) r.IncludeRow).OrderBy(Function(r) r.ProductId).ThenBy(Function(r) r.InstanceUserId).ToList()
            Dim activateRows = selectedRows.Where(Function(r) r.PlannedAction.Equals("Activate", StringComparison.OrdinalIgnoreCase)).ToList()
            Dim deactivateRows = selectedRows.Where(Function(r) r.PlannedAction.Equals("Deactivate", StringComparison.OrdinalIgnoreCase)).ToList()
            Dim keepRows = selectedRows.Where(Function(r) r.PlannedAction.Equals("Keep active", StringComparison.OrdinalIgnoreCase)).ToList()

            sb.AppendLine("Preview of selected changes")
            sb.AppendLine("---------------------------")
            sb.AppendLine($"Selected rows: {selectedRows.Count}")
            sb.AppendLine($"Activate: {activateRows.Count}")
            sb.AppendLine($"Deactivate: {deactivateRows.Count}")
            sb.AppendLine($"Keep active: {keepRows.Count}")

            sb.AppendLine()
            sb.AppendLine("Selected rows and planned actions:")
            If selectedRows.Count = 0 Then
                sb.AppendLine("- No rows are currently selected.")
            Else
                For Each row As AdminLicenseRow In selectedRows
                    sb.AppendLine($"- Product {If(String.IsNullOrWhiteSpace(row.ProductId), "(none)", row.ProductId)}, User '{row.InstanceUserId}': {row.PlannedAction}")
                Next
            End If

            If Not String.IsNullOrWhiteSpace(capacitySummary) Then
                sb.AppendLine()
                sb.AppendLine("Capacity warnings:")
                sb.AppendLine(capacitySummary)
            End If

            Return sb.ToString().TrimEnd()
        End Function

        Private Shared Function ApplyAdminPlannedChanges(rows As IList(Of AdminLicenseRow),
                                                         licenseKey As String,
                                                         allowedProductIds As HashSet(Of String),
                                                         ByRef correlationId As String,
                                                         Optional progressCallback As Action(Of Integer, Integer, String) = Nothing) As String

            correlationId = System.Guid.NewGuid().ToString("N")

            Dim activateRows = rows.Where(Function(r) r.IncludeRow AndAlso r.PlannedAction.Equals("Activate", StringComparison.OrdinalIgnoreCase)).ToList()
            Dim deactivateRows = rows.Where(Function(r) r.IncludeRow AndAlso r.PlannedAction.Equals("Deactivate", StringComparison.OrdinalIgnoreCase)).ToList()

            Dim activateSuccess As Integer = 0
            Dim deactivateSuccess As Integer = 0
            Dim errorCount As Integer = 0
            Dim totalActions As Integer = deactivateRows.Count + activateRows.Count
            Dim currentAction As Integer = 0

            For Each row As AdminLicenseRow In deactivateRows
                currentAction += 1
                If progressCallback IsNot Nothing Then
                    progressCallback(currentAction, totalActions, $"Deactivating Product ID {row.ProductId}, User ID '{row.InstanceUserId}'...")
                End If

                row.LastCorrelationId = correlationId
                LogLicenseEvent("Organization Bulk", $"CorrelationID={correlationId}, Action=deactivate, ProductID={row.ProductId}, UserID={row.InstanceUserId}", alwaysLog:=True)

                Dim response As LicenseApiResponse = CallLicenseApi("deactivate", row.ProductId, licenseKey, row.InstanceUserId)

                If response.Success Then
                    Dim statusResponse As LicenseApiResponse = CallLicenseApi("status", row.ProductId, licenseKey, row.InstanceUserId)

                    If statusResponse.Success AndAlso Not statusResponse.Activated Then
                        row.VerificationState = "Verified not active"
                        row.IsVerifiedActive = False
                        row.RequiresReverify = False
                        row.ApiResult = "Deactivated"
                        row.PlannedAction = "Deactivated"
                        row.IncludeRow = False
                        row.LastKnownActivationsRemaining = statusResponse.ActivationsRemaining
                        row.LastKnownTotalActivations = statusResponse.TotalActivations
                        row.LastKnownTotalPurchased = statusResponse.TotalActivationsPurchased
                        deactivateSuccess += 1
                        LogLicenseEvent("Organization Bulk", $"CorrelationID={correlationId}, Result=success, Action=deactivate, ProductID={row.ProductId}, UserID={row.InstanceUserId}", alwaysLog:=True)
                    Else
                        row.VerificationState = "API error"
                        row.ApiResult = "Deactivation could not be verified"
                        row.PlannedAction = "Deactivation verify failed"
                        row.RequiresReverify = True
                        errorCount += 1
                        LogLicenseEvent("Organization Bulk", $"CorrelationID={correlationId}, Result=verify-failed, Action=deactivate, ProductID={row.ProductId}, UserID={row.InstanceUserId}", alwaysLog:=True)
                    End If
                Else
                    row.VerificationState = ClassifyAdminApiFailure(response.ErrorMessage)
                    row.ApiResult = ExtractShortAdminApiMessage(response.ErrorMessage)
                    row.PlannedAction = "Deactivation failed"
                    row.RequiresReverify = True
                    AppendDistinctAdminWarning(row, ExtractShortAdminApiMessage(response.ErrorMessage))
                    errorCount += 1
                    LogLicenseEvent("Organization Bulk", $"CorrelationID={correlationId}, Result=error, Action=deactivate, ProductID={row.ProductId}, UserID={row.InstanceUserId}, Message={TruncateForLog(response.ErrorMessage, 160)}", alwaysLog:=True)
                End If
            Next

            For Each row As AdminLicenseRow In activateRows
                currentAction += 1
                If progressCallback IsNot Nothing Then
                    progressCallback(currentAction, totalActions, $"Activating Product ID {row.ProductId}, User ID '{row.InstanceUserId}'...")
                End If

                row.LastCorrelationId = correlationId
                LogLicenseEvent("Organization Bulk", $"CorrelationID={correlationId}, Action=activate, ProductID={row.ProductId}, UserID={row.InstanceUserId}", alwaysLog:=True)

                Dim response As LicenseApiResponse = CallLicenseApi("activate", row.ProductId, licenseKey, row.InstanceUserId)

                If response.Success AndAlso response.Activated Then
                    Dim statusResponse As LicenseApiResponse = CallLicenseApi("status", row.ProductId, licenseKey, row.InstanceUserId)

                    If statusResponse.Success AndAlso statusResponse.Activated Then
                        row.VerificationState = "Verified active"
                        row.IsVerifiedActive = True
                        row.RequiresReverify = False
                        row.ApiResult = "Activated"
                        row.PlannedAction = "Activated"
                        row.IncludeRow = False
                        row.LastKnownActivationsRemaining = statusResponse.ActivationsRemaining
                        row.LastKnownTotalActivations = statusResponse.TotalActivations
                        row.LastKnownTotalPurchased = statusResponse.TotalActivationsPurchased
                        If Not String.IsNullOrWhiteSpace(statusResponse.ProductTitle) AndAlso statusResponse.ProductTitle <> "(no product title available)" Then
                            row.ProductTitle = statusResponse.ProductTitle
                        End If
                        activateSuccess += 1
                        LogLicenseEvent("Organization Bulk", $"CorrelationID={correlationId}, Result=success, Action=activate, ProductID={row.ProductId}, UserID={row.InstanceUserId}", alwaysLog:=True)
                    Else
                        row.VerificationState = "API error"
                        row.ApiResult = "Activation could not be verified"
                        row.PlannedAction = "Activation verify failed"
                        row.RequiresReverify = True
                        errorCount += 1
                        LogLicenseEvent("Organization Bulk", $"CorrelationID={correlationId}, Result=verify-failed, Action=activate, ProductID={row.ProductId}, UserID={row.InstanceUserId}", alwaysLog:=True)
                    End If
                Else
                    Dim statusResponse As LicenseApiResponse = CallLicenseApi("status", row.ProductId, licenseKey, row.InstanceUserId)

                    If statusResponse.Success AndAlso statusResponse.Activated Then
                        row.VerificationState = "Verified active"
                        row.IsVerifiedActive = True
                        row.RequiresReverify = False
                        row.ApiResult = "Already active"
                        row.PlannedAction = "Already active"
                        row.IncludeRow = False
                        row.LastKnownActivationsRemaining = statusResponse.ActivationsRemaining
                        row.LastKnownTotalActivations = statusResponse.TotalActivations
                        row.LastKnownTotalPurchased = statusResponse.TotalActivationsPurchased
                        activateSuccess += 1
                        LogLicenseEvent("Organization Bulk", $"CorrelationID={correlationId}, Result=already-active, Action=activate, ProductID={row.ProductId}, UserID={row.InstanceUserId}", alwaysLog:=True)
                    Else
                        row.VerificationState = ClassifyAdminApiFailure(response.ErrorMessage)
                        row.ApiResult = ExtractShortAdminApiMessage(response.ErrorMessage)
                        row.PlannedAction = "Activation failed"
                        row.RequiresReverify = True

                        If statusResponse.Success AndAlso statusResponse.ActivationsRemaining <= 0 Then
                            row.PlannedAction = "Capacity problem"
                            AppendDistinctAdminWarning(row, "No slots remaining for this product.")
                        Else
                            AppendDistinctAdminWarning(row, ExtractShortAdminApiMessage(response.ErrorMessage))
                        End If

                        errorCount += 1
                        LogLicenseEvent("Organization Bulk", $"CorrelationID={correlationId}, Result=error, Action=activate, ProductID={row.ProductId}, UserID={row.InstanceUserId}, Message={TruncateForLog(response.ErrorMessage, 160)}", alwaysLog:=True)
                    End If
                End If
            Next

            Dim sb As New System.Text.StringBuilder()
            Dim appliedCorrelationId As String = correlationId

            sb.AppendLine("Bulk license operation completed.")
            sb.AppendLine($"Correlation ID: {appliedCorrelationId}")
            sb.AppendLine($"Activated successfully: {activateSuccess}")
            sb.AppendLine($"Deactivated successfully: {deactivateSuccess}")
            sb.AppendLine($"Errors / partial failures: {errorCount}")

            If totalActions > 0 Then
                sb.AppendLine()
                sb.AppendLine("Processed rows:")
                For Each row As AdminLicenseRow In rows.Where(Function(r) Not String.IsNullOrWhiteSpace(r.LastCorrelationId) AndAlso r.LastCorrelationId.Equals(appliedCorrelationId, StringComparison.Ordinal))
                    sb.AppendLine($"- Product {row.ProductId}, User '{row.InstanceUserId}': {row.ApiResult}")
                Next
            End If

            Return sb.ToString().TrimEnd()
        End Function

#End Region

#Region "Organization Bulk License Management - Import / Mapping"

        Private Shared Function ImportDesiredRowsFromFile(rows As BindingList(Of AdminLicenseRow),
                                                          allowedProductIds As HashSet(Of String),
                                                          parsedProductTitles As Dictionary(Of String, String),
                                                          mappingConfig As AdminUserIdMappingConfig,
                                                          ByRef markedForDeactivationCount As Integer) As Integer

            markedForDeactivationCount = 0

            Dim fileName As String = ""

            Using dialog As New System.Windows.Forms.OpenFileDialog()
                dialog.Title = $"{AN} - Sync with Imported List"
                dialog.Filter = "Supported files (*.csv;*.txt;*.json)|*.csv;*.txt;*.json|CSV files (*.csv)|*.csv|Text files (*.txt)|*.txt|JSON files (*.json)|*.json|All files (*.*)|*.*"
                dialog.Multiselect = False

                Dim __safeDialogOwner2059 As System.Windows.Forms.IWin32Window = Global.SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
                If If(__safeDialogOwner2059 IsNot Nothing, dialog.ShowDialog(__safeDialogOwner2059), dialog.ShowDialog()) <> System.Windows.Forms.DialogResult.OK Then
                    Return 0
                End If

                fileName = dialog.FileName
            End Using

            Dim records As List(Of Dictionary(Of String, String)) = LoadAdminImportRecords(fileName)
            If records.Count = 0 Then
                ShowCustomMessageBox("No importable rows were found in the selected file.", $"{AN} - Organization License Manager")
                Return 0
            End If

            Dim defaultProductId As String = SelectDefaultAdminProductId(allowedProductIds, parsedProductTitles, "Select the default product for imported users.")
            If String.IsNullOrWhiteSpace(defaultProductId) Then
                Return 0
            End If

            Dim pendingRows As New List(Of AdminPendingDesiredRow)()
            Dim previewLines As New List(Of String)()

            For Each record As Dictionary(Of String, String) In records
                Dim resolvedProductId As String = ResolveAdminImportedProductId(record, defaultProductId, allowedProductIds)
                Dim resolvedProductTitle As String = GetProductTitleFromLookup(parsedProductTitles, resolvedProductId)
                Dim generatedUserId As String = GenerateAdminUserIdFromRecord(record, mappingConfig)
                Dim pending As New AdminPendingDesiredRow() With {
                    .ProductId = resolvedProductId,
                    .ProductTitle = resolvedProductTitle,
                    .InstanceUserId = generatedUserId,
                    .SourceType = "Imported sync list",
                    .SourceKind = "Import"
                }

                If String.IsNullOrWhiteSpace(resolvedProductId) Then
                    pending.WarningText = "No valid Product ID could be assigned from the parsed backend products."
                ElseIf String.IsNullOrWhiteSpace(generatedUserId) Then
                    pending.WarningText = "No User ID could be generated from the imported record and current mapping rules."
                End If

                pendingRows.Add(pending)

                If previewLines.Count < 12 Then
                    previewLines.Add($"{If(String.IsNullOrWhiteSpace(resolvedProductId), "(no product)", resolvedProductId)}: {If(String.IsNullOrWhiteSpace(generatedUserId), "(blank)", generatedUserId)}")
                End If
            Next

            Dim previewMessage As New System.Text.StringBuilder()
            previewMessage.AppendLine($"Sync file: {System.IO.Path.GetFileName(fileName)}")
            previewMessage.AppendLine($"Records parsed: {records.Count}")
            previewMessage.AppendLine()
            previewMessage.AppendLine("Preview of generated User IDs:")
            For Each previewLine As String In previewLines
                previewMessage.AppendLine($"- {previewLine}")
            Next
            If pendingRows.Count > previewLines.Count Then
                previewMessage.AppendLine($"- ... and {pendingRows.Count - previewLines.Count} more")
            End If
            previewMessage.AppendLine()
            previewMessage.AppendLine("Proceed with the synchronization?")
            previewMessage.AppendLine("Users in the imported list will remain active or be added. Existing activations for synchronized product IDs that are not present in the imported list will be marked for deactivation.")

            Dim proceed As Integer = ShowCustomYesNoBox(previewMessage.ToString().TrimEnd(), "Sync", "Cancel", $"{AN} - Sync with Imported List")
            If proceed <> 1 Then
                Return 0
            End If

            Dim desiredKeys As New HashSet(Of String)(StringComparer.Ordinal)
            Dim synchronizedProductIds As New HashSet(Of String)(StringComparer.Ordinal)
            Dim importedCount As Integer = 0

            For Each pending As AdminPendingDesiredRow In pendingRows
                If String.IsNullOrWhiteSpace(pending.WarningText) Then
                    Dim syncKey As String = BuildAdminKey(pending.ProductId.Trim(), pending.InstanceUserId.Trim())

                    If desiredKeys.Add(syncKey) Then
                        synchronizedProductIds.Add(pending.ProductId.Trim())

                        Dim row As AdminLicenseRow = FindAdminRowByProductAndInstance(rows, pending.ProductId, pending.InstanceUserId)

                        If row Is Nothing Then
                            row = New AdminLicenseRow() With {
                                .IncludeRow = True,
                                .ProductId = pending.ProductId,
                                .ProductTitle = pending.ProductTitle,
                                .InstanceUserId = pending.InstanceUserId,
                                .ParsedActivationDate = "",
                                .SourceType = pending.SourceType,
                                .SourceKind = pending.SourceKind,
                                .VerificationState = "Not verified",
                                .DesiredState = "DesiredActive",
                                .PlannedAction = "Pending activation",
                                .ApiResult = "",
                                .WarningText = "",
                                .Confidence = "",
                                .ProductAssociationInferred = False,
                                .IsAmbiguous = False,
                                .SourceOffset = 0,
                                .IsVerifiedActive = False,
                                .RequiresReverify = True,
                                .DesiredOccurrences = 1
                            }
                            rows.Add(row)
                        Else
                            Dim rowRepresentsExistingActivation As Boolean = AdminRowRepresentsExistingActivation(row)

                            If String.IsNullOrWhiteSpace(row.SourceType) Then
                                row.SourceType = pending.SourceType
                            ElseIf row.SourceType.IndexOf(pending.SourceType, StringComparison.OrdinalIgnoreCase) < 0 Then
                                row.SourceType = $"{row.SourceType} + {pending.SourceType}"
                            End If

                            If rowRepresentsExistingActivation Then
                                row.IncludeRow = False
                                row.DesiredState = "Ignore"
                                row.PlannedAction = "Ignore"
                                row.ApiResult = ""
                                row.RequiresReverify = False
                                row.DesiredOccurrences = 1
                            Else
                                row.SourceKind = pending.SourceKind
                                row.IncludeRow = True
                                row.DesiredState = "DesiredActive"
                                row.PlannedAction = "Pending activation"
                                row.ApiResult = ""
                                row.RequiresReverify = True
                                row.DesiredOccurrences = 1
                            End If

                            If String.IsNullOrWhiteSpace(row.ProductTitle) AndAlso Not String.IsNullOrWhiteSpace(pending.ProductTitle) Then
                                row.ProductTitle = pending.ProductTitle
                            End If
                        End If

                        importedCount += 1
                    Else
                        Dim duplicateRow As AdminLicenseRow = FindAdminRowByProductAndInstance(rows, pending.ProductId, pending.InstanceUserId)
                        If duplicateRow IsNot Nothing Then
                            AppendDistinctAdminWarning(duplicateRow, "Duplicate imported user under the same product. Only one synchronized entry was kept.")
                        End If
                    End If
                Else
                    rows.Add(New AdminLicenseRow() With {
                        .IncludeRow = False,
                        .ProductId = pending.ProductId,
                        .ProductTitle = pending.ProductTitle,
                        .InstanceUserId = pending.InstanceUserId,
                        .ParsedActivationDate = "",
                        .SourceType = pending.SourceType,
                        .SourceKind = pending.SourceKind,
                        .VerificationState = "Not verified",
                        .DesiredState = "Ignore",
                        .PlannedAction = "Invalid/missing user",
                        .ApiResult = "",
                        .WarningText = pending.WarningText,
                        .Confidence = "",
                        .RequiresReverify = False
                    })
                End If
            Next

            For Each row As AdminLicenseRow In rows
                If row Is Nothing Then
                    Continue For
                End If

                If String.IsNullOrWhiteSpace(row.ProductId) OrElse String.IsNullOrWhiteSpace(row.InstanceUserId) Then
                    Continue For
                End If

                If synchronizedProductIds.Count > 0 AndAlso Not synchronizedProductIds.Contains(row.ProductId.Trim()) Then
                    Continue For
                End If

                Dim syncKey As String = BuildAdminKey(row.ProductId.Trim(), row.InstanceUserId.Trim())

                If desiredKeys.Contains(syncKey) Then
                    Continue For
                End If

                If AdminRowRepresentsExistingActivation(row) Then
                    row.IncludeRow = True
                    row.DesiredState = "DesiredInactive"
                    row.PlannedAction = "Pending deactivation"
                    row.ApiResult = ""
                    row.RequiresReverify = True
                    row.DesiredOccurrences = 0
                    markedForDeactivationCount += 1
                Else
                    row.IncludeRow = False
                    row.DesiredState = "Ignore"
                    row.PlannedAction = "Ignore"
                    row.ApiResult = ""
                    row.RequiresReverify = True
                    row.DesiredOccurrences = 0
                End If
            Next

            Return importedCount
        End Function

        Private Shared Function LoadAdminImportRecords(fileName As String) As List(Of Dictionary(Of String, String))
            Dim extension As String = System.IO.Path.GetExtension(fileName).ToLowerInvariant()

            Select Case extension
                Case ".csv"
                    Return LoadAdminDelimitedFile(fileName, DetectAdminDelimiter(System.IO.File.ReadLines(fileName).FirstOrDefault(Function(l) Not String.IsNullOrWhiteSpace(l))))
                Case ".txt"
                    Dim firstLine As String = System.IO.File.ReadLines(fileName).FirstOrDefault(Function(l) Not String.IsNullOrWhiteSpace(l))
                    Dim delimiter As Char = DetectAdminDelimiter(firstLine)

                    If delimiter = ControlChars.NullChar Then
                        Dim records As New List(Of Dictionary(Of String, String))()
                        For Each line As String In System.IO.File.ReadAllLines(fileName, System.Text.Encoding.UTF8)
                            Dim trimmedLine As String = line.Trim()
                            If trimmedLine.Length = 0 Then Continue For

                            records.Add(New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
                                {"Column1", trimmedLine}
                            })
                        Next
                        Return records
                    End If

                    Return LoadAdminDelimitedFile(fileName, delimiter)

                Case ".json"
                    Return LoadAdminJsonFile(fileName)

                Case Else
                    Throw New System.InvalidOperationException("Unsupported import file format.")
            End Select
        End Function

        Private Shared Function LoadAdminDelimitedFile(fileName As String, delimiter As Char) As List(Of Dictionary(Of String, String))
            Dim records As New List(Of Dictionary(Of String, String))()

            Dim firstRowIsHeader As Boolean =
                ShowCustomYesNoBox(
                    "Should the first row of the imported file be treated as header names?",
                    "Yes",
                    "No",
                    $"{AN} - Sync with Imported List") = 1

            Using parser As New Microsoft.VisualBasic.FileIO.TextFieldParser(fileName, System.Text.Encoding.UTF8)
                parser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited
                parser.SetDelimiters(delimiter.ToString())
                parser.HasFieldsEnclosedInQuotes = True

                Dim headers As String() = Nothing
                Dim rowIndex As Integer = 0

                While Not parser.EndOfData
                    Dim fields As String() = parser.ReadFields()
                    If fields Is Nothing Then Continue While

                    If rowIndex = 0 Then
                        If firstRowIsHeader Then
                            headers = fields.Select(Function(f, i) If(String.IsNullOrWhiteSpace(f), $"Column{i + 1}", f.Trim())).ToArray()
                            rowIndex += 1
                            Continue While
                        Else
                            headers = Enumerable.Range(1, fields.Length).Select(Function(i) $"Column{i}").ToArray()
                        End If
                    End If

                    Dim record As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                    For i As Integer = 0 To headers.Length - 1
                        Dim fieldValue As String = If(i < fields.Length, fields(i), "")
                        record(headers(i)) = If(fieldValue, "").Trim()
                    Next
                    records.Add(record)
                    rowIndex += 1
                End While
            End Using

            Return records
        End Function

        Private Shared Function LoadAdminJsonFile(fileName As String) As List(Of Dictionary(Of String, String))
            Dim records As New List(Of Dictionary(Of String, String))()
            Dim root As JToken = JToken.Parse(System.IO.File.ReadAllText(fileName, System.Text.Encoding.UTF8))

            If root.Type = JTokenType.Object Then
                Dim obj As JObject = CType(root, JObject)
                Dim firstArrayProperty = obj.Properties().FirstOrDefault(Function(p) p.Value.Type = JTokenType.Array)
                If firstArrayProperty IsNot Nothing Then
                    root = firstArrayProperty.Value
                Else
                    root = New JArray(obj)
                End If
            End If

            If root.Type <> JTokenType.Array Then
                Return records
            End If

            Dim array As JArray = CType(root, JArray)
            Dim firstRowIsHeader As Boolean = False

            If array.Count > 0 AndAlso array(0).Type = JTokenType.Array Then
                firstRowIsHeader =
                    ShowCustomYesNoBox(
                        "Should the first JSON array row be treated as header names?",
                        "Yes",
                        "No",
                        $"{AN} - Sync with Imported List") = 1
            End If

            Dim headers As String() = Nothing
            Dim index As Integer = 0

            For Each item As JToken In array
                If item.Type = JTokenType.Object Then
                    Dim obj As JObject = CType(item, JObject)
                    Dim record As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                    For Each [property] As JProperty In obj.Properties()
                        record([property].Name) = [property].Value.ToString().Trim()
                    Next
                    records.Add(record)
                ElseIf item.Type = JTokenType.Array Then
                    Dim values As JArray = CType(item, JArray)

                    If headers Is Nothing Then
                        If firstRowIsHeader Then
                            headers = values.Select(Function(v, i) If(String.IsNullOrWhiteSpace(v.ToString()), $"Column{i + 1}", v.ToString().Trim())).ToArray()
                            firstRowIsHeader = False
                            index += 1
                            Continue For
                        Else
                            headers = Enumerable.Range(1, values.Count).Select(Function(i) $"Column{i}").ToArray()
                        End If
                    End If

                    Dim record As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                    For i As Integer = 0 To headers.Length - 1
                        record(headers(i)) = If(i < values.Count, values(i).ToString().Trim(), "")
                    Next
                    records.Add(record)
                Else
                    Dim record As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
                        {"Column1", item.ToString().Trim()}
                    }
                    records.Add(record)
                End If

                index += 1
            Next

            Return records
        End Function

        Private Shared Function DetectAdminDelimiter(line As String) As Char
            If String.IsNullOrWhiteSpace(line) Then
                Return ControlChars.NullChar
            End If

            If line.IndexOf(ControlChars.Tab) >= 0 Then
                Return ControlChars.Tab
            End If

            Dim semicolonCount As Integer = line.Count(Function(c) c = ";"c)
            Dim commaCount As Integer = line.Count(Function(c) c = ","c)

            If semicolonCount > commaCount AndAlso semicolonCount > 0 Then
                Return ";"c
            End If

            If commaCount > 0 Then
                Return ","c
            End If

            Return ControlChars.NullChar
        End Function

        Private Shared Function ResolveAdminImportedProductId(record As Dictionary(Of String, String),
                                                              defaultProductId As String,
                                                              allowedProductIds As HashSet(Of String)) As String
            Dim importedProductId As String = GetAdminRecordFieldValue(record, "Product ID")

            If Not String.IsNullOrWhiteSpace(importedProductId) AndAlso allowedProductIds.Contains(importedProductId.Trim()) Then
                Return importedProductId.Trim()
            End If

            Return defaultProductId
        End Function

        Private Shared Function LoadAdminUserIdMappingConfig() As AdminUserIdMappingConfig
            Dim config As New AdminUserIdMappingConfig()

            Try
                Dim json As String = If(My.Settings.License_AdminUserIdMappingJson, "").Trim()
                If json.Length = 0 Then
                    Return config
                End If

                Dim obj As JObject = JObject.Parse(json)
                config.PrimaryField = obj.Value(Of String)("PrimaryField")
                config.UseSecondaryField = obj.Value(Of Boolean?)("UseSecondaryField").GetValueOrDefault(False)
                config.SecondaryField = obj.Value(Of String)("SecondaryField")
                config.PrimaryFieldPart = If(obj.Value(Of String)("PrimaryFieldPart"), "Whole")
                config.SecondaryFieldPart = If(obj.Value(Of String)("SecondaryFieldPart"), "Whole")
                config.Separator = If(obj.Value(Of String)("Separator"), ".")
                config.UseEmailLocalPart = obj.Value(Of Boolean?)("UseEmailLocalPart").GetValueOrDefault(False)
                config.Lowercase = obj.Value(Of Boolean?)("Lowercase").GetValueOrDefault(True)
                config.Uppercase = obj.Value(Of Boolean?)("Uppercase").GetValueOrDefault(False)
                config.TrimWhitespace = obj.Value(Of Boolean?)("TrimWhitespace").GetValueOrDefault(True)
                config.RemoveSpaces = obj.Value(Of Boolean?)("RemoveSpaces").GetValueOrDefault(False)
                config.ReplaceSpacesWith = If(obj.Value(Of String)("ReplaceSpacesWith"), "")
                config.RemoveDiacritics = obj.Value(Of Boolean?)("RemoveDiacritics").GetValueOrDefault(False)
            Catch
            End Try

            If String.IsNullOrWhiteSpace(config.PrimaryField) Then config.PrimaryField = "User ID"
            If String.IsNullOrWhiteSpace(config.PrimaryFieldPart) Then config.PrimaryFieldPart = "Whole"
            If String.IsNullOrWhiteSpace(config.SecondaryFieldPart) Then config.SecondaryFieldPart = "Whole"
            If config.Separator Is Nothing Then config.Separator = "."

            Return config
        End Function

        Private Shared Sub SaveAdminUserIdMappingConfig(config As AdminUserIdMappingConfig)
            Dim obj As New JObject() From {
                {"PrimaryField", config.PrimaryField},
                {"UseSecondaryField", config.UseSecondaryField},
                {"SecondaryField", config.SecondaryField},
                {"PrimaryFieldPart", config.PrimaryFieldPart},
                {"SecondaryFieldPart", config.SecondaryFieldPart},
                {"Separator", config.Separator},
                {"UseEmailLocalPart", config.UseEmailLocalPart},
                {"Lowercase", config.Lowercase},
                {"Uppercase", config.Uppercase},
                {"TrimWhitespace", config.TrimWhitespace},
                {"RemoveSpaces", config.RemoveSpaces},
                {"ReplaceSpacesWith", config.ReplaceSpacesWith},
                {"RemoveDiacritics", config.RemoveDiacritics}
            }

            My.Settings.License_AdminUserIdMappingJson = obj.ToString(Newtonsoft.Json.Formatting.None)
            My.Settings.Save()
        End Sub

        Private Shared Function ShowAdminUserIdMappingDialog(ByRef config As AdminUserIdMappingConfig) As Boolean
            Dim workingCopy As New AdminUserIdMappingConfig() With {
                .PrimaryField = config.PrimaryField,
                .UseSecondaryField = config.UseSecondaryField,
                .SecondaryField = config.SecondaryField,
                .PrimaryFieldPart = config.PrimaryFieldPart,
                .SecondaryFieldPart = config.SecondaryFieldPart,
                .Separator = config.Separator,
                .UseEmailLocalPart = config.UseEmailLocalPart,
                .Lowercase = config.Lowercase,
                .Uppercase = config.Uppercase,
                .TrimWhitespace = config.TrimWhitespace,
                .RemoveSpaces = config.RemoveSpaces,
                .ReplaceSpacesWith = config.ReplaceSpacesWith,
                .RemoveDiacritics = config.RemoveDiacritics
            }

            Using form As New System.Windows.Forms.Form()
                form.AutoScaleDimensions = New System.Drawing.SizeF(96.0F, 96.0F)
                form.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
                form.Text = $"{AN} - User-ID Mapping"
                form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
                form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
                form.MinimizeBox = False
                form.MaximizeBox = False
                form.ShowInTaskbar = False
                form.TopMost = True
                form.MinimumSize = New System.Drawing.Size(760, 700)
                form.Size = New System.Drawing.Size(840, 760)
                form.AutoScroll = True
                form.Font = New System.Drawing.Font("Segoe UI", 9.5F)

                Try
                    Dim bmp As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                    form.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())
                Catch
                End Try

                Dim layout As New System.Windows.Forms.TableLayoutPanel() With {
                    .Dock = System.Windows.Forms.DockStyle.Fill,
                    .Padding = New System.Windows.Forms.Padding(18),
                    .ColumnCount = 2,
                    .RowCount = 13
                }
                layout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize))
                layout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
                form.Controls.Add(layout)

                Dim rowIndex As Integer = 0

                Dim lblIntro As New System.Windows.Forms.Label() With {
                    .Text = "Configure how imported records generate desired User IDs / instances. This mapping is used only for imported desired lists, not for parsed backend activations.",
                    .AutoSize = True,
                    .MaximumSize = New System.Drawing.Size(640, 0),
                    .Margin = New System.Windows.Forms.Padding(0, 0, 0, 12)
                }
                layout.Controls.Add(lblIntro, 0, rowIndex)
                layout.SetColumnSpan(lblIntro, 2)
                rowIndex += 1

                Dim txtPrimaryField As New System.Windows.Forms.TextBox() With {.Text = workingCopy.PrimaryField, .Dock = System.Windows.Forms.DockStyle.Fill}
                AddAdminMappingRow(layout, rowIndex, "Primary source field:", txtPrimaryField)
                rowIndex += 1

                Dim cboPrimaryPart As New System.Windows.Forms.ComboBox() With {.Dock = System.Windows.Forms.DockStyle.Fill, .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList}
                cboPrimaryPart.Items.AddRange(New Object() {"Whole", "First token", "Last token"})
                cboPrimaryPart.SelectedItem = NormalizeAdminMappingPart(workingCopy.PrimaryFieldPart)
                AddAdminMappingRow(layout, rowIndex, "Primary field part:", cboPrimaryPart)
                rowIndex += 1

                Dim chkUseSecondary As New System.Windows.Forms.CheckBox() With {.Text = "Combine a second field", .Checked = workingCopy.UseSecondaryField, .AutoSize = True}
                layout.Controls.Add(chkUseSecondary, 0, rowIndex)
                layout.SetColumnSpan(chkUseSecondary, 2)
                rowIndex += 1

                Dim txtSecondaryField As New System.Windows.Forms.TextBox() With {.Text = workingCopy.SecondaryField, .Dock = System.Windows.Forms.DockStyle.Fill}
                AddAdminMappingRow(layout, rowIndex, "Secondary source field:", txtSecondaryField)
                rowIndex += 1

                Dim cboSecondaryPart As New System.Windows.Forms.ComboBox() With {.Dock = System.Windows.Forms.DockStyle.Fill, .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList}
                cboSecondaryPart.Items.AddRange(New Object() {"Whole", "First token", "Last token"})
                cboSecondaryPart.SelectedItem = NormalizeAdminMappingPart(workingCopy.SecondaryFieldPart)
                AddAdminMappingRow(layout, rowIndex, "Secondary field part:", cboSecondaryPart)
                rowIndex += 1

                Dim txtSeparator As New System.Windows.Forms.TextBox() With {.Text = workingCopy.Separator, .Dock = System.Windows.Forms.DockStyle.Left, .Width = 80}
                AddAdminMappingRow(layout, rowIndex, "Combine separator:", txtSeparator)
                rowIndex += 1

                Dim chkEmailLocalPart As New System.Windows.Forms.CheckBox() With {.Text = "If a source value contains an email address, use only the local part before '@'.", .Checked = workingCopy.UseEmailLocalPart, .AutoSize = True}
                layout.Controls.Add(chkEmailLocalPart, 0, rowIndex)
                layout.SetColumnSpan(chkEmailLocalPart, 2)
                rowIndex += 1

                Dim chkTrim As New System.Windows.Forms.CheckBox() With {.Text = "Trim leading/trailing whitespace", .Checked = workingCopy.TrimWhitespace, .AutoSize = True}
                layout.Controls.Add(chkTrim, 0, rowIndex)
                layout.SetColumnSpan(chkTrim, 2)
                rowIndex += 1

                Dim chkLower As New System.Windows.Forms.CheckBox() With {.Text = "Lowercase", .Checked = workingCopy.Lowercase, .AutoSize = True}
                layout.Controls.Add(chkLower, 0, rowIndex)
                layout.SetColumnSpan(chkLower, 2)
                rowIndex += 1

                Dim chkUpper As New System.Windows.Forms.CheckBox() With {.Text = "Uppercase", .Checked = workingCopy.Uppercase, .AutoSize = True}
                layout.Controls.Add(chkUpper, 0, rowIndex)
                layout.SetColumnSpan(chkUpper, 2)
                rowIndex += 1

                Dim chkRemoveSpaces As New System.Windows.Forms.CheckBox() With {.Text = "Remove spaces entirely", .Checked = workingCopy.RemoveSpaces, .AutoSize = True}
                layout.Controls.Add(chkRemoveSpaces, 0, rowIndex)
                layout.SetColumnSpan(chkRemoveSpaces, 2)
                rowIndex += 1

                Dim txtReplaceSpaces As New System.Windows.Forms.TextBox() With {.Text = workingCopy.ReplaceSpacesWith, .Dock = System.Windows.Forms.DockStyle.Left, .Width = 80}
                AddAdminMappingRow(layout, rowIndex, "Replace spaces with:", txtReplaceSpaces)
                rowIndex += 1

                Dim chkRemoveDiacritics As New System.Windows.Forms.CheckBox() With {.Text = "Remove accents / diacritics", .Checked = workingCopy.RemoveDiacritics, .AutoSize = True}
                layout.Controls.Add(chkRemoveDiacritics, 0, rowIndex)
                layout.SetColumnSpan(chkRemoveDiacritics, 2)
                rowIndex += 1

                Dim buttonPanel As New System.Windows.Forms.FlowLayoutPanel() With {
                    .Dock = System.Windows.Forms.DockStyle.Fill,
                    .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
                    .AutoSize = True,
                    .WrapContents = False
                }
                Dim btnOk As New System.Windows.Forms.Button() With {.Text = "Save", .AutoSize = True}
                Dim btnCancel As New System.Windows.Forms.Button() With {.Text = "Cancel", .AutoSize = True}
                buttonPanel.Controls.Add(btnOk)
                buttonPanel.Controls.Add(btnCancel)

                layout.RowCount += 1
                layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
                layout.Controls.Add(buttonPanel, 0, rowIndex)
                layout.SetColumnSpan(buttonPanel, 2)

                txtSecondaryField.Enabled = chkUseSecondary.Checked
                cboSecondaryPart.Enabled = chkUseSecondary.Checked

                AddHandler chkUseSecondary.CheckedChanged,
                    Sub()
                        txtSecondaryField.Enabled = chkUseSecondary.Checked
                        cboSecondaryPart.Enabled = chkUseSecondary.Checked
                    End Sub

                AddHandler chkLower.CheckedChanged,
                    Sub()
                        If chkLower.Checked Then
                            chkUpper.Checked = False
                        End If
                    End Sub

                AddHandler chkUpper.CheckedChanged,
                    Sub()
                        If chkUpper.Checked Then
                            chkLower.Checked = False
                        End If
                    End Sub

                AddHandler btnOk.Click,
                    Sub()
                        workingCopy.PrimaryField = txtPrimaryField.Text.Trim()
                        workingCopy.UseSecondaryField = chkUseSecondary.Checked
                        workingCopy.SecondaryField = txtSecondaryField.Text.Trim()
                        workingCopy.PrimaryFieldPart = CStr(cboPrimaryPart.SelectedItem)
                        workingCopy.SecondaryFieldPart = CStr(cboSecondaryPart.SelectedItem)
                        workingCopy.Separator = txtSeparator.Text
                        workingCopy.UseEmailLocalPart = chkEmailLocalPart.Checked
                        workingCopy.TrimWhitespace = chkTrim.Checked
                        workingCopy.Lowercase = chkLower.Checked
                        workingCopy.Uppercase = chkUpper.Checked
                        workingCopy.RemoveSpaces = chkRemoveSpaces.Checked
                        workingCopy.ReplaceSpacesWith = txtReplaceSpaces.Text
                        workingCopy.RemoveDiacritics = chkRemoveDiacritics.Checked

                        If workingCopy.ReplaceSpacesWith.Length > 0 Then
                            workingCopy.RemoveSpaces = False
                        End If

                        If String.IsNullOrWhiteSpace(workingCopy.PrimaryField) Then
                            ShowCustomMessageBox("Please define at least a primary source field.", $"{AN} - User-ID Mapping")
                            Return
                        End If

                        form.DialogResult = System.Windows.Forms.DialogResult.OK
                        form.Close()
                    End Sub

                AddHandler btnCancel.Click,
                    Sub()
                        form.DialogResult = System.Windows.Forms.DialogResult.Cancel
                        form.Close()
                    End Sub

                Dim ownerWnd As System.Windows.Forms.IWin32Window = ResolveSameThreadDialogOwner()
                Dim ownerScope As System.IDisposable = Nothing

                Try
                    ownerScope = PushDialogOwner(form)

                    If ownerWnd IsNot Nothing Then
                        form.ShowDialog(ownerWnd)
                    Else
                        form.ShowDialog()
                    End If
                Finally
                    If ownerScope IsNot Nothing Then
                        Try
                            ownerScope.Dispose()
                        Catch
                        End Try
                    End If
                End Try

                If form.DialogResult <> System.Windows.Forms.DialogResult.OK Then
                    Return False
                End If

                config.PrimaryField = workingCopy.PrimaryField
                config.UseSecondaryField = workingCopy.UseSecondaryField
                config.SecondaryField = workingCopy.SecondaryField
                config.PrimaryFieldPart = workingCopy.PrimaryFieldPart
                config.SecondaryFieldPart = workingCopy.SecondaryFieldPart
                config.Separator = workingCopy.Separator
                config.UseEmailLocalPart = workingCopy.UseEmailLocalPart
                config.Lowercase = workingCopy.Lowercase
                config.Uppercase = workingCopy.Uppercase
                config.TrimWhitespace = workingCopy.TrimWhitespace
                config.RemoveSpaces = workingCopy.RemoveSpaces
                config.ReplaceSpacesWith = workingCopy.ReplaceSpacesWith
                config.RemoveDiacritics = workingCopy.RemoveDiacritics

                Return True
            End Using
        End Function

        Private Shared Sub AddAdminMappingRow(layout As System.Windows.Forms.TableLayoutPanel,
                                              rowIndex As Integer,
                                              labelText As String,
                                              control As System.Windows.Forms.Control)
            If layout.RowStyles.Count <= rowIndex Then
                layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            End If

            Dim label As New System.Windows.Forms.Label() With {
                .Text = labelText,
                .AutoSize = True,
                .Anchor = System.Windows.Forms.AnchorStyles.Left,
                .Margin = New System.Windows.Forms.Padding(0, 3, 10, 3)
            }

            layout.Controls.Add(label, 0, rowIndex)
            layout.Controls.Add(control, 1, rowIndex)
        End Sub

        Private Shared Function NormalizeAdminMappingPart(value As String) As String
            If value.Equals("First token", StringComparison.OrdinalIgnoreCase) Then Return "First token"
            If value.Equals("Last token", StringComparison.OrdinalIgnoreCase) Then Return "Last token"
            Return "Whole"
        End Function

        Private Shared Function GenerateAdminUserIdFromRecord(record As Dictionary(Of String, String),
                                                              config As AdminUserIdMappingConfig) As String
            Dim primaryValue As String = GetAdminRecordFieldValue(record, config.PrimaryField)

            If config.UseEmailLocalPart Then
                primaryValue = AdminExtractEmailLocalPart(primaryValue)
            End If

            primaryValue = SelectAdminFieldPart(primaryValue, config.PrimaryFieldPart)

            Dim combinedValue As String = primaryValue

            If config.UseSecondaryField Then
                Dim secondaryValue As String = GetAdminRecordFieldValue(record, config.SecondaryField)

                If config.UseEmailLocalPart Then
                    secondaryValue = AdminExtractEmailLocalPart(secondaryValue)
                End If

                secondaryValue = SelectAdminFieldPart(secondaryValue, config.SecondaryFieldPart)

                If Not String.IsNullOrWhiteSpace(secondaryValue) Then
                    If String.IsNullOrWhiteSpace(combinedValue) Then
                        combinedValue = secondaryValue
                    Else
                        combinedValue &= If(config.Separator, "") & secondaryValue
                    End If
                End If
            End If

            Return ApplyAdminUserIdTransforms(combinedValue, config)
        End Function

        Private Shared Function GetAdminRecordFieldValue(record As Dictionary(Of String, String), fieldName As String) As String
            If record Is Nothing OrElse record.Count = 0 Then Return ""

            Dim normalizedFieldName As String = If(fieldName, "").Trim()
            If normalizedFieldName.Length = 0 Then
                If record.Count = 1 Then
                    Return record.Values.First().Trim()
                End If
                Return ResolveAdminFallbackFieldValue(record)
            End If

            If record.ContainsKey(normalizedFieldName) Then
                Return If(record(normalizedFieldName), "").Trim()
            End If

            Dim compactFieldName As String = normalizedFieldName.Replace(" ", "")
            For Each kvp As KeyValuePair(Of String, String) In record
                If kvp.Key.Replace(" ", "").Equals(compactFieldName, StringComparison.OrdinalIgnoreCase) Then
                    Return If(kvp.Value, "").Trim()
                End If
            Next

            Select Case normalizedFieldName.ToLowerInvariant()
                Case "user id", "userid", "instance"
                    Return GetAdminRecordValueByAliases(record, {"User ID", "UserID", "Instance", "Username", "Login", "User", "Column1"})
                Case "first name", "firstname"
                    Return GetAdminRecordValueByAliases(record, {"First Name", "Firstname", "First", "Given Name"})
                Case "last name", "lastname", "surname"
                    Return GetAdminRecordValueByAliases(record, {"Last Name", "Lastname", "Surname", "Family Name"})
                Case "full name", "name"
                    Return GetAdminRecordValueByAliases(record, {"Full Name", "Name", "Display Name"})
                Case "email", "e-mail"
                    Return GetAdminRecordValueByAliases(record, {"Email", "E-mail", "Mail"})
                Case "email local part"
                    Return AdminExtractEmailLocalPart(GetAdminRecordValueByAliases(record, {"Email", "E-mail", "Mail"}))
                Case "organization", "org", "company"
                    Return GetAdminRecordValueByAliases(record, {"Organization", "Org", "Company"})
                Case Else
                    Return ""
            End Select
        End Function

        Private Shared Function ResolveAdminFallbackFieldValue(record As Dictionary(Of String, String)) As String
            Dim fallbackAliases As String() = {"User ID", "UserID", "Instance", "Email", "Full Name", "Name", "Column1"}

            For Each aliasName As String In fallbackAliases
                Dim value As String = GetAdminRecordFieldValue(record, aliasName)
                If Not String.IsNullOrWhiteSpace(value) Then
                    Return value
                End If
            Next

            Return record.Values.FirstOrDefault(Function(v) Not String.IsNullOrWhiteSpace(v))
        End Function

        Private Shared Function GetAdminRecordValueByAliases(record As Dictionary(Of String, String), aliases As IEnumerable(Of String)) As String
            For Each aliasName As String In aliases
                If record.ContainsKey(aliasName) Then
                    Return If(record(aliasName), "").Trim()
                End If
            Next

            For Each aliasName As String In aliases
                Dim compactAlias As String = aliasName.Replace(" ", "")
                For Each kvp As KeyValuePair(Of String, String) In record
                    If kvp.Key.Replace(" ", "").Equals(compactAlias, StringComparison.OrdinalIgnoreCase) Then
                        Return If(kvp.Value, "").Trim()
                    End If
                Next
            Next

            Return ""
        End Function

        Private Shared Function SelectAdminFieldPart(value As String, partSelection As String) As String
            Dim trimmedValue As String = If(value, "").Trim()
            If trimmedValue.Length = 0 Then Return ""

            Dim tokens As String() = trimmedValue.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
            If tokens.Length = 0 Then Return ""

            Select Case NormalizeAdminMappingPart(partSelection)
                Case "First token"
                    Return tokens(0)
                Case "Last token"
                    Return tokens(tokens.Length - 1)
                Case Else
                    Return trimmedValue
            End Select
        End Function

        Private Shared Function ApplyAdminUserIdTransforms(value As String, config As AdminUserIdMappingConfig) As String
            Dim transformed As String = If(value, "")

            If config.TrimWhitespace Then
                transformed = transformed.Trim()
            End If

            If config.RemoveDiacritics Then
                transformed = RemoveAdminDiacritics(transformed)
            End If

            If Not String.IsNullOrEmpty(config.ReplaceSpacesWith) Then
                transformed = System.Text.RegularExpressions.Regex.Replace(transformed, "\s+", config.ReplaceSpacesWith)
            ElseIf config.RemoveSpaces Then
                transformed = System.Text.RegularExpressions.Regex.Replace(transformed, "\s+", "")
            End If

            If config.Lowercase Then
                transformed = transformed.ToLowerInvariant()
            ElseIf config.Uppercase Then
                transformed = transformed.ToUpperInvariant()
            End If

            Return transformed.Trim()
        End Function

        Private Shared Function RemoveAdminDiacritics(value As String) As String
            Dim normalizedString As String = value.Normalize(NormalizationForm.FormD)
            Dim sb As New System.Text.StringBuilder()

            For Each c As Char In normalizedString
                If CharUnicodeInfo.GetUnicodeCategory(c) <> UnicodeCategory.NonSpacingMark Then
                    sb.Append(c)
                End If
            Next

            Return sb.ToString().Normalize(NormalizationForm.FormC)
        End Function

        Private Shared Function AdminExtractEmailLocalPart(value As String) As String
            Dim trimmedValue As String = If(value, "").Trim()
            Dim atIndex As Integer = trimmedValue.IndexOf("@"c)
            If atIndex > 0 Then
                Return trimmedValue.Substring(0, atIndex)
            End If
            Return trimmedValue
        End Function

#End Region

#Region "Organization Bulk License Management - UI Helpers / Filters / Export"

        Private Shared Function ParseOpaqueUserList(inputText As String) As List(Of String)
            Dim values As New List(Of String)()
            Dim normalizedText As String = NormalizeAdminLineEndings(inputText)

            For Each line As String In normalizedText.Split({vbLf}, StringSplitOptions.None)
                Dim trimmedLine As String = line.Trim()
                If trimmedLine.Length = 0 Then Continue For

                Dim splitValues As String() = trimmedLine.Split(New Char() {ControlChars.Tab, ","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)

                If splitValues.Length > 1 Then
                    For Each splitValue As String In splitValues
                        Dim candidate As String = splitValue.Trim()
                        If candidate.Length > 0 Then values.Add(candidate)
                    Next
                Else
                    values.Add(trimmedLine)
                End If
            Next

            Return values
        End Function

        Private Shared Function SelectDefaultAdminProductId(allowedProductIds As HashSet(Of String),
                                                            productTitles As Dictionary(Of String, String),
                                                            prompt As String) As String
            If allowedProductIds Is Nothing OrElse allowedProductIds.Count = 0 Then
                Return ""
            End If

            If allowedProductIds.Count = 1 Then
                Return allowedProductIds.First()
            End If

            Dim options As New List(Of String)()
            For Each productId As String In allowedProductIds.OrderBy(Function(v) v)
                Dim productTitle As String = GetProductTitleFromLookup(productTitles, productId)
                If String.IsNullOrWhiteSpace(productTitle) Then
                    options.Add(productId)
                Else
                    options.Add($"{productId} - {productTitle}")
                End If
            Next

            Dim selected As String = ShowSelectionForm(prompt, $"{AN} - Select Product", options)
            If selected = "ESC" Then Return ""

            Dim dashIndex As Integer = selected.IndexOf(" - ", StringComparison.Ordinal)
            If dashIndex > 0 Then
                Return selected.Substring(0, dashIndex).Trim()
            End If

            Return selected.Trim()
        End Function

        Private Shared Function GetProductTitleFromLookup(productTitles As Dictionary(Of String, String), productId As String) As String
            If productTitles Is Nothing Then Return ""
            If String.IsNullOrWhiteSpace(productId) Then Return ""

            Dim title As String = Nothing
            If productTitles.TryGetValue(productId, title) Then
                Return title
            End If

            Return ""
        End Function

        Private Shared Sub RefreshAdminFilterChoices(cboFilterProduct As System.Windows.Forms.ComboBox,
                                                     cboFilterDesired As System.Windows.Forms.ComboBox,
                                                     cboFilterAction As System.Windows.Forms.ComboBox,
                                                     cboFilterVerify As System.Windows.Forms.ComboBox,
                                                     cboFilterWarnings As System.Windows.Forms.ComboBox,
                                                     rows As IEnumerable(Of AdminLicenseRow))

            RefreshAdminFilterCombo(cboFilterProduct, "All", rows.Select(Function(r) If(String.IsNullOrWhiteSpace(r.ProductId), "(blank)", r.ProductId)).Distinct().OrderBy(Function(v) v))
            RefreshAdminFilterCombo(cboFilterDesired, "All", rows.Select(Function(r) If(String.IsNullOrWhiteSpace(r.DesiredState), "(blank)", r.DesiredState)).Distinct().OrderBy(Function(v) v))
            RefreshAdminFilterCombo(cboFilterAction, "All", rows.Select(Function(r) If(String.IsNullOrWhiteSpace(r.PlannedAction), "(blank)", r.PlannedAction)).Distinct().OrderBy(Function(v) v))
            RefreshAdminFilterCombo(cboFilterVerify, "All", rows.Select(Function(r) If(String.IsNullOrWhiteSpace(r.VerificationState), "(blank)", r.VerificationState)).Distinct().OrderBy(Function(v) v))
            RefreshAdminFilterCombo(cboFilterWarnings, "All", New String() {"All", "With warnings", "Without warnings"})
        End Sub

        Private Shared Sub RefreshAdminFilterCombo(comboBox As System.Windows.Forms.ComboBox,
                                                   defaultValue As String,
                                                   values As IEnumerable(Of String))
            Dim previousValue As String = If(comboBox.SelectedItem, defaultValue).ToString()

            comboBox.BeginUpdate()
            comboBox.Items.Clear()
            comboBox.Items.Add(defaultValue)

            Dim uniqueValues As New List(Of String)()
            For Each value As String In values
                If value Is Nothing Then Continue For
                If value.Equals(defaultValue, StringComparison.OrdinalIgnoreCase) Then Continue For
                If Not uniqueValues.Any(Function(v) v.Equals(value, StringComparison.OrdinalIgnoreCase)) Then
                    uniqueValues.Add(value)
                End If
            Next

            For Each value As String In uniqueValues
                comboBox.Items.Add(value)
            Next
            comboBox.EndUpdate()

            If comboBox.Items.Contains(previousValue) Then
                comboBox.SelectedItem = previousValue
            Else
                comboBox.SelectedItem = defaultValue
            End If
        End Sub

        Private Shared Sub ApplyAdminGridFilters(grid As System.Windows.Forms.DataGridView,
                                                 cboFilterProduct As System.Windows.Forms.ComboBox,
                                                 cboFilterDesired As System.Windows.Forms.ComboBox,
                                                 cboFilterAction As System.Windows.Forms.ComboBox,
                                                 cboFilterVerify As System.Windows.Forms.ComboBox,
                                                 cboFilterWarnings As System.Windows.Forms.ComboBox)
            Dim visibilityByRow As New Dictionary(Of System.Windows.Forms.DataGridViewRow, Boolean)()

            For Each gridRow As System.Windows.Forms.DataGridViewRow In grid.Rows
                Dim row As AdminLicenseRow = TryCast(gridRow.DataBoundItem, AdminLicenseRow)
                If row Is Nothing Then Continue For

                Dim visible As Boolean = True

                Dim productFilter As String = CStr(If(cboFilterProduct.SelectedItem, "All"))
                If Not productFilter.Equals("All", StringComparison.OrdinalIgnoreCase) Then
                    Dim rowProduct As String = If(String.IsNullOrWhiteSpace(row.ProductId), "(blank)", row.ProductId)
                    visible = visible AndAlso rowProduct.Equals(productFilter, StringComparison.OrdinalIgnoreCase)
                End If

                Dim desiredFilter As String = CStr(If(cboFilterDesired.SelectedItem, "All"))
                If Not desiredFilter.Equals("All", StringComparison.OrdinalIgnoreCase) Then
                    Dim rowDesired As String = If(String.IsNullOrWhiteSpace(row.DesiredState), "(blank)", row.DesiredState)
                    visible = visible AndAlso rowDesired.Equals(desiredFilter, StringComparison.OrdinalIgnoreCase)
                End If

                Dim actionFilter As String = CStr(If(cboFilterAction.SelectedItem, "All"))
                If Not actionFilter.Equals("All", StringComparison.OrdinalIgnoreCase) Then
                    Dim rowAction As String = If(String.IsNullOrWhiteSpace(row.PlannedAction), "(blank)", row.PlannedAction)
                    visible = visible AndAlso rowAction.Equals(actionFilter, StringComparison.OrdinalIgnoreCase)
                End If

                Dim verifyFilter As String = CStr(If(cboFilterVerify.SelectedItem, "All"))
                If Not verifyFilter.Equals("All", StringComparison.OrdinalIgnoreCase) Then
                    Dim rowVerify As String = If(String.IsNullOrWhiteSpace(row.VerificationState), "(blank)", row.VerificationState)
                    visible = visible AndAlso rowVerify.Equals(verifyFilter, StringComparison.OrdinalIgnoreCase)
                End If

                Dim warningFilter As String = CStr(If(cboFilterWarnings.SelectedItem, "All"))
                If warningFilter.Equals("With warnings", StringComparison.OrdinalIgnoreCase) Then
                    visible = visible AndAlso Not String.IsNullOrWhiteSpace(row.WarningText)
                ElseIf warningFilter.Equals("Without warnings", StringComparison.OrdinalIgnoreCase) Then
                    visible = visible AndAlso String.IsNullOrWhiteSpace(row.WarningText)
                End If

                visibilityByRow(gridRow) = visible
            Next

            Dim currencyManager As System.Windows.Forms.CurrencyManager = Nothing
            If grid.BindingContext IsNot Nothing AndAlso grid.DataSource IsNot Nothing Then
                currencyManager = TryCast(grid.BindingContext(grid.DataSource), System.Windows.Forms.CurrencyManager)
            End If

            If currencyManager IsNot Nothing Then
                currencyManager.SuspendBinding()
            End If

            Try
                Dim fallbackRow As System.Windows.Forms.DataGridViewRow = visibilityByRow.
                    Where(Function(kvp) kvp.Value).
                    Select(Function(kvp) kvp.Key).
                    FirstOrDefault()

                If grid.CurrentCell IsNot Nothing Then
                    Dim currentRow As System.Windows.Forms.DataGridViewRow = grid.Rows(grid.CurrentCell.RowIndex)
                    Dim currentRowVisible As Boolean = True

                    If visibilityByRow.TryGetValue(currentRow, currentRowVisible) AndAlso Not currentRowVisible Then
                        grid.ClearSelection()

                        If fallbackRow IsNot Nothing Then
                            Dim targetCell As System.Windows.Forms.DataGridViewCell =
                                fallbackRow.Cells.Cast(Of System.Windows.Forms.DataGridViewCell)().
                                    FirstOrDefault(Function(c) c.Visible)

                            If targetCell IsNot Nothing Then
                                grid.CurrentCell = targetCell
                            Else
                                grid.CurrentCell = Nothing
                            End If
                        Else
                            grid.CurrentCell = Nothing
                        End If
                    End If
                End If

                For Each kvp As KeyValuePair(Of System.Windows.Forms.DataGridViewRow, Boolean) In visibilityByRow
                    kvp.Key.Visible = kvp.Value
                Next
            Finally
                If currencyManager IsNot Nothing Then
                    currencyManager.ResumeBinding()
                End If
            End Try
        End Sub

        Private Shared Sub ExportAdminRowsToCsv(rows As IEnumerable(Of AdminLicenseRow),
                                                resultsOnly As Boolean,
                                                correlationId As String)
            Dim delimiterChoice As String = ShowSelectionForm(
                "Choose the CSV delimiter for export:",
                $"{AN} - Export CSV",
                New String() {"Comma (,)", "Semicolon (;)"})

            If delimiterChoice = "ESC" Then
                Return
            End If

            Dim delimiter As Char = If(delimiterChoice.StartsWith("Semicolon", StringComparison.OrdinalIgnoreCase), ";"c, ","c)

            Using dialog As New System.Windows.Forms.SaveFileDialog()
                dialog.Title = If(resultsOnly, $"{AN} - Export Result Report", $"{AN} - Export Current Table")
                dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
                dialog.FileName = If(resultsOnly, "license-manager-result-report.csv", "license-manager-current-table.csv")

                Dim __safeDialogOwner3166 As System.Windows.Forms.IWin32Window = Global.SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
                If If(__safeDialogOwner3166 IsNot Nothing, dialog.ShowDialog(__safeDialogOwner3166), dialog.ShowDialog()) <> System.Windows.Forms.DialogResult.OK Then
                    Return
                End If

                Dim exportRows = If(resultsOnly,
                    rows.Where(Function(r) Not String.IsNullOrWhiteSpace(r.ApiResult) OrElse
                                           Not r.PlannedAction.Equals("Ignore", StringComparison.OrdinalIgnoreCase) OrElse
                                           Not String.IsNullOrWhiteSpace(r.LastCorrelationId)),
                    rows)

                Dim sb As New System.Text.StringBuilder()
                sb.AppendLine(String.Join(delimiter, New String() {
                    AdminCsvEscape("Product ID", delimiter),
                    AdminCsvEscape("Product Title", delimiter),
                    AdminCsvEscape("User ID / Instance", delimiter),
                    AdminCsvEscape("Parsed Activation Date", delimiter),
                    AdminCsvEscape("Source", delimiter),
                    AdminCsvEscape("Verification State", delimiter),
                    AdminCsvEscape("Desired State", delimiter),
                    AdminCsvEscape("Planned Action", delimiter),
                    AdminCsvEscape("Final Result", delimiter),
                    AdminCsvEscape("Warning / Error", delimiter),
                    AdminCsvEscape("Correlation ID", delimiter)
                }))

                For Each row As AdminLicenseRow In exportRows
                    sb.AppendLine(String.Join(delimiter, New String() {
                        AdminCsvEscape(row.ProductId, delimiter),
                        AdminCsvEscape(row.ProductTitle, delimiter),
                        AdminCsvEscape(row.InstanceUserId, delimiter),
                        AdminCsvEscape(row.ParsedActivationDate, delimiter),
                        AdminCsvEscape(row.SourceType, delimiter),
                        AdminCsvEscape(row.VerificationState, delimiter),
                        AdminCsvEscape(row.DesiredState, delimiter),
                        AdminCsvEscape(row.PlannedAction, delimiter),
                        AdminCsvEscape(row.ApiResult, delimiter),
                        AdminCsvEscape(row.WarningText, delimiter),
                        AdminCsvEscape(If(String.IsNullOrWhiteSpace(row.LastCorrelationId), correlationId, row.LastCorrelationId), delimiter)
                    }))
                Next

                System.IO.File.WriteAllText(dialog.FileName, sb.ToString(), New System.Text.UTF8Encoding(True))
            End Using
        End Sub

        Private Shared Function AdminCsvEscape(value As String, delimiter As Char) As String
            Dim text As String = If(value, "")
            Dim mustQuote As Boolean =
                text.IndexOf(delimiter) >= 0 OrElse
                text.IndexOf(""""c) >= 0 OrElse
                text.IndexOf(vbCr) >= 0 OrElse
                text.IndexOf(vbLf) >= 0

            If mustQuote Then
                Return $"""{text.Replace("""", """""")}"""
            End If

            Return text
        End Function

        Private Shared Function MaskLicenseKey(value As String) As String
            If String.IsNullOrWhiteSpace(value) Then
                Return "(none)"
            End If

            Dim trimmedValue As String = value.Trim()
            If trimmedValue.Length <= 8 Then
                Return New String("*"c, trimmedValue.Length)
            End If

            Return $"{trimmedValue.Substring(0, 4)}...{trimmedValue.Substring(trimmedValue.Length - 4)}"
        End Function

        Private Shared Function BuildAdminKey(productId As String, instanceUserId As String) As String
            Return $"{productId}{ChrW(31)}{instanceUserId}"
        End Function

        Private Shared Sub AppendDistinctAdminWarning(row As AdminLicenseRow, warningText As String)
            If String.IsNullOrWhiteSpace(warningText) Then Return

            If String.IsNullOrWhiteSpace(row.WarningText) Then
                row.WarningText = warningText.Trim()
                Return
            End If

            If row.WarningText.IndexOf(warningText, StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return
            End If

            row.WarningText &= "; " & warningText.Trim()
        End Sub

#End Region

    End Class
End Namespace
