' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: KnowledgeStoreForm.vb
' Purpose:
'   WinForms administration dialog for configuring and maintaining Knowledge
'   Stores.
'
' Responsibilities:
'   - List all configured stores discovered from `KnowledgeStoreCatalog`.
'   - Edit local store metadata such as name, path, owner, role, and flags.
'   - Display per-store statistics including indexed documents, wiki pages,
'     embeddings, and writeability state.
'   - Trigger maintenance actions such as index, re-index, health check,
'     repair, revectorization, and schema editing.
'   - Run long-running maintenance work asynchronously and surface progress via
'     status text and tray notifications.
'
' Notes:
'   - This dialog is shared between Word and Outlook.
'   - Health Check, Lint Wiki, and Repair delegate into `KnowledgeWikiService`.
'   - Foreground indexing delegates into `KnowledgeStoreForegroundIndexer`.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports System.Threading.Tasks

Namespace SharedLibrary

    ''' <summary>
    ''' Modal dialog for full Knowledge Store administration: listing, editing,
    ''' statistics, and maintenance operations.
    ''' </summary>
    Public Class KnowledgeStoreForm
        Inherits Form

        Private ReadOnly _context As ISharedContext
        Private _stores As New List(Of KnowledgeStoreCatalog.KnowledgeStoreDefinition)()
        Private _dialogOwnerScope As System.IDisposable = Nothing

        ' ── Left panel: Store list ──
        Private ReadOnly _lstStores As New ListBox()
        Private ReadOnly _btnAddStore As New Button() With {.Text = "New Store", .Margin = New Padding(4)}
        Private ReadOnly _btnDeleteStore As New Button() With {.Text = "Delete Store", .Margin = New Padding(4)}

        ' ── Right panel: Detail editing ──
        Private ReadOnly _txtName As New TextBox() With {.Width = 320}
        Private ReadOnly _txtSourcePath As New TextBox() With {.Width = 420}
        Private ReadOnly _btnBrowsePath As New Button() With {.Text = "Browse...", .Margin = New Padding(4)}
        Private ReadOnly _txtOwner As New TextBox() With {.Width = 320}
        Private ReadOnly _cboRole As New ComboBox() With {.Width = 180, .DropDownStyle = ComboBoxStyle.DropDownList}
        Private ReadOnly _chkActive As New CheckBox() With {.Text = "Active", .AutoSize = True}
        Private ReadOnly _chkScanSub As New CheckBox() With {.Text = "Scan subdirectories", .AutoSize = True}
        Private ReadOnly _btnSave As New Button() With {.Text = "Save Changes", .Margin = New Padding(4)}
        Private ReadOnly _btnReformatPage As New Button() With {.Text = "Reformat Page...", .Margin = New Padding(4)}

        ' ── Statistics ──
        Private ReadOnly _lblStats As New Label() With {
            .AutoSize = False,
            .Width = 420,
            .Height = 124,
            .BorderStyle = BorderStyle.FixedSingle,
            .Padding = New Padding(6),
            .BackColor = SystemColors.Info,
            .ForeColor = SystemColors.InfoText
        }

        ' ── Operation buttons ──
        Private ReadOnly _btnIndex As New Button() With {.Text = "Index", .Margin = New Padding(4)}
        Private ReadOnly _btnReindex As New Button() With {.Text = "Full Re-Index", .Margin = New Padding(4)}
        Private ReadOnly _btnHealthCheck As New Button() With {.Text = "Health Check", .Margin = New Padding(4)}
        Private ReadOnly _btnRepair As New Button() With {.Text = "Repair", .Margin = New Padding(4)}
        Private ReadOnly _btnRevectorize As New Button() With {.Text = "Rebuild Embeddings", .Margin = New Padding(4)}
        Private ReadOnly _btnLint As New Button() With {.Text = "Lint Wiki", .Margin = New Padding(4)}
        Private ReadOnly _btnRepairPage As New Button() With {.Text = "Repair Page...", .Margin = New Padding(4)}
        Private ReadOnly _btnEditSchema As New Button() With {.Text = "Edit Schema", .Margin = New Padding(4)}
        Private ReadOnly _btnClose As New Button() With {.Text = "Close", .Margin = New Padding(4)}

        ' ── Status bar ──
        Private ReadOnly _lblStatus As New Label() With {.AutoSize = True, .Dock = DockStyle.Bottom, .Padding = New Padding(8)}

        ''' <summary>
        ''' Initializes the Knowledge Store administration form.
        ''' </summary>
        Public Sub New(context As ISharedContext)
            _context = context
            InitializeForm()
            LoadStores()
        End Sub

        Private _runningOperations As Integer = 0

        Protected Overrides Sub OnHandleCreated(e As System.EventArgs)
            MyBase.OnHandleCreated(e)

            If _dialogOwnerScope Is Nothing Then
                _dialogOwnerScope = PushDialogOwner(Me)
            End If
        End Sub

        Protected Overrides Sub OnHandleDestroyed(e As System.EventArgs)
            Dim scope As System.IDisposable = _dialogOwnerScope
            _dialogOwnerScope = Nothing

            If scope IsNot Nothing Then
                Try
                    scope.Dispose()
                Catch
                End Try
            End If

            MyBase.OnHandleDestroyed(e)
        End Sub

#Region "Form Initialization"

        Private Sub InitializeForm()
            Me.SuspendLayout()

            Me.Text = $"{AN} — Knowledge Store Administration"
            Me.AutoScaleMode = AutoScaleMode.Dpi
            Me.AutoScaleDimensions = New SizeF(96.0F, 96.0F)
            Me.Size = New Size(1120, 780)
            Me.MinimumSize = New Size(980, 700)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.FormBorderStyle = FormBorderStyle.Sizable
            Me.KeyPreview = True
            Me.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)

            ' Set icon
            Try
                Dim bmp As New Bitmap(GetLogoBitmap(LogoType.Standard))
                Me.Icon = Icon.FromHandle(bmp.GetHicon())
                bmp.Dispose()
            Catch
            End Try

            ConfigureStandardButton(_btnAddStore)
            ConfigureStandardButton(_btnDeleteStore)
            ConfigureStandardButton(_btnSave)
            ConfigureStandardButton(_btnIndex)
            ConfigureStandardButton(_btnReindex)
            ConfigureStandardButton(_btnHealthCheck)
            ConfigureStandardButton(_btnRepair)
            ConfigureStandardButton(_btnRepairPage)
            ConfigureStandardButton(_btnReformatPage)
            ConfigureStandardButton(_btnRevectorize)
            ConfigureStandardButton(_btnLint)
            ConfigureStandardButton(_btnEditSchema)
            ConfigureStandardButton(_btnClose)


            ConfigureBrowseButton(_btnBrowsePath, _txtSourcePath)

            ' ── Role combo items ──
            _cboRole.Items.AddRange(New Object() {"personal", "shared", "readonly"})
            _cboRole.SelectedIndex = 0

            ' ── Left panel: store list ──
            Dim leftPanel As New Panel() With {
                .Dock = DockStyle.Left,
                .Width = 260,
                .Padding = New Padding(10)
            }

            Dim lblStores As New Label() With {
                .Text = "Knowledge Stores",
                .Font = New Font(Me.Font, FontStyle.Bold),
                .AutoSize = True,
                .Dock = DockStyle.Top,
                .Padding = New Padding(0, 0, 0, 4)
            }

            _lstStores.Dock = DockStyle.Fill
            _lstStores.IntegralHeight = False

            Dim leftBtnPanel As New FlowLayoutPanel() With {
                .Dock = DockStyle.Bottom,
                .FlowDirection = FlowDirection.LeftToRight,
                .AutoSize = True,
                .Padding = New Padding(0, 4, 0, 0)
            }
            leftBtnPanel.Controls.AddRange(New Control() {_btnAddStore, _btnDeleteStore})

            leftPanel.Controls.Add(_lstStores)
            leftPanel.Controls.Add(leftBtnPanel)
            leftPanel.Controls.Add(lblStores)

            ' ── Right panel: details + operations ──
            Dim rightPanel As New Panel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(10),
                .AutoScroll = True
            }

            Dim detailLayout As New TableLayoutPanel() With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .ColumnCount = 2,
                .Padding = New Padding(0)
            }
            detailLayout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            detailLayout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))

            Dim row As Integer = 0
            AddLabeledRow(detailLayout, row, "Name:", _txtName) : row += 1

            ' Source path with browse button
            Dim pathPanel As New FlowLayoutPanel() With {
                .FlowDirection = FlowDirection.LeftToRight,
                .AutoSize = True,
                .WrapContents = False,
                .Margin = New Padding(0),
                .AutoSizeMode = AutoSizeMode.GrowAndShrink
            }
            pathPanel.Controls.Add(_txtSourcePath)
            pathPanel.Controls.Add(_btnBrowsePath)
            AddLabeledRow(detailLayout, row, "Source Path:", pathPanel) : row += 1

            AddLabeledRow(detailLayout, row, "Owner:", _txtOwner) : row += 1
            AddLabeledRow(detailLayout, row, "Role:", _cboRole) : row += 1

            Dim checkPanel As New FlowLayoutPanel() With {
                .FlowDirection = FlowDirection.LeftToRight,
                .AutoSize = True,
                .WrapContents = False,
                .Margin = New Padding(0)
            }
            checkPanel.Controls.Add(_chkActive)
            checkPanel.Controls.Add(_chkScanSub)
            AddLabeledRow(detailLayout, row, "", checkPanel) : row += 1

            ' Save button
            Dim savePanel As New FlowLayoutPanel() With {
                .FlowDirection = FlowDirection.LeftToRight,
                .AutoSize = True,
                .Dock = DockStyle.Top,
                .Padding = New Padding(0, 6, 0, 6)
            }
            savePanel.Controls.Add(_btnSave)

            ' Statistics
            Dim lblStatsHeader As New Label() With {
                .Text = "Statistics",
                .Font = New Font(Me.Font, FontStyle.Bold),
                .AutoSize = True,
                .Dock = DockStyle.Top,
                .Padding = New Padding(0, 8, 0, 4)
            }

            Dim statsPanel As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 132,
                .Padding = New Padding(0, 0, 0, 4)
            }
            _lblStats.Dock = DockStyle.Fill
            statsPanel.Controls.Add(_lblStats)

            ' Operations
            Dim lblOpsHeader As New Label() With {
                .Text = "Operations",
                .Font = New Font(Me.Font, FontStyle.Bold),
                .AutoSize = True,
                .Dock = DockStyle.Top,
                .Padding = New Padding(0, 6, 0, 4)
            }

            Dim opsPanel As New FlowLayoutPanel() With {
                .Dock = DockStyle.Top,
                .FlowDirection = FlowDirection.LeftToRight,
                .AutoSize = True,
                .WrapContents = True,
                .Padding = New Padding(0, 0, 0, 4)
            }
            opsPanel.Controls.AddRange(New Control() {
                        _btnIndex, _btnReindex, _btnHealthCheck, _btnRepair, _btnRepairPage, _btnReformatPage,
                        _btnRevectorize, _btnLint, _btnEditSchema
                    })

            ' Bottom close button
            Dim bottomPanel As New FlowLayoutPanel() With {
                .Dock = DockStyle.Bottom,
                .FlowDirection = FlowDirection.RightToLeft,
                .AutoSize = True,
                .Padding = New Padding(8, 4, 8, 8)
            }
            bottomPanel.Controls.Add(_btnClose)

            ' Add right panel children in reverse dock order (bottom-up for Top docking)
            rightPanel.Controls.Add(opsPanel)
            rightPanel.Controls.Add(lblOpsHeader)
            rightPanel.Controls.Add(statsPanel)
            rightPanel.Controls.Add(lblStatsHeader)
            rightPanel.Controls.Add(savePanel)
            rightPanel.Controls.Add(detailLayout)

            ' Splitter
            Dim splitter As New Splitter() With {.Dock = DockStyle.Left, .Width = 4}

            ' Add to form (order matters for docking)
            Me.Controls.Add(rightPanel)
            Me.Controls.Add(splitter)
            Me.Controls.Add(leftPanel)
            Me.Controls.Add(bottomPanel)
            Me.Controls.Add(_lblStatus)

            ' Wire events
            AddHandler _lstStores.SelectedIndexChanged, AddressOf OnStoreSelected
            AddHandler _btnAddStore.Click, AddressOf OnAddStore
            AddHandler _btnDeleteStore.Click, AddressOf OnDeleteStore
            AddHandler _btnBrowsePath.Click, AddressOf OnBrowsePath
            AddHandler _btnSave.Click, AddressOf OnSaveChanges
            AddHandler _btnIndex.Click, AddressOf OnIndex
            AddHandler _btnReindex.Click, AddressOf OnReindex
            AddHandler _btnHealthCheck.Click, AddressOf OnHealthCheck
            AddHandler _btnRepair.Click, AddressOf OnRepair
            AddHandler _btnRepairPage.Click, AddressOf OnRepairPage
            AddHandler _btnReformatPage.Click, AddressOf OnReformatPage
            AddHandler _btnRevectorize.Click, AddressOf OnRevectorize
            AddHandler _btnLint.Click, AddressOf OnLint
            AddHandler _btnEditSchema.Click, AddressOf OnEditSchema
            AddHandler _btnClose.Click, Sub(s, ev) Me.Close()
            AddHandler Me.KeyDown, AddressOf OnKeyDown
            Me.ResumeLayout(performLayout:=True)
        End Sub

        Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
            Dim cancellationPending As Boolean = False

            Try
                cancellationPending = ProgressBarModule.CancelOperation
            Catch
                cancellationPending = False
            End Try

            If _runningOperations > 0 AndAlso Not cancellationPending Then
                Dim answer = ShowCustomYesNoBox(
            "A maintenance operation is still running. Closing now will let it finish in the background, " &
            "but progress will no longer be shown. Close anyway?",
            "Yes, close", "No, keep open")
                If answer <> 1 Then
                    e.Cancel = True
                    Return
                End If
            End If

            MyBase.OnFormClosing(e)
        End Sub

        Private Shared Sub AddLabeledRow(table As TableLayoutPanel, row As Integer, labelText As String, control As Control)
            table.RowCount = row + 1
            table.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            If Not String.IsNullOrEmpty(labelText) Then
                Dim lbl As New Label() With {
                    .Text = labelText,
                    .AutoSize = True,
                    .Anchor = AnchorStyles.Left,
                    .Margin = New Padding(0, 6, 8, 2)
                }
                table.Controls.Add(lbl, 0, row)
            End If
            control.Margin = New Padding(0, 4, 0, 2)
            table.Controls.Add(control, 1, row)
        End Sub

        Private Shared Sub ConfigureStandardButton(button As Button)
            button.AutoSize = True
            button.AutoSizeMode = AutoSizeMode.GrowAndShrink
            button.UseVisualStyleBackColor = True
            button.Padding = New Padding(10, 4, 10, 4)
            button.MinimumSize = New Size(0, 0)
        End Sub

        Private Shared Sub ConfigureBrowseButton(button As Button, relatedTextBox As TextBox)
            button.AutoSize = False
            button.UseVisualStyleBackColor = True
            button.Padding = New Padding(10, 0, 10, 0)
            button.Height = relatedTextBox.PreferredHeight + 2
            button.Width = Math.Max(90, TextRenderer.MeasureText(button.Text, button.Font).Width + 24)
        End Sub

#End Region

#Region "Store Loading"

        Private Sub LoadStores()
            _stores = KnowledgeStoreCatalog.LoadAll(_context)
            _lstStores.Items.Clear()

            For Each store In _stores
                _lstStores.Items.Add(KnowledgeStoreCatalog.GetDisplayLabel(store))
            Next

            UpdateStatus()

            If _lstStores.Items.Count > 0 Then
                _lstStores.SelectedIndex = 0
            Else
                ClearDetails()
            End If
        End Sub

        Private Sub UpdateStatus()
            Dim total = _stores.Count
            Dim active = _stores.Where(Function(s) s.Active).Count()
            _lblStatus.Text = $"{total} store(s) configured, {active} active"
        End Sub

#End Region

#Region "Store Selection & Details"

        Private Sub OnStoreSelected(sender As Object, e As EventArgs)
            Dim idx = _lstStores.SelectedIndex
            If idx < 0 OrElse idx >= _stores.Count Then
                ClearDetails()
                Return
            End If

            Dim store = _stores(idx)
            _txtName.Text = store.Name
            _txtSourcePath.Text = store.SourcePath
            _txtOwner.Text = store.Owner
            _chkActive.Checked = store.Active
            _chkScanSub.Checked = store.ScanSubdirectories

            Dim roleIdx = _cboRole.Items.IndexOf(If(store.Role, "personal").ToLowerInvariant())
            _cboRole.SelectedIndex = If(roleIdx >= 0, roleIdx, 0)

            ' Enable/disable editing for central stores
            Dim isLocal = Not store.IsFromCentralCatalog
            _txtName.ReadOnly = Not isLocal
            _txtSourcePath.ReadOnly = Not isLocal
            _txtOwner.ReadOnly = Not isLocal
            _cboRole.Enabled = isLocal
            _chkActive.Enabled = isLocal
            _chkScanSub.Enabled = isLocal
            _btnBrowsePath.Enabled = isLocal
            _btnSave.Enabled = isLocal
            _btnDeleteStore.Enabled = isLocal

            LoadStatistics(store)
        End Sub

        Private Sub ClearDetails()
            _txtName.Text = ""
            _txtSourcePath.Text = ""
            _txtOwner.Text = ""
            _cboRole.SelectedIndex = 0
            _chkActive.Checked = False
            _chkScanSub.Checked = False
            _lblStats.Text = "No store selected."
            _btnSave.Enabled = False
            _btnDeleteStore.Enabled = False
        End Sub

        Private Sub LoadStatistics(store As KnowledgeStoreCatalog.KnowledgeStoreDefinition)
            Try
                Dim manifest = KnowledgeStoreManifest.Load(store)
                Dim docCount = manifest.Entries.Count
                Dim lastIndexed As String = "Never"
                If docCount > 0 Then
                    Dim maxDate = manifest.Entries.Max(Function(e) e.IndexedDate)
                    If maxDate <> Date.MinValue Then
                        lastIndexed = maxDate.ToString("yyyy-MM-dd HH:mm", System.Globalization.CultureInfo.InvariantCulture)
                    End If
                End If

                Dim wikiPageCount As Integer = 0
                Dim embeddingCount As Integer = 0
                Dim sourceLabel = If(store.IsFromCentralCatalog, "Central", "Local")
                Dim resolvedPath = If(store.ResolvedSourcePath, ExpandEnvironmentVariables(store.SourcePath))
                Dim pathExists = Not String.IsNullOrWhiteSpace(resolvedPath) AndAlso Directory.Exists(resolvedPath)

                If pathExists Then
                    Dim wikiRoot = Path.Combine(resolvedPath, ".redink", KnowledgeStoreCatalog.WikiFolder)
                    If Directory.Exists(wikiRoot) Then
                        wikiPageCount = Directory.GetFiles(wikiRoot, "*.md", SearchOption.AllDirectories).
                            Where(Function(f)
                                      Dim n = Path.GetFileName(f)
                                      Return Not n.Equals("index.md", StringComparison.OrdinalIgnoreCase) AndAlso
                                             Not n.Equals("log.md", StringComparison.OrdinalIgnoreCase) AndAlso
                                             Not n.Equals("health_report.md", StringComparison.OrdinalIgnoreCase) AndAlso
                                             Not n.Equals("review_queue.md", StringComparison.OrdinalIgnoreCase)
                                  End Function).Count()
                    End If

                    Dim embPath = Path.Combine(resolvedPath, ".redink", ".embeddings.json")
                    If File.Exists(embPath) Then
                        Try
                            Dim embJson = File.ReadAllText(embPath, System.Text.Encoding.UTF8)
                            Dim arr = Newtonsoft.Json.Linq.JArray.Parse(embJson)
                            embeddingCount = arr.Count
                        Catch
                        End Try
                    End If
                End If

                Dim canWrite = KnowledgeStoreCatalog.CanCurrentUserWrite(store, _context)

                _lblStats.Text =
                    $"Source: {sourceLabel}  |  Path exists: {If(pathExists, "Yes", "No")}" & vbCrLf &
                    $"Documents indexed: {docCount}" & vbCrLf &
                    $"Wiki pages: {wikiPageCount}" & vbCrLf &
                    $"Embedding vectors: {embeddingCount}" & vbCrLf &
                    $"Last indexed: {lastIndexed}" & vbCrLf &
                    $"Writable: {If(canWrite, "Yes", "No")}"
            Catch ex As Exception
                _lblStats.Text = $"Error loading statistics: {ex.Message}"
            End Try
        End Sub

#End Region

#Region "Add / Delete Store"

        Private Sub OnAddStore(sender As Object, e As EventArgs)
            If String.IsNullOrWhiteSpace(_context.INI_KnowledgeStorePathLocal) Then
                ShowCustomMessageBox(
                    "No local Knowledge Store catalog path is configured. " &
                    "Please set 'KnowledgeStorePathLocal' in your configuration file.", AN)
                Return
            End If

            Dim storeName = ShowCustomInputBox(
                "Enter a name for the new Knowledge Store (e.g., 'Contracts', 'Research'):",
                $"{AN} — New Knowledge Store", True, "")
            If String.IsNullOrWhiteSpace(storeName) Then Return

            Dim existing = KnowledgeStoreCatalog.GetStoreByName(storeName.Trim(), _context)
            If existing IsNot Nothing Then
                ShowCustomMessageBox($"A Knowledge Store named '{storeName.Trim()}' already exists.", AN)
                Return
            End If

            Using fbd As New FolderBrowserDialog()
                fbd.Description = "Select the root directory for this Knowledge Store"
                fbd.ShowNewFolderButton = True
                If fbd.ShowDialog() <> DialogResult.OK Then Return

                Try
                    Dim resolvedPath = fbd.SelectedPath
                    If Not Directory.Exists(resolvedPath) Then
                        Directory.CreateDirectory(resolvedPath)
                    End If

                    KnowledgeWikiService.InitializeWikiStructure(resolvedPath)

                    Dim def = KnowledgeStoreCatalog.CreateDefinition(
                        storeName.Trim(), fbd.SelectedPath, _context)

                    Dim allDefs = KnowledgeStoreCatalog.LoadAll(_context)
                    allDefs.Add(def)
                    KnowledgeStoreCatalog.SaveLocalCatalog(allDefs, _context)

                    ShowCustomMessageBox(
                        $"Knowledge Store '{def.Name}' created at:{vbCrLf}{def.ResolvedSourcePath}",
                        $"{AN} Knowledge Store")

                    LoadStores()

                    ' Select the newly added store
                    For i = 0 To _stores.Count - 1
                        If _stores(i).Name.Equals(storeName.Trim(), StringComparison.OrdinalIgnoreCase) Then
                            _lstStores.SelectedIndex = i
                            Exit For
                        End If
                    Next

                Catch ex As Exception
                    ShowCustomMessageBox($"Error creating Knowledge Store: {ex.Message}", AN)
                End Try
            End Using
        End Sub

        Private Sub OnDeleteStore(sender As Object, e As EventArgs)
            Dim idx = _lstStores.SelectedIndex
            If idx < 0 OrElse idx >= _stores.Count Then Return

            Dim store = _stores(idx)
            If store.IsFromCentralCatalog Then
                ShowCustomMessageBox("Central stores cannot be deleted from here. Contact your administrator.", AN)
                Return
            End If

            Dim answer = ShowCustomYesNoBox(
                $"Delete the Knowledge Store definition '{store.Name}'?{vbCrLf}" &
                $"(Source files at '{store.ResolvedSourcePath}' will NOT be deleted.)",
                "Yes, delete", "Cancel")
            If answer <> 1 Then Return

            Dim allDefs = KnowledgeStoreCatalog.LoadAll(_context)
            allDefs.RemoveAll(Function(d)
                                  Return Not d.IsFromCentralCatalog AndAlso
                                         String.Equals(d.StoreId, store.StoreId, StringComparison.OrdinalIgnoreCase)
                              End Function)
            KnowledgeStoreCatalog.SaveLocalCatalog(allDefs, _context)

            LoadStores()
        End Sub

#End Region

#Region "Edit & Save"

        Private Sub OnBrowsePath(sender As Object, e As EventArgs)
            Using fbd As New FolderBrowserDialog()
                fbd.Description = "Select the root directory for this Knowledge Store"
                fbd.ShowNewFolderButton = True
                If Not String.IsNullOrWhiteSpace(_txtSourcePath.Text) Then
                    Dim expanded = ExpandEnvironmentVariables(_txtSourcePath.Text)
                    If Directory.Exists(expanded) Then
                        fbd.SelectedPath = expanded
                    End If
                End If
                If fbd.ShowDialog() = DialogResult.OK Then
                    _txtSourcePath.Text = fbd.SelectedPath
                End If
            End Using
        End Sub

        Private Sub OnSaveChanges(sender As Object, e As EventArgs)
            Dim idx = _lstStores.SelectedIndex
            If idx < 0 OrElse idx >= _stores.Count Then Return

            Dim store = _stores(idx)
            If store.IsFromCentralCatalog Then
                ShowCustomMessageBox("Central store definitions cannot be modified here.", AN)
                Return
            End If

            If String.IsNullOrWhiteSpace(_txtName.Text) Then
                ShowCustomMessageBox("Store name cannot be empty.", AN)
                Return
            End If

            If String.IsNullOrWhiteSpace(_txtSourcePath.Text) Then
                ShowCustomMessageBox("Source path cannot be empty.", AN)
                Return
            End If

            ' Update the in-memory definition
            store.Name = _txtName.Text.Trim()
            store.SourcePath = KnowledgeStoreCatalog.StripQuotes(_txtSourcePath.Text)
            store.Owner = _txtOwner.Text.Trim()
            store.Role = If(_cboRole.SelectedItem IsNot Nothing, _cboRole.SelectedItem.ToString(), "personal")
            store.Active = _chkActive.Checked
            store.ScanSubdirectories = _chkScanSub.Checked
            store.ResolvedSourcePath = ExpandEnvironmentVariables(store.SourcePath)

            ' Save back to local catalog
            Try
                Dim allDefs = KnowledgeStoreCatalog.LoadAll(_context)

                ' Find and replace the matching local definition
                Dim found = False
                For i = 0 To allDefs.Count - 1
                    If Not allDefs(i).IsFromCentralCatalog AndAlso
                       String.Equals(allDefs(i).StoreId, store.StoreId, StringComparison.OrdinalIgnoreCase) Then
                        allDefs(i) = store
                        found = True
                        Exit For
                    End If
                Next

                If Not found Then
                    ' Match by old name if StoreId changed
                    For i = 0 To allDefs.Count - 1
                        If Not allDefs(i).IsFromCentralCatalog AndAlso
                           String.Equals(allDefs(i).Name, _stores(idx).Name, StringComparison.OrdinalIgnoreCase) Then
                            allDefs(i) = store
                            found = True
                            Exit For
                        End If
                    Next
                End If

                KnowledgeStoreCatalog.SaveLocalCatalog(allDefs, _context)

                ShowCustomMessageBox($"Store '{store.Name}' saved successfully.", AN)

                ' Reload to reflect changes
                Dim selectedName = store.Name
                LoadStores()
                For i = 0 To _stores.Count - 1
                    If _stores(i).Name.Equals(selectedName, StringComparison.OrdinalIgnoreCase) Then
                        _lstStores.SelectedIndex = i
                        Exit For
                    End If
                Next
            Catch ex As Exception
                ShowCustomMessageBox($"Error saving store: {ex.Message}", AN)
            End Try
        End Sub

#End Region

#Region "Operations"

        Private Async Function RunMaintenanceJobAsync(store As KnowledgeStoreCatalog.KnowledgeStoreDefinition,
                                                      operationName As String,
                                                      work As Func(Of Action(Of String), Task(Of String))) As Task(Of String)
            Dim trayIcon As NotifyIcon = Nothing
            Dim lastBalloonKey As String = ""

            Try
                Try
                    trayIcon = New NotifyIcon() With {
                        .Icon = SystemIcons.Information,
                        .Visible = True
                    }
                Catch
                    trayIcon = Nothing
                End Try

                UpdateMaintenanceStatusSafe(store.Name, operationName, "Queued as a background job.", trayIcon)
                ShowMaintenanceBalloonSafe(
                    trayIcon,
                    $"{AN} Knowledge Store",
                    $"{operationName} for '{store.Name}' started.")

                Dim progressCallback As Action(Of String) =
                    Sub(statusText)
                        UpdateMaintenanceStatusSafe(store.Name, operationName, statusText, trayIcon)

                        Dim nextBalloonKey As String = ""
                        Dim nextBalloonText As String = ""

                        If statusText.IndexOf("waiting for", StringComparison.OrdinalIgnoreCase) >= 0 Then
                            nextBalloonKey = "wait"
                            nextBalloonText = $"{operationName} for '{store.Name}' is waiting for {KnowledgeStoreHostGate.HostDisplayName} to become idle."
                        ElseIf statusText.IndexOf("starting ai step", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                               statusText.IndexOf("continuing", StringComparison.OrdinalIgnoreCase) >= 0 Then
                            nextBalloonKey = "run"
                            nextBalloonText = $"{operationName} for '{store.Name}' resumed."
                        End If

                        If nextBalloonKey <> "" AndAlso
                           Not String.Equals(lastBalloonKey, nextBalloonKey, StringComparison.Ordinal) Then
                            lastBalloonKey = nextBalloonKey
                            ShowMaintenanceBalloonSafe(trayIcon, $"{AN} Knowledge Store", nextBalloonText)
                        End If
                    End Sub

                Dim result = Await Task.Run(
                    Async Function()
                        Return Await work(progressCallback).ConfigureAwait(False)
                    End Function)

                UpdateMaintenanceStatusSafe(store.Name, operationName, "Completed.", trayIcon)
                ShowMaintenanceBalloonSafe(
                    trayIcon,
                    $"{AN} Knowledge Store",
                    $"{operationName} for '{store.Name}' completed.")

                Return result
            Catch ex As Exception
                UpdateMaintenanceStatusSafe(store.Name, operationName, "Failed.", trayIcon)
                ShowMaintenanceBalloonSafe(
                    trayIcon,
                    $"{AN} Knowledge Store",
                    $"{operationName} for '{store.Name}' failed: {ex.Message}")
                Throw
            Finally
                If trayIcon IsNot Nothing Then
                    Dim iconToDispose = trayIcon
                    trayIcon = Nothing

                    Try
                        If Me.IsDisposed OrElse Me.Disposing OrElse Not Me.IsHandleCreated Then
                            Try
                                iconToDispose.Visible = False
                                iconToDispose.Dispose()
                            Catch
                            End Try
                        ElseIf Me.InvokeRequired Then
                            Try
                                Me.BeginInvoke(
                        New Action(
                            Sub()
                                Try
                                    iconToDispose.Visible = False
                                    iconToDispose.Dispose()
                                Catch
                                End Try
                            End Sub))
                            Catch
                                Try
                                    iconToDispose.Visible = False
                                    iconToDispose.Dispose()
                                Catch
                                End Try
                            End Try
                        Else
                            Try
                                iconToDispose.Visible = False
                                iconToDispose.Dispose()
                            Catch
                            End Try
                        End If
                    Catch
                    End Try
                End If
            End Try
        End Function

        Private Sub UpdateMaintenanceStatusSafe(storeName As String,
                                        operationName As String,
                                        statusText As String,
                                        trayIcon As NotifyIcon)
            If Me.IsDisposed OrElse Me.Disposing OrElse Not Me.IsHandleCreated Then
                Return
            End If

            If Me.InvokeRequired Then
                Try
                    Me.Invoke(
                New Action(Of String, String, String, NotifyIcon)(AddressOf UpdateMaintenanceStatusSafe),
                storeName,
                operationName,
                statusText,
                trayIcon)
                Catch
                End Try
                Return
            End If

            Dim cleaned = statusText
            If cleaned IsNot Nothing AndAlso
       cleaned.StartsWith(operationName, StringComparison.OrdinalIgnoreCase) Then
                Dim rest = cleaned.Substring(operationName.Length)
                If rest.Length = 0 OrElse "(: —-".IndexOf(rest(0)) >= 0 OrElse Char.IsWhiteSpace(rest(0)) Then
                    cleaned = rest.TrimStart(":"c, " "c, ChrW(9))
                End If
            End If

            _lblStatus.Text = $"{operationName} — {storeName}: {cleaned}"
            _lblStatus.Refresh()

            If trayIcon IsNot Nothing Then
                Dim toolTipText = $"{operationName}: {cleaned}"
                If String.IsNullOrWhiteSpace(toolTipText) Then
                    toolTipText = $"{operationName}: working..."
                End If
                If toolTipText.Length > 63 Then
                    toolTipText = toolTipText.Substring(0, 60).TrimEnd() & "..."
                End If
                Try
                    trayIcon.Text = toolTipText
                Catch
                End Try
            End If
        End Sub


        Private Sub ShowMaintenanceBalloonSafe(trayIcon As NotifyIcon, title As String, text As String)
            If trayIcon Is Nothing OrElse String.IsNullOrWhiteSpace(text) OrElse Me.IsDisposed Then
                Return
            End If

            If Me.InvokeRequired Then
                Try
                    Me.BeginInvoke(
                        New Action(Of NotifyIcon, String, String)(AddressOf ShowMaintenanceBalloonSafe),
                        trayIcon,
                        title,
                        text)
                Catch
                End Try
                Return
            End If

            Try
                trayIcon.BalloonTipIcon = ToolTipIcon.Info
                trayIcon.BalloonTipTitle = title
                trayIcon.BalloonTipText = text
                trayIcon.ShowBalloonTip(3000)
            Catch
            End Try
        End Sub

        Private Function GetSelectedStore() As KnowledgeStoreCatalog.KnowledgeStoreDefinition
            Dim idx = _lstStores.SelectedIndex
            If idx < 0 OrElse idx >= _stores.Count Then
                ShowCustomMessageBox("Please select a Knowledge Store first.", AN)
                Return Nothing
            End If
            Return _stores(idx)
        End Function

        Private Async Sub OnIndex(sender As Object, e As EventArgs)
            Dim store = GetSelectedStore()
            If store Is Nothing Then Return

            If Not KnowledgeStoreCatalog.CanCurrentUserWrite(store, _context) Then
                ShowCustomMessageBox($"You do not have write permission to '{store.Name}'.", AN)
                Return
            End If

            Dim buttonsRestoredEarly As Boolean = False

            SetOperationButtonsEnabled(False)
            System.Threading.Interlocked.Increment(_runningOperations)
            Try
                Dim result = Await KnowledgeStoreForegroundIndexer.RunAsync(
            _context, storeName:=store.StoreId, forceReindex:=False)

                If Me.IsDisposed OrElse Me.Disposing Then Return

                System.Threading.Interlocked.Decrement(_runningOperations)
                SetOperationButtonsEnabled(True)
                UpdateStatus()
                buttonsRestoredEarly = True

                ShowCustomMessageBox(
            $"Indexing complete for '{store.Name}':{vbCrLf}" &
            $"Total: {result.TotalFiles}, Indexed: {result.IndexedFiles}, " &
            $"Skipped: {result.SkippedFiles}, Failed: {result.FailedFiles}" &
            If(result.WasCancelled, $"{vbCrLf}(Cancelled by user)", ""),
            $"{AN} Knowledge Store")

                LoadStatistics(store)
                UpdateStatus()
            Catch ex As Exception
                If Not Me.IsDisposed Then ShowCustomMessageBox($"Error during indexing: {ex.Message}", AN)
            Finally
                If Not buttonsRestoredEarly Then
                    System.Threading.Interlocked.Decrement(_runningOperations)
                    If Not Me.IsDisposed Then
                        SetOperationButtonsEnabled(True)
                        UpdateStatus()
                    End If
                End If
            End Try
        End Sub

        Private Async Sub OnReindex(sender As Object, e As EventArgs)
            Dim store = GetSelectedStore()
            If store Is Nothing Then Return

            If Not KnowledgeStoreCatalog.CanCurrentUserWrite(store, _context) Then
                ShowCustomMessageBox($"You do not have write permission to '{store.Name}'.", AN)
                Return
            End If

            Dim answer = ShowCustomYesNoBox(
        $"This will force a full re-index of '{store.Name}', regenerating all metadata and Wiki summaries. " &
        "This may take a while and use API credits. Continue?",
        "Yes, re-index", "No, cancel")
            If answer <> 1 Then Return

            Dim buttonsRestoredEarly As Boolean = False

            SetOperationButtonsEnabled(False)
            System.Threading.Interlocked.Increment(_runningOperations)
            Try
                Dim result = Await KnowledgeStoreForegroundIndexer.RunAsync(
            _context, storeName:=store.StoreId, forceReindex:=True)

                If Me.IsDisposed OrElse Me.Disposing Then Return

                System.Threading.Interlocked.Decrement(_runningOperations)
                SetOperationButtonsEnabled(True)
                UpdateStatus()
                buttonsRestoredEarly = True

                ShowCustomMessageBox(
            $"Re-indexing complete for '{store.Name}':{vbCrLf}" &
            $"Total: {result.TotalFiles}, Indexed: {result.IndexedFiles}, " &
            $"Skipped: {result.SkippedFiles}, Failed: {result.FailedFiles}" &
            If(result.WasCancelled, $"{vbCrLf}(Cancelled by user)", ""),
            $"{AN} Knowledge Store")

                LoadStatistics(store)
                UpdateStatus()
            Catch ex As Exception
                If Not Me.IsDisposed Then ShowCustomMessageBox($"Error during re-indexing: {ex.Message}", AN)
            Finally
                If Not buttonsRestoredEarly Then
                    System.Threading.Interlocked.Decrement(_runningOperations)
                    If Not Me.IsDisposed Then
                        SetOperationButtonsEnabled(True)
                        UpdateStatus()
                    End If
                End If
            End Try
        End Sub

        Private Async Sub OnHealthCheck(sender As Object, e As EventArgs)
            Dim store = GetSelectedStore()
            If store Is Nothing Then Return

            If String.IsNullOrWhiteSpace(store.ResolvedSourcePath) Then
                ShowCustomMessageBox("The selected store does not have a valid source path.", AN)
                Return
            End If

            Dim buttonsRestoredEarly As Boolean = False

            SetOperationButtonsEnabled(False)
            System.Threading.Interlocked.Increment(_runningOperations)
            Try
                Dim report = Await RunMaintenanceJobAsync(
            store:=store,
            operationName:="Health Check",
            work:=Function(progressCallback)
                      Return KnowledgeWikiService.LintWikiAsync(
                          kbRootPath:=store.ResolvedSourcePath,
                          context:=_context,
                          autoApply:=False,
                          progressCallback:=progressCallback)
                  End Function)

                If Me.IsDisposed OrElse Me.Disposing Then Return

                System.Threading.Interlocked.Decrement(_runningOperations)
                SetOperationButtonsEnabled(True)
                UpdateStatus()
                buttonsRestoredEarly = True

                ShowCustomWindow(
            $"Health check report for '{store.Name}':",
            report,
            "You can copy this report to the clipboard.",
            $"{AN} Knowledge Store")

                LoadStatistics(store)
                UpdateStatus()
            Catch ex As Exception
                If Not Me.IsDisposed Then ShowCustomMessageBox($"Error during health check: {ex.Message}", AN)
            Finally
                If Not buttonsRestoredEarly Then
                    System.Threading.Interlocked.Decrement(_runningOperations)
                    If Not Me.IsDisposed Then
                        SetOperationButtonsEnabled(True)
                        UpdateStatus()
                    End If
                End If
            End Try
        End Sub

        Private Async Sub OnReformatPage(sender As Object, e As EventArgs)
            Dim store = GetSelectedStore()
            If store Is Nothing Then Return

            If String.IsNullOrWhiteSpace(store.ResolvedSourcePath) Then
                ShowCustomMessageBox("The selected store does not have a valid source path.", AN)
                Return
            End If

            If Not KnowledgeStoreCatalog.CanCurrentUserWrite(store, _context) Then
                ShowCustomMessageBox($"You do not have write permission to '{store.Name}'.", AN)
                Return
            End If

            Dim wikiRoot = Path.Combine(store.ResolvedSourcePath, ".redink", KnowledgeStoreCatalog.WikiFolder)
            If Not Directory.Exists(wikiRoot) Then
                ShowCustomMessageBox("This store has no wiki yet. Index it first.", AN)
                Return
            End If

            Dim reservedNames = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
        KnowledgeStoreCatalog.IndexFile, KnowledgeStoreCatalog.LogFile,
        "health_report.md", "review_queue.md"
    }

            Dim allPages = Directory.GetFiles(wikiRoot, "*.md", SearchOption.AllDirectories).
        Where(Function(f) Not reservedNames.Contains(Path.GetFileName(f))).
        OrderBy(Function(f) f, StringComparer.OrdinalIgnoreCase).
        ToList()

            If allPages.Count = 0 Then
                ShowCustomMessageBox("No wiki pages found in this store.", AN)
                Return
            End If

            Dim selected = ShowWikiPageSelector(wikiRoot, allPages)
            If selected Is Nothing OrElse selected.Count = 0 Then Return

            Dim confirm = ShowCustomYesNoBox(
        $"Reformat {selected.Count} wiki page(s) in '{store.Name}'?{vbCrLf}{vbCrLf}" &
        "This deterministically normalizes front matter and removes malformed metadata blocks. " &
        "No LLM call is made.",
        "Yes, reformat", "Cancel")
            If confirm <> 1 Then Return

            Dim buttonsRestoredEarly As Boolean = False

            SetOperationButtonsEnabled(False)
            System.Threading.Interlocked.Increment(_runningOperations)

            Try
                Dim summary = Await RunMaintenanceJobAsync(
            store:=store,
            operationName:="Reformat Page",
            work:=Function(progressCallback)
                      Return KnowledgeWikiService.ReformatMalformedPagesAsync(
                          kbRootPath:=store.ResolvedSourcePath,
                          context:=_context,
                          pagePaths:=selected,
                          progressCallback:=progressCallback,
                          operationName:="Reformat Page")
                  End Function)

                If Me.IsDisposed OrElse Me.Disposing Then Return

                System.Threading.Interlocked.Decrement(_runningOperations)
                SetOperationButtonsEnabled(True)
                UpdateStatus()
                buttonsRestoredEarly = True

                ShowCustomWindow(
            $"Page reformat summary for '{store.Name}':",
            summary,
            "You can copy this summary to the clipboard.",
            $"{AN} Knowledge Store")

                LoadStatistics(store)
                UpdateStatus()
            Catch ex As Exception
                If Not Me.IsDisposed Then ShowCustomMessageBox($"Error during page reformat: {ex.Message}", AN)
            Finally
                If Not buttonsRestoredEarly Then
                    System.Threading.Interlocked.Decrement(_runningOperations)
                    If Not Me.IsDisposed Then
                        SetOperationButtonsEnabled(True)
                        UpdateStatus()
                    End If
                End If
            End Try
        End Sub

        Private Async Sub OnRepairPage(sender As Object, e As EventArgs)
            Dim store = GetSelectedStore()
            If store Is Nothing Then Return

            If String.IsNullOrWhiteSpace(store.ResolvedSourcePath) Then
                ShowCustomMessageBox("The selected store does not have a valid source path.", AN)
                Return
            End If

            If Not KnowledgeStoreCatalog.CanCurrentUserWrite(store, _context) Then
                ShowCustomMessageBox($"You do not have write permission to '{store.Name}'.", AN)
                Return
            End If

            Dim wikiRoot = Path.Combine(store.ResolvedSourcePath, ".redink", KnowledgeStoreCatalog.WikiFolder)
            If Not Directory.Exists(wikiRoot) Then
                ShowCustomMessageBox("This store has no wiki yet. Index it first.", AN)
                Return
            End If

            Dim reservedNames = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
        KnowledgeStoreCatalog.IndexFile, KnowledgeStoreCatalog.LogFile,
        "health_report.md", "review_queue.md"
    }

            Dim allPages = Directory.GetFiles(wikiRoot, "*.md", SearchOption.AllDirectories).
        Where(Function(f) Not reservedNames.Contains(Path.GetFileName(f))).
        OrderBy(Function(f) f, StringComparer.OrdinalIgnoreCase).
        ToList()

            If allPages.Count = 0 Then
                ShowCustomMessageBox("No wiki pages found in this store.", AN)
                Return
            End If

            Dim selected = ShowWikiPageSelector(wikiRoot, allPages)
            If selected Is Nothing OrElse selected.Count = 0 Then Return

            Dim confirm = ShowCustomYesNoBox(
        $"Repair {selected.Count} wiki page(s) in '{store.Name}'?{vbCrLf}{vbCrLf}" &
        "For each page the corresponding source document (if known) will be re-ingested by the LLM, " &
        "overwriting the page. Pages without a known source will simply be deleted.",
        "Yes, repair", "Cancel")
            If confirm <> 1 Then Return

            Dim buttonsRestoredEarly As Boolean = False

            SetOperationButtonsEnabled(False)
            System.Threading.Interlocked.Increment(_runningOperations)
            Try
                Dim summary = Await RunMaintenanceJobAsync(
            store:=store,
            operationName:="Repair Page",
            work:=Function(progressCallback) RepairWikiPagesAsync(store, selected, progressCallback))

                If Me.IsDisposed OrElse Me.Disposing Then Return

                System.Threading.Interlocked.Decrement(_runningOperations)
                SetOperationButtonsEnabled(True)
                UpdateStatus()
                buttonsRestoredEarly = True

                ShowCustomWindow(
            $"Page repair summary for '{store.Name}':",
            summary,
            "You can copy this summary to the clipboard.",
            $"{AN} Knowledge Store")

                LoadStatistics(store)
                UpdateStatus()
            Catch ex As Exception
                If Not Me.IsDisposed Then ShowCustomMessageBox($"Error during page repair: {ex.Message}", AN)
            Finally
                If Not buttonsRestoredEarly Then
                    System.Threading.Interlocked.Decrement(_runningOperations)
                    If Not Me.IsDisposed Then
                        SetOperationButtonsEnabled(True)
                        UpdateStatus()
                    End If
                End If
            End Try
        End Sub

        ''' <summary>
        ''' Modal multi-select dialog listing wiki page relative paths.
        ''' Returns the selected absolute file paths, or Nothing on cancel.
        ''' </summary>
        Private Function ShowWikiPageSelector(wikiRoot As String, pages As List(Of String)) As List(Of String)
            Using dlg As New Form()
                dlg.Text = $"{AN} — Select Wiki Pages To Repair"
                dlg.StartPosition = FormStartPosition.CenterParent
                dlg.Size = New Size(720, 560)
                dlg.MinimumSize = New Size(560, 400)
                dlg.FormBorderStyle = FormBorderStyle.Sizable
                dlg.MinimizeBox = False
                dlg.MaximizeBox = True
                dlg.ShowInTaskbar = False
                dlg.Font = Me.Font
                dlg.AutoScaleMode = AutoScaleMode.Dpi
                dlg.AutoScaleDimensions = New SizeF(96.0F, 96.0F)
                dlg.BackColor = SystemColors.Control

                ' Match main form icon
                Try
                    Dim bmp As New Bitmap(GetLogoBitmap(LogoType.Standard))
                    dlg.Icon = Icon.FromHandle(bmp.GetHicon())
                    bmp.Dispose()
                Catch
                End Try

                ' Header label — auto-sized so it cannot be cut off
                Dim header As New Label() With {
                    .Text = "Select one or more wiki pages to repair." & vbCrLf &
                            "Each selected page will be deleted and, if its source can be located, regenerated by the LLM.",
                    .Dock = DockStyle.Top,
                    .AutoSize = True,
                    .MaximumSize = New Size(dlg.ClientSize.Width - 24, 0),
                    .Padding = New Padding(12, 12, 12, 8),
                    .Font = New Font(Me.Font, FontStyle.Regular),
                    .BackColor = SystemColors.Control
                }
                AddHandler dlg.SizeChanged, Sub(s, ev) header.MaximumSize = New Size(dlg.ClientSize.Width - 24, 0)

                ' Padded list container
                Dim listHost As New Panel() With {
                    .Dock = DockStyle.Fill,
                    .Padding = New Padding(12, 4, 12, 8),
                    .BackColor = SystemColors.Control
                }
                Dim list As New CheckedListBox() With {
                    .Dock = DockStyle.Fill,
                    .CheckOnClick = True,
                    .IntegralHeight = False,
                    .BorderStyle = BorderStyle.FixedSingle
                }
                For Each absPath In pages
                    list.Items.Add(GetRelativePathSafe(wikiRoot, absPath))
                Next
                listHost.Controls.Add(list)

                ' Select-all helpers
                Dim btnAll As New Button() With {.Text = "Select All", .Margin = New Padding(4)}
                Dim btnNone As New Button() With {.Text = "Clear", .Margin = New Padding(4)}
                ConfigureStandardButton(btnAll)
                ConfigureStandardButton(btnNone)
                AddHandler btnAll.Click, Sub(s, ev)
                                             For i = 0 To list.Items.Count - 1
                                                 list.SetItemChecked(i, True)
                                             Next
                                         End Sub
                AddHandler btnNone.Click, Sub(s, ev)
                                              For i = 0 To list.Items.Count - 1
                                                  list.SetItemChecked(i, False)
                                              Next
                                          End Sub

                Dim btnOk As New Button() With {.Text = "Repair Selected", .DialogResult = DialogResult.OK, .Margin = New Padding(4)}
                Dim btnCancel As New Button() With {.Text = "Cancel", .DialogResult = DialogResult.Cancel, .Margin = New Padding(4)}
                ConfigureStandardButton(btnOk)
                ConfigureStandardButton(btnCancel)

                Dim btnPanel As New TableLayoutPanel() With {
                    .Dock = DockStyle.Bottom,
                    .AutoSize = True,
                    .ColumnCount = 4,
                    .Padding = New Padding(12, 8, 12, 12),
                    .BackColor = SystemColors.Control
                }
                btnPanel.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
                btnPanel.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
                btnPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
                btnPanel.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
                btnPanel.Controls.Add(btnAll, 0, 0)
                btnPanel.Controls.Add(btnNone, 1, 0)
                Dim spacer As New Panel() With {.Dock = DockStyle.Fill, .Height = 1}
                btnPanel.Controls.Add(spacer, 2, 0)
                Dim rightFlow As New FlowLayoutPanel() With {
                    .FlowDirection = FlowDirection.RightToLeft,
                    .AutoSize = True,
                    .WrapContents = False
                }
                rightFlow.Controls.Add(btnCancel)
                rightFlow.Controls.Add(btnOk)
                btnPanel.Controls.Add(rightFlow, 3, 0)

                ' Order matters for docking (Fill last among non-bottom/top)
                dlg.Controls.Add(listHost)
                dlg.Controls.Add(btnPanel)
                dlg.Controls.Add(header)
                dlg.AcceptButton = btnOk
                dlg.CancelButton = btnCancel

                If dlg.ShowDialog(Me) <> DialogResult.OK Then Return Nothing

                Dim picked As New List(Of String)
                For i = 0 To list.Items.Count - 1
                    If list.GetItemChecked(i) Then picked.Add(pages(i))
                Next
                Return picked
            End Using
        End Function

        Private Shared Function GetRelativePathSafe(root As String, fullPath As String) As String
            Try
                Dim rootUri As New Uri(If(root.EndsWith(Path.DirectorySeparatorChar), root, root & Path.DirectorySeparatorChar))
                Dim fileUri As New Uri(fullPath)
                Return Uri.UnescapeDataString(rootUri.MakeRelativeUri(fileUri).ToString()).Replace("/"c, Path.DirectorySeparatorChar)
            Catch
                Return fullPath
            End Try
        End Function

        Private Async Function RepairWikiPagesAsync(store As KnowledgeStoreCatalog.KnowledgeStoreDefinition,
                                                    pageFullPaths As List(Of String),
                                                    progressCallback As Action(Of String)) As Task(Of String)
            Dim manifest = KnowledgeStoreManifest.Load(store)
            Dim regenerated As Integer = 0
            Dim deletedOnly As Integer = 0
            Dim failed As New List(Of String)()
            Dim total = pageFullPaths.Count
            Dim idx = 0

            For Each pagePath In pageFullPaths
                idx += 1
                Dim relForLog = GetRelativePathSafe(store.ResolvedSourcePath, pagePath)
                progressCallback?.Invoke($"Repair Page ({idx}/{total}) {relForLog}: locating source...")

                Dim sourcePath = TryResolveSourceForWikiPage(pagePath, manifest)

                Try
                    If File.Exists(pagePath) Then File.Delete(pagePath)
                Catch ex As Exception
                    failed.Add($"{relForLog} — could not delete: {ex.Message}")
                    Continue For
                End Try

                If String.IsNullOrWhiteSpace(sourcePath) OrElse Not File.Exists(sourcePath) Then
                    deletedOnly += 1
                    failed.Add($"{relForLog} — no recoverable source path; page was deleted only.")
                    Continue For
                End If

                ' Heartbeat so the status bar / tray keep moving while
                ' IngestSourceAsync is inside the host gate or LLM call
                ' (it does not accept a progressCallback of its own).
                Dim heartbeatCts As New System.Threading.CancellationTokenSource()
                Dim ingestStartedUtc = DateTime.UtcNow
                Dim sourceName = Path.GetFileName(sourcePath)
                progressCallback?.Invoke($"Repair Page ({idx}/{total}) {relForLog}: starting LLM ingest of '{sourceName}'.")

                Dim heartbeatTask = Task.Run(
                    Async Function()
                        Try
                            While Not heartbeatCts.Token.IsCancellationRequested
                                Await Task.Delay(2000, heartbeatCts.Token).ConfigureAwait(False)
                                If heartbeatCts.Token.IsCancellationRequested Then Exit While
                                Dim elapsed = CInt(Math.Max(1, (DateTime.UtcNow - ingestStartedUtc).TotalSeconds))
                                progressCallback?.Invoke(
                                    $"Repair Page ({idx}/{total}) {relForLog}: still working ({elapsed}s) — waiting for LLM/host.")
                            End While
                        Catch
                        End Try
                    End Function,
                    heartbeatCts.Token)

                Dim ok As Boolean = False
                Dim ingestError As String = Nothing
                Try
                    manifest.RemoveByPath(sourcePath)
                    manifest.Save(store)

                    ok = Await KnowledgeWikiService.IngestSourceAsync(
                        kbRootPath:=store.ResolvedSourcePath,
                        sourceFilePath:=sourcePath,
                        context:=_context,
                        isBackground:=False).ConfigureAwait(False)
                Catch ex As Exception
                    ingestError = ex.Message
                End Try

                ' Stop heartbeat (Await must live outside Finally in VB).
                heartbeatCts.Cancel()
                Try
                    Await heartbeatTask.ConfigureAwait(False)
                Catch
                End Try
                heartbeatCts.Dispose()

                If ingestError IsNot Nothing Then
                    failed.Add($"{relForLog} — {ingestError}")
                ElseIf ok Then
                    Dim elapsedTotal = CInt(Math.Max(1, (DateTime.UtcNow - ingestStartedUtc).TotalSeconds))
                    progressCallback?.Invoke(
                        $"Repair Page ({idx}/{total}) {relForLog}: regenerated in {elapsedTotal}s.")
                    regenerated += 1
                Else
                    failed.Add($"{relForLog} — re-ingestion of '{sourcePath}' returned no content (LLM empty or host gate cancelled).")
                End If
            Next

            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine($"# Page Repair — {store.Name}")
            sb.AppendLine()
            sb.AppendLine($"- Pages requested: {total}")
            sb.AppendLine($"- Regenerated from source: {regenerated}")
            sb.AppendLine($"- Deleted (no source available): {deletedOnly}")
            sb.AppendLine($"- Failed: {failed.Count}")
            If failed.Count > 0 Then
                sb.AppendLine()
                sb.AppendLine("## Failures")
                For Each f In failed
                    sb.AppendLine($"- {f}")
                Next
            End If
            Return sb.ToString()
        End Function

        ''' <summary>
        ''' Tries to recover the originating source file path for a wiki page.
        ''' Order of preference:
        '''   1. YAML front matter `source_paths:` (LLM-authored pages).
        '''   2. `**Source:** \`...\`` line (auto summary pages).
        '''   3. Manifest entry whose Title matches the page H1.
        ''' </summary>
        Private Shared Function TryResolveSourceForWikiPage(pagePath As String,
                                                            manifest As KnowledgeStoreManifest) As String
            Try
                If Not File.Exists(pagePath) Then Return ""
                Dim text = File.ReadAllText(pagePath, System.Text.Encoding.UTF8)

                ' (1) YAML front matter: source_paths: \n  - 'path'  (or "path" or path)
                Dim mFront = System.Text.RegularExpressions.Regex.Match(
                    text, "\A---\s*\r?\n(?<fm>[\s\S]*?)\r?\n---\s*\r?\n")
                If mFront.Success Then
                    Dim fm = mFront.Groups("fm").Value
                    Dim mList = System.Text.RegularExpressions.Regex.Match(
                        fm, "(?ms)^source_paths\s*:\s*\r?\n(?<items>(?:[ \t]+-.*\r?\n?)+)")
                    If mList.Success Then
                        For Each line In mList.Groups("items").Value.
                                Split(New String() {vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)
                            Dim m = System.Text.RegularExpressions.Regex.Match(
                                line, "^\s*-\s*(?:'(?<v>(?:[^']|'')*)'|""(?<v>[^""]*)""|(?<v>.+?))\s*$")
                            If m.Success Then
                                Dim raw = m.Groups("v").Value.Replace("''", "'").Trim()
                                Dim expanded = ExpandEnvironmentVariables(raw)
                                If Not String.IsNullOrWhiteSpace(expanded) AndAlso File.Exists(expanded) Then
                                    Return expanded
                                End If
                            End If
                        Next
                    End If

                    ' Single-line variant: source_paths: ['x','y']  or  source_path: 'x'
                    Dim mSingle = System.Text.RegularExpressions.Regex.Match(
                        fm, "(?im)^source_path[s]?\s*:\s*(?<v>.+?)\s*$")
                    If mSingle.Success Then
                        Dim v = mSingle.Groups("v").Value.Trim().Trim("["c, "]"c)
                        For Each part In v.Split(","c)
                            Dim raw = part.Trim().Trim("'"c, """"c)
                            Dim expanded = ExpandEnvironmentVariables(raw)
                            If File.Exists(expanded) Then Return expanded
                        Next
                    End If
                End If

                ' (2) Legacy `**Source:**` line
                Dim mSrc = System.Text.RegularExpressions.Regex.Match(
                    text, "(?im)^\s*\*\*\s*Source\s*:\s*\*\*\s*`?([^`\r\n]+)`?\s*$")
                If mSrc.Success Then
                    Dim candidate = ExpandEnvironmentVariables(mSrc.Groups(1).Value.Trim())
                    If File.Exists(candidate) Then Return candidate
                End If

                ' (3) Manifest title fallback
                Dim mTitle = System.Text.RegularExpressions.Regex.Match(text, "(?m)^\s*#\s+(.+?)\s*$")
                If mTitle.Success AndAlso manifest IsNot Nothing Then
                    Dim title = mTitle.Groups(1).Value.Trim()
                    Dim hit = manifest.Entries.FirstOrDefault(
                        Function(en) Not String.IsNullOrWhiteSpace(en.Title) AndAlso
                                     en.Title.Equals(title, StringComparison.OrdinalIgnoreCase))
                    If hit IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(hit.FilePath) Then
                        Dim expanded = ExpandEnvironmentVariables(hit.FilePath)
                        If File.Exists(expanded) Then Return expanded
                    End If
                End If
            Catch
            End Try
            Return ""
        End Function

        Private Async Sub OnRepair(sender As Object, e As EventArgs)
            Dim store = GetSelectedStore()
            If store Is Nothing Then Return

            If String.IsNullOrWhiteSpace(store.ResolvedSourcePath) Then
                ShowCustomMessageBox("The selected store does not have a valid source path.", AN)
                Return
            End If

            If Not KnowledgeStoreCatalog.CanCurrentUserWrite(store, _context) Then
                ShowCustomMessageBox($"You do not have write permission to '{store.Name}'.", AN)
                Return
            End If

            Dim answer = ShowCustomYesNoBox(
        $"This will run automatic repairs on '{store.Name}', including LLM-assisted fixes. Continue?",
        "Yes, repair", "No, cancel")
            If answer <> 1 Then Return

            Dim buttonsRestoredEarly As Boolean = False

            SetOperationButtonsEnabled(False)
            System.Threading.Interlocked.Increment(_runningOperations)
            Try
                Dim summary = Await RunMaintenanceJobAsync(
            store:=store,
            operationName:="Repair",
            work:=Function(progressCallback)
                      Return KnowledgeWikiService.ApplyWikiHealthFixesAsync(
                          kbRootPath:=store.ResolvedSourcePath,
                          context:=_context,
                          includeLlmRepairs:=True,
                          progressCallback:=progressCallback,
                          operationName:="Repair")
                  End Function)

                If Me.IsDisposed OrElse Me.Disposing Then Return

                System.Threading.Interlocked.Decrement(_runningOperations)
                SetOperationButtonsEnabled(True)
                UpdateStatus()
                buttonsRestoredEarly = True

                ShowCustomWindow(
            $"Repair summary for '{store.Name}':",
            summary,
            "You can copy this summary to the clipboard.",
            $"{AN} Knowledge Store")

                LoadStatistics(store)
                UpdateStatus()
            Catch ex As Exception
                If Not Me.IsDisposed Then ShowCustomMessageBox($"Error during repair: {ex.Message}", AN)
            Finally
                If Not buttonsRestoredEarly Then
                    System.Threading.Interlocked.Decrement(_runningOperations)
                    If Not Me.IsDisposed Then
                        SetOperationButtonsEnabled(True)
                        UpdateStatus()
                    End If
                End If
            End Try
        End Sub

        Private Async Sub OnRevectorize(sender As Object, e As EventArgs)
            Dim store = GetSelectedStore()
            If store Is Nothing Then Return

            If String.IsNullOrWhiteSpace(store.ResolvedSourcePath) Then
                ShowCustomMessageBox("The selected store does not have a valid source path.", AN)
                Return
            End If

            Dim wikiRoot = Path.Combine(store.ResolvedSourcePath, ".redink", KnowledgeStoreCatalog.WikiFolder)
            If Not Directory.Exists(wikiRoot) Then
                ShowCustomMessageBox("No wiki pages found to rebuild embeddings for.", AN)
                Return
            End If

            Dim wikiPages = Directory.GetFiles(wikiRoot, "*.md", SearchOption.AllDirectories).
        Where(Function(f)
                  Dim n = Path.GetFileName(f)
                  Return Not n.Equals(KnowledgeStoreCatalog.IndexFile, StringComparison.OrdinalIgnoreCase) AndAlso
                         Not n.Equals(KnowledgeStoreCatalog.LogFile, StringComparison.OrdinalIgnoreCase) AndAlso
                         Not n.Equals("health_report.md", StringComparison.OrdinalIgnoreCase)
              End Function).
        ToList()

            If wikiPages.Count = 0 Then
                ShowCustomMessageBox("No wiki pages found to rebuild embeddings for.", AN)
                Return
            End If

            Dim answer = ShowCustomYesNoBox(
        $"This will rebuild the embedding index for '{store.Name}' from {wikiPages.Count} existing wiki page(s) " &
        "using the currently configured embedding model. Continue?",
        "Yes, rebuild embeddings", "No, cancel")
            If answer <> 1 Then Return

            Dim buttonsRestoredEarly As Boolean = False

            SetOperationButtonsEnabled(False)
            System.Threading.Interlocked.Increment(_runningOperations)
            Try
                ProgressBarModule.GlobalProgressValue = 0
                ProgressBarModule.GlobalProgressMax = wikiPages.Count
                ProgressBarModule.GlobalProgressLabel = "Preparing embedding rebuild..."
                ProgressBarModule.CancelOperation = False
                ProgressBarModule.ShowProgressBarInSeparateThread(
            $"{AN} Knowledge Store — Embeddings",
            "Refreshing embeddings...")

                Dim rebuiltPages = Await KnowledgeEmbeddingService.RebuildAllWikiEmbeddingsAsync(
            kbRootPath:=store.ResolvedSourcePath,
            context:=_context,
            progressPrefix:=store.Name,
            progressOffset:=0,
            progressTotal:=wikiPages.Count)

                Dim wasCancelled = ProgressBarModule.CancelOperation
                ProgressBarModule.CancelOperation = True

                If Me.IsDisposed OrElse Me.Disposing Then Return

                System.Threading.Interlocked.Decrement(_runningOperations)
                SetOperationButtonsEnabled(True)
                UpdateStatus()
                buttonsRestoredEarly = True

                ShowCustomMessageBox(
            If(wasCancelled,
               $"Embedding rebuild cancelled. Refreshed {rebuiltPages} wiki page embedding set(s).",
               $"Embedding rebuild complete. Refreshed {rebuiltPages} wiki page embedding set(s)."),
            $"{AN} Knowledge Store")

                LoadStatistics(store)
                UpdateStatus()
            Catch ex As Exception
                ProgressBarModule.CancelOperation = True
                If Not Me.IsDisposed Then ShowCustomMessageBox($"Error during embedding rebuild: {ex.Message}", AN)
            Finally
                If Not buttonsRestoredEarly Then
                    ProgressBarModule.CancelOperation = True
                    System.Threading.Interlocked.Decrement(_runningOperations)
                    If Not Me.IsDisposed Then
                        SetOperationButtonsEnabled(True)
                        UpdateStatus()
                    End If
                End If
            End Try
        End Sub

        Private Async Sub OnLint(sender As Object, e As EventArgs)
            Dim store = GetSelectedStore()
            If store Is Nothing Then Return

            If String.IsNullOrWhiteSpace(store.ResolvedSourcePath) Then
                ShowCustomMessageBox("The selected store does not have a valid source path.", AN)
                Return
            End If

            Dim buttonsRestoredEarly As Boolean = False

            SetOperationButtonsEnabled(False)
            System.Threading.Interlocked.Increment(_runningOperations)
            Try
                Dim report = Await RunMaintenanceJobAsync(
            store:=store,
            operationName:="Lint Wiki",
            work:=Function(progressCallback)
                      Return KnowledgeWikiService.LintWikiAsync(
                          kbRootPath:=store.ResolvedSourcePath,
                          context:=_context,
                          autoApply:=True,
                          progressCallback:=progressCallback)
                  End Function)

                If Me.IsDisposed OrElse Me.Disposing Then Return

                System.Threading.Interlocked.Decrement(_runningOperations)
                SetOperationButtonsEnabled(True)
                UpdateStatus()
                buttonsRestoredEarly = True

                ShowCustomWindow(
            $"Lint report for '{store.Name}':",
            report,
            "Auto-fixes have been applied. You can copy this report to the clipboard.",
            $"{AN} Knowledge Store")

                LoadStatistics(store)
                UpdateStatus()
            Catch ex As Exception
                If Not Me.IsDisposed Then ShowCustomMessageBox($"Error during wiki lint: {ex.Message}", AN)
            Finally
                If Not buttonsRestoredEarly Then
                    System.Threading.Interlocked.Decrement(_runningOperations)
                    If Not Me.IsDisposed Then
                        SetOperationButtonsEnabled(True)
                        UpdateStatus()
                    End If
                End If
            End Try
        End Sub



        Private Sub OnEditSchema(sender As Object, e As EventArgs)
            Dim store = GetSelectedStore()
            If store Is Nothing Then Return

            If String.IsNullOrWhiteSpace(store.ResolvedSourcePath) Then
                ShowCustomMessageBox("The selected store does not have a valid source path.", AN)
                Return
            End If

            Try
                KnowledgeStoreSchema.LoadOrCreate(store.ResolvedSourcePath)

                Dim schemaPath = KnowledgeStoreSchema.GetSchemaPath(store.ResolvedSourcePath)
                If String.IsNullOrWhiteSpace(schemaPath) Then
                    ShowCustomMessageBox("Could not resolve the schema path for the selected store.", AN)
                    Return
                End If

                ShowTextFileEditor(
                    schemaPath,
                    $"Edit schema for Knowledge Store '{store.Name}'.",
                    ForceJson:=True,
                    _context:=_context)
            Catch ex As Exception
            Catch ex As Exception
                ShowCustomMessageBox($"Error opening schema: {ex.Message}", AN)
            End Try
        End Sub

        Private Sub SetOperationButtonsEnabled(enabled As Boolean)
            _btnIndex.Enabled = enabled
            _btnReindex.Enabled = enabled
            _btnHealthCheck.Enabled = enabled
            _btnRepair.Enabled = enabled
            _btnRepairPage.Enabled = enabled
            _btnReformatPage.Enabled = enabled
            _btnRevectorize.Enabled = enabled
            _btnLint.Enabled = enabled
            _btnEditSchema.Enabled = enabled
            _btnAddStore.Enabled = enabled
            _btnDeleteStore.Enabled = enabled
            _btnSave.Enabled = enabled
            _lstStores.Enabled = enabled
        End Sub

#End Region

#Region "Key Handling"

        Private Sub OnKeyDown(sender As Object, e As KeyEventArgs)
            If e.KeyCode = Keys.Escape Then
                e.Handled = True
                Me.Close()
            End If
        End Sub

#End Region

    End Class

End Namespace
