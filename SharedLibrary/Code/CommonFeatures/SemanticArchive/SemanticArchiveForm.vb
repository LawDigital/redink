' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
Option Strict On
Option Explicit On

Namespace SharedLibrary

    Partial Public Class SharedMethods
        Public Shared Sub ShowSemanticArchiveConsole(context As SharedContext.ISharedContext)
            If Not SemanticArchiveHostIntegration.IsConfigured(context) Then Return
            SharedMethods.RequireInteractiveExecution("semantic_archive_console")
            Using dialog As New SemanticArchiveForm(context)
                Dim owner As System.Windows.Forms.IWin32Window = ResolveSameThreadDialogOwner()
                If owner IsNot Nothing Then
                    dialog.ShowDialog(owner)
                Else
                    dialog.ShowDialog()
                End If
            End Using
        End Sub
    End Class

    ''' <summary>Separate archive administration. Definitions are edited by stable IDs and saved with
    ''' a catalog revision check. I/O and builds run on workers, never on the Office UI thread.</summary>
    Public NotInheritable Class SemanticArchiveForm
        Inherits System.Windows.Forms.Form

        Private ReadOnly _context As SharedContext.ISharedContext
        Private _store As SemanticArchiveStore
        Private _catalog As SemanticArchiveCatalog
        Private _archive As SemanticArchiveDefinition
        Private _root As SemanticArchiveSourceBinding
        Private _editedBudgets As SemanticArchiveRetrievalBudgets
        Private _ownerScope As System.IDisposable
        Private _automaticPause As System.IDisposable
        Private _permissionsPause As System.IDisposable
        Private _operationCancellation As System.Threading.CancellationTokenSource
        Private _busy As Boolean
        Private _consoleOperations As System.Int32
        Private _closeRequested As System.Boolean
        Private _closeContinuationQueued As System.Boolean
        Private _diagnosticDisplayTrimmed As System.Boolean
        Private Const DiagnosticDisplayOmissionNotice As System.String = "[Older displayed diagnostic messages were omitted. Refresh status reads the stored archive and per-source status.]"
        Private Const DiagnosticEntryOmissionNotice As System.String = "[End of this overlong diagnostic entry omitted from the display.]"
        Private _loading As Boolean
        Private _catalogActivationBlocked As System.Boolean
        Private _catalogActivationMessage As System.String = ""
        Private _archiveIndexUnsupported As System.Boolean
        Private _archiveDirty As System.Boolean
        Private _backgroundDirty As System.Boolean
        Private _updatingWrapping As System.Boolean
        Private _layingOutSections As System.Boolean
        Private _layingOutSources As System.Boolean
        Private _archiveSelectionIndex As System.Int32 = -1
        Private _rootSelectionIndex As System.Int32 = -1
        Private _ownedIcon As System.Drawing.Icon
        Private ReadOnly _toolTips As New System.Windows.Forms.ToolTip() With {.AutoPopDelay = 20000, .InitialDelay = 400, .ReshowDelay = 100, .ShowAlways = True}
        ' Outer rows never ask nested panels for an unconstrained preferred height.
        ' Header, commands and status are measured from their leaf controls; the editor
        ' owns all remaining height. Long personal settings live in a scrollable tab.
        Private ReadOnly _layout As New System.Windows.Forms.TableLayoutPanel() With {.Dock = System.Windows.Forms.DockStyle.Fill, .AutoSize = False, .ColumnCount = 1, .RowCount = 5, .GrowStyle = System.Windows.Forms.TableLayoutPanelGrowStyle.FixedSize, .Padding = New System.Windows.Forms.Padding(16)}
        Private ReadOnly _catalogHeader As New System.Windows.Forms.Panel() With {.Dock = System.Windows.Forms.DockStyle.Fill, .AutoSize = False}
        Private ReadOnly _catalogCaption As New System.Windows.Forms.Label() With {.Text = "Personal catalog location", .AutoSize = False}
        Private ReadOnly _catalogButtons As New System.Windows.Forms.FlowLayoutPanel() With {.AutoSize = False, .WrapContents = True}
        Private ReadOnly _split As New System.Windows.Forms.SplitContainer() With {.Dock = System.Windows.Forms.DockStyle.Fill, .FixedPanel = System.Windows.Forms.FixedPanel.Panel1, .Size = New System.Drawing.Size(1150, 600), .SplitterDistance = 290}
        Private ReadOnly _archiveButtons As New System.Windows.Forms.FlowLayoutPanel() With {.Dock = System.Windows.Forms.DockStyle.Bottom, .AutoSize = False, .WrapContents = True}
        Private ReadOnly _commands As New System.Windows.Forms.Panel() With {.Dock = System.Windows.Forms.DockStyle.Fill, .AutoSize = False}
        Private ReadOnly _operations As New System.Windows.Forms.FlowLayoutPanel() With {.AutoSize = False, .AutoScroll = True, .WrapContents = True}
        Private ReadOnly _closeButton As System.Windows.Forms.Button = MakeButton("Close")
        Private ReadOnly _backgroundTab As New System.Windows.Forms.TabPage("Personal automatic processing") With {.AutoScroll = True}
        Private ReadOnly _rootEntry As New System.Windows.Forms.Panel() With {.Dock = System.Windows.Forms.DockStyle.Fill, .AutoSize = False, .Padding = New System.Windows.Forms.Padding(8)}
        Private ReadOnly _rootCaption As New System.Windows.Forms.Label() With {.Text = "Source folder to add", .AutoSize = False, .AutoEllipsis = True, .TextAlign = System.Drawing.ContentAlignment.MiddleLeft, .Visible = True}
        Private ReadOnly _rootButtons As New System.Windows.Forms.FlowLayoutPanel() With {.AutoSize = False, .WrapContents = True}
        Private ReadOnly _registeredSourcesPanel As New System.Windows.Forms.Panel() With {.Dock = System.Windows.Forms.DockStyle.Fill, .AutoSize = False}
        Private ReadOnly _registeredSourcesCaption As New System.Windows.Forms.Label() With {.Dock = System.Windows.Forms.DockStyle.Top, .AutoSize = False, .Text = "No source folders registered", .Padding = New System.Windows.Forms.Padding(8, 4, 8, 4)}
        Private ReadOnly _editorNotice As New System.Windows.Forms.Label() With {.AutoSize = True, .Dock = System.Windows.Forms.DockStyle.Fill, .AutoEllipsis = True, .Padding = New System.Windows.Forms.Padding(0, 3, 0, 0)}
        Private ReadOnly _statusArea As New System.Windows.Forms.TableLayoutPanel() With {.Dock = System.Windows.Forms.DockStyle.Fill, .AutoSize = False, .ColumnCount = 2, .RowCount = 1, .GrowStyle = System.Windows.Forms.TableLayoutPanelGrowStyle.FixedSize}
        Private ReadOnly _background As New System.Windows.Forms.GroupBox() With {.Dock = System.Windows.Forms.DockStyle.Top, .AutoSize = True, .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink, .Text = "Your automatic processing settings"}
        Private _libraryTab As System.Windows.Forms.TabPage
        Private ReadOnly _publishLibrary As System.Windows.Forms.Button = MakeButton("Publish / Update library")
        Private ReadOnly _withdrawLibrary As System.Windows.Forms.Button = MakeButton("Withdraw from library")
        Private ReadOnly _syncLibrary As System.Windows.Forms.Button = MakeButton("Sync library")
        Private ReadOnly _libraryStatus As New System.Windows.Forms.Label() With {.AutoSize = True, .MaximumSize = New System.Drawing.Size(530, 0)}
        Private ReadOnly _catalogLocation As New System.Windows.Forms.TextBox() With {.Dock = System.Windows.Forms.DockStyle.Fill, .ReadOnly = True}
        Private _paused As Boolean
        Private _resumeArchiveId As String = ""
        Private _resumeOperationId As String = ""
        Private _resumeCommand As MaintenanceCommand
        Private _resumeDocumentIds As System.Collections.Generic.List(Of String)
        Private _documentCursor As System.Collections.Generic.IEnumerator(Of SemanticArchiveDocumentRecord)
        Private _documentGenerationId As String = ""
        Private _documentListingComplete As Boolean

        Private Enum MaintenanceCommand
            RefreshContent
            Permissions
            SemanticReindex
            Extract
            Retry
        End Enum

        Private ReadOnly _archives As New System.Windows.Forms.ListBox() With {.Dock = System.Windows.Forms.DockStyle.Fill, .HorizontalScrollbar = True, .IntegralHeight = False}
        Private ReadOnly _roots As New System.Windows.Forms.ListBox() With {.Dock = System.Windows.Forms.DockStyle.Top, .Height = 120, .HorizontalScrollbar = True, .IntegralHeight = False}
        Private ReadOnly _tabs As New System.Windows.Forms.TabControl() With {.Dock = System.Windows.Forms.DockStyle.Fill, .Multiline = True}
        Private ReadOnly _name As New System.Windows.Forms.TextBox() With {.Dock = System.Windows.Forms.DockStyle.Fill}
        Private ReadOnly _description As New System.Windows.Forms.TextBox() With {.Dock = System.Windows.Forms.DockStyle.Fill, .Multiline = True, .Height = 75, .ScrollBars = System.Windows.Forms.ScrollBars.Vertical}
        Private ReadOnly _visible As New System.Windows.Forms.CheckBox() With {.Text = "Enable archive", .AutoSize = True}
        Private ReadOnly _defaultScope As New System.Windows.Forms.CheckBox() With {.Text = "Use by default", .AutoSize = True}
        Private ReadOnly _archiveBackground As New System.Windows.Forms.CheckBox() With {.Text = "Update this archive automatically", .AutoSize = True}
        Private ReadOnly _archiveWindow As New System.Windows.Forms.TextBox() With {.Dock = System.Windows.Forms.DockStyle.Fill}
        Private ReadOnly _partialSearch As New System.Windows.Forms.CheckBox() With {.Text = "Allow incomplete extracted text", .AutoSize = True}
        Private ReadOnly _threshold As New System.Windows.Forms.NumericUpDown() With {.Minimum = 0D, .Maximum = 9223372036854775807D, .Increment = 1024D, .Width = 180}
        Private ReadOnly _children As New System.Windows.Forms.NumericUpDown() With {.Minimum = 2D, .Maximum = 256D, .Width = 180}
        Private ReadOnly _routingCharacters As New System.Windows.Forms.NumericUpDown() With {.Minimum = 4096D, .Maximum = 120000D, .Increment = 1024D, .Width = 180}
        Private ReadOnly _budgets As New System.Windows.Forms.PropertyGrid() With {.Dock = System.Windows.Forms.DockStyle.Fill, .ToolbarVisible = False, .HelpVisible = True}
        Private ReadOnly _rootPath As New System.Windows.Forms.TextBox() With {.Dock = System.Windows.Forms.DockStyle.Fill, .ReadOnly = True}
        Private ReadOnly _newRootPath As New System.Windows.Forms.TextBox() With {.Dock = System.Windows.Forms.DockStyle.Fill, .ReadOnly = False}
        Private ReadOnly _rootLayout As New System.Windows.Forms.TableLayoutPanel() With {.Dock = System.Windows.Forms.DockStyle.Fill, .ColumnCount = 1, .RowCount = 3, .GrowStyle = System.Windows.Forms.TableLayoutPanelGrowStyle.FixedSize}
        Private ReadOnly _recursive As New System.Windows.Forms.CheckBox() With {.Text = "Include subfolders", .AutoSize = True}
        Private ReadOnly _ocr As New System.Windows.Forms.CheckBox() With {.Text = "Allow OCR", .AutoSize = True}
        Private ReadOnly _ocrBatchPages As New System.Windows.Forms.NumericUpDown() With {.Minimum = 1D, .Maximum = 75D, .Width = 100}
        Private ReadOnly _fileTypeFilter As New System.Windows.Forms.CheckBox() With {.Text = "Filter source file types", .AutoSize = True}
        Private ReadOnly _restoreFileTypes As System.Windows.Forms.Button = MakeButton("Restore office/image defaults")
        Private ReadOnly _extensions As New System.Windows.Forms.TextBox() With {.Dock = System.Windows.Forms.DockStyle.Fill}
        Private ReadOnly _exclusions As New System.Windows.Forms.TextBox() With {.Dock = System.Windows.Forms.DockStyle.Fill, .Multiline = True, .Height = 100, .ScrollBars = System.Windows.Forms.ScrollBars.Vertical}
        Private ReadOnly _placementMode As New System.Windows.Forms.ComboBox() With {.Dock = System.Windows.Forms.DockStyle.Fill, .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList}
        Private ReadOnly _sharedArtifactRoot As New System.Windows.Forms.TextBox() With {.Dock = System.Windows.Forms.DockStyle.Fill}
        Private ReadOnly _shadowArtifactRoot As New System.Windows.Forms.TextBox() With {.Dock = System.Windows.Forms.DockStyle.Fill}
        Private ReadOnly _documents As New System.Windows.Forms.ListView() With {.Dock = System.Windows.Forms.DockStyle.Fill, .View = System.Windows.Forms.View.Details, .FullRowSelect = True, .MultiSelect = True, .HideSelection = False}
        Private ReadOnly _documentFilter As New System.Windows.Forms.TextBox() With {.Width = 260}
        Private ReadOnly _documentStateFilter As New System.Windows.Forms.ComboBox() With {.Width = 210, .DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList}
        Private ReadOnly _documentTechnical As New System.Windows.Forms.CheckBox() With {.Text = "Show IDs / technical details", .AutoSize = True}
        Private ReadOnly _documentDetails As New System.Windows.Forms.TextBox() With {.Dock = System.Windows.Forms.DockStyle.Fill, .Multiline = True, .ReadOnly = True, .ScrollBars = System.Windows.Forms.ScrollBars.Vertical}
        Private ReadOnly _documentSelectionStatus As New System.Windows.Forms.Label() With {.AutoSize = True, .Padding = New System.Windows.Forms.Padding(4, 7, 4, 0)}
        Private ReadOnly _selectLoadedDocuments As System.Windows.Forms.Button = MakeButton("Select loaded matches")
        Private ReadOnly _clearDocumentSelection As System.Windows.Forms.Button = MakeButton("Clear selection")
        Private ReadOnly _documentAdvanced As New System.Windows.Forms.Panel() With {.Dock = System.Windows.Forms.DockStyle.Top, .Height = 70, .Visible = False}
        Private _updatingDocumentRows As System.Boolean
        Private _documentSortColumn As System.Int32
        Private _documentSortDescending As System.Boolean
        Private Const MaintenancePageSize As System.Int32 = 200
        Private Const MaintenanceLoadedLimit As System.Int32 = 5000
        Private ReadOnly _selectedDocumentIds As New System.Windows.Forms.TextBox() With {.Dock = System.Windows.Forms.DockStyle.Fill, .Multiline = True, .Height = 60, .ScrollBars = System.Windows.Forms.ScrollBars.Vertical}
        Private ReadOnly _documentPageStatus As New System.Windows.Forms.Label() With {.Dock = System.Windows.Forms.DockStyle.Bottom, .Height = 38, .AutoEllipsis = True}
        Private ReadOnly _scopeTags As New System.Windows.Forms.TextBox() With {.Dock = System.Windows.Forms.DockStyle.Fill}
        Private ReadOnly _rootEditor As New System.Windows.Forms.Panel() With {.Dock = System.Windows.Forms.DockStyle.Fill, .AutoScroll = True}
        Private ReadOnly _diagnostics As New System.Windows.Forms.TextBox() With {.Dock = System.Windows.Forms.DockStyle.Fill, .Multiline = True, .ReadOnly = True, .ScrollBars = System.Windows.Forms.ScrollBars.Both, .WordWrap = False, .MaxLength = SharedMethods.DEFAULT_SEMANTICARCHIVE_MAX_DIAGNOSTIC_DISPLAY_CHARACTERS}
        Private ReadOnly _diagnosticActions As New System.Windows.Forms.Panel() With {.Dock = System.Windows.Forms.DockStyle.Top, .AutoSize = False}
        Private ReadOnly _copyDiagnostics As System.Windows.Forms.Button = MakeButton("Copy diagnostics")
        Private ReadOnly _showTechnicalDiagnostics As New System.Windows.Forms.CheckBox() With {
            .Text = "Show technical details", .AutoSize = True,
            .Checked = SharedMethods.DEFAULT_SEMANTICARCHIVE_SHOW_TECHNICAL_DIAGNOSTICS}
        Private ReadOnly _status As New System.Windows.Forms.Label() With {.Dock = System.Windows.Forms.DockStyle.Fill, .AutoSize = True, .MinimumSize = New System.Drawing.Size(0, 38), .AutoEllipsis = True, .Padding = New System.Windows.Forms.Padding(0, 8, 0, 0)}
        Private ReadOnly _progress As New System.Windows.Forms.ProgressBar() With {.Dock = System.Windows.Forms.DockStyle.Bottom, .Height = 12, .Style = System.Windows.Forms.ProgressBarStyle.Continuous}
        Private ReadOnly _globalEnabled As New System.Windows.Forms.CheckBox() With {.Text = "Enable automatic content indexing", .AutoSize = True}
        Private ReadOnly _globalWindow As New System.Windows.Forms.TextBox() With {.Width = 280}
        Private ReadOnly _permissionsEnabled As New System.Windows.Forms.CheckBox() With {.Text = "Enable automatic permission maintenance", .AutoSize = True}
        Private ReadOnly _permissionsWindow As New System.Windows.Forms.TextBox() With {.Width = 280}
        Private ReadOnly _chooseCatalog As System.Windows.Forms.Button = MakeButton("Enter catalog path…")
        Private ReadOnly _browseCatalog As System.Windows.Forms.Button = MakeButton("Browse…")
        Private ReadOnly _suggestCatalog As System.Windows.Forms.Button = MakeButton("Use recommended location…")
        Private ReadOnly _newArchive As System.Windows.Forms.Button = MakeButton("New archive")
        Private ReadOnly _removeArchive As System.Windows.Forms.Button = MakeButton("Unregister")
        Private ReadOnly _reload As System.Windows.Forms.Button = MakeButton("Reload")
        Private ReadOnly _save As System.Windows.Forms.Button = MakeButton("Save changes")
        Private ReadOnly _addRoot As System.Windows.Forms.Button = MakeButton("Browse directory…")
        Private ReadOnly _addRootPath As System.Windows.Forms.Button = MakeButton("Add source folder")
        Private ReadOnly _removeRoot As System.Windows.Forms.Button = MakeButton("Remove root")
        Private ReadOnly _saveRoot As System.Windows.Forms.Button = MakeButton("Save source and archive settings")
        Private ReadOnly _refresh As System.Windows.Forms.Button = MakeButton("Refresh archive")
        Private ReadOnly _pause As System.Windows.Forms.Button = MakeButton("Pause")
        Private ReadOnly _retry As System.Windows.Forms.Button = MakeButton("Retry failed")
        Private ReadOnly _rebuild As System.Windows.Forms.Button = MakeButton("Rebuild semantic index")
        Private ReadOnly _extractAll As System.Windows.Forms.Button = MakeButton("Re-extract/OCR all")
        Private ReadOnly _permissionsAll As System.Windows.Forms.Button = MakeButton("Permissions: all")
        Private ReadOnly _rebuildSelected As System.Windows.Forms.Button = MakeButton("Rebuild semantic index (selected)")
        Private ReadOnly _extractSelected As System.Windows.Forms.Button = MakeButton("Re-extract/OCR selected")
        Private ReadOnly _permissionsSelected As System.Windows.Forms.Button = MakeButton("Permissions: selected")
        Private ReadOnly _retrySelected As System.Windows.Forms.Button = MakeButton("Retry selected")
        Private ReadOnly _loadDocuments As System.Windows.Forms.Button = MakeButton("Find documents")
        Private ReadOnly _nextDocuments As System.Windows.Forms.Button = MakeButton("Load more matches")
        Private ReadOnly _inspect As System.Windows.Forms.Button = MakeButton("Refresh status")
        Private ReadOnly _saveBackground As System.Windows.Forms.Button = MakeButton("Save personal settings")

        Public Sub New(context As SharedContext.ISharedContext)
            If context Is Nothing Then Throw New System.ArgumentNullException(NameOf(context))
            If Not SemanticArchiveHostIntegration.IsConfigured(context) Then Throw New System.InvalidOperationException("SemanticArchiveCatalogPathLocal is not configured; Semantic Archives are disabled.")
            _context = context
            InitializeForm()
        End Sub

        Protected Overrides Sub OnHandleCreated(e As System.EventArgs)
            MyBase.OnHandleCreated(e)
            If _ownerScope Is Nothing Then _ownerScope = SharedMethods.PushDialogOwner(Me)
        End Sub

        Protected Overrides Sub OnHandleDestroyed(e As System.EventArgs)
            If _ownerScope IsNot Nothing Then _ownerScope.Dispose()
            _ownerScope = Nothing
            MyBase.OnHandleDestroyed(e)
        End Sub

        Protected Overrides Async Sub OnShown(e As System.EventArgs)
            MyBase.OnShown(e)
            InitializeBranding()
            ConstrainToWorkingArea()
            LayoutConsoleSections()
            TopMost = True
            SharedMethods.ForceDialogToForeground(Me)
            SharedMethods.AttachForeignForegroundWatchdog(Me)
            Await RunConsoleOperationAsync(Function() ReloadCatalogAsync())
        End Sub

        Protected Overrides Sub OnFormClosing(e As System.Windows.Forms.FormClosingEventArgs)
            If _busy OrElse _consoleOperations > 0 Then
                ' Retain this STA, its owner scope and any active Office reader until
                ' the entire user operation finishes, including nested save/reload work.
                _closeRequested = True
                If _operationCancellation IsNot Nothing Then _operationCancellation.Cancel()
                e.Cancel = True
                _status.Text = "Stopping at a safe checkpoint. This window will close automatically when the operation has stopped."
                AppendDiagnostic(_status.Text)
                Return
            End If
            If e.CloseReason = System.Windows.Forms.CloseReason.UserClosing AndAlso Not ConfirmDiscardChanges(True) Then
                _closeRequested = False
                e.Cancel = True
                Return
            End If
            _closeRequested = False
            ReleaseAutomaticPause()
            ResetDocumentListing()
            MyBase.OnFormClosing(e)
        End Sub

        ''' <summary>Tracks the full button/Shown operation, not individual SetBusy intervals.</summary>
        Private Async Function RunConsoleOperationAsync(operation As System.Func(Of System.Threading.Tasks.Task)) As System.Threading.Tasks.Task
            If operation Is Nothing Then Throw New System.ArgumentNullException(NameOf(operation))
            If IsDisposed OrElse _closeRequested Then Return
            _consoleOperations += 1
            Try
                Await operation.Invoke()
            Catch ex As System.Exception
                ReportError("Archive console operation failed", ex)
            Finally
                _consoleOperations -= 1
                CompletePendingClose()
            End Try
        End Function

        Private Sub CompletePendingClose()
            If Not _closeRequested OrElse _busy OrElse _consoleOperations > 0 OrElse
               _closeContinuationQueued OrElse IsDisposed OrElse Not IsHandleCreated Then Return
            _closeContinuationQueued = True
            Try
                BeginInvoke(New System.Windows.Forms.MethodInvoker(
                    Sub()
                        _closeContinuationQueued = False
                        If _closeRequested AndAlso Not _busy AndAlso _consoleOperations = 0 AndAlso Not IsDisposed Then Close()
                    End Sub))
            Catch ex As System.InvalidOperationException
                _closeContinuationQueued = False
                System.Diagnostics.Debug.WriteLine("Semantic Archive pending close: " & ex.Message)
            End Try
        End Sub

        Private Sub ConstrainToWorkingArea()
            Dim workingArea As System.Drawing.Rectangle = System.Windows.Forms.Screen.FromControl(Me).WorkingArea
            MinimumSize = New System.Drawing.Size(System.Math.Min(MinimumSize.Width, workingArea.Width),
                                                 System.Math.Min(MinimumSize.Height, workingArea.Height))
            Dim width As System.Int32 = System.Math.Min(Me.Width, workingArea.Width)
            Dim height As System.Int32 = System.Math.Min(Me.Height, workingArea.Height)
            Dim left As System.Int32 = System.Math.Max(workingArea.Left, System.Math.Min(Me.Left, workingArea.Right - width))
            Dim top As System.Int32 = System.Math.Max(workingArea.Top, System.Math.Min(Me.Top, workingArea.Bottom - height))
            Bounds = New System.Drawing.Rectangle(left, top, width, height)
        End Sub

        Protected Overrides Sub Dispose(disposing As System.Boolean)
            If disposing Then
                _toolTips.Dispose()
                If _ownedIcon IsNot Nothing Then
                    Icon = Nothing
                    _ownedIcon.Dispose()
                    _ownedIcon = Nothing
                End If
            End If
            MyBase.Dispose(disposing)
        End Sub

        <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
        Private Shared Function DestroyIcon(iconHandle As System.IntPtr) As System.Boolean
        End Function

        Private Sub InitializeBranding()
            If _ownedIcon IsNot Nothing Then Return
            Try
                Using bitmap As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                    Dim iconHandle As System.IntPtr = bitmap.GetHicon()
                    Try
                        Using borrowed As System.Drawing.Icon = System.Drawing.Icon.FromHandle(iconHandle)
                            _ownedIcon = DirectCast(borrowed.Clone(), System.Drawing.Icon)
                            Icon = _ownedIcon
                        End Using
                    Finally
                        DestroyIcon(iconHandle)
                    End Try
                End Using
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("Semantic Archive branding: " & ex.Message)
            End Try
        End Sub

        Private Shared Function MakeButton(text As String) As System.Windows.Forms.Button
            Return New System.Windows.Forms.Button() With {.Text = text, .AutoSize = True, .Margin = New System.Windows.Forms.Padding(4), .MinimumSize = New System.Drawing.Size(85, 28)}
        End Function

        Private Shared Function NewTable() As System.Windows.Forms.TableLayoutPanel
            Dim table As New System.Windows.Forms.TableLayoutPanel() With {.Dock = System.Windows.Forms.DockStyle.Top, .AutoSize = True, .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink, .ColumnCount = 2, .RowCount = 0, .GrowStyle = System.Windows.Forms.TableLayoutPanelGrowStyle.FixedSize, .Padding = New System.Windows.Forms.Padding(8)}
            table.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 38))
            table.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100))
            Return table
        End Function

        Private Shared Sub AddRow(table As System.Windows.Forms.TableLayoutPanel, caption As String, control As System.Windows.Forms.Control)
            Dim row = table.RowCount
            table.RowCount += 1
            table.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            Dim label As New System.Windows.Forms.Label() With {.Text = caption, .AutoSize = True, .MaximumSize = New System.Drawing.Size(275, 0), .Padding = New System.Windows.Forms.Padding(0, 5, 4, 0)}
            control.Margin = New System.Windows.Forms.Padding(3, 3, 3, 7)
            table.Controls.Add(label, 0, row)
            table.Controls.Add(control, 1, row)
        End Sub

        Private Sub AddParameterRow(table As System.Windows.Forms.TableLayoutPanel, caption As String,
                                           parameterName As String, control As System.Windows.Forms.Control)
            AddRow(table, caption, control)
            _toolTips.SetToolTip(control, "Catalog parameter: " & parameterName & ".")
        End Sub

        Private Sub InitializeForm()
            Text = SharedMethods.AN & " — Semantic Archives"
            Font = New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
            AutoScaleDimensions = New System.Drawing.SizeF(96.0F, 96.0F)
            Size = New System.Drawing.Size(1440, 960)
            MinimumSize = New System.Drawing.Size(1000, 740)
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen

            _split.Panel1.Padding = New System.Windows.Forms.Padding(0, 0, 8, 0)
            _split.Panel1.Controls.Add(_archives)
            _archiveButtons.Controls.AddRange(New System.Windows.Forms.Control() {_newArchive, _removeArchive, _reload})
            _split.Panel1.Controls.Add(_archiveButtons)
            _split.Panel2.Controls.Add(_tabs)

            Dim archiveTab As New System.Windows.Forms.TabPage("Archive") With {.AutoScroll = True}
            _editorNotice.Text = "Select an archive, or create one after loading a writable catalog. Background settings can be saved independently."
            Dim archiveTable = NewTable()
            AddParameterRow(archiveTable, "Archive name", "SemanticArchiveName", _name)
            AddParameterRow(archiveTable, "Description", "SemanticArchiveDescription", _description)
            AddParameterRow(archiveTable, "Search visibility", "SemanticArchiveEnabled", _visible)
            AddParameterRow(archiveTable, "Default selection", "SemanticArchiveDefaultArchiveIds", _defaultScope)
            AddParameterRow(archiveTable, "Archive background", "SemanticArchiveBackgroundEnabled", _archiveBackground)
            AddParameterRow(archiveTable, "Archive processing window", "SemanticArchiveBackgroundWindow", _archiveWindow)
            AddParameterRow(archiveTable, "Extraction coverage", "SemanticArchiveAllowPartialSearch", _partialSearch)
            AddRow(archiveTable, "Generated file access", New System.Windows.Forms.Label() With {.Text = "Personal navigation and private shadows stay in this Windows user's access domain. Cooperative document artifacts follow verified source permissions. Source files are never made writable or have their permissions changed.", .AutoSize = True, .MaximumSize = New System.Drawing.Size(530, 0)})
            AddParameterRow(archiveTable, "Create section indexes from (bytes; 0 = off)", "SemanticArchiveSourceIndexThresholdBytes", _threshold)
            AddParameterRow(archiveTable, "Navigation branches per node", "SemanticArchiveMaxChildrenPerNode", _children)
            AddParameterRow(archiveTable, "Navigation summary size (characters)", "SemanticArchiveMaxRoutingCharacters", _routingCharacters)
            archiveTab.Controls.Add(archiveTable)
            _tabs.TabPages.Add(archiveTab)
            _libraryTab = CreateLibraryTab()
            _tabs.TabPages.Add(_libraryTab)

            Dim rootsTab As New System.Windows.Forms.TabPage("Source roots")
            _newRootPath.Dock = System.Windows.Forms.DockStyle.None
            _rootButtons.Controls.AddRange(New System.Windows.Forms.Control() {_addRootPath, _addRoot, _removeRoot})
            _rootEntry.Controls.AddRange(New System.Windows.Forms.Control() {_rootCaption, _newRootPath, _rootButtons})
            Dim rootTable = NewTable()
            AddParameterRow(rootTable, "Selected source folder", "SemanticArchiveRootPath", _rootPath)
            AddParameterRow(rootTable, "Traversal", "SemanticArchiveRecursive", _recursive)
            AddParameterRow(rootTable, "Extraction", "SemanticArchiveEnableOcr", _ocr)
            AddParameterRow(rootTable, "OCR pages per model call", "SemanticArchiveOcrBatchPages", _ocrBatchPages)
            AddParameterRow(rootTable, "Source type filter", "SemanticArchiveFileTypeFilterEnabled", _fileTypeFilter)
            AddParameterRow(rootTable, "Allowed extensions (blank = defaults)", "SemanticArchiveSupportedExtensions", _extensions)
            AddRow(rootTable, "", _restoreFileTypes)
            AddHandler _restoreFileTypes.Click, Sub(sender, args)
                                                   _extensions.Text = SharedMethods.DEFAULT_SEMANTICARCHIVE_SUPPORTED_EXTENSIONS.Replace(";", "; ")
                                               End Sub
            AddHandler _fileTypeFilter.CheckedChanged, Sub(sender, args)
                                                           _extensions.Enabled = _fileTypeFilter.Checked
                                                           _restoreFileTypes.Enabled = _fileTypeFilter.Checked
                                                       End Sub
            AddParameterRow(rootTable, "Exclude files or folders (one per line)", "SemanticArchiveExclusions", _exclusions)
            _placementMode.Items.AddRange(New Object() {"auto", "private"})
            AddParameterRow(rootTable, "Generated file storage", "SemanticArchiveArtifactPlacementMode", _placementMode)
            AddParameterRow(rootTable, "Shared derivative base folder", "SemanticArchiveSharedArtifactRoot", _sharedArtifactRoot)
            AddParameterRow(rootTable, "Private derivative folder", "SemanticArchiveShadowArtifactRoot", _shadowArtifactRoot)
            AddRow(rootTable, "Placement and access", New System.Windows.Forms.Label() With {.Text = "auto reuses/publishes cooperative document artifacts when source permissions can be enforced, otherwise uses a writable private shadow. private uses shadows only. Source read access is sufficient; diagnostics explain fallback and Windows path-length failures. Re-extract/OCR uses the OCR setting above.", .AutoSize = True, .MaximumSize = New System.Drawing.Size(530, 0)})
            AddParameterRow(rootTable, "Source labels (optional)", "SemanticArchiveScopeTags", _scopeTags)
            AddRow(rootTable, "", _saveRoot)
            _rootEditor.Controls.Add(rootTable)
            _rootLayout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100))
            _rootLayout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 1))
            _rootLayout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 0))
            _rootLayout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100))
            For Each section As System.Windows.Forms.Control In New System.Windows.Forms.Control() {_rootEntry, _registeredSourcesPanel, _rootEditor}
                section.Margin = New System.Windows.Forms.Padding(0)
            Next
            _roots.Dock = System.Windows.Forms.DockStyle.Fill
            _roots.AccessibleName = "Registered source folders"
            _registeredSourcesPanel.Controls.Add(_roots)
            _registeredSourcesPanel.Controls.Add(_registeredSourcesCaption)
            _rootLayout.Controls.Add(_rootEntry, 0, 0)
            _rootLayout.Controls.Add(_registeredSourcesPanel, 0, 1)
            _rootLayout.Controls.Add(_rootEditor, 0, 2)
            rootsTab.Controls.Add(_rootLayout)
            _tabs.TabPages.Add(rootsTab)
            Dim budgetTab As New System.Windows.Forms.TabPage("Retrieval budgets")
            budgetTab.Controls.Add(_budgets)
            budgetTab.Controls.Add(New System.Windows.Forms.Label() With {.Dock = System.Windows.Forms.DockStyle.Top, .AutoSize = True, .Padding = New System.Windows.Forms.Padding(8), .Text = "Limits for one retrieval call; the lowest limit across selected archives applies. Select a setting to read its explanation below."})
            _tabs.TabPages.Add(budgetTab)
            _tabs.TabPages.Add(CreateDocumentsTab())
            Dim statusTab As New System.Windows.Forms.TabPage("Status and diagnostics")
            statusTab.Controls.Add(_diagnostics)
            _diagnosticActions.Controls.Add(_copyDiagnostics)
            _diagnosticActions.Controls.Add(_showTechnicalDiagnostics)
            statusTab.Controls.Add(_diagnosticActions)
            _copyDiagnostics.Enabled = False
            AddHandler _copyDiagnostics.Click, Sub(sender, args) CopyDiagnostics()
            AddHandler _showTechnicalDiagnostics.CheckedChanged, Async Sub(sender, args)
                If Not _busy Then Await RunConsoleOperationAsync(Function() RefreshDiagnosticsAsync())
            End Sub
            AddHandler _diagnostics.TextChanged, Sub(sender, args) _copyDiagnostics.Enabled = _diagnostics.TextLength > 0
            _tabs.TabPages.Add(statusTab)

            _operations.Controls.AddRange(New System.Windows.Forms.Control() {_save, _refresh, _permissionsAll, _rebuild, _extractAll, _retry, _pause, _inspect})
            AddHandler _closeButton.Click, Sub(sender, args) Me.Close()
            _commands.Controls.AddRange(New System.Windows.Forms.Control() {_operations, _closeButton})
            Dim backgroundTable As System.Windows.Forms.TableLayoutPanel = NewTable()
            backgroundTable.Dock = System.Windows.Forms.DockStyle.Top
            AddRow(backgroundTable, "Content indexing", _globalEnabled)
            AddRow(backgroundTable, "Indexing hours (local time)", _globalWindow)
            AddRow(backgroundTable, "Source and artifact permissions", _permissionsEnabled)
            AddRow(backgroundTable, "Permission check hours (local time)", _permissionsWindow)
            AddRow(backgroundTable, "", _saveBackground)
            _globalWindow.Dock = System.Windows.Forms.DockStyle.Fill
            _permissionsWindow.Dock = System.Windows.Forms.DockStyle.Fill
            _background.Controls.Add(backgroundTable)

            _backgroundTab.Controls.Add(_background)
            _tabs.TabPages.Add(_backgroundTab)
            _catalogLocation.Dock = System.Windows.Forms.DockStyle.None
            _catalogButtons.Controls.AddRange(New System.Windows.Forms.Control() {_chooseCatalog, _browseCatalog, _suggestCatalog})
            _catalogHeader.Controls.AddRange(New System.Windows.Forms.Control() {_catalogCaption, _catalogLocation, _catalogButtons})
            _statusArea.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 70))
            _statusArea.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 30))
            _statusArea.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100))
            _statusArea.Controls.Add(_status, 0, 0)
            _statusArea.Controls.Add(_editorNotice, 1, 0)
            _editorNotice.Padding = New System.Windows.Forms.Padding(8, 8, 0, 0)
            For Each section As System.Windows.Forms.Control In New System.Windows.Forms.Control() {_catalogHeader, _split, _commands, _progress, _statusArea, _status, _editorNotice}
                section.Margin = New System.Windows.Forms.Padding(0)
            Next
            _layout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100))
            _layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 1))
            _layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100))
            _layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 1))
            _layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 12))
            _layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 1))
            _layout.Controls.Add(_catalogHeader, 0, 0)
            _layout.Controls.Add(_split, 0, 1)
            _layout.Controls.Add(_commands, 0, 2)
            _layout.Controls.Add(_progress, 0, 3)
            _layout.Controls.Add(_statusArea, 0, 4)
            Controls.Add(_layout)
            AddHandler _layout.Layout, Sub(sender, args) LayoutConsoleSections()
            AddHandler _layout.SizeChanged, Sub(sender, args) LayoutConsoleSections()
            AddHandler _split.Panel1.SizeChanged, Sub(sender, args) LayoutConsoleSections()
            AddHandler _rootLayout.Layout, Sub(sender, args) LayoutSourceSections()
            AddHandler _tabs.SelectedIndexChanged, Sub(sender, args)
                                                      LayoutSourceSections()
                                                      ConstrainWrappingText(_tabs)
                                                  End Sub
            AddHandler FontChanged, Sub(sender, args) LayoutConsoleSections()
            ConfigureToolTips(_closeButton)
            WireEditorChanges()
            _catalogLocation.Text = _context.INI_SemanticArchiveCatalogPathLocal

            AddHandler _archives.SelectedIndexChanged, AddressOf ArchiveSelected
            AddHandler _roots.SelectedIndexChanged, AddressOf RootSelected
            AddHandler _chooseCatalog.Click, Async Sub(sender, args) Await RunConsoleOperationAsync(Function() ChooseCatalogPathAsync(False, False))
            AddHandler _browseCatalog.Click, Async Sub(sender, args) Await RunConsoleOperationAsync(Function() ChooseCatalogPathAsync(True, False))
            AddHandler _suggestCatalog.Click, Async Sub(sender, args) Await RunConsoleOperationAsync(Function() ChooseCatalogPathAsync(False, True))
            AddHandler _newArchive.Click, Async Sub(sender, args) Await RunConsoleOperationAsync(Function() CreateArchiveAsync())
            AddHandler _removeArchive.Click, Async Sub(sender, args) Await RunConsoleOperationAsync(Function() RemoveArchiveAsync())
            AddHandler _publishLibrary.Click, Async Sub(sender, args) Await RunConsoleOperationAsync(Function() PublishLibraryAsync(False))
            AddHandler _withdrawLibrary.Click, Async Sub(sender, args) Await RunConsoleOperationAsync(Function() PublishLibraryAsync(True))
            AddHandler _syncLibrary.Click, Async Sub(sender, args) Await RunConsoleOperationAsync(Function() SyncLibraryAsync())
            AddHandler _reload.Click, Async Sub(sender, args) Await RunConsoleOperationAsync(Function() RequestReloadCatalogAsync())
            AddHandler _save.Click, Async Sub(sender, args) Await RunConsoleOperationAsync(Function() SaveArchiveAsync())
            AddHandler _addRoot.Click, Async Sub(sender, args) Await RunConsoleOperationAsync(Function() AddRootAsync())
            AddHandler _addRootPath.Click, Async Sub(sender, args) Await RunConsoleOperationAsync(Function() AddRootPathAsync())
            AddHandler _removeRoot.Click, Async Sub(sender, args) Await RunConsoleOperationAsync(Function() RemoveRootAsync())
            AddHandler _saveRoot.Click, Async Sub(sender, args) Await RunConsoleOperationAsync(Function() SaveArchiveAsync())
            AddHandler _saveBackground.Click, Async Sub(sender, args) Await RunConsoleOperationAsync(Function() SaveBackgroundAsync())
            AddHandler _refresh.Click, Async Sub(sender, args) Await RunConsoleOperationAsync(Function() RunBuildAsync(MaintenanceCommand.RefreshContent, False))
            AddHandler _retry.Click, Async Sub(sender, args) Await RunConsoleOperationAsync(Function() RunBuildAsync(MaintenanceCommand.Retry, False))
            AddHandler _rebuild.Click, Async Sub(sender, args) Await RunConsoleOperationAsync(Function() RunBuildAsync(MaintenanceCommand.SemanticReindex, False))
            AddHandler _extractAll.Click, Async Sub(sender, args) Await RunConsoleOperationAsync(Function() RunBuildAsync(MaintenanceCommand.Extract, False))
            AddHandler _permissionsAll.Click, Async Sub(sender, args) Await RunConsoleOperationAsync(Function() RunBuildAsync(MaintenanceCommand.Permissions, False))
            AddHandler _retrySelected.Click, Async Sub(sender, args) Await RunConsoleOperationAsync(Function() RunBuildAsync(MaintenanceCommand.Retry, True))
            AddHandler _rebuildSelected.Click, Async Sub(sender, args) Await RunConsoleOperationAsync(Function() RunBuildAsync(MaintenanceCommand.SemanticReindex, True))
            AddHandler _extractSelected.Click, Async Sub(sender, args) Await RunConsoleOperationAsync(Function() RunBuildAsync(MaintenanceCommand.Extract, True))
            AddHandler _permissionsSelected.Click, Async Sub(sender, args) Await RunConsoleOperationAsync(Function() RunBuildAsync(MaintenanceCommand.Permissions, True))
            AddHandler _loadDocuments.Click, Async Sub(sender, args) Await RunConsoleOperationAsync(Function() LoadDocumentPageAsync(True))
            AddHandler _nextDocuments.Click, Async Sub(sender, args) Await RunConsoleOperationAsync(Function() LoadDocumentPageAsync(False))
            AddHandler _documents.SelectedIndexChanged, AddressOf DocumentSelectionChanged
            AddHandler _documents.ColumnClick, AddressOf DocumentColumnClicked
            AddHandler _documentFilter.TextChanged, AddressOf DocumentFilterChanged
            AddHandler _documentStateFilter.SelectedIndexChanged, AddressOf DocumentFilterChanged
            AddHandler _selectLoadedDocuments.Click, AddressOf SelectLoadedDocuments
            AddHandler _clearDocumentSelection.Click, Sub(sender, args)
                                                         _updatingDocumentRows = True
                                                         _documents.BeginUpdate()
                                                         Try
                                                             For Each item As System.Windows.Forms.ListViewItem In _documents.Items
                                                                 item.Selected = False
                                                             Next
                                                         Finally
                                                             _documents.EndUpdate()
                                                             _updatingDocumentRows = False
                                                         End Try
                                                         _selectedDocumentIds.Clear()
                                                         UpdateDocumentDetails()
                                                     End Sub
            AddHandler _documentTechnical.CheckedChanged, Sub(sender, args)
                                                             _documents.Columns(4).Width = If(_documentTechnical.Checked, 270, 0)
                                                             _documentAdvanced.Visible = _documentTechnical.Checked
                                                             UpdateDocumentDetails()
                                                         End Sub
            AddHandler _selectedDocumentIds.TextChanged, Sub(sender, args)
                                                            UpdateSelectedActions()
                                                            UpdateDocumentDetails()
                                                        End Sub
            AddHandler _pause.Click, Async Sub(sender, args) Await RunConsoleOperationAsync(Function() PauseResumeAsync())
            AddHandler _inspect.Click, Async Sub(sender, args) Await RunConsoleOperationAsync(Function() RefreshDiagnosticsAsync())
            SetBusy(False)
            LayoutConsoleSections()
        End Sub

        Private Sub WireEditorChanges()
            For Each input As System.Windows.Forms.TextBox In New System.Windows.Forms.TextBox() {_name, _description, _archiveWindow, _extensions, _exclusions, _sharedArtifactRoot, _shadowArtifactRoot, _scopeTags}
                AddHandler input.TextChanged, AddressOf ArchiveEditorChanged
            Next
            For Each input As System.Windows.Forms.CheckBox In New System.Windows.Forms.CheckBox() {_visible, _defaultScope, _archiveBackground, _partialSearch, _recursive, _ocr, _fileTypeFilter}
                AddHandler input.CheckedChanged, AddressOf ArchiveEditorChanged
            Next
            For Each input As System.Windows.Forms.NumericUpDown In New System.Windows.Forms.NumericUpDown() {_threshold, _children, _routingCharacters, _ocrBatchPages}
                AddHandler input.ValueChanged, AddressOf ArchiveEditorChanged
            Next
            AddHandler _newRootPath.TextChanged, AddressOf NewSourcePathChanged
            AddHandler _placementMode.SelectedIndexChanged, AddressOf ArchiveEditorChanged
            AddHandler _budgets.PropertyValueChanged, Sub(sender, args) ArchiveEditorChanged(sender, System.EventArgs.Empty)
            AddHandler _globalEnabled.CheckedChanged, AddressOf BackgroundEditorChanged
            AddHandler _globalWindow.TextChanged, AddressOf BackgroundEditorChanged
            AddHandler _permissionsEnabled.CheckedChanged, AddressOf BackgroundEditorChanged
            AddHandler _permissionsWindow.TextChanged, AddressOf BackgroundEditorChanged
        End Sub

        Private Sub ArchiveEditorChanged(sender As System.Object, e As System.EventArgs)
            If _loading OrElse _archive Is Nothing Then Return
            _archiveDirty = True
            UpdateSaveState()
        End Sub

        Private Sub BackgroundEditorChanged(sender As System.Object, e As System.EventArgs)
            If _loading Then Return
            _backgroundDirty = True
            UpdateSaveState()
        End Sub

        Private Sub NewSourcePathChanged(sender As System.Object, e As System.EventArgs)
            If _loading Then Return
            UpdateSaveState()
        End Sub

        Private Function HasPendingSourcePath() As System.Boolean
            Return _archive IsNot Nothing AndAlso Not System.String.IsNullOrWhiteSpace(_newRootPath.Text)
        End Function

        Private Sub MarkArchiveSaved()
            _archiveDirty = False
            UpdateSaveState()
        End Sub

        Private Sub MarkBackgroundSaved()
            _backgroundDirty = False
            UpdateSaveState()
        End Sub

        Private Sub UpdateSaveState()
            If IsDisposed Then Return
            _save.Enabled = Not _busy AndAlso Not _catalogActivationBlocked AndAlso _archive IsNot Nothing AndAlso _archiveDirty
            _saveRoot.Enabled = _save.Enabled AndAlso _root IsNot Nothing
            _saveBackground.Enabled = Not _busy AndAlso _backgroundDirty
            _addRootPath.Enabled = Not _busy AndAlso Not _catalogActivationBlocked AndAlso Not SemanticArchiveLibrary.IsSubscriber(_archive) AndAlso HasPendingSourcePath()
            If _catalogActivationBlocked Then
                _editorNotice.Text = "Catalog actions are locked until configuration and catalog reload succeed. The displayed catalog and your draft are retained. Use Reload or choose a catalog location; personal background settings remain available."
                _toolTips.SetToolTip(_editorNotice, _editorNotice.Text & System.Environment.NewLine & _catalogActivationMessage)
                Return
            End If
            If _archive Is Nothing Then
                _editorNotice.Text = If(_catalog Is Nothing,
                    "Choose a writable personal catalog location to create archives. Personal background settings can be saved independently.",
                    "Create an archive to edit its settings and add source folders. Personal background settings can be saved independently.")
            Else
                _editorNotice.Text = If(_archiveDirty, "Unsaved archive changes — Save changes also saves source options and retrieval budgets.", "Archive settings are saved.")
            End If
            If HasPendingSourcePath() Then _editorNotice.Text &= " A source folder has not been added yet; choose Add source folder."
            If _backgroundDirty Then _editorNotice.Text &= " Personal background settings have separate unsaved changes."
            _toolTips.SetToolTip(_editorNotice, _editorNotice.Text & System.Environment.NewLine & "Archive changes and personal background settings are saved separately.")
        End Sub

        Private Function ConfirmDiscardChanges(includeBackground As System.Boolean) As System.Boolean
            Dim pendingSource As System.Boolean = HasPendingSourcePath()
            If Not _archiveDirty AndAlso Not pendingSource AndAlso (Not includeBackground OrElse Not _backgroundDirty) Then Return True
            Dim scope As System.String = If(_archiveDirty AndAlso includeBackground AndAlso _backgroundDirty, "archive and personal background", If(_archiveDirty, "archive", "personal background"))
            If pendingSource Then
                scope = If(_archiveDirty, "archive and source folder", "source folder")
                If includeBackground AndAlso _backgroundDirty Then scope &= " and personal background"
            End If
            Dim reminder As System.String = If(pendingSource, "Choose Add source folder to register the entered path, and save any other changes, if you want to keep them.", "Save them first if you want to keep them.")
            Return ConfirmConsoleAction("Discard unsaved " & scope & " changes? " & reminder, "Discard changes", "Keep editing")
        End Function

        Private Function ConfirmConsoleAction(message As System.String, acceptText As System.String, cancelText As System.String) As System.Boolean
            Using dialogOwner As System.IDisposable = SharedMethods.PushDialogOwner(Me)
                Return SharedMethods.ShowCustomYesNoBox(message, acceptText, cancelText, "Semantic Archives") = 1
            End Using
        End Function

        Private Sub ConstrainWrappingText(parent As System.Windows.Forms.Control)
            If _updatingWrapping Then Return
            _updatingWrapping = True
            Try
                ConstrainWrappingChildren(parent)
            Finally
                _updatingWrapping = False
            End Try
        End Sub

        Private Sub ConstrainWrappingChildren(parent As System.Windows.Forms.Control)
            Dim table As System.Windows.Forms.TableLayoutPanel = TryCast(parent, System.Windows.Forms.TableLayoutPanel)
            Dim widths As System.Int32() = If(table Is Nothing, Nothing, table.GetColumnWidths())
            For Each control As System.Windows.Forms.Control In parent.Controls
                If TypeOf control Is System.Windows.Forms.Label OrElse TypeOf control Is System.Windows.Forms.CheckBox Then
                    Dim available As System.Int32 = parent.ClientSize.Width - parent.Padding.Horizontal - control.Margin.Horizontal
                    If table IsNot Nothing Then
                        Dim column As System.Int32 = table.GetColumn(control)
                        If column >= 0 AndAlso column < widths.Length Then
                            available = 0
                            For index As System.Int32 = column To System.Math.Min(widths.Length - 1, column + table.GetColumnSpan(control) - 1)
                                available += widths(index)
                            Next
                            available -= control.Margin.Horizontal
                        End If
                    End If
                    If available > 0 Then
                        Dim maximumHeight As System.Int32 = 0
                        If System.Object.ReferenceEquals(control, _status) Then maximumHeight = _status.Font.Height * 3 + _status.Padding.Vertical
                        If System.Object.ReferenceEquals(control, _editorNotice) Then maximumHeight = _editorNotice.Font.Height * 2 + _editorNotice.Padding.Vertical
                        control.MaximumSize = New System.Drawing.Size(available, maximumHeight)
                    End If
                End If
                If control.HasChildren AndAlso Not TypeOf control Is System.Windows.Forms.PropertyGrid Then ConstrainWrappingChildren(control)
            Next
        End Sub

        Private Function ScaleSpacing(logicalPixels As System.Int32) As System.Int32
            Return System.Math.Max(1, CInt(System.Math.Ceiling(logicalPixels * DeviceDpi / 96.0R)))
        End Function

        ''' <summary>Measures buttons only, never a nested panel's default/preferred size.</summary>
        Private Shared Function MeasureButtonRows(panel As System.Windows.Forms.FlowLayoutPanel, availableWidth As System.Int32) As System.Int32
            Dim width As System.Int32 = System.Math.Max(1, availableWidth - panel.Padding.Horizontal)
            Dim usedWidth As System.Int32 = 0
            Dim rowHeight As System.Int32 = 0
            Dim height As System.Int32 = panel.Padding.Vertical
            For Each button As System.Windows.Forms.Control In panel.Controls
                Dim preferred As System.Drawing.Size = button.GetPreferredSize(System.Drawing.Size.Empty)
                Dim buttonWidth As System.Int32 = System.Math.Max(button.MinimumSize.Width, preferred.Width) + button.Margin.Horizontal
                Dim buttonHeight As System.Int32 = System.Math.Max(button.MinimumSize.Height, preferred.Height) + button.Margin.Vertical
                If usedWidth > 0 AndAlso usedWidth + buttonWidth > width Then
                    height += rowHeight
                    usedWidth = 0
                    rowHeight = 0
                End If
                usedWidth += buttonWidth
                rowHeight = System.Math.Max(rowHeight, buttonHeight)
            Next
            Return height + rowHeight
        End Function

        Private Function ArrangePathEntry(panel As System.Windows.Forms.Panel, caption As System.Windows.Forms.Label,
                                          input As System.Windows.Forms.TextBox, buttons As System.Windows.Forms.FlowLayoutPanel,
                                          availableWidth As System.Int32, alignButtonsWithInput As System.Boolean) As System.Int32
            Dim gap As System.Int32 = ScaleSpacing(6)
            Dim width As System.Int32 = System.Math.Max(1, availableWidth - panel.Padding.Horizontal)
            Dim captionWidth As System.Int32 = System.Math.Max(1, System.Math.Min(ScaleSpacing(275), CInt(width * 0.28R)))
            caption.MaximumSize = System.Drawing.Size.Empty
            Dim captionHeight As System.Int32 = System.Windows.Forms.TextRenderer.MeasureText(caption.Text, caption.Font,
                New System.Drawing.Size(captionWidth, System.Int32.MaxValue), System.Windows.Forms.TextFormatFlags.WordBreak Or System.Windows.Forms.TextFormatFlags.TextBoxControl).Height
            Dim inputHeight As System.Int32 = input.PreferredHeight
            Dim rowHeight As System.Int32 = System.Math.Max(captionHeight, inputHeight)
            caption.SetBounds(panel.Padding.Left, panel.Padding.Top, captionWidth, rowHeight)
            input.SetBounds(panel.Padding.Left + captionWidth + gap, panel.Padding.Top,
                            System.Math.Max(1, width - captionWidth - gap), inputHeight)
            Dim buttonLeft As System.Int32 = If(alignButtonsWithInput, input.Left, panel.Padding.Left)
            Dim buttonWidth As System.Int32 = System.Math.Max(1, availableWidth - panel.Padding.Right - buttonLeft)
            Dim buttonHeight As System.Int32 = MeasureButtonRows(buttons, buttonWidth)
            buttons.SetBounds(buttonLeft, panel.Padding.Top + rowHeight + gap, buttonWidth, buttonHeight)
            Return buttons.Bottom + panel.Padding.Bottom + gap
        End Function

        Private Sub LayoutConsoleSections()
            If _layingOutSections OrElse IsDisposed OrElse _layout.RowStyles.Count <> 5 Then Return
            _layingOutSections = True
            Try
                Dim width As System.Int32 = System.Math.Max(1, _layout.ClientSize.Width - _layout.Padding.Horizontal)
                Dim headerHeight As System.Int32 = ArrangePathEntry(_catalogHeader, _catalogCaption, _catalogLocation, _catalogButtons, width, True)
                Dim closeSize As System.Drawing.Size = _closeButton.GetPreferredSize(System.Drawing.Size.Empty)
                Dim closeWidth As System.Int32 = System.Math.Max(_closeButton.MinimumSize.Width, closeSize.Width)
                Dim closeHeight As System.Int32 = System.Math.Max(_closeButton.MinimumSize.Height, closeSize.Height)
                Dim operationWidth As System.Int32 = System.Math.Max(1, width - closeWidth - _closeButton.Margin.Horizontal)
                Dim fullCommandHeight As System.Int32 = MeasureButtonRows(_operations, operationWidth - System.Windows.Forms.SystemInformation.VerticalScrollBarWidth)
                Dim commandHeight As System.Int32 = System.Math.Max(closeHeight + _closeButton.Margin.Vertical,
                    System.Math.Min(fullCommandHeight, (closeHeight + _closeButton.Margin.Vertical) * 2))
                _operations.SetBounds(0, 0, operationWidth, commandHeight)
                _operations.AutoScrollMinSize = New System.Drawing.Size(0, fullCommandHeight)
                _closeButton.SetBounds(width - closeWidth - _closeButton.Margin.Right, _closeButton.Margin.Top, closeWidth, closeHeight)
                Dim statusHeight As System.Int32 = System.Math.Max(_status.Font.Height * 3 + _status.Padding.Vertical,
                    _editorNotice.Font.Height * 2 + _editorNotice.Padding.Vertical)
                _layout.SuspendLayout()
                Try
                    _layout.RowStyles(0).Height = headerHeight
                    _layout.RowStyles(2).Height = commandHeight
                    _layout.RowStyles(3).Height = ScaleSpacing(12)
                    _layout.RowStyles(4).Height = statusHeight
                Finally
                    _layout.ResumeLayout(True)
                End Try
                Dim archiveWidth As System.Int32 = System.Math.Max(1, _split.Panel1.ClientSize.Width - _split.Panel1.Padding.Horizontal)
                _archiveButtons.Height = MeasureButtonRows(_archiveButtons, archiveWidth)
                Dim copySize As System.Drawing.Size = _copyDiagnostics.GetPreferredSize(System.Drawing.Size.Empty)
                _copyDiagnostics.SetBounds(_copyDiagnostics.Margin.Left, _copyDiagnostics.Margin.Top, copySize.Width, copySize.Height)
                Dim technicalSize As System.Drawing.Size = _showTechnicalDiagnostics.GetPreferredSize(System.Drawing.Size.Empty)
                Dim technicalLeft As System.Int32 = _copyDiagnostics.Right + _copyDiagnostics.Margin.Right + _showTechnicalDiagnostics.Margin.Left
                Dim technicalTop As System.Int32 = _showTechnicalDiagnostics.Margin.Top
                If technicalLeft + technicalSize.Width + _showTechnicalDiagnostics.Margin.Right > _diagnosticActions.ClientSize.Width Then
                    technicalLeft = _showTechnicalDiagnostics.Margin.Left
                    technicalTop = _copyDiagnostics.Bottom + _copyDiagnostics.Margin.Bottom + _showTechnicalDiagnostics.Margin.Top
                End If
                _showTechnicalDiagnostics.SetBounds(technicalLeft, technicalTop, technicalSize.Width, technicalSize.Height)
                _diagnosticActions.Height = System.Math.Max(_copyDiagnostics.Bottom + _copyDiagnostics.Margin.Bottom,
                    _showTechnicalDiagnostics.Bottom + _showTechnicalDiagnostics.Margin.Bottom)
                ConstrainWrappingText(_layout)
                LayoutSourceSections()
            Finally
                _layingOutSections = False
            End Try
        End Sub

        Private Sub LayoutSourceSections()
            If _layingOutSources OrElse IsDisposed OrElse _rootLayout.RowStyles.Count <> 3 Then Return
            _layingOutSources = True
            Try
                Dim width As System.Int32 = System.Math.Max(1, _rootLayout.ClientSize.Width - _rootLayout.Padding.Horizontal)
                _rootCaption.Visible = True
                _rootCaption.BringToFront()
                Dim entryHeight As System.Int32 = ArrangePathEntry(_rootEntry, _rootCaption, _newRootPath, _rootButtons, width, False)
                Dim listHeight As System.Int32 = 0
                Dim captionHeight As System.Int32 = _registeredSourcesCaption.Font.Height + _registeredSourcesCaption.Padding.Vertical
                _registeredSourcesCaption.Height = captionHeight
                _registeredSourcesCaption.Text = If(_roots.Items.Count = 0, "No source folders registered",
                    "Registered source folders (" & _roots.Items.Count.ToString() & ")")
                _roots.Visible = _roots.Items.Count > 0
                If _roots.Items.Count > 0 Then
                    Dim rows As System.Int32 = System.Math.Min(4, _roots.Items.Count)
                    Dim borderHeight As System.Int32 = System.Windows.Forms.SystemInformation.BorderSize.Height * 2 + System.Windows.Forms.SystemInformation.HorizontalScrollBarHeight
                    Dim optionHeight As System.Int32 = _rootPath.PreferredHeight + _rootPath.Margin.Vertical + ScaleSpacing(16)
                    Dim available As System.Int32 = _rootLayout.ClientSize.Height - entryHeight - captionHeight - optionHeight - borderHeight
                    rows = System.Math.Min(rows, System.Math.Max(1, available \ System.Math.Max(1, _roots.ItemHeight)))
                    listHeight = rows * _roots.ItemHeight + borderHeight
                End If
                _rootLayout.SuspendLayout()
                Try
                    _rootLayout.RowStyles(0).Height = entryHeight
                    _rootLayout.RowStyles(1).Height = captionHeight + listHeight
                Finally
                    _rootLayout.ResumeLayout(True)
                End Try
                ConstrainWrappingText(_rootLayout)
            Finally
                _layingOutSources = False
            End Try
        End Sub

        Private Sub SetHelp(control As System.Windows.Forms.Control, explanation As System.String, Optional parameterName As System.String = Nothing)
            Dim help As System.String = explanation & If(System.String.IsNullOrEmpty(parameterName), "", System.Environment.NewLine & "Parameter: " & parameterName & ".")
            _toolTips.SetToolTip(control, help)
            Dim table As System.Windows.Forms.TableLayoutPanel = TryCast(control.Parent, System.Windows.Forms.TableLayoutPanel)
            If table IsNot Nothing AndAlso table.GetColumn(control) = 1 Then
                Dim label As System.Windows.Forms.Control = table.GetControlFromPosition(0, table.GetRow(control))
                If label IsNot Nothing Then _toolTips.SetToolTip(label, help)
            End If
        End Sub

        Private Sub ConfigureToolTips(closeButton As System.Windows.Forms.Button)
            SetHelp(_chooseCatalog, "Enter a personal catalog directory. The location is checked before the existing personal Red Ink INI configuration is updated. Original and shared derivative folders are independent.")
            SetHelp(_browseCatalog, "Choose an existing personal catalog directory with the Windows folder picker. Red Ink checks that private catalog and index files can be protected here; original and shared derivative folders are separate.")
            SetHelp(_suggestCatalog, "Offer the recommended short folder under your Windows LocalAppData directory. Existing catalog data is retained at its current location; selecting a different folder does not move it.")
            SetHelp(_archives, "Select an archive to edit. Unsaved changes must be saved or explicitly discarded before switching archives.")
            SetHelp(_roots, "Registered source folders in this archive. Select a row to inspect its full path and source options below. Changes to each source are retained in the archive draft until Save changes.")
            SetHelp(_registeredSourcesCaption, "The number of source folders already registered in the selected archive. Add a new folder in the entry above; existing folders remain in the list below and do not need to be entered again.")
            SetHelp(_tabs, "Archive settings, source folders and retrieval budgets are saved together. Status and document actions operate on the selected archive.")
            SetHelp(_catalogLocation, "Your personal catalog and navigation directory, containing redink-sa-catalog.json; enter a folder, not a JSON filename. Original files and shared derivatives are separate and may be on a network share.", "SemanticArchiveCatalogPathLocal")
            SetHelp(_name, "Name used in sa: requests and in the model-visible catalog overview. Choose a clear, descriptive name; stable IDs distinguish duplicate names.", "SemanticArchiveName")
            SetHelp(_description, "Describe the topics, document types and intended use. This description is visible to the model within its allowed archive scope so it can choose suitable sources. It is navigation metadata, not document evidence; it grants no file permissions.", "SemanticArchiveDescription")
            SetHelp(_visible, "Allow this archive in search and automatic content indexing. Turning this off does not delete data or stop independent permission maintenance.", "SemanticArchiveEnabled")
            SetHelp(_defaultScope, "Use this archive as a saved default when no session selection exists. Without saved defaults, local sessions may use enabled archives. An explicit empty session selection still disables archive use. This setting does not start indexing or require archive search for every prompt.", "SemanticArchiveDefaultArchiveIds")
            SetHelp(_archiveBackground, "Include this archive in background content indexing when your personal indexing setting is on and both time windows allow it. Manual commands remain available.", "SemanticArchiveBackgroundEnabled")
            SetHelp(_archiveWindow, "Local time. Blank means any time. Examples: allow:22:00-06:00 or deny:08:00-18:00; separate ranges with ;. Both this and your personal indexing window must allow automatic work. Permission maintenance has its own window.", "SemanticArchiveBackgroundWindow")
            SetHelp(_partialSearch, "Allow available text whose extraction is incomplete or unknown. Results still report incomplete coverage and require current source access. Unchecked: only complete extractions are eligible.", "SemanticArchiveAllowPartialSearch")
            SetHelp(_threshold, "Create an additional semantic section index for source files at least this many bytes long. 65,536 bytes = 64 KiB. 0 disables this additional index; archive navigation and document metadata are still built. Indexed documents are read through selected sections.", "SemanticArchiveSourceIndexThresholdBytes")
            SetHelp(_children, "Maximum child branches or document cards in one archive navigation node (2 to 256). Higher values make the hierarchy broader.", "SemanticArchiveMaxChildrenPerNode")
            SetHelp(_routingCharacters, "Maximum routing text retained in one navigation node (4,096 to 120,000 characters). This is not a limit on extracted document text.", "SemanticArchiveMaxRoutingCharacters")
            SetHelp(_budgets, "Select a retrieval limit to read its description below the grid. These limits bound model calls, traversal, candidate selection and returned text; all are saved with the archive.", "SemanticArchiveRetrievalBudgets")
            SetHelp(_newRootPath, "Type or paste an absolute source folder, for example C:\Documents or \\server\share\Archive, then choose Add source folder. Environment variables and a quoted pasted path are accepted. Browse directory also adds a source. The entry is kept if adding fails; it is not registered until adding succeeds. Red Ink never changes original-file permissions.", "SemanticArchiveRootPath")
            _toolTips.SetToolTip(_rootCaption, _toolTips.GetToolTip(_newRootPath))
            _toolTips.SetToolTip(_catalogCaption, _toolTips.GetToolTip(_catalogLocation))
            SetHelp(_backgroundTab, "Personal content indexing and permission-maintenance settings apply across your archives and are saved separately. This tab also remains available before an archive is created.")
            SetHelp(_rootPath, "Registered source folder containing originals. Environment variables such as %APPDATA% and %LOCALAPPDATA% are supported. This display is read-only because the binding has a stable identity. To use another directory, add it above and remove the old root if appropriate; source files and existing artifacts are retained. Red Ink never changes original-file permissions.", "SemanticArchiveRootPath")
            SetHelp(_recursive, "Discover supported files in this folder and all included subfolders. Exclusions still apply.", "SemanticArchiveRecursive")
            SetHelp(_ocr, "Permit the existing extractor to use OCR only for PDF pages whose native text layer appears insufficient. Re-extract commands respect this setting.", "SemanticArchiveEnableOcr")
            SetHelp(_ocrBatchPages, "Maximum contiguous OCR candidate pages sent to one model call. Default 16; allowed 1-75. Larger batches reduce model calls; failed batches remain retryable without changing document identity.", "SemanticArchiveOcrBatchPages")
            SetHelp(_fileTypeFilter, "On by default: process only the allowed Office, PDF, text, mail and image extensions below. Turn off to allow every format supported by the existing converter, including technical text formats. This never includes generated Knowledge Store (.redink) or archive outputs. Save and Refresh after changing it.", "SemanticArchiveFileTypeFilterEnabled")
            SetHelp(_extensions, "Separate extensions with ;, for example .pdf; .docx; .txt. Blank restores the office/image defaults. Only formats supported by text_export_to_text can be processed. A leading * is accepted (for example *.pdf); use the exclusion field for path patterns. The list is preserved but ignored when the source type filter is off.", "SemanticArchiveSupportedExtensions")
            SetHelp(_restoreFileTypes, "Restore the central office/image extension list for this source. This edits the draft only; use Save changes, then Refresh to apply it.")
            SetHelp(_exclusions, "One full path, source-relative path or * / ? pattern per line. A pattern without a slash matches file names; with a slash it matches relative paths. Excluded folders exclude their contents. Examples: Temp and *.tmp.", "SemanticArchiveExclusions")
            SetHelp(_placementMode, "auto: reuse or publish shared derivatives with verified source-reader permissions where supported, otherwise use private storage. private: use private derivatives only. Source files are never made writable.", "SemanticArchiveArtifactPlacementMode")
            SetHelp(_sharedArtifactRoot, "Optional local or UNC base folder for reusable document text and indexes. Blank uses the source folder. Red Ink adds .redink-sa automatically; do not append it yourself. Environment variables such as %APPDATA% and %LOCALAPPDATA% are supported. This is separate from your personal catalog.", "SemanticArchiveSharedArtifactRoot")
            SetHelp(_shadowArtifactRoot, "Preferred private output folder. Blank uses the per-user LocalAppData shadow. Unavailable or overlong preferred locations can fall back to LocalAppData. Environment variables such as %APPDATA% and %LOCALAPPDATA% are supported. This is not shared storage.", "SemanticArchiveShadowArtifactRoot")
            SetHelp(_scopeTags, "Optional labels stored with this source; separate with ; or ,. These labels currently do not filter search or grant access.", "SemanticArchiveScopeTags")
            SetHelp(_globalEnabled, "Allow Outlook to perform background discovery, extraction and indexing for enabled archives that also allow automatic indexing. Word does not host Semantic Archive automatic maintenance. Work runs only while Outlook is open and idle; manual indexing commands remain available in Word and Outlook. Stored for your Windows user and overrides the INI default.", "SemanticArchiveBackgroundIndexingEnabled")
            SetHelp(_globalWindow, "Local time for Outlook-hosted Semantic Archive automatic maintenance. Blank = any time. Examples: allow:22:00-06:00 or deny:08:00-18:00; separate ranges with ;. This and each archive's window must allow automatic content work.", "SemanticArchiveBackgroundIndexingWindow")
            SetHelp(_permissionsEnabled, "Allow Outlook to independently check generated-file access against originals while Outlook is open and idle. Word does not host this automatic Semantic Archive maintenance. This does not enable content indexing, rewrite source permissions or grant access to originals. Stored for your Windows user and overrides the INI default.", "SemanticArchivePermissionMaintenanceEnabled")
            SetHelp(_permissionsWindow, "Local time. Blank = any time. Examples: allow:22:00-06:00 or deny:08:00-18:00; separate ranges with ;. Independent of content indexing and archive time windows.", "SemanticArchivePermissionMaintenanceWindow")
            SetHelp(_saveBackground, "Save both personal switches and time windows. These apply to all your archives and are separate from Save changes. Values remain editable when no archive is selected.")
            SetHelp(_save, "Save changed archive settings, all edited source options and retrieval budgets together. Requires an archive and a writable catalog; disabled while busy or when nothing has changed.")
            SetHelp(_saveRoot, "Save all changed archive settings, all edited source options and retrieval budgets together, just like Save changes.")
            SetHelp(_newArchive, "Create a named entry in the personal catalog. Does not create source folders or index files. Existing archive changes are saved first.")
            SetHelp(_removeArchive, "Unregister this archive. Original files and generated artifacts are retained.")
            SetHelp(_reload, "Reload saved catalog and personal background settings. After an activation failure, first reload the normal Red Ink configuration. Unsaved changes are discarded only after confirmation and successful loading; a failure retains the draft.")
            SetHelp(_addRoot, "Choose and add a source folder with the Windows folder picker. Archive changes are saved before opening the picker, even if it is then cancelled. A cancelled picker retains the typed source entry.")
            SetHelp(_addRootPath, "Add the source folder typed above to this archive. Absolute local and UNC paths are accepted; an existing identical source is selected instead of duplicated. Source read access is enough when derivatives use writable private storage. Existing archive changes are saved first, and a failed addition retains the typed path.")
            SetHelp(_removeRoot, "Remove the selected source from this archive. Original files are retained. Existing changes are saved first.")
            SetHelp(_refresh, "Save archive changes, discover new or changed files, and reuse valid extracted text and indexes.")
            SetHelp(_rebuild, "Save archive changes, then rebuild semantic cards, applicable section indexes and routing from validated existing extracted text. This operation never invokes extraction or OCR; documents without a compatible valid extract are reported as needing extraction.")
            SetHelp(_extractAll, "Save archive changes, then extract all text again. OCR is used only where enabled for the source.")
            SetHelp(_permissionsAll, "Save archive changes, then reconcile generated-file permissions for all documents. Original-file permissions never change; insufficient rights can defer repairs.")
            SetHelp(_retry, "Save archive changes and retry failed or coverage-excluded work across the selected archive. Missing/empty/incomplete/unknown extraction is freshly extracted/OCRed; verified complete extracts are reused when only indexing or semantic work failed.")
            SetHelp(_pause, "Pause content and permission work in this Office context at safe checkpoints; published search remains available. Resume retains the exact operation and selection. Closing releases this temporary automatic pause.")
            SetHelp(_inspect, "Read current published-generation status and work diagnostics without starting indexing.")
            SetHelp(_documents, "Matching published documents. Click column headers to sort loaded matches. Multi-select rows for targeted actions. Names, paths and source status require current source access; inaccessible sources appear only in All documents or Source access unavailable.")
            SetHelp(_documentFilter, "Filter names/paths or document IDs. Changing the filter clears results and selection; Find documents applies it across the published metadata.")
            SetHelp(_selectedDocumentIds, "One stable document ID per line, up to 1,024. Use the first column of the document list. Empty never means all documents.")
            SetHelp(_loadDocuments, "Find matching documents across the published metadata. Empty chunks are skipped automatically. Pause cancels the search. No original-directory scan is started.")
            SetHelp(_nextDocuments, "Append matching documents from the same published snapshot. Loaded rows and selections remain. Sort applies to loaded matches; Find documents starts with the current snapshot.")
            SetHelp(_rebuildSelected, "Save archive changes and rebuild semantic indexes only for the explicit document IDs below. Valid extracted text is retained; this operation never invokes extraction or OCR.")
            SetHelp(_extractSelected, "Save archive changes and extract text again only for explicit document IDs. Each source's OCR option still applies.")
            SetHelp(_permissionsSelected, "Save archive changes and reconcile generated-file permissions only for explicit document IDs. Original permissions remain unchanged.")
            SetHelp(_retrySelected, "Save archive changes and retry failed or coverage-excluded work only for explicit document IDs. Extraction is repeated only when the stored extraction is missing/empty/incomplete/unknown; complete extracts are reused. Empty selection never means all.")
            SetHelp(_diagnostics, "Read-only operation details and errors, including path or permission failures. Select text and use Copy diagnostics, or copy the whole view when no text is selected.")
            SetHelp(_copyDiagnostics, "Copy selected diagnostic text to the clipboard; if no text is selected, copy all displayed diagnostics.")
            SetHelp(_showTechnicalDiagnostics, "Include deduplicated routing-metadata reductions and local hierarchy splits. These are technical observations, not failed document conversions. Saved details are bounded; source-specific messages require current access to the original.")
            SetHelp(_status, "The latest operation result. Full errors also appear under Status and diagnostics.")
            SetHelp(_progress, "Activity indicator for an ongoing operation. Pause or Close requests a stop at a safe checkpoint.")
            SetHelp(closeButton, "Close this console. Running work first stops at a safe checkpoint; unsaved changes require confirmation.")
        End Sub

        Private Async Function ChooseCatalogPathAsync(browse As System.Boolean, suggested As System.Boolean) As System.Threading.Tasks.Task
            If _closeRequested OrElse _busy Then Return
            Dim selected As System.String
            Try
                If browse Then
                    Using picker As New System.Windows.Forms.FolderBrowserDialog() With {.Description = "Choose your personal Semantic Archive catalog directory. Source and shared derivative folders are configured separately.", .ShowNewFolderButton = False}
                        Using dialogOwner As System.IDisposable = SharedMethods.PushDialogOwner(Me)
                            Dim owner As System.Windows.Forms.IWin32Window = SharedMethods.ResolveSameThreadDialogOwner()
                            Dim result As System.Windows.Forms.DialogResult = If(owner Is Nothing, picker.ShowDialog(), picker.ShowDialog(owner))
                            If result <> System.Windows.Forms.DialogResult.OK Then Return
                            selected = picker.SelectedPath
                        End Using
                    End Using
                Else
                    Dim initial As System.String = If(suggested, SemanticArchiveStore.GetSuggestedPrivateCatalogDirectory(), _context.INI_SemanticArchiveCatalogPathLocal)
                    Using dialogOwner As System.IDisposable = SharedMethods.PushDialogOwner(Me)
                        selected = SharedMethods.ShowCustomInputBox(
                            "Enter your personal catalog directory. This location stores private catalog and index files for your Windows user, SYSTEM and administrators." & System.Environment.NewLine &
                            "Originals and shared derivatives may remain on a network share. A new catalog location does not move or delete existing catalog data.",
                            "Personal Semantic Archive catalog", True, initial)
                    End Using
                End If
                If System.String.IsNullOrWhiteSpace(selected) Then Return
                selected = selected.Trim()
                If Not ConfirmDiscardChanges(False) Then Return
                SetBusy(True)
                Await System.Threading.Tasks.Task.Run(Sub() SemanticArchiveStore.ValidatePrivateCatalogLocation(selected))
                If _closeRequested Then Return
                Dim configurationFile As System.String = SemanticArchiveConfiguration.SavePersonalCatalogPath(_context, selected)
                MarkCatalogActivationFailed("The personal catalog location was saved in '" & configurationFile & "'. The selected catalog must finish loading before catalog actions resume.")
                _status.Text = "Personal catalog location saved in " & configurationFile & ". Loading the selected catalog."
            Catch ex As SemanticArchiveConfiguration.CatalogActivationException
                MarkCatalogActivationFailed(ex.Message)
                ReportError("Catalog location was saved, but activation failed. The displayed catalog and draft are retained; reload before continuing", ex)
                Return
            Catch ex As System.Exception
                ReportError("Catalog location could not be applied", ex)
                Return
            Finally
                SetBusy(False)
            End Try
            Await ReloadCatalogAsync(discardSourceInput:=True)
        End Function

        Private Sub MarkCatalogActivationFailed(message As System.String)
            _catalogActivationBlocked = True
            _catalogActivationMessage = If(message, "Configuration activation is incomplete.")
            _catalogLocation.Text = If(_store Is Nothing, "", _store.DirectoryPath)
            SetBusy(_busy)
        End Sub

        Private Sub MarkCatalogReloaded()
            _catalogActivationBlocked = False
            _catalogActivationMessage = ""
            SetBusy(_busy)
        End Sub

        Private Async Function RequestReloadCatalogAsync() As System.Threading.Tasks.Task
            If _closeRequested OrElse _busy OrElse Not ConfirmDiscardChanges(True) Then Return
            Await ReloadCatalogAsync(discardBackgroundEdits:=True, discardSourceInput:=True)
        End Function

        Private Async Function ReloadCatalogAsync(Optional selectArchiveId As String = Nothing,
                                                  Optional discardBackgroundEdits As System.Boolean = False,
                                                  Optional discardSourceInput As System.Boolean = False) As System.Threading.Tasks.Task
            If _closeRequested OrElse _busy Then Return
            Dim sourceInput As System.String = _newRootPath.Text
            Dim sourceArchiveId As System.String = If(_archive Is Nothing, "", _archive.ArchiveId)
            SetBusy(True)
            _loading = True
            Try
                If _catalogActivationBlocked Then
                    Dim activeContext As SharedContext.ISharedContext = _context
                    SharedMethods.InitializeConfig(activeContext, False, True)
                    If Not _context.INIloaded OrElse _context.GPTSetupError Then Throw New System.Configuration.ConfigurationErrorsException("The normal configuration loader did not complete successfully. The existing catalog and draft remain locked.")
                    If System.String.IsNullOrWhiteSpace(_context.INI_SemanticArchiveCatalogPathLocal) Then Throw New System.Configuration.ConfigurationErrorsException("The active configuration has no personal catalog location. Choose a catalog path before continuing.")
                End If
                Dim reloadBackground As System.Boolean = discardBackgroundEdits OrElse Not _backgroundDirty
                Dim permissions As SemanticArchivePermissionMaintenanceSettings.Controls = Nothing
                Dim globalEnabled As System.Boolean = False
                Dim globalWindow As System.String = ""
                If reloadBackground Then
                    Await System.Threading.Tasks.Task.Run(Sub()
                                                              SemanticArchiveBackgroundSettings.ReadInto(_context)
                                                              permissions = SemanticArchivePermissionMaintenanceSettings.ReadControls(_context)
                                                              globalEnabled = _context.INI_SemanticArchiveBackgroundIndexing
                                                              globalWindow = _context.INI_SemanticArchiveBackgroundIndexingWindow
                                                          End Sub)
                End If
                Dim catalogPath As System.String = _context.INI_SemanticArchiveCatalogPathLocal
                If System.String.IsNullOrWhiteSpace(catalogPath) Then
                    _catalog = Nothing
                    _store = Nothing
                    _archive = Nothing
                    _root = Nothing
                    _archiveDirty = False
                    _archives.Items.Clear()
                    _catalogLocation.Clear()
                    _newRootPath.Clear()
                    If reloadBackground Then ApplyLoadedBackgroundSettings(permissions, globalEnabled, globalWindow)
                    _status.Text = "No personal catalog is configured. Use Enter catalog path, Browse or Use recommended location above. Personal background settings can be saved independently."
                    ResetDiagnostics("Configuration key: SemanticArchiveCatalogPathLocal" & System.Environment.NewLine &
                        "The personal catalog directory contains redink-sa-catalog.json, private routing indexes and work queues. Original and shared derivative folders remain separate." & System.Environment.NewLine &
                        "Choose a private writable directory above. Changing location does not move or delete an existing catalog. Background indexing defaults to off.")
                    _tabs.SelectedTab = DirectCast(_diagnostics.Parent, System.Windows.Forms.TabPage)
                    Return
                End If
                Dim store As SemanticArchiveStore = Nothing
                Dim loaded As SemanticArchiveCatalog = Nothing
                Await System.Threading.Tasks.Task.Run(
                    Sub()
                        store = New SemanticArchiveStore(catalogPath)
                        loaded = store.LoadCatalog()
                    End Sub)
                _store = store
                _catalog = loaded
                _catalogLocation.Text = store.DirectoryPath
                If reloadBackground Then ApplyLoadedBackgroundSettings(permissions, globalEnabled, globalWindow)
                MarkCatalogReloaded()
                Dim requestedId = If(selectArchiveId, If(_archive Is Nothing, "", _archive.ArchiveId))
                _archives.Items.Clear()
                For Each definition In loaded.Archives
                    Dim index = _archives.Items.Add(New ArchiveItem(definition))
                    If System.String.Equals(definition.ArchiveId, requestedId, System.StringComparison.Ordinal) Then _archives.SelectedIndex = index
                Next
                If _archives.SelectedIndex < 0 AndAlso _archives.Items.Count > 0 Then _archives.SelectedIndex = 0
                _archiveDirty = False
                _newRootPath.Clear()
                _loading = False
                ArchiveSelected(Me, System.EventArgs.Empty)
                If Not discardSourceInput AndAlso _archive IsNot Nothing AndAlso System.String.Equals(_archive.ArchiveId, sourceArchiveId, System.StringComparison.Ordinal) Then _newRootPath.Text = sourceInput
                _status.Text = "Catalog loaded: " & store.CatalogPath & " — " & loaded.Archives.Count.ToString() & " archive(s)."
            Catch ex As System.Exception
                ReportError("Could not load archives; the previous catalog and draft are retained", ex)
            Finally
                _loading = False
                If _catalog Is Nothing Then
                    _archives.Items.Clear()
                    ArchiveSelected(Me, System.EventArgs.Empty)
                End If
                SetBusy(False)
            End Try
            If Not _catalogActivationBlocked Then Await RefreshDiagnosticsAsync()
        End Function

        Private Sub ApplyLoadedBackgroundSettings(permissions As SemanticArchivePermissionMaintenanceSettings.Controls,
                                                   enabled As System.Boolean, window As System.String)
            _permissionsEnabled.Checked = permissions.Enabled
            _permissionsWindow.Text = permissions.Window
            _globalEnabled.Checked = enabled
            _globalWindow.Text = window
            MarkBackgroundSaved()
        End Sub

        Private Async Sub ArchiveSelected(sender As Object, e As System.EventArgs)
            If _loading Then Return
            If _catalogActivationBlocked Then
                _loading = True
                _archives.SelectedIndex = _archiveSelectionIndex
                _loading = False
                Return
            End If
            If (_archiveDirty OrElse HasPendingSourcePath()) AndAlso Not ConfirmDiscardChanges(False) Then
                _loading = True
                _archives.SelectedIndex = _archiveSelectionIndex
                _loading = False
                Return
            End If
            Dim item = TryCast(_archives.SelectedItem, ArchiveItem)
            _archive = If(item Is Nothing, Nothing, SemanticArchiveMetadata.Clone(item.Definition))
            _archiveSelectionIndex = _archives.SelectedIndex
            _rootSelectionIndex = -1
            _archiveDirty = False
            _archiveIndexUnsupported = False
            ResetDocumentListing()
            _documents.Items.Clear()
            _selectedDocumentIds.Clear()
            _documentDetails.Clear()
            _documentPageStatus.Text = "Choose a status and Find documents. Needs attention lists incomplete, empty or unsuccessful processing. Source access unavailable is a separate view."
            _root = Nothing
            _loading = True
            _newRootPath.Clear()
            _roots.Items.Clear()
            If _archive IsNot Nothing Then
                _name.Text = _archive.Name
                _description.Text = _archive.Description
                _visible.Checked = If(SemanticArchiveLibrary.IsSubscriber(_archive), Not _archive.Library.OptOut, _archive.Enabled)
                _defaultScope.Checked = _catalog.DefaultArchiveIds.Contains(_archive.ArchiveId)
                _archiveBackground.Checked = _archive.BackgroundEnabled
                _archiveWindow.Text = _archive.BackgroundWindow
                _partialSearch.Checked = _archive.AllowPartialSearch
                _threshold.Value = System.Math.Max(_threshold.Minimum, System.Math.Min(_threshold.Maximum, CDec(_archive.SectionIndexThresholdBytes)))
                _children.Value = System.Math.Max(_children.Minimum, System.Math.Min(_children.Maximum, CDec(_archive.MaxChildrenPerNode)))
                _routingCharacters.Value = System.Math.Max(_routingCharacters.Minimum, System.Math.Min(_routingCharacters.Maximum, CDec(_archive.MaxRoutingCharacters)))
                _editedBudgets = CloneBudgets(_archive.RetrievalBudgets)
                _budgets.SelectedObject = _editedBudgets
                For Each binding In _archive.Roots
                    _roots.Items.Add(New RootItem(binding))
                Next
                If _roots.Items.Count > 0 Then _roots.SelectedIndex = 0
            Else
                _name.Clear()
                _description.Clear()
                _visible.Checked = False
                _defaultScope.Checked = False
                _archiveBackground.Checked = False
                _archiveWindow.Clear()
                _partialSearch.Checked = False
                _editedBudgets = Nothing
                _budgets.SelectedObject = Nothing
            End If
            _loading = False
            RootSelected(Me, System.EventArgs.Empty)
            SetBusy(_busy)
            If _archive IsNot Nothing AndAlso Not _busy AndAlso Not _closeRequested Then
                Await RefreshDiagnosticsAsync()
            End If
        End Sub

        Private Sub RootSelected(sender As Object, e As System.EventArgs)
            If _loading Then Return
            If _root IsNot Nothing Then
                Try
                    ApplyRootEditors()
                Catch ex As System.Exception
                    _loading = True
                    _roots.SelectedIndex = _rootSelectionIndex
                    _loading = False
                    ReportError("Correct the source options before switching roots", ex)
                    Return
                End Try
            End If
            Dim item = TryCast(_roots.SelectedItem, RootItem)
            _root = If(item Is Nothing, Nothing, item.Binding)
            _rootSelectionIndex = _roots.SelectedIndex
            _rootEditor.Enabled = _root IsNot Nothing AndAlso Not SemanticArchiveLibrary.IsSubscriber(_archive) AndAlso Not _busy
            _loading = True
            Try
                If _root Is Nothing Then
                    For Each input As System.Windows.Forms.TextBox In New System.Windows.Forms.TextBox() {_rootPath, _extensions, _exclusions, _sharedArtifactRoot, _shadowArtifactRoot, _scopeTags}
                        input.Clear()
                    Next
                    _recursive.Checked = False
                    _ocr.Checked = False
                    _ocrBatchPages.Value = SharedMethods.DEFAULT_SEMANTICARCHIVE_OCR_BATCH_PAGES
                    _fileTypeFilter.Checked = SharedMethods.DEFAULT_SEMANTICARCHIVE_FILE_TYPE_FILTER_ENABLED
                    _placementMode.SelectedIndex = -1
                    Return
                End If
                _rootPath.Text = _root.RootPath
                _recursive.Checked = _root.Recursive
                _ocr.Checked = _root.EnableOcr
                _ocrBatchPages.Value = System.Math.Max(CInt(_ocrBatchPages.Minimum), System.Math.Min(CInt(_ocrBatchPages.Maximum), _root.OcrBatchPages))
                _fileTypeFilter.Checked = _root.FileTypeFilterEnabled
                _extensions.Enabled = _fileTypeFilter.Checked
                _restoreFileTypes.Enabled = _fileTypeFilter.Checked
                _extensions.Text = System.String.Join("; ", SemanticArchiveStore.GetSourceFilterExtensions(_root))
                _exclusions.Text = System.String.Join(System.Environment.NewLine, _root.Exclusions)
                _placementMode.SelectedItem = _root.ArtifactPlacementMode
                If _placementMode.SelectedIndex < 0 Then _placementMode.SelectedItem = SharedMethods.DEFAULT_SEMANTICARCHIVE_ARTIFACT_PLACEMENT_MODE
                _sharedArtifactRoot.Text = _root.SharedArtifactRoot
                _shadowArtifactRoot.Text = _root.ShadowArtifactRoot
                _scopeTags.Text = System.String.Join("; ", _root.ScopeTags)
            Finally
                _loading = False
                UpdateSaveState()
            End Try
        End Sub

        Private Shared Function ParseLines(text As String, separators As Char()) As System.Collections.Generic.List(Of String)
            Dim result As New System.Collections.Generic.List(Of String)()
            Dim seen As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.OrdinalIgnoreCase)
            For Each part In If(text, "").Split(separators, System.StringSplitOptions.RemoveEmptyEntries)
                Dim value = part.Trim()
                If value.Length > 0 AndAlso seen.Add(value) Then result.Add(value)
            Next
            Return result
        End Function

        Private Shared Function CloneBudgets(source As SemanticArchiveRetrievalBudgets) As SemanticArchiveRetrievalBudgets
            If source Is Nothing Then Return New SemanticArchiveRetrievalBudgets()
            ' Preserve future catalog fields as well as every current budget while editing.
            Return SemanticArchiveMetadata.Clone(source)
        End Function

        Private Sub ApplyEditors()
            If _archive Is Nothing Then Throw New System.InvalidOperationException("Select or create an archive first.")
            If SemanticArchiveLibrary.IsSubscriber(_archive) Then
                _archive.Library.OptOut = Not _visible.Checked
                _archive.Enabled = _visible.Checked AndAlso _archive.Library.State = "available"
                Return ' Publisher-managed definition; only local visibility/default selection is editable.
            End If
            If System.String.IsNullOrWhiteSpace(_name.Text) Then Throw New System.ArgumentException("Enter an archive name.")
            If Not BackgroundProcessingWindow.IsValid(_archiveWindow.Text) Then Throw New System.ArgumentException("The archive processing window is invalid.")
            _archive.Name = _name.Text.Trim()
            _archive.Description = _description.Text.Trim()
            _archive.Enabled = _visible.Checked
            _archive.BackgroundEnabled = _archiveBackground.Checked
            _archive.BackgroundWindow = _archiveWindow.Text.Trim()
            _archive.AllowPartialSearch = _partialSearch.Checked
            _archive.SectionIndexThresholdBytes = CLng(_threshold.Value)
            _archive.MaxChildrenPerNode = CInt(_children.Value)
            _archive.MaxRoutingCharacters = CInt(_routingCharacters.Value)
            Dim budgets = _editedBudgets
            If budgets Is Nothing OrElse budgets.MaxNodesVisited < 1 OrElse budgets.MaxModelCalls < 1 OrElse
               budgets.MaxElapsedSeconds < 1 OrElse budgets.MaxCandidateFiles < 1 OrElse budgets.MaxEvidenceBytes < 1 OrElse
               budgets.InitialBranches < 1 OrElse budgets.MaxExactLookupDocuments < 0 OrElse budgets.MaxSectionCandidates < 1 OrElse
               budgets.MaxPromptCharacters < 1024 OrElse budgets.MaxRequestTokens < 1 OrElse budgets.MaxRequestTokens > 262144 OrElse
               budgets.MaxLiteralScanBytes < 1 OrElse budgets.MaxLiteralScanBytes > 67108864 OrElse
               budgets.MaxLiteralScanDocuments < 1 OrElse budgets.MaxLiteralScanDocuments > 1024 Then
                Throw New System.ArgumentException("Retrieval budgets must be positive (MaxExactLookupDocuments can be zero); MaxPromptCharacters must be at least 1024, MaxRequestTokens must be between 1 and 262144, MaxLiteralScanBytes must be between 1 and 67108864, and MaxLiteralScanDocuments must be between 1 and 1024.")
            End If
            _archive.RetrievalBudgets = CloneBudgets(budgets)
            ApplyRootEditors()
        End Sub

        Private Sub ApplyRootEditors()
            If _root Is Nothing OrElse SemanticArchiveLibrary.IsSubscriber(_archive) Then Return
            _root.Recursive = _recursive.Checked
            _root.EnableOcr = _ocr.Checked
            _root.OcrBatchPages = System.Convert.ToInt32(_ocrBatchPages.Value, System.Globalization.CultureInfo.InvariantCulture)
            _root.FileTypeFilterEnabled = _fileTypeFilter.Checked
            _root.SupportedExtensions = ParseLines(_extensions.Text, New Char() {";"c, ","c, " "c})
            For index As Integer = 0 To _root.SupportedExtensions.Count - 1
                Dim extension = _root.SupportedExtensions(index)
                If Not extension.StartsWith(".", System.StringComparison.Ordinal) Then extension = "." & extension
                If extension.IndexOfAny(New Char() {"*"c, "?"c, "/"c, "\"c, ":"c}) >= 0 Then Throw New System.ArgumentException("File types must be extensions such as .pdf; .docx; .txt.")
                _root.SupportedExtensions(index) = extension.ToLowerInvariant()
            Next
            _root.Exclusions = ParseLines(_exclusions.Text, New Char() {Microsoft.VisualBasic.ChrW(13), Microsoft.VisualBasic.ChrW(10)})
            _root.ArtifactPlacementMode = If(TryCast(_placementMode.SelectedItem, System.String), SharedMethods.DEFAULT_SEMANTICARCHIVE_ARTIFACT_PLACEMENT_MODE)
            _root.SharedArtifactRoot = _sharedArtifactRoot.Text.Trim()
            _root.ShadowArtifactRoot = _shadowArtifactRoot.Text.Trim()
            _root.ScopeTags = ParseLines(_scopeTags.Text, New Char() {";"c, ","c})
        End Sub

        Private Async Function SaveArchiveAsync() As System.Threading.Tasks.Task(Of Boolean)
            If _closeRequested OrElse _busy OrElse _catalogActivationBlocked OrElse _catalog Is Nothing OrElse _store Is Nothing Then Return False
            SetBusy(True)
            Try
                ApplyEditors()
                Dim pending As SemanticArchiveCatalog = SemanticArchiveMetadata.Clone(_catalog)
                Dim archiveIndex As System.Int32 = pending.Archives.FindIndex(Function(definition) System.String.Equals(definition.ArchiveId, _archive.ArchiveId, System.StringComparison.Ordinal))
                If archiveIndex < 0 Then Throw New System.InvalidOperationException("The archive is no longer in the catalog. Reload before saving.")
                pending.Archives(archiveIndex) = SemanticArchiveMetadata.Clone(_archive)
                pending.DefaultArchiveIds.RemoveAll(Function(id) System.String.Equals(id, _archive.ArchiveId, System.StringComparison.Ordinal))
                If _defaultScope.Checked Then pending.DefaultArchiveIds.Add(_archive.ArchiveId)
                Dim revision As System.Int64 = _catalog.Revision
                Await System.Threading.Tasks.Task.Run(Sub() _store.SaveCatalog(pending, revision))
                _catalog = pending
                Dim index = _archives.SelectedIndex
                If index >= 0 Then
                    _loading = True
                    _archives.Items(index) = New ArchiveItem(SemanticArchiveMetadata.Clone(_archive))
                    _archives.SelectedIndex = index
                    _loading = False
                End If
                MarkArchiveSaved()
                RetrievalSourceDiscovery.RequestRefresh(_context, True)
                _status.Text = "Saved archive " & _archive.Name & " [" & _archive.ArchiveId & "]."
                Return True
            Catch ex As System.Exception
                ReportError("Archive save could not be completed or confirmed; edits are retained", ex)
                Return False
            Finally
                _loading = False
                SetBusy(False)
            End Try
        End Function

        Private Async Function CreateArchiveAsync() As System.Threading.Tasks.Task
            If _closeRequested OrElse _busy OrElse _catalogActivationBlocked OrElse _catalog Is Nothing Then Return
            If _archiveDirty AndAlso Not Await SaveArchiveAsync() Then Return
            If _closeRequested Then Return
            If HasPendingSourcePath() AndAlso Not ConfirmDiscardChanges(False) Then Return
            Dim name As System.String
            Using dialogOwner As System.IDisposable = SharedMethods.PushDialogOwner(Me)
                name = SharedMethods.ShowCustomInputBox("Name of the new semantic archive:", "Create Semantic Archive", True)
            End Using
            If System.String.IsNullOrWhiteSpace(name) Then Return
            Dim definition As New SemanticArchiveDefinition() With {.ArchiveId = SemanticArchiveIdentity.NewId(), .Name = name.Trim()}
            SetBusy(True)
            Try
                Dim pending As SemanticArchiveCatalog = SemanticArchiveMetadata.Clone(_catalog)
                Dim revision As System.Int64 = pending.Revision
                pending.Archives.Add(definition)
                Await System.Threading.Tasks.Task.Run(Sub() _store.SaveCatalog(pending, revision))
                _catalog = pending
            Catch ex As System.Exception
                ReportError("Archive could not be created", ex)
                Return
            Finally
                SetBusy(False)
            End Try
            Await ReloadCatalogAsync(definition.ArchiveId)
        End Function

        Private Async Function RemoveArchiveAsync() As System.Threading.Tasks.Task
            If _closeRequested OrElse _busy OrElse _catalogActivationBlocked OrElse _archive Is Nothing Then Return
            If Not ConfirmDiscardChanges(False) Then Return
            Dim withdrawRequired As System.Boolean = Not SemanticArchiveLibrary.IsSubscriber(_archive) AndAlso
                _archive.Library IsNot Nothing AndAlso _archive.Library.State = "available"
            Dim message As System.String = "Unregister '" & _archive.Name & "'? Original files and generated artifacts are retained."
            If withdrawRequired Then message = "Withdraw the published library definition of '" & _archive.Name & "' and then unregister this local archive? Library readers will no longer receive the definition. Original files and generated artifacts are retained."
            If Not ConfirmConsoleAction(message, If(withdrawRequired, "Withdraw and unregister", "Unregister"), "Cancel") Then Return
            Dim id As System.String = _archive.ArchiveId
            Dim cancellation As New System.Threading.CancellationTokenSource()
            _operationCancellation = cancellation
            SetBusy(True)
            Dim removed As System.Boolean = False
            Dim failure As System.Exception = Nothing
            Dim cancelled As System.Boolean = False
            Try
                Await System.Threading.Tasks.Task.Run(
                    Sub()
                        cancellation.Token.ThrowIfCancellationRequested()
                        If withdrawRequired Then
                            If Not SemanticArchiveLibrary.IsConfigured(_context) Then Throw New System.InvalidOperationException("Restore the archive's configured library connection first. Its published definition must be withdrawn before local unregistration.")
                            SemanticArchiveLibrary.Publish(_context, id, True, cancellation.Token)
                        End If
                        ' Store keeps its publication/revision/subscriber guards. No index is opened.
                        cancellation.Token.ThrowIfCancellationRequested()
                        _store.RemoveArchive(id)
                        removed = True
                    End Sub, cancellation.Token)
            Catch ex As System.OperationCanceledException
                cancelled = True
            Catch ex As System.Exception
                failure = ex
            Finally
                _operationCancellation = Nothing
                cancellation.Dispose()
                SetBusy(False)
            End Try
            If removed Then _archiveDirty = False
            If Not _closeRequested Then Await ReloadCatalogAsync()
            If failure IsNot Nothing Then ReportError("Archive could not be unregistered. A failed withdrawal leaves the local registration in place; reload to verify a partially completed operation", failure)
            If cancelled AndAlso Not IsDisposed Then _status.Text = "Unregistration stopped. A publication already withdrawn stays withdrawn; reload before continuing."
        End Function

        Private Async Function AddRootAsync() As System.Threading.Tasks.Task
            If SemanticArchiveLibrary.IsSubscriber(_archive) Then Return
            If _closeRequested OrElse _busy OrElse _catalogActivationBlocked OrElse _archive Is Nothing Then Return
            If Not Await SaveArchiveAsync() Then Return
            If _closeRequested Then Return
            Dim selected As String = Nothing
            Using picker As New System.Windows.Forms.FolderBrowserDialog() With {.Description = "Select an original source directory to add. The source entry above also accepts local and UNC directories.", .ShowNewFolderButton = False}
                Using dialogOwner As System.IDisposable = SharedMethods.PushDialogOwner(Me)
                    Dim owner As System.Windows.Forms.IWin32Window = SharedMethods.ResolveSameThreadDialogOwner()
                    Dim result As System.Windows.Forms.DialogResult = If(owner Is Nothing, picker.ShowDialog(), picker.ShowDialog(owner))
                    If result <> System.Windows.Forms.DialogResult.OK Then Return
                    selected = picker.SelectedPath
                End Using
            End Using
            _newRootPath.Text = selected
            Await RegisterSourceRootAsync(selected)
        End Function

        Private Async Function AddRootPathAsync() As System.Threading.Tasks.Task
            If _closeRequested OrElse _busy OrElse _catalogActivationBlocked OrElse _archive Is Nothing Then Return
            Await RegisterSourceRootAsync(_newRootPath.Text)
        End Function

        Private Shared Function NormalizeSourceRootInput(value As System.String) As System.String
            Dim selected As System.String = If(value, "").Trim()
            If selected.Length >= 2 AndAlso selected(0) = Microsoft.VisualBasic.ChrW(34) AndAlso selected(selected.Length - 1) = Microsoft.VisualBasic.ChrW(34) Then selected = selected.Substring(1, selected.Length - 2).Trim()
            selected = SharedMethods.ExpandEnvironmentVariables(selected).Replace("/"c, "\"c)
            Dim localAbsolute As System.Boolean = selected.Length >= 3 AndAlso
                ((selected(0) >= "A"c AndAlso selected(0) <= "Z"c) OrElse (selected(0) >= "a"c AndAlso selected(0) <= "z"c)) AndAlso
                selected(1) = ":"c AndAlso selected(2) = "\"c
            Dim uncAbsolute As System.Boolean = False
            If selected.StartsWith("\\", System.StringComparison.Ordinal) AndAlso Not selected.StartsWith("\\?\", System.StringComparison.Ordinal) AndAlso Not selected.StartsWith("\\.\", System.StringComparison.Ordinal) Then
                Dim parts As System.String() = selected.Substring(2).Split(New System.Char() {"\"c}, System.StringSplitOptions.None)
                uncAbsolute = parts.Length >= 2 AndAlso parts(0).Length > 0 AndAlso parts(1).Length > 0 AndAlso
                    parts(0) <> "." AndAlso parts(0) <> ".." AndAlso parts(1) <> "." AndAlso parts(1) <> ".."
            End If
            If Not localAbsolute AndAlso Not uncAbsolute Then Throw New System.ArgumentException("Enter an absolute source directory, for example C:\Documents or \\server\share\Archive. Relative and drive-relative paths are not accepted.")
            ' Pure path validation: do not contact a local disk or network share on the UI thread.
            ' The existing guard retains Windows source path and component length limits.
            Return SemanticArchivePathGuard.RequireWindowsSourcePath(selected)
        End Function

        Private Async Function RegisterSourceRootAsync(selected As String) As System.Threading.Tasks.Task
            If SemanticArchiveLibrary.IsSubscriber(_archive) Then Return
            If _closeRequested OrElse _busy OrElse _catalogActivationBlocked OrElse _archive Is Nothing Then Return
            Dim sourcePath As System.String
            Try
                sourcePath = NormalizeSourceRootInput(selected)
            Catch ex As System.Exception
                ReportError("Source folder could not be added; the entered path is retained", ex)
                Return
            End Try
            If _archiveDirty AndAlso Not Await SaveArchiveAsync() Then Return
            If _closeRequested Then Return
            For Each existing As SemanticArchiveSourceBinding In _archive.Roots
                If System.String.Equals(SemanticArchivePathGuard.CanonicalPath(existing.RootPath), sourcePath, System.StringComparison.OrdinalIgnoreCase) Then
                    If SelectSourceRoot(existing.BindingId) Then
                        _newRootPath.Clear()
                        _status.Text = "This source folder is already registered. Its existing source settings are selected."
                    End If
                    Return
                End If
            Next
            Dim wasDirty As System.Boolean = _archiveDirty
            Dim binding As New SemanticArchiveSourceBinding() With {.BindingId = SemanticArchiveIdentity.NewId(), .RootPath = sourcePath, .Recursive = SharedMethods.DEFAULT_SEMANTICARCHIVE_RECURSIVE}
            _archive.Roots.Add(binding)
            _archiveDirty = True
            ' SaveCatalog retains the catalog revision, source scope and output checks for
            ' both typed and browsed paths. Source directories and their ACLs are unchanged.
            If Await SaveArchiveAsync() Then
                Dim archiveId As System.String = _archive.ArchiveId
                _newRootPath.Clear()
                Await ReloadCatalogAsync(archiveId)
                ' Reload and diagnostics own the operation result, including any errors.
                ' Selecting the saved root must not replace that result with a success notice.
                SelectSourceRoot(binding.BindingId)
            Else
                _archive.Roots.Remove(binding)
                _archiveDirty = wasDirty
                UpdateSaveState()
            End If
        End Function

        Private Function SelectSourceRoot(bindingId As System.String) As System.Boolean
            For index As System.Int32 = 0 To _roots.Items.Count - 1
                Dim item As RootItem = TryCast(_roots.Items(index), RootItem)
                If item IsNot Nothing AndAlso System.String.Equals(item.Binding.BindingId, bindingId, System.StringComparison.Ordinal) Then
                    _roots.SelectedIndex = index
                    Return _root IsNot Nothing AndAlso System.String.Equals(_root.BindingId, bindingId, System.StringComparison.Ordinal)
                End If
            Next
            Return False
        End Function

        Private Async Function RemoveRootAsync() As System.Threading.Tasks.Task
            If SemanticArchiveLibrary.IsSubscriber(_archive) Then Return
            If _closeRequested OrElse _busy OrElse _catalogActivationBlocked OrElse _archive Is Nothing OrElse _root Is Nothing Then Return
            If _archiveDirty AndAlso Not Await SaveArchiveAsync() Then Return
            If _closeRequested Then Return
            If Not ConfirmConsoleAction("Remove this root from the archive? Original files are retained.", "Remove root", "Cancel") Then Return
            Dim archiveId = _archive.ArchiveId
            Dim bindingId = _root.BindingId
            SetBusy(True)
            Try
                Await System.Threading.Tasks.Task.Run(Sub() _store.RemoveRoot(archiveId, bindingId))
            Catch ex As System.Exception
                ReportError("Root could not be removed", ex)
                Return
            Finally
                SetBusy(False)
            End Try
            Await ReloadCatalogAsync(archiveId)
        End Function

        Private Async Function SaveBackgroundAsync() As System.Threading.Tasks.Task
            If _closeRequested OrElse _busy Then Return
            Dim enabled = _globalEnabled.Checked
            Dim window = _globalWindow.Text
            Dim rightsEnabled = _permissionsEnabled.Checked
            Dim rightsWindow = _permissionsWindow.Text
            SetBusy(True)
            Try
                If Not BackgroundProcessingWindow.IsValid(window) OrElse Not BackgroundProcessingWindow.IsValid(rightsWindow) Then Throw New System.ArgumentException("One of the background processing windows is invalid.")
                Await System.Threading.Tasks.Task.Run(Sub()
                                                          SemanticArchiveBackgroundSettings.Save(_context, enabled, window)
                                                          SemanticArchivePermissionMaintenanceSettings.Save(rightsEnabled, rightsWindow, _context)
                                                      End Sub)
                MarkBackgroundSaved()
                _status.Text = "Semantic Archive background settings saved. Participating hosts synchronize controls at the next activity tick."
            Catch ex As System.Exception
                ReportError("Could not save all background settings; some values may already be saved. Correct the error and save again", ex)
            Finally
                SetBusy(False)
            End Try
        End Function

        Private Sub AcquireAutomaticPause()
            If _automaticPause Is Nothing Then _automaticPause = SemanticArchiveMaintenanceProvider.PauseAutomatic(_context)
            If _permissionsPause Is Nothing Then _permissionsPause = SemanticArchivePermissionMaintenanceProvider.PauseAutomatic(_context)
        End Sub

        Private Sub ReleaseAutomaticPause()
            If _automaticPause IsNot Nothing Then _automaticPause.Dispose()
            If _permissionsPause IsNot Nothing Then _permissionsPause.Dispose()
            _automaticPause = Nothing
            _permissionsPause = Nothing
        End Sub

        Private Function CanResumeSelectedArchive() As Boolean
            Return _archive IsNot Nothing AndAlso _resumeOperationId.Length > 0 AndAlso
                System.String.Equals(_resumeArchiveId, _archive.ArchiveId, System.StringComparison.Ordinal)
        End Function

        Private Async Function PauseResumeAsync() As System.Threading.Tasks.Task
            If _closeRequested OrElse _catalogActivationBlocked Then Return
            If Not _busy AndAlso (_paused OrElse CanResumeSelectedArchive()) Then
                _paused = False
                _pause.Text = "Pause"
                If CanResumeSelectedArchive() Then
                    Await RunBuildAsync(_resumeCommand, _resumeDocumentIds IsNot Nothing, True)
                Else
                    ReleaseAutomaticPause()
                    _status.Text = "Independent automatic content and permission eligibility resumed."
                End If
                Return
            End If
            _paused = True
            AcquireAutomaticPause()
            If _operationCancellation IsNot Nothing Then _operationCancellation.Cancel()
            _pause.Text = "Resume"
            _status.Text = "Content and permission maintenance pause at a safe checkpoint. Published search remains available."
        End Function

        Private Shared Function UserOperationText(command As MaintenanceCommand, selectedCount As System.Nullable(Of System.Int32)) As System.String
            Dim scopeText As System.String = If(selectedCount.HasValue, selectedCount.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) & " selected document(s)", "all documents in this archive")
            Select Case command
                Case MaintenanceCommand.RefreshContent
                    Return "Refreshing " & scopeText & ". Unchanged documents and valid extracted text are reused; only necessary work is performed."
                Case MaintenanceCommand.SemanticReindex
                    Return "Rebuilding the semantic index for " & scopeText & ". Existing validated extracted text is reused; this operation never runs extraction or OCR."
                Case MaintenanceCommand.Extract
                    Return "Re-extracting text for " & scopeText & ". PDF text layers are checked first; OCR runs only for pages that need it, using the configured batch size."
                Case MaintenanceCommand.Permissions
                    Return "Checking generated-file permissions for " & scopeText & ". Original document permissions are not changed."
                Case MaintenanceCommand.Retry
                    Return "Retrying unfinished or coverage-excluded work for " & scopeText & ". Missing/empty/incomplete/unknown extraction is refreshed; verified complete extracts are reused."
                Case Else
                    Return "Processing " & scopeText & "."
            End Select
        End Function

        Private Shared Function UserProgressStage(stage As System.String) As System.String
            Select Case If(stage, "")
                Case "scanning" : Return "Checking source folders"
                Case "processing" : Return "Processing document"
                Case "hierarchy" : Return "Updating search structure"
                Case "publishing" : Return "Activating updated archive"
                Case "permissions" : Return "Checking permissions"
                Case Else : Return "Working"
            End Select
        End Function

        Private Async Function RunBuildAsync(command As MaintenanceCommand, selectedOnly As Boolean,
                                             Optional resumeExistingOperation As System.Boolean = False) As System.Threading.Tasks.Task
            If _closeRequested OrElse _busy OrElse _catalogActivationBlocked OrElse _archive Is Nothing OrElse _store Is Nothing Then Return
            If Not resumeExistingOperation Then
                Dim selected As System.Collections.Generic.List(Of String) = Nothing
                Try
                    If selectedOnly Then selected = ReadSelectedDocumentIds()
                Catch ex As System.Exception
                    ReportError("Select document IDs before targeted maintenance", ex)
                    Return
                End Try
                If Not Await SaveArchiveAsync() OrElse _closeRequested Then Return
                If SemanticArchiveLibrary.IsSubscriber(_archive) Then
                    Dim currentId As System.String = _archive.ArchiveId
                    Dim synchronizationCancellation As New System.Threading.CancellationTokenSource()
                    _operationCancellation = synchronizationCancellation
                    SetBusy(True)
                    Dim synchronized As System.Boolean = False
                    Try
                        _status.Text = "Checking the subscribed archive definition..."
                        Await SemanticArchiveLibrary.SynchronizeAsync(_context, True, synchronizationCancellation.Token)
                        synchronized = True
                    Catch ex As System.OperationCanceledException
                        _status.Text = "Library synchronization cancelled before content processing."
                    Catch ex As System.Exception
                        ReportError("The subscribed archive definition could not be refreshed", ex)
                    Finally
                        _operationCancellation = Nothing
                        synchronizationCancellation.Dispose()
                        SetBusy(False)
                    End Try
                    If Not synchronized OrElse _closeRequested Then Return
                    Await ReloadCatalogAsync(currentId)
                    If _archive Is Nothing OrElse _archive.ArchiveId <> currentId OrElse Not _archive.Enabled Then Return
                End If
                _resumeArchiveId = _archive.ArchiveId
                _resumeOperationId = SemanticArchiveIdentity.NewId()
                _resumeCommand = command
                _resumeDocumentIds = selected
            End If
            If Not System.String.Equals(_resumeArchiveId, _archive.ArchiveId, System.StringComparison.Ordinal) OrElse _resumeOperationId.Length = 0 Then Return
            Dim archiveId = _resumeArchiveId
            Dim operationId = _resumeOperationId
            command = _resumeCommand
            Dim selectedIds = If(_resumeDocumentIds Is Nothing, Nothing, New System.Collections.Generic.List(Of String)(_resumeDocumentIds))
            _paused = False
            _pause.Text = "Pause"
            AcquireAutomaticPause()
            Dim cancellation As New System.Threading.CancellationTokenSource()
            _operationCancellation = cancellation
            SetBusy(True)
            ResetDiagnostics()
            _tabs.SelectedTab = DirectCast(_diagnostics.Parent, System.Windows.Forms.TabPage)
            Dim includeTechnicalDiagnostics As System.Boolean = _showTechnicalDiagnostics.Checked
            If includeTechnicalDiagnostics Then
                AppendDiagnostic("Operation " & operationId & ": " & command.ToString() & "; scope " & If(selectedIds Is Nothing, "all documents in this archive", selectedIds.Count.ToString() & " explicitly selected document ID(s)") & ".")
            Else
                AppendDiagnostic(UserOperationText(command, If(selectedIds Is Nothing, New System.Nullable(Of System.Int32)(), New System.Nullable(Of System.Int32)(selectedIds.Count))))
                AppendDiagnostic("The operation runs in bounded batches and continues automatically until the current work queue is finished or needs attention.")
            End If
            Try
                Dim builder As New SemanticArchiveBuilder(_context, _store)
                Dim progress As New System.Progress(Of SemanticArchiveBuildProgress)(AddressOf ShowProgress)
                Dim totalProcessed As Integer = 0
                Dim totalReused As Integer = 0
                Dim totalFailed As Integer = 0
                Dim totalCoverageExcluded As Integer = 0
                Dim totalRepaired As Integer = 0
                Dim totalQuarantined As Integer = 0
                Do
                    cancellation.Token.ThrowIfCancellationRequested()
                    Dim options As New SemanticArchiveBuildOptions() With {
                        .OperationId = operationId,
                        .SelectedDocumentIds = If(selectedIds Is Nothing, Nothing, New System.Collections.Generic.List(Of String)(selectedIds)),
                        .RebuildSemanticMetadata = command = MaintenanceCommand.SemanticReindex,
                        .IndexOnlyRebuild = command = MaintenanceCommand.SemanticReindex,
                        .ForceReextract = command = MaintenanceCommand.Extract,
                        .RetryFailures = command = MaintenanceCommand.Retry,
                        .ReconcilePermissionsOnly = command = MaintenanceCommand.Permissions,
                        .ForceScan = True, .IsBackground = False, .MaximumFilesPerBatch = SharedMethods.DEFAULT_SEMANTICARCHIVE_MAXIMUM_FILES_PER_BATCH,
                        .MaxDiscoveryEntries = SharedMethods.DEFAULT_SEMANTICARCHIVE_MAX_DISCOVERY_ENTRIES, .MaxDiscoverySeconds = SharedMethods.DEFAULT_SEMANTICARCHIVE_MAX_DISCOVERY_SECONDS,
                        .HostReaderDispatcher = AddressOf DispatchHostReaderAsync
                    }
                    Dim result = Await builder.BuildAsync(archiveId, options, progress, cancellation.Token)
                    totalProcessed += result.ProcessedFiles
                    totalReused += result.ReusedFiles
                    totalFailed += result.FailedFiles
                    totalCoverageExcluded += result.CoverageExcludedFiles
                    totalRepaired += result.PermissionArtifactsRepaired
                    totalQuarantined += result.PermissionArtifactsQuarantined
                    Dim visibleDiagnostics As System.Collections.Generic.List(Of System.String) = Await System.Threading.Tasks.Task.Run(
                        Function() SemanticArchiveBuilder.PresentDiagnostics(_store, archiveId, result.Diagnostics, includeTechnicalDiagnostics))
                    For Each message As System.String In visibleDiagnostics
                        AppendDiagnostic(message)
                    Next
                    If result.SelectionRequired Then Throw New System.InvalidOperationException("The targeted command has no selected document IDs; no archive-wide work was started.")
                    If includeTechnicalDiagnostics Then
                        AppendDiagnostic("Processed: " & totalProcessed.ToString() & "; reused: " & totalReused.ToString() &
                            "; failed: " & totalFailed.ToString() & "; coverage excluded: " & totalCoverageExcluded.ToString() &
                            "; pending: " & result.PendingFiles.ToString() & "; deferred: " & result.DeferredFiles.ToString() &
                            "; permissions repaired/quarantined: " & totalRepaired.ToString() & "/" & totalQuarantined.ToString() &
                            "; discovery pending: " & result.DiscoveryPending.ToString() & "; published: " & result.Published.ToString() & ".")
                        If result.Published Then AppendDiagnostic("Validated generation activated: " & result.GenerationId)
                    End If
                    If result.Inventory IsNot Nothing Then
                        If includeTechnicalDiagnostics Then
                            AppendDiagnostic(result.Inventory.ToTechnicalDiagnosticText())
                        End If
                        _status.Text = "Searchable: " & result.Inventory.SearchableDocuments.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                            "/" & result.Inventory.CurrentSources.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                            "; remaining work: " & result.PendingFiles.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                            If(result.DeferredFiles > 0, "; deferred: " & result.DeferredFiles.ToString(System.Globalization.CultureInfo.InvariantCulture), "")
                    Else
                        _status.Text = "Processed: " & totalProcessed.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                            "; remaining work: " & result.PendingFiles.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    End If
                    If result.Cancelled Then Throw New System.OperationCanceledException(cancellation.Token)
                    If result.PermissionsDeferred Then
                        AppendDiagnostic("Some permission work is deferred or cannot be repaired with the current rights. Unrelated discovery continues; source-specific retries remain durable and this operation remains resumable.")
                        If Not result.DiscoveryPending OrElse result.DiscoveryEntriesInspected = 0 Then Exit Do
                    End If
                    Dim pending As Boolean = If(command = MaintenanceCommand.Permissions,
                        result.PermissionsPending OrElse result.DiscoveryPending,
                        result.PendingFiles > 0 OrElse result.DiscoveryPending)
                    If Not pending Then
                        If Not includeTechnicalDiagnostics Then
                            AppendDiagnostic("Completed. " & totalProcessed.ToString(System.Globalization.CultureInfo.InvariantCulture) & " document job(s) processed; " &
                                totalReused.ToString(System.Globalization.CultureInfo.InvariantCulture) & " reused; " & totalFailed.ToString(System.Globalization.CultureInfo.InvariantCulture) & " failed; " &
                                totalCoverageExcluded.ToString(System.Globalization.CultureInfo.InvariantCulture) & " coverage-excluded result(s) recorded during this operation.")
                            If result.Inventory IsNot Nothing Then AppendDiagnostic(result.Inventory.ToDiagnosticText())
                        End If
                        _resumeArchiveId = ""
                        _resumeOperationId = ""
                        _resumeDocumentIds = Nothing
                        Exit Do
                    ElseIf Not includeTechnicalDiagnostics Then
                        AppendDiagnostic("Continuing the same operation: " & totalProcessed.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                            " document job(s) completed; " & result.PendingFiles.ToString(System.Globalization.CultureInfo.InvariantCulture) & " remain queued.")
                    End If
                    If result.ProcessedFiles + result.ReusedFiles + result.DiscoveryEntriesInspected +
                       result.PermissionSourcesChecked + result.PermissionArtifactsChecked = 0 Then
                        AppendDiagnostic("No progress in this bounded batch. Work remains checkpointed; inspect diagnostics and retry or resume.")
                        Exit Do
                    End If
                Loop
            Catch ex As System.OperationCanceledException
                _paused = True
                _pause.Text = "Resume"
                _status.Text = "Paused at a checkpoint. Resume retains this operation and its exact document selection."
            Catch ex As System.Exception
                ReportError("Archive maintenance failed", ex)
            Finally
                _operationCancellation = Nothing
                cancellation.Dispose()
                If Not _paused Then ReleaseAutomaticPause()
                SetBusy(False)
                If Not IsDisposed Then
                    _progress.Style = System.Windows.Forms.ProgressBarStyle.Continuous
                    _progress.Value = 0
                End If
            End Try
        End Function

        ''' <summary>Marshals only the existing Office-dependent reader to this console's owning STA.</summary>
        Private Async Function DispatchHostReaderAsync(reader As System.Func(Of String), cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of String)
            If reader Is Nothing Then Throw New System.ArgumentNullException(NameOf(reader))
            cancellationToken.ThrowIfCancellationRequested()
            Dim completion As New System.Threading.Tasks.TaskCompletionSource(Of String)(System.Threading.Tasks.TaskCreationOptions.RunContinuationsAsynchronously)
            Dim state As Integer = 0 ' 0=queued, 1=executing on the host STA, 2=cancelled before start
            Using registration = cancellationToken.Register(
                Sub()
                    If System.Threading.Interlocked.CompareExchange(state, 2, 0) = 0 Then completion.TrySetCanceled()
                End Sub)
                Try
                    If IsDisposed OrElse Not IsHandleCreated Then Throw New System.ObjectDisposedException(NameOf(SemanticArchiveForm))
                    BeginInvoke(New System.Windows.Forms.MethodInvoker(
                        Sub()
                            If System.Threading.Interlocked.CompareExchange(state, 1, 0) <> 0 Then Return
                            Try
                                cancellationToken.ThrowIfCancellationRequested()
                                Using dialogOwner As System.IDisposable = SharedMethods.PushDialogOwner(Me)
                                    Using interactionScope As SharedMethods.HeadlessExecutionScope = SharedMethods.BeginHeadlessExecution()
                                        Dim text As System.String = reader.Invoke()
                                        interactionScope.ThrowIfInteractionRequested()
                                        cancellationToken.ThrowIfCancellationRequested()
                                        completion.TrySetResult(text)
                                    End Using
                                End Using
                            Catch ex As System.Exception
                                completion.TrySetException(ex)
                            End Try
                        End Sub))
                Catch ex As System.Exception
                    completion.TrySetException(ex)
                End Try
                Return Await completion.Task.ConfigureAwait(False)
            End Using
        End Function

        Private Sub ShowProgress(progress As SemanticArchiveBuildProgress)
            If IsDisposed OrElse progress Is Nothing Then Return
            Dim remaining As System.Int32 = System.Math.Max(0, progress.PendingFiles - progress.CompletedFiles)
            Dim sourceName As System.String = If(System.String.IsNullOrWhiteSpace(progress.SourcePath), "", " — " & System.IO.Path.GetFileName(progress.SourcePath))
            _status.Text = UserProgressStage(progress.Stage) & sourceName & " (" & progress.CompletedFiles.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                " completed" & If(progress.PendingFiles > 0, "; about " & remaining.ToString(System.Globalization.CultureInfo.InvariantCulture) & " remaining in this batch", "") & ")"
            _progress.Style = System.Windows.Forms.ProgressBarStyle.Marquee
            If System.String.Equals(progress.Stage, "failed", System.StringComparison.Ordinal) OrElse
               System.String.Equals(progress.Stage, "pending_host", System.StringComparison.Ordinal) OrElse
               System.String.Equals(progress.Stage, "coverage_excluded", System.StringComparison.Ordinal) OrElse
               System.String.Equals(progress.Stage, "shared_claim_deferred", System.StringComparison.Ordinal) OrElse
               System.String.Equals(progress.Stage, "diagnostic_warning", System.StringComparison.Ordinal) Then AppendDiagnostic(progress.Message)
        End Sub

        Private Async Function RefreshDiagnosticsAsync() As System.Threading.Tasks.Task
            If _closeRequested OrElse _busy OrElse _catalogActivationBlocked OrElse _store Is Nothing OrElse _archive Is Nothing Then Return
            Dim archiveId = _archive.ArchiveId
            Dim includeTechnical As System.Boolean = _showTechnicalDiagnostics.Checked
            SetBusy(True)
            Try
                Dim publishedDiagnostics As New System.Collections.Generic.List(Of System.String)()
                Dim queueDiagnostics As System.Collections.Generic.List(Of String) = Nothing
                Dim manifest As SemanticArchiveGenerationManifest = Nothing
                Dim unsupportedIndex As System.Boolean = False
                Dim unsupportedDetail As System.String = ""
                Await System.Threading.Tasks.Task.Run(
                    Sub()
                        Try
                            manifest = _store.PinGenerationForAdministration(archiveId)
                        Catch ex As System.IO.InvalidDataException When ex.Message.StartsWith("unsupported_semantic_index:", System.StringComparison.Ordinal)
                            unsupportedIndex = True
                            If includeTechnical Then unsupportedDetail = ex.ToString()
                        End Try
                        queueDiagnostics = SemanticArchiveBuilder.ReadDiagnostics(_store, archiveId, includeTechnical)
                        If manifest IsNot Nothing Then publishedDiagnostics = SemanticArchiveBuilder.PresentDiagnostics(_store, archiveId, manifest.Diagnostics, includeTechnical)
                    End Sub)
                If includeTechnical Then
                    ResetDiagnostics("Personal catalog: " & _store.CatalogPath & System.Environment.NewLine & "Archive: " & _archive.Name & " [" & archiveId & "]" & System.Environment.NewLine &
                        "Source roots: " & _archive.Roots.Count.ToString() & System.Environment.NewLine &
                        "Search visibility: " & _archive.Enabled.ToString() & System.Environment.NewLine &
                        "Automatic eligibility: " & _archive.BackgroundEnabled.ToString())
                    AppendDiagnostic("Personal routing metadata uses this Windows user's private access domain. Cooperative per-document artifacts follow verified source permissions; writable private shadows are available when shared publication is unavailable.")
                    AppendDiagnostic("Automatic permission reconciliation is controlled independently of content indexing and never changes original source ACLs.")
                Else
                    ResetDiagnostics("Archive: " & _archive.Name & System.Environment.NewLine &
                        "Source folders: " & _archive.Roots.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) & System.Environment.NewLine &
                        "Search: " & If(_archive.Enabled, "enabled", "disabled") & System.Environment.NewLine &
                        "Automatic background processing: " & If(_archive.BackgroundEnabled, "enabled", "disabled"))
                    AppendDiagnostic("Only items that need attention are listed here. Enable Show technical details for paths, IDs, processing signatures, routing and successful shared-publication events.")
                End If
                For Each message In queueDiagnostics
                    AppendDiagnostic(message)
                Next
                _archiveIndexUnsupported = unsupportedIndex
                If unsupportedIndex Then
                    AppendDiagnostic("This archive uses an unsupported old index format. Create a new archive index; repairing or migrating this old index is not supported. Unregister remains available and will first withdraw a published library definition. Original files and generated artifacts are retained.")
                    If includeTechnical Then AppendDiagnostic(unsupportedDetail)
                    _status.Text = "Old index format: create a new archive or use Unregister to withdraw and remove its registration."
                ElseIf manifest Is Nothing Then
                    AppendDiagnostic("No published generation is available.")
                    _status.Text = "No published generation is available for " & _archive.Name & ". See diagnostics for queued or pending work."
                Else
                    AppendDiagnostic("Generation: " & manifest.GenerationId)
                    AppendDiagnostic("Published UTC: " & manifest.CreatedUtc.ToString("O"))
                    AppendDiagnostic(If(includeTechnical, manifest.Inventory.ToTechnicalDiagnosticText(), manifest.Inventory.ToDiagnosticText()))
                    _status.Text = "Published searchable documents: " & manifest.DocumentCount.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                        "/" & manifest.Inventory.CurrentSources.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                        "; excluded unknown/incomplete: " & manifest.Inventory.ExcludedUnknown.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                        "/" & manifest.Inventory.ExcludedIncomplete.ToString(System.Globalization.CultureInfo.InvariantCulture) & ". See diagnostics for details."
                    For Each message As System.String In publishedDiagnostics
                        AppendDiagnostic(message)
                    Next
                End If
            Catch ex As System.Exception
                ReportError("Archive status could not be read", ex)
            Finally
                SetBusy(False)
            End Try
        End Function

        Private Sub SetBusy(value As Boolean)
            _busy = value
            If IsDisposed Then Return
            _archives.Enabled = Not value AndAlso Not _catalogActivationBlocked
            For Each page As System.Windows.Forms.TabPage In _tabs.TabPages
                If System.Object.ReferenceEquals(page, _backgroundTab) OrElse System.Object.ReferenceEquals(page, _libraryTab) Then
                    page.Enabled = Not value
                ElseIf Not System.Object.ReferenceEquals(page, _diagnostics.Parent) Then
                    page.Enabled = Not value AndAlso _archive IsNot Nothing
                End If
            Next
            _showTechnicalDiagnostics.Enabled = Not value
            _chooseCatalog.Enabled = Not value
            _browseCatalog.Enabled = Not value
            _suggestCatalog.Enabled = Not value
            _newArchive.Enabled = Not value AndAlso Not _catalogActivationBlocked AndAlso _catalog IsNot Nothing
            _removeArchive.Enabled = Not value AndAlso Not _catalogActivationBlocked AndAlso _archive IsNot Nothing
            _reload.Enabled = Not value
            _addRoot.Enabled = Not value AndAlso Not _catalogActivationBlocked AndAlso _archive IsNot Nothing
            _newRootPath.Enabled = Not value AndAlso Not _catalogActivationBlocked AndAlso _archive IsNot Nothing
            _removeRoot.Enabled = Not value AndAlso Not _catalogActivationBlocked AndAlso _root IsNot Nothing
            _rootEditor.Enabled = Not value AndAlso _root IsNot Nothing
            LayoutSourceSections()
            _refresh.Enabled = Not value AndAlso Not _catalogActivationBlocked AndAlso _archive IsNot Nothing AndAlso Not _archiveIndexUnsupported
            _retry.Enabled = Not value AndAlso Not _catalogActivationBlocked AndAlso _archive IsNot Nothing AndAlso Not _archiveIndexUnsupported
            _rebuild.Enabled = Not value AndAlso Not _catalogActivationBlocked AndAlso _archive IsNot Nothing AndAlso Not _archiveIndexUnsupported
            _extractAll.Enabled = Not value AndAlso Not _catalogActivationBlocked AndAlso _archive IsNot Nothing AndAlso Not _archiveIndexUnsupported
            _permissionsAll.Enabled = Not value AndAlso Not _catalogActivationBlocked AndAlso _archive IsNot Nothing
            _loadDocuments.Enabled = Not value AndAlso Not _catalogActivationBlocked AndAlso _archive IsNot Nothing
            _nextDocuments.Enabled = Not value AndAlso Not _catalogActivationBlocked AndAlso _archive IsNot Nothing AndAlso _documentCursor IsNot Nothing AndAlso Not _documentListingComplete AndAlso _documents.Items.Count < MaintenanceLoadedLimit
            _documentFilter.Enabled = Not value
            _documentStateFilter.Enabled = Not value
            _documentTechnical.Enabled = Not value
            _selectedDocumentIds.ReadOnly = value
            _documents.Enabled = Not value
            _selectLoadedDocuments.Enabled = Not value AndAlso _documents.Items.Count > 0
            _clearDocumentSelection.Enabled = Not value
            UpdateSelectedActions()
            _inspect.Enabled = Not value AndAlso Not _catalogActivationBlocked AndAlso _archive IsNot Nothing
            _globalEnabled.Enabled = Not value
            _globalWindow.Enabled = Not value
            _permissionsEnabled.Enabled = Not value
            _permissionsWindow.Enabled = Not value
            UpdateSaveState()
            _pause.Enabled = Not _catalogActivationBlocked AndAlso _archive IsNot Nothing AndAlso (Not value OrElse _operationCancellation IsNot Nothing)
            If Not value Then _pause.Text = If(_paused OrElse CanResumeSelectedArchive(), "Resume", "Pause")
            ApplyLibraryControls(value)
            If Not value Then CompletePendingClose()
        End Sub

        Private Function CreateLibraryTab() As System.Windows.Forms.TabPage
            Dim tab As New System.Windows.Forms.TabPage("Library") With {.AutoScroll = True}
            Dim table As System.Windows.Forms.TableLayoutPanel = NewTable()
            Dim path As New System.Windows.Forms.TextBox() With {.Text = If(_context.INI_SemanticArchiveCatalogLibraryPath, ""), .ReadOnly = True, .Dock = System.Windows.Forms.DockStyle.Fill}
            AddParameterRow(table, "Central library directory", "SemanticArchiveCatalogLibraryPath", path)
            SetHelp(path, "This central directory is configured in redink.ini. It stores archive definitions, not documents. A local catalog path is also required. SharedMethods environment/Red Ink placeholders are expanded before use.")
            AddRow(table, "Publication / subscription", _libraryStatus)
            AddRow(table, "", _publishLibrary)
            AddRow(table, "", _withdrawLibrary)
            AddRow(table, "", _syncLibrary)
            AddRow(table, "How it works", New System.Windows.Forms.Label() With {.AutoSize = True, .MaximumSize = New System.Drawing.Size(530, 0),
                .Text = "Create and test an archive locally, then publish its definition. Readable library entries are subscribed automatically. Local indexes stay private and source-file permissions are checked again at retrieval. New files under existing roots need a refresh, not a new publication. Publish again after changing roots or archive settings. Withdrawal stops subscribers after reconciliation; it does not delete originals or your local archive. Unregistering a subscription is a persistent local opt-out; enable it again here to resubscribe."})
            SetHelp(_publishLibrary, "Save the local definition and publish a new revision. Only the publisher can update it, and Windows must allow the write. Names/descriptions become visible to the definition's readers, as inherited from the library folder. No original documents or private output paths are copied. Network libraries require UNC source paths.")
            SetHelp(_withdrawLibrary, "Publish a withdrawn revision for your own library entry. Subscriptions stop being usable; your local archive and all original documents remain. No permissions are broadened.")
            SetHelp(_syncLibrary, "Read the configured central library and reconcile local subscriptions now. This does not run OCR. New or changed document content is prepared by Outlook's idle maintenance, or by Refresh archive. Library subscribers are prepared automatically even when automatic indexing of your personal archives is disabled; configured time windows still apply.")
            tab.Controls.Add(table)
            Return tab
        End Function

        Private Sub ApplyLibraryControls(busy As System.Boolean)
            Dim configured As System.Boolean = SemanticArchiveLibrary.IsConfigured(_context)
            Dim subscriber As System.Boolean = SemanticArchiveLibrary.IsSubscriber(_archive)
            _syncLibrary.Enabled = configured AndAlso Not busy
            _publishLibrary.Enabled = configured AndAlso Not busy AndAlso Not _catalogActivationBlocked AndAlso _archive IsNot Nothing AndAlso Not subscriber
            _withdrawLibrary.Enabled = _publishLibrary.Enabled AndAlso _archive IsNot Nothing AndAlso _archive.Library IsNot Nothing AndAlso _archive.Library.State <> "withdrawn"
            If Not configured Then
                _libraryStatus.Text = "No central library configured in redink.ini. Personal archives remain available."
            ElseIf _archive Is Nothing OrElse _archive.Library Is Nothing Then
                _libraryStatus.Text = "This is a personal archive. Publish it when its definition is ready to share."
            Else
                _libraryStatus.Text = If(subscriber, "Subscribed", "Published locally") & " — " & _archive.Library.State &
                    "; revision " & _archive.Library.Revision.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                    If(_archive.Library.OptOut, ". You have disabled this subscription.", ".")
            End If
            _name.ReadOnly = subscriber
            _description.ReadOnly = subscriber
            For Each control As System.Windows.Forms.Control In New System.Windows.Forms.Control() {_archiveBackground, _archiveWindow, _partialSearch, _threshold, _children, _routingCharacters, _budgets}
                control.Enabled = Not busy AndAlso Not subscriber AndAlso _archive IsNot Nothing
            Next
            If subscriber Then
                _addRoot.Enabled = False
                _addRootPath.Enabled = False
                _removeRoot.Enabled = False
                _newRootPath.Enabled = False
                _rootEditor.Enabled = False
            End If
            _removeArchive.Text = If(subscriber, "Unsubscribe", "Unregister")
        End Sub

        Private Async Function SyncLibraryAsync() As System.Threading.Tasks.Task
            If _busy OrElse Not SemanticArchiveLibrary.IsConfigured(_context) Then Return
            If _archiveDirty AndAlso Not Await SaveArchiveAsync() Then Return
            Dim cancellation As New System.Threading.CancellationTokenSource()
            _operationCancellation = cancellation
            SetBusy(True)
            _status.Text = "Synchronizing available archive definitions..."
            Dim success As System.Boolean = False
            Try
                Await SemanticArchiveLibrary.SynchronizeAsync(_context, True, cancellation.Token)
                RetrievalSourceDiscovery.RequestRefresh(_context, True)
                success = True
            Catch ex As System.OperationCanceledException
                _status.Text = "Library operation cancelled before completion."
            Catch ex As System.Exception
                ReportError("Library synchronization could not be completed; private archives are retained", ex)
            Finally
                _operationCancellation = Nothing
                cancellation.Dispose()
                SetBusy(False)
            End Try
            If success AndAlso Not _closeRequested Then Await ReloadCatalogAsync()
        End Function

        Private Async Function PublishLibraryAsync(withdraw As System.Boolean) As System.Threading.Tasks.Task
            If _busy OrElse _archive Is Nothing OrElse Not SemanticArchiveLibrary.IsConfigured(_context) OrElse SemanticArchiveLibrary.IsSubscriber(_archive) Then Return
            If Not Await SaveArchiveAsync() OrElse _closeRequested Then Return
            If Not ConfirmConsoleAction(If(withdraw, "Withdraw this archive from the central library? Your local archive and original files remain.",
                "Publish this archive definition? Its name, description and source folders will be visible to the library entry's readers. Original files keep their own permissions and are not copied."), If(withdraw, "Withdraw", "Publish"), "Cancel") Then Return
            Dim archiveId As System.String = _archive.ArchiveId
            Dim cancellation As New System.Threading.CancellationTokenSource()
            _operationCancellation = cancellation
            SetBusy(True)
            _status.Text = If(withdraw, "Withdrawing library definition...", "Publishing library definition...")
            Dim success As System.Boolean = False
            Try
                Await System.Threading.Tasks.Task.Run(Sub() SemanticArchiveLibrary.Publish(_context, archiveId, withdraw, cancellation.Token))
                success = True
            Catch ex As System.OperationCanceledException
                _status.Text = "Library operation cancelled before completion."
            Catch ex As System.Exception
                ReportError("The library change could not be completed or confirmed", ex)
            Finally
                _operationCancellation = Nothing
                cancellation.Dispose()
                SetBusy(False)
            End Try
            If success AndAlso Not _closeRequested Then
                _status.Text = If(withdraw, "Withdrawn from the library; your local archive remains available.", "Published. Readers receive this definition through automatic library synchronization.")
                ' Reload failures retain their own diagnostic; do not overwrite them with success.
                Await ReloadCatalogAsync(archiveId)
            End If
        End Function

        Private Function CreateDocumentsTab() As System.Windows.Forms.TabPage
            Dim tab As New System.Windows.Forms.TabPage("Documents and selected actions")
            Dim layout As New System.Windows.Forms.TableLayoutPanel() With {.Dock = System.Windows.Forms.DockStyle.Fill, .ColumnCount = 1, .RowCount = 6, .Padding = New System.Windows.Forms.Padding(4)}
            layout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 76))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            _documents.Columns.Add("Document", 260)
            _documents.Columns.Add("Status", 180)
            _documents.Columns.Add("Recommended action", 285)
            _documents.Columns.Add("Original folder", 280)
            _documents.Columns.Add("Document ID", 0)
            _documents.ShowItemToolTips = True
            _documents.ListViewItemSorter = New MaintenanceDocumentComparer(0, False)
            _documentStateFilter.Items.AddRange(New System.Object() {"Needs attention", "Incomplete / unknown extraction", "No readable text", "Processing failed", "Needs extraction / indexing", "Source access unavailable", "Searchable", "All documents", "Removed"})
            _documentStateFilter.SelectedIndex = 0
            Dim top As New System.Windows.Forms.FlowLayoutPanel() With {.Dock = System.Windows.Forms.DockStyle.Fill, .AutoSize = True, .WrapContents = True}
            top.Controls.Add(_documentStateFilter)
            top.Controls.Add(New System.Windows.Forms.Label() With {.Text = "Name or path contains", .AutoSize = True, .Padding = New System.Windows.Forms.Padding(0, 7, 0, 0)})
            top.Controls.Add(_documentFilter)
            top.Controls.Add(_loadDocuments)
            top.Controls.Add(_nextDocuments)
            top.Controls.Add(_documentTechnical)
            layout.Controls.Add(top, 0, 0)
            layout.Controls.Add(_documents, 0, 1)
            layout.Controls.Add(_documentDetails, 0, 2)
            Dim actions As New System.Windows.Forms.FlowLayoutPanel() With {.Dock = System.Windows.Forms.DockStyle.Fill, .AutoSize = True, .WrapContents = True}
            actions.Controls.AddRange(New System.Windows.Forms.Control() {_selectLoadedDocuments, _clearDocumentSelection, _documentSelectionStatus, _retrySelected, _extractSelected, _rebuildSelected, _permissionsSelected})
            layout.Controls.Add(actions, 0, 3)
            _selectedDocumentIds.Dock = System.Windows.Forms.DockStyle.Fill
            _selectedDocumentIds.AccessibleName = "Explicit document IDs (one per line)"
            SetHelp(_selectedDocumentIds, "Advanced: paste stable document IDs, one per line. Hidden IDs still belong to the selected actions; Clear selection removes them. An empty selection never means all.")
            _documentAdvanced.Controls.Add(_selectedDocumentIds)
            layout.Controls.Add(_documentAdvanced, 0, 4)
            _documentPageStatus.Dock = System.Windows.Forms.DockStyle.Fill
            _documentPageStatus.AutoSize = True
            _documentPageStatus.Text = "Choose a status and Find documents. Column headers sort loaded matches. Changing filters clears selection."
            layout.Controls.Add(_documentPageStatus, 0, 5)
            tab.Controls.Add(layout)
            Return tab
        End Function

        Private Sub ResetDocumentListing()
            If _documentCursor IsNot Nothing Then _documentCursor.Dispose()
            _documentCursor = Nothing
            _documentGenerationId = ""
            _documentListingComplete = False
        End Sub

        Private Sub DocumentFilterChanged(sender As System.Object, e As System.EventArgs)
            If _busy Then Return
            ResetDocumentListing()
            _documents.Items.Clear()
            _selectedDocumentIds.Clear()
            _documentDetails.Clear()
            _nextDocuments.Enabled = False
            _documentPageStatus.Text = "Filter changed. Find documents searches the published metadata; only matching documents are displayed."
        End Sub

        Private Async Function LoadDocumentPageAsync(restart As System.Boolean) As System.Threading.Tasks.Task
            If _closeRequested OrElse _busy OrElse _catalogActivationBlocked OrElse _archive Is Nothing OrElse _store Is Nothing Then Return
            Dim archiveId As System.String = _archive.ArchiveId
            Dim filter As System.String = _documentFilter.Text.Trim()
            Dim stateFilter As System.Int32 = _documentStateFilter.SelectedIndex
            Dim rows As New System.Collections.Generic.List(Of MaintenanceDocumentRow)()
            Dim inspected As System.Int32 = 0
            Dim cancellation As New System.Threading.CancellationTokenSource()
            _operationCancellation = cancellation
            SetBusy(True)
            Try
                If restart Then
                    ResetDocumentListing()
                    _documents.Items.Clear()
                    _selectedDocumentIds.Clear()
                    _documentDetails.Clear()
                End If
                Dim pageSize As System.Int32 = System.Math.Min(MaintenancePageSize, MaintenanceLoadedLimit - _documents.Items.Count)
                ' Continue across empty metadata chunks until a matching page or the end is reached.
                ' UI/cancellation stay live between bounded worker chunks; no original-folder scan occurs.
                While rows.Count < pageSize AndAlso Not _documentListingComplete
                    Await System.Threading.Tasks.Task.Run(
                        Sub()
                            cancellation.Token.ThrowIfCancellationRequested()
                            Dim currentArchive As SemanticArchiveDefinition = _store.GetArchive(archiveId)
                            If currentArchive Is Nothing Then Throw New System.InvalidOperationException("The archive was unregistered; reload the catalog.")
                            Dim access As SemanticArchiveAccessContext = SemanticArchiveAccessContext.CreateForCurrentUser()
                            If _documentCursor Is Nothing Then
                                Dim manifest As SemanticArchiveGenerationManifest = _store.PinGenerationForAdministration(archiveId)
                                If manifest Is Nothing Then
                                    _documentListingComplete = True
                                    Return
                                End If
                                _documentGenerationId = manifest.GenerationId
                                _documentCursor = _store.EnumerateDocuments(manifest).GetEnumerator()
                            End If
                            Dim timer As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew()
                            Dim chunkInspected As System.Int32 = 0
                            While rows.Count < pageSize AndAlso chunkInspected < 1000 AndAlso timer.Elapsed.TotalSeconds < 1
                                cancellation.Token.ThrowIfCancellationRequested()
                                If Not _documentCursor.MoveNext() Then
                                    _documentListingComplete = True
                                    Exit While
                                End If
                                inspected += 1
                                chunkInspected += 1
                                Dim document As SemanticArchiveDocumentRecord = _documentCursor.Current
                                Dim state As System.String = MaintenanceState(document)
                                ' Status-only views skip irrelevant records before costly source checks.
                                ' Restricted records are NEVER emitted based on their hidden cached status.
                                If stateFilter <> 5 AndAlso stateFilter <> 7 AndAlso Not MatchesMaintenanceState(state, stateFilter, SemanticArchiveInventory.IsSearchable(document)) Then Continue While
                                Dim readable As System.Boolean = CanDiscloseMaintenanceSource(document, currentArchive, access)
                                If Not readable AndAlso stateFilter <> 5 AndAlso stateFilter <> 7 Then Continue While
                                If readable AndAlso stateFilter = 5 Then Continue While
                                Dim row As MaintenanceDocumentRow = CreateMaintenanceRow(document, readable, state)
                                If filter.Length > 0 AndAlso row.DocumentId.IndexOf(filter, System.StringComparison.OrdinalIgnoreCase) < 0 AndAlso
                                   (Not readable OrElse row.SourceDisplay.IndexOf(filter, System.StringComparison.OrdinalIgnoreCase) < 0) Then Continue While
                                rows.Add(row)
                            End While
                        End Sub, cancellation.Token)
                    If IsDisposed Then Return
                    _documentPageStatus.Text = "Finding matching documents: " & rows.Count.ToString() & " found; " & inspected.ToString() & " metadata records checked. Pause cancels this search."
                End While
                If IsDisposed Then Return
                _updatingDocumentRows = True
                _documents.BeginUpdate()
                Try
                    For Each document As MaintenanceDocumentRow In rows
                        Dim item As New System.Windows.Forms.ListViewItem(document.DisplayName) With {.Tag = document, .ToolTipText = document.SourceDisplay}
                        item.SubItems.Add(document.Status)
                        item.SubItems.Add(document.RecommendedAction)
                        item.SubItems.Add(document.Folder)
                        item.SubItems.Add(document.DocumentId)
                        _documents.Items.Add(item)
                    Next
                    _documents.Sort()
                Finally
                    _documents.EndUpdate()
                    _updatingDocumentRows = False
                End Try
                _documentPageStatus.Text = _documents.Items.Count.ToString() & " matching documents loaded. " &
                    If(_documentListingComplete, "All matches in this published snapshot have been checked.",
                       If(_documents.Items.Count >= MaintenanceLoadedLimit, "Display limit reached; refine the filters to narrow the results.", "Load more matches continues; existing selections are retained.")) &
                    " Sorting applies to loaded matches."
            Catch ex As System.OperationCanceledException
                ResetDocumentListing()
                If Not IsDisposed Then _documentPageStatus.Text = "Search cancelled. Previously loaded matches and selections remain; Find documents starts a fresh search."
            Catch ex As System.Exception
                ResetDocumentListing()
                ReportError("Document metadata could not be listed", ex)
            Finally
                _operationCancellation = Nothing
                cancellation.Dispose()
                If IsDisposed Then ResetDocumentListing()
                SetBusy(False)
            End Try
        End Function

        Private Shared Function MaintenanceState(document As SemanticArchiveDocumentRecord) As System.String
            If document Is Nothing Then Return "unavailable"
            Dim status As System.String = If(document.ProcessingStatus, System.String.Empty).Trim().ToLowerInvariant()
            If status = "removed" OrElse status = "failed" OrElse status = "unavailable" OrElse status = "pending_host" OrElse status = "empty" Then Return status
            If document.Representation Is Nothing Then Return "needs_extraction"
            Dim coverage As System.String = If(document.Representation.Completeness, "unknown").ToLowerInvariant()
            If coverage = "empty" OrElse coverage = "incomplete" OrElse coverage = "unknown" Then Return coverage
            If SemanticArchiveInventory.IsSearchable(document) Then Return "searchable"
            Return "needs_indexing"
        End Function

        Private Shared Function MatchesMaintenanceState(state As System.String, filter As System.Int32, searchable As System.Boolean) As System.Boolean
            Select Case filter
                Case 0 : Return state <> "searchable" AndAlso state <> "removed"
                Case 1 : Return state = "incomplete" OrElse state = "unknown"
                Case 2 : Return state = "empty"
                Case 3 : Return state = "failed" OrElse state = "unavailable"
                Case 4 : Return state = "needs_extraction" OrElse state = "needs_indexing" OrElse state = "pending_host"
                Case 6 : Return searchable
                Case 8 : Return state = "removed"
                Case Else : Return True
            End Select
        End Function

        Private Shared Function CreateMaintenanceRow(document As SemanticArchiveDocumentRecord, readable As System.Boolean, state As System.String) As MaintenanceDocumentRow
            If Not readable Then Return New MaintenanceDocumentRow(document.DocumentId, "Source access unavailable", "", "Source access unavailable", "Check source access; reconcile permissions if needed.", "", "", "The original source could not be verified. Its cached name, path, processing status and diagnostics are hidden.", "")
            Dim name As System.String = System.IO.Path.GetFileName(document.SourcePath)
            Dim folder As System.String = System.IO.Path.GetDirectoryName(document.SourcePath)
            Dim status As System.String
            Dim action As System.String
            Dim explanation As System.String
            Select Case state
                Case "searchable"
                    status = "Searchable" : action = "No repair needed." : explanation = "Extracted text and semantic index are available."
                Case "incomplete"
                    status = "Incomplete extraction" : action = "Retry selected; check OCR settings if needed." : explanation = "Some source content could not be validated. Re-extraction/OCR uses this source folder's settings."
                Case "unknown"
                    status = "Extraction not verified" : action = "Retry selected to verify extraction." : explanation = "The completeness of the extracted content has not been established."
                Case "empty"
                    status = "No readable text" : action = "Check original; re-extract/OCR selected." : explanation = "No readable text was extracted. The source may be empty, image-only or unsupported by its configured reader."
                Case "failed", "unavailable"
                    status = "Processing failed" : action = "Retry selected; inspect details if it fails again." : explanation = "Extraction or indexing did not complete successfully."
                Case "pending_host"
                    status = "Needs Office reader" : action = "Process through a supported Office host." : explanation = "This source requires a host reader; the standalone worker cannot read it."
                Case "needs_extraction"
                    status = "Needs extraction" : action = "Re-extract/OCR selected." : explanation = "No usable text representation is available."
                Case "removed"
                    status = "Removed" : action = "Check source registration if unexpected." : explanation = "This retained record is no longer a current source."
                Case Else
                    status = "Needs semantic indexing" : action = "Rebuild semantic index (selected)." : explanation = "Extracted text is available, but this document is not currently searchable."
            End Select
            If state = "incomplete" OrElse state = "unknown" Then explanation &= If(SemanticArchiveInventory.IsSearchable(document), " Partial-text search is currently allowed for this document.", " This document is currently excluded from search.")
            Return New MaintenanceDocumentRow(document.DocumentId, name, document.SourcePath, status, action, folder, state, explanation, If(document.Diagnostic, ""))
        End Function

        Private Shared Function CanDiscloseMaintenanceSource(document As SemanticArchiveDocumentRecord,
                                                             archive As SemanticArchiveDefinition,
                                                             access As SemanticArchiveAccessContext) As Boolean
            If document Is Nothing OrElse archive Is Nothing OrElse access Is Nothing OrElse document.BindingIds Is Nothing Then Return False
            For Each binding In archive.Roots
                If Not document.BindingIds.Contains(binding.BindingId) Then Continue For
                Try
                    If SemanticArchiveStore.IsSourceExcluded(binding, document.SourcePath) OrElse
                       Not SemanticArchiveStore.IsSupportedSource(binding, document.SourcePath) Then Continue For
                    Dim identity = SemanticArchivePathGuard.GetVerifiedSourceIdentity(binding.RootPath, document.SourcePath)
                    If Not System.String.Equals(identity, document.CanonicalSourceKey, System.StringComparison.OrdinalIgnoreCase) Then Continue For
                    If access.CanReadSource(document.SourcePath) Then Return True
                Catch ex As System.Exception
                    ' Missing, moved, replaced, excluded and inaccessible originals all hide cached labels.
                End Try
            Next
            Return False
        End Function

        Private NotInheritable Class MaintenanceDocumentRow
            Public ReadOnly DocumentId As System.String
            Public ReadOnly DisplayName As System.String
            Public ReadOnly SourceDisplay As System.String
            Public ReadOnly Status As System.String
            Public ReadOnly RecommendedAction As System.String
            Public ReadOnly Folder As System.String
            Public ReadOnly State As System.String
            Public ReadOnly Explanation As System.String
            Public ReadOnly Diagnostic As System.String
            Public Sub New(documentId As System.String, displayName As System.String, sourceDisplay As System.String, status As System.String, action As System.String, folder As System.String, state As System.String, explanation As System.String, diagnostic As System.String)
                Me.DocumentId = documentId : Me.DisplayName = displayName : Me.SourceDisplay = sourceDisplay
                Me.Status = status : Me.RecommendedAction = action : Me.Folder = folder : Me.State = state
                Me.Explanation = explanation : Me.Diagnostic = diagnostic
            End Sub
        End Class

        Private NotInheritable Class MaintenanceDocumentComparer
            Implements System.Collections.IComparer
            Private ReadOnly _column As System.Int32
            Private ReadOnly _descending As System.Boolean
            Public Sub New(column As System.Int32, descending As System.Boolean)
                _column = column : _descending = descending
            End Sub
            Public Function Compare(x As System.Object, y As System.Object) As System.Int32 Implements System.Collections.IComparer.Compare
                Dim left As System.Windows.Forms.ListViewItem = DirectCast(x, System.Windows.Forms.ListViewItem)
                Dim right As System.Windows.Forms.ListViewItem = DirectCast(y, System.Windows.Forms.ListViewItem)
                Dim result As System.Int32 = System.StringComparer.CurrentCultureIgnoreCase.Compare(left.SubItems(_column).Text, right.SubItems(_column).Text)
                If result = 0 Then result = System.StringComparer.Ordinal.Compare(DirectCast(left.Tag, MaintenanceDocumentRow).DocumentId, DirectCast(right.Tag, MaintenanceDocumentRow).DocumentId)
                Return If(_descending, -result, result)
            End Function
        End Class

        Private Sub DocumentColumnClicked(sender As System.Object, e As System.Windows.Forms.ColumnClickEventArgs)
            If _busy Then Return
            If _documentSortColumn = e.Column Then
                _documentSortDescending = Not _documentSortDescending
            Else
                _documentSortColumn = e.Column
                _documentSortDescending = False
            End If
            _updatingDocumentRows = True
            Try
                _documents.ListViewItemSorter = New MaintenanceDocumentComparer(_documentSortColumn, _documentSortDescending)
                _documents.Sort()
            Finally
                _updatingDocumentRows = False
            End Try
        End Sub

        Private Sub DocumentSelectionChanged(sender As System.Object, e As System.EventArgs)
            If _updatingDocumentRows Then Return
            Dim ids As New System.Collections.Generic.List(Of System.String)()
            For Each item As System.Windows.Forms.ListViewItem In _documents.SelectedItems
                ids.Add(DirectCast(item.Tag, MaintenanceDocumentRow).DocumentId)
            Next
            _selectedDocumentIds.Text = System.String.Join(System.Environment.NewLine, ids)
            UpdateDocumentDetails()
        End Sub

        Private Sub SelectLoadedDocuments(sender As System.Object, e As System.EventArgs)
            If _busy Then Return
            If _documents.Items.Count > 1024 Then
                _documentPageStatus.Text = "At most 1,024 documents can be selected for one command. Refine the filter or select a smaller set."
                Return
            End If
            _updatingDocumentRows = True
            _documents.BeginUpdate()
            Try
                For Each item As System.Windows.Forms.ListViewItem In _documents.Items
                    item.Selected = True
                Next
            Finally
                _documents.EndUpdate()
                _updatingDocumentRows = False
            End Try
            DocumentSelectionChanged(sender, e)
        End Sub

        Private Sub UpdateDocumentDetails()
            _documentSelectionStatus.Text = ParseLines(_selectedDocumentIds.Text, New Char() {Microsoft.VisualBasic.ChrW(13), Microsoft.VisualBasic.ChrW(10), ";"c, ","c}).Count.ToString() & " selected"
            If _documents.SelectedItems.Count <> 1 Then
                _documentDetails.Text = "Select one document for an explanation and recommended action. Actions below apply only to the explicitly selected documents."
                Return
            End If
            Dim row As MaintenanceDocumentRow = DirectCast(_documents.SelectedItems(0).Tag, MaintenanceDocumentRow)
            _documentDetails.Text = row.DisplayName & " — " & row.Status & System.Environment.NewLine & row.Explanation & System.Environment.NewLine & row.RecommendedAction
            If _documentTechnical.Checked Then
                _documentDetails.AppendText(System.Environment.NewLine & "Document ID: " & row.DocumentId)
                If row.Diagnostic.Length > 0 Then _documentDetails.AppendText(System.Environment.NewLine & row.Diagnostic)
            End If
        End Sub

        Private Function ReadSelectedDocumentIds() As System.Collections.Generic.List(Of String)
            Dim ids = ParseLines(_selectedDocumentIds.Text, New Char() {Microsoft.VisualBasic.ChrW(13), Microsoft.VisualBasic.ChrW(10), ";"c, ","c})
            If ids.Count = 0 Then Throw New System.ArgumentException("Select one or more stable document IDs. Use an explicitly labelled all-documents action for the whole archive.")
            If ids.Count > 1024 Then Throw New System.ArgumentException("Select at most 1,024 document IDs per targeted command.")
            For Each id In ids
                SemanticArchiveIdentity.ValidateId(id, "selectedDocumentId")
            Next
            Return ids
        End Function

        Private Sub UpdateSelectedActions()
            Dim enabled As Boolean = Not _busy AndAlso Not _catalogActivationBlocked AndAlso _archive IsNot Nothing AndAlso Not System.String.IsNullOrWhiteSpace(_selectedDocumentIds.Text)
            _rebuildSelected.Enabled = enabled AndAlso Not _archiveIndexUnsupported
            _extractSelected.Enabled = enabled AndAlso Not _archiveIndexUnsupported
            _permissionsSelected.Enabled = enabled
            _retrySelected.Enabled = enabled AndAlso Not _archiveIndexUnsupported
        End Sub

        Private Sub ResetDiagnostics(Optional initialText As System.String = Nothing)
            If IsDisposed Then Return
            _diagnosticDisplayTrimmed = False
            _diagnostics.Clear()
            If initialText IsNot Nothing Then AppendDiagnostic(initialText)
        End Sub

        Private Shared Function CompleteDiagnosticPrefix(value As System.String, maximumCharacters As System.Int32) As System.String
            Dim length As System.Int32 = System.Math.Min(value.Length, System.Math.Max(0, maximumCharacters))
            If length = 0 Then Return ""
            Dim lastLineEnd As System.Int32 = value.LastIndexOf(Microsoft.VisualBasic.ChrW(10), length - 1)
            If lastLineEnd >= 0 Then Return value.Substring(0, lastLineEnd + 1)
            ' A single overlong line has no complete line to retain. Keep a valid
            ' Unicode prefix and explicitly mark the omitted end in the caller.
            If length < value.Length AndAlso System.Char.IsHighSurrogate(value(length - 1)) AndAlso System.Char.IsLowSurrogate(value(length)) Then length -= 1
            If length > 0 AndAlso value(length - 1) = Microsoft.VisualBasic.ChrW(13) Then length -= 1
            Return value.Substring(0, length)
        End Function

        Private Sub AppendDiagnostic(message As String)
            If IsDisposed Then Return
            Dim maximum As System.Int32 = SharedMethods.DEFAULT_SEMANTICARCHIVE_MAX_DIAGNOSTIC_DISPLAY_CHARACTERS
            Dim prefix As System.String = DiagnosticDisplayOmissionNotice & System.Environment.NewLine
            Dim latest As System.String = If(message, "")
            Dim available As System.Int32 = maximum - prefix.Length
            If latest.Length > available - System.Environment.NewLine.Length Then
                Dim shortened As System.String = System.Environment.NewLine & DiagnosticEntryOmissionNotice & System.Environment.NewLine
                latest = CompleteDiagnosticPrefix(latest, available - shortened.Length) & shortened
            Else
                latest &= System.Environment.NewLine
            End If
            Dim selectionStart As System.Int32 = _diagnostics.SelectionStart
            Dim selectionLength As System.Int32 = _diagnostics.SelectionLength
            If latest.Length <= maximum - _diagnostics.TextLength Then
                _diagnostics.AppendText(latest)
                If selectionLength > 0 Then _diagnostics.Select(selectionStart, selectionLength)
                Return
            End If
            Dim current As System.String = _diagnostics.Text
            Dim oldPrefixLength As System.Int32 = If(_diagnosticDisplayTrimmed AndAlso current.StartsWith(prefix, System.StringComparison.Ordinal), prefix.Length, 0)
            If oldPrefixLength > 0 Then current = current.Substring(oldPrefixLength)
            Dim requiredRemoval As System.Int32 = System.Math.Max(0, current.Length - (available - latest.Length))
            ' Trim in bounded blocks instead of rebuilding the full display for every
            ' small progress message once the archive has produced many diagnostics.
            Dim removeAtLeast As System.Int32 = System.Math.Max(requiredRemoval, System.Math.Min(current.Length, maximum \ 4))
            Dim nextLineEnd As System.Int32 = If(removeAtLeast > 0, current.IndexOf(Microsoft.VisualBasic.ChrW(10), removeAtLeast - 1), -1)
            Dim removed As System.Int32 = If(nextLineEnd >= 0, nextLineEnd + 1, current.Length)
            Dim retained As System.String = current.Substring(removed)
            _diagnostics.Text = prefix & retained & latest
            _diagnosticDisplayTrimmed = True
            Dim retainedSelectionStart As System.Int32 = System.Math.Max(selectionStart, oldPrefixLength + removed)
            Dim retainedSelectionEnd As System.Int32 = selectionStart + selectionLength
            If selectionLength > 0 AndAlso retainedSelectionEnd > retainedSelectionStart Then
                _diagnostics.Select(prefix.Length + retainedSelectionStart - oldPrefixLength - removed, retainedSelectionEnd - retainedSelectionStart)
            Else
                _diagnostics.Select(_diagnostics.TextLength, 0)
            End If
        End Sub

        Private Function DiagnosticTextForCopy() As System.String
            Return If(_diagnostics.SelectionLength > 0, _diagnostics.SelectedText, _diagnostics.Text)
        End Function

        Private Sub CopyDiagnostics()
            Dim text As System.String = DiagnosticTextForCopy()
            If System.String.IsNullOrEmpty(text) Then Return
            Try
                ' This synchronous click already runs on the form's STA. Keep clipboard
                ' errors on that thread so they can be reported instead of escaping a worker.
                System.Windows.Forms.Clipboard.SetText(text, System.Windows.Forms.TextDataFormat.UnicodeText)
                _status.Text = "Diagnostic text copied to the clipboard."
            Catch ex As System.Exception
                ReportError("Diagnostic text could not be copied", ex)
            End Try
        End Sub

        Private Sub ReportError(operation As String, ex As System.Exception)
            If IsDisposed Then Return
            _status.Text = operation & ": " & ex.Message
            _toolTips.SetToolTip(_status, _status.Text)
            AppendDiagnostic(_status.Text)
            AppendDiagnostic(ex.ToString())
            _tabs.SelectedTab = DirectCast(_diagnostics.Parent, System.Windows.Forms.TabPage)
            System.Diagnostics.Debug.WriteLine("Semantic Archive console: " & operation & ": " & ex.ToString())
        End Sub

        Private NotInheritable Class ArchiveItem
            Public ReadOnly Definition As SemanticArchiveDefinition
            Public Sub New(definition As SemanticArchiveDefinition)
                Me.Definition = definition
            End Sub
            Public Overrides Function ToString() As String
                Return Definition.Name & " [" & Definition.ArchiveId & "]"
            End Function
        End Class

        Private NotInheritable Class RootItem
            Public ReadOnly Binding As SemanticArchiveSourceBinding
            Public Sub New(binding As SemanticArchiveSourceBinding)
                Me.Binding = binding
            End Sub
            Public Overrides Function ToString() As String
                Return Binding.RootPath
            End Function
        End Class
    End Class
End Namespace
