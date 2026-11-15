' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. For license to use see https://redink.ai.
' Offline WebView2 review controller. The browser sends intentions and opaque IDs;
' this host validates every mutation, owns file access and commits through the adapter.
Option Explicit On
Option Strict On

Namespace Global.SharedLibrary.TabularReview
    Public NotInheritable Class TabularReviewForm
        Inherits Global.SharedLibrary.SharedLibrary.WebView2HostForm

        Private ReadOnly _context As Global.SharedLibrary.SharedLibrary.SharedContext.ISharedContext
        Private ReadOnly _callbacks As ReviewHostCallbacks
        Private ReadOnly _web As New Global.SharedLibrary.SharedLibrary.DpiAwareWebView2()
        Private ReadOnly _lifetime As System.Threading.CancellationTokenSource
        Private ReadOnly _hostLifetime As System.Threading.CancellationToken
        Private ReadOnly _session As System.String = System.Guid.NewGuid().ToString("N")
        Private ReadOnly _assetRoot As System.String
        Private ReadOnly _origin As System.String
        Private ReadOnly _allowedPaths As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
        Private ReadOnly _models As System.Collections.Generic.List(Of ReviewModelChoice)
        Private _project As ReviewProject
        Private Shared _sessionReviewer As System.String = ""
        Private _ready As System.Boolean
        Private _busy As System.Boolean
        Private _closed As System.Boolean
        Private _closing As System.Boolean
        Private _allowClose As System.Boolean
        Private _operation As System.Threading.CancellationTokenSource
        Private _operationDone As System.Threading.Tasks.TaskCompletionSource(Of System.Boolean)
        Private _sourceRequest As System.Threading.CancellationTokenSource
        Private _flush As System.Threading.Tasks.TaskCompletionSource(Of System.Boolean)
        Private _flushId As System.String = ""
        Private _persistedViewJson As System.String = ""
        Private _pendingMetadata As System.Boolean
        Private _savingWorkbook As System.Boolean
        Private _filenameDeclined As System.Boolean
        Private _lastSavedUtc As System.String = ""
        Private _persistenceMessage As System.String = ""
        Private _portableNotice As System.String = ""
        Private _consistencyPlanId As System.String = ""
        Private _consistencyPlanRevision As System.Int64 = -1
        Private _consistencyPlanScope As System.String = ""
        Private _consistencyPlanModel As System.String = ""
        Private _consistencyPlanInstruction As System.String = ""
        Private _cachedLogoDataUrl As System.String
        Private ReadOnly _columnTemplateEntries As New System.Collections.Generic.Dictionary(Of System.String, ReviewTemplateLibraryEntry)(System.StringComparer.Ordinal)
        Private ReadOnly _columnTemplateTargets As New System.Collections.Generic.Dictionary(Of System.String, ReviewTemplateWriteTarget)(System.StringComparer.Ordinal)
        Private ReadOnly _browserStateResolver As New BrowserStateContractResolver()
        Private _message As System.String = "Choose a source folder or import an Excel result. Save the workbook to keep your review."

        Public Sub New(context As Global.SharedLibrary.SharedLibrary.SharedContext.ISharedContext,
                       project As ReviewProject, callbacks As ReviewHostCallbacks,
                       hostLifetime As System.Threading.CancellationToken)
            If context Is Nothing OrElse callbacks Is Nothing Then Throw New System.ArgumentNullException("reviewHost")
            ReviewContract.Validate(project)
            _context = context
            _project = project
            _callbacks = callbacks
            _persistedViewJson = ReviewContract.Serialize(_project.View)
            If _project.View.LayoutVersion < 3 Then
                If _project.View.LayoutVersion < 2 Then _project.View.Mode = "matrix"
                ApplyCriterionOrder(_project, _project.View.ColumnOrder)
                _project.View.LayoutVersion = 3
                _pendingMetadata = True
            End If
            _hostLifetime = hostLifetime
            _lifetime = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(hostLifetime)
            _assetRoot = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "RedInk", "TabularReview", _session)
            _origin = "https://redink-review-" & _session & ".invalid/"
            _models = ReviewModelCatalog.Load(context, Sub(detail As System.String) _message = detail)
            If System.String.IsNullOrEmpty(_project.ModelKey) AndAlso _models.Count > 0 Then
                Dim currentKey As System.String = "primary:" & Global.SharedLibrary.SharedLibrary.PrimaryModelManager.GetCurrentModelNumber().ToString(System.Globalization.CultureInfo.InvariantCulture)
                Dim choice As ReviewModelChoice = _models.Find(Function(item As ReviewModelChoice) item.Key = currentKey)
                _project.ModelKey = If(choice Is Nothing, _models(0).Key, choice.Key)
                _pendingMetadata = True
            End If
            For Each run As ReviewConsistencyRun In _project.ConsistencyRuns
                If run.Status = "running" Then
                    run.Status = "partial"
                    run.CompletedUtc = System.DateTime.UtcNow.ToString("o")
                    run.CoverageMessage &= " The previous session ended before this check completed. Retained findings remain available; run the scope again to complete its coverage."
                    _pendingMetadata = True
                End If
            Next
            Me.Text = "Red Ink — Tabular Review"
            Dim area As System.Drawing.Rectangle = System.Windows.Forms.Screen.FromPoint(System.Windows.Forms.Cursor.Position).WorkingArea
            Me.Size = New System.Drawing.Size(System.Math.Min(1440, area.Width), System.Math.Min(920, area.Height))
            Me.MinimumSize = New System.Drawing.Size(System.Math.Min(900, area.Width), System.Math.Min(620, area.Height))
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
            Me.Font = New System.Drawing.Font("Segoe UI", 9.0F)
            Global.SharedLibrary.SharedLibrary.SharedMethods.ApplyStandardDialogIcon(Me)
            _web.Dock = System.Windows.Forms.DockStyle.Fill
            Me.Controls.Add(_web)
            AddHandler Me.Shown, AddressOf InitializeAsync
        End Sub

        Public ReadOnly Property Project As ReviewProject
            Get
                Return _project
            End Get
        End Property

        Public ReadOnly Property IsWorking As System.Boolean
            Get
                Return _busy
            End Get
        End Property

        Private NotInheritable Class SourceFileCheck
            Public Property Hash As System.String = ""
            Public Property Status As System.String = ""
            Public Property Message As System.String = ""
        End Class

        Private Shared Async Function CheckSourceFileAsync(path As System.String, cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of SourceFileCheck)
            Dim result As New SourceFileCheck()
            Try
                cancellation.ThrowIfCancellationRequested()
                result.Hash = Await System.Threading.Tasks.Task.Run(Function() ReviewSourceService.FileHash(path, cancellation), cancellation).ConfigureAwait(False)
            Catch ex As System.OperationCanceledException
                Throw
            Catch ex As System.Exception
                result.Status = If(System.IO.File.Exists(path), "error", "missing")
                result.Message = "The original source is unavailable or unreadable. Link the source or restore its file; the saved extract remains available. " & ex.Message
                System.Diagnostics.Trace.WriteLine("[TabularReview:source_check] " & ex.ToString())
            End Try
            Return result
        End Function

        Private Shared Function WorkbookDirectory(value As System.String) As System.String
            ' Native workbook locations may be URLs. Save still belongs to Excel;
            ' portable filesystem metadata only applies to local or UNC locations.
            If System.String.IsNullOrWhiteSpace(value) OrElse value.IndexOf("://", System.StringComparison.Ordinal) >= 0 OrElse Not System.IO.Path.IsPathRooted(value) Then Return ""
            Return If(System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(value)), "")
        End Function

        Private Shared Function ContainedRelativePath(root As System.String, path As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(root) OrElse System.String.IsNullOrWhiteSpace(path) OrElse Not System.IO.Path.IsPathRooted(root) OrElse Not System.IO.Path.IsPathRooted(path) Then Return ""
            Dim fullRoot As System.String = System.IO.Path.GetFullPath(root).TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar)
            Dim fullPath As System.String = System.IO.Path.GetFullPath(path).TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar)
            If System.String.Equals(fullRoot, fullPath, System.StringComparison.OrdinalIgnoreCase) Then Return "."
            Dim prefix As System.String = fullRoot & System.IO.Path.DirectorySeparatorChar
            If Not fullPath.StartsWith(prefix, System.StringComparison.OrdinalIgnoreCase) Then Return ""
            Return fullPath.Substring(prefix.Length).Replace("\", "/")
        End Function

        Private Shared Function DocumentPortablePath(project As ReviewProject, document As ReviewDocument) As System.String
            Dim relative As System.String = ContainedRelativePath(project.SourceRoot, document.SourcePath)
            If relative.Length > 0 AndAlso relative <> "." Then Return relative
            If System.IO.Path.IsPathRooted(If(document.SourcePath, "")) Then Return ""
            If Not System.IO.Path.IsPathRooted(If(document.RelativePath, "")) AndAlso ReviewContract.IsSafeRelativePath(document.RelativePath) AndAlso document.RelativePath <> "." Then Return document.RelativePath.Replace("\", "/")
            Return ""
        End Function

        ' The Excel adapter calls this after confirmation/path selection and before
        ' its native write, so that one saved file includes the portable metadata.
        Public Sub PrepareWorkbookSavePath(targetPath As System.String, Optional forceCheckpoint As System.Boolean = False)
            EnsureAvailable()
            Dim folder As System.String = WorkbookDirectory(targetPath)
            Dim relative As System.String = ContainedRelativePath(folder, _project.SourceRoot)
            If relative = _project.SourceRootRelativePath AndAlso Not forceCheckpoint Then Return
            Dim candidate As ReviewProject = ReviewContract.Clone(_project)
            candidate.SourceRootRelativePath = relative
            Commit(candidate, "portable_source_root",
                If(relative.Length > 0, "Saved a source-root reference relative to the workbook folder.",
                    "The source folder lies outside the workbook folder. Keep the originals available or save the workbook beside its source folder."), targetWorkbookPath:=targetPath)
        End Sub

        Private Function PortabilityState() As System.Object
            Dim external As System.Int32 = 0
            For Each document As ReviewDocument In _project.Documents
                If Not document.Archived AndAlso DocumentPortablePath(_project, document).Length = 0 Then external += 1
            Next
            Dim relativeRoot As System.String = _project.SourceRootRelativePath
            Dim workbook As System.String = If(_callbacks.WorkbookPath Is Nothing, "", _callbacks.WorkbookPath.Invoke())
            If workbook.Length > 0 Then relativeRoot = ContainedRelativePath(WorkbookDirectory(workbook), _project.SourceRoot)
            Dim message As System.String = _portableNotice
            If message.Length = 0 Then
                If relativeRoot.Length > 0 AndAlso external = 0 Then
                    message = "Share the entire workbook folder and its source subfolders. The workbook also keeps the saved extracts and review history."
                ElseIf external > 0 Then
                    message = external.ToString(System.Globalization.CultureInfo.InvariantCulture) & " source(s) lie outside the package folder. Copy them into that folder and use Link source before sharing."
                Else
                    message = "Save the workbook in the document folder to create a portable folder package."
                End If
            End If
            If _portableNotice.Length > 0 AndAlso external > 0 Then message &= " " & external.ToString(System.Globalization.CultureInfo.InvariantCulture) & " source(s) remain outside the package; copy them into it and use Link source before sharing."
            Return New With {.relativeRoot = relativeRoot, .externalDocumentCount = external, .message = message}
        End Function

        Private Async Function RelocatePortableSourcesAsync(cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            If System.String.IsNullOrWhiteSpace(_project.SourceRootRelativePath) OrElse _callbacks.WorkbookPath Is Nothing Then Return
            Dim workbookPath As System.String = _callbacks.WorkbookPath.Invoke()
            If System.String.IsNullOrWhiteSpace(workbookPath) Then Return
            Dim directory As System.String = WorkbookDirectory(workbookPath)
            If directory.Length = 0 Then Return
            Dim root As System.String = System.IO.Path.GetFullPath(System.IO.Path.Combine(directory, _project.SourceRootRelativePath.Replace("/"c, System.IO.Path.DirectorySeparatorChar)))
            If ContainedRelativePath(directory, root).Length = 0 Then Throw New System.IO.InvalidDataException("The portable source root leaves the workbook folder.")
            Dim rootChanged As System.Boolean = Not System.String.Equals(root.TrimEnd(System.IO.Path.DirectorySeparatorChar), _project.SourceRoot.TrimEnd(System.IO.Path.DirectorySeparatorChar), System.StringComparison.OrdinalIgnoreCase)
            If Not rootChanged AndAlso Not _project.Documents.Exists(
                Function(item As ReviewDocument) DocumentPortablePath(_project, item).Length > 0 AndAlso
                    (item.ExtractionStatus = "missing" OrElse item.ExtractionStatus = "error" OrElse item.ExtractionStatus = "changed")) Then Return
            ' Adopt the expected package root even if its folder has not arrived yet.
            ' Later edits must retain the relative hint, and a later open can verify
            ' restored originals against their hashes without repeating extraction.
            Dim rootMissing As System.Boolean = Not System.IO.Directory.Exists(root)
            Dim candidate As ReviewProject = ReviewContract.Clone(_project)
            Dim changed As System.Boolean = rootChanged
            Dim checks As New System.Collections.Generic.Dictionary(Of System.String, SourceFileCheck)(System.StringComparer.OrdinalIgnoreCase)
            Dim matched As System.Int32 = 0
            Dim unavailable As System.Int32 = 0
            For Each document As ReviewDocument In candidate.Documents
                cancellation.ThrowIfCancellationRequested()
                Dim relative As System.String = DocumentPortablePath(_project, document)
                If relative.Length = 0 Then Continue For
                Dim relocated As System.String = System.IO.Path.GetFullPath(System.IO.Path.Combine(root, relative.Replace("/"c, System.IO.Path.DirectorySeparatorChar)))
                If ContainedRelativePath(root, relocated).Length = 0 Then Throw New System.IO.InvalidDataException("A document path leaves its portable source folder.")
                Dim check As SourceFileCheck = Nothing
                If Not checks.TryGetValue(relocated, check) Then
                    check = Await CheckSourceFileAsync(relocated, cancellation)
                    checks(relocated) = check
                End If
                Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
                EnsureAvailable()
                If document.SourcePath <> relocated OrElse document.RelativePath <> relative Then changed = True
                document.SourcePath = relocated
                document.RelativePath = relative
                Dim previousStatus As System.String = document.ExtractionStatus
                Dim previousMessage As System.String = document.ExtractionMessage
                If check.Status.Length = 0 AndAlso check.Hash = document.SourceHash AndAlso document.Segments.Count > 0 Then
                    matched += 1
                    If document.StoredExtractionStatus = "complete" OrElse document.StoredExtractionStatus = "partial" Then
                        document.ExtractionStatus = document.StoredExtractionStatus
                        document.ExtractionMessage = document.StoredExtractionMessage
                    End If
                Else
                    unavailable += 1
                    document.ExtractionStatus = If(check.Status.Length > 0, check.Status, "changed")
                    document.ExtractionMessage = If(check.Status.Length > 0, check.Message, "The packaged file differs from the saved source. Refresh or relink it before generating new results; saved quotations remain available.")
                    If rootChanged OrElse previousStatus <> document.ExtractionStatus OrElse previousMessage <> document.ExtractionMessage Then ReviewContract.Invalidate(document)
                End If
                If previousStatus <> document.ExtractionStatus OrElse previousMessage <> document.ExtractionMessage Then changed = True
            Next
            candidate.SourceRoot = root
            For Each report As ReviewReportFile In candidate.ReportFiles
                If report.RelativePath.Length > 0 AndAlso ReviewContract.IsSafeRelativePath(report.RelativePath) Then
                    report.Path = System.IO.Path.GetFullPath(System.IO.Path.Combine(root, report.RelativePath.Replace("/"c, System.IO.Path.DirectorySeparatorChar)))
                End If
            Next
            _portableNotice = If(rootMissing,
                "The expected source subfolder is missing beside this workbook. Its relative location has been retained: copy the source subfolder beside the workbook and reopen the review. Saved extracts remain available.",
                "Opened the package at its current location: " & matched.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                    " source(s) match the saved hashes; " & unavailable.ToString(System.Globalization.CultureInfo.InvariantCulture) & " source(s) need attention. Saved extracts remain in the workbook.")
            If changed Then Commit(candidate, "package_relocated", _portableNotice)
        End Function

        Private Sub EnsureAvailable()
            If _closed OrElse Me.IsDisposed OrElse _lifetime.IsCancellationRequested Then Throw New System.OperationCanceledException("The review or Office host has closed.")
            If Me.InvokeRequired Then Throw New System.InvalidOperationException("Review changes must run on the owning Office UI thread.")
            _callbacks.EnsureAvailable.Invoke()
        End Sub

        Private Async Sub InitializeAsync(sender As System.Object, args As System.EventArgs)
            Try
                EnsureAvailable()
                Await RelocatePortableSourcesAsync(_lifetime.Token)
                Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
                EnsureAvailable()
                System.IO.Directory.CreateDirectory(_assetRoot)
                Using stream As System.IO.Stream = GetType(TabularReviewForm).Assembly.GetManifestResourceStream("RedInk.TabularReview.html")
                    If stream Is Nothing Then Throw New System.IO.FileNotFoundException("The embedded Tabular Review interface is missing. Rebuild Red Ink with its HTML resource.")
                    Using reader As New System.IO.StreamReader(stream, System.Text.Encoding.UTF8)
                        System.IO.File.WriteAllText(System.IO.Path.Combine(_assetRoot, "index.html"), reader.ReadToEnd(), New System.Text.UTF8Encoding(False))
                    End Using
                End Using
                _allowedPaths.Add("/index.html")
                Dim environment As Microsoft.Web.WebView2.Core.CoreWebView2Environment = Await Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(
                    Nothing, Global.SharedLibrary.SharedLibrary.SharedMethods.GetWebView2UserDataFolder())
                Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
                EnsureAvailable()
                Await Global.SharedLibrary.SharedLibrary.WebView2DpiHost.EnsureInitializedAsync(_web, environment)
                Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
                EnsureAvailable()
                Dim core As Microsoft.Web.WebView2.Core.CoreWebView2 = _web.CoreWebView2
                core.Settings.AreDevToolsEnabled = False
                core.Settings.AreHostObjectsAllowed = False
                core.Settings.AreDefaultContextMenusEnabled = False
                core.Settings.IsWebMessageEnabled = True
                core.Settings.IsStatusBarEnabled = False
                core.SetVirtualHostNameToFolderMapping(New System.Uri(_origin).Host, _assetRoot, Microsoft.Web.WebView2.Core.CoreWebView2HostResourceAccessKind.DenyCors)
                core.AddWebResourceRequestedFilter("*", Microsoft.Web.WebView2.Core.CoreWebView2WebResourceContext.All)
                AddHandler core.WebResourceRequested,
                    Sub(s As System.Object, e As Microsoft.Web.WebView2.Core.CoreWebView2WebResourceRequestedEventArgs)
                        If Not IsAllowedResource(e.Request.Uri) Then
                            e.Response = core.Environment.CreateWebResourceResponse(Nothing, 403, "Offline review", "Content-Type: text/plain")
                        End If
                    End Sub
                AddHandler core.NavigationStarting,
                    Sub(s As System.Object, e As Microsoft.Web.WebView2.Core.CoreWebView2NavigationStartingEventArgs)
                        If Not System.String.Equals(e.Uri, _origin & "index.html", System.StringComparison.Ordinal) Then e.Cancel = True
                    End Sub
                AddHandler core.FrameNavigationStarting,
                    Sub(s As System.Object, e As Microsoft.Web.WebView2.Core.CoreWebView2NavigationStartingEventArgs)
                        If e.Uri <> "about:blank" AndAlso Not IsAllowedResource(e.Uri) Then e.Cancel = True
                    End Sub
                AddHandler core.NewWindowRequested, Sub(s As System.Object, e As Microsoft.Web.WebView2.Core.CoreWebView2NewWindowRequestedEventArgs) e.Handled = True
                AddHandler core.PermissionRequested, Sub(s As System.Object, e As Microsoft.Web.WebView2.Core.CoreWebView2PermissionRequestedEventArgs) e.State = Microsoft.Web.WebView2.Core.CoreWebView2PermissionState.Deny
                AddHandler core.DownloadStarting, Sub(s As System.Object, e As Microsoft.Web.WebView2.Core.CoreWebView2DownloadStartingEventArgs) e.Cancel = True
                AddHandler core.WebMessageReceived, AddressOf ReceiveAsync
                AddHandler core.ProcessFailed,
                    Sub(s As System.Object, e As Microsoft.Web.WebView2.Core.CoreWebView2ProcessFailedEventArgs)
                        _ready = False
                        If _operation IsNot Nothing Then _operation.Cancel()
                        Global.SharedLibrary.SharedLibrary.SharedMethods.LogWebView2ProcessFailed("TabularReview", e.ProcessFailedKind.ToString(), e.ExitCode.ToString())
                        ShowMessage("The review browser stopped. Answers already acknowledged remain in Excel. Close and reopen Tabular Review; re-enter any text you had not saved.")
                    End Sub
                core.Navigate(_origin & "index.html")
            Catch ex As System.OperationCanceledException
                ' The host owns shutdown; do not revive UI or write a workbook afterwards.
            Catch ex As System.Exception
                If Not _closed AndAlso Not Me.IsDisposed Then ShowMessage("Tabular Review could not open. Check the Microsoft Edge WebView2 Runtime and rebuild Red Ink if necessary. " & ex.Message)
            End Try
        End Sub

        Private Function IsAllowedResource(value As System.String) As System.Boolean
            Dim uri As System.Uri = Nothing
            If Not System.Uri.TryCreate(value, System.UriKind.Absolute, uri) Then Return False
            If uri.Scheme = "data" Then Return value.StartsWith("data:image/", System.StringComparison.OrdinalIgnoreCase)
            Dim expected As New System.Uri(_origin)
            Return uri.Scheme = "https" AndAlso uri.Host = expected.Host AndAlso uri.IsDefaultPort AndAlso uri.UserInfo.Length = 0 AndAlso _allowedPaths.Contains(uri.AbsolutePath)
        End Function

        Private Async Sub ReceiveAsync(sender As System.Object, e As Microsoft.Web.WebView2.Core.CoreWebView2WebMessageReceivedEventArgs)
            If _closed OrElse e.Source <> _origin & "index.html" Then Return
            Dim requestId As System.String = ""
            Dim command As System.String = ""
            Dim enteredOperation As System.Boolean = False
            Dim selectedFolder As System.String = Nothing
            Dim saveResult As ReviewWorkbookSaveResult = Nothing
            Dim reportResult As ReviewReportExportResult = Nothing
            Dim consistencyPlan As ReviewConsistencyRun = Nothing
            Dim templateResult As System.Object = Nothing
            Dim criteriaResult As System.Object = Nothing
            Dim initialRevision As System.Int64 = _project.Revision
            Try
                Dim raw As System.String = e.WebMessageAsJson
                If raw.Length > 16000000 Then Throw New System.IO.InvalidDataException("The edit exceeds the browser message limit. Split this edit into smaller columns; no result was changed.")
                Dim request As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(raw)
                requestId = TextField(request, "request")
                If requestId.Length > 128 Then Throw New System.IO.InvalidDataException("Invalid request identity.")
                command = TextField(request, "command")
                If command = "flushed" Then
                    If requestId = _flushId AndAlso _flush IsNot Nothing Then _flush.TrySetResult(BooleanField(request, "ok"))
                    Return
                End If
                If command = "rememberReviewer" Then
                    _sessionReviewer = TextField(request, "name").Trim()
                    For Each window As System.Windows.Forms.Form In System.Windows.Forms.Application.OpenForms
                        Dim review As TabularReviewForm = TryCast(window, TabularReviewForm)
                        If review IsNot Nothing AndAlso review IsNot Me AndAlso review._ready Then review.Post(New With {.command = "sessionReviewer", .name = _sessionReviewer})
                    Next
                    Return
                End If
                EnsureAvailable()
                Select Case command
                    Case "ready"
                        _ready = True
                        Post(New With {.command = "sessionReviewer", .name = _sessionReviewer})
                        SendState()
                    Case "cancel"
                        If _operation IsNot Nothing Then _operation.Cancel()
                        _message = "Stopping after the active step; completed results remain in Excel."
                        SendState()
                    Case "source"
                        Await ShowSourceAsync(request, requestId)
                    Case "openOriginal"
                        OpenOriginal(FindDocument(TextField(request, "documentId")))
                    Case "close"
                        Me.BeginInvoke(New System.Action(Sub() Me.Close()))
                    Case "autoSaveWorkbook"
                        If _busy OrElse _savingWorkbook OrElse _closing Then
                            saveResult = New ReviewWorkbookSaveResult With {.Status = "skipped", .Message = "Autosave will resume when the review is idle."}
                        Else
                            CheckRevision(request)
                            saveResult = SaveWorkbookCore(False, True)
                        End If
                    Case Else
                        If _busy OrElse _savingWorkbook Then Throw New System.InvalidOperationException("A review operation is running. Stop it or wait before editing the review.")
                        CheckRevision(request)
                        Select Case command
                            Case "saveSettings"
                                SaveSettings(request)
                            Case "chooseFolder"
                                selectedFolder = ChooseFolder(TextField(request, "currentPath", _project.SourceRoot))
                            Case "saveCriteria"
                                SaveCriteria(request)
                                ' A complete receipt keeps host-assigned column identities authoritative
                                ' even if a separate browser state message cannot be accepted.
                                criteriaResult = CriteriaSaveReceipt()
                            Case "saveReportDefinition", "deleteReportDefinition"
                                SaveReportDefinition(request, command = "deleteReportDefinition")
                            Case "reorderCriteria"
                                ReorderCriteria(request)
                            Case "saveFlagColors"
                                SaveFlagColors(request)
                            Case "removeDocuments", "restoreDocuments"
                                SetDocumentsArchived(request, command = "removeDocuments")
                            Case "saveFinding"
                                SaveFinding(request)
                            Case "removeFindings", "restoreFindings"
                                SetFindingsArchived(request, command = "removeFindings")
                            Case "saveAnswer"
                                SaveAnswer(request)
                            Case "dissolveGroup"
                                DissolveGroup(request)
                            Case "assignGroup", "assignCategory", "groupByFolder"
                                AssignReviewScope(request, command)
                            Case "saveDocument"
                                SaveDocument(request)
                            Case "view"
                                SaveView(request)
                            Case "saveAutosave"
                                _filenameDeclined = False
                                Dim candidate As ReviewProject = ReviewContract.Clone(_project)
                                candidate.View.AutoSave = BooleanField(request, "enabled", True)
                                Commit(candidate, "autosave_preference", If(candidate.View.AutoSave, "Enabled review autosave.", "Disabled review autosave."), actor:=RequireReviewer(request))
                            Case "saveWorkbook", "saveWorkbookAs"
                                saveResult = SaveWorkbookCore(command = "saveWorkbookAs", False)
                            Case "reloadExcel"
                                Dim refreshed As ReviewProject = _callbacks.ReloadProject.Invoke()
                                If refreshed Is Nothing Then Throw New System.InvalidOperationException("This workbook no longer contains the review.")
                                _project = refreshed
                                _persistedViewJson = ReviewContract.Serialize(_project.View)
                                _pendingMetadata = False
                                _persistenceMessage = ""
                                _message = "Loaded direct Excel edits. Changed questions and answers require review."
                                SendState()
                            Case "scanFolder", "importSelection", "refreshDocument", "relinkDocument", "generate", "generateAll", "classifyDocuments", "reviewGroups", "planContradictions", "checkContradictions", "exportReport",
                                 "listColumnTemplates", "loadColumnTemplate", "importColumnTemplates", "saveColumnTemplate", "exportColumnTemplate"
                                _busy = True
                                enteredOperation = True
                                _operation = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(_lifetime.Token)
                                _operationDone = New System.Threading.Tasks.TaskCompletionSource(Of System.Boolean)(System.Threading.Tasks.TaskCreationOptions.RunContinuationsAsynchronously)
                                SendState()
                                If command = "scanFolder" Then
                                    Await ScanFolderAsync(request, _operation.Token)
                                ElseIf command = "importSelection" Then
                                    Dim importedIds As System.Collections.Generic.List(Of System.String) = ImportSelection()
                                    Await LoadRegisteredSourcesAsync(importedIds, _operation.Token)
                                ElseIf command = "refreshDocument" OrElse command = "relinkDocument" Then
                                    Await RefreshDocumentAsync(TextField(request, "documentId"), command = "relinkDocument", _operation.Token)
                                ElseIf command = "classifyDocuments" Then
                                    Await ClassifyDocumentsAsync(request, _operation.Token)
                                ElseIf command = "reviewGroups" Then
                                    Await ReviewGroupsAsync(request, _operation.Token)
                                ElseIf command = "planContradictions" Then
                                    consistencyPlan = Await PlanConsistencyAsync(request, _operation.Token)
                                ElseIf command = "checkContradictions" Then
                                    Await CheckConsistencyAsync(request, _operation.Token)
                                ElseIf command = "exportReport" Then
                                    reportResult = Await ExportReportAsync(request, _operation.Token)
                                ElseIf command = "listColumnTemplates" Then
                                    templateResult = Await ListColumnTemplatesAsync("list", _operation.Token)
                                ElseIf command = "loadColumnTemplate" Then
                                    templateResult = Await LoadColumnTemplateAsync(request, _operation.Token)
                                ElseIf command = "importColumnTemplates" Then
                                    templateResult = Await ImportColumnTemplatesAsync(_operation.Token)
                                ElseIf command = "saveColumnTemplate" Then
                                    templateResult = Await SaveColumnTemplateAsync(request, _operation.Token)
                                ElseIf command = "exportColumnTemplate" Then
                                    templateResult = Await ExportColumnTemplateAsync(request, _operation.Token)
                                Else
                                    Await GenerateAsync(request, command = "generateAll", _operation.Token)
                                End If
                            Case Else
                                Throw New System.IO.InvalidDataException("Unknown review command.")
                        End Select
                End Select
                Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
                SendAck(requestId, True, "", selectedFolder, saveResult, reportResult, consistencyPlan, templateResult, criteriaResult)
            Catch ex As System.OperationCanceledException
                If command = "source" Then
                    SendAck(requestId, False, "The selected source changed; the new source request takes precedence.")
                ElseIf Not _closed AndAlso Not _hostLifetime.IsCancellationRequested Then
                    _message = "Stopped. Completed changes remain in Excel; save the workbook to keep them."
                    SendAck(requestId, False, _message)
                End If
            Catch ex As System.Exception
                If Not _closed AndAlso Not _hostLifetime.IsCancellationRequested Then
                    _message = ex.Message
                    System.Diagnostics.Trace.WriteLine("[TabularReview] command=" & command & "; " & ex.ToString())
                    If command = "saveWorkbook" OrElse command = "saveWorkbookAs" OrElse command = "autoSaveWorkbook" Then
                        _persistenceMessage = ex.Message
                        saveResult = New ReviewWorkbookSaveResult With {.Status = "failed", .Message = ex.Message}
                    End If
                    SendAck(requestId, False, ex.Message, saveResult:=saveResult)
                End If
            Finally
                If enteredOperation Then
                    _busy = False
                    If _operation IsNot Nothing Then _operation.Dispose()
                    _operation = Nothing
                    If _operationDone IsNot Nothing Then _operationDone.TrySetResult(True)
                End If
                If Not _closed AndAlso _ready AndAlso (command <> "autoSaveWorkbook" OrElse _project.Revision <> initialRevision) Then SendState()
            End Try
        End Sub

        Private Sub CheckRevision(request As Newtonsoft.Json.Linq.JObject)
            Dim token As Newtonsoft.Json.Linq.JToken = request("revision")
            Dim revision As System.Int64
            If token Is Nothing OrElse Not System.Int64.TryParse(token.ToString(), revision) OrElse revision <> _project.Revision Then
                Throw New System.InvalidOperationException("The review has changed since this edit began. Your draft remains in the editor. Reload the latest state before saving it again.")
            End If
        End Sub

        Private Shared Function TextField(request As Newtonsoft.Json.Linq.JObject, field As System.String, Optional defaultValue As System.String = "") As System.String
            Dim token As Newtonsoft.Json.Linq.JToken = request(field)
            If token Is Nothing OrElse token.Type = Newtonsoft.Json.Linq.JTokenType.Null Then Return defaultValue
            If token.Type <> Newtonsoft.Json.Linq.JTokenType.String Then Throw New System.IO.InvalidDataException("Invalid text field: " & field)
            Return token.ToString()
        End Function

        Private Shared Function BooleanField(request As Newtonsoft.Json.Linq.JObject, field As System.String, Optional defaultValue As System.Boolean = False) As System.Boolean
            Dim token As Newtonsoft.Json.Linq.JToken = request(field)
            If token Is Nothing Then Return defaultValue
            If token.Type <> Newtonsoft.Json.Linq.JTokenType.Boolean Then Throw New System.IO.InvalidDataException("Invalid checkbox field: " & field)
            Return token.ToObject(Of System.Boolean)()
        End Function

        Private Function FindDocument(id As System.String, Optional project As ReviewProject = Nothing) As ReviewDocument
            ReviewContract.RequireId(id)
            Dim document As ReviewDocument = If(project, _project).Documents.Find(Function(item As ReviewDocument) item.Id = id)
            If document Is Nothing OrElse document.Archived Then Throw New System.IO.InvalidDataException("This document no longer belongs to the active review.")
            Return document
        End Function

        Private Function FindCriterion(id As System.String, Optional project As ReviewProject = Nothing) As ReviewCriterion
            ReviewContract.RequireId(id)
            Dim criterion As ReviewCriterion = If(project, _project).Criteria.Find(Function(item As ReviewCriterion) item.Id = id)
            If criterion Is Nothing OrElse criterion.Archived Then Throw New System.IO.InvalidDataException("This column no longer belongs to the active review.")
            Return criterion
        End Function

        Private Sub Commit(candidate As ReviewProject, action As System.String, detail As System.String,
                           Optional documentId As System.String = "", Optional criterionId As System.String = "",
                           Optional targetWorkbookPath As System.String = Nothing, Optional actor As System.String = "")
            EnsureAvailable()
            Dim workbookPath As System.String = targetWorkbookPath
            If workbookPath Is Nothing Then workbookPath = If(_callbacks.WorkbookPath Is Nothing, "", _callbacks.WorkbookPath.Invoke())
            candidate.SourceRootRelativePath = If(System.String.IsNullOrWhiteSpace(workbookPath), "",
                ContainedRelativePath(WorkbookDirectory(workbookPath), candidate.SourceRoot))
            For Each report As ReviewReportFile In candidate.ReportFiles
                report.RelativePath = ContainedRelativePath(candidate.SourceRoot, report.Path)
            Next
            SynchronizeFindingFreshness(candidate)
            ReviewContract.AddAudit(candidate, action, detail, documentId, criterionId, actor)
            ReviewContract.Validate(candidate)
            _callbacks.SaveProject.Invoke(candidate)
            _project = candidate
            _persistedViewJson = ReviewContract.Serialize(_project.View)
            _pendingMetadata = False
            _persistenceMessage = ""
            _message = If(_project.View.AutoSave, "Updated in Excel. Autosave keeps the named workbook up to date.", "Updated in Excel. Choose Save to keep the review on disk.")
            SendState()
        End Sub

        Private Shared Function RequireReviewer(request As Newtonsoft.Json.Linq.JObject) As System.String
            Dim name As System.String = TextField(request, "actor").Trim()
            If name.Length = 0 AndAlso BooleanField(request, "automatic") Then Return System.Environment.UserName
            If name.Length = 0 Then Throw New System.InvalidOperationException("Enter your name in Reviewer before saving manual changes.")
            Return name
        End Function

        Private Function HasPendingView() As System.Boolean
            Return _pendingMetadata OrElse Not System.String.Equals(_persistedViewJson, ReviewContract.Serialize(_project.View), System.StringComparison.Ordinal)
        End Function

        Private Function HasReviewContent() As System.Boolean
            Return _project.Revision > 0 OrElse _project.Criteria.Count > 0 OrElse _project.Documents.Count > 0
        End Function

        Private Sub PersistPendingView(Optional includeEmptyReview As System.Boolean = False)
            EnsureAvailable()
            If Not includeEmptyReview AndAlso Not HasReviewContent() Then Return
            ' Layout/schema migration and a native Excel Save As can reach this
            ' path without a semantic edit. Keep the portable hint bound to the
            ' actual workbook before persisting that pending snapshot as well.
            Dim workbookPath As System.String = If(_callbacks.WorkbookPath Is Nothing, "", _callbacks.WorkbookPath.Invoke())
            Dim relativeRoot As System.String = ContainedRelativePath(WorkbookDirectory(workbookPath), _project.SourceRoot)
            If _project.SourceRootRelativePath <> relativeRoot Then
                _project.SourceRootRelativePath = relativeRoot
                _pendingMetadata = True
            End If
            If includeEmptyReview OrElse _project.Revision = 0 OrElse HasPendingView() Then
                ' Keep the same optimistic workbook/fingerprint checks as every edit.
                ' A layout-only save must never overwrite independent Excel edits.
                _callbacks.SaveProject.Invoke(_project)
                _persistedViewJson = ReviewContract.Serialize(_project.View)
                _pendingMetadata = False
            End If
        End Sub

        Private Function SaveWorkbookCore(saveAs As System.Boolean, automatic As System.Boolean) As ReviewWorkbookSaveResult
            EnsureAvailable()
            If _savingWorkbook Then Return New ReviewWorkbookSaveResult With {.Status = "skipped", .Message = "A workbook save is already running."}
            If automatic Then
                If Not _project.View.AutoSave Then Return New ReviewWorkbookSaveResult With {.Status = "skipped", .Message = "Autosave is off."}
                If _filenameDeclined AndAlso _callbacks.WorkbookPath IsNot Nothing AndAlso System.String.IsNullOrWhiteSpace(_callbacks.WorkbookPath.Invoke()) Then Return New ReviewWorkbookSaveResult With {.Status = "skipped", .Message = "Choose Save to select a filename."}
                If Not HasReviewContent() Then Return New ReviewWorkbookSaveResult With {.Status = "skipped", .Message = "Add documents or columns to begin your review."}
                Dim state As ReviewWorkbookSaveState = If(_callbacks.WorkbookSaveState Is Nothing, Nothing, _callbacks.WorkbookSaveState.Invoke())
                If state Is Nothing OrElse Not state.CanAutoSave Then
                    Return New ReviewWorkbookSaveResult With {.Status = "skipped", .Message = If(state Is Nothing, "Choose Save to check this workbook's save location.", state.Message)}
                End If
            End If
            Dim callback As System.Func(Of ReviewWorkbookSaveResult) = If(automatic, _callbacks.AutoSaveWorkbook, If(saveAs, _callbacks.SaveWorkbookAs, _callbacks.SaveWorkbook))
            If callback Is Nothing Then Throw New System.InvalidOperationException("The workbook host does not support this save action. Rebuild the Excel and SharedLibrary projects together.")
            _savingWorkbook = True
            Try
                PersistPendingView(Not automatic)
                ' The native adapter owns file dialogs and workbook identity. A cancelled
                ' picker is different from an Excel event refusing or changing a save.
                Dim result As ReviewWorkbookSaveResult
                Using Global.SharedLibrary.SharedLibrary.SharedMethods.PushDialogOwner(Me)
                    result = callback.Invoke()
                End Using
                If result Is Nothing OrElse Not ReviewContract.Contains(New System.String() {"saved", "cancelled", "skipped", "failed"}, result.Status) Then
                    Throw New System.IO.InvalidDataException("Excel returned an invalid save result. Check the workbook before closing it.")
                End If
                If result.Status = "cancelled" Then _filenameDeclined = True
                If result.Status = "saved" Then _filenameDeclined = False
                If result.Status = "failed" Then Throw New System.IO.IOException(result.Message)
                _persistenceMessage = result.Message
                If result.Status = "saved" Then
                    _lastSavedUtc = System.DateTime.UtcNow.ToString("o")
                    _message = If(automatic, "Workbook autosaved.", "Workbook saved. You can reopen it to resume this review.")
                ElseIf Not automatic Then
                    _message = result.Message
                End If
                Return result
            Finally
                _savingWorkbook = False
            End Try
        End Function

        Private Function PersistenceState() As System.Object
            Dim state As ReviewWorkbookSaveState = Nothing
            Try
                If _callbacks.WorkbookSaveState IsNot Nothing Then state = _callbacks.WorkbookSaveState.Invoke()
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine("[TabularReview:save_state] " & ex.ToString())
                state = New ReviewWorkbookSaveState With {.Message = "Excel could not report the save status. Choose Save to check the workbook."}
            End Try
            If state Is Nothing Then state = New ReviewWorkbookSaveState With {.Message = "Choose Save to check the workbook."}
            Dim saved As System.Boolean = state.Saved AndAlso Not HasPendingView()
            Dim message As System.String = If(_persistenceMessage.Length > 0, _persistenceMessage, state.Message)
            If Not state.CanAutoSave Then message = state.Message
            If saved AndAlso _persistenceMessage.Length = 0 Then message = "Workbook saved."
            Return New With {.autoSaveEnabled = _project.View.AutoSave, .canAutoSave = state.CanAutoSave,
                .saved = saved, .message = message, .lastSavedUtc = _lastSavedUtc}
        End Function

        Private Sub SaveSettings(request As Newtonsoft.Json.Linq.JObject)
            Dim candidate As ReviewProject = ReviewContract.Clone(_project)
            candidate.Title = TextField(request, "title").Trim()
            If candidate.Title.Length = 0 Then Throw New System.IO.InvalidDataException("Enter a review title.")
            candidate.OutputLanguage = TextField(request, "outputLanguage").Trim()
            Dim modelKey As System.String = TextField(request, "modelKey")
            If modelKey <> candidate.ModelKey AndAlso Not _models.Exists(Function(item As ReviewModelChoice) item.Key = modelKey) Then Throw New System.IO.InvalidDataException("Choose an available model.")
            candidate.ModelKey = modelKey
            candidate.SourceRoot = TextField(request, "sourceRoot").Trim()
            If candidate.SourceRoot.Length > 0 Then candidate.SourceRoot = NormalizeLocalPath(candidate.SourceRoot)
            candidate.IncludeSubfolders = BooleanField(request, "includeSubfolders", True)
            candidate.UseOcr = BooleanField(request, "useOcr")
            Commit(candidate, "settings", "Updated review settings and selected model.", actor:=RequireReviewer(request))
        End Sub

        ' Picking a folder must not write the workbook. The UI places the selected
        ' path in the active dialog and saves it only when the user confirms.
        Private Function ChooseFolder(currentPath As System.String) As System.String
            Using dialog As New System.Windows.Forms.FolderBrowserDialog With {.Description = "Choose the local or UNC folder containing the review documents", .ShowNewFolderButton = False}
                If System.IO.Directory.Exists(currentPath) Then dialog.SelectedPath = currentPath
                Using Global.SharedLibrary.SharedLibrary.SharedMethods.PushDialogOwner(Me)
                    If dialog.ShowDialog(Me) <> System.Windows.Forms.DialogResult.OK Then Return Nothing
                End Using
                Return NormalizeLocalPath(dialog.SelectedPath)
            End Using
        End Function

        Private Shared Function NormalizeLocalPath(value As System.String) As System.String
            Dim expanded As System.String = Global.SharedLibrary.SharedLibrary.SharedMethods.ExpandEnvironmentVariables(value)
            If Not System.IO.Path.IsPathRooted(expanded) OrElse expanded.IndexOf("://", System.StringComparison.Ordinal) >= 0 Then Throw New System.IO.InvalidDataException("Use an absolute local path or a UNC path.")
            Return System.IO.Path.GetFullPath(expanded)
        End Function

        Private Function CriteriaSaveReceipt() As System.Object
            Dim documents As New System.Collections.Generic.List(Of System.Object)()
            For Each document As ReviewDocument In _project.Documents
                Dim statuses As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.Ordinal)
                For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, ReviewAnswer) In document.Answers
                    statuses.Add(pair.Key, pair.Value.Status)
                Next
                documents.Add(New With {.documentId = document.Id, .reviewStatus = document.ReviewStatus,
                    .lastReviewedUtc = document.LastReviewedUtc, .answerStatuses = statuses})
            Next
            Return New With {.projectId = _project.Id, .revision = _project.Revision, .criteria = _project.Criteria,
                .columnWidths = _project.View.ColumnWidths, .flagColors = _project.FlagColors, .reportDefinitions = _project.ReportDefinitions, .documentReviewStates = documents}
        End Function

        Private Sub SaveCriteria(request As Newtonsoft.Json.Linq.JObject)
            Dim items As Newtonsoft.Json.Linq.JArray = TryCast(request("criteria"), Newtonsoft.Json.Linq.JArray)
            If items Is Nothing Then Throw New System.IO.InvalidDataException("Missing column definitions.")
            Dim candidate As ReviewProject = ReviewContract.Clone(_project)
            Dim definitions As New System.Collections.Generic.List(Of ReviewCriterion)()
            Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            For Each item As Newtonsoft.Json.Linq.JToken In items
                Dim definition As Newtonsoft.Json.Linq.JObject = TryCast(item, Newtonsoft.Json.Linq.JObject)
                If definition Is Nothing Then Throw New System.IO.InvalidDataException("Invalid column definition.")
                Dim id As System.String = TextField(definition, "id")
                Dim previous As ReviewCriterion = candidate.Criteria.Find(Function(existing As ReviewCriterion) existing.Id = id)
                If previous Is Nothing AndAlso id.Length > 0 Then Throw New System.IO.InvalidDataException("A new column must not reuse an unknown identity.")
                Dim criterion As New ReviewCriterion With {
                    .Id = If(previous IsNot Nothing, previous.Id, System.Guid.NewGuid().ToString("N")),
                    .Title = TextField(definition, "title"), .Instruction = TextField(definition, "instruction"),
                    .Kind = TextField(definition, "kind", "analysis"), .Archived = BooleanField(definition, "archived"),
                    .Category = TextField(definition, "category").Trim(), .FlagScale = TextField(definition, "flagScale"),
                    .FlagConditions = If(definition("flagConditions") Is Nothing, New System.Collections.Generic.Dictionary(Of System.String, System.String)(), definition("flagConditions").ToObject(Of System.Collections.Generic.Dictionary(Of System.String, System.String))()),
                    .Options = ExactStringList(definition("options"))
                }
                If criterion.FlagConditions Is Nothing Then Throw New System.IO.InvalidDataException("Flag conditions must map labels to instructions.")
                For Each label As System.String In New System.Collections.Generic.List(Of System.String)(criterion.FlagConditions.Keys)
                    If criterion.Kind = "analysis" AndAlso Not criterion.FlagValues().Contains(label) Then criterion.FlagConditions.Remove(label)
                Next
                If Not seen.Add(criterion.Id) Then Throw New System.IO.InvalidDataException("A column appears more than once.")
                If previous IsNot Nothing Then
                    criterion.Revision = previous.Revision
                    Dim changed As System.Boolean = criterion.Category <> previous.Category OrElse criterion.FlagScale <> previous.FlagScale OrElse ReviewContract.Serialize(criterion.FlagConditions) <> ReviewContract.Serialize(previous.FlagConditions) OrElse criterion.Title <> previous.Title OrElse criterion.Instruction <> previous.Instruction OrElse criterion.Kind <> previous.Kind OrElse ReviewContract.Serialize(criterion.Options) <> ReviewContract.Serialize(previous.Options)
                    If changed Then
                        criterion.Revision += 1
                    End If
                    If changed OrElse (previous.Archived AndAlso Not criterion.Archived) Then
                        For Each document As ReviewDocument In candidate.Documents
                            If criterion.AppliesTo(document) OrElse previous.AppliesTo(document) Then ReviewContract.Invalidate(document, criterion.Id)
                        Next
                    End If
                ElseIf Not criterion.Archived Then
                    For Each document As ReviewDocument In candidate.Documents
                        If criterion.AppliesTo(document) Then ReviewContract.Invalidate(document, criterion.Id)
                    Next
                End If
                Dim width As System.Nullable(Of System.Int32) = OptionalColumnWidth(definition("width"), False)
                If width.HasValue AndAlso Not criterion.Archived Then candidate.View.ColumnWidths(criterion.Id) = width.Value
                definitions.Add(criterion)
            Next
            For Each previous As ReviewCriterion In candidate.Criteria
                If Not seen.Contains(previous.Id) Then
                    previous.Archived = True
                    definitions.Add(previous)
                End If
            Next
            candidate.Criteria = definitions
            If request("reportDefinitions") IsNot Nothing Then
                Dim reports As System.Collections.Generic.List(Of ReviewReportDefinition) = ReadReportDefinitions(request("reportDefinitions"))
                Dim columnKeys As New System.Collections.Generic.List(Of System.String)()
                For index As System.Int32 = 0 To definitions.Count - 1
                    columnKeys.Add(index.ToString(System.Globalization.CultureInfo.InvariantCulture))
                Next
                ReviewContract.ValidateReportDefinitions(reports, columnKeys)
                For Each report As ReviewReportDefinition In reports
                    For index As System.Int32 = 0 To report.Fields.Count - 1
                        Dim parts As System.String() = report.Fields(index).Split(":"c)
                        If parts.Length = 2 Then report.Fields(index) = parts(0) & ":" & definitions(System.Int32.Parse(parts(1), System.Globalization.CultureInfo.InvariantCulture)).Id
                    Next
                Next
                candidate.ReportDefinitions = reports
            End If
            Dim documentWidth As System.Nullable(Of System.Int32) = OptionalColumnWidth(request("documentColumnWidth"), True)
            If documentWidth.HasValue Then candidate.View.ColumnWidths("document") = documentWidth.Value
            If request("flagColors") IsNot Nothing Then
                Dim palette As Newtonsoft.Json.Linq.JObject = TryCast(request("flagColors"), Newtonsoft.Json.Linq.JObject)
                If palette Is Nothing Then Throw New System.IO.InvalidDataException("Template flag colors must be named color values.")
                For Each pair As Newtonsoft.Json.Linq.JProperty In palette.Properties()
                    If System.String.IsNullOrWhiteSpace(pair.Name) OrElse pair.Value.Type <> Newtonsoft.Json.Linq.JTokenType.String OrElse
                        Not System.Text.RegularExpressions.Regex.IsMatch(pair.Value.ToString(), "^#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{6})$") Then Throw New System.IO.InvalidDataException("Choose a valid color for every template flag label.")
                    ' Import changes only supplied labels; unrelated review colors remain.
                    candidate.FlagColors(pair.Name) = pair.Value.ToString()
                Next
            End If
            candidate.View.LayoutVersion = 3
            candidate.View.ColumnOrder = ActiveCriterionIds(candidate)
            Commit(candidate, "criteria", ReviewContract.Serialize(New With {.previous = _project.Criteria, .current = definitions,
                .previousWidths = _project.View.ColumnWidths, .currentWidths = candidate.View.ColumnWidths,
                .previousColors = _project.FlagColors, .currentColors = candidate.FlagColors,
                .previousReports = _project.ReportDefinitions, .currentReports = candidate.ReportDefinitions}), actor:=RequireReviewer(request))
        End Sub

        Private NotInheritable Class BrowserStateContractResolver
            Inherits Newtonsoft.Json.Serialization.DefaultContractResolver

            Public Sub New()
                Me.NamingStrategy = New Newtonsoft.Json.Serialization.CamelCaseNamingStrategy With {.ProcessDictionaryKeys = False, .OverrideSpecifiedNames = False}
            End Sub

            Protected Overrides Function CreateProperty(member As System.Reflection.MemberInfo, serialization As Newtonsoft.Json.MemberSerialization) As Newtonsoft.Json.Serialization.JsonProperty
                Dim result As Newtonsoft.Json.Serialization.JsonProperty = MyBase.CreateProperty(member, serialization)
                If member.DeclaringType Is GetType(ReviewDocument) AndAlso
                   (member.Name = NameOf(ReviewDocument.Segments) OrElse member.Name = NameOf(ReviewDocument.SourceHistory)) Then result.Ignored = True
                Return result
            End Function
        End Class

        Private Shared Function ActiveCriterionIds(project As ReviewProject) As System.Collections.Generic.List(Of System.String)
            Dim ids As New System.Collections.Generic.List(Of System.String)()
            For Each criterion As ReviewCriterion In project.Criteria
                If Not criterion.Archived Then ids.Add(criterion.Id)
            Next
            Return ids
        End Function

        Private Shared Sub ApplyCriterionOrder(project As ReviewProject, ids As System.Collections.Generic.List(Of System.String))
            Dim ordered As New System.Collections.Generic.List(Of ReviewCriterion)()
            If ids IsNot Nothing Then
                For Each id As System.String In ids
                    Dim criterion As ReviewCriterion = project.Criteria.Find(Function(item As ReviewCriterion) item.Id = id AndAlso Not item.Archived)
                    If criterion IsNot Nothing AndAlso Not ordered.Contains(criterion) Then ordered.Add(criterion)
                Next
            End If
            For Each criterion As ReviewCriterion In project.Criteria
                If Not criterion.Archived AndAlso Not ordered.Contains(criterion) Then ordered.Add(criterion)
            Next
            For Each criterion As ReviewCriterion In project.Criteria
                If criterion.Archived Then ordered.Add(criterion)
            Next
            project.Criteria = ordered
            project.View.ColumnOrder = ActiveCriterionIds(project)
        End Sub

        Private Sub ReorderCriteria(request As Newtonsoft.Json.Linq.JObject)
            Dim ids As System.Collections.Generic.List(Of System.String) = StringList(request("criterionIds"))
            Dim active As System.Collections.Generic.List(Of System.String) = ActiveCriterionIds(_project)
            If ids.Count <> active.Count OrElse ids.Exists(Function(id As System.String) Not active.Contains(id)) Then Throw New System.IO.InvalidDataException("Reorder all current criteria once. Reload the current columns and retry.")
            Dim candidate As ReviewProject = ReviewContract.Clone(_project)
            ApplyCriterionOrder(candidate, ids)
            candidate.View.LayoutVersion = 3
            Commit(candidate, "criteria_order", "Updated the shared Matrix and Criteria order; criterion identities and answers remain unchanged.", actor:=RequireReviewer(request))
        End Sub

        Private Sub SaveFlagColors(request As Newtonsoft.Json.Linq.JObject)
            Dim values As Newtonsoft.Json.Linq.JObject = TryCast(request("flagColors"), Newtonsoft.Json.Linq.JObject)
            If values Is Nothing Then Throw New System.IO.InvalidDataException("Choose the flag colors to save.")
            Dim candidate As ReviewProject = ReviewContract.Clone(_project)
            Dim colors As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.Ordinal)
            For Each propertyValue As Newtonsoft.Json.Linq.JProperty In values.Properties()
                If propertyValue.Value.Type <> Newtonsoft.Json.Linq.JTokenType.String Then Throw New System.IO.InvalidDataException("A flag color must be a text value.")
                colors(propertyValue.Name) = propertyValue.Value.ToString()
            Next
            Dim deleted As System.Collections.Generic.List(Of System.String) = ExactStringList(request("deletedLabels"))
            For Each label As System.String In deleted
                colors.Remove(label)
            Next
            candidate.FlagColors = colors
            ReviewContract.NormalizeFlagColors(candidate)
            For Each criterion As ReviewCriterion In candidate.Criteria
                If criterion.Kind = "analysis" Then
                    Dim labels = criterion.FlagValues()
                    If labels.RemoveAll(Function(label As System.String) deleted.Contains(label)) > 0 Then
                        criterion.FlagScale = System.String.Join(System.Environment.NewLine, labels)
                        For Each label As System.String In deleted
                            criterion.FlagConditions.Remove(label)
                        Next
                        criterion.Revision += 1
                        For Each document As ReviewDocument In candidate.Documents
                            Dim prior As ReviewAnswer = Nothing
                            If document.Answers.TryGetValue(criterion.Id, prior) AndAlso deleted.Contains(prior.RiskFlag) Then
                                ReviewContract.ArchiveAnswer(prior, "Deleted question flag")
                                prior.RiskFlag = ""
                                ReviewContract.Invalidate(document, criterion.Id)
                            End If
                        Next
                    End If
                End If
                If criterion.Kind <> "flag" Then Continue For
                Dim removedOptions As System.Int32 = criterion.Options.RemoveAll(Function(label As System.String) deleted.Contains(label))
                If removedOptions > 0 Then criterion.Revision += 1
                For Each document As ReviewDocument In candidate.Documents
                    Dim answer As ReviewAnswer = Nothing
                    If Not document.Answers.TryGetValue(criterion.Id, answer) Then Continue For
                    If deleted.Contains(answer.Text) Then
                        ReviewContract.ArchiveAnswer(answer, "Deleted flag label")
                        answer.Text = ""
                        answer.CriterionRevision = criterion.Revision
                        answer.Status = "draft"
                        answer.UserEdited = True
                        answer.UpdatedUtc = System.DateTime.UtcNow.ToString("o")
                        answer.UpdatedBy = RequireReviewer(request)
                        answer.ReviewedBy = ""
                        answer.ReviewedUtc = ""
                        document.UpdatedBy = answer.UpdatedBy
                        document.UpdatedUtc = answer.UpdatedUtc
                        document.ReviewedBy = ""
                        document.ReviewedUtc = ""
                        document.ReviewStatus = "in_review"
                        document.LastReviewedUtc = ""
                    End If
                    ' Removing a different option does not invalidate a retained value.
                    If removedOptions > 0 Then answer.CriterionRevision = criterion.Revision
                Next
            Next
            For Each document As ReviewDocument In candidate.Documents
                If Not document.Flags.Exists(Function(label As System.String) deleted.Contains(label)) Then Continue For
                Dim previous As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(ReviewContract.Serialize(DocumentMetadata(document)))
                document.Flags.RemoveAll(Function(label As System.String) deleted.Contains(label))
                ReviewContract.AddAudit(candidate, "document_flags_deleted", ReviewContract.Serialize(New With {.previous = previous, .current = DocumentMetadata(document)}), document.Id, actor:=RequireReviewer(request))
            Next
            If deleted.Contains(candidate.View.FlagFilter) Then candidate.View.FlagFilter = ""
            If deleted.Count > 0 Then ReviewContract.AddAudit(candidate, "flag_definitions_deleted", ReviewContract.Serialize(New With {
                .deletedLabels = deleted, .previousColors = _project.FlagColors, .previousCriteria = _project.Criteria
            }), actor:=RequireReviewer(request))
            Commit(candidate, "flag_colors", "Updated the flag palette. Deleted " & deleted.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) & " label(s) and their exact current assignments; prior answer versions and audit history remain available.", actor:=RequireReviewer(request))
        End Sub

        Private Sub SetDocumentsArchived(request As Newtonsoft.Json.Linq.JObject, archived As System.Boolean)
            Dim ids As System.Collections.Generic.List(Of System.String) = StringList(request("documentIds"))
            If ids.Count = 0 Then Throw New System.IO.InvalidDataException("Select at least one document.")
            Dim candidate As ReviewProject = ReviewContract.Clone(_project)
            For Each id As System.String In ids
                ReviewContract.RequireId(id)
                Dim document As ReviewDocument = candidate.Documents.Find(Function(item As ReviewDocument) item.Id = id)
                If document Is Nothing Then Throw New System.IO.InvalidDataException("A selected document no longer belongs to this review.")
                document.Archived = archived
            Next
            If _sourceRequest IsNot Nothing Then _sourceRequest.Cancel()
            Dim selected As ReviewDocument = candidate.Documents.Find(Function(item As ReviewDocument) item.Id = candidate.View.DocumentId AndAlso Not item.Archived)
            If selected Is Nothing Then
                selected = candidate.Documents.Find(Function(item As ReviewDocument) Not item.Archived)
                candidate.View.DocumentId = If(selected Is Nothing, "", selected.Id)
            End If
            Commit(candidate, If(archived, "documents_removed", "documents_restored"),
                If(archived, "Removed ", "Restored ") & ids.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                " document(s) in the active review. Original files, saved extracts and review history remain unchanged.", actor:=RequireReviewer(request))
        End Sub

        Private Function ResolveOperationModel(request As Newtonsoft.Json.Linq.JObject) As System.String
            Dim key As System.String = TextField(request, "modelKey").Trim()
            If key.Length = 0 Then key = _project.ModelKey
            If Not _models.Exists(Function(item As ReviewModelChoice) item.Key = key) Then Throw New System.InvalidOperationException("Select an available model for this analysis. The saved default model remains unchanged.")
            Return key
        End Function

        Private Function SelectedDocuments(request As Newtonsoft.Json.Linq.JObject) As System.Collections.Generic.List(Of ReviewDocument)
            Dim result As New System.Collections.Generic.List(Of ReviewDocument)()
            For Each id As System.String In StringList(request("documentIds"))
                result.Add(FindDocument(id))
            Next
            If result.Count = 0 Then Throw New System.InvalidOperationException("Select the documents to analyze.")
            Return result
        End Function

        Private Shared Function AllFindingEvidence(finding As ReviewFinding) As System.Collections.Generic.List(Of ReviewCrossDocumentEvidence)
            Dim result As New System.Collections.Generic.List(Of ReviewCrossDocumentEvidence)(finding.Evidence)
            For Each context As ReviewFindingContext In finding.Context
                result.AddRange(context.Evidence)
            Next
            Return result
        End Function

        Private Shared Function FindingSourcesCurrent(project As ReviewProject, finding As ReviewFinding) As System.Boolean
            If finding.Kind = "review_finding" Then
                Dim run As ReviewConsistencyRun = project.ConsistencyRuns.Find(Function(item As ReviewConsistencyRun) item.Id = finding.RunId)
                If run Is Nothing OrElse run.Documents.Count < 2 Then Return False
                For Each coverage As ReviewConsistencyDocumentCoverage In run.Documents
                    Dim source As ReviewDocument = project.Documents.Find(Function(item As ReviewDocument) item.Id = coverage.DocumentId)
                    If source Is Nothing OrElse source.Archived OrElse source.SourceHash <> coverage.SourceHash OrElse source.TextHash <> coverage.TextHash OrElse
                       (source.ExtractionStatus <> "complete" AndAlso source.ExtractionStatus <> "partial") Then Return False
                Next
            End If
            For Each citation As ReviewCrossDocumentEvidence In AllFindingEvidence(finding)
                Dim document As ReviewDocument = project.Documents.Find(Function(item As ReviewDocument) item.Id = citation.DocumentId)
                If document Is Nothing OrElse document.Archived OrElse document.SourceHash <> citation.SourceHash OrElse document.TextHash <> citation.TextHash OrElse
                   (document.ExtractionStatus <> "complete" AndAlso document.ExtractionStatus <> "partial") Then Return False
            Next
            Return True
        End Function

        Private Shared Sub SynchronizeFindingFreshness(project As ReviewProject)
            For Each finding As ReviewFinding In project.Findings
                If finding.Status <> "stale" AndAlso finding.Status <> "dismissed" AndAlso Not FindingSourcesCurrent(project, finding) Then
                    ReviewContract.ArchiveFinding(finding, "Source changed or left the active review")
                    finding.Status = "stale"
                    finding.UpdatedUtc = System.DateTime.UtcNow.ToString("o")
                End If
            Next
        End Sub

        Private Sub SetFindingsArchived(request As Newtonsoft.Json.Linq.JObject, archived As System.Boolean)
            Dim ids As System.Collections.Generic.List(Of System.String) = StringList(request("findingIds"))
            If ids.Count = 0 Then Throw New System.IO.InvalidDataException("Select at least one finding.")
            Dim candidate As ReviewProject = ReviewContract.Clone(_project)
            For Each id As System.String In ids
                ReviewContract.RequireId(id)
                Dim finding As ReviewFinding = candidate.Findings.Find(Function(item As ReviewFinding) item.Id = id)
                If finding Is Nothing Then Throw New System.IO.InvalidDataException("A selected finding no longer belongs to this review.")
                If finding.Archived = archived Then Continue For
                ReviewContract.ArchiveFinding(finding, If(archived, "Removed finding", "Restored finding"))
                finding.Archived = archived
                finding.UpdatedUtc = System.DateTime.UtcNow.ToString("o")
            Next
            Commit(candidate, If(archived, "findings_removed", "findings_restored"), If(archived, "Removed ", "Restored ") &
                ids.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) & " finding(s). The review retains their sources, run and edit history.", actor:=RequireReviewer(request))
        End Sub

        Private Sub SaveFinding(request As Newtonsoft.Json.Linq.JObject)
            Dim id As System.String = TextField(request, "findingId")
            ReviewContract.RequireId(id)
            Dim actor As System.String = RequireReviewer(request)
            Dim candidate As ReviewProject = ReviewContract.Clone(_project)
            Dim finding As ReviewFinding = candidate.Findings.Find(Function(item As ReviewFinding) item.Id = id)
            If finding Is Nothing OrElse finding.Archived Then Throw New System.IO.InvalidDataException("Restore this finding before editing it.")
            Dim status As System.String = TextField(request, "status", finding.Status)
            If Not ReviewContract.Contains(New System.String() {"draft", "reviewed", "dismissed", "stale"}, status) Then Throw New System.IO.InvalidDataException("Choose a valid finding status.")
            If status = "reviewed" AndAlso Not FindingSourcesCurrent(candidate, finding) Then Throw New System.InvalidOperationException("Refresh or restore the cited sources and recheck this finding before marking it reviewed.")
            ReviewContract.ArchiveFinding(finding, "Manual finding edit")
            finding.Title = TextField(request, "title", finding.Title).Trim()
            finding.Description = TextField(request, "description", finding.Description)
            finding.Assessment = TextField(request, "assessment", finding.Assessment)
            finding.Priority = TextField(request, "priority", finding.Priority)
            finding.Status = status
            finding.Question = TextField(request, "question", finding.Question)
            finding.RequestedAction = TextField(request, "requestedAction", finding.RequestedAction)
            finding.Notes = TextField(request, "notes", finding.Notes)
            finding.Reviewer = TextField(request, "reviewer", finding.Reviewer)
            If status = "reviewed" AndAlso System.String.IsNullOrWhiteSpace(finding.Reviewer) Then finding.Reviewer = actor
            finding.UserEdited = True
            finding.UpdatedUtc = System.DateTime.UtcNow.ToString("o")
            finding.UpdatedBy = actor
            finding.ReviewedBy = If(status = "reviewed", actor, "")
            finding.ReviewedUtc = If(status = "reviewed", finding.UpdatedUtc, "")
            For Each edits As System.Collections.Generic.Dictionary(Of System.String, ReviewAnswer) In candidate.GroupEdits.Values
                For Each edited As ReviewAnswer In edits.Values
                    If edited.SourceHash <> finding.RunId Then Continue For
                    ReviewContract.ArchiveAnswer(edited, "Finding edit")
                    edited.Text = finding.Description
                    edited.Assessment = finding.Assessment
                    edited.Status = finding.Status
                    edited.UpdatedUtc = finding.UpdatedUtc
                    edited.UpdatedBy = actor
                    edited.ReviewedUtc = finding.ReviewedUtc
                    edited.ReviewedBy = finding.ReviewedBy
                Next
            Next
            Commit(candidate, "finding_edit", "Saved the finding's assessment, priority and review status.", actor:=RequireReviewer(request))
        End Sub

        ' Column options and flag labels are exact authored values, not identifiers.
        Private Shared Function ExactStringList(token As Newtonsoft.Json.Linq.JToken) As System.Collections.Generic.List(Of System.String)
            Dim result As New System.Collections.Generic.List(Of System.String)()
            If token Is Nothing OrElse token.Type = Newtonsoft.Json.Linq.JTokenType.Null Then Return result
            Dim values As Newtonsoft.Json.Linq.JArray = TryCast(token, Newtonsoft.Json.Linq.JArray)
            If values Is Nothing Then Throw New System.IO.InvalidDataException("Expected a list of text values.")
            For Each item As Newtonsoft.Json.Linq.JToken In values
                If item.Type <> Newtonsoft.Json.Linq.JTokenType.String OrElse System.String.IsNullOrWhiteSpace(item.ToString()) Then Throw New System.IO.InvalidDataException("Column options and flag labels must contain text.")
                Dim value As System.String = item.ToString()
                If result.Contains(value) Then Throw New System.IO.InvalidDataException("A column option or flag label appears more than once.")
                result.Add(value)
            Next
            Return result
        End Function

        Private Shared Function StringList(token As Newtonsoft.Json.Linq.JToken) As System.Collections.Generic.List(Of System.String)
            Dim result As New System.Collections.Generic.List(Of System.String)()
            If token Is Nothing OrElse token.Type = Newtonsoft.Json.Linq.JTokenType.Null Then Return result
            Dim values As Newtonsoft.Json.Linq.JArray = TryCast(token, Newtonsoft.Json.Linq.JArray)
            If values Is Nothing Then Throw New System.IO.InvalidDataException("Expected a list of text values.")
            For Each item As Newtonsoft.Json.Linq.JToken In values
                If item.Type <> Newtonsoft.Json.Linq.JTokenType.String Then Throw New System.IO.InvalidDataException("A list contains a non-text value.")
                Dim value As System.String = item.ToString().Trim()
                If value.Length > 0 AndAlso Not result.Contains(value) Then result.Add(value)
            Next
            Return result
        End Function

        Private Sub SaveAnswer(request As Newtonsoft.Json.Linq.JObject)
            Dim actor As System.String = RequireReviewer(request)
            Dim candidate As ReviewProject = ReviewContract.Clone(_project)
            Dim documentId As System.String = TextField(request, "documentId")
            Dim group As System.Boolean = documentId.StartsWith("group:", System.StringComparison.Ordinal)
            Dim document As ReviewDocument = If(group, ReviewContract.CurrentGroupRows(candidate).Find(Function(item As ReviewDocument) item.Id = documentId), FindDocument(documentId, candidate))
            Dim criterion As ReviewCriterion = FindCriterion(TextField(request, "criterionId"), candidate)
            If document Is Nothing Then Throw New System.IO.InvalidDataException("The group no longer exists.")
            If Not If(group, candidate.Documents.Exists(Function(member As ReviewDocument) Not member.Archived AndAlso member.GroupName = document.GroupName AndAlso criterion.AppliesTo(member)), criterion.AppliesTo(document)) Then Throw New System.IO.InvalidDataException("This question does not apply to the selected row.")
            Dim answer As ReviewAnswer = Nothing
            If Not document.Answers.TryGetValue(criterion.Id, answer) Then
                answer = New ReviewAnswer()
                document.Answers(criterion.Id) = answer
            End If
            Dim status As System.String = TextField(request, "status", "draft")
            If status <> "draft" AndAlso status <> "reviewed" Then Throw New System.IO.InvalidDataException("Choose Draft or Reviewed for an edited answer.")
            Dim text As System.String = TextField(request, "text")
            If status = "reviewed" AndAlso System.String.IsNullOrWhiteSpace(text) Then Throw New System.IO.InvalidDataException("Enter the answer or explain the absence of a finding before marking this answer reviewed.")
            If (criterion.Kind = "priority" OrElse criterion.Kind = "flag") AndAlso Not group AndAlso criterion.Options.Count > 0 AndAlso text.Length > 0 AndAlso Not criterion.Options.Contains(text) Then Throw New System.IO.InvalidDataException("Choose one of the values configured for this column.")
            Dim riskFlag As System.String = TextField(request, "riskFlag", answer.RiskFlag)
            Dim labels As System.Collections.Generic.List(Of System.String) = criterion.FlagValues()
            If labels.Count > 0 AndAlso (riskFlag.Length > 0 OrElse status = "reviewed") AndAlso Not labels.Contains(riskFlag) Then Throw New System.IO.InvalidDataException("Select a configured flag for this answer.")
            ReviewContract.ArchiveAnswer(answer, "Manual edit")
            answer.RiskFlag = If(labels.Count > 0, riskFlag, "")
            answer.TextStyle = ReadFieldStyle(request, "textStyle", answer.TextStyle)
            answer.AssessmentStyle = ReadFieldStyle(request, "assessmentStyle", answer.AssessmentStyle)
            answer.Text = text
            answer.Assessment = TextField(request, "assessment")
            answer.Status = status
            answer.UserEdited = True
            answer.CriterionRevision = criterion.Revision
            If Not group Then answer.SourceHash = document.SourceHash
            answer.UpdatedUtc = System.DateTime.UtcNow.ToString("o")
            answer.UpdatedBy = actor
            answer.ReviewedBy = If(status = "reviewed", actor, "")
            answer.ReviewedUtc = If(status = "reviewed", answer.UpdatedUtc, "")
            If group Then
                answer.GroupContext = document.TextHash
                If Not candidate.GroupEdits.ContainsKey(document.Id) Then candidate.GroupEdits.Add(document.Id, New System.Collections.Generic.Dictionary(Of System.String, ReviewAnswer)())
                Dim finding As ReviewFinding = candidate.Findings.Find(Function(item As ReviewFinding) item.RunId = answer.SourceHash AndAlso Not item.Archived)
                If finding IsNot Nothing Then
                    ReviewContract.ArchiveFinding(finding, "Group matrix edit")
                    finding.Description = answer.Text
                    finding.Assessment = answer.Assessment
                    finding.UserEdited = True
                    finding.UpdatedUtc = answer.UpdatedUtc
                    finding.UpdatedBy = actor
                    finding.Status = status
                    finding.ReviewedUtc = answer.ReviewedUtc
                    finding.ReviewedBy = answer.ReviewedBy
                End If
                candidate.GroupEdits(document.Id)(criterion.Id) = answer
                Commit(candidate, "group_answer_edit", status, document.Id, criterion.Id, actor:=actor)
                Return
            End If
            document.UpdatedBy = actor
            document.UpdatedUtc = answer.UpdatedUtc
            document.ReviewedBy = ""
            document.ReviewedUtc = ""
            If status = "reviewed" AndAlso System.String.IsNullOrWhiteSpace(document.Reviewer) Then document.Reviewer = actor
            document.ReviewStatus = "in_review"
            document.LastReviewedUtc = ""
            Commit(candidate, "answer_edit", ReviewContract.Serialize(New With {.status = status, .previousReviewer = FindDocument(document.Id).Reviewer, .reviewer = document.Reviewer}), document.Id, criterion.Id, actor:=RequireReviewer(request))
        End Sub

        Private Shared Function ReadFieldStyle(request As Newtonsoft.Json.Linq.JObject, key As System.String, fallback As System.Int32) As System.Int32
            If request(key) Is Nothing Then Return fallback
            Dim value As System.Int32
            If Not System.Int32.TryParse(request(key).ToString(), value) OrElse value < 0 OrElse value > 3 Then Throw New System.IO.InvalidDataException("Invalid field formatting.")
            Return value
        End Function

        Private Sub DissolveGroup(request As Newtonsoft.Json.Linq.JObject)
            Dim name As System.String = TextField(request, "groupName")
            If System.String.IsNullOrWhiteSpace(name) Then Throw New System.IO.InvalidDataException("Select a named group.")
            Dim candidate As ReviewProject = ReviewContract.Clone(_project)
            For Each document As ReviewDocument In candidate.Documents
                If document.GroupName = name Then document.GroupName = ""
            Next
            For Each run As ReviewConsistencyRun In candidate.ConsistencyRuns
                If run.GroupName = name Then run.GroupName = ""
            Next
            candidate.GroupEdits.Remove("group:" & ReviewContract.Digest(name))
            Commit(candidate, "dissolve_group", name, actor:=RequireReviewer(request))
        End Sub

        Private Sub AssignReviewScope(request As Newtonsoft.Json.Linq.JObject, command As System.String)
            Dim actor As System.String = RequireReviewer(request)
            Dim selected As System.Collections.Generic.List(Of ReviewDocument) = SelectedDocuments(request)
            Dim candidate As ReviewProject = ReviewContract.Clone(_project)
            Dim changes As New System.Collections.Generic.List(Of System.Object)()
            Dim value As System.String = TextField(request, "value").Trim()
            For Each original As ReviewDocument In selected
                Dim document As ReviewDocument = FindDocument(original.Id, candidate)
                Dim previous As System.String = If(command = "assignCategory", document.Category, document.GroupName)
                Dim assignment As System.String = value
                If command = "groupByFolder" Then
                    If previous.Length > 0 OrElse System.String.IsNullOrWhiteSpace(document.RelativePath) OrElse Not ReviewContract.IsSafeRelativePath(document.RelativePath.Replace("\", "/")) Then Continue For
                    Dim relative As System.String = document.RelativePath.Replace("\", "/")
                    Dim separator As System.Int32 = relative.LastIndexOf("/"c)
                    If separator <= 0 Then Continue For
                    assignment = relative.Substring(0, separator)
                End If
                If previous = assignment Then Continue For
                If command = "assignCategory" Then
                    document.Category = assignment
                    document.ReviewStatus = "unreviewed"
                    document.LastReviewedUtc = ""
                    document.ReviewedBy = ""
                    document.ReviewedUtc = ""
                Else
                    document.GroupName = assignment
                End If
                document.UpdatedBy = actor
                document.UpdatedUtc = System.DateTime.UtcNow.ToString("o")
                changes.Add(New With {.documentId = document.Id, .previous = previous, .current = assignment})
            Next
            If changes.Count > 0 Then Commit(candidate, command, ReviewContract.Serialize(changes), actor:=actor)
            _message = "Updated assignments: " & changes.Count.ToString() & " of " & selected.Count.ToString() & " selected documents."
        End Sub

        Private Sub SaveDocument(request As Newtonsoft.Json.Linq.JObject)
            Dim actor As System.String = RequireReviewer(request)
            Dim candidate As ReviewProject = ReviewContract.Clone(_project)
            Dim document As ReviewDocument = FindDocument(TextField(request, "documentId"), candidate)
            document.Category = TextField(request, "category", document.Category).Trim()
            document.GroupName = TextField(request, "groupName", document.GroupName).Trim()
            Dim status As System.String = TextField(request, "reviewStatus", "unreviewed")
            If Not ReviewContract.Contains(New System.String() {"unreviewed", "in_review", "reviewed"}, status) Then Throw New System.IO.InvalidDataException("Unknown review status.")
            If status = "reviewed" Then
                For Each criterion As ReviewCriterion In candidate.Criteria
                    If criterion.Archived OrElse criterion.Kind = "text" OrElse Not criterion.AppliesTo(document) Then Continue For
                    Dim answer As ReviewAnswer = Nothing
                    If Not document.Answers.TryGetValue(criterion.Id, answer) OrElse System.String.IsNullOrWhiteSpace(answer.Text) OrElse answer.Status = "stale" OrElse answer.Status = "error" OrElse answer.Status = "pending" OrElse (criterion.FlagValues().Count > 0 AndAlso Not criterion.FlagValues().Contains(answer.RiskFlag)) Then
                        Throw New System.InvalidOperationException("Resolve empty, stale or failed answers before marking this document reviewed.")
                    End If
                Next
                document.LastReviewedUtc = System.DateTime.UtcNow.ToString("o")
            Else
                document.LastReviewedUtc = ""
            End If
            document.Number = TextField(request, "number", document.Number)
            document.ReviewStatus = status
            document.Priority = TextField(request, "priority")
            document.Flags = StringList(request("flags"))
            document.Reviewer = TextField(request, "reviewer")
            document.Notes = TextField(request, "notes")
            document.UpdatedBy = actor
            document.UpdatedUtc = System.DateTime.UtcNow.ToString("o")
            document.ReviewedBy = If(status = "reviewed", actor, "")
            document.ReviewedUtc = If(status = "reviewed", document.LastReviewedUtc, "")
            If status = "reviewed" AndAlso System.String.IsNullOrWhiteSpace(document.Reviewer) Then document.Reviewer = actor
            Commit(candidate, "document_review", ReviewContract.Serialize(New With {.previous = DocumentMetadata(FindDocument(document.Id)), .current = DocumentMetadata(document)}), document.Id, actor:=RequireReviewer(request))
        End Sub

        Private Shared Function DocumentMetadata(document As ReviewDocument) As System.Object
            Return New With {.category = document.Category, .groupName = document.GroupName, .number = document.Number, .reviewStatus = document.ReviewStatus, .priority = document.Priority,
                .flags = document.Flags, .reviewer = document.Reviewer, .notes = document.Notes, .lastReviewedUtc = document.LastReviewedUtc,
                .updatedBy = document.UpdatedBy, .updatedUtc = document.UpdatedUtc, .reviewedBy = document.ReviewedBy, .reviewedUtc = document.ReviewedUtc}
        End Function

        Private Sub SaveView(request As Newtonsoft.Json.Linq.JObject)
            Dim value As Newtonsoft.Json.Linq.JObject = TryCast(request("view"), Newtonsoft.Json.Linq.JObject)
            If value Is Nothing Then Throw New System.IO.InvalidDataException("Missing view state.")
            Dim candidate As ReviewProject = ReviewContract.Clone(_project)
            Dim view As ReviewViewState = Newtonsoft.Json.JsonConvert.DeserializeObject(Of ReviewViewState)(value.ToString(), ReviewContract.JsonSettings)
            If view Is Nothing Then Throw New System.IO.InvalidDataException("Missing view state.")
            ' Autosave is changed explicitly, so an older queued layout message
            ' cannot silently turn it on again after the reviewer switches it off.
            view.AutoSave = _project.View.AutoSave
            view.LayoutVersion = 3
            ' Order changes use their own revision-checked command. An older queued
            ' layout must not undo a newer matrix or Criteria-list reorder.
            view.ColumnOrder = ActiveCriterionIds(candidate)
            If Not System.String.IsNullOrEmpty(view.DocumentId) Then FindDocument(view.DocumentId, candidate)
            If Not System.String.IsNullOrEmpty(view.CriterionId) Then FindCriterion(view.CriterionId, candidate)
            candidate.View = view
            ReviewContract.NormalizeView(candidate)
            ' Retain view changes in memory and include them in the next real commit,
            ' Save workbook or close-flush. Never rewrite the workbook on each search key.
            _project.View = candidate.View
            _persistenceMessage = ""
        End Sub

        Private Function ImportSelection() As System.Collections.Generic.List(Of System.String)
            Dim importedIds As New System.Collections.Generic.List(Of System.String)()
            Dim imported As ReviewProject = _callbacks.ImportSelection.Invoke()
            If imported Is Nothing Then Return importedIds
            ReviewContract.Validate(imported)
            Dim candidate As ReviewProject = ReviewContract.Clone(_project)
            Dim map As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.Ordinal)
            Dim usedTargets As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            For Each criterion As ReviewCriterion In imported.Criteria
                ' Equal headers can represent different source columns. Match each
                ' target at most once and retain every original column/value.
                Dim same As ReviewCriterion = candidate.Criteria.Find(Function(item As ReviewCriterion) Not item.Archived AndAlso Not usedTargets.Contains(item.Id) AndAlso item.Category = criterion.Category AndAlso item.FlagScale = criterion.FlagScale AndAlso ReviewContract.Serialize(item.FlagConditions) = ReviewContract.Serialize(criterion.FlagConditions) AndAlso item.Title = criterion.Title AndAlso item.Instruction = criterion.Instruction AndAlso item.Kind = criterion.Kind AndAlso ReviewContract.Serialize(item.Options) = ReviewContract.Serialize(criterion.Options))
                If same Is Nothing Then
                    candidate.Criteria.Add(criterion)
                    map(criterion.Id) = criterion.Id
                    For Each document As ReviewDocument In candidate.Documents
                        If criterion.AppliesTo(document) Then ReviewContract.Invalidate(document, criterion.Id)
                    Next
                Else
                    map(criterion.Id) = same.Id
                End If
                usedTargets.Add(map(criterion.Id))
            Next
            For Each document As ReviewDocument In imported.Documents
                If document.ImportedFrom.Length > 0 AndAlso candidate.Documents.Exists(Function(item As ReviewDocument) item.ImportedFrom = document.ImportedFrom) Then Throw New System.InvalidOperationException("This Excel result row already exists in the review: " & document.ImportedFrom & ". Edit that review row or choose a different source range.")
                Dim answers As New System.Collections.Generic.Dictionary(Of System.String, ReviewAnswer)(System.StringComparer.Ordinal)
                For Each answer As System.Collections.Generic.KeyValuePair(Of System.String, ReviewAnswer) In document.Answers
                    answers.Add(map(answer.Key), answer.Value)
                Next
                document.Answers = answers
                candidate.Documents.Add(document)
                importedIds.Add(document.Id)
            Next
            If candidate.SourceRoot.Length = 0 Then candidate.SourceRoot = imported.SourceRoot
            Commit(candidate, "import", "Copied " & imported.Documents.Count.ToString() & " result rows from Excel. Original cells remain unchanged; imported answers require source verification.")
            Return importedIds
        End Function

        Private Async Function LoadRegisteredSourcesAsync(ids As System.Collections.Generic.List(Of System.String), cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            Dim paths As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            For Each id As System.String In ids
                Dim document As ReviewDocument = FindDocument(id)
                If Not System.String.IsNullOrWhiteSpace(document.SourcePath) AndAlso System.IO.File.Exists(document.SourcePath) Then paths.Add(document.SourcePath)
            Next
            For Each path As System.String In paths
                cancellation.ThrowIfCancellationRequested()
                _message = "Reading the linked source of the imported result: " & System.IO.Path.GetFileName(path)
                SendState()
                Dim extracted As ReviewDocument = Await ReviewSourceService.ExtractAsync(path, _project.SourceRoot, _context, _project.UseOcr, cancellation)
                Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
                EnsureAvailable()
                Dim candidate As ReviewProject = ReviewContract.Clone(_project)
                For Each id As System.String In ids
                    Dim document As ReviewDocument = FindDocument(id, candidate)
                    If System.String.Equals(document.SourcePath, path, System.StringComparison.OrdinalIgnoreCase) Then ApplySource(document, extracted)
                Next
                Commit(candidate, "import_source", extracted.ExtractionStatus & ": " & extracted.ExtractionMessage)
            Next
            If ids.Count > 0 Then _message = "Imported answers remain unchanged. Read " & paths.Count.ToString() & " linked source file(s). Use Link source or scan the source folder for unresolved files; verify imported answers against their sources."
        End Function

        Private Async Function ScanFolderAsync(request As Newtonsoft.Json.Linq.JObject, cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            ' Validate and enumerate before modifying Excel. Picker and typed paths
            ' now follow the same pipeline and only the first imported file commits.
            Dim rawRoot As System.String = TextField(request, "sourceRoot", _project.SourceRoot).Trim()
            If rawRoot.Length = 0 Then Throw New System.InvalidOperationException("Choose a source folder first.")
            Dim root As System.String = NormalizeLocalPath(rawRoot)
            Dim recursive As System.Boolean = BooleanField(request, "includeSubfolders", _project.IncludeSubfolders)
            Dim useOcr As System.Boolean = BooleanField(request, "useOcr", _project.UseOcr)
            Dim allowLegacy As System.Boolean = _context.INI_AllowLegacyDocFiles
            Dim excluded As New System.Collections.Generic.List(Of System.String)()
            If _callbacks.WorkbookPath IsNot Nothing Then excluded.Add(_callbacks.WorkbookPath.Invoke())
            For Each report As ReviewReportFile In _project.ReportFiles
                excluded.Add(report.Path)
            Next
            _message = "Checking the source folder and excluding Tabular Review reports and review workbooks."
            SendState()
            Dim files As System.Collections.Generic.List(Of System.String) = Await System.Threading.Tasks.Task.Run(Function() ReviewSourceService.EnumerateFiles(root, recursive, allowLegacy, excluded, cancellation), cancellation)
            Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
            EnsureAvailable()
            If files.Count = 0 Then Throw New System.InvalidOperationException("No supported source documents remain after excluding Tabular Review reports and review workbooks. Choose another folder or check its file types.")
            Dim linked As ReviewProject = ReviewContract.Clone(_project)
            linked.SourceRoot = root
            linked.IncludeSubfolders = recursive
            linked.UseOcr = useOcr
            If ResolveMissingSources(linked, files, root) Then Commit(linked, "source_links", "Resolved missing sources by exact relative path or a unique filename within the selected folder. Ambiguous files remain unlinked.")
            For index As System.Int32 = 0 To files.Count - 1
                cancellation.ThrowIfCancellationRequested()
                Dim path As System.String = files(index)
                _message = "Reading document " & (index + 1).ToString() & " of " & files.Count.ToString() & ": " & System.IO.Path.GetFileName(path)
                SendState()
                Dim hash As System.String = ""
                Dim failure As System.String = ""
                Try
                    hash = Await System.Threading.Tasks.Task.Run(Function() ReviewSourceService.FileHash(path, cancellation), cancellation)
                Catch ex As System.OperationCanceledException
                    Throw
                Catch ex As System.Exception
                    failure = ex.Message
                End Try
                Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
                EnsureAvailable()
                Dim matches As System.Collections.Generic.List(Of ReviewDocument) = _project.Documents.FindAll(Function(item As ReviewDocument) Not item.Archived AndAlso System.String.Equals(item.SourcePath, path, System.StringComparison.OrdinalIgnoreCase))
                If matches.Count = 0 AndAlso _project.Documents.Exists(Function(item As ReviewDocument) item.Archived AndAlso System.String.Equals(item.SourcePath, path, System.StringComparison.OrdinalIgnoreCase)) Then Continue For
                If failure.Length = 0 AndAlso matches.Count > 0 AndAlso matches.TrueForAll(Function(item As ReviewDocument) CanReuseExtraction(item, hash)) Then
                    Dim restored As ReviewProject = Nothing
                    Dim relative As System.String = RelativePath(root, path)
                    For Each matching As ReviewDocument In matches
                        If matching.ExtractionStatus <> matching.StoredExtractionStatus OrElse matching.ExtractionMessage <> matching.StoredExtractionMessage OrElse matching.RelativePath <> relative Then
                            If restored Is Nothing Then restored = ReviewContract.Clone(_project)
                            RestoreCachedSource(FindDocument(matching.Id, restored), path, root)
                        End If
                    Next
                    If restored IsNot Nothing Then
                        restored.SourceRoot = root
                        restored.IncludeSubfolders = recursive
                        restored.UseOcr = useOcr
                        Commit(restored, "source_cache_reused", "The original hashes match the saved extracts. Restored source availability without repeating text extraction or OCR.")
                    End If
                    Continue For
                End If
                Dim extracted As ReviewDocument = Nothing
                If failure.Length = 0 Then
                    Try
                        extracted = Await ReviewSourceService.ExtractAsync(path, root, _context, useOcr, cancellation)
                    Catch ex As System.OperationCanceledException
                        Throw
                    Catch ex As System.Exception
                        failure = ex.Message
                    End Try
                End If
                Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
                EnsureAvailable()
                Dim candidate As ReviewProject = ReviewContract.Clone(_project)
                candidate.SourceRoot = root
                candidate.IncludeSubfolders = recursive
                candidate.UseOcr = useOcr
                If extracted Is Nothing Then extracted = New ReviewDocument With {.FileName = System.IO.Path.GetFileName(path), .SourcePath = path, .RelativePath = RelativePath(root, path), .Number = DocumentNumber(root, path), .ExtractionStatus = "error", .ExtractionMessage = failure}
                If matches.Count = 0 Then
                    If extracted.Number.Length = 0 Then extracted.Number = DocumentNumber(root, path)
                    candidate.Documents.Add(extracted)
                Else
                    For Each matching As ReviewDocument In matches
                        ApplySource(FindDocument(matching.Id, candidate), extracted)
                    Next
                End If
                Commit(candidate, "source_import", If(failure.Length > 0, failure, extracted.ExtractionStatus & ": " & extracted.ExtractionMessage), extracted.Id)
            Next
            ' A scan can encounter only unchanged sources. Persist the confirmed
            ' folder settings in that case too, without forcing a pre-scan save.
            If _project.SourceRoot <> root OrElse _project.IncludeSubfolders <> recursive OrElse _project.UseOcr <> useOcr Then
                Dim candidate As ReviewProject = ReviewContract.Clone(_project)
                candidate.SourceRoot = root
                candidate.IncludeSubfolders = recursive
                candidate.UseOcr = useOcr
                Commit(candidate, "source_folder", "Updated source folder and scan options.")
            End If
            _message = "Folder import completed. Review source warnings, configure columns, then generate answers. Save the workbook to keep the imported sources."
        End Function

        Private Shared Function ResolveMissingSources(project As ReviewProject, files As System.Collections.Generic.List(Of System.String), root As System.String) As System.Boolean
            Dim changed As System.Boolean = False
            For Each document As ReviewDocument In project.Documents
                If document.Archived OrElse System.IO.File.Exists(document.SourcePath) Then Continue For
                Dim matches As System.Collections.Generic.List(Of System.String) = files.FindAll(
                    Function(path As System.String) System.String.Equals(RelativePath(root, path).Replace("/", "\"), If(document.RelativePath, "").Replace("/", "\"), System.StringComparison.OrdinalIgnoreCase))
                If matches.Count = 0 AndAlso Not System.String.IsNullOrWhiteSpace(document.FileName) Then
                    matches = files.FindAll(Function(path As System.String) System.String.Equals(System.IO.Path.GetFileName(path), document.FileName, System.StringComparison.OrdinalIgnoreCase))
                End If
                If matches.Count = 1 Then
                    document.SourcePath = matches(0)
                    changed = True
                ElseIf matches.Count > 1 Then
                    document.ExtractionMessage = "More than one file matches this imported source name. Use Link source to choose the correct document."
                    changed = True
                End If
            Next
            Return changed
        End Function

        Private Shared Function CanReuseExtraction(document As ReviewDocument, hash As System.String) As System.Boolean
            Return hash.Length > 0 AndAlso document.SourceHash = hash AndAlso document.Segments.Count > 0 AndAlso
                (document.StoredExtractionStatus = "complete" OrElse document.StoredExtractionStatus = "partial")
        End Function

        Private Shared Sub RestoreCachedSource(document As ReviewDocument, path As System.String, root As System.String)
            document.SourcePath = path
            document.RelativePath = RelativePath(root, path)
            document.FileName = System.IO.Path.GetFileName(path)
            document.ExtractionStatus = document.StoredExtractionStatus
            document.ExtractionMessage = document.StoredExtractionMessage
        End Sub

        Private Async Function RefreshDocumentAsync(id As System.String, choosePath As System.Boolean, cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            Dim document As ReviewDocument = FindDocument(id)
            Dim path As System.String = document.SourcePath
            Dim sourceChosen As System.Boolean = choosePath OrElse Not System.IO.File.Exists(path)
            If sourceChosen Then
                Using dialog As New System.Windows.Forms.OpenFileDialog With {.Title = "Link the source document", .CheckFileExists = True, .Multiselect = False,
                    .Filter = "Review documents|*.pdf;*.docx;*.docm;*.doc;*.rtf;*.xlsx;*.xlsm;*.pptx;*.pptm;*.txt;*.md;*.csv;*.tsv;*.json;*.xml;*.html;*.htm;*.log;*.ini;*.yaml;*.yml;*.eml;*.msg|All files|*.*"}
                    If System.IO.Directory.Exists(_project.SourceRoot) Then dialog.InitialDirectory = _project.SourceRoot
                    If dialog.ShowDialog(Me) <> System.Windows.Forms.DialogResult.OK Then Return
                    path = NormalizeLocalPath(dialog.FileName)
                End Using
            End If
            If sourceChosen Then
                Dim check As SourceFileCheck = Await CheckSourceFileAsync(path, cancellation)
                Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
                EnsureAvailable()
                If check.Status.Length = 0 AndAlso CanReuseExtraction(document, check.Hash) Then
                    Dim restored As ReviewProject = ReviewContract.Clone(_project)
                    RestoreCachedSource(FindDocument(id, restored), path, _project.SourceRoot)
                    Commit(restored, "source_relinked", "The selected original matches the saved hash. Reused the saved extraction and retained review history.", id)
                    Return
                End If
            End If
            Dim extracted As ReviewDocument = Await ReviewSourceService.ExtractAsync(path, _project.SourceRoot, _context, _project.UseOcr, cancellation)
            Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
            EnsureAvailable()
            Dim candidate As ReviewProject = ReviewContract.Clone(_project)
            ApplySource(FindDocument(id, candidate), extracted)
            Commit(candidate, "source_refresh", extracted.ExtractionStatus & ": " & extracted.ExtractionMessage, id)
        End Function

        Private Shared Sub ApplySource(document As ReviewDocument, extracted As ReviewDocument)
            If extracted.Segments.Count = 0 Then
                ' Preserve the previous source snapshot on a failed refresh.
                document.ExtractionStatus = extracted.ExtractionStatus
                document.ExtractionMessage = extracted.ExtractionMessage
                ReviewContract.Invalidate(document)
                Return
            End If
            Dim changed As System.Boolean = document.SourceHash.Length > 0 AndAlso (document.SourceHash <> extracted.SourceHash OrElse document.TextHash <> extracted.TextHash)
            If changed Then
                document.SourceHistory.Add(New ReviewSourceVersion With {.SourcePath = document.SourcePath, .FileName = document.FileName,
                    .RelativePath = document.RelativePath, .SourceHash = document.SourceHash, .TextHash = document.TextHash,
                    .ExtractionStatus = If(System.String.IsNullOrEmpty(document.StoredExtractionStatus), document.ExtractionStatus, document.StoredExtractionStatus),
                    .ExtractionMessage = If(System.String.IsNullOrEmpty(document.StoredExtractionStatus), document.ExtractionMessage, document.StoredExtractionMessage), .Segments = document.Segments})
                For Each answer As ReviewAnswer In document.Answers.Values
                    ReviewContract.ArchiveAnswer(answer, "Source changed")
                    answer.Evidence = New System.Collections.Generic.List(Of ReviewEvidence)()
                    answer.Error = "The source changed. Previous quotations remain in the answer history; regenerate or verify this answer against the current source."
                Next
                ReviewContract.Invalidate(document)
            End If
            document.SourcePath = extracted.SourcePath
            document.FileName = extracted.FileName
            document.RelativePath = extracted.RelativePath
            If document.Number.Length = 0 Then document.Number = extracted.Number
            document.SourceHash = extracted.SourceHash
            document.SourceLength = extracted.SourceLength
            document.SourceLastWriteUtc = extracted.SourceLastWriteUtc
            document.TextHash = extracted.TextHash
            document.Segments = extracted.Segments
            document.ExtractionStatus = extracted.ExtractionStatus
            document.ExtractionMessage = extracted.ExtractionMessage
            If extracted.ExtractionStatus = "complete" OrElse extracted.ExtractionStatus = "partial" Then
                document.StoredExtractionStatus = extracted.ExtractionStatus
                document.StoredExtractionMessage = extracted.ExtractionMessage
            Else
                document.StoredExtractionStatus = ""
                document.StoredExtractionMessage = ""
            End If
        End Sub

        Private Shared Function RelativePath(root As System.String, path As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(root) Then Return System.IO.Path.GetFileName(path)
            Dim prefix As System.String = System.IO.Path.GetFullPath(root).TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar) & System.IO.Path.DirectorySeparatorChar
            Dim full As System.String = System.IO.Path.GetFullPath(path)
            If full.StartsWith(prefix, System.StringComparison.OrdinalIgnoreCase) Then Return full.Substring(prefix.Length)
            Return full
        End Function

        Private Shared Function DocumentNumber(root As System.String, path As System.String) As System.String
            Dim relative As System.String = RelativePath(root, path)
            Dim extension As System.String = System.IO.Path.GetExtension(relative)
            Return ReviewContract.ShortDocumentReference(If(extension.Length > 0, relative.Substring(0, relative.Length - extension.Length), relative).Replace("\", " / "), ReviewContract.Digest(relative))
        End Function

        Private Async Function ClassifyDocumentsAsync(request As Newtonsoft.Json.Linq.JObject, cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            Dim selected = SelectedDocuments(request)
            Dim operationModelKey As System.String = ResolveOperationModel(request)
            Dim choice As ReviewModelChoice = _models.Find(Function(item As ReviewModelChoice) item.Key = operationModelKey)
            Dim invoke As System.Func(Of System.String, System.String, System.Threading.CancellationToken, System.Threading.Tasks.Task(Of System.String)) = ReviewModelCatalog.CreateInvoker(_context, choice)
            Dim configuredBudget As System.Int32 = _context.INI_TextAnalyzeMaxSourceChars
            Dim budget As System.Int32 = If(configuredBudget >= 2048, System.Math.Min(28000, configuredBudget), 28000)
            Dim engine As New ReviewAnalysisEngine(invoke, budget)
            Dim types As New System.Collections.Generic.List(Of System.String)()
            Dim known As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            For Each criterion As ReviewCriterion In _project.Criteria
                If Not criterion.Archived AndAlso criterion.Category.Length > 0 AndAlso known.Add(criterion.Category) Then types.Add(criterion.Category)
            Next
            For Each document As ReviewDocument In _project.Documents
                If Not document.Archived AndAlso document.Category.Length > 0 AndAlso known.Add(document.Category) Then types.Add(document.Category)
            Next
            If types.Count = 0 Then Throw New System.InvalidOperationException("Define document types on questions or documents first.")
            Dim options As New System.Collections.Generic.List(Of System.String) From {"0"}
            For index As System.Int32 = 1 To types.Count
                options.Add(index.ToString(System.Globalization.CultureInfo.InvariantCulture))
            Next
            Dim question As New ReviewCriterion With {.Title = "Document type", .Kind = "flag", .Options = options,
                .Instruction = "Classify the supplied document by its substantive purpose. Return the one-based index of the best matching type in this ordered list, or 0 for ambiguous or unsupported. Explain the choice and cite the source. Types: " & ReviewContract.Serialize(types)}
            Dim assigned As System.Int32
            For Each original As ReviewDocument In selected
                cancellation.ThrowIfCancellationRequested()
                Dim singleton As New System.Collections.Generic.List(Of ReviewDocument) From {original}
                Await VerifySelectedSourcesAsync(singleton, cancellation)
                Dim document As ReviewDocument = FindDocument(original.Id)
                If document.ExtractionStatus <> "complete" AndAlso document.ExtractionStatus <> "partial" Then Throw New System.IO.InvalidDataException("Refresh the source before classifying " & document.FileName)
                Dim answer As ReviewAnswer = Await System.Threading.Tasks.Task.Run(Function() engine.GenerateAsync(document, question, _project.OutputLanguage, operationModelKey, "", cancellation), cancellation)
                Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
                singleton(0) = FindDocument(original.Id)
                Await VerifySelectedSourcesAsync(singleton, cancellation)
                Dim candidate As ReviewProject = ReviewContract.Clone(_project)
                Dim current As ReviewDocument = FindDocument(original.Id, candidate)
                If current.SourceHash <> document.SourceHash OrElse (current.ExtractionStatus <> "complete" AndAlso current.ExtractionStatus <> "partial") Then Throw New System.IO.InvalidDataException("The source changed during classification.")
                Dim index As System.Int32 = System.Int32.Parse(answer.Text, System.Globalization.CultureInfo.InvariantCulture)
                Dim previous As System.String = current.Category
                If index > 0 AndAlso answer.Evidence.Count > 0 Then
                    current.Category = types(index - 1)
                    assigned += 1
                    If current.Category <> previous Then
                        current.ReviewStatus = "unreviewed"
                        current.LastReviewedUtc = ""
                        current.ReviewedUtc = ""
                        current.ReviewedBy = ""
                    End If
                End If
                cancellation.ThrowIfCancellationRequested()
                Commit(candidate, "classify_document", ReviewContract.Serialize(New With {.previous = previous, .category = current.Category, .choice = index, .reason = answer.Assessment, .evidence = answer.Evidence, .model = operationModelKey}), original.Id)
            Next
            _message = "Classified " & assigned.ToString() & " of " & selected.Count.ToString() & " documents. Ambiguous assignments were retained."
        End Function

        Private Async Function GenerateAsync(request As Newtonsoft.Json.Linq.JObject, allDocuments As System.Boolean, cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            Dim operationModelKey As System.String = ResolveOperationModel(request)
            Dim choice As ReviewModelChoice = _models.Find(Function(item As ReviewModelChoice) item.Key = operationModelKey)
            Dim invoke As System.Func(Of System.String, System.String, System.Threading.CancellationToken, System.Threading.Tasks.Task(Of System.String)) = ReviewModelCatalog.CreateInvoker(_context, choice)
            Dim configuredBudget As System.Int32 = _context.INI_TextAnalyzeMaxSourceChars
            Dim budget As System.Int32 = If(configuredBudget >= 2048, System.Math.Min(28000, configuredBudget), 28000)
            Dim engine As New ReviewAnalysisEngine(invoke, budget)
            Dim targets As System.Collections.Generic.List(Of ReviewDocument)
            If allDocuments Then
                targets = _project.Documents.FindAll(Function(item As ReviewDocument) Not item.Archived)
            ElseIf request("documentIds") IsNot Nothing Then
                targets = SelectedDocuments(request)
            Else
                targets = New System.Collections.Generic.List(Of ReviewDocument) From {FindDocument(TextField(request, "documentId"))}
            End If
            Dim protectExisting As System.Boolean = allDocuments OrElse targets.Count > 1
            Dim criterionId As System.String = TextField(request, "criterionId")
            Dim criteria As System.Collections.Generic.List(Of ReviewCriterion)
            If criterionId.Length > 0 Then
                criteria = New System.Collections.Generic.List(Of ReviewCriterion) From {FindCriterion(criterionId)}
            ElseIf request("criterionIds") IsNot Nothing Then
                criteria = New System.Collections.Generic.List(Of ReviewCriterion)()
                For Each id As System.String In StringList(request("criterionIds"))
                    criteria.Add(FindCriterion(id))
                Next
            Else
                criteria = _project.Criteria.FindAll(Function(item As ReviewCriterion) Not item.Archived AndAlso item.Kind <> "text" AndAlso Not System.String.IsNullOrWhiteSpace(item.Instruction))
            End If
            If criteria.Count = 0 Then Throw New System.InvalidOperationException("Add an analysis column with a detailed question first.")
            If targets.Count = 0 Then Throw New System.InvalidOperationException("Import documents first.")
            Dim completed As System.Int32 = 0
            Dim failures As System.Int32 = 0
            Dim progress As New System.Progress(Of System.String)(
                Sub(message As System.String)
                    If Not _closed AndAlso Not _hostLifetime.IsCancellationRequested AndAlso Not cancellation.IsCancellationRequested Then
                        _message = message
                        SendState()
                    End If
                End Sub)
            For Each target As ReviewDocument In targets
                cancellation.ThrowIfCancellationRequested()
                EnsureAvailable()
                Dim document As ReviewDocument = FindDocument(target.Id)
                If Not criteria.Exists(Function(item As ReviewCriterion) item.AppliesTo(document)) Then Continue For
                Dim sourceIssue As System.String = ""
                Dim sourceState As System.String = ""
                Try
                    If Not System.IO.File.Exists(document.SourcePath) Then
                        sourceIssue = "The original source is unavailable. Link the document before generating new answers; its saved extract remains readable."
                        sourceState = "missing"
                    Else
                        Dim hash As System.String = Await System.Threading.Tasks.Task.Run(Function() ReviewSourceService.FileHash(document.SourcePath, cancellation), cancellation)
                        If hash <> document.SourceHash Then
                            sourceIssue = "The source file changed. Refresh this document before generating new answers."
                            sourceState = "changed"
                        End If
                    End If
                Catch ex As System.OperationCanceledException
                    Throw
                Catch ex As System.Exception
                    sourceIssue = "The source could not be checked: " & ex.Message
                    sourceState = "error"
                End Try
                Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
                EnsureAvailable()
                If sourceIssue.Length = 0 AndAlso document.Segments.Count > 0 AndAlso
                   (document.StoredExtractionStatus = "complete" OrElse document.StoredExtractionStatus = "partial") AndAlso
                   document.ExtractionStatus <> document.StoredExtractionStatus Then
                    Dim recovered As ReviewProject = ReviewContract.Clone(_project)
                    Dim restored As ReviewDocument = FindDocument(document.Id, recovered)
                    restored.ExtractionStatus = restored.StoredExtractionStatus
                    restored.ExtractionMessage = restored.StoredExtractionMessage
                    Commit(recovered, "source_available", "The original bytes match the saved extract again; reused the cached text.", document.Id)
                End If
                For Each criterion As ReviewCriterion In criteria
                    If Not criterion.AppliesTo(document) Then Continue For
                    cancellation.ThrowIfCancellationRequested()
                    document = FindDocument(target.Id)
                    Dim existing As ReviewAnswer = Nothing
                    document.Answers.TryGetValue(criterion.Id, existing)
                    If protectExisting AndAlso existing IsNot Nothing AndAlso (existing.UserEdited OrElse existing.Status = "reviewed" OrElse existing.Status = "imported" OrElse (existing.Status <> "pending" AndAlso existing.Status <> "stale" AndAlso existing.Status <> "error")) Then Continue For
                    Dim result As ReviewAnswer = Nothing
                    Dim failure As System.String = sourceIssue
                    If failure.Length = 0 Then
                        Try
                            Dim language As System.String = _project.OutputLanguage
                            Dim instruction As System.String = TextField(request, "instruction")
                            result = Await System.Threading.Tasks.Task.Run(
                                Function() engine.GenerateAsync(document, criterion, language, operationModelKey,
                                    instruction, cancellation, Sub(value As System.String) DirectCast(progress, System.IProgress(Of System.String)).Report(value)), cancellation)
                        Catch ex As System.OperationCanceledException
                            Throw
                        Catch ex As System.Exception
                            failure = ex.Message
                        End Try
                    End If
                    Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
                    EnsureAvailable()
                    If result IsNot Nothing Then
                        Try
                            Dim currentHash As System.String = Await System.Threading.Tasks.Task.Run(Function() ReviewSourceService.FileHash(document.SourcePath, cancellation), cancellation)
                            If currentHash <> document.SourceHash Then
                                failure = "The source changed while the model was running. Refresh the document and run this criterion again."
                                sourceState = "changed"
                            End If
                        Catch ex As System.OperationCanceledException
                            Throw
                        Catch ex As System.Exception
                            failure = "The original source could not be checked after analysis: " & ex.Message
                            sourceState = "missing"
                        End Try
                    End If
                    Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
                    EnsureAvailable()
                    Dim candidate As ReviewProject = ReviewContract.Clone(_project)
                    Dim edited As ReviewDocument = FindDocument(target.Id, candidate)
                    If failure.Length > 0 Then
                        failures += 1
                        If sourceState.Length > 0 Then
                            edited.ExtractionStatus = sourceState
                            edited.ExtractionMessage = failure
                            ReviewContract.Invalidate(edited)
                        End If
                        Dim failed As ReviewAnswer = Nothing
                        If Not edited.Answers.TryGetValue(criterion.Id, failed) Then
                            failed = New ReviewAnswer With {.Status = "error"}
                            edited.Answers(criterion.Id) = failed
                        End If
                        failed.Error = failure
                        ' A failed attempt does not change the previous accepted text/status.
                        Commit(candidate, "analysis_failed", failure, edited.Id, criterion.Id)
                    Else
                        completed += 1
                        edited.Answers(criterion.Id) = result
                        edited.ReviewStatus = "in_review"
                        edited.LastReviewedUtc = ""
                        Commit(candidate, "analysis", "Generated a draft using " & choice.Label & ". Host validated every cited quotation against the stored source.", edited.Id, criterion.Id)
                    End If
                Next
            Next
            _message = completed.ToString() & " answer(s) generated; " & failures.ToString() & " failed. Existing manual, imported and reviewed answers remain protected in batch runs. Save the workbook."
        End Function

        Private Async Function VerifySelectedSourcesAsync(documents As System.Collections.Generic.List(Of ReviewDocument), cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            Dim candidate As ReviewProject = Nothing
            Dim checks As New System.Collections.Generic.Dictionary(Of System.String, SourceFileCheck)(System.StringComparer.OrdinalIgnoreCase)
            For Each document As ReviewDocument In documents
                cancellation.ThrowIfCancellationRequested()
                Dim path As System.String = If(document.SourcePath, "")
                Dim check As SourceFileCheck = Nothing
                If Not checks.TryGetValue(path, check) Then
                    check = Await CheckSourceFileAsync(path, cancellation)
                    checks(path) = check
                End If
                Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
                EnsureAvailable()
                Dim status As System.String = document.ExtractionStatus
                Dim message As System.String = document.ExtractionMessage
                If check.Status.Length > 0 Then
                    status = check.Status
                    message = check.Message
                ElseIf check.Hash <> document.SourceHash Then
                    status = "changed"
                    message = "The original bytes differ from the saved source. Refresh the document before analyzing it."
                ElseIf document.Segments.Count > 0 AndAlso (document.StoredExtractionStatus = "complete" OrElse document.StoredExtractionStatus = "partial") Then
                    status = document.StoredExtractionStatus
                    message = document.StoredExtractionMessage
                End If
                If status <> document.ExtractionStatus OrElse message <> document.ExtractionMessage Then
                    If candidate Is Nothing Then candidate = ReviewContract.Clone(_project)
                    Dim edited As ReviewDocument = candidate.Documents.Find(Function(item As ReviewDocument) item.Id = document.Id)
                    edited.ExtractionStatus = status
                    edited.ExtractionMessage = message
                    If status <> "complete" AndAlso status <> "partial" Then ReviewContract.Invalidate(edited)
                End If
            Next
            If candidate IsNot Nothing Then Commit(candidate, "source_verification", "Checked current source hashes and retained the saved text; matching originals reuse the cached extraction.")
        End Function

        Private Function ConsistencyEngine(modelKey As System.String) As ReviewContradictionEngine
            Dim choice As ReviewModelChoice = _models.Find(Function(item As ReviewModelChoice) item.Key = modelKey)
            Dim invoke As System.Func(Of System.String, System.String, System.Threading.CancellationToken, System.Threading.Tasks.Task(Of System.String)) = ReviewModelCatalog.CreateInvoker(_context, choice)
            Dim budget As System.Int32 = If(_context.INI_TextAnalyzeMaxSourceChars >= 2048, System.Math.Min(28000, _context.INI_TextAnalyzeMaxSourceChars), 28000)
            Return New ReviewContradictionEngine(
                Function(key As System.String, systemPrompt As System.String, requestJson As System.String, cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of System.String)
                    If key <> modelKey Then Throw New System.InvalidOperationException("The model for this operation changed. Plan the check again.")
                    Return invoke(systemPrompt, requestJson, cancellation)
                End Function, budget)
        End Function

        Private Async Function PlanConsistencyAsync(request As Newtonsoft.Json.Linq.JObject, cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of ReviewConsistencyRun)
            Dim modelKey As System.String = ResolveOperationModel(request)
            Dim documents As System.Collections.Generic.List(Of ReviewDocument) = SelectedDocuments(request)
            If documents.Count < 2 Then Throw New System.InvalidOperationException("Select at least two documents for a cross-document check.")
            Dim instruction As System.String = TextField(request, "instruction")
            If System.String.IsNullOrWhiteSpace(instruction) Then Throw New System.InvalidOperationException("Enter the prompt to run over the selected documents.")
            _message = "Checking the selected sources and planning the combined prompt. No model call has started."
            SendState()
            Await VerifySelectedSourcesAsync(documents, cancellation)
            Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
            EnsureAvailable()
            documents = SelectedDocuments(request)
            Dim engine As ReviewContradictionEngine = ConsistencyEngine(modelKey)
            Dim language As System.String = _project.OutputLanguage
            Dim criterion As ReviewCriterion = If(TextField(request, "groupCriterionId").Length > 0, FindCriterion(TextField(request, "groupCriterionId")), Nothing)
            Dim plan As ReviewConsistencyRun = Await System.Threading.Tasks.Task.Run(Function() engine.PlanPrompt(documents, instruction, language, modelKey, cancellation, criterion), cancellation)
            Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
            EnsureAvailable()
            cancellation.ThrowIfCancellationRequested()
            _consistencyPlanId = plan.Id
            _consistencyPlanRevision = _project.Revision
            _consistencyPlanScope = ReviewContract.Serialize(StringList(request("documentIds")))
            _consistencyPlanModel = modelKey
            _consistencyPlanInstruction = instruction
            _message = plan.CoverageMessage
            Return plan
        End Function

        Private Async Function ReviewGroupsAsync(request As Newtonsoft.Json.Linq.JObject, cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            Dim requested As System.String = TextField(request, "groupName")
            Dim groups As New System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of System.String))(System.StringComparer.Ordinal)
            For Each document As ReviewDocument In _project.Documents
                If document.Archived OrElse System.String.IsNullOrWhiteSpace(document.GroupName) OrElse (requested.Length > 0 AndAlso document.GroupName <> requested) Then Continue For
                If Not groups.ContainsKey(document.GroupName) Then groups.Add(document.GroupName, New System.Collections.Generic.List(Of System.String)())
                groups(document.GroupName).Add(document.Id)
            Next
            Dim completed As System.Int32
            For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, System.Collections.Generic.List(Of System.String)) In groups
                If pair.Value.Count < 2 Then Continue For
                cancellation.ThrowIfCancellationRequested()
                Dim scope As Newtonsoft.Json.Linq.JObject = DirectCast(request.DeepClone(), Newtonsoft.Json.Linq.JObject)
                scope("groupName") = New Newtonsoft.Json.Linq.JValue(pair.Key)
                scope("documentIds") = Newtonsoft.Json.Linq.JArray.FromObject(pair.Value)
                If BooleanField(request, "perQuestion") Then
                    Dim planned As System.Int32
                    For Each criterion As ReviewCriterion In _project.Criteria
                        If criterion.Archived OrElse criterion.Kind = "text" OrElse System.String.IsNullOrWhiteSpace(criterion.Instruction) Then Continue For
                        If Not pair.Value.Exists(Function(id As System.String) criterion.AppliesTo(FindDocument(id))) Then Continue For
                        scope("groupCriterionId") = New Newtonsoft.Json.Linq.JValue(criterion.Id)
                        scope("instruction") = New Newtonsoft.Json.Linq.JValue(criterion.Instruction)
                        Await RunGroupPromptAsync(scope, cancellation)
                        planned += 1
                    Next
                    If planned = 0 Then Throw New System.InvalidOperationException("No applicable AI questions for group " & pair.Key)
                Else
                    Await RunGroupPromptAsync(scope, cancellation)
                End If
                cancellation.ThrowIfCancellationRequested()
                completed += 1
            Next
            If completed = 0 Then Throw New System.InvalidOperationException("Assign at least two documents to a named group before running this review.")
            _message = "Reviewed groups: " & completed.ToString() & ". See their summary rows in Excel and their source quotations in Findings."
        End Function

        Private Async Function RunGroupPromptAsync(scope As Newtonsoft.Json.Linq.JObject, cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            Dim plan As ReviewConsistencyRun = Await PlanConsistencyAsync(scope, cancellation)
            scope("planId") = New Newtonsoft.Json.Linq.JValue(plan.Id)
            Await CheckConsistencyAsync(scope, cancellation)
        End Function

        Private Async Function CheckConsistencyAsync(request As Newtonsoft.Json.Linq.JObject, cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            Dim modelKey As System.String = ResolveOperationModel(request)
            Dim instruction As System.String = TextField(request, "instruction")
            If TextField(request, "planId") <> _consistencyPlanId OrElse _consistencyPlanId.Length = 0 OrElse
               _consistencyPlanRevision <> _project.Revision OrElse _consistencyPlanModel <> modelKey OrElse _consistencyPlanInstruction <> instruction OrElse
               _consistencyPlanScope <> ReviewContract.Serialize(StringList(request("documentIds"))) Then
                Throw New System.InvalidOperationException("The prompt, selected documents, model or review changed. Preview the work again before starting.")
            End If
            Dim plannedRevision As System.Int64 = _project.Revision
            Dim documents As System.Collections.Generic.List(Of ReviewDocument) = SelectedDocuments(request)
            Await VerifySelectedSourcesAsync(documents, cancellation)
            Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
            EnsureAvailable()
            If _project.Revision <> plannedRevision Then Throw New System.InvalidOperationException("The source availability changed. Preview the work again.")
            _consistencyPlanId = ""
            documents = SelectedDocuments(request)
            Dim engine As ReviewContradictionEngine = ConsistencyEngine(modelKey)
            Dim progress As New System.Progress(Of System.String)(
                Sub(message As System.String)
                    If Not _closed AndAlso Not _hostLifetime.IsCancellationRequested AndAlso Not cancellation.IsCancellationRequested Then
                        _message = message
                        SendState()
                    End If
                End Sub)
            Dim language As System.String = _project.OutputLanguage
            Dim criterion As ReviewCriterion = If(TextField(request, "groupCriterionId").Length > 0, FindCriterion(TextField(request, "groupCriterionId")), Nothing)
            Dim groupName As System.String = TextField(request, "groupName")
            If groupName.Length > 0 Then
                Dim members As System.Collections.Generic.List(Of ReviewDocument) = _project.Documents.FindAll(Function(item As ReviewDocument) Not item.Archived AndAlso item.GroupName = groupName)
                If members.Count <> documents.Count OrElse Not documents.TrueForAll(Function(item As ReviewDocument) item.GroupName = groupName) Then Throw New System.InvalidOperationException("The group membership changed. Start the group review again.")
            End If
            Dim run As ReviewConsistencyRun = Await System.Threading.Tasks.Task.Run(
                Function() engine.CheckPromptAsync(documents, instruction, language, modelKey, cancellation,
                    Function(progressRun As ReviewConsistencyRun, finding As ReviewFinding)
                        progressRun.GroupName = groupName
                        If criterion IsNot Nothing Then
                            progressRun.GroupCriterionId = criterion.Id
                            progressRun.GroupCriterionRevision = criterion.Revision
                        End If
                        Return PersistConsistencyAsync(progressRun, finding)
                    End Function, Sub(message As System.String) DirectCast(progress, System.IProgress(Of System.String)).Report(message), criterion), cancellation)
            Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
            EnsureAvailable()
            If run.PersistenceFailed Then Throw New System.InvalidOperationException("Excel could not confirm a combined-review checkpoint. Reload Excel before continuing; completed checkpoints may already be present. " & run.Error)
            ' Completed findings already passed source checks at their commit. Stop
            ' must not launch another full-file pass after the operator cancels.
            If Not cancellation.IsCancellationRequested Then
                Await VerifySelectedSourcesAsync(SelectedDocuments(request), cancellation)
                Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
                EnsureAvailable()
            End If
            Dim saved As ReviewConsistencyRun = _project.ConsistencyRuns.Find(Function(item As ReviewConsistencyRun) item.Id = run.Id)
            Dim changed As System.Boolean = run.Documents.Exists(
                Function(coverage As ReviewConsistencyDocumentCoverage)
                    Dim document As ReviewDocument = _project.Documents.Find(Function(item As ReviewDocument) item.Id = coverage.DocumentId)
                    Return document Is Nothing OrElse document.SourceHash <> coverage.SourceHash OrElse document.TextHash <> coverage.TextHash OrElse
                        (document.ExtractionStatus <> "complete" AndAlso document.ExtractionStatus <> "partial")
                End Function)
            If changed AndAlso saved IsNot Nothing AndAlso saved.Status = "complete" Then
                Dim candidate As ReviewProject = ReviewContract.Clone(_project)
                Dim edited As ReviewConsistencyRun = candidate.ConsistencyRuns.Find(Function(item As ReviewConsistencyRun) item.Id = run.Id)
                edited.Status = "partial"
                edited.CoverageMessage &= " A source changed or became unavailable during the check. Restore/refresh it and rerun the scope before relying on completeness."
                Commit(candidate, "consistency_source_changed", edited.CoverageMessage)
                saved = edited
            End If
            _message = If(saved Is Nothing, run.CoverageMessage, saved.CoverageMessage)
            If Not System.String.IsNullOrWhiteSpace(run.Error) Then Throw New System.InvalidOperationException(run.Error & " " & run.CoverageMessage)
        End Function

        Private Async Function PersistConsistencyAsync(run As ReviewConsistencyRun, finding As ReviewFinding) As System.Threading.Tasks.Task
            Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
            EnsureAvailable()
            If finding IsNot Nothing Then
                Dim cited As New System.Collections.Generic.List(Of ReviewDocument)()
                Dim ids As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
                If finding.Kind = "review_finding" Then
                    For Each coverage As ReviewConsistencyDocumentCoverage In run.Documents
                        If ids.Add(coverage.DocumentId) Then cited.Add(FindDocument(coverage.DocumentId))
                    Next
                Else
                    For Each citation As ReviewCrossDocumentEvidence In AllFindingEvidence(finding)
                        If ids.Add(citation.DocumentId) Then cited.Add(FindDocument(citation.DocumentId))
                    Next
                End If
                ' Honor completed model work after Stop, but never after host shutdown.
                Await VerifySelectedSourcesAsync(cited, _lifetime.Token)
                Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
                EnsureAvailable()
            End If
            Dim candidate As ReviewProject = ReviewContract.Clone(_project)
            Dim storedRun As ReviewConsistencyRun = Newtonsoft.Json.JsonConvert.DeserializeObject(Of ReviewConsistencyRun)(ReviewContract.Serialize(run), ReviewContract.JsonSettings)
            Dim index As System.Int32 = candidate.ConsistencyRuns.FindIndex(Function(item As ReviewConsistencyRun) item.Id = storedRun.Id)
            If index < 0 Then
                candidate.ConsistencyRuns.Add(storedRun)
            Else
                candidate.ConsistencyRuns(index) = storedRun
            End If
            If finding IsNot Nothing Then
                Dim storedFinding As ReviewFinding = Newtonsoft.Json.JsonConvert.DeserializeObject(Of ReviewFinding)(ReviewContract.Serialize(finding), ReviewContract.JsonSettings)
                Dim findingIndex As System.Int32 = candidate.Findings.FindIndex(Function(item As ReviewFinding) item.Id = storedFinding.Id)
                If findingIndex < 0 Then
                    candidate.Findings.Add(storedFinding)
                Else
                    Dim previous As ReviewFinding = candidate.Findings(findingIndex)
                    If previous.UserEdited Then Throw New System.InvalidOperationException("A user edited this finding while the check was running. Keep that edit and start a new check.")
                    storedFinding.History = previous.History
                    storedFinding.Archived = previous.Archived
                    If previous.Status = "stale" Then storedFinding.Status = "stale"
                    candidate.Findings(findingIndex) = storedFinding
                End If
            End If
            Commit(candidate, "consistency_" & storedRun.Status, storedRun.CoverageMessage)
        End Function

        Private Shared Function ReadReportDefinitions(token As Newtonsoft.Json.Linq.JToken) As System.Collections.Generic.List(Of ReviewReportDefinition)
            If token Is Nothing OrElse token.Type <> Newtonsoft.Json.Linq.JTokenType.Array Then Throw New System.IO.InvalidDataException("Report definitions must be an array.")
            Return Newtonsoft.Json.JsonConvert.DeserializeObject(Of System.Collections.Generic.List(Of ReviewReportDefinition))(token.ToString(), ReviewContract.JsonSettings)
        End Function

        Private Sub SaveReportDefinition(request As Newtonsoft.Json.Linq.JObject, delete As System.Boolean)
            Dim candidate As ReviewProject = ReviewContract.Clone(_project)
            Dim report As ReviewReportDefinition = Newtonsoft.Json.JsonConvert.DeserializeObject(Of ReviewReportDefinition)(request("definition").ToString(), ReviewContract.JsonSettings)
            If report Is Nothing Then Throw New System.IO.InvalidDataException("Choose a report definition.")
            If delete Then
                candidate.ReportDefinitions.RemoveAll(Function(item As ReviewReportDefinition) item.Id = report.Id)
            Else
                Dim index As System.Int32 = candidate.ReportDefinitions.FindIndex(Function(item As ReviewReportDefinition) item.Id = report.Id)
                If index < 0 Then
                    candidate.ReportDefinitions.Add(report)
                Else
                    candidate.ReportDefinitions(index) = report
                End If
            End If
            Commit(candidate, "report_definition", ReviewContract.Serialize(New With {.previous = _project.ReportDefinitions, .current = candidate.ReportDefinitions}), actor:=RequireReviewer(request))
        End Sub

        Private Function BuildMatrixReportData(request As Newtonsoft.Json.Linq.JObject) As ReviewReportData
            Dim definition As New ReviewReportDefinition With {.Name = TextField(request, "title"), .Kind = TextField(request, "kind"),
                .Orientation = TextField(request, "orientation", "landscape"), .Design = TextField(request, "design", "redink"), .PaperSize = TextField(request, "paperSize", "A4"), .Fields = ExactStringList(request("fields"))}
            ReviewContract.ValidateReportDefinitions(New System.Collections.Generic.List(Of ReviewReportDefinition) From {definition}, ActiveCriterionIds(_project))
            Dim data As New ReviewReportData With {.Title = definition.Name, .Kind = definition.Kind, .Orientation = definition.Orientation,
                .ProjectId = _project.Id, .ProjectRevision = _project.Revision, .Design = definition.Design, .PaperSize = definition.PaperSize}
            data.TableLayout = TextField(request, "tableLayout", "combined")
            If data.TableLayout <> "combined" AndAlso data.TableLayout <> "types" Then Throw New System.IO.InvalidDataException("Unknown report layout.")
            Dim ids As System.Collections.Generic.List(Of System.String) = ExactStringList(request("documentIds"))
            If ids.Count = 0 Then Throw New System.IO.InvalidDataException("Choose at least one document for the report.")
            Dim selected As System.Collections.Generic.List(Of ReviewDocument) = ids.ConvertAll(Function(key As System.String) FindDocument(key))
            If selected.Exists(Function(doc As ReviewDocument) Not System.String.IsNullOrEmpty(doc.GroupName)) AndAlso Not definition.Fields.Contains("document") Then definition.Fields.Insert(0, "document")
            For Each key As System.String In definition.Fields
                Dim column As New ReviewReportColumn With {.Key = key, .Title = key}
                Select Case key
                    Case "number" : column.Title = "Document reference" : column.Width = 14
                    Case "document" : column.Title = "Document" : column.Width = 38
                    Case "category" : column.Title = "Category"
                    Case "group" : column.Title = "Group"
                    Case "location" : column.Title = "Location in source folder"
                    Case "status" : column.Title = "Review status" : column.Width = 22
                    Case "priority" : column.Title = "Priority" : column.Width = 20
                    Case "flags" : column.Title = "Flags" : column.Width = 22
                    Case "reviewer" : column.Title = "Reviewer" : column.Width = 24
                    Case "notes" : column.Title = "Notes"
                    Case "sourceStatus" : column.Title = "Source status" : column.Width = 24
                    Case Else
                        Dim parts As System.String() = key.Split(":"c)
                        Dim criterion As ReviewCriterion = _project.Criteria.Find(Function(item As ReviewCriterion) item.Id = parts(1) AndAlso Not item.Archived)
                        column.Category = criterion.Category
                        column.Title = criterion.Title
                        Select Case parts(0)
                            Case "riskFlag" : column.Title &= " — Flag"
                            Case "assessment" : column.Title &= " — Assessment"
                            Case "answerStatus" : column.Title &= " — Status" : column.Width = 22
                            Case "evidence" : column.Title &= " — Sources" : column.Width = 60
                            Case "answer"
                                Dim width As System.Int32
                                If _project.View.ColumnWidths.TryGetValue(criterion.Id, width) Then column.Width = System.Math.Max(22, System.Math.Min(80, width / 7.0R))
                        End Select
                End Select
                data.Columns.Add(column)
            Next
            Dim ordered As New System.Collections.Generic.List(Of ReviewDocument)()
            Dim emitted As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            Dim summaries As System.Collections.Generic.List(Of ReviewDocument) = ReviewContract.CurrentGroupRows(_project)
            For Each doc As ReviewDocument In selected
                Dim name As System.String = doc.GroupName
                If System.String.IsNullOrEmpty(name) Then
                    ordered.Add(doc)
                ElseIf emitted.Add(name) Then
                    Dim members As System.Collections.Generic.List(Of ReviewDocument) = selected.FindAll(Function(item As ReviewDocument) item.GroupName = name)
                    Dim summary As ReviewDocument = summaries.Find(Function(item As ReviewDocument) item.GroupName = name)
                    If summary Is Nothing OrElse _project.Documents.Exists(Function(item As ReviewDocument) Not item.Archived AndAlso item.GroupName = name AndAlso Not ids.Contains(item.Id)) OrElse
                       _project.ConsistencyRuns.Exists(Function(run As ReviewConsistencyRun) run.GroupName = name AndAlso run.Documents.Exists(Function(entry As ReviewConsistencyDocumentCoverage) Not ids.Contains(entry.DocumentId))) Then
                        summary = New ReviewDocument With {.Id = "group:" & ReviewContract.Digest(name), .FileName = "Group — " & name, .GroupName = name, .Notes = "Group analysis omitted for this selection."}
                    End If
                    ordered.Add(summary)
                    ordered.AddRange(members)
                End If
            Next
            For Each document As ReviewDocument In ordered
                Dim id As System.String = document.Id
                Dim group As System.Boolean = id.StartsWith("group:", System.StringComparison.Ordinal)
                If document.Archived Then Throw New System.IO.InvalidDataException("Restore the selected document or remove it from the report selection.")
                data.Documents.Add(New ReviewReportDocument With {.Id = document.Id, .Number = document.Number, .FileName = document.FileName,
                    .RelativePath = document.RelativePath, .SourcePath = document.SourcePath, .SourceHash = document.SourceHash,
                    .ExtractionStatus = document.ExtractionStatus, .ExtractionMessage = document.ExtractionMessage})
                Dim row As New ReviewReportRow With {.DocumentId = id, .Category = document.Category, .IsGroup = group}
                For Each column As ReviewReportColumn In data.Columns
                    Dim value As System.String = ""
                    row.Styles.Add(If(group, 1, 0))
                    Select Case column.Key
                        Case "number" : If Not group Then value = ReviewContract.ShortDocumentReference(document.Number, document.Id)
                        Case "document"
                            value = document.FileName
                            If group AndAlso Not definition.Fields.Contains("notes") AndAlso document.Notes <> "" Then value &= System.Environment.NewLine & document.Notes
                            If Not group AndAlso document.GroupName <> "" Then value = "    ↳ " & If(data.TableLayout = "types", document.GroupName & " · ", "") & value
                        Case "category" : value = document.Category
                        Case "group" : value = document.GroupName
                        Case "location" : value = SafeReportLocation(document)
                        Case "status" : value = MatrixStatusLabel(document.ReviewStatus)
                        Case "priority" : value = document.Priority
                        Case "flags" : value = System.String.Join(System.Environment.NewLine, document.Flags)
                        Case "reviewer" : value = document.Reviewer
                        Case "notes" : value = document.Notes
                        Case "sourceStatus" : value = MatrixStatusLabel(document.ExtractionStatus)
                        Case Else
                            Dim parts As System.String() = column.Key.Split(":"c)
                            Dim answer As ReviewAnswer = Nothing
                            If Not group AndAlso Not FindCriterion(parts(1)).AppliesTo(document) Then
                                row.Values.Add("N/A")
                                Continue For
                            End If
                            If Not group AndAlso parts(0) = "answerStatus" Then value = MatrixStatusLabel("pending")
                            If document.Answers.TryGetValue(parts(1), answer) Then
                                Select Case parts(0)
                                    Case "answer" : value = answer.Text : row.Styles(row.Styles.Count - 1) = answer.TextStyle
                                    Case "riskFlag" : value = answer.RiskFlag
                                    Case "assessment" : value = answer.Assessment : row.Styles(row.Styles.Count - 1) = answer.AssessmentStyle
                                    Case "answerStatus" : value = MatrixStatusLabel(answer.Status)
                                    Case "evidence"
                                        Dim quotations As New System.Collections.Generic.List(Of System.String)()
                                        For Each evidence As ReviewEvidence In answer.Evidence
                                            Dim verification As System.String = evidence.MatchStatus
                                            If Not group AndAlso Not System.String.Equals(answer.SourceHash, document.SourceHash, System.StringComparison.Ordinal) Then verification &= "; earlier source version"
                                            quotations.Add(evidence.Locator & System.Environment.NewLine & evidence.Quote & System.Environment.NewLine & "[" & verification & "]")
                                        Next
                                        value = System.String.Join(System.Environment.NewLine & System.Environment.NewLine, quotations)
                                End Select
                            End If
                    End Select
                    row.Values.Add(If(value, ""))
                Next
                data.Rows.Add(row)
            Next
            Return data
        End Function

        Private Shared Function SafeReportLocation(document As ReviewDocument) As System.String
            If ReviewContract.IsSafeRelativePath(document.RelativePath) Then Return document.RelativePath
            Return ""
        End Function

        Private Shared Function MatrixStatusLabel(value As System.String) As System.String
            Select Case value
                Case "unreviewed" : Return "Unreviewed"
                Case "in_review" : Return "In progress"
                Case "reviewed" : Return "Reviewed"
                Case "pending" : Return "Pending"
                Case "draft" : Return "Draft"
                Case "stale" : Return "Stale"
                Case "error" : Return "Error"
                Case "complete" : Return "Extracted"
                Case "partial" : Return "Partial extraction"
                Case "missing" : Return "File missing"
                Case "changed" : Return "Source changed"
                Case "imported" : Return "Imported"
                Case Else : Return value
            End Select
        End Function

        Private Function BuildReportData(request As Newtonsoft.Json.Linq.JObject) As ReviewReportData
            Dim kind As System.String = TextField(request, "kind", "issues")
            If kind = "matrix_xlsx" OrElse kind = "matrix_docx" OrElse kind = "document_report" Then Return BuildMatrixReportData(request)
            Throw New System.IO.InvalidDataException("Choose a configurable Matrix or Documents with cover report. The earlier issue, Q&A and Word reports are no longer available.")
        End Function

        Private Function ReportModelLabel(modelKey As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(modelKey) Then Return "Manual content"
            If modelKey = "imported" Then Return "Imported content"
            Dim choice As ReviewModelChoice = _models.Find(Function(item As ReviewModelChoice) item.Key = modelKey)
            Return If(choice Is Nothing, "Previously selected model", choice.Label)
        End Function

        Private Shared Function ReportSourceCoverage(document As ReviewDocument) As System.String
            Select Case document.ExtractionStatus
                Case "complete"
                    Return "The review used the complete saved extract."
                Case "partial"
                    Return "The review used a partial extract. Check the original for missing content."
                Case "changed"
                    Return "The original has changed. Refresh and verify this answer before relying on it."
                Case "missing", "error"
                    Return "The original is unavailable or unreadable. The report contains saved review content; restore the source and verify it."
                Case Else
                    Return "The review has no complete verified extract for this source. Link or refresh it and verify the answer."
            End Select
        End Function

        Private Shared Function ContextLabel(relationship As System.String) As System.String
            Select Case relationship
                Case "supports"
                    Return "Supporting context"
                Case "qualifies"
                    Return "Qualification"
                Case "apparent_change"
                    Return "Apparent change"
                Case "ambiguity"
                    Return "Ambiguity"
                Case Else
                    Return "Additional context"
            End Select
        End Function

        Private Sub AddReportCitation(item As ReviewReportItem, citation As ReviewCrossDocumentEvidence, role As System.String,
                                      documents As System.Collections.Generic.HashSet(Of System.String))
            Dim document As ReviewDocument = FindDocument(citation.DocumentId)
            If Not item.DocumentIds.Contains(document.Id) Then item.DocumentIds.Add(document.Id)
            documents.Add(document.Id)
            Dim evidence As ReviewReportEvidence = ReportEvidence(citation.DocumentId, citation.SourceHash, citation.Evidence)
            evidence.Role = role
            If Not item.Evidence.Exists(Function(existing As ReviewReportEvidence) existing.DocumentId = evidence.DocumentId AndAlso
                existing.SourceHash = evidence.SourceHash AndAlso existing.SegmentId = evidence.SegmentId AndAlso existing.Start = evidence.Start AndAlso
                existing.Length = evidence.Length AndAlso existing.Quote = evidence.Quote AndAlso existing.Locator = evidence.Locator AndAlso
                existing.MatchStatus = evidence.MatchStatus AndAlso existing.Role = evidence.Role) Then item.Evidence.Add(evidence)
        End Sub

        Private Function ReportPriority(document As ReviewDocument) As System.String
            If Not System.String.IsNullOrWhiteSpace(document.Priority) Then Return document.Priority
            Dim values As New System.Collections.Generic.List(Of System.String)()
            For Each criterion As ReviewCriterion In _project.Criteria
                If criterion.Archived OrElse criterion.Kind <> "priority" Then Continue For
                Dim answer As ReviewAnswer = Nothing
                If document.Answers.TryGetValue(criterion.Id, answer) AndAlso Not System.String.IsNullOrWhiteSpace(answer.Text) Then values.Add(criterion.Title & ": " & answer.Text)
            Next
            Return System.String.Join(System.Environment.NewLine, values)
        End Function

        Private Shared Function ReportEvidence(documentId As System.String, sourceHash As System.String, evidence As ReviewEvidence) As ReviewReportEvidence
            Return New ReviewReportEvidence With {.DocumentId = documentId, .SourceHash = sourceHash, .SegmentId = evidence.SegmentId,
                .Locator = evidence.Locator, .Quote = evidence.Quote, .MatchStatus = evidence.MatchStatus, .Start = evidence.Start, .Length = evidence.Length}
        End Function

        Private NotInheritable Class TemplateCatalogState
            Public Property Catalog As ReviewTemplateLibraryResult
            Public Property Targets As New System.Collections.Generic.Dictionary(Of System.String, ReviewTemplateWriteTarget)(System.StringComparer.Ordinal)
            Public Property TargetMessages As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.Ordinal)
        End Class

        Private Shared Function ReadTemplateCatalog(personalPath As System.String, sharedPath As System.String,
                cancellation As System.Threading.CancellationToken) As TemplateCatalogState
            Dim result As New TemplateCatalogState With {.Catalog = ReviewTemplateService.LoadLibraries(personalPath, sharedPath, cancellation)}
            Dim paths As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.Ordinal) From {
                {"personal", personalPath}, {"shared", sharedPath}
            }
            For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, System.String) In paths
                cancellation.ThrowIfCancellationRequested()
                If System.String.IsNullOrWhiteSpace(pair.Value) Then
                    result.TargetMessages(pair.Key) = "Set " & If(pair.Key = "personal", "TabularReviewPathLocal", "TabularReviewPath") & " in Settings to save templates in this library. JSON import and export are also available."
                    Continue For
                End If
                Try
                    result.Targets(pair.Key) = ReviewTemplateService.ResolveWriteTarget(pair.Value, cancellation)
                Catch ex As System.OperationCanceledException
                    Throw
                Catch ex As System.Exception
                    result.TargetMessages(pair.Key) = ex.Message
                    System.Diagnostics.Trace.WriteLine("[TabularReview:template_target] scope=" & pair.Key & "; " & ex.ToString())
                End Try
            Next
            Return result
        End Function

        Private Function RegisterTemplateEntry(entry As ReviewTemplateLibraryEntry) As System.Object
            Dim key As System.String = System.Guid.NewGuid().ToString("N")
            _columnTemplateEntries.Add(key, entry)
            Return New With {.key = key, .label = entry.DisplayName, .name = entry.Template.Name,
                .description = entry.Template.Description, .columnCount = entry.Template.Columns.Count,
                .scope = entry.Scope, .source = System.IO.Path.GetFileName(entry.FilePath)}
        End Function

        Private Async Function ListColumnTemplatesAsync(action As System.String, cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of System.Object)
            Dim personalPath As System.String = If(_context.INI_TabularReviewPathLocal, "")
            Dim sharedPath As System.String = If(_context.INI_TabularReviewPath, "")
            Dim catalog As TemplateCatalogState = Await System.Threading.Tasks.Task.Run(Function() ReadTemplateCatalog(personalPath, sharedPath, cancellation), cancellation)
            Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
            EnsureAvailable()
            cancellation.ThrowIfCancellationRequested()
            _columnTemplateEntries.Clear()
            _columnTemplateTargets.Clear()
            Dim entries As New System.Collections.Generic.List(Of System.Object)()
            For Each entry As ReviewTemplateLibraryEntry In catalog.Catalog.Entries
                entries.Add(RegisterTemplateEntry(entry))
            Next
            Dim targets As New System.Collections.Generic.List(Of System.Object)()
            For Each scope As System.String In New System.String() {"personal", "shared"}
                Dim target As ReviewTemplateWriteTarget = Nothing
                Dim available As System.Boolean = catalog.Targets.TryGetValue(scope, target)
                If available Then _columnTemplateTargets.Add(scope, target)
                Dim message As System.String = ""
                catalog.TargetMessages.TryGetValue(scope, message)
                targets.Add(New With {.key = scope, .label = If(scope = "personal", "Personal library", "Shared library"),
                    .available = available, .message = If(message, "")})
            Next
            Dim warnings As New System.Collections.Generic.List(Of System.String)()
            For Each notice As ReviewTemplateLibraryNotice In catalog.Catalog.Notices
                warnings.Add(If(notice.Scope = "personal", "Personal", "Shared") & " library — " & notice.FileName & ": " & notice.Message)
            Next
            Return New With {.action = action, .entries = entries, .targets = targets, .warnings = warnings}
        End Function

        Private Function FindColumnTemplateEntry(key As System.String) As ReviewTemplateLibraryEntry
            Dim entry As ReviewTemplateLibraryEntry = Nothing
            If key.Length = 0 OrElse Not _columnTemplateEntries.TryGetValue(key, entry) Then Throw New System.InvalidOperationException("Refresh the template list and select the template again.")
            Return entry
        End Function

        Private Async Function LoadColumnTemplateAsync(request As Newtonsoft.Json.Linq.JObject, cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of System.Object)
            Dim entry As ReviewTemplateLibraryEntry = FindColumnTemplateEntry(TextField(request, "key"))
            Dim template As ReviewColumnTemplate = Await System.Threading.Tasks.Task.Run(Function() ReviewTemplateService.ReadEntry(entry, cancellation), cancellation)
            Return New With {.action = "load", .template = Newtonsoft.Json.Linq.JObject.Parse(ReviewTemplateService.SerializeTemplate(template))}
        End Function

        Private Async Function ImportColumnTemplatesAsync(cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of System.Object)
            Dim path As System.String = ""
            Using dialog As New System.Windows.Forms.OpenFileDialog With {.Title = "Import column templates",
                .Filter = "Tabular Review column templates (*.json)|*.json", .DefaultExt = "json", .Multiselect = False,
                .CheckFileExists = True, .CheckPathExists = True, .RestoreDirectory = True}
                Using Global.SharedLibrary.SharedLibrary.SharedMethods.PushDialogOwner(Me)
                    If dialog.ShowDialog(Me) <> System.Windows.Forms.DialogResult.OK Then Return New With {.action = "import", .cancelled = True}
                End Using
                path = NormalizeLocalPath(dialog.FileName)
            End Using
            EnsureAvailable()
            Dim snapshot As ReviewTemplateFileSnapshot = Await System.Threading.Tasks.Task.Run(Function() ReviewTemplateService.ReadFile(path, cancellation), cancellation)
            Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
            EnsureAvailable()
            cancellation.ThrowIfCancellationRequested()
            Dim oldImports As New System.Collections.Generic.List(Of System.String)()
            For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, ReviewTemplateLibraryEntry) In _columnTemplateEntries
                If pair.Value.Scope = "import" Then oldImports.Add(pair.Key)
            Next
            For Each key As System.String In oldImports
                _columnTemplateEntries.Remove(key)
            Next
            Dim entries As New System.Collections.Generic.List(Of System.Object)()
            For index As System.Int32 = 0 To snapshot.Templates.Count - 1
                Dim template As ReviewColumnTemplate = snapshot.Templates(index)
                entries.Add(RegisterTemplateEntry(New ReviewTemplateLibraryEntry With {
                    .DisplayName = template.Name & " (import — " & System.IO.Path.GetFileName(snapshot.FilePath) & ")",
                    .Scope = "import", .FilePath = snapshot.FilePath, .FileHash = snapshot.FileHash,
                    .Index = index, .Template = template
                }))
            Next
            Return New With {.action = "import", .entries = entries}
        End Function

        Private Shared Function ColumnTemplateFromRequest(request As Newtonsoft.Json.Linq.JObject) As ReviewColumnTemplate
            Dim definition As Newtonsoft.Json.Linq.JObject = TryCast(request("template"), Newtonsoft.Json.Linq.JObject)
            If definition Is Nothing Then Throw New System.IO.InvalidDataException("Choose the column settings to save as a template.")
            Return ReviewTemplateService.ParseTemplateJson(definition.ToString(Newtonsoft.Json.Formatting.None))
        End Function

        Private Function ProtectedTemplatePaths() As System.Collections.Generic.List(Of System.String)
            Dim paths As New System.Collections.Generic.List(Of System.String)()
            For Each document As ReviewDocument In _project.Documents
                If Not System.String.IsNullOrWhiteSpace(document.SourcePath) Then paths.Add(document.SourcePath)
                For Each version As ReviewSourceVersion In document.SourceHistory
                    If Not System.String.IsNullOrWhiteSpace(version.SourcePath) Then paths.Add(version.SourcePath)
                Next
            Next
            For Each report As ReviewReportFile In _project.ReportFiles
                If Not System.String.IsNullOrWhiteSpace(report.Path) Then paths.Add(report.Path)
            Next
            If _callbacks.WorkbookPath IsNot Nothing Then
                Dim workbook As System.String = _callbacks.WorkbookPath.Invoke()
                If Not System.String.IsNullOrWhiteSpace(workbook) Then paths.Add(workbook)
            End If
            Return paths
        End Function

        Private Shared Sub EnsureIndependentTemplateDestination(path As System.String, protectedPaths As System.Collections.Generic.IEnumerable(Of System.String))
            Dim fullPath As System.String = System.IO.Path.GetFullPath(path)
            Dim existing As System.Boolean = System.IO.File.Exists(fullPath)
            Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            For Each protectedPath As System.String In protectedPaths
                If System.String.IsNullOrWhiteSpace(protectedPath) OrElse Not seen.Add(protectedPath) Then Continue For
                If Not System.IO.Path.IsPathRooted(protectedPath) OrElse protectedPath.IndexOf("://", System.StringComparison.Ordinal) >= 0 Then
                    If existing Then Throw New System.IO.IOException("Choose a new template filename. An existing destination cannot be verified against a workbook or source stored outside the filesystem.")
                    Continue For
                End If
                If System.String.Equals(fullPath, System.IO.Path.GetFullPath(protectedPath), System.StringComparison.OrdinalIgnoreCase) Then Throw New System.IO.IOException("Choose a separate template file. Templates cannot replace the review workbook, an original source or an exported report.")
                If existing Then
                    Dim sameFile As System.Boolean
                    Try
                        sameFile = Global.SharedLibrary.SharedLibrary.GeneratedOutputRegistry.IsPhysicalPathAtOrBelow(protectedPath, fullPath)
                    Catch ex As System.Exception When TypeOf ex Is System.IO.IOException OrElse TypeOf ex Is System.UnauthorizedAccessException OrElse TypeOf ex Is System.ArgumentException OrElse TypeOf ex Is System.NotSupportedException
                        Throw New System.IO.IOException("Choose a new template filename. This existing destination could not be verified as separate from every protected review file.", ex)
                    End Try
                    If sameFile Then Throw New System.IO.IOException("Choose a separate template file. This destination resolves to a protected review file.")
                End If
            Next
            If existing AndAlso (System.IO.File.GetAttributes(fullPath) And System.IO.FileAttributes.ReparsePoint) <> 0 Then Throw New System.IO.IOException("Choose a new template filename instead of replacing a linked file.")
        End Sub

        Private Async Function SaveColumnTemplateAsync(request As Newtonsoft.Json.Linq.JObject, cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of System.Object)
            Dim template As ReviewColumnTemplate = ColumnTemplateFromRequest(request)
            Dim scope As System.String = TextField(request, "targetKey")
            If scope <> "personal" AndAlso scope <> "shared" Then Throw New System.IO.InvalidDataException("Choose the personal or shared template library.")
            Dim path As System.String
            Dim expectedHash As System.String
            Dim replaceName As System.String = Nothing
            Dim replaceKey As System.String = TextField(request, "replaceKey")
            If replaceKey.Length > 0 Then
                Dim entry As ReviewTemplateLibraryEntry = FindColumnTemplateEntry(replaceKey)
                If entry.Scope <> scope Then Throw New System.IO.InvalidDataException("Choose the same library as the template you want to update.")
                path = entry.FilePath
                expectedHash = entry.FileHash
                replaceName = entry.Template.Name
            Else
                Dim target As ReviewTemplateWriteTarget = Nothing
                If Not _columnTemplateTargets.TryGetValue(scope, target) Then Throw New System.InvalidOperationException("Configure this template library in Settings, then refresh the template list before saving.")
                path = target.FilePath
                expectedHash = target.FileHash
            End If
            Dim protectedPaths As System.Collections.Generic.List(Of System.String) = ProtectedTemplatePaths()
            Await System.Threading.Tasks.Task.Run(Function() ReviewTemplateService.SaveTemplate(path, template, expectedHash, cancellation, replaceName,
                Sub(destination As System.String) EnsureIndependentTemplateDestination(destination, protectedPaths)), cancellation)
            Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
            EnsureAvailable()
            _message = "Saved the column template. Apply the Columns draft separately to change this review."
            ' Publication has completed; a late Stop must not report the saved file as cancelled.
            Return Await ListColumnTemplatesAsync("save", _lifetime.Token)
        End Function

        Private Async Function ExportColumnTemplateAsync(request As Newtonsoft.Json.Linq.JObject, cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of System.Object)
            Dim template As ReviewColumnTemplate = ColumnTemplateFromRequest(request)
            Dim path As System.String = ""
            Using dialog As New System.Windows.Forms.SaveFileDialog With {.Title = "Export column settings",
                .Filter = "Tabular Review column settings (*.json)|*.json", .DefaultExt = "json", .FileName = "TabularReviewColumns.json",
                .AddExtension = True, .CheckPathExists = True, .OverwritePrompt = True, .RestoreDirectory = True}
                Using Global.SharedLibrary.SharedLibrary.SharedMethods.PushDialogOwner(Me)
                    If dialog.ShowDialog(Me) <> System.Windows.Forms.DialogResult.OK Then Return New With {.action = "export", .cancelled = True}
                End Using
                path = NormalizeLocalPath(dialog.FileName)
            End Using
            EnsureAvailable()
            Dim protectedPaths As System.Collections.Generic.List(Of System.String) = ProtectedTemplatePaths()
            Await System.Threading.Tasks.Task.Run(
                Function()
                    EnsureIndependentTemplateDestination(path, protectedPaths)
                    Dim target As ReviewTemplateWriteTarget = ReviewTemplateService.InspectDestination(path, cancellation)
                    Return ReviewTemplateService.ExportTemplate(target.FilePath, template, target.FileHash, cancellation,
                        Sub(destination As System.String) EnsureIndependentTemplateDestination(destination, protectedPaths))
                End Function, cancellation)
            Return New With {.action = "export", .fileName = System.IO.Path.GetFileName(path)}
        End Function

        Private Shared Function OptionalColumnWidth(value As Newtonsoft.Json.Linq.JToken, documentColumn As System.Boolean) As System.Nullable(Of System.Int32)
            If value Is Nothing OrElse value.Type = Newtonsoft.Json.Linq.JTokenType.Null Then Return Nothing
            Dim width As System.Int32
            If value.Type <> Newtonsoft.Json.Linq.JTokenType.Integer OrElse Not System.Int32.TryParse(value.ToString(), System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, width) Then Throw New System.IO.InvalidDataException("A column width must be a whole number.")
            If width < If(documentColumn, 160, 140) OrElse width > If(documentColumn, 600, 1200) Then Throw New System.IO.InvalidDataException("A column width is outside the supported range.")
            Return width
        End Function


        Private Shared Function ReportFileStem(title As System.String) As System.String
            Dim name As System.String = If(title, "").Trim()
            For Each invalid As System.Char In System.IO.Path.GetInvalidFileNameChars()
                name = name.Replace(invalid, "_"c)
            Next
            name = name.TrimEnd("."c, " "c)
            While name.EndsWith(".xlsx", System.StringComparison.OrdinalIgnoreCase) OrElse name.EndsWith(".docx", System.StringComparison.OrdinalIgnoreCase)
                name = name.Substring(0, name.Length - 5).TrimEnd("."c, " "c)
            End While
            If name.Length = 0 Then name = "Tabular Review"
            If name.Length > 100 Then name = name.Substring(0, If(System.Char.IsHighSurrogate(name(99)), 99, 100)).TrimEnd("."c, " "c)
            Return name
        End Function

        Private Async Function ExportReportAsync(request As Newtonsoft.Json.Linq.JObject, cancellation As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of ReviewReportExportResult)
            Dim data As ReviewReportData = BuildReportData(request)
            Dim selectedSources As New System.Collections.Generic.List(Of ReviewDocument)()
            For Each reference As ReviewReportDocument In data.Documents
                selectedSources.Add(FindDocument(reference.Id))
            Next
            _message = "Checking the selected sources before exporting their saved findings."
            SendState()
            Await VerifySelectedSourcesAsync(selectedSources, cancellation)
            Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
            EnsureAvailable()
            data = BuildReportData(request)
            Dim extension As System.String = If(data.Kind = "report" OrElse data.Kind = "matrix_docx" OrElse data.Kind = "document_report", ".docx", ".xlsx")
            Dim name As System.String = ReportFileStem(data.Title)
            Dim target As System.String = ""
            Using dialog As New System.Windows.Forms.SaveFileDialog With {.Title = "Save report",
                .Filter = If(extension = ".docx", "Word report (*.docx)|*.docx", "Excel report (*.xlsx)|*.xlsx"),
                .FileName = name & extension, .DefaultExt = extension.TrimStart("."c),
                .AddExtension = True, .CheckPathExists = True, .OverwritePrompt = True, .RestoreDirectory = True}
                If System.IO.Directory.Exists(_project.SourceRoot) Then dialog.InitialDirectory = _project.SourceRoot
                Using Global.SharedLibrary.SharedLibrary.SharedMethods.PushDialogOwner(Me)
                    If dialog.ShowDialog(Me) <> System.Windows.Forms.DialogResult.OK Then
                        Return New ReviewReportExportResult With {.Message = "Export cancelled. Your selected findings remain in the review."}
                    End If
                End Using
                target = NormalizeLocalPath(dialog.FileName)
            End Using
            EnsureAvailable()
            Dim protectedPaths As New System.Collections.Generic.List(Of System.String)()
            For Each document As ReviewDocument In _project.Documents
                If Not System.String.IsNullOrWhiteSpace(document.SourcePath) Then protectedPaths.Add(document.SourcePath)
                For Each version As ReviewSourceVersion In document.SourceHistory
                    If Not System.String.IsNullOrWhiteSpace(version.SourcePath) Then protectedPaths.Add(version.SourcePath)
                Next
            Next
            If _callbacks.WorkbookPath IsNot Nothing Then
                Dim workbook As System.String = _callbacks.WorkbookPath.Invoke()
                If workbook.Length > 0 Then protectedPaths.Add(workbook)
            End If
            _message = "Creating a separate report from the selected findings and saved source quotations."
            SendState()
            Dim overwrite As System.Boolean = System.IO.File.Exists(target)
            Dim result As ReviewReportExportResult = Await System.Threading.Tasks.Task.Run(
                Function() TabularReviewReportService.Export(data, target, overwrite, protectedPaths, cancellation), cancellation)
            Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
            EnsureAvailable()
            ' Publication already succeeded. A later workbook conflict must not rerun
            ' the export or hide the completed file from its user.
            Try
                Dim candidate As ReviewProject = ReviewContract.Clone(_project)
                candidate.ReportFiles.RemoveAll(Function(item As ReviewReportFile) System.String.Equals(item.Path, result.Path, System.StringComparison.OrdinalIgnoreCase))
                candidate.ReportFiles.Add(New ReviewReportFile With {.Path = result.Path, .RelativePath = ContainedRelativePath(candidate.SourceRoot, result.Path), .Kind = data.Kind})
                Commit(candidate, "report_export", "Created a separate " & data.Kind & " export containing " & result.ItemCount.ToString(System.Globalization.CultureInfo.InvariantCulture) & " selected report entries.")
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine("[TabularReview:report_registry_failed] " & ex.ToString())
                result.Message &= " The file is ready, but Excel could not record its export reference. Keep the file and reload Excel before further edits."
            End Try
            _message = result.Message
            Return result
        End Function

        Private Async Function ShowSourceAsync(request As Newtonsoft.Json.Linq.JObject, requestId As System.String) As System.Threading.Tasks.Task
            If _sourceRequest IsNot Nothing Then _sourceRequest.Cancel()
            Dim cancellation As System.Threading.CancellationTokenSource = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(_lifetime.Token)
            _sourceRequest = cancellation
            Try
                Dim document As ReviewDocument = FindDocument(TextField(request, "documentId"))
                Dim snapshot As ReviewDocument = Newtonsoft.Json.JsonConvert.DeserializeObject(Of ReviewDocument)(ReviewContract.Serialize(document), ReviewContract.JsonSettings)
                Dim evidence As ReviewEvidence = Nothing
                Dim evidenceId As System.String = TextField(request, "evidenceId")
                If evidenceId.Length > 0 Then
                    Dim findingId As System.String = TextField(request, "findingId")
                    If findingId.Length > 0 Then
                        Dim finding As ReviewFinding = _project.Findings.Find(Function(item As ReviewFinding) item.Id = findingId)
                        If finding Is Nothing Then Throw New System.IO.InvalidDataException("This finding no longer exists.")
                        Dim citation As ReviewCrossDocumentEvidence = AllFindingEvidence(finding).Find(Function(item As ReviewCrossDocumentEvidence) item.DocumentId = document.Id AndAlso item.Evidence.Id = evidenceId)
                        If citation Is Nothing Then Throw New System.IO.InvalidDataException("This finding does not cite the selected source.")
                        evidence = citation.Evidence
                        If snapshot.SourceHash <> citation.SourceHash OrElse snapshot.TextHash <> citation.TextHash Then
                            Dim version As ReviewSourceVersion = snapshot.SourceHistory.Find(Function(item As ReviewSourceVersion) item.SourceHash = citation.SourceHash AndAlso item.TextHash = citation.TextHash)
                            If version Is Nothing Then Throw New System.IO.InvalidDataException("The finding's saved source version is unavailable.")
                            snapshot.Segments = version.Segments
                            snapshot.SourceHash = version.SourceHash
                            snapshot.TextHash = version.TextHash
                            snapshot.ExtractionStatus = version.ExtractionStatus
                            snapshot.ExtractionMessage = "Saved source version: " & version.ArchivedUtc & ". " & version.ExtractionMessage
                        End If
                    Else
                        Dim criterion As ReviewCriterion = FindCriterion(TextField(request, "criterionId"))
                        Dim answer As ReviewAnswer = Nothing
                        If Not snapshot.Answers.TryGetValue(criterion.Id, answer) Then Throw New System.IO.InvalidDataException("This answer has no source citation.")
                        evidence = answer.Evidence.Find(Function(item As ReviewEvidence) item.Id = evidenceId)
                    End If
                    If evidence Is Nothing Then Throw New System.IO.InvalidDataException("This citation is no longer current.")
                    If evidence.MatchStatus <> "exact" AndAlso evidence.MatchStatus <> "whitespace" Then Throw New System.IO.InvalidDataException("This citation has not passed source validation.")
                End If
                ' Deliver the persisted text immediately; native preview conversion can take longer.
                SendSource(requestId, snapshot, evidence, "", "Preparing original layout…", False, "info")
                Dim preview As ReviewPreviewResult = Nothing
                Dim failure As System.String = ""
                Try
                    preview = Await ReviewPreviewService.CreateAsync(snapshot, evidence, System.IO.Path.Combine(_assetRoot, "preview"), cancellation.Token)
                Catch ex As System.OperationCanceledException
                    Throw
                Catch ex As System.Exception
                    failure = ex.Message
                End Try
                Await Global.SharedLibrary.SharedLibrary.SharedMethods.SwitchToHostUiThread()
                EnsureAvailable()
                cancellation.Token.ThrowIfCancellationRequested()
                Dim url As System.String = ""
                If preview IsNot Nothing AndAlso preview.FilePath.Length > 0 Then
                    Dim prefix As System.String = System.IO.Path.GetFullPath(System.IO.Path.Combine(_assetRoot, "preview")) & System.IO.Path.DirectorySeparatorChar
                    Dim full As System.String = System.IO.Path.GetFullPath(preview.FilePath)
                    If Not full.StartsWith(prefix, System.StringComparison.OrdinalIgnoreCase) OrElse System.IO.Path.GetExtension(full).ToLowerInvariant() <> ".pdf" Then Throw New System.IO.InvalidDataException("The preview adapter returned an invalid local file.")
                    Dim relative As System.String = "/preview/" & System.Uri.EscapeDataString(System.IO.Path.GetFileName(full))
                    _allowedPaths.Add(relative)
                    url = _origin.TrimEnd("/"c) & relative & "#page=" & System.Math.Max(1, preview.PageNumber).ToString(System.Globalization.CultureInfo.InvariantCulture) & "&view=FitH"
                End If
                Dim message As System.String = If(preview Is Nothing, "Original layout unavailable. " & failure & " The stored text remains available.", preview.Message)
                SendSource(requestId, snapshot, evidence, url, message, preview IsNot Nothing AndAlso preview.Highlighted,
                    If(preview Is Nothing, "warning", preview.MessageKind))
                If url.Length = 0 OrElse (evidence IsNot Nothing AndAlso Not preview.Highlighted) Then
                    System.Diagnostics.Debug.WriteLine("[TabularReview:preview_fallback] document=" & snapshot.Id & "; " & message)
                End If
            Finally
                If System.Object.ReferenceEquals(_sourceRequest, cancellation) Then _sourceRequest = Nothing
                cancellation.Dispose()
            End Try
        End Function

        Private Sub SendSource(requestId As System.String, document As ReviewDocument, evidence As ReviewEvidence,
                               previewUrl As System.String, previewMessage As System.String, highlighted As System.Boolean,
                               Optional previewMessageKind As System.String = "warning")
            Post(New With {.command = "source", .request = requestId, .documentId = document.Id, .segments = document.Segments,
                .evidence = evidence, .previewUrl = previewUrl, .previewMessage = previewMessage, .previewMessageKind = previewMessageKind, .highlighted = highlighted,
                .sourceStatus = document.ExtractionStatus, .sourceNotice = document.ExtractionMessage})
        End Sub

        Private Sub OpenOriginal(document As ReviewDocument)
            Dim path As System.String = NormalizeLocalPath(document.SourcePath)
            If Not System.IO.File.Exists(path) Then Throw New System.IO.FileNotFoundException("The source file is unavailable. Link it again or use its stored text.")
            If Not ReviewSourceService.IsSupportedFile(path, _context.INI_AllowLegacyDocFiles) Then Throw New System.InvalidOperationException("Open this file yourself in its native application; the review only launches supported document types.")
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(path) With {.UseShellExecute = True})
        End Sub

        Private Sub SendState()
            If Me.InvokeRequired Then
                If Me.IsHandleCreated AndAlso Not Me.IsDisposed Then Me.BeginInvoke(New System.Action(Sub() SendState()))
                Return
            End If
            If Not _ready OrElse _closed OrElse _web.CoreWebView2 Is Nothing Then Return
            Try
                Dim serializer As Newtonsoft.Json.JsonSerializer = Newtonsoft.Json.JsonSerializer.Create(ReviewContract.JsonSettings)
                ' Do not materialize every cached source string just to remove it again.
                ' The workbook serializer still retains the complete source and its history.
                serializer.ContractResolver = _browserStateResolver
                Dim state As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.FromObject(_project, serializer)
                state.Remove("groupEdits")
                state("groupRows") = Newtonsoft.Json.Linq.JArray.FromObject(ReviewContract.CurrentGroupRows(_project), serializer)
                Dim documents As Newtonsoft.Json.Linq.JArray = DirectCast(state("documents"), Newtonsoft.Json.Linq.JArray)
                For index As System.Int32 = 0 To documents.Count - 1
                    Dim document As Newtonsoft.Json.Linq.JObject = DirectCast(documents(index), Newtonsoft.Json.Linq.JObject)
                    document("segmentCount") = _project.Documents(index).Segments.Count
                    Dim characters As System.Int64 = 0
                    For Each segment As ReviewSourceSegment In _project.Documents(index).Segments
                        characters += segment.Text.Length
                    Next
                    document("sourceCharacters") = characters
                Next
                If _cachedLogoDataUrl Is Nothing Then _cachedLogoDataUrl = LogoDataUrl()
                Post(New With {.command = "state", .project = state, .models = _models, .logo = _cachedLogoDataUrl,
                    .workbook = _callbacks.WorkbookLabel.Invoke(), .busy = _busy OrElse _savingWorkbook,
                    .operationBusy = _busy, .saveBusy = _savingWorkbook, .message = _message,
                    .persistence = PersistenceState(), .portability = PortabilityState()})
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("[TabularReview:state] " & ex.ToString())
            End Try
        End Sub

        Private Shared Function LogoDataUrl() As System.String
            Dim sharedImage As System.Drawing.Bitmap = Global.SharedLibrary.SharedLibrary.SharedMethods.GetLogoBitmap(Global.SharedLibrary.SharedLibrary.SharedMethods.LogoType.Standard, True)
            If sharedImage Is Nothing Then Return ""
            Using bitmap As System.Drawing.Bitmap = DirectCast(sharedImage.Clone(), System.Drawing.Bitmap)
                Using stream As New System.IO.MemoryStream()
                    bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Png)
                    Return "data:image/png;base64," & System.Convert.ToBase64String(stream.ToArray())
                End Using
            End Using
        End Function

        Private Sub Post(value As System.Object)
            If Me.InvokeRequired Then
                If Me.IsHandleCreated AndAlso Not Me.IsDisposed Then Me.BeginInvoke(New System.Action(Sub() Post(value)))
                Return
            End If
            If _closed OrElse Me.IsDisposed OrElse _web.IsDisposed OrElse _web.CoreWebView2 Is Nothing Then Return
            _web.CoreWebView2.PostWebMessageAsJson(ReviewContract.Serialize(value))
        End Sub

        Private Sub SendAck(requestId As System.String, success As System.Boolean, message As System.String,
                            Optional selectedFolder As System.String = Nothing, Optional saveResult As ReviewWorkbookSaveResult = Nothing,
                            Optional reportResult As ReviewReportExportResult = Nothing, Optional consistencyPlan As ReviewConsistencyRun = Nothing,
                            Optional templateResult As System.Object = Nothing, Optional criteriaResult As System.Object = Nothing)
            If requestId.Length = 0 Then Return
            Post(New With {.command = "ack", .request = requestId, .ok = success, .error = message, .revision = _project.Revision,
                .folder = selectedFolder, .saveResult = saveResult, .persistence = PersistenceState(),
                .reportResult = reportResult, .consistencyPlan = consistencyPlan, .templateResult = templateResult, .criteriaResult = criteriaResult})
        End Sub

        Private Sub ShowMessage(message As System.String)
            If Me.InvokeRequired Then
                If Me.IsHandleCreated AndAlso Not Me.IsDisposed Then Me.BeginInvoke(New System.Action(Sub() ShowMessage(message)))
                Return
            End If
            If _closed OrElse Me.IsDisposed Then Return
            Using Global.SharedLibrary.SharedLibrary.SharedMethods.PushDialogOwner(Me)
                Global.SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox(message, Me.Text)
            End Using
        End Sub

        Private Async Function ConfirmCloseWithoutBrowserFlushAsync(reason As System.String) As System.Threading.Tasks.Task(Of System.Boolean)
            Return Await Global.SharedLibrary.SharedLibrary.SharedMethods.PostToHostUiAsync(Of System.Boolean)(
                Function() As System.Boolean
                    If _closed OrElse Me.IsDisposed Then Return True
                    Using Global.SharedLibrary.SharedLibrary.SharedMethods.PushDialogOwner(Me)
                        Dim answer As System.Int32 = Global.SharedLibrary.SharedLibrary.SharedMethods.ShowCustomYesNoBox(
                            reason & System.Environment.NewLine & "Unsaved edits may be lost. Close the review anyway?",
                            "Stay", "Close anyway", "Red Ink — Tabular Review")
                        Return answer = 2
                    End Using
                End Function)
        End Function

        Private Async Function FinishCloseOnOwnerThreadAsync(Optional confirmedFlush As System.Boolean = False) As System.Threading.Tasks.Task(Of System.Boolean)
            Return Await Global.SharedLibrary.SharedLibrary.SharedMethods.PostToHostUiAsync(Of System.Boolean)(
                Function() As System.Boolean
                    If _closed OrElse Me.IsDisposed Then Return True
                    ' Commit already persisted acknowledged edits. Flush only pending
                    ' layout changes here, on the Excel STA after a confirmed browser flush.
                    ' If the browser failed, preserve the previous acknowledged checkpoint.
                    If confirmedFlush Then
                        PersistPendingView()
                        If _project.View.AutoSave Then SaveWorkbookCore(False, True)
                    End If
                    _allowClose = True
                    Try
                        Me.Close()
                        Return Me.IsDisposed OrElse _closed
                    Finally
                        If Not _closed Then _allowClose = False
                    End Try
                End Function)
        End Function

        Public Async Function FlushAndCloseAsync() As System.Threading.Tasks.Task(Of System.Boolean)
            If _closed OrElse Me.IsDisposed Then Return True
            Dim confirmedFlush As System.Boolean = False
            If _busy AndAlso _operation IsNot Nothing Then
                _operation.Cancel()
                Dim completion As System.Threading.Tasks.Task = _operationDone.Task
                If (Await System.Threading.Tasks.Task.WhenAny(completion, System.Threading.Tasks.Task.Delay(30000))) IsNot completion Then
                    If Not Await ConfirmCloseWithoutBrowserFlushAsync("The active review operation did not stop after cancellation.") Then Return False
                    Return Await FinishCloseOnOwnerThreadAsync()
                End If
            End If
            If _ready Then
                If _flush IsNot Nothing Then Return False
                _flush = New System.Threading.Tasks.TaskCompletionSource(Of System.Boolean)(System.Threading.Tasks.TaskCreationOptions.RunContinuationsAsynchronously)
                _flushId = System.Guid.NewGuid().ToString("N")
                Post(New With {.command = "flush", .request = _flushId})
                Dim completed As System.Threading.Tasks.Task = Await System.Threading.Tasks.Task.WhenAny(_flush.Task, System.Threading.Tasks.Task.Delay(60000))
                Dim success As System.Boolean = completed Is _flush.Task AndAlso _flush.Task.Result
                confirmedFlush = success
                _flush = Nothing
                _flushId = ""
                If Not success Then
                    If Not Await ConfirmCloseWithoutBrowserFlushAsync("The review browser did not confirm that all pending edits were saved.") Then Return False
                End If
            End If
            Return Await FinishCloseOnOwnerThreadAsync(confirmedFlush)
        End Function

        Protected Overrides Async Sub OnFormClosing(e As System.Windows.Forms.FormClosingEventArgs)
            If _allowClose OrElse _hostLifetime.IsCancellationRequested OrElse e.CloseReason = System.Windows.Forms.CloseReason.WindowsShutDown Then
                MyBase.OnFormClosing(e)
                Return
            End If
            e.Cancel = True
            MyBase.OnFormClosing(e)
            If _closing Then Return
            _closing = True
            Try
                Await FlushAndCloseAsync()
            Catch ex As System.Exception
                ShowMessage("The review could not finish saving its edits: " & ex.Message)
            Finally
                _closing = False
            End Try
        End Sub

        Protected Overrides Sub OnFormClosed(e As System.Windows.Forms.FormClosedEventArgs)
            _closed = True
            _lifetime.Cancel()
            If _operation IsNot Nothing Then _operation.Cancel()
            If _sourceRequest IsNot Nothing Then _sourceRequest.Cancel()
            If _flush IsNot Nothing Then _flush.TrySetResult(False)
            _web.Dispose()
            Try
                If System.IO.Directory.Exists(_assetRoot) Then System.IO.Directory.Delete(_assetRoot, True)
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("[TabularReview:cache_cleanup] " & ex.Message)
            End Try
            MyBase.OnFormClosed(e)
        End Sub
    End Class
End Namespace
