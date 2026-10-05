' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' Shared modeless offline editor; native file privileges are limited to user-selected paths.

' =============================================================================
' File: MarkdownEditorForm.vb
' Purpose:
'   Shared modeless WebView2 Markdown editor with native files, persistent profiles,
'   floating mode and recovery.
'
' Architecture / Function:
'   Mediates selected file/folder access, conflict-checked atomic saves and acknowledged
'   recovery; supports saved Outlook startup placement and repeat-click rescue.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary
    Public NotInheritable Class MarkdownEditorForm
        Inherits System.Windows.Forms.Form

        Private Shared ReadOnly Instances As New System.Collections.Generic.Dictionary(Of System.String, MarkdownEditorForm)(System.StringComparer.OrdinalIgnoreCase)
        Private Const Origin As System.String = "https://redink-markdown.invalid/"
        Private ReadOnly _host As System.String
        Private ReadOnly _registryPath As System.String
        Private ReadOnly _web As New Microsoft.Web.WebView2.WinForms.WebView2()
        Private ReadOnly _bar As New System.Windows.Forms.FlowLayoutPanel()
        Private ReadOnly _compactButton As New System.Windows.Forms.Button()
        Private ReadOnly _paths As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.Ordinal)
        Private ReadOnly _hashes As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.Ordinal)
        Private ReadOnly _pending As New System.Collections.Generic.Dictionary(Of System.String, Newtonsoft.Json.Linq.JObject)(System.StringComparer.Ordinal)
        Private ReadOnly _queuedPaths As New System.Collections.Generic.List(Of System.String)()
        Private _ready As System.Boolean
        Private _compact As System.Boolean
        Private _resizing As System.Boolean
        Private _normalBounds As System.Drawing.Rectangle
        Private _flushCompletion As System.Threading.Tasks.TaskCompletionSource(Of System.Boolean)
        Private _closing As System.Boolean
        Private _allowClose As System.Boolean
        Private _startupCompact As System.Boolean
        Private _startupCompactLocation As System.Nullable(Of System.Drawing.Point)
        Private _hostClosing As System.Boolean
        Private _fileBusy As System.Boolean
        Private ReadOnly _floating As New System.Windows.Forms.PictureBox()
        Private ReadOnly _recoveryTimer As New System.Windows.Forms.Timer() With {.Interval = 200}
        Private ReadOnly _recovery As New System.Collections.Generic.Dictionary(Of System.String, Newtonsoft.Json.Linq.JObject)(System.StringComparer.Ordinal)
        Private _recoveryLoaded As System.Boolean
        Private _recoveryFailureReported As System.Boolean
        Private _dragStart As System.Drawing.Point
        Private _dragLocation As System.Drawing.Point
        Private _dragged As System.Boolean
        Private _lease As System.Threading.Mutex
        Private _ownsLease As System.Boolean

        Public Shared Sub ShowEditor(host As System.String)
            SharedMethods.RequireInteractiveExecution("markdown_editor")
            Dim existing As MarkdownEditorForm = Nothing
            If Instances.TryGetValue(host, existing) AndAlso Not existing.IsDisposed Then
                existing.SetCompact(False)
                existing.WindowState = System.Windows.Forms.FormWindowState.Normal
                Dim area As System.Drawing.Rectangle = System.Windows.Forms.Screen.FromPoint(System.Windows.Forms.Cursor.Position).WorkingArea
                existing.MinimumSize = New System.Drawing.Size(System.Math.Min(560, area.Width), System.Math.Min(360, area.Height))
                Dim width As System.Int32 = System.Math.Min(existing.Width, area.Width)
                Dim height As System.Int32 = System.Math.Min(existing.Height, area.Height)
                existing.Bounds = New System.Drawing.Rectangle(area.Left + (area.Width - width) \ 2, area.Top + (area.Height - height) \ 2, width, height)
                existing.Show()
                Dim keepOnTop As System.Boolean = existing.TopMost
                existing.TopMost = True
                existing.BringToFront()
                existing.Activate()
                existing.TopMost = keepOnTop
                Return
            End If
            Try
                Dim editor As New MarkdownEditorForm(host)
                Instances(host) = editor
                editor.Show() ' Modeless and independent: no Office HWND is made a modal owner.
            Catch ex As System.Exception
                SharedMethods.ShowCustomMessageBox("Markdown Editor could not start: " & ex.Message, SharedMethods.AN)
            End Try
        End Sub

        Public Shared Sub RestoreIfPreviouslyOpen(host As System.String)
            If host <> "Word" AndAlso host <> "Outlook" Then Throw New System.ArgumentException("Unknown editor host.", NameOf(host))
            Try
                Using key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\LawDigital\Red Ink\MarkdownEditor\" & host)
                    If key Is Nothing OrElse CInt(key.GetValue("Reopen", 0)) = 0 Then Return
                End Using
                ' Startup must not enlarge an already-open compact editor.
                Dim editor As MarkdownEditorForm = Nothing
                If Instances.TryGetValue(host, editor) AndAlso Not editor.IsDisposed Then Return
                ShowEditor(host)
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("[MarkdownEditor] Startup restore failed: " & ex.ToString())
            End Try
        End Sub

        Private Sub New(host As System.String)
            If host <> "Word" AndAlso host <> "Outlook" Then Throw New System.ArgumentException("Unknown editor host.", NameOf(host))
            _host = host
            _registryPath = "Software\LawDigital\Red Ink\MarkdownEditor\" & host
            _lease = New System.Threading.Mutex(False, "Local\RedInk.MarkdownEditor." & host & "." & System.Security.Principal.WindowsIdentity.GetCurrent().User.Value)
            Try
                _ownsLease = _lease.WaitOne(0)
            Catch ex As System.Threading.AbandonedMutexException
                _ownsLease = True
            End Try
            If Not _ownsLease Then
                _lease.Dispose()
                Throw New System.InvalidOperationException("The Markdown Editor is already open in another " & host & " process.")
            End If
            Try
                Text = "Red Ink - Markdown Editor"
                AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
                MinimumSize = New System.Drawing.Size(560, 360)
                Size = New System.Drawing.Size(1000, 720)
                StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
                ShowInTaskbar = True
                TopMost = True
                Icon = SharedMethods.CreateMarkdownEditorIcon()
                _bar.Dock = System.Windows.Forms.DockStyle.Top
                _bar.AutoSize = True
                _bar.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
                _bar.WrapContents = False
                _bar.Padding = New System.Windows.Forms.Padding(3)
                _compactButton.Text = "Compact"
                _compactButton.AutoSize = True
                AddHandler _compactButton.Click, Sub() SetCompact(Not _compact)
                Dim openButton As New System.Windows.Forms.Button() With {.Text = "Open file", .AutoSize = True}
                AddHandler openButton.Click, Async Sub()
                                                 If _ready AndAlso Not _fileBusy Then Await OpenFilesAsync()
                                             End Sub
                Dim top As New System.Windows.Forms.CheckBox() With {.Text = "Always on top", .Checked = True, .AutoSize = True}
                AddHandler top.CheckedChanged, Sub() TopMost = top.Checked
                _bar.Controls.Add(_compactButton)
                _bar.Controls.Add(openButton)
                _bar.Controls.Add(top)
                _web.Dock = System.Windows.Forms.DockStyle.Fill
                _web.AllowExternalDrop = True
                Controls.Add(_web)
                Controls.Add(_bar)
                _floating.Dock = System.Windows.Forms.DockStyle.Fill
                _floating.Image = SharedMethods.CreateMarkdownDocumentBitmap(64)
                _floating.BackColor = System.Drawing.Color.Magenta
                _floating.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
                _floating.Visible = False
                _floating.Cursor = System.Windows.Forms.Cursors.Hand
                _floating.AccessibleName = "Red Ink Markdown Editor — click to restore, drag to move"
                _floating.AllowDrop = True
                Controls.Add(_floating)
                AddHandler _floating.MouseDown, Sub(sender, args)
                                                    If args.Button <> System.Windows.Forms.MouseButtons.Left Then Return
                                                    _dragStart = System.Windows.Forms.Cursor.Position
                                                    _dragLocation = Location
                                                    _dragged = False
                                                    _floating.Capture = True
                                                End Sub
                AddHandler _floating.MouseMove, Sub(sender, args)
                                                    If Not _floating.Capture OrElse args.Button <> System.Windows.Forms.MouseButtons.Left Then Return
                                                    Dim delta = System.Windows.Forms.Cursor.Position - New System.Drawing.Size(_dragStart)
                                                    If System.Math.Abs(delta.X) + System.Math.Abs(delta.Y) > 4 Then _dragged = True
                                                    If _dragged Then Location = New System.Drawing.Point(_dragLocation.X + delta.X, _dragLocation.Y + delta.Y)
                                                End Sub
                AddHandler _floating.MouseUp, Sub(sender, args)
                                                  If args.Button <> System.Windows.Forms.MouseButtons.Left Then Return
                                                  _floating.Capture = False
                                                  If Not _dragged Then SetCompact(False)
                                              End Sub
                AddHandler _floating.DragEnter, AddressOf FileDragEnter
                AddHandler _floating.DragDrop, AddressOf FileDragDrop
                AddHandler _recoveryTimer.Tick, Sub()
                                                   _recoveryTimer.Stop()
                                                   Try
                                                       PersistRecovery()
                                                       _recoveryFailureReported = False
                                                   Catch ex As System.Exception
                                                       System.Diagnostics.Debug.WriteLine("[MarkdownEditor] Recovery write failed: " & ex.ToString())
                                                       If Not _recoveryFailureReported Then
                                                           _recoveryFailureReported = True
                                                           Using SharedMethods.PushDialogOwner(Me)
                                                               SharedMethods.ShowCustomMessageBox("Recovery storage failed. Save or export your edited notes before closing the host. " & ex.Message, Text)
                                                           End Using
                                                       End If
                                                   End Try
                                               End Sub
                AllowDrop = True
                AddHandler DragEnter, AddressOf FileDragEnter
                AddHandler DragDrop, AddressOf FileDragDrop
                AddHandler _bar.DragEnter, AddressOf FileDragEnter
                AddHandler _bar.DragDrop, AddressOf FileDragDrop
                _bar.AllowDrop = True
                AddHandler Shown, AddressOf InitializeAsync
                AddHandler FormClosing, AddressOf ClosingAsync
                AddHandler FormClosed, Sub()
                                           _recoveryTimer.Dispose()
                                           If _floating.Image IsNot Nothing Then _floating.Image.Dispose()
                                           Instances.Remove(_host)
                                           If _ownsLease Then _lease.ReleaseMutex()
                                           _ownsLease = False
                                           _lease.Dispose()
                                           If Icon IsNot Nothing Then Icon.Dispose()
                                       End Sub
                AddHandler Deactivate, Sub() SharedMethods.PromoteForeignForegroundDialog(Me)
                SharedMethods.AttachForeignForegroundWatchdog(Me)
                LoadState()
                top.Checked = TopMost
            Catch ex As System.Exception
                If _ownsLease Then _lease.ReleaseMutex()
                _ownsLease = False
                _lease.Dispose()
                If Icon IsNot Nothing Then Icon.Dispose()
                Throw
            End Try
        End Sub

        Private Function EditorDialogOwner(dialogType As System.String) As System.Windows.Forms.IWin32Window
            OfficeWindowWatchdog.InspectDialogOwner(Me, dialogType, "MarkdownEditor")
            Return SharedMethods.IfOwnerOnCurrentThread(Me)
        End Function

        Private Sub LoadState()
            Try
                Using key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(_registryPath)
                    If key Is Nothing Then Return
                    _folderRoot = CStr(key.GetValue("Folder", ""))
                    Dim paths = Newtonsoft.Json.Linq.JObject.Parse(CStr(key.GetValue("Files", "{}")))
                    For Each p In paths.Properties()
                        Dim record = TryCast(p.Value, Newtonsoft.Json.Linq.JObject)
                        If record Is Nothing Then Continue For
                        _paths(p.Name) = CStr(record("path"))
                        _hashes(p.Name) = CStr(record("hash"))
                    Next
                    Dim rect As New System.Drawing.Rectangle(CInt(key.GetValue("X", 100)), CInt(key.GetValue("Y", 100)), CInt(key.GetValue("Width", 1000)), CInt(key.GetValue("Height", 720)))
                    rect.Width = System.Math.Max(MinimumSize.Width, rect.Width)
                    rect.Height = System.Math.Max(MinimumSize.Height, rect.Height)
                    Dim area = System.Windows.Forms.Screen.FromRectangle(rect).WorkingArea
                    MinimumSize = New System.Drawing.Size(System.Math.Min(560, area.Width), System.Math.Min(360, area.Height))
                    rect.Width = System.Math.Min(rect.Width, area.Width)
                    rect.Height = System.Math.Min(rect.Height, area.Height)
                    rect.X = System.Math.Max(area.Left, System.Math.Min(rect.X, area.Right - rect.Width))
                    rect.Y = System.Math.Max(area.Top, System.Math.Min(rect.Y, area.Bottom - rect.Height))
                    Bounds = rect
                    StartPosition = System.Windows.Forms.FormStartPosition.Manual
                    _startupCompact = CInt(key.GetValue("Compact", 0)) <> 0
                    If _startupCompact Then
                        _startupCompactLocation = New System.Drawing.Point(CInt(key.GetValue("CompactX", CInt(key.GetValue("X", rect.X)))), CInt(key.GetValue("CompactY", CInt(key.GetValue("Y", rect.Y)))))
                    ElseIf CInt(key.GetValue("Maximized", 0)) <> 0 Then
                        WindowState = System.Windows.Forms.FormWindowState.Maximized
                    End If
                    TopMost = CInt(key.GetValue("TopMost", 1)) <> 0
                End Using
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("[MarkdownEditor] State load failed: " & ex.Message)
                SharedMethods.ShowCustomMessageBox("Saved window/file state could not be loaded. Editor notes are kept in the browser profile. " & ex.Message, Text)
            End Try
        End Sub

        Private Sub SaveState()
            Dim records As New Newtonsoft.Json.Linq.JObject()
            For Each p In _paths
                records(p.Key) = New Newtonsoft.Json.Linq.JObject(New Newtonsoft.Json.Linq.JProperty("path", p.Value), New Newtonsoft.Json.Linq.JProperty("hash", If(_hashes.ContainsKey(p.Key), _hashes(p.Key), "")))
            Next
            Dim rect = If(_compact, _normalBounds, If(WindowState = System.Windows.Forms.FormWindowState.Normal, Bounds, RestoreBounds))
            Using key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(_registryPath)
                key.SetValue("Folder", If(_folderRoot, ""))
                key.SetValue("Files", records.ToString(Newtonsoft.Json.Formatting.None))
                key.SetValue("X", rect.X)
                key.SetValue("Y", rect.Y)
                key.SetValue("Width", rect.Width)
                key.SetValue("Height", rect.Height)
                key.SetValue("Compact", If(_compact, 1, 0))
                key.SetValue("CompactX", Location.X)
                key.SetValue("CompactY", Location.Y)
                key.SetValue("Maximized", If(Not _compact AndAlso WindowState = System.Windows.Forms.FormWindowState.Maximized, 1, 0))
                key.SetValue("Reopen", 1)
                key.SetValue("TopMost", If(TopMost, 1, 0))
            End Using
        End Sub

        Private Async Sub InitializeAsync(sender As System.Object, e As System.EventArgs)
            Try
                Dim root = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData), "RedInk", "MarkdownEditor", _host)
                Dim assets = System.IO.Path.Combine(root, "Assets")
                System.IO.Directory.CreateDirectory(assets)
                Using stream = GetType(MarkdownEditorForm).Assembly.GetManifestResourceStream("RedInk.MarkdownEditor.html")
                    If stream Is Nothing Then Throw New System.IO.FileNotFoundException("Embedded Markdown Editor is missing.")
                    Using reader As New System.IO.StreamReader(stream, System.Text.Encoding.UTF8)
                        System.IO.File.WriteAllText(System.IO.Path.Combine(assets, "index.html"), reader.ReadToEnd(), New System.Text.UTF8Encoding(False))
                    End Using
                End Using
                Dim environment = Await Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(Nothing, System.IO.Path.Combine(root, "Profile"))
                Await _web.EnsureCoreWebView2Async(environment)
                Dim core = _web.CoreWebView2
                core.Settings.AreDevToolsEnabled = False
                core.Settings.AreHostObjectsAllowed = False
                core.Settings.IsWebMessageEnabled = True
                core.SetVirtualHostNameToFolderMapping("redink-markdown.invalid", assets, Microsoft.Web.WebView2.Core.CoreWebView2HostResourceAccessKind.DenyCors)
                core.AddWebResourceRequestedFilter("*", Microsoft.Web.WebView2.Core.CoreWebView2WebResourceContext.All)
                AddHandler core.WebResourceRequested,
                    Sub(s, args)
                        Dim uri As System.Uri = Nothing
                        If Not System.Uri.TryCreate(args.Request.Uri, System.UriKind.Absolute, uri) Then Return
                        If uri.Scheme = "data" OrElse uri.Scheme = "blob" Then Return
                        If Not args.Request.Uri.StartsWith(Origin, System.StringComparison.Ordinal) Then
                            args.Response = core.Environment.CreateWebResourceResponse(Nothing, 403, "Offline editor", "Content-Type: text/plain")
                        End If
                    End Sub
                AddHandler core.NavigationStarting,
                    Sub(s, args)
                        If Not args.Uri.StartsWith(Origin, System.StringComparison.Ordinal) Then args.Cancel = True
                    End Sub
                AddHandler core.NewWindowRequested,
                    Sub(s, args)
                        args.Handled = True
                        If Not args.IsUserInitiated Then Return
                        ' Explicit link clicks retain navigation through a user confirmation; background requests remain blocked.
                        Using SharedMethods.PushDialogOwner(Me)
                            Dim target As System.Uri = Nothing
                            If System.Uri.TryCreate(args.Uri, System.UriKind.Absolute, target) AndAlso (target.Scheme = "http" OrElse target.Scheme = "https" OrElse target.Scheme = "mailto") Then
                                If SharedMethods.ShowCustomYesNoBox("Open this link in your default application?" & System.Environment.NewLine & args.Uri, "Open", "Cancel", Text) = 1 Then
                                    System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(args.Uri) With {.UseShellExecute = True})
                                End If
                            Else
                                SharedMethods.ShowCustomMessageBox("Unsupported link (not opened): " & args.Uri, Text)
                            End If
                        End Using
                    End Sub
                AddHandler core.PermissionRequested, Sub(s, args) args.State = Microsoft.Web.WebView2.Core.CoreWebView2PermissionState.Deny
                AddHandler core.WebMessageReceived, AddressOf MessageAsync
                AddHandler core.DownloadStarting, AddressOf DownloadStarting
                AddHandler core.ProcessFailed,
                    Sub(s, args)
                        _ready = False
                        SharedMethods.LogWebView2ProcessFailed("MarkdownEditor", args.ProcessFailedKind.ToString(), args.ExitCode.ToString())
                        Using SharedMethods.PushDialogOwner(Me)
                            SharedMethods.ShowCustomMessageBox("The editor browser stopped. Close and reopen the editor to recover autosaved notes. The latest unsaved keystrokes may need re-entry.", Text)
                        End Using
                    End Sub
                core.Navigate(Origin & "index.html")
                If _startupCompact Then
                    SetCompact(True)
                    If _startupCompactLocation.HasValue Then Location = _startupCompactLocation.Value
                    ClampToWorkingArea()
                End If
            Catch ex As System.Exception
                Using SharedMethods.PushDialogOwner(Me)
                    SharedMethods.ShowCustomMessageBox("Markdown Editor could not load. Install/repair the Microsoft Edge WebView2 Runtime. " & ex.Message, Text)
                End Using
            End Try
        End Sub

        Private Async Sub MessageAsync(sender As System.Object, e As Microsoft.Web.WebView2.Core.CoreWebView2WebMessageReceivedEventArgs)
            If e.Source <> Origin & "index.html" Then Return
            Dim request As System.String = ""
            Try
                Dim message = Newtonsoft.Json.Linq.JObject.Parse(e.WebMessageAsJson)
                request = CStr(message("request"))
                Select Case CStr(message("command"))
                    Case "recovery"
                        Dim id = CStr(message("id"))
                        If System.String.IsNullOrWhiteSpace(id) OrElse message("text") Is Nothing OrElse message("text").Type <> Newtonsoft.Json.Linq.JTokenType.String Then Throw New System.IO.InvalidDataException("Invalid recovery text.")
                        _recovery(id) = New Newtonsoft.Json.Linq.JObject(New Newtonsoft.Json.Linq.JProperty("id", id), New Newtonsoft.Json.Linq.JProperty("name", CStr(message("name"))), New Newtonsoft.Json.Linq.JProperty("text", CStr(message("text"))), New Newtonsoft.Json.Linq.JProperty("updated", System.DateTime.UtcNow.ToString("o")))
                        If Not _recoveryTimer.Enabled Then _recoveryTimer.Start()
                    Case "recoveryAccepted", "recoveryClear"
                        Dim id = CStr(message("id"))
                        Dim snapshot As Newtonsoft.Json.Linq.JObject = Nothing
                        If _recovery.TryGetValue(id, snapshot) Then
                            Dim matches = If(CStr(message("command")) = "recoveryAccepted", CStr(snapshot("updated")) = CStr(message("updated")), CStr(snapshot("text")) = CStr(message("text")))
                            If matches Then
                                _recovery.Remove(id)
                                PersistRecovery()
                            End If
                        End If
                    Case "flushed"
                        If _flushCompletion IsNot Nothing Then _flushCompletion.TrySetResult(message.Value(Of System.Boolean)("ok"))
                    Case "ready"
                        _ready = True
                        Dim recoveryPath = RecoveryPathForHost(_host)
                        If System.IO.File.Exists(recoveryPath) Then
                            Dim records = Newtonsoft.Json.Linq.JArray.Parse(System.IO.File.ReadAllText(recoveryPath, System.Text.Encoding.UTF8))
                            For Each record In records
                                Dim snapshot = TryCast(record, Newtonsoft.Json.Linq.JObject)
                                If snapshot Is Nothing OrElse snapshot("text") Is Nothing OrElse snapshot("text").Type <> Newtonsoft.Json.Linq.JTokenType.String OrElse System.String.IsNullOrWhiteSpace(CStr(snapshot("id"))) Then Throw New System.IO.InvalidDataException("Invalid recovery snapshot; original recovery file retained.")
                                _recovery(CStr(snapshot("id"))) = snapshot
                            Next
                            For Each snapshot In _recovery.Values
                                _web.CoreWebView2.PostWebMessageAsJson(New Newtonsoft.Json.Linq.JObject(New Newtonsoft.Json.Linq.JProperty("command", "recover"), New Newtonsoft.Json.Linq.JProperty("snapshot", snapshot)).ToString())
                            Next
                        End If
                        _recoveryLoaded = True
                        _web.CoreWebView2.PostWebMessageAsJson(New Newtonsoft.Json.Linq.JObject(New Newtonsoft.Json.Linq.JProperty("command", "paths"), New Newtonsoft.Json.Linq.JProperty("paths", Newtonsoft.Json.Linq.JObject.FromObject(_paths))).ToString())
                        For Each path In _queuedPaths.ToArray()
                            Await ImportPathAsync(path)
                        Next
                        _queuedPaths.Clear()
                        If Not System.String.IsNullOrWhiteSpace(_folderRoot) Then
                            Dim handle = FolderHandle(_folderRoot)
                            _web.CoreWebView2.PostWebMessageAsJson(New Newtonsoft.Json.Linq.JObject(New Newtonsoft.Json.Linq.JProperty("command", "folderRestore"), New Newtonsoft.Json.Linq.JProperty("handle", handle)).ToString())
                        End If
                    Case "pickFolder", "disconnectFolder", "folderList", "folderChild", "folderRead", "folderAcceptRead", "folderWrite", "folderRemove", "folderRename"
                        FolderMessage(message, request)
                    Case "bind"
                        Dim token = CStr(message("token"))
                        Dim id = CStr(message("id"))
                        Dim file As Newtonsoft.Json.Linq.JObject = Nothing
                        If System.String.IsNullOrWhiteSpace(id) OrElse Not _pending.TryGetValue(token, file) Then Throw New System.IO.InvalidDataException("Unknown file selection.")
                        _paths(id) = CStr(file("path"))
                        If Not _hashes.ContainsKey(id) Then _hashes(id) = CStr(file("hash"))
                        _pending.Remove(token)
                        SaveState()
                    Case "open"
                        If _fileBusy Then Throw New System.InvalidOperationException("Another file action is active.")
                        Await OpenFilesAsync()
                        Reply(request)
                    Case "openDropped"
                        For Each item In e.AdditionalObjects
                            Dim file = TryCast(item, Microsoft.Web.WebView2.Core.CoreWebView2File)
                            If file IsNot Nothing AndAlso IsTextPath(file.Path) Then Await ImportPathAsync(file.Path)
                        Next
                    Case "save"
                        If _fileBusy Then Throw New System.InvalidOperationException("Another file action is active.")
                        _fileBusy = True
                        Try
                            Dim id = CStr(message("id"))
                            If System.String.IsNullOrWhiteSpace(id) Then Throw New System.IO.InvalidDataException("Missing note ID.")
                            Dim text = message("text")
                            If text Is Nothing OrElse text.Type <> Newtonsoft.Json.Linq.JTokenType.String Then Throw New System.IO.InvalidDataException("Missing note content.")
                            Dim path As System.String = Nothing
                            Dim saveCopy = message.Value(Of System.Boolean)("copy")
                            If Not saveCopy Then _paths.TryGetValue(id, path)
                            If message.Value(Of System.Boolean)("autosave") AndAlso (saveCopy OrElse message.Value(Of System.Boolean)("saveAs") OrElse System.String.IsNullOrWhiteSpace(path) OrElse Not System.IO.File.Exists(path)) Then Throw New System.IO.IOException("Autosave requires an existing linked file. Use Save as to select a destination.")
                            If message.Value(Of System.Boolean)("saveAs") OrElse System.String.IsNullOrWhiteSpace(path) Then
                                Using dialog As New System.Windows.Forms.SaveFileDialog() With {.Filter = "Markdown / text|*.md;*.markdown;*.txt|All files|*.*", .DefaultExt = "md", .FileName = System.IO.Path.GetFileName(CStr(message("name"))), .OverwritePrompt = True}
                                    If dialog.ShowDialog(EditorDialogOwner(dialog.GetType().FullName)) <> System.Windows.Forms.DialogResult.OK Then
                                        Reply(request, New Newtonsoft.Json.Linq.JProperty("cancelled", True))
                                        Return
                                    End If
                                    path = dialog.FileName
                                End Using
                            ElseIf System.IO.File.Exists(path) AndAlso (Not _hashes.ContainsKey(id) OrElse Fingerprint(path) <> _hashes(id)) Then
                                Throw New System.IO.IOException("The file changed outside the editor. Use Save as or explicitly Reload from disk.")
                            End If
                            AtomicWrite(path, CStr(text))
                            If Not saveCopy Then
                                _paths(id) = path
                                _hashes(id) = Fingerprint(path)
                            End If
                            SaveState()
                            Reply(request, New Newtonsoft.Json.Linq.JProperty("path", path))
                        Finally
                            _fileBusy = False
                        End Try
                    Case "reload"
                        Dim id = CStr(message("id"))
                        Dim path As System.String = Nothing
                        If Not _paths.TryGetValue(id, path) Then Throw New System.IO.FileNotFoundException("This note has no file path.")
                        Dim snapshot = ReadTextSnapshot(path)
                        Dim token = System.Guid.NewGuid().ToString("N")
                        snapshot("path") = path
                        snapshot("id") = id
                        _pending(token) = snapshot
                        Reply(request, New Newtonsoft.Json.Linq.JProperty("text", snapshot("text")), New Newtonsoft.Json.Linq.JProperty("token", token))
                    Case "reloadCommitted"
                        Dim token = CStr(message("token"))
                        Dim snapshot As Newtonsoft.Json.Linq.JObject = Nothing
                        If Not _pending.TryGetValue(token, snapshot) OrElse CStr(snapshot("id")) <> CStr(message("id")) Then Throw New System.IO.InvalidDataException("Unknown reload snapshot.")
                        Dim path As System.String = Nothing
                        If Not _paths.TryGetValue(CStr(message("id")), path) OrElse path <> CStr(snapshot("path")) Then Throw New System.IO.InvalidDataException("Reload file binding changed.")
                        _hashes(CStr(message("id"))) = CStr(snapshot("hash"))
                        _pending.Remove(token)
                        SaveState()
                    Case "discardSnapshot"
                        _pending.Remove(CStr(message("token")))
                End Select
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("[MarkdownEditor] File action failed: " & ex.Message)
                If Not System.String.IsNullOrWhiteSpace(request) Then
                    Reply(request, New Newtonsoft.Json.Linq.JProperty("error", ex.Message))
                Else
                    Using SharedMethods.PushDialogOwner(Me)
                        SharedMethods.ShowCustomMessageBox("File action failed: " & ex.Message, Text)
                    End Using
                End If
            End Try
        End Sub

        Private _folderRoot As System.String
        Private ReadOnly _handles As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.Ordinal)
        Private ReadOnly _handleTokens As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)
        Private ReadOnly _folderHashes As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)
        Private ReadOnly _folderReadHashes As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)

        Private Function FolderHandle(path As System.String) As Newtonsoft.Json.Linq.JObject
            ValidateFolderPath(path)
            Dim token As System.String = Nothing
            If Not _handleTokens.TryGetValue(path, token) Then
                token = System.Guid.NewGuid().ToString("N")
                _handleTokens(path) = token
                _handles(token) = path
            End If
            Return New Newtonsoft.Json.Linq.JObject(New Newtonsoft.Json.Linq.JProperty("token", token), New Newtonsoft.Json.Linq.JProperty("name", System.IO.Path.GetFileName(path.TrimEnd(System.IO.Path.DirectorySeparatorChar))), New Newtonsoft.Json.Linq.JProperty("kind", If(System.IO.Directory.Exists(path), "directory", "file")))
        End Function

        Private Sub ValidateFolderPath(path As System.String)
            If System.String.IsNullOrWhiteSpace(_folderRoot) Then Throw New System.UnauthorizedAccessException("No selected folder.")
            Dim root = System.IO.Path.GetFullPath(_folderRoot).TrimEnd(System.IO.Path.DirectorySeparatorChar)
            Dim full = System.IO.Path.GetFullPath(path)
            If Not System.String.Equals(full.TrimEnd(System.IO.Path.DirectorySeparatorChar), root, System.StringComparison.OrdinalIgnoreCase) AndAlso
               Not full.StartsWith(root & System.IO.Path.DirectorySeparatorChar, System.StringComparison.OrdinalIgnoreCase) Then Throw New System.UnauthorizedAccessException("Path is outside the selected folder.")
            Dim current = full
            While Not System.String.IsNullOrEmpty(current)
                If (System.IO.File.Exists(current) OrElse System.IO.Directory.Exists(current)) AndAlso
                   (System.IO.File.GetAttributes(current) And System.IO.FileAttributes.ReparsePoint) <> 0 Then Throw New System.UnauthorizedAccessException("Linked files/folders are not followed.")
                Dim parent = System.IO.Path.GetDirectoryName(current)
                If parent = current Then Exit While
                current = parent
            End While
        End Sub

        Private Function HandlePath(token As System.String) As System.String
            Dim path As System.String = Nothing
            If token Is Nothing OrElse Not _handles.TryGetValue(token, path) Then Throw New System.UnauthorizedAccessException("Unknown folder handle.")
            ValidateFolderPath(path)
            Return path
        End Function

        Private Function ChildPath(parent As System.String, name As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(name) OrElse name = "." OrElse name = ".." OrElse name.IndexOfAny(System.IO.Path.GetInvalidFileNameChars()) >= 0 OrElse name <> System.IO.Path.GetFileName(name) Then Throw New System.IO.InvalidDataException("Invalid child name.")
            If name.EndsWith(".", System.StringComparison.Ordinal) OrElse name.EndsWith(" ", System.StringComparison.Ordinal) OrElse
               System.Text.RegularExpressions.Regex.IsMatch(name, "^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)", System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then Throw New System.IO.InvalidDataException("Reserved Windows filename.")
            Dim path = System.IO.Path.Combine(parent, name)
            ValidateFolderPath(path)
            Return path
        End Function

        Private Sub FolderMessage(message As Newtonsoft.Json.Linq.JObject, request As System.String)
            Dim command = CStr(message("command"))
            If command = "pickFolder" Then
                Using dialog As New System.Windows.Forms.FolderBrowserDialog() With {.Description = "Choose a folder for the Markdown Editor. Files in this folder can be edited.", .ShowNewFolderButton = True}
                    If dialog.ShowDialog(EditorDialogOwner(dialog.GetType().FullName)) <> System.Windows.Forms.DialogResult.OK Then
                        Reply(request, New Newtonsoft.Json.Linq.JProperty("cancelled", True))
                        Return
                    End If
                    _folderRoot = System.IO.Path.GetFullPath(dialog.SelectedPath)
                    _handles.Clear()
                    _handleTokens.Clear()
                    _folderHashes.Clear()
                    _folderReadHashes.Clear()
                    Dim handle = FolderHandle(_folderRoot)
                    SaveState()
                    Reply(request, New Newtonsoft.Json.Linq.JProperty("handle", handle))
                    Return
                End Using
            End If
            If command = "disconnectFolder" Then
                _folderRoot = Nothing
                _handles.Clear()
                _handleTokens.Clear()
                _folderHashes.Clear()
                _folderReadHashes.Clear()
                SaveState()
                Reply(request)
                Return
            End If
            Dim path = HandlePath(CStr(message("token")))
            Select Case command
                Case "folderList"
                    Dim entries As New Newtonsoft.Json.Linq.JArray()
                    For Each child In System.IO.Directory.EnumerateFileSystemEntries(path)
                        ' A rejected reparse entry must not prevent the remaining authorized tree from loading.
                        If (System.IO.File.GetAttributes(child) And System.IO.FileAttributes.ReparsePoint) <> 0 Then Continue For
                        entries.Add(FolderHandle(child))
                    Next
                    Reply(request, New Newtonsoft.Json.Linq.JProperty("entries", entries))
                Case "folderChild"
                    Dim child = ChildPath(path, CStr(message("name")))
                    Dim directory = CStr(message("kind")) = "directory"
                    If message.Value(Of System.Boolean)("create") Then
                        If directory Then
                            System.IO.Directory.CreateDirectory(child)
                        ElseIf Not System.IO.File.Exists(child) Then
                            Using stream = New System.IO.FileStream(child, System.IO.FileMode.CreateNew, System.IO.FileAccess.Write)
                            End Using
                            _folderHashes(child) = Fingerprint(child)
                        End If
                    End If
                    If (directory AndAlso Not System.IO.Directory.Exists(child)) OrElse (Not directory AndAlso Not System.IO.File.Exists(child)) Then Throw New System.IO.FileNotFoundException("Child does not exist.")
                    Reply(request, New Newtonsoft.Json.Linq.JProperty("handle", FolderHandle(child)))
                Case "folderRead"
                    Dim bytes = System.IO.File.ReadAllBytes(path)
                    Using hash = System.Security.Cryptography.SHA256.Create()
                        _folderReadHashes(path) = System.Convert.ToBase64String(hash.ComputeHash(bytes))
                    End Using
                    Reply(request, New Newtonsoft.Json.Linq.JProperty("base64", System.Convert.ToBase64String(bytes)), New Newtonsoft.Json.Linq.JProperty("hash", _folderReadHashes(path)), New Newtonsoft.Json.Linq.JProperty("modified", CLng((System.IO.File.GetLastWriteTimeUtc(path) - New System.DateTime(1970, 1, 1, 0, 0, 0, System.DateTimeKind.Utc)).TotalMilliseconds)))
                Case "folderAcceptRead"
                    Dim hash As System.String = Nothing
                    If Not _folderReadHashes.TryGetValue(path, hash) OrElse hash <> CStr(message("hash")) Then Throw New System.IO.IOException("The read snapshot is no longer current; reload the file.")
                    _folderHashes(path) = hash
                    Reply(request)
                Case "folderRename"
                    RenameFolderFile(path, CStr(message("token")), CStr(message("name")))
                    Reply(request)
                Case "folderWrite"
                    If System.IO.File.Exists(path) AndAlso (Not _folderHashes.ContainsKey(path) OrElse Fingerprint(path) <> _folderHashes(path)) Then Throw New System.IO.IOException("The file changed outside the editor. Rescan/reload or export your editor version.")
                    Dim temporary = path & "." & System.Guid.NewGuid().ToString("N") & ".tmp"
                    Try
                        System.IO.File.WriteAllBytes(temporary, System.Convert.FromBase64String(CStr(message("base64"))))
                        If System.IO.File.Exists(path) Then
                            System.IO.File.Replace(temporary, path, Nothing)
                        Else
                            System.IO.File.Move(temporary, path)
                        End If
                        _folderHashes(path) = Fingerprint(path)
                    Finally
                        If System.IO.File.Exists(temporary) Then System.IO.File.Delete(temporary)
                    End Try
                    Reply(request)
                Case "folderRemove"
                    Dim child = ChildPath(path, CStr(message("name")))
                    If System.IO.Directory.Exists(child) Then
                        If Not message.Value(Of System.Boolean)("recursive") Then
                            Using entries = System.IO.Directory.EnumerateFileSystemEntries(child).GetEnumerator()
                                If entries.MoveNext() Then Throw New System.IO.IOException("Folder is not empty.")
                            End Using
                        End If
                        Microsoft.VisualBasic.FileIO.FileSystem.DeleteDirectory(child, Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs, Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin)
                    Else
                        Microsoft.VisualBasic.FileIO.FileSystem.DeleteFile(child, Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs, Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin)
                    End If
                    _folderHashes.Remove(child)
                    Reply(request)
            End Select
        End Sub

        Private Sub RenameFolderFile(path As System.String, token As System.String, name As System.String)
            If Not System.IO.File.Exists(path) Then Throw New System.IO.FileNotFoundException("Rename source does not exist.")
            If Not _folderHashes.ContainsKey(path) OrElse Fingerprint(path) <> _folderHashes(path) Then Throw New System.IO.IOException("The file changed outside the editor; reload before renaming.")
            Dim destination = ChildPath(System.IO.Path.GetDirectoryName(path), name)
            If System.String.Equals(path, destination, System.StringComparison.Ordinal) Then Return
            If System.String.Equals(path, destination, System.StringComparison.OrdinalIgnoreCase) Then
                Dim temporary = path & "." & System.Guid.NewGuid().ToString("N") & ".rename"
                System.IO.File.Move(path, temporary)
                Try
                    System.IO.File.Move(temporary, destination)
                Catch ex As System.Exception
                    Try
                        System.IO.File.Move(temporary, path)
                    Catch rollback As System.Exception
                        Throw New System.IO.IOException("Rename failed; original bytes remain at " & temporary & ". " & rollback.Message, ex)
                    End Try
                    Throw
                End Try
            Else
                System.IO.File.Move(path, destination) ' Never overwrite an existing destination.
            End If
            Dim hash = _folderHashes(path)
            _folderHashes.Remove(path)
            _folderReadHashes.Remove(path)
            _handleTokens.Remove(path)
            _handles(token) = destination
            _handleTokens(destination) = token
            _folderHashes(destination) = hash
        End Sub

        Private Sub Reply(request As System.String, ParamArray fields As Newtonsoft.Json.Linq.JProperty())
            Dim reply As New Newtonsoft.Json.Linq.JObject(New Newtonsoft.Json.Linq.JProperty("request", request))
            For Each field In fields
                reply.Add(field)
            Next
            _web.CoreWebView2.PostWebMessageAsJson(reply.ToString(Newtonsoft.Json.Formatting.None))
        End Sub

        Private Async Function OpenFilesAsync() As System.Threading.Tasks.Task
            _fileBusy = True
            Try
                Using dialog As New System.Windows.Forms.OpenFileDialog() With {.Filter = "Markdown / text|*.md;*.markdown;*.mdown;*.mkd;*.txt|All files|*.*", .Multiselect = True}
                    If dialog.ShowDialog(EditorDialogOwner(dialog.GetType().FullName)) <> System.Windows.Forms.DialogResult.OK Then Return
                    For Each path In dialog.FileNames
                        Await ImportPathAsync(path)
                    Next
                End Using
            Catch ex As System.Exception
                Using SharedMethods.PushDialogOwner(Me)
                    SharedMethods.ShowCustomMessageBox("File open failed: " & ex.Message, Text)
                End Using
            Finally
                _fileBusy = False
            End Try
        End Function

        Private Async Function ImportPathAsync(path As System.String) As System.Threading.Tasks.Task
            path = System.IO.Path.GetFullPath(path)
            For Each pending In _pending.Values
                If System.String.Equals(CStr(pending("path")), path, System.StringComparison.OrdinalIgnoreCase) Then Return
            Next
            If Not IsTextPath(path) Then Throw New System.IO.InvalidDataException("Drop Markdown or plain-text files only.")
            If Not _ready Then
                _queuedPaths.Add(path)
                Return
            End If
            Dim id As System.String = Nothing
            For Each p In _paths
                If System.String.Equals(p.Value, path, System.StringComparison.OrdinalIgnoreCase) Then
                    id = p.Key
                    Exit For
                End If
            Next
            Dim token = System.Guid.NewGuid().ToString("N")
            Dim snapshot = ReadTextSnapshot(path)
            Dim file As New Newtonsoft.Json.Linq.JObject(New Newtonsoft.Json.Linq.JProperty("id", id), New Newtonsoft.Json.Linq.JProperty("token", token), New Newtonsoft.Json.Linq.JProperty("path", path), New Newtonsoft.Json.Linq.JProperty("name", System.IO.Path.GetFileName(path)), New Newtonsoft.Json.Linq.JProperty("text", snapshot("text")), New Newtonsoft.Json.Linq.JProperty("hash", snapshot("hash")))
            _pending(token) = file
            _web.CoreWebView2.PostWebMessageAsJson(New Newtonsoft.Json.Linq.JObject(New Newtonsoft.Json.Linq.JProperty("command", "import"), New Newtonsoft.Json.Linq.JProperty("file", file)).ToString(Newtonsoft.Json.Formatting.None))
            Await System.Threading.Tasks.Task.CompletedTask
        End Function

        Private Shared Function IsTextPath(path As System.String) As System.Boolean
            Return (New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase) From {".md", ".markdown", ".mdown", ".mkd", ".txt"}).Contains(System.IO.Path.GetExtension(path))
        End Function

        Private Shared Function ReadTextSnapshot(path As System.String) As Newtonsoft.Json.Linq.JObject
            Dim bytes = System.IO.File.ReadAllBytes(path)
            Dim encoding As System.Text.Encoding = New System.Text.UTF8Encoding(False, True)
            Dim offset As System.Int32 = 0
            If bytes.Length >= 4 AndAlso bytes(0) = &HFF AndAlso bytes(1) = &HFE AndAlso bytes(2) = 0 AndAlso bytes(3) = 0 Then
                encoding = New System.Text.UTF32Encoding(False, False, True)
                offset = 4
            ElseIf bytes.Length >= 4 AndAlso bytes(0) = 0 AndAlso bytes(1) = 0 AndAlso bytes(2) = &HFE AndAlso bytes(3) = &HFF Then
                encoding = New System.Text.UTF32Encoding(True, False, True)
                offset = 4
            ElseIf bytes.Length >= 2 AndAlso bytes(0) = &HFF AndAlso bytes(1) = &HFE Then
                encoding = New System.Text.UnicodeEncoding(False, False, True)
                offset = 2
            ElseIf bytes.Length >= 2 AndAlso bytes(0) = &HFE AndAlso bytes(1) = &HFF Then
                encoding = New System.Text.UnicodeEncoding(True, False, True)
                offset = 2
            ElseIf bytes.Length >= 3 AndAlso bytes(0) = &HEF AndAlso bytes(1) = &HBB AndAlso bytes(2) = &HBF Then
                offset = 3
            End If
            Using hash = System.Security.Cryptography.SHA256.Create()
                Return New Newtonsoft.Json.Linq.JObject(New Newtonsoft.Json.Linq.JProperty("text", encoding.GetString(bytes, offset, bytes.Length - offset)), New Newtonsoft.Json.Linq.JProperty("hash", System.Convert.ToBase64String(hash.ComputeHash(bytes))))
            End Using
        End Function

        Private Shared Function Fingerprint(path As System.String) As System.String
            Using hash = System.Security.Cryptography.SHA256.Create(), stream = System.IO.File.OpenRead(path)
                Return System.Convert.ToBase64String(hash.ComputeHash(stream))
            End Using
        End Function

        Private Shared Sub AtomicWrite(path As System.String, text As System.String)
            Dim temporary = path & "." & System.Guid.NewGuid().ToString("N") & ".tmp"
            Try
                System.IO.File.WriteAllText(temporary, text, New System.Text.UTF8Encoding(False, True))
                If System.IO.File.Exists(path) Then
                    System.IO.File.Replace(temporary, path, Nothing)
                Else
                    System.IO.File.Move(temporary, path)
                End If
            Finally
                If System.IO.File.Exists(temporary) Then System.IO.File.Delete(temporary)
            End Try
        End Sub

        Private Sub DownloadStarting(sender As System.Object, e As Microsoft.Web.WebView2.Core.CoreWebView2DownloadStartingEventArgs)
            ' Keep HTML/ZIP/image exports, but never silently download to a default folder.
            Using dialog As New System.Windows.Forms.SaveFileDialog() With {.FileName = System.IO.Path.GetFileName(e.ResultFilePath), .OverwritePrompt = True}
                If dialog.ShowDialog(EditorDialogOwner(dialog.GetType().FullName)) = System.Windows.Forms.DialogResult.OK Then
                    e.ResultFilePath = dialog.FileName
                    e.Handled = True
                Else
                    e.Cancel = True
                End If
            End Using
        End Sub

        Private Sub FileDragEnter(sender As System.Object, e As System.Windows.Forms.DragEventArgs)
            If e.Data.GetDataPresent(System.Windows.Forms.DataFormats.FileDrop) Then e.Effect = System.Windows.Forms.DragDropEffects.Copy
        End Sub

        Private Async Sub FileDragDrop(sender As System.Object, e As System.Windows.Forms.DragEventArgs)
            Try
                Dim paths = TryCast(e.Data.GetData(System.Windows.Forms.DataFormats.FileDrop), System.String())
                If paths Is Nothing Then Return
                For Each path In paths
                    If IsTextPath(path) Then Await ImportPathAsync(path)
                Next
                SetCompact(False)
            Catch ex As System.Exception
                Using SharedMethods.PushDialogOwner(Me)
                    SharedMethods.ShowCustomMessageBox("File drop failed: " & ex.Message, Text)
                End Using
            End Try
        End Sub

        Private Sub ClampToWorkingArea()
            Dim area = System.Windows.Forms.Screen.FromRectangle(Bounds).WorkingArea
            If Not _compact Then
                MinimumSize = New System.Drawing.Size(System.Math.Min(560, area.Width), System.Math.Min(360, area.Height))
                Size = New System.Drawing.Size(System.Math.Min(Width, area.Width), System.Math.Min(Height, area.Height))
            End If
            Location = New System.Drawing.Point(System.Math.Max(area.Left, System.Math.Min(Left, area.Right - Width)), System.Math.Max(area.Top, System.Math.Min(Top, area.Bottom - Height)))
        End Sub

        Private Shared Function RecoveryPathForHost(host As System.String) As System.String
            Return System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData), "RedInk", "MarkdownEditor", host, "Recovery.json")
        End Function

        Private Sub PersistRecovery()
            Dim path = RecoveryPathForHost(_host)
            If Not _recoveryLoaded Then
                If System.IO.File.Exists(path) Then Throw New System.IO.IOException("Recovery data has not been loaded; it will not be overwritten.")
                Return
            End If
            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path))
            If _recovery.Count = 0 Then
                If System.IO.File.Exists(path) Then System.IO.File.Delete(path)
            Else
                AtomicWrite(path, New Newtonsoft.Json.Linq.JArray(_recovery.Values).ToString(Newtonsoft.Json.Formatting.None))
            End If
        End Sub

        Public Shared Sub HostClosing(host As System.String)
            Dim editor As MarkdownEditorForm = Nothing
            If Not Instances.TryGetValue(host, editor) OrElse editor.IsDisposed Then Return
            ' Quit is too late for async WebView/UI handshakes. Persist the latest acknowledged edit;
            ' never implicitly overwrite a linked source file or block Office waiting for its UI thread.
            Try
                editor._hostClosing = True
                editor._recoveryTimer.Stop()
                editor.SaveState()
                editor.PersistRecovery()
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("[MarkdownEditor] Host recovery checkpoint failed: " & ex.ToString())
                SharedMethods.ShowCustomMessageBox("Markdown recovery could not be saved. Already autosaved notes remain in the editor profile. " & ex.Message, "Red Ink - Markdown Editor")
            End Try
        End Sub

        Private Sub SetCompact(value As System.Boolean)
            If value = _compact OrElse _resizing Then Return
            _resizing = True
            Try
                If value Then
                    _normalBounds = If(WindowState = System.Windows.Forms.FormWindowState.Normal, Bounds, RestoreBounds)
                    WindowState = System.Windows.Forms.FormWindowState.Normal
                    _compact = True
                    _web.Visible = False
                    _bar.Visible = False
                    FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
                    MinimumSize = System.Drawing.Size.Empty
                    Dim edge = CInt(System.Math.Ceiling(72.0 * DeviceDpi / 96.0))
                    ClientSize = New System.Drawing.Size(edge, edge)
                    BackColor = System.Drawing.Color.Magenta
                    TransparencyKey = System.Drawing.Color.Magenta
                    _floating.Visible = True
                    _floating.BringToFront()
                    ClampToWorkingArea()
                Else
                    _compact = False
                    _floating.Visible = False
                    TransparencyKey = System.Drawing.Color.Empty
                    BackColor = System.Drawing.SystemColors.Control
                    FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
                    MinimumSize = New System.Drawing.Size(560, 360)
                    _bar.Visible = True
                    Bounds = New System.Drawing.Rectangle(Location, _normalBounds.Size)
                    _web.Visible = True
                    _compactButton.Text = "Compact"
                    ClampToWorkingArea()
                End If
                Text = "Red Ink - Markdown Editor"
            Finally
                _resizing = False
            End Try
        End Sub

        Protected Overrides Sub OnResize(e As System.EventArgs)
            MyBase.OnResize(e)
            If _bar Is Nothing OrElse _resizing Then Return
            If WindowState = System.Windows.Forms.FormWindowState.Minimized Then
                If _compact Then
                    _resizing = True
                    Try
                        WindowState = System.Windows.Forms.FormWindowState.Normal
                    Finally
                        _resizing = False
                    End Try
                Else
                    SetCompact(True)
                End If
            End If
        End Sub

        Private Async Sub ClosingAsync(sender As System.Object, e As System.Windows.Forms.FormClosingEventArgs)
            If _allowClose OrElse _hostClosing Then Return
            e.Cancel = True
            If _closing Then Return
            _closing = True
            Try
                If _fileBusy Then Throw New System.IO.IOException("A file action is still active; finish it first.")
                If _ready Then
                    _flushCompletion = New System.Threading.Tasks.TaskCompletionSource(Of System.Boolean)()
                    Await _web.CoreWebView2.ExecuteScriptAsync("window.redInkFlush().then(ok => chrome.webview.postMessage({command:'flushed',ok})).catch(error => chrome.webview.postMessage({command:'flushed',ok:false}));")
                    Dim completed = Await System.Threading.Tasks.Task.WhenAny(_flushCompletion.Task, System.Threading.Tasks.Task.Delay(5000))
                    If completed IsNot _flushCompletion.Task OrElse Not _flushCompletion.Task.Result Then Throw New System.IO.IOException("Browser note/file storage did not confirm saving. Export the current note before closing.")
                ElseIf _web.CoreWebView2 IsNot Nothing Then
                    Using SharedMethods.PushDialogOwner(Me)
                        If SharedMethods.ShowCustomYesNoBox("The editor browser is unavailable. Close with the last autosaved state? Recent unsaved edits cannot be confirmed.", "Close", "Keep open", Text) <> 1 Then Return
                    End Using
                End If
                PersistRecovery()
                SaveState()
                Using key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(_registryPath)
                    key.SetValue("Reopen", 0) ' Only a successful explicit editor close cancels startup restoration.
                End Using
                _allowClose = True
                Close()
            Catch ex As System.Exception
                Using SharedMethods.PushDialogOwner(Me)
                    SharedMethods.ShowCustomMessageBox("Editor stays open because saving could not be confirmed: " & ex.Message, Text)
                End Using
            Finally
                _closing = False
            End Try
        End Sub
    End Class

    Partial Public Class SharedMethods
        <System.Runtime.InteropServices.DllImport("user32.dll")>
        Private Shared Function DestroyIcon(handle As System.IntPtr) As System.Boolean
        End Function

        Public Shared Function CreateMarkdownDocumentBitmap(Optional edge As System.Int32 = 32) As System.Drawing.Bitmap
            Dim bitmap As New System.Drawing.Bitmap(edge, edge, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
            Using graphics = System.Drawing.Graphics.FromImage(bitmap)
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias
                graphics.ScaleTransform(edge / 64.0F, edge / 64.0F)
                Using outline As New System.Drawing.Pen(System.Drawing.Color.FromArgb(55, 65, 81), 2.0F), paper As New System.Drawing.SolidBrush(System.Drawing.Color.White), accent As New System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(180, 35, 24))
                    Dim points As System.Drawing.PointF() = {New System.Drawing.PointF(10, 3), New System.Drawing.PointF(41, 3), New System.Drawing.PointF(54, 16), New System.Drawing.PointF(54, 61), New System.Drawing.PointF(10, 61)}
                    graphics.FillPolygon(paper, points)
                    graphics.DrawPolygon(outline, points)
                    graphics.DrawLines(outline, {New System.Drawing.PointF(41, 3), New System.Drawing.PointF(41, 16), New System.Drawing.PointF(54, 16)})
                    Using font As New System.Drawing.Font("Segoe UI", 17.0F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Pixel)
                        graphics.DrawString(".md", font, accent, 13, 27)
                    End Using
                    graphics.DrawLine(outline, 17, 21, 34, 21)
                    graphics.DrawLine(outline, 17, 52, 45, 52)
                End Using
            End Using
            Return bitmap
        End Function

        Public Shared Function CreateMarkdownEditorIcon() As System.Drawing.Icon
            Using bitmap As New System.Drawing.Bitmap(GetLogoBitmap(LogoType.Standard))
                Dim handle = bitmap.GetHicon()
                Try
                    Using borrowed = System.Drawing.Icon.FromHandle(handle)
                        Return DirectCast(borrowed.Clone(), System.Drawing.Icon)
                    End Using
                Finally
                    DestroyIcon(handle)
                End Try
            End Using
        End Function
    End Class
End Namespace
